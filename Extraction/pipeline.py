import argparse
import base64
import csv
import json
import os
import re
import shutil
import sys
from pathlib import Path
import openpyxl
import time
from openai import OpenAI, APIError, APITimeoutError, RateLimitError
from pdf_to_indented_text import pdf_to_indented_text
import functools, builtins

# ── Intercept print() → also write to DuckDB log ─────────────────────────────
_pipeline_log_writer = None   # set in main() after args are parsed
_log_processing_id   = 0 
_original_print = builtins.print
def _next_log_pid(filename: str = "") -> None:
    """Set current filename as ProcessingId on the log writer."""
    if _pipeline_log_writer is not None:
        _pipeline_log_writer.set_context(
            processing_id = filename,   # ← filename string, e.g. "LG_CIT_WI_600005841_2024.pdf"
            stage         = "p",
        )
    return _log_processing_id
def _logging_print(*args, **kwargs):
    kwargs.setdefault("flush", True)
    _original_print(*args, **kwargs)
    if _pipeline_log_writer is not None:
        line = " ".join(str(a) for a in args)
        # Strip the sep/end kwargs if present, write each non-blank line
        for sub in line.split("\n"):
            sub = sub.strip()
            if sub:
                _pipeline_log_writer.write(sub)

builtins.print = _logging_print
#
from concurrent.futures import ThreadPoolExecutor, as_completed
import threading
try:
    from total_check_engine import run_total_check
    _TOTAL_CHECK_AVAILABLE = True
except ImportError:
    _TOTAL_CHECK_AVAILABLE = False
    print("[WARN] total_check_engine.py not found — Total Check will be skipped.")
from compact_schema import build_short_key_instruction, expand_compact_json


_gemini_semaphore = threading.Semaphore(3)
_claude_semaphore = threading.Semaphore(3)

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

try:
    from page_extractor import (
        find_snp_pages, extract_pages_to_pdf,
        process_file, _PRODUCED_RE,
    )
except ImportError:
    sys.exit(
        "[ERROR] page_extractor.py not found next to pipeline.py.\n"
        "        Place page_extractor.py in the same folder and retry."
    )

try:
    from jsonToCsv import run_json_to_csv_pipeline
    from jsonToCsv import merge_deal_csvs_to_excel
except ImportError as e:
    sys.exit(f"[ERROR] Could not import jsonToCsv.py: {e}")


# ─────────────────────────────────────────────────────────────────────────────
# STATEMENT TYPE → PROMPT FILE MAPPING
# ─────────────────────────────────────────────────────────────────────────────

SUFFIX_TO_PROMPT: dict[str, str] = {
    "_SNP":      "SNP_Prompt.txt",
    "_SOA":      "SOA_Prompt.txt",
    "_GOV_BS":   "GOV_BS_Prompt.txt",
    "_GOV_IS":   "GOV_IS_Prompt.txt",
    "_PROP_SNP": "PROP_SNP_Prompt.txt",
    "_PROP_IS":  "PROP_IS_Prompt.txt",
    "_PROP_CFS": "PROP_CFS_Prompt.txt",
    "_DSR":      "DSR_Prompt.txt",
    "_DEBT":     "DEBT_Prompt.txt",
}


TYPE_PATTERNS: list[tuple[re.Pattern, str]] = [
    (re.compile(r"_PROP_SNP", re.IGNORECASE), "_PROP_SNP"),
    (re.compile(r"_PROP_CFS", re.IGNORECASE), "_PROP_CFS"),
    (re.compile(r"_PROP_IS",  re.IGNORECASE), "_PROP_IS"),
    (re.compile(r"_GOV_BS",   re.IGNORECASE), "_GOV_BS"),
    (re.compile(r"_GOV_IS",   re.IGNORECASE), "_GOV_IS"),
    (re.compile(r"_SNP",      re.IGNORECASE), "_SNP"),
    (re.compile(r"_SOA",      re.IGNORECASE), "_SOA"),
    (re.compile(r"_DSR",      re.IGNORECASE), "_DSR"),
    (re.compile(r"_DEBT",     re.IGNORECASE), "_DEBT"),
]


KNOWN_SUFFIXES = tuple(s.lower() for s in SUFFIX_TO_PROMPT)

# ─────────────────────────────────────────────────────────────────────────────
# PROP_SNP / PROP_IS / PROP_CFS MULTI-TABLE SPLIT HELPERS
# ─────────────────────────────────────────────────────────────────────────────

PROP_SNP_TABLE_BREAK = "---PROP_SNP_TABLE_BREAK---"
PROP_IS_TABLE_BREAK  = "---PROP_IS_TABLE_BREAK---"
PROP_CFS_TABLE_BREAK = "---PROP_CFS_TABLE_BREAK---"

STMT_TABLE_BREAK: dict[str, str] = {
    "_PROP_SNP": PROP_SNP_TABLE_BREAK,
    "_PROP_IS":  PROP_IS_TABLE_BREAK,
    "_PROP_CFS": PROP_CFS_TABLE_BREAK,
}

def get_table_break(stmt_type: str) -> str | None:
    return STMT_TABLE_BREAK.get(stmt_type)


def split_multi_table_response(raw_llm_output: str, stmt_type: str) -> list[str]:
    delimiter = get_table_break(stmt_type)
    if delimiter and delimiter in raw_llm_output:
        parts = raw_llm_output.split(delimiter)
        return [p.strip() for p in parts if p.strip()]
    return [raw_llm_output.strip()]


def prop_stmt_output_stem(base_stem: str, metadata: dict, table_index: int) -> str:
    core = re.sub(r"_p[\d\-]+$", "", base_stem)
    page_no = metadata.get("Page No", "")
    pages_str = "p" + "-".join(p.strip() for p in page_no.split(",") if p.strip())
    return f"{core}_{pages_str}"


def save_multi_table_results(
    raw_llm_output: str,
    original_pdf_path: str,
    stmt_type: str,
) -> list[tuple[str, dict]]:
    parts = split_multi_table_response(raw_llm_output, stmt_type)
    base_stem = Path(original_pdf_path).stem

    results = []
    for idx, json_str in enumerate(parts):
        try:
            parsed = parse_json_response(json_str)
        except ValueError as e:
            print(f"  [{stmt_type} SPLIT] JSON parse error on table {idx + 1}: {e}")
            continue

        expanded = expand_compact_json(parsed, stmt_type=stmt_type)
        metadata = expanded.get("Metadata", {})
        out_stem = prop_stmt_output_stem(base_stem, metadata, idx)
        results.append((out_stem, expanded))

    return results

def split_prop_snp_response(raw_llm_output: str) -> list[str]:
    if PROP_SNP_TABLE_BREAK in raw_llm_output:
        parts = raw_llm_output.split(PROP_SNP_TABLE_BREAK)
        return [p.strip() for p in parts if p.strip()]
    return [raw_llm_output.strip()]


def prop_snp_output_stem(base_stem: str, metadata: dict, table_index: int) -> str:
    core = re.sub(r"_p[\d\-]+$", "", base_stem)
    page_no = metadata.get("Page No", "")
    pages_str = "p" + "-".join(p.strip() for p in page_no.split(",") if p.strip())
    return f"{core}_{pages_str}"

def save_prop_snp_results(
    raw_llm_output: str,
    original_pdf_path: str,
    stmt_type: str,
) -> list[tuple[str, dict]]:
    parts = split_prop_snp_response(raw_llm_output)
    base_stem = Path(original_pdf_path).stem

    results = []
    for idx, json_str in enumerate(parts):
        try:
            parsed = parse_json_response(json_str)
        except ValueError as e:
            print(f"  [PROP_SNP SPLIT] JSON parse error on table {idx + 1}: {e}")
            continue

        expanded = expand_compact_json(parsed, stmt_type=stmt_type)
        metadata = expanded.get("Metadata", {})
        out_stem = prop_snp_output_stem(base_stem, metadata, idx)
        results.append((out_stem, expanded))

    return results

_PAGE_TAG_RE = re.compile(
    r"_(?:SNP|SOA|GOV_BS|GOV_IS|PROP_SNP|PROP_IS|PROP_CFS|DSR|DEBT)"
    r"_p(\d+)((?:-\d+)*)$",
    re.IGNORECASE,
)


# ─────────────────────────────────────────────────────────────────────────────
# OPENAI COST TABLE  ($ per 1M tokens)
# ─────────────────────────────────────────────────────────────────────────────

OPENAI_COST_TABLE: dict[str, tuple[float, float]] = {
    "gpt-4o-mini":      (0.15,   0.60),
    "gpt-4.1-mini":     (0.40,   1.60),
    "gpt-4o":           (2.50,  10.00),
    "gpt-4.1":          (2.00,   8.00),
    "gpt-5.4-mini":     (0.40,   1.60),
    "gpt-5.5":          (5.00,  30.00),
}


# ─────────────────────────────────────────────────────────────────────────────
# CLAUDE COST TABLE  ($ per 1M tokens)
# ─────────────────────────────────────────────────────────────────────────────

CLAUDE_COST_TABLE: dict[str, tuple[float, float]] = {
    "claude-sonnet-4-6": (3.00, 15.00),
    "claude-sonnet-4-5": (3.00, 15.00),
    "claude-opus-4-8":   (5.00, 25.00),
    "claude-opus-4-5":   (5.00, 25.00),
}


def calc_claude_cost(model: str, prompt_tokens: int, completion_tokens: int) -> float:
    input_rate, output_rate = CLAUDE_COST_TABLE.get(model, (3.00, 15.00))
    return (
        (prompt_tokens / 1_000_000) * input_rate
        + (completion_tokens / 1_000_000) * output_rate
    )

# ─────────────────────────────────────────────────────────────────────────────
# CLI
# ─────────────────────────────────────────────────────────────────────────────

def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(
        description="End-to-end pipeline: PDF folder → page extraction → JSON normalization.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    p.add_argument("--folder",         default="./03_Validated_Report",
                   help="Folder containing raw PDFs.")
    p.add_argument("--xlsx",           required=True)
    p.add_argument("--prompts",        required=True)
    p.add_argument("--output",         default="./04_Validated_output",
                   help="Output folder for all-PASS results.")
    p.add_argument("--manual-output",  default="./05_Manual_Validation_Required",
                   help="Output folder for FAIL results.")
    p.add_argument(
        "--provider",
        choices=["openai", "gemini", "claude"],
        default="openai",
        help="LLM provider: openai, gemini, or claude.",
    )
    p.add_argument(
        "--tier",
        choices=["free", "paid"],
        default="free",
        help="Tier. Claude always forces paid.",
    )
    p.add_argument(
        "--model",
        default=None,
        help="Model name. If omitted, provider default is used.",
    )
    p.add_argument(
        "--batch",
        action="store_true",
        help="Use async Batch API instead of sync calls. Paid tier only.",
    )
    p.add_argument(
        "--coa-mapping",
        default="",
        help="Folder containing sector-specific COA mapping prompt files.",
    )
    p.add_argument("--max-tokens",         type=int, default=65536)
    p.add_argument("--skip-extraction",    action="store_true")
    p.add_argument("--reporting-columns",  nargs="*", default=None, metavar="COL")
    p.add_argument("--llm-page-id",        action="store_true",
                   help="Use claude-sonnet-4-6 to identify pages before extraction (one call per PDF).")
    p.add_argument("--id-model",           default="claude-sonnet-4-6",
                   help="Model to use for LLM page identification.")
    p.add_argument("--skip-normalization", action="store_true",
                   help="Run page extraction only — do not call any LLM for normalization.")
    return p.parse_args()


# ─────────────────────────────────────────────────────────────────────────────
# MODEL DEFAULTING
# ─────────────────────────────────────────────────────────────────────────────

DEFAULT_MODEL_BY_PROVIDER: dict[str, str] = {
    "openai": "gpt-4o-mini",
    "gemini": "gemini-3.5-flash",
    "claude": "claude-sonnet-4-6",
}


def resolve_model(provider: str, model: str | None) -> str:
    if model:
        return model
    resolved = DEFAULT_MODEL_BY_PROVIDER.get(provider, "gpt-4o-mini")
    print(f"[INFO] --model not specified — defaulting to '{resolved}' for provider '{provider}'")
    return resolved


def cleanup_extracted_pdfs(folder: str):
    deleted = 0
    for f in os.listdir(folder):
        if f.lower().endswith(".pdf") and _PRODUCED_RE.search(f):
            try:
                os.remove(os.path.join(folder, f))
                deleted += 1
            except Exception as e:
                print(f"[WARN] Could not delete {f}: {e}")
    #print(f"\n[Cleanup] Deleted {deleted} extracted PDF(s) from {folder}")


# ─────────────────────────────────────────────────────────────────────────────
# SECTOR DETECTION & TABLE ROUTING
# ─────────────────────────────────────────────────────────────────────────────

def detect_sector(filename: str) -> str:
    stem = Path(filename).stem
    if re.search(r'NON[-_]?LG$', stem, re.IGNORECASE):
        return "NON-LG"
    return "LG"


SECTOR_TABLE_SUFFIXES: dict[str, list[str]] = {
    "LG": [
        "_SNP", "_SOA",
        "_GOV_BS", "_GOV_IS",
        "_PROP_SNP", "_PROP_IS", "_PROP_CFS",
        "_DSR",
        "_DEBT",
    ],
    "NON-LG": [
        "_PROP_SNP",
        "_PROP_IS",
        "_PROP_CFS",
    ],
}


# ─────────────────────────────────────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────────────────────────────────────

def load_xlsx_as_pipe_text(path: str) -> str:
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb.worksheets[0]
    lines = []
    for row in ws.iter_rows(values_only=True):
        if all(c is None or str(c).strip() == "" for c in row):
            continue
        lines.append(" | ".join(str(c).strip() if c is not None else "" for c in row))
    wb.close()
    return "\n".join(lines)


def detect_type(filename: str) -> str | None:
    stem = Path(filename).stem
    for pattern, key in TYPE_PATTERNS:
        if pattern.search(stem):
            return key
    return None


def parse_json_response(raw: str) -> dict:
    text = raw.strip()
    for attempt in [
        lambda t: json.loads(t),
        lambda t: json.loads(re.search(r"```(?:json)?\s*([\s\S]+?)\s*```", t).group(1)),
        lambda t: json.loads(t[t.index("{"):t.rindex("}") + 1]),
    ]:
        try:
            return attempt(text)
        except Exception:
            pass
    raise ValueError(f"Cannot parse JSON. First 400 chars:\n{raw[:400]}")


def calc_openai_cost(model: str, prompt_tokens: int, completion_tokens: int) -> float:
    input_rate, output_rate = OPENAI_COST_TABLE.get(model, (0.0, 0.0))
    return (prompt_tokens / 1_000_000) * input_rate + \
           (completion_tokens / 1_000_000) * output_rate


def extract_page_info_from_filename(pdf_path: str) -> dict | None:
    stem = Path(pdf_path).stem
    m = _PAGE_TAG_RE.search(stem)
    if not m:
        return None

    first_page = int(m.group(1))
    rest = m.group(2)

    if rest:
        all_nums = [first_page] + [int(x) for x in rest.strip("-").split("-") if x]
    else:
        all_nums = [first_page]

    start = min(all_nums)
    end   = max(all_nums)
    pages = sorted(all_nums)

    if start == end:
        label = f"page {start}"
    else:
        label = f"pages {start}–{end}"

    return {"start": start, "end": end, "pages": pages, "label": label}


# ─────────────────────────────────────────────────────────────────────────────
# STEP 1 — PAGE EXTRACTION
# ─────────────────────────────────────────────────────────────────────────────

def get_available_types(prompts_folder: str | None) -> set[str]:
    if not prompts_folder:
        return set(SUFFIX_TO_PROMPT.keys())

    abs_path = os.path.abspath(prompts_folder)
    print(f"\n[Prompts] Checking folder: {abs_path}")

    if not os.path.isdir(prompts_folder):
        print(f"[ERROR] Prompts folder does NOT exist: {abs_path}")
        return set()

    actual_files = os.listdir(prompts_folder)
    print(f"[Prompts] Files in folder ({len(actual_files)} total):")
    for f in sorted(actual_files):
        print(f"           - {f}")

    available = set()
    for suffix, expected_filename in SUFFIX_TO_PROMPT.items():
        full_path = os.path.join(prompts_folder, expected_filename)
        if os.path.isfile(full_path):
            available.add(suffix)
            print(f"[Prompts]   ✓ Found: {expected_filename}")
        else:
            print(f"[Prompts]   ✗ Missing: {expected_filename}")

    missing = set(SUFFIX_TO_PROMPT.keys()) - available
    if missing:
        print(f"[INFO] No prompt file found for: {', '.join(sorted(missing))}")

    return available


def run_extraction(
    folder: str,
    prompts_folder: str = None,
    use_llm_id: bool = False,
    id_model: str = "claude-sonnet-4-6",
) -> list[str]:
    folder_path = Path(folder)

    raw_pdfs = [
        str(folder_path / f)
        for f in sorted(os.listdir(folder))
        if f.lower().endswith(".pdf") and not _PRODUCED_RE.search(f)
    ]

    if not raw_pdfs:
        print(f"[WARN] No raw PDFs found in {folder}")
        return []

    available_prompts: set[str] = set()
    if prompts_folder and os.path.isdir(prompts_folder):
        for suffix, prompt_filename in SUFFIX_TO_PROMPT.items():
            prompt_path = os.path.join(prompts_folder, prompt_filename)
            if os.path.isfile(prompt_path):
                available_prompts.add(suffix)

        missing_prompts = set(SUFFIX_TO_PROMPT.keys()) - available_prompts
        if missing_prompts:
            print(f"[INFO] Prompts folder: {prompts_folder}")
            print(f"[INFO] Available prompts: {sorted(available_prompts)}")
            print(f"[INFO] MISSING prompts (will be skipped): {sorted(missing_prompts)}")
    else:
        print(f"[WARN] Prompts folder not found: {prompts_folder}")
        available_prompts = set(SUFFIX_TO_PROMPT.keys())

    all_extracted: list[str] = []
    total_files = len(raw_pdfs)          # ← ADD

    for file_idx, pdf_path in enumerate(raw_pdfs, start=1):
        fname  = os.path.basename(pdf_path)
        _next_log_pid(Path(fname).stem)                           # ← only this, no manual set_context
        sector = detect_sector(fname)            # ← only once
        sector_allowed  = set(SECTOR_TABLE_SUFFIXES.get(sector, SECTOR_TABLE_SUFFIXES["LG"]))
        effective_allowed = sector_allowed & available_prompts
        skipped_no_prompt = sector_allowed - available_prompts

        if skipped_no_prompt:
            print(f"  Skipped    : {sorted(skipped_no_prompt)}  (no prompt file)")

        if not effective_allowed:
            print(f"  [WARN] No allowed tables for {fname} — skipping extraction entirely.")
            continue

        try:
            if use_llm_id:
                from page_extractor import process_file_llm_guided
                produced = process_file_llm_guided(
                    src              = pdf_path,
                    sector           = sector,
                    allowed_suffixes = effective_allowed,
                    id_model         = id_model,
                    use_llm_id       = True,
                    file_index       = file_idx,
                    file_total       = total_files,
                )
            else:
                produced = process_file(
                    pdf_path,
                    sector           = sector,
                    allowed_suffixes = effective_allowed,
                )
        except TypeError:
            produced = process_file(pdf_path, sector=sector)

        if not produced:
            print(f"  [WARN] No extracted PDFs produced for {fname}")
            continue

        filtered      = []
        removed_count = 0
        for p in produced:
            matched_type = detect_type(p)
            if matched_type is None:
                print(f"  [SKIP] {os.path.basename(p)} (unknown type)")
                try:
                    os.remove(p)
                    removed_count += 1
                except OSError:
                    pass
                continue
            if matched_type not in effective_allowed:
                if matched_type in sector_allowed:
                    reason = f"no prompt file ({SUFFIX_TO_PROMPT.get(matched_type, '?')})"
                else:
                    reason = f"not allowed for sector {sector}"
                print(f"  [SKIP] {os.path.basename(p)} ({reason})")
                try:
                    os.remove(p)
                    removed_count += 1
                except OSError:
                    pass
                continue
            filtered.append(p)

        all_extracted.extend(filtered)
        print(f"  → {sector}: kept {len(filtered)} table(s), removed {removed_count}")
        print("")

    return all_extracted

def collect_existing_extracted(folder: str, prompts_folder: str = None) -> list[str]:
    folder_path = Path(folder)
    found: list[str] = []

    _EXTRACTED_RE = re.compile(
        r"_(?:SNP|SOA|GOV_BS|GOV_IS|PROP_SNP|PROP_IS|PROP_CFS|DSR|DEBT)"
        r"_p\d+(?:-\d+)*\.pdf$",
        re.IGNORECASE,
    )

    available_prompts: set[str] = set()
    if prompts_folder and os.path.isdir(prompts_folder):
        for suffix, prompt_filename in SUFFIX_TO_PROMPT.items():
            if os.path.isfile(os.path.join(prompts_folder, prompt_filename)):
                available_prompts.add(suffix)
    else:
        available_prompts = set(SUFFIX_TO_PROMPT.keys())

    skipped_files = []

    for f in sorted(os.listdir(folder)):
        if not (f.lower().endswith(".pdf") and _EXTRACTED_RE.search(f)):
            continue

        full = str(folder_path / f)
        stem = Path(f).stem
        base_name = get_base_pdf_name(stem)
        sector = detect_sector(base_name + ".pdf")
        sector_allowed = set(SECTOR_TABLE_SUFFIXES.get(sector, SECTOR_TABLE_SUFFIXES["LG"]))
        effective_allowed = sector_allowed & available_prompts

        matched_type = detect_type(f)

        if matched_type and matched_type in effective_allowed:
            found.append(full)
        else:
            if matched_type not in sector_allowed:
                reason = f"not needed for sector {sector}"
            elif matched_type not in available_prompts:
                reason = f"no prompt file"
            else:
                reason = "unknown type"
            skipped_files.append((f, reason))

    if skipped_files:
        for fn, reason in skipped_files:
            print(f"  [SKIP] {fn} — {reason}")

    #print(f"[Skip-extraction] Found {len(found)} previously extracted PDFs "
          #f"(skipped {len(skipped_files)})")
    return found

# ─────────────────────────────────────────────────────────────────────────────
# STEP 2a — NORMALIZE via OPENAI
# ─────────────────────────────────────────────────────────────────────────────

def normalize_one_openai(
    pdf_path: str,
    prompt_path: str,
    coa_text: str,
    reporting_columns: list | None,
    client,
    model: str,
    max_tokens: int,
    page_info: dict | None = None,
    stmt_type: str = "UNKNOWN",
    coa_map_rules: str = "",
) -> dict:
    system_prompt = Path(prompt_path).read_text(encoding="utf-8", errors="replace")
    system_prompt = system_prompt + "\n\n" + build_short_key_instruction(stmt_type)
    if coa_map_rules:
        system_prompt = system_prompt + "\n\n" + coa_map_rules
    pdf_bytes     = Path(pdf_path).read_bytes()
    b64           = base64.b64encode(pdf_bytes).decode("utf-8")
    pdf_data_uri  = f"data:application/pdf;base64,{b64}"
    pdf_name      = Path(pdf_path).name

    if page_info is None:
        page_info = extract_page_info_from_filename(pdf_path)

    try:
        indented_text = pdf_to_indented_text(pdf_path)
    except Exception as e:
        #print(f"  [WARN] pdf_to_indented_text failed ({e}), falling back to no indented text")
        indented_text = None

    instruction_lines = [
        "Normalize the financial statement from the attached PDF.",
        "Follow all steps in the system prompt exactly.",
        "Use the COA Master table above for COA Datapoint mapping.",
        "Return valid JSON only — no markdown, no extra text.",
    ]
    if reporting_columns:
        instruction_lines += ["", "[[REPORTING COLUMNS]]"] + reporting_columns

    user_content = [
        {"type": "file", "file": {"filename": pdf_name, "file_data": pdf_data_uri}},
        {
            "type": "text",
            "text": (
                "STANDARD COA MASTER (pipe-delimited):\n"
                "Format: Statement | Section | COA Flag | COA Datapoint\n\n"
                + coa_text
            ),
        },
    ]

    if page_info:
        page_note = build_page_note(page_info)
        user_content.append({"type": "text", "text": page_note})
        #print(f"  Source pages : {page_info['label']}")
    # else:
    #     print(f"  Source pages : [WARN] no page tag found in filename — skipping page context")

    if indented_text:
        user_content.append({
            "type": "text",
            "text": (
                "INDENTED TEXT EXTRACTION OF THE PDF (HIERARCHY-PRESERVING):\n"
                "The following is the financial statement text extracted with visual\n"
                "indentation preserved. Leading spaces indicate hierarchy level:\n"
                "  0 spaces = root level\n"
                "  2 spaces = level-1 child\n"
                "  4 spaces = level-2 grandchild\n"
                "  6 spaces = total/subtotal row\n\n"
                "USE THIS INDENTED TEXT (NOT the raw PDF) for Section E hierarchy "
                "flattening. Trust the leading spaces -- do not override them with "
                "semantic reasoning about label names.\n\n"
                "--- BEGIN INDENTED TEXT ---\n"
                + indented_text
                + "\n--- END INDENTED TEXT ---"
            ),
        })

    user_content.append({
        "type": "text",
        "text": "\n".join(instruction_lines),
    })

    messages = [
        {"role": "system", "content": system_prompt},
        {"role": "user", "content": user_content},
    ]

    try:
        response = client.chat.completions.create(
            model=model,
            max_completion_tokens=max_tokens,
            messages=messages,
        )
    except APITimeoutError:
        raise RuntimeError("OpenAI request timed out.")
    except RateLimitError:
        raise RuntimeError("Rate limit hit — wait and retry.")
    except APIError as e:
        raise RuntimeError(f"OpenAI API error: {e}")

    u = response.usage
    prompt_tokens     = u.prompt_tokens     or 0
    completion_tokens = u.completion_tokens or 0
    total_tokens      = u.total_tokens      or 0
    print(f"  Tokens — prompt: {prompt_tokens:,}  completion: {completion_tokens:,}  total: {total_tokens:,}")

    raw = (response.choices[0].message.content or "").strip()

    delimiter = get_table_break(stmt_type)
    if delimiter and delimiter in raw:
        parts = split_multi_table_response(raw, stmt_type)
        split_tables = []
        for idx, part in enumerate(parts):
            try:
                parsed = parse_json_response(part)
                expanded = expand_compact_json(parsed, stmt_type=stmt_type)
                split_tables.append(expanded)
            except Exception as e:
                print(f"  [{stmt_type} SPLIT] Parse error on table {idx + 1}: {e}")
        return {
            "data":           split_tables[0] if split_tables else None,
            "prop_snp_split": split_tables,
            "usage": {
                "prompt_tokens":     prompt_tokens,
                "completion_tokens": completion_tokens,
                "total_tokens":      total_tokens,
            },
        }

    parsed_data = parse_json_response(raw)
    expanded_data = expand_compact_json(parsed_data, stmt_type=stmt_type)

    return {
        "data":           expanded_data,
        "prop_snp_split": [],
        "usage": {
            "prompt_tokens":     prompt_tokens,
            "completion_tokens": completion_tokens,
            "total_tokens":      total_tokens,
        },
    }


# ─────────────────────────────────────────────────────────────────────────────
# STEP 2b — NORMALIZE via GEMINI
# ─────────────────────────────────────────────────────────────────────────────

def normalize_one_gemini_wrapper(
    pdf_path: str,
    prompt_path: str,
    coa_text: str,
    reporting_columns: list[str] | None,
    model: str,
    max_tokens: int,
    page_info: dict | None = None,
    stmt_type: str = "UNKNOWN",
    coa_map_rules: str = "",
) -> dict:
    try:
        from gemini_client import normalize_one_gemini
    except ImportError as e:
        raise RuntimeError(f"gemini_client.py not found: {e}")

    if page_info is None:
        page_info = extract_page_info_from_filename(pdf_path)

    # if page_info:
    #     print(f"  Source pages : {page_info['label']}")
    # else:
    #     print(f"  Source pages : [WARN] no page tag found in filename — skipping page context")

    try:
        result = normalize_one_gemini(
            pdf_path=pdf_path,
            prompt_path=prompt_path,
            coa_text=coa_text,
            reporting_columns=reporting_columns,
            model=model,
            max_tokens=max_tokens,
            page_info=page_info,
            stmt_type=stmt_type,
            coa_map_rules=coa_map_rules,
        )
    except RuntimeError:
        raise
    except Exception as e:
        raise RuntimeError(f"  Unexpected Gemini error: {type(e).__name__}: {e}")

    result.setdefault("prop_snp_split", [])
    return result


# ─────────────────────────────────────────────────────────────────────────────
# STEP 2c — NORMALIZE via CLAUDE
# ─────────────────────────────────────────────────────────────────────────────

def normalize_one_claude_wrapper(
    pdf_path: str,
    prompt_path: str,
    coa_text: str,
    reporting_columns: list[str] | None,
    model: str,
    max_tokens: int,
    page_info: dict | None = None,
    stmt_type: str = "UNKNOWN",
    coa_map_rules: str = "",
) -> dict:
    """
    Claude paid sync call using prompt caching.
    Claude has no free tier in this project.
    """
    try:
        from claude_cache_client import normalize_one_claude_cached
    except ImportError as e:
        raise RuntimeError(f"claude_cache_client.py not found/importable: {e}")

    result = normalize_one_claude_cached(
        pdf_path=pdf_path,
        prompt_path=prompt_path,
        coa_text=coa_text,
        reporting_columns=reporting_columns,
        model=model,
        max_tokens=max_tokens,
        page_info=page_info,
        stmt_type=stmt_type,
        coa_map_rules=coa_map_rules,
    )
    # Ensure prop_snp_split key always present, same as Gemini wrapper.
    result.setdefault("prop_snp_split", [])
    return result


def get_base_pdf_name(stem: str) -> str:
    result = re.sub(r"_p\d+(?:-\d+)*$", "", stem)
    result = re.sub(
        r"_(PROP_SNP|PROP_CFS|PROP_IS|GOV_BS|GOV_IS|SNP|SOA|DSR|DEBT)$",
        "", result, flags=re.IGNORECASE
    )
    return result

SUFFIX_TO_COA_KEYWORD: dict[str, str] = {
    "_SNP":      "SNP",
    "_SOA":      "SOA",
    "_GOV_BS":   "GOV_BS",
    "_GOV_IS":   "GOV_IS",
    "_PROP_SNP": "PROP_SNP",
    "_PROP_IS":  "PROP_IS",
    "_PROP_CFS": "PROP_CFS",
    "_DSR":      "",
    "_DEBT":     "",
}
def load_coa_mapping(coa_mapping_folder: str, sector: str, stmt_type: str, silent: bool = False) -> str:
    if not coa_mapping_folder:
        return ""
    path = os.path.join(coa_mapping_folder, sector,
                        f"{stmt_type.lstrip('_')}_COA.txt")
    if not os.path.isfile(path):
        if stmt_type.lstrip("_") not in ("DEBT", "DSR") and not silent:
            print(f"  [COA-MAP] No mapping file for sector={sector} "
                  f"stmt_type={stmt_type} — skipping (path: {path})")
        return ""
    text = Path(path).read_text(encoding="utf-8", errors="replace")
    #if not silent:
        #print(f"  [COA-MAP] Loaded mapping: {os.path.relpath(path)}")
    return text

def filter_coa_for_type(coa_text: str, stmt_type: str, sector: str = "ALL") -> str:
    keyword = SUFFIX_TO_COA_KEYWORD.get(stmt_type, "")
    if not keyword:
        return ""

    lines = coa_text.split('\n')
    header = lines[0]
    filtered = [header]

    for line in lines[1:]:
        if not line.strip():
            continue
        parts = line.split(" | ")
        if len(parts) < 1:
            continue
        stmt_col = parts[0].strip()
        stmt_types_in_row = [s.strip() for s in stmt_col.split("|")]
        if keyword in stmt_types_in_row:
            filtered.append(line)

    return '\n'.join(filtered)


def compress_coa_text(coa_text: str) -> str:
    lines = coa_text.split("\n")
    if not lines:
        return coa_text
    body = []
    for line in lines[1:]:
        if not line.strip():
            continue
        parts = line.split(" | ")
        if len(parts) < 2:
            continue
        body.append(f"{parts[1]} | {parts[2]}")
    return "\n".join(body)

def dynamic_coa_filter(coa_text: str, page_text: str) -> str:
    keywords = set(
        word.lower()
        for word in page_text.split()
        if len(word) > 4
    )
    lines = coa_text.split("\n")
    header = lines[0]
    selected = [header]
    for line in lines[1:]:
        if any(k in line.lower() for k in keywords):
            selected.append(line)
    return "\n".join(selected)


# ─────────────────────────────────────────────────────────────────────────────
# WORKER — process_one_pdf
# ─────────────────────────────────────────────────────────────────────────────

def process_one_pdf(args_tuple):
    """
    Worker function for parallel execution.
    Routes to the correct LLM path based on (provider, tier).

    BUG FIX: The Claude branch previously used `return` inside
    `with _claude_semaphore`, which bypassed the common return block at
    the bottom of the try and caused prop_snp_split / cached_hit / cost
    to be lost. Now Claude stores into `result` exactly like the other
    providers and falls through to the shared return.
    """
    (pdf_path, prompt_path, coa_text, model, max_tokens,
     provider, client, reporting_columns, tier, stmt_type,
     coa_map_rules) = args_tuple

    pdf_name  = os.path.basename(pdf_path)
    page_info = extract_page_info_from_filename(pdf_path)

    try:
        # ── PATH 1 — OPENAI PAID (context caching) ───────────────────────
        if provider == "openai" and tier == "paid":
            try:
                from openai_cache_client import normalize_one_openai_cached
            except ImportError as e:
                raise RuntimeError(f"openai_cache_client.py not found: {e}")
            print(f"   [PAID-OPENAI] Cache-aware call → {pdf_name}")
            result = normalize_one_openai_cached(
                pdf_path=pdf_path,
                prompt_path=prompt_path,
                coa_text=coa_text,
                reporting_columns=reporting_columns,
                model=model,
                max_tokens=max_tokens,
                page_info=page_info,
                stmt_type=stmt_type,
                coa_map_rules=coa_map_rules,
            )

        # ── PATH 2 — OPENAI FREE ─────────────────────────────────────────
        elif provider == "openai":
            #print(f"   [FREE-OPENAI] Plain call → {pdf_name}")
            result = normalize_one_openai(
                pdf_path=pdf_path,
                prompt_path=prompt_path,
                coa_text=coa_text,
                reporting_columns=reporting_columns,
                client=client,
                model=model,
                max_tokens=max_tokens,
                page_info=page_info,
                stmt_type=stmt_type,
                coa_map_rules=coa_map_rules,
            )

        # ── PATH 3 — GEMINI PAID (cached, with fallback) ─────────────────
        elif provider == "gemini" and tier == "paid":
            try:
                from gemini_cache_client import (
                    normalize_one_gemini_cached,
                    CacheUnavailableError,
                )
            except ImportError as e:
                print(f"   [PAID→FREE FALLBACK] gemini_cache_client.py not "
                      f"found ({e}). Using plain Gemini call.")
                result = normalize_one_gemini_wrapper(
                    pdf_path=pdf_path,
                    prompt_path=prompt_path,
                    coa_text=coa_text,
                    reporting_columns=reporting_columns,
                    model=model,
                    max_tokens=max_tokens,
                    page_info=page_info,
                    stmt_type=stmt_type,
                    coa_map_rules=coa_map_rules,
                )
            else:
                print(f"   [PAID-GEMINI] Cache-aware call → {pdf_name}")
                try:
                    result = normalize_one_gemini_cached(
                        pdf_path=pdf_path,
                        prompt_path=prompt_path,
                        coa_text=coa_text,
                        reporting_columns=reporting_columns,
                        model=model,
                        max_tokens=max_tokens,
                        page_info=page_info,
                        stmt_type=stmt_type,
                        coa_map_rules=coa_map_rules,
                    )
                except CacheUnavailableError as e:
                    print(f"   [PAID→FREE FALLBACK] Caching unavailable ({e}). Using plain Gemini call.")
                    result = normalize_one_gemini_wrapper(
                        pdf_path=pdf_path,
                        prompt_path=prompt_path,
                        coa_text=coa_text,
                        reporting_columns=reporting_columns,
                        model=model,
                        max_tokens=max_tokens,
                        page_info=page_info,
                        stmt_type=stmt_type,
                        coa_map_rules=coa_map_rules,
                    )
                except Exception as e:
                    print(f"   [PAID→FREE FALLBACK] Gemini cache path failed unexpectedly "
                          f"({type(e).__name__}: {e}). Using plain Gemini call.")
                    result = normalize_one_gemini_wrapper(
                        pdf_path=pdf_path,
                        prompt_path=prompt_path,
                        coa_text=coa_text,
                        reporting_columns=reporting_columns,
                        model=model,
                        max_tokens=max_tokens,
                        page_info=page_info,
                        stmt_type=stmt_type,
                        coa_map_rules=coa_map_rules,
                    )

        # ── PATH 4 — GEMINI FREE ─────────────────────────────────────────
        elif provider == "gemini":
            print(f"   [FREE-GEMINI] Plain call → {pdf_name}")
            result = normalize_one_gemini_wrapper(
                pdf_path=pdf_path,
                prompt_path=prompt_path,
                coa_text=coa_text,
                reporting_columns=reporting_columns,
                model=model,
                max_tokens=max_tokens,
                page_info=page_info,
                stmt_type=stmt_type,
                coa_map_rules=coa_map_rules,
            )

        # ── PATH 5 — CLAUDE PAID (prompt caching) ────────────────────────
        # FIX: store into `result`, do NOT return early.
        # Returning inside the semaphore context skips the shared return
        # block below and loses prop_snp_split, cached_hit, and cost.
        elif provider == "claude":
            if tier != "paid":
                print("[CLAUDE] Claude selected — forcing paid tier.")
                tier = "paid"
            print(f"   [PAID-CLAUDE] Cache-aware call → {pdf_name}")
            with _claude_semaphore:
                result = normalize_one_claude_wrapper(
                    pdf_path=pdf_path,
                    prompt_path=prompt_path,
                    coa_text=coa_text,
                    reporting_columns=reporting_columns,
                    model=model,
                    max_tokens=max_tokens,
                    page_info=page_info,
                    stmt_type=stmt_type,
                    coa_map_rules=coa_map_rules,
                )

        else:
            raise ValueError(f"Unknown provider: {provider}")

        # ── Shared return block (ALL providers reach here) ────────────────
        return {
            "pdf":            pdf_name,
            "ok":             True,
            "data":           result.get("data"),
            "prop_snp_split": result.get("prop_snp_split", []),
            "usage":          result.get("usage", {}),
            "cost":           result.get("cost_usd", 0.0),
            "cached_hit":     result.get("cached_hit", False),
            "model":          result.get("model", model),
        }

    except Exception as e:
        print(f"   [ERROR] {pdf_name}: {e}")
        return {
            "pdf":            pdf_name,
            "ok":             False,
            "error":          str(e),
            "data":           None,
            "prop_snp_split": [],
            "usage":          {},
            "cost":           0.0,
            "cached_hit":     False,
            "model":          model,
        }


def make_batch_custom_id(pdf_path: str) -> str:
    fname = Path(pdf_path).name
    return re.sub(r"[^a-zA-Z0-9_\-]", "_", fname)


# ─────────────────────────────────────────────────────────────────────────────
# BATCH NORMALIZATION
# ─────────────────────────────────────────────────────────────────────────────

def run_normalization_batch(
    extracted_pdfs: list[str],
    prompts_folder: str,
    xlsx_path: str,
    output_folder: str,
    manual_output: str,
    raw_folder: str,
    provider: str,
    model: str,
    max_tokens: int,
    reporting_columns,
    coa_mapping_folder: str = "",
) -> None:
    """
    Async Batch-API path (50% cost, ≤24 h turnaround).

    NOTE: "Total Check Status" is intentionally NOT computed during JSON
    parsing/saving. It is CSV-only output computed by run_json_to_csv_pipeline().
    """
    os.makedirs(output_folder, exist_ok=True)
    os.makedirs(manual_output, exist_ok=True)

    coa_full = load_xlsx_as_pipe_text(xlsx_path)
    print(f"\n[COA] Loaded {len(coa_full.splitlines())} rows from {xlsx_path}")

    # ── Build job list ────────────────────────────────────────────────────────
    jobs: list[dict] = []

    for pdf_path in extracted_pdfs:
        stmt_type = detect_type(os.path.basename(pdf_path))
        if not stmt_type:
            print(f"[BATCH-SKIP] Unknown type: {os.path.basename(pdf_path)}")
            continue

        prompt_filename = SUFFIX_TO_PROMPT.get(stmt_type)
        if not prompt_filename:
            print(f"[BATCH-SKIP] No prompt mapping for {stmt_type}")
            continue

        prompt_path = os.path.join(prompts_folder, prompt_filename)
        if not os.path.exists(prompt_path):
            print(f"[BATCH-SKIP] Prompt file missing: {prompt_path}")
            continue

        prompt_text   = Path(prompt_path).read_text(encoding="utf-8", errors="replace")
        sector        = detect_sector(os.path.basename(pdf_path))
        coa_filtered  = filter_coa_for_type(coa_full, stmt_type, sector)
        coa_map_rules = load_coa_mapping(coa_mapping_folder, sector, stmt_type, silent=True)

        cid       = make_batch_custom_id(pdf_path)
        page_info = extract_page_info_from_filename(pdf_path)

        jobs.append({
            "custom_id":          cid,
            "pdf_path":           pdf_path,
            "prompt_text":        prompt_text,
            "coa_text":           coa_filtered,
            "coa_map_rules":      coa_map_rules,
            "reporting_columns":  reporting_columns,
            "stmt_type":          stmt_type,
            "max_tokens":         max_tokens,
            "page_info":          page_info,
            "coa_mapping_folder": coa_mapping_folder,
        })
    # ── Print COA mapping summary (deduplicated) ──────────────────────────────
    loaded_combos = {(j["stmt_type"], detect_sector(j["pdf_path"])) for j in jobs if j["coa_map_rules"]}
    # for stmt_t, sec in sorted(loaded_combos):
    #     print(f"  [COA-MAP] Loaded mapping: {sec}/{stmt_t.lstrip('_')}_COA.txt")
    # print(f"[BATCH] COA mappings loaded for {len(loaded_combos)} unique (sector, type) combination(s).")

    if not jobs:
        print("[BATCH] No jobs to submit — exiting.")
        return

    print(f"\n[BATCH] Submitting {len(jobs)} job(s) to {provider.upper()} Batch API ...")

    # ── Submit and wait ───────────────────────────────────────────────────────
    if provider == "openai":
        from openai_batch_client import submit_openai_batch, wait_and_download_openai
        batch_id = submit_openai_batch(jobs, model)
        raw_results: dict[str, dict] = wait_and_download_openai(batch_id, model)

    elif provider == "gemini":
        from gemini_batch_client import submit_gemini_batch, wait_and_download_gemini
        batch_name = submit_gemini_batch(jobs, model)
        raw_results: dict[str, dict] = wait_and_download_gemini(batch_name, model)

    elif provider == "claude":
        # ── FIX: use raw_results (not batch_results) consistently ────────
        from claude_batch_client import submit_claude_batch, wait_and_download_claude

        batch_id = submit_claude_batch(
            jobs=jobs,
            model=model,
            display_name="fs-claude-batch",
        )

        raw_results: dict[str, dict] = wait_and_download_claude(
            batch_id=batch_id,
            model=model,
        )

        # Re-expand compact JSON using each job's real stmt_type.
        # (claude_batch_client expanded with "" fallback — redo with real type)
        from compact_schema import expand_compact_json as _expand
        job_by_id = {j["custom_id"]: j for j in jobs}
        for cid, res in raw_results.items():
            if res.get("ok") and res.get("data") is not None and cid in job_by_id:
                res["data"] = _expand(
                    res["data"],
                    job_by_id[cid].get("stmt_type", ""),
                )

    else:
        raise RuntimeError(
            f"[BATCH] provider='{provider}' does not support Batch API.\n"
            "        Supported: openai, gemini, claude."
        )

    # print(f"[BATCH] Received {len(raw_results)} result(s).")

    # ── Group results by deal ─────────────────────────────────────────────────
    deal_groups: dict[str, list[str]] = {}
    for pdf_path in extracted_pdfs:
        base = get_base_pdf_name(Path(pdf_path).stem)
        deal_groups.setdefault(base, []).append(pdf_path)

    total_deals = len(deal_groups)
    pass_count  = 0
    fail_count  = 0
    total_cost  = 0.0

    print(f"\n{'='*70}")
    print(f"  BATCH POST-PROCESSING: {total_deals} deal(s)")
    print(f"{'='*70}")

    for deal_idx, (deal_name, deal_pdfs) in enumerate(deal_groups.items(), 1):
      try:
        sector   = detect_sector(deal_name + ".pdf")
        expected = SECTOR_TABLE_SUFFIXES.get(sector, SECTOR_TABLE_SUFFIXES["LG"])

        _next_log_pid(deal_name) 
        print(f"\n{'─'*70}")
        print(f"  [{deal_idx}/{total_deals}] Deal : {deal_name} || Sector: {sector}")

        sector_allowed_pdfs = [
            p for p in deal_pdfs if detect_type(p) in expected
        ]

        deal_passed = True
        deal_cost   = 0.0
        deal_json_paths: list[str] = []

        for pdf_path in sector_allowed_pdfs:
            fname     = os.path.basename(pdf_path)
            cid       = make_batch_custom_id(pdf_path)
            res       = raw_results.get(cid)
            stmt_type = detect_type(fname) or "UNKNOWN"
            base_stem = Path(pdf_path).stem

            if res is None:
                print(f"    ✘ MISSING  : {fname}  (no result returned from batch)")
                deal_passed = False
                continue

            if not res.get("ok"):
                print(f"    ✘ FAIL     : {fname}  — {res.get('error', 'unknown error')}")
                deal_passed = False
                continue

            cost = res.get("cost_usd", 0.0)
            deal_cost += cost

            # ── Route by result shape ─────────────────────────────────────
            # Claude batch client (claude_batch_client.py) returns already-
            # parsed dicts in "data" / "prop_snp_split" — no "json_text".
            # OpenAI / Gemini batch clients return raw "json_text" string.
            # Detect which shape we have and handle each path cleanly.

            if provider == "claude":
                # ── CLAUDE PATH: data already parsed & expanded ───────────
                prop_snp_split = res.get("prop_snp_split", [])
                split_stmt_types = {"_PROP_SNP", "_PROP_IS", "_PROP_CFS"}

                if stmt_type in split_stmt_types and prop_snp_split:
                    # Multi-table split result
                    for idx, sub_data in enumerate(prop_snp_split):
                        out_stem = prop_stmt_output_stem(
                            base_stem, sub_data.get("Metadata", {}), idx
                        )
                        sub_filename   = f"{out_stem}.json"
                        temp_json_path = os.path.join(
                            os.path.dirname(pdf_path), sub_filename
                        )
                        with open(temp_json_path, "w", encoding="utf-8") as fh:
                            json.dump(sub_data, fh, indent=2, ensure_ascii=False)
                        deal_json_paths.append(temp_json_path)
                        #print(f"    ✔ PASS  : {sub_filename}  [{stmt_type} SPLIT]  "
                              #f"(cost ≈ ${cost:.4f})")
                else:
                    # Single-table result
                    data = res.get("data")
                    if not data:
                        print(f"    ✘ FAIL     : {fname}  — no data in Claude result")
                        deal_passed = False
                        continue

                    json_filename  = f"{base_stem}.json"
                    temp_json_path = os.path.join(
                        os.path.dirname(pdf_path), json_filename
                    )
                    with open(temp_json_path, "w", encoding="utf-8") as fh:
                        json.dump(data, fh, indent=2, ensure_ascii=False)
                    deal_json_paths.append(temp_json_path)
                    cache_note = "  [cache HIT]" if res.get("cached_hit") else ""
                    #print(f"    ✔ PASS  : {fname}  (cost ≈ ${cost:.4f}){cache_note}")

            else:
                # ── OPENAI / GEMINI PATH: raw json_text string ────────────
                json_text = res.get("json_text", "")
                if not json_text:
                    print(f"    ✘ FAIL     : {fname}  — empty json_text in result")
                    deal_passed = False
                    continue

                delimiter = get_table_break(stmt_type)
                if delimiter and delimiter in json_text:
                    split_results = save_multi_table_results(json_text, pdf_path, stmt_type)
                    if not split_results:
                        print(f"    ✘ PARSE-FAIL : {fname}  — split delimiter detected "
                              f"but no sub-table parsed successfully")
                        deal_passed = False
                        continue

                    for out_stem, sub_data in split_results:
                        sub_filename   = f"{out_stem}.json"
                        temp_json_path = os.path.join(
                            os.path.dirname(pdf_path), sub_filename
                        )
                        with open(temp_json_path, "w", encoding="utf-8") as fh:
                            json.dump(sub_data, fh, indent=2, ensure_ascii=False)
                        deal_json_paths.append(temp_json_path)
                        #print(f"    ✔ PASS  : {sub_filename}  [{stmt_type} SPLIT]  "
                              #f"(cost ≈ ${cost:.4f})")

                else:
                    try:
                        parsed   = parse_json_response(json_text)
                        expanded = expand_compact_json(parsed, stmt_type=stmt_type)
                    except Exception as e:
                        print(f"    ✘ PARSE-FAIL : {fname}  — {e}")
                        deal_passed = False
                        continue

                    json_filename  = f"{base_stem}.json"
                    temp_json_path = os.path.join(
                        os.path.dirname(pdf_path), json_filename
                    )
                    with open(temp_json_path, "w", encoding="utf-8") as fh:
                        json.dump(expanded, fh, indent=2, ensure_ascii=False)
                    deal_json_paths.append(temp_json_path)
                    #print(f"    ✔ PASS  : {fname}  (cost ≈ ${cost:.4f})")

        total_cost += deal_cost

        if deal_passed:
            parent_folder = output_folder
            pass_count   += 1
            tag           = "✅ ALL EXTRACTION PASS"
        else:
            parent_folder = manual_output
            fail_count   += 1
            tag           = "❌ FAIL → Manual"

        dest_folder = os.path.join(parent_folder, deal_name)
        os.makedirs(dest_folder, exist_ok=True)

        print(f"\n    {tag} | Cost: ${deal_cost:.4f}")
        print("")
        #print(f"    [Folder] {dest_folder}")

        for temp_json_path in deal_json_paths:
            if os.path.isfile(temp_json_path):
                dest = os.path.join(dest_folder, os.path.basename(temp_json_path))
                shutil.move(temp_json_path, dest)
                #print(f"    [JSON] Saved: {os.path.basename(dest)}")

        json_files_in_dest = [
            os.path.join(dest_folder, f)
            for f in os.listdir(dest_folder)
            if f.endswith(".json") and deal_name in f
        ]

        total_check_failed = False

        for json_file in json_files_in_dest:
            csv_file = json_file.replace(".json", ".csv")
            try:
                csv_pass = run_json_to_csv_pipeline(json_file, csv_file)
            except Exception as e:
                print(f"    [WARN] JSON→CSV failed for "
                      f"{os.path.basename(json_file)}: {e}")
                total_check_failed = True
                continue

            if csv_pass is False:
                total_check_failed = True
                #print(f"    ✘ TOTAL-CHECK FAIL : {os.path.basename(json_file)} "
                      #f"— Total Check Status contains FAIL")

        if deal_passed and total_check_failed:
            #print(f"\n    ⚠ TOTAL CHECK FAIL detected — re-routing deal "
                  #f"from PASS → Manual Validation")
            new_dest = os.path.join(manual_output, deal_name)
            if os.path.exists(new_dest):
                shutil.rmtree(new_dest)
            shutil.move(dest_folder, new_dest)
            dest_folder = new_dest
            pass_count -= 1
            fail_count += 1
            deal_passed = False
            #print(f"    [Folder] {dest_folder}")

        try:
            merge_deal_csvs_to_excel(dest_folder, deal_name)
        except Exception as e:
            print(f"    [WARN] CSV→Excel merge failed for {deal_name}: {e}")

        raw_pdf_candidates = [
            os.path.join(raw_folder, f)
            for f in os.listdir(raw_folder)
            if f.lower().endswith(".pdf")
            and not _PRODUCED_RE.search(f)
            and get_base_pdf_name(Path(f).stem) == deal_name
        ]
        for raw_pdf in raw_pdf_candidates:
            dest = os.path.join(dest_folder, os.path.basename(raw_pdf))
            if not os.path.exists(dest):
                shutil.copy2(raw_pdf, dest)

      except Exception as e:
        print(f"\n    ❌ UNEXPECTED ERROR processing deal '{deal_name}' — "
              f"skipping to next deal. Reason: {type(e).__name__}: {e}")
        fail_count += 1
        continue

    print(f"\n{'='*70}")
    print(f"  ✅ BATCH pipeline completed!")
    #print(f"     TOTAL COST : ${total_cost:.4f}  (at batch-discounted rate)")
    print(f"     Deals PASS : {pass_count}  |  Deals FAIL: {fail_count}")
    #print(f"     Output     : {output_folder}")
    #print(f"     Manual     : {manual_output}")
    print(f"{'='*70}")


# ─────────────────────────────────────────────────────────────────────────────
# PARALLEL DEAL PROCESSING
# ─────────────────────────────────────────────────────────────────────────────

def process_deal_tables_parallel(
    deal_pdfs: list[str],
    prompts_folder: str,
    coa_text: str,
    model: str,
    max_tokens: int,
    provider: str,
    client,
    reporting_columns,
    sector: str = "LG",
    tier: str = "free",
    coa_mapping_folder: str = "",
) -> dict[str, dict]:
    tasks = []
    for pdf_path in deal_pdfs:
        stmt_type = detect_type(os.path.basename(pdf_path))
        if not stmt_type:
            print(f"[SKIP] Unknown statement type: {os.path.basename(pdf_path)}")
            continue

        prompt_filename = SUFFIX_TO_PROMPT.get(stmt_type)
        if not prompt_filename:
            print(f"[SKIP] No prompt mapping for {stmt_type}")
            continue

        prompt_path = os.path.join(prompts_folder, prompt_filename)
        if not os.path.exists(prompt_path):
            print(f"[SKIP] Prompt file missing: {prompt_path}")
            continue

        coa_filtered  = filter_coa_for_type(coa_text, stmt_type, sector)
        coa_map_rules = load_coa_mapping(coa_mapping_folder, sector, stmt_type)

        args_tuple = (
            pdf_path,
            prompt_path,
            coa_filtered,
            model,
            max_tokens,
            provider,
            client,
            reporting_columns,
            tier,
            stmt_type or "UNKNOWN",
            coa_map_rules,
        )
        tasks.append((pdf_path, args_tuple))

    if provider == "gemini":
        max_workers = 3
    elif provider == "ollama":
        max_workers = 1
    else:
        max_workers = 3

    results: dict[str, dict] = {}
    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        future_to_pdf = {
            executor.submit(process_one_pdf, args): pdf_path
            for pdf_path, args in tasks
        }
        for future in as_completed(future_to_pdf):
            pdf_path = future_to_pdf[future]
            try:
                results[pdf_path] = future.result()
            except Exception as e:
                print(f"[ERROR] Worker crashed on {os.path.basename(pdf_path)}: {e}")
                results[pdf_path] = {"pdf_path": pdf_path, "error": str(e)}

    return results


def build_page_note(page_info: dict) -> str:
    pages_str = ",".join(str(p) for p in page_info["pages"])

    if page_info["start"] == page_info["end"]:
        return (
            "SOURCE PAGE REFERENCE:\n"
            f"This extracted PDF corresponds to {page_info['label']} of the "
            "original full financial report (1-based page numbering), as "
            "determined from the source filename — NOT from any page "
            "number printed inside the PDF.\n"
            "Metadata.Page No MUST be set to exactly this string:\n"
            f"  \"{pages_str}\"\n"
            "Do NOT read, infer, or substitute any page number printed in "
            "the PDF body, footer, or header. Use the value above verbatim, "
            "regardless of what page number(s) appear in the document text."
        )

    return (
        "SOURCE PAGE REFERENCE:\n"
        f"This extracted PDF corresponds to {page_info['label']} of the "
        "original full financial report (1-based page numbering), as "
        "determined from the source filename — NOT from any page number(s) "
        "printed inside the PDF.\n"
        f"The content spans these pages, in order: {pages_str}.\n"
        "Metadata.Page No MUST be set to exactly this comma-joined string "
        "(no spaces), covering the FULL extracted range:\n"
        f"  \"{pages_str}\"\n"
        "Do NOT read, infer, or substitute any page number printed in the "
        "PDF body, footer, or header. Do NOT pick just one page from the "
        "range. Do NOT determine page numbers per data point/row — "
        "Metadata.Page No is a single value for the entire statement. Use "
        "the value above verbatim, regardless of what page number(s) "
        "appear in the document text."
    )


# ─────────────────────────────────────────────────────────────────────────────
# SYNC NORMALIZATION
# ─────────────────────────────────────────────────────────────────────────────

def run_normalization(
    extracted_pdfs: list[str],
    prompts_folder: str,
    xlsx_path: str,
    output_folder: str,
    manual_output: str,
    raw_folder: str,
    provider: str,
    model: str,
    max_tokens: int,
    reporting_columns,
    tier: str = "free",
    coa_mapping_folder: str = "",
) -> None:
    os.makedirs(output_folder, exist_ok=True)
    os.makedirs(manual_output, exist_ok=True)

    coa_text = load_xlsx_as_pipe_text(xlsx_path)
    print(f"\n[COA] Loaded {len(coa_text.splitlines())} rows from {xlsx_path}")

    client = None
    if provider == "openai":
        client = OpenAI()

    deal_groups: dict[str, list[str]] = {}
    for pdf_path in extracted_pdfs:
        stem = Path(pdf_path).stem
        base = get_base_pdf_name(stem)
        deal_groups.setdefault(base, []).append(pdf_path)

    total_deals = len(deal_groups)
    total_cost  = 0.0
    pass_count  = 0
    fail_count  = 0

    #print(f"\n{'='*70}")
    #print(f"  NORMALIZATION: {total_deals} deal(s) to process")

    for deal_idx, (deal_name, deal_pdfs) in enumerate(deal_groups.items(), 1):
      try:
        sector   = detect_sector(deal_name + ".pdf")
        expected = SECTOR_TABLE_SUFFIXES.get(sector, SECTOR_TABLE_SUFFIXES["LG"])  # ← RESTORE

        _next_log_pid(deal_name)                          # ← only this line needed

        print(f"\n{'─'*70}")
        print(f"  [{deal_idx}/{total_deals}] Deal : {deal_name} || Sector: {sector}")

        sector_allowed_pdfs = [
            p for p in deal_pdfs
            if detect_type(p) in expected        # ← now works
        ]
        sector_skipped = len(deal_pdfs) - len(sector_allowed_pdfs)
        if sector_skipped > 0:
            print(f"    [SECTOR FILTER] Removed {sector_skipped} PDF(s) "
                  f"not allowed for sector={sector}")

        if not sector_allowed_pdfs:
            print(f"    [SKIP] No valid tables for sector={sector} — skipping deal")
            continue

        deal_results = process_deal_tables_parallel(
            deal_pdfs          = sector_allowed_pdfs,
            prompts_folder     = prompts_folder,
            coa_text           = coa_text,
            model              = model,
            max_tokens         = max_tokens,
            provider           = provider,
            client             = client,
            reporting_columns  = reporting_columns,
            sector             = sector,
            tier               = tier,
            coa_mapping_folder = coa_mapping_folder,
        )

        deal_passed = True
        deal_cost   = 0.0

        for pdf_path, wrapped in deal_results.items():
            fname = os.path.basename(pdf_path)
            error = wrapped.get("error")

            if error:
                deal_passed = False
                print(f"    ✘ FAIL  : {fname}")
                print(f"              Reason: {error}")
                continue

            json_data = wrapped.get("data")

            if json_data:
                usage = wrapped.get("usage", {})
                pt    = usage.get("prompt_tokens", 0)
                ct    = usage.get("completion_tokens", 0)

                if provider == "openai":
                    cost = calc_openai_cost(model, pt, ct)
                elif wrapped.get("cost"):
                    cost = wrapped["cost"]
                else:
                    cost = 0.0

                cache_note = ""
                if wrapped.get("cached_hit"):
                    cache_note = "  [cache HIT]"

                deal_cost += cost

                base_stem          = Path(pdf_path).stem
                stmt_type_detected = detect_type(os.path.basename(pdf_path))

                prop_snp_split   = wrapped.get("prop_snp_split")
                split_stmt_types = {"_PROP_SNP", "_PROP_IS", "_PROP_CFS"}

                if stmt_type_detected in split_stmt_types and prop_snp_split:
                    saved_paths = []
                    for idx, sub_data in enumerate(prop_snp_split):
                        out_stem = prop_stmt_output_stem(
                            base_stem, sub_data.get("Metadata", {}), idx
                        )
                        sub_filename = f"{out_stem}.json"
                        sub_path = os.path.join(
                            os.path.dirname(pdf_path), sub_filename
                        )
                        with open(sub_path, "w", encoding="utf-8") as f:
                            json.dump(sub_data, f, indent=2, ensure_ascii=False)
                        saved_paths.append(sub_path)
                    wrapped["json_path"]       = saved_paths[0] if saved_paths else None
                    wrapped["json_path_extra"] = saved_paths[1:]
                else:
                    json_filename  = f"{base_stem}.json"
                    temp_json_path = os.path.join(os.path.dirname(pdf_path), json_filename)
                    with open(temp_json_path, "w", encoding="utf-8") as f:
                        json.dump(json_data, f, indent=2, ensure_ascii=False)
                    wrapped["json_path"]       = temp_json_path
                    wrapped["json_path_extra"] = []

            else:
                deal_passed = False
                print(f"    ✘ FAIL  : {fname}")
                print(f"              Reason: No data in response")

        total_cost += deal_cost

        if deal_passed:
            parent_folder = output_folder
            pass_count   += 1
            tag           = "✅ ALL PASS"
        else:
            parent_folder = manual_output
            fail_count   += 1
            tag           = "❌ FAIL → Manual"

        dest_folder = os.path.join(parent_folder, deal_name)
        os.makedirs(dest_folder, exist_ok=True)

        print(f"\n    {tag} ")

        for pdf_path, wrapped in deal_results.items():
            json_path = wrapped.get("json_path")
            if json_path and os.path.isfile(json_path):
                dest = os.path.join(dest_folder, os.path.basename(json_path))
                shutil.move(json_path, dest)

            for extra_path in wrapped.get("json_path_extra", []):
                if extra_path and os.path.isfile(extra_path):
                    dest = os.path.join(dest_folder, os.path.basename(extra_path))
                    shutil.move(extra_path, dest)
                    print(f"    [JSON] Saved (split): {os.path.basename(dest)}")

        json_files = [
            os.path.join(dest_folder, f)
            for f in os.listdir(dest_folder)
            if f.endswith(".json") and deal_name in f
        ]

        total_check_failed = False

        for json_file in json_files:
            csv_file = json_file.replace(".json", ".csv")
            csv_pass = None
            try:
                csv_pass = run_json_to_csv_pipeline(json_file, csv_file)
            except TypeError:
                try:
                    csv_pass = run_json_to_csv_pipeline(json_file)
                except TypeError:
                    try:
                        csv_pass = run_json_to_csv_pipeline(dest_folder, deal_name)
                    except Exception as e:
                        print(f"    [WARN] JSON→CSV failed for "
                              f"{os.path.basename(json_file)}: {e}")
                        total_check_failed = True
            except Exception as e:
                print(f"    [WARN] JSON→CSV failed for "
                      f"{os.path.basename(json_file)}: {e}")
                total_check_failed = True

            if csv_pass is False:
                total_check_failed = True

        if deal_passed and total_check_failed:
            print(f"\n    ⚠ TOTAL CHECK FAIL detected — re-routing deal "
                  f"from PASS → Manual Validation")
            new_dest = os.path.join(manual_output, deal_name)
            if os.path.exists(new_dest):
                shutil.rmtree(new_dest)
            shutil.move(dest_folder, new_dest)
            dest_folder = new_dest
            pass_count -= 1
            fail_count += 1
            deal_passed = False

        try:
            merge_deal_csvs_to_excel(dest_folder, deal_name)
        except TypeError:
            try:
                merge_deal_csvs_to_excel(folder=dest_folder, deal_name=deal_name)
            except Exception as e:
                print(f"    [WARN] CSV merge failed for {deal_name}: {e}")
        except Exception as e:
            print(f"    [WARN] CSV merge failed for {deal_name}: {e}")

        raw_pdf_candidates = [
            os.path.join(raw_folder, f)
            for f in os.listdir(raw_folder)
            if f.lower().endswith(".pdf")
            and not _PRODUCED_RE.search(f)
            and get_base_pdf_name(Path(f).stem) == deal_name
        ]
        for raw_pdf in raw_pdf_candidates:
            dest = os.path.join(dest_folder, os.path.basename(raw_pdf))
            if not os.path.exists(dest):
                shutil.copy2(raw_pdf, dest)

      except Exception as e:
        print(f"\n    ❌ UNEXPECTED ERROR processing deal '{deal_name}' — "
              f"skipping to next deal. Reason: {type(e).__name__}: {e}")
        fail_count += 1
        continue

    # ── Provider-specific cleanup ─────────────────────────────────────────────
    if tier == "paid":
        if provider == "gemini":
            try:
                from gemini_cache_client import cleanup_caches
                cleanup_caches()
            except Exception as e:
                print(f"[WARN] Gemini cache cleanup skipped: {e}")
        elif provider == "openai":
            try:
                from openai_cache_client import cleanup_caches
                cleanup_caches()
            except Exception as e:
                print(f"[WARN] OpenAI cache cleanup skipped: {e}")
        elif provider == "claude":
            try:
                from claude_cache_client import cleanup_caches
                cleanup_caches()
            except Exception as e:
                print(f"[WARN] Claude cache cleanup skipped: {e}")

    print(f"\n{'='*70}")
    #print(f"     TOTAL COST: ${total_cost:.4f}")
    print(f"     Deals PASS: {pass_count}  |  Deals FAIL: {fail_count}")

# ─────────────────────────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────────────────────────

def main() -> None:
    args = parse_args()

    if args.batch and args.tier != "paid":
        print("[ERROR] --batch requires --tier paid. Re-run with --tier paid --batch.")
        sys.exit(1)

    if args.provider == "claude":
        if args.tier != "paid":
            print("[INFO] Claude provider selected — forcing --tier paid because Claude has no free tier.")
        args.tier = "paid"

    if args.batch and args.provider == "ollama":
        print("[ERROR] --batch is not supported for provider=ollama.")
        sys.exit(1)

    model = resolve_model(args.provider, args.model)

    # ── Start DuckDB log writer FIRST — before any other prints ──────────────
    # This ensures the run header and ALL subsequent output is captured.
    global _pipeline_log_writer
    try:
        from log_writer import LogWriter
        _pipeline_log_writer = LogWriter()
    except ImportError:
        #print("[WARN] log_writer.py not found — DuckDB logging disabled.")
        _pipeline_log_writer = None
        _log_processing_id   = 0
    # ── Print run header (mirrors UI display, now captured in parquet) ────────
    mode_label = (
        "⚡ BATCH (async, ≤24 h, 50% cost)"
        if (args.tier == "paid" and getattr(args, "batch", False))
        else f"🔄 SYNC (tier={args.tier})"
    )


    # ── Now announce the log file (LogWriter already open above) ─────────────
    # if _pipeline_log_writer is not None:
    #     print(f"[LOG] Logging to: logs/pipeline_logs.parquet")
    # ─────────────────────────────────────────────────────────────────────────

    if args.skip_extraction:
        extracted_pdfs = collect_existing_extracted(
            args.folder, prompts_folder=args.prompts
        )
    else:
        existing = collect_existing_extracted(args.folder, prompts_folder=args.prompts)

        # Find raw PDFs that have NO extracted outputs yet
        all_raw = [
            f for f in os.listdir(args.folder)
            if f.lower().endswith(".pdf") and not _PRODUCED_RE.search(f)
        ]
        already_processed_bases = {
            get_base_pdf_name(Path(p).stem) for p in existing
        }
        unprocessed_raw = [
            os.path.join(args.folder, f) for f in all_raw
            if get_base_pdf_name(Path(f).stem) not in already_processed_bases
        ]

        if unprocessed_raw:
            newly_extracted = run_extraction(
                args.folder,
                prompts_folder = args.prompts,
                use_llm_id     = getattr(args, "llm_page_id", False),
                id_model       = getattr(args, "id_model", "claude-sonnet-4-6"),
            )
            extracted_pdfs = existing + newly_extracted
        else:
            print(f"[INFO] All raw PDFs already extracted — using existing sliced PDFs.")
            extracted_pdfs = existing

    if not extracted_pdfs:
        print("[INFO] No extracted PDFs to normalize. Exiting.")
        return

    if getattr(args, "skip_normalization", False):
        print(f"\n[INFO] --skip-normalization set — extraction complete.")
        return

    # ── Banner now appears AFTER extraction, BEFORE normalization ────────
    print(f"\n{'='*70}")                                                  # ← KEEP
    print(f"DATA Extraction IS RUNNING....")                              # ← KEEP (static, no file label needed here)
    print(f"{'='*70}")                                                    # ← ADD

    if args.tier == "paid" and args.batch:
        print(f"[INFO] Mode: PAID + BATCH  → async Batch API  (provider={args.provider})")
        run_normalization_batch(
            extracted_pdfs     = extracted_pdfs,
            prompts_folder     = args.prompts,
            xlsx_path          = args.xlsx,
            output_folder      = args.output,
            manual_output      = args.manual_output,
            raw_folder         = args.folder,
            provider           = args.provider,
            model              = model,
            max_tokens         = args.max_tokens,
            reporting_columns  = args.reporting_columns,
            coa_mapping_folder = args.coa_mapping,
        )
    else:
        tier_label = args.tier
        run_normalization(
            extracted_pdfs     = extracted_pdfs,
            prompts_folder     = args.prompts,
            xlsx_path          = args.xlsx,
            output_folder      = args.output,
            manual_output      = args.manual_output,
            raw_folder         = args.folder,
            provider           = args.provider,
            model              = model,
            max_tokens         = args.max_tokens,
            reporting_columns  = args.reporting_columns,
            tier               = tier_label,
            coa_mapping_folder = args.coa_mapping,
        )
    cleanup_extracted_pdfs(args.folder)

    # ── Provider-specific cache cleanup ───────────────────────────────────────
    if args.tier == "paid" and not args.batch:
        if args.provider == "gemini":
            try:
                from gemini_cache_client import cleanup_caches
                cleanup_caches()
            except Exception as e:
                print(f"[WARN] Gemini cache cleanup skipped: {e}")
        elif args.provider == "openai":
            try:
                from openai_cache_client import cleanup_caches
                cleanup_caches()
            except Exception as e:
                print(f"[WARN] OpenAI cache cleanup skipped: {e}")
        elif args.provider == "claude":
            try:
                from claude_cache_client import cleanup_caches
                cleanup_caches()
            except Exception as e:
                print(f"[WARN] Claude cache cleanup skipped: {e}")

    # ── Close log writer → flush + export parquet ─────────────────────────────
    if _pipeline_log_writer is not None:
        try:
            _pipeline_log_writer.close()
            print(f"[LOG] Parquet saved → logs/pipeline_logs.parquet")
        except Exception as e:
            print(f"[WARN] Log writer close failed: {e}")
    # ─────────────────────────────────────────────────────────────────────────


if __name__ == "__main__":
    main()