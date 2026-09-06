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

# ── Self-bootstrap this engine folder onto sys.path ───────────────────────────
# Every engine module imports its siblings by BARE name (from pdf_to_indented_text
# import ..., from coordinate_extractor import ..., import page_extractor). Those
# resolve only when THIS folder is on sys.path. Ensure it — so the engine runs no
# matter how it was launched (manual, Dagster op, subprocess, package import,
# differing cwd). MUST stay above the sibling imports below, and it is re-asserted
# at the start of the runtime entry points in case sys.path changes after import.
def _ensure_engine_on_path() -> None:
    _d = os.path.dirname(os.path.abspath(__file__))
    if _d not in sys.path:
        sys.path.insert(0, _d)

_ensure_engine_on_path()

from pdf_to_indented_text import (
    extract_all_pages_words,
    pdf_to_indented_text_from_words,
)
from coordinate_extractor import attach_coordinates_from_words
import functools, builtins
import hashlib as _hashlib

# ── Intercept print() → also write to DuckDB log ─────────────────────────────
_pipeline_log_writer = None   # set in main() (CLI) or via start_pipeline_logging() (DB workflow)
_log_processing_id   = 0
_log_pid_override    = None   # when set, wins over the per-file filename (see _next_log_pid)
_original_print = builtins.print
def _enable_windows_long_paths() -> None:
    """
    Enable long path support (> 260 chars) for this process on Windows.
    Requires either:
      - Windows 10 version 1607+ with LongPathsEnabled registry key = 1, OR
      - Python 3.6+ (which calls SetFileInformationByHandle internally)
    This call is a no-op on Linux/macOS.
    """
    if os.name != "nt":
        return
    try:
        import ctypes
        # FILE_ATTRIBUTE_NORMAL = 0x80; kernel32.SetFileShortNameW is not
        # what we want — we need SetConsoleCP or the manifest approach.
        # The reliable programmatic way on Python is to set the
        # PYTHONLEGACYWINDOWSSTDIO env and call the Win32 API directly.
        kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
        # SetFileInformationByHandle is already called by Python 3.6+.
        # What we need is to toggle the process-level long-path flag via
        # the registry — but that requires admin rights.
        # SAFEST option: just prepend \\?\ to all absolute paths.
        print("[INFO] Windows detected — long-path mode active (\\\\?\\ prefix).")
    except Exception:
        pass
def _next_log_pid(filename: str = "") -> None:
    """Set the current ProcessingId on the log writer.

    Uses the caller-locked override (e.g. the DB TProcessStatus.ProcessingId set via
    set_log_context) when present; otherwise falls back to the per-file filename stem
    (the standalone-CLI behavior — unchanged when no override is set).
    """
    # FIX (Bug 2): removed the dead `return _log_processing_id` — the function
    # signature is None and _log_processing_id was never incremented, so the
    # return value was always 0.  Callers don't use the return value anyway.
    if _pipeline_log_writer is not None:
        _pipeline_log_writer.set_context(
            processing_id = _log_pid_override if _log_pid_override else filename,
            stage         = "p",
        )

def _logging_print(*args, **kwargs):
    kwargs.setdefault("flush", True)
    _original_print(*args, **kwargs)
    # FIX (Concern 1): wrap log writer call in try/except so a DuckDB lock
    # or any other log_writer error never silently crashes the print intercept.
    if _pipeline_log_writer is not None:
        line = " ".join(str(a) for a in args)
        for sub in line.split("\n"):
            sub = sub.strip()
            if sub:
                try:
                    _pipeline_log_writer.write(sub)
                except Exception:
                    pass  # Never let logging break the pipeline

builtins.print = _logging_print


# ── Public logging API (used by the DB workflow, PFG_Extraction.py) ──────────
# The standalone CLI (main() below) manages the writer inline; these let an
# external caller drive the SAME writer without duplicating that logic.
def start_pipeline_logging(log_dir=None):
    """Attach a LogWriter so intercepted print() output is captured to the
    processing-log parquet. `log_dir` lets a caller (PFG_Extraction) inject its
    centralized log location; None → LogWriter's standalone default. Idempotent;
    returns the active writer (or None if log_writer is unavailable).

    `log_writer` is imported LAZILY here, and this is the FIRST pipeline function
    the DB workflow calls — before run_extraction/process_one_deal have had a
    chance to re-assert sys.path. Under Dagster that made the import the earliest
    casualty of a clobbered sys.path, and because the failure was swallowed
    silently the whole run produced NO processing log with nothing to say why.
    Hence the path assertion, and the warning below."""
    global _pipeline_log_writer
    _ensure_engine_on_path()
    if _pipeline_log_writer is None:
        try:
            from log_writer import LogWriter
            _pipeline_log_writer = LogWriter(log_dir)
        except Exception as exc:
            # Never fatal — extraction must still run without its log. But say so:
            # a missing processing log used to be indistinguishable from a quiet run.
            _original_print(
                f"[WARN] processing log DISABLED — could not start log_writer: "
                f"{type(exc).__name__}: {exc}", flush=True)
            _pipeline_log_writer = None
    return _pipeline_log_writer


def stop_pipeline_logging():
    """Flush and detach the LogWriter (safe to call when none is attached)."""
    global _pipeline_log_writer
    if _pipeline_log_writer is not None:
        try:
            _pipeline_log_writer.close()   # flushes the buffer
        finally:
            _pipeline_log_writer = None


def set_log_context(pid) -> None:
    """Lock the log ProcessingId to `pid` (e.g. the DB TProcessStatus.ProcessingId),
    overriding the per-file filename that run_extraction sets internally. Persists
    until clear_log_context()."""
    global _log_pid_override
    _log_pid_override = None if pid is None else str(pid)
    if _pipeline_log_writer is not None and _log_pid_override:
        _pipeline_log_writer.set_context(_log_pid_override, "p")


def clear_log_context() -> None:
    """Release the ProcessingId lock; _next_log_pid falls back to the filename."""
    global _log_pid_override
    _log_pid_override = None


from concurrent.futures import ThreadPoolExecutor, as_completed
import threading
try:
    from coordinate_extractor import attach_coordinates
    _COORD_AVAILABLE = True
except ImportError:
    _COORD_AVAILABLE = False
    print("[WARN] coordinate_extractor.py not found — coordinates will be skipped.")
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
    from compact_schema import STATEMENT_SUFFIX_ALTERNATION
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
    # ── Notes / RSI / Statistical-Section tabs ──
    "_OVERVIEW":       "OVERVIEW_Prompt.txt",
    "_CAPITAL_ASSETS": "CAPITAL_ASSETS_Prompt.txt",
    "_TAX_BASE":       "TAX_BASE_Prompt.txt",
    "_PEN":            "PEN_Prompt.txt",
    "_OPEB":           "OPEB_Prompt.txt",
    "_FAQS":           "FAQS_Prompt.txt",
}


# ORDER IS SIGNIFICANT — detect_type() returns the FIRST pattern that matches,
# so any suffix that CONTAINS a shorter suffix must be listed before it
# (PROP_SNP before SNP, CAPITAL_ASSETS before any substring, etc.).
TYPE_PATTERNS: list[tuple[re.Pattern, str]] = [
    (re.compile(r"_CAPITAL_ASSETS", re.IGNORECASE), "_CAPITAL_ASSETS"),
    (re.compile(r"_PROP_SNP", re.IGNORECASE), "_PROP_SNP"),
    (re.compile(r"_PROP_CFS", re.IGNORECASE), "_PROP_CFS"),
    (re.compile(r"_PROP_IS",  re.IGNORECASE), "_PROP_IS"),
    (re.compile(r"_GOV_BS",   re.IGNORECASE), "_GOV_BS"),
    (re.compile(r"_GOV_IS",   re.IGNORECASE), "_GOV_IS"),
    (re.compile(r"_TAX_BASE", re.IGNORECASE), "_TAX_BASE"),
    (re.compile(r"_OVERVIEW", re.IGNORECASE), "_OVERVIEW"),
    (re.compile(r"_OPEB",     re.IGNORECASE), "_OPEB"),
    (re.compile(r"_FAQS",     re.IGNORECASE), "_FAQS"),
    (re.compile(r"_PEN",      re.IGNORECASE), "_PEN"),
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

# FIX (Concern 3): removed save_prop_snp_results() and prop_snp_output_stem()
# — they were exact duplicates of save_multi_table_results() and
# prop_stmt_output_stem() respectively, and were never called anywhere in the
# normalization paths. The generic versions handle all three split types.

_PAGE_TAG_RE = re.compile(
    rf"_(?:{STATEMENT_SUFFIX_ALTERNATION})"
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
    p.add_argument(
    "--xlsx",
    required=False,
    default=None,
    help="Financial: path to COA master XLSX (e.g. ./Master/standard_coa_master.xlsx). "
         "ESG: path to Master folder containing guide XLSXs (e.g. ./Master).",
)
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

    # ═══════════════════════════════════════════════════════════════
    # ★ ESG PATCH — NEW ARGUMENTS
    # ═══════════════════════════════════════════════════════════════
    p.add_argument(
        "--esg-mode",
        choices=["none", "indian", "global", "auto"],
        default="none",
        help="Run ESG extraction. 'auto' detects Indian vs Global from filenames.",
    )
    # ═══════════════════════════════════════════════════════════════
    args = p.parse_args()
    if not args.xlsx:
        if args.esg_mode == "none":
            p.error("--xlsx is required for the financial pipeline "
                    "(e.g. --xlsx ./Master/standard_coa_master.xlsx)")
        else:
            p.error("--xlsx is required for the ESG pipeline "
                    "(e.g. --xlsx ./Master)")

    return args


# ─────────────────────────────────────────────────────────────────────────────
# MODEL DEFAULTING
# ─────────────────────────────────────────────────────────────────────────────

DEFAULT_MODEL_BY_PROVIDER: dict[str, str] = {
    "openai": "gpt-4o-mini",
    "gemini": "gemini-3.6-flash",
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
        # ── Notes / RSI / Statistical-Section tabs ──
        "_OVERVIEW",
        "_CAPITAL_ASSETS",
        "_TAX_BASE",
        "_PEN",
        "_OPEB",
        "_FAQS",
    ],
    # NON-LG deliberately stays at the three proprietary statements. The notes
    # tabs are NOT added here: TAX_BASE has no analogue outside a taxing
    # authority, and the remaining five need their own NON-LG prompts and COA
    # vocabulary before they can be enabled. Adding a suffix here without those
    # files would send the LG prompt at a NON-LG document.
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
    # Check for sidecar file written when the page list was too long for the filename.
    sidecar = str(pdf_path)[:-4] + ".pages"
    if os.path.isfile(sidecar):
        import json as _json_sidecar
        try:
            with open(sidecar, "r", encoding="utf-8") as _fh:
                pages = _json_sidecar.load(_fh)
            pages = sorted(int(p) for p in pages)
            start, end = min(pages), max(pages)
            label = f"page {start}" if start == end else f"pages {start}–{end}"
            return {"start": start, "end": end, "pages": pages, "label": label}
        except Exception:
            pass  # fall through to filename parsing

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
    # Runtime entry point — re-assert sys.path in case it changed after import
    # (a Dagster op can lose it, breaking the lazy sibling imports downstream).
    _ensure_engine_on_path()
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
    total_files = len(raw_pdfs)

    for file_idx, pdf_path in enumerate(raw_pdfs, start=1):
        fname  = os.path.basename(pdf_path)
        _next_log_pid(Path(fname).stem)
        sector = detect_sector(fname)
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


def run_extraction_batch(
    folder: str,
    prompts_folder: str = None,
    id_model: str = "claude-sonnet-4-6",
) -> list[str]:
    """
    Batch-API variant of run_extraction for Claude ID models.
    Submits ALL page-identification LLM calls as one Claude batch job,
    waits for results, then slices PDFs locally.

    Falls back to sync run_extraction if the id_model is not a Claude model.
    NOTE for the DB workflow: with a Gemini/OpenAI `id_model` this ALWAYS delegates
    to run_extraction, so page identification stays synchronous; only normalization
    is batched. Batch page ID engages only for a Claude id_model.
    """
    # Runtime entry point — see run_extraction.
    _ensure_engine_on_path()
    provider = id_model.lower().split("-")[0] if id_model else "gemini"

    if not id_model.lower().startswith("claude"):
        print(
            f"[ID-BATCH] Batch page ID is only supported for Claude models. "
            f"Falling back to sync for id_model={id_model!r}."
        )
        return run_extraction(
            folder=folder,
            prompts_folder=prompts_folder,
            use_llm_id=True,
            id_model=id_model,
        )

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
            if os.path.isfile(os.path.join(prompts_folder, prompt_filename)):
                available_prompts.add(suffix)
    else:
        available_prompts = set(SUFFIX_TO_PROMPT.keys())

    # ── Phase 1: Prepare batch jobs for all PDFs ──────────────────────
    from llm_page_identifier import prepare_page_id_claude_jobs
    from page_extractor import process_file_llm_guided

    all_jobs: list[dict] = []
    pdf_meta: dict[str, dict] = {}  # pdf_path → {sector, allowed_suffixes}

    print(f"\n[ID-BATCH] Preparing page-ID batch jobs for {len(raw_pdfs)} PDF(s)…")
    for pdf_path in raw_pdfs:
        fname   = os.path.basename(pdf_path)
        sector  = detect_sector(fname)
        sector_allowed  = set(SECTOR_TABLE_SUFFIXES.get(sector, SECTOR_TABLE_SUFFIXES["LG"]))
        effective_allowed = sector_allowed & available_prompts

        if not effective_allowed:
            print(f"  [SKIP] {fname} — no allowed tables")
            continue

        pdf_meta[pdf_path] = {"sector": sector, "allowed_suffixes": effective_allowed}
        jobs = prepare_page_id_claude_jobs(
            pdf_path=pdf_path,
            id_model=id_model,
            allowed_suffixes=effective_allowed,
            prompt_dir=prompts_folder,
        )
        all_jobs.extend(jobs)
        print(f"  {fname}: {len(jobs)} job(s) queued")

    if not all_jobs:
        print("[ID-BATCH] No jobs to submit.")
        return []

    print(f"[ID-BATCH] Submitting {len(all_jobs)} request(s) to Claude Batch API…")

    # ── Phase 2: Submit batch ─────────────────────────────────────────
    from claude_batch_client import _safe_custom_id
    import anthropic

    api_key = os.environ.get("ANTHROPIC_API_KEY", "")
    if not api_key:
        try:
            import winreg
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, "Environment") as _key:
                api_key, _ = winreg.QueryValueEx(_key, "ANTHROPIC_API_KEY")
                os.environ["ANTHROPIC_API_KEY"] = api_key
        except Exception:
            pass

    client = anthropic.Anthropic(api_key=api_key)

    id_map: dict[str, str] = {}  # safe_id → original custom_id
    requests = []
    seen: set[str] = set()
    for j in all_jobs:
        raw_id  = j["custom_id"]
        safe_id = _safe_custom_id(raw_id)
        if safe_id in seen:
            continue
        seen.add(safe_id)
        id_map[safe_id] = raw_id
        requests.append({"custom_id": safe_id, "params": j["params"]})

    batch     = client.messages.batches.create(requests=requests)
    batch_id  = getattr(batch, "id", None)
    if not batch_id:
        raise RuntimeError(f"Claude batch created but no id returned: {batch}")

    print(f"[ID-BATCH] Batch submitted: batch_id={batch_id}")
    print(f"[ID-BATCH] Polling every 60 s…")

    # ── Phase 3: Poll until complete ─────────────────────────────────
    import time as _time
    while True:
        b = client.messages.batches.retrieve(batch_id)
        status = (
            getattr(b, "processing_status", None)
            or getattr(b, "status", None)
            or ""
        )
        counts = getattr(b, "request_counts", None)
        print(f"  [ID-BATCH] status={status}  counts={counts}")
        if status in ("ended", "complete", "completed"):
            break
        _time.sleep(60)

    # ── Phase 4: Collect results ──────────────────────────────────────
    raw_results: dict[str, str] = {}  # original custom_id → raw LLM text
    for result_item in client.messages.batches.results(batch_id):
        safe_id = result_item.custom_id
        orig_id = id_map.get(safe_id, safe_id)
        res     = result_item.result
        if getattr(res, "type", None) == "succeeded":
            msg = res.message
            raw = "".join(
                block.text for block in getattr(msg, "content", [])
                if hasattr(block, "text")
            )
            raw_results[orig_id] = raw
        else:
            print(f"  [ID-BATCH] FAILED  custom_id={orig_id}: {getattr(res, 'error', res)}")

    # ── Phase 5: Post-process + slice PDFs ───────────────────────────
    from llm_page_identifier import process_page_id_batch_results

    all_extracted: list[str] = []
    total_files = len(pdf_meta)

    for file_idx, (pdf_path, meta) in enumerate(pdf_meta.items(), start=1):
        fname             = os.path.basename(pdf_path)
        sector            = meta["sector"]
        allowed_suffixes  = meta["allowed_suffixes"]
        stem              = Path(pdf_path).stem

        _next_log_pid(stem)
        print(f"\n[{file_idx}/{total_files}] {fname}")

        raw_by_call = {}
        for call_type in ("dsr_debt", "main", "notes"):
            cid = f"{stem}__{call_type}"
            if cid in raw_results:
                raw_by_call[call_type] = raw_results[cid]

        if not raw_by_call:
            print(f"  [WARN] No batch results for {fname} — skipping")
            continue

        try:
            page_map = process_page_id_batch_results(
                pdf_path        = pdf_path,
                allowed_suffixes= allowed_suffixes,
                raw_by_call_type= raw_by_call,
            )
        except Exception as e:
            print(f"  [ERROR] Post-processing failed for {fname}: {e}")
            continue

        try:
            produced = process_file_llm_guided(
                src              = pdf_path,
                sector           = sector,
                allowed_suffixes = allowed_suffixes,
                id_model         = id_model,
                use_llm_id       = True,
                file_index       = file_idx,
                file_total       = total_files,
                _precomputed_page_map = page_map,
            )
        except TypeError:
            from page_extractor import process_file_llm_guided as _pfg
            import inspect
            sig = inspect.signature(_pfg)
            if "_precomputed_page_map" in sig.parameters:
                produced = _pfg(
                    src=pdf_path, sector=sector,
                    allowed_suffixes=allowed_suffixes,
                    id_model=id_model, use_llm_id=True,
                    file_index=file_idx, file_total=total_files,
                    _precomputed_page_map=page_map,
                )
            else:
                produced = _pfg(
                    src=pdf_path, sector=sector,
                    allowed_suffixes=allowed_suffixes,
                    id_model=id_model, use_llm_id=True,
                    file_index=file_idx, file_total=total_files,
                )

        if not produced:
            print(f"  [WARN] No extracted PDFs produced for {fname}")
            continue

        filtered = []
        for p in produced:
            matched_type = detect_type(p)
            if matched_type is None or matched_type not in allowed_suffixes:
                try:
                    os.remove(p)
                except OSError:
                    pass
                continue
            filtered.append(p)

        all_extracted.extend(filtered)
        print(f"  → kept {len(filtered)} table PDF(s)")

    return all_extracted


def collect_existing_extracted(folder: str, prompts_folder: str = None) -> list[str]:
    folder_path = Path(folder)
    found: list[str] = []

    _EXTRACTED_RE = re.compile(
        rf"_(?:{STATEMENT_SUFFIX_ALTERNATION})"
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
    indented_text: str | None = None,       # ← NEW: pre-extracted, no second pdfplumber open
) -> dict:
    system_prompt = Path(prompt_path).read_text(encoding="utf-8", errors="replace")
    system_prompt = system_prompt + "\n\n" + build_short_key_instruction(stmt_type)
    if coa_map_rules:
        system_prompt = system_prompt + "\n\n" + coa_map_rules
    pdf_bytes    = Path(pdf_path).read_bytes()
    b64          = base64.b64encode(pdf_bytes).decode("utf-8")
    pdf_data_uri = f"data:application/pdf;base64,{b64}"
    pdf_name     = Path(pdf_path).name

    if page_info is None:
        page_info = extract_page_info_from_filename(pdf_path)

    # Use pre-extracted indented_text if provided; only open pdfplumber as fallback
    if indented_text is None:
        try:
            from pdf_to_indented_text import pdf_to_indented_text as _pit
            indented_text = _pit(pdf_path)
        except Exception:
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
        page_note = build_page_note(page_info, stmt_type)
        user_content.append({"type": "text", "text": page_note})

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
        {"role": "user",   "content": user_content},
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
                parsed   = parse_json_response(part)
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

    parsed_data   = parse_json_response(raw)
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
    indented_text: str | None = None,       # ← NEW
) -> dict:
    try:
        from gemini_client import normalize_one_gemini
    except ImportError as e:
        raise RuntimeError(f"gemini_client.py not found: {e}")

    if page_info is None:
        page_info = extract_page_info_from_filename(pdf_path)

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
            indented_text=indented_text,    # ← NEW: passed through to gemini_client
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
    indented_text: str | None = None,       # ← NEW
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
        indented_text=indented_text,        # ← NEW: passed through to claude_cache_client
    )
    result.setdefault("prop_snp_split", [])
    return result


def get_base_pdf_name(stem: str) -> str:
    result = re.sub(r"_p\d+(?:-\d+)*$", "", stem)
    result = re.sub(
        rf"_(?:{STATEMENT_SUFFIX_ALTERNATION})$",
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
    # DSR and DEBT intentionally have no COA keyword — these statement types
    # do not require COA master mapping, so filter_coa_for_type() returns ""
    # for them and the LLM prompt receives an empty COA section.
    "_DSR":      "",
    "_DEBT":     "",
    # The six notes/RSI/statistical tabs likewise have NO COA-master keyword:
    #   OVERVIEW / PEN / OPEB / FAQs  → COA Flag and COA Datapoint are "n/a";
    #                                   there is no mapping to perform.
    #   TAX_BASE                      → its 7 COA Datapoints are fixed constants
    #                                   baked into TAX_BASE_Prompt.txt.
    #   CAPITAL_ASSETS                → maps to a CLOSED 4-value vocabulary that
    #                                   is not in standard_coa_master.xlsx; the
    #                                   rules live in CAPITAL_ASSETS_COA.txt and
    #                                   are loaded by load_coa_mapping().
    "_OVERVIEW":       "",
    "_CAPITAL_ASSETS": "",
    "_TAX_BASE":       "",
    "_PEN":            "",
    "_OPEB":           "",
    "_FAQS":           "",
}

# Statement types that legitimately have NO <TYPE>_COA.txt mapping file, so
# load_coa_mapping() must stay silent instead of warning on every run.
_NO_COA_MAPPING_FILE: set[str] = {
    "DEBT", "DSR", "OVERVIEW", "TAX_BASE", "PEN", "OPEB", "FAQS",
}

def load_coa_mapping(coa_mapping_folder: str, sector: str, stmt_type: str, silent: bool = False) -> str:
    if not coa_mapping_folder:
        return ""
    path = os.path.join(coa_mapping_folder, sector,
                        f"{stmt_type.lstrip('_')}_COA.txt")
    if not os.path.isfile(path):
        if stmt_type.lstrip("_") not in _NO_COA_MAPPING_FILE and not silent:
            print(f"  [COA-MAP] No mapping file for sector={sector} "
                  f"stmt_type={stmt_type} — skipping (path: {path})")
        return ""
    text = Path(path).read_text(encoding="utf-8", errors="replace")
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

    OPTIMISATION: pdfplumber is opened ONCE via extract_all_pages_words().
    The same word list is passed to both:
      • pdf_to_indented_text_from_words()  — builds indented text for LLM
      • attach_coordinates_from_words()    — injects _coord keys post-LLM
    This eliminates the duplicate pdfplumber open that previously occurred.
    """
    (pdf_path, prompt_path, coa_text, model, max_tokens,
     provider, client, reporting_columns, tier, stmt_type,
     coa_map_rules) = args_tuple

    pdf_name  = os.path.basename(pdf_path)
    page_info = extract_page_info_from_filename(pdf_path)

    # ── SINGLE pdfplumber open — extract words once for entire worker ─────
    try:
        from pdf_to_indented_text import (
            extract_all_pages_words,
            pdf_to_indented_text_from_words,
        )
        all_pages_words = extract_all_pages_words(pdf_path)
        indented_text   = pdf_to_indented_text_from_words(all_pages_words)
    except Exception as _e:
        print(f"   [WARN] pdfplumber extraction failed for {pdf_name}: {_e}")
        all_pages_words = None
        # Fallback: try the original single-call version
        try:
            from pdf_to_indented_text import pdf_to_indented_text
            indented_text = pdf_to_indented_text(pdf_path)
        except Exception:
            indented_text = None

    try:
        # ── PATH 1 — OPENAI PAID (context caching) ────────────────────────
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
                indented_text=indented_text,
            )

        # ── PATH 2 — OPENAI FREE ──────────────────────────────────────────
        elif provider == "openai":
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
                indented_text=indented_text,
            )

        # ── PATH 3 — GEMINI PAID (cached, with fallback) ──────────────────
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
                    indented_text=indented_text,
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
                        indented_text=indented_text,
                    )
                except CacheUnavailableError as e:
                    print(f"   [PAID→FREE FALLBACK] Caching unavailable "
                          f"({e}). Using plain Gemini call.")
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
                        indented_text=indented_text,
                    )
                except Exception as e:
                    print(f"   [PAID→FREE FALLBACK] Gemini cache path failed "
                          f"unexpectedly ({type(e).__name__}: {e}). "
                          f"Using plain Gemini call.")
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
                        indented_text=indented_text,
                    )

        # ── PATH 4 — GEMINI FREE ──────────────────────────────────────────
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
                indented_text=indented_text,
            )

        # ── PATH 5 — CLAUDE PAID (prompt caching) ─────────────────────────
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
                    indented_text=indented_text,
                )

        else:
            raise ValueError(f"Unknown provider: {provider}")

        # ── Post-LLM: attach coordinates using the SAME words ─────────────
        # No second pdfplumber open needed — reuse all_pages_words from above.
        json_data = result.get("data")
        if json_data and _COORD_AVAILABLE:
            try:
                if all_pages_words is not None:
                    from coordinate_extractor import attach_coordinates_from_words
                    json_data = attach_coordinates_from_words(
                        json_data, all_pages_words
                    )
                else:
                    # Fallback to original path if word extraction failed
                    from coordinate_extractor import attach_coordinates
                    json_data = attach_coordinates(json_data, pdf_path)
                result["data"] = json_data
            except Exception as coord_err:
                print(f"   [WARN] Coordinate attachment failed for "
                      f"{pdf_name}: {coord_err}")

        # ── Column-shift repair (Taxes/Inventories identity-trap fix) ──────
        if json_data:
            try:
                from column_shift_repair import repair_column_shifts
                json_data = repair_column_shifts(json_data, stmt_type)
                result["data"] = json_data
            except Exception as repair_err:
                print(f"   [WARN] Column shift repair failed for "
                      f"{pdf_name}: {repair_err}")

        # ── Shared return block (ALL providers reach here) ─────────────────
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
    """
    Anthropic enforces a 64-character hard limit on custom_id.
    Includes a hash of the full path to prevent collisions between
    files with the same filename in different directories.
    """
    fname     = Path(pdf_path).name
    sanitized = re.sub(r"[^a-zA-Z0-9_\-]", "_", fname)
    # Always append an 8-char hash of the full absolute path
    path_hash = _hashlib.sha256(
        os.path.abspath(pdf_path).encode()
    ).hexdigest()[:8]
    base = f"{sanitized}_{path_hash}"
    if len(base) <= 64:
        return base
    # Truncate sanitized part to fit, keep full hash suffix
    suffix = _hashlib.sha256(sanitized.encode()).hexdigest()[:32]
    return sanitized[:23] + "_" + path_hash + suffix  # 23+1+8+32 = 64
# ─────────────────────────────────────────────────────────────────────────────
# PATH HELPERS — Windows long-path safe
# ─────────────────────────────────────────────────────────────────────────────
def _win_safe(path: str) -> str:
    r"""
    Prefix path with \\?\ on Windows to bypass the 260-char MAX_PATH limit.
    No-op on Linux/macOS.
    """
    if os.name != "nt":
        return path
    path = os.path.abspath(path)
    if not path.startswith("\\\\?\\"):
        path = "\\\\?\\" + path
    return path


def _safe_makedirs(parent: str, child: str) -> str:
    r"""
    Create parent/child directory and return the plain (non-\\?\) path.
    The \\?\ prefix is applied internally for the makedirs call only.
    Returned path is plain so it can be used in further os.path.join calls.
    """
    folder = os.path.join(parent, child)
    os.makedirs(_win_safe(folder), exist_ok=True)
    return folder          # return PLAIN path — callers join further


def _safe_join(folder: str, filename: str) -> str:
    """Join folder + filename and return plain path."""
    return os.path.join(folder, filename)


def _write_json(path: str, data: dict) -> None:
    r"""Write JSON using \\?\ prefix on Windows for long-path safety."""
    with open(_win_safe(path), "w", encoding="utf-8") as fh:
        json.dump(data, fh, indent=2, ensure_ascii=False)


def _safe_listdir(folder: str) -> list[str]:
    r"""os.listdir with \\?\ prefix on Windows."""
    return os.listdir(_win_safe(folder))


def _safe_isfile(path: str) -> bool:
    r"""os.path.isfile with \\?\ prefix on Windows."""
    return os.path.isfile(_win_safe(path))


def _safe_exists(path: str) -> bool:
    r"""os.path.exists with \\?\ prefix on Windows."""
    return os.path.exists(_win_safe(path))


def _safe_move(src: str, dst: str) -> None:
    r"""shutil.move with \\?\ prefix on Windows for both src and dst."""
    shutil.move(_win_safe(src), _win_safe(dst))


def _safe_copy2(src: str, dst: str) -> None:
    r"""shutil.copy2 with \\?\ prefix on Windows."""
    shutil.copy2(_win_safe(src), _win_safe(dst))


def _safe_rmtree(path: str, ignore_errors: bool = False) -> None:
    r"""shutil.rmtree with \\?\ prefix on Windows."""
    shutil.rmtree(_win_safe(path), ignore_errors=ignore_errors)


def _move_folder_contents(src: str, dst: str) -> None:
    r"""
    Move all files from src into dst.
    Safer than shutil.move(folder) which fails if dst already exists on Windows.
    Both paths get \\?\ prefix applied via helpers.
    """
    _safe_makedirs(os.path.dirname(dst), os.path.basename(dst))
    for item in _safe_listdir(src):
        s = os.path.join(src, item)
        d = os.path.join(dst, item)
        if _safe_isfile(s):
            _safe_move(s, d)
    """
    Move all files from src into dst (both must exist).
    Safer than shutil.move(folder) which fails if dst already exists on Windows.
    """
    os.makedirs(_win_safe(dst), exist_ok=True)
    for item in os.listdir(src):
        s = os.path.join(src, item)
        d = os.path.join(dst, item)
        if os.path.isfile(s):
            shutil.move(s, _win_safe(d))

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
    # Runtime entry point — this function lazily imports the provider batch
    # clients further down, so assert sys.path before any of them fire.
    _ensure_engine_on_path()
    os.makedirs(output_folder, exist_ok=True)
    os.makedirs(manual_output, exist_ok=True)

    # ── Deduplicate by absolute path to prevent duplicate custom_ids ──────
    seen_paths: set[str] = set()
    deduped: list[str] = []
    for p in extracted_pdfs:
        abs_p = os.path.abspath(p)
        if abs_p not in seen_paths:
            seen_paths.add(abs_p)
            deduped.append(p)
        else:
            print(f"[BATCH-DEDUP] Skipping duplicate path: {os.path.basename(p)}")
    extracted_pdfs = deduped
    # ─────────────────────────────────────────────────────────────────────

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
        coa_filtered = filter_coa_for_type(coa_full, stmt_type, sector)
        coa_filtered = compress_coa_text(coa_filtered) 
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

        # ── Sync retry for failed Gemini batch jobs ───────────────────────────
        # Large PDFs (e.g. OVERVIEW with 11 pages) can exceed the per-request
        # inline-data limit in the batch JSONL, causing the model to return an
        # empty candidate ("Candidate had no text part"). Retry those jobs
        # synchronously so the deal is not silently dropped.
        failed_cids = [cid for cid, r in raw_results.items() if not r.get("ok")]
        if failed_cids:
            job_by_cid = {j["custom_id"]: j for j in jobs}
            retry_jobs = [job_by_cid[cid] for cid in failed_cids if cid in job_by_cid]
            print(f"\n[BATCH-RETRY] {len(retry_jobs)} failed job(s) — retrying synchronously ...")
            for rj in retry_jobs:
                rj_fname = os.path.basename(rj["pdf_path"])
                print(f"  → {rj_fname}")
                try:
                    retry_res = normalize_one_gemini_wrapper(
                        pdf_path         = rj["pdf_path"],
                        prompt_path      = os.path.join(
                            prompts_folder,
                            SUFFIX_TO_PROMPT.get(rj["stmt_type"], ""),
                        ),
                        coa_text         = rj["coa_text"],
                        reporting_columns= rj["reporting_columns"],
                        model            = model,
                        max_tokens       = rj["max_tokens"],
                        page_info        = rj.get("page_info"),
                        stmt_type        = rj["stmt_type"],
                        coa_map_rules    = rj.get("coa_map_rules", ""),
                    )
                    raw_results[rj["custom_id"]] = retry_res
                    status = "✔ OK" if retry_res.get("ok") else f"✘ FAIL — {retry_res.get('error','')}"
                    print(f"    {status}: {rj_fname}")
                except Exception as exc:
                    print(f"    ✘ RETRY ERROR: {rj_fname} — {exc}")

    elif provider == "claude":
        from claude_batch_client import submit_claude_batch, wait_and_download_claude
 
        batch_id, id_map = submit_claude_batch(
            jobs=jobs,
            model=model,
            display_name="fs-claude-batch",
        )
 
        raw_results: dict[str, dict] = wait_and_download_claude(
            batch_id=batch_id,
            model=model,
            jobs=jobs,
            id_map=id_map,
        )
 
        # Re-expand compact JSON using each job's real stmt_type.
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

        deal_passed      = True
        deal_cost        = 0.0
        deal_json_paths: list[str] = []

        # Create dest_folder upfront — write JSONs directly here, no temp→move
        dest_folder = _safe_makedirs(output_folder, deal_name)

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

            deal_cost += res.get("cost_usd", 0.0)

            if provider == "claude":
                prop_snp_split   = res.get("prop_snp_split", [])
                split_stmt_types = {"_PROP_SNP", "_PROP_IS", "_PROP_CFS"}

                if stmt_type in split_stmt_types and prop_snp_split:
                    for idx, sub_data in enumerate(prop_snp_split):
                        out_stem = prop_stmt_output_stem(
                            base_stem, sub_data.get("Metadata", {}), idx
                        )
                        out_path = _safe_join(dest_folder, f"{out_stem}.json")
                        if _COORD_AVAILABLE:
                            sub_data = attach_coordinates(sub_data, pdf_path)
                        _write_json(out_path, sub_data)
                        deal_json_paths.append(out_path)
                else:
                    data = res.get("data")
                    if not data:
                        print(f"    ✘ FAIL     : {fname}  — no data in Claude result")
                        deal_passed = False
                        continue
                    out_path = _safe_join(dest_folder, f"{base_stem}.json")
                    if _COORD_AVAILABLE:
                        data = attach_coordinates(data, pdf_path)
                    _write_json(out_path, data)
                    deal_json_paths.append(out_path)

            else:
                json_text = res.get("json_text", "")
                if not json_text:
                    print(f"    ✘ FAIL     : {fname}  — empty json_text in result")
                    deal_passed = False
                    continue

                delimiter = get_table_break(stmt_type)
                if delimiter and delimiter in json_text:
                    split_results = save_multi_table_results(json_text, pdf_path, stmt_type)
                    if not split_results:
                        print(f"    ✘ PARSE-FAIL : {fname}  — split delimiter found "
                              f"but no sub-table parsed")
                        deal_passed = False
                        continue
                    for out_stem, sub_data in split_results:
                        out_path = _safe_join(dest_folder, f"{out_stem}.json")
                        if _COORD_AVAILABLE:
                            sub_data = attach_coordinates(sub_data, pdf_path)
                        _write_json(out_path, sub_data)
                        deal_json_paths.append(out_path)
                else:
                    try:
                        parsed   = parse_json_response(json_text)
                        expanded = expand_compact_json(parsed, stmt_type=stmt_type)
                    except Exception as e:
                        print(f"    ✘ PARSE-FAIL : {fname}  — {e}")
                        deal_passed = False
                        continue
                    out_path = _safe_join(dest_folder, f"{base_stem}.json")
                    if _COORD_AVAILABLE:
                        expanded = attach_coordinates(expanded, pdf_path)
                    _write_json(out_path, expanded)
                    deal_json_paths.append(out_path)

        total_cost += deal_cost

        # ── Re-route to manual_output if any table failed ─────────────────
        if not deal_passed:
            new_dest = _safe_makedirs(manual_output, deal_name)
            _move_folder_contents(dest_folder, new_dest)
            _safe_rmtree(dest_folder, ignore_errors=True)
            dest_folder = new_dest
            fail_count += 1
            tag = "❌ FAIL → Manual"
        else:
            pass_count += 1
            tag = "✅ ALL EXTRACTION PASS"

        print(f"\n    {tag}")

        # ── JSON → CSV ────────────────────────────────────────────────────
        json_files_in_dest = [
            _safe_join(dest_folder, f)
            for f in _safe_listdir(dest_folder)
            if f.endswith(".json")
        ]

        total_check_failed = False
        for json_file in json_files_in_dest:
            csv_file = json_file.replace(".json", ".csv")
            try:
                csv_pass = run_json_to_csv_pipeline(
                    _win_safe(json_file), _win_safe(csv_file)
                )
            except Exception as e:
                print(f"    [WARN] JSON→CSV failed for "
                      f"{os.path.basename(json_file)}: {e}")
                total_check_failed = True
                continue
            if csv_pass is False:
                total_check_failed = True

        # ── Re-route to manual if total check failed ───────────────────────
        if deal_passed and total_check_failed:
            new_dest = _safe_makedirs(manual_output, deal_name)
            _safe_rmtree(new_dest, ignore_errors=True)
            _safe_makedirs(manual_output, deal_name)
            _move_folder_contents(dest_folder, new_dest)
            _safe_rmtree(dest_folder, ignore_errors=True)
            dest_folder = new_dest
            pass_count -= 1
            fail_count += 1
            deal_passed = False

        try:
            merge_deal_csvs_to_excel(dest_folder, deal_name)
        except Exception as e:
            print(f"    [WARN] CSV→Excel merge failed for {deal_name}: {e}")

        # ── Copy raw PDF into dest ─────────────────────────────────────────
        raw_pdf_candidates = [
            os.path.join(raw_folder, f)
            for f in os.listdir(raw_folder)
            if f.lower().endswith(".pdf")
            and not _PRODUCED_RE.search(f)
            and get_base_pdf_name(Path(f).stem) == deal_name
        ]
        for raw_pdf in raw_pdf_candidates:
            dest = _safe_join(dest_folder, os.path.basename(raw_pdf))
            if not _safe_exists(dest):
                _safe_copy2(raw_pdf, dest)

      except Exception as e:
        print(f"\n    ❌ UNEXPECTED ERROR processing deal '{deal_name}' — "
              f"skipping to next deal. Reason: {type(e).__name__}: {e}")
        fail_count += 1
        continue

    print(f"\n{'='*70}")
    print(f"  ✅ BATCH pipeline completed!")
    print(f"     Deals PASS : {pass_count}  |  Deals FAIL: {fail_count}")
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
        coa_filtered = compress_coa_text(coa_filtered) 
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

    # FIX (Concern 4): collapsed the Gemini branch — it was identical to the
    # else/default (both set max_workers=3). Ollama remains distinct at 1.
    if provider == "ollama":
        max_workers = 1
    else:
        max_workers = 3   # OpenAI, Gemini, Claude all use 3
                          # (_claude_semaphore limits Claude concurrency independently)

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


# Statement types whose OUTPUT CARRIES A PER-ROW "Page No" column, as opposed
# to a single Metadata.Page No for the whole statement. These need the slice →
# original page map (see build_page_note) because the sliced PDF handed to the
# LLM is renumbered 1..N, so a page number read from the slice would be wrong.
_PER_ROW_PAGE_NO_TYPES: set[str] = {
    "_OVERVIEW", "_TAX_BASE", "_PEN", "_OPEB", "_FAQS",
}


def build_page_note(page_info: dict, stmt_type: str = "UNKNOWN") -> str:
    pages_str = ",".join(str(p) for p in page_info["pages"])

    # ── Per-row Page No tabs: supply the slice → original page mapping ──
    # The other statement types get ONE Metadata.Page No covering the whole
    # table, and the note below explicitly forbids per-row page numbers. That
    # instruction is correct for them and WRONG for these five, whose specs
    # require a physical page on every row (and "-" where untraceable).
    if stmt_type in _PER_ROW_PAGE_NO_TYPES:
        pages = page_info["pages"]
        mapping = "\n".join(
            f"  slice page {i} = original page {orig}"
            for i, orig in enumerate(pages, start=1)
        )
        return (
            "SOURCE PAGE REFERENCE — PER-ROW PAGE NUMBERS REQUIRED:\n"
            "The attached PDF is a SLICE of the original full financial report. "
            "Its pages have been RENUMBERED 1.." f"{len(pages)}"
            ". A page number you read from this slice, or from any footer "
            "printed inside it, is NOT the original page number.\n\n"
            "SLICE → ORIGINAL PAGE MAP (use this for every page reference):\n"
            f"{mapping}\n\n"
            f"Metadata.Page No MUST be exactly this comma-joined string:\n"
            f"  \"{pages_str}\"\n\n"
            "This statement's output carries a PER-ROW \"Page No\" column. For "
            "EVERY row, set \"Page No\" to the ORIGINAL page number(s) — "
            "translated through the map above — on which that row's evidence "
            "physically appears. Join multiple pages with comma+space. Use the "
            "literal \"-\" where the row's value is genuinely untraceable.\n"
            "Never emit a slice-local page number (1.." f"{len(pages)}"
            ") as a row's Page No, and never emit a page number printed in the "
            "PDF body, footer, or header."
        )

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
# SINGLE-DEAL ENTRY POINT (used by the DB workflow — PFG_Extraction.py)
# ─────────────────────────────────────────────────────────────────────────────
# run_normalization / run_normalization_batch are folder-driven: they discover
# every deal under a folder, loop, and print a tally. The DB stage instead owns
# ONE TProcessStatus row == ONE PDF == ONE deal, and needs the per-deal verdict
# returned rather than printed. process_one_deal is that seam: same work, one
# deal, an outcome dict instead of a summary.
#
# It is self-contained (obtains its own LLM results) so it works for a single
# deal in either mode:
#   batch=True  → build jobs for this deal's slices, submit ONE batch, download
#   batch=False → process_deal_tables_parallel (ThreadPoolExecutor, 3 at a time)
# Both are normalized to a {pdf_path: result} map, and the post-processing below
# branches on whether a result carries raw `json_text` (gemini/openai batch) or
# already-parsed `data` (sync, and claude batch) — NOT on provider.

def _submit_batch_for_pdfs(
    pdfs: list[str],
    prompts_folder: str,
    coa_text: str,
    provider: str,
    model: str,
    max_tokens: int,
    reporting_columns,
    coa_mapping_folder: str,
    sector: str,
) -> dict[str, dict]:
    """Build + submit + await one Batch-API job covering `pdfs`, returning
    {pdf_path: result}. Mirrors run_normalization_batch's job phase, except the
    sector is passed IN (the DB stage derives it from TCOAMaster.SegmentId, not
    from the filename) so COA filtering and mapping rules match the DB row."""
    jobs: list[dict] = []
    cid_to_pdf: dict[str, str] = {}

    for pdf_path in pdfs:
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
        coa_filtered  = compress_coa_text(filter_coa_for_type(coa_text, stmt_type, sector))
        coa_map_rules = load_coa_mapping(coa_mapping_folder, sector, stmt_type, silent=True)
        cid           = make_batch_custom_id(pdf_path)
        cid_to_pdf[cid] = pdf_path

        jobs.append({
            "custom_id":          cid,
            "pdf_path":           pdf_path,
            "prompt_text":        prompt_text,
            "coa_text":           coa_filtered,
            "coa_map_rules":      coa_map_rules,
            "reporting_columns":  reporting_columns,
            "stmt_type":          stmt_type,
            "max_tokens":         max_tokens,
            "page_info":          extract_page_info_from_filename(pdf_path),
            "coa_mapping_folder": coa_mapping_folder,
        })

    if not jobs:
        return {}

    print(f"\n[BATCH] Submitting {len(jobs)} job(s) to {provider.upper()} Batch API ...")

    if provider == "openai":
        from openai_batch_client import submit_openai_batch, wait_and_download_openai
        raw_results = wait_and_download_openai(submit_openai_batch(jobs, model), model)

    elif provider == "gemini":
        from gemini_batch_client import submit_gemini_batch, wait_and_download_gemini
        raw_results = wait_and_download_gemini(submit_gemini_batch(jobs, model), model)
        # Large slices can blow the per-request inline-data limit and come back with
        # no candidate; retry those synchronously so the deal is not silently dropped.
        job_by_cid  = {j["custom_id"]: j for j in jobs}
        failed_cids = [c for c, r in raw_results.items() if not r.get("ok")]
        if failed_cids:
            print(f"\n[BATCH-RETRY] {len(failed_cids)} failed job(s) — retrying synchronously ...")
            for cid in failed_cids:
                rj = job_by_cid.get(cid)
                if rj is None:
                    continue
                try:
                    raw_results[cid] = normalize_one_gemini_wrapper(
                        pdf_path          = rj["pdf_path"],
                        prompt_path       = os.path.join(
                            prompts_folder, SUFFIX_TO_PROMPT.get(rj["stmt_type"], "")),
                        coa_text          = rj["coa_text"],
                        reporting_columns = rj["reporting_columns"],
                        model             = model,
                        max_tokens        = rj["max_tokens"],
                        page_info         = rj.get("page_info"),
                        stmt_type         = rj["stmt_type"],
                        coa_map_rules     = rj.get("coa_map_rules", ""),
                    )
                except Exception as exc:
                    print(f"    ✘ RETRY ERROR: {os.path.basename(rj['pdf_path'])} — {exc}")

    elif provider == "claude":
        from claude_batch_client import submit_claude_batch, wait_and_download_claude
        batch_id, id_map = submit_claude_batch(
            jobs=jobs, model=model, display_name="fs-claude-batch")
        raw_results = wait_and_download_claude(
            batch_id=batch_id, model=model, jobs=jobs, id_map=id_map)
        # Re-expand compact JSON using each job's real stmt_type.
        job_by_id = {j["custom_id"]: j for j in jobs}
        for cid, res in raw_results.items():
            if res.get("ok") and res.get("data") is not None and cid in job_by_id:
                res["data"] = expand_compact_json(
                    res["data"], job_by_id[cid].get("stmt_type", ""))

    else:
        raise RuntimeError(
            f"[BATCH] provider={provider!r} does not support Batch API. "
            "Supported: openai, gemini, claude.")

    return {cid_to_pdf[c]: r for c, r in raw_results.items() if c in cid_to_pdf}


def process_one_deal(
    deal_name: str,
    deal_pdfs: list[str],
    prompts_folder: str,
    coa_text: str,
    output_folder: str,
    manual_output: str,
    raw_folder: str,
    provider: str,
    model: str,
    max_tokens: int,
    reporting_columns,
    tier: str = "free",
    coa_mapping_folder: str = "",
    client=None,
    sector: str | None = None,
    deal_idx: int | None = None,
    total_deals: int | None = None,
    batch: bool = False,
) -> dict:
    """Normalize ONE deal end to end and RETURN its verdict.

    `sector` overrides filename-based detection — the DB stage supplies it from
    TCOAMaster.SegmentId. `batch` selects the Batch API over the sync thread pool.

    Returns:
        {
          "passed":             bool,        # every table parsed AND Total Check passed
          "output_folder":      str | None,  # where the JSON/CSV/xlsx ended up
          "json_files":         list[str],   # persisted .json paths in output_folder
          "total_check_failed": bool,
          "error":              str | None,  # only for an unexpected failure
        }
    """
    _ensure_engine_on_path()

    outcome: dict = {
        "passed":             False,
        "output_folder":      None,
        "json_files":         [],
        "total_check_failed": False,
        "error":              None,
    }

    try:
        sector   = sector or detect_sector(deal_name + ".pdf")
        expected = SECTOR_TABLE_SUFFIXES.get(sector, SECTOR_TABLE_SUFFIXES["LG"])

        _next_log_pid(deal_name)
        print(f"\n{'─'*70}")
        idx_note = f"[{deal_idx}/{total_deals}] " if deal_idx and total_deals else ""
        print(f"  {idx_note}Deal : {deal_name} || Sector: {sector} || "
              f"mode: {'BATCH' if batch else 'SYNC'}")

        sector_allowed_pdfs = [p for p in deal_pdfs if detect_type(p) in expected]
        skipped = len(deal_pdfs) - len(sector_allowed_pdfs)
        if skipped:
            print(f"    [SECTOR FILTER] Removed {skipped} PDF(s) not allowed for sector={sector}")
        if not sector_allowed_pdfs:
            outcome["error"] = f"no valid tables for sector={sector}"
            print(f"    [SKIP] {outcome['error']}")
            return outcome

        # ── Obtain LLM results, normalized to {pdf_path: result} ──────────────
        if batch:
            results = _submit_batch_for_pdfs(
                pdfs               = sector_allowed_pdfs,
                prompts_folder     = prompts_folder,
                coa_text           = coa_text,
                provider           = provider,
                model              = model,
                max_tokens         = max_tokens,
                reporting_columns  = reporting_columns,
                coa_mapping_folder = coa_mapping_folder,
                sector             = sector,
            )
        else:
            results = process_deal_tables_parallel(
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
        deal_json_paths: list[str] = []
        dest_folder = _safe_makedirs(output_folder, deal_name)

        for pdf_path in sector_allowed_pdfs:
            fname     = os.path.basename(pdf_path)
            stmt_type = detect_type(fname) or "UNKNOWN"
            base_stem = Path(pdf_path).stem
            res       = results.get(pdf_path)

            if res is None:
                print(f"    ✘ MISSING  : {fname}  (no result returned)")
                deal_passed = False
                continue
            if res.get("error") or res.get("ok") is False:
                print(f"    ✘ FAIL     : {fname}  — {res.get('error', 'unknown error')}")
                deal_passed = False
                continue

            json_text = res.get("json_text")
            if json_text:
                # Raw text (gemini / openai batch) — split or parse+expand here.
                delimiter = get_table_break(stmt_type)
                if delimiter and delimiter in json_text:
                    split_results = save_multi_table_results(json_text, pdf_path, stmt_type)
                    if not split_results:
                        print(f"    ✘ PARSE-FAIL : {fname} — delimiter found but no sub-table parsed")
                        deal_passed = False
                        continue
                    for out_stem, sub_data in split_results:
                        out_path = _safe_join(dest_folder, f"{out_stem}.json")
                        if _COORD_AVAILABLE:
                            sub_data = attach_coordinates(sub_data, pdf_path)
                        _write_json(out_path, sub_data)
                        deal_json_paths.append(out_path)
                else:
                    try:
                        expanded = expand_compact_json(
                            parse_json_response(json_text), stmt_type=stmt_type)
                    except Exception as e:
                        print(f"    ✘ PARSE-FAIL : {fname} — {e}")
                        deal_passed = False
                        continue
                    out_path = _safe_join(dest_folder, f"{base_stem}.json")
                    if _COORD_AVAILABLE:
                        expanded = attach_coordinates(expanded, pdf_path)
                    _write_json(out_path, expanded)
                    deal_json_paths.append(out_path)
            else:
                # Already-parsed data (sync any provider, and claude batch).
                prop_split = res.get("prop_snp_split") or []
                if stmt_type in {"_PROP_SNP", "_PROP_IS", "_PROP_CFS"} and prop_split:
                    for i, sub_data in enumerate(prop_split):
                        out_stem = prop_stmt_output_stem(
                            base_stem, sub_data.get("Metadata", {}), i)
                        out_path = _safe_join(dest_folder, f"{out_stem}.json")
                        if _COORD_AVAILABLE:
                            sub_data = attach_coordinates(sub_data, pdf_path)
                        _write_json(out_path, sub_data)
                        deal_json_paths.append(out_path)
                else:
                    data = res.get("data")
                    if not data:
                        print(f"    ✘ FAIL     : {fname} — no data in result")
                        deal_passed = False
                        continue
                    out_path = _safe_join(dest_folder, f"{base_stem}.json")
                    if _COORD_AVAILABLE:
                        data = attach_coordinates(data, pdf_path)
                    _write_json(out_path, data)
                    deal_json_paths.append(out_path)

        # ── Route to manual_output if any table failed ─────────────────────────
        if not deal_passed:
            new_dest = _safe_makedirs(manual_output, deal_name)
            _move_folder_contents(dest_folder, new_dest)
            _safe_rmtree(dest_folder, ignore_errors=True)
            dest_folder = new_dest
        print(f"\n    {'❌ FAIL → Manual' if not deal_passed else '✅ ALL EXTRACTION PASS'}")

        # ── JSON → CSV (this is where Total Check is computed) ────────────────
        json_files_in_dest = [
            _safe_join(dest_folder, f)
            for f in _safe_listdir(dest_folder) if f.endswith(".json")
        ]
        total_check_failed = False
        for json_file in json_files_in_dest:
            try:
                csv_pass = run_json_to_csv_pipeline(
                    _win_safe(json_file), _win_safe(json_file.replace(".json", ".csv")))
            except Exception as e:
                print(f"    [WARN] JSON→CSV failed for {os.path.basename(json_file)}: {e}")
                total_check_failed = True
                continue
            if csv_pass is False:
                total_check_failed = True

        # ── A Total Check failure demotes an otherwise-passing deal ───────────
        if deal_passed and total_check_failed:
            print("\n    ⚠ TOTAL CHECK FAIL — re-routing PASS → Manual Validation")
            new_dest = _safe_makedirs(manual_output, deal_name)
            _safe_rmtree(new_dest, ignore_errors=True)
            new_dest = _safe_makedirs(manual_output, deal_name)
            _move_folder_contents(dest_folder, new_dest)
            _safe_rmtree(dest_folder, ignore_errors=True)
            dest_folder = new_dest
            deal_passed = False

        try:
            merge_deal_csvs_to_excel(dest_folder, deal_name)
        except Exception as e:
            print(f"    [WARN] CSV→Excel merge failed for {deal_name}: {e}")

        # ── Keep the source PDF beside its output ─────────────────────────────
        for raw_pdf in [
            os.path.join(raw_folder, f)
            for f in os.listdir(raw_folder)
            if f.lower().endswith(".pdf")
            and not _PRODUCED_RE.search(f)
            and get_base_pdf_name(Path(f).stem) == deal_name
        ]:
            dest = _safe_join(dest_folder, os.path.basename(raw_pdf))
            if not _safe_exists(dest):
                _safe_copy2(raw_pdf, dest)

        outcome["passed"]             = deal_passed
        outcome["output_folder"]      = dest_folder
        outcome["total_check_failed"] = total_check_failed
        outcome["json_files"]         = [
            _safe_join(dest_folder, f)
            for f in _safe_listdir(dest_folder) if f.endswith(".json")
        ]
        return outcome

    except Exception as e:
        outcome["error"] = f"{type(e).__name__}: {e}"
        print(f"\n    ❌ UNEXPECTED ERROR processing deal '{deal_name}': {outcome['error']}")
        return outcome


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
    # Runtime entry point — lazily imports the provider cache clients downstream.
    _ensure_engine_on_path()
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
 
    for deal_idx, (deal_name, deal_pdfs) in enumerate(deal_groups.items(), 1):
      try:
        sector   = detect_sector(deal_name + ".pdf")
        expected = SECTOR_TABLE_SUFFIXES.get(sector, SECTOR_TABLE_SUFFIXES["LG"])
 
        _next_log_pid(deal_name)
 
        print(f"\n{'─'*70}")
        print(f"  [{deal_idx}/{total_deals}] Deal : {deal_name} || Sector: {sector}")
 
        sector_allowed_pdfs = [
            p for p in deal_pdfs
            if detect_type(p) in expected
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
 
                deal_cost += cost
 
                base_stem          = Path(pdf_path).stem
                stmt_type_detected = detect_type(os.path.basename(pdf_path))
 
                prop_snp_split   = wrapped.get("prop_snp_split")
                split_stmt_types = {"_PROP_SNP", "_PROP_IS", "_PROP_CFS"}
 
                # ── SPLIT CASE (PROP_SNP / PROP_IS / PROP_CFS) ──────────────
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
 
                        # ★ COORD ADDED
                        if _COORD_AVAILABLE:
                            sub_data = attach_coordinates(sub_data, pdf_path)
 
                        with open(sub_path, "w", encoding="utf-8") as f:
                            json.dump(sub_data, f, indent=2, ensure_ascii=False)
                        saved_paths.append(sub_path)
 
                    wrapped["json_path"]       = saved_paths[0] if saved_paths else None
                    wrapped["json_path_extra"] = saved_paths[1:]
 
                # ── NORMAL CASE (SNP, SOA, GOV_BS, GOV_IS, DSR, DEBT) ───────
                else:
                    json_filename  = f"{base_stem}.json"
                    temp_json_path = os.path.join(os.path.dirname(pdf_path), json_filename)
 
                    # ★ COORD ADDED
                    if _COORD_AVAILABLE:
                        json_data = attach_coordinates(json_data, pdf_path)
 
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
    print(f"     Deals PASS: {pass_count}  |  Deals FAIL: {fail_count}")

# ─────────────────────────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────────────────────────

def main() -> None:
    _enable_windows_long_paths()
    args = parse_args()

    if args.batch and args.tier != "paid":
        print("[ERROR] --batch requires --tier paid. Re-run with --tier paid --batch.")
        sys.exit(1)

    if args.provider == "claude":
        if args.tier != "paid":
            print("[INFO] Claude provider selected — forcing --tier paid because Claude has no free tier.")
        args.tier = "paid"


    model = resolve_model(args.provider, args.model)

    # ── Initialize DuckDB log writer (captures both ESG and financial runs) ──
    global _pipeline_log_writer
    try:
        from log_writer import LogWriter
        _pipeline_log_writer = LogWriter()
    except ImportError:
        _pipeline_log_writer = None

    # ═════════════════════════════════════════════════════════════════════════
    # ★ ESG PATCH — DISPATCH TO ESG PIPELINE INSTEAD OF FINANCIAL PIPELINE
    # ═════════════════════════════════════════════════════════════════════════
    if args.esg_mode != "none":
        try:
            from esg_pipeline import run_esg_pipeline
        except ImportError as e:
            sys.exit(f"[ERROR] esg_pipeline.py not found next to pipeline.py: {e}")

        _next_log_pid(f"ESG_{args.esg_mode.upper()}")
        esg_input = os.path.join(args.folder, "ESG")

        try:
            run_esg_pipeline(
                esg_folder     = esg_input,
                prompts_folder = args.prompts,
                master_folder  = args.xlsx,
                output_folder  = args.output,
                manual_folder  = args.manual_output,
                provider       = args.provider,
                model          = model,
                max_tokens     = args.max_tokens,
                tier           = args.tier,
                mode           = args.esg_mode,
                batch          = args.batch,              # ← NEW
                concurrent_deals   = None,                # ← NEW (env var override)
                intra_deal_workers = None,                # ← NEW (env var override)
            )
        finally:
            if _pipeline_log_writer is not None:
                try:
                    _pipeline_log_writer.close()
                    print(f"[LOG] Parquet saved → logs/pipeline_logs.parquet")
                except Exception as e:
                    print(f"[WARN] Log writer close failed: {e}")
        return
    # ═════════════════════════════════════════════════════════════════════════
    # END OF ESG PATCH — remainder is the financial pipeline
    # ═════════════════════════════════════════════════════════════════════════

    # FIX (Bug 1 / Concern 2): mode_label is now actually printed, and
    # tier_label alias removed — args.tier passed directly to run_normalization().
    mode_label = (
        "⚡ BATCH (async, ≤24 h, 50% cost)"
        if (args.tier == "paid" and getattr(args, "batch", False))
        else f"🔄 SYNC (tier={args.tier})"
    )
    #print(f"\n[INFO] Provider : {args.provider.upper()}  |  Model : {model}  |  Mode : {mode_label}")

    if args.skip_extraction:
        extracted_pdfs = collect_existing_extracted(
            args.folder, prompts_folder=args.prompts
        )
    else:
        existing = collect_existing_extracted(args.folder, prompts_folder=args.prompts)

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
            use_llm_id = getattr(args, "llm_page_id", False)
            id_model   = getattr(args, "id_model", "claude-sonnet-4-6")
            use_batch  = args.tier == "paid" and getattr(args, "batch", False)

            if use_llm_id and use_batch and getattr(args, "skip_normalization", False):
                # Batch page-extraction mode: submit all page-ID LLM calls as one
                # Claude batch job, wait for results, then slice PDFs.
                print(f"\n[INFO] Mode: BATCH PAGE EXTRACTION (id_model={id_model})")
                newly_extracted = run_extraction_batch(
                    folder         = args.folder,
                    prompts_folder = args.prompts,
                    id_model       = id_model,
                )
            else:
                newly_extracted = run_extraction(
                    args.folder,
                    prompts_folder = args.prompts,
                    use_llm_id     = use_llm_id,
                    id_model       = id_model,
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

    print(f"\n{'='*70}")
    print(f"DATA Extraction IS RUNNING....")
    print(f"{'='*70}")

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
            tier               = args.tier,          # FIX (Concern 2): pass args.tier directly
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


if __name__ == "__main__":
    main()