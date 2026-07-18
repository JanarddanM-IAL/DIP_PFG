import argparse
import base64
import json
import os
import re
import shutil
import sys
from pathlib import Path
import openpyxl
from openai import OpenAI, APIError, APITimeoutError, RateLimitError
from pdf_to_indented_text import pdf_to_indented_text
import pyodbc
from dotenv import load_dotenv



# Load .env (GEMINI_API_KEY / OPENAI_API_KEY / DB creds live here)
load_dotenv()


try:
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")
except Exception:
    pass

try:
    from page_extractor import find_snp_pages, extract_pages_to_pdf
except ImportError:
    sys.exit(
        "[ERROR] page_extractor.py not found next to pipeline.py.\n"
        "        Place page_extractor.py in the same folder and retry."
    )

try:
    from jsonToCsv import run_json_to_csv_pipeline
except ImportError as e:
    sys.exit(f"[ERROR] Could not import jsonToCsv.py: {e}")


# ─────────────────────────────────────────────────────────────────────────────
# CONFIG (was CLI args) — paths relative to this file where co-located
# ─────────────────────────────────────────────────────────────────────────────

BASE_DIR = Path(__file__).resolve().parent

# Co-located with the code:
XLSX_PATH      = str(BASE_DIR / "standard_coa_master.xlsx")
PROMPTS_FOLDER = str(BASE_DIR / "prompts")

# Absolute output locations:
OUTPUT_BASE_PATH = r"C:\Test Code\Public Finance\04_Validated_Output"
MANUAL_OUTPUT    = r"C:\Test Code\Public Finance\05_Manual_validation_required"

# LLM settings:
PROVIDER   = "gemini"
# MODEL      = "gemini-2.5-pro"
MODEL      = "gemini-3.5-flash"
MAX_TOKENS = 30000

# Optional reporting columns (was --reporting-columns)
REPORTING_COLUMNS = None


# ─────────────────────────────────────────────────────────────────────────────
# DATABASE
# ─────────────────────────────────────────────────────────────────────────────

def get_db_connection():
    """Open a pyodbc connection using credentials from .env."""
    driver   = os.environ.get("DB_DRIVER", "ODBC Driver 17 for SQL Server")  # no braces in default
    server   = os.environ.get("DB_SERVER")
    database = os.environ.get("DB_DATABASE")
    username = os.environ.get("DB_USERNAME")
    password = os.environ.get("DB_PASSWORD")
    trusted  = os.environ.get("DB_TRUSTED_CONNECTION", "no").strip().lower() in ("yes", "true", "1")

    if not server or not database:
        raise RuntimeError("DB_SERVER and DB_DATABASE must be set in .env file.")

    if trusted:
        conn_str = (
            f"DRIVER={{{driver}}};"        # wraps to → {ODBC Driver 17 for SQL Server}
            f"SERVER={server};"
            f"DATABASE={database};"
            f"Trusted_Connection=yes;"
            f"TrustServerCertificate=yes;"
        )
    else:
        if not username or not password:
            raise RuntimeError("DB_USERNAME and DB_PASSWORD must be set in .env.")
        conn_str = (
            f"DRIVER={{{driver}}};"        # wraps to → {ODBC Driver 17 for SQL Server}
            f"SERVER={server};"
            f"DATABASE={database};"
            f"UID={username};"
            f"PWD={password};"
            f"TrustServerCertificate=yes;"
        )

    return pyodbc.connect(conn_str)

# ─────────────────────────────────────────────────────────────────────────────
# DB STATUS UPDATES (statuses are int 0/1, flags are char 'p'/'c')
# ─────────────────────────────────────────────────────────────────────────────

def update_extraction_flag(conn, processing_id, flag):
    """Set ExtractionFlag to 'p' (in progress) or 'c' (complete)."""
    cur = conn.cursor()
    cur.execute(
        "UPDATE TProcessStatus SET ExtractionFlag = ?, ModifiedOn = GETDATE() "
        "WHERE ProcessingId = ?", flag, processing_id)
    conn.commit()


def update_extraction_status(conn, processing_id, status):
    """Set ExtractionStatus 1 (success) / 0 (fail)."""
    cur = conn.cursor()
    cur.execute(
        "UPDATE TProcessStatus SET ExtractionStatus = ?, ModifiedOn = GETDATE() "
        "WHERE ProcessingId = ?", status, processing_id)
    conn.commit()


def update_data_validation(conn, processing_id, status=None, flag=None):
    """Set DataValidationStatus (1/0) and/or DataValidationFlag ('p'/'c')."""
    sets, params = ["ModifiedOn = GETDATE()"], []
    if status is not None:
        sets.append("DataValidationStatus = ?"); params.append(status)
    if flag is not None:
        sets.append("DataValidationFlag = ?"); params.append(flag)
    params.append(processing_id)
    cur = conn.cursor()
    cur.execute(f"UPDATE TProcessStatus SET {', '.join(sets)} WHERE ProcessingId = ?", *params)
    conn.commit()


def update_output_path(conn, processing_id, output_path):
    """Store the produced .xlsx path in OutputPath."""
    cur = conn.cursor()
    cur.execute(
        "UPDATE TProcessStatus SET OutputPath = ?, ModifiedOn = GETDATE() "
        "WHERE ProcessingId = ?", str(output_path), processing_id)
    conn.commit()


def update_completion_status(conn, processing_id, extraction_status, data_validation_status):
    """
    CompletionStatus = 1 iff ExtractionStatus=1 AND DataValidationStatus=1.
    (SourcingStatus / SourcingValidationStatus are already 1 by the SELECT filter.)
    """
    completion = 1 if (extraction_status == 1 and data_validation_status == 1) else 0
    cur = conn.cursor()
    cur.execute(
        "UPDATE TProcessStatus SET CompletionStatus = ?, ModifiedOn = GETDATE() "
        "WHERE ProcessingId = ?", completion, processing_id)
    conn.commit()
    return completion


def update_remarks(conn, processing_id, remarks):
    """Append an operational note to Remarks (extraction edge cases)."""
    cur = conn.cursor()
    cur.execute(
        "UPDATE TProcessStatus SET Remarks = ?, ModifiedOn = GETDATE() "
        "WHERE ProcessingId = ?", remarks, processing_id)
    conn.commit()

# ─────────────────────────────────────────────────────────────────────────────
# STATEMENT TYPE → PROMPT FILE MAPPING
# ─────────────────────────────────────────────────────────────────────────────

SUFFIX_TO_PROMPT: dict[str, str] = {
    "_SNP": "SNP_Prompt.txt",
    "_SOA": "SOA_Prompt.txt",
}

TYPE_PATTERNS: list[tuple[re.Pattern, str]] = [
    (re.compile(r"_SNP", re.IGNORECASE), "_SNP"),
    (re.compile(r"_SOA", re.IGNORECASE), "_SOA"),
    (re.compile(r"_BS",  re.IGNORECASE), "_BS"),
    (re.compile(r"_IS",  re.IGNORECASE), "_IS"),
    (re.compile(r"_CF",  re.IGNORECASE), "_CF"),
]

KNOWN_SUFFIXES = tuple(s.lower() for s in SUFFIX_TO_PROMPT)

# Matches the page-number tag appended by financial_statement_extractor.py
# Examples: _SNP_p15  _SOA_p16-17  _BS_p22
_PAGE_TAG_RE = re.compile(
    r"_(?:SNP|SOA|BS|IS|CF)_p(\d+)(?:-(\d+))?$",
    re.IGNORECASE,
)


# ─────────────────────────────────────────────────────────────────────────────
# OPENAI COST TABLE  ($ per 1M tokens)
# ─────────────────────────────────────────────────────────────────────────────

OPENAI_COST_TABLE: dict[str, tuple[float, float]] = {
    # model              input   output
    "gpt-4o-mini":      (0.15,   0.60),
    "gpt-4.1-mini":     (0.40,   1.60),
    "gpt-4o":           (2.50,  10.00),
    "gpt-4.1":          (2.00,   8.00),
    "gpt-5.4-mini":     (0.40,   1.60),
    "gpt-5.5":          (2.00,   8.00),
}


# ─────────────────────────────────────────────────────────────────────────────
# CLI
# ─────────────────────────────────────────────────────────────────────────────

def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(
        description="End-to-end pipeline: PDF folder → page extraction → JSON normalization.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__,
    )
    p.add_argument("--folder",         default="./03_Validated_Report",
                   help="Folder containing raw PDFs (default: ./03_Validated_Report).")
    p.add_argument("--xlsx",           required=True)
    p.add_argument("--prompts",        required=True)
    p.add_argument("--output",         default="./04_Validated_output",
                   help="Output folder for all-PASS results (default: ./04_Validated_output).")
    p.add_argument("--manual-output",  default="./05_Manual_Validation_Required",
                   help="Output folder for FAIL results (default: ./05_Manual_Validation_Required).")
    p.add_argument(
        "--provider",
        default="openai",
        choices=["openai", "gemini"],
        help="LLM provider to use: openai (default) or gemini.",
    )
    p.add_argument(
        "--model",
        default=None,
        help=(
            "Model name. Defaults: openai→gpt-4o-mini, gemini→gemini-3.5-flash. "
            "For Gemini table parsing, recommended: gemini-3.5-flash."
        ),
    )
    p.add_argument("--max-tokens",     type=int, default=16000)
    p.add_argument("--skip-extraction", action="store_true")
    p.add_argument("--reporting-columns", nargs="*", default=None, metavar="COL")
    return p.parse_args()


def cleanup_extracted_pdfs(folder: str):
    """
    Deletes extracted PDFs with page tags like:
    *_SNP_pXX.pdf, *_SOA_pXX-YY.pdf, *_BS_pXX.pdf etc.
    """
    pattern = re.compile(r"_(?:SNP|SOA|BS|IS|CF)_p\d+(?:-\d+)?\.pdf$", re.IGNORECASE)

    deleted = 0
    for f in os.listdir(folder):
        if f.lower().endswith(".pdf") and pattern.search(f):
            try:
                os.remove(os.path.join(folder, f))
                deleted += 1
            except Exception as e:
                print(f"[WARN] Could not delete {f}: {e}")

    print(f"\n[Cleanup] Deleted {deleted} extracted PDF(s) from {folder}")

# ─────────────────────────────────────────────────────────────────────────────
# STEP 1 — PAGE EXTRACTION (single source PDF)
# ─────────────────────────────────────────────────────────────────────────────

def run_extraction_for_pdf(src_pdf: str) -> list[str]:
    """
    Split ONE source PDF into tagged SNP/SOA sub-PDFs written next to it.
    Returns the list of produced sub-PDF paths.
    """
    from page_extractor import find_snp_pages, find_soa_pages, extract_pages_to_pdf, _pages_suffix, START_PAGE

    folder = str(Path(src_pdf).parent)
    stem   = Path(src_pdf).stem
    produced: list[str] = []

    # ── SNP ──────────────────────────────────────────────────────────────────
    snp_pages = None
    try:
        snp_pages  = find_snp_pages(src_pdf)
        snp_suffix = _pages_suffix(snp_pages)
        snp_out    = os.path.join(folder, f"{stem}_SNP{snp_suffix}.pdf")
        extract_pages_to_pdf(src_pdf, snp_pages, snp_out)
        print(f"  {Path(src_pdf).name}  → SNP pages {snp_pages}  →  {Path(snp_out).name}")
        produced.append(snp_out)
    except Exception as e:
        print(f"  {Path(src_pdf).name}  → SNP error: {e}")

    # ── SOA — always starts AFTER last SNP page ──────────────────────────────
    soa_search_start = (max(snp_pages) + 1) if snp_pages else START_PAGE
    try:
        soa_pages  = find_soa_pages(src_pdf, start_page=soa_search_start)
        soa_suffix = _pages_suffix(soa_pages)
        soa_out    = os.path.join(folder, f"{stem}_SOA{soa_suffix}.pdf")
        extract_pages_to_pdf(src_pdf, soa_pages, soa_out)
        print(f"  {Path(src_pdf).name}  → SOA pages {soa_pages}  →  {Path(soa_out).name}")
        produced.append(soa_out)
    except Exception as e:
        print(f"  {Path(src_pdf).name}  → SOA error: {e}")

    return produced

def cleanup_extracted_for_pdf(src_pdf: str):
    """Delete the tagged SNP/SOA sub-PDFs produced from this one source PDF."""
    folder = Path(src_pdf).parent
    stem   = Path(src_pdf).stem
    pattern = re.compile(rf"^{re.escape(stem)}_(?:SNP|SOA|BS|IS|CF)_p\d+(?:-\d+)?\.pdf$", re.IGNORECASE)
    deleted = 0
    for f in os.listdir(folder):
        if pattern.match(f):
            try:
                os.remove(folder / f); deleted += 1
            except Exception as e:
                print(f"[WARN] Could not delete {f}: {e}")
    print(f"  [Cleanup] Deleted {deleted} extracted sub-PDF(s) for {Path(src_pdf).name}")


# ─────────────────────────────────────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────────────────────────────────────

# ─────────────────────────────────────────────────────────────────────────────
# OUTPUT MERGE — combine per-statement CSVs into ONE xlsx (sheet per statement)
# ─────────────────────────────────────────────────────────────────────────────

def merge_csvs_to_xlsx(csv_by_sheet: dict, out_xlsx_path: Path) -> Path:
    """
    Write one .xlsx with one worksheet per statement.

    csv_by_sheet: {"SNP": Path(snp.csv), "SOA": Path(soa.csv), ...}
    Returns the written .xlsx path.
    """
    import csv as _csv
    out_xlsx_path.parent.mkdir(parents=True, exist_ok=True)

    wb = openpyxl.Workbook()
    wb.remove(wb.active)  # drop default sheet

    for sheet_name, csv_path in csv_by_sheet.items():
        ws = wb.create_sheet(title=sheet_name[:31])  # Excel 31-char sheet limit
        with open(csv_path, "r", encoding="utf-8-sig", newline="") as f:
            for row in _csv.reader(f):
                ws.append(row)

    wb.save(str(out_xlsx_path))
    wb.close()
    return out_xlsx_path

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


def load_pdf_as_data_uri(path: str) -> tuple[str, str]:
    raw = Path(path).read_bytes()
    b64 = base64.b64encode(raw).decode("utf-8")
    return f"data:application/pdf;base64,{b64}", Path(path).name


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
    """
    Parse the page-number tag that financial_statement_extractor.py embeds in
    the filename, e.g.:

        CityReport_SNP_p15.pdf      → {"start": 15, "end": 15, "label": "page 15"}
        CityReport_SOA_p16-17.pdf   → {"start": 16, "end": 17, "label": "pages 16–17"}

    Returns None if no tag is found (graceful fallback for old-style filenames).
    """
    stem = Path(pdf_path).stem          # e.g. "CityReport_SNP_p16-17"
    m = _PAGE_TAG_RE.search(stem)
    if not m:
        return None

    start = int(m.group(1))
    end   = int(m.group(2)) if m.group(2) else start

    if start == end:
        label = f"page {start}"
        pages = [start]
    else:
        label = f"pages {start}–{end}"
        pages = list(range(start, end + 1))

    return {"start": start, "end": end, "pages": pages, "label": label}


# ─────────────────────────────────────────────────────────────────────────────
# STEP 1 — PAGE EXTRACTION
# ─────────────────────────────────────────────────────────────────────────────


# all the pdfs from the provided foler
# def run_extraction(folder: str) -> list[str]:
#     from page_extractor import find_snp_pages, find_soa_pages, extract_pages_to_pdf, _pages_suffix, START_PAGE

#     _produced_re = re.compile(r"_(?:SNP|SOA)_p\d+(?:-\d+)?\.pdf$", re.IGNORECASE)

#     raw_pdfs = sorted(
#         f for f in os.listdir(folder)
#         if f.lower().endswith(".pdf")
#         and not any(Path(f).stem.lower().endswith(s) for s in KNOWN_SUFFIXES)
#         and not _produced_re.search(f)
#     )

#     if not raw_pdfs:
#         print("  [WARN] No raw PDFs found in folder (or all already extracted).")
#         return []

#     print(f"  Found {len(raw_pdfs)} raw PDF(s).\n")
#     produced: list[str] = []

#     for fname in raw_pdfs:
#         src  = os.path.join(folder, fname)
#         stem = Path(fname).stem

#         # ── SNP ──────────────────────────────────────────────────────────
#         snp_pages = None
#         try:
#             snp_pages  = find_snp_pages(src)
#             snp_suffix = _pages_suffix(snp_pages)
#             snp_out    = os.path.join(folder, f"{stem}_SNP{snp_suffix}.pdf")
#             extract_pages_to_pdf(src, snp_pages, snp_out)
#             print(f"  {fname}  → SNP pages {snp_pages}  →  {Path(snp_out).name}")
#             produced.append(snp_out)
#         except Exception as e:
#             print(f"  {fname}  → SNP error: {e}")

#         # ── SOA — always starts AFTER last SNP page ───────────────────────
#         soa_search_start = (max(snp_pages) + 1) if snp_pages else START_PAGE
#         try:
#             soa_pages  = find_soa_pages(src, start_page=soa_search_start)
#             soa_suffix = _pages_suffix(soa_pages)
#             soa_out    = os.path.join(folder, f"{stem}_SOA{soa_suffix}.pdf")
#             extract_pages_to_pdf(src, soa_pages, soa_out)
#             print(f"  {fname}  → SOA pages {soa_pages}  →  {Path(soa_out).name}")
#             produced.append(soa_out)
#         except Exception as e:
#             print(f"  {fname}  → SOA error: {e}")

#     return produced

# def collect_existing_extracted(folder: str) -> list[str]:
#     """
#     For each raw PDF in the folder, checks each statement type (_SNP, _BS, _IS, _CF)
#     independently. If an extracted file with a page tag already exists for that
#     type → skip. If extracted but missing page tag → re-extract. If not extracted
#     at all → extract fresh.
#     """
#     folder_path = Path(folder)
#     final = []

#     # Get all raw (non-extracted) PDFs
#     raw_pdfs = sorted(
#         f for f in os.listdir(folder)
#         if f.lower().endswith(".pdf")
#         and not any(Path(f).stem.lower().endswith(s) for s in KNOWN_SUFFIXES)
#     )

#     if not raw_pdfs:
#         print("  [WARN] No raw PDFs found in folder.")
#         return []

#     for raw_fname in raw_pdfs:
#         raw_path = str(folder_path / raw_fname)
#         raw_stem = Path(raw_fname).stem
#         print(f"\n  Checking : {raw_fname}")

#         for suffix, _ in SUFFIX_TO_PROMPT.items():
#             suffix_clean = suffix.lstrip("_")  # e.g. "SNP", "BS", "IS", "CF"

#             # Find any existing extracted file for this raw PDF + suffix combo
#             # Matches: RawStem_SNP.pdf  OR  RawStem_SNP_p15.pdf  OR  RawStem_SNP_p15-16.pdf
#             pattern = re.compile(
#                 rf"^{re.escape(raw_stem)}[_\-]{suffix_clean}(_p\d+(?:-\d+)?)?\.pdf$",
#                 re.IGNORECASE,
#             )
#             matches = [
#                 f for f in os.listdir(folder)
#                 if pattern.match(f)
#             ]

#             if not matches:
#                 # Not extracted at all for this type — skip silently
#                 # (page_extractor decides if this statement type exists in the PDF)
#                 continue

#             existing_path = str(folder_path / matches[0])
#             existing_stem = Path(matches[0]).stem

#             if _PAGE_TAG_RE.search(existing_stem):
#                 # Already has page tag — perfect, use as-is
#                 print(f"    [{suffix_clean}] ✓ Already extracted with page tag: {matches[0]}")
#                 final.append(existing_path)

#             else:
#                 # Extracted but missing page tag — re-extract this type only
#                 print(f"    [{suffix_clean}] ↻ Missing page tag — re-extracting from {raw_fname}")
#                 try:
#                     pages = find_snp_pages(raw_path)

#                     if len(pages) == 1:
#                         page_tag = f"p{pages[0]}"
#                     else:
#                         page_tag = f"p{pages[0]}-{pages[-1]}"

#                     new_out = str(folder_path / f"{raw_stem}_{suffix_clean}_{page_tag}.pdf")
#                     extract_pages_to_pdf(raw_path, pages, new_out)

#                     # Remove old tagless file
#                     os.remove(existing_path)
#                     print(f"    [{suffix_clean}] → Re-extracted: {Path(new_out).name}")
#                     final.append(new_out)

#                 except Exception as e:
#                     print(f"    [{suffix_clean}] [WARN] Re-extraction failed ({e}) — using original as-is")
#                     final.append(existing_path)

#     return final


# ─────────────────────────────────────────────────────────────────────────────
# STEP 2a — NORMALIZE via OPENAI
# ─────────────────────────────────────────────────────────────────────────────

def normalize_one_openai(
    pdf_path: str,
    prompt_path: str,
    coa_text: str,
    reporting_columns: list | None,
    client,          # OpenAI client
    model: str,
    max_tokens: int,
) -> dict:
    import base64, json, re
    from pathlib import Path
    from openai import APITimeoutError, RateLimitError, APIError

    # ── Load inputs ──────────────────────────────────────────────────────────
    system_prompt = Path(prompt_path).read_text(encoding="utf-8")
    pdf_bytes     = Path(pdf_path).read_bytes()
    b64           = base64.b64encode(pdf_bytes).decode("utf-8")
    pdf_data_uri  = f"data:application/pdf;base64,{b64}"
    pdf_name      = Path(pdf_path).name

    # ── Parse page numbers embedded in the filename ──────────────────────────
    page_info = extract_page_info_from_filename(pdf_path)

    # ── Extract indented text from PDF ──────────────────────────────────────
    try:
        from pdf_to_indented_text import pdf_to_indented_text
        indented_text = pdf_to_indented_text(pdf_path)
    except Exception as e:
        print(f"  [WARN] pdf_to_indented_text failed ({e}), falling back to no indented text")
        indented_text = None

    # ── Build instruction lines ──────────────────────────────────────────────
    instruction_lines = [
        "Normalize the financial statement from the attached PDF.",
        "Follow all steps in the system prompt exactly.",
        "Use the COA Master table above for COA Datapoint mapping.",
        "Return valid JSON only — no markdown, no extra text.",
    ]
    if reporting_columns:
        instruction_lines += ["", "[[REPORTING COLUMNS]]"] + reporting_columns

    # ── Build user content blocks ────────────────────────────────────────────
    user_content = [
        # Block 1: PDF file (base64)
        {"type": "file", "file": {"filename": pdf_name, "file_data": pdf_data_uri}},
        # Block 2: COA master
        {
            "type": "text",
            "text": (
                "STANDARD COA MASTER (pipe-delimited):\n"
                "Format: COA Flag | COA Datapoint | Statement | Section\n\n"
                + coa_text
            ),
        },
    ]

    # Block 3 (NEW): Source page numbers from the original full report
    # This tells the LLM exactly which pages of the full document it is reading,
    # so it can correctly populate any "Source Page" or "Page Reference" fields.
    if page_info:
        if page_info["start"] == page_info["end"]:
            page_note = (
                f"SOURCE PAGE REFERENCE:\n"
                f"This extracted PDF corresponds to {page_info['label']} of the "
                f"original full financial report (1-based page numbering).\n"
                f"When populating any 'Source Page', 'Page Number', or 'Page Reference' "
                f"field in the output JSON, use: {page_info['start']}"
            )
        else:
            pages_str = ", ".join(str(p) for p in page_info["pages"])
            page_note = (
                f"SOURCE PAGE REFERENCE:\n"
                f"This extracted PDF corresponds to {page_info['label']} of the "
                f"original full financial report (1-based page numbering).\n"
                f"The content spans pages: {pages_str}.\n"
                f"When populating any 'Source Page', 'Page Number', or 'Page Reference' "
                f"field in the output JSON, use the page where each data point appears."
            )
        user_content.append({"type": "text", "text": page_note})
        print(f"  Source pages : {page_info['label']}")
    else:
        print(f"  Source pages : [WARN] no page tag found in filename — skipping page context")

    # Block 4: Indented text representation of the PDF
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

    # Block 5: Final instruction
    user_content.append({
        "type": "text",
        "text": "\n".join(instruction_lines),
    })

    messages = [
        {"role": "system", "content": system_prompt},
        {"role": "user", "content": user_content},
    ]

    # ── Call OpenAI ──────────────────────────────────────────────────────────
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

    def parse_json_response(text):
        for attempt in [
            lambda t: json.loads(t),
            lambda t: json.loads(re.search(r"```(?:json)?\s*([\s\S]+?)\s*```", t).group(1)),
            lambda t: json.loads(t[t.index("{"):t.rindex("}") + 1]),
        ]:
            try:
                return attempt(text)
            except Exception:
                pass
        raise ValueError(f"Cannot parse JSON. First 400 chars:\n{text[:400]}")

    return {
        "data": parse_json_response(raw),
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
) -> dict:
    try:
        from gemini_client import normalize_one_gemini
    except ImportError as e:
        raise RuntimeError(f"gemini_client.py not found or google-genai not installed: {e}")

    # Parse page info from filename and forward it to the Gemini client
    page_info = extract_page_info_from_filename(pdf_path)
    if page_info:
        print(f"  Source pages : {page_info['label']}")
    else:
        print(f"  Source pages : [WARN] no page tag found in filename — skipping page context")

    try:
        return normalize_one_gemini(
            pdf_path=pdf_path,
            prompt_path=prompt_path,
            coa_text=coa_text,
            reporting_columns=reporting_columns,
            model=model,
            max_tokens=max_tokens,
            page_info=page_info,   # ← forwarded so gemini_client can inject the same context
        )
    except RuntimeError:
        raise
    except Exception as e:
        raise RuntimeError(f"  Unexpected Gemini error: {type(e).__name__}: {e}")


# ─────────────────────────────────────────────────────────────────────────────
# RAW PDF LOCATOR
# ─────────────────────────────────────────────────────────────────────────────

def find_raw_pdf(raw_folder: str, base_name: str) -> Path | None:
    """
    Locate the original (non-extracted) PDF in raw_folder whose stem matches
    base_name.  Returns the Path if found, else None.
    """
    exact = Path(raw_folder) / f"{base_name}.pdf"
    if exact.exists():
        return exact

    candidates = [
        c for c in Path(raw_folder).glob("*.pdf")
        if not any(c.stem.lower().endswith(s) for s in KNOWN_SUFFIXES)
        and c.stem.lower().startswith(base_name.lower())
    ]
    return candidates[0] if candidates else None


# ─────────────────────────────────────────────────────────────────────────────
# STEP 2 — NORMALIZATION ORCHESTRATOR
# ─────────────────────────────────────────────────────────────────────────────

# def run_normalization(
#     extracted_pdfs: list[str],
#     prompts_folder: str,
#     xlsx_path: str,
#     output_folder: str,
#     manual_output: str,
#     raw_folder: str,
#     provider: str,
#     model: str,
#     max_tokens: int,
#     reporting_columns: list[str] | None,
# ) -> None:
#     import csv

#     Path(output_folder).mkdir(parents=True, exist_ok=True)
#     Path(manual_output).mkdir(parents=True, exist_ok=True)

#     # ── Provider-specific setup ──────────────────────────
#     openai_client = None
#     if provider == "openai":
#         api_key = os.environ.get("OPENAI_API_KEY")
#         if not api_key:
#             sys.exit("[ERROR] OPENAI_API_KEY not set.")
#         openai_client = OpenAI(api_key=api_key)
#     else:
#         if not os.environ.get("GEMINI_API_KEY"):
#             sys.exit("[ERROR] GEMINI_API_KEY not set.")

#     print(f"  Loading COA master : {xlsx_path}")
#     coa_text = load_xlsx_as_pipe_text(xlsx_path)
#     print(f"  COA rows loaded    : {coa_text.count(chr(10)) + 1}")
#     print(f"  Provider           : {provider.upper()}  |  Model: {model}")

#     # ── Token log setup ──────────────────────────────────
#     log_file = os.path.join(output_folder, "token_log.csv")
#     if not Path(log_file).exists():
#         with open(log_file, "w", newline="", encoding="utf-8") as f:
#             writer = csv.writer(f)
#             writer.writerow([
#                 "Filename", "Provider", "Model",
#                 "Prompt Tokens", "Completion Tokens", "Total Tokens", "Cost ($)"
#             ])

#     total    = len(extracted_pdfs)
#     ok_count = 0

#     for idx, pdf_path in enumerate(extracted_pdfs, 1):
#         fname    = Path(pdf_path).name
#         stem     = Path(pdf_path).stem

#         print(f"\n  [{idx}/{total}] {fname}")

#         type_key = detect_type(fname)
#         if type_key is None:
#             print(f"  [SKIP] Unknown type — no suffix in: {fname}")
#             continue

#         prompt_filename = SUFFIX_TO_PROMPT[type_key]
#         prompt_path     = os.path.join(prompts_folder, prompt_filename)

#         if not Path(prompt_path).exists():
#             print(f"  [SKIP] Prompt file not found: {prompt_path}")
#             continue

#         print(f"  Type     : {type_key.lstrip('_')}  ->  prompt: {prompt_filename}")
#         print(f"  Calling {provider.upper()} ({model}) ...")

#         # ── Dispatch to correct provider ──────────────────
#         try:
#             if provider == "openai":
#                 response = normalize_one_openai(
#                     pdf_path=pdf_path,
#                     prompt_path=prompt_path,
#                     coa_text=coa_text,
#                     reporting_columns=reporting_columns,
#                     client=openai_client,
#                     model=model,
#                     max_tokens=max_tokens,
#                 )
#             else:
#                 response = normalize_one_gemini_wrapper(
#                     pdf_path=pdf_path,
#                     prompt_path=prompt_path,
#                     coa_text=coa_text,
#                     reporting_columns=reporting_columns,
#                     model=model,
#                     max_tokens=max_tokens,
#                 )
#         except (RuntimeError, ValueError) as e:
#             err_msg = str(e)
#             print(f"\n  [ERROR] {err_msg}")
#             if any(k in err_msg.lower() for k in (
#                 "quota", "exhausted", "429", "resource", "permission",
#                 "unauthenticated", "api key", "deadline", "timeout"
#             )):
#                 print("\n    Fatal API error — aborting pipeline.")
#                 sys.exit(1)
#             continue

#         result = response["data"]
#         usage  = response["usage"]

#         # ── Cost calculation ─────────────────────────────
#         if provider == "openai":
#             total_cost = calc_openai_cost(
#                 model, usage["prompt_tokens"], usage["completion_tokens"]
#             )
#         else:
#             from gemini_client import calculate_cost as gemini_cost
#             total_cost = gemini_cost(
#                 model, usage["prompt_tokens"], usage["completion_tokens"]
#             )

#         print(f"  Tokens   : {usage['total_tokens']:,}")
#         print(f"  Cost     : ${total_cost:.4f}")

#         # ── Derive base name (strip _SNP_pN / _BS_pN / etc. suffix) ──────────
#         base_name = re.sub(
#             r"[_\-](?:SNP|SOA|BS|IS|CF)_p\d+(?:-\d+)?$", "", stem, flags=re.IGNORECASE
#         ).strip("_-")

#         # ── Read all_pass from the LLM-set Total Check Status ─────────────────
#         all_pass = all(
#             str(item.get("Total Check Status", "")) in {
#                 "", "PASS", "Skipped - no membership rule defined"
#             }
#             for section_items in result.get("Sections", {}).values()
#             for item in (section_items if isinstance(section_items, list) else [])
#             if str(item.get("COA Flag", "")).strip().upper() == "CP"
#         )

#         if all_pass:
#             per_pdf_out_dir = Path(output_folder) / base_name
#             print(f"  Check    :  ALL PASS → 04_Validated_output")
#         else:
#             per_pdf_out_dir = Path(manual_output) / base_name
#             print(f"  Check    :  FAIL(S) detected → 05_Manual_Validation_Required")

#         per_pdf_out_dir.mkdir(parents=True, exist_ok=True)

#         # ── Write final JSON ──────────────────────────────────────────────────
#         out_path = per_pdf_out_dir / f"{stem}.json"
#         out_path.write_text(
#             json.dumps(result, indent=2, ensure_ascii=False), encoding="utf-8"
#         )

#         # ── Write final CSV ───────────────────────────────────────────────────
#         csv_path = per_pdf_out_dir / f"{stem}.csv"
#         run_json_to_csv_pipeline(json_path=str(out_path), csv_path=str(csv_path))

#         # ── Copy original raw PDF into the output folder ──────────────────────
#         raw_pdf = find_raw_pdf(raw_folder, base_name)
#         if raw_pdf:
#             dest_pdf = per_pdf_out_dir / raw_pdf.name
#             shutil.copy2(str(raw_pdf), str(dest_pdf))
#             print(f"  Raw PDF  : {raw_pdf.name} copied ✓")
#         else:
#             print(f"  Raw PDF  : [WARN] original PDF not found for '{base_name}' — skipped")

#         print(f"  JSON     : {out_path}")
#         print(f"  CSV      : {csv_path}")

#         # ── Append token log ──────────────────────────────────────────────────
#         with open(log_file, "a", newline="", encoding="utf-8") as f:
#             writer = csv.writer(f)
#             writer.writerow([
#                 fname, provider, model,
#                 usage["prompt_tokens"],
#                 usage["completion_tokens"],
#                 usage["total_tokens"],
#                 round(total_cost, 6),
#             ])

#         # ── Summary ───────────────────────────────────────
#         meta       = result.get("Metadata", {})
#         total_rows = sum(len(v) for v in result.get("Sections", {}).values())
#         print(f"  Issuer   : {meta.get('Issuer Name', 'N/A')}")
#         print(f"  FYE      : {meta.get('FYE', 'N/A')}")
#         print(f"  Rows     : {total_rows}")

#         ok_count += 1

#     print(f"\n  [ok] Normalization complete: {ok_count}/{total} file(s) succeeded.")
#     print(f"  [INFO] Token log saved at: {log_file}")
# ─────────────────────────────────────────────────────────────────────────────
# PER-ROW EXTRACTION + NORMALIZATION + VALIDATION + DB UPDATE
# ─────────────────────────────────────────────────────────────────────────────

def process_one_row(conn, row, coa_text):
    """
    Full extraction pipeline for a single DB row (one processing_id / one PDF).

    Steps:
      1. mark ExtractionFlag='p', DataValidationFlag='p'
      2. split source PDF → SNP/SOA sub-PDFs
      3. LLM-normalize each → per-statement JSON + CSV (in a temp work dir)
      4. DataValidationStatus = AND of each statement's all_pass
      5. merge CSVs → one xlsx (sheet per statement), routed to validated/manual
      6. update OutputPath, ExtractionStatus, DataValidationStatus,
         flags='c', CompletionStatus
    """
    processing_id   = row.ProcessingId
    processing_code = row.ProcessingCode
    src_pdf         = row.PdfFilePath

    base_name = Path(src_pdf).stem  # e.g. "{guid}_AR"
    print(f"\n=== ProcessingId {processing_id} | {row.IssuerName} ===")
    print(f"  Source PDF : {src_pdf}")

    update_extraction_flag(conn, processing_id, 'p')
    update_data_validation(conn, processing_id, flag='p')

    extraction_status = 0
    data_validation_status = 0

    try:
        if not src_pdf or not Path(src_pdf).exists():
            update_remarks(conn, processing_id, "Extraction: source PDF not found")
            print(f"  [WARN] Source PDF missing.")
            extraction_status = 0
            data_validation_status = 0
            return  # finally-block finalizes the row

        # ── 1. Split into SNP/SOA sub-PDFs ───────────────────────────────────
        extracted_pdfs = run_extraction_for_pdf(src_pdf)
        if not extracted_pdfs:
            update_remarks(conn, processing_id, "Extraction: no SNP/SOA pages found")
            print("  [WARN] No statements extracted.")
            return

        # ── 2-3. Normalize each statement → temp CSVs ────────────────────────
        # Work dir for per-statement json/csv before merge
        work_dir = Path(src_pdf).parent / f"_work_{base_name}"
        work_dir.mkdir(parents=True, exist_ok=True)

        csv_by_sheet = {}        # {"SNP": csv_path, "SOA": csv_path}
        statement_pass = []      # per-statement all_pass bools
        produced_any = False

        for pdf_path in extracted_pdfs:
            fname = Path(pdf_path).name
            stem  = Path(pdf_path).stem
            type_key = detect_type(fname)
            if type_key is None:
                print(f"  [SKIP] Unknown type: {fname}")
                continue

            prompt_path = os.path.join(PROMPTS_FOLDER, SUFFIX_TO_PROMPT[type_key])
            if not Path(prompt_path).exists():
                print(f"  [SKIP] Prompt not found: {prompt_path}")
                continue

            sheet_name = type_key.lstrip("_")  # "SNP" / "SOA"
            print(f"  Normalizing {sheet_name} via {PROVIDER} ({MODEL}) ...")

            # ── Dispatch to provider ──
            if PROVIDER == "openai":
                response = normalize_one_openai(
                    pdf_path=pdf_path, prompt_path=prompt_path, coa_text=coa_text,
                    reporting_columns=REPORTING_COLUMNS, client=_OPENAI_CLIENT,
                    model=MODEL, max_tokens=MAX_TOKENS)
            else:
                response = normalize_one_gemini_wrapper(
                    pdf_path=pdf_path, prompt_path=prompt_path, coa_text=coa_text,
                    reporting_columns=REPORTING_COLUMNS, model=MODEL, max_tokens=MAX_TOKENS)

            result = response["data"]

            # per-statement all_pass (same rule as before)
            stmt_pass = all(
                str(item.get("Total Check Status", "")) in {
                    "", "PASS", "Skipped - no membership rule defined"}
                for section_items in result.get("Sections", {}).values()
                for item in (section_items if isinstance(section_items, list) else [])
                if str(item.get("COA Flag", "")).strip().upper() == "CP"
            )
            statement_pass.append(stmt_pass)

            # write json + csv into work_dir
            json_path = work_dir / f"{stem}.json"
            json_path.write_text(json.dumps(result, indent=2, ensure_ascii=False), encoding="utf-8")
            csv_path = work_dir / f"{stem}.csv"
            run_json_to_csv_pipeline(json_path=str(json_path), csv_path=str(csv_path))

            csv_by_sheet[sheet_name] = csv_path
            produced_any = True

        # ── 4. Overall verdict ───────────────────────────────────────────────
        if not produced_any:
            update_remarks(conn, processing_id, "Extraction: LLM produced no usable output")
            extraction_status = 0
            data_validation_status = 0
            return

        extraction_status = 1  # JSON+CSV produced
        data_validation_status = 1 if (statement_pass and all(statement_pass)) else 0

        # ── 5. Merge → one xlsx routed by verdict ────────────────────────────
        dest_base = Path(OUTPUT_BASE_PATH) if data_validation_status == 1 else Path(MANUAL_OUTPUT)
        out_dir = dest_base / base_name
        out_xlsx = out_dir / f"{base_name}.xlsx"
        merge_csvs_to_xlsx(csv_by_sheet, out_xlsx)
        print(f"  Output xlsx : {out_xlsx}  ({'PASS' if data_validation_status==1 else 'FAIL → manual'})")

        # ── 6. Persist OutputPath ────────────────────────────────────────────
        update_output_path(conn, processing_id, out_xlsx)

        # cleanup the work dir (keep only the merged xlsx)
        try:
            import shutil as _sh
            _sh.rmtree(work_dir, ignore_errors=True)
        except Exception:
            pass

    except (RuntimeError, ValueError) as e:
        err = str(e)
        print(f"  [ERROR] {err}")
        update_remarks(conn, processing_id, f"Extraction error: {err[:400]}")
        # Fatal API errors should abort the whole batch
        if any(k in err.lower() for k in (
            "quota", "exhausted", "429", "resource", "permission",
            "unauthenticated", "api key")):   # ← dropped "deadline", "timeout"
            # finalize this row first, then re-raise to stop the batch
            extraction_status = 0
            data_validation_status = 0
            # _finalize_row(conn, processing_id, extraction_status, data_validation_status)
            # cleanup_extracted_for_pdf(src_pdf)
            raise
        extraction_status = 0
        data_validation_status = 0

    finally:
        _finalize_row(conn, processing_id, extraction_status, data_validation_status)
        cleanup_extracted_for_pdf(src_pdf)


def _finalize_row(conn, processing_id, extraction_status, data_validation_status):
    """Write the four status columns + completion in one place."""
    update_extraction_status(conn, processing_id, extraction_status)
    update_data_validation(conn, processing_id, status=data_validation_status, flag='c')
    update_extraction_flag(conn, processing_id, 'c')
    completion = update_completion_status(conn, processing_id, extraction_status, data_validation_status)
    print(f"  DB → Extraction={extraction_status} DataValidation={data_validation_status} "
          f"Completion={completion}")

# ─────────────────────────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────────────────────────
# ─────────────────────────────────────────────────────────────────────────────
# DB ROW SELECTION
# ─────────────────────────────────────────────────────────────────────────────

SELECT_PENDING_SQL = """
    SELECT
        cm.CompanyId, ps.ProcessingId, ps.ProcessingCode,
        cm.IssuerName, cm.State, cm.Sector, cm.SubSector,
        cm.UEI, cm.EIN, ps.ProcessYear, ps.FyeDate, ps.PdfFilePath
    FROM TCompanyMaster cm
    JOIN TProcessStatus ps ON ps.CompanyId = cm.CompanyId
    WHERE cm.IsActive = 1
      AND ps.IsActive = 1
      AND ps.SourcingStatus = 1
      AND ps.SourcingFlag = 'c'
      AND ps.SourcingValidationStatus = 1
      AND ps.SourcingValidationFlag = 'c'
      AND (ps.ExtractionFlag = 'c' OR ps.ExtractionFlag IS NULL)
      AND (ps.ExtractionStatus = 0 OR ps.ExtractionStatus IS NULL)
      AND (ps.CompletionStatus = 0 OR ps.CompletionStatus IS NULL)
      AND cm.ModuleId = 1
"""

# Provider client created once (OpenAI only; Gemini reads key per-call)
_OPENAI_CLIENT = None


def _init_provider():
    """Validate API keys and (for OpenAI) build the shared client."""
    global _OPENAI_CLIENT
    if PROVIDER == "openai":
        key = os.environ.get("OPENAI_API_KEY")
        if not key:
            sys.exit("[ERROR] OPENAI_API_KEY not set in .env")
        _OPENAI_CLIENT = OpenAI(api_key=key)
    else:
        if not os.environ.get("GEMINI_API_KEY"):
            sys.exit("[ERROR] GEMINI_API_KEY not set in .env")


def main(processing_id: int | None = None) -> None:
    """
    Extraction pipeline driver.

    - Standalone batch mode: main() processes ALL pending rows from the DB.
    - Dagster single mode:   main(processing_id=N) processes just that row.
    """
    _init_provider()

    print("=" * 62)
    print("  Financial Statement Extraction Pipeline (DB-driven)")
    print("=" * 62)
    print(f"  Provider : {PROVIDER.upper()}  |  Model: {MODEL}")
    print(f"  COA xlsx : {XLSX_PATH}")
    print(f"  Prompts  : {PROMPTS_FOLDER}")
    print(f"  Output   : {OUTPUT_BASE_PATH}")
    print(f"  Manual   : {MANUAL_OUTPUT}")

    if not Path(XLSX_PATH).exists():
        sys.exit(f"[ERROR] COA master not found: {XLSX_PATH}")
    if not Path(PROMPTS_FOLDER).exists():
        sys.exit(f"[ERROR] Prompts folder not found: {PROMPTS_FOLDER}")

    coa_text = load_xlsx_as_pipe_text(XLSX_PATH)
    print(f"  COA rows : {coa_text.count(chr(10)) + 1}")

    conn = get_db_connection()
    cur = conn.cursor()

    if processing_id is not None:
        cur.execute(SELECT_PENDING_SQL + " AND ps.ProcessingId = ?", processing_id)
    else:
        cur.execute(SELECT_PENDING_SQL)

    rows = cur.fetchall()
    print(f"  Pending  : {len(rows)} row(s)\n")

    if not rows:
        print("[DONE] Nothing to extract.")
        conn.close()
        return

    for row in rows:
        try:
            process_one_row(conn, row, coa_text)
        except SystemExit:
            raise
        except Exception as e:
            # process_one_row already finalized + re-raised on fatal API errors
            print(f"  [FATAL] Aborting batch: {e}")
            break

    conn.close()
    print("\n[DONE] Extraction pipeline complete.")


if __name__ == "__main__":
    # Optional: allow `python pipeline.py <processing_id>` for a single row
    pid = int(sys.argv[1]) if len(sys.argv) > 1 and sys.argv[1].isdigit() else None
    main(processing_id=pid)