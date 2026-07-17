# -*- coding: utf-8 -*-
"""
STEP 1:
- Load XLSM workbook, read headers, create maps.

STEP 2:
- Loop Daily Sourcing rows:
  - Open FAC advanced search
  - Apply filters (Audit Year, UEI, State, Entity type)
  - Search
  - Download "Download all" Excel

STEP 3:
- Open downloaded Excel -> General sheet
- Identify "reliant row" (scoring match)
- Extract and write to Daily Sourcing columns

STEP 4:
- Build PDF URL from Report ID
- Download PDF using browser download (expect_download)
- Save PDF with naming convention (similar to attached py)

STEP 5:
- Update PDF LINK column with pdf url
- Save XLSM after each row
"""

import re
from pathlib import Path
from typing import Dict, List, Optional, Any, Tuple
from datetime import datetime, date
import unicodedata
import pdfplumber
import os
import shutil
import subprocess
import pyodbc
from dotenv import load_dotenv

# Shared, reusable DB layer (see db.py). All DB access goes through this now
# instead of a per-file get_db_connection() + raw pyodbc cursors.
from db import Database, build_in_clause

# For auto-closing proof popup
try:
    import tkinter as tk
    from PIL import Image, ImageTk
    _TK_AVAILABLE = True
except Exception:
    _TK_AVAILABLE = False





# =========================================================
# DB CONNECTION
# =========================================================
# Connection handling now lives in db.py (get_db_connection() + Database).
# Use the Database class (imported at the top) as a context manager:
#     with Database() as db:
#         res = db.fetch_all(sql, params)
#         db.update(sql, params)
# The old per-file get_db_connection() has been removed in favour of that layer.

# =========================================================
# ADD THESE IMPORTS TO THE TOP OF YOUR EXISTING SCRIPT
# =========================================================
import pandas as pd
import threading
from filelock import FileLock     # pip install filelock


# =========================================================
# VALIDATION PARQUET FILE PATH (global variable)
# =========================================================

VALIDATION_PARQUET_PATH = Path(r"C:\Test Code\Public Finance\999_Log_Trackers\validation_results.parquet")
VALIDATION_LOCK_PATH    = VALIDATION_PARQUET_PATH.with_suffix(".parquet.lock")

# --- All columns: "Id" (physical TProcessStatus.Id, unique key) first, then
#     the 17 columns of the validation.xlsx template ---
VALIDATION_COLUMNS = [
    "Id",
    "ProcessingID",
    "ISSUER NAME",
    "UEI",
    "Name",
    "FYE (mm/yy)",
    "State",
    "Availability of Audit Report",
    "Audit Opinion",
    "Audit Report FYE with Report FYE",
    "Auditor's Signature",
    "Statement of Net Position with FYE",
    "Statement of Activities with FYE",
    "Balance Sheet with FYE",
    "Statement of Revenues, Expenditures and changes in Fund Balances with FYE",
    "Statement of Net Position of Proprietary Funds with FYE",
    "Statements of Revenues, Expenses And Changes In Net Position with FYE",
    "Statement of Cash Receipts and Disbursements with FYE",
]

# Concurrency primitives (same pattern as log parquet)
_VALIDATION_THREAD_LOCK = threading.Lock()
_VALIDATION_FILE_LOCK   = FileLock(str(VALIDATION_LOCK_PATH), timeout=60)


# =========================================================
# VALIDATION PARQUET HELPERS (THREAD + PROCESS SAFE)
# =========================================================

def _ensure_validation_file():
    """
    Create the validation parquet file with the correct schema if it doesn't exist.
    Must be called INSIDE the lock.
    """
    VALIDATION_PARQUET_PATH.parent.mkdir(parents=True, exist_ok=True)

    if not VALIDATION_PARQUET_PATH.exists():
        empty_df = pd.DataFrame(columns=VALIDATION_COLUMNS)
        # Id and ProcessingID are ints, all other columns are strings
        empty_df = empty_df.astype({col: "object" for col in VALIDATION_COLUMNS})
        empty_df["Id"] = empty_df["Id"].astype("Int64")
        empty_df["ProcessingID"] = empty_df["ProcessingID"].astype("Int64")
        empty_df.to_parquet(VALIDATION_PARQUET_PATH, index=False)


def write_validation_row(row_id: int, processing_id: int, validation_data: Dict[str, Any]):
    """
    Insert or update a validation row in the parquet file.
    Thread-safe AND process-safe.

    Parameters:
      row_id          → unique key (Id column) = physical TProcessStatus.Id.
                        ProcessingId is NON-unique (merge-siblings share it), so the
                        report is keyed by the physical row Id to avoid collisions.
      processing_id   → deliverable grouping key, stored in ProcessingID for readability.
      validation_data → dict where keys are column names from VALIDATION_COLUMNS
                        (you can pass ANY subset — missing keys stay NULL/unchanged)

    Behavior:
      - If Id does not exist  → INSERT new row
      - If Id already exists  → UPDATE only the provided columns
                                (other column values are preserved)

    Usage examples:

      # Initial write at start of validation (name + uei)
      write_validation_row(row_id, processing_id, {
          "ISSUER NAME": "CITY OF MOUNTAIN VIEW",
          "UEI": "JJZLKAS3G111",
      })

      # After Name check passes
      write_validation_row(row_id, processing_id, {"Name": "PASS"})

      # Mark all validations as failed (e.g., PDF Not Found case)
      write_validation_row(row_id, processing_id, {
          "Name": "Fail",
          "FYE (mm/yy)": "Fail",
          # ... etc for all 14 validation columns
      })
    """
    # Build the new row with only the provided columns (rest are None)
    new_row_dict = {col: None for col in VALIDATION_COLUMNS}
    new_row_dict["Id"] = int(row_id)
    new_row_dict["ProcessingID"] = int(processing_id)

    for key, value in (validation_data or {}).items():
        if key in VALIDATION_COLUMNS:
            new_row_dict[key] = value
        # silently ignore unknown keys to avoid breaking callers

    # ---- THREAD LOCK ----
    with _VALIDATION_THREAD_LOCK:
        # ---- FILE LOCK ----
        try:
            with _VALIDATION_FILE_LOCK:
                _ensure_validation_file()

                try:
                    existing_df = pd.read_parquet(VALIDATION_PARQUET_PATH)
                except Exception:
                    existing_df = pd.DataFrame(columns=VALIDATION_COLUMNS)

                # Check if a row for this physical Id already exists
                mask = existing_df["Id"] == int(row_id)

                if mask.any():
                    # ---- UPDATE: only overwrite columns that the caller provided ----
                    row_idx = existing_df.index[mask][0]
                    for key, value in (validation_data or {}).items():
                        if key in VALIDATION_COLUMNS:
                            existing_df.at[row_idx, key] = value
                    updated_df = existing_df
                else:
                    # ---- INSERT new row ----
                    new_row_df = pd.DataFrame([new_row_dict])
                    updated_df = pd.concat([existing_df, new_row_df], ignore_index=True)

                updated_df.to_parquet(VALIDATION_PARQUET_PATH, index=False)

        except Exception as e:
            # Never crash the main workflow because of validation file write
            print(f"[VALIDATION-LOG-ERROR] Failed to write validation row for Id={row_id}: {e}")


# =========================================================
# PERFORMANCE HELPERS (reuse PDF handle + cache page text/words per row)
# =========================================================
from contextlib import contextmanager

_SHARED_PDF_PATH: Optional[str] = None
_SHARED_PDF_HANDLE = None

_PAGE_TEXT_CACHE: Dict[int, str] = {}
_PAGE_WORDS_CACHE: Dict[Tuple[int, bool], list] = {}

def clear_pdf_page_caches():
    """Clear per-page caches so memory does not grow across rows."""
    _PAGE_TEXT_CACHE.clear()
    _PAGE_WORDS_CACHE.clear()

def cached_page_text(page) -> str:
    """Return page text, cached per page object (avoids re-extracting many times)."""
    k = id(page)
    if k not in _PAGE_TEXT_CACHE:
        _PAGE_TEXT_CACHE[k] = page.extract_text() or ""
    return _PAGE_TEXT_CACHE[k]

def cached_page_words(page, use_text_flow: bool = True):
    """Return page words, cached per page object (avoids re-extracting many times)."""
    k = (id(page), bool(use_text_flow))
    if k not in _PAGE_WORDS_CACHE:
        _PAGE_WORDS_CACHE[k] = page.extract_words(use_text_flow=use_text_flow) or []
    return _PAGE_WORDS_CACHE[k]

@contextmanager
def open_pdf_maybe_shared(pdf_path):
    """Open a PDF unless a shared row-level PDF handle is already set for this path."""
    global _SHARED_PDF_PATH, _SHARED_PDF_HANDLE
    p = None if pdf_path is None else str(pdf_path)
    if _SHARED_PDF_HANDLE is not None and _SHARED_PDF_PATH == p:
        yield _SHARED_PDF_HANDLE
        return
    with pdfplumber.open(p) as pdf:
        yield pdf


PDF_NOT_FOUND_TEXT = "PDF Not Found"

ADV_URL = "https://app.fac.gov/dissemination/search/advanced/"

DOWNLOAD_DIR = Path(r"C:\Test Code\Public Finance\01_Metadata_update")
DOWNLOAD_DIR.mkdir(parents=True, exist_ok=True)

PDF_DIR = Path(r"C:\Test Code\Public Finance\02_Downloaded_Report")
PDF_DIR.mkdir(parents=True, exist_ok=True)

VALIDATED_PDF_DIR = Path(r"C:\Test Code\Public Finance\03_Validated_Report")
VALIDATED_PDF_DIR.mkdir(parents=True, exist_ok=True)

FAILED_VALIDATION_PDF_DIR = Path(r"C:\Test Code\Public Finance\05_Manual_validation_required")
FAILED_VALIDATION_PDF_DIR.mkdir(parents=True, exist_ok=True)
# =========================================================
# PRE-RUN CLEANUP + VALIDATION SETTINGS
# =========================================================

PROOF_ROOT = Path(r"C:\Test Code\Public Finance\999_Log_Trackers\Validation Proofs")
PROOF_DIR_NAME = PROOF_ROOT / "NAME"
PROOF_DIR_STATE = PROOF_ROOT / "STATE"
PROOF_DIR_FYE = PROOF_ROOT / "FYE"
PROOF_DIR_AUDIT = PROOF_ROOT / "AUDIT_REPORT"
PROOF_DIR_OPINION = PROOF_ROOT / "OPINION"
PROOF_DIR_AUDIT_FYE = PROOF_ROOT / "AUDIT_FYE"
PROOF_DIR_SIGNATURE = PROOF_ROOT / "SIGNATURE"
PROOF_DIR_NET_POSITION = PROOF_ROOT / "NET_POSITION"
PROOF_DIR_ACTIVITIES = PROOF_ROOT / "STATEMENT_OF_ACTIVITIES"
PROOF_DIR_BALANCE_SHEET = PROOF_ROOT / "BALANCE_SHEET"
PROOF_DIR_REV_EXP_FUND_BAL = PROOF_ROOT / "REV_EXP_CHG_FUND_BAL"
PROOF_DIR_REV_EXP_CHG_NET_POSITION = PROOF_ROOT / "REV_EXP_CHG_NET_POSITION"
PROOF_DIR_CASH_RECEIPTS_DISBURSEMENTS = PROOF_ROOT / "CASH_RECEIPTS_DISBURSEMENTS"


PROOF_DIR_NAME.mkdir(parents=True, exist_ok=True)
PROOF_DIR_STATE.mkdir(parents=True, exist_ok=True)
PROOF_DIR_FYE.mkdir(parents=True, exist_ok=True)
PROOF_DIR_AUDIT.mkdir(parents=True, exist_ok=True)
PROOF_DIR_OPINION.mkdir(parents=True, exist_ok=True)
PROOF_DIR_AUDIT_FYE.mkdir(parents=True, exist_ok=True)
PROOF_DIR_SIGNATURE.mkdir(parents=True, exist_ok=True)
PROOF_DIR_NET_POSITION.mkdir(parents=True, exist_ok=True)
PROOF_DIR_ACTIVITIES.mkdir(parents=True, exist_ok=True)
PROOF_DIR_BALANCE_SHEET.mkdir(parents=True, exist_ok=True)
PROOF_DIR_REV_EXP_FUND_BAL.mkdir(parents=True, exist_ok=True)
PROOF_DIR_REV_EXP_CHG_NET_POSITION.mkdir(parents=True, exist_ok=True)
PROOF_DIR_CASH_RECEIPTS_DISBURSEMENTS.mkdir(parents=True, exist_ok=True)


# Always show a snippet popup for every row (PASS/FAIL/NOT EXTRACTABLE)
SHOW_PROOF_POPUP = False
# Popup auto closes after N seconds (set 0 to disable popup)
PROOF_POPUP_SECONDS = 1

# If popup is not visible (or Tk is unavailable), reveal/open the saved proof file
REVEAL_PROOF_FILES = False

# 'explorer' = open File Explorer selecting the PNG (most reliable)
# 'viewer'   = open default image viewer
REVEAL_METHOD = "explorer"

# Proof image resolution
PROOF_IMAGE_RESOLUTION = 120  # 120–200 good; higher = clearer but slower

# Highlight style (Yellow)
HIGHLIGHT_FILL_RGBA = (255, 255, 0, 90)
HIGHLIGHT_STROKE = None  # set to "yellow" if you want a border too

# Stop searching after reaching a Table of Contents page
TOC_STOP_ENABLED = True

HEADLESS = False          # set True for final run
SLOW_MO_MS = 150          # set 0 for faster run
UI_TIMEOUT_MS = 30000

CHROME_CHANNEL = "chrome"
IGNORE_HTTPS_ERRORS = True  # helps in SSL-inspected networks


# =========================================================
# STEP 1 HELPERS
# =========================================================


# =========================================================
# SOURCING (FAC SEARCH + DOWNLOADS + EXCEL UPDATE)
# =========================================================









# SOURCING: Turn any cell value into a clean trimmed string (blank-safe).
def normalize(v) -> str:
    """Return a trimmed string for any value (``None`` becomes an empty string)."""
    return "" if v is None else str(v).strip()


# SOURCING: Clean text so it can be safely used as a Windows file name.
def sanitize_filename(s: str, max_len: int = 140) -> str:
    """Sanitize ``s`` into a safe Windows filename, collapsing illegal characters and truncating to ``max_len``."""
    s = (s or "").strip()
    s = re.sub(r'[\\/:"*?<>|]+', "_", s)
    s = re.sub(r"\s+", " ", s).strip(" ._")
    if not s:
        s = "UNKNOWN"
    if len(s) > max_len:
        s = s[:max_len].rstrip(" ._")
    return s


US_STATE_CODES = {
    "Alabama":"AL","Alaska":"AK","Arizona":"AZ","Arkansas":"AR",
    "California":"CA","Colorado":"CO","Connecticut":"CT","Delaware":"DE",
    "Florida":"FL","Georgia":"GA","Hawaii":"HI","Idaho":"ID",
    "Illinois":"IL","Indiana":"IN","Iowa":"IA","Kansas":"KS",
    "Kentucky":"KY","Louisiana":"LA","Maine":"ME","Maryland":"MD",
    "Massachusetts":"MA","Michigan":"MI","Minnesota":"MN","Mississippi":"MS",
    "Missouri":"MO","Montana":"MT","Nebraska":"NE","Nevada":"NV",
    "New Hampshire":"NH","New Jersey":"NJ","New Mexico":"NM",
    "New York":"NY","North Carolina":"NC","North Dakota":"ND",
    "Ohio":"OH","Oklahoma":"OK","Oregon":"OR","Pennsylvania":"PA",
    "Rhode Island":"RI","South Carolina":"SC","South Dakota":"SD",
    "Tennessee":"TN","Texas":"TX","Utah":"UT","Vermont":"VT",
    "Virginia":"VA","Washington":"WA","West Virginia":"WV",
    "Wisconsin":"WI","Wyoming":"WY"
}


# =========================================================
# DOWNLOAD RETRY WRAPPERS
# =========================================================


# SOURCING: Remove bracketed text so opinion classification does not get confused by exceptions.
def _remove_parentheses(text: str) -> str:
    """Remove parenthetical parts like (except for ...) so they don't trigger 'Qualified' wrongly."""
    if not text:
        return ""
    return re.sub(r"\([^)]*\)", "", text)


# SOURCING: Get the first few meaningful lines of a page (used for heading detection).
def _top_non_empty_lines(page_text: str, max_lines: int = 5) -> List[str]:
    """
    Return the first `max_lines` non-empty lines from extracted page text.
    """
    if not page_text:
        return []
    lines = [ln.strip() for ln in (page_text or "").splitlines() if ln.strip()]
    return lines[:max_lines]

# SOURCING: Aggressively clean text to letters/numbers/spaces for robust contains matching.
def clean_for_contains_match(s: str) -> str:
    """Aggressive cleanup: keep only letters/numbers/spaces (lowercase)."""
    s = normalize_pdf_text(s).lower()
    s = re.sub(r"[^a-z0-9 ]+", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s


# SOURCING: Normalize PDF text so matching works even with unusual hyphens, quotes, and spacing.
def normalize_pdf_text(text: str) -> str:
    """Normalize extracted PDF text for robust matching."""
    if not text:
        return ""
    text = unicodedata.normalize("NFKD", text)
    for a, b in {"–": "-", "—": "-", "’": "'", " ": " "}.items():
        text = text.replace(a, b)
    text = re.sub(r"\s+", " ", text).strip()
    return text


# SOURCING: Parse the FY End Date from Excel into a usable date value.
def parse_fy_end_date(value) -> Optional[date]:
    """Parse FY end date from Excel (date/datetime or string like 30-06-2022)."""
    if value is None:
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value

    s = str(value).strip()
    if not s:
        return None

    for fmt in (
        "%d-%m-%Y", "%d/%m/%Y", "%m/%d/%Y", "%Y-%m-%d",
        "%d-%m-%y", "%d/%m/%y", "%m/%d/%y",
        "%d.%m.%Y", "%m.%d.%Y",
        "%B %d, %Y", "%B %d %Y", "%b %d, %Y", "%b %d %Y",
        "%d %B %Y", "%d %b %Y", "%d %B, %Y", "%d %b, %Y",
    ):
        try:
            return datetime.strptime(s, fmt).date()
        except Exception:
            pass

    return None

# SOURCING: Split text into sentences to help opinion extraction and classification.
def split_sentences(text: str) -> List[str]:
    """Same idea as friend's: split by sentence punctuation."""
    if not text:
        return []
    return [s.strip() for s in re.split(r"(?<=[.!?])\s+", text) if s.strip()]



# =========================================================
# VALIDATION (PDF CHECKS + PROOFS + DETAILED VALIDATION CHECK)
# =========================================================

# VALIDATION: Split text into cleaned tokens (words) for matching across PDFs.
def tokenize(s: str) -> list:
    """Tokenize ``s`` into a list of cleaned lowercase word tokens for robust text matching."""
    return [t for t in clean_for_contains_match(s).split() if t]


# VALIDATION: Find where a token sequence appears inside a larger token list (used for highlighting).
def find_token_window(page_tokens: list, target_tokens: list) -> tuple:
    """Locate the sub-sequence ``target_tokens`` inside ``page_tokens``.

Returns a ``(start, end)`` index tuple for an exact match, a high-overlap
fuzzy match (>=0.8) with matching endpoints, or ``(None, None)`` if not found.
    """
    n = len(page_tokens)
    m = len(target_tokens)
    if m == 0 or n == 0 or m > n:
        return (None, None)

    for i in range(0, n - m + 1):
        if page_tokens[i:i+m] == target_tokens:
            return (i, i+m)

    best = (0.0, None, None)
    target_set = set(target_tokens)
    for i in range(0, n - m + 1):
        window = page_tokens[i:i+m]
        if window[0] != target_tokens[0] or window[-1] != target_tokens[-1]:
            continue
        overlap = len(target_set.intersection(window))
        score = overlap / max(1, len(target_set))
        if score > best[0]:
            best = (score, i, i+m)

    if best[0] >= 0.8:
        return (best[1], best[2])

    return (None, None)


# VALIDATION: Detect Table of Contents pages so validations can stop/skip where needed.
def is_table_of_contents_page(page_text: str) -> bool:
    """Return ``True`` if the page text looks like a Table of Contents / index page."""
    t = normalize_pdf_text(page_text).lower()
    top_part = " ".join((page_text or "").splitlines()[:25]).lower()

    patterns = [
        r"table\s+of\s+contents?",
        r"table\s+of\s+content",
        r"contents",
        r"index",
        r"summary\s+of\s+contents",
        r"content\s+outline",
        r"contents\s+at\s+a\s+glance",
        r"contents\s+page",
        r"table\s+of\s+figures",
        r"table\s+of\s+tables",
    ]

    if re.search(patterns[0], t, re.I):
        return True

    for p in patterns[1:]:
        if re.search(p, top_part, re.I):
            return True

    if re.search(r"contents", top_part, re.I) and re.search(r"\.{3,}\s*\d+", top_part):
        return True

    return False

def _main_audit_report_signals(page_text: str) -> dict:
    """
    Detect whether page text looks like the MAIN financial-statement audit report.

    Returns:
        {
            "heading_hit": bool,
            "structure_hit": bool,
            "excluded_heading_hit": bool,
            "matched_heading_line": Optional[str],
        }

    IMPORTANT:
    - Heading detection is broader than before.
    - Structural detection helps pages like Seattle where extracted text may not show
      a clean standalone Independent Auditors' Report heading.
    - Exclusion applies only when internal-control / compliance wording appears
      in the heading line itself.
    """
    out = {
        "heading_hit": False,
        "structure_hit": False,
        "excluded_heading_hit": False,
        "matched_heading_line": None,
    }

    if not page_text:
        return out

    lines = [ln.strip() for ln in (page_text or "").splitlines() if ln.strip()]
    scan_lines = lines[:80]

    def _norm(s: str) -> str:
        s = normalize_pdf_text(s)
        s = s.replace("’", "'").replace("‘", "'")
        s = re.sub(r"\s+", " ", s).strip()
        s = s.rstrip(" .:-").strip()
        s = re.sub(r"^[^A-Za-z0-9]+", "", s)
        s = re.sub(r"[^A-Za-z0-9]+$", "", s)
        return s.strip()

    heading_res = [
        re.compile(r"^INDEPENDENT\s+AUDITOR(?:S)?\s*'?S?\s+REPORT(?:S)?$", re.I),
        re.compile(r"^REPORT\s+OF\s+INDEPENDENT\s+AUDITOR(?:S)?$", re.I),
        re.compile(r"^REPORT\s+OF\s+THE\s+INDEPENDENT\s+AUDITOR(?:S)?$", re.I),
        re.compile(r"^REPORT\s+OF\s+INDEPENDENT\s+CERTIFIED\s+PUBLIC\s+ACCOUNTANTS$", re.I),
        re.compile(r"^REPORT\s+OF\s+THE\s+INDEPENDENT\s+CERTIFIED\s+PUBLIC\s+ACCOUNTANTS$", re.I),
        re.compile(r"^INDEPENDENT\s+CERTIFIED\s+PUBLIC\s+ACCOUNTANTS\s*'?S?\s+REPORT$", re.I),
        re.compile(r"^INDEPENDENT\s+PUBLIC\s+ACCOUNTANTS\s*'?S?\s+REPORT$", re.I),
        re.compile(r"^REPORT\s+OF\s+INDEPENDENT\s+PUBLIC\s+ACCOUNTANTS$", re.I),
    ]

    excluded_heading_res = [
        re.compile(r"\binternal\s+control\b", re.I),
        re.compile(r"\bcompliance\b", re.I),
        re.compile(r"\bmajor\s+program\b", re.I),
        re.compile(r"\bsingle\s+audit\b", re.I),
        re.compile(r"\buniform\s+guidance\b", re.I),
        re.compile(r"\bgovernment\s+auditing\s+standards\b", re.I),
        re.compile(r"\bother\s+matters\b", re.I),
    ]

    # 1) Direct heading-line detection
    for ln in scan_lines:
        n = _norm(ln)
        if any(rx.search(n) for rx in excluded_heading_res):
            out["excluded_heading_hit"] = True
            continue
        if any(rx.match(n) for rx in heading_res):
            out["heading_hit"] = True
            out["matched_heading_line"] = ln.strip()
            break

    # 2) Structural detection for cases like Seattle
    full = normalize_pdf_text(page_text)

    has_opinion_heading = bool(re.search(r"\bopinion\b", full, re.I))
    has_basis_heading = bool(re.search(r"\bbasis\s+for\s+opinion\b|\bbasis\s+for\s+opinions\b", full, re.I))
    has_audit_scope = bool(
        re.search(r"\bwe\s+have\s+audited\b", full, re.I) or
        re.search(r"\bwe\s+audited\b", full, re.I) or
        re.search(r"\bthe\s+financial\s+statements\b", full, re.I) or
        re.search(r"\bthe\s+consolidated\s+financial\s+statements\b", full, re.I)
    )
    has_present_fairly = bool(
        re.search(r"\bin\s+our\s+opinion\b", full, re.I) or
        re.search(r"\bpresent\s+fairly\b", full, re.I) or
        re.search(r"\bfairly\s+presented\b", full, re.I) or
        re.search(r"\btrue\s+and\s+fair\b", full, re.I)
    )

    has_fin_stmt_report_heading = bool(
        re.search(r"\breport\s+on\s+the\s+financial\s+statements\b", full, re.I) or
        re.search(r"\breport\s+on\s+the\s+audit\s+of\s+the\s+consolidated\s+financial\s+statements\b", full, re.I)
    )

    if (has_fin_stmt_report_heading and has_opinion_heading and has_present_fairly) or \
       (has_opinion_heading and has_basis_heading and has_audit_scope and has_present_fairly):
        out["structure_hit"] = True

    return out

# VALIDATION: Detect whether a page is the Independent Auditor’s Report page based on its heading.
def is_independent_auditors_report_page(page_text: str) -> bool:
    """
    True if a page contains the MAIN financial-statement audit report.

    PASS if either:
    - a supported audit heading line is found, OR
    - the page has a strong audit-report structure (Opinion / Basis / audited financial statements / present fairly)

    Excludes internal-control/compliance pages only when such wording appears
    in the heading line itself.
    """
    sig = _main_audit_report_signals(page_text)
    if sig.get("heading_hit"):
        return True
    if sig.get("structure_hit"):
        return True
    return False

# VALIDATION: Find the bounding box of the audit report heading so it can be highlighted in proof images.
def find_audit_report_heading_box(page) -> Optional[dict]:
    """
    Returns bounding box of the MAIN audit report heading line if present.

    First tries direct heading-line match.
    If no heading line is found but the page has strong audit-report structure,
    falls back to highlighting one of:
      - Report on the financial statements
      - Report on the audit of the consolidated financial statements
      - Opinion
    """
    page_text = cached_page_text(page)
    if not page_text.strip():
        return None


    sig = _main_audit_report_signals(page_text)
    heading_line = sig.get("matched_heading_line")

    words = cached_page_words(page, use_text_flow=True)
    if not words:
        return None

    def _box_for_phrase(phrase: str) -> Optional[dict]:
        target_tokens = tokenize(phrase)
        if not target_tokens:
            return None

        page_tokens = []
        token_word_map = []

        for w in words:
            wtoks = tokenize(w.get("text", ""))
            for t in wtoks:
                page_tokens.append(t)
                token_word_map.append(w)

        s, e = find_token_window(page_tokens, target_tokens)
        if s is None or e is None:
            return None

        matched_words = token_word_map[s:e]
        return {
            "x0": min(w["x0"] for w in matched_words),
            "top": min(w["top"] for w in matched_words),
            "x1": max(w["x1"] for w in matched_words),
            "bottom": max(w["bottom"] for w in matched_words),
        }

    # 1) Direct heading line
    if heading_line:
        box = _box_for_phrase(heading_line)
        if box:
            return box

    # 2) Structural fallback
    fallback_phrases = [
        "Report on the financial statements",
        "Report on the audit of the consolidated financial statements",
        "Opinion",
    ]

    for ph in fallback_phrases:
        box = _box_for_phrase(ph)
        if box:
            return box

    return None
# VALIDATION: Open File Explorer (or image viewer) to show the saved proof image to the user.
def reveal_proof_file(png_path: Path):
    """Make the proof file visible to the user (Explorer select or open viewer)."""
    try:
        if REVEAL_METHOD.lower() == "explorer":
            # Opens File Explorer and highlights the file (very reliable)
            subprocess.Popen(["explorer", f"/select,{str(png_path)}"])
        else:
            os.startfile(str(png_path))
    except Exception:
        try:
            os.startfile(str(png_path))
        except Exception:
            pass


# VALIDATION: Show a quick image popup for proof (PASS only) and auto-close after a few seconds.
def show_proof_popup(png_path: Path, title: str, seconds: float = 3):
    """
    Shows image popup ONLY for PASS cases.
    If title contains FAIL / NOT EXTRACTABLE, it will NOT show any popup (and will NOT reveal files).
    """

    # -------------------------------------------------
    # ✅ DO NOT SHOW FAILED VALIDATION POPUPS
    # -------------------------------------------------
    t = (title or "").strip().lower()

    # Covers: "FAIL", "NOT EXTRACTABLE", and similar titles
    if (" fail" in t) or (t.endswith("fail")) or ("not extractable" in t):
        return

    # Existing guard
    if not SHOW_PROOF_POPUP or seconds <= 0:
        return

    # If Tk not available, only reveal for PASS (we already returned on FAIL)
    if not _TK_AVAILABLE:
        if REVEAL_PROOF_FILES:
            reveal_proof_file(png_path)
        return

    try:
        root = tk.Tk()
        root.title(title)

        # Always on top
        try:
            root.attributes("-topmost", True)
        except Exception:
            pass

        # Fullscreen / maximize (your current behavior)
        try:
            root.state("zoomed")  # Windows maximize
        except Exception:
            pass
        try:
            root.attributes("-fullscreen", True)
        except Exception:
            pass

        # Load image
        img = Image.open(png_path)

        # Fit to screen size but keep aspect ratio
        screen_w = root.winfo_screenwidth()
        screen_h = root.winfo_screenheight()

        max_w = max(200, screen_w - 40)
        max_h = max(200, screen_h - 80)

        w, h = img.size
        scale = min(max_w / w, max_h / h, 1.0)
        if scale < 1.0:
            img = img.resize((int(w * scale), int(h * scale)))

        photo = ImageTk.PhotoImage(img)
        label = tk.Label(root, image=photo, bg="black")
        label.image = photo
        label.pack(expand=True, fill="both")

        # ESC closes
        root.bind("<Escape>", lambda e: root.destroy())

        root.after(int(seconds * 1000), root.destroy)
        root.mainloop()

    except Exception:
        # If popup fails, reveal file (PASS only; FAIL already returned)
        if REVEAL_PROOF_FILES:
            reveal_proof_file(png_path)
        else:
            try:
                os.startfile(str(png_path))
            except Exception:
                pass

# VALIDATION: Create one proof image with multiple highlighted boxes (for bundled validations).
def save_combined_proof_png(
    pdf_path: Path,
    page_num: int,
    box_color_groups: List[Tuple[List[dict], Tuple[int, int, int, int]]],
    out_png: Path,
    stroke=None
) -> bool:
    """
    Re-render the same PDF page and draw multiple highlight boxes on one image
    with different colors per group.

    box_color_groups = [
        (boxes_for_validation_1, rgba_color_1),
        (boxes_for_validation_2, rgba_color_2),
        (boxes_for_validation_3, rgba_color_3),
    ]
    """
    try:
        with open_pdf_maybe_shared(pdf_path) as pdf:
            if not pdf.pages:
                return False
            if page_num < 1 or page_num > len(pdf.pages):
                return False

            page = pdf.pages[page_num - 1]
            im = page.to_image(resolution=PROOF_IMAGE_RESOLUTION)

            for boxes, color in (box_color_groups or []):
                for b in (boxes or []):
                    if not b:
                        continue
                    try:
                        im.draw_rect(b, stroke=stroke, fill=color)
                    except Exception:
                        pass

            out_png.parent.mkdir(parents=True, exist_ok=True)
            im.save(out_png)
            return True

    except Exception:
        return False

# VALIDATION: Show one combined proof popup when 3 validations pass on the same page, else show separately.
def show_threeway_bundle_or_individual(
    processing_id: int,
    pdf_path: Path,
    res1: dict,
    res2: dict,
    res3: dict,
    out_dir: Path,
    bundle_tag: str,
    color1: Tuple[int, int, int, int],
    color2: Tuple[int, int, int, int],
    color3: Tuple[int, int, int, int],
    popup_title: Optional[str] = None,
    show_popup = False
):
    """
    If res1+res2+res3 are PASS on the SAME page => ONE popup with 3 colors.
    Else => show individual popups using each proof_file.

    Each res dict must have:
      - status, page, proof_file, boxes, title
    """
    try:
        if not pdf_path or not Path(pdf_path).exists():
            return

        s1, s2, s3 = (res1 or {}).get("status"), (res2 or {}).get("status"), (res3 or {}).get("status")
        p1, p2, p3 = (res1 or {}).get("page"), (res2 or {}).get("page"), (res3 or {}).get("page")

        same_page = (
            s1 == "PASS" and s2 == "PASS" and s3 == "PASS" and
            p1 and p2 and p3 and int(p1) == int(p2) == int(p3)
        )

        if same_page:
            page_num = int(p1)

            boxes1 = (res1 or {}).get("boxes", []) or []
            boxes2 = (res2 or {}).get("boxes", []) or []
            boxes3 = (res3 or {}).get("boxes", []) or []

            combined_png = out_dir / f"Row{processing_id}_{bundle_tag}_Page{page_num}.png"

            ok = save_combined_proof_png(
                pdf_path=Path(pdf_path),
                page_num=page_num,
                box_color_groups=[
                    (boxes1, color1),
                    (boxes2, color2),
                    (boxes3, color3),
                ],
                out_png=combined_png,
                stroke=None
            )

            if ok:
                title = popup_title or f"Row {processing_id} | {bundle_tag.replace('_',' ')} (Single Page {page_num})"
                if(show_popup):
                    show_proof_popup(combined_png, title, seconds=PROOF_POPUP_SECONDS)
                return

        # Otherwise: show individual popups
        if show_popup:
            for res in (res1, res2, res3):
                if not res:
                    continue
                pf = res.get("proof_file")
                title = res.get("title", f"Row {processing_id}")
                if pf and Path(pf).exists():
                    show_proof_popup(Path(pf), title, seconds=PROOF_POPUP_SECONDS)

    except Exception:
        return

# VALIDATION: Save a bundled audit proof image (popup disabled).
def show_single_or_individual_popups(processing_id: int, pdf_path: Path, audit_res: dict, opinion_res: dict, audit_fye_res: dict):
    """
    Audit bundle:
      - Audit Report (Orange)
      - Audit Opinion (Cyan)
      - Audit Report FYE (Red)

    If all PASS on same page => ONE combined proof PNG with 3 colors.
    Else => individual proof PNGs.
    (Popups disabled; combined/individual PNGs still saved for manual review.)
    """
    AUDIT_REPORT_COLOR = (255, 165, 0, 90)   # Orange
    OPINION_COLOR      = (0, 255, 255, 90)   # Cyan
    AUDIT_FYE_COLOR    = (255, 0, 0, 90)     # Red

    show_threeway_bundle_or_individual(
        processing_id=processing_id,
        pdf_path=pdf_path,
        res1=audit_res,
        res2=opinion_res,
        res3=audit_fye_res,
        out_dir=PROOF_DIR_AUDIT,
        bundle_tag="AUDIT_REPORT_OPINION_AUDIT_FYE",
        color1=AUDIT_REPORT_COLOR,
        color2=OPINION_COLOR,
        color3=AUDIT_FYE_COLOR,
        popup_title=f"ProcessingId {processing_id} | Audit Report + Opinion + Audit FYE (Single Page)",
        show_popup=False
    )

# VALIDATION: Search the PDF for the entity name and save a proof image showing the match.
def visible_name_validation_with_proof(pdf_path: Path, issuer_name: str, proof_png_path: Path) -> dict:
    """Search pages until TOC is reached (TOC page included)."""
    issuer_name = (issuer_name or "").strip()
    if not issuer_name:
        return {"status": "FAIL", "reason": "Blank ISSUER NAME in Daily Sourcing", "page": None}

    try:
        with open_pdf_maybe_shared(pdf_path) as pdf:
            if not pdf.pages:
                return {"status": "NOT EXTRACTABLE", "reason": "PDF has no pages", "page": None}

            target_tokens = tokenize(issuer_name)
            last_im = None
            last_page_num = None

            for page_index, page in enumerate(pdf.pages):
                page_num = page_index + 1
                page_text = cached_page_text(page)
                last_page_num = page_num

                im = page.to_image(resolution=PROOF_IMAGE_RESOLUTION)
                last_im = im

                if page_text.strip():
                    words = cached_page_words(page, use_text_flow=True)
                    page_tokens = []
                    token_word_map = []
                    for w in words:
                        wtoks = tokenize(w.get('text', ''))
                        for t in wtoks:
                            page_tokens.append(t)
                            token_word_map.append(w)

                    s, e = find_token_window(page_tokens, target_tokens)
                    if s is not None and e is not None:
                        matched_words = token_word_map[s:e]
                        x0 = min(w['x0'] for w in matched_words)
                        top = min(w['top'] for w in matched_words)
                        x1 = max(w['x1'] for w in matched_words)
                        bottom = max(w['bottom'] for w in matched_words)
                        box = {"x0": x0, "top": top, "x1": x1, "bottom": bottom}
                        im.draw_rect(box, stroke=HIGHLIGHT_STROKE, fill=HIGHLIGHT_FILL_RGBA)
                        im.save(proof_png_path)
                        return {"status": "PASS", "reason": f"Issuer name found on page {page_num}", "page": page_num}

                    if TOC_STOP_ENABLED and is_table_of_contents_page(page_text):
                        # Stop after saving TOC page proof
                        im.save(proof_png_path)
                        return {"status": "FAIL", "reason": f"Issuer not found before/at TOC (stopped at page {page_num})", "page": page_num}
                else:
                    # No text; still allow continuing to next page
                    pass

            if last_im is not None:
                last_im.save(proof_png_path)
            return {"status": "FAIL", "reason": "Issuer name not found in processed pages", "page": last_page_num}

    except Exception as e:
        return {"status": "FAIL", "reason": f"Error reading PDF: {e}", "page": None}


# =========================================================
# STATE FALLBACK HELPERS
# =========================================================
# These helpers are used ONLY when the existing State validation fails.
# Existing State logic remains primary and unchanged.



# =========================================================
# FYE (mm/yy) VALIDATION HELPERS
# =========================================================

_MONTHS = {
    1: ("january", "jan"),
    2: ("february", "feb"),
    3: ("march", "mar"),
    4: ("april", "apr"),
    5: ("may", "may"),
    6: ("june", "jun"),
    7: ("july", "jul"),
    8: ("august", "aug"),
    9: ("september", "sep"),
    10: ("october", "oct"),
    11: ("november", "nov"),
    12: ("december", "dec"),
}

# VALIDATION: Generate many possible token formats for the same FY date (e.g., June 30 2025, 06/30/2025).
def build_fye_token_candidates(d: date) -> List[List[str]]:
    """Build token sequences representing the same date in PDFs across many formats."""
    day = str(d.day)
    day2 = f"{d.day:02d}"
    mon = d.month
    mon2 = f"{mon:02d}"
    year = str(d.year)
    year2 = f"{d.year % 100:02d}"
    full, abbr = _MONTHS[mon]

    cands = []

    # June 30 2025 / June 30, 2025
    cands += [[full, day, year], [abbr, day, year], [full, day2, year], [abbr, day2, year]]
    # 30 June 2025
    cands += [[day, full, year], [day, abbr, year], [day2, full, year], [day2, abbr, year]]
    # 06/30/2025 or 6/30/2025
    cands += [[mon2, day2, year], [str(mon), day2, year], [mon2, day, year], [str(mon), day, year]]
    # 30/06/2025
    cands += [[day2, mon2, year], [day2, str(mon), year], [day, mon2, year], [day, str(mon), year]]

    # two digit year variants
    cands += [[full, day, year2], [abbr, day, year2], [day, full, year2], [day, abbr, year2]]
    cands += [[mon2, day2, year2], [day2, mon2, year2], [str(mon), day, year2], [day, str(mon), year2]]

    # de-dup
    seen = set()
    out = []
    for c in cands:
        t = tuple(c)
        if t not in seen:
            seen.add(t)
            out.append(c)
    return out

# ---------------------------------------------------------
# FYE CONTEXT RULES (REQUIRED KEYWORDS BEFORE DATE)
# ---------------------------------------------------------

FYE_PREFIX_PATTERNS = [
    r"\bfiscal\s+year\s+ended\b",
    r"\bfiscal\s+year\s+ending\b",
    r"\bfor\s+the\s+fiscal\s+year\s+ended\b",
    r"\byear\s+ended\b",
    r"\byears\s+ended\b",
    r"\bfor\s+the\s+year\s+ended\b",
    r"\bfiscal\s+year\b",
    r"\bfiscal\s+period\b",
    r"\bperiod\s+ended\b",
    r"\bannual\s+period\s+ended\b",
]

# Optional words that might appear around the FYE line (not mandatory)
FYE_SOFT_PATTERNS = [
    r"\bfinancial\s+statements?\b",
    r"\bbasic\s+financial\s+statements?\b",
    r"\bannual\s+financial\s+report\b",
    r"\bacfr\b",
    r"\bcafr\b",
]

# VALIDATION: Confirm the FY date is preceded by context like "year ended" to avoid false matches.
def fye_context_ok(page_tokens: list, date_start_index: int) -> bool:
    """
    Returns True only if FYE keywords appear BEFORE the matched date.
    We look back a window of tokens so date must be 'followed after' the keyword.
    """
    lookback = 35  # you can tune this (25–50 works well)
    left = max(0, date_start_index - lookback)
    before_text = " ".join(page_tokens[left:date_start_index])

    # hard requirement: one of the prefix patterns must exist BEFORE the date
    for pat in FYE_PREFIX_PATTERNS:
        if re.search(pat, before_text, flags=re.I):
            return True

    # (Optional fallback): if no hard prefix found, we still allow if soft patterns exist
    # AND 'ended' exists somewhere before date (covers rare layouts)
    if re.search(r"\bended\b", before_text, flags=re.I):
        for pat in FYE_SOFT_PATTERNS:
            if re.search(pat, before_text, flags=re.I):
                return True

    return False

# VALIDATION: Find the FY date on a PDF page only when valid FY context appears before it.
def _find_fye_on_page_with_context(page, fye_date: date) -> Optional[Tuple[int, int, dict]]:
    """
    Searches for the FYE date on a single pdfplumber page:
    - tries many date formats (token candidates)
    - requires fye_context_ok(page_tokens, date_start_index) == True
    Returns:
      (start_token_index, end_token_index, box_dict) if found
      None if not found
    """
    words = cached_page_words(page, use_text_flow=True)

    page_tokens = []
    token_word_map = []
    for w in words:
        wtoks = tokenize(w.get("text", ""))
        for t in wtoks:
            page_tokens.append(t)
            token_word_map.append(w)

    for cand in build_fye_token_candidates(fye_date):
        s, e = find_token_window(page_tokens, cand)
        if s is not None and e is not None:
            # IMPORTANT: FYE must be after "year ended/fiscal year ended/..." etc
            if not fye_context_ok(page_tokens, s):
                continue

            matched_words = token_word_map[s:e]
            x0 = min(w["x0"] for w in matched_words)
            top = min(w["top"] for w in matched_words)
            x1 = max(w["x1"] for w in matched_words)
            bottom = max(w["bottom"] for w in matched_words)
            box = {"x0": x0, "top": top, "x1": x1, "bottom": bottom}
            return (s, e, box)

    return None

# VALIDATION: Validate FY End Date in the PDF and save a proof image highlighting the date.
def visible_fye_validation_with_proof(
    pdf_path: Path,
    fy_end_value,
    name_found_page: int,
    proof_png_path: Path
) -> dict:
    """
    FYE Validation (improved to avoid cover-page false FAILs)

    Strategy:
      1) Try STRICT context match on the Name-found page (existing rule).
      2) If not found, try HEADER/TOP-LINES match (no context) on:
            - Name-found page
            - Page 1 (cover page)
            - Page 2 and 3 (TOC/intro often repeats the date)
      3) If still not found, fallback to Independent Auditor's Report page with strict context.
    """
    d = parse_fy_end_date(fy_end_value)
    if not d:
        return {"status": "FAIL", "reason": "Invalid/blank fy_end_date from General sheet", "page": name_found_page}

    try:
        with open_pdf_maybe_shared(pdf_path) as pdf:
            if not pdf.pages:
                return {"status": "NOT EXTRACTABLE", "reason": "PDF has no pages", "page": name_found_page}

            total_pages = len(pdf.pages)

            # Defensive: if name_found_page is missing/wrong, fall back to page 1
            if not name_found_page or name_found_page < 1 or name_found_page > total_pages:
                name_found_page = 1

            # Candidate pages to check (unique, ordered)
            candidate_pages = []
            for pno in [name_found_page, 1, 2, 3]:
                if 1 <= pno <= total_pages and pno not in candidate_pages:
                    candidate_pages.append(pno)

            # ----------------------------------------------------
            # A) Check candidate pages:
            #    1) strict context
            #    2) top header lines (no context)
            #    3) anywhere (no context)
            # ----------------------------------------------------
            for pno in candidate_pages:
                pg = pdf.pages[pno - 1]
                pg_text = pg.extract_text() or ""
                if not pg_text.strip():
                    continue

                # NOTE:
                # - We skip TOC pages for STRICT context attempt,
                #   but we still allow header-based date match on TOC because TOC often repeats date.
                is_toc = bool(TOC_STOP_ENABLED and is_table_of_contents_page(pg_text))

                im = pg.to_image(resolution=PROOF_IMAGE_RESOLUTION)

                # 1) STRICT context match (date must be after "year ended" etc.)
                box = None
                if not is_toc:
                    try:
                        box = find_fye_box_on_page_with_context(pg, fy_end_value)
                    except Exception:
                        box = None

                if box:
                    im.draw_rect(box, stroke=HIGHLIGHT_STROKE, fill=HIGHLIGHT_FILL_RGBA)
                    im.save(proof_png_path)
                    return {
                        "status": "PASS",
                        "reason": f"FYE found with strict context on page {pno}",
                        "page": pno
                    }

                # 2) HEADER/TOP-LINES fallback (no strict context)
                # Use more lines to survive margin/vertical text cases
                try:
                    box = find_fye_box_on_page_top_lines(pg, fy_end_value, max_lines=20)
                except Exception:
                    box = None

                if box:
                    im.draw_rect(box, stroke=HIGHLIGHT_STROKE, fill=HIGHLIGHT_FILL_RGBA)
                    im.save(proof_png_path)
                    return {
                        "status": "PASS",
                        "reason": f"FYE found in top header area (no context) on page {pno}",
                        "page": pno
                    }

                # 3) Anywhere fallback (no strict context)
                try:
                    box = find_fye_box_on_page_anywhere(pg, fy_end_value)
                except Exception:
                    box = None

                if box:
                    im.draw_rect(box, stroke=HIGHLIGHT_STROKE, fill=HIGHLIGHT_FILL_RGBA)
                    im.save(proof_png_path)
                    return {
                        "status": "PASS",
                        "reason": f"FYE found anywhere on page {pno} (no context)",
                        "page": pno
                    }

            # ----------------------------------------------------
            # B) Fallback: Independent Auditor's Report page (strict context)
            # ----------------------------------------------------
            for idx, pg in enumerate(pdf.pages):
                page_num = idx + 1
                pg_text = pg.extract_text() or ""
                if not pg_text.strip():
                    continue

                if TOC_STOP_ENABLED and is_table_of_contents_page(pg_text):
                    continue

                if not is_independent_auditors_report_page(pg_text):
                    continue

                hit = None
                try:
                    hit = _find_fye_on_page_with_context(pg, d)
                except Exception:
                    hit = None

                auditor_im = pg.to_image(resolution=PROOF_IMAGE_RESOLUTION)

                if hit:
                    _, _, box = hit
                    auditor_im.draw_rect(box, stroke=HIGHLIGHT_STROKE, fill=HIGHLIGHT_FILL_RGBA)
                    auditor_im.save(proof_png_path)
                    return {
                        "status": "PASS",
                        "reason": f"FYE found on Independent Auditor's Report page {page_num}",
                        "page": page_num
                    }

                auditor_im.save(proof_png_path)
                return {
                    "status": "FAIL",
                    "reason": f"Independent Auditor's Report found (page {page_num}) but FYE not found (strict context)",
                    "page": page_num
                }

            # If no auditor report page found:
            # Save the name page as proof and fail
            fallback_pg = pdf.pages[name_found_page - 1]
            fallback_im = fallback_pg.to_image(resolution=PROOF_IMAGE_RESOLUTION)
            fallback_im.save(proof_png_path)

            return {
                "status": "FAIL",
                "reason": f"FYE not found on candidate pages {candidate_pages} and no Auditor Report page matched",
                "page": name_found_page
            }

    except Exception as e:
        return {"status": "FAIL", "reason": f"Error reading PDF for FYE validation: {e}", "page": name_found_page}
    

# VALIDATION: Run FYE PASS/FAIL check and generate proof image. (status returned, parent batches parquet write)
def write_fye_validation(
    processing_id: int,
    fy_end_value,
    pdf_path: Optional[Path],
    name_result: dict,
    show_popup: bool = True
) -> dict:
    """
    FYE validation
    - Saves proof PNG to PROOF_DIR_FYE
    - Returns dict with: status, page, proof_file, boxes, title
    - Does NOT write to parquet (parent run_validations_for_row batches the write)
    """
    if not pdf_path or not Path(pdf_path).exists():
        return {
            "status": "FAIL",
            "page": None,
            "proof_file": None,
            "boxes": [],
            "title": f"ProcessingId {processing_id} | FYE Check = FAIL (PDF missing)"
        }

    name_page = (name_result or {}).get("page") or 1

    d = parse_fy_end_date(fy_end_value)
    tag = d.strftime("%Y%m%d") if d else sanitize_filename(str(fy_end_value), 20)
    proof_file = PROOF_DIR_FYE / f"Row{processing_id}_FYE_{tag}.png"

    res = visible_fye_validation_with_proof(Path(pdf_path), fy_end_value, int(name_page), proof_file)
    status = res.get("status", "FAIL")
    page_no = res.get("page", name_page)

    # Best-effort box for bundling: re-find date on the returned page using fallbacks
    boxes = []
    if status == "PASS" and page_no:
        try:
            with open_pdf_maybe_shared(pdf_path) as pdf:
                if 1 <= int(page_no) <= len(pdf.pages):
                    pg = pdf.pages[int(page_no) - 1]

                    b = None

                    # 1) strict context
                    try:
                        b = find_fye_box_on_page_with_context(pg, fy_end_value)
                    except Exception:
                        b = None

                    # 2) top-lines
                    if not b:
                        try:
                            b = find_fye_box_on_page_top_lines(pg, fy_end_value, max_lines=20)
                        except Exception:
                            b = None

                    # 3) anywhere
                    if not b:
                        try:
                            b = find_fye_box_on_page_anywhere(pg, fy_end_value)
                        except Exception:
                            b = None

                    if b:
                        boxes.append(b)

        except Exception:
            pass

    title = f"ProcessingId {processing_id} | FYE Check = {status}" + (f" | Page {page_no}" if page_no else "")

    if show_popup and proof_file.exists():
        show_proof_popup(proof_file, title, seconds=PROOF_POPUP_SECONDS)

    return {
        "status": status,
        "page": page_no,
        "proof_file": proof_file,
        "boxes": boxes,
        "title": title
    }

# =========================================================
# AVAILABILITY OF AUDIT REPORT (Detailed Validation Check - Column F)
# =========================================================

# VALIDATION: Find a specific phrase on a page and return its bounding box (for highlighting).
def find_phrase_box_on_page(page, phrase: str) -> Optional[dict]:
    """Find issuer name phrase on a given page; return bbox dict if found else None."""
    phrase = (phrase or "").strip()
    target_tokens = tokenize(phrase)
    if not target_tokens:
        return None

    words = cached_page_words(page, use_text_flow=True)
    page_tokens = []
    token_word_map = []
    for w in words:
        wtoks = tokenize(w.get("text", ""))
        for t in wtoks:
            page_tokens.append(t)
            token_word_map.append(w)

    s, e = find_token_window(page_tokens, target_tokens)
    if s is None or e is None:
        return None

    matched_words = token_word_map[s:e]
    x0 = min(w["x0"] for w in matched_words)
    top = min(w["top"] for w in matched_words)
    x1 = max(w["x1"] for w in matched_words)
    bottom = max(w["bottom"] for w in matched_words)
    return {"x0": x0, "top": top, "x1": x1, "bottom": bottom}


# VALIDATION: Find the FY End Date on a page using strict "year ended" style context rules.
def find_fye_box_on_page_with_context(page, fy_end_value) -> Optional[dict]:
    """
    Find FY end date on a page using your strict context rule (date must be AFTER
    'year ended / fiscal year ended / years ended / ...').
    Returns bbox dict if found else None.
    """
    d = parse_fy_end_date(fy_end_value)
    if not d:
        return None

    # If your helper _find_fye_on_page_with_context exists, use it (best)
    if "_find_fye_on_page_with_context" in globals():
        hit = _find_fye_on_page_with_context(page, d)
        if hit:
            _, _, box = hit
            return box

    # Otherwise fall back to token window + context check
    words = cached_page_words(page, use_text_flow=True)
    page_tokens = []
    token_word_map = []
    for w in words:
        wtoks = tokenize(w.get("text", ""))
        for t in wtoks:
            page_tokens.append(t)
            token_word_map.append(w)

    for cand in build_fye_token_candidates(d):
        s, e = find_token_window(page_tokens, cand)
        if s is None or e is None:
            continue
        if "fye_context_ok" in globals() and not fye_context_ok(page_tokens, s):
            continue

        matched_words = token_word_map[s:e]
        x0 = min(w["x0"] for w in matched_words)
        top = min(w["top"] for w in matched_words)
        x1 = max(w["x1"] for w in matched_words)
        bottom = max(w["bottom"] for w in matched_words)
        return {"x0": x0, "top": top, "x1": x1, "bottom": bottom}

    return None

# =========================================================
# AUDIT OPINION VALIDATION HELPERS (Column G)
# =========================================================


# VALIDATION: Pick the best opinion sentence (prefer "in our opinion" if present).
def find_opinion_sentence(opinion_text: str) -> Optional[str]:
    """
    Priority:
      1) Sentence containing 'in our opinion'
      2) Otherwise, fallback to friend-style patterns
    """
    if not opinion_text:
        return None

    sentences = split_sentences(opinion_text)

    # ✅ Priority 1: "In our opinion"
    for s in sentences:
        if re.search(r"\bin\s+our\s+opinion\b", s, re.I):
            return s.strip()

    # Fallback: friend-style detection
    opinion_sentence_patterns = [
        r"\bpresent\s+fairly\b",
        r"\bfairly\s+presented\b",
        r"\btrue\s+and\s+fair\b",
        r"\bunmodified\s+opinion\b",
        r"\bqualified\s+opinion\b",
        r"\badverse\s+opinion\b",
        r"\bdisclaimer\s+of\s+opinion\b",
        r"\bdo\s+not\s+express\s+an\s+opinion\b",
        r"\bexcept\s+for\b",
    ]

    for sentence in sentences:
        if any(re.search(pat, sentence, re.I) for pat in opinion_sentence_patterns):
            return sentence.strip()

    return None

# VALIDATION: Classify audit opinion as unqualified/qualified/adverse/disclaimer based on keywords.
def classify_opinion(opinion_scope_text: str, opinion_sentence: Optional[str] = None) -> Optional[str]:
    """
    Classify based on opinion sentence FIRST (if available).
    If not available, fallback to full opinion scope.

    'except for' => Qualified only if it looks like qualification text
    and is NOT inside parentheses.
    """
    if not opinion_scope_text and not opinion_sentence:
        return None

    # ✅ Prefer the 'In our opinion' sentence
    scope = opinion_sentence.strip() if opinion_sentence else opinion_scope_text.strip()

    # Work on a version without parentheses for "except for" checks
    scope_no_paren = _remove_parentheses(scope)

    # Strong explicit labels (highest confidence)
    if re.search(r"\badverse\s+opinion\b", scope, re.I):
        return "Adverse"
    if re.search(r"\bdisclaimer\s+of\s+opinion\b|\bdo\s+not\s+express\s+an\s+opinion\b", scope, re.I):
        return "Disclaimer"
    if re.search(r"\bqualified\s+opinion\b", scope, re.I):
        return "Qualified"
    if re.search(r"\bunmodified\s+opinion\b", scope, re.I):
        return "Unmodified"

    # ✅ Qualified via "except for" ONLY if it looks like real qualification language
    # Examples of real qualification: "except for the effects of..." / "except for the possible effects of..."
    if re.search(r"\bexcept\s+for\b", scope_no_paren, re.I):
        if re.search(r"\bexcept\s+for\b.*\b(effects|possible\s+effects)\b", scope_no_paren, re.I):
            return "Qualified"
        # If it's just "(except for ...)" parenthetical or a date note, we DO NOT call it qualified.

    # Positive unmodified indicators
    if re.search(r"\bpresent\s+fairly\b", scope, re.I):
        return "Present Fairly"
    if re.search(r"\bfairly\s+presented\b", scope, re.I):
        return "Fairly Presented"
    if re.search(r"\btrue\s+and\s+fair\b", scope, re.I):
        return "True and Fair"

    return None

def find_opinion_anchor_in_pdf(pdf, max_scan_pages: int = 120) -> Optional[dict]:
    """
    Opinion-anchored audit discovery (NO TOC
    Finds the FIRST page in early PDF that:
      - contains a strong opinion trigger phrase
      - produces a non-empty classify_opinion() label
      - has a MAIN audit report on:
           same page OR previous 1-2 pages
    """
    if not pdf or not getattr(pdf, "pages", None):
        return None

    limit = min(max_scan_pages, len(pdf.pages))

    opinion_res = [
        re.compile(r"\bin\s+our\s+opinion\b", re.I),
        re.compile(r"\bin\s+my\s+opinion\b", re.I),
        re.compile(r"\bpresent\s+fairly\b", re.I),
        re.compile(r"\bfairly\s+presented\b", re.I),
        re.compile(r"\btrue\s+and\s+fair\b", re.I),
        re.compile(r"\bgive\s+a\s+true\s+and\s+fair\s+view\b", re.I),
        re.compile(r"\bthe\s+financial\s+statements\s+referred\s+to\s+above\s+present\b", re.I),
        re.compile(r"\bthe\s+consolidated\s+financial\s+statements\b", re.I),
    ]

    for idx in range(limit):
        page = pdf.pages[idx]
        txt = cached_page_text(page) or ""
        if not txt.strip():
            continue

        if not any(rx.search(txt) for rx in opinion_res):
            continue

        op_sentence = find_opinion_sentence(txt) or ""
        label = classify_opinion(txt, opinion_sentence=op_sentence)
        if not label:
            continue

        heading_page = None

        try:
            if is_independent_auditors_report_page(txt):
                heading_page = idx + 1
        except Exception:
            heading_page = None

        if heading_page is None:
            for back in (1, 2):
                if idx - back < 0:
                    break

                prev_txt = cached_page_text(pdf.pages[idx - back]) or ""
                if not prev_txt.strip():
                    continue

                try:
                    if is_independent_auditors_report_page(prev_txt):
                        heading_page = (idx - back) + 1
                        break
                except Exception:
                    pass

        if heading_page is None:
            continue

        return {
            "opinion_page": idx + 1,
            "heading_page": heading_page,
            "label": label,
            "opinion_sentence": op_sentence
        }

    return None

def build_audit_anchor_window_pages(
    pdf,
    heading_page_num: int,
    opinion_page_num: int,
    max_total_pages: int = 20,
    post_opinion_pages: int = 3
) -> List[Tuple[int, Any]]:
    """
    Build a compact audit window from the discovered anchor.
    NO TOC skipping.

    Window starts at heading_page_num
    Window ends at the greater of:
      - opinion_page_num + post_opinion_pages
      - heading_page_num + max_total_pages - 1
    but never beyond PDF length.
    """
    if not pdf or not getattr(pdf, "pages", None):
        return []

    n = len(pdf.pages)
    if heading_page_num < 1:
        heading_page_num = 1
    if opinion_page_num < 1:
        opinion_page_num = 1

    start_idx = heading_page_num - 1
    target_end = max(opinion_page_num + post_opinion_pages, heading_page_num + max_total_pages - 1)
    end_idx = min(target_end - 1, n - 1)

    return [(i + 1, pdf.pages[i]) for i in range(start_idx, end_idx + 1)]




# VALIDATION: Find the first audit report page in the PDF (skipping TOC pages).
def find_first_audit_report_page(pdf) -> Optional[Tuple[int, Any]]:
    """
    Returns (start_idx, page_obj) for the FIRST page that matches the strict
    Independent Auditor's Report heading, skipping TOC pages.
    start_idx is 0-based index in pdf.pages.
    """
    for idx, page in enumerate(pdf.pages):
        txt = page.extract_text() or ""
        if not txt.strip():
            continue

        # Skip Table of Contents pages (do NOT stop; just ignore them)
        if TOC_STOP_ENABLED and is_table_of_contents_page(txt):
            continue

        if is_independent_auditors_report_page(txt):
            return (idx, page)

    return None

# VALIDATION: Validate the presence of the Audit Report section and write PASS/FAIL with proof image.
def _orig_write_audit_report_validation(
    processing_id: int,
    issuer_name: str,
    fy_end_value,
    pdf_path: Optional[Path],
    show_popup: bool = False
) -> dict:
    """
    Availability of Audit Report.

    PASS if within the opinion-anchored window we find the MAIN audit report page.
    This function validates report availability only; it should not fail merely
    because issuer-name/FYE matching is imperfect on that same page.

    - Does NOT write to parquet (parent run_validations_for_row batches the write)
    """
    if not pdf_path or not Path(pdf_path).exists():
        return {
            "status": "FAIL",
            "page": None,
            "proof_file": None,
            "boxes": [],
            "title": f"ProcessingId {processing_id} | Audit Report = FAIL (PDF missing)"
        }

    try:
        with open_pdf_maybe_shared(pdf_path) as pdf:
            if not pdf.pages:
                return {
                    "status": "NOT EXTRACTABLE",
                    "page": None,
                    "proof_file": None,
                    "boxes": [],
                    "title": f"ProcessingId {processing_id} | Audit Report = NOT EXTRACTABLE (No pages)"
                }

            anchor = find_opinion_anchor_in_pdf(pdf, max_scan_pages=120)

            selected_pages = []
            if anchor:
                heading_page_num = int(anchor.get("heading_page") or 1)
                opinion_page_num = int(anchor.get("opinion_page") or heading_page_num)
                selected_pages = build_audit_anchor_window_pages(
                    pdf,
                    heading_page_num=heading_page_num,
                    opinion_page_num=opinion_page_num,
                    max_total_pages=20,
                    post_opinion_pages=3
                )

            # fallback if anchor window is empty
            if not selected_pages:
                first_hit = find_first_audit_report_page(pdf)
                if first_hit:
                    selected_pages = [(first_hit[0] + 1, first_hit[1])]

            for page_num, page in selected_pages:
                txt = cached_page_text(page) or ""
                if not txt.strip():
                    continue

                if not is_independent_auditors_report_page(txt):
                    continue

                heading_box = find_audit_report_heading_box(page)

                proof_file = PROOF_DIR_AUDIT / f"Row{processing_id}_AUDIT_REPORT_PASS_Page{page_num}.png"
                im = page.to_image(resolution=PROOF_IMAGE_RESOLUTION)
                boxes = []

                if heading_box:
                    im.draw_rect(heading_box, stroke=HIGHLIGHT_STROKE, fill=HIGHLIGHT_FILL_RGBA)
                    boxes.append(heading_box)

                im.save(proof_file)
                title = f"ProcessingId {processing_id} | Audit Report = PASS | Page {page_num}"
                
                if show_popup:
                    show_proof_popup(proof_file, title, seconds=PROOF_POPUP_SECONDS)

                return {
                    "status": "PASS",
                    "page": page_num,
                    "proof_file": proof_file,
                    "boxes": boxes,
                    "title": title
                }

            return {
                "status": "FAIL",
                "page": None,
                "proof_file": None,
                "boxes": [],
                "title": f"ProcessingId {processing_id} | Audit Report = FAIL"
            }
    except Exception as e:
        return {
            "status": "FAIL",
            "page": None,
            "proof_file": None,
            "boxes": [],
            "title": f"ProcessingId {processing_id} | Audit Report = FAIL ({e})"
        }



def write_audit_report_validation(
    processing_id: int,
    issuer_name: str,
    fy_end_value,
    pdf_path: Optional[Path],
    show_popup: bool = True
) -> dict:
    """
    Hybrid wrapper:
    1) keep the existing Audit Report validation as primary;
    2) if it fails, apply C4F-style audit-report scoring fallback.

    - Does NOT write to parquet (parent run_validations_for_row batches the write)
    """
    res = _orig_write_audit_report_validation(
        processing_id=processing_id,
        issuer_name=issuer_name,
        fy_end_value=fy_end_value,
        pdf_path=pdf_path,
        show_popup=show_popup,
    )
    if isinstance(res, dict) and str(res.get('status', '')).upper().startswith('PASS'):
        return res

    if not pdf_path or not Path(pdf_path).exists():
        return res if isinstance(res, dict) else {
            "status": "FAIL", "page": None, "proof_file": None, "boxes": [],
            "title": f"ProcessingId {processing_id} | Audit Report = FAIL (PDF missing)"
        }

    try:
        with open_pdf_maybe_shared(pdf_path) as pdf:
            found = _c4f_find_audit_report_page_fallback(pdf, max_scan_pages=120)
            if not found:
                return res if isinstance(res, dict) else {
                    "status": "FAIL", "page": None, "proof_file": None, "boxes": [],
                    "title": f"ProcessingId {processing_id} | Audit Report = FAIL"
                }

            page_num, page, score, signals = found
            proof_file = PROOF_DIR_AUDIT / f"Row{processing_id}_AUDIT_REPORT_C4F_PASS_Page{page_num}.png"
            box = _c4f_audit_heading_or_signal_box(page)
            boxes = [box] if box else []
            try:
                im = page.to_image(resolution=PROOF_IMAGE_RESOLUTION)
                if box:
                    im.draw_rect(box, stroke=HIGHLIGHT_STROKE, fill=HIGHLIGHT_FILL_RGBA)
                im.save(proof_file)
            except Exception:
                proof_file = None

            title = f"ProcessingId {processing_id} | Audit Report = PASS | Page {page_num} | C4F score {score:.2f}"
            if show_popup and proof_file:
                show_proof_popup(proof_file, title, seconds=PROOF_POPUP_SECONDS)
            return {
                "status": "PASS",
                "page": page_num,
                "proof_file": proof_file,
                "boxes": boxes,
                "title": title,
                "fallback": "C4F audit scoring",
                "score": score,
                "signals": signals,
            }
    except Exception:
        return res if isinstance(res, dict) else {
            "status": "FAIL", "page": None, "proof_file": None, "boxes": [],
            "title": f"ProcessingId {processing_id} | Audit Report = FAIL"
        }
# =========================================================
# AUDIT OPINION VALIDATION v3 FAST INTEGRATION HELPERS
# =========================================================
# Scope: AUDIT OPINION validation only (Column G).
# Existing non-opinion logic is not changed.
# Runtime improvement:
#   1) Try existing Audit Report page/window first when Column F passed.
#   2) If not found there, scan deeper but stop after high-confidence opinion page.


def auditop_v3_compact_text(text: str) -> str:
    """Return ``text`` lowercased with all whitespace removed (handles spaced-out PDF extraction)."""
    return re.sub(r"[^a-z0-9]+", "", normalize_pdf_text(text or "").lower())


def auditop_v3_top_text(page_text: str, max_lines: int = 45) -> str:
    """Return the normalized lowercase text of the first ``max_lines`` non-empty lines of a page."""
    return normalize_pdf_text(
        " ".join([ln.strip() for ln in (page_text or "").splitlines() if ln.strip()][:max_lines])
    )


def auditop_v3_is_toc_page(page_text: str) -> bool:
    """Return ``True`` if the page is a Table of Contents page (audit-opinion v3 detector)."""
    t = auditop_v3_top_text(page_text, 25).lower()
    c = auditop_v3_compact_text(t)

    if "wehaveaudited" in c or "inouropinion" in c or "presentfairly" in c:
        return False

    return bool(
        "tableofcontents" in c
        or re.search(r"^\s*table\s+of\s+contents\b", t, re.I)
    )


def auditop_v3_is_divider_page(page_text: str) -> bool:
    """Return ``True`` if the page is a section divider/cover page with little real content."""
    lines = [ln.strip() for ln in (page_text or "").splitlines() if ln.strip()]
    if len(lines) > 4:
        return False

    c = auditop_v3_compact_text(" ".join(lines))
    return c in {
        "financialsection",
        "basicfinancialstatements",
        "introductorysection",
        "statisticalsection",
        "compliancesection",
    }


def auditop_v3_is_non_financial_audit_report(page_text: str) -> bool:
    """
    Reject Government Auditing Standards internal-control reports and Uniform Guidance compliance reports.
    These can contain "Independent Auditor's Report" and "In our opinion" but are not the main
    financial-statement opinion.
    """
    top = auditop_v3_top_text(page_text, 55).lower()
    comp = auditop_v3_compact_text(top)

    reject_patterns = [
        r"report\s+on\s+internal\s+control\s+over\s+financial\s+reporting",
        r"report\s+on\s+compliance\s+for\s+each\s+major",
        r"report\s+on\s+compliance\s+with\s+requirements",
        r"internal\s+control\s+over\s+compliance",
        r"uniform\s+guidance",
        r"schedule\s+of\s+expenditures\s+of\s+federal\s+awards",
        r"major\s+federal\s+program",
        r"major\s+state\s+program",
    ]

    if any(re.search(p, top, re.I) for p in reject_patterns):
        # Keep true financial-statement report if it clearly says report on audit of financial statements
        # and also has the financial-statement opinion wording.
        if (
            ("reportontheauditofthefinancialstatements" in comp or "reportonthefinancialstatements" in comp)
            and "presentfairly" in auditop_v3_compact_text(page_text)
        ):
            return False
        return True

    return False


def auditop_v3_financial_opinion_score(page_text: str) -> Tuple[float, Dict[str, bool]]:
    """
    Score page for main financial-statement audit opinion.
    This is adapted from the standalone Audit Opinion v3 FAST logic.
    """
    if not page_text or not page_text.strip():
        return 0.0, {}

    if (
        auditop_v3_is_toc_page(page_text)
        or auditop_v3_is_divider_page(page_text)
        or auditop_v3_is_non_financial_audit_report(page_text)
    ):
        return 0.0, {"rejected": True}

    txt = normalize_pdf_text(page_text)
    low = txt.lower()
    top = auditop_v3_top_text(txt, 45).lower()
    comp = auditop_v3_compact_text(txt)
    top_comp = auditop_v3_compact_text(top)

    sig = {
        "independent_auditor_report_heading": bool(
            re.search(r"\bindependent\s+au\s*ditor(?:s)?'?s?\s+report\b", top, re.I)
            or "independentauditorsreport" in top_comp
            or "independentauditorreport" in top_comp
        ),
        "report_on_audit_of_financial_statements": bool(
            "reportontheauditofthefinancialstatements" in comp
            or "reportonthefinancialstatements" in comp
        ),
        "opinions_heading": bool(re.search(r"\bopinions?\b", top, re.I) or "opinions" in top_comp),
        "we_have_audited": bool(re.search(r"\bwe\s+have\s+audited\b", low, re.I) or "wehaveaudited" in comp),
        "in_our_opinion": bool(re.search(r"\bin\s+our\s+opinion\b", low, re.I) or "inouropinion" in comp),
        "present_fairly": bool(re.search(r"\bpresent\s+fairly\b", low, re.I) or "presentfairly" in comp),
        "fairly_presented": bool(re.search(r"\bfairly\s+presented\b", low, re.I) or "fairlypresented" in comp),
        "in_all_material_respects": bool(
            re.search(r"\bin\s+all\s+material\s+respects\b", low, re.I)
            or "inallmaterialrespects" in comp
        ),
        "basis_for_opinion": bool(
            re.search(r"\bbasis\s+for\s+opinions?\b", low, re.I)
            or "basisforopinion" in comp
            or "basisforopinions" in comp
        ),
        "gaap": bool(
            "accountingprinciplesgenerallyaccepted" in comp
            or "generallyacceptedaccountingprinciples" in comp
        ),
    }

    # Must look like the main financial-statement opinion.
    if not (sig["present_fairly"] or sig["fairly_presented"]):
        return 0.0, sig

    if not (sig["in_our_opinion"] or sig["opinions_heading"]):
        return 0.0, sig

    if not (
        sig["independent_auditor_report_heading"]
        or sig["report_on_audit_of_financial_statements"]
        or sig["we_have_audited"]
    ):
        return 0.0, sig

    score = 0.0
    if sig["independent_auditor_report_heading"]:
        score += 0.20
    if sig["report_on_audit_of_financial_statements"]:
        score += 0.15
    if sig["opinions_heading"]:
        score += 0.10
    if sig["we_have_audited"]:
        score += 0.15
    if sig["in_our_opinion"]:
        score += 0.15
    if sig["present_fairly"] or sig["fairly_presented"]:
        score += 0.15
    if sig["in_all_material_respects"]:
        score += 0.05
    if sig["basis_for_opinion"]:
        score += 0.10
    if sig["gaap"]:
        score += 0.05

    return min(score, 1.0), sig



def auditop_v3_window_from_anchor(pdf, anchor: dict, window_pages: int = 8) -> dict:
    """Build a list of ``(page_num, page)`` tuples forming the audit window around the opinion ``anchor``."""
    pages = []

    for i in range(anchor["page_index"], min(len(pdf.pages), anchor["page_index"] + window_pages)):
        txt = cached_page_text(pdf.pages[i]) or ""
        ttop = auditop_v3_top_text(txt, 12).lower()

        # Stop after report if MD&A / basic statements starts.
        if i > anchor["page_index"] and re.search(
            r"\bmanagement['’]?s\s+discussion\s+and\s+analysis\b"
            r"|\bbasic\s+financial\s+statements\b"
            r"|\bfinancial\s+statements\b\s*$",
            ttop,
            re.I,
        ):
            break

        pages.append((i + 1, pdf.pages[i]))

    anchor["pages"] = pages
    return anchor


def auditop_v3_try_audit_res_window(pdf, audit_res: Optional[dict], window_pages: int = 8) -> Optional[dict]:
    """
    Runtime shortcut only.
    If Column F already passed, check that page and nearby pages first.
    If not a financial-statement opinion page, return None and full scan continues.
    """
    if not audit_res or not str(audit_res.get("status", "")).upper().startswith("PASS"):
        return None

    try:
        page_num = int(audit_res.get("page") or 0)
    except Exception:
        page_num = 0

    if page_num < 1 or page_num > len(pdf.pages):
        return None

    best = None
    start_idx = page_num - 1
    end_idx = min(len(pdf.pages) - 1, start_idx + max(0, window_pages - 1))

    for idx in range(start_idx, end_idx + 1):
        txt = cached_page_text(pdf.pages[idx]) or ""
        score, sig = auditop_v3_financial_opinion_score(txt)

        if score <= 0:
            continue

        candidate = {
            "page_index": idx,
            "page_num": idx + 1,
            "score": score,
            "signals": sig,
        }

        if best is None or score > best["score"]:
            best = candidate

        if score >= 0.90:
            break

    if best and best["score"] >= 0.50:
        return auditop_v3_window_from_anchor(pdf, best, window_pages=window_pages)

    return None


def auditop_v3_find_financial_audit_report(
    pdf,
    audit_res: Optional[dict] = None,
    max_scan_pages: int = 300,
    window_pages: int = 8,
) -> Optional[dict]:
    """
    FAST v3 scanner.
    Identification logic uses the same financial-opinion score and threshold >= 0.50.
    Runtime improvement:
    1) try Column F page/window first when available;
    2) stop full scan once score >= 0.90.
    """
    anchor = auditop_v3_try_audit_res_window(pdf, audit_res=audit_res, window_pages=window_pages)
    if anchor:
        anchor["source"] = "AUDIT_RES_WINDOW"
        return anchor

    HIGH_CONFIDENCE_STOP_SCORE = 0.90
    limit = min(len(pdf.pages), max_scan_pages)
    best = None

    for idx in range(limit):
        page = pdf.pages[idx]
        txt = cached_page_text(page) or ""
        score, sig = auditop_v3_financial_opinion_score(txt)

        if score <= 0:
            continue

        candidate = {
            "page_index": idx,
            "page_num": idx + 1,
            "score": score,
            "signals": sig,
        }

        if best is None or score > best["score"]:
            best = candidate

        if score >= HIGH_CONFIDENCE_STOP_SCORE:
            break

    if not best or best["score"] < 0.50:
        return None

    best["source"] = "FULL_SCAN"
    return auditop_v3_window_from_anchor(pdf, best, window_pages=window_pages)


def auditop_v3_extract_opinion_scope(report_text: str) -> str:
    """Extract the opinion paragraph (text between 'Opinion' and 'Basis for Opinion') from ``report_text``."""
    txt = normalize_pdf_text(report_text or "")
    if not txt:
        return ""

    # Normal Opinion(s) -> Basis for Opinion(s)
    m = re.search(r"\bopinions?\b(.+?)\bbasis\s+for\s+opinions?\b", txt, re.I | re.S)
    if m:
        return normalize_pdf_text(m.group(1))

    # Around In our opinion -> next section heading
    m = re.search(
        r"(\bin\s+our\s+opinion\b.+?)"
        r"(?:\bbasis\s+for\s+opinions?\b"
        r"|\bemphasis\s+of\s+matter\b"
        r"|\bresponsibilities\s+of\s+management\b"
        r"|\bauditor['’]?s\s+responsibilities\b)",
        txt,
        re.I | re.S,
    )
    if m:
        return normalize_pdf_text(m.group(1))

    # Compact fallback around inouropinion / presentfairly.
    comp = auditop_v3_compact_text(txt)

    for marker in ["inouropinion", "presentfairly", "fairlypresented"]:
        pos = comp.find(marker)

        if pos >= 0:
            comp_to_orig = []

            for i, ch in enumerate(txt):
                if ch.isalnum():
                    comp_to_orig.append(i)

            if pos < len(comp_to_orig):
                orig_start = max(0, comp_to_orig[pos] - 250)
                orig_end = min(
                    len(txt),
                    comp_to_orig[min(len(comp_to_orig) - 1, pos + 2200)]
                    if comp_to_orig
                    else len(txt),
                )
                return normalize_pdf_text(txt[orig_start:orig_end])

    return txt[:2500]


def auditop_v3_split_sentences(text: str) -> List[str]:
    """Split ``text`` into a list of sentences on sentence-ending punctuation."""
    text = normalize_pdf_text(text or "")
    if not text:
        return []
    return [s.strip() for s in re.split(r"(?<=[.!?])\s+", text) if s.strip()]


def auditop_v3_find_opinion_sentence(scope: str) -> Optional[str]:
    """Return the most relevant opinion sentence from ``scope`` (prefers 'in our opinion')."""
    if not scope:
        return None

    for s in auditop_v3_split_sentences(scope):
        c = auditop_v3_compact_text(s)
        if re.search(r"\bin\s+our\s+opinion\b", s, re.I) or "inouropinion" in c:
            return s

    for s in auditop_v3_split_sentences(scope):
        c = auditop_v3_compact_text(s)
        if (
            re.search(r"\bpresent\s+fairly\b|\bfairly\s+presented\b|\btrue\s+and\s+fair\b", s, re.I)
            or any(x in c for x in ["presentfairly", "fairlypresented", "trueandfair"])
        ):
            return s

    c = auditop_v3_compact_text(scope)
    if any(x in c for x in ["inouropinion", "presentfairly", "fairlypresented"]):
        return normalize_pdf_text(scope[:1400])

    return None


def auditop_v3_classify_opinion(scope: str, sentence: Optional[str] = None) -> Optional[str]:
    """Classify an audit opinion (Unmodified/Qualified/Adverse/Disclaimer/etc.) from the scope text and sentence."""
    text = normalize_pdf_text(sentence or scope or "")
    if not text:
        return None

    comp = auditop_v3_compact_text(text)
    no_paren = re.sub(r"\([^)]*\)", "", text)
    comp_no_paren = auditop_v3_compact_text(no_paren)

    if re.search(r"\badverse\s+opinion\b", text, re.I) or "adverseopinion" in comp:
        return "Adverse"

    if (
        re.search(r"\bdisclaimer\s+of\s+opinion\b|\bdo\s+not\s+express\s+an\s+opinion\b", text, re.I)
        or "disclaimerofopinion" in comp
        or "donotexpressanopinion" in comp
    ):
        return "Disclaimer"

    if re.search(r"\bqualified\s+opinion\b", text, re.I) or "qualifiedopinion" in comp:
        return "Qualified"

    if (
        re.search(r"\bexcept\s+for\b.*\b(effects|possible\s+effects)\b", no_paren, re.I)
        or "exceptfortheeffects" in comp_no_paren
        or "exceptforthepossibleeffects" in comp_no_paren
    ):
        return "Qualified"

    if re.search(r"\bunmodified\s+opinion\b", text, re.I) or "unmodifiedopinion" in comp:
        return "Unmodified"

    if re.search(r"\bpresent\s+fairly\b", text, re.I) or "presentfairly" in comp:
        return "Present Fairly"

    if re.search(r"\bfairly\s+presented\b", text, re.I) or "fairlypresented" in comp:
        return "Fairly Presented"

    if re.search(r"\btrue\s+and\s+fair\b", text, re.I) or "trueandfair" in comp:
        return "True and Fair"

    if (
        ("presentfairly" in comp and "inallmaterialrespects" in comp)
        or ("presentfairly" in comp and "accountingprinciplesgenerallyaccepted" in comp)
    ):
        return "Present Fairly"

    return None


def auditop_v3_box_from_words(words: List[dict]) -> Optional[dict]:
    """Return a bounding-box dict spanning all pdfplumber word dicts in ``words`` (or ``None`` if empty)."""
    words = [w for w in (words or []) if w]
    if not words:
        return None

    return {
        "x0": min(float(w["x0"]) for w in words),
        "top": min(float(w["top"]) for w in words),
        "x1": max(float(w["x1"]) for w in words),
        "bottom": max(float(w["bottom"]) for w in words),
    }


def auditop_v3_page_tokens_with_boxes(page) -> Tuple[List[str], List[dict]]:
    """Return aligned ``(page_tokens, token_word_map)`` lists for a page, mapping each token to its source word."""
    page_tokens, token_word_map = [], []

    for w in cached_page_words(page, use_text_flow=True):
        for tok in tokenize(w.get("text", "")):
            page_tokens.append(tok)
            token_word_map.append(w)

    return page_tokens, token_word_map


def auditop_v3_find_phrase_box_on_page(page, phrase: str) -> Optional[dict]:
    """Find ``phrase`` on a page and return its highlight bounding box, or ``None`` if not found."""
    target = tokenize(phrase)
    if not target:
        return None

    tokens, word_map = auditop_v3_page_tokens_with_boxes(page)
    s, e = find_token_window(tokens, target)

    if s is None:
        return None

    return auditop_v3_box_from_words(word_map[s:e + 1])


def auditop_v3_find_compact_phrase_box_on_page(page, compact_phrase: str) -> Optional[dict]:
    """Find a whitespace-insensitive ``compact_phrase`` on a page and return its bounding box, or ``None``."""
    compact_phrase = re.sub(r"[^a-z0-9]+", "", (compact_phrase or "").lower())

    if not compact_phrase:
        return None

    words = cached_page_words(page, use_text_flow=True)

    for i in range(len(words)):
        accum = ""
        selected = []

        for j in range(i, min(len(words), i + 14)):
            wtxt = auditop_v3_compact_text(words[j].get("text", ""))

            if not wtxt:
                continue

            accum += wtxt
            selected.append(words[j])

            if compact_phrase in accum:
                return auditop_v3_box_from_words(selected)

            if len(accum) > len(compact_phrase) + 60:
                break

    return None


def auditop_v3_opinion_proof_box(page) -> Optional[dict]:
    """Return a bounding box for the opinion heading on the page for proof highlighting, or ``None``."""
    for phrase in ["present fairly", "fairly presented", "In our opinion", "Opinions", "Opinion"]:
        box = auditop_v3_find_phrase_box_on_page(page, phrase)
        if box:
            return box

    for cphrase in ["presentfairly", "fairlypresented", "inouropinion", "opinions", "opinion"]:
        box = auditop_v3_find_compact_phrase_box_on_page(page, cphrase)
        if box:
            return box

    return None


def auditop_v3_save_final_proof(page, boxes: List[dict], proof_path: Path) -> bool:
    """Render the page, draw the supplied ``boxes``, and save the proof PNG to ``proof_path``. Returns success bool."""
    proof_path.parent.mkdir(parents=True, exist_ok=True)

    try:
        im = page.to_image(resolution=PROOF_IMAGE_RESOLUTION)

        for b in boxes or []:
            if b:
                im.draw_rect(b, stroke=HIGHLIGHT_STROKE, fill=HIGHLIGHT_FILL_RGBA)

        im.save(str(proof_path))
        return True

    except Exception as e:
        print(f"[WARN] Could not save Audit Opinion proof image: {e}")
        return False


def auditop_v3_validate_integrated(
    pdf_path: Path,
    proof_path: Path,
    audit_res: Optional[dict] = None,
    max_scan_pages: int = 300,
) -> dict:
    """
    Validate Audit Opinion using v3 FAST logic inside the main script.
    """
    try:
        if proof_path.exists():
            proof_path.unlink()
    except Exception:
        pass

    if not pdf_path or not Path(pdf_path).exists():
        return {
            "status": "FAIL",
            "method": "PDF",
            "reason": "PDF file not found",
            "page": None,
            "proof_file": None,
            "boxes": [],
        }

    with open_pdf_maybe_shared(str(pdf_path)) as pdf:
        anchor = auditop_v3_find_financial_audit_report(
            pdf,
            audit_res=audit_res,
            max_scan_pages=max_scan_pages,
            window_pages=8,
        )

        if not anchor:
            return {
                "status": "FAIL",
                "method": "AUDIT_REPORT_DETECTION",
                "reason": "Main financial-statement audit opinion page not found",
                "page": None,
                "proof_file": None,
                "boxes": [],
            }

        report_text = normalize_pdf_text(
            "\n".join((cached_page_text(p) or "") for _, p in anchor.get("pages", []))
        )

        scope = auditop_v3_extract_opinion_scope(report_text)
        sentence = auditop_v3_find_opinion_sentence(scope)
        label = auditop_v3_classify_opinion(scope, sentence)

        if not label:
            return {
                "status": "FAIL",
                "method": "OPINION_CLASSIFICATION",
                "reason": "Financial audit report found but opinion could not be classified",
                "page": anchor.get("page_num"),
                "proof_file": None,
                "boxes": [],
                "audit_report_page": anchor.get("page_num"),
                "audit_report_score": anchor.get("score"),
                "audit_report_signals": anchor.get("signals"),
                "opinion_scope_preview": scope[:1200],
            }

        proof_page_num = anchor["page_num"]
        proof_page = pdf.pages[proof_page_num - 1]
        box = auditop_v3_opinion_proof_box(proof_page)

        if not box:
            for pnum, page in anchor.get("pages", []):
                box = auditop_v3_opinion_proof_box(page)
                if box:
                    proof_page_num = pnum
                    proof_page = page
                    break

        boxes = [box] if box else []
        proof_file = None

        if box and auditop_v3_save_final_proof(proof_page, boxes, proof_path):
            proof_file = proof_path

        return {
            "status": f"PASS - {label}",
            "method": "AUDIT_OPINION_V3_FAST",
            "reason": "Financial statement audit opinion found and classified",
            "page": proof_page_num,
            "proof_file": proof_file,
            "boxes": boxes,
            "label": label,
            "audit_report_page": anchor.get("page_num"),
            "audit_report_score": anchor.get("score"),
            "audit_report_signals": anchor.get("signals"),
            "audit_report_source": anchor.get("source"),
            "opinion_sentence": sentence,
            "opinion_scope_preview": scope[:1200],
        }

# VALIDATION: Validate and classify the audit opinion and return the result with proof image.
def write_audit_opinion_validation(
    processing_id: int,
    pdf_path: Optional[Path],
    audit_res: Optional[dict] = None,
    show_popup: bool = True
) -> dict:
    """
    Audit Opinion.

    Integrated v3 FAST logic only for Audit Opinion:
    - skips TOC/divider/compliance/internal-control reports;
    - finds the main financial-statement audit opinion;
    - handles merged extraction text such as Inour / presentfairly;
    - scans deeper for long ACFRs but stops early once high-confidence opinion is found;
    - if Audit Report validation already passed, checks that page/window first for speed.

    - Does NOT write to parquet (parent run_validations_for_row batches the write)
    """
    if not pdf_path or not Path(pdf_path).exists():
        return {
            "status": "FAIL",
            "page": None,
            "proof_file": None,
            "boxes": [],
            "title": f"ProcessingId {processing_id} | Audit Opinion = FAIL (PDF missing)",
            "method": "PDF",
            "reason": "PDF file not found",
        }

    proof_file = PROOF_DIR_OPINION / f"Row{processing_id}_OPINION_PASS.png"

    try:
        result = auditop_v3_validate_integrated(
            pdf_path=Path(pdf_path),
            proof_path=proof_file,
            audit_res=audit_res,
            max_scan_pages=300,
        )

        status = result.get("status", "FAIL")

        page_no = result.get("page")
        boxes = result.get("boxes", []) or []
        proof = result.get("proof_file")
        label = result.get("label")
        method = result.get("method")
        reason = result.get("reason", "")

        title = f"ProcessingId {processing_id} | Audit Opinion = {status}"
        if page_no:
            title += f" | Page {page_no}"
        if label:
            title = f"ProcessingId {processing_id} | Audit Opinion = PASS - {label} | Page {page_no}"

        if show_popup and str(status).upper().startswith("PASS") and proof:
            show_proof_popup(Path(proof), title, seconds=PROOF_POPUP_SECONDS)

        return {
            "status": status,
            "page": page_no,
            "proof_file": Path(proof) if proof else None,
            "boxes": boxes,
            "title": title,
            "method": method,
            "reason": reason,
            "label": label,
            "audit_report_page": result.get("audit_report_page"),
            "audit_report_score": result.get("audit_report_score"),
            "audit_report_source": result.get("audit_report_source"),
            "opinion_sentence": result.get("opinion_sentence"),
        }

    except Exception as e:
        return {
            "status": "FAIL",
            "page": None,
            "proof_file": None,
            "boxes": [],
            "title": f"ProcessingId {processing_id} | Audit Opinion = FAIL ({e})",
            "method": "EXCEPTION",
            "reason": str(e),
        }         
# =========================================================
# AUDIT REPORT FYE vs REPORT FYE (Detailed Validation Check - Column H)
# =========================================================


# =========================================================
# AUDIT REPORT FYE VALIDATION v1 FAST INTEGRATION HELPERS
# =========================================================
# Added from standalone Audit FYE.py / Audit_Report_FYE_Test.py.
# Scope: Column H only - Audit Report FYE with Report FYE.
# Existing non-Column-H logic is not changed.
# Runtime improvement:
#   1) Try Column F audit report page/window first when available.
#   2) If not found there, scan deeper but stop once high-confidence audit-scope/FYE page is found.

AUDITFYE_CONTEXT_FILL_RGBA = (255, 165, 0, 90)


def auditfye_compact_text(text: str) -> str:
    """Return ``text`` lowercased with whitespace removed (audit-FYE matcher)."""
    return re.sub(r"[^a-z0-9]+", "", normalize_pdf_text(text or "").lower())


def auditfye_top_text(page_text: str, max_lines: int = 45) -> str:
    """Return normalized lowercase text from the first ``max_lines`` non-empty lines of a page."""
    return normalize_pdf_text(
        " ".join([ln.strip() for ln in (page_text or "").splitlines() if ln.strip()][:max_lines])
    )


def auditfye_clean_for_match(s: str) -> str:
    """Aggressively clean ``s`` to lowercase alphanumerics and single spaces for contains-matching."""
    s = normalize_pdf_text(s or "").lower()
    s = re.sub(r"[^a-z0-9 ]+", " ", s)
    return re.sub(r"\s+", " ", s).strip()


def auditfye_tokenize(s: str) -> List[str]:
    """Tokenize ``s`` into cleaned lowercase tokens (audit-FYE matcher)."""
    return [t for t in auditfye_clean_for_match(s).split() if t]


def auditfye_find_token_window(page_tokens: List[str], target_tokens: List[str]) -> Tuple[Optional[int], Optional[int]]:
    """Find ``target_tokens`` within ``page_tokens``; returns ``(start, end)`` or ``(None, None)``."""
    n, m = len(page_tokens), len(target_tokens)
    if not n or not m or m > n:
        return None, None
    for i in range(n - m + 1):
        if page_tokens[i:i + m] == target_tokens:
            return i, i + m - 1
    return None, None


def auditfye_box_from_words(words: List[dict]) -> Optional[dict]:
    """Return a bounding box spanning the given word dicts, or ``None`` if empty."""
    words = [w for w in (words or []) if w]
    if not words:
        return None
    return {
        "x0": min(float(w["x0"]) for w in words),
        "top": min(float(w["top"]) for w in words),
        "x1": max(float(w["x1"]) for w in words),
        "bottom": max(float(w["bottom"]) for w in words),
    }


def auditfye_page_tokens_with_boxes(page) -> Tuple[List[str], List[dict]]:
    """Return aligned ``(tokens, word_map)`` for a page (audit-FYE matcher)."""
    page_tokens, token_word_map = [], []
    for w in cached_page_words(page, use_text_flow=True):
        for tok in auditfye_tokenize(w.get("text", "")):
            page_tokens.append(tok)
            token_word_map.append(w)
    return page_tokens, token_word_map


_AUDITFYE_MONTHS = {
    1: ("january", "jan"), 2: ("february", "feb"), 3: ("march", "mar"),
    4: ("april", "apr"), 5: ("may", "may"), 6: ("june", "jun"),
    7: ("july", "jul"), 8: ("august", "aug"), 9: ("september", "sep", "sept"),
    10: ("october", "oct"), 11: ("november", "nov"), 12: ("december", "dec"),
}
_AUDITFYE_MONTH_NAME_TO_NUM = {nm: m for m, names in _AUDITFYE_MONTHS.items() for nm in names}


def auditfye_parse_fye_date(value) -> Optional[date]:
    """Parse ``value`` into a ``date`` across many formats; returns ``None`` if unparseable."""
    if value is None:
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    s = str(value).strip()
    if not s:
        return None
    s = normalize_pdf_text(s)

    # Excel serial date fallback, harmless for normal date strings.
    if re.fullmatch(r"\d{5}(?:\.0)?", s):
        try:
            from datetime import timedelta
            serial = int(float(s))
            return date(1899, 12, 30) + timedelta(days=serial)
        except Exception:
            pass

    for fmt in (
        "%m/%d/%Y", "%m/%d/%y", "%Y-%m-%d", "%d-%m-%Y", "%d/%m/%Y",
        "%B %d, %Y", "%b %d, %Y", "%B %d %Y", "%b %d %Y",
    ):
        try:
            return datetime.strptime(s, fmt).date()
        except Exception:
            pass

    m = re.search(r"\b([A-Za-z]+)\s+(\d{1,2})(?:st|nd|rd|th)?\s*,?\s*((?:19|20)\d{2})\b", s)
    if m:
        raw_mon = m.group(1).lower().rstrip(".")
        mon = _AUDITFYE_MONTH_NAME_TO_NUM.get(raw_mon) or _AUDITFYE_MONTH_NAME_TO_NUM.get(raw_mon[:3]) or _AUDITFYE_MONTH_NAME_TO_NUM.get(raw_mon[:4])
        if mon:
            return date(int(m.group(3)), mon, int(m.group(2)))

    # Fallback to the main script parser if available.
    try:
        return parse_fy_end_date(value)
    except Exception:
        return None


def auditfye_date_strings(d: date) -> List[str]:
    """Return a set of textual representations of date ``d`` across common formats."""
    full, abbr = _AUDITFYE_MONTHS[d.month][0], _AUDITFYE_MONTHS[d.month][1]
    month_full = full.title()
    month_abbr = abbr.title()
    day = str(d.day)
    day2 = f"{d.day:02d}"
    mon = str(d.month)
    mon2 = f"{d.month:02d}"
    year = str(d.year)
    year2 = f"{d.year % 100:02d}"
    vals = [
        f"{month_full} {day}, {year}",
        f"{month_full} {day2}, {year}",
        f"{month_full} {day} {year}",
        f"{month_abbr} {day}, {year}",
        f"{month_abbr}. {day}, {year}",
        f"{mon}/{day}/{year}", f"{mon2}/{day2}/{year}", f"{mon}/{day}/{year2}", f"{mon2}/{day2}/{year2}",
        f"{mon}-{day}-{year}", f"{mon2}-{day2}-{year}",
        f"{year}-{mon2}-{day2}",
    ]
    out, seen = [], set()
    for v in vals:
        key = v.lower()
        if key not in seen:
            seen.add(key)
            out.append(v)
    return out


def auditfye_build_fye_token_candidates(d: date) -> List[List[str]]:
    """Build candidate token sequences representing date ``d`` in many PDF date formats."""
    cands = []
    for s in auditfye_date_strings(d):
        toks = auditfye_tokenize(s)
        if toks and toks not in cands:
            cands.append(toks)
    return cands


def auditfye_find_date_token_box_on_page(page, d: date) -> Tuple[Optional[dict], Optional[int], Optional[int], Optional[List[str]]]:
    """Find date ``d`` on a page and return its bounding box, or ``None``."""
    page_tokens, word_map = auditfye_page_tokens_with_boxes(page)
    for cand in auditfye_build_fye_token_candidates(d):
        s, e = auditfye_find_token_window(page_tokens, cand)
        if s is not None:
            return auditfye_box_from_words(word_map[s:e + 1]), s, e, cand
    return None, None, None, None


def auditfye_find_phrase_box_on_page(page, phrase: str) -> Optional[dict]:
    """Find ``phrase`` on a page and return its bounding box, or ``None``."""
    target = auditfye_tokenize(phrase)
    if not target:
        return None
    tokens, word_map = auditfye_page_tokens_with_boxes(page)
    s, e = auditfye_find_token_window(tokens, target)
    if s is None:
        return None
    return auditfye_box_from_words(word_map[s:e + 1])


def auditfye_is_toc_page(page_text: str) -> bool:
    """Return ``True`` if the page is a Table of Contents page (audit-FYE detector)."""
    t = auditfye_top_text(page_text, 25).lower()
    c = auditfye_compact_text(t)
    if "wehaveaudited" in c or "inouropinion" in c or "presentfairly" in c:
        return False
    return bool("tableofcontents" in c or re.search(r"^\s*table\s+of\s+contents\b", t, re.I))


def auditfye_is_divider_page(page_text: str) -> bool:
    """Return ``True`` if the page is a divider/cover page (audit-FYE detector)."""
    lines = [ln.strip() for ln in (page_text or "").splitlines() if ln.strip()]
    if len(lines) > 4:
        return False
    c = auditfye_compact_text(" ".join(lines))
    return c in {"financialsection", "basicfinancialstatements", "introductorysection", "statisticalsection", "compliancesection"}


def auditfye_is_non_financial_audit_report(page_text: str) -> bool:
    """Return ``True`` if the page is a non-financial audit report (compliance/internal-control/Uniform Guidance)."""
    top = auditfye_top_text(page_text, 60).lower()
    comp = auditfye_compact_text(top)
    reject_patterns = [
        r"report\s+on\s+internal\s+control\s+over\s+financial\s+reporting",
        r"report\s+on\s+compliance\s+for\s+each\s+major",
        r"report\s+on\s+compliance\s+with\s+requirements",
        r"internal\s+control\s+over\s+compliance",
        r"uniform\s+guidance",
        r"schedule\s+of\s+expenditures\s+of\s+federal\s+awards",
        r"major\s+federal\s+program",
        r"major\s+state\s+program",
    ]
    if any(re.search(p, top, re.I) for p in reject_patterns):
        if ("reportontheauditofthefinancialstatements" in comp or "reportonthefinancialstatements" in comp) and "wehaveaudited" in auditfye_compact_text(page_text):
            return False
        return True
    return False


def auditfye_report_score(page_text: str, fye_date: Optional[date] = None) -> Tuple[float, Dict[str, bool]]:
    """Score how strongly a page resembles the main financial-statement audit report containing ``fye_date``."""
    if not page_text or not page_text.strip():
        return 0.0, {}
    if auditfye_is_toc_page(page_text) or auditfye_is_divider_page(page_text) or auditfye_is_non_financial_audit_report(page_text):
        return 0.0, {"rejected": True}

    txt = normalize_pdf_text(page_text)
    low = txt.lower()
    top = auditfye_top_text(txt, 55).lower()
    comp = auditfye_compact_text(txt)
    top_comp = auditfye_compact_text(top)

    has_fye = False
    if fye_date:
        has_fye = any(auditfye_compact_text(s) in comp for s in auditfye_date_strings(fye_date))

    sig = {
        "independent_auditor_report_heading": bool(re.search(r"\bindependent\s+au\s*ditor(?:s)?'?s?\s+report\b", top, re.I) or "independentauditorsreport" in top_comp or "independentauditorreport" in top_comp),
        "report_on_financial_statements": bool("reportontheauditofthefinancialstatements" in comp or "reportonthefinancialstatements" in comp or re.search(r"\breport\s+on\s+the\s+financial\s+statements\b", low, re.I)),
        "we_have_audited": bool(re.search(r"\bwe\s+have\s+audited\b", low, re.I) or "wehaveaudited" in comp),
        "financial_statements": bool(re.search(r"\bfinancial\s+statements\b", low, re.I) or "financialstatements" in comp),
        "as_of": bool(re.search(r"\bas\s+of\b", low, re.I) or "asof" in comp),
        "year_ended": bool(re.search(r"\byear\s+ended\b|\byears\s+ended\b", low, re.I) or "yearended" in comp or "yearsended" in comp),
        "in_our_opinion": bool(re.search(r"\bin\s+our\s+opinion\b", low, re.I) or "inouropinion" in comp),
        "present_fairly": bool(re.search(r"\bpresent\s+fairly\b", low, re.I) or "presentfairly" in comp),
        "basis_for_opinion": bool(re.search(r"\bbasis\s+for\s+opinions?\b", low, re.I) or "basisforopinion" in comp or "basisforopinions" in comp),
        "expected_fye_on_page": has_fye,
    }

    if not sig["we_have_audited"]:
        return 0.0, sig
    if not (sig["financial_statements"] or sig["report_on_financial_statements"]):
        return 0.0, sig

    score = 0.0
    if sig["independent_auditor_report_heading"]: score += 0.20
    if sig["report_on_financial_statements"]: score += 0.15
    if sig["we_have_audited"]: score += 0.20
    if sig["financial_statements"]: score += 0.08
    if sig["as_of"]: score += 0.08
    if sig["year_ended"]: score += 0.10
    if sig["expected_fye_on_page"]: score += 0.12
    if sig["in_our_opinion"]: score += 0.06
    if sig["present_fairly"]: score += 0.06
    if sig["basis_for_opinion"]: score += 0.05
    return min(score, 1.0), sig


def auditfye_window_from_anchor(pdf, anchor: dict, window_pages: int = 8) -> dict:
    """Build ``(page_num, page)`` tuples for the audit window around ``anchor``."""
    pages = []
    for i in range(anchor["page_index"], min(len(pdf.pages), anchor["page_index"] + window_pages)):
        txt = cached_page_text(pdf.pages[i]) or ""
        ttop = auditfye_top_text(txt, 12).lower()
        if i > anchor["page_index"] and re.search(
            r"\bmanagement['’]?s\s+discussion\s+and\s+analysis\b|\bbasic\s+financial\s+statements\b|\bfinancial\s+statements\b\s*$",
            ttop,
            re.I,
        ):
            break
        pages.append((i + 1, pdf.pages[i]))
    anchor["pages"] = pages
    return anchor


def auditfye_try_audit_res_window(pdf, audit_res: Optional[dict], fye_date: Optional[date] = None, window_pages: int = 8) -> Optional[dict]:
    """Runtime-only shortcut: try Column F audit report page/window before full scan."""
    if not audit_res or not str(audit_res.get("status", "")).upper().startswith("PASS"):
        return None
    try:
        page_num = int(audit_res.get("page") or 0)
    except Exception:
        page_num = 0
    if page_num < 1 or page_num > len(pdf.pages):
        return None

    best = None
    start_idx = page_num - 1
    end_idx = min(len(pdf.pages) - 1, start_idx + max(0, window_pages - 1))
    for idx in range(start_idx, end_idx + 1):
        txt = cached_page_text(pdf.pages[idx]) or ""
        score, sig = auditfye_report_score(txt, fye_date=fye_date)
        if score <= 0:
            continue
        candidate = {"page_index": idx, "page_num": idx + 1, "score": score, "signals": sig}
        if best is None or score > best["score"]:
            best = candidate
        if score >= 0.86 and (not fye_date or sig.get("expected_fye_on_page")):
            break
    if best and best["score"] >= 0.45:
        best["source"] = "AUDIT_RES_WINDOW"
        return auditfye_window_from_anchor(pdf, best, window_pages=window_pages)
    return None


def auditfye_find_main_audit_report(pdf, fye_date: Optional[date] = None, audit_res: Optional[dict] = None, max_scan_pages: int = 300, window_pages: int = 8) -> Optional[dict]:
    """Locate the main financial-statement audit report page/window for ``fye_date``; uses ``audit_res`` to speed up."""
    anchor = auditfye_try_audit_res_window(pdf, audit_res=audit_res, fye_date=fye_date, window_pages=window_pages)
    if anchor:
        return anchor

    high_confidence_stop_score = 0.86
    limit = min(len(pdf.pages), max_scan_pages)
    best = None
    for idx in range(limit):
        txt = cached_page_text(pdf.pages[idx]) or ""
        score, sig = auditfye_report_score(txt, fye_date=fye_date)
        if score <= 0:
            continue
        candidate = {"page_index": idx, "page_num": idx + 1, "score": score, "signals": sig}
        if best is None or score > best["score"]:
            best = candidate
        if score >= high_confidence_stop_score and (not fye_date or sig.get("expected_fye_on_page")):
            break
    if not best or best["score"] < 0.45:
        return None
    best["source"] = "FULL_SCAN"
    return auditfye_window_from_anchor(pdf, best, window_pages=window_pages)


def auditfye_sentence_windows(text: str) -> List[str]:
    """Yield candidate text windows/sentences from ``text`` for audit-scope FYE matching."""
    txt = normalize_pdf_text(text or "")
    if not txt:
        return []
    return [p.strip() for p in re.split(r"(?<=[.!?])\s+", txt) if p.strip()]


def auditfye_text_contains_fye(text: str, d: date) -> bool:
    """Return ``True`` if ``text`` contains a representation of date ``d``."""
    comp = auditfye_compact_text(text)
    return any(auditfye_compact_text(s) in comp for s in auditfye_date_strings(d))


def auditfye_scope_sentence_with_fye(page_text: str, d: date) -> Optional[str]:
    """Return the audit-scope sentence ('We have audited' ...) containing date ``d``, or ``None``."""
    txt = normalize_pdf_text(page_text or "")
    if not txt:
        return None

    for s in auditfye_sentence_windows(txt):
        sc = auditfye_compact_text(s)
        if auditfye_text_contains_fye(s, d) and (
            "wehaveaudited" in sc
            or ("financialstatements" in sc and ("yearended" in sc or "asof" in sc or "asofandfortheyearended" in sc))
        ):
            return s

    comp = auditfye_compact_text(txt)
    for marker in ["wehaveaudited", "financialstatements"]:
        pos = comp.find(marker)
        if pos >= 0:
            comp_to_orig = []
            for i, ch in enumerate(txt):
                if ch.isalnum():
                    comp_to_orig.append(i)
            if pos < len(comp_to_orig):
                start = max(0, comp_to_orig[pos] - 100)
                end = min(len(txt), comp_to_orig[min(len(comp_to_orig) - 1, pos + 2500)])
                chunk = txt[start:end]
                if auditfye_text_contains_fye(chunk, d):
                    return normalize_pdf_text(chunk)
    return None


def auditfye_token_context_ok(tokens: List[str], date_start: int) -> bool:
    """Return ``True`` if valid audit-period context precedes the date at ``date_start`` in ``tokens``."""
    if date_start is None:
        return False
    left = max(0, date_start - 90)
    before = " ".join(tokens[left:date_start]).lower()
    patterns = [
        r"\bwe\s+have\s+audited\b",
        r"\bfinancial\s+statements\b",
        r"\bas\s+of\b",
        r"\byear\s+ended\b",
        r"\byears\s+ended\b",
        r"\bfor\s+the\s+year\s+ended\b",
        r"\bfor\s+the\s+years\s+ended\b",
    ]
    return any(re.search(p, before, re.I) for p in patterns) and bool(
        re.search(r"\bfinancial\s+statements\b|\bwe\s+have\s+audited\b", before, re.I)
    )


def auditfye_find_on_page(page, fye_date: date) -> Optional[dict]:
    """Find ``fye_date`` near the audit-scope sentence on a page; returns box info or ``None``."""
    page_text = cached_page_text(page) or ""
    if not page_text.strip():
        return None

    scope_sentence = auditfye_scope_sentence_with_fye(page_text, fye_date)
    if not scope_sentence:
        return None

    date_box, date_start, date_end, date_tokens = auditfye_find_date_token_box_on_page(page, fye_date)
    if not date_box:
        return None

    tokens, _word_map = auditfye_page_tokens_with_boxes(page)
    context_ok = auditfye_token_context_ok(tokens, date_start) if date_start is not None else False

    context_box = auditfye_find_phrase_box_on_page(page, "we have audited")
    if not context_box:
        context_box = auditfye_find_phrase_box_on_page(page, "financial statements")

    return {
        "date_box": date_box,
        "context_box": context_box,
        "scope_sentence": scope_sentence,
        "date_tokens": date_tokens,
        "context_ok": context_ok,
    }


def auditfye_save_final_proof(page, date_boxes: List[dict], context_boxes: List[dict], proof_path: Path) -> bool:
    """Render the page, draw date/context boxes, and save the proof PNG. Returns success bool."""
    proof_path.parent.mkdir(parents=True, exist_ok=True)
    try:
        im = page.to_image(resolution=PROOF_IMAGE_RESOLUTION)
        for b in context_boxes or []:
            if b:
                im.draw_rect(b, stroke=HIGHLIGHT_STROKE, fill=AUDITFYE_CONTEXT_FILL_RGBA)
        for b in date_boxes or []:
            if b:
                im.draw_rect(b, stroke=HIGHLIGHT_STROKE, fill=HIGHLIGHT_FILL_RGBA)
        im.save(str(proof_path))
        return True
    except Exception as e:
        print(f"[WARN] Could not save Audit Report FYE proof image: {e}")
        return False


def auditfye_validate_integrated(pdf_path: Path, fy_end_value, proof_path: Path, audit_res: Optional[dict] = None, max_scan_pages: int = 300) -> dict:
    """Validate the audit-report FYE against ``fy_end_value`` and save a proof.

Returns a result dict with at least ``status`` (PASS/FAIL/NOT EXTRACTABLE), ``page``,
``proof_file``, ``boxes``, and ``reason``.
    """
    try:
        if proof_path.exists():
            proof_path.unlink()
    except Exception:
        pass

    if not pdf_path or not Path(pdf_path).exists():
        return {"status": "FAIL", "method": "PDF", "reason": "PDF file not found", "page": None, "proof_file": None, "boxes": []}

    fye_date = auditfye_parse_fye_date(fy_end_value)
    if not fye_date:
        return {"status": "FAIL", "method": "FYE_INPUT", "reason": f"Could not parse FYE input: {fy_end_value}", "page": None, "proof_file": None, "boxes": []}

    with open_pdf_maybe_shared(str(pdf_path)) as pdf:
        anchor = auditfye_find_main_audit_report(pdf, fye_date=fye_date, audit_res=audit_res, max_scan_pages=max_scan_pages, window_pages=8)
        if not anchor:
            return {"status": "FAIL", "method": "AUDIT_REPORT_DETECTION", "reason": "Main financial-statement audit report page not found", "page": None, "proof_file": None, "boxes": []}

        for page_num, page in anchor.get("pages", []):
            hit = auditfye_find_on_page(page, fye_date)
            if hit:
                date_boxes = [hit["date_box"]] if hit.get("date_box") else []
                context_boxes = [hit["context_box"]] if hit.get("context_box") else []
                proof_file = None
                if date_boxes and auditfye_save_final_proof(page, date_boxes, context_boxes, proof_path):
                    proof_file = proof_path
                return {
                    "status": "PASS",
                    "method": "AUDIT_REPORT_FYE",
                    "reason": "Expected FYE found in/near the audit report scope sentence",
                    "page": page_num,
                    "proof_file": proof_file,
                    "boxes": date_boxes + context_boxes,
                    "audit_report_page": anchor.get("page_num"),
                    "audit_report_score": anchor.get("score"),
                    "audit_report_source": anchor.get("source"),
                    "audit_report_signals": anchor.get("signals"),
                    "fye_date": fye_date.isoformat(),
                    "matched_scope_sentence": hit.get("scope_sentence"),
                    "date_tokens": hit.get("date_tokens"),
                    "context_ok": hit.get("context_ok"),
                }

        # Fallback scan from the detected audit report page; still limited by max_scan_pages.
        start_idx = anchor.get("page_index", 0)
        limit = min(len(pdf.pages), max_scan_pages)
        for idx in range(start_idx, limit):
            txt = cached_page_text(pdf.pages[idx]) or ""
            if auditfye_is_non_financial_audit_report(txt) and idx > start_idx + 10:
                continue
            hit = auditfye_find_on_page(pdf.pages[idx], fye_date)
            if hit:
                date_boxes = [hit["date_box"]] if hit.get("date_box") else []
                context_boxes = [hit["context_box"]] if hit.get("context_box") else []
                proof_file = None
                if date_boxes and auditfye_save_final_proof(pdf.pages[idx], date_boxes, context_boxes, proof_path):
                    proof_file = proof_path
                return {
                    "status": "PASS",
                    "method": "AUDIT_REPORT_FYE_FALLBACK_SCAN",
                    "reason": "Expected FYE found in/near audit-scope sentence by fallback scan",
                    "page": idx + 1,
                    "proof_file": proof_file,
                    "boxes": date_boxes + context_boxes,
                    "audit_report_page": anchor.get("page_num"),
                    "audit_report_score": anchor.get("score"),
                    "audit_report_source": anchor.get("source"),
                    "audit_report_signals": anchor.get("signals"),
                    "fye_date": fye_date.isoformat(),
                    "matched_scope_sentence": hit.get("scope_sentence"),
                    "date_tokens": hit.get("date_tokens"),
                    "context_ok": hit.get("context_ok"),
                }

        return {
            "status": "FAIL",
            "method": "AUDIT_REPORT_FYE",
            "reason": "Main audit report found but expected FYE not found in audit-scope context",
            "page": anchor.get("page_num"),
            "proof_file": None,
            "boxes": [],
            "audit_report_page": anchor.get("page_num"),
            "audit_report_score": anchor.get("score"),
            "audit_report_source": anchor.get("source"),
            "audit_report_signals": anchor.get("signals"),
            "fye_date": fye_date.isoformat(),
        }




# VALIDATION: Validate that the audit report refers to the same FY End Date and return PASS/FAIL.
def write_audit_report_fye_validation(
    processing_id: int,
    fy_end_value,
    pdf_path: Optional[Path],
    audit_res: Optional[dict] = None,
    show_popup: bool = True
) -> dict:
    """
    Audit Report FYE with Report FYE.

    Integrated Audit FYE standalone logic only:
    - finds the main financial-statement audit report;
    - skips TOC/divider/compliance/internal-control/Uniform Guidance reports;
    - validates expected FYE inside/near the "We have audited" audit-scope sentence;
    - if Audit Report validation already passed, checks that page/window first for speed;
    - falls back to a deeper scan only when needed.

    - Does NOT write to parquet (parent run_validations_for_row batches the write)
    """
    if not pdf_path or not Path(pdf_path).exists():
        return {
            "status": "FAIL",
            "page": None,
            "proof_file": None,
            "boxes": [],
            "title": f"ProcessingId {processing_id} | Audit Report FYE = FAIL (PDF missing)",
            "method": "PDF",
            "reason": "PDF file not found",
        }

    proof_file = PROOF_DIR_AUDIT_FYE / f"Row{processing_id}_AUDIT_FYE_PASS.png"

    try:
        result = auditfye_validate_integrated(
            pdf_path=Path(pdf_path),
            fy_end_value=fy_end_value,
            proof_path=proof_file,
            audit_res=audit_res,
            max_scan_pages=300,
        )

        status = result.get("status", "FAIL")

        page_no = result.get("page")
        boxes = result.get("boxes", []) or []
        proof = result.get("proof_file")
        method = result.get("method")
        reason = result.get("reason", "")
        fye_date = result.get("fye_date")

        title = f"ProcessingId {processing_id} | Audit Report FYE = {status}"
        if fye_date:
            title += f" | FYE {fye_date}"
        if page_no:
            title += f" | Page {page_no}"

        if show_popup and str(status).upper().startswith("PASS") and proof:
            show_proof_popup(Path(proof), title, seconds=PROOF_POPUP_SECONDS)

        return {
            "status": status,
            "page": page_no,
            "proof_file": Path(proof) if proof else None,
            "boxes": boxes,
            "title": title,
            "method": method,
            "reason": reason,
            "fye_date": fye_date,
            "audit_report_page": result.get("audit_report_page"),
            "audit_report_score": result.get("audit_report_score"),
            "audit_report_source": result.get("audit_report_source"),
            "matched_scope_sentence": result.get("matched_scope_sentence"),
            "context_ok": result.get("context_ok"),
        }

    except Exception as e:
        return {
            "status": "FAIL",
            "page": None,
            "proof_file": None,
            "boxes": [],
            "title": f"ProcessingId {processing_id} | Audit Report FYE = FAIL ({e})",
            "method": "EXCEPTION",
            "reason": str(e),
        }

# =========================================================
# AUDITOR SIGNATURE / AUDITOR FIRM NAME VALIDATION v2 HELPERS
# =========================================================
# Scope: Column I only.
# Integrated from latest Signature.py standalone.
# Existing signature workflow is preserved:
#   1) Find opinion/audit anchor.
#   2) Search next 10 pages after anchor.
#   3) If not found, use full-PDF fallback.
# Enhancement:
#   - When searching AUDITOR'S FIRM NAME, create break variants.
#     Example: "Mauldin &amp; Jenkins, CPAs and Advisors"
#       -> "Mauldin & Jenkins, CPAs and Advisors"
#       -> "Mauldin & Jenkins, CPAs"
#       -> "Mauldin & Jenkins"
#   - Keeps previous normalization/suffix tolerance.
#   - Full-PDF fallback will NOT pass on the single generic word "State".

SIGNATURE_LEGAL_SUFFIXES = {
    "llc", "llp", "ltd", "limited", "inc", "incorporated", "pllc", "pc", "pa", "pca", "lp", "lllp",
    # split legal forms after punctuation removal
    "p", "c", "a",
}

SIGNATURE_DESCRIPTOR_TOKENS = {
    "cpa", "cpas", "accountant", "accountants", "accounting",
    "advisor", "advisors", "advisory",
    "consultant", "consultants", "consulting",
    "assurance", "audit", "audits", "auditor", "auditors",
    "tax", "business", "certified", "public", "professional", "services", "service", "firm",
}

SIGNATURE_CONNECTORS = {"and", "&"}
SIGNATURE_STOPWORDS = {"the", "a", "an", "of", "for", "to", "in", "on", "at", "by", "with", "company", "co"}
SIGNATURE_GENERIC_SINGLE_TOKEN_EXCEPTIONS = {"state"}


def _sig_basic_tokens(s: str) -> List[str]:
    """Tokenize ``s`` into basic lowercase alphanumeric tokens for auditor-name matching."""
    s = normalize_pdf_text(s or "").lower()
    s = s.replace("&amp;", "&")
    s = s.replace("&", " and ")
    s = re.sub(r"[^a-z0-9]+", " ", s)
    return [t for t in re.sub(r"\s+", " ", s).strip().split() if t]


def _sig_match_tokens(s: str, remove_descriptors: bool = False, remove_suffix: bool = False) -> List[str]:
    """Tokenize ``s`` for auditor matching, optionally removing descriptor words and legal suffixes."""
    toks = _sig_basic_tokens(s)
    out = []
    for t in toks:
        if t in SIGNATURE_CONNECTORS:
            continue
        if t in SIGNATURE_STOPWORDS:
            continue
        if remove_suffix and t in SIGNATURE_LEGAL_SUFFIXES:
            continue
        if remove_descriptors and t in SIGNATURE_DESCRIPTOR_TOKENS:
            continue
        out.append(t)
    return out


def _sig_core_tokens(s: str) -> List[str]:
    """Return the core distinctive tokens of ``s`` (descriptors and legal suffixes removed)."""
    return _sig_match_tokens(s, remove_descriptors=True, remove_suffix=True)


def _sig_is_unsafe_generic_single_token(tokens: List[str]) -> bool:
    """Do not allow a full-PDF fallback PASS based only on a generic word like 'State'."""
    return len(tokens or []) == 1 and (tokens[0] or "").lower() in SIGNATURE_GENERIC_SINGLE_TOKEN_EXCEPTIONS


# VALIDATION: Generate auditor firm-name break variants.
def _firm_name_variants(firm_name: str) -> List[str]:
    """
    Ordered auditor firm-name variants.

    Keeps previous normalization/suffix behavior and adds requested break logic.

    Example:
      Mauldin &amp; Jenkins, CPAs and Advisors
        -> Mauldin & Jenkins, CPAs and Advisors
        -> Mauldin & Jenkins, CPAs
        -> Mauldin & Jenkins
        -> Mauldin Jenkins
    """
    raw = normalize_pdf_text(firm_name or "").strip()
    if not raw:
        return []

    variants: List[str] = []

    def add(v: str):
        v = normalize_pdf_text(v or "")
        v = re.sub(r"\s+", " ", v).strip(" ,;:-")
        if v and v not in variants:
            variants.append(v)

    add(raw)
    add(raw.replace("&amp;", "&"))

    # Break comma right side one trailing word at a time.
    if "," in raw:
        left, right = raw.split(",", 1)
        left = left.strip()
        right_tokens = right.strip().split()
        for keep in range(len(right_tokens) - 1, 0, -1):
            add(left + ", " + " ".join(right_tokens[:keep]))
        add(left)

    # Phrase-level removals, even if comma is absent.
    phrase_patterns = [
        r"\bCPAs?\s+and\s+Advisors?\b",
        r"\bCertified\s+Public\s+Accountants?\s+and\s+Advisors?\b",
        r"\band\s+Advisors?\b",
        r"\bCPAs?\b",
        r"\bCertified\s+Public\s+Accountants?\b",
        r"\bAdvisors?\b",
        r"\bLLC\b", r"\bLLP\b", r"\bPLLC\b", r"\bLTD\b", r"\bLimited\b", r"\bInc\b", r"\bP\.C\.\b", r"\bP\.A\.\b",
    ]
    for pat in phrase_patterns:
        add(re.sub(pat, " ", raw, flags=re.I))

    # Previous style punctuation/and/compact variants.
    for v in list(variants):
        v_no_punct = re.sub(r"[^A-Za-z0-9 &]+", " ", v)
        v_no_punct = re.sub(r"\s+", " ", v_no_punct).strip()
        add(v_no_punct)
        add(v_no_punct.replace("&", " and "))
        add(re.sub(r"\band\b", " ", v_no_punct.replace("&", " and "), flags=re.I))

    core = " ".join(_sig_core_tokens(raw))
    if core:
        add(core)

    for v in list(variants):
        compact = re.sub(r"\s+", "", v)
        add(compact)

    return variants


# VALIDATION: Convert an auditor firm name into tokens so matching works across formatting differences.
def _auditor_tokens(s: str, remove_suffix: bool = False, remove_descriptors: bool = False) -> List[str]:
    """
    Normalize text into tokens for auditor matching.
    Rules:
      - '&amp;' -> '&'
      - '&' -> 'and' -> connector removed
      - punctuation -> space
      - optional legal suffix removal
      - optional descriptor removal such as CPAs / Advisors
    """
    return _sig_match_tokens(s, remove_descriptors=remove_descriptors, remove_suffix=remove_suffix)


def _auditor_token_variants(firm_name: str) -> List[List[str]]:
    """Return ordered name variants of ``firm_name`` (ampersand/and, punctuation, suffix, spacing)."""
    token_variants: List[List[str]] = []

    def add_tokens(toks: List[str]):
        toks = [t for t in toks if t]
        if not toks:
            return
        # Avoid weak single-token matches unless the token is long and is not generic.
        if len(toks) == 1:
            if toks[0].lower() in SIGNATURE_GENERIC_SINGLE_TOKEN_EXCEPTIONS:
                return
            if len(toks[0]) < 6:
                return
        if toks not in token_variants:
            token_variants.append(toks)

    for v in _firm_name_variants(firm_name):
        add_tokens(_auditor_tokens(v, remove_suffix=False, remove_descriptors=False))
        add_tokens(_auditor_tokens(v, remove_suffix=True, remove_descriptors=False))
        add_tokens(_auditor_tokens(v, remove_suffix=True, remove_descriptors=True))

    # Prefer longer / more specific variants first.
    token_variants.sort(key=lambda x: (len(x), len("".join(x))), reverse=True)
    return token_variants


def _contains_token_sequence(page_tokens: List[str], target_tokens: List[str]) -> Optional[Tuple[int, int]]:
    """Return ``True`` if ``target_tokens`` appears as a contiguous sub-sequence of ``page_tokens``."""
    if not page_tokens or not target_tokens or len(target_tokens) > len(page_tokens):
        return None
    for i in range(len(page_tokens) - len(target_tokens) + 1):
        if page_tokens[i:i + len(target_tokens)] == target_tokens:
            return i, i + len(target_tokens) - 1
    return None


# VALIDATION: Check if the auditor firm name appears in text using flexible token rules.
def _auditor_match_in_text(auditor_name: str, page_text: str, allow_generic_single_word: bool = False) -> bool:
    """
    True if auditor_name matches inside page_text using controlled break variants.
    Full-PDF fallback should call with allow_generic_single_word=False so 'State' alone cannot pass.
    """
    page_streams = [
        _auditor_tokens(page_text, remove_suffix=False, remove_descriptors=False),
        _auditor_tokens(page_text, remove_suffix=True, remove_descriptors=False),
        _auditor_tokens(page_text, remove_suffix=True, remove_descriptors=True),
    ]

    for target in _auditor_token_variants(auditor_name):
        if _sig_is_unsafe_generic_single_token(target) and not allow_generic_single_word:
            continue
        for page_tokens in page_streams:
            if _contains_token_sequence(page_tokens, target):
                return True
            # compact fallback for spacing issues like BerganKDV vs Bergan KDV
            target_compact = "".join(target)
            page_compact = "".join(page_tokens)
            if target_compact and len(target_compact) >= 6 and target_compact in page_compact:
                return True
    return False


# VALIDATION: Find the bounding box for the auditor firm name on a page for proof image.
def _find_auditor_name_box_on_page(page, auditor_name: str) -> Optional[dict]:
    """
    Find a highlight bbox for auditor name using token-aligned word mapping.
    """
    words = cached_page_words(page, use_text_flow=True)
    if not words:
        return None

    page_tokens: List[str] = []
    token_word_map: List[dict] = []

    for w in words:
        wtoks = _auditor_tokens(w.get("text", ""), remove_suffix=True, remove_descriptors=True)
        for t in wtoks:
            page_tokens.append(t)
            token_word_map.append(w)

    if not page_tokens:
        return None

    for target_tokens in _auditor_token_variants(auditor_name):
        reduced_target = [t for t in target_tokens if t not in SIGNATURE_DESCRIPTOR_TOKENS and t not in SIGNATURE_LEGAL_SUFFIXES and t not in SIGNATURE_CONNECTORS]
        if not reduced_target or _sig_is_unsafe_generic_single_token(reduced_target):
            continue

        s, e = find_token_window(page_tokens, reduced_target)
        if s is None or e is None:
            continue

        matched_words = token_word_map[s:e + 1]
        return {
            "x0": min(w["x0"] for w in matched_words),
            "top": min(w["top"] for w in matched_words),
            "x1": max(w["x1"] for w in matched_words),
            "bottom": max(w["bottom"] for w in matched_words),
        }

    return None


# VALIDATION: If signature check fails, search the whole PDF for the auditor firm name as fallback.
def find_auditor_firm_anywhere_in_pdf(pdf, firm_name: str) -> Optional[Tuple[int, Optional[dict]]]:
    """
    Full-PDF fallback search using same matching logic.
    Important exception: a match based only on the single word 'State' is ignored.
    """
    firm_name = (firm_name or "").strip()
    if not firm_name:
        return None

    # Pre-check: if all possible variants reduce to only 'state', do not run fallback.
    safe_variants = [v for v in _auditor_token_variants(firm_name) if not _sig_is_unsafe_generic_single_token(v)]
    if not safe_variants:
        return None

    for idx, page in enumerate(pdf.pages):
        page_num = idx + 1
        txt = cached_page_text(page) or ""
        if not txt.strip():
            continue

        if not _auditor_match_in_text(firm_name, txt, allow_generic_single_word=False):
            continue

        box = None
        try:
            box = _find_auditor_name_box_on_page(page, firm_name)
        except Exception:
            box = None

        return (page_num, box)

    return None




# VALIDATION: Last-resort signature fallback - auditor short name must appear in an audit-context sentence.
def _signature_sentence_windows(text: str) -> List[str]:
    """Split page text into sentence-like windows, preserving enough context for OCR/PDF spacing issues."""
    txt = normalize_pdf_text(text or "")
    if not txt:
        return []
    parts = re.split(r"(?<=[.!?])\s+", txt)
    out = []
    for p in parts:
        p = normalize_pdf_text(p)
        if p:
            out.append(p)
    # Fallback for pages with poor punctuation: add rolling chunks around audit words.
    if len(out) <= 2 and re.search(r"\baudit\w*\b", txt, re.I):
        words = txt.split()
        for i, w in enumerate(words):
            if re.search(r"\baudit\w*\b", w, re.I):
                out.append(" ".join(words[max(0, i - 25): min(len(words), i + 35)]))
    return out


def _auditor_short_context_variants(firm_name: str) -> List[List[str]]:
    """
    Build short auditor variants for last fallback only.
    Example:
      Rehmann Robson LLC -> [rehmann robson], [rehmann]
      Thompson, Price, Scott, Adams & Co. -> [thompson price scott adams], [thompson price], [thompson]
    Single-token variants are allowed only if non-generic and length >= 6.
    """
    core = [t for t in _sig_core_tokens(firm_name) if t]
    variants: List[List[str]] = []

    def add(toks: List[str]):
        toks = [t for t in toks if t]
        if not toks:
            return
        if len(toks) == 1:
            if toks[0].lower() in SIGNATURE_GENERIC_SINGLE_TOKEN_EXCEPTIONS:
                return
            if len(toks[0]) < 6:
                return
        if toks not in variants:
            variants.append(toks)

    add(core)
    if len(core) >= 2:
        add(core[:2])
    if len(core) >= 1:
        add(core[:1])

    # Also consider comma-left firm phrase before descriptors/suffixes.
    raw = normalize_pdf_text(firm_name or "")
    if "," in raw:
        left_core = [t for t in _sig_core_tokens(raw.split(",", 1)[0]) if t]
        add(left_core)
        if len(left_core) >= 2:
            add(left_core[:2])
        if len(left_core) >= 1:
            add(left_core[:1])

    return variants


def _find_short_auditor_box_on_page(page, short_tokens: List[str]) -> Optional[dict]:
    """Find a short auditor-name token sequence on a page and return its bounding box, or ``None``."""
    if not short_tokens or _sig_is_unsafe_generic_single_token(short_tokens):
        return None
    words = cached_page_words(page, use_text_flow=True)
    if not words:
        return None

    page_tokens: List[str] = []
    token_word_map: List[dict] = []
    for w in words:
        wtoks = _auditor_tokens(w.get("text", ""), remove_suffix=True, remove_descriptors=True)
        for t in wtoks:
            page_tokens.append(t)
            token_word_map.append(w)

    s, e = find_token_window(page_tokens, short_tokens)
    if s is None or e is None:
        return None
    matched_words = token_word_map[s:e + 1]
    if not matched_words:
        return None
    return {
        "x0": min(w["x0"] for w in matched_words),
        "top": min(w["top"] for w in matched_words),
        "x1": max(w["x1"] for w in matched_words),
        "bottom": max(w["bottom"] for w in matched_words),
    }


def find_auditor_firm_in_audit_sentence_fallback(pdf, firm_name: str) -> Optional[Tuple[int, Optional[dict], str]]:
    """
    Last fallback only:
      PASS if a safe auditor short-name variant appears in the same sentence/window as an audit word.

    This catches cases like:
      "... financial statements have been audited by Rehmann, a firm of licensed certified public accountants."
    for FAC auditor name "Rehmann Robson LLC".

    Safety:
      - Does not allow a single generic token like "State".
      - Requires an audit-context word: audit/audited/auditor/auditors/auditing/etc.
    """
    firm_name = (firm_name or "").strip()
    if not firm_name:
        return None

    short_variants = _auditor_short_context_variants(firm_name)
    if not short_variants:
        return None

    audit_word_re = re.compile(r"\baudit\w*\b", re.I)

    for idx, page in enumerate(pdf.pages):
        txt = cached_page_text(page) or ""
        if not txt.strip() or not audit_word_re.search(txt):
            continue

        for sent in _signature_sentence_windows(txt):
            if not audit_word_re.search(sent):
                continue
            sent_tokens = _auditor_tokens(sent, remove_suffix=True, remove_descriptors=True)
            sent_compact = "".join(sent_tokens)
            for short_tokens in short_variants:
                if _sig_is_unsafe_generic_single_token(short_tokens):
                    continue
                hit = find_token_window(sent_tokens, short_tokens)
                compact_hit = "".join(short_tokens) in sent_compact if len("".join(short_tokens)) >= 6 else False
                if hit[0] is not None or compact_hit:
                    box = None
                    try:
                        box = _find_short_auditor_box_on_page(page, short_tokens)
                    except Exception:
                        box = None
                    return (idx + 1, box, sent)

    return None

# VALIDATION: Validate that the auditor firm name/signature appears after the audit report and write PASS/FAIL.
# VALIDATION: Standalone Signature.py audit-anchor fallback helpers.
# Scope: Column I only. These helpers are intentionally prefixed with _signature_
# so no existing audit-report/opinion/FYE logic is changed.
def _signature_compact_text(text: str) -> str:
    """Return ``text`` lowercased with whitespace removed (signature matcher)."""
    return re.sub(r"[^a-z0-9]+", "", normalize_pdf_text(text or "").lower())


def _signature_top_text(page_text: str, lines: int = 45) -> str:
    """Return normalized lowercase text from the top ``lines`` lines of a page."""
    return normalize_pdf_text(
        " ".join([ln.strip() for ln in (page_text or "").splitlines() if ln.strip()][:lines])
    )


def _signature_is_toc_page(page_text: str) -> bool:
    """Return ``True`` if the page is a Table of Contents page (signature detector)."""
    t = _signature_top_text(page_text, 25).lower()
    c = _signature_compact_text(t)
    if "wehaveaudited" in c or "inouropinion" in c or "presentfairly" in c:
        return False
    return "tableofcontents" in c or bool(re.search(r"^\s*table\s+of\s+contents\b", t, re.I))


def _signature_is_non_financial_report(page_text: str) -> bool:
    """Return ``True`` if the page is a non-financial report (compliance/internal control) for signature scanning."""
    t = _signature_top_text(page_text, 60).lower()
    patterns = [
        r"report\s+on\s+internal\s+control\s+over\s+financial\s+reporting",
        r"report\s+on\s+compliance\s+for\s+each\s+major",
        r"internal\s+control\s+over\s+compliance",
        r"uniform\s+guidance",
        r"schedule\s+of\s+expenditures\s+of\s+federal\s+awards",
    ]
    return any(re.search(p, t, re.I) for p in patterns)


def _signature_audit_anchor_score(page_text: str) -> Tuple[float, dict]:
    """
    Standalone Signature.py-style anchor scoring for the main independent audit/opinion report.
    Used only as a Column-I fallback when the main opinion-anchor detector returns None.
    """
    if not page_text or not page_text.strip():
        return 0.0, {}
    if _signature_is_toc_page(page_text) or _signature_is_non_financial_report(page_text):
        return 0.0, {"rejected": True}

    txt = normalize_pdf_text(page_text or "")
    low = txt.lower()
    top = _signature_top_text(page_text, 45).lower()
    comp = _signature_compact_text(txt)
    top_comp = _signature_compact_text(top)

    sig = {
        "independent_auditor_heading": bool(re.search(r"\bindependent\s+auditors?'?\s+report\b", top, re.I))
                                      or "independentauditorsreport" in top_comp
                                      or "independentauditorreport" in top_comp,
        "report_on_financial_statements": bool(re.search(r"\breport\s+on\s+the\s+financial\s+statements\b", low, re.I)),
        "opinions_heading": bool(re.search(r"\bopinions?\b", top, re.I)),
        "we_have_audited": "wehaveaudited" in comp or bool(re.search(r"\bwe\s+have\s+audited\b", low, re.I)),
        "in_our_opinion": "inouropinion" in comp or bool(re.search(r"\bin\s+our\s+opinion\b", low, re.I)),
        "present_fairly": "presentfairly" in comp or bool(re.search(r"\bpresent\s+fairly\b", low, re.I)),
        "basis_for_opinion": bool(re.search(r"\bbasis\s+for\s+opinions?\b", low, re.I)),
    }

    score = 0.0
    if sig["independent_auditor_heading"]:
        score += 0.32
    if sig["report_on_financial_statements"]:
        score += 0.12
    if sig["opinions_heading"]:
        score += 0.08
    if sig["we_have_audited"]:
        score += 0.18
    if sig["in_our_opinion"]:
        score += 0.14
    if sig["present_fairly"]:
        score += 0.08
    if sig["basis_for_opinion"]:
        score += 0.08
    return min(score, 1.0), sig


def _signature_find_audit_anchor(pdf, max_scan_pages: int = 120) -> Optional[int]:
    """Return 1-based audit anchor page using standalone Signature.py-style scoring."""
    best = None
    limit = min(len(pdf.pages), max_scan_pages)
    for idx in range(limit):
        txt = cached_page_text(pdf.pages[idx]) or ""
        score, sig = _signature_audit_anchor_score(txt)
        if score <= 0:
            continue
        cand = (score, idx + 1, sig)
        if best is None or cand[0] > best[0]:
            best = cand
        if score >= 0.80:
            break
    return best[1] if best and best[0] >= 0.45 else None


def _signature_save_auditor_proof_safe(page, box: Optional[dict], proof_file: Path) -> bool:
    """
    Save signature proof without allowing proof-rendering issues to flip a valid PASS to FAIL.
    This matches the standalone Signature.py behavior where proof save is best-effort.
    """
    try:
        proof_file.parent.mkdir(parents=True, exist_ok=True)
        im = page.to_image(resolution=PROOF_IMAGE_RESOLUTION)
        if box:
            im.draw_rect(box, stroke=HIGHLIGHT_STROKE, fill=HIGHLIGHT_FILL_RGBA)
        im.save(str(proof_file))
        return True
    except Exception as e:
        try:
            print(f"[WARN] Signature proof save failed: {e}")
        except Exception:
            pass
        return False


def write_auditor_signature_validation(
    processing_id: int,
    pdf_path: Optional[Path],
    audit_res: Optional[dict] = None,
    auditor_firm_name: str = ""
) -> dict:
    """
    Auditor Name / Signature validation.

    Aligned with the standalone Signature.py flow:
      1) Try the existing main-code opinion anchor first.
      2) If that anchor is not found, use the standalone Signature.py-style audit anchor.
      3) If an anchor is found, search the next 10 pages after the anchor.
      4) Whether or not an anchor is found, run the full-PDF firm-name fallback.
      5) If still not found, run the short auditor-name + audit-context sentence fallback.

    - Does NOT write to parquet (parent run_validations_for_row batches the write)
    """
    if not pdf_path or not Path(pdf_path).exists():
        return {"status": "FAIL", "page": None, "proof_file": None, "reason": "PDF missing"}

    firm_name = (auditor_firm_name or "").strip()
    if not firm_name:
        return {"status": "FAIL", "page": None, "proof_file": None, "reason": "AUDITOR'S FIRM NAME is blank"}

    try:
        with open_pdf_maybe_shared(pdf_path) as pdf:
            if not pdf.pages:
                return {"status": "NOT EXTRACTABLE", "page": None, "proof_file": None, "reason": "PDF has no pages"}

            anchor_page_num = None

            # Existing main-code opinion anchor remains primary.
            try:
                anchor = find_opinion_anchor_in_pdf(pdf, max_scan_pages=120)
                if anchor:
                    anchor_page_num = int(anchor.get("opinion_page") or 0) or int(anchor.get("heading_page") or 0) or None
            except Exception:
                anchor_page_num = None

            # Standalone Signature.py audit-anchor fallback.
            if not anchor_page_num:
                try:
                    anchor_page_num = _signature_find_audit_anchor(pdf, max_scan_pages=120)
                except Exception:
                    anchor_page_num = None

            # Search next 10 pages after anchor, same as standalone Signature.py.
            if anchor_page_num:
                start_idx = anchor_page_num  # 0-based index of page AFTER 1-based anchor page
                end_idx = min(start_idx + 10, len(pdf.pages))
                for idx in range(start_idx, end_idx):
                    page_num = idx + 1
                    page = pdf.pages[idx]
                    txt = cached_page_text(page) or ""
                    if not txt.strip():
                        continue
                    if not _auditor_match_in_text(firm_name, txt, allow_generic_single_word=False):
                        continue
                    try:
                        box = _find_auditor_name_box_on_page(page, firm_name)
                    except Exception:
                        box = None
                    proof_file = PROOF_DIR_SIGNATURE / f"Row{processing_id}_SIGNATURE_PASS_Page{page_num}.png"
                    proof_ok = _signature_save_auditor_proof_safe(page, box, proof_file)
                    if proof_ok and SHOW_PROOF_POPUP:
                        show_proof_popup(proof_file, f"ProcessingId {processing_id} | Signature = PASS | Page {page_num}", seconds=PROOF_POPUP_SECONDS)
                    return {"status": "PASS", "page": page_num, "anchor_page": anchor_page_num, "proof_file": str(proof_file) if proof_ok else None, "method": "ANCHOR_NEXT_10"}

            # Full-PDF fallback must run even when no anchor is found, matching standalone Signature.py.
            hit = find_auditor_firm_anywhere_in_pdf(pdf, firm_name)
            if hit:
                page_num, box = hit
                page = pdf.pages[page_num - 1]
                proof_file = PROOF_DIR_SIGNATURE / f"Row{processing_id}_SIGNATURE_PASS_Fallback_Page{page_num}.png"
                proof_ok = _signature_save_auditor_proof_safe(page, box, proof_file)
                if proof_ok and SHOW_PROOF_POPUP:
                    show_proof_popup(proof_file, f"ProcessingId {processing_id} | Signature = PASS (Fallback) | Page {page_num}", seconds=PROOF_POPUP_SECONDS)
                return {"status": "PASS", "page": page_num, "anchor_page": anchor_page_num, "proof_file": str(proof_file) if proof_ok else None, "method": "FULL_PDF_FALLBACK"}

            # Last fallback: short auditor firm name in a sentence/window containing an audit word.
            audit_sentence_hit = find_auditor_firm_in_audit_sentence_fallback(pdf, firm_name)
            if audit_sentence_hit:
                page_num, box, sentence = audit_sentence_hit
                page = pdf.pages[page_num - 1]
                proof_file = PROOF_DIR_SIGNATURE / f"Row{processing_id}_SIGNATURE_PASS_AuditSentenceFallback_Page{page_num}.png"
                proof_ok = _signature_save_auditor_proof_safe(page, box, proof_file)
                if proof_ok and SHOW_PROOF_POPUP:
                    show_proof_popup(proof_file, f"ProcessingId {processing_id} | Signature = PASS (Audit Sentence Fallback) | Page {page_num}", seconds=PROOF_POPUP_SECONDS)
                return {"status": "PASS", "page": page_num, "anchor_page": anchor_page_num, "proof_file": str(proof_file) if proof_ok else None, "method": "AUDIT_SENTENCE_FALLBACK", "matched_sentence": sentence[:500]}

            return {"status": "FAIL", "page": anchor_page_num, "anchor_page": anchor_page_num, "proof_file": None, "reason": "Auditor firm name not found"}

    except Exception as e:
        return {"status": "FAIL", "page": None, "proof_file": None, "reason": str(e)}
def find_fye_box_on_page_anywhere(page, fy_end_value) -> Optional[dict]:
    """
    Find FY end date on a page WITHOUT strict context.
    Used for Financial Statements where header usually contains only the date.
    Returns bbox dict if found else None.
    """
    d = parse_fy_end_date(fy_end_value)
    if not d:
        return None

    words = cached_page_words(page, use_text_flow=True)
    if not words:
        return None

    page_tokens = []
    token_word_map = []

    for w in words:
        wtoks = tokenize(w.get("text", ""))
        for t in wtoks:
            page_tokens.append(t)
            token_word_map.append(w)

    for cand in build_fye_token_candidates(d):
        s, e = find_token_window(page_tokens, cand)
        if s is None or e is None:
            continue

        matched_words = token_word_map[s:e]
        x0 = min(w["x0"] for w in matched_words)
        top = min(w["top"] for w in matched_words)
        x1 = max(w["x1"] for w in matched_words)
        bottom = max(w["bottom"] for w in matched_words)
        return {"x0": x0, "top": top, "x1": x1, "bottom": bottom}

    return None

# VALIDATION: Skip pages that look like reconciliation/other excluded financial statement pages.
def is_excluded_fin_stmt_page(page_text: str) -> bool:
    """Return ``True`` if the page should be excluded from financial-statement detection (TOC/MD&A/notes/reconciliation/etc.)."""
    if not page_text:
        return False

    t = normalize_pdf_text(page_text).lower()

    # remove a very common boilerplate so it doesn't trigger false excludes
    t = re.sub(r"\bsee\s+notes\s+to\s+(?:the\s+)?basic\s+financial\s+statements\b", "", t, flags=re.I)

    excluded_re = re.compile(
        r"\b("
        r"contents?|table\s+of\s+contents?|"
        r"foreword|introduction|"
        r"md\s*&\s*a|management'?s?\s+discussion\s*(?:&|and)\s*analysis|"
        r"budget|supplement(?:al|ary)?|statistical|"
        r"reconciliation"
        r")\b",
        re.I
    )

    return bool(excluded_re.search(t))

# ---------------------------------------------------------
# EXCLUDED HEADER KEYWORDS (applies to ALL statement headers)
# ---------------------------------------------------------

EXCLUDED_HEADER_KEYWORDS = [
    "content",
    "management discussion",
    "managementdiscussion",
    "management's discussion",
    "management'sdiscussion",
    "internal",
    "nonmajor",
    "non major",
    "non-major",
    "budgetary",
    "statistical",
    "admin",
    "supplement",
]

def header_has_excluded_keywords(header_text: str) -> bool:
    """Return True if header_text contains any excluded keyword.

    Handles spaced/unspaced variants by also checking a compact form with spaces/apostrophes removed.

    Controlled exception:
    - Do NOT treat the word "internal" as excluded when it appears in the valid financial-statement
      title phrase "Internal Service Fund" / "Internal Service Funds". This restores the previous
      proprietary-fund validations while keeping the top-5 excluded keyword workflow intact for
      other uses of "internal" such as internal control/report sections.
    """
    ht = (header_text or "").lower()
    compact = re.sub(r"[\s’']", "", ht)

    internal_service_fund_ok = bool(re.search(r"\binternal\s+service\s+funds?\b", ht, flags=re.I))

    for kw in EXCLUDED_HEADER_KEYWORDS:
        k = (kw or "").lower()
        if not k:
            continue

        # Allow valid proprietary/internal service fund statement titles.
        if k == "internal" and internal_service_fund_ok:
            continue

        if k in ht:
            return True
        k_compact = re.sub(r"[\s’']", "", k)
        if k_compact and k_compact in compact:
            return True

    return False

def excluded_keywords_above_statement_and_fye(page, heading_box: dict, fye_box: dict, max_lines: int = 12) -> bool:
    """
    WORKFLOW RULE (Top-5-title-lines gate):
    Step 1: Read the visually top lines of the page.
    Step 2: If a statement date / FYE line appears in those top lines, inspect only the title block
            up to and including that date/FYE line. This avoids rejecting valid statements because
            fund-column labels such as "Nonmajor" appear immediately below the title block.
    Step 3: If no date/FYE line is found, fall back to the top 5 visual lines.
    Step 4: If any excluded header keyword appears in the inspected title area, reject the page.

    heading_box/fye_box/max_lines are kept only for compatibility with existing calls.
    """
    try:
        top_words = _top_words_by_lines(page, max_lines=max(12, max_lines))
        if not top_words:
            return False

        # Build visual lines from the top words.
        line_map = {}
        for w in top_words:
            key = round(float(w.get("top", 0)), 0)
            line_map.setdefault(key, []).append(w)

        visual_lines = []
        for k in sorted(line_map.keys()):
            ws = sorted(line_map[k], key=lambda x: float(x.get("x0", 0)))
            line_text = " ".join(
                (w.get("text", "") or "").strip()
                for w in ws
                if (w.get("text", "") or "").strip()
            )
            if line_text.strip():
                visual_lines.append(line_text.strip())

        if not visual_lines:
            return False

        month_pattern = (
            r"january|jan|february|feb|march|mar|april|apr|may|"
            r"june|jun|july|jul|august|aug|september|sep|sept|"
            r"october|oct|november|nov|december|dec"
        )
        date_line_re = re.compile(
            rf"(?i)(?:fiscal\s+year\s+ended\s+)?(?:{month_pattern})\s+\d{{1,2}},?\s+\d{{4}}|\d{{1,2}}[/-]\d{{1,2}}[/-]\d{{2,4}}"
        )

        selected_lines = []
        found_date_line = False
        for ln in visual_lines:
            selected_lines.append(ln)
            if date_line_re.search(ln):
                found_date_line = True
                break

        if found_date_line:
            header_text = " ".join(selected_lines)
        else:
            header_text = " ".join(visual_lines[:5])

        return header_has_excluded_keywords(header_text)

    except Exception:
        return False

# VALIDATION: Locate the "Statement of Net Position" heading near the top of a page.
def find_net_position_heading_box_top_lines(page, max_lines: int = 12) -> Optional[dict]:
    """
    Find header bbox for either:
      - 'Statement of Net Position'
      - 'Statements of Net Position'
    ONLY within top `max_lines` visual lines.

    Also supports optional '(Deficit)' appearing after the phrase.
    """
    # Try plural first
    box = find_phrase_box_on_page_top_lines(page, "Statements of Net Position", max_lines=max_lines)
    if box:
        return box

    # Try singular
    box = find_phrase_box_on_page_top_lines(page, "Statement of Net Position", max_lines=max_lines)
    if box:
        return box

    # Some PDFs might break words oddly; try looser fallback
    box = find_phrase_box_on_page_top_lines(page, "Net Position", max_lines=max_lines)
    return box

# VALIDATION: Confirm Net Position statement header includes the FY date.
def page_header_contains_net_position_and_fye(page, fy_end_value, max_lines: int = 12) -> bool:
    """
    Header check using TOP VISUAL LINES (same robust method as Statement of Activities):
    - Must contain 'statement of net position' OR 'statements of net position'
    - Must contain FY end date tokens in the same header area
    """
    header_text = top_header_text_from_page(page, max_lines=max_lines)
    if not header_text:
        return False

    # Accept singular or plural
    if ("statement of net position" not in header_text) and ("statements of net position" not in header_text):
        return False

    d = parse_fy_end_date(fy_end_value)
    if not d:
        return False

    header_tokens = tokenize(header_text)

    for cand_tokens in build_fye_token_candidates(d):
        s, e = find_token_window(header_tokens, cand_tokens)
        if s is not None and e is not None:
            return True

    return False

# VALIDATION: Locate the "Statement of Activities" heading near the top of a page.
def find_statement_of_activities_heading_box_top_lines(page, max_lines: int = 12) -> Optional[dict]:
    """
    Find header bbox for 'Statement of Activities' ONLY within top `max_lines` visual lines.
    """
    box = find_phrase_box_on_page_top_lines(page, "Statement of Activities", max_lines=max_lines)
    if box:
        return box

    # fallback: sometimes split words or formatting
    return find_phrase_box_on_page_top_lines(page, "Activities", max_lines=max_lines)

# VALIDATION: Confirm Activities statement header includes the FY date.
def page_header_contains_activities_and_fye(page, fy_end_value, max_lines: int = 12) -> bool:
    """
    Header check using TOP VISUAL LINES (robust for landscape/ACFR tables):
    - Must contain 'statement of activities'
    - Must contain FY end date tokens in same header area
    """
    header_text = top_header_text_from_page(page, max_lines=max_lines)
    if not header_text:
        return False

    if "statement of activities" not in header_text:
        return False

    d = parse_fy_end_date(fy_end_value)
    if not d:
        return False

    header_tokens = tokenize(header_text)

    for cand_tokens in build_fye_token_candidates(d):
        s, e = find_token_window(header_tokens, cand_tokens)
        if s is not None and e is not None:
            return True

    return False

# VALIDATION: Decide if a page is the Statement of Activities (strict detection).
def is_statement_of_activities_page(page_text: str) -> bool:
    """
    STRICT detector for Statement of Activities page.

    Rules:
      - Must NOT be excluded (TOC/MD&A/Notes/Reconciliation etc.)
      - Must contain 'statement of activities' in top ~20 lines (handles running headers)
      - Must contain table-sanity words:
          * 'expenses'
          * ('program revenues' OR 'charges for services')
      - Avoid reconciliation pages explicitly
    """
    if not page_text:
        return False

    # Exclude TOC/MD&A/Notes/etc.
    if is_excluded_fin_stmt_page(page_text):
        return False

    t_all = normalize_pdf_text(page_text).lower()
    top20 = normalize_pdf_text(" ".join(_top_non_empty_lines(page_text, max_lines=20))).lower()

    if "statement of activities" not in top20:
        return False

    # avoid reconciliation pages
    if "reconciliation" in t_all:
        return False

    # sanity checks (similar to your reference code)
    if "expenses" not in t_all:
        return False

    if ("program revenues" not in t_all) and ("charges for services" not in t_all):
        return False

    return True

# VALIDATION: Validate Statement of Activities header + FY date and return PASS/FAIL with proof.
def _orig_write_statement_of_activities_validation(
    processing_id: int,
    fy_end_value,
    pdf_path: Optional[Path],
    show_popup: bool = True
) -> dict:
    """
    Statement of Activities with FYE.

    PASS only if:
      - Candidate page is Statement of Activities (strict detector)
      - Header area (top visual lines) contains:
            * 'Statement of Activities'
            * FY End Date (from fy_end_value)
      - We can locate highlight boxes for both heading and date within top visual lines.

    - Does NOT write to parquet (parent run_validations_for_row batches the write)
    """

    _clear_page_words_cache()
    HEADER_MAX_LINES = 12

    if not pdf_path or not Path(pdf_path).exists():
        return {"status": "FAIL", "page": None, "proof_file": None, "boxes": [], "title": f"ProcessingId {processing_id} | Statement of Activities = FAIL (PDF missing)"}

    try:
        with open_pdf_maybe_shared(pdf_path) as pdf:
            if not pdf.pages:
                return {"status": "NOT EXTRACTABLE", "page": None, "proof_file": None, "boxes": [], "title": f"ProcessingId {processing_id} | Statement of Activities = NOT EXTRACTABLE"}

            first_candidate_page = None

            for idx, page in enumerate(pdf.pages):
                page_num = idx + 1
                txt = page.extract_text() or ""
                if not txt.strip():
                    continue

                # skip TOC pages
                if TOC_STOP_ENABLED and is_table_of_contents_page(txt):
                    continue

                # Candidate statement page?

                # WORKFLOW: reject page first if excluded header keywords appear in TOP 5 lines
                if excluded_keywords_above_statement_and_fye(page, None, None):
                    continue
                if not is_statement_of_activities_page(txt):
                    continue

                if first_candidate_page is None:
                    first_candidate_page = page_num

                # header rule uses TOP VISUAL lines from page words
                if not page_header_contains_activities_and_fye(page, fy_end_value, max_lines=HEADER_MAX_LINES):
                    continue

                # Boxes for proof (within top visual lines)
                heading_box = find_statement_of_activities_heading_box_top_lines(page, max_lines=HEADER_MAX_LINES)
                fye_box = find_fye_box_on_page_top_lines(page, fy_end_value, max_lines=HEADER_MAX_LINES)

                if not heading_box or not fye_box:
                    continue

                # PASS
                proof_file = PROOF_DIR_ACTIVITIES / f"Row{processing_id}_ACTIVITIES_PASS_Page{page_num}.png"
                im = page.to_image(resolution=PROOF_IMAGE_RESOLUTION)

                boxes = []
                im.draw_rect(heading_box, stroke=HIGHLIGHT_STROKE, fill=HIGHLIGHT_FILL_RGBA)
                boxes.append(heading_box)

                im.draw_rect(fye_box, stroke=HIGHLIGHT_STROKE, fill=HIGHLIGHT_FILL_RGBA)
                boxes.append(fye_box)

                im.save(proof_file)

                title = f"ProcessingId {processing_id} | Statement of Activities = PASS | Page {page_num}"
                if show_popup:
                    show_proof_popup(proof_file, title, seconds=PROOF_POPUP_SECONDS)

                return {"status": "PASS", "page": page_num, "proof_file": proof_file, "boxes": boxes, "title": title}

            # FAIL
            if first_candidate_page is None:
                return {"status": "FAIL", "page": None, "proof_file": None, "boxes": [], "title": f"ProcessingId {processing_id} | Statement of Activities = FAIL (Statement not found)"}

            # Save FAIL proof (popup suppressed by your global rule)
            proof_file = PROOF_DIR_ACTIVITIES / f"Row{processing_id}_ACTIVITIES_FAIL_Page{first_candidate_page}.png"
            try:
                im = pdf.pages[first_candidate_page - 1].to_image(resolution=PROOF_IMAGE_RESOLUTION)
                im.save(proof_file)
            except Exception:
                pass

            title = f"ProcessingId {processing_id} | Statement of Activities = FAIL (Header missing Activities or FYE) | Page {first_candidate_page}"
            if show_popup:
                show_proof_popup(proof_file, title, seconds=PROOF_POPUP_SECONDS)

            return {"status": "FAIL", "page": first_candidate_page, "proof_file": proof_file, "boxes": [], "title": title}

    except Exception as e:
        return {"status": "FAIL", "page": None, "proof_file": None, "boxes": [], "title": f"ProcessingId {processing_id} | Statement of Activities = FAIL ({e})"} 


# VALIDATION: Locate Balance Sheet heading near top of page (supports common variants).
def find_balance_sheet_heading_box_top_lines(page, max_lines: int = 12) -> Optional[dict]:
    """
    Find header bbox for Balance Sheet ONLY within top `max_lines` meaningful visual lines.
    Tries several common headings:
      - 'Balance Sheet'
      - 'Balance Sheet - Governmental Funds'
      - 'Balance Sheet Governmental Funds'
    """
    # Try longer/more specific first
    for phrase in [
        "Balance Sheet - Governmental Funds",
        "Balance Sheet Governmental Funds",
        "Balance Sheet",
    ]:
        box = find_phrase_box_on_page_top_lines(page, phrase, max_lines=max_lines)
        if box:
            return box

    return None

# VALIDATION: Confirm Balance Sheet header includes the FY date.
def page_header_contains_balance_sheet_and_fye(page, fy_end_value, max_lines: int = 12) -> bool:
    """
    Header check using TOP VISUAL LINES (robust like Activities/Net Position):
      - Must contain 'balance sheet' (any variant)
      - Must contain FY end date tokens in the same header area
    """
    header_text = top_header_text_from_page(page, max_lines=max_lines)
    if not header_text:
        return False

    # Normalize to be tolerant: "balance-sheet", "balance sheet", etc.
    header_clean = clean_for_contains_match(header_text)

    if "balance sheet" not in header_clean:
        return False

    d = parse_fy_end_date(fy_end_value)
    if not d:
        return False

    header_tokens = tokenize(header_text)

    for cand_tokens in build_fye_token_candidates(d):
        s, e = find_token_window(header_tokens, cand_tokens)
        if s is not None and e is not None:
            return True

    return False

# VALIDATION: Robustly detect Balance Sheet (Governmental Funds) pages.
def is_balance_sheet_governmental_funds_page(page_text: str) -> bool:
    """
    ROBUST detector for Balance Sheet (Governmental Funds).

    Pass if:
      - not excluded
      - contains Balance Sheet context (either 'balance sheet' OR strong balance-sheet structure)
      - contains 'assets' and 'liabilities'
      - contains 'fund balance' or 'fund balances'
    Extra boosters (NOT required):
      - 'governmental funds'
      - 'total liabilities and fund balances'
      - 'total fund balance(s)'
    """
    if not page_text:
        return False

    if is_excluded_fin_stmt_page(page_text):
        return False

    t = normalize_pdf_text(page_text).lower()

    # Exclude reconciliation pages
    if "reconciliation" in t:
        return False

    # Core table sanity
    if ("assets" not in t) or ("liabilities" not in t):
        return False

    if ("fund balance" not in t) and ("fund balances" not in t):
        return False

    # Balance sheet hint (either explicit header phrase or common footer)
    has_balance_sheet_phrase = ("balance sheet" in t)
    has_total_liab_fb = ("total liabilities and fund balances" in t) or ("total liabilities & fund balances" in t)

    # If header phrase not in body text, allow structure-based pass
    if not has_balance_sheet_phrase and not has_total_liab_fb:
        # Some PDFs omit the phrase in extracted text; still accept based on structure.
        # However, require at least one of these stronger indicators:
        if ("total fund balance" not in t) and ("total fund balances" not in t) and ("governmental funds" not in t):
            return False

    # Optional: If you want to exclude Nonmajor always (like your friend's code), keep it:
    # if "nonmajor" in t:
    #     return False

    return True


# VALIDATION: Validate Balance Sheet header + FY date and return PASS/FAIL with proof.
def _orig_write_balance_sheet_validation(
    processing_id: int,
    fy_end_value,
    pdf_path: Optional[Path],
    show_popup: bool = True
) -> dict:
    """
    Balance Sheet with FYE.

    - Does NOT write to parquet (parent run_validations_for_row batches the write)
    """
    HEADER_MAX_LINES = 12

    if not pdf_path or not Path(pdf_path).exists():
        return {"status": "FAIL", "page": None, "proof_file": None, "boxes": [], "title": f"ProcessingId {processing_id} | Balance Sheet = FAIL (PDF missing)"}

    try:
        with open_pdf_maybe_shared(pdf_path) as pdf:
            if not pdf.pages:
                return {"status": "NOT EXTRACTABLE", "page": None, "proof_file": None, "boxes": [], "title": f"ProcessingId {processing_id} | Balance Sheet = NOT EXTRACTABLE"}

            first_candidate_page = None

            for idx, page in enumerate(pdf.pages):
                page_num = idx + 1
                txt = page.extract_text() or ""
                if not txt.strip():
                    continue

                # Skip TOC pages (your global setting)
                if TOC_STOP_ENABLED and is_table_of_contents_page(txt):
                    continue

                # Candidate balance sheet page?

                # WORKFLOW: reject page first if excluded header keywords appear in TOP 5 lines
                if excluded_keywords_above_statement_and_fye(page, None, None):
                    continue
                # Existing text-based detector remains primary.
                # Fallback: some PDFs extract the table body with spaced characters,
                # causing the text detector to miss the correct Balance Sheet page even though
                # the visual header clearly contains Balance Sheet + FYE.
                # Do NOT allow the fallback to accept reconciliation pages such as:
                # "Reconciliation of the Governmental Funds Balance Sheet to the Statement of Net Position".
                balance_header_text = top_header_text_from_page(page, max_lines=HEADER_MAX_LINES)
                balance_header_clean = clean_for_contains_match(balance_header_text)
                balance_header_fallback_ok = (
                    page_header_contains_balance_sheet_and_fye(page, fy_end_value, max_lines=HEADER_MAX_LINES)
                    and "reconciliation" not in balance_header_clean
                )

                if not (
                    is_balance_sheet_governmental_funds_page(txt)
                    or balance_header_fallback_ok
                ):
                    continue

                if first_candidate_page is None:
                    first_candidate_page = page_num

                # Header must contain Balance Sheet + FYE in top visual lines
                if not page_header_contains_balance_sheet_and_fye(page, fy_end_value, max_lines=HEADER_MAX_LINES):
                    continue

                # Boxes for proof (within top visual lines)
                heading_box = find_balance_sheet_heading_box_top_lines(page, max_lines=HEADER_MAX_LINES)
                fye_box = find_fye_box_on_page_top_lines(page, fy_end_value, max_lines=HEADER_MAX_LINES)

                if not heading_box or not fye_box:
                    continue

                # PASS
                proof_file = PROOF_DIR_BALANCE_SHEET / f"Row{processing_id}_BALANCE_SHEET_PASS_Page{page_num}.png"
                im = page.to_image(resolution=PROOF_IMAGE_RESOLUTION)

                boxes = []
                im.draw_rect(heading_box, stroke=HIGHLIGHT_STROKE, fill=HIGHLIGHT_FILL_RGBA)
                boxes.append(heading_box)

                im.draw_rect(fye_box, stroke=HIGHLIGHT_STROKE, fill=HIGHLIGHT_FILL_RGBA)
                boxes.append(fye_box)

                im.save(proof_file)

                title = f"ProcessingId {processing_id} | Balance Sheet = PASS | Page {page_num}"
                if show_popup:
                    show_proof_popup(proof_file, title, seconds=PROOF_POPUP_SECONDS)

                return {"status": "PASS", "page": page_num, "proof_file": proof_file, "boxes": boxes, "title": title}

            # FAIL
            if first_candidate_page is None:
                return {"status": "FAIL", "page": None, "proof_file": None, "boxes": [], "title": f"ProcessingId {processing_id} | Balance Sheet = FAIL (Not found)"}

            proof_file = PROOF_DIR_BALANCE_SHEET / f"Row{processing_id}_BALANCE_SHEET_FAIL_Page{first_candidate_page}.png"
            try:
                im = pdf.pages[first_candidate_page - 1].to_image(resolution=PROOF_IMAGE_RESOLUTION)
                im.save(proof_file)
            except Exception:
                pass

            title = f"ProcessingId {processing_id} | Balance Sheet = FAIL (Header/FYE not proven) | Page {first_candidate_page}"
            if show_popup:
                show_proof_popup(proof_file, title, seconds=PROOF_POPUP_SECONDS)

            return {"status": "FAIL", "page": first_candidate_page, "proof_file": proof_file, "boxes": [], "title": title}

    except Exception as e:
        return {"status": "FAIL", "page": None, "proof_file": None, "boxes": [], "title": f"ProcessingId {processing_id} | Balance Sheet = FAIL ({e})"}      

# VALIDATION: Locate the Revenues/Expenditures/Fund Balances statement heading near the top.
def find_rev_exp_fund_bal_heading_box_top_lines(page, max_lines: int = 12) -> Optional[dict]:
    """
    Find header bbox for the statement heading within top `max_lines` meaningful visual lines.

    Common heading variants:
      - "Statement of Revenues, Expenditures and Changes in Fund Balances"
      - "Statement of Revenues, Expenditures, and Changes in Fund Balances"
      - Sometimes line breaks split the title.
    """
    # Try longer phrases first (more precise)
    phrases = [
        "Statement of Revenues, Expenditures and Changes in Fund Balances",
        "Statement of Revenues, Expenditures, and Changes in Fund Balances",
        "Statement of Revenues Expenditures and Changes in Fund Balances",
        "Statement of Revenues",
    ]
    for ph in phrases:
        box = find_phrase_box_on_page_top_lines(page, ph, max_lines=max_lines)
        if box:
            return box

    return None

# VALIDATION: Confirm the revenues/expenditures statement header includes the FY date.
def page_header_contains_rev_exp_fund_bal_and_fye(page, fy_end_value, max_lines: int = 12) -> bool:
    """
    Header check using TOP VISUAL LINES:
      - Must contain 'statement of revenues'
      - Must contain 'expenditures' (or 'expenses' fallback)
      - Must contain 'fund balances' (or 'fund balance')
      - Must contain FY end date tokens in the same header area
    """
    header_text = top_header_text_from_page(page, max_lines=max_lines)
    if not header_text:
        return False

    header_clean = clean_for_contains_match(header_text)

    if "statement of revenues" not in header_clean:
        return False

    if ("expenditures" not in header_clean) and ("expenses" not in header_clean):
        return False

    if ("fund balances" not in header_clean) and ("fund balance" not in header_clean):
        return False

    d = parse_fy_end_date(fy_end_value)
    if not d:
        return False

    header_tokens = tokenize(header_text)

    for cand_tokens in build_fye_token_candidates(d):
        s, e = find_token_window(header_tokens, cand_tokens)
        if s is not None and e is not None:
            return True

    return False

# VALIDATION: Detect the Revenues/Expenditures/Changes in Fund Balances statement page.
def is_rev_exp_changes_fund_balances_page(page_text: str) -> bool:
    """
    ROBUST detector for:
      'Statement of Revenues, Expenditures, and Changes in Fund Balances'

    IMPORTANT FIX:
      - Do NOT reject pages just because they contain 'nonmajor'.
        Many valid governmental fund statements include a 'Nonmajor' column (like Littleton). 【1-9a3573】

    Pass if:
      - not excluded (TOC/MD&A/etc.)
      - not reconciliation
      - not budget/actual schedule
      - contains revenues + expenditures/expenses
      - contains fund balance/balances
      - contains strong indicators (net change / excess-deficiency / other financing / fund balance begin-end)
    """
    if not page_text:
        return False

    if is_excluded_fin_stmt_page(page_text):
        return False

    t = normalize_pdf_text(page_text).lower()

    # Exclude reconciliation pages
    if "reconciliation" in t:
        return False

    # Exclude budget schedules (RSI) that look similar but are not the basic statement
    if "budget and actual" in t or "budgetary" in t:
        return False

    # Core terms (must exist)
    if "revenues" not in t:
        return False

    if ("expenditures" not in t) and ("expenses" not in t):
        return False

    if ("fund balances" not in t) and ("fund balance" not in t):
        return False

    # Strong signals typical of this statement
    strong_signals = [
        "net change in fund balances",
        "net change in fund balance",
        "excess (deficiency) of revenues",
        "excess (deficiency) of revenues over",
        "excess (deficiency) of revenues and other",
        "other financing sources",
        "other financing uses",
        "other financing sources (uses)",
        "fund balances - beginning",
        "fund balance - beginning",
        "fund balances - end",
        "fund balance - end",
        "beginning of year",
        "end of year",
    ]

    if not any(sig in t for sig in strong_signals):
        return False

    # If the page is a narrative explanation mentioning these words, it may pass incorrectly.
    # Add one extra “table-ish” check: at least a few money-like numbers.
    money_like = len(re.findall(r"\$|\b\d{1,3}(?:,\d{3})+\b", page_text))
    if money_like < 8:
        return False

    return True

# VALIDATION: Validate that statement header + FY date exist and return PASS/FAIL with proof.
def _orig_write_rev_exp_changes_fund_balances_validation(
    processing_id: int,
    fy_end_value,
    pdf_path: Optional[Path],
    show_popup: bool = True
) -> dict:
    """
    Statement of Revenues, Expenditures and Changes in Fund Balances with FYE.

    - Does NOT write to parquet (parent run_validations_for_row batches the write)
    """
    HEADER_MAX_LINES = 12

    if not pdf_path or not Path(pdf_path).exists():
        return {"status": "FAIL", "page": None, "proof_file": None, "boxes": [], "title": f"ProcessingId {processing_id} | Rev/Exp/Fund Bal = FAIL (PDF missing)"}

    try:
        with open_pdf_maybe_shared(pdf_path) as pdf:
            if not pdf.pages:
                return {"status": "NOT EXTRACTABLE", "page": None, "proof_file": None, "boxes": [], "title": f"ProcessingId {processing_id} | Rev/Exp/Fund Bal = NOT EXTRACTABLE"}

            first_candidate_page = None

            for idx, page in enumerate(pdf.pages):
                page_num = idx + 1
                txt = page.extract_text() or ""
                if not txt.strip():
                    continue

                # Skip TOC pages
                if TOC_STOP_ENABLED and is_table_of_contents_page(txt):
                    continue

                # Candidate statement page?

                # WORKFLOW: reject page first if excluded header keywords appear in TOP 5 lines
                if excluded_keywords_above_statement_and_fye(page, None, None):
                    continue
                if not is_rev_exp_changes_fund_balances_page(txt):
                    continue

                if first_candidate_page is None:
                    first_candidate_page = page_num

                # Header must contain statement + FYE in top visual lines
                if not page_header_contains_rev_exp_fund_bal_and_fye(page, fy_end_value, max_lines=HEADER_MAX_LINES):
                    continue

                # Boxes for proof
                heading_box = find_rev_exp_fund_bal_heading_box_top_lines(page, max_lines=HEADER_MAX_LINES)
                fye_box = find_fye_box_on_page_top_lines(page, fy_end_value, max_lines=HEADER_MAX_LINES)

                if not heading_box or not fye_box:
                    continue

                # PASS
                proof_file = PROOF_DIR_REV_EXP_FUND_BAL / f"Row{processing_id}_REV_EXP_FUND_BAL_PASS_Page{page_num}.png"
                im = page.to_image(resolution=PROOF_IMAGE_RESOLUTION)

                boxes = []
                im.draw_rect(heading_box, stroke=HIGHLIGHT_STROKE, fill=HIGHLIGHT_FILL_RGBA)
                boxes.append(heading_box)

                im.draw_rect(fye_box, stroke=HIGHLIGHT_STROKE, fill=HIGHLIGHT_FILL_RGBA)
                boxes.append(fye_box)

                im.save(proof_file)

                title = f"ProcessingId {processing_id} | Rev/Exp/Fund Bal = PASS | Page {page_num}"
                if show_popup:
                    show_proof_popup(proof_file, title, seconds=PROOF_POPUP_SECONDS)

                return {"status": "PASS", "page": page_num, "proof_file": proof_file, "boxes": boxes, "title": title}

            # FAIL
            if first_candidate_page is None:
                return {"status": "FAIL", "page": None, "proof_file": None, "boxes": [], "title": f"ProcessingId {processing_id} | Rev/Exp/Fund Bal = FAIL (Not found)"}

            proof_file = PROOF_DIR_REV_EXP_FUND_BAL / f"Row{processing_id}_REV_EXP_FUND_BAL_FAIL_Page{first_candidate_page}.png"
            try:
                im = pdf.pages[first_candidate_page - 1].to_image(resolution=PROOF_IMAGE_RESOLUTION)
                im.save(proof_file)
            except Exception:
                pass

            title = f"ProcessingId {processing_id} | Rev/Exp/Fund Bal = FAIL (Header/FYE not proven) | Page {first_candidate_page}"
            if show_popup:
                show_proof_popup(proof_file, title, seconds=PROOF_POPUP_SECONDS)

            return {"status": "FAIL", "page": first_candidate_page, "proof_file": proof_file, "boxes": [], "title": title}

    except Exception as e:
        return {"status": "FAIL", "page": None, "proof_file": None, "boxes": [], "title": f"ProcessingId {processing_id} | Rev/Exp/Fund Bal = FAIL ({e})"}


# ---------------------------------------------------------
# PERFORMANCE: cache pdfplumber extract_words per page
# ---------------------------------------------------------
_PAGE_WORDS_CACHE = {}

# VALIDATION: Cache extracted PDF words for a page to speed up repeated validations.
def _get_cached_words(page):
    """
    Cache pdfplumber page.extract_words() results so we don't re-extract
    multiple times per page during validations (big speed improvement).
    """
    key = id(page)
    if key in _PAGE_WORDS_CACHE:
        return _PAGE_WORDS_CACHE[key]

    words = cached_page_words(page, use_text_flow=True)
    _PAGE_WORDS_CACHE[key] = words
    return words


# VALIDATION: Clear the PDF words cache between validations to avoid memory growth.
def _clear_page_words_cache():
    """
    Clear cached words between validations to avoid memory growth.
    """
    _PAGE_WORDS_CACHE.clear()

# VALIDATION: Collect words from the top visual lines of a PDF page (works for rotated/landscape pages).
def _top_words_by_lines(page, max_lines: int = 5) -> List[dict]:
    """
    Returns word dicts belonging to the visually top `max_lines` MEANINGFUL lines.

    IMPORTANT FIX:
    Some PDFs (especially landscape financial statements) have vertical running text
    in page margins. pdfplumber may extract those as single-letter lines (M, e, c...),
    which wrongly consume the first N lines and causes header checks to fail.

    This version:
      - uses cached extract_words()
      - groups by rounded 'top'
      - SKIPS tiny/single-letter lines
      - counts only meaningful lines toward max_lines
    """
    words = _get_cached_words(page)
    if not words:
        return []

    # group words by line (rounded y)
    line_map = {}
    for w in words:
        key = round(float(w.get("top", 0)), 0)
        line_map.setdefault(key, []).append(w)

    # decide if a line is meaningful
    def _is_meaningful_line(ws: List[dict]) -> bool:
        # build a raw line text
        ws_sorted = sorted(ws, key=lambda x: float(x.get("x0", 0)))
        raw = " ".join((w.get("text", "") or "").strip() for w in ws_sorted if (w.get("text", "") or "").strip())
        raw = raw.strip()
        if not raw:
            return False

        # remove spaces and punctuation for length check
        compact = re.sub(r"[^A-Za-z0-9]+", "", raw)

        # count tokens that are more than 1 alphanumeric character
        tokens = [re.sub(r"[^A-Za-z0-9]+", "", (w.get("text", "") or "")) for w in ws_sorted]
        long_tokens = [t for t in tokens if len(t) >= 2]

        # Meaningful if:
        # - enough overall text OR
        # - has at least 2 "real" tokens OR
        # - contains key header words
        if len(compact) >= 10:
            return True
        if len(long_tokens) >= 2:
            return True
        low = raw.lower()
        if ("statement" in low) or ("activities" in low) or ("position" in low) or ("june" in low) or ("december" in low):
            return True

        return False

    # collect top meaningful line keys
    chosen_keys = []
    for k in sorted(line_map.keys()):
        if _is_meaningful_line(line_map[k]):
            chosen_keys.append(k)
            if len(chosen_keys) >= max_lines:
                break

    # flatten chosen lines' words
    out = []
    for k in chosen_keys:
        out.extend(line_map[k])

    # stable order
    out.sort(key=lambda x: (round(float(x.get("top", 0)), 0), float(x.get("x0", 0))))
    return out

# VALIDATION: Build a clean header string from the top area of a PDF page using word coordinates.
def top_header_text_from_page(page, max_lines: int = 12) -> str:
    """
    Build a clean header text string from the TOP visual lines of the page
    using pdfplumber word coordinates (robust for landscape pages).

    Returns normalized lowercase header text.
    """
    top_words = _top_words_by_lines(page, max_lines=max_lines)
    if not top_words:
        return ""

    # group by rounded top -> line
    line_map = {}
    for w in top_words:
        key = round(float(w.get("top", 0)), 0)
        line_map.setdefault(key, []).append(w)

    lines = []
    for k in sorted(line_map.keys()):
        # left-to-right
        ws = sorted(line_map[k], key=lambda x: float(x.get("x0", 0)))
        line_text = " ".join((w.get("text", "") or "").strip() for w in ws if (w.get("text", "") or "").strip())
        if line_text.strip():
            lines.append(line_text.strip())

    return normalize_pdf_text(" ".join(lines)).lower()

# VALIDATION: Find a phrase only in the top header area and return a highlight box.
def find_phrase_box_on_page_top_lines(page, phrase: str, max_lines: int = 5) -> Optional[dict]:
    """
    Find phrase bounding box ONLY within the top `max_lines` visual lines of the page.
    Returns bbox dict or None.
    """
    phrase = (phrase or "").strip()
    target_tokens = tokenize(phrase)
    if not target_tokens:
        return None

    top_words = _top_words_by_lines(page, max_lines=max_lines)
    if not top_words:
        return None

    page_tokens = []
    token_word_map = []

    for w in top_words:
        wtoks = tokenize(w.get("text", ""))
        for t in wtoks:
            page_tokens.append(t)
            token_word_map.append(w)

    s, e = find_token_window(page_tokens, target_tokens)
    if s is None or e is None:
        return None

    matched_words = token_word_map[s:e]
    return {
        "x0": min(w["x0"] for w in matched_words),
        "top": min(w["top"] for w in matched_words),
        "x1": max(w["x1"] for w in matched_words),
        "bottom": max(w["bottom"] for w in matched_words),
    }

# VALIDATION: Find the FY date only in the top header area and return a highlight box.
def find_fye_box_on_page_top_lines(page, fy_end_value, max_lines: int = 5) -> Optional[dict]:
    """
    Find FY end date bbox ONLY within the top `max_lines` visual lines of the page.
    No strict 'year ended' context needed for statement headers.
    """
    d = parse_fy_end_date(fy_end_value)
    if not d:
        return None

    top_words = _top_words_by_lines(page, max_lines=max_lines)
    if not top_words:
        return None

    page_tokens = []
    token_word_map = []

    for w in top_words:
        wtoks = tokenize(w.get("text", ""))
        for t in wtoks:
            page_tokens.append(t)
            token_word_map.append(w)

    for cand in build_fye_token_candidates(d):
        s, e = find_token_window(page_tokens, cand)
        if s is None or e is None:
            continue

        matched_words = token_word_map[s:e]
        return {
            "x0": min(w["x0"] for w in matched_words),
            "top": min(w["top"] for w in matched_words),
            "x1": max(w["x1"] for w in matched_words),
            "bottom": max(w["bottom"] for w in matched_words),
        }

    return None

def _first_word_box(page, word: str) -> Optional[dict]:
    """
    Return first bbox for a single word (e.g., ASSETS / LIABILITIES).
    Used to prove table structure exists.
    """
    word = (word or "").strip().lower()
    if not word:
        return None

    words = _get_cached_words(page)
    if not words:
        return None

    for w in words:
        t = (w.get("text", "") or "").strip().lower()
        if t == word:
            return {"x0": w["x0"], "top": w["top"], "x1": w["x1"], "bottom": w["bottom"]}
    return None

# VALIDATION: Strictly detect Statement of Net Position pages.
def is_statement_of_net_position_page(page_text: str) -> bool:
    """
    STRICT detector for Statement(s) of Net Position pages.

    NOTE:
    Many ACFR PDFs have running headers, so the statement title may not appear
    within first 5 lines. We detect the statement by:
      - Not excluded page
      - Contains 'statement(s) of net position' somewhere near top (first ~20 lines)
      - Contains table sanity words anywhere on page
    """
    if not page_text:
        return False

    # Exclude TOC / MD&A / reconciliation / etc.
    if is_excluded_fin_stmt_page(page_text):
        return False

    t_all = normalize_pdf_text(page_text).lower()

    # Look in a slightly larger top window (handles running headers)
    top20 = normalize_pdf_text(" ".join(_top_non_empty_lines(page_text, max_lines=20))).lower()

    if not (
        "statement of net position" in top20
        or "statements of net position" in top20
    ):
        return False

    # Table sanity check anywhere on page
    must_have = ["assets", "liabilities", "net position"]
    if not all(w in t_all for w in must_have):
        return False

    return True

# VALIDATION: Validate Statement of Net Position header + FY date and write PASS/FAIL with proof.
def _orig_write_statement_of_net_position_validation(
    processing_id: int,
    fy_end_value,
    pdf_path: Optional[Path],
    show_popup: bool = True
) -> dict:
    """
    Statement of Net Position with FYE.

    PASS only if:
      - Candidate page is a net position statement page (strict detector)
      - Header area (top visual lines) contains:
            * 'Statement(s) of Net Position'
            * FY End Date (from fy_end_value)
      - We can locate highlight boxes for both heading and date within top visual lines.

    - Does NOT write to parquet (parent run_validations_for_row batches the write)
    """
    HEADER_MAX_LINES = 12

    if not pdf_path or not Path(pdf_path).exists():
        return {"status": "FAIL", "page": None, "proof_file": None, "boxes": [], "title": f"ProcessingId {processing_id} | Net Position = FAIL (PDF missing)"}

    try:
        with open_pdf_maybe_shared(pdf_path) as pdf:
            if not pdf.pages:
                return {"status": "NOT EXTRACTABLE", "page": None, "proof_file": None, "boxes": [], "title": f"ProcessingId {processing_id} | Net Position = NOT EXTRACTABLE"}

            first_candidate_page = None

            for idx, page in enumerate(pdf.pages):
                page_num = idx + 1
                txt = page.extract_text() or ""
                if not txt.strip():
                    continue

                # skip TOC pages
                if TOC_STOP_ENABLED and is_table_of_contents_page(txt):
                    continue

                # Candidate statement page?

                # WORKFLOW: reject page first if excluded header keywords appear in TOP 5 lines
                if excluded_keywords_above_statement_and_fye(page, None, None):
                    continue
                if not is_statement_of_net_position_page(txt):
                    continue

                if first_candidate_page is None:
                    first_candidate_page = page_num

                # header rule uses TOP VISUAL lines from page words
                if not page_header_contains_net_position_and_fye(page, fy_end_value, max_lines=HEADER_MAX_LINES):
                    continue

                # Boxes for proof (within top visual lines)
                heading_box = find_net_position_heading_box_top_lines(page, max_lines=HEADER_MAX_LINES)
                fye_box = find_fye_box_on_page_top_lines(page, fy_end_value, max_lines=HEADER_MAX_LINES)

                if not heading_box or not fye_box:
                    continue

                # PASS
                proof_file = PROOF_DIR_NET_POSITION / f"Row{processing_id}_NET_POSITION_PASS_Page{page_num}.png"
                im = page.to_image(resolution=PROOF_IMAGE_RESOLUTION)

                boxes = []
                im.draw_rect(heading_box, stroke=HIGHLIGHT_STROKE, fill=HIGHLIGHT_FILL_RGBA)
                boxes.append(heading_box)

                im.draw_rect(fye_box, stroke=HIGHLIGHT_STROKE, fill=HIGHLIGHT_FILL_RGBA)
                boxes.append(fye_box)

                im.save(proof_file)

                title = f"ProcessingId {processing_id} | Net Position = PASS | Page {page_num}"
                if show_popup:
                    show_proof_popup(proof_file, title, seconds=PROOF_POPUP_SECONDS)

                return {"status": "PASS", "page": page_num, "proof_file": proof_file, "boxes": boxes, "title": title}

            # FAIL
            if first_candidate_page is None:
                return {"status": "FAIL", "page": None, "proof_file": None, "boxes": [], "title": f"ProcessingId {processing_id} | Net Position = FAIL (Statement not found)"}

            # Save FAIL proof (popup suppressed by your global rule)
            proof_file = PROOF_DIR_NET_POSITION / f"Row{processing_id}_NET_POSITION_FAIL_Page{first_candidate_page}.png"
            try:
                im = pdf.pages[first_candidate_page - 1].to_image(resolution=PROOF_IMAGE_RESOLUTION)
                im.save(proof_file)
            except Exception:
                pass

            title = f"ProcessingId {processing_id} | Net Position = FAIL (Header missing Net Position or FYE) | Page {first_candidate_page}"
            if show_popup:
                show_proof_popup(proof_file, title, seconds=PROOF_POPUP_SECONDS)

            return {"status": "FAIL", "page": first_candidate_page, "proof_file": proof_file, "boxes": [], "title": title}

    except Exception as e:
        return {"status": "FAIL", "page": None, "proof_file": None, "boxes": [], "title": f"ProcessingId {processing_id} | Net Position = FAIL ({e})"}

def find_proprietary_net_position_heading_box_top_lines(page, max_lines: int = 12) -> Optional[dict]:
    """
    Find heading bbox inside TOP VISUAL lines for proprietary funds net position.
    """
    phrases = [
        "Statement of Net Position - Proprietary Funds",
        "Statement of Net Position Proprietary Funds",
        "Statements of Net Position - Proprietary Funds",
        "Statements of Net Position Proprietary Funds",
        "Proprietary Funds",
    ]
    for ph in phrases:
        box = find_phrase_box_on_page_top_lines(page, ph, max_lines=max_lines)
        if box:
            return box

    # fallback: Net Position if title is truncated
    return find_phrase_box_on_page_top_lines(page, "Net Position", max_lines=max_lines)


def page_header_contains_proprietary_net_position_and_fye(page, fy_end_value, max_lines: int = 12) -> bool:
    """
    Header check (TOP VISUAL LINES):
    Must contain:
      - statement
      - net position
      - proprietary
      - fund/funds
      - FY end date tokens
    """
    header_text = top_header_text_from_page(page, max_lines=max_lines)
    if not header_text:
        return False

    if "statement" not in header_text:
        return False
    if "net position" not in header_text:
        return False
    if "proprietary" not in header_text:
        return False
    if ("fund" not in header_text) and ("funds" not in header_text):
        return False

    d = parse_fy_end_date(fy_end_value)
    if not d:
        return False

    header_tokens = tokenize(header_text)
    for cand_tokens in build_fye_token_candidates(d):
        s, e = find_token_window(header_tokens, cand_tokens)
        if s is not None and e is not None:
            return True

    return False


def is_statement_of_net_position_proprietary_funds_page(page_text: str) -> bool:
    """
    STRICT detector for proprietary funds net position table pages.
    """
    if not page_text:
        return False

    if is_excluded_fin_stmt_page(page_text):
        return False

    t = normalize_pdf_text(page_text).lower()

    # Exclude reconciliation & budget pages
    if "reconciliation" in t:
        return False
    if "budget and actual" in t or "budgetary" in t:
        return False

    # Core required words
    if "net position" not in t:
        return False
    if "proprietary" not in t:
        return False
    if ("fund" not in t) and ("funds" not in t):
        return False

    # Table structure words
    if ("assets" not in t) or ("liabilities" not in t):
        return False

    return True


def _orig_write_statement_of_net_position_proprietary_funds_validation(
    processing_id: int,
    fy_end_value,
    pdf_path: Optional[Path],
    show_popup: bool = True
) -> dict:
    """
    Statement of Net Position of Proprietary Funds with FYE (TABLE FORMAT).

    - Does NOT write to parquet (parent run_validations_for_row batches the write)
    """
    HEADER_MAX_LINES = 12

    if not pdf_path or not Path(pdf_path).exists():
        return {"status": "FAIL", "page": None, "proof_file": None, "boxes": [], "title": f"ProcessingId {processing_id} | Prop Net Pos = FAIL (PDF missing)"}

    try:
        with open_pdf_maybe_shared(pdf_path) as pdf:
            first_candidate_page = None

            for idx, page in enumerate(pdf.pages):
                page_num = idx + 1
                txt = page.extract_text() or ""
                if not txt.strip():
                    continue

                if TOC_STOP_ENABLED and is_table_of_contents_page(txt):
                    continue


                # WORKFLOW: reject page first if excluded header keywords appear in TOP 5 lines
                if excluded_keywords_above_statement_and_fye(page, None, None):
                    continue
                if not is_statement_of_net_position_proprietary_funds_page(txt):
                    continue

                if first_candidate_page is None:
                    first_candidate_page = page_num

                if not page_header_contains_proprietary_net_position_and_fye(page, fy_end_value, max_lines=HEADER_MAX_LINES):
                    continue

                heading_box = find_proprietary_net_position_heading_box_top_lines(page, max_lines=HEADER_MAX_LINES)
                fye_box = find_fye_box_on_page_top_lines(page, fy_end_value, max_lines=HEADER_MAX_LINES)
                if not heading_box or not fye_box:
                    continue

                # Excluded header check
                if excluded_keywords_above_statement_and_fye(page, heading_box, fye_box, max_lines=HEADER_MAX_LINES):
                    continue

                # Table proof: ASSETS or LIABILITIES word
                struct_box = _first_word_box(page, "assets") or _first_word_box(page, "liabilities")
                if not struct_box:
                    continue

                # PASS
                proof_file = PROOF_DIR_NET_POSITION / f"Row{processing_id}_PROP_NET_POSITION_PASS_Page{page_num}.png"
                im = page.to_image(resolution=PROOF_IMAGE_RESOLUTION)

                boxes = []
                im.draw_rect(heading_box, stroke=HIGHLIGHT_STROKE, fill=HIGHLIGHT_FILL_RGBA); boxes.append(heading_box)
                im.draw_rect(fye_box, stroke=HIGHLIGHT_STROKE, fill=HIGHLIGHT_FILL_RGBA); boxes.append(fye_box)
                im.draw_rect(struct_box, stroke=HIGHLIGHT_STROKE, fill=HIGHLIGHT_FILL_RGBA); boxes.append(struct_box)
                im.save(proof_file)

                title = f"ProcessingId {processing_id} | Prop Net Pos = PASS | Page {page_num}"
                if show_popup:
                    show_proof_popup(proof_file, title, seconds=PROOF_POPUP_SECONDS)

                return {"status": "PASS", "page": page_num, "proof_file": proof_file, "boxes": boxes, "title": title}

            if first_candidate_page is None:
                return {"status": "FAIL", "page": None, "proof_file": None, "boxes": [], "title": f"ProcessingId {processing_id} | Prop Net Pos = FAIL (Not found)"}
            return {"status": "FAIL", "page": first_candidate_page, "proof_file": None, "boxes": [], "title": f"ProcessingId {processing_id} | Prop Net Pos = FAIL"}

    except Exception as e:
        return {"status": "FAIL", "page": None, "proof_file": None, "boxes": [], "title": f"ProcessingId {processing_id} | Prop Net Pos = FAIL ({e})"}


def find_rev_exp_changes_net_position_heading_box_top_lines(page, max_lines: int = 12) -> Optional[dict]:
    """
    Find heading bbox for Rev/Exp/Changes in Net Position OR Net Assets within top lines.

    Supports:
    - Statement(ss, Expenses and Changes in Net Position
    - Statement(s) of Revenues, Expenses and Changes in Net Assets
    """
    phrases = [
        # Net Position variants - existing logic
        "Statement of Revenues, Expenses and Changes in Net Position",
        "Statement of Revenues, Expenses, and Changes in Net Position",
        "Statement of Revenues Expenses and Changes in Net Position",
        "Statements of Revenues, Expenses and Changes in Net Position",
        "Statements of Revenues, Expenses, and Changes in Net Position",
        "Statements of Revenues Expenses and Changes in Net Position",

        # Net Assets variants - newly added
        "Statement of Revenues, Expenses and Changes in Net Assets",
        "Statement of Revenues, Expenses, and Changes in Net Assets",
        "Statement of Revenues Expenses and Changes in Net Assets",
        "Statements of Revenues, Expenses and Changes in Net Assets",
        "Statements of Revenues, Expenses, and Changes in Net Assets",
        "Statements of Revenues Expenses and Changes in Net Assets",

        # Existing fallback
        "Statement of Revenues",
        "Statements of Revenues",
    ]

    for ph in phrases:
        box = find_phrase_box_on_page_top_lines(page, ph, max_lines=max_lines)
        if box:
            return box

    return None

def page_header_contains_rev_exp_changes_net_position_and_fye(page, fy_end_value, max_lines: int = 12) -> bool:
    """
    Header check using TOP VISUAL LINES.

    Supports:
    - Statement(s) of Revenues, Expenses and Changes in Net Position
    - Statement(s) of Revenues, Expenses and Changes in Net Assets

    Logic:
    - Must contain statement + revenue + expense
    - Must contain either "net position" OR "net assets"
    - For Net Position → require fund/funds (keep existing stricter logic)
    - For Net Assets → do NOT require fund/funds
    - Must contain valid FYE
    """

    header_text = top_header_text_from_page(page, max_lines=max_lines)
    if not header_text:
        return False

    header_text = header_text.lower()

    # -----------------------------
    # CORE STRUCTURE CHECK
    # -----------------------------
    if "statement" not in header_text:
        return False

    if "revenu" not in header_text:
        return False

    if ("expense" not in header_text) and ("expenses" not in header_text):
        return False

    # -----------------------------
    # NET POSITION / NET ASSETS CHECK
    # -----------------------------
    has_net_position = "net position" in header_text
    has_net_assets = "net assets" in header_text

    if not (has_net_position or has_net_assets):
        return False

    # Net assets → no fund requirement (intentional)

    # -----------------------------
    # FYE CHECK
    # -----------------------------
    d = parse_fy_end_date(fy_end_value)
    if not d:
        return False

    header_tokens = tokenize(header_text)

    for cand_tokens in build_fye_token_candidates(d):
        s, e = find_token_window(header_tokens, cand_tokens)
        if s is not None and e is not None:
            return True

    return False

def is_rev_exp_changes_net_position_page(page_text: str) -> bool:
    """
    ROBUST detector for:
    - Statement of Revenues, Expenses and Changes in Net Position
    - Statements of Revenues, Expenses and Changes in Net Position
    - Statement of Revenues, Expenses and Changes in Net Assets
    - Statements of Revenues, Expenses and Changes in Net Assets

    Must include table-ish signals and exclude junk pages.
    """
    if not page_text:
        return False

    if is_excluded_fin_stmt_page(page_text):
        return False

    t = normalize_pdf_text(page_text).lower()

    if "budget and actual" in t or "budgetary" in t:
        return False

    if "reconciliation" in t:
        return False

    if "statement" not in t:
        return False

    if ("revenue" not in t) and ("revenues" not in t):
        return False

    if ("expense" not in t) and ("expenses" not in t):
        return False

    has_net_position = "net position" in t
    has_net_assets = "net assets" in t

    if not (has_net_position or has_net_assets):
        return False

    # Preserve existing stricter proprietary-fund guard for Net Position pages.
    # Do not require fund/funds for Net Assets pages.
    if has_net_position and not has_net_assets:
        if ("fund" not in t) and ("funds" not in t):
            return False

    # Table-ish: operating/nonoperating words help avoid false pages.
    # Kept unchanged from existing logic.
    if ("operating" not in t) and ("nonoperating" not in t) and ("non-operating" not in t):
        return False

    return True

def _orig_write_rev_exp_changes_net_position_validation(
    processing_id: int,
    fy_end_value,
    pdf_path: Optional[Path],
    show_popup: bool = True
) -> dict:
    """
    Statement of Revenues, Expenses and Changes in Net Position with FYE.

    - Does NOT write to parquet (parent run_validations_for_row batches the write)
    """
    HEADER_MAX_LINES = 12

    if not pdf_path or not Path(pdf_path).exists():
        return {"status": "FAIL", "page": None, "proof_file": None, "boxes": [], "title": f"ProcessingId {processing_id} | Rev/Exp/Chg Net Pos = FAIL (PDF missing)"}

    try:
        with open_pdf_maybe_shared(pdf_path) as pdf:
            first_candidate_page = None

            for idx, page in enumerate(pdf.pages):
                page_num = idx + 1
                txt = page.extract_text() or ""
                if not txt.strip():
                    continue

                if TOC_STOP_ENABLED and is_table_of_contents_page(txt):
                    continue


                # WORKFLOW: reject page first if excluded header keywords appear in TOP 5 lines
                if excluded_keywords_above_statement_and_fye(page, None, None):
                    continue
                if not is_rev_exp_changes_net_position_page(txt):
                    continue

                if first_candidate_page is None:
                    first_candidate_page = page_num

                if not page_header_contains_rev_exp_changes_net_position_and_fye(page, fy_end_value, max_lines=HEADER_MAX_LINES):
                    continue

                heading_box = find_rev_exp_changes_net_position_heading_box_top_lines(page, max_lines=HEADER_MAX_LINES)
                fye_box = find_fye_box_on_page_top_lines(page, fy_end_value, max_lines=HEADER_MAX_LINES)
                if not heading_box or not fye_box:
                    continue

                if excluded_keywords_above_statement_and_fye(page, heading_box, fye_box, max_lines=HEADER_MAX_LINES):
                    continue

                proof_file = PROOF_DIR_REV_EXP_CHG_NET_POSITION / f"Row{processing_id}_REV_EXP_CHG_NET_POS_PASS_Page{page_num}.png"
                im = page.to_image(resolution=PROOF_IMAGE_RESOLUTION)

                boxes = []
                im.draw_rect(heading_box, stroke=HIGHLIGHT_STROKE, fill=HIGHLIGHT_FILL_RGBA); boxes.append(heading_box)
                im.draw_rect(fye_box, stroke=HIGHLIGHT_STROKE, fill=HIGHLIGHT_FILL_RGBA); boxes.append(fye_box)
                im.save(proof_file)

                title = f"ProcessingId {processing_id} | Rev/Exp/Chg Net Pos = PASS | Page {page_num}"
                if show_popup:
                    show_proof_popup(proof_file, title, seconds=PROOF_POPUP_SECONDS)

                return {"status": "PASS", "page": page_num, "proof_file": proof_file, "boxes": boxes, "title": title}

            if first_candidate_page is None:
                return {"status": "FAIL", "page": None, "proof_file": None, "boxes": [], "title": f"ProcessingId {processing_id} | Rev/Exp/Chg Net Pos = FAIL (Not found)"}
            return {"status": "FAIL", "page": first_candidate_page, "proof_file": None, "boxes": [], "title": f"ProcessingId {processing_id} | Rev/Exp/Chg Net Pos = FAIL"}

    except Exception as e:
        return {"status": "FAIL", "page": None, "proof_file": None, "boxes": [], "title": f"ProcessingId {processing_id} | Rev/Exp/Chg Net Pos = FAIL ({e})"}

def find_cash_receipts_disbursements_heading_box_top_lines(page, max_lines: int = 12) -> Optional[dict]:
    """
    Find heading bbox for Cash Receipts / Disbursements style statements within top lines.

    Supported common variants:
    - Statement of Cash Receipts and Disbursements
    - Statements of Cash Receipts and Disbursements
    - Statement of Receipts, Disbursements, and Cash and Investment Balances
    - Statements of Receipts, Disbursements, and Cash and Investment Balances
    - Statement of Receipts, Disbursements, and Changes in Cash and Investment Balances
    - Statements of Receipts, Disbursements, and Changes in Cash and Investment Balances
    """
    phrases = [
        "Statement of Cash Receipts and Disbursements",
        "Statements of Cash Receipts and Disbursements",
        "Statement of Cash Receipts, and Disbursements",
        "Statements of Cash Receipts, and Disbursements",

        "Statement of Receipts, Disbursements, and Cash and Investment Balances",
        "Statements of Receipts, Disbursements, and Cash and Investment Balances",
        "Statement of Receipts Disbursements and Cash and Investment Balances",
        "Statements of Receipts Disbursements and Cash and Investment Balances",

        "Statement of Receipts, Disbursements, and Changes in Cash and Investment Balances",
        "Statements of Receipts, Disbursements, and Changes in Cash and Investment Balances",
        "Statement of Receipts Disbursements and Changes in Cash and Investment Balances",
        "Statements of Receipts Disbursements and Changes in Cash and Investment Balances",

        # shorter safe fallbacks
        "Statement of Cash Receipts",
        "Statements of Cash Receipts",
        "Statement of Receipts, Disbursements",
        "Statements of Receipts, Disbursements",
    ]

    for ph in phrases:
        box = find_phrase_box_on_page_top_lines(page, ph, max_lines=max_lines)
        if box:
            return box

    return None


def page_header_contains_cash_receipts_disbursements_and_fye(page, fy_end_value, max_lines: int = 12) -> bool:
    """
    Header check using TOP VISUAL LINES.

    Accept either:
    1) 'Statement(s) of Cash Receipts and Disbursements'
    OR
    2) 'Statement(s) of Receipts, Disbursements, and Cash and Investment Balances'
       (also accepts 'Changes in Cash and Investment Balances')

    Must also contain valid FYE in the same header area.
    Explicitly excludes cash flow statements.
    """
    header_text = top_header_text_from_page(page, max_lines=max_lines)
    if not header_text:
        return False

    header_text = header_text.lower()

    if "statement" not in header_text:
        return False

    # Hard exclusion: avoid confusing with cash flow statements
    if "cash flow" in header_text or "cash flows" in header_text:
        return False

    has_receipts = ("receipt" in header_text) or ("receipts" in header_text)
    has_disbursements = ("disbursement" in header_text) or ("disbursements" in header_text)

    # Variant A: Statement of Cash Receipts and Disbursements
    variant_a = ("cash" in header_text) and has_receipts and has_disbursements

    # Variant B: Statement of Receipts, Disbursements, and Cash and Investment Balances
    has_cash_investment_balances = (
        ("cash" in header_text)
        and ("investment" in header_text or "investments" in header_text)
        and ("balance" in header_text or "balances" in header_text)
    )
    variant_b = has_receipts and has_disbursements and has_cash_investment_balances

    if not (variant_a or variant_b):
        return False

    d = parse_fy_end_date(fy_end_value)
    if not d:
        return False

    header_tokens = tokenize(header_text)
    for cand_tokens in build_fye_token_candidates(d):
        s, e = find_token_window(header_tokens, cand_tokens)
        if s is not None and e is not None:
            return True

    return False


def is_statement_of_cash_receipts_disbursements_page(page_text: str) -> bool:
    """
    ROBUST detector for:
    - Statement(s) of Cash Receipts and Disbursements
    - Statement(s) of Receipts, Disbursements, and Cash and Investment Balances
    - Statement(s) of Receipts, Disbursements, and Changes in Cash and Investment Balances
    """
    if not page_text:
        return False

    if is_excluded_fin_stmt_page(page_text):
        return False

    t = normalize_pdf_text(page_text).lower()

    # Avoid false positives
    if "cash flow" in t or "cash flows" in t:
        return False
    if "budget and actual" in t or "budgetary" in t:
        return False
    if "reconciliation" in t:
        return False

    if "statement" not in t:
        return False

    has_receipts = ("receipt" in t) or ("receipts" in t)
    has_disbursements = ("disbursement" in t) or ("disbursements" in t)

    if not (has_receipts and has_disbursements):
        return False

    # Accept either the shorter 'cash receipts and disbursements' style
    has_cash_receipts_style = ("cash" in t and has_receipts and has_disbursements)

    # Or the longer 'cash and investment balances' style
    has_cash_investment_balance_style = (
        ("cash" in t)
        and ("investment" in t or "investments" in t)
        and ("balance" in t or "balances" in t)
    )

    if not (has_cash_receipts_style or has_cash_investment_balance_style):
        return False

    return True


def _orig_write_statement_of_cash_receipts_disbursements_validation(
    processing_id: int,
    fy_end_value,
    pdf_path: Optional[Path],
    show_popup: bool = True
) -> dict:
    """
    Statement of Cash Receipts and Disbursements with FYE.

    - Does NOT write to parquet (parent run_validations_for_row batches the write)
    """
    HEADER_MAX_LINES = 12

    if not pdf_path or not Path(pdf_path).exists():
        return {
            "status": "FAIL",
            "page": None,
            "proof_file": None,
            "boxes": [],
            "title": f"ProcessingId {processing_id} | Cash Receipts/Disb = FAIL (PDF missing)"
        }

    try:
        with open_pdf_maybe_shared(pdf_path) as pdf:
            first_candidate_page = None

            for idx, page in enumerate(pdf.pages):
                page_num = idx + 1
                txt = page.extract_text() or ""
                if not txt.strip():
                    continue

                if TOC_STOP_ENABLED and is_table_of_contents_page(txt):
                    continue

                # Same workflow guard as your other statement validations
                if excluded_keywords_above_statement_and_fye(page, None, None):
                    continue

                if not is_statement_of_cash_receipts_disbursements_page(txt):
                    continue

                if first_candidate_page is None:
                    first_candidate_page = page_num

                if not page_header_contains_cash_receipts_disbursements_and_fye(
                    page, fy_end_value, max_lines=HEADER_MAX_LINES
                ):
                    continue

                heading_box = find_cash_receipts_disbursements_heading_box_top_lines(
                    page, max_lines=HEADER_MAX_LINES
                )
                fye_box = find_fye_box_on_page_top_lines(
                    page, fy_end_value, max_lines=HEADER_MAX_LINES
                )

                if not heading_box or not fye_box:
                    continue

                if excluded_keywords_above_statement_and_fye(
                    page, heading_box, fye_box, max_lines=HEADER_MAX_LINES
                ):
                    continue

                proof_file = (
                    PROOF_DIR_CASH_RECEIPTS_DISBURSEMENTS
                    / f"Row{processing_id}_CASH_RECEIPTS_DISB_PASS_Page{page_num}.png"
                )
                im = page.to_image(resolution=PROOF_IMAGE_RESOLUTION)

                boxes = []
                im.draw_rect(heading_box, stroke=HIGHLIGHT_STROKE, fill=HIGHLIGHT_FILL_RGBA); boxes.append(heading_box)
                im.draw_rect(fye_box, stroke=HIGHLIGHT_STROKE, fill=HIGHLIGHT_FILL_RGBA); boxes.append(fye_box)
                im.save(proof_file)

                title = f"ProcessingId {processing_id} | Cash Receipts/Disb = PASS | Page {page_num}"
                if show_popup:
                    show_proof_popup(proof_file, title, seconds=PROOF_POPUP_SECONDS)

                return {
                    "status": "PASS",
                    "page": page_num,
                    "proof_file": proof_file,
                    "boxes": boxes,
                    "title": title
                }

            if first_candidate_page is None:
                return {
                    "status": "FAIL",
                    "page": None,
                    "proof_file": None,
                    "boxes": [],
                    "title": f"ProcessingId {processing_id} | Cash Receipts/Disb = FAIL (Not found)"
                }

            return {
                "status": "FAIL",
                "page": first_candidate_page,
                "proof_file": None,
                "boxes": [],
                "title": f"ProcessingId {processing_id} | Cash Receipts/Disb = FAIL"
            }

    except Exception as e:
        return {
            "status": "FAIL",
            "page": None,
            "proof_file": None,
            "boxes": [],
            "title": f"ProcessingId {processing_id} | Cash Receipts/Disb = FAIL ({e})"
        }


# =========================================================
# C4F-INSPIRED LIGHTWEIGHT VALIDATION HELPERS
# =========================================================
# Incorporated selectively from C4F_AR_Validator.py:
# - stronger false-positive rejection for notes/prose/management/auditor pages
# - table-structure checks using right-side financial numbers
# - simple number-column clustering
# - audit-report scoring fallback
#
# Important: these helpers are used ONLY as secondary/fallback support.
# Existing validations remain primary and are not removed.

C4F_DYN_EXTRA_EXCLUDED_TITLE_CONTEXTS = [
    r"\bsupplementary\b",
    r"\bsupplemental\b",
    r"\badditional\s+information\b",
    r"\bnotes?\s+(?:to|on|for)\b",
    r"\bin\s+accordance\s+with\b",
    r"\bcondensed\b",
    r"\babbreviated\b",
    r"\bextract(?:s|ed)?\b",
    r"\bhighlights?\b",
    r"\banalysis\s+of\b",
    r"\breview\s+of\b",
    r"\bdiscussion\s+(?:of|on|and)\b",
    r"\bcommentary\s+on\b",
    r"\bimpact\s+on\s+(?:the\s+)?(?:balance|income|cash|statement)\b",
    r"\beffect\s+on\b",
    r"\bpro\s*forma\b",
    r"\bforecast(?:s|ed)?\b",
    r"\bprojection\b",
    r"\bfive[\s-]year\b",
    r"\bten[\s-]year\b",
    r"\bmulti[\s-]year\b",
    r"\bhistorical\b",
    r"\bsegment(?:al|ed)?\s+(?:information|reporting|analysis)\b",
    r"\binterim\b",
    r"\bquarterly\b",
    r"\bkey\s+figures?\b",
    r"\bat\s+a\s+glance\b",
    r"\breconciliation\s+(?:of|to)\b",
    r"\bmanagement\s+(?:report|discussion|commentary|review)\b",
]

C4F_DYN_PAGE_REJECTION_PATTERNS = [
    r"\bnotes?\s+to\s+(?:the\s+)?(?:basic\s+)?financial\s+statements?\b",
    r"\bindependent\s+auditor",
    r"\bauditor'?s?\s+report\b",
    r"\bwe\s+have\s+audited\b",
    r"\bin\s+our\s+opinion\b",
    r"\baccounting\s+polic(?:y|ies)\b",
    r"\bsignificant\s+accounting\b",
    r"\bsummary\s+of\s+(?:significant\s+)?accounting\b",
    r"\btable\s+of\s+contents\b",
    r"\bmanagement(?:'s)?\s+(?:discussion|report|commentary|review)\b",
    r"\bdirectors?'?\s+report\b",
    r"\bcorporate\s+governance\b",
    r"\brisk\s+management\b",
    r"\bfinancial\s+highlights?\b",
    r"\boperational\s+review\b",
    r"\bregistered\s+public\s+accounting\b",
]

C4F_DYN_REFERENCE_PATTERNS = [
    r"\bsee\s+(?:the\s+)?(?:basic\s+)?(?:financial|statement|note)",
    r"\brefer\s+to\s+(?:the\s+)?",
    r"\bas\s+(?:shown|reported|presented|disclosed)\s+(?:in|on)\b",
    r"\bthe\s+accompanying\s+notes\b",
    r"\bnote\s+\d+",
    r"\bbalance\s+sheet\s+(?:position|exposure|date|item)\b",
    r"\bincome\s+statement\s+(?:impact|effect|line)\b",
    r"\bcash\s+flow\s+(?:impact|effect|hedge|projection|forecast)\b",
    r"\bfor\s+the\s+purposes?\s+of\s+the\s+statements?\b",
]

_C4F_DYN_EXTRA_EXCLUDED_TITLE_RES = [re.compile(p, re.I) for p in C4F_DYN_EXTRA_EXCLUDED_TITLE_CONTEXTS]
_C4F_DYN_PAGE_REJECTION_RES = [re.compile(p, re.I) for p in C4F_DYN_PAGE_REJECTION_PATTERNS]
_C4F_DYN_REFERENCE_RES = [re.compile(p, re.I) for p in C4F_DYN_REFERENCE_PATTERNS]


def _c4f_top_page_text(page_text: str, max_lines: int = 18) -> str:
    """Top-page text used for fast section/prose rejection."""
    lines = [ln.strip() for ln in (page_text or '').splitlines() if ln.strip()]
    return normalize_pdf_text(' '.join(lines[:max_lines])).lower()


def _c4f_dyn_page_rejection_reason(page_text: str) -> Optional[str]:
    """
    Reject obvious non-statement pages before table fallback scans them.
    This mirrors the C4F validator's strong rejection idea but stays lightweight.
    """
    if not page_text:
        return 'blank'
    top = _c4f_top_page_text(page_text, max_lines=18)
    if not top:
        return 'blank'
    for rx in _C4F_DYN_PAGE_REJECTION_RES:
        if rx.search(top):
            return 'page_rejection_context'
    return None


def _c4f_dyn_title_context_is_excluded(text: str) -> bool:
    """Extra C4F-style title-context rejection for multi-line heading candidates."""
    t = normalize_pdf_text(text or '').lower()
    if not t:
        return False
    return any(rx.search(t) for rx in _C4F_DYN_EXTRA_EXCLUDED_TITLE_RES)


def _c4f_dyn_is_reference_line(text: str) -> bool:
    """Reject prose/reference lines when collecting table evidence."""
    t = normalize_pdf_text(text or '').lower()
    return any(rx.search(t) for rx in _C4F_DYN_REFERENCE_RES)


def _c4f_dyn_number_column_count(page, numeric_lines: List[dict]) -> int:
    """
    Approximate C4F number-column detection:
    collect x positions of financial numbers in table lines and cluster them.
    """
    xs = []
    try:
        for ln in numeric_lines or []:
            for w in ln.get('words', []):
                if _dyn_is_financial_number(w.get('text', '')):
                    xs.append(float(w.get('x0', 0)))
        if not xs:
            return 0
        xs = sorted(xs)
        clusters = []
        tolerance = max(14.0, float(getattr(page, 'width', 600)) * 0.025)
        for x in xs:
            if not clusters or abs(x - clusters[-1][-1]) > tolerance:
                clusters.append([x])
            else:
                clusters[-1].append(x)
        # Require at least two numeric appearances in a column cluster so page numbers/stray values do not count.
        return sum(1 for c in clusters if len(c) >= 2)
    except Exception:
        return 0


def _c4f_audit_report_signals(page_text: str) -> Dict[str, bool]:
    """C4F-style audit-report signal set used as fallback only."""
    txt = normalize_pdf_text(page_text or '')
    txt_l = txt.lower()
    top_text = _c4f_top_page_text(txt, max_lines=15)
    return {
        'heading_independent_auditor': bool(re.search(r"\bindependent\s+auditor", top_text, re.I)),
        'heading_auditors_report': bool(re.search(r"\bauditor'?s?\s+report\b", top_text, re.I)),
        'we_have_audited': bool(re.search(r"\bwe\s+have\s+audited\b", txt_l, re.I)),
        'in_our_opinion': bool(re.search(r"\bin\s+our\s+opinion\b", txt_l, re.I)),
        'present_fairly': bool(re.search(r"\bpresent\s+fairly\b", txt_l, re.I)),
        'in_all_material_respects': bool(re.search(r"\bin\s+all\s+material\s+respects\b", txt_l, re.I)),
        'true_and_fair_view': bool(re.search(r"\btrue\s+and\s+fair\s+view\b", txt_l, re.I)),
        'financial_statements_give': bool(re.search(r"\bfinancial\s+statements\s+give\b", txt_l, re.I)),
    }


def _c4f_audit_report_score(page_text: str) -> Tuple[float, Dict[str, bool], Optional[str]]:
    """Score a page for being the main independent auditor report."""
    if not page_text or not page_text.strip():
        return 0.0, {}, 'blank'
    top = _c4f_top_page_text(page_text, max_lines=15)
    if re.search(r"\btable\s+of\s+contents\b|^\s*contents\s*$", top, re.I):
        return 0.0, {}, 'contents_page'
    if re.search(r"\bnotes?\s+to\s+(?:the\s+)?financial\s+statements?\b", top, re.I):
        return 0.0, {}, 'notes_page'
    sig = _c4f_audit_report_signals(page_text)
    score = 0.0
    if sig.get('heading_independent_auditor'):
        score += 0.30
    if sig.get('heading_auditors_report'):
        score += 0.22
    if sig.get('we_have_audited'):
        score += 0.18
    if sig.get('in_our_opinion'):
        score += 0.14
    if sig.get('present_fairly') or sig.get('in_all_material_respects'):
        score += 0.08
    if sig.get('true_and_fair_view') or sig.get('financial_statements_give'):
        score += 0.08
    if not (sig.get('heading_independent_auditor') or sig.get('heading_auditors_report')) and score < 0.45:
        return score, sig, 'no_audit_heading'
    return min(1.0, score), sig, None


def _c4f_find_audit_report_page_fallback(pdf, max_scan_pages: int = 120) -> Optional[Tuple[int, Any, float, Dict[str, bool]]]:
    """Find main auditor report page using C4F-style scoring as secondary fallback."""
    best = None
    try:
        limit = min(len(pdf.pages), max_scan_pages)
        for idx in range(limit):
            page = pdf.pages[idx]
            txt = cached_page_text(page) or ''
            if not txt.strip():
                continue
            score, sig, reason = _c4f_audit_report_score(txt)
            if score <= 0:
                continue
            if best is None or score > best[2]:
                best = (idx + 1, page, score, sig)
            if score >= 0.52:
                return (idx + 1, page, score, sig)
    except Exception:
        pass
    if best and best[2] >= 0.48:
        return best
    return None


def _c4f_audit_heading_or_signal_box(page) -> Optional[dict]:
    """Get a proof box for audit report fallback."""
    try:
        box = find_audit_report_heading_box(page)
        if box:
            return box
    except Exception:
        pass
    for phrase in ["Independent Auditor", "Auditor's Report", "Auditors' Report", "We have audited", "In our opinion"]:
        try:
            box = find_phrase_box_on_page(page, phrase)
            if box:
                return box
        except Exception:
            pass
    return None

# =========================================================
# DYNAMIC TABLE-AWARE STATEMENT FALLBACK HELPERS
# =========================================================
# These helpers DO NOT change your existing validation workflow.
# They are only used as a fallback after each original statement validator fails.
# So your current logic stays primary; the dynamic engine is secondary.

_DYN_FIN_NUM_RE = re.compile(
    r'^\s*'
    r'[(-]?'
    r'[$€£¥₹]?'
    r'\s*'
    r'(?:\d{1,3}(?:[,\s]\d{3})+|\d+)'
    r'(?:\.\d+)?'
    r'\s*'
    r'[)]?'
    r'\s*$'
)
_DYN_DASH_RE = re.compile(r'^\s*[-–—]{1,3}\s*$')
_DYN_YEAR_RE = re.compile(r'^\s*(?:19|20)\d{2}\s*$')

DYN_TITLE_SEARCH_ZONE_RATIO = 0.78
DYN_RIGHT_ZONE_RATIO = 0.55
DYN_MIN_NUMERIC_ROW_RATIO = 0.16
DYN_RIGHT_ZONE_MIN_RATIO = 0.10
DYN_COMBO_MAX_LINES = 3

DYN_EXCLUDED_TITLE_CONTEXTS = [
    r'\bsupplement(?:al|ary)?\b',
    r'\bnotes?\s+(?:to|on|for)\b',
    r'\bcondensed\b',
    r'\babbreviated\b',
    r'\breconciliation\b',
    r'\bbudget(?:ary)?\b',
    r'\bstatistical\b',
    r"\bmanagement(?:'s)?\s+(?:discussion|report|commentary|review)\b",
    r'\bdirectors?\s+report\b',
    r'\bfinancial\s+highlights?\b',
    r'\btable\s+of\s+contents\b',
    r'\bcontents?\b',
    r'\bin\s+accordance\s+with\b',
    r'\bnonmajor\b',
    r'\binternal\s+control\b',
]
DYN_REFERENCE_PATTERNS = [
    r'\bsee\s+(?:the\s+)?notes?\b',
    r'\brefer\s+to\b',
    r'\bthe\s+accompanying\s+notes\b',
    r'\bnote\s+\d+\b',
    r'\bas\s+(?:shown|reported|presented|disclosed)\b',
]
DYN_MANAGEMENT_PATTERNS = [
    r'management\s+discussion',
    r'management\s+report',
    r'financial\s+review',
    r'group\s+management',
    r'directors?\s+report',
    r'strategic\s+report',
    r'business\s+review',
]

DYNAMIC_STATEMENT_CONFIGS = {
    'NET_POSITION': {
        'col': 10, 'proof_dir_name': 'PROOF_DIR_NET_POSITION', 'proof_tag': 'NET_POSITION_DYN', 'label_short': 'Net Position',
        'phrases': ['Statement of Net Position', 'Statements of Net Position', 'Statement of Net Position (Deficit)', 'Statements of Net Position (Deficit)'],
        'title_required_any': [['statement of net position', 'statements of net position'], ['net position']],
        'title_boosters': ['governmental activities', 'business-type activities', 'component units'],
        'title_forbidden': ['proprietary funds', 'cash flows'],
        'group_keywords': {
            'assets': ['assets', 'cash and investments', 'receivables', 'capital assets'],
            'liabilities': ['liabilities', 'deferred inflows', 'payables', 'long term debt'],
            'net': ['net position', 'net investment in capital assets', 'restricted', 'unrestricted'],
        },
        'min_groups': 2, 'min_numeric_rows': 4,
    },
    'ACTIVITIES': {
        'col': 11, 'proof_dir_name': 'PROOF_DIR_ACTIVITIES', 'proof_tag': 'ACTIVITIES_DYN', 'label_short': 'Activities',
        'phrases': ['Statement of Activities', 'Statements of Activities'],
        'title_required_any': [['statement of activities', 'statements of activities'], ['activities']],
        'title_boosters': ['program revenues', 'governmental activities', 'business-type activities'],
        'title_forbidden': ['cash flows', 'budget'],
        'group_keywords': {
            'expenses': ['expenses', 'functions/programs', 'functions programs', 'program expenses', 'function expenses'],
            'revenues': ['program revenues', 'charges for services', 'operating grants', 'capital grants', 'general revenues'],
            'net': ['change in net position', 'net revenue', 'net expense', 'net program', 'net assets'],
        },
        'min_groups': 2, 'min_numeric_rows': 4,
    },
    'BALANCE_SHEET': {
        'col': 12, 'proof_dir_name': 'PROOF_DIR_BALANCE_SHEET', 'proof_tag': 'BALANCE_SHEET_DYN', 'label_short': 'Balance Sheet',
        'phrases': ['Balance Sheet - Governmental Funds', 'Balance Sheet Governmental Funds', 'Balance Sheet'],
        'title_required_any': [['balance sheet']],
        'title_boosters': ['governmental funds', 'major funds'],
        'title_forbidden': ['cash flows', 'notes to'],
        'group_keywords': {
            'assets': ['assets', 'cash and investments', 'receivables', 'due from other funds'],
            'liabilities': ['liabilities', 'payables', 'deferred inflows', 'due to other funds'],
            'funds': ['fund balance', 'fund balances', 'nonspendable', 'restricted', 'committed', 'assigned', 'unassigned'],
        },
        'min_groups': 2, 'min_numeric_rows': 4,
    },
    'REV_EXP_FUND_BAL': {
        'col': 13, 'proof_dir_name': 'PROOF_DIR_REV_EXP_FUND_BAL', 'proof_tag': 'REV_EXP_CHG_FUND_BAL_DYN', 'label_short': 'Rev/Exp/Fund Bal',
        'phrases': ['Statement of Revenues, Expenditures and Changes in Fund Balances', 'Statement of Revenues, Expenditures, and Changes in Fund Balances', 'Statements of Revenues, Expenditures and Changes in Fund Balances', 'Statements of Revenues, Expenditures, and Changes in Fund Balances', 'Statement of Revenues Expenditures and Changes in Fund Balances'],
        'title_required_any': [['revenue'], ['expend'], ['fund balanc']],
        'title_boosters': ['other financing sources', 'governmental funds'],
        'title_forbidden': ['cash flows', 'budget'],
        'group_keywords': {
            'revenues': ['revenues', 'taxes', 'intergovernmental', 'licenses and permits', 'charges for services'],
            'expenditures': ['expenditures', 'current', 'debt service', 'capital outlay'],
            'funds': ['excess deficiency', 'other financing sources', 'net change in fund balances', 'fund balances'],
        },
        'min_groups': 2, 'min_numeric_rows': 4,
    },
    'PROP_NET_POSITION': {
        'col': 14, 'proof_dir_name': 'PROOF_DIR_NET_POSITION', 'proof_tag': 'PROP_NET_POSITION_DYN', 'label_short': 'Prop Net Pos',
        'phrases': ['Statement of Net Position - Proprietary Funds', 'Statement of Net Position Proprietary Funds', 'Statements of Net Position - Proprietary Funds', 'Statements of Net Position Proprietary Funds', 'Statement of Net Position - Internal Service Funds', 'Statements of Net Position - Internal Service Funds'],
        'title_required_any': [['net position'], ['proprietary', 'internal service'], ['fund']],
        'title_boosters': ['enterprise funds', 'internal service funds'],
        'title_forbidden': ['cash flows'],
        'group_keywords': {
            'assets': ['assets', 'current assets', 'capital assets', 'cash and investments'],
            'liabilities': ['liabilities', 'current liabilities', 'noncurrent liabilities', 'deferred inflows'],
            'net': ['net position', 'net investment in capital assets', 'restricted', 'unrestricted'],
        },
        'min_groups': 2, 'min_numeric_rows': 4,
    },
    'REV_EXP_NET_POSITION': {
        'col': 15, 'proof_dir_name': 'PROOF_DIR_REV_EXP_CHG_NET_POSITION', 'proof_tag': 'REV_EXP_CHG_NET_POS_DYN', 'label_short': 'Rev/Exp/Chg Net Pos',
        'phrases': ['Statement of Revenues, Expenses and Changes in Net Position', 'Statement of Revenues, Expenses, and Changes in Net Position', 'Statements of Revenues, Expenses and Changes in Net Position', 'Statements of Revenues, Expenses, and Changes in Net Position', 'Statement of Revenues, Expenses and Changes in Net Assets', 'Statement of Revenues, Expenses, and Changes in Net Assets', 'Statements of Revenues, Expenses and Changes in Net Assets', 'Statements of Revenues, Expenses, and Changes in Net Assets'],
        'title_required_any': [['revenue'], ['expense'], ['net position', 'net assets']],
        'title_boosters': ['operating revenues', 'operating expenses', 'nonoperating'],
        'title_forbidden': ['cash flows', 'budget'],
        'group_keywords': {
            'revenues': ['operating revenues', 'nonoperating revenues', 'revenues'],
            'expenses': ['operating expenses', 'depreciation', 'expenses'],
            'net': ['change in net position', 'change in net assets', 'net position', 'net assets'],
        },
        'min_groups': 2, 'min_numeric_rows': 4,
    },
    'CASH_RECEIPTS_DISB': {
        'col': 16, 'proof_dir_name': 'PROOF_DIR_CASH_RECEIPTS_DISBURSEMENTS', 'proof_tag': 'CASH_RECEIPTS_DISB_DYN', 'label_short': 'Cash Receipts/Disb',
        'phrases': ['Statement of Cash Receipts and Disbursements', 'Statements of Cash Receipts and Disbursements', 'Statement of Receipts, Disbursements, and Cash and Investment Balances', 'Statements of Receipts, Disbursements, and Cash and Investment Balances', 'Statement of Receipts, Disbursements, and Changes in Cash and Investment Balances', 'Statements of Receipts, Disbursements, and Changes in Cash and Investment Balances'],
        'title_required_any': [['receipt'], ['disbursement'], ['cash', 'investment balanc']],
        'title_boosters': ['ending cash and investment balances', 'beginning cash and investment balances'],
        'title_forbidden': ['cash flow', 'cash flows'],
        'group_keywords': {
            'receipts': ['receipts', 'cash receipts', 'total receipts'],
            'disbursements': ['disbursements', 'cash disbursements', 'total disbursements'],
            'cash': ['cash and investment balances', 'ending cash and investment balances', 'beginning cash and investment balances'],
        },
        'min_groups': 2, 'min_numeric_rows': 3,
    },
}

_DYN_EXCLUDED_TITLE_RES = [re.compile(p, re.I) for p in DYN_EXCLUDED_TITLE_CONTEXTS]
_DYN_EXCLUDED_TITLE_RES.extend(_C4F_DYN_EXTRA_EXCLUDED_TITLE_RES)
_DYN_REFERENCE_RES = [re.compile(p, re.I) for p in DYN_REFERENCE_PATTERNS]
_DYN_MANAGEMENT_RES = [re.compile(p, re.I) for p in DYN_MANAGEMENT_PATTERNS]


def _dyn_clean_text(text: str) -> str:
    """Clean ``text`` to lowercase alphanumerics and single spaces for the dynamic statement detector."""
    s = normalize_pdf_text(text or '').lower()
    s = re.sub(r"[^a-z0-9/&'\-]+", ' ', s)
    s = re.sub(r'\s+', ' ', s).strip()
    return s


def _dyn_is_financial_number(text: str) -> bool:
    """Return ``True`` if ``text`` looks like a financial figure (currency/large comma-grouped number)."""
    clean = (text or '').strip()
    if not clean:
        return False
    if _DYN_DASH_RE.match(clean):
        return True
    if clean.endswith('%'):
        return False
    if _DYN_YEAR_RE.match(clean):
        return False
    digits_only = re.sub(r'[^\d]', '', clean)
    if digits_only and len(digits_only) <= 2 and ',' not in clean and '.' not in clean:
        try:
            val = int(digits_only)
            if val <= 50 and not (clean.startswith('(') and clean.endswith(')')):
                return False
        except Exception:
            pass
    return bool(_DYN_FIN_NUM_RE.match(clean))


def _dyn_union_box(items):
    """Return a bounding box that is the union of all boxes in ``items`` (or ``None`` if empty)."""
    objs = [x for x in (items or []) if x]
    if not objs:
        return None
    return {
        'x0': min(float(o['x0']) for o in objs),
        'top': min(float(o['top']) for o in objs),
        'x1': max(float(o['x1']) for o in objs),
        'bottom': max(float(o['bottom']) for o in objs),
    }


def _dyn_get_visual_lines(page, y_tol: float = 3.0):
    """Group page words into visual lines by their ``top`` coordinate within ``y_tol`` tolerance."""
    words = cached_page_words(page, use_text_flow=True)
    if not words:
        return []
    words = sorted(words, key=lambda w: (round(float(w.get('top', 0)) / y_tol), float(w.get('x0', 0))))
    lines = []
    current = []
    current_top = None
    for w in words:
        top = float(w.get('top', 0))
        if current_top is None or abs(top - current_top) <= y_tol:
            current.append(w)
            if current_top is None:
                current_top = top
        else:
            current = sorted(current, key=lambda x: float(x.get('x0', 0)))
            txt = ' '.join((c.get('text', '') or '').strip() for c in current).strip()
            if txt:
                lines.append({'words': current, 'text': txt, 'y_top': min(float(c.get('top', 0)) for c in current)})
            current = [w]
            current_top = top
    if current:
        current = sorted(current, key=lambda x: float(x.get('x0', 0)))
        txt = ' '.join((c.get('text', '') or '').strip() for c in current).strip()
        if txt:
            lines.append({'words': current, 'text': txt, 'y_top': min(float(c.get('top', 0)) for c in current)})
    return lines


def _dyn_text_has_any(clean_text: str, alternatives) -> bool:
    """Return ``True`` if ``clean_text`` contains any phrase in ``alternatives``."""
    for alt in alternatives:
        ac = _dyn_clean_text(alt)
        if ac and ac in clean_text:
            return True
    return False


def _dyn_header_combo_candidates(page, max_lines: int = 12):
    """Build candidate combined header strings from the top ``max_lines`` visual lines of a page."""
    lines = _dyn_get_visual_lines(page)
    if not lines:
        return []
    zone_bottom = float(page.height) * DYN_TITLE_SEARCH_ZONE_RATIO
    top_lines = [ln for ln in lines if ln['text'] and ln['y_top'] <= zone_bottom][:max_lines]
    combos = []
    for i in range(len(top_lines)):
        for span in range(1, min(DYN_COMBO_MAX_LINES, len(top_lines) - i) + 1):
            subset = top_lines[i:i + span]
            txt = ' '.join(ln['text'] for ln in subset).strip()
            if txt:
                combos.append({'text': txt, 'clean': _dyn_clean_text(txt), 'box': _dyn_union_box([w for ln in subset for w in ln['words']])})
    combos.sort(key=lambda c: -len(c['clean']))
    return combos


def _dyn_combo_is_excluded(combo_text: str) -> bool:
    """Return ``True`` if a combined header string ``combo_text`` matches an excluded pattern (e.g. reconciliation)."""
    t = normalize_pdf_text(combo_text or '').lower()
    return (
        any(rx.search(t) for rx in _DYN_EXCLUDED_TITLE_RES)
        or any(rx.search(t) for rx in _DYN_MANAGEMENT_RES)
        or any(rx.search(t) for rx in _DYN_REFERENCE_RES)
        or _c4f_dyn_title_context_is_excluded(t)
    )


def _dyn_find_heading_candidate(page, cfg, max_lines: int = 12):
    """Find a statement-heading candidate on a page for config ``cfg``; returns heading info or ``None``."""
    best_exact = None
    for phrase in cfg.get('phrases', []):
        box = find_phrase_box_on_page_top_lines(page, phrase, max_lines=max_lines)
        if box:
            c = _dyn_clean_text(phrase)
            score = 2.8 + min(0.5, len(c.split()) * 0.03)
            cand = {'box': box, 'text': phrase, 'clean': c, 'score': score}
            if best_exact is None or cand['score'] > best_exact['score']:
                best_exact = cand
    best_combo = None
    for combo in _dyn_header_combo_candidates(page, max_lines=max_lines):
        ctext = combo['clean']
        if not ctext or _dyn_combo_is_excluded(combo['text']):
            continue
        if any(_dyn_clean_text(forb) in ctext for forb in cfg.get('title_forbidden', [])):
            continue
        score = 0.0
        exact_hits = sum(1 for ph in cfg.get('phrases', []) if _dyn_clean_text(ph) in ctext)
        if exact_hits:
            score += 2.2 + 0.2 * min(exact_hits, 2)
        req_groups = cfg.get('title_required_any', [])
        req_hits = 0
        for group in req_groups:
            if _dyn_text_has_any(ctext, group):
                req_hits += 1
                score += 0.75
        if req_groups and req_hits < max(1, len(req_groups) - 1) and exact_hits == 0:
            continue
        for booster in cfg.get('title_boosters', []):
            if _dyn_clean_text(booster) in ctext:
                score += 0.15
        if 'statement' in ctext or 'statements' in ctext:
            score += 0.20
        if score < 1.20:
            continue
        cand = {'box': combo['box'], 'text': combo['text'], 'clean': ctext, 'score': score}
        if best_combo is None or cand['score'] > best_combo['score']:
            best_combo = cand
    if best_exact and best_combo:
        return best_exact if best_exact['score'] >= best_combo['score'] else best_combo
    return best_exact or best_combo


def _dyn_collect_table_evidence(page, cut_top: float, cfg):
    """
    C4F-enhanced table evidence:
    - counts financial-number rows
    - verifies numbers appear in right-side columns
    - clusters numeric x-positions to prove table-column structure
    - validates grouped statement keywords below the heading
    """
    lines = _dyn_get_visual_lines(page)
    if not lines:
        return None
    lines = [ln for ln in lines if min(float(w.get('top', 0)) for w in ln['words']) >= cut_top][:55]
    if not lines:
        return None

    numeric_lines = []
    right_numeric_lines = []
    keyword_group_hits = {k: False for k in cfg.get('group_keywords', {})}
    evidence_lines = []

    for ln in lines:
        txt = normalize_pdf_text(ln['text']).lower()
        clean = _dyn_clean_text(txt)
        if not clean:
            continue
        if any(rx.search(txt) for rx in _DYN_REFERENCE_RES) or _c4f_dyn_is_reference_line(txt):
            continue

        evidence_lines.append(ln)
        has_num = any(_dyn_is_financial_number(w.get('text', '')) for w in ln['words'])
        has_right_num = any(
            _dyn_is_financial_number(w.get('text', ''))
            and float(w.get('x0', 0)) >= float(page.width) * DYN_RIGHT_ZONE_RATIO
            for w in ln['words']
        )
        if has_num:
            numeric_lines.append(ln)
        if has_right_num:
            right_numeric_lines.append(ln)

        for grp, kws in cfg.get('group_keywords', {}).items():
            if keyword_group_hits[grp]:
                continue
            if any(_dyn_clean_text(kw) in clean for kw in kws):
                keyword_group_hits[grp] = True

    total_lines = max(1, len(evidence_lines))
    numeric_count = len(numeric_lines)
    right_count = len(right_numeric_lines)
    number_column_count = _c4f_dyn_number_column_count(page, numeric_lines)
    matched_groups = sum(1 for v in keyword_group_hits.values() if v)
    total_groups = max(1, len(keyword_group_hits))
    numeric_ratio = numeric_count / float(total_lines)
    right_ratio = right_count / float(total_lines)

    confidence = 0.0
    confidence += min(0.32, numeric_count * 0.045)
    confidence += min(0.22, matched_groups / float(total_groups) * 0.22)
    confidence += 0.14 if numeric_ratio >= DYN_MIN_NUMERIC_ROW_RATIO else numeric_ratio * 0.45
    confidence += 0.14 if right_ratio >= DYN_RIGHT_ZONE_MIN_RATIO else right_ratio * 0.60
    if number_column_count >= 1:
        confidence += 0.08
    if number_column_count >= 2:
        confidence += 0.05
    if numeric_count >= cfg.get('min_numeric_rows', 4):
        confidence += 0.08

    subset = numeric_lines[:12] if numeric_lines else evidence_lines[:12]
    table_box = _dyn_union_box([w for ln in subset for w in ln['words']]) if subset else None
    return {
        'numeric_count': numeric_count,
        'right_count': right_count,
        'number_column_count': number_column_count,
        'matched_groups': matched_groups,
        'total_groups': total_groups,
        'table_box': table_box,
        'confidence': confidence,
    }


def _dyn_detect_statement_page(page, page_text: str, fy_end_value, cfg, max_lines: int = 12):
    """Detect whether a page is the target financial statement for ``cfg`` with matching FYE.

Returns a detection dict with keys such as ``pass``, ``title_candidate``, ``confidence``,
``heading``, ``fye_box`` and ``table_box``.
    """
    txt = page_text or ''
    if not txt.strip() or is_excluded_fin_stmt_page(txt):
        return {'title_candidate': False, 'pass': False}

    # C4F-style hard page rejection for notes/prose/audit/management pages.
    reject_reason = _c4f_dyn_page_rejection_reason(txt)
    if reject_reason:
        return {'title_candidate': False, 'pass': False, 'reason': reject_reason}

    clean_page = normalize_pdf_text(txt).lower()
    if any(rx.search(clean_page) for rx in _DYN_MANAGEMENT_RES):
        return {'title_candidate': False, 'pass': False}

    heading = _dyn_find_heading_candidate(page, cfg, max_lines=max_lines)
    if not heading:
        return {'title_candidate': False, 'pass': False}

    fye_box = find_fye_box_on_page_top_lines(page, fy_end_value, max_lines=max_lines)
    if not fye_box:
        return {'title_candidate': True, 'pass': False, 'heading': heading}

    if excluded_keywords_above_statement_and_fye(page, heading['box'], fye_box, max_lines=max_lines):
        return {'title_candidate': True, 'pass': False, 'heading': heading}

    cut_top = max(float(heading['box']['bottom']), float(fye_box['bottom'])) + 2.0
    evidence = _dyn_collect_table_evidence(page, cut_top, cfg)
    if not evidence:
        return {'title_candidate': True, 'pass': False, 'heading': heading, 'fye_box': fye_box}

    min_groups = cfg.get('min_groups', 2)
    min_numeric = min(3, int(cfg.get('min_numeric_rows', 4)))
    has_table_shape = (evidence.get('right_count', 0) >= 1 or evidence.get('number_column_count', 0) >= 1)
    pass_ok = (
        evidence['matched_groups'] >= min_groups
        and evidence['numeric_count'] >= min_numeric
        and has_table_shape
        and evidence['confidence'] >= 0.45
    )
    total_conf = min(1.0, 0.35 + heading['score'] * 0.12 + evidence['confidence'])
    return {
        'title_candidate': True,
        'pass': pass_ok,
        'heading': heading,
        'fye_box': fye_box,
        'table_box': evidence.get('table_box'),
        'confidence': total_conf,
        'evidence': evidence,
    }


def _dynamic_statement_proof_dir(cfg):
    """Resolve and return the proof-output directory ``Path`` for config ``cfg`` (via its ``proof_dir_name``)."""
    return globals()[cfg['proof_dir_name']]


def _write_dynamic_statement_validation(processing_id: int, fy_end_value, pdf_path: Optional[Path], cfg_key: str, show_popup: bool = True, previous_result: Optional[dict] = None) -> dict:
    """Dynamic table-aware fallback for a financial-statement validation.

Scans pages for the statement described by ``cfg_key`` with the expected ``fy_end_value``,
saves a proof PNG, and returns a result dict with ``status``/``page``/``proof_file``/``boxes``/``title``.
Does NOT write to the parquet (the caller batches that).
    """
    cfg = DYNAMIC_STATEMENT_CONFIGS[cfg_key]
    label_short = cfg['label_short']
    proof_dir = _dynamic_statement_proof_dir(cfg)
    proof_tag = cfg['proof_tag']
    HEADER_MAX_LINES = 12
    if not pdf_path or not Path(pdf_path).exists():
        return previous_result or {'status': 'FAIL', 'page': None, 'proof_file': None, 'boxes': [], 'title': f'ProcessingId {processing_id} | {label_short} = FAIL (PDF missing)'}
    try:
        with open_pdf_maybe_shared(pdf_path) as pdf:
            best_fail = None
            first_candidate_page = previous_result.get('page') if isinstance(previous_result, dict) else None
            for idx, page in enumerate(pdf.pages):
                page_num = idx + 1
                txt = cached_page_text(page) or ''
                if not txt.strip():
                    continue
                if TOC_STOP_ENABLED and is_table_of_contents_page(txt):
                    continue
                result = _dyn_detect_statement_page(page, txt, fy_end_value, cfg, max_lines=HEADER_MAX_LINES)
                if result.get('title_candidate') and first_candidate_page is None:
                    first_candidate_page = page_num
                if not result.get('pass'):
                    if result.get('title_candidate'):
                        score = float(result.get('confidence') or 0.0)
                        if best_fail is None or score > best_fail[0]:
                            best_fail = (score, page_num)
                    continue
                proof_file = proof_dir / f'Row{processing_id}_{proof_tag}_PASS_Page{page_num}.png'
                im = page.to_image(resolution=PROOF_IMAGE_RESOLUTION)
                boxes = []
                for box in [result['heading']['box'], result.get('fye_box'), result.get('table_box')]:
                    if box:
                        im.draw_rect(box, stroke=HIGHLIGHT_STROKE, fill=HIGHLIGHT_FILL_RGBA)
                        boxes.append(box)
                im.save(proof_file)
                title = f'ProcessingId {processing_id} | {label_short} = PASS | Page {page_num}'
                if show_popup:
                    show_proof_popup(proof_file, title, seconds=PROOF_POPUP_SECONDS)
                return {'status': 'PASS', 'page': page_num, 'proof_file': proof_file, 'boxes': boxes, 'title': title}
            fail_page = first_candidate_page or (best_fail[1] if best_fail else None)
            return previous_result or {'status': 'FAIL', 'page': fail_page, 'proof_file': None, 'boxes': [], 'title': f'ProcessingId {processing_id} | {label_short} = FAIL'}
    except Exception:
        return previous_result or {'status': 'FAIL', 'page': None, 'proof_file': None, 'boxes': [], 'title': f'ProcessingId {processing_id} | {label_short} = FAIL'}


def write_statement_of_net_position_validation(
    processing_id: int,
    fy_end_value,
    pdf_path: Optional[Path],
    show_popup: bool = True
) -> dict:
    """
    Hybrid wrapper: keep existing logic first, then apply dynamic table-aware fallback
    only if original logic fails.

    - Does NOT write to parquet (parent run_validations_for_row batches the write)
    """
    res = _orig_write_statement_of_net_position_validation(processing_id, fy_end_value, pdf_path, show_popup=show_popup)
    if isinstance(res, dict) and str(res.get('status', '')).upper().startswith('PASS'):
        return res
    return _write_dynamic_statement_validation(
        processing_id=processing_id,
        fy_end_value=fy_end_value,
        pdf_path=pdf_path,
        cfg_key='NET_POSITION',
        show_popup=show_popup,
        previous_result=res if isinstance(res, dict) else None,
    )

def write_statement_of_activities_validation(
    processing_id: int,
    fy_end_value,
    pdf_path: Optional[Path],
    show_popup: bool = True
) -> dict:
    """
    Hybrid wrapper: keep existing logic first, then apply dynamic table-aware fallback
    only if original logic fails.

    - Does NOT write to parquet (parent run_validations_for_row batches the write)
    """
    res = _orig_write_statement_of_activities_validation(processing_id, fy_end_value, pdf_path, show_popup=show_popup)
    if isinstance(res, dict) and str(res.get('status', '')).upper().startswith('PASS'):
        return res
    return _write_dynamic_statement_validation(
        processing_id=processing_id,
        fy_end_value=fy_end_value,
        pdf_path=pdf_path,
        cfg_key='ACTIVITIES',
        show_popup=show_popup,
        previous_result=res if isinstance(res, dict) else None,
    )

def write_balance_sheet_validation(
    processing_id: int,
    fy_end_value,
    pdf_path: Optional[Path],
    show_popup: bool = True
) -> dict:
    """
    Hybrid wrapper: keep existing logic first, then apply dynamic table-aware fallback
    only if original logic fails.

    - Does NOT write to parquet (parent run_validations_for_row batches the write)
    """
    res = _orig_write_balance_sheet_validation(processing_id, fy_end_value, pdf_path, show_popup=show_popup)
    if isinstance(res, dict) and str(res.get('status', '')).upper().startswith('PASS'):
        return res
    return _write_dynamic_statement_validation(
        processing_id=processing_id,
        fy_end_value=fy_end_value,
        pdf_path=pdf_path,
        cfg_key='BALANCE_SHEET',
        show_popup=show_popup,
        previous_result=res if isinstance(res, dict) else None,
    )

def write_rev_exp_changes_fund_balances_validation(
    processing_id: int,
    fy_end_value,
    pdf_path: Optional[Path],
    show_popup: bool = True
) -> dict:
    """
    Hybrid wrapper: keep existing logic first, then apply dynamic table-aware fallback
    only if original logic fails.

    - Does NOT write to parquet (parent run_validations_for_row batches the write)
    """
    res = _orig_write_rev_exp_changes_fund_balances_validation(processing_id, fy_end_value, pdf_path, show_popup=show_popup)
    if isinstance(res, dict) and str(res.get('status', '')).upper().startswith('PASS'):
        return res
    return _write_dynamic_statement_validation(
        processing_id=processing_id,
        fy_end_value=fy_end_value,
        pdf_path=pdf_path,
        cfg_key='REV_EXP_FUND_BAL',
        show_popup=show_popup,
        previous_result=res if isinstance(res, dict) else None,
    )

def write_statement_of_net_position_proprietary_funds_validation(
    processing_id: int,
    fy_end_value,
    pdf_path: Optional[Path],
    show_popup: bool = True
) -> dict:
    """
    Hybrid wrapper: keep existing logic first, then apply dynamic table-aware fallback
    only if original logic fails.

    - Does NOT write to parquet (parent run_validations_for_row batches the write)
    """
    res = _orig_write_statement_of_net_position_proprietary_funds_validation(processing_id, fy_end_value, pdf_path, show_popup=show_popup)
    if isinstance(res, dict) and str(res.get('status', '')).upper().startswith('PASS'):
        return res
    return _write_dynamic_statement_validation(
        processing_id=processing_id,
        fy_end_value=fy_end_value,
        pdf_path=pdf_path,
        cfg_key='PROP_NET_POSITION',
        show_popup=show_popup,
        previous_result=res if isinstance(res, dict) else None,
    )

def write_rev_exp_changes_net_position_validation(
    processing_id: int,
    fy_end_value,
    pdf_path: Optional[Path],
    show_popup: bool = True
) -> dict:
    """
    Hybrid wrapper: keep existing logic first, then apply dynamic table-aware fallback
    only if original logic fails.

    - Does NOT write to parquet (parent run_validations_for_row batches the write)
    """
    res = _orig_write_rev_exp_changes_net_position_validation(processing_id, fy_end_value, pdf_path, show_popup=show_popup)
    if isinstance(res, dict) and str(res.get('status', '')).upper().startswith('PASS'):
        return res
    return _write_dynamic_statement_validation(
        processing_id=processing_id,
        fy_end_value=fy_end_value,
        pdf_path=pdf_path,
        cfg_key='REV_EXP_NET_POSITION',
        show_popup=show_popup,
        previous_result=res if isinstance(res, dict) else None,
    )

def write_statement_of_cash_receipts_disbursements_validation(
    processing_id: int,
    fy_end_value,
    pdf_path: Optional[Path],
    show_popup: bool = True
) -> dict:
    """
    Hybrid wrapper: keep existing logic first, then apply dynamic table-aware fallback
    only if original logic fails.

    - Does NOT write to parquet (parent run_validations_for_row batches the write)
    """
    res = _orig_write_statement_of_cash_receipts_disbursements_validation(processing_id, fy_end_value, pdf_path, show_popup=show_popup)
    if isinstance(res, dict) and str(res.get('status', '')).upper().startswith('PASS'):
        return res
    return _write_dynamic_statement_validation(
        processing_id=processing_id,
        fy_end_value=fy_end_value,
        pdf_path=pdf_path,
        cfg_key='CASH_RECEIPTS_DISB',
        show_popup=show_popup,
        previous_result=res if isinstance(res, dict) else None,
    )


# =========================================================
# PDF VALIDATION HELPERS (Detailed Validation Check)
# =========================================================

# =========================================================
# PROCESS DOWNLOADED EXCEL -> UPDATE MASTER -> DOWNLOAD PDF
# =========================================================
# VALIDATION: Write name validation.
def write_name_validation(
    processing_id: int,
    issuer_name: str,
    uei: str,
    pdf_path: Optional[Path],
    show_popup: bool = True
) -> dict:
    """
    NAME validation
    - Saves proof PNG to PROOF_DIR_NAME
    - Returns dict with: status, page, proof_file, boxes, title
    - Does NOT write to parquet (parent run_validations_for_row batches the write)
    """
    if not pdf_path or not Path(pdf_path).exists():
        return {
            "status": "FAIL",
            "page": None,
            "proof_file": None,
            "boxes": [],
            "title": f"ProcessingId {processing_id} | Name Check = FAIL (PDF missing)"
        }

    proof_file = PROOF_DIR_NAME / f"Row{processing_id}_NAME_{sanitize_filename(issuer_name, 60)}_{sanitize_filename(uei, 25)}.png"

    res = visible_name_validation_with_proof(Path(pdf_path), issuer_name, proof_file)
    status = res.get("status", "FAIL")
    page_no = res.get("page")

    # Try to compute a box for bundling (best-effort)
    boxes = []
    if status == "PASS" and page_no:
        try:
            with open_pdf_maybe_shared(pdf_path) as pdf:
                if 1 <= int(page_no) <= len(pdf.pages):
                    pg = pdf.pages[int(page_no) - 1]
                    b = find_phrase_box_on_page(pg, issuer_name)
                    if b:
                        boxes.append(b)
        except Exception:
            pass

    title = f"ProcessingId {processing_id} | Name Check = {status}" + (f" | Page {page_no}" if page_no else "")

    if show_popup and proof_file.exists():
        show_proof_popup(proof_file, title, seconds=PROOF_POPUP_SECONDS)

    return {
        "status": status,
        "page": page_no,
        "proof_file": proof_file,
        "boxes": boxes,
        "title": title
    }

# VALIDATION: Write state validation.

# =========================================================
# STATE VALIDATION v4 INTEGRATION HELPERS
# =========================================================
# Added from State.py v4 fixed logic.
# Scope: STATE validation only. No other validation logic is changed.
# Proof behavior: save ONLY the final PASS State image in PROOF_DIR_STATE.

STATE_V4_US_STATE_CODES = {
    "Alabama":"AL","Alaska":"AK","Arizona":"AZ","Arkansas":"AR","California":"CA","Colorado":"CO",
    "Connecticut":"CT","Delaware":"DE","Florida":"FL","Georgia":"GA","Hawaii":"HI","Idaho":"ID",
    "Illinois":"IL","Indiana":"IN","Iowa":"IA","Kansas":"KS","Kentucky":"KY","Louisiana":"LA",
    "Maine":"ME","Maryland":"MD","Massachusetts":"MA","Michigan":"MI","Minnesota":"MN",
    "Mississippi":"MS","Missouri":"MO","Montana":"MT","Nebraska":"NE","Nevada":"NV",
    "New Hampshire":"NH","New Jersey":"NJ","New Mexico":"NM","New York":"NY",
    "North Carolina":"NC","North Dakota":"ND","Ohio":"OH","Oklahoma":"OK","Oregon":"OR",
    "Pennsylvania":"PA","Rhode Island":"RI","South Carolina":"SC","South Dakota":"SD",
    "Tennessee":"TN","Texas":"TX","Utah":"UT","Vermont":"VT","Virginia":"VA","Washington":"WA",
    "West Virginia":"WV","Wisconsin":"WI","Wyoming":"WY","District of Columbia":"DC",
}
STATE_V4_ABBR_TO_STATE = {v.upper(): k for k, v in STATE_V4_US_STATE_CODES.items()}
STATE_V4_NAME_TO_ABBR = {k.lower(): v for k, v in STATE_V4_US_STATE_CODES.items()}
STATE_V4_NAME_TO_ABBR["dc"] = "DC"


def state_v4_result_dict(status: str, method: str, reason: str, page=None, proof_file=None, boxes=None, extra=None) -> dict:
    """Build and return a standardized State-v4 result dict from the given fields."""
    d = {
        "status": status,
        "method": method,
        "reason": reason,
        "page": page,
        "proof_file": str(proof_file) if proof_file else None,
        "boxes": boxes or [],
    }
    if extra:
        d.update(extra)
    return d


def state_v4_clean_for_contains_match(s: str) -> str:
    """Clean ``s`` to lowercase alphanumerics and single spaces (State-v4 matcher)."""
    s = normalize_pdf_text(s or "").lower()
    s = re.sub(r"[^a-z0-9 ]+", " ", s)
    return re.sub(r"\s+", " ", s).strip()


def state_v4_tokenize(s: str) -> List[str]:
    """Tokenize ``s`` into cleaned lowercase tokens (State-v4 matcher)."""
    return [t for t in state_v4_clean_for_contains_match(s).split() if t]


def state_v4_find_token_window(page_tokens: List[str], target_tokens: List[str]) -> Tuple[Optional[int], Optional[int]]:
    """Find ``target_tokens`` within ``page_tokens``; returns ``(start, end)`` or ``(None, None)``."""
    n, m = len(page_tokens), len(target_tokens)
    if not n or not m or m > n:
        return None, None
    for i in range(n - m + 1):
        if page_tokens[i:i + m] == target_tokens:
            return i, i + m - 1
    return None, None


def state_v4_box_from_words(words: List[dict]) -> Optional[dict]:
    """Return a bounding box spanning the given word dicts, or ``None`` if empty."""
    words = [w for w in (words or []) if w]
    if not words:
        return None
    return {
        "x0": min(float(w["x0"]) for w in words),
        "top": min(float(w["top"]) for w in words),
        "x1": max(float(w["x1"]) for w in words),
        "bottom": max(float(w["bottom"]) for w in words),
    }


def state_v4_page_tokens_with_boxes(page) -> Tuple[List[str], List[dict]]:
    """Return aligned ``(tokens, word_map)`` for a page (State-v4 matcher)."""
    page_tokens, token_word_map = [], []
    for w in cached_page_words(page, use_text_flow=True):
        for tok in state_v4_tokenize(w.get("text", "")):
            page_tokens.append(tok)
            token_word_map.append(w)
    return page_tokens, token_word_map


def state_v4_save_final_proof(page, boxes: List[dict], proof_path: Path) -> bool:
    """Save ONLY the final PASS State proof image."""
    proof_path.parent.mkdir(parents=True, exist_ok=True)
    try:
        im = page.to_image(resolution=PROOF_IMAGE_RESOLUTION)
        for b in boxes or []:
            if b:
                im.draw_rect(b, stroke=HIGHLIGHT_STROKE, fill=HIGHLIGHT_FILL_RGBA)
        im.save(str(proof_path))
        return True
    except Exception as e:
        print(f"[WARN] Could not save final State proof image: {e}")
        return False


def state_v4_state_target_tokens(master_state: str) -> Tuple[str, str, List[str], List[str]]:
    """Return ``(full_state_tokens, abbr_tokens)`` for matching ``master_state`` (full name and 2-letter code)."""
    raw = normalize(master_state)
    raw_upper = raw.upper()
    raw_lower = raw.lower()
    if raw_upper in STATE_V4_ABBR_TO_STATE:
        abbr = raw_upper
        full_state = STATE_V4_ABBR_TO_STATE[abbr]
    elif raw_lower in STATE_V4_NAME_TO_ABBR:
        abbr = STATE_V4_NAME_TO_ABBR[raw_lower].upper()
        full_state = STATE_V4_ABBR_TO_STATE.get(abbr, raw.title())
    else:
        abbr = raw_upper
        full_state = STATE_V4_ABBR_TO_STATE.get(abbr, raw.title() if raw else "")
    return abbr, full_state, state_v4_tokenize(full_state), [abbr.lower()] if abbr else []


def state_v4_issuer_token_variants(issuer_name: str) -> List[List[str]]:
    """Return tokenized variants of ``issuer_name`` for flexible issuer matching."""
    issuer_name = normalize_pdf_text(issuer_name or "").strip()
    variants, seen = [], set()

    def add(txt: str):
        toks = state_v4_tokenize(txt)
        if toks and tuple(toks) not in seen:
            seen.add(tuple(toks))
            variants.append(toks)

    if not issuer_name:
        return []

    clean = re.sub(r"\s+", " ", issuer_name).strip()
    no_the = re.sub(r"^\s*the\s+", "", clean, flags=re.I).strip()

    add(clean)
    add(no_the)

    # County suffix -> County of X and X-only variant.
    # Critical for reports like "County of Luna, New Mexico".
    m = re.match(r"^(.+?)\s+county$", no_the, flags=re.I)
    if m:
        base = m.group(1).strip()
        if base:
            add(base)
            add(f"{base} County")
            add(f"County of {base}")
            add(f"The County of {base}")

    # County of X -> X County and X-only variant.
    m = re.match(r"^(?:the\s+)?county\s+of\s+(.+)$", no_the, flags=re.I)
    if m:
        base = m.group(1).strip()
        if base:
            add(base)
            add(f"{base} County")
            add(f"County of {base}")
            add(f"The County of {base}")

    return variants


def state_v4_is_table_of_contents_page(page_text: str) -> bool:
    """Robust TOC detector used only by the integrated State v4 logic."""
    lines = [ln.strip() for ln in (page_text or "").splitlines() if ln.strip()]
    top8 = normalize_pdf_text(" ".join(lines[:8])).lower()
    top20 = normalize_pdf_text(" ".join(lines[:20])).lower()
    compact8 = re.sub(r"[^a-z]", "", top8)
    compact20 = re.sub(r"[^a-z]", "", top20)

    # Do not classify a true auditor report as TOC merely because its paragraph
    # says "listed in the table of contents".
    audit_signals = bool(re.search(r"\bwe\s+have\s+audited\b|\bin\s+our\s+opinion\b", top20, re.I))
    if audit_signals:
        return False

    return bool(
        compact8.startswith("tableofcontents")
        or "tableofcontentscontinued" in compact20
        or "tableofcontents" in compact20
        or re.search(r"^\s*table\s+of\s+contents\b", top8, re.I)
    )


def state_v4_is_main_auditor_report_page(page_text: str) -> bool:
    """
    Robust State-specific main auditor report detector.
    Avoids TOC/divider pages and waits for real audit-body text.
    """
    if not page_text or state_v4_is_table_of_contents_page(page_text):
        return False

    txt = normalize_pdf_text(page_text)
    low = txt.lower()
    lines = [normalize_pdf_text(ln.strip()) for ln in (page_text or "").splitlines() if ln.strip()]
    top_lines = lines[:35]
    top = normalize_pdf_text(" ".join(top_lines)).lower()

    # Reject later compliance / single-audit reports unless it is clearly the simple main heading.
    if re.search(
        r"\b(compliance|major\s+federal\s+program|uniform\s+guidance|internal\s+control\s+over\s+compliance|internal\s+control\s+over\s+financial\s+reporting)\b",
        top,
        re.I,
    ):
        simple_heading = any(
            re.sub(r"[^a-z]", "", ln.lower()) in {"independentauditorsreport", "independentauditorreport"}
            for ln in top_lines
        )
        if not simple_heading:
            return False

    if re.search(r"\bsee\s+accompanying\s+(?:independent\s+)?auditor", top, re.I):
        return False

    has_audit_body = bool(
        re.search(r"\bwe\s+have\s+audited\b", low, re.I)
        or re.search(r"\breport\s+on\s+the\s+audit\s+of\s+the\s+financial\s+statements\b", low, re.I)
        or re.search(r"\bin\s+our\s+opinion\b", low, re.I)
        or re.search(r"\bopinions?\b", low, re.I)
    )

    for ln in top_lines:
        l = ln.lower().strip(" .:-")
        compact = re.sub(r"[^a-z]", "", l)
        if compact in {"independentauditorsreport", "independentauditorreport"}:
            return has_audit_body
        if re.search(r"^independent\s+au\s*ditor(?:s)?'?s?\s+report$", l, re.I):
            return has_audit_body
        if re.search(r"^report\s+of\s+(?:the\s+)?independent\s+auditor", l, re.I):
            return has_audit_body

    return bool(
        re.search(r"\bindependent\s+au\s*ditor", top, re.I)
        and re.search(r"\bwe\s+have\s+audited\b", low, re.I)
        and re.search(r"\bin\s+our\s+opinion\b", low, re.I)
    )


def state_v4_is_statistical_supplementary_or_schedule_page(page_text: str) -> bool:
    """Return ``True`` if the page is a statistical/supplementary/schedule section (excluded for state matching)."""
    top = normalize_pdf_text(" ".join((page_text or "").splitlines()[:30])).lower()
    reject_patterns = [
        r"\bstatistical\b", r"\btable\s+\d+\b", r"\bprincipal\s+employers\b",
        r"\brequired\s+supplementary\s+information\b", r"\bother\s+supplementary\s+information\b", r"\bsupplementary\s+information\b",
        r"\bschedule\s+\d+[a-z]?\b", r"\bcombining\s+schedule\b", r"\bschedule\s+of\s+expenditures\s+of\s+federal\s+awards\b",
        r"\bcompliance\b", r"\bsingle\s+audit\b", r"\buniform\s+guidance\b",
    ]
    return any(re.search(p, top, re.I) for p in reject_patterns)


def state_v4_has_main_financial_statement_heading(page_text: str) -> bool:
    """Return ``True`` if the page contains a main financial-statement heading."""
    if not page_text or state_v4_is_table_of_contents_page(page_text) or state_v4_is_statistical_supplementary_or_schedule_page(page_text):
        return False
    lines = [ln.strip() for ln in (page_text or "").splitlines() if ln.strip()]
    top = normalize_pdf_text(" ".join(lines[:12])).lower()
    if re.search(r"\bnotes?\s+to\s+(?:the\s+)?financial\s+statements?\b|\bmanagement['’]?s\s+discussion\b|\breconciliation\b|\bcombining\b|\bnonmajor\b|\bbudget", top, re.I):
        return False
    patterns = [
        r"\bstatements?\s+of\s+net\s+position\b",
        r"\bstatements?\s+of\s+activities\b",
        r"\bbalance\s+sheet\b",
        r"\bstatements?\s+of\s+revenues?\s*,?\s*expenditures?\s*,?\s*(?:and\s+)?changes?\s+in\s+fund\s+balances?\b",
        r"\bstatements?\s+of\s+(?:fund\s+)?net\s+position\b",
        r"\bstatements?\s+of\s+revenues?\s*,?\s*expenses?\s*,?\s*(?:and\s+)?changes?\s+in\s+(?:fund\s+)?net\s+position\b",
        r"\bstatements?\s+of\s+cash\s+flows?\b",
        r"\bstatements?\s+of\s+fiduciary\s+net\s+position\b",
        r"\bstatements?\s+of\s+changes?\s+in\s+fiduciary\s+net\s+position\b",
    ]
    return any(re.search(p, top, re.I) for p in patterns)


def state_v4_has_notes_or_note1_start(page_text: str) -> bool:
    """Return ``True`` if the page begins the Notes section or Note 1."""
    if not page_text:
        return False
    lines = [normalize_pdf_text(ln.strip()) for ln in page_text.splitlines() if ln.strip()]
    top_lines = lines[:15]
    top = normalize_pdf_text(" ".join(top_lines)).lower()
    if re.search(r"\bsee\s+notes?\s+to\s+financial\s+statements\b", top, re.I):
        return False
    for ln in top_lines:
        l = ln.lower().strip(" .:-")
        if re.match(r"^notes?\s+to\s+(?:the\s+)?financial\s+statements?$", l, re.I):
            return True
        if re.match(r"^note\s+1\b", l, re.I):
            return True
        if re.search(r"\bsummary\s+of\s+significant\s+accounting\s+polic(?:y|ies)\b|\bbasis\s+of\s+presentation\b|\breporting\s+entity\b", l, re.I):
            return True
    return False


def state_v4_find_post_statement_note_window(pdf, max_scan_pages: int = 260, max_pages_after_statement: int = 16) -> Optional[Tuple[int, int, int]]:
    """Find the safe Notes window after the basic financial statements; returns page-window info."""
    last_stmt = None
    notes_start = None
    limit = min(len(pdf.pages), max_scan_pages)
    for idx in range(limit):
        txt = cached_page_text(pdf.pages[idx]) or ""
        if not txt.strip():
            continue
        if last_stmt is not None and state_v4_is_statistical_supplementary_or_schedule_page(txt):
            break
        if state_v4_has_main_financial_statement_heading(txt):
            last_stmt = idx
            continue
        if last_stmt is not None and idx > last_stmt:
            if state_v4_has_notes_or_note1_start(txt):
                notes_start = idx
                break
            if idx - last_stmt > max_pages_after_statement:
                break
    if last_stmt is None:
        return None
    start = notes_start if notes_start is not None else last_stmt + 1
    end = min(len(pdf.pages) - 1, start + max_pages_after_statement)
    return (last_stmt, start, end) if start <= end else None


def state_v4_basis_header_hit(text: str) -> bool:
    """Return ``True`` if ``text`` contains a 'Basis of Presentation' style header."""
    t = normalize_pdf_text(text or "").lower()
    return bool(re.search(r"\bbasis\s+of\s+presentation\b", t, re.I) or re.search(r"\bbasis\s+of\s+accounting\s+and\s+presentation\b", t, re.I))


def state_v4_note1_reporting_entity_candidate(text: str) -> bool:
    """Return ``True`` if ``text`` is a Note 1 / Reporting Entity / Legal Identity candidate."""
    if not text or state_v4_is_statistical_supplementary_or_schedule_page(text):
        return False
    t = normalize_pdf_text(text).lower()
    top = normalize_pdf_text(" ".join(text.splitlines()[:40])).lower()
    return bool(
        re.search(r"\bnotes?\s+to\s+(?:the\s+)?financial\s+statements?\b|\bnote\s+1\b|\bsummary\s+of\s+significant\s+accounting\s+polic(?:y|ies)\b|\bbasis\s+of\s+presentation\b|\breporting\s+entity\b", top, re.I)
        or re.search(r"\bpolitical\s+subdivision\s+of\s+the\s+state\s+of\b|\bcounty\s+is\s+(?:a\s+)?(?:political\s+subdivision|organized)\b|\borganized\s+under\s+the\s+laws\s+of\b|\bbody\s+politic\s+and\s+corporate\b", t, re.I)
    )


def state_v4_find_issuer_and_state_in_tokens(tokens, word_map, issuer_variants, full_state_tokens, abbr_tokens):
    """Find issuer and state co-located in the token stream; returns match/box info or ``None``."""
    for issuer_tokens in issuer_variants:
        is_, ie = state_v4_find_token_window(tokens, issuer_tokens)
        if is_ is None:
            continue
        issuer_box = state_v4_box_from_words(word_map[is_:ie + 1])
        if not issuer_box:
            continue
        for state_tokens, label in [(full_state_tokens, "full_state"), (abbr_tokens, "state_abbr")]:
            if not state_tokens:
                continue
            ss, se = state_v4_find_token_window(tokens, state_tokens)
            if ss is not None:
                state_box = state_v4_box_from_words(word_map[ss:se + 1])
                if state_box:
                    return issuer_box, state_box, label
    return None


def state_v4_find_state_after_issuer_with_gap(tokens, word_map, issuer_variants, state_tokens, max_gap_words: int = 5):
    """Find the state appearing within ``max_gap_words`` after an issuer variant; returns box info or ``None``."""
    if not state_tokens:
        return None
    for issuer_tokens in issuer_variants:
        for i in range(0, len(tokens) - len(issuer_tokens) + 1):
            if tokens[i:i + len(issuer_tokens)] != issuer_tokens:
                continue
            issuer_end = i + len(issuer_tokens) - 1
            max_state_start = min(len(tokens) - len(state_tokens), issuer_end + max_gap_words + 1)
            for j in range(issuer_end + 1, max_state_start + 1):
                if tokens[j:j + len(state_tokens)] == state_tokens:
                    issuer_box = state_v4_box_from_words(word_map[i:issuer_end + 1])
                    state_box = state_v4_box_from_words(word_map[j:j + len(state_tokens)])
                    if issuer_box and state_box:
                        return issuer_box, state_box
    return None


def state_v4_split_sentences(text: str) -> List[str]:
    """Split ``text`` into sentences (State-v4 matcher)."""
    text = normalize_pdf_text(text or "")
    if not text:
        return []
    return [s.strip() for s in re.split(r"(?<=[.!?])\s+", text) if s.strip()]


def state_v4_same_sentence_issuer_state_hit(page_text, page_tokens, word_map, issuer_variants, full_state_tokens, abbr_tokens):
    """Detect issuer and state occurring in the same sentence; returns match/box info or ``None``."""
    for sent in state_v4_split_sentences(page_text):
        sent_tokens = state_v4_tokenize(sent)
        if not sent_tokens:
            continue

        state_tokens_used = None
        state_label = None
        if full_state_tokens and state_v4_find_token_window(sent_tokens, full_state_tokens)[0] is not None:
            state_tokens_used = full_state_tokens
            state_label = "full_state"
        elif abbr_tokens and state_v4_find_token_window(sent_tokens, abbr_tokens)[0] is not None:
            state_tokens_used = abbr_tokens
            state_label = "state_abbr"
        else:
            continue

        issuer_tokens_used = None
        for issuer_tokens in issuer_variants:
            if state_v4_find_token_window(sent_tokens, issuer_tokens)[0] is not None:
                issuer_tokens_used = issuer_tokens
                break
        if not issuer_tokens_used:
            continue

        sent_s, sent_e = state_v4_find_token_window(page_tokens, sent_tokens)
        if sent_s is not None:
            local_tokens = page_tokens[sent_s:sent_e + 1]
            local_map = word_map[sent_s:sent_e + 1]
            hit = state_v4_find_issuer_and_state_in_tokens(local_tokens, local_map, [issuer_tokens_used], state_tokens_used, [])
        else:
            hit = state_v4_find_issuer_and_state_in_tokens(page_tokens, word_map, [issuer_tokens_used], state_tokens_used, [])

        if hit:
            return hit[0], hit[1], state_label, sent
    return None


def state_v4_primary_state_same_page(pdf_path: Path, state_val: str, name_page: int, proof_path: Path) -> dict:
    """Primary State check: look for the state on the same page where the issuer name was found.

Returns a State-v4 result dict.
    """
    abbr, full_state, full_tokens, abbr_tokens = state_v4_state_target_tokens(state_val)
    with open_pdf_maybe_shared(str(pdf_path)) as pdf:
        if not name_page or name_page < 1 or name_page > len(pdf.pages):
            return state_v4_result_dict("FAIL", "PRIMARY_SAME_PAGE", "Name page unavailable")
        page = pdf.pages[name_page - 1]
        tokens, word_map = state_v4_page_tokens_with_boxes(page)
        for stoks, label in [(full_tokens, "full_state"), (abbr_tokens, "state_abbr")]:
            s, e = state_v4_find_token_window(tokens, stoks)
            if s is not None:
                box = state_v4_box_from_words(word_map[s:e + 1])
                if box:
                    state_v4_save_final_proof(page, [box], proof_path)
                return state_v4_result_dict("PASS", "PRIMARY_SAME_PAGE", f"State {label} found on issuer/name page", name_page, proof_path, [box] if box else [])
    return state_v4_result_dict("FAIL", "PRIMARY_SAME_PAGE", "State not found on issuer/name page", name_page)


def state_v4_fallback_auditor_report(pdf_path: Path, issuer_name: str, state_val: str, proof_path: Path, max_gap_words: int = 5) -> dict:
    """State fallback: search the main auditor report for issuer-followed-by-state within a word gap. Returns a result dict."""
    abbr, full_state, full_tokens, abbr_tokens = state_v4_state_target_tokens(state_val)
    issuer_vars = state_v4_issuer_token_variants(issuer_name)
    with open_pdf_maybe_shared(str(pdf_path)) as pdf:
        for idx, page in enumerate(pdf.pages[:120]):
            txt = cached_page_text(page) or ""
            if not txt.strip() or not state_v4_is_main_auditor_report_page(txt):
                continue

            tokens, word_map = state_v4_page_tokens_with_boxes(page)

            # Rule 1: issuer + max 5 words + state on auditor report page.
            for stoks, label in [(full_tokens, "full_state"), (abbr_tokens, "state_abbr")]:
                hit = state_v4_find_state_after_issuer_with_gap(tokens, word_map, issuer_vars, stoks, max_gap_words=max_gap_words)
                if hit:
                    issuer_box, state_box = hit
                    state_v4_save_final_proof(page, [issuer_box, state_box], proof_path)
                    return state_v4_result_dict(
                        "PASS", "FALLBACK_1_AUDITOR_REPORT",
                        f"State {label} found after issuer/issuer-variant on main auditor report page within {max_gap_words} words",
                        idx + 1, proof_path, [issuer_box, state_box]
                    )

            # Rule 2: issuer and state in same sentence on auditor report page.
            # This catches Broward County: "Broward County, Florida (the County)".
            hit = state_v4_same_sentence_issuer_state_hit(txt, tokens, word_map, issuer_vars, full_tokens, abbr_tokens)
            if hit:
                issuer_box, state_box, label, sentence = hit
                state_v4_save_final_proof(page, [issuer_box, state_box], proof_path)
                return state_v4_result_dict(
                    "PASS", "FALLBACK_1_AUDITOR_REPORT_SAME_SENTENCE",
                    f"Issuer and state found in the same sentence on main auditor report page ({label})",
                    idx + 1, proof_path, [issuer_box, state_box], {"matched_sentence": sentence}
                )

            return state_v4_result_dict("FAIL", "FALLBACK_1_AUDITOR_REPORT", "Main auditor report found, but state not found after issuer or in same sentence", idx + 1)
    return state_v4_result_dict("FAIL", "FALLBACK_1_AUDITOR_REPORT", "Main auditor report page not found")


def state_v4_fallback_basis_of_presentation(pdf_path: Path, issuer_name: str, state_val: str, proof_path: Path) -> dict:
    """State fallback: search the 'Basis of Presentation' note window for the state. Returns a result dict."""
    abbr, full_state, full_tokens, abbr_tokens = state_v4_state_target_tokens(state_val)
    issuer_vars = state_v4_issuer_token_variants(issuer_name)
    with open_pdf_maybe_shared(str(pdf_path)) as pdf:
        window = state_v4_find_post_statement_note_window(pdf)
        if not window:
            return state_v4_result_dict("FAIL", "FALLBACK_2_BASIS", "No safe Notes/Note 1 window immediately after financial statements")
        _, start_idx, end_idx = window
        for idx in range(start_idx, end_idx + 1):
            page = pdf.pages[idx]
            txt = cached_page_text(page) or ""
            if state_v4_is_statistical_supplementary_or_schedule_page(txt):
                break
            if not state_v4_basis_header_hit(txt):
                continue
            tokens, word_map = state_v4_page_tokens_with_boxes(page)
            hit = state_v4_find_issuer_and_state_in_tokens(tokens, word_map, issuer_vars, full_tokens, abbr_tokens)
            if hit:
                issuer_box, state_box, label = hit
                state_v4_save_final_proof(page, [issuer_box, state_box], proof_path)
                return state_v4_result_dict("PASS", "FALLBACK_2_BASIS", f"Issuer and state found in Basis/Note page immediately after financial statements ({label})", idx + 1, proof_path, [issuer_box, state_box])
    return state_v4_result_dict("FAIL", "FALLBACK_2_BASIS", "Issuer/state not found in Basis area immediately after financial statements")


def state_v4_fallback_note1_same_sentence(pdf_path: Path, issuer_name: str, state_val: str, proof_path: Path) -> dict:
    # Fallback 3 rule: issuer and state must be in the same sentence.
    """State fallback: search Note 1 / Reporting Entity for issuer and state in the same sentence. Returns a result dict."""
    abbr, full_state, full_tokens, abbr_tokens = state_v4_state_target_tokens(state_val)
    issuer_vars = state_v4_issuer_token_variants(issuer_name)
    with open_pdf_maybe_shared(str(pdf_path)) as pdf:
        window = state_v4_find_post_statement_note_window(pdf)
        if not window:
            return state_v4_result_dict("FAIL", "FALLBACK_3_NOTE1_SAME_SENTENCE", "No safe Notes/Note 1 window immediately after financial statements")
        _, start_idx, end_idx = window
        for idx in range(start_idx, end_idx + 1):
            page = pdf.pages[idx]
            txt = cached_page_text(page) or ""
            if state_v4_is_statistical_supplementary_or_schedule_page(txt):
                break
            if not state_v4_note1_reporting_entity_candidate(txt):
                continue
            tokens, word_map = state_v4_page_tokens_with_boxes(page)
            hit = state_v4_same_sentence_issuer_state_hit(txt, tokens, word_map, issuer_vars, full_tokens, abbr_tokens)
            if hit:
                issuer_box, state_box, label, sentence = hit
                state_v4_save_final_proof(page, [issuer_box, state_box], proof_path)
                return state_v4_result_dict(
                    "PASS", "FALLBACK_3_NOTE1_SAME_SENTENCE",
                    f"Issuer and state found in same sentence inside Note 1 / Reporting Entity / Legal Identity ({label})",
                    idx + 1, proof_path, [issuer_box, state_box], {"matched_sentence": sentence}
                )
    return state_v4_result_dict("FAIL", "FALLBACK_3_NOTE1_SAME_SENTENCE", "Issuer/state not found in same sentence inside safe Note 1 window")


def state_v4_run_integrated(pdf_path: Path, issuer_name: str, state_val: str, name_page: int, final_proof_path: Path) -> dict:
    """Run State v4 validation using the existing Name result page from the main workflow."""
    final_proof_path.parent.mkdir(parents=True, exist_ok=True)

    # Remove any prior State proof for this row first, so the STATE folder contains only the current final PASS image.
    try:
        if final_proof_path.exists():
            final_proof_path.unlink()
    except Exception:
        pass

    out = {
        "pdf_path": str(pdf_path),
        "issuer_name": issuer_name,
        "target_state": state_val,
        "steps": [],
        "final_status": "FAIL",
        "final_method": None,
        "final_proof_file": None,
    }

    if not pdf_path or not Path(pdf_path).exists():
        out["steps"].append(state_v4_result_dict("FAIL", "PDF", "PDF file not found"))
        return out

    checks = [state_v4_primary_state_same_page(Path(pdf_path), state_val, int(name_page or 0), final_proof_path)]
    if checks[-1].get("status") != "PASS":
        checks.append(state_v4_fallback_auditor_report(Path(pdf_path), issuer_name, state_val, final_proof_path))
    if checks[-1].get("status") != "PASS":
        checks.append(state_v4_fallback_basis_of_presentation(Path(pdf_path), issuer_name, state_val, final_proof_path))
    if checks[-1].get("status") != "PASS":
        checks.append(state_v4_fallback_note1_same_sentence(Path(pdf_path), issuer_name, state_val, final_proof_path))

    out["steps"].extend(checks)
    for step in out["steps"]:
        if step.get("status") == "PASS":
            out["final_status"] = "PASS"
            out["final_method"] = step.get("method")
            out["final_proof_file"] = str(final_proof_path) if final_proof_path.exists() else step.get("proof_file")
            break
    return out

def write_state_validation(
    processing_id: int,
    issuer_name: str,
    master_state_abbr: str,
    pdf_path: Optional[Path],
    name_result: dict,
    show_popup: bool = True
) -> dict:
    """
    STATE validation - Column E.

    Integrated State v4 logic only:
    1) Primary: state on the same page as the existing Name validation page.
    2) Fallback 1: main auditor report page:
       - issuer followed by state within 5 words, OR
       - issuer and state in the same sentence (Broward-style auditor report wording).
    3) Fallback 2: Basis of Presentation in a safe Notes window.
    4) Fallback 3: Note 1 / Reporting Entity / Legal Identity, with issuer and state in the same sentence.

    Proof behavior:
    - Saves ONLY the final PASS State proof image in PROOF_DIR_STATE.
    - Does not save intermediate State images.

    - Does NOT write to parquet (parent run_validations_for_row batches the write)
    """
    if not pdf_path or not Path(pdf_path).exists():
        return {
            "status": "FAIL",
            "page": None,
            "proof_file": None,
            "boxes": [],
            "title": f"ProcessingId {processing_id} | State Check = FAIL (PDF missing)"
        }

    name_page = (name_result or {}).get("page")
    if not name_page:
        return {
            "status": "FAIL",
            "page": None,
            "proof_file": None,
            "boxes": [],
            "title": f"ProcessingId {processing_id} | State Check = FAIL (No name page)"
        }

    proof_file = PROOF_DIR_STATE / f"Row{processing_id}_STATE_{sanitize_filename(master_state_abbr, 10)}.png"

    state_result = state_v4_run_integrated(
        pdf_path=Path(pdf_path),
        issuer_name=issuer_name,
        state_val=master_state_abbr,
        name_page=int(name_page),
        final_proof_path=proof_file
    )

    status = state_result.get("final_status", "FAIL")
    final_step = None
    for step in state_result.get("steps", []):
        if step.get("status") == "PASS":
            final_step = step
            break
    if final_step is None:
        final_step = state_result.get("steps", [{}])[-1] if state_result.get("steps") else {}

    page_no = final_step.get("page", name_page)
    boxes = final_step.get("boxes", []) or []
    proof = state_result.get("final_proof_file") or final_step.get("proof_file")
    method = state_result.get("final_method") or final_step.get("method")
    reason = final_step.get("reason", "")

    title = f"ProcessingId {processing_id} | State Check = {status} | Page {page_no}"
    if method:
        title += f" | {method}"

    return {
        "status": status,
        "page": page_no,
        "proof_file": Path(proof) if proof else None,
        "boxes": boxes,
        "title": title,
        "method": method,
        "reason": reason,
        "steps": state_result.get("steps", []),
    }

def clear_directory_contents(dir_path: Path) -> None:
    """Delete all files/subfolders inside dir_path (but keep the folder)."""
    dir_path.mkdir(parents=True, exist_ok=True)
    for item in dir_path.iterdir():
        try:
            if item.is_file() or item.is_symlink():
                item.unlink(missing_ok=True)
            elif item.is_dir():
                shutil.rmtree(item, ignore_errors=True)
        except Exception:
            pass



def move_pdf(src_path: Path, dest_dir: Path) -> Tuple[Optional[Path], Optional[str]]:
    """
    Move a PDF to dest_dir, keeping the same filename.

    Returns:
      (new_path, None)        on success
      (None, error_message)   on failure (file stays at src)
    """
    try:
        src_path = Path(src_path)
        if not src_path.exists():
            return (None, f"Source PDF not found: {src_path}")

        dest_dir.mkdir(parents=True, exist_ok=True)
        target = dest_dir / src_path.name

        # If a file with the same name already exists at the target, remove it (latest wins)
        if target.exists():
            try:
                target.unlink()
            except Exception:
                pass

        # Same-drive rename is fastest; shutil.move handles cross-drive too
        try:
            src_path.replace(target)
        except Exception:
            shutil.move(str(src_path), str(target))

        return (target, None)

    except Exception as e:
        return (None, f"Move failed: {e}")



# =========================================================
# FINAL VALIDATION LOGIC - DAILY SOURCING COLUMN H ONLY
# =========================================================

def final_normalize(val) -> str:
    """Return ``val`` as a trimmed lowercase string for final pass/fail comparison."""
    if val is None:
        return ""
    return str(val).strip().lower()


def final_is_pass(val) -> bool:
    """Treat PASS / PASSED / PASS - something as pass."""
    return final_normalize(val).startswith("pass")


# row wide check
def final_compute_status_from_results(results: dict) -> Tuple[bool, str]:
    """
    Final validation logic using the in-memory results dict from run_validations_for_row.
    Identical 5-step logic; reads by column name instead of Excel cells.
    """
    def v(col_name):
        return (results or {}).get(col_name)

    # Step 1: Issuer Detail Validation
    issuer_ok = (
        final_is_pass(v("Name"))
        and final_is_pass(v("FYE (mm/yy)"))
        and final_is_pass(v("State"))
    )
    if not issuer_ok:
        return False, "FAILED at Step 1: Issuer Detail Validation"

    # Step 2: Audit Validation
    audit_ok = (
        final_is_pass(v("Availability of Audit Report"))
        and final_is_pass(v("Audit Opinion"))
        and final_is_pass(v("Audit Report FYE with Report FYE"))
        and final_is_pass(v("Auditor's Signature"))
    )
    if not audit_ok:
        return False, "FAILED at Step 2: Audit Validation"

    # Step 3: Statement Validation
    snp           = final_is_pass(v("Statement of Net Position with FYE"))
    act           = final_is_pass(v("Statement of Activities with FYE"))
    bs            = final_is_pass(v("Balance Sheet with FYE"))
    rev_fund      = final_is_pass(v("Statement of Revenues, Expenditures and changes in Fund Balances with FYE"))
    prop_snp      = final_is_pass(v("Statement of Net Position of Proprietary Funds with FYE"))
    rev_net       = final_is_pass(v("Statements of Revenues, Expenses And Changes In Net Position with FYE"))
    cash_receipts = final_is_pass(v("Statement of Cash Receipts and Disbursements with FYE"))

    logic_3a = snp and act and bs and rev_fund and prop_snp and rev_net
    logic_3b = snp and act and bs and rev_fund
    logic_3c = bs and rev_fund
    logic_3d = (prop_snp or snp) and rev_net
    logic_3e = cash_receipts

    if logic_3a:
        return True, "PASSED by Step 3a"
    if logic_3b:
        return True, "PASSED by Step 3b"
    if logic_3c:
        return True, "PASSED by Step 3c"
    if logic_3d:
        return True, "PASSED by Step 3d"
    if logic_3e:
        return True, "PASSED by Step 3e"

    return False, "FAILED at Step 3: Statement Validation"



# =========================================================
# ENTRYPOINT
# =========================================================
# OUTPUT-FORMATTING: Main workflow: search FAC for each row, download Excel/PDF, validate PDFs, and update workbook.

def apply_fye_fallback_using_audit_report(processing_id: int, fye_res: dict, audit_fye_res: dict) -> dict:
    """
    If FYE(mm/yy) failed but Audit Report FYE passed, force FYE(mm/yy) to PASS for this row.
    - Does NOT write to parquet (parent run_validations_for_row re-captures fye_res["status"])
    """
    try:
        fye_status = (fye_res or {}).get("status", "")
        audit_fye_status = (audit_fye_res or {}).get("status", "")

        if str(fye_status).upper() != "PASS" and str(audit_fye_status).upper() == "PASS":
            # Update the returned dict so the parent + popups/bundles show PASS consistently
            fye_res = dict(fye_res or {})
            fye_res["status"] = "PASS"
            reason = fye_res.get("reason", "")
            suffix = " (Auto-PASS because Audit Report FYE matched)"
            if suffix not in str(reason):
                fye_res["reason"] = (str(reason) + suffix).strip()
            title = fye_res.get("title")
            if title and "PASS" not in str(title):
                fye_res["title"] = str(title).replace("FAIL", "PASS") + suffix

    except Exception:
        # Never break the run due to fallback logic
        pass

    return fye_res


def run_validations_for_row(db, row_id: int, processing_id: int, issuer_name: str, master_state_abbr: str, uei: str, fy_end_val, pdf_saved: Path):
    """Run all validations for a row in one go (one PDF open + cached page extraction).

    row_id        -> physical TProcessStatus.Id (unique) — parquet report key.
    processing_id -> deliverable grouping (non-unique) — display + AdditionalInfo lookup.
    """
    global _SHARED_PDF_PATH, _SHARED_PDF_HANDLE

    # Colors for proof bundles (moved in from old signature params)
    NAME_COLOR  = (0, 120, 255, 90)   # Blue
    STATE_COLOR = (0, 200, 0, 90)     # Green
    FYE_COLOR   = (255, 0, 255, 90)   # Magenta

    # Accumulator for the whole row → single parquet write at the end
    results = {
        "ISSUER NAME": issuer_name,
        "UEI": uei,
    }

    clear_pdf_page_caches()
    _SHARED_PDF_PATH = str(pdf_saved)

    try:
        with pdfplumber.open(str(pdf_saved)) as _pdf_shared:
            _SHARED_PDF_HANDLE = _pdf_shared

            # NAME + STATE + FYE
            name_res = write_name_validation(processing_id, issuer_name, uei, pdf_saved, show_popup=False)#done
            results["Name"] = name_res["status"]

            state_res = write_state_validation(processing_id, issuer_name, master_state_abbr, pdf_saved, name_res, show_popup=False)
            results["State"] = state_res["status"]

            fye_res = write_fye_validation(processing_id, fy_end_val, pdf_saved, name_res, show_popup=False)
            results["FYE (mm/yy)"] = fye_res["status"]

            show_threeway_bundle_or_individual(
                processing_id=processing_id,
                pdf_path=pdf_saved,
                res1=name_res,
                res2=state_res,
                res3=fye_res,
                out_dir=PROOF_DIR_NAME,
                bundle_tag="NAME_STATE_FYE",
                color1=NAME_COLOR,
                color2=STATE_COLOR,
                color3=FYE_COLOR,
                popup_title=f"ProcessingId {processing_id} | NAME/STATE/FYE",
                show_popup=False
            )

            # AUDIT + OPINION + AUDIT FYE --done
            audit_res = write_audit_report_validation(processing_id, issuer_name, fy_end_val, pdf_saved, show_popup=False)
            results["Availability of Audit Report"] = audit_res["status"]

            opinion_res = write_audit_opinion_validation(processing_id, pdf_saved, audit_res=audit_res, show_popup=False)
            results["Audit Opinion"] = opinion_res["status"]
            
            # all the below methods are Unreviewed 
            audit_fye_res = write_audit_report_fye_validation(processing_id, fy_end_val, pdf_saved, audit_res=audit_res, show_popup=False)
            results["Audit Report FYE with Report FYE"] = audit_fye_res["status"]
            
            # FYE fallback: if the cover page is an image and FYE(mm/yy) fails, accept PASS when Audit Report FYE matche   
            fye_res = apply_fye_fallback_using_audit_report(processing_id, fye_res, audit_fye_res)
            results["FYE (mm/yy)"] = fye_res["status"]   # may have been upgraded to PASS

            show_single_or_individual_popups(processing_id, pdf_saved, audit_res, opinion_res, audit_fye_res)

            # SIGNATURE — auditor firm name now comes from DB
            # TProcessingAdditionalInfo stays keyed on ProcessingId (not part of the Id re-key).
            firm_res = db.fetch_one("""
                SELECT AuditorFirmName
                FROM TProcessingAdditionalInfo
                WHERE ProcessingId = ?
            """, processing_id)
            row = firm_res.data if firm_res.success else None
            auditor_firm_name = normalize(str(row.AuditorFirmName or "")) if row else ""

            sig_res = write_auditor_signature_validation(processing_id, pdf_saved, audit_res=audit_res, auditor_firm_name=auditor_firm_name)
            results["Auditor's Signature"] = sig_res["status"]


            # FINANCIAL STATEMENTS
            net_pos_res = write_statement_of_net_position_validation(processing_id, fy_end_val, pdf_saved, show_popup=False)
            results["Statement of Net Position with FYE"] = net_pos_res["status"]

            activities_res = write_statement_of_activities_validation(processing_id, fy_end_val, pdf_saved, show_popup=False)
            results["Statement of Activities with FYE"] = activities_res["status"]

            balance_res = write_balance_sheet_validation(processing_id, fy_end_val, pdf_saved, show_popup=False)
            results["Balance Sheet with FYE"] = balance_res["status"]

            rev_exp_res = write_rev_exp_changes_fund_balances_validation(processing_id, fy_end_val, pdf_saved, show_popup=False)
            results["Statement of Revenues, Expenditures and changes in Fund Balances with FYE"] = rev_exp_res["status"]

            net_pos_prop_res = write_statement_of_net_position_proprietary_funds_validation(processing_id, fy_end_val, pdf_saved, show_popup=False)
            results["Statement of Net Position of Proprietary Funds with FYE"] = net_pos_prop_res["status"]

            rev_exp_net_res = write_rev_exp_changes_net_position_validation(processing_id, fy_end_val, pdf_saved, show_popup=False)
            results["Statements of Revenues, Expenses And Changes In Net Position with FYE"] = rev_exp_net_res["status"]

            cash_res = write_statement_of_cash_receipts_disbursements_validation(processing_id, fy_end_val, pdf_saved, show_popup=False)
            results["Statement of Cash Receipts and Disbursements with FYE"] = cash_res["status"]
    finally:
        _SHARED_PDF_HANDLE = None
        _SHARED_PDF_PATH = None
        clear_pdf_page_caches()
        # ---- SINGLE parquet write for the whole row (even on partial failure) ----
        write_validation_row(row_id, processing_id, results)
    return results


# =========================================================
# DOWNLOAD FAILURE HANDLING (Excel download failed after all retries)
# =========================================================

FAILED_TO_DOWNLOAD_TEXT = "Failed to Download"


# =========================================================
# DB Status & Flag update
# =========================================================
def update_sourcing_validation_status(db, row_id, status, flag, pdf_path=None, remarks=None):
    """
    status: None=pending, 0=fail, 1=pass
    flag:   None=not s=started, p=progress, c=complete
    pdf_path: if provided, updates PdfFilePath
    remarks:  if provided, updates Remarks

    Keyed on the physical TProcessStatus.Id (row_id).
    """
    sets = ["SourcingValidationStatus = ?", "SourcingValidationFlag = ?", "ModifiedOn = GETDATE()"]
    params = [status, flag]

    if pdf_path is not None:
        sets.append("PdfFilePath = ?")
        params.append(str(pdf_path))

    if remarks is not None:
        sets.append("Remarks = ?")
        params.append(remarks)

    params.append(row_id)

    res = db.update(f"""
        UPDATE TProcessStatus
        SET {", ".join(sets)}
        WHERE Id = ?
    """, params)
    if not res.success:
        print(f"[FAILED] row_id {row_id}: {res.error}")

def main_validate():
    """Run ONLY PDF validations using already-downloaded PDFs. Does NOT download anything."""


    # --------------------------------------------------
    # 1. CONNECT TO DATABASE (replaces load_workbook)
    #    Uses the shared db.py layer as a context manager: commits on clean
    #    exit, rolls back on exception, and closes the connection.
    # --------------------------------------------------
    with Database() as db:
        print("[DB] Connected to database successfully.")

        # --------------------------------------------------
        # 2. FETCH ROWS — ADD FyeDate + PdfFilePath to SELECT
        #    ps.Id AS RowId → the physical-row PK (unique) used for writes
        # --------------------------------------------------
        res = db.fetch_all("""
            SELECT
                cm.CompanyId,
                ps.Id AS RowId,
                ps.ProcessingId,
                ps.ProcessingCode,
                cm.IssuerName,
                cm.State,
                cm.Sector,
                cm.SubSector,
                cm.UEI,
                cm.EIN,
                ps.ProcessYear,
                ps.FyeDate,
                ps.PdfFilePath
            FROM TCompanyMaster cm
            JOIN TProcessStatus ps
                ON ps.CompanyId = cm.CompanyId
            WHERE cm.IsActive  = 1
              AND ps.IsActive   = 1
              AND ps.SourcingStatus = 1
              AND ps.SourcingFlag = 'c'
              AND (ps.SourcingValidationFlag = 'c' OR ps.SourcingValidationFlag IS NULL)
              AND (ps.SourcingValidationStatus = 0 OR ps.SourcingValidationStatus IS NULL)
              AND (ps.CompletionStatus = 0      OR ps.CompletionStatus IS NULL)
              AND ps.COAID IN (1, 2)
        """)
        if not res.success:
            print(f"[DB-ERROR] Failed to fetch rows pending validation: {res.error}")
            return
        master_rows = res.data
        print(f"[DB] Found {len(master_rows)} rows pending validation.")

        if not master_rows:
            print("[INFO] No rows pending validation. Exiting.")
            return

        # Claim the fetched physical rows by their PK (Id), not ProcessingId —
        # ProcessingId is non-unique so an IN(ProcessingId) list would also flip
        # sibling rows that share a deliverable.
        row_ids = [row.RowId for row in master_rows]
        clause, params = build_in_clause("Id", row_ids)
        claim = db.update(f"""
            UPDATE TProcessStatus
            SET SourcingValidationFlag = 's'
            WHERE IsActive = 1
              AND {clause}
        """, params)
        print(f"[DB] Updated SourcingValidationFlag to 's' for {claim.rowcount} rows.")

        # Clear proof folders only
        clear_directory_contents(PROOF_DIR_NAME)
        clear_directory_contents(PROOF_DIR_STATE)
        clear_directory_contents(PROOF_DIR_FYE)
        clear_directory_contents(PROOF_DIR_AUDIT)
        clear_directory_contents(PROOF_DIR_OPINION)
        clear_directory_contents(PROOF_DIR_AUDIT_FYE)
        clear_directory_contents(PROOF_DIR_SIGNATURE)
        clear_directory_contents(PROOF_DIR_NET_POSITION)
        clear_directory_contents(PROOF_DIR_ACTIVITIES)
        clear_directory_contents(PROOF_DIR_BALANCE_SHEET)
        clear_directory_contents(PROOF_DIR_REV_EXP_FUND_BAL)

        for db_row in master_rows:
            row_id          = db_row.RowId          # physical TProcessStatus.Id — DB row key
            processing_id   = db_row.ProcessingId   # deliverable grouping (non-unique) — display/label
            processing_code = db_row.ProcessingCode
            uei             = normalize(str(db_row.UEI         or ""))
            ein             = normalize(str(db_row.EIN         or ""))
            issuer_name_val = normalize(str(db_row.IssuerName  or ""))
            year            = normalize(str(db_row.ProcessYear or ""))
            state           = normalize(str(db_row.State       or ""))
            fy_end_val      = db_row.FyeDate        # date/None — same type as before
            pdf_file_path   = db_row.PdfFilePath    # stored path from sourcing run

            update_sourcing_validation_status(db, row_id, None, 'p')

            print(f"\n--- Validating ProcessingId: {processing_id} (row Id {row_id}) | {issuer_name_val} ---")

            # --------------------------------------------------
            # 4a. LOCATE PDF
            # --------------------------------------------------
            pdf_path = Path(pdf_file_path)
            # pdf_path = _find_matching_pdf(PDF_DIR, year, uei, ein)

            # --------------------------------------------------
            # 4b. PDF NOT FOUND → mark all validations failed
            # --------------------------------------------------
            if not pdf_path or not pdf_path.exists():
                write_validation_row(row_id, processing_id, {
                    "ISSUER NAME": issuer_name_val,
                    "UEI":         uei,
                    "Name":                                                          "Fail",
                    "FYE (mm/yy)":                                                   "Fail",
                    "State":                                                         "Fail",
                    "Availability of Audit Report":                                  "Fail",
                    "Audit Opinion":                                                 "Fail",
                    "Audit Report FYE with Report FYE":                              "Fail",
                    "Auditor's Signature":                                           "Fail",
                    "Statement of Net Position with FYE":                            "Fail",
                    "Statement of Activities with FYE":                              "Fail",
                    "Balance Sheet with FYE":                                        "Fail",
                    "Statement of Revenues, Expenditures and changes in Fund Balances with FYE": "Fail",
                    "Statement of Net Position of Proprietary Funds with FYE":       "Fail",
                    "Statements of Revenues, Expenses And Changes In Net Position with FYE":     "Fail",
                    "Statement of Cash Receipts and Disbursements with FYE":         "Fail",
                })
                update_sourcing_validation_status(
                    db, row_id, status=0, flag='c',
                    remarks="PDF not found for validation"
                )
                print(f"[WARN] ProcessingId {processing_id} (row Id {row_id}): PDF not found for validation.")
                continue


            # --------------------------------------------------
            # RUN ALL VALIDATIONS (returns results dict)
            # --------------------------------------------------
            results = run_validations_for_row(
                db=db,
                row_id=row_id,
                processing_id=processing_id,
                issuer_name=issuer_name_val,
                master_state_abbr=state,
                uei=uei,
                fy_end_val=fy_end_val,
                pdf_saved=pdf_path
            )

            print(f"[INFO] ProcessingId {processing_id} (row Id {row_id}): validations completed.")

            # --------------------------------------------------
            # COMPUTE VERDICT + MOVE PDF + UPDATE DB
            # --------------------------------------------------
            final_pass, reason = final_compute_status_from_results(results)

            dest_dir = VALIDATED_PDF_DIR if final_pass else FAILED_VALIDATION_PDF_DIR
            new_path, move_err = move_pdf(pdf_path, dest_dir)

            verdict_status = 1 if final_pass else 0

            if move_err:
                # Validation verdict stands; only the move failed → keep original path, note remarks
                update_sourcing_validation_status(
                    db, row_id,
                    status=verdict_status, flag='c',
                    remarks=f"{reason} | PDF move failed: {move_err}"
                )
                print(f"[WARN] ProcessingId {processing_id} (row Id {row_id}): {reason} | move failed: {move_err}")
            else:
                # Move succeeded → update PdfFilePath to new location
                update_sourcing_validation_status(
                    db, row_id,
                    status=verdict_status, flag='c',
                    pdf_path=new_path
                )
                print(f"[INFO] ProcessingId {processing_id} (row Id {row_id}): {reason} | moved -> {new_path}")

    # --------------------------------------------------
    # FINALIZE
    #    The `with Database()` block above committed and closed the connection.
    # --------------------------------------------------
    print("[DONE] Validation completed.")


if __name__ == "__main__":
    main_validate()