# -*- coding: utf-8 -*-
"""
STEP 1:
- Load XLSM workbook, read headers, create maps.

STEP 2:
- Loop MASTER LIST rows:
  - Open FAC advanced search
  - Apply filters (Audit Year, UEI, State, Entity type)
  - Search
  - Download "Download all" Excel

STEP 3:
- Open downloaded Excel -> General sheet
- Identify "reliant row" (scoring match)
- Extract and write to MASTER LIST columns

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
import time
import shutil
import subprocess
import pyodbc
import pandas as pd
import threading
from filelock import FileLock
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


from openpyxl import load_workbook
# from openpyxl.styles import PatternFill, Border, Side
from playwright.sync_api import (
    sync_playwright,
    TimeoutError as PlaywrightTimeoutError,
    Error as PlaywrightError,
)



# =========================================================
# DB CONNECTION
# =========================================================
# Connection handling now lives in db.py (get_db_connection() + Database).
# Import `Database` at the top of this file and use it as a context manager:
#     with Database() as db:
#         res = db.fetch_all(sql, params)
#         db.update(sql, params)
# The old per-file get_db_connection() has been removed in favour of that layer.


# =========================================================
# PARQUET LOG FILE PATH (global variable)
# =========================================================

LOG_PARQUET_PATH = Path(r"C:\S2\Public Finance\999_Log_Trackers\processing_log.parquet")
LOG_LOCK_PATH    = LOG_PARQUET_PATH.with_suffix(".parquet.lock")   # sibling lock file
 
# Stage codes (as per the schema)
STAGE_SOURCING             = "s"
STAGE_PROCESSING           = "p"
STAGE_SOURCING_VALIDATION  = "sv"
STAGE_PROCESSING_VALIDATION= "pv"

# --- Concurrency primitives ---
# Thread lock: protects multiple threads in the SAME Python process.
# File lock : protects multiple PROCESSES (e.g., 5 separate script runs).
# We need BOTH because filelock alone is too slow for in-process threads,
# and threading.Lock alone can't see other processes.
_LOG_THREAD_LOCK = threading.Lock()
_LOG_FILE_LOCK   = FileLock(str(LOG_LOCK_PATH), timeout=60)
 


# =========================================================
# PARQUET LOG HELPERS
# =========================================================

def _ensure_log_file():
    """
    Check if the parquet log file exists.
    If not, create it with the correct empty schema.

    Columns:
      ProcessingLogId  → processing_code  (string)
      Id               → row_id           (int)  physical TProcessStatus.Id (unique per row)
      ProcessingId     → processing_id    (int)  deliverable grouping key (non-unique)
      Stage            → s / p / sv / pv  (string)
      Remarks          → log message text (string)
      Time             → datetime of insert (datetime)
    """
    LOG_PARQUET_PATH.parent.mkdir(parents=True, exist_ok=True)

    if not LOG_PARQUET_PATH.exists():
        empty_df = pd.DataFrame(columns=[
            "ProcessingLogId",   # processing_code
            "Id",                # row_id (physical TProcessStatus.Id)
            "ProcessingId",      # processing_id (deliverable grouping)
            "Stage",             # s / p / sv / pv
            "Remarks",           # log message text
            "Time",              # datetime
        ])
        empty_df = empty_df.astype({
            "ProcessingLogId": "object",
            "Id":              "Int64",    # nullable int
            "ProcessingId":    "Int64",    # nullable int
            "Stage":           "object",
            "Remarks":         "object",
            "Time":            "datetime64[ns]",
        })
        empty_df.to_parquet(LOG_PARQUET_PATH, index=False)
        # print(f"[LOG] Created new parquet log file at: {LOG_PARQUET_PATH}")


def write_log(processing_code: str, row_id: int, processing_id: int, stage: str, remarks: str):
    """
    Append one row to the parquet log file.
    Thread-safe AND process-safe.

    row_id       = physical TProcessStatus.Id (unique per row) — the row identity.
    processing_id = deliverable grouping key (non-unique) — kept for readability.
 
    Concurrency strategy:
      1. Acquire in-process THREAD lock first  (fast - blocks sibling threads)
      2. Then acquire cross-process FILE lock  (slow - blocks other script runs)
      3. Read -> concat -> write inside both locks
      4. Release both locks
 
    With 5 parallel threads:
      - Only ONE thread at a time will read+write the parquet
      - Other 4 threads wait until the active thread releases the lock
      - No lost rows, no corrupted files
    """
    new_row = pd.DataFrame([{
        "ProcessingLogId": str(processing_code or ""),
        "Id":              int(row_id) if row_id else pd.NA,
        "ProcessingId":    int(processing_id) if processing_id else pd.NA,
        "Stage":           str(stage or ""),
        "Remarks":         str(remarks or ""),
        "Time":            datetime.now(),
    }])
 
    # ---- THREAD LOCK (fastest layer) ----
    with _LOG_THREAD_LOCK:
        # ---- FILE LOCK (cross-process layer) ----
        try:
            with _LOG_FILE_LOCK:
                _ensure_log_file()
 
                try:
                    existing_df = pd.read_parquet(LOG_PARQUET_PATH)
                    updated_df  = pd.concat([existing_df, new_row], ignore_index=True)
                except Exception:
                    # Corrupted/empty file -> start fresh with this single row
                    updated_df = new_row
 
                updated_df.to_parquet(LOG_PARQUET_PATH, index=False)
 
        except Exception as e:
            # Logging must NEVER crash the main workflow.
            # Print to console as a last-resort fallback.
            print(f"[LOG-ERROR] Failed to write log row: {e} | Remarks: {remarks}")

# =========================================================
# PERFORMANCE HELPERS (reuse PDF handle + cache page text/words per row)
# =========================================================
from contextlib import contextmanager

_SHARED_PDF_PATH: Optional[str] = None
_SHARED_PDF_HANDLE = None

_PAGE_TEXT_CACHE: Dict[int, str] = {}
_PAGE_WORDS_CACHE: Dict[Tuple[int, bool], list] = {}


# No Changes
def cached_page_text(page) -> str:
    """Return page text, cached per page object (avoids re-extracting many times)."""
    k = id(page)
    if k not in _PAGE_TEXT_CACHE:
        _PAGE_TEXT_CACHE[k] = page.extract_text() or ""
    return _PAGE_TEXT_CACHE[k]

# No Changes
def cached_page_words(page, use_text_flow: bool = True):
    """Return page words, cached per page object (avoids re-extracting many times)."""
    k = (id(page), bool(use_text_flow))
    if k not in _PAGE_WORDS_CACHE:
        _PAGE_WORDS_CACHE[k] = page.extract_words(use_text_flow=use_text_flow) or []
    return _PAGE_WORDS_CACHE[k]

# No Changes
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

# =========================================================
# CONFIG
# =========================================================
# FILE_PATH = r"C:\S2\Public Finance\PDF to be Sourced and Extracted.xlsm"

SHEET_MASTER = "MASTER LIST"
SHEET_VALIDATION = "Detailed Validation Check"
PDF_NOT_FOUND_TEXT = "PDF Not Found"

BASE = "https://app.fac.gov"
ADV_URL = "https://app.fac.gov/dissemination/search/advanced/"

DOWNLOAD_DIR = Path(r"C:\Test Code\Public Finance\01_Metadata_update")
DOWNLOAD_DIR.mkdir(parents=True, exist_ok=True)

PDF_DIR = Path(r"C:\Test Code\Public Finance\02_Downloaded_Report")
PDF_DIR.mkdir(parents=True, exist_ok=True)



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

# Always show a snippet popup for every row (PASS/FAIL/NOT EXTRACTABLE)
SHOW_PROOF_POPUP = True
# Popup auto closes after N seconds (set 0 to disable popup)
PROOF_POPUP_SECONDS = 1

# If popup is not visible (or Tk is unavailable), reveal/open the saved proof file
REVEAL_PROOF_FILES = True

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

# Retry settings (4 more retries => total 5 attempts)
DOWNLOAD_RETRIES = 4
RETRY_WAIT_SECONDS = 2
HEADLESS = True          # set True for final run
SLOW_MO_MS = 150          # set 0 for faster run
PAGE_LOAD_TIMEOUT_MS = 60000
UI_TIMEOUT_MS = 30000
DOWNLOAD_TIMEOUT_MS = 90000
PDF_DOWNLOAD_TIMEOUT_MS = 90000

CHROME_CHANNEL = "chrome"
IGNORE_HTTPS_ERRORS = True  # helps in SSL-inspected networks


# =========================================================
# STEP 1 HELPERS
# =========================================================


# =========================================================
# SOURCING (FAC SEARCH + DOWNLOADS + EXCEL UPDATE)
# =========================================================

# SOURCING: Find the last non-empty cell index in a list (used to detect real header width).
# No Changes
def _last_non_empty_index(values):
    last = 0
    for i, v in enumerate(values, start=1):
        if v is not None and str(v).strip() != "":
            last = i
    return last



# SOURCING: Convert an Excel column header into a safe Python-style key (lowercase with underscores).
# No Changes
def make_python_key(header: str) -> str:
    s = header.strip().lower()
    s = re.sub(r"[^\w]+", "_", s, flags=re.UNICODE)
    s = re.sub(r"_+", "_", s).strip("_")
    return s or "col"


# SOURCING: Turn any cell value into a clean trimmed string (blank-safe).
# No Changes
def normalize(v) -> str:
    return "" if v is None else str(v).strip()


# SOURCING: Clean text so it can be safely used as a Windows file name.
# No Changes
def sanitize_filename(s: str, max_len: int = 140) -> str:
    s = (s or "").strip()
    s = re.sub(r'[\\/:"*?<>|]+', "_", s)
    s = re.sub(r"\s+", " ", s).strip(" ._")
    if not s:
        s = "UNKNOWN"
    if len(s) > max_len:
        s = s[:max_len].rstrip(" ._")
    return s


# SOURCING: Normalize an Excel header cell so smart quotes and extra spaces do not break matching.
# No Changes
def normalize_header_text(v) -> str:
    """Normalize header strings so Excel smart-quotes / NBSP / spacing don't break matching."""
    if v is None:
        return ""
    s = str(v).strip()

    # NBSP -> normal space
    s = s.replace("\u00A0", " ")

    # Smart quotes -> straight quotes
    s = (s.replace("’", "'")
           .replace("‘", "'")
           .replace("“", '"')
           .replace("”", '"'))

    # Collapse whitespace
    s = re.sub(r"\s+", " ", s).strip()
    return s


# SOURCING: Create a comparable key from a header name to match columns reliably.
# No Changes
def header_key(v) -> str:
    """Convert a header to a comparable python-key style token."""
    return make_python_key(normalize_header_text(v))


# SOURCING: Build a dictionary of header-key to column-number for fast column lookup.
# No Changes
def build_header_index_map(ws, header_row: int = 1) -> Dict[str, int]:
    """
    Map normalized header_key -> column index.
    This makes 'FY End Date' match 'fy_end_date' automatically.
    """
    m = {}
    for idx, cell in enumerate(ws[header_row], start=1):
        k = header_key(cell.value)
        if k:
            # if duplicates exist, keep first occurrence
            m.setdefault(k, idx)
    return m


# SOURCING: Find a column number in a sheet by trying multiple possible header names.
# No Changes
def find_col(ws, candidates: List[str], header_row: int = 1) -> Optional[int]:
    """
    Find a column using normalized header_key matching.
    Includes a fallback contains-match for extra-safe detection.
    """
    hmap = build_header_index_map(ws, header_row=header_row)

    # exact normalized-key match
    for c in candidates:
        ck = header_key(c)
        if ck in hmap:
            return hmap[ck]

    # fallback: loose/contains match (handles "FY End Date (mm/dd/yyyy)" etc.)
    cand_keys = [header_key(c) for c in candidates if header_key(c)]
    for hk, col in hmap.items():
        for ck in cand_keys:
            if ck and (ck in hk or hk in ck):
                return col

    return None

# SOURCING: Read a single Excel cell value.
# No Changes
def get_cell(ws, row: int, col: int):
    return ws.cell(row=row, column=col).value


# SOURCING: Write a value into a single Excel cell.
# No Changes
def set_cell(ws, row: int, col: int, value):
    ws.cell(row=row, column=col).value = value


# SOURCING: Locate a worksheet name even if capitalization is different.
# No Changes (Works on downloaded excel not the metadata excel)
def find_sheet_case_insensitive(wb, desired: str) -> Optional[str]:
    d = desired.strip().lower()
    for s in wb.sheetnames:
        if s.strip().lower() == d:
            return s
    return None


# =========================================================
# FAC (STATE + SECTOR) MAPPINGS
# =========================================================
STATE_NAME_TO_ABBR = {
    "alabama": "AL", "alaska": "AK", "arizona": "AZ", "arkansas": "AR", "california": "CA",
    "colorado": "CO", "connecticut": "CT", "delaware": "DE", "florida": "FL", "georgia": "GA",
    "hawaii": "HI", "idaho": "ID", "illinois": "IL", "indiana": "IN", "iowa": "IA", "kansas": "KS",
    "kentucky": "KY", "louisiana": "LA", "maine": "ME", "maryland": "MD", "massachusetts": "MA",
    "michigan": "MI", "minnesota": "MN", "mississippi": "MS", "missouri": "MO", "montana": "MT",
    "nebraska": "NE", "nevada": "NV", "new hampshire": "NH", "new jersey": "NJ", "new mexico": "NM",
    "new york": "NY", "north carolina": "NC", "north dakota": "ND", "ohio": "OH", "oklahoma": "OK",
    "oregon": "OR", "pennsylvania": "PA", "rhode island": "RI", "south carolina": "SC",
    "south dakota": "SD", "tennessee": "TN", "texas": "TX", "utah": "UT", "vermont": "VT",
    "virginia": "VA", "washington": "WA", "west virginia": "WV", "wisconsin": "WI", "wyoming": "WY",
    "district of columbia": "DC", "dc": "DC",
    "american samoa": "AS", "guam": "GU", "northern mariana islands": "MP", "puerto rico": "PR",
    "virgin islands": "VI", "palau": "PW", "micronesia": "FM", "marshall islands": "MH",
}

SECTOR_TO_LABEL = {
    "LG": "Local Government",
    "LOCAL GOVERNMENT": "Local Government",
    "STATE": "State",
    "TRIBAL": "Indian tribe or tribal organization",
    "IHE": "Institution of Higher Education (IHE)",
    "HIGHER ED": "Institution of Higher Education (IHE)",
    "HIGHER-ED": "Institution of Higher Education (IHE)",
    "HIGHER_ED": "Institution of Higher Education (IHE)",
    "NON-PROFIT": "Non-profit",
    "NON PROFIT": "Non-profit",
    "NPO": "Non-profit",
    "UNKNOWN": "Unknown",
}

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

ABBR_TO_STATE = {v.upper(): k for k, v in US_STATE_CODES.items()}

# SOURCING: Convert a state name like "North Carolina" into its 2-letter code like "NC".
# No Changes
def to_state_abbr(state_val: str) -> str:
    s = normalize(state_val)
    if not s:
        return ""
    s2 = s.strip().upper()
    if len(s2) == 2 and s2.isalpha():
        return s2
    return STATE_NAME_TO_ABBR.get(s.strip().lower(), s2[:2])


# SOURCING: Convert your sector text into the label used by the FAC website filters.
# No Changes
def sector_to_fac_label(sector_val: str) -> str:
    s = normalize(sector_val)
    if not s:
        return ""
    key = s.strip().upper()
    if key in SECTOR_TO_LABEL:
        return SECTOR_TO_LABEL[key]
    key2 = re.sub(r"\s+", " ", s.strip().upper())
    if key2 in SECTOR_TO_LABEL:
        return SECTOR_TO_LABEL[key2]
    return s.strip()


# =========================================================
# PLAYWRIGHT UI HELPERS
# =========================================================
# SOURCING: Open a web page in Playwright and wait until it fully loads.
# No Changes
def safe_goto(page, url: str):
    page.goto(url, wait_until="domcontentloaded", timeout=PAGE_LOAD_TIMEOUT_MS)
    page.wait_for_load_state("networkidle", timeout=PAGE_LOAD_TIMEOUT_MS)
    page.wait_for_timeout(500)

# SOURCING: Maximize the Playwright browser window for stable clicking and downloads.
# No Changes
def maximize_playwright_page(page):
    """
    Make the current Playwright page/window full screen (Windows-friendly).
    Works best when context viewport=None and launch args include --start-maximized.
    """
    try:
        # Try to resize to full available screen
        page.evaluate("""
            () => {
                try {
                    window.moveTo(0, 0);
                    window.resizeTo(screen.width, screen.height);
                } catch(e) {}
            }
        """)
    except Exception:
        pass

    # Also try viewport resize (helps if viewport is set)
    try:
        vw = page.evaluate("() => screen.width") or 1920
        vh = page.evaluate("() => screen.height") or 1080
        page.set_viewport_size({"width": int(vw), "height": int(vh)})
    except Exception:
        pass

# SOURCING: Open the correct filter section on the FAC Advanced Search page.
# No Changes
def ensure_accordion_open(page, controls_id: str):
    btn = page.locator(f"button.usa-accordion__button[aria-controls='{controls_id}']").first
    btn.wait_for(state="attached", timeout=UI_TIMEOUT_MS)

    expanded = btn.get_attribute("aria-expanded")
    if expanded is None:
        try:
            btn.click(timeout=UI_TIMEOUT_MS)
        except Exception:
            pass
        return

    if expanded.strip().lower() == "false":
        btn.click(timeout=UI_TIMEOUT_MS)
        page.wait_for_timeout(200)


# SOURCING: Fill the UEI field on the FAC Advanced Search page.
# No Changes
def fill_uei(page, uei: str):
    ensure_accordion_open(page, "uei-or-ein")
    uei = normalize(uei)
    if not uei:
        raise ValueError("UEI is blank in MASTER LIST row")

    page.locator("#input_uei-or-ein").wait_for(state="visible", timeout=UI_TIMEOUT_MS)
    page.fill("#input_uei-or-ein", uei)
    page.wait_for_timeout(200)


# SOURCING: Set the Entity Type/Sector filter on the FAC Advanced Search page.
# No Changes
def set_entity_type(page, sector_val: str):
    ensure_accordion_open(page, "entity-type")

    label = sector_to_fac_label(sector_val)
    if not label:
        raise ValueError("Sector/Entity type is blank in MASTER LIST row")

    checkbox_id = f"entity_type-{label}"

    page.evaluate(
        """
        (targetId) => {
          const boxes = Array.from(document.querySelectorAll("input[name='entity_type'][type='checkbox']"));
          boxes.forEach(b => {
            b.checked = false;
            b.dispatchEvent(new Event('input', {bubbles:true}));
            b.dispatchEvent(new Event('change', {bubbles:true}));
          });

          const el = document.getElementById(targetId);
          if (el){
            el.checked = true;
            el.dispatchEvent(new Event('input', {bubbles:true}));
            el.dispatchEvent(new Event('change', {bubbles:true}));
          }
        }
        """,
        checkbox_id
    )

    exists = page.evaluate("(targetId) => !!document.getElementById(targetId)", checkbox_id)
    if not exists:
        raise RuntimeError(
            f"Entity type checkbox not found for sector='{sector_val}'. "
            f"Tried label='{label}' and id='{checkbox_id}'. "
            f"Update SECTOR_TO_LABEL mapping if needed."
        )

    page.wait_for_timeout(200)


# SOURCING: Click the Search button on the FAC website and wait for results.
# No Changes
def click_search(page):
    btn = page.locator("input.usa-button[type='submit'][value='Search']").first
    btn.wait_for(state="visible", timeout=UI_TIMEOUT_MS)
    btn.scroll_into_view_if_needed(timeout=UI_TIMEOUT_MS)
    page.wait_for_timeout(150)
    btn.click(timeout=UI_TIMEOUT_MS)
    page.wait_for_load_state("networkidle", timeout=PAGE_LOAD_TIMEOUT_MS)
    page.wait_for_timeout(700)


# SOURCING: Download the "Download all" Excel from the FAC results and return the saved file path.
# No Changes
def download_all_excel(page, download_dir: Path, file_tag: str) -> Path:
    candidates = [
        page.get_by_role("button", name=re.compile(r"download all", re.I)),
        page.get_by_role("link", name=re.compile(r"download all", re.I)),
        page.locator("button:has-text('Download all')"),
        page.locator("a:has-text('Download all')"),
    ]

    download_btn = None
    last_err = None

    for loc in candidates:
        try:
            if loc.count() > 0:
                loc.first.wait_for(state="visible", timeout=UI_TIMEOUT_MS)
                download_btn = loc.first
                break
        except Exception as e:
            last_err = e

    if download_btn is None:
        raise RuntimeError("Could not find 'Download all' button/link after search.") from last_err

    download_dir.mkdir(parents=True, exist_ok=True)

    with page.expect_download(timeout=DOWNLOAD_TIMEOUT_MS) as d:
        download_btn.scroll_into_view_if_needed(timeout=UI_TIMEOUT_MS)
        page.wait_for_timeout(200)
        download_btn.click(timeout=UI_TIMEOUT_MS)
    download = d.value

    suggested = download.suggested_filename
    suffix = Path(suggested).suffix if suggested else ".xlsx"
    if not suffix:
        suffix = ".xlsx"

    # out_name = sanitize_filename(file_tag) + suffix
    out_name = file_tag + suffix
    out_path = download_dir / out_name
    download.save_as(out_path)
    return out_path

# No Changes
def rename_downloaded_excel_to_match_pdf_name(downloaded_xlsx: Path, download_key: str, download_dir: Path) -> Path:
    """
    Rename/move the downloaded FAC Excel to match the SAME naming rule as PDFs.
    Output name: <download_key>.xlsx inside DOWNLOAD_DIR.

    - Keeps folder = download_dir.
    - Replaces existing file if the target name already exists.
    - If rename fails, returns the original path (does not break run).
    """
    try:
        if not downloaded_xlsx or not Path(downloaded_xlsx).exists():
            return downloaded_xlsx

        download_dir.mkdir(parents=True, exist_ok=True)
        target_path = download_dir / f"{sanitize_filename(download_key)}.xlsx"

        # If target exists, remove it (latest wins)
        if target_path.exists():
            try:
                target_path.unlink()
            except Exception:
                pass

        # Rename/move (same drive rename is fastest)
        try:
            Path(downloaded_xlsx).replace(target_path)
            return target_path
        except Exception:
            # Fallback move (works even if replace fails)
            try:
                shutil.move(str(downloaded_xlsx), str(target_path))
                return target_path
            except Exception:
                return downloaded_xlsx

    except Exception:
        return downloaded_xlsx

# =========================================================
# DOWNLOAD RETRY WRAPPERS
# =========================================================

# SOURCING: Check if a downloaded file exists and is not empty.
# No Changes
def _file_ok(p: Path) -> bool:
    try:
        return p.exists() and p.is_file() and p.stat().st_size > 0
    except Exception:
        return False


# SOURCING: Retry the FAC Excel download a few times if it fails.
# No Changes
def download_all_excel_with_retry(page, download_dir: Path, file_tag: str, retries: int = DOWNLOAD_RETRIES) -> Path:
    """Retry Download-all Excel up to 1 + retries attempts."""
    last_err = None
    for attempt in range(1, retries + 2):
        try:
            p = download_all_excel(page, download_dir, file_tag)
            if _file_ok(p):
                return p
            last_err = RuntimeError(f"Excel downloaded but file missing/empty: {p}")
        except Exception as e:
            last_err = e
        
        write_log
        print(f"[RETRY] Excel download attempt {attempt}/{retries+1} failed: {last_err}")
        try:
            page.wait_for_timeout(int(RETRY_WAIT_SECONDS * 1000))
        except Exception:
            time.sleep(RETRY_WAIT_SECONDS)

    raise RuntimeError(f"Excel download failed after {retries+1} attempts") from last_err


# SOURCING: Download the PDF report using the browser download flow (Playwright).
# No Changes
def download_pdf_via_browser(context, pdf_url: str, save_path: Path) -> None:
    """
    Trigger PDF download using browser context.
    Fullscreens the temporary window used for download.
    """
    save_path.parent.mkdir(parents=True, exist_ok=True)

    tmp = context.new_page()
    tmp.set_default_timeout(UI_TIMEOUT_MS)

    # ✅ Make the temporary download window full screen
    maximize_playwright_page(tmp)

    try:
        with tmp.expect_download(timeout=PDF_DOWNLOAD_TIMEOUT_MS) as d:
            try:
                tmp.goto(pdf_url, wait_until="commit", timeout=PAGE_LOAD_TIMEOUT_MS)
            except PlaywrightError as e:
                # When download starts, Playwright can raise: "Download is starting"
                msg = str(e)
                if "Download is starting" not in msg:
                    raise

        download = d.value
        download.save_as(save_path)

    finally:
        try:
            tmp.close()
        except Exception:
            pass

# =========================================================
# MASTER LIST: ROWS + OUTPUT COLS
# =========================================================
# SOURCING: Retry the PDF download a few times if it fails.
# No Changes
def download_pdf_via_browser_with_retry(context, pdf_url: str, save_path: Path, retries: int = DOWNLOAD_RETRIES) -> None:
    """Retry PDF download up to 1 + retries attempts."""
    last_err = None
    for attempt in range(1, retries + 2):
        try:
            download_pdf_via_browser(context, pdf_url, save_path)
            if _file_ok(save_path):
                return
            last_err = RuntimeError(f"PDF downloaded but file missing/empty: {save_path}")
        except Exception as e:
            last_err = e

        print(f"[RETRY] PDF download attempt {attempt}/{retries+1} failed: {last_err}")
        time.sleep(RETRY_WAIT_SECONDS)

    raise RuntimeError(f"PDF download failed after {retries+1} attempts") from last_err



# =========================================================
# GENERAL SHEET: RELIANT ROW + FIELD EXTRACTION
# =========================================================
# SOURCING: Convert different date formats from Excel into a Python date object.
# No Changes
def as_date(v) -> Optional[date]:
    if v is None:
        return None
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    s = str(v).strip()
    if not s:
        return None
    for fmt in ("%m/%d/%Y", "%m/%d/%y", "%Y-%m-%d", "%d-%m-%Y", "%d/%m/%Y"):
        try:
            return datetime.strptime(s, fmt).date()
        except Exception:
            pass
    return None


GENERAL_COL_CANDIDATES = {
    "report_id": ["Report ID", "REPORT ID", "report_id"],
    "audit_year": ["Audit Year", "AUDIT YEAR", "Year", "AY"],
    "auditee_uei": ["Auditee UEI", "AUDITEE UEI", "UEI", "UEI Number"],
    "auditee_ein": ["Auditee EIN", "AUDITEE EIN", "EIN"],
    "auditee_state": ["Auditee State", "AUDITEE STATE", "State", "STATE"],
    "entity_type": ["Entity Type", "ENTITY TYPE", "Auditee Entity Type"],

    "fy_end_date": ["FY End Date", "FYE Date", "Fiscal Year End Date", "FY END DATE"],
    "submitted_date": ["Submitted Date", "SUBMITTED DATE", "Date of Release", "FAC acceptance date", "FAC Acceptance Date"],
    "auditee_certified_date": ["Auditee Certified Date", "AUDITEE CERTIFIED DATE", "Auditor's Signature Date", "AUDITOR'S SIGNATURE DATE"],
    "auditee_city": ["Auditee City", "AUDITEE CITY", "City", "CITY"],
    "auditee_contact_name": ["Auditee Contact Name", "AUDITEE CONTACT NAME", "Entity Contact: Name", "ENTITY CONTACT: NAME"],
    "auditee_contact_title": ["Auditee Contact Title", "AUDITEE CONTACT TITLE", "Entity Contact: Title", "ENTITY CONTACT: TITLE"],
    "auditee_phone": ["Auditee Phone", "AUDITEE PHONE", "Entity Contact: Phone", "ENTITY CONTACT: PHONE"],
    "auditee_email": ["Auditee Email", "AUDITEE EMAIL", "Entity Contact: Email", "ENTITY CONTACT: EMAIL ID", "Entity Contact: Email ID"],

    "auditor_firm_name": ["Auditor Firm Name", "AUDITOR FIRM NAME", "Auditor's Firm Name", "AUDITOR'S FARM NAME", "Auditor's Farm Name"],
    "auditor_contact_name": ["Auditor Contact Name", "AUDITOR CONTACT NAME", "Auditor's Name", "AUDITOR'S NAME"],
    "auditor_phone": ["Auditor Phone", "AUDITOR PHONE", "Auditor's Contact: Phone", "AUDITOR'S CONTACT: PHONE"],
    "auditor_email": ["Auditor Email", "AUDITOR EMAIL", "Auditor's Contact: Email", "AUDITOR'S CONTACT: EMAIL ID"],
}


# SOURCING: Find the key columns in the downloaded FAC Excel "General" sheet.
# No Changes (Works on downloaded excel not the metadata excel)
def locate_general_columns(ws_general) -> Dict[str, int]:
    cols = {}
    for key, candidates in GENERAL_COL_CANDIDATES.items():
        c = find_col(ws_general, candidates, header_row=1)
        if c is not None:
            cols[key] = c
    return cols


# SOURCING: Pick the best matching row in the FAC "General" sheet using UEI/Year/State/Sector.
# No Changes (Works on downloaded excel not the metadata excel)
def best_reliant_row(ws_general, gen_cols: Dict[str, int], criteria: Dict[str, str]) -> Optional[int]:
    uei = normalize(criteria.get("uei"))
    ay = normalize(criteria.get("year"))
    state = to_state_abbr(criteria.get("state", ""))
    ent = sector_to_fac_label(criteria.get("sector", ""))

    c_report_id = gen_cols.get("report_id")
    c_ay = gen_cols.get("audit_year")
    c_uei = gen_cols.get("auditee_uei")
    c_state = gen_cols.get("auditee_state")
    c_entity = gen_cols.get("entity_type")
    c_submitted = gen_cols.get("submitted_date")

    best_score = -1
    best_date = None
    best_row = None

    for r in range(2, ws_general.max_row + 1):
        rid = normalize(get_cell(ws_general, r, c_report_id)) if c_report_id else ""
        row_uei = normalize(get_cell(ws_general, r, c_uei)) if c_uei else ""
        if not (rid or row_uei):
            continue

        score = 0

        if uei and row_uei and uei.replace(" ", "") == row_uei.replace(" ", ""):
            score += 50

        if ay and c_ay:
            row_ay = normalize(get_cell(ws_general, r, c_ay))
            if row_ay == ay:
                score += 20

        if state and c_state:
            row_st = to_state_abbr(get_cell(ws_general, r, c_state))
            if row_st == state:
                score += 10

        if ent and c_entity:
            row_ent = normalize(get_cell(ws_general, r, c_entity))
            if row_ent.strip().lower() == ent.strip().lower():
                score += 8
            elif ent.strip().lower() in row_ent.strip().lower() or row_ent.strip().lower() in ent.strip().lower():
                score += 5

        sub_dt = as_date(get_cell(ws_general, r, c_submitted)) if c_submitted else None

        if score > best_score:
            best_score = score
            best_date = sub_dt
            best_row = r
        elif score == best_score:
            if sub_dt and (best_date is None or sub_dt > best_date):
                best_date = sub_dt
                best_row = r

    return best_row


# SOURCING: Build the FAC PDF URL from the Report ID (or use it directly if already a URL).
# No Changes
def build_pdf_url(report_id_value) -> Optional[str]:
    rid = normalize(report_id_value)
    if not rid:
        return None
    if rid.startswith("http://") or rid.startswith("https://"):
        return rid
    return f"{BASE}/dissemination/report/pdf/{rid}"


# =========================================================
# PDF DOWNLOAD (UPDATED FIX)
# =========================================================
# SOURCING: Loop through MASTER LIST rows and return the input values needed for searching.
#!!! Not called
def iter_master_rows(ws_master, year_col: int, uei_col: int, state_col: int, sector_col: int, start_row: int = 2):
    for r in range(start_row, ws_master.max_row + 1):
        year = normalize(get_cell(ws_master, r, year_col))
        uei = normalize(get_cell(ws_master, r, uei_col))
        state = normalize(get_cell(ws_master, r, state_col))
        sector = normalize(get_cell(ws_master, r, sector_col))

        if not (year or uei or state or sector):
            continue

        yield {"row_index": r, "year": year, "uei": uei, "state": state, "sector": sector}


# SOURCING: Open the downloaded FAC Excel, pick the best row, update MASTER LIST, and prepare PDF download.
#!!!Done
def process_downloaded_excel_and_update_master(
    downloaded_xlsx: Path,
    db,
    row_id: int,
    processing_id: int,
    processing_code: str,
    criteria: Dict[str, str],
    playwright_context
) -> Tuple[Optional[str], Optional[Path], Any]:

    if not downloaded_xlsx.exists():
        raise FileNotFoundError(f"Downloaded excel not found: {downloaded_xlsx}")

    xwb = load_workbook(downloaded_xlsx, data_only=True)
    sheet_name = find_sheet_case_insensitive(xwb, "General")
    if not sheet_name:
        xwb.close()
        raise RuntimeError(f"'General' sheet not found in downloaded file: {downloaded_xlsx.name}")

    ws_general = xwb[sheet_name]
    gen_cols = locate_general_columns(ws_general)

    reliant_row = best_reliant_row(ws_general, gen_cols, criteria)
    if reliant_row is None:
        reliant_row = 2  # fallback

    g = extract_general_fields(ws_general, gen_cols, reliant_row)
    pdf_url = build_pdf_url(g.get("report_id"))

    # Update TProcessStatus & TProcessingAdditionalInfo using the fac excel
    update_master_from_general(db, row_id, processing_id, g, pdf_url) #!!!done

    # entity name for better file name
    # entity_name_col = find_col(ws_general, ["Auditee Name", "AUDITEE NAME", "Entity Name", "ENTITY NAME", "Name"], 1)
    # entity_name = normalize(get_cell(ws_general, reliant_row, entity_name_col)) if entity_name_col else "UNKNOWN"

    # download_key = build_download_key(
    #     name=entity_name or "UNKNOWN",
    #     ay=criteria.get("year"),
    #     auditee_city=normalize(g.get("auditee_city")),
    #     auditee_state=to_state_abbr(g.get("auditee_state")) or to_state_abbr(criteria.get("state", "")),
    #     entity_type=normalize(g.get("entity_type")) or sector_to_fac_label(criteria.get("sector", "")),
    #     auditee_uei=normalize(g.get("auditee_uei")) or normalize(criteria.get("uei")),
    #     auditee_ein=normalize(g.get("auditee_ein")),
    # )

    # IMPORTANT: Close workbook before renaming file (Windows lock safety)
    xwb.close()

    # ✅ NEW: Rename Excel using same naming rule as PDF
    # downloaded_xlsx = rename_downloaded_excel_to_match_pdf_name(downloaded_xlsx, download_key, DOWNLOAD_DIR)

    # -----------------------------
    # EXISTING HANDLING: PDF NOT FOUND
    # -----------------------------
    if not pdf_url:
        mark_master_row_pdf_not_found(db, row_id, processing_id)
        return PDF_NOT_FOUND_TEXT, None, None

    # pdf_saved_path = PDF_DIR / f"{download_key}.pdf"
    pdf_saved_path = PDF_DIR / f"{processing_code}_AR.pdf"

    try:
        download_pdf_via_browser_with_retry(playwright_context, pdf_url, pdf_saved_path)

        if not pdf_saved_path.exists() or pdf_saved_path.stat().st_size == 0:
            mark_master_row_pdf_not_found(db, row_id, processing_id)
            return PDF_NOT_FOUND_TEXT, None, None

        return pdf_url, pdf_saved_path, g.get("fy_end_date")

    except RuntimeError as e:
        cause = getattr(e, "__cause__", None)
        if isinstance(cause, PlaywrightTimeoutError) or ("timeout" in str(cause).lower()):
            mark_master_row_pdf_not_found(db, row_id, processing_id)
            return PDF_NOT_FOUND_TEXT, None, None

        raise

    except PlaywrightTimeoutError:
        mark_master_row_pdf_not_found(db, row_id, processing_id)
        return PDF_NOT_FOUND_TEXT, None, None
    
# =========================================================
# MAIN RUN: FAC SEARCH -> EXCEL -> GENERAL -> PDF -> MASTER UPDATE
# =========================================================
# SOURCING: Norm space.
# No Changes
def _norm_space(s: str) -> str:
    return re.sub(r"\s+", " ", s or "").strip()


# SOURCING: Remove bracketed text so opinion classification does not get confused by exceptions.
# No Changes
def _remove_parentheses(text: str) -> str:
    """Remove parenthetical parts like (except for ...) so they don't trigger 'Qualified' wrongly."""
    if not text:
        return ""
    return re.sub(r"\([^)]*\)", "", text)


# SOURCING: Get the first few meaningful lines of a page (used for heading detection).
# No Changes
def _top_non_empty_lines(page_text: str, max_lines: int = 5) -> List[str]:
    """
    Return the first `max_lines` non-empty lines from extracted page text.
    """
    if not page_text:
        return []
    lines = [ln.strip() for ln in (page_text or "").splitlines() if ln.strip()]
    return lines[:max_lines]

# SOURCING: Aggressively clean text to letters/numbers/spaces for robust contains matching.
# No Changes
def clean_for_contains_match(s: str) -> str:
    """Aggressive cleanup: keep only letters/numbers/spaces (lowercase)."""
    s = normalize_pdf_text(s).lower()
    s = re.sub(r"[^a-z0-9 ]+", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s


# SOURCING: Normalize PDF text so matching works even with unusual hyphens, quotes, and spacing.
# No Changes
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
# No Changes
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
# No Changes
def split_sentences(text: str) -> List[str]:
    """Same idea as friend's: split by sentence punctuation."""
    if not text:
        return []
    return [s.strip() for s in re.split(r"(?<=[.!?])\s+", text) if s.strip()]



# =========================================================
# VALIDATION (PDF CHECKS + PROOFS + DETAILED VALIDATION CHECK)
# =========================================================

# VALIDATION: Split text into cleaned tokens (words) for matching across PDFs.
# No Changes
def tokenize(s: str) -> list:
    return [t for t in clean_for_contains_match(s).split() if t]


# VALIDATION: Find where a token sequence appears inside a larger token list (used for highlighting).
# No Changes
def find_token_window(page_tokens: list, target_tokens: list) -> tuple:
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
# No Changes
def is_table_of_contents_page(page_text: str) -> bool:
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

# VALIDATION: Detect whether a page is the Independent Auditor’s Report page based on its heading.
# No Changes
def is_independent_auditors_report_page(page_text: str) -> bool:
    """
    True if a page contains a standalone audit-report heading line.

    FIXES:
    1) Handles bold/markdown markers like **Independent Auditors’ Report**
       by stripping leading/trailing non-alphanumeric wrappers.
    2) Accepts additional common heading family used by firms like Grant Thornton:
       "REPORT OF INDEPENDENT CERTIFIED PUBLIC ACCOUNTANTS"
    3) Scans more than the first 5 lines because some PDFs place the heading
       after letterhead or within a boxed/table header.
    """
    if not page_text:
        return False

    lines = [ln.strip() for ln in (page_text or "").splitlines() if ln.strip()]
    # Scan first 80 non-empty lines (safe + catches letterhead/table headers)
    scan_lines = lines[:80]

    def _norm(s: str) -> str:
        s = normalize_pdf_text(s)
        s = s.replace("’", "'").replace("‘", "'")
        s = re.sub(r"\s+", " ", s).strip()
        s = s.rstrip(" .:-").strip()
        # Strip wrappers like '**', '__', bullets, etc. at BOTH ENDS
        s = re.sub(r"^[^A-Za-z0-9]+", "", s)
        s = re.sub(r"[^A-Za-z0-9]+$", "", s)
        return s.strip()

    heading_res = [
        # Independent Auditor's Report / Independent Auditors' Report
        re.compile(r"^INDEPENDENT\s+AUDITOR(?:S)?\s*'?S?\s+REPORT(?:S)?$", re.I),

        # Report of Independent Auditor(s)
        re.compile(r"^REPORT\s+OF\s+INDEPENDENT\s+AUDITOR(?:S)?$", re.I),

        # Report of Independent Certified Public Accountants (Nassau-style)
        re.compile(r"^REPORT\s+OF\s+INDEPENDENT\s+CERTIFIED\s+PUBLIC\s+ACCOUNTANTS$", re.I),

        # Independent Certified Public Accountants' Report (variant)
        re.compile(r"^INDEPENDENT\s+CERTIFIED\s+PUBLIC\s+ACCOUNTANTS\s*'?S?\s+REPORT$", re.I),
    ]

    for ln in scan_lines:
        n = _norm(ln)
        for rx in heading_res:
            if rx.match(n):
                return True

# VALIDATION: Find the bounding box of the audit report heading so it can be highlighted in proof images.
# No Changes
def find_audit_report_heading_box(page) -> Optional[dict]:
    """
    Returns bounding box of the audit report heading line if present.

    Works with BOTH:
      - "INDEPENDENT AUDITOR'S REPORT" / "INDEPENDENT AUDITORS' REPORT"
      - "REPORT OF INDEPENDENT AUDITOR" / "REPORT OF INDEPENDENT AUDITORS"
    """
    page_text = cached_page_text(page)
    if not page_text.strip():
        return None

    lines = [ln.strip() for ln in page_text.splitlines() if ln.strip()]
    top_lines = lines[:5]

    def _norm(s: str) -> str:
        s = normalize_pdf_text(s)
        s = s.replace("’", "'").replace("‘", "'")
        s = re.sub(r"\s+", " ", s).strip()
        s = s.rstrip(" .:-").strip()
        return s

    heading_res = [
        re.compile(r"^INDEPENDENT\s+AUDITOR(?:S)?\s*'?S?\s+REPORT(?:S)?$", re.I),
        re.compile(r"^REPORT\s+OF\s+INDEPENDENT\s+AUDITOR(?:S)?$", re.I),
    ]

    heading_line = None
    for ln in top_lines:
        n = _norm(ln)
        for rx in heading_res:
            if rx.match(n):
                heading_line = ln.strip()
                break
        if heading_line:
            break

    if not heading_line:
        return None

    words = cached_page_words(page, use_text_flow=True)
    if not words:
        return None

    target_tokens = tokenize(heading_line)
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

# VALIDATION: Open File Explorer (or image viewer) to show the saved proof image to the user.
# No Changes
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
# No Changes
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
# No Changes
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


# VALIDATION: Search the PDF for the entity name and save a proof image showing the match.
# No changes
def visible_name_validation_with_proof(pdf_path: Path, issuer_name: str, proof_png_path: Path) -> dict:
    """Search pages until TOC is reached (TOC page included)."""
    issuer_name = (issuer_name or "").strip()
    if not issuer_name:
        return {"status": "FAIL", "reason": "Blank ISSUER NAME in MASTER LIST", "page": None}

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

# VALIDATION: Check the state appears on the same page as the entity name and save a proof image.
# No changes
def visible_state_validation_with_proof(
    pdf_path: Path,
    master_state_abbr: str,
    name_found_page: int,
    proof_png_path: Path
) -> dict:
    """
    STATE Validation (Same page only):
      - Use the page where NAME was found (name_found_page)
      - Search for full state name first (e.g. Massachusetts)
      - If not found, search for abbreviation (e.g. MA)
      - If found, highlight it with yellow fill and save PNG to proof_png_path
      - No popup for state

    Returns: {"status": PASS/FAIL/NOT EXTRACTABLE, "reason":..., "page": <page>}
    """

    abbr = (master_state_abbr or "").strip().upper()
    if not abbr:
        return {"status": "FAIL", "reason": "Blank State Abbreviation in MASTER LIST", "page": name_found_page}

    full_state = ABBR_TO_STATE.get(abbr, "")
    full_tokens = tokenize(full_state) if full_state else []
    abbr_tokens = [abbr.lower()]  # for token compare

    try:
        with open_pdf_maybe_shared(pdf_path) as pdf:
            if not pdf.pages:
                return {"status": "NOT EXTRACTABLE", "reason": "PDF has no pages", "page": name_found_page}

            if not name_found_page or name_found_page < 1 or name_found_page > len(pdf.pages):
                return {"status": "FAIL", "reason": "Invalid Name-found page for State validation", "page": name_found_page}

            page = pdf.pages[name_found_page - 1]
            page_text = cached_page_text(page)

            im = page.to_image(resolution=PROOF_IMAGE_RESOLUTION)

            if not page_text.strip():
                # scanned/image PDF
                im.save(proof_png_path)
                return {"status": "NOT EXTRACTABLE", "reason": "No extractable text on name page", "page": name_found_page}

            # Build page tokens aligned to words
            words = cached_page_words(page, use_text_flow=True)
            page_tokens = []
            token_word_map = []
            for w in words:
                wtoks = tokenize(w.get("text", ""))
                for t in wtoks:
                    page_tokens.append(t)
                    token_word_map.append(w)

            # 1) Try FULL STATE NAME first (e.g., "Massachusetts" / "New Hampshire")
            if full_tokens:
                s, e = find_token_window(page_tokens, full_tokens)
                if s is not None and e is not None:
                    matched_words = token_word_map[s:e]
                    x0 = min(w["x0"] for w in matched_words)
                    top = min(w["top"] for w in matched_words)
                    x1 = max(w["x1"] for w in matched_words)
                    bottom = max(w["bottom"] for w in matched_words)
                    box = {"x0": x0, "top": top, "x1": x1, "bottom": bottom}
                    im.draw_rect(box, stroke=HIGHLIGHT_STROKE, fill=HIGHLIGHT_FILL_RGBA)
                    im.save(proof_png_path)
                    return {"status": "PASS", "reason": f"State full name found: {full_state}", "page": name_found_page}

            # 2) If full name not found, try ABBREVIATION (e.g., "MA")
            # Strict match: token must equal abbr (lowercase)
            abbr_l = abbr.lower()
            for w in words:
                wtoks = tokenize(w.get("text", ""))
                if any(t == abbr_l for t in wtoks):
                    box = {"x0": w["x0"], "top": w["top"], "x1": w["x1"], "bottom": w["bottom"]}
                    im.draw_rect(box, stroke=HIGHLIGHT_STROKE, fill=HIGHLIGHT_FILL_RGBA)
                    im.save(proof_png_path)
                    return {"status": "PASS", "reason": f"State abbreviation found: {abbr}", "page": name_found_page}

            # Not found
            im.save(proof_png_path)
            return {"status": "FAIL", "reason": f"State not found (full/abbr) on page {name_found_page}", "page": name_found_page}

    except Exception as e:
        return {"status": "FAIL", "reason": f"Error reading PDF for State validation: {e}", "page": name_found_page}

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
# No Changes
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
# No Changes
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
# No Changes
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
# No Changes
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
    
# =========================================================
# AVAILABILITY OF AUDIT REPORT (Detailed Validation Check - Column F)
# =========================================================

# VALIDATION: Find a specific phrase on a page and return its bounding box (for highlighting).
# No Changes
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
# No Changes
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

# VALIDATION: Extract the opinion paragraph from the audit report (between Opinion and Basis sections).
# No Changes
def get_opinion_paragraph(report_text: str) -> Optional[str]:
    """
    Friend's logic:
    opinions? .... basis for opinions?
    Return the text in between as the 'opinion paragraph'.
    """
    if not report_text:
        return None
    m = re.search(r"\bopinions?\b(.+?)\bbasis\s+for\s+opinions?\b", report_text, re.I | re.S)
    if m:
        return _norm_space(m.group(1))
    return None

# VALIDATION: Pick the best opinion sentence (prefer "in our opinion" if present).
# No Changes
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
# No Changes
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

# No Changes
def find_opinion_anchor_in_pdf(pdf, max_scan_pages: int = 120) -> Optional[dict]:
    """
   hored audit discovery (NO TOC skipping).

    Finds the FIRST page in early PDF that:
      - contains "in our opinion"
      - produces a non-empty classify_opinion() label
      - has an Independent Auditor(s) Report heading on:
           same page OR previous 1-2 pages

    Returns:
      {
        "opinion_page": int,          # 1-based
        "heading_page": int,          # 1-based (page where heading exists)
        "label": str,                # classify_opinion label
        "opinion_sentence": str      # extracted sentence (may be "")
      }
    or None if not found.
    """
    if not pdf or not getattr(pdf, "pages", None):
        return None

    limit = min(max_scan_pages, len(pdf.pages))

    for idx in range(limit):
        page = pdf.pages[idx]
        txt = cached_page_text(page) or ""
        if not txt.strip():
            continue

        # Must contain opinion phrase
        if not re.search(r"\bin\s+our\s+opinion\b", txt, re.I):
            continue

        # Try to classify based on this page text directly
        op_sentence = find_opinion_sentence(txt) or ""
        label = classify_opinion(txt, opinion_sentence=op_sentence)

        if not label:
            continue

        # Heading check: same page or previous 1-2 pages
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

# No Changes
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

# VALIDATION: Safely detect audit report pages using available heading check functions.
# No Changes
def is_audit_report_page_safe(page) -> bool:
    """
    Safe detector:
    - If is_independent_auditors_report_page_obj exists, use it
    - Else use is_independent_auditors_report_page(extract_text)
    """
    try:
        fn = globals().get("is_independent_auditors_report_page_obj")
        if callable(fn):
            return bool(fn(page))
    except Exception:
        pass

    # Fallback: text-based
    try:
        txt = page.extract_text() or ""
        return is_independent_auditors_report_page(txt)
    except Exception:
        return False

# VALIDATION: Find the audit report start page and return a small page range to search validations in.
# No Changes
def find_audit_report_page_range(pdf) -> List[Tuple[int, Any]]:
    """
    Find audit report start page using your audit-report heading logic,
    then return up to 10 pages from there (friend used start:start+10).
    Stop early if we hit MD&A (management discussion) heading patterns.
    """
    start_idx = None
    pages_out = []

    # --- Find first audit-report heading page ---
    for idx, page in enumerate(pdf.pages):
        page_num = idx + 1
        txt = page.extract_text() or ""
        if not txt.strip():
            continue

        if TOC_STOP_ENABLED and is_table_of_contents_page(txt):
            continue

        # Prefer object-based detector if exists in your script; otherwise use text-based
        try:
            is_audit = is_audit_report_page_safe(page)
        except Exception:
            is_audit = is_independent_auditors_report_page(txt)

        if is_audit:
            start_idx = idx
            break

    if start_idx is None:
        return []

    # --- Collect up to 10 pages from start (friend style) ---
    end = min(start_idx + 10, len(pdf.pages))
    for idx in range(start_idx, end):
        page = pdf.pages[idx]
        page_num = idx + 1
        txt = page.extract_text() or ""
        norm = _norm_space(txt).lower()

        # stop if MD&A begins (friend used a condition like this)
        if pages_out and re.search(r"\bmanagement'?s?\s+discussion\s*(?:and|&)\s*analysis\b", norm, re.I):
            break

        pages_out.append((page_num, page))

    return pages_out

# VALIDATION: Select a short, stable phrase to highlight inside a long opinion sentence.
# No Changes
def pick_highlight_phrase(opinion_sentence: str) -> Optional[str]:
    """
    Instead of trying to highlight the entire long sentence (can be fragile),
    highlight the key phrase inside it.
    """
    if not opinion_sentence:
        return None

    phrases = [
        "adverse opinion",
        "disclaimer of opinion",
        "do not express an opinion",
        "qualified opinion",
        "unmodified opinion",
        "present fairly",
        "fairly presented",
        "true and fair",
        "except for",
    ]

    s = opinion_sentence.lower()
    for ph in phrases:
        if ph in s:
            return ph
    return None

# VALIDATION: Find the first audit report page in the PDF (skipping TOC pages).
# No Changes
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

      
# VALIDATION: Validate and classify the audit opinion and write the result with proof image.

def write_audit_opinion_validation(
    ws_validation,
    row_index: int,
    pdf_path: Optional[Path],
    audit_res: Optional[dict] = None,
    show_popup: bool = True
) -> dict:
    """
    Audit Opinion -> Detailed Validation Check Column G (7)

    NEW LOGIC (Opinion-anchored, NO TOC skipping):
      1) Locate opinion anchor page + heading page
      2) Build anchor window pages (<=20 pages)
      3) Extract opinion paragraph (Opinion..Basis) from window text
      4) classify_opinion(...) => PASS - <label>
    """
    # Column G = 7
    if not pdf_path or not Path(pdf_path).exists():
        ws_validation.cell(row=row_index, column=7).value = "FAIL"
        return {"status": "FAIL", "page": None, "audit_start_page": None, "proof_file": None, "boxes": [], "title": f"Row {row_index} | Audit Opinion = FAIL (PDF missing)"}

    try:
        with open_pdf_maybe_shared(pdf_path) as pdf:
            if not pdf.pages:
                ws_validation.cell(row=row_index, column=7).value = "NOT EXTRACTABLE"
                return {"status": "NOT EXTRACTABLE", "page": None, "audit_start_page": None, "proof_file": None, "boxes": [], "title": f"Row {row_index} | Audit Opinion = NOT EXTRACTABLE"}

            anchor = find_opinion_anchor_in_pdf(pdf, max_scan_pages=120)
            if not anchor:
                ws_validation.cell(row=row_index, column=7).value = "FAIL"
                proof_file = PROOF_DIR_OPINION / f"Row{row_index}_OPINION_FAIL_NoOpinionAnchor.png"
                try:
                    im = pdf.pages[0].to_image(resolution=PROOF_IMAGE_RESOLUTION)
                    im.save(proof_file)
                except Exception:
                    proof_file = None
                title = f"Row {row_index} | Audit Opinion = FAIL (Opinion anchor not found)"
                if show_popup and proof_file:
                    show_proof_popup(proof_file, title, seconds=PROOF_POPUP_SECONDS)
                return {"status": "FAIL", "page": None, "audit_start_page": None, "proof_file": proof_file, "boxes": [], "title": title}

            heading_page_num = int(anchor["heading_page"])
            opinion_page_num = int(anchor["opinion_page"])

            selected_pages = build_audit_anchor_window_pages(
                pdf,
                heading_page_num=heading_page_num,
                opinion_page_num=opinion_page_num,
                max_total_pages=20,
                post_opinion_pages=3
            )

            if not selected_pages:
                ws_validation.cell(row=row_index, column=7).value = "FAIL"
                return {"status": "FAIL", "page": None, "audit_start_page": None, "proof_file": None, "boxes": [], "title": f"Row {row_index} | Audit Opinion = FAIL (No pages selected)"}

            audit_start_page = selected_pages[0][0]

            report_text = "\n".join([(p.extract_text() or "") for _, p in selected_pages])
            report_norm = _norm_space(report_text)

            op_para = get_opinion_paragraph(report_norm)
            opinion_scope = op_para if op_para else report_norm

            op_sentence = find_opinion_sentence(op_para or opinion_scope) or ""
            label = classify_opinion(opinion_scope, opinion_sentence=op_sentence)

            if not label:
                ws_validation.cell(row=row_index, column=7).value = "FAIL"
                first_page_num, first_page = selected_pages[0]
                proof_file = PROOF_DIR_OPINION / f"Row{row_index}_OPINION_FAIL_Page{first_page_num}.png"
                im = first_page.to_image(resolution=PROOF_IMAGE_RESOLUTION)
                im.save(proof_file)
                title = f"Row {row_index} | Audit Opinion = FAIL | Anchor {heading_page_num}->{opinion_page_num}"
                if show_popup:
                    show_proof_popup(proof_file, title, seconds=PROOF_POPUP_SECONDS)
                return {"status": "FAIL", "page": first_page_num, "audit_start_page": audit_start_page, "proof_file": proof_file, "boxes": [], "title": title}

            ws_validation.cell(row=row_index, column=7).value = f"PASS - {label}"

            highlight_phrase = pick_highlight_phrase(op_sentence) or "in our opinion"

            chosen_page_num = opinion_page_num
            chosen_page = pdf.pages[opinion_page_num - 1]
            chosen_target = highlight_phrase

            proof_label = sanitize_filename(label, 30)
            proof_file = PROOF_DIR_OPINION / f"Row{row_index}_OPINION_PASS_{proof_label}_Page{chosen_page_num}.png"
            im = chosen_page.to_image(resolution=PROOF_IMAGE_RESOLUTION)

            boxes = []

            # heading highlight (on heading page only if same page)
            try:
                if chosen_page_num == heading_page_num:
                    hb = find_audit_report_heading_box(chosen_page)
                    if hb:
                        im.draw_rect(hb, stroke=HIGHLIGHT_STROKE, fill=HIGHLIGHT_FILL_RGBA)
                        boxes.append(hb)
            except Exception:
                pass

            # highlight chosen target phrase if bbox found
            try:
                box = find_phrase_box_on_page(chosen_page, chosen_target)
                if box:
                    im.draw_rect(box, stroke=HIGHLIGHT_STROKE, fill=HIGHLIGHT_FILL_RGBA)
                    boxes.append(box)
            except Exception:
                pass

            im.save(proof_file)

            title = f"Row {row_index} | Audit Opinion = PASS - {label} | Anchor {heading_page_num}->{opinion_page_num} | Page {chosen_page_num}"
            if show_popup:
                show_proof_popup(proof_file, title, seconds=PROOF_POPUP_SECONDS)

            return {"status": "PASS", "page": chosen_page_num, "audit_start_page": audit_start_page, "proof_file": proof_file, "boxes": boxes, "title": title}

    except Exception as e:
        ws_validation.cell(row=row_index, column=7).value = "FAIL"
        return {"status": "FAIL", "page": None, "audit_start_page": None, "proof_file": None, "boxes": [], "title": f"Row {row_index} | Audit Opinion = FAIL ({e})"}
    
# =========================================================
# AUDIT REPORT FYE vs REPORT FYE (Detailed Validation Check - Column H)
# =========================================================

# VALIDATION: Find the FY date near the "We have audited" paragraph to confirm audit report period.
# No Changes
def audit_fye_context_ok(page_tokens: list, date_start_index: int) -> bool:
    """
    Flexible audit-period context check for Column H.

    Accepts any of these before the date within a small lookback window:
    - year ended / years ended
    - period ended / for the period
    - through / thru / to (date ranges)
    - as of (common in audit scope sentences)
    - for the year(s) ended

    This is ONLY used inside Audit FYE validation (Column H).
    """
    if not page_tokens or date_start_index is None:
        return False

    lookback = 45
    left = max(0, date_start_index - lookback)
    before = " ".join(page_tokens[left:date_start_index]).lower()

    patterns = [
        r"\byear\s+ended\b",
        r"\byears\s+ended\b",
        r"\bfiscal\s+year\s+ended\b",
        r"\bperiod\s+ended\b",
        r"\bfor\s+the\s+period\b",
        r"\bfor\s+the\s+year\s+ended\b",
        r"\bas\s+of\b",
        r"\bthrough\b",
        r"\bthru\b",
        r"\bto\b",
        r"\bfrom\b",
    ]
    return any(re.search(p, before, re.I) for p in patterns)

# No Changes
def find_fye_near_we_have_audited(page, fy_end_value):
    """
    Find FY end date that appears in/near the paragraph starting with 'we have audited'.

    LOGICAL FIX:
    - First try strict context (existing logic via fye_context_ok).
    - If that fails, allow audit_fye_context_ok (supports 'through', 'period', etc.)
    - Still requires 'we have audited' to appear before the date within a token window
      to avoid random date matches elsewhere.
    Returns: (date_box, we_box) or (None, None)
    """
    d = parse_fy_end_date(fy_end_value)
    if not d:
        return (None, None)

    words = cached_page_words(page, use_text_flow=True)
    if not words:
        return (None, None)

    # Build page tokens using your existing tokenizer approach
    page_text = cached_page_text(page) or ""
    tokens = tokenize(page_text)

    # Must contain 'we have audited' somewhere
    if not re.search(r"\bwe\s+have\s+audited\b", page_text, re.I):
        return (None, None)

    # 1) Find candidate date box using existing strict finder (if available)
    date_box = None
    try:
        date_box = find_fye_box_on_page_with_context(page, fy_end_value)
        if date_box:
            # optionally also find bbox for 'we have audited'
            we_box = None
            try:
                we_box = find_phrase_box_on_page(page, "we have audited")
            except Exception:
                we_box = None
            return (date_box, we_box)
    except Exception:
        date_box = None

    # 2) Flexible fallback:
    # find date anywhere on page, but accept only if audit_fye_context_ok holds
    try:
        any_box = find_fye_box_on_page_anywhere(page, fy_end_value)
    except Exception:
        any_box = None

    if not any_box:
        return (None, None)

    # To apply audit_fye_context_ok we need the token index of the date match.
    # We do a token-sequence search using your date token candidates.
    try:
        candidates = build_fye_token_candidates(d)  # existing helper in your script
    except Exception:
        candidates = []

    # Locate a matching token window for the date in tokens
    for cand in candidates:
        if not cand:
            continue
        start_idx, end_idx = find_token_window(tokens, cand)
        if start_idx is None:
            continue

        # Require that 'we have audited' occurs before the date within a reasonable window
        lookback = 80
        left = max(0, start_idx - lookback)
        before_text = " ".join(tokens[left:start_idx])
        if not re.search(r"\bwe\s+have\s+audited\b", before_text, re.I):
            continue

        # Accept if either strict FYE context OR audit-flex context holds
        if fye_context_ok(tokens, start_idx) or audit_fye_context_ok(tokens, start_idx):
            we_box = None
            try:
                we_box = find_phrase_box_on_page(page, "we have audited")
            except Exception:
                we_box = None
            return (any_box, we_box)

    return (None, None)

# VALIDATION: Fallback scan to find audit-report FY date near "We have audited" when page selection fails.
# No Changes
def find_audit_report_fye_anywhere(pdf, fy_end_value, max_scan_pages: int = 80) -> Optional[Tuple[int, dict, Optional[dict]]]:
    """
    Fallback for Column H:
    Scan early pages for the first occurrence where:
      - 'we have audited' exists
      - FY end date is found near that paragraph (find_fye_near_we_have_audited)

    Returns:
      (page_num, date_box, we_box) where page_num is 1-based.
    """
    limit = min(max_scan_pages, len(pdf.pages))
    for idx in range(limit):
        pg = pdf.pages[idx]
        txt = pg.extract_text() or ""
        if not txt.strip():
            continue

        # Do NOT skip TOC here; just keep it simple and safe.
        date_box, we_box = find_fye_near_we_have_audited(pg, fy_end_value)
        if date_box:
            return (idx + 1, date_box, we_box)

    return None


# =========================================================
# AUDITOR'S SIGNATURE / AUDITOR NAME VALIDATION (Detailed Validation Check - Column I)
# =========================================================

# VALIDATION: Generate safe variants of an auditor firm name (handles &, and, punctuation, suffixes).
# No Changes
def _firm_name_variants(firm_name: str) -> List[str]:
    """
    Ordered auditor firm-name variants.
    - Do NOT add extra words like Advisors / CPAs.
    - Allow '&amp;' <-> '&' <-> 'and'
    - Allow punctuation differences
    - Allow legal suffix removal only (LLC/LLP/LTD/LIMITED/INC/PC/P.C.)
    - Allow compact spacing differences (BerganKDV vs Bergan KDV)

    Example:
      REDW, LLC -> ['REDW, LLC', 'REDW, LLC' (amp normalized), 'REDW LLC', 'REDW']
    """
    s = (firm_name or "").strip()
    if not s:
        return []

    variants: List[str] = []

    def add(x: str):
        x = re.sub(r"\s+", " ", (x or "").strip())
        if x and len(x) >= 2 and x not in variants:
            variants.append(x)

    # 1) exact
    add(s)

    # 2) normalize &amp; -> &
    s_amp = s.replace("&amp;", "&")
    add(s_amp)

    # 3) punctuation removed (keeps words): "REDW, LLC" -> "REDW LLC"
    s_no_punct = re.sub(r"[^A-Za-z0-9 &]+", " ", s_amp)
    s_no_punct = re.sub(r"\s+", " ", s_no_punct).strip()
    add(s_no_punct)

    # 4) normalize '&' to 'and' (allowed substitution)
    s_and = s_no_punct.replace("&", " and ")
    s_and = re.sub(r"\s+", " ", s_and).strip()
    add(s_and)

    # 5) remove connector "and" (to match '&' or punctuation use)
    s_no_and = re.sub(r"\band\b", " ", s_and, flags=re.I)
    s_no_and = re.sub(r"\s+", " ", s_no_and).strip()
    add(s_no_and)

    # 6) remove ONLY legal suffixes (NOT CPAs, Advisors, etc.)
    s_no_suffix = re.sub(r"\b(llp|llc|ltd|limited|inc|p\.c\.|pc)\b", " ", s_no_and, flags=re.I)
    s_no_suffix = re.sub(r"\s+", " ", s_no_suffix).strip()
    add(s_no_suffix)

    # 7) compact spacing variants (for BerganKDV vs Bergan KDV)
    for v in list(variants):
        compact = re.sub(r"\s+", "", v)
        add(compact)

    return variants


# VALIDATION: Convert an auditor firm name into tokens so matching works across formatting differences.
# No Changes
def _auditor_tokens(s: str, remove_suffix: bool = False) -> List[str]:
    """
    Normalize text into tokens for auditor matching.
    Rules:
      - '&amp;' -> '&'
      - '&' -> 'and'
      - punctuation -> space
      - drop 'and' token (so punctuation can act as connector)
      - optional legal suffix removal only
    """
    s = (s or "").lower().strip()
    s = s.replace("&amp;", "&")
    s = s.replace("&", " and ")
    s = re.sub(r"[^a-z0-9 ]+", " ", s)
    s = re.sub(r"\s+", " ", s).strip()

    tokens = s.split()

    # remove connector
    tokens = [t for t in tokens if t != "and"]

    if remove_suffix:
        suffixes = {"llp", "llc", "ltd", "limited", "inc", "pc"}
        tokens = [t for t in tokens if t not in suffixes]

    return tokens


# VALIDATION: Check if the auditor firm name appears in text using flexible token rules.
# No Changes
def _auditor_match_in_text(auditor_name: str, page_text: str) -> bool:
    """
    True if auditor_name matches inside page_text using controlled permutations.
    No extra words are added beyond '&'->'and' normalization.
    """
    page_keep = " ".join(_auditor_tokens(page_text, remove_suffix=False))
    page_nosfx = " ".join(_auditor_tokens(page_text, remove_suffix=True))

    if not page_keep:
        return False

    # Try ordered variants
    for cand in _firm_name_variants(auditor_name):
        cand_keep = " ".join(_auditor_tokens(cand, remove_suffix=False))
        cand_nosfx = " ".join(_auditor_tokens(cand, remove_suffix=True))

        if cand_keep and cand_keep in page_keep:
            return True
        if cand_nosfx and cand_nosfx in page_nosfx:
            return True

        # compact fallback (spacing issues like BerganKDV)
        if cand_keep and cand_keep.replace(" ", "") in page_keep.replace(" ", ""):
            return True
        if cand_nosfx and cand_nosfx.replace(" ", "") in page_nosfx.replace(" ", ""):
            return True

    return False


# VALIDATION: Find the bounding box for the auditor firm name on a page for highlighting.
# No Changes
def _find_auditor_name_box_on_page(page, auditor_name: str) -> Optional[dict]:
    """
    Find a highlight bbox for auditor name using token-aligned word mapping.
    """
    words = cached_page_words(page, use_text_flow=True)
    if not words:
        return None

    # Build page token stream aligned to words
    page_tokens: List[str] = []
    token_word_map: List[dict] = []

    for w in words:
        wtoks = _auditor_tokens(w.get("text", ""), remove_suffix=True)
        for t in wtoks:
            page_tokens.append(t)
            token_word_map.append(w)

    if not page_tokens:
        return None

    # Try each variant (suffix removed token matching is the most stable)
    for cand in _firm_name_variants(auditor_name):
        target_tokens = _auditor_tokens(cand, remove_suffix=True)
        target_tokens = [t for t in target_tokens if t]  # safety

        if not target_tokens:
            continue

        s, e = find_token_window(page_tokens, target_tokens)
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

# VALIDATION: If signature check fails, search the whole PDF for the auditor firm name as fallback.
# No Changes
def find_auditor_firm_anywhere_in_pdf(pdf, firm_name: str) -> Optional[Tuple[int, Optional[dict]]]:
    """
    Fallback search:
    Scan the entire PDF for auditor firm name using the same matching logic:
      - '&amp;' <-> '&' <-> 'and'
      - punctuation tolerant
      - legal suffix removal
      - compact spacing tolerance

    Returns:
      (page_num, box) where page_num is 1-based and box is bbox dict (may be None).
      None if not found.
    """
    firm_name = (firm_name or "").strip()
    if not firm_name:
        return None

    for idx, page in enumerate(pdf.pages):
        page_num = idx + 1
        txt = page.extract_text() or ""
        if not txt.strip():
            continue

        # IMPORTANT: Entire-PDF search means we do NOT skip TOC pages here.
        if not _auditor_match_in_text(firm_name, txt):
            continue

        box = None
        try:
            box = _find_auditor_name_box_on_page(page, firm_name)
        except Exception:
            box = None

        return (page_num, box)

    return None

# VALIDATION: Validate that the auditor firm name/signature appears after the audit report and write PASS/FAIL.

# def write_auditor_signature_validation(
#     ws_validation,
#     row_index: int,
#     pdf_path: Optional[Path],
#     fy_end_value=None,  # kept only to avoid changing your call signature
#     audit_res: Optional[dict] = None,
#     auditor_firm_name: str = ""
# ):
#     """
#     Column I (9): Auditor Name / Signature validation

#     NEW LOGIC (Opinion-anchored, NO TOC skipping):
#       1) Find opinion anchor (opinion_page + heading_page)
#       2) Anchor page = opinion_page (preferred) else heading_page
#       3) Search NEXT 10 pages AFTER anchor for auditor firm name
#       4) If not found, keep your existing full-PDF fallback search
#     """
#     # Column I = 9
#     if not pdf_path or not Path(pdf_path).exists():
#         ws_validation.cell(row=row_index, column=9).value = "FAIL"
#         return

#     firm_name = (auditor_firm_name or "").strip()
#     if not firm_name:
#         ws_validation.cell(row=row_index, column=9).value = "FAIL"
#         return

#     try:
#         with open_pdf_maybe_shared(pdf_path) as pdf:
#             if not pdf.pages:
#                 ws_validation.cell(row=row_index, column=9).value = "NOT EXTRACTABLE"
#                 return

#             anchor = find_opinion_anchor_in_pdf(pdf, max_scan_pages=120)

#             anchor_page_num = None
#             if anchor:
#                 anchor_page_num = int(anchor.get("opinion_page") or 0) or int(anchor.get("heading_page") or 0)

#             if not anchor_page_num:
#                 ws_validation.cell(row=row_index, column=9).value = "FAIL"
#                 return

#             # Search AFTER anchor page, next 10 pages
#             start_idx = anchor_page_num  # 1-based anchor; start_idx=anchor_page_num means next page index (0-based)
#             end_idx = min(start_idx + 10, len(pdf.pages))

#             for idx in range(start_idx, end_idx):
#                 page_num = idx + 1
#                 page = pdf.pages[idx]
#                 txt = cached_page_text(page) or ""
#                 if not txt.strip():
#                     continue

#                 # NO TOC skipping here
#                 if not _auditor_match_in_text(firm_name, txt):
#                     continue

#                 ws_validation.cell(row=row_index, column=9).value = "PASS"
#                 proof_file = PROOF_DIR_SIGNATURE / f"Row{row_index}_SIGNATURE_PASS_Page{page_num}.png"
#                 im = page.to_image(resolution=PROOF_IMAGE_RESOLUTION)

#                 boxes = []
#                 try:
#                     box = _find_auditor_name_box_on_page(page, firm_name)
#                     if box:
#                         im.draw_rect(box, stroke=HIGHLIGHT_STROKE, fill=HIGHLIGHT_FILL_RGBA)
#                         boxes.append(box)
#                 except Exception:
#                     pass

#                 im.save(proof_file)

#                 if SHOW_PROOF_POPUP:
#                     show_proof_popup(proof_file, f"Row {row_index} | Signature = PASS | Page {page_num}", seconds=PROOF_POPUP_SECONDS)

#                 return

#             # Fallback: full PDF search (keep existing behavior)
#             hit = find_auditor_firm_anywhere_in_pdf(pdf, firm_name)
#             if hit:
#                 page_num, box = hit
#                 ws_validation.cell(row=row_index, column=9).value = "PASS"

#                 page = pdf.pages[page_num - 1]
#                 proof_file = PROOF_DIR_SIGNATURE / f"Row{row_index}_SIGNATURE_PASS_Fallback_Page{page_num}.png"
#                 im = page.to_image(resolution=PROOF_IMAGE_RESOLUTION)

#                 if box:
#                     im.draw_rect(box, stroke=HIGHLIGHT_STROKE, fill=HIGHLIGHT_FILL_RGBA)

#                 im.save(proof_file)

#                 if SHOW_PROOF_POPUP:
#                     show_proof_popup(proof_file, f"Row {row_index} | Signature = PASS (Fallback) | Page {page_num}", seconds=PROOF_POPUP_SECONDS)

#                 return

#             # FAIL
#             ws_validation.cell(row=row_index, column=9).value = "FAIL"
#             return

#     except Exception:
#         ws_validation.cell(row=row_index, column=9).value = "FAIL"
#         return
       
# VALIDATION: Find an FY date on a page without context (used for statement headers).
# No Changes
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
# No Changes
def is_excluded_fin_stmt_page(page_text: str) -> bool:
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

# VALIDATION: Locate the "Statement of Net Position" heading near the top of a page.
# No Changes
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
# No Changes
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
# No Changes
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
# No Changes
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
# No Changes
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

   
# VALIDATION: Locate Balance Sheet heading near top of page (supports common variants).
# No Changes
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
# No Changes
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
# No Changes
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

# VALIDATION: Strictly detect Balance Sheet pages based on heading patterns.
# No Changes
def is_balance_sheet_page(page_text: str) -> bool:
    """
    STRICT detector for Balance Sheet (Governmental Funds style).

    Rules:
      - Must NOT be excluded (TOC/MD&A/Notes/Reconciliation etc.)
      - Must contain 'balance sheet'
      - Must contain 'assets' and 'liabilities'
      - Should contain 'fund balances' OR 'fund balance'
      - Prefer 'governmental funds' (helps avoid proprietary balance sheets)
      - Exclude 'reconciliation' and (optionally) 'nonmajor'
    """
    if not page_text:
        return False

    if is_excluded_fin_stmt_page(page_text):
        return False

    t_all = normalize_pdf_text(page_text).lower()

    if "reconciliation" in t_all:
        return False

    # Optional: avoid Nonmajor pages (like reference code)
    if "nonmajor" in t_all:
        return False

    if "balance sheet" not in t_all:
        return False

    if "assets" not in t_all or "liabilities" not in t_all:
        return False

    if ("fund balances" not in t_all) and ("fund balance" not in t_all):
        return False

    # Strong signal for the correct balance sheet type
    if "governmental funds" not in t_all:
        # Still allow if fund balances exist (some PDFs omit the phrase)
        pass

    return True
       
# VALIDATION: Locate the Revenues/Expenditures/Fund Balances statement heading near the top.
# No Changes
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
# No Changes
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
# No Changes
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


# ---------------------------------------------------------
# PERFORMANCE: cache pdfplumber extract_words per page
# ---------------------------------------------------------
_PAGE_WORDS_CACHE = {}

# VALIDATION: Cache extracted PDF words for a page to speed up repeated validations.
# No Changes
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
# No Changes
def _clear_page_words_cache():
    """
    Clear cached words between validations to avoid memory growth.
    """
    _PAGE_WORDS_CACHE.clear()

# VALIDATION: Collect words from the top visual lines of a PDF page (works for rotated/landscape pages).
# No Changes
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
# No Changes
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
# No Changes
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
# No Changes
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

# VALIDATION: Strictly detect Statement of Net Position pages.
# No Changes
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

 
# =========================================================
# EXCEL COLUMN DISCOVERY
# =========================================================
# VALIDATION: Create a unique, readable file tag used to name downloaded FAC files.
# No Changes (Fac Naming convention)
def build_download_key(name, ay, auditee_city, auditee_state, entity_type, auditee_uei, auditee_ein) -> str:
    """
    Naming convention similar to attached PublicFinance.py:
    sanitize each part and join with underscore. 【1-07881c】
    """
    parts = [
        sanitize_filename(name or "", 80),
        sanitize_filename(str(ay) or "", 10),
        sanitize_filename(auditee_city or "", 40),
        sanitize_filename(auditee_state or "", 10),
        sanitize_filename(entity_type or "", 25),
        sanitize_filename(auditee_uei or "", 25),
        sanitize_filename(auditee_ein or "", 25),
    ]
    return "_".join([p if p else "UNKNOWN" for p in parts])




# =========================================================
# PDF VALIDATION HELPERS (Detailed Validation Check)
# =========================================================

# VALIDATION: Extract needed values (Report ID, FYE, etc.) from the selected General row.
def extract_general_fields(ws_general, gen_cols: Dict[str, int], row_index: int) -> Dict[str, Any]:
    def gv(key: str):
        c = gen_cols.get(key)
        return get_cell(ws_general, row_index, c) if c else None

    return {
        "report_id": gv("report_id"),
        "fy_end_date": gv("fy_end_date"),
        "submitted_date": gv("submitted_date"),
        "auditee_certified_date": gv("auditee_certified_date"),
        "auditee_city": gv("auditee_city"),
        "auditee_state": gv("auditee_state"),
        "auditee_contact_name": gv("auditee_contact_name"),
        "auditee_contact_title": gv("auditee_contact_title"),
        "auditee_phone": gv("auditee_phone"),
        "auditee_email": gv("auditee_email"),
        "auditor_firm_name": gv("auditor_firm_name"),
        "auditor_contact_name": gv("auditor_contact_name"),
        "auditor_phone": gv("auditor_phone"),
        "auditor_email": gv("auditor_email"),
        "auditee_uei": gv("auditee_uei"),
        "auditee_ein": gv("auditee_ein"),
        "entity_type": gv("entity_type"),
    }


# VALIDATION: Find the output columns in MASTER LIST where extracted values will be written.
# now will be written in the db 
# def locate_master_output_columns(ws_master) -> Dict[str, int]:
#     m = {}
#     m["fye_date"] = find_col(ws_master, ["FYE DATE"], 1)
#     m["date_of_release"] = find_col(ws_master, ["DATE OF RELEASE"], 1)
#     m["date_of_download"] = find_col(ws_master, ["DATE OF DOWNLOAD"], 1)
#     m["auditor_signature_date"] = find_col(ws_master, ["AUDITOR'S SIGNATURE DATE"], 1)

#     m["auditee_city"] = find_col(ws_master, ["ENTITY'S CITY"], 1)
#     m["auditee_state"] = find_col(ws_master, ["ENTITY'S STATE"], 1)

#     m["auditee_contact_name"] = find_col(ws_master, ["ENTITY CONTACT: NAME"], 1)
#     m["auditee_contact_title"] = find_col(ws_master, ["ENTITY CONTACT: TITLE"], 1)
#     m["auditee_phone"] = find_col(ws_master, ["ENTITY CONTACT: PHONE"], 1)
#     m["auditee_email"] = find_col(ws_master, ["ENTITY CONTACT: EMAIL ID"], 1)

#     m["auditor_firm_name"] = find_col(ws_master, ["AUDITOR'S FARM NAME", "AUDITOR'S FIRM NAME"], 1)
#     m["auditor_name"] = find_col(ws_master, ["AUDITOR'S NAME"], 1)
#     m["auditor_phone"] = find_col(ws_master, ["AUDITOR'S CONTACT: PHONE"], 1)
#     m["auditor_email"] = find_col(ws_master, ["AUDITOR'S CONTACT: EMAIL ID"], 1)

#     m["pdf_link"] = find_col(ws_master, ["PDF LINK"], 1)
#     return m

# If no PDF button is found
# !!! Done Update DB using conn instead of passing ws_master 
def mark_master_row_pdf_not_found(db, row_id: int, processing_id: int):
    """
    Set all output columns to 'PDF Not Found' for this row.
    Replaces: writing PDF_NOT_FOUND_TEXT to all master_cols in MASTER LIST.

    row_id        -> keys the single physical TProcessStatus row (new PK).
    processing_id -> keys TProcessingAdditionalInfo (kept on ProcessingId).
    """
    db.update("""
        UPDATE TProcessStatus
        SET
            SourcingStatus          = 0,
            SourcingFlag            = 'c',
            SourcingValidationStatus= NULL,
            SourcingValidationFlag  = NULL,
            ExtractionStatus        = NULL,
            pdf_download_link       = 'PDF Not Found',
            ExtractionFlag          = NULL,
            CompletionStatus        = 0,
            PdfFilePath             = NULL,
            OutputPath              = NULL,
            FyeDate                 = NULL,
            ReleaseDate             = NULL,
            DownloadDate            = NULL,
            ModifiedOn              = GETDATE()
        WHERE Id = ?
    """, row_id, commit=False)

    db.update("""
        UPDATE TProcessingAdditionalInfo
        SET
            AuditorsSignatureDate = NULL,
            EntityCity = NULL,
            EntityState = NULL,
            EntityContactName   = NULL,
            EntityContactTitle  = NULL,
            EntityContactPhone  = NULL,
            EntityContactEmail  = NULL,
            AuditorFirmName     = NULL,
            AuditorName         = NULL,
            AuditorContactPhone = NULL,
            AuditorContactEmail = NULL,
            ModifiedOn          = GETDATE()
        WHERE ProcessingId = ?
    """, processing_id, commit=False)

    db.commit()



#!!! Change it to parquet
def mark_all_validations_failed(ws_validation, row_index: int, issuer_name: str, uei: str):
    """
    Force all validations to FAIL in Detailed Validation Check for this row (and write Issuer Name + UEI).
    Assumes columns A..M => A=Name, B=UEI, C..M=validations.
    """
    # Write issuer name & UEI (A,B)
    ws_validation.cell(row=row_index, column=1).value = issuer_name
    ws_validation.cell(row=row_index, column=2).value = uei

    # Force FAIL for all validation columns C..M (3..13)
    for col in range(3, 14):
        ws_validation.cell(row=row_index, column=col).value = "Fail"

# VALIDATION: Set the Audit Year filter on the FAC Advanced Search page.
# No Changes
def set_audit_year_only(page, year: str):
    ensure_accordion_open(page, "audit-year")
    year = normalize(year)
    if not year:
        raise ValueError("Audit Year is blank in MASTER LIST row")

    page.evaluate(
        """
        (targetYear) => {
          const boxes = Array.from(document.querySelectorAll("input[name='audit_year'][type='checkbox']"));
          boxes.forEach(b => {
            b.checked = false;
            b.dispatchEvent(new Event('input', {bubbles:true}));
            b.dispatchEvent(new Event('change', {bubbles:true}));
          });

          const match = boxes.find(b => String(b.value) === String(targetYear));
          if (match){
            match.checked = true;
            match.dispatchEvent(new Event('input', {bubbles:true}));
            match.dispatchEvent(new Event('change', {bubbles:true}));
          }
        }
        """,
        year
    )
    page.wait_for_timeout(200)


# VALIDATION: Set the State filter on the FAC Advanced Search page.
# No Changes
def set_state(page, state_val: str):
    ensure_accordion_open(page, "state")
    abbr = to_state_abbr(state_val)
    if not abbr:
        raise ValueError("State is blank in MASTER LIST row")

    sel = page.locator("#auditee_state")
    sel.wait_for(state="visible", timeout=UI_TIMEOUT_MS)
    sel.select_option(value=abbr)
    page.wait_for_timeout(200)


# VALIDATION: Write extracted values from FAC Excel into the MASTER LIST row and set download dates.
# !!! Done instead of writing in the ws_master, pass the db connection and update the values from g
def update_master_from_general(db, row_id: int, processing_id: int, g: Dict[str, Any], pdf_url: Optional[str]):
    today = date.today()

    # --------------------------------------------------
    # UPDATE TProcessStatus
    # Columns: FyeDate, ReleaseDate, DownloadDate, pdf_download_link
    # Keyed on the physical row PK (Id).
    # --------------------------------------------------
    db.update("""
        UPDATE TProcessStatus
        SET
            FyeDate          = ?,
            ReleaseDate      = ?,
            DownloadDate     = ?,
            pdf_download_link= ?,
            ModifiedOn       = GETDATE()
        WHERE Id = ?
    """, [
        g.get("fy_end_date"),
        g.get("submitted_date"),
        today,
        pdf_url if pdf_url else None,
        row_id,
    ], commit=False)

    # --------------------------------------------------
    # UPDATE TProcessingAdditionalInfo
    # Columns: AuditorsSignatureDate, EntityContact*, Auditor*
    # Kept keyed on ProcessingId (this table is not part of the Id re-key).
    # --------------------------------------------------
    db.update("""
        UPDATE TProcessingAdditionalInfo
        SET
            AuditorsSignatureDate = ?,
            EntityCity = ?,
            EntityState = ?,
            EntityContactName     = ?,
            EntityContactTitle    = ?,
            EntityContactPhone    = ?,
            EntityContactEmail    = ?,
            AuditorFirmName       = ?,
            AuditorName           = ?,
            AuditorContactPhone   = ?,
            AuditorContactEmail   = ?,
            ModifiedOn            = GETDATE()
        WHERE ProcessingId = ?
    """, [
        g.get("auditee_certified_date"),
        g.get("auditee_city"),
        g.get("auditee_state"),
        g.get("auditee_contact_name"),
        g.get("auditee_contact_title"),
        g.get("auditee_phone"),
        g.get("auditee_email"),
        g.get("auditor_firm_name"),
        g.get("auditor_contact_name"),
        g.get("auditor_phone"),
        g.get("auditor_email"),
        processing_id,
    ], commit=False)

    db.commit()


# =========================================================
# OUTPUT FORMATTING (BORDERS / FILLS / FILTERS / OPEN FILE)
# =========================================================

# OUTPUT-FORMATTING: Delete old proof images/files inside a folder before starting a fresh run.
# No Changes
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


# OUTPUT-FORMATTING: Clear cell values in a sheet for a specific row/column range.
def clear_sheet_columns(ws, start_row: int, start_col: int, end_col: int) -> None:
    """Clear values in ws from start_row to ws.max_row for given column range."""
    max_r = ws.max_row
    for r in range(start_row, max_r + 1):
        for c in range(start_col, end_col + 1):
            ws.cell(row=r, column=c).value = None



# =========================================================
# ENTRYPOINT
# =========================================================
# OUTPUT-FORMATTING: Main workflow: search FAC for each row, download Excel/PDF, validate PDFs, and update workbook.

# =========================================================
# DOWNLOAD FAILURE HANDLING (Excel download failed after all retries)
# =========================================================

FAILED_TO_DOWNLOAD_TEXT = "Failed to Download"

# Yellow fill to clearly flag download failures in Excel
# FILL_DOWNLOAD_FAIL = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")


# !!! (DONE)Update DB and set failed to download for the requierd columns
def mark_master_list_ix_failed_to_download(db, row_id: int, processing_id: int):
    """
    Mark a row as 'Failed to Download' in the DB.
    Replaces: writing 'Failed to Download' to MASTER LIST columns I..X (9..24)

    Updates:
      - TProcessStatus           → sets all status/flag/path columns + CompletionStatus=1
      - TProcessingAdditionalInfo → clears all contact/auditor fields

    row_id        -> keys the single physical TProcessStatus row (new PK).
    processing_id -> keys TProcessingAdditionalInfo (kept on ProcessingId).
    """
    db.update("""
        UPDATE TProcessStatus
        SET
            SourcingStatus          = 0,
            SourcingFlag            = 'c',
            SourcingValidationStatus= NULL,
            SourcingValidationFlag  = NULL,
            ExtractionStatus        = NULL,
            pdf_download_link       = 'Failed to Download',
            ExtractionFlag          = NULL,
            CompletionStatus        = 0,
            PdfFilePath             = NULL,
            OutputPath              = NULL,
            FyeDate                 = NULL,
            ReleaseDate             = NULL,
            DownloadDate            = NULL,
            ModifiedOn              = GETDATE()
        WHERE Id = ?
    """, row_id, commit=False)

    db.update("""
        UPDATE TProcessingAdditionalInfo
        SET
            AuditorsSignatureDate = NULL,
            EntityCity = NULL,
            EntityState = NULL,
            EntityContactName   = NULL,
            EntityContactTitle  = NULL,
            EntityContactPhone  = NULL,
            EntityContactEmail  = NULL,
            AuditorFirmName     = NULL,
            AuditorName         = NULL,
            AuditorContactPhone = NULL,
            AuditorContactEmail = NULL,
            ModifiedOn          = GETDATE()
        WHERE ProcessingId = ?
    """, processing_id, commit=False)

    db.commit()


# !!! Need to convert it indo Parquet, all the rows from the Detailed Validation Check sheet
def mark_validation_failed_to_download(ws_validation, row_index: int, issuer_name: str, uei: str,
                                      start_col: int = 3, end_col: int = 13):
    """
    Detailed Validation Check sheet handling:
    - A/B written normally (Issuer Name, UEI)
    - C..M marked as 'Failed to Download' and filled yellow
    """
    # A = Issuer Name, B = UEI
    ws_validation.cell(row=row_index, column=1).value = issuer_name
    ws_validation.cell(row=row_index, column=2).value = uei

    # C..M = validations
    for c in range(start_col, end_col + 1):
        cell = ws_validation.cell(row=row_index, column=c)
        cell.value = FAILED_TO_DOWNLOAD_TEXT


# =========================================================
# DB Status & Flag update
# =========================================================
def update_sourcing_status(db, row_id, status, flag, pdf_path=None):
    """
    status: None=pending, 0=fail, 1=success
    flag:   None=not started, s=stared, p=processing, c=competed

    Keyed on the physical TProcessStatus.Id (row_id).
    """
    if pdf_path is not None:
        res = db.update("""
            UPDATE TProcessStatus
            SET SourcingStatus = ?, SourcingFlag = ?, PdfFilePath = ?, ModifiedOn = GETDATE()
            WHERE Id = ?
        """, [status, flag, str(pdf_path), row_id])
    else:
        res = db.update("""
            UPDATE TProcessStatus
            SET SourcingStatus = ?, SourcingFlag = ?, ModifiedOn = GETDATE()
            WHERE Id = ?
        """, [status, flag, row_id])
    if not res.success:
        print(f"[FAILED] row_id {row_id}: {res.error}")

# =========================================================
# ENTRYPOINT
# =========================================================

# ENTRYPOINT: Run the full program: load workbook, perform sourcing + validation, save, and open the output.

def run_fac_for_rows_source_only(rows: List[dict], db):

    with sync_playwright() as p:
        browser = p.chromium.launch(
            headless=HEADLESS,
            slow_mo=SLOW_MO_MS,
            channel=CHROME_CHANNEL,
            args=["--start-maximized"]
        )

        # ✅ viewport=None lets the browser use the full window size
        context = browser.new_context(
            accept_downloads=True,
            ignore_https_errors=IGNORE_HTTPS_ERRORS,
            viewport=None
        )

        page = context.new_page()
        page.set_default_timeout(UI_TIMEOUT_MS)

        # ✅ Make the main Chrome page full screen/maximized
        maximize_playwright_page(page)

        for item in rows:
            r = item["processing_id"]
            year = item["year"]
            uei = item["uei"]
            state = item["state"]
            sector = item["sector"]
            # company_id = item["company_id"]
            processing_code = item["processing_code"]
            processing_id = item["processing_id"]
            row_id = item["row_id"]            # physical TProcessStatus.Id — DB row key
            # issuer_name = item["issuer_name"]
            # sub_sector = item["sub_sector"]
            # ein = item["ein"]

            print(f"\n--- Processing id: {processing_id} | Year={year} UEI={uei} State={state} Sector={sector} ---")

            # SOURCING START → in progress, result pending (signal for the frontend)
            update_sourcing_status(db, row_id, status=None, flag='p')

            try:
                safe_goto(page, ADV_URL)
                set_audit_year_only(page, year)
                fill_uei(page, uei)
                set_state(page, state)
                set_entity_type(page, sector)
                click_search(page)

                tag = f"{processing_code}_FAC"
                try:
                    downloaded_xlsx = download_all_excel_with_retry(page, DOWNLOAD_DIR, tag)
                except RuntimeError as e:
                    # If Excel download fails even after all retries (default: 5 attempts),
                    # mark MASTER LIST I..X and validation C..M as 'Failed to Download' (yellow), then move to next row.
                    if "Excel download failed after" in str(e):
                        mark_master_list_ix_failed_to_download(db, row_id, processing_id)#!!!done
                        print(f"[FAILED] processing_id {r}: Excel download failed after retries => Marked as '{FAILED_TO_DOWNLOAD_TEXT}'")
                        continue
                    raise
                print(f"[SUCCESS] Excel downloaded: {downloaded_xlsx}")

                criteria = {"year": year, "uei": uei, "state": state, "sector": sector}

                pdf_url, pdf_saved, fy_end_val = process_downloaded_excel_and_update_master(
                    downloaded_xlsx, db, row_id, processing_id, processing_code, criteria, context
                )

                # -----------------------------
                # ✅ MARK SOURCING COMPLETE
                # -----------------------------
                # decide outcome
                pdf_ok = (pdf_url and pdf_url != PDF_NOT_FOUND_TEXT
                        and pdf_saved and Path(pdf_saved).exists())

                if pdf_ok:
                    update_sourcing_status(db, row_id, status=1, flag='c', pdf_path=pdf_saved)  # success
                    print(f"[SUCCESS] PDF saved: {pdf_saved}")
                else:
                    update_sourcing_status(db, row_id, status=0, flag='c')   # done + fail
                    print(f"[INFO] {processing_id}: PDF not found/saved => marked FAIL.")

                # -----------------------------
                # PERFORMANCE: run all validations at once (single PDF open + cached extraction)
                page.wait_for_timeout(500)


            except PlaywrightTimeoutError as te:
                print(f"[TIMEOUT] processing_id {r}: {te}")
                update_sourcing_status(db, row_id, status=0, flag='c')   # done + fail
            except Exception as e:
                print(f"[FAILED] processing_id {r}: {e}")
                update_sourcing_status(db, row_id, status=0, flag='c')   # done + fail

        try:
            context.close()
        except Exception:
            pass
        try:
            browser.close()
        except Exception:
            pass


# =========================================================
# PRE-RUN CLEANUP HELPERS
# =========================================================

# =========================================================
# ENTRYPOINT
# =========================================================

def main_source():
    """Run ONLY sourcing: download Excel+PDF and update DB. No PDF validations."""

    # --------------------------------------------------
    # 1. CONNECT TO DATABASE (replaces load_workbook)
    #    Uses the shared db.py layer as a context manager: commits on clean
    #    exit, rolls back on exception, and closes the connection.
    # --------------------------------------------------
    with Database() as db:
        print("[DB] Connected to database successfully.")

        # --------------------------------------------------
        # 2. FETCH MASTER ROWS (replaces ws_master = wb[SHEET_MASTER])
        #
        #    This JOIN gives us every row that still needs sourcing.
        #    SourcingStatus = 0  → only pending/un-sourced rows
        #    cm.ModuleId = 1     → Public Finance module
        #    ps.Id AS RowId      → the physical-row PK (unique) used for writes
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
                ps.ProcessYear
            FROM TCompanyMaster cm
            JOIN TProcessStatus ps
                ON ps.CompanyId = cm.CompanyId
            WHERE cm.IsActive   = 1
              AND ps.IsActive    = 1
              AND (ps.SourcingFlag = 'c' or ps.SourcingFlag is NULL)
              AND (ps.SourcingStatus = 0 OR ps.SourcingStatus is NULL)
              AND (ps.CompletionStatus=0 OR ps.CompletionStatus is NULL)
              AND ps.COAID IN (1, 2)
        """)
        if not res.success:
            print(f"[DB-ERROR] Failed to fetch master rows: {res.error}")
            return
        master_rows = res.data

        if not master_rows:
            print("[INFO] No pending rows to process. Exiting.")
            return

        # Claim the fetched physical rows by their PK (Id), not ProcessingId —
        # ProcessingId is non-unique so an IN(ProcessingId) list would also flip
        # sibling rows that share a deliverable.
        row_ids = [row.RowId for row in master_rows]
        clause, params = build_in_clause("Id", row_ids)
        claim = db.update(f"""
            UPDATE TProcessStatus
            SET SourcingFlag = 's'
            WHERE IsActive = 1
              AND {clause}
        """, params)
        print(f"[DB] Updated SourcingFlag to 's' for {claim.rowcount} rows.")

        # print(f"[DB] Found {len(master_rows)} pending rows in TCompanyMaster ⟕ TProcessStatus.")

        # --------------------------------------------------
        # 4. PRE-RUN CLEANUP  (directory clearing stays the same)
        #
        #    NOTE: clear_sheet_columns(ws_master, 9, 24) is NO LONGER NEEDED
        #    because the DB query already filters SourcingStatus = 0,
        #    so we only get un-sourced rows. Output columns start as NULL.
        # --------------------------------------------------
        print("[PRE-RUN] Clearing download folders...")
        clear_directory_contents(DOWNLOAD_DIR)
        clear_directory_contents(PDF_DIR)

        # --------------------------------------------------
        # 5. BUILD ROWS LIST FROM DB
        #    (replaces: find_col + for r in range(2, max_row) + get_cell)
        #
        #    Every field the FAC search needs is already in the DB row.
        #    No header discovery or column-index guessing required.
        # --------------------------------------------------
        rows = []
        for db_row in master_rows:
            year   = normalize(str(db_row.ProcessYear or ""))
            uei    = normalize(str(db_row.UEI or ""))
            state  = normalize(str(db_row.State or ""))
            sector = normalize(str(db_row.Sector or ""))

            if not year and not uei:
                continue

            rows.append({
                # --- DB identifiers (NEW — used for DB writes later) ---
                "row_id":          db_row.RowId,          # physical TProcessStatus.Id (unique) — DB row key
                "processing_id":   db_row.ProcessingId,   # deliverable grouping (non-unique) — display/label
                "company_id":      db_row.CompanyId,
                "processing_code": db_row.ProcessingCode,

                # --- FAC search inputs (same keys as before) ---
                "year":   year,
                "uei":    uei,
                "state":  state,
                "sector": sector,

                # --- Extra fields available from DB (no extra query needed) ---
                "sub_sector":   normalize(str(db_row.SubSector or "")),
                "issuer_name":  normalize(str(db_row.IssuerName or "")),
                "ein":          normalize(str(db_row.EIN or "")),
            })

        print(f"[INFO] {len(rows)} rows ready for FAC sourcing.")

        # master_output_cols = master sheet output columns -> DB TProcessingAdditionalInfo
        # run_fac_for_rows_source_only(rows, ws_master, ws_validation, issuer_name_col, master_output_cols, wb, xlsm_path)
        run_fac_for_rows_source_only(rows, db)

    # --------------------------------------------------
    # 7. FINALIZE
    #    The `with Database()` block above committed and closed the connection.
    # --------------------------------------------------
    print("[DONE] Sourcing completed.")
 
 
if __name__ == "__main__":
    main_source()
