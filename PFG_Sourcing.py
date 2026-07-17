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
import time
import shutil
import pandas as pd
import threading
from filelock import FileLock

# Shared, reusable DB layer (see db.py). All DB access goes through this now
# instead of a per-file get_db_connection() + raw pyodbc cursors.
from db import Database, build_in_clause


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
# CONFIG
# =========================================================
# FILE_PATH = r"C:\S2\Public Finance\PDF to be Sourced and Extracted.xlsm"

PDF_NOT_FOUND_TEXT = "PDF Not Found"

BASE = "https://app.fac.gov"
ADV_URL = "https://app.fac.gov/dissemination/search/advanced/"

DOWNLOAD_DIR = Path(r"C:\Test Code\Public Finance\01_Metadata_update")
DOWNLOAD_DIR.mkdir(parents=True, exist_ok=True)

PDF_DIR = Path(r"C:\Test Code\Public Finance\02_Downloaded_Report")
PDF_DIR.mkdir(parents=True, exist_ok=True)


# Retry settings (4 more retries => total 5 attempts)
DOWNLOAD_RETRIES = 4
RETRY_WAIT_SECONDS = 2
HEADLESS = False          # set True for final run
SLOW_MO_MS = 150          # set 0 for faster run
PAGE_LOAD_TIMEOUT_MS = 60000
UI_TIMEOUT_MS = 30000
DOWNLOAD_TIMEOUT_MS = 90000
PDF_DOWNLOAD_TIMEOUT_MS = 90000

CHROME_CHANNEL = "chrome"
IGNORE_HTTPS_ERRORS = True  # helps in SSL-inspected networks


# =========================================================
# SOURCING (FAC SEARCH + DOWNLOADS + EXCEL UPDATE)
# =========================================================


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
        set_remarks(db, row_id, "PDF not found: no Report ID / PDF URL in FAC data.")
        return PDF_NOT_FOUND_TEXT, None, None

    # pdf_saved_path = PDF_DIR / f"{download_key}.pdf"
    pdf_saved_path = PDF_DIR / f"{processing_code}_AR.pdf"

    try:
        download_pdf_via_browser_with_retry(playwright_context, pdf_url, pdf_saved_path)

        if not pdf_saved_path.exists() or pdf_saved_path.stat().st_size == 0:
            mark_master_row_pdf_not_found(db, row_id, processing_id)
            set_remarks(db, row_id, "PDF download failed: file missing or empty after retries.")
            return PDF_NOT_FOUND_TEXT, None, None

        return pdf_url, pdf_saved_path, g.get("fy_end_date")

    except RuntimeError as e:
        cause = getattr(e, "__cause__", None)
        if isinstance(cause, PlaywrightTimeoutError) or ("timeout" in str(cause).lower()):
            mark_master_row_pdf_not_found(db, row_id, processing_id)
            set_remarks(db, row_id, "PDF download timed out.")
            return PDF_NOT_FOUND_TEXT, None, None

        raise

    except PlaywrightTimeoutError:
        mark_master_row_pdf_not_found(db, row_id, processing_id)
        set_remarks(db, row_id, "PDF download timed out.")
        return PDF_NOT_FOUND_TEXT, None, None


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


# If no PDF button is found
# !!! Done Update DB using conn instead of passing ws_master
def mark_master_row_pdf_not_found(db, row_id: int, processing_id: int):
    """
    Set all output columns to 'PDF Not Found' for this row.
    Replaces: writing PDF_NOT_FOUND_TEXT to all master_cols in MASTER LIST.

    row_id        -> keys the single physical TProcessStatus row (PK) AND
                     TProcessingAdditionalInfo (its ID column FKs to TProcessStatus.Id).
    processing_id -> retained for display/logging only; no longer a DB key.
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
        WHERE ID = ?
    """, row_id, commit=False)

    db.commit()


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
    # Keyed on row_id via its ID column (FK -> TProcessStatus.Id).
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
        WHERE ID = ?
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
        row_id,
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

    row_id        -> keys the single physical TProcessStatus row (PK) AND
                     TProcessingAdditionalInfo (its ID column FKs to TProcessStatus.Id).
    processing_id -> retained for display/logging only; no longer a DB key.
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
        WHERE ID = ?
    """, row_id, commit=False)

    db.commit()


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
# DB Remarks update (TProcessStatus.Remarks)
# =========================================================
def set_remarks(db, row_id, remark: Optional[str]):
    """
    Set TProcessStatus.Remarks for one physical row (keyed on Id).

    - Call with remark=None at the START of a run to CLEAR any stale note,
      so a fresh attempt never shows a message from a previous run.
    - Call with a short informative string when something goes wrong
      (exception, FAC Excel not downloaded, PDF not downloaded, etc.).

    Never raises: a remark write must not break the sourcing flow.
    """
    text = None
    if remark is not None:
        text = str(remark).strip()
        # Keep it a *small* informative note.
        if len(text) > 500:
            text = text[:497] + "..."

    try:
        res = db.update("""
            UPDATE TProcessStatus
            SET Remarks = ?, ModifiedOn = GETDATE()
            WHERE Id = ?
        """, [text, row_id])
        if not res.success:
            print(f"[REMARKS-ERROR] row_id {row_id}: {res.error}")
    except Exception as e:
        print(f"[REMARKS-ERROR] row_id {row_id}: {e}")


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

            # Fresh run for this id: clear any stale Remarks from a previous attempt.
            set_remarks(db, row_id, None)

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
                        set_remarks(db, row_id, f"FAC Excel download failed after {DOWNLOAD_RETRIES + 1} attempts.")
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
                set_remarks(db, row_id, f"Timed out during FAC sourcing: {te}")
                update_sourcing_status(db, row_id, status=0, flag='c')   # done + fail
            except Exception as e:
                print(f"[FAILED] processing_id {r}: {e}")
                set_remarks(db, row_id, f"Error during sourcing: {e}")
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
# DB ROW -> WORK ITEM (shared by batch + single-id paths)
# =========================================================

# One SELECT list, reused by both the batch fetch (main_source) and the
# single-id fetch (source_one_id) so every path returns an identical row shape.
_SOURCING_SELECT = """
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
"""


def _build_row_item(db_row) -> Optional[dict]:
    """
    Convert one joined TCompanyMaster x TProcessStatus row into the work-item
    dict that run_fac_for_rows_source_only() consumes.

    Returns None if the row has neither Year nor UEI (nothing to search on).
    """
    year   = normalize(str(db_row.ProcessYear or ""))
    uei    = normalize(str(db_row.UEI or ""))
    state  = normalize(str(db_row.State or ""))
    sector = normalize(str(db_row.Sector or ""))

    if not year and not uei:
        return None

    return {
        # --- DB identifiers (used for DB writes later) ---
        "row_id":          db_row.RowId,          # physical TProcessStatus.Id (unique) — DB row key
        "processing_id":   db_row.ProcessingId,   # deliverable grouping (non-unique) — display/label
        "company_id":      db_row.CompanyId,
        "processing_code": db_row.ProcessingCode,

        # --- FAC search inputs ---
        "year":   year,
        "uei":    uei,
        "state":  state,
        "sector": sector,

        # --- Extra fields available from DB (no extra query needed) ---
        "sub_sector":   normalize(str(db_row.SubSector or "")),
        "issuer_name":  normalize(str(db_row.IssuerName or "")),
        "ein":          normalize(str(db_row.EIN or "")),
    }


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
        res = db.fetch_all(_SOURCING_SELECT + """
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
            item = _build_row_item(db_row)
            if item is not None:
                rows.append(item)

        print(f"[INFO] {len(rows)} rows ready for FAC sourcing.")

        # master_output_cols = master sheet output columns -> DB TProcessingAdditionalInfo
        # run_fac_for_rows_source_only(rows, ws_master, ws_validation, issuer_name_col, master_output_cols, wb, xlsm_path)
        run_fac_for_rows_source_only(rows, db)

    # --------------------------------------------------
    # 7. FINALIZE
    #    The `with Database()` block above committed and closed the connection.
    # --------------------------------------------------
    print("[DONE] Sourcing completed.")


# =========================================================
# ENTRYPOINT (SINGLE ID)
# =========================================================

def source_one_id(db, row_id: int) -> bool:
    """
    Run the SAME sourcing process as main_source(), but for a SINGLE
    TProcessStatus row identified by its physical Id (ps.Id), using an
    already-open Database handle `db`.

    Use this when the caller already owns a Database connection (e.g. an
    external dispatcher). For a self-contained one-shot call, use
    main_source_by_id(row_id) instead.

    Returns True if the row was dispatched to FAC sourcing, False if it
    could not be found / had no searchable Year or UEI.

    NOTE: unlike the batch main_source(), this does NOT wipe DOWNLOAD_DIR /
    PDF_DIR. Those are shared folders and files are named per ProcessingCode,
    so a targeted single-id run must not delete other rows' outputs.
    """
    # Fetch exactly this physical row (by Id). We intentionally filter only on
    # IsActive — when a caller explicitly passes an Id they want THAT row
    # sourced, regardless of its current Sourcing/Completion flags.
    res = db.fetch_one(_SOURCING_SELECT + """
        WHERE ps.Id = ?
          AND cm.IsActive = 1
          AND ps.IsActive = 1
    """, [row_id])

    if not res.success:
        print(f"[DB-ERROR] Failed to fetch row Id={row_id}: {res.error}")
        return False
    if res.data is None:
        print(f"[INFO] No active TProcessStatus row found for Id={row_id}.")
        return False

    item = _build_row_item(res.data)
    if item is None:
        print(f"[INFO] Row Id={row_id} has no Year/UEI — nothing to search on.")
        return False

    # Claim this single physical row (SourcingFlag='s'), same as the batch path.
    claim = db.update("""
        UPDATE TProcessStatus
        SET SourcingFlag = 's'
        WHERE IsActive = 1 AND Id = ?
    """, [row_id])
    print(f"[DB] Claimed row_id={row_id} (SourcingFlag='s'), rows={claim.rowcount}.")

    # Reuse the exact same pipeline as the batch run (single-element list).
    run_fac_for_rows_source_only([item], db)
    return True


def main_source_by_id(row_id: int):
    """
    Source ONE TProcessStatus row by its Id (ps.Id).

    Opens its own DB connection (mirrors main_source) and runs the same
    sourcing pipeline for just that row:

        main_source_by_id(12345)
    """
    with Database() as db:
        print(f"[DB] Connected. Sourcing single row_id={row_id}.")
        source_one_id(db, row_id)
    print(f"[DONE] Sourcing completed for row_id={row_id}.")


if __name__ == "__main__":
    import sys

    # Usage:
    #   python PFG_Sourcing.py          -> source ALL pending rows (batch)
    #   python PFG_Sourcing.py <id>     -> source ONE TProcessStatus row by its Id
    if len(sys.argv) > 1:
        main_source_by_id(int(sys.argv[1]))
    else:
        main_source()
