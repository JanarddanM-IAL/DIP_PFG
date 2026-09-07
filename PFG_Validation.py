# -*- coding: utf-8 -*-
"""
PFG_Validation.py — DB-integrated audit-report validation runner.

Polls TProcessStatus for rows pending sourcing-validation, opens each
already-downloaded PDF, and runs the gate-based validation engine
(Gate 1 issuer details, Gate 2 audit compliance, Gate 3 financial-statement
structure, plus framework/basis detection and OCR detection) that was lifted
verbatim from the FAAC "US PUBLIC FINANCE AUDIT SOURCING & VALIDATION" tool.
Results are written back to the DB (SourcingValidationStatus / SourcingValidationFlag
/ Remarks / PdfFilePath) and to a parquet result log.

No Excel, no sourcing, no upload: the DB is the sole work queue and status sink,
and the parquet file is the detailed per-gate result log.

Run modes (mirrors PFG_Sourcing.py):
    python PFG_Validation.py           # BATCH  — validate every pending row
    python PFG_Validation.py <id>      # SINGLE — validate one TProcessStatus.Id

Sector rule: TProcessStatus.COAID -> TCOAMaster.SegmentId; SegmentId 1 = LG,
SegmentId 2 = NONLG (default LG when unknown).
"""

# ---------- standard library ----------
import re, time, sys, shutil, warnings, unicodedata
import traceback as _traceback
from pathlib import Path
from datetime import datetime, date
from dataclasses import dataclass
from typing import Optional, List, Dict, Tuple, Any

# ---------- third-party ----------
import pandas as pd
import threading
from filelock import FileLock          # cross-process lock around the parquet log
import fitz                            # PyMuPDF — PDF parsing engine
from rapidfuzz import fuzz             # fuzzy string matching for titles/names
from dateutil import parser as dtparser

# ---------- project DB layer (db.py) ----------
from db import Database, build_in_clause

# ---------- user-facing display log (optional; never breaks a run) ----------
# Short, human-readable "what happened to this report?" log (see
# user_display_log_pfg.PFG_LOG_DIR). A missing module degrades every udl.* call
# to a safe no-op so validation is never affected.
try:
    import user_display_log_pfg as udl
except Exception:
    class _UDLNoop:
        def __getattr__(self, _):
            return lambda *a, **k: None
    udl = _UDLNoop()

# Display allow-list: validation is one 'sv' sub-stage with a per-gate check row.
_DISPLAY_STAGES = [
    ("Validation", "sv", "Sourcing validation"),
]

# Human labels for the gate step codes shown in the display log (unknown codes
# fall back to the raw code).
_CHECK_LABELS = {
    "G1-1": "Entity Name", "G1-2": "State", "G1-3": "FYE",
    "G2-1": "Audit Opinion", "G2-2": "FYE in Opinion",
    "G2-3": "Auditor Name", "G2-4": "Auditor Signature",
    "G3-1": "Cash Receipts & Disbursements", "G3-2": "Net Position / Balance Sheet",
    "G3-3": "Governmental Funds Balance Sheet", "G3-4": "Statement of Activities",
    "G3-5": "Revenue, Expenditure & Fund Balances", "G3-6": "Cash Flows",
}

warnings.filterwarnings("ignore")

# The lifted engine prints emoji/arrows in its diagnostics. When stdout is a
# non-UTF-8 pipe/console (cp1252), those raise UnicodeEncodeError. Reconfigure to
# UTF-8 with replacement so console output can never crash a gate.
for _stream in (sys.stdout, sys.stderr):
    try:
        _stream.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass


# ---------- tqdm shim ----------
# The lifted engine helpers emit progress via tqdm.write(...). We don't want the
# tqdm dependency (or its progress bars) in this headless runner, so route
# tqdm.write to print(). The __call__ fallback makes `tqdm(iterable)` degrade to
# the bare iterable in case any lifted helper iterates it. write() is defensive
# against any residual encoding issue so a diagnostic line can never crash a gate.
class _TqdmShim:
    @staticmethod
    def write(*args, **kwargs):
        try:
            print(*args, **kwargs)
        except UnicodeEncodeError:
            safe = [str(a).encode("ascii", "replace").decode("ascii") for a in args]
            print(*safe, **kwargs)

    def __call__(self, iterable=None, *args, **kwargs):
        return iterable if iterable is not None else []

tqdm = _TqdmShim()


# =========================================================
# OUTPUT DIRECTORIES  (verdict routing)
# =========================================================
PDF_DIR                   = Path(r"D:\S2\Public Finance\02_Downloaded_Report")
VALIDATED_PDF_DIR         = Path(r"D:\S2\Public Finance\03_Validated_Report")
REVIEW_PDF_DIR            = Path(r"D:\S2\Public Finance\04_Review_Required")
FAILED_VALIDATION_PDF_DIR = Path(r"D:\S2\Public Finance\05_Manual_validation_required")
OCR_REQUIRED_PDF_DIR      = Path(r"D:\S2\Public Finance\06_OCR_Required")
PROOF_ROOT                = Path(r"D:\S2\Public Finance\999_Log_Trackers\Validation Proofs")

for _d in (PDF_DIR, VALIDATED_PDF_DIR, REVIEW_PDF_DIR,
           FAILED_VALIDATION_PDF_DIR, OCR_REQUIRED_PDF_DIR, PROOF_ROOT):
    _d.mkdir(parents=True, exist_ok=True)


# =========================================================
# VALIDATION PARQUET RESULT LOG  (new gate-based schema)
# =========================================================
# NOTE: new filename (v2) — the gate-based columns are incompatible with the old
# pdfplumber-era validation_results.parquet, so we do not reuse that file.
VALIDATION_PARQUET_PATH = Path(r"D:\S2\Public Finance\999_Log_Trackers\validation_results_v2.parquet")
VALIDATION_LOCK_PATH    = VALIDATION_PARQUET_PATH.with_suffix(".parquet.lock")

# Concurrency primitives (same pattern the sourcing/validation tools use).
_VALIDATION_THREAD_LOCK = threading.Lock()
_VALIDATION_FILE_LOCK   = FileLock(str(VALIDATION_LOCK_PATH), timeout=60)

# "Id" (physical TProcessStatus.Id, unique key) + "ProcessingID" (deliverable
# grouping, non-unique) + the gate-based result columns.
VALIDATION_COLUMNS = [
    "Id", "ProcessingID", "ISSUER NAME", "UEI",
    "OVERALL", "GATE1", "GATE2", "GATE3", "OCR_REQUIRED",
    "FRAMEWORK", "FRAMEWORK_CONFIDENCE",
    "G1-1", "G1-2", "G1-3",
    "G2-1", "G2-2", "G2-3", "G2-4",
    "G3-1", "G3-2", "G3-3", "G3-4", "G3-5", "G3-6",
]


# =========================================================
# PARQUET RESULT-LOG HELPERS (thread + process safe) — reused from prior runner
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

def normalize(v) -> str:
    """Return a trimmed string for any value (``None`` becomes an empty string)."""
    return "" if v is None else str(v).strip()

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
# DB STATUS WRITER
# =========================================================
def update_sourcing_validation_status(db, row_id, status, flag,
                                      pdf_path=None, remarks=None, clear_remarks=False):
    """
    Update the sourcing-validation columns of a single TProcessStatus row.

      status  : None = pending, 0 = fail, 1 = pass  (SourcingValidationStatus)
      flag    : 's' = started, 'p' = processing, 'c' = complete  (SourcingValidationFlag)
      pdf_path: if not None, also updates PdfFilePath
      remarks : if not None, sets Remarks = <text>
      clear_remarks: if True, sets Remarks = NULL. Used at the very start of a row
                     to wipe any stale remark from a prior run. Ignored when an
                     explicit `remarks` text is supplied (an explicit remark wins).

    Keyed on the physical TProcessStatus.Id (row_id). ProcessingId is non-unique
    and must NOT be used as the write key.
    """
    sets   = ["SourcingValidationStatus = ?", "SourcingValidationFlag = ?", "ModifiedOn = GETDATE()"]
    params = [status, flag]

    if pdf_path is not None:
        sets.append("PdfFilePath = ?")
        params.append(str(pdf_path))

    if remarks is not None:
        sets.append("Remarks = ?")
        params.append(remarks)
    elif clear_remarks:
        sets.append("Remarks = NULL")   # no bound param

    params.append(row_id)

    res = db.update(f"""
        UPDATE TProcessStatus
        SET {", ".join(sets)}
        WHERE Id = ?
    """, params)
    if not res.success:
        print(f"[FAILED] update_sourcing_validation_status row_id={row_id}: {res.error}")
    return res



# ============================================================================
# ============================================================================
# VALIDATION ENGINE  (lifted verbatim from the FAAC sourcing/validation tool)
#   Gate 1: issuer details | Gate 2: audit compliance | Gate 3: statements
#   + framework/basis detection + OCR detection. Excel/sourcing/tqdm-bar code
#   was dropped; the non-timed run_gate1/2/3 drivers were dropped in favour of
#   the crash-hardened run_gate*_timed runners used by validate_pdf().
#   Single-threaded ONLY: run_gate3_timed mutates STATEMENT_SPECS in place and
#   restores it in a finally — do NOT parallelize rows without de-coupling that.
# ============================================================================
# ============================================================================

CURRENT_YEAR      = datetime.now().year

CLR_LIGHT_GREEN  = "C6EFCE"

CLR_LIGHT_YELLOW = "FFEB9C"

CLR_LIGHT_RED    = "FFC7CE"

G1_STEPS = ["G1-1", "G1-2", "G1-3"]

G2_STEPS = ["G2-1", "G2-2", "G2-3", "G2-4"]

G3_LG_STEPS    = ["G3-1", "G3-2", "G3-3", "G3-4", "G3-5"]

G3_NONLG_STEPS = ["G3-1", "G3-2", "G3-4", "G3-6"]

US_STATES = {
 "AL":"Alabama","AK":"Alaska","AZ":"Arizona","AR":"Arkansas","CA":"California",
 "CO":"Colorado","CT":"Connecticut","DE":"Delaware","DC":"District of Columbia",
 "FL":"Florida","GA":"Georgia","HI":"Hawaii","ID":"Idaho","IL":"Illinois",
 "IN":"Indiana","IA":"Iowa","KS":"Kansas","KY":"Kentucky","LA":"Louisiana",
 "ME":"Maine","MD":"Maryland","MA":"Massachusetts","MI":"Michigan","MN":"Minnesota",
 "MS":"Mississippi","MO":"Missouri","MT":"Montana","NE":"Nebraska","NV":"Nevada",
 "NH":"New Hampshire","NJ":"New Jersey","NM":"New Mexico","NY":"New York",
 "NC":"North Carolina","ND":"North Dakota","OH":"Ohio","OK":"Oklahoma","OR":"Oregon",
 "PA":"Pennsylvania","RI":"Rhode Island","SC":"South Carolina","SD":"South Dakota",
 "TN":"Tennessee","TX":"Texas","UT":"Utah","VT":"Vermont","VA":"Virginia",
 "WA":"Washington","WV":"West Virginia","WI":"Wisconsin","WY":"Wyoming","PR":"Puerto Rico",
 "GU":"Guam"
}

STOPWORDS = {
    "the","of","and","for","in","on","at","to","a","an","by","with","or",
    "inc","incorporated","llc","plc","ltd","co","corp","corporation","company",
    "dba","dbaof","aka","akaof","fka","fkaof",
    "d/b/a","d.b.a","a/k/a","a.k.a","f/k/a","f.k.a"
}

SPECIAL_CHARS_RE = re.compile(r"[^\w\s]")   # keeps alphanumerics + underscore + space

def normalize_entity_name(name: str):
    """Lowercase → strip specials/symbols/non-printable → tokenize → drop stopwords."""
    if not isinstance(name, str) or not name.strip():
        return []
    # Unicode NFKD → drop non-printable
    s = unicodedata.normalize("NFKD", name)
    s = "".join(ch for ch in s if ch.isprintable())
    s = s.lower()
    # Remove d/b/a-style separators before stripping
    s = re.sub(r"\bd[./]?b[./]?a\b", " ", s)
    s = re.sub(r"\ba[./]?k[./]?a\b", " ", s)
    s = re.sub(r"\bf[./]?k[./]?a\b", " ", s)
    s = SPECIAL_CHARS_RE.sub(" ", s)
    tokens = [t for t in s.split() if t and t not in STOPWORDS and not t.isdigit()]
    # De-duplicate while preserving order
    seen, out = set(), []
    for t in tokens:
        if t not in seen:
            out.append(t); seen.add(t)
    return out

def match_pct(workbook_tokens: List[str], sentence: str) -> float:
    """Percentage of workbook_tokens that appear in the sentence (case-insensitive)."""
    if not workbook_tokens:
        return 0.0
    s_low = sentence.lower()
    hits = sum(1 for t in workbook_tokens if re.search(rf"\b{re.escape(t)}\b", s_low))
    return round(100.0 * hits / len(workbook_tokens), 2)

FYE_DATE_RE = re.compile(
    r"""
    (?P<full>
        # Month-Name day, year   e.g. June 30, 2025 / June 30 - 2025 / 30-June-2025
        (?:
            (?:jan(?:uary)?|feb(?:ruary)?|mar(?:ch)?|apr(?:il)?|may|jun(?:e)?|
               jul(?:y)?|aug(?:ust)?|sep(?:t|tember)?|oct(?:ober)?|nov(?:ember)?|dec(?:ember)?)
            [\s\-.]+\d{1,2}(?:st|nd|rd|th)?[\s\-.,]+\d{2,4}
          |
            \d{1,2}(?:st|nd|rd|th)?[\s\-.]+
            (?:jan(?:uary)?|feb(?:ruary)?|mar(?:ch)?|apr(?:il)?|may|jun(?:e)?|
               jul(?:y)?|aug(?:ust)?|sep(?:t|tember)?|oct(?:ober)?|nov(?:ember)?|dec(?:ember)?)
            [\s\-.,]+\d{2,4}
          |
            # numeric  mm/dd/yy(yy)  mm-dd-yy(yy)  mm.dd.yy(yy)
            \b\d{1,2}[./\-]\d{1,2}[./\-]\d{2,4}\b
        )
    )
    """,
    re.IGNORECASE | re.VERBOSE,
)

def parse_fye_from_workbook(cell_value) -> Optional[Tuple[int, int]]:
    """
    Robust workbook FYE parser. Returns (month, year) tuple or None.
    
    Handles:
      • Excel serial numbers (int/float, e.g. 45900 = 2025-08-31)
      • datetime and date objects
      • ISO-format strings ('2025-08-31')
      • US-format strings ('8/31/2025', '08/31/2025')
      • 2-digit year strings ('8/31/25' → assumes 2000s)
      • Long-form strings ('August 31, 2025', 'Aug. 31, 2025')
      • Prefixed strings ('FYE: 8/31/2025', 'Fiscal Year End August 31, 2025')
      • Curly quotes, non-breaking spaces, extra whitespace
    
    Rejects (returns None):
      • Empty/None/NaN values
      • Values without a full year+month (e.g. '8/31', '2025')
      • Values without any recognizable date pattern
      • Out-of-realistic-range values (year < 1990 or > current+2)
    """
    # ---- Reject null/empty ----
    if cell_value is None:
        return None
    if isinstance(cell_value, float) and pd.isna(cell_value):
        return None
    
    # ---- Handle datetime and date objects directly ----
    if isinstance(cell_value, datetime):
        return _validate_fye_range(cell_value.month, cell_value.year)
    if isinstance(cell_value, date):
        return _validate_fye_range(cell_value.month, cell_value.year)
    
    # ---- Handle Excel serial numbers (sanity-check the range) ----
    if isinstance(cell_value, (int, float)) and not pd.isna(cell_value):
        try:
            serial = int(cell_value)
            # Reasonable Excel serial range: 32874 (1990-01-01) to 55153 (2050-12-31)
            if 32874 <= serial <= 55153:
                base = datetime(1899, 12, 30)
                dt = base + pd.Timedelta(days=serial)
                return _validate_fye_range(dt.month, dt.year)
            else:
                # Number outside plausible Excel-date range
                # (could be a year like 2025 misplaced in the cell — reject)
                return None
        except Exception:
            return None
    
    # ---- Normalize string input ----
    s = str(cell_value).strip()
    if not s:
        return None
    
    # Clean up common Unicode issues
    s = (s.replace("\u00A0", " ")     # non-breaking space
         .replace("\u2019", "'")       # curly apostrophe
         .replace("\u2018", "'")
         .replace("\u2013", "-")       # en-dash
         .replace("\u2014", "-"))      # em-dash
    s = re.sub(r"\s+", " ", s).strip()
    
    # ---- Try explicit format patterns FIRST (before fuzzy parsing) ----
    # These are strict and prevent dateutil from inventing years.
    
    # Pattern 1: mm/dd/yyyy or m/d/yyyy or mm-dd-yyyy
    m = re.search(r"\b(\d{1,2})/\-\.\d{4}\b", s)
    if m:
        month, day, year = int(m.group(1)), int(m.group(2)), int(m.group(3))
        if 1 <= month <= 12 and 1 <= day <= 31:
            return _validate_fye_range(month, year)
    
    # Pattern 2: mm/dd/yy (2-digit year → assume 20xx)
    m = re.search(r"\b(\d{1,2})/\-\.\d{2}\b", s)
    if m:
        month, day, yy = int(m.group(1)), int(m.group(2)), int(m.group(3))
        year = 2000 + yy if yy < 50 else 1900 + yy
        if 1 <= month <= 12 and 1 <= day <= 31:
            return _validate_fye_range(month, year)
    
    # Pattern 3: yyyy-mm-dd (ISO)
    m = re.search(r"\b(\d{4})/\-\.\d{1,2}\b", s)
    if m:
        year, month, day = int(m.group(1)), int(m.group(2)), int(m.group(3))
        if 1 <= month <= 12 and 1 <= day <= 31:
            return _validate_fye_range(month, year)
    
    # Pattern 4: Month-name day, year  (e.g. "August 31, 2025")
    month_map = {
        'jan': 1, 'january': 1, 'feb': 2, 'february': 2,
        'mar': 3, 'march': 3, 'apr': 4, 'april': 4,
        'may': 5, 'jun': 6, 'june': 6,
        'jul': 7, 'july': 7, 'aug': 8, 'august': 8,
        'sep': 9, 'sept': 9, 'september': 9,
        'oct': 10, 'october': 10, 'nov': 11, 'november': 11,
        'dec': 12, 'december': 12,
    }
    m = re.search(
        r"\b(jan(?:uary)?|feb(?:ruary)?|mar(?:ch)?|apr(?:il)?|may|jun(?:e)?|"
        r"jul(?:y)?|aug(?:ust)?|sep(?:t|tember)?|oct(?:ober)?|nov(?:ember)?|dec(?:ember)?)"
        r"[\s\.\-,]+(\d{1,2})(?:st|nd|rd|th)?[\s\.\-,]+(\d{4})\b",
        s, re.IGNORECASE)
    if m:
        month = month_map.get(m.group(1).lower().rstrip('.'), 0)
        day = int(m.group(2))
        year = int(m.group(3))
        if month and 1 <= day <= 31:
            return _validate_fye_range(month, year)
    
    # Pattern 5: day Month-name year  (e.g. "31 August 2025")
    m = re.search(
        r"\b(\d{1,2})(?:st|nd|rd|th)?[\s\.\-,]+"
        r"(jan(?:uary)?|feb(?:ruary)?|mar(?:ch)?|apr(?:il)?|may|jun(?:e)?|"
        r"jul(?:y)?|aug(?:ust)?|sep(?:t|tember)?|oct(?:ober)?|nov(?:ember)?|dec(?:ember)?)"
        r"[\s\.\-,]+(\d{4})\b",
        s, re.IGNORECASE)
    if m:
        day = int(m.group(1))
        month = month_map.get(m.group(2).lower().rstrip('.'), 0)
        year = int(m.group(3))
        if month and 1 <= day <= 31:
            return _validate_fye_range(month, year)
    
    # ---- Fallback: dateutil with strict validation ----
    # Only accept if BOTH month and year come from actual text (not defaulted).
    try:
        dt_default = datetime(1900, 1, 1)   # sentinel default
        dt = dtparser.parse(s, dayfirst=False, fuzzy=True, default=dt_default)
        # If year didn't change from default, dateutil invented it → reject
        if dt.year == 1900:
            return None
        return _validate_fye_range(dt.month, dt.year)
    except Exception:
        return None

def _validate_fye_range(month: int, year: int) -> Optional[Tuple[int, int]]:
    """
    Sanity-check parsed FYE against realistic ranges.
    Rejects years before 1990 or more than 3 years in the future.
    """
    if not (1 <= month <= 12):
        return None
    current_year = datetime.now().year
    if year < 1990 or year > current_year + 3:
        return None
    return (month, year)

def extract_dates_from_text(text: str) -> List[Tuple[int, int, str]]:
    """Find all date-like patterns in text → list of (month, year, matched_string)."""
    results = []
    for m in FYE_DATE_RE.finditer(text):
        raw = m.group("full")
        try:
            dt = dtparser.parse(raw, dayfirst=False, fuzzy=True)
            results.append((dt.month, dt.year, raw))
        except Exception:
            continue
    return results

def fye_matches(target: Tuple[int, int], candidate: Tuple[int, int]) -> str:
    """Return 'EXACT' | 'MONTH' | 'YEAR' | 'NONE'.
       Includes optional diagnostic printing when DEBUG_FYE is enabled."""
    if not target or not candidate:
        return "NONE"
    tm, ty = target
    cm, cy = candidate
    if tm == cm and ty == cy:
        return "EXACT"
    if tm == cm:
        return "MONTH"
    if ty == cy:
        return "YEAR"
    return "NONE"

DEBUG_FYE = False    # TWEAK ZONE: enable for debugging

@dataclass
class StepResult:
    code: str                                          # e.g. "G1-1"
    status: str = "FAIL"                                # PASS / FAIL / REVIEW / NOT FOUND
    score:  float = 0.0                                 # 0-100
    page_idx: Optional[int] = None
    screenshot: Optional[str] = None
    notes: str = ""
    page_range: Optional[str] = None                    # e.g. "5-7" for multi-page statements
    reconciliation_note: Optional[str] = None           # ← NEW: appended to cell display with " - "
    classification: Optional[str] = None      # ← NEW: e.g. "Present Fairly" for G2-1

@dataclass
class PageLine:
    page_idx: int          # 0-based
    line_idx: int          # 0-based within page
    text: str
    bbox:   Tuple[float, float, float, float]   # x0, y0, x1, y1

def extract_page_lines(doc: fitz.Document, page_idx: int):
    """
    Return line-level text with bboxes, SORTED VISUALLY top-to-bottom.
    
    CRITICAL FIX: PyMuPDF's default block ordering places table content BEFORE
    surrounding captions/titles for many financial statement pages (esp. FASB
    nonprofit healthcare). Without y0-sort, 'Combined Statements of Cash Flows'
    lands at line 40+ and is invisible to top-N-line matching.
    """
    page = doc[page_idx]
    d = page.get_text("dict")
    raw = []
    for block in d.get("blocks", []):
        if block.get("type", 0) != 0:
            continue
        for line in block.get("lines", []):
            text = " ".join(span["text"] for span in line.get("spans", []))
            text = text.strip()
            if not text:
                continue
            x0, y0, x1, y1 = line["bbox"]
            raw.append((y0, x0, text, (x0, y0, x1, y1)))
    
    # Sort by y0 (top-to-bottom), then x0 (left-to-right)
    raw.sort(key=lambda t: (t[0], t[1]))
    
    lines = []
    for li, (_, _, text, bbox) in enumerate(raw):
        lines.append(PageLine(page_idx, li, text, bbox))
    return lines

def page_top_lines(doc: fitz.Document, page_idx: int, n: int = 5):
    """
    First n visible text lines on a page — used for headings & top-of-page date checks.
    Filters out common short header/footer artifacts that push real content down:
         • 'Exhibit N', 'Schedule N', 'Table N', 'Note N' single-line markers
         • Standalone page numbers (like '47' or 'ii' or 'Page 3')
         • Very short strings (< 3 chars) that are typically noise
    """
    all_lines = extract_page_lines(doc, page_idx)
    filtered = []
    for ln in all_lines:
        text = ln.text.strip()
        # Skip empty lines
        if not text:
            continue
        # Skip standalone exhibit/schedule/table markers
        if re.match(r"^(exhibit|schedule|table|note|section)\s+\w+\s*$",
                    text, re.IGNORECASE):
            continue
        # Skip standalone page numbers (arabic or roman)
        if re.match(r"^\d{1,3}$", text):
            continue
        if re.match(r"^[ivxlc]{1,6}$", text, re.IGNORECASE):
            continue
        # Skip "Page N" / "Page N of M" markers
        if re.match(r"^page\s+\d+(\s+of\s+\d+)?$", text, re.IGNORECASE):
            continue
        # Skip very short strings (typically noise like "1", "-", "*")
        if len(text) < 3:
            continue
        filtered.append(ln)
        if len(filtered) >= n:
            break
    return filtered

def iter_sentences(doc: fitz.Document,
                   page_range: Optional[Tuple[int, int]] = None):
    total = doc.page_count
    lo, hi = (0, total) if page_range is None else (page_range[0], min(page_range[1], total))
    buf_lines: List[PageLine] = []
    buf_text  = ""
    last_page = lo                              # ← NEW: initialize
    for p in range(lo, hi):
        for ln in extract_page_lines(doc, p):
            buf_lines.append(ln)
            last_page = p                       # ← NEW: track most recent page seen
            buf_text = (buf_text + " " + ln.text).strip() if buf_text else ln.text
            if re.search(r"[.?!]\s*$", ln.text):
                yield (p, buf_text, buf_lines[:])
                buf_lines, buf_text = [], ""
    if buf_text:
        yield (buf_lines[-1].page_idx if buf_lines else last_page, buf_text, buf_lines[:])

def _sanitize_filename(s: str, max_len: int = 34) -> str:
    """Safe filesystem name; trims to max_len chars."""
    s = re.sub(r"[^A-Za-z0-9_\- ]", "", str(s)).strip()
    s = re.sub(r"\s+", "_", s)
    return s[:max_len] or "UNKNOWN"

def build_screenshot_name(sector: str, state: str, audit_year,
                          step_code: str, page_idx: int) -> str:
    """Construct a standardized screenshot filename.
    
    Pattern: <SECTOR>_<STATE>_<AUDIT_YEAR>_<STEP_CODE>_p<PAGE_NUM>.png
    
    Example: LG_TX_2025_G1-1_p13.png
              │  │  │    │    │
              │  │  │    │    └── 1-based page number (page_idx + 1)
              │  │  │    └────── validation step (G1-1, G3-4, etc.)
              │  │  └────────── audit year (2025)
              │  └──────────── 2-char state code (TX)
              └──────────── sector (LG or NONLG)
    
    All components are normalized: uppercase, whitespace-stripped, and
    non-safe characters removed. Missing components are replaced with 'UNK'.
    """
    sector_str = str(sector).upper().strip() if sector else "UNK"
    state_str = str(state).upper().strip() if state else "UNK"
    
    # Handle audit_year: could be int, str, or None
    try:
        year_str = str(int(audit_year)) if audit_year else "UNKYR"
    except (ValueError, TypeError):
        year_str = str(audit_year) if audit_year else "UNKYR"
    
    step_str = str(step_code).upper().strip() if step_code else "UNK"
    
    # Convert 0-based page_idx to 1-based page number for user-friendly filenames
    try:
        page_num = int(page_idx) + 1 if page_idx is not None else 0
    except (ValueError, TypeError):
        page_num = 0
    
    # Sanitize each component: remove characters unsafe for filesystems
    def _clean(s: str) -> str:
        return re.sub(r"[^A-Za-z0-9\-_]", "", str(s))
    
    filename = (f"{_clean(sector_str)}_{_clean(state_str)}_"
                f"{_clean(year_str)}_{_clean(step_str)}_p{page_num}.png")
    return filename

def save_highlighted_screenshot(
    doc:  fitz.Document, page_idx: int, bboxes: List[Tuple[float, float, float, float]],
    highlight_hex: str, out_dir: Path, filename: str, dpi: int = 150) -> Path:
    """
    Draws colored rectangles over the given bboxes on a copy of the page,
    renders to PNG, saves and returns the file path.
    Includes robust error handling so failures don't crash validation.
    """
    out_dir.mkdir(parents=True, exist_ok=True)
    out = out_dir / filename
    tmp = None
    try:
        # Guard: page index must be valid
        if page_idx < 0 or page_idx >= doc.page_count:
            raise ValueError(f"Page {page_idx} out of range (doc has {doc.page_count})")

        # Copy the target page into a new PDF so annotations don't affect the original
        tmp = fitz.open()
        tmp.insert_pdf(doc, from_page=page_idx, to_page=page_idx)
        page = tmp[0]

        # HEX color → 0-1 RGB
        r = int(highlight_hex[0:2], 16) / 255
        g = int(highlight_hex[2:4], 16) / 255
        b = int(highlight_hex[4:6], 16) / 255

        # Add semi-transparent highlight rectangles
        for bb in (bboxes or []):
            try:
                rect = fitz.Rect(*bb)
                if rect.is_empty or rect.width < 1 or rect.height < 1:
                    continue
                annot = page.add_rect_annot(rect)
                annot.set_colors(stroke=(r, g, b), fill=(r, g, b))
                annot.set_opacity(0.35)
                annot.update()
            except Exception as bbox_ex:
                print(f"     ⚠️  Bad bbox {bb}: {bbox_ex}")

        # Render page → PNG
        pix = page.get_pixmap(dpi=dpi)
        pix.save(str(out))
        
        # ---- Verify PNG actually exists on disk ----
        if not out.exists():
            print(f"     ❌ Screenshot not on disk after save: {out}")
            return None
        size_kb = out.stat().st_size / 1024
        tqdm.write(f"     🖼  Saved: {out.relative_to(out.parents[2])} ({size_kb:.1f} KB)")
        return out

    except Exception as ex:
        print(f"     ❌ Screenshot save failed → {filename}: {ex}")
        raise
    finally:
        if tmp is not None:
            try: tmp.close()
            except Exception: pass

def take_diagnostic_screenshot(
    doc: fitz.Document, page_idx: Optional[int], sector: str,
    state: str, audit_year, step_code: str, screenshot_out: Path,
    status: str = "FAIL", bboxes: Optional[List] = None):
    """
    Always-on screenshot helper — captures the relevant page even on FAIL
    so the human auditor can see WHY the step failed.  For FAIL cases,
    captures page 0 (cover) or the first non-empty page as a fallback.
    """
    if not doc or doc.page_count == 0:
        return None
    if page_idx is None or page_idx >= doc.page_count:
        # No specific page — grab cover page for diagnostics
        page_idx = 0
    highlight = (CLR_LIGHT_GREEN if status == "PASS" else
                 CLR_LIGHT_YELLOW if status == "REVIEW" else
                 CLR_LIGHT_RED)
    try:
        shot_name = build_screenshot_name(sector, state, audit_year, step_code, page_idx)
        shot_path = save_highlighted_screenshot(
            doc, page_idx, bboxes or [], highlight, screenshot_out, shot_name
        )
        return str(shot_path)
    except Exception as ex:
        print(f"     ⚠️  Screenshot failed for {step_code} p{page_idx+1}: {ex}")
        return None

OCR_STATUS           = "OCR REQUIRED"

OCR_MIN_TEXT_CHARS   = 40      # a genuine content page yields more than this

OCR_MIN_IMG_AREA     = 20000   # px² — ignore tiny logos/rules

OCR_SCANNED_FRACTION = 0.60    # ≥60% of content pages image-only → doc needs OCR

OCR_SAMPLE_MAX       = 40      # cap pages sampled (perf on 300-page audits)

def _page_text_len(doc: fitz.Document, pi: int) -> int:
    try:
        return len(doc[pi].get_text("text").strip())
    except Exception:
        return 0

def _page_has_raster(doc: fitz.Document, pi: int) -> bool:
    """
    True if page pi carries a meaningful raster image OR flattened full-page
    vector art with a negligible text layer (i.e., a scanned/rendered page).
     """
    try:
        page = doc[pi]
    except Exception:
        return False
    # 1) Embedded raster via get_image_info (original path)
    try:
        try:
            info = page.get_image_info(xrefs=True) or []
        except TypeError:
            info = page.get_image_info() or []
    except Exception:
        info = []
    for i in info:
        b = i.get("bbox")
        if b and (b[2] - b[0]) * (b[3] - b[1]) >= OCR_MIN_IMG_AREA:
            return True
    # 2) Some scanners register the image only through get_images()
    try:
        if page.get_images(full=True):
            return True
    except Exception:
        pass
    # 3) Flattened vector line-art (glyph outlines) with no text layer
    try:
        if len(page.get_drawings()) >= 10:
            return True
    except Exception:
        pass
    return False

def page_needs_ocr(doc: fitz.Document, pi: int) -> bool:
    "True if page pi is image-only (raster present, negligible text layer)."
    if pi < 0 or pi >= doc.page_count:
        return False
    return _page_has_raster(doc, pi) and _page_text_len(doc, pi) < OCR_MIN_TEXT_CHARS

def doc_requires_ocr(doc: fitz.Document) -> bool:
    """Sample up to OCR_SAMPLE_MAX evenly-spaced pages. True when the PDF lacks a
       usable text layer across its CONTENT pages (blank pages ignored)."""
    n = doc.page_count
    if n == 0:
        return False
    if n <= OCR_SAMPLE_MAX:
        idxs = list(range(n))
    else:
        step = n / float(OCR_SAMPLE_MAX)
        idxs = sorted({int(i * step) for i in range(OCR_SAMPLE_MAX)})
    scanned = considered = 0
    for p in idxs:
        txt = _page_text_len(doc, p)
        if txt >= OCR_MIN_TEXT_CHARS:
            considered += 1                       # usable text layer
        elif _page_has_raster(doc, p):
            considered += 1; scanned += 1         # image-only content page
        # else: truly blank page → ignored
    return considered > 0 and (scanned / considered) >= OCR_SCANNED_FRACTION

def apply_ocr_overlay(ocr_required: bool, gate_status: str,
                      results: Dict[str, StepResult]) -> str:
    """
    When the doc requires OCR, relabel every step that could NOT be evaluated
    from text (FAIL / NOT FOUND) as 'OCR REQUIRED', and surface it at GATE
    level. PASS/REVIEW (residual text) and NOT APPLICABLE are left intact.
    """
    if not ocr_required:
        return gate_status
    any_ocr = False
    for sr in results.values():
        if sr is None:
            continue
        if sr.status in ("FAIL", "NOT FOUND"):
            sr.status = OCR_STATUS; sr.score = 0.0
            note = "OCR REQUIRED: evidence page(s) are image-only (no extractable text)."
            sr.notes = f"{sr.notes} | {note}".strip(" |") if sr.notes else note
            any_ocr = True
    return OCR_STATUS if any_ocr else gate_status

def relabel_step_if_page_ocr(sr, doc):
    """PER-PAGE OCR guard for ONE step (works for G1, G2, G3 - LG and NONLG).
    If the step's evidence page is an image-only scan, mark it 'OCR REQUIRED'."""
    if sr is None or sr.status == OCR_STATUS:
        return False
    pi = sr.page_idx
    if pi is None or not page_needs_ocr(doc, pi):
        return False
    note = ("OCR REQUIRED: evidence page is image-only "
            "(content not in extractable text layer).")
    sr.status = OCR_STATUS
    sr.score  = 0.0
    sr.notes  = f"{sr.notes} | {note}".strip(" |") if sr.notes else note
    return True

def relabel_notfound_stmt_if_ocr(sr, doc, candidate_pages=None) -> bool:
    "Gate-3 companion to relabel_step_if_page_ocr: if a statement step is "
    "FAIL/NOT FOUND *because* the evidence pages are image-only, relabel to "
    "OCR REQUIRED and pin a representative page for the diagnostic screenshot."
    if sr is None or sr.status not in ("FAIL", "NOT FOUND"):
        return False
    if not doc_requires_ocr(doc):
        return False
    pi = None
    scan_range = candidate_pages if candidate_pages else range(doc.page_count)
    for p in scan_range:
        if page_needs_ocr(doc, p):
            pi = p
            break
    sr.status   = OCR_STATUS
    sr.score    = 0.0
    sr.page_idx = pi if pi is not None else sr.page_idx
    note = ("OCR REQUIRED: statement page(s) are image-only / flattened "
            "(no extractable text layer).")
    sr.notes = (sr.notes + " | " + note).strip(" |") if sr.notes else note
    return True

G1_FIRST_N_PAGES  = 30      # spec: "first 30 pages"

G1_REVIEW_MIN_PCT = 90.0    # >= 90 and < 100 → REVIEW   (G1-1 entity-name scale)

G1_PASS_PCT       = 100.0   # ==100 → PASS               (G1-1 entity-name scale)

FYE_TOP_LINES     = 5       # spec: "first five lines from top" of a page

def validate_g1_1_entity_name(doc: fitz.Document,
                              workbook_entity: str,
                              sector: str, state: str, audit_year,
                              screenshot_out: Path) -> StepResult:
    """Scan first 30 pages sentence-by-sentence AND line-by-line.
       Title pages (no punctuation) are caught by the line-level scan."""
    result = StepResult(code="G1-1")
    tokens = normalize_entity_name(workbook_entity)
    if not tokens:
        result.status = "FAIL"; result.notes = "Empty workbook entity name."
        return result

    best_score = 0.0
    best_page  = None
    best_lines: List[PageLine] = []
    best_sent  = ""

    page_limit = min(G1_FIRST_N_PAGES, doc.page_count)

    # ---- PASS 1: Line-by-line scan of top 10 lines of first 5 pages
    #      (catches cover / title pages that have no sentence punctuation) ----
    for p in range(min(5, page_limit)):
        top = page_top_lines(doc, p, 10)
        for ln in top:
            pct = match_pct(tokens, ln.text)
            if pct >= G1_PASS_PCT:
                best_score, best_page, best_lines, best_sent = pct, p, [ln], ln.text
                break
            if pct > best_score and pct >= G1_REVIEW_MIN_PCT:
                best_score, best_page, best_lines, best_sent = pct, p, [ln], ln.text
        if best_score >= G1_PASS_PCT:
            break

    # ---- PASS 2: Sentence-level scan (existing logic) - only if no perfect line hit ----
    if best_score < G1_PASS_PCT:
        for p_idx, sentence, lines in iter_sentences(doc, (0, page_limit)):
            pct = match_pct(tokens, sentence)
            if pct >= G1_PASS_PCT:
                best_score, best_page, best_lines, best_sent = pct, p_idx, lines, sentence
                break
            if pct > best_score and pct >= G1_REVIEW_MIN_PCT:
                best_score, best_page, best_lines, best_sent = pct, p_idx, lines, sentence

    # ---- Categorize ----
    if best_page is not None and best_score >= G1_PASS_PCT:
        highlight = CLR_LIGHT_GREEN
        status    = "PASS"
    elif best_page is not None and G1_REVIEW_MIN_PCT <= best_score < G1_PASS_PCT:
        highlight = CLR_LIGHT_YELLOW
        status    = "REVIEW"
    else:
        result.status = "FAIL"
        result.score  = round(best_score, 2)
        result.notes  = "No sentence reached 90% entity-name match in first 30 pages."
        return result

    # ---- Screenshot ----
    bboxes = [ln.bbox for ln in best_lines] if best_lines else []
    shot_name = build_screenshot_name(sector, state, audit_year, "G1-1", best_page)
    shot_path = save_highlighted_screenshot(
        doc, best_page, bboxes, highlight, screenshot_out, shot_name
    )
    result.status     = status
    result.score      = round(best_score, 2)
    result.page_idx   = best_page
    result.screenshot = str(shot_path)
    result.notes      = f"Matched: {best_sent[:120]!r}"
    return result

def _is_sole_state_line(line_text: str, state_code: str, state_name: str) -> bool:
    """
    True when the line is a STATE-IDENTITY line (→ 100% PASS), NOT a prose description.

    Identity line = the State (full name case-INSENSITIVE / 2-letter code UPPERCASE)
    surrounded ONLY by proper-noun place tokens (city / county / entity words in
    Title-case or ALL-CAPS) plus a few location connectors.
    A prose 'description' (long, or containing lowercase verb/function words) is NOT
    an identity line → caller correctly falls through to the lower REVIEW tier.
    """
    if not line_text:
        return False
    norm = line_text.replace("\u00A0", " ").replace("\u2019", "'").strip()

    # State must actually be present on this line.
    has_full = bool(state_name and re.search(rf"\b{re.escape(state_name)}\b", norm, re.IGNORECASE))
    has_code = bool(state_code and re.search(rf"\b{re.escape(state_code.upper())}\b", norm))
    if not (has_full or has_code):
        return False

    # Strip the matched State, then inspect the RESIDUAL tokens.
    residual = norm
    if has_full:
        residual = re.sub(rf"\b{re.escape(state_name)}\b", " ", residual, flags=re.IGNORECASE)
    if has_code:
        residual = re.sub(rf"\b{re.escape(state_code.upper())}\b", " ", residual)
    residual = re.sub(r"[^A-Za-z ]", " ", residual)      # drop commas/periods/digits
    tokens = [t for t in residual.split() if t]

    # Nothing left → the line is essentially just the State.
    if not tokens:
        return True

    # Identity heuristic: short line, and every residual token is a proper noun
    # (Title-case or ALL-CAPS) or a permitted location connector.
    if len(tokens) > 8:
        return False                                     # long line → prose, not identity
    _CONNECTORS = {"of", "and", "the"}
    for t in tokens:
        if t.lower() in _CONNECTORS:
            continue
        if not (t.isupper() or t[0].isupper()):
            return False                                 # a lowercase prose word → description
    return True

def _find_state_in_line(line_text: str, state_code: str, state_name: str) -> bool:
    "Full state name → case-INSENSITIVE.  Two-letter code → case-SENSITIVE (UPPERCASE only)."
    if not line_text:
        return False
    norm = line_text.replace("\u00A0", " ").replace("\u2019", "'")
    norm = re.sub(r"\s+", " ", norm).strip()

    # (1) FULL NAME → case-INSENSITIVE  ← this branch was effectively broken
    if state_name and re.search(rf"\b{re.escape(state_name)}\b", norm, re.IGNORECASE):
        return True

    # (2) TWO-LETTER CODE → CASE-SENSITIVE, uppercase only (NO IGNORECASE flag)
    if state_code and re.search(rf"\b{re.escape(state_code.upper())}\b", norm):
        return True

    return False

def validate_g1_2_state(doc: fitz.Document,
                        workbook_state: str,
                        entity_g1_1_result: StepResult,
                        sector: str, audit_year,
                        screenshot_out: Path) -> StepResult:
    """
    Spec-driven state validation with nonprofit fallback tier.

    ═══ PRIMARY SPEC (Passes 1 + 2) ═══
      • If a line SOLELY contains the state name/code → PASS @100%, green
      • Else measure line-distance from ENTITY-NAME line
            same line  → PASS 100% green
            +1 line    → PASS  95% green
            +2 lines   → REVIEW 90% yellow
            >2 lines   → FAIL   [100 - 5*d]% red

    ═══ FALLBACK TIER (Pass 3, only when Pass 2 → FAIL) ═══
      • Search first 30 pages for high-confidence organizational phrases like
        "New Mexico nonprofit corporation", "state of New Mexico",
        "hospitals in New Mexico" — common in FASB nonprofit healthcare/
        university reports where state appears deep in Note 1 "Organization".
      • Fallback match scores 95% (slight discount from spec's 100%).
    """
    result = StepResult(code="G1-2")
    state_code = str(workbook_state or "").strip().upper()
    state_name = US_STATES.get(state_code, "")
    if not state_code or not state_name:
        result.status = "FAIL"; result.notes = f"Unknown state code: {workbook_state!r}"
        return result

    page_limit = min(G1_FIRST_N_PAGES, doc.page_count)
    entity_page = entity_g1_1_result.page_idx

    # ═══════════════════════════════════════════════════════════════════
    # PASS 1 — Search first 30 pages for a SOLE state-line
    # ═══════════════════════════════════════════════════════════════════
    for p in range(page_limit):
        for ln in extract_page_lines(doc, p):
            if _is_sole_state_line(ln.text, state_code, state_name):
                shot_name = build_screenshot_name(sector, state_code, audit_year, "G1-2", p)
                shot_path = save_highlighted_screenshot(
                    doc, p, [ln.bbox], CLR_LIGHT_GREEN, screenshot_out, shot_name
                )
                result.status = "PASS"; result.score = 100.0
                result.page_idx = p; result.screenshot = str(shot_path)
                result.notes = f"Sole state line: {ln.text!r}"
                return result

    # ═══════════════════════════════════════════════════════════════════
    # PASS 2 — Distance-from-entity-name scoring
    # ═══════════════════════════════════════════════════════════════════
    if entity_page is None:
        # No entity anchor → still try to find state anywhere in first 30 pages
        for p in range(page_limit):
            for ln in extract_page_lines(doc, p):
                if _find_state_in_line(ln.text, state_code, state_name):
                    shot_name = build_screenshot_name(sector, state_code, audit_year, "G1-2", p)
                    shot_path = save_highlighted_screenshot(
                        doc, p, [ln.bbox], CLR_LIGHT_RED, screenshot_out, shot_name
                    )
                    result.status = "FAIL"; result.score = 50.0
                    result.page_idx = p; result.screenshot = str(shot_path)
                    result.notes = "State found but entity anchor missing."
                    # ✨ Don't return yet — try Pass 3 fallback first
                    break
            if result.status == "FAIL" and result.score == 50.0:
                break
        # ✨ Fall through to Pass 3 (fallback) instead of early-returning
    else:
        # Get lines of the entity page and find the entity-name line index.
        # ✅ FIX 1: single tokenization (the previous version computed the same
        #           normalize_entity_name(...) twice on consecutive lines).
        e_lines = extract_page_lines(doc, entity_page)
        tokens  = normalize_entity_name(str(entity_g1_1_result.notes))
        entity_line_idx = None
        best_hits = -1
        for i, ln in enumerate(e_lines):
            low = ln.text.lower()
            hits = sum(1 for t in tokens if t and re.search(rf"\b{re.escape(t)}\b", low))
            if hits > best_hits:
                best_hits, entity_line_idx = hits, i

        # Walk downward through this page and subsequent pages counting line distance
        def _walk_lines_from(page_idx: int, start_line: int):
            for i, ln in enumerate(extract_page_lines(doc, page_idx)):
                if i < start_line:
                    continue
                yield page_idx, i, ln
            for pp in range(page_idx + 1, page_limit):
                for j, ln2 in enumerate(extract_page_lines(doc, pp)):
                    yield pp, j, ln2

        distance = 0
        found_page = None
        found_line = None
        for pp, li, ln in _walk_lines_from(entity_page, entity_line_idx or 0):
            if _find_state_in_line(ln.text, state_code, state_name):
                # Skip if it's the exact same line as the entity anchor
                if pp == entity_page and li == (entity_line_idx or 0):
                    distance = 0
                found_page, found_line = pp, li
                break
            distance += 1

        if found_page is not None:
            # Score tiers per spec
            if distance == 0:
                score, status, highlight = 100.0, "PASS", CLR_LIGHT_GREEN
            elif distance == 1:
                score, status, highlight =  95.0, "PASS", CLR_LIGHT_GREEN
            elif distance == 2:
                score, status, highlight =  90.0, "REVIEW", CLR_LIGHT_YELLOW
            else:
                score = max(0.0, 100.0 - 5.0 * distance)
                status, highlight = "FAIL", CLR_LIGHT_RED

            # Screenshot bbox
            hit_lines = extract_page_lines(doc, found_page)
            bb = hit_lines[found_line].bbox if 0 <= found_line < len(hit_lines) else None
            shot_name = build_screenshot_name(sector, state_code, audit_year, "G1-2", found_page)
            shot_path = save_highlighted_screenshot(
                doc, found_page, [bb] if bb else [], highlight, screenshot_out, shot_name
            )
            result.status = status; result.score = score
            result.page_idx = found_page; result.screenshot = str(shot_path)
            result.notes = f"Distance from ENTITY NAME: {distance} line(s)"

            # ✨ Only return PASS/REVIEW here — let FAIL fall through to Pass 3
            if status in ("PASS", "REVIEW"):
                return result
        else:
            # State not found at/after entity name — set FAIL and try Pass 3
            result.status = "FAIL"; result.score = 0.0
            result.notes = "State not found at/after the entity-name line."

    # ═══════════════════════════════════════════════════════════════════
    # PASS 3 (FALLBACK) — Nonprofit-friendly state detection
    # For FASB nonprofits (hospitals, universities, foundations) where the
    # state appears in Note 1 "Organization" or footer rather than adjacent
    # to the entity name on cover/auditor's report pages.
    # ═══════════════════════════════════════════════════════════════════
    high_conf_state_patterns = [
        # Organizational/incorporation declarations (highest confidence)
        rf"\b{re.escape(state_name)}\s+nonprofit\s+corporation\b",
        rf"\b{re.escape(state_name)}\s+(?:non-?profit|not-?for-?profit)\b",
        rf"\bstate\s+of\s+{re.escape(state_name)}\b",
        rf"\borganized\s+(?:under|in)\s+(?:the\s+)?(?:laws?\s+of\s+)?"
        rf"(?:the\s+state\s+of\s+)?{re.escape(state_name)}\b",
        rf"\bincorporated\s+(?:in|under\s+(?:the\s+)?laws?\s+of)\s+"
        rf"(?:the\s+state\s+of\s+)?{re.escape(state_name)}\b",
        rf"\bdomiciled\s+in\s+(?:the\s+state\s+of\s+)?{re.escape(state_name)}\b",
        # Physical presence patterns (medium confidence)
        rf"\bhospitals?\s+in\s+{re.escape(state_name)}\b",
        rf"\bthroughout\s+{re.escape(state_name)}\b",
        rf"\bfacilities\s+(?:in|throughout|across)\s+{re.escape(state_name)}\b",
        rf"\boperates?\s+(?:in|throughout)\s+{re.escape(state_name)}\b",
        rf"\blocated\s+in\s+(?:the\s+state\s+of\s+)?{re.escape(state_name)}\b",
        rf"\bheadquartered\s+in\s+(?:the\s+state\s+of\s+)?{re.escape(state_name)}\b",
        # Contractual/regulatory relationships (medium confidence)
        rf"\bcontract\s+with\s+the\s+state\s+of\s+{re.escape(state_name)}\b",
        rf"\bstate\s+of\s+{re.escape(state_name)}\s+(?:department|division|agency)\b",
        rf"\b{re.escape(state_name)}\s+department\s+of\s+(?:health|education|revenue)\b",
    ]

    for pp in range(page_limit):
        for ln in extract_page_lines(doc, pp):
            text_lower = ln.text.lower()
            for pattern in high_conf_state_patterns:
                if re.search(pattern, text_lower, re.IGNORECASE):
                    # High-confidence organizational phrase found
                    shot_name = build_screenshot_name(sector, state_code, audit_year,
                                                      "G1-2", pp)
                    try:
                        shot_path = save_highlighted_screenshot(
                            doc, pp, [ln.bbox], CLR_LIGHT_GREEN,
                            screenshot_out, shot_name
                        )
                        screenshot_str = str(shot_path)
                    except Exception:
                        screenshot_str = None

                    result.status = "PASS"
                    result.score = 95.0    # slight discount from 100% for fallback
                    result.page_idx = pp
                    result.screenshot = screenshot_str
                    result.notes = (
                        f"State '{state_name}' found via organizational phrase "
                        f"on page {pp+1} (nonprofit fallback): "
                        f"{ln.text[:80]!r}"
                    )
                    return result

    # ═══════════════════════════════════════════════════════════════════
    # All 3 passes exhausted — return the best FAIL result we have
    # ═══════════════════════════════════════════════════════════════════
    if result.status not in ("PASS", "REVIEW", "FAIL"):
        result.status = "FAIL"
        result.score = 0.0
    if not result.notes:
        result.notes = (
            f"State '{state_name}' ({state_code}) not found within {page_limit} pages "
            f"via sole-line, line-distance, or organizational-phrase methods."
        )

    # Ensure diagnostic screenshot exists even on FAIL for auditor visibility
    if result.screenshot is None:
        try:
            result.screenshot = take_diagnostic_screenshot(
                doc, entity_page or 0, sector, state_code, audit_year,
                "G1-2", screenshot_out, status="FAIL"
            )
        except Exception:
            pass

    return result

def _best_date_in_line(line_text: str) -> Optional[Tuple[int, int, str]]:
    """When multiple dates exist in one line, spec says use the HIGHEST date."""
    dates = extract_dates_from_text(line_text)
    if not dates:
        return None
    # Sort by (year, month) desc
    dates.sort(key=lambda t: (t[1], t[0]), reverse=True)
    return dates[0]

def validate_g1_3_fye(doc: fitz.Document,
                      workbook_fye,
                      sector: str, state: str, audit_year,
                      screenshot_out: Path,
                      financial_stmt_pages: Optional[List[int]] = None) -> StepResult:
    """
    Order of Priority:
         A. Cover / title pages (first 10 pages — CAFRs often place the FYE
            title on p2/p3 behind an image cover)                → score 100 PASS
         B. Top 5 lines of any financial-statement page          → score 100 PASS
         C. Anywhere else, multiple dates in a line → 'highest'  → score  90 REVIEW
       Fallbacks:
         • Only month matches → REVIEW at <90% (per spec C.3)
         • Only year matches  → REVIEW at <90% (per spec C.3)
    """
    result = StepResult(code="G1-3")
    target = parse_fye_from_workbook(workbook_fye)
    if not target:
        result.status = "FAIL"; result.notes = f"Unparseable workbook FYE: {workbook_fye!r}"
        return result

    def _finalize(page_idx, bbox_list, level, matched_dt, score, highlight, status, note):
        shot_name = build_screenshot_name(sector, state, audit_year, "G1-3", page_idx)
        shot_path = save_highlighted_screenshot(
            doc, page_idx, bbox_list, highlight, screenshot_out, shot_name
        )
        result.status = status; result.score = float(score)
        result.page_idx = page_idx; result.screenshot = str(shot_path)
        result.notes = f"[{level}] target={target}  matched={matched_dt}  ({note})"
        return result

    # ---------- A. Cover / title pages (search first 10 pages) ----------
    # Many CAFRs have an image cover on page 1 and the FYE title on page 2 or 3.
    for cover_p in range(min(10, doc.page_count)):
        cover_lines = extract_page_lines(doc, cover_p)
        for ln in cover_lines:
            best = _best_date_in_line(ln.text)
            if best:
                m, y, raw = best
                if fye_matches(target, (m, y)) == "EXACT":
                    return _finalize(cover_p, [ln.bbox], f"A-Cover-p{cover_p+1}",
                                     (m, y), 100, CLR_LIGHT_GREEN, "PASS", raw)

    # ---------- B. Top 5 lines of financial-statement pages ----------
    if financial_stmt_pages:
        for p in financial_stmt_pages:
            if p >= doc.page_count:
                continue
            for ln in page_top_lines(doc, p, FYE_TOP_LINES):
                best = _best_date_in_line(ln.text)
                if best:
                    m, y, raw = best
                    if fye_matches(target, (m, y)) == "EXACT":
                        return _finalize(p, [ln.bbox], "B-Top5", (m, y), 100,
                                         CLR_LIGHT_GREEN, "PASS", raw)

    # ---------- C. Anywhere in doc; multi-date line → highest date ----------
    best_ever_match = None    # (score, page, line, matched_dt, raw, level)
    for p in range(doc.page_count):
        for ln in extract_page_lines(doc, p):
            dates = extract_dates_from_text(ln.text)
            if not dates:
                continue
            # Consider all dates; keep any EXACT (score 90 → REVIEW), else month/year
            dates.sort(key=lambda t: (t[1], t[0]), reverse=True)
            for m, y, raw in dates:
                cat = fye_matches(target, (m, y))
                if cat == "EXACT":
                    return _finalize(p, [ln.bbox], "C-Line", (m, y), 90,
                                     CLR_LIGHT_YELLOW, "REVIEW",
                                     f"line-multi-date, top={raw}")
                elif cat in ("MONTH", "YEAR"):
                    sub_score = 80 if cat == "MONTH" else 70
                    cand = (sub_score, p, ln, (m, y), raw, cat)
                    # ✅ FIX: split the fused line into a proper test + assignment
                    if best_ever_match is None or cand[0] > best_ever_match[0]:
                        best_ever_match = cand

    # ---------- C fallbacks: month-only or year-only ----------
    if best_ever_match:
        score, p, ln, dt, raw, cat = best_ever_match
        return _finalize(p, [ln.bbox], f"C-{cat}", dt, score,
                         CLR_LIGHT_RED, "REVIEW",
                         f"only-{cat.lower()}-match: {raw}")

    result.status = "FAIL"; result.score = 0.0
    result.notes  = "No fiscal-year-end date found in report."
    return result

def aggregate_gate(steps: List[StepResult]) -> str:
    """
    Aggregate a list of StepResults into a single gate verdict.

   ✅ FIX 4 — robustness:
     • NOT APPLICABLE steps are IGNORED (they must not drag a gate down —
       relevant now that Gate 3 emits NOT APPLICABLE for LG-only steps).
     • FAIL or NOT FOUND on any effective step → FAIL.
     • Any REVIEW (rest PASS) → REVIEW.
     • All effective steps PASS → PASS.
     • No effective steps, or an unrecognised status → REVIEW (needs a
       human look; never silently PASS).
    """
    statuses = [s.status for s in steps if s is not None]
    effective = [s for s in statuses if s != "NOT APPLICABLE"]
    if not effective:
        return "REVIEW"                                   # nothing to judge
    if any(s in ("FAIL", "NOT FOUND") for s in effective): return "FAIL"
    if any(s == "REVIEW" for s in effective):              return "REVIEW"
    if all(s == "PASS" for s in effective):                return "PASS"
    return "REVIEW"   # unknown status → conservative REVIEW (never silent PASS)

def _safe_step(step_code: str, fn, doc, sector, state, audit_year, screenshot_out):
    """Run a single-StepResult validator; on ANY exception return a FAIL
       StepResult (never let it bubble up as gate 'ERROR'). Returns StepResult."""
    try:
        return fn()
    except Exception as ex:
        tqdm.write(f"     ⚠️  {step_code} crashed: {type(ex).__name__}: {ex}")
        _traceback.print_exc()
        sr = StepResult(code=step_code)
        sr.status = "FAIL"; sr.score = 0.0
        sr.notes = f"Validator crashed ({type(ex).__name__}: {ex}); marked FAIL."
        try:
            sr.screenshot = take_diagnostic_screenshot(
                doc, 0, sector, state, audit_year, step_code,
                screenshot_out, status="FAIL")
        except Exception:
            pass
        return sr

def _safe_step_pages(step_code: str, fn, doc, sector, state, audit_year, screenshot_out):
    """Same as _safe_step but for validators returning (StepResult, pages_list).
       Returns (StepResult, List[int])."""
    try:
        out = fn()
        if isinstance(out, tuple):
            sr, pages = out
            return sr, (pages or [])
        return out, []
    except Exception as ex:
        tqdm.write(f"     ⚠️  {step_code} crashed: {type(ex).__name__}: {ex}")
        _traceback.print_exc()
        sr = StepResult(code=step_code)
        sr.status = "FAIL"; sr.score = 0.0
        sr.notes = f"Validator crashed ({type(ex).__name__}: {ex}); marked FAIL."
        try:
            sr.screenshot = take_diagnostic_screenshot(
                doc, 0, sector, state, audit_year, step_code,
                screenshot_out, status="FAIL")
        except Exception:
            pass
        return sr, []

def run_gate1_timed(doc, workbook_entity, workbook_state, workbook_fye,
                    sector, audit_year, screenshot_out,
                    financial_stmt_pages, step_times: Dict[str, float]
                    ) -> Tuple[str, Dict[str, StepResult]]:
    """Timed wrapper for Gate 1 — records per-step elapsed time (crash-hardened)."""
    results: Dict[str, StepResult] = {}

    t = time.perf_counter()
    g11 = _safe_step("G1-1",
        lambda: validate_g1_1_entity_name(
            doc, workbook_entity, sector, workbook_state, audit_year, screenshot_out),
        doc, sector, workbook_state, audit_year, screenshot_out)
    step_times["G1-1"] = time.perf_counter() - t
    results["G1-1"] = g11

    t = time.perf_counter()
    g12 = _safe_step("G1-2",
        lambda: validate_g1_2_state(
            doc, workbook_state, g11, sector, audit_year, screenshot_out),
        doc, sector, workbook_state, audit_year, screenshot_out)
    step_times["G1-2"] = time.perf_counter() - t
    results["G1-2"] = g12

    t = time.perf_counter()
    g13 = _safe_step("G1-3",
        lambda: validate_g1_3_fye(
            doc, workbook_fye, sector, workbook_state, audit_year,
            screenshot_out, financial_stmt_pages),
        doc, sector, workbook_state, audit_year, screenshot_out)
    step_times["G1-3"] = time.perf_counter() - t
    results["G1-3"] = g13

    gate_status = aggregate_gate([g11, g12, g13])
    return gate_status, results

def run_gate2_timed(doc, workbook_fye, auditor_name_raw,
                    sector, state, audit_year, screenshot_out,
                    step_times: Dict[str, float]
                    ) -> Tuple[str, Dict[str, StepResult]]:
    """Timed wrapper for Gate 2 — records per-step elapsed time (crash-hardened).

       ✅ FIX 1 — G2-4 (Auditor's Signature) is a HITL step:
          • never PASS (only REVIEW=found / FAIL=not found)
          • EXCLUDED from finalization → gate aggregates over G2-1..G2-3 ONLY.
          The old REVIEW→PASS promotion hack has been REMOVED; a signature
          FAIL can no longer sink Gate 2, and a REVIEW no longer needs faking.
    """
    results: Dict[str, StepResult] = {}

    # G2-1 also yields report_pages + opinion_lines reused by G2-2 / G2-4.
    t = time.perf_counter()
    try:
        g21, report_pages, opinion_lines = validate_g2_1_opinion(
            doc, sector, state, audit_year, screenshot_out)
    except Exception as ex:
        tqdm.write(f"     ⚠️  G2-1 crashed: {type(ex).__name__}: {ex}")
        _traceback.print_exc()
        g21 = StepResult(code="G2-1"); g21.status = "FAIL"; g21.score = 0.0
        g21.notes = f"Validator crashed ({type(ex).__name__}: {ex}); marked FAIL."
        report_pages, opinion_lines = [], []
    step_times["G2-1"] = time.perf_counter() - t
    results["G2-1"] = g21

    t = time.perf_counter()
    g22 = _safe_step("G2-2",
        lambda: validate_g2_2_fye_in_opinion(
            doc, workbook_fye, opinion_lines, sector, state, audit_year, screenshot_out),
        doc, sector, state, audit_year, screenshot_out)
    step_times["G2-2"] = time.perf_counter() - t
    results["G2-2"] = g22

    t = time.perf_counter()
    g23 = _safe_step("G2-3",
        lambda: validate_g2_3_auditor_name(
            doc, auditor_name_raw, sector, state, audit_year, screenshot_out),
        doc, sector, state, audit_year, screenshot_out)
    step_times["G2-3"] = time.perf_counter() - t
    results["G2-3"] = g23

    t = time.perf_counter()
    g24 = _safe_step("G2-4",
        lambda: validate_g2_4_signature(
            doc, report_pages, sector, state, audit_year, screenshot_out),
        doc, sector, state, audit_year, screenshot_out)
    step_times["G2-4"] = time.perf_counter() - t
    results["G2-4"] = g24

    # ---- GATE-2 FINALIZATION: G2-4 EXCLUDED (HITL-only) ----
    gate_status = aggregate_gate([g21, g22, g23])
    return gate_status, results

def run_gate3_timed(doc, sector, state, audit_year, workbook_fye,
                    screenshot_out, step_times: Dict[str, float]
                    ) -> Tuple[str, Dict[str, StepResult], List[int]]:
    """Timed wrapper for Gate 3 — records per-step elapsed time (crash-hardened,
       sector-aware). Mirrors the Part-6 run_gate3 plan:

         • NONLG → G3-1 (generic) + dedicated G3-2/G3-4/G3-6;
                   G3-3 & G3-5 are LG-only → NOT APPLICABLE (no crash, no blank).
         • LG    → generic G3-1..G3-5, PLUS a G3-6 probe (forced NONLG sector)
                   so BTA-style LG (community colleges / water / transit) can
                   still satisfy the NONLG-style combination in aggregate_gate3().
         • EVERY step runs under _safe_step_pages() → a crash becomes FAIL, never
           an unhandled 'ERROR' that blanks all six columns.
    """
    sector_u = sector.upper()
    heading_cache = {p: page_heading_status(doc, p) for p in range(doc.page_count)}
    candidate_pages = [p for p, s in heading_cache.items() if s != "EXCLUDED"]

    results: Dict[str, StepResult] = {}
    all_stmt_pages: List[int] = []

    def _generic(step_code: str):
        "Generic STATEMENT_SPECS validator with both sectors temporarily enabled."
        spec = STATEMENT_SPECS[step_code]
        orig_sectors = spec["sectors"]
        spec["sectors"] = {"LG", "NONLG"}
        try:
            return validate_single_statement(
                doc, step_code, sector, state, audit_year, workbook_fye,
                screenshot_out, candidate_pages=candidate_pages)
        finally:
            spec["sectors"] = orig_sectors

    if sector_u == "NONLG":
        plan = [
            ("G3-1", lambda: _generic("G3-1")),
            ("G3-2", lambda: validate_g3_2_net_position_nonlg(
                        doc, sector, state, audit_year, workbook_fye, screenshot_out)),
            ("G3-4", lambda: validate_g3_4_activities_nonlg(
                        doc, sector, state, audit_year, workbook_fye, screenshot_out)),
            ("G3-6", lambda: validate_g3_6_cash_flows(
                        doc, sector, state, audit_year, workbook_fye, screenshot_out)),
        ]
        for code, fn in plan:
            t = time.perf_counter()
            sr, pages = _safe_step_pages(code, fn, doc, sector, state,
                                         audit_year, screenshot_out)
            step_times[code] = time.perf_counter() - t
            results[code] = sr
            all_stmt_pages.extend(pages)

        # G3-3 / G3-5 are LG-only → explicit NOT APPLICABLE
        for code in ("G3-3", "G3-5"):
            sr = StepResult(code=code)
            sr.status = "NOT APPLICABLE"; sr.notes = f"{code} not applicable to NONLG"
            results[code] = sr
            step_times[code] = 0.0

    else:
        # LG: generic G3-1..G3-5
        for code in ("G3-1", "G3-2", "G3-3", "G3-4", "G3-5"):
            t = time.perf_counter()
            sr, pages = _safe_step_pages(code, (lambda c=code: _generic(c)),
                                         doc, sector, state, audit_year, screenshot_out)
            step_times[code] = time.perf_counter() - t
            results[code] = sr
            all_stmt_pages.extend(pages)

        # G3-6 probe for BTA-style LG (force NONLG sector so validator runs)
        t = time.perf_counter()
        sr6, pages6 = _safe_step_pages("G3-6",
            lambda: validate_g3_6_cash_flows(
                doc, "NONLG", state, audit_year, workbook_fye, screenshot_out),
            doc, sector, state, audit_year, screenshot_out)
        step_times["G3-6"] = time.perf_counter() - t
        results["G3-6"] = sr6
        all_stmt_pages.extend(pages6)

    gate_status = aggregate_gate3(sector, results)
    return gate_status, results, sorted(set(all_stmt_pages))

G2_HEADING_TOP_LINES = 15     # top-N lines examined for the report heading

G2_OPINION_WINDOW    = 40    # lines after heading considered "OPINION paragraph area"

G2_NAME_FUZZY_MIN    = 85    # RapidFuzz partial_ratio threshold for firm/contact match

G2_SIG_SECTION_MAX_PAGES = 6 # spec: signature may span multiple pages incl. blanks

G2_SIG_MIN_IMG_AREA  = 5000  # px² — filter out tiny logos when searching for signature

_OCR_AUD = r"a[ua]dit[o0]r"

AUDITOR_REPORT_HEADING_RE = re.compile(
    # Standard order: "Independent/State Auditor's Report"
    r"\b(?:independent|state)\s+auditor(?:['\u2019]?s|s['\u2019]?)?\s*report(?:s)?\b"
    r"|"
    # Reversed order: "Report of Independent Auditors" (Big-4 style)
    r"\breport(?:s)?\s+of\s+(?:independent|state)\s+auditor(?:s|['\u2019]?s|s['\u2019]?)?\b"
    r"|"
    # OCR-tolerant 'auditor' fragment: absorbs 'AADITOR' (U→A), 'AUDIT0R' (O→0)
    rf"\b(?:independent|state)\s+{_OCR_AUD}(?:['\u2019]?s|s['\u2019]?)?\s*report(?:s)?\b"
    r"|"
    rf"\breport(?:s)?\s+of\s+(?:independent|state)\s+{_OCR_AUD}(?:s|['\u2019]?s|s['\u2019]?)?\b"
    r"|"
    # Full CPA variant: "Report of Independent Certified Public Accountants"
    r"\breport(?:s)?\s+of\s+independent\s+"
    r"(?:registered\s+public|certified\s+public)\s+"
    r"accountant(?:s|['\u2019]?s|s['\u2019]?)?\b",
    re.IGNORECASE
)

OPINION_HEADING_RE = re.compile(r"^\s*opinion(?:s)?\b[:\s\-–—]*$", re.IGNORECASE)

OPINION_SUB_HEADING_RE = re.compile(r"^\s*opinion(?:s)?\s*$", re.IGNORECASE)

OPINION_PARAGRAPH_RE = re.compile(r"\bin\s+our\s+opinion\b", re.IGNORECASE)

OPINION_STOP_RE = re.compile(
    r"^\s*(?:"
    r"basis\s+for\b.*\bopinion(?:s)?\b"                 # Basis for [Qualified/Unmodified/…] Opinion(s)
    r"|emphasis\s+of\s+matter(?:s)?"
    r"|(?:management(?:'?s)?|responsibilit(?:y|ies))\s+.*\bfinancial\s+statement"
    r"|responsibilit(?:y|ies)\s+of\s+management"
    r"|auditor(?:'?s|s'?)?\s+responsibilit(?:y|ies)"
    r"|other\s+matter(?:s)?"
    r"|other\s+information"
    r"|report\s+on\s+other\s+legal"
    r"|other\s+reporting\s+required"
    r"|required\s+supplementary\s+information"
    r"|supplementary\s+information"
    r"|report\s+on\s+internal\s+control"
    r"|report\s+on\s+compliance"
    r")\b",
    re.IGNORECASE)

OTHER_REPORTING_RE = re.compile(
    r"other\s+report(?:ing|s)?\s+required\s+by\s+(?:the\s+)?"
    r"government\s+auditing\s+standards",
    re.IGNORECASE
)

OPINION_PASS_RE = re.compile(
    r"\bpresent(?:s|ed)?\s+fairly\b"
    r"|\btrue\s+and\s+fair\b"
    r"|\bunmodified\b"
    r"|\bunqualified\b",
    re.IGNORECASE
)

OPINION_REVIEW_RE = re.compile(
    r"\bqualified\b(?!\s+opinion\s+is\s+not)",
    re.IGNORECASE
)

OPINION_FAIL_RE = re.compile(
    r"\bdoes\s+not\s+present\s+fairly\b"
    r"|\badverse\b"
    r"|\bdisclaimer\b",
    re.IGNORECASE
)

_OPINION_SUBHEAD_RE = re.compile(
    r"^\s*(?:qualified|unmodified|unqualified|adverse|disclaimer)\s+"
    r"opinion(?:s)?\b", re.IGNORECASE)

_REGULATORY_BASIS_RE = re.compile(
    r"regulatory\s+basis(?:\s+of\s+accounting)?"
    r"|prescribed\s+or\s+permitted\s+by\s+the\s+division"
    r"|modified\s+cash\s+basis"
    r"|statutory\s+basis\s+of\s+accounting"
    r"|other\s+comprehensive\s+basis\s+of\s+accounting", re.IGNORECASE)

_GAAP_OPINION_RE = re.compile(
    r"u\.?s\.?\s+generally\s+accepted\s+accounting\s+principles"
    r"|accounting\s+principles\s+generally\s+accepted", re.IGNORECASE)

AUDITOR_SUFFIX_RE = re.compile(
    r"\b(?:llp|l\.l\.p\.|pllc|p\.l\.l\.c\.|pc|p\.c\.|pa|p\.a\.|"
    r"cpa|cpas|c\.p\.a\.|c\.p\.a\.s\.|"
    r"inc(?:orporated)?|corp(?:oration)?|company|co|ltd|limited|"
    r"a\s+professional\s+(?:accounting|corporation|association)|"
    r"chartered\s+accountants?)\b",
    re.IGNORECASE
)

def normalize_auditor_name(s: str) -> str:
    """Aggressive normalization: unicode strip → lowercase → drop suffixes →
       collapse spaces/punctuation.  Handles LLP / PC / P.A. / PLLC / CPAs etc."""
    if not isinstance(s, str) or not s.strip():
        return ""
    s = unicodedata.normalize("NFKD", s)
    s = "".join(ch for ch in s if ch.isprintable())
    s = s.lower()
    s = AUDITOR_SUFFIX_RE.sub(" ", s)
    s = re.sub(r"[^\w\s&]", " ", s)          # keep alphanumerics and &
    s = re.sub(r"\s+", " ", s).strip()
    return s

def split_firm_contact(raw: str) -> Tuple[str, str]:
    """Split '<firm> (<contact>)' back into (firm, contact).  Handles blanks."""
    if not raw:
        return "", ""
    m = re.match(r"^(?P<firm>.*?)\s*\((?P<c>[^)]+)\)\s*$", raw.strip())
    if m:
        return m.group("firm").strip(), m.group("c").strip()
    return raw.strip(), ""

_COMPLIANCE_REPORT_RE = re.compile(
    r"report\s+on\s+compliance"
    r"|for\s+each\s+major\s+federal\s+program"
    r"|report\s+on\s+internal\s+control\s+over\s+"
    r"(?:financial\s+reporting|compliance)"
    r"|uniform\s+guidance"
    r"|schedule\s+of\s+(?:findings|expenditures\s+of\s+federal\s+awards)"
    r"|single\s+audit",
    re.IGNORECASE,
)

_G2_NONOPINION_HEADING_RE = re.compile(
    r"internal\s+control"                                   # Internal Control (K-1)
    r"|\bcompliance\b"                                      # …on Compliance…
    r"|other\s+matters?"                                    # …and Other Matters (K-1)
    r"|for\s+each\s+major"                                  # …for Each Major … Program (K-2)
    r"|(?:federal|state)\s+(?:and\s+(?:federal|state)\s+)?" # Federal/State (…)
    r"(?:financial\s+assistance\s+)?program"                #   … Program (K-2)
    r"|(?:federal|state)\s+(?:financial\s+)?award"          # Federal/State Awards (K-3)
    r"|schedule\s+of\s+(?:findings|expenditures)"           # Schedules (K-3/K-6)
    r"|single\s+audit",                                     # Single-Audit section title
    re.IGNORECASE,
)

_FINSTMT_REPORT_RE = re.compile(
    r"report\s+on\s+the\s+audit\s+of\s+the\s+financial\s+statement"
    r"|audit\s+of\s+the\s+financial\s+statement"
    r"|\bopinion(?:s)?\b",
    re.IGNORECASE,
)

_TOC_RE = re.compile(r"\btable\s+of\s+content(?:s)?\b", re.IGNORECASE)

def _deglue_heading(text: str) -> str:
    """
    Normalize a heading line for matching. Handles glued lowercase→UPPERCASE
    boundaries AND space-injected extractions (e.g. "AUDITOR ' S", "E - mail").
    """
    t = (text.replace("\u00A0", " ")
             .replace("\u2019", "'")
             .replace("\u2018", "'"))
    t = re.sub(r"(?<=[a-z0-9])(?=[A-Z])", " ", t)   # de-glue camel-joins
    t = re.sub(r"\s*'\s*", "'", t)                  # ← FIX: "AUDITOR ' S" → "AUDITOR'S"
    return re.sub(r"\s+", " ", t).strip()

def find_auditor_report_pages(doc: fitz.Document) -> List:
    """
    FINANCIAL-STATEMENT auditor's report pages ONLY.
    ✅ De-glues extraction artifacts before matching (fixes 'comINDEPENDENT').
    ✅ Skips TABLE OF CONTENTS pages (they LIST the report title as an entry).
    ✅ Judges admission on the TITLE region (before the opinion body begins).
    """
    strict_pages, fallback_pages = [], []
    _FINSTMT_TITLE_RE = re.compile(
        r"report\s+on\s+the\s+audit\s+of\s+the\s+financial\s+statement(?:s)?"
        r"|audit\s+of\s+the\s+financial\s+statement(?:s)?", re.IGNORECASE)
    _BODY_START_RE = re.compile(
        r"\bwe\s+have\s+audited\b|\bin\s+our\s+opinion\b", re.IGNORECASE)

    for p in range(doc.page_count):
        top_lines = page_top_lines(doc, p, G2_HEADING_TOP_LINES)

        title_parts = []
        for ln in top_lines:
            t = _deglue_heading(ln.text)
            if _BODY_START_RE.search(t):
                break                       # opinion body starts → stop the title region
            title_parts.append(t)
        title_blob = re.sub(r"\s+", " ", " ".join(title_parts)).strip()

        # TOC guard — the contents pages enumerate 'Independent Auditor's Report'.
        if _TOC_RE.search(title_blob):
            continue
        # (1) POSITIVE
        if not AUDITOR_REPORT_HEADING_RE.search(title_blob):
            continue
        # (2) NEGATIVE — K-1..K-6
        if (_COMPLIANCE_REPORT_RE.search(title_blob) or
                _G2_NONOPINION_HEADING_RE.search(title_blob)):
            continue
        # (3) strict vs. fallback
        (strict_pages if _FINSTMT_TITLE_RE.search(title_blob)
         else fallback_pages).append(p)

    return strict_pages if strict_pages else fallback_pages

def locate_opinion_section(doc: fitz.Document, heading_page: int) -> List:
    """
    Return the block of lines representing the OPINION paragraph area.

    Tier 1: Standalone 'Opinion'/'Opinions' heading line, then collect
            following lines until a stop-heading (Basis for Opinion, etc.)
            or the line budget is exhausted.
    Tier 2: Bold 'Opinion(s)' sub-heading immediately followed by paragraph.
    Tier 3: Fallback — anchor on the 'In our opinion' sentence and collect a
            window of surrounding lines (handles firms that omit the heading).
    """
    collected: List[PageLine] = []
    line_budget = G2_OPINION_WINDOW
    max_page = min(heading_page + 3, doc.page_count)   # opinion rarely spans >3 pp

    def _norm(text: str) -> str:
        return re.sub(
            r"\s+", " ",
            text.replace("\u00A0", " ")
                .replace("\u2019", "'")
                .replace("\u2018", "'")
        ).strip()

    # ── Build a flat, ordered list of lines across the heading + continuation pages ──
    flat: List[PageLine] = []
    for p in range(heading_page, max_page):
        flat.extend(extract_page_lines(doc, p))

    # ═══ TIER 1 & 2: explicit 'Opinion(s)' heading or sub-heading ═══
    start_idx = None
    for i, ln in enumerate(flat):
        norm = _norm(ln.text)
        if OPINION_HEADING_RE.match(norm) or OPINION_SUB_HEADING_RE.match(norm):
            start_idx = i + 1           # begin collecting AFTER the heading line
            break

    if start_idx is not None:
        for ln in flat[start_idx:]:
            norm = _norm(ln.text)
            if not norm:
                continue
            if OPINION_STOP_RE.match(norm):
                break
            collected.append(ln)
            line_budget -= 1
            if line_budget <= 0:
                break
        if collected:
            return collected

    # ═══ TIER 3 (FALLBACK): anchor on 'In our opinion' sentence ═══
    para_idx = None
    for i, ln in enumerate(flat):
        if OPINION_PARAGRAPH_RE.search(_norm(ln.text)):
            para_idx = i
            break

    if para_idx is not None:
        lo = max(0, para_idx - 2)
        for ln in flat[lo:]:
            norm = _norm(ln.text)
            if not norm:
                continue
            if collected and OPINION_STOP_RE.match(norm):
                break
            collected.append(ln)
            line_budget -= 1
            if line_budget <= 0:
                break

    return collected

def select_operative_opinion_block(opinion_lines: List) -> List:
    """
    Pick the OPERATIVE opinion block for classification.

    NJ regulatory-basis reports carry TWO opinions inside one 'Opinions' section:
      • 'Unmodified Opinion(s) on Regulatory Basis of Accounting'          ← OPERATIVE
      • 'Adverse Opinion on U.S. Generally Accepted Accounting Principles' ← DISCARD
    Per spec: when the Regulatory-Basis sub-opinion is present, classify ONLY from
    that block (so the GAAP 'adverse / do not present fairly' never wins). When it
    is ABSENT (ordinary single-opinion report), return the whole 'Opinion' section.
    """
    if not opinion_lines:
        return opinion_lines

    def _norm(t: str) -> str:
        return re.sub(r"\s+", " ",
                      t.replace("\u00A0", " ").replace("\u2019", "'")).strip()

    section_blob = _norm(" ".join(l.text for l in opinion_lines))
    has_reg  = bool(_REGULATORY_BASIS_RE.search(section_blob))
    has_gaap = bool(_GAAP_OPINION_RE.search(section_blob))

    # Only a dual (regulatory + GAAP) report needs disambiguation.
    if not (has_reg and has_gaap):
        return opinion_lines

    # Discriminating sub-opinion HEADING phrases (unanchored → tolerant to gluing).
    _REG_HEAD  = re.compile(r"opinion(?:s)?\s+on\s+(?:the\s+)?regulatory\s+basis", re.IGNORECASE)
    _GAAP_HEAD = re.compile(
        r"\bopinion(?:s)?\s+on\b.{0,60}?generally\s+accepted"   # …Opinion on … Generally Accepted
        r"|\bopinion(?:s)?\s+on\b.{0,60}?accounting\s+principles" # …Opinion on … Accounting Principles
        r"|\badverse\s+opinion\b",                               # explicit 'Adverse Opinion' heading
        re.IGNORECASE)


    reg_idx  = next((i for i, ln in enumerate(opinion_lines)
                     if _REG_HEAD.search(_norm(ln.text))),  None)
    gaap_idx = next((i for i, ln in enumerate(opinion_lines)
                     if _GAAP_HEAD.search(_norm(ln.text))), None)

    # 1) Clean split on the two sub-headings.
    if reg_idx is not None and gaap_idx is not None and reg_idx != gaap_idx:
        return (opinion_lines[reg_idx:gaap_idx] if reg_idx < gaap_idx
                else opinion_lines[reg_idx:])
    if reg_idx is not None and gaap_idx is None:
        return opinion_lines[reg_idx:]

    # 2) Fallback (both opinions glued on shared lines): drop the GAAP/adverse
    #    lines so the residual regulatory 'present fairly' can be classified.
    kept = [ln for ln in opinion_lines
            if not (_GAAP_HEAD.search(_norm(ln.text))
                    or re.search(r"\badverse\b", ln.text, re.IGNORECASE)
                    or (OPINION_FAIL_RE.search(ln.text)
                        and _GAAP_OPINION_RE.search(ln.text)))]
    return kept or opinion_lines

def _opinion_label(matched_text: str) -> str:
    "Map a matched opinion phrase to a clean display label."
    t = (matched_text or "").lower()
    if "does not present fairly" in t: return "Does Not Present Fairly"
    if "present" in t and "fairly" in t: return "Present Fairly"
    if "true and fair"  in t: return "True and Fair"
    if "unmodified"     in t: return "Unmodified"
    if "unqualified"    in t: return "Unqualified"
    if "qualified"      in t: return "Qualified"
    if "adverse"        in t: return "Adverse"
    if "disclaimer"     in t: return "Disclaimer"
    return matched_text.strip().title()[:40]     # fallback

def validate_g2_1_opinion(doc: fitz.Document,
                          sector: str, state: str, audit_year,
                          screenshot_out: Path
                          ) -> Tuple[StepResult, List[int], List[PageLine]]:
    """Returns (StepResult, report_pages, opinion_section_lines).
       ✅ FIX: the 3rd return value is now the OPERATIVE opinion block (regulatory
       block for NJ dual-opinion reports), NOT the whole 'Opinions' section, so
       G2-2 checks the FYE against the SAME opinion G2-1 classified — the GAAP
       'adverse / do not present fairly' sub-opinion can never leak into either."""
    result = StepResult(code="G2-1")
    report_pages = find_auditor_report_pages(doc)
    if not report_pages:
        result.status = "FAIL"; result.notes = "Independent/State Auditor's Report heading not found."
        return result, [], []

    operative_lines_all: List[PageLine] = []   # ← operative (regulatory) block → handed to G2-2
    opinion_lines_all:   List[PageLine] = []   # ← full section (fallback only)
    hit_line: Optional[PageLine] = None
    hit_tier: Optional[str] = None    # "PASS" | "REVIEW" | "FAIL"
    hit_match: Optional[str] = None   # ← exact phrase that matched
    regulatory_basis_used = False     # ← for labeling

    for rp in report_pages:
        section = locate_opinion_section(doc, rp)

        operative = select_operative_opinion_block(section)
        if operative is not section and operative:
            regulatory_basis_used = True

        opinion_lines_all.extend(section)
        operative_lines_all.extend(operative or section)   # ← accumulate operative for G2-2

        for ln in operative:
            if (m := OPINION_FAIL_RE.search(ln.text)):
                hit_line, hit_tier, hit_match = ln, "FAIL", m.group(0); break
            if (m := OPINION_PASS_RE.search(ln.text)) and hit_tier is None:
                hit_line, hit_tier, hit_match = ln, "PASS", m.group(0)
            elif (m := OPINION_REVIEW_RE.search(ln.text)) and hit_tier not in ("PASS", "FAIL"):
                hit_line, hit_tier, hit_match = ln, "REVIEW", m.group(0)
        if hit_tier == "FAIL":
            break
        if hit_tier == "PASS":
            break

    # G2-2 should inherit the operative block; fall back to full section if empty.
    g2_2_lines = operative_lines_all or opinion_lines_all

    if hit_line is None or hit_tier is None:
        p = report_pages[0]
        shot_name = build_screenshot_name(sector, state, audit_year, "G2-1", p)
        shot_path = save_highlighted_screenshot(
            doc, p, [], CLR_LIGHT_RED, screenshot_out, shot_name
        )
        result.status = "FAIL"; result.score = 0.0
        result.page_idx = p; result.screenshot = str(shot_path)
        result.notes = "Opinion paragraph found but no recognizable opinion keyword."
        return result, report_pages, g2_2_lines

    tier_map = {
        "PASS":   (100.0, CLR_LIGHT_GREEN,  "PASS"),
        "REVIEW": ( 90.0, CLR_LIGHT_YELLOW, "REVIEW"),
        "FAIL":   ( 50.0, CLR_LIGHT_RED,    "FAIL"),
    }
    score, highlight, status = tier_map[hit_tier]
    shot_name = build_screenshot_name(sector, state, audit_year, "G2-1", hit_line.page_idx)
    shot_path = save_highlighted_screenshot(
        doc, hit_line.page_idx, [hit_line.bbox], highlight, screenshot_out, shot_name
    )
    result.status = status; result.score = score
    result.page_idx = hit_line.page_idx; result.screenshot = str(shot_path)

    label = _opinion_label(hit_match)
    if regulatory_basis_used and label:
        label = f"{label} - Regulatory Basis"     # e.g., "Present Fairly - Regulatory Basis"
    result.classification = label

    note_prefix = "Opinion (regulatory basis) " if regulatory_basis_used else "Opinion "
    result.notes = f"{note_prefix}tier={hit_tier}: {hit_line.text[:120]!r}"
    return result, report_pages, g2_2_lines     # ← operative block flows to G2-2

def validate_g2_2_fye_in_opinion(doc: fitz.Document,
                                 workbook_fye,
                                 opinion_lines: List[PageLine],
                                 sector: str, state: str, audit_year,
                                 screenshot_out: Path) -> StepResult:
    """Search the (operative) OPINION block for the workbook FYE month+year.
       PRIORITY: EXACT match always wins over partial.
       ✅ FIX: adds a CROSS-LINE combined-text EXACT pass so an FYE split across a
       PDF line break (e.g. 'December 31,' / '2025') is still caught. Safe now that
       G2-1 hands G2-2 the operative regulatory block (no stray report/sig dates)."""
    result = StepResult(code="G2-2")
    target = parse_fye_from_workbook(workbook_fye)
    if not target:
        result.status = "FAIL"; result.notes = f"Unparseable workbook FYE: {workbook_fye!r}"
        return result
    if not opinion_lines:
        result.status = "FAIL"; result.notes = "No opinion section available (G2-1 failed)."
        return result

    if DEBUG_FYE:
        tqdm.write(f"     [G2-2 debug] target FYE: month={target[0]}, year={target[1]}")

    def _pick_line_for(year: int, month_hint: str = "") -> PageLine:
        "Best line to highlight when the match came from combined text."
        yr = str(year)
        for ln in opinion_lines:
            if yr in ln.text and (not month_hint or month_hint.lower() in ln.text.lower()):
                return ln
        for ln in opinion_lines:                     # relax: year only
            if yr in ln.text:
                return ln
        return opinion_lines[0]

    # ---- PASS 1a: line-level EXACT (preferred — gives a precise bbox) ----
    exact_hit = None
    for ln in opinion_lines:
        for m, y, raw in extract_dates_from_text(ln.text):
            if fye_matches(target, (m, y)) == "EXACT":
                exact_hit = (ln, raw); break
        if exact_hit:
            break

    # ---- PASS 1b: cross-line EXACT (combined text) — catches wrapped dates ----
    if not exact_hit:
        combined = " ".join(
            l.text.replace("\u00A0", " ") for l in opinion_lines)
        combined = re.sub(r"\s+", " ", combined)
        for m, y, raw in extract_dates_from_text(combined):
            if fye_matches(target, (m, y)) == "EXACT":
                mon_hint = raw.split()[0] if raw and raw[0].isalpha() else ""
                exact_hit = (_pick_line_for(y, mon_hint), raw)
                if DEBUG_FYE:
                    tqdm.write(f"     [G2-2 debug] cross-line EXACT: '{raw}'")
                break

    if exact_hit:
        ln, raw = exact_hit
        if DEBUG_FYE:
            tqdm.write(f"     [G2-2 debug] EXACT match: '{raw}' on page {ln.page_idx+1}")
        shot_name = build_screenshot_name(sector, state, audit_year, "G2-2", ln.page_idx)
        shot_path = save_highlighted_screenshot(
            doc, ln.page_idx, [ln.bbox], CLR_LIGHT_GREEN, screenshot_out, shot_name
        )
        result.status = "PASS"; result.score = 100.0
        result.page_idx = ln.page_idx; result.screenshot = str(shot_path)
        result.notes = f"FYE matched exactly in opinion: {raw}"
        return result

    # ---- PASS 2: no EXACT → partial (MONTH or YEAR only) ----
    other_hit = None
    all_partials = []
    for ln in opinion_lines:
        for m, y, raw in extract_dates_from_text(ln.text):
            cat = fye_matches(target, (m, y))
            if cat in ("MONTH", "YEAR"):
                all_partials.append((ln, raw, cat, m, y))
                if other_hit is None:
                    other_hit = (ln, raw, cat)

    if DEBUG_FYE:
        if all_partials:
            tqdm.write(f"     [G2-2 debug] {len(all_partials)} partial matches found:")
            for ln, raw, cat, m, y in all_partials[:5]:
                tqdm.write(f"                  - '{raw}' → ({m},{y}) [{cat}] on page {ln.page_idx+1}")
        else:
            tqdm.write(f"     [G2-2 debug] no partial matches; target month={target[0]}, year={target[1]}")

    if other_hit:
        ln, raw, cat = other_hit
        shot_name = build_screenshot_name(sector, state, audit_year, "G2-2", ln.page_idx)
        shot_path = save_highlighted_screenshot(
            doc, ln.page_idx, [ln.bbox], CLR_LIGHT_YELLOW, screenshot_out, shot_name
        )
        result.status = "REVIEW"; result.score = 80.0
        result.page_idx = ln.page_idx; result.screenshot = str(shot_path)
        result.notes = f"Only partial date match (target={target}, found={raw}, cat={cat})"
        return result

    result.status = "FAIL"; result.score = 0.0
    result.notes = "No date found in the opinion paragraph/section."
    return result

def validate_g2_3_auditor_name(doc: fitz.Document,
                               auditor_name_raw: str,
                               sector: str, state: str, audit_year,
                               screenshot_out: Path) -> StepResult:
    """Search entire PDF for the auditor firm and/or contact name.
       Uses normalized firm-name matching + RapidFuzz partial_ratio fallback."""
    result = StepResult(code="G2-3")
    firm, contact = split_firm_contact(auditor_name_raw)
    firm_norm    = normalize_auditor_name(firm)
    contact_norm = normalize_auditor_name(contact)
    if not firm_norm and not contact_norm:
        result.status = "FAIL"; result.notes = "No auditor name captured during sourcing."
        return result

    best_hit: Optional[Tuple[int, PageLine, int, str]] = None   # (score, line, page, matched)

    for p in range(doc.page_count):
        for ln in extract_page_lines(doc, p):
            line_norm = normalize_auditor_name(ln.text)
            if not line_norm:
                continue

            # 1) Direct substring match on firm
            if firm_norm and firm_norm in line_norm:
                best_hit = (100, ln, p, firm); break

            # 2) Fuzzy firm match
            if firm_norm:
                fscore = fuzz.partial_ratio(firm_norm, line_norm)
                if fscore >= G2_NAME_FUZZY_MIN:
                    if not best_hit or fscore > best_hit:
                        best_hit = (fscore, ln, p, firm)

            # 3) Contact fallback (only if firm not matched at 100 yet)
            if contact_norm and (not best_hit or best_hit[0] < 100):
                cscore = fuzz.partial_ratio(contact_norm, line_norm)
                if cscore >= G2_NAME_FUZZY_MIN:
                    if not best_hit or cscore > best_hit:
                        best_hit = (cscore, ln, p, contact)

        if best_hit and best_hit[0] == 100:
            break

    if not best_hit:
        result.status = "FAIL"; result.score = 0.0
        result.notes = f"Auditor '{firm or contact}' not found anywhere in PDF."
        return result

    score, ln, p, matched = best_hit
    shot_name = build_screenshot_name(sector, state, audit_year, "G2-3", p)
    shot_path = save_highlighted_screenshot(
        doc, p, [ln.bbox], CLR_LIGHT_GREEN, screenshot_out, shot_name
    )
    result.status = "PASS"; result.score = float(score)
    result.page_idx = p; result.screenshot = str(shot_path)
    result.notes = f"Matched '{matched}' at page {p+1} (fuzzy={score}%)"
    return result

def _find_other_reporting_anchor(doc: fitz.Document,
                                 report_pages: List[int]
                                 ) -> Optional[Tuple[int, PageLine]]:
    """Locate the 'Other Reporting Required by Government Auditing Standards' line."""
    start_page = min(report_pages) if report_pages else 0
    for p in range(start_page, doc.page_count):
        for ln in extract_page_lines(doc, p):
            if OTHER_REPORTING_RE.search(ln.text):
                return (p, ln)
    return None

def _page_has_signature_image(doc: fitz.Document, page_idx: int
                              ) -> Optional[Tuple[float, float, float, float]]:
    """Return the bbox of the first non-trivial image on the page, or None."""
    page = doc[page_idx]
    try:
        try:
            img_info = page.get_image_info(xrefs=True) or []
        except TypeError:
            img_info = page.get_image_info() or []
    except Exception:
        img_info = []
    for info in img_info:
        bbox = info.get("bbox")
        if not bbox:
            continue
        w = bbox[2] - bbox[0]; h = bbox[3] - bbox[1]
        if (w * h) >= G2_SIG_MIN_IMG_AREA:
            return tuple(bbox)
    return None

SIGNATURE_FONT_KEYWORDS = [
    "italic", "script", "handwrit", "signature",
    "corsiva", "vivaldi", "kunstler", "brush", "autograph",
    "signer", "edwardian", "bickham", "chancery", "monotype corsiva",
    "segoe script", "lucida hand", "rage", "mistral", "bradley hand",
    "zapfino", "cursive", "handwriting", "chalkboard",
    "coronet", "monotype", "palace script", "shelley", "vladimir",
    "sign", "curs", "hand",
]

def _page_has_signature_font(doc: fitz.Document, page_idx: int
                              ) -> Optional[Tuple[float, float, float, float]]:
    """Heuristic: detect a line rendered in a script/italic/cursive font (signature-style)."""
    try:
        d = doc[page_idx].get_text("dict")
    except Exception:
        return None

    _EXCLUDE_ITALIC_PHRASES = [
        "government auditing standards",
        "code of federal regulations",
        "uniform administrative requirements",
        "office of management and budget",
        "office of the comptroller",
        "see accompanying notes",
        "see notes to",
        "in thousands",
        "in millions",
    ]

    for block in d.get("blocks", []):
        for line in block.get("lines", []):
            for span in line.get("spans", []):
                font_name = (span.get("font") or "").lower()
                if any(k in font_name for k in SIGNATURE_FONT_KEYWORDS):
                    txt = span.get("text", "").strip()
                    txt_lower = txt.lower()
                    if 3 <= len(txt) <= 80 and not any(
                            phrase in txt_lower for phrase in _EXCLUDE_ITALIC_PHRASES
                    ):
                        return tuple(span["bbox"])
    return None

_SIG_STOP_HEADING_RE = re.compile(
    r"management(?:['\u2019]?s)?\s+discussion\s+and\s+analysis"
    r"|md\s?&\s?a"
    r"|basic\s+financial\s+statement"
    r"|statement\s+of\s+net\s+position"
    r"|balance\s+sheet"
    r"|notes?\s+to\s+the\s+(?:basic\s+)?financial\s+statement"
    r"|table\s+of\s+content"
    r"|required\s+supplementary\s+information",
    re.IGNORECASE,
)

def _is_report_section_boundary(doc: fitz.Document, page_idx: int) -> bool:
    "True if this page's top lines start a NEW (non-report) section."
    for ln in page_top_lines(doc, page_idx, G2_HEADING_TOP_LINES):
        norm = ln.text.replace("\u00A0", " ").replace("\u2019", "'")
        if _SIG_STOP_HEADING_RE.search(norm):
            return True
    return False

def validate_g2_4_signature(doc: fitz.Document,
                            report_pages: List[int],
                            sector: str, state: str, audit_year,
                            screenshot_out: Path) -> StepResult:
    """
    Locate the auditor's signature starting AT the 'Other Reporting' anchor.
    PAGE-major within each tier; NEVER crosses a report-section boundary.

    HITL RULE (per spec): signature can NEVER be PASS.
      • Found (image OR script-font OR firm/City-State sign-off) → REVIEW
      • Nothing found in the signature area                      → FAIL
    """
    result = StepResult(code="G2-4")
    anchor = _find_other_reporting_anchor(doc, report_pages)
    start_page = anchor[0] if anchor else (report_pages[0] if report_pages else 0)
    end_page = min(start_page + G2_SIG_SECTION_MAX_PAGES, doc.page_count)

    _STATE_FULL = (
        r"Alabama|Alaska|Arizona|Arkansas|California|Colorado|Connecticut|"
        r"Delaware|Florida|Georgia|Hawaii|Idaho|Illinois|Indiana|Iowa|Kansas|"
        r"Kentucky|Louisiana|Maine|Maryland|Massachusetts|Michigan|Minnesota|"
        r"Mississippi|Missouri|Montana|Nebraska|Nevada|New\s+Hampshire|"
        r"New\s+Jersey|New\s+Mexico|New\s+York|North\s+Carolina|North\s+Dakota|"
        r"Ohio|Oklahoma|Oregon|Pennsylvania|Rhode\s+Island|South\s+Carolina|"
        r"South\s+Dakota|Tennessee|Texas|Utah|Vermont|Virginia|Washington|"
        r"West\s+Virginia|Wisconsin|Wyoming"
    )
    _STATE_CODE = (
        r"AL|AK|AZ|AR|CA|CO|CT|DE|FL|GA|HI|ID|IL|IN|IA|KS|KY|LA|ME|MD|MA|MI|"
        r"MN|MS|MO|MT|NE|NV|NH|NJ|NM|NY|NC|ND|OH|OK|OR|PA|RI|SC|SD|TN|TX|UT|"
        r"VT|VA|WA|WV|WI|WY"
    )
    _FIRM_SUFFIX = (
        r"llp|l\.?l\.?p|llc|l\.?l\.?c|pllc|p\.?l\.?l\.?c|"
        r"pa|p\.?a|pc|p\.?\s?c|cpas?|"
        r"professional\s+(?:association|corporation)|corporation|inc"
    )

    firm_city_pattern = re.compile(
        rf"\b(?:{_FIRM_SUFFIX})\b[\s\S]{{0,80}}?"
        r"[A-Z][A-Za-z.\-]+(?:\s+[A-Z][A-Za-z.\-]+)*,\s+"
        rf"(?:{_STATE_FULL}|{_STATE_CODE})\b",
        re.IGNORECASE,
    )
    signoff_full_re = re.compile(
        r"\b[A-Z][A-Za-z.\-]+(?:\s+[A-Z][A-Za-z.\-]+)*,\s+"
        rf"(?:{_STATE_FULL})\b",
        re.IGNORECASE,
    )
    signoff_code_re = re.compile(
        r"\b[A-Z][A-Za-z.\-]+(?:\s+[A-Z][A-Za-z.\-]+)*,\s+"
        rf"(?:{_STATE_CODE})\b(?=\s|,|\.|$)"
    )

    # ── TIER 1: signature image ──
    for pp in range(start_page, end_page):
        if pp > start_page and _is_report_section_boundary(doc, pp):
            break
        img_bbox = _page_has_signature_image(doc, pp)
        if img_bbox:
            try:
                shot_name = build_screenshot_name(sector, state, audit_year, "G2-4", pp)
                shot = str(save_highlighted_screenshot(
                    doc, pp, [img_bbox], CLR_LIGHT_YELLOW, screenshot_out, shot_name))
            except Exception:
                shot = None
            result.status = "REVIEW"; result.score = 90.0
            result.page_idx = pp; result.screenshot = shot
            result.notes = f"Signature image detected at bbox {img_bbox} on page {pp+1}"
            return result

    # ── TIER 2: script/cursive font ──
    for pp in range(start_page, end_page):
        if pp > start_page and _is_report_section_boundary(doc, pp):
            break
        font_bbox = _page_has_signature_font(doc, pp)
        if font_bbox:
            try:
                shot_name = build_screenshot_name(sector, state, audit_year, "G2-4", pp)
                shot = str(save_highlighted_screenshot(
                    doc, pp, [font_bbox], CLR_LIGHT_YELLOW, screenshot_out, shot_name))
            except Exception:
                shot = None
            result.status = "REVIEW"; result.score = 90.0
            result.page_idx = pp; result.screenshot = shot
            result.notes = f"Signature-style font detected at bbox {font_bbox} on page {pp+1}"
            return result

    # ── TIER 3a: firm name + City, State ──
    for pp in range(start_page, end_page):
        if pp > start_page and _is_report_section_boundary(doc, pp):
            break
        try:
            page_text = "\n".join(ln.text for ln in extract_page_lines(doc, pp))
        except Exception:
            continue
        if firm_city_pattern.search(page_text):
            try:
                shot = take_diagnostic_screenshot(
                    doc, pp, sector, state, audit_year, "G2-4",
                    screenshot_out, status="REVIEW")
            except Exception:
                shot = None
            result.status = "REVIEW"; result.score = 90.0
            result.page_idx = pp; result.screenshot = shot
            result.notes = f"Firm signature block (firm name + City, State) on page {pp+1}"
            return result

    # ── TIER 3b: bare City, State sign-off ──
    for pp in range(start_page, end_page):
        if pp > start_page and _is_report_section_boundary(doc, pp):
            break
        try:
            page_text = "\n".join(ln.text for ln in extract_page_lines(doc, pp))
        except Exception:
            continue
        if signoff_full_re.search(page_text) or signoff_code_re.search(page_text):
            try:
                shot = take_diagnostic_screenshot(
                    doc, pp, sector, state, audit_year, "G2-4",
                    screenshot_out, status="REVIEW")
            except Exception:
                shot = None
            result.status = "REVIEW"; result.score = 90.0
            result.page_idx = pp; result.screenshot = shot
            result.notes = (f"Auditor sign-off block (City, State) detected on "
                            f"page {pp+1} — firm signs with graphic/location block.")
            return result

    # ── No signature found via any tier → FAIL (never PASS) ──
    fallback_page = start_page if (0 <= start_page < doc.page_count) else 0
    try:
        shot = take_diagnostic_screenshot(
            doc, fallback_page, sector, state, audit_year,
            "G2-4", screenshot_out, status="FAIL")
    except Exception:
        shot = None
    result.status = "FAIL"; result.score = 0.0
    result.page_idx = fallback_page; result.screenshot = shot
    result.notes = ("No signature detected: no image, no script/cursive font, "
                    "and no firm-city or City/State sign-off block in the "
                    "auditor's report signature area.")
    return result

FRAMEWORK_MAX_PAGES  = 150      # scan first N pages (audits typically declare framework early)

FRAMEWORK_SECONDARY_MAX_PAGES = 250  # for extremely long audits (>200 pages)

FRAMEWORK_MIN_SCORE  = 2       # minimum weighted score to classify

FRAMEWORK_HIGH_CONF  = 8       # score ≥ this = high confidence (90-100%)

FRAMEWORK_MED_CONF   = 4       # score ≥ this = medium confidence (70-89%)

FRAMEWORK_PATTERNS = {
    "GASB": [
        (re.compile(r"\bgovernmental\s+accounting\s+standards\s+board\b", re.IGNORECASE), 10),
        (re.compile(r"\bGASB\s+statement\s+no\.?\s+\d+", re.IGNORECASE), 8),
        (re.compile(r"\bGASB\b", re.IGNORECASE), 5),
        (re.compile(r"\bgovernment-?wide\s+financial\s+statement", re.IGNORECASE), 6),
        (re.compile(r"\bgovernmental\s+activit(?:y|ies)\b", re.IGNORECASE), 4),
        (re.compile(r"\bbusiness-?type\s+activit(?:y|ies)\b", re.IGNORECASE), 4),
        (re.compile(r"\bstatement\s+of\s+net\s+position\b", re.IGNORECASE), 3),
        (re.compile(r"\bcomprehensive\s+annual\s+financial\s+report\b", re.IGNORECASE), 3),
        (re.compile(r"\bannual\s+comprehensive\s+financial\s+report\b", re.IGNORECASE), 3),
        (re.compile(r"\bmodified\s+accrual\s+basis\b", re.IGNORECASE), 3),
        (re.compile(r"\bfund\s+balance(?:s)?\b", re.IGNORECASE), 2),
        (re.compile(r"\bgeneral\s+fund\b", re.IGNORECASE), 2),
        (re.compile(r"\bspecial\s+revenue\s+fund", re.IGNORECASE), 2),
        (re.compile(r"\bcapital\s+projects?\s+fund", re.IGNORECASE), 2),
        (re.compile(r"\bdebt\s+service\s+fund", re.IGNORECASE), 2),
        (re.compile(r"\bproprietary\s+fund", re.IGNORECASE), 2),
    ],
    "FASB": [
        (re.compile(r"\bfinancial\s+accounting\s+standards\s+board\b", re.IGNORECASE), 10),
        (re.compile(r"\bFASB\s+ASC\s+958\b", re.IGNORECASE), 10),
        (re.compile(r"\bASC\s+958\b", re.IGNORECASE), 8),
        (re.compile(r"\bFASB\b", re.IGNORECASE), 5),
        (re.compile(r"\bnet\s+assets?\s+without\s+donor\s+restrictions?\b", re.IGNORECASE), 8),
        (re.compile(r"\bnet\s+assets?\s+with\s+donor\s+restrictions?\b", re.IGNORECASE), 8),
        (re.compile(r"\bstatement\s+of\s+financial\s+position\b", re.IGNORECASE), 4),
        (re.compile(r"\bunrestricted\s+net\s+assets?\b", re.IGNORECASE), 3),
        (re.compile(r"\bFASB\s+ASU\b", re.IGNORECASE), 8),
        (re.compile(r"\bFASB'?s\s+guidance\b", re.IGNORECASE), 6),
        (re.compile(r"\bAccounting\s+Standards\s+Codification\b", re.IGNORECASE), 6),
        (re.compile(r"\bdonor-restricted\s+contributions?\b", re.IGNORECASE), 4),
        (re.compile(r"\b(?:consolidated|combined)\s+balance\s+sheets?\b", re.IGNORECASE), 3),
        (re.compile(r"\b(?:consolidated|combined)\s+statement(?:s)?\s+of\s+operations\s+"
                    r"and\s+changes\s+in\s+net\s+assets?\b", re.IGNORECASE), 5),
        (re.compile(r"\bchanges?\s+in\s+net\s+assets?\b", re.IGNORECASE), 2),
        (re.compile(r"\bASC\s+606\b", re.IGNORECASE), 3),
        (re.compile(r"\bASC\s+820\b", re.IGNORECASE), 3),
        (re.compile(r"\bASC\s+350\b", re.IGNORECASE), 2),
        (re.compile(r"\bASC\s+715\b", re.IGNORECASE), 2),
        (re.compile(r"\bASC\s+740\b", re.IGNORECASE), 2),
        (re.compile(r"\bASC\s+815\b", re.IGNORECASE), 2),
        (re.compile(r"\bASC\s+230\b", re.IGNORECASE), 2),
        (re.compile(r"\bSection\s+501\s*\(\s*c\s*\)\s*\(\s*3\s*\)\b", re.IGNORECASE), 5),
        (re.compile(r"\bnonprofit\s+organization", re.IGNORECASE), 2),
        (re.compile(r"\bnot-?for-?profit\s+(?:corporation|organization|entity)\b", re.IGNORECASE), 3),
    ],
}

BASIS_PATTERNS = {
    "ACCRUAL": [
        (re.compile(r"\bfull\s+accrual\s+basis(?:\s+of\s+accounting)?\b", re.IGNORECASE), 10),
        (re.compile(r"\bfull\s+accrual\s+method(?:\s+of\s+accounting)?\b", re.IGNORECASE), 10),
        (re.compile(r"\bfull\s+accrual\s+accounting\b", re.IGNORECASE), 9),
        (re.compile(r"\baccrual\s+basis\s+of\s+accounting\b", re.IGNORECASE), 8),
        (re.compile(r"\baccrual\s+method\s+of\s+accounting\b", re.IGNORECASE), 8),
        (re.compile(r"\bmodified\s+accrual\s+basis\b", re.IGNORECASE), -20),  # suppress
        (re.compile(r"\bprepared\s+(?:on|using|under)\s+(?:the\s+|a\s+)?"
                    r"(?:full\s+)?accrual\s+(?:basis|method)\b", re.IGNORECASE), 8),
        (re.compile(r"\breported\s+(?:on|using|under)\s+(?:the\s+|a\s+)?"
                    r"(?:full\s+)?accrual\s+(?:basis|method)\b", re.IGNORECASE), 8),
        (re.compile(r"\bmeasured\s+(?:on|using|under)\s+(?:the\s+|a\s+)?"
                    r"(?:full\s+)?accrual\s+(?:basis|method)\b", re.IGNORECASE), 8),
        (re.compile(r"\butilizing\s+(?:the\s+|a\s+)?(?:full\s+)?accrual\b", re.IGNORECASE), 6),
        (re.compile(r"\bfollowing\s+the\s+(?:full\s+)?accrual\s+(?:basis|method)\b", re.IGNORECASE), 7),
        (re.compile(r"\beconomic\s+resources\s+measurement\s+focus\b", re.IGNORECASE), 7),
        (re.compile(r"\bflow\s+of\s+economic\s+resources\b", re.IGNORECASE), 6),
        (re.compile(r"\beconomic\s+resources\s+basis\b", re.IGNORECASE), 5),
        (re.compile(r"\brevenues?\s+are\s+recognized\s+when\s+earned\b", re.IGNORECASE), 5),
        (re.compile(r"\brevenues?\s+are\s+recognized\s+in\s+the\s+period\s+earned\b", re.IGNORECASE), 5),
        (re.compile(r"\bexpenses?\s+are\s+recognized\s+when\s+incurred\b", re.IGNORECASE), 5),
        (re.compile(r"\bexpenses?\s+are\s+(?:recorded|recognized)\s+"
                    r"in\s+the\s+period\s+(?:they\s+are\s+)?incurred\b", re.IGNORECASE), 5),
        (re.compile(r"\baccrual[-\s]based\b", re.IGNORECASE), 4),
        (re.compile(r"\bon\s+an?\s+accrual\s+basis\b", re.IGNORECASE), 5),
    ],
    "MODIFIED ACCRUAL": [
        (re.compile(r"\bmodified\s+accrual\s+basis(?:\s+of\s+accounting)?\b", re.IGNORECASE), 12),
        (re.compile(r"\bcurrent\s+financial\s+resources\s+measurement\s+focus\b", re.IGNORECASE), 8),
        (re.compile(r"\brevenues?\s+are\s+recognized\s+when\s+measurable\s+and\s+available\b", re.IGNORECASE), 6),
    ],
    "CASH": [
        (re.compile(r"\bcash\s+basis\s+of\s+accounting\b", re.IGNORECASE), 10),
        (re.compile(r"\bstatement\s+of\s+cash\s+receipts?\s+and\s+disbursements?\b", re.IGNORECASE), 8),
        (re.compile(r"\brevenues?\s+are\s+recognized\s+when\s+(?:cash\s+is\s+)?received\b", re.IGNORECASE), 6),
        (re.compile(r"\bexpenditures?\s+are\s+recorded\s+when\s+paid\b", re.IGNORECASE), 6),
        (re.compile(r"\bmodified\s+cash\s+basis\b", re.IGNORECASE), -20),
    ],
    "MODIFIED CASH": [
        (re.compile(r"\bmodified\s+cash\s+basis(?:\s+of\s+accounting)?\b", re.IGNORECASE), 12),
        (re.compile(r"\bmodified\s+cash\s+(?:method|method\s+of\s+accounting)\b", re.IGNORECASE), 10),
        (re.compile(r"\bregulatory\s+basis\s+of\s+accounting\b", re.IGNORECASE), 8),
        (re.compile(r"\bother\s+comprehensive\s+basis\s+of\s+accounting\b", re.IGNORECASE), 7),
        (re.compile(r"\bOCBOA\b", re.IGNORECASE), 6),
    ],
}

def detect_framework_and_basis(doc: fitz.Document,
                                 max_pages: int = FRAMEWORK_MAX_PAGES,
                                 verbose: bool = False
                                 ) -> Dict[str, Any]:
    """Detect the accounting framework and basis-of-accounting for an audit report."""
    def _aggregate_text(page_start: int, page_end: int) -> str:
        parts = []
        for p in range(page_start, min(page_end, doc.page_count)):
            try:
                text = doc[p].get_text("text")
                text = (text.replace("\u00A0", " ")
                            .replace("\u2019", "'")
                            .replace("\u2018", "'"))
                text = re.sub(r"\s+", " ", text)
                parts.append(text)
            except Exception:
                continue
        return " ".join(parts)

    primary_text = _aggregate_text(0, max_pages)
    full_text = primary_text

    def _score_category(pattern_list, text):
        total = 0
        matches = []
        for rx, weight in pattern_list:
            found = rx.findall(text)
            if found:
                capped_count = min(len(found), 3)
                total += weight * capped_count
                if verbose:
                    matches.append((rx.pattern[:70], len(found), weight))
        return total, matches

    framework_scores = {}
    framework_evidence = {}
    for category, patterns in FRAMEWORK_PATTERNS.items():
        score, matches = _score_category(patterns, full_text)
        if score > 0:
            framework_scores[category] = score
            framework_evidence[category] = matches

    basis_scores = {}
    basis_evidence = {}
    for category, patterns in BASIS_PATTERNS.items():
        score, matches = _score_category(patterns, full_text)
        if score > 0:
            basis_scores[category] = score
            basis_evidence[category] = matches

    best_framework = max(framework_scores, key=framework_scores.get) if framework_scores else "UNCLASSIFIED"
    best_framework_score = framework_scores.get(best_framework, 0)
    if best_framework_score < FRAMEWORK_MIN_SCORE:
        best_framework = "UNCLASSIFIED"; best_framework_score = 0

    best_basis = max(basis_scores, key=basis_scores.get) if basis_scores else "UNCLASSIFIED"
    best_basis_score = basis_scores.get(best_basis, 0)
    if best_basis_score < FRAMEWORK_MIN_SCORE:
        best_basis = "UNCLASSIFIED"; best_basis_score = 0

    # ── SECONDARY SCAN ──
    if (best_framework == "UNCLASSIFIED" or best_basis == "UNCLASSIFIED") \
            and doc.page_count > max_pages:
        secondary_text = _aggregate_text(max_pages, FRAMEWORK_SECONDARY_MAX_PAGES)
        if secondary_text:
            full_text = primary_text + " " + secondary_text
            if verbose:
                tqdm.write(f"     🔎 Framework primary scan gave "
                           f"({best_framework}, {best_basis}) — expanding to secondary scan "
                           f"({max_pages+1} to {min(FRAMEWORK_SECONDARY_MAX_PAGES, doc.page_count)})")

            framework_scores = {}
            framework_evidence = {}
            for category, patterns in FRAMEWORK_PATTERNS.items():
                score, matches = _score_category(patterns, full_text)
                if score > 0:
                    framework_scores[category] = score
                    framework_evidence[category] = matches

            basis_scores = {}
            basis_evidence = {}
            for category, patterns in BASIS_PATTERNS.items():
                score, matches = _score_category(patterns, full_text)
                if score > 0:
                    basis_scores[category] = score
                    basis_evidence[category] = matches

            best_framework = max(framework_scores, key=framework_scores.get) if framework_scores else "UNCLASSIFIED"
            best_framework_score = framework_scores.get(best_framework, 0)
            if best_framework_score < FRAMEWORK_MIN_SCORE:
                best_framework = "UNCLASSIFIED"; best_framework_score = 0

            best_basis = max(basis_scores, key=basis_scores.get) if basis_scores else "UNCLASSIFIED"
            best_basis_score = basis_scores.get(best_basis, 0)
            if best_basis_score < FRAMEWORK_MIN_SCORE:
                best_basis = "UNCLASSIFIED"; best_basis_score = 0

    combined_score = best_framework_score + best_basis_score
    if combined_score >= FRAMEWORK_HIGH_CONF * 2:
        confidence = "HIGH"
        confidence_pct = min(100, 90 + (combined_score - FRAMEWORK_HIGH_CONF * 2) // 4)
    elif combined_score >= FRAMEWORK_MED_CONF * 2:
        confidence = "MEDIUM"
        confidence_pct = 70 + (combined_score - FRAMEWORK_MED_CONF * 2) * 20 \
                              // max(1, FRAMEWORK_HIGH_CONF * 2 - FRAMEWORK_MED_CONF * 2)
    elif combined_score >= FRAMEWORK_MIN_SCORE * 2:
        confidence = "LOW"
        confidence_pct = 40 + (combined_score - FRAMEWORK_MIN_SCORE * 2) * 30 \
                              // max(1, FRAMEWORK_MED_CONF * 2 - FRAMEWORK_MIN_SCORE * 2)
    else:
        confidence = "UNCERTAIN"
        confidence_pct = min(39, combined_score * 10)

    if best_framework != "UNCLASSIFIED" and best_basis != "UNCLASSIFIED":
        display_text = f"{best_framework} - {best_basis}"
    elif best_framework != "UNCLASSIFIED":
        display_text = f"{best_framework} - UNCLASSIFIED BASIS"
    elif best_basis != "UNCLASSIFIED":
        display_text = f"UNCLASSIFIED FRAMEWORK - {best_basis}"
    else:
        display_text = "UNCLASSIFIED"

    if verbose:
        tqdm.write(f"     🏛  Framework scores: {framework_scores}")
        tqdm.write(f"     🏛  Basis scores:     {basis_scores}")
        if framework_evidence.get(best_framework):
            tqdm.write(f"     🏛  Framework top matches: {framework_evidence[best_framework][:3]}")
        if basis_evidence.get(best_basis):
            tqdm.write(f"     🏛  Basis top matches:     {basis_evidence[best_basis][:3]}")

    return {
        "framework":       best_framework,
        "framework_score": best_framework_score,
        "basis":           best_basis,
        "basis_score":     best_basis_score,
        "confidence":      confidence,
        "confidence_pct":  confidence_pct,
        "display_text":    display_text,
        "evidence": {
            "framework": framework_evidence,
            "basis":     basis_evidence,
        },
    }

G3_HEADING_TOP_LINES  = 8     # top-N lines examined for page heading

G3_TITLE_SPAN_MAX     = 3     # statement name may continue up to 3 consecutive lines

G3_ROW_SEARCH_WINDOW  = 60    # lines below title inspected for row/column labels

G3_STMT_MAX_CONT_PAGES = 6    # a single statement may span up to N pages

G3_CONT_MAX_PAGES      = G3_STMT_MAX_CONT_PAGES   # alias (kept for API compat)

EXCLUDED_TERM_RES = [
    re.compile(r"\bcontent(?:s)?\b|\btable\s+of\s+content(?:s)?\b", re.IGNORECASE),
    re.compile(r"\bmanagement(?:'?s)?\s+discussion\s+and\s+analysis\b|\bmd\s?&\s?a\b", re.IGNORECASE),
    re.compile(r"\bnon[-\s]?major\b", re.IGNORECASE),
    re.compile(r"\bcondensed?\b", re.IGNORECASE),
    re.compile(r"\bbudget(?:ary)?\b", re.IGNORECASE),
    re.compile(r"\bsupplement(?:ary|al)?\b", re.IGNORECASE),
    re.compile(r"\bstatistical\b", re.IGNORECASE),
    re.compile(r"\bfoundation\b", re.IGNORECASE),
]

ALLOWED_TERM_RES = [
    re.compile(r"\bcombined\b", re.IGNORECASE),
    re.compile(r"\bconsolidated\b", re.IGNORECASE),
    re.compile(r"\bgroup(?:ed)?\b", re.IGNORECASE),
]

_STMT       = r"statement(?:s)?"

_RECEIPT    = r"receipt(?:s)?"

_DISBURSE   = r"disbursement(?:s)?"

_EXPEND     = r"expenditure(?:s)?"

_EXPENSE    = r"expense(?:s)?"

_REVENUE    = r"revenue(?:s)?"

_POSITION   = r"position(?:s)?"

_ASSET      = r"asset(?:s)?"

_LIAB       = r"liabilit(?:y|ies)"

_BALANCE    = r"balance(?:s)?"

_RESERVE    = r"reserve(?:s)?"

_SHEET      = r"sheet(?:s)?"

_FUND       = r"fund(?:s)?"

_ACTIVITY   = r"activit(?:y|ies)"

_CHANGE     = r"change(?:s)?"

_OPERATION  = r"operation(?:s)?"

_FLOW       = r"flow(?:s)?"

_CHARGE     = r"charge(?:s)?"

_SERVICE    = r"service(?:s)?"

_GRANT      = r"grant(?:s)?"

_CONTRIB    = r"contribution(?:s)?"

_PROGRAM    = r"program(?:s)?"

_FUNCTION   = r"function(?:s)?"

_MEMBER_EQ  = r"member(?:s'?|'?s)?\s+equit(?:y|ies)"

_NP_OR_NA   = rf"net\s+{_POSITION}|net\s+{_ASSET}|{_MEMBER_EQ}"

_AND        = r"(?:and|&)"

_AND_OR_COMMA = r"[\s,]+(?:and|&)?\s*"

_BARE_BS_EXCLUDE = (r"current\s+fund|trust\s+fund(?:s)?|general\s+capital\s+fund|"
                    r"water\s+utility|sewer\s+utility|public\s+assistance\s+fund|"
                    r"assessment\s+trust\s+fund|govt|governmental?\s+fund(?:s)?")

STATEMENT_SPECS: Dict[str, Dict[str, Any]] = {

    # ─────────────────────────────────────────────────────────────────────
    # G3-1: STATEMENT OF CASH RECEIPTS AND DISBURSEMENTS
    # ─────────────────────────────────────────────────────────────────────
    "G3-1": {
        "label": "STATEMENT OF CASH RECEIPTS AND DISBURSEMENTS",
        "sectors": {"LG", "NONLG"},
        "name_res": [
            re.compile(
                rf"\b{_STMT}\s+of\s+cash\s+{_RECEIPT}\s+{_AND}\s+(?:{_DISBURSE}|{_EXPEND})\b",
                re.IGNORECASE
            ),
        ],
        "row_res": [
            [re.compile(rf"\b{_RECEIPT}\b", re.IGNORECASE)],
            [re.compile(rf"\b(?:{_DISBURSE}|{_EXPEND})\b", re.IGNORECASE)],
        ],
        "col_res": [],
        "col_all": False,
    },

    # ─────────────────────────────────────────────────────────────────────
    # G3-2: STATEMENT OF NET POSITION / BALANCE SHEET
    #  (+ OCBOA modified-cash: "Statement of Assets, Liabilities and Reserves")
    #  (+ NJ regulatory: General Capital Fund Balance Sheet — fund embedded)
    # ─────────────────────────────────────────────────────────────────────
    "G3-2": {
        "label": "STATEMENT OF NET POSITION / BALANCE SHEET",
        "sectors": {"LG", "NONLG"},
        "name_res": [
            re.compile(rf"\b{_STMT}\s+of\s+net\s+{_POSITION}\b", re.IGNORECASE),
            re.compile(rf"\b{_STMT}\s+of\s+financial\s+{_POSITION}\b", re.IGNORECASE),
            re.compile(rf"\b{_STMT}\s+of\s+net\s+financial\s+{_POSITION}\b", re.IGNORECASE),
            re.compile(rf"\b{_STMT}\s+of\s+net\s+{_ASSET}\b", re.IGNORECASE),
            re.compile(rf"\b{_STMT}\s+of\s+financial\s+{_ASSET}\b", re.IGNORECASE),
            re.compile(rf"\b{_STMT}\s+of\s+net\s+financial\s+{_ASSET}\b", re.IGNORECASE),
            # Bare Balance Sheet (NONLG / BTA) — GUARDED against NJ funds AND
            # governmental funds (the latter belongs to G3-3).
            re.compile(rf"^(?!.*(?:{_BARE_BS_EXCLUDE})).*\b{_BALANCE}\s+{_SHEET}\b",
                       re.IGNORECASE),
            # OCBOA modified-cash equivalent: "Statement of Assets, Liabilities and Reserves"
            re.compile(rf"{_STMT}\s+of\s+{_ASSET}[\s,]+{_LIAB}{_AND_OR_COMMA}{_RESERVE}", re.IGNORECASE),
            # NJ regulatory: General Capital Fund balance sheet (FUND EMBEDDED) → Net Position
            re.compile(rf"general\s+capital\s+{_FUND}[\s\S]*?{_BALANCE}\s+{_SHEET}", re.IGNORECASE),
            re.compile(rf"{_BALANCE}\s+{_SHEET}[\s\S]*?general\s+capital\s+{_FUND}", re.IGNORECASE),
        ],
        "row_res": [
            [re.compile(rf"\b{_ASSET}\b", re.IGNORECASE)],
            [re.compile(rf"\b{_LIAB}\b", re.IGNORECASE)],
            [re.compile(rf"\b(?:{_NP_OR_NA})\b|\b(?:total\s+)?{_RESERVE}\b|\bfund\s+{_BALANCE}\b",
                        re.IGNORECASE)],
        ],
        "col_res": [],
        "col_all": False,
    },

    # ─────────────────────────────────────────────────────────────────────
    # G3-3: BALANCE SHEET OF GOVERNMENTAL FUNDS  (LG only)
    #  (+ NJ regulatory: Current Fund Balance Sheet — fund embedded)
    # ─────────────────────────────────────────────────────────────────────
    "G3-3": {
        "label": "BALANCE SHEET OF GOVERNMENTAL FUNDS",
        "sectors": {"LG"},
        "name_res": [
            re.compile(rf"\b{_BALANCE}\s+{_SHEET}\b.*\b(?:govt|governmental?)\b.*\b{_FUND}\b", re.IGNORECASE),
            re.compile(rf"\bgovernmental\s+{_FUND}\s+{_BALANCE}\s+{_SHEET}\b", re.IGNORECASE),
            # NJ regulatory: Current Fund balance sheet (FUND EMBEDDED) → Governmental-fund BS
            re.compile(rf"current\s+{_FUND}[\s\S]*?{_BALANCE}\s+{_SHEET}", re.IGNORECASE),
            re.compile(rf"{_BALANCE}\s+{_SHEET}[\s\S]*?current\s+{_FUND}", re.IGNORECASE),
        ],
        "row_res": [
            [re.compile(rf"\b{_ASSET}\b", re.IGNORECASE)],
            [re.compile(rf"\b{_LIAB}\b", re.IGNORECASE)],
            [re.compile(rf"\bfund\s+{_BALANCE}\b|\b(?:total\s+)?{_RESERVE}\b", re.IGNORECASE)],
        ],
        "col_res": [
            re.compile(
                rf"\btotal\s+governmental\s+{_FUND}\b|\baggregat(?:e|ed)\s+{_FUND}\b"
                r"|regulatory\s+basis|december\s+31|\b20\d{2}\b",
                re.IGNORECASE
            ),
        ],
        "col_all": False,
    },

    # ─────────────────────────────────────────────────────────────────────
    # G3-4: STATEMENT OF ACTIVITIES / REV,EXP,ΔNET POSITION
    # ─────────────────────────────────────────────────────────────────────
    "G3-4": {
        "label": "STATEMENT OF ACTIVITIES / REV,EXP,DELTA-NET POSITION",
        "sectors": {"LG", "NONLG"},
        "name_res": [
            re.compile(rf"\b{_STMT}\s+of\s+(?:net\s+)?{_ACTIVITY}\b", re.IGNORECASE),
            re.compile(rf"\b{_STMT}\s+of\s+{_ACTIVITY}\b", re.IGNORECASE),
            re.compile(rf"\b{_STMT}\s+of\s+financial\s+{_ACTIVITY}\b", re.IGNORECASE),
            re.compile(rf"\b{_STMT}\s+of\s+net\s+financial\s+{_ACTIVITY}\b", re.IGNORECASE),
            re.compile(rf"\b{_STMT}\s+of\s+{_OPERATION}\b"
                       rf"(?![\s\S]*{_CHANGE}\s+in\s+{_FUND}\s+{_BALANCE})", re.IGNORECASE),
            re.compile(rf"\b{_STMT}\s+of\s+financial\s+{_OPERATION}\b", re.IGNORECASE),
            re.compile(
                rf"\b{_STMT}\s+of\s+(?:{_REVENUE}|{_RECEIPT})"
                rf"{_AND_OR_COMMA}(?:{_EXPENSE}|{_EXPEND})"
                rf"{_AND_OR_COMMA}{_CHANGE}\s+in\s+(?:{_NP_OR_NA})\b",
                re.IGNORECASE
            ),
            re.compile(rf"{_STMT}\s+of\s+{_REVENUE}[\s,]+{_EXPEND}{_AND_OR_COMMA}"
                       rf"{_CHANGE}\s+in\s+{_RESERVE}", re.IGNORECASE),
            re.compile(rf"current\s+{_FUND}[\s\S]*?{_STMT}\s+of\s+{_REVENUE}\b"
                       rf"(?![\s\S]*{_CHANGE}\s+in\s+{_FUND}\s+{_BALANCE})", re.IGNORECASE),
            re.compile(rf"current\s+{_FUND}[\s\S]*?{_STMT}\s+of\s+{_EXPEND}\b", re.IGNORECASE),
        ],
        "row_res": [
            [re.compile(
                rf"\b{_FUNCTION}\b|\b{_PROGRAM}\b|\b{_REVENUE}\b|\boperating\s+{_REVENUE}\b"
                rf"|\banticipated\b",
                re.IGNORECASE
            )],
            [re.compile(
                rf"\bgeneral\s+{_REVENUE}\b|\b{_EXPEND}\b|\boperating\s+{_EXPENSE}\b"
                rf"|\brealized\b|\bappropriation(?:s)?\b",
                re.IGNORECASE
            )],
            [re.compile(
                rf"\b(?:increase|decrease|change|net\s+change)\s+in\s+(?:{_NP_OR_NA})\b"
                rf"|\bexcess\s+of\s+{_REVENUE}\s+over\b"
                rf"|\b{_RESERVE}[\s,]+(?:beginning|end)\s+of\s+(?:the\s+)?year\b"
                rf"|\bexcess\s+or\s+deficit\b|\bfund\s+{_BALANCE}\b|\btotal\b",
                re.IGNORECASE
            )],
        ],
        "col_res": [],
        "col_all": False,
    },

    # ─────────────────────────────────────────────────────────────────────
    # G3-5: STATEMENT OF REVENUE, EXPENDITURE AND CHANGE IN FUND BALANCES  (LG only)
    # ─────────────────────────────────────────────────────────────────────
    "G3-5": {
        "label": "STATEMENT OF REVENUE, EXPENDITURE AND CHANGE IN FUND BALANCES",
        "sectors": {"LG"},
        "name_res": [
            re.compile(
                rf"\b{_STMT}\s+of\s+{_REVENUE}"
                rf"{_AND_OR_COMMA}{_EXPEND}"
                rf"{_AND_OR_COMMA}{_CHANGE}\s+in\s+fund\s+{_BALANCE}\b", re.IGNORECASE
            ),
            re.compile(rf"current\s+{_FUND}[\s\S]*?{_STMT}\s+of\s+{_OPERATION}"
                       rf"[\s\S]*?{_CHANGE}\s+in\s+{_FUND}\s+{_BALANCE}", re.IGNORECASE),
        ],
        "row_res": [
            [re.compile(rf"\b{_REVENUE}\b", re.IGNORECASE)],
            [re.compile(rf"\b{_EXPEND}\b", re.IGNORECASE)],
            [re.compile(
                rf"\b(?:increase|decrease|change|net\s+change)\s+in\s+fund\s+{_BALANCE}\b"
                rf"|\bbalance\s+december\s+31\b|\bstatutory\s+excess\b|\bfund\s+{_BALANCE}\b",
                re.IGNORECASE
            )],
        ],
        "col_res": [],
        "col_all": False,
    },

    # ─────────────────────────────────────────────────────────────────────
    # G3-6: STATEMENT OF CASH FLOWS  (NONLG only)
    # ─────────────────────────────────────────────────────────────────────
    "G3-6": {
        "label": "STATEMENT OF CASH FLOWS",
        "sectors": {"NONLG"},
        "name_res": [
            re.compile(
                rf"\b(?:direct\s+|indirect\s+)?{_STMT}\s+of\s+cash\s+{_FLOW}\b",
                re.IGNORECASE
            ),
            re.compile(
                rf"\b(?:direct\s+|indirect\s+)?cash\s+{_FLOW}\s+{_STMT}\b",
                re.IGNORECASE
            ),
        ],
        "row_res": [
            [re.compile(
                rf"\b(?:cash\s+{_FLOW}[\s\-–—]+from\s+)?"
                rf"(?:{_OPERATION}|operating\s+{_ACTIVITY})\b",
                re.IGNORECASE
            )],
            [re.compile(
                r"cash(?:\s+and\s+cash\s+equivalents?)?(?:\s*,?\s*and\s+restricted\s+cash)?"
                r"[\s,\-]{1,10}"
                r"(?:balance\s+)?(?:at\s+(?:the\s+)?)?beg(?:in(?:n)?ing)"
                r"(?:\s+of\s+(?:the\s+)?(?:year|period|fiscal\s+year))?"
                r"|"
                r"(?:at\s+(?:the\s+)?)?beg(?:in(?:n)?ing)"
                r"(?:\s+of\s+(?:the\s+)?(?:year|period|fiscal\s+year))?[\s,\-]{1,10}"
                r"(?:balance\s+of\s+)?cash"
                r"(?:\s+and\s+cash\s+equivalents?)?(?:\s*,?\s*and\s+restricted\s+cash)?"
                r"|"
                r"beg(?:in(?:n)?ing)?\s+(?:balance\s+of\s+cash|cash\s+balance)"
                r"|"
                r"cash\s+balance\s+at\s+the\s+beg(?:in(?:n)?ing)?\s+of\s+the\s+year",
                re.IGNORECASE
            )],
            [re.compile(
                r"cash(?:\s+and\s+cash\s+equivalents?)?(?:\s*,?\s*and\s+restricted\s+cash)?"
                r"[\s,\-]{1,10}"
                r"(?:balance\s+)?(?:at\s+(?:the\s+)?)?end(?:ing)?"
                r"(?:\s+of\s+(?:the\s+)?(?:year|period|fiscal\s+year))?"
                r"|"
                r"(?:at\s+(?:the\s+)?)?end(?:ing)?"
                r"(?:\s+of\s+(?:the\s+)?(?:year|period|fiscal\s+year))?[\s,\-]{1,10}"
                r"(?:balance\s+of\s+)?cash"
                r"(?:\s+and\s+cash\s+equivalents?)?(?:\s*,?\s*and\s+restricted\s+cash)?"
                r"|"
                r"end(?:ing)?\s+(?:balance\s+of\s+cash|cash\s+balance)"
                r"|"
                r"cash\s+balance\s+at\s+the\s+end\s+of\s+the\s+year",
                re.IGNORECASE
            )],
        ],
        "col_res": [],
        "col_all": False,
    },
}

def page_heading_status(doc: fitz.Document, page_idx: int) -> str:
    """
    Inspect first N lines of a page. Returns:
         'EXCLUDED'  → skip entirely
         'ALLOWED'   → do NOT skip
         'NEUTRAL'   → treat as normal
    CRITICAL: EXCLUDED terms must appear as PAGE TITLES only.
    """
    top = page_top_lines(doc, page_idx, G3_HEADING_TOP_LINES)
    if not top:
        return "NEUTRAL"

    normalized_lines = []
    for ln in top:
        text = (ln.text
                 .replace("\u00A0", " ")
                 .replace("\u2019", "'")
                 .replace("\u2018", "'"))
        text = re.sub(r"\s+", " ", text).strip()
        normalized_lines.append(text)

    # ALLOWED check first — always accepts combined/consolidated statements
    for text in normalized_lines:
        text_lower = text.lower()
        for rx in ALLOWED_TERM_RES:
            if rx.search(text_lower):
                return "ALLOWED"

    # EXCLUDED check — only in TOP 3 lines, only if short + not mid-sentence
    for i, text in enumerate(normalized_lines[:3]):
        text_lower = text.lower()
        if len(text) >= 80:
            continue
        if re.search(r"(?:[,;]|\b(?:and|or|to|of|in|for|the|a|an)\s*)$", text_lower):
            continue
        for rx in EXCLUDED_TERM_RES:
            if rx.search(text_lower):
                return "EXCLUDED"

    return "NEUTRAL"

def match_statement_title(doc: fitz.Document, page_idx: int,
                          spec: Dict[str, Any]) -> Optional[List[PageLine]]:
    """Match statement title STRICTLY within top N lines of the page."""
    top = page_top_lines(doc, page_idx, G3_HEADING_TOP_LINES)
    if not top or len(top) == 0:
        return None
    n = len(top)
    for start in range(n):
        for span in range(1, G3_TITLE_SPAN_MAX + 1):
            if start + span > n:
                break
            raw_chunk = " ".join(top[i].text for i in range(start, start + span))
            normalized_chunk = (raw_chunk
                                .replace("\u00A0", " ")
                                .replace("\u2019", "'")
                                .replace("\u2018", "'")
                                .replace("\u201C", '"')
                                .replace("\u201D", '"'))
            normalized_chunk = re.sub(r"\s+", " ", normalized_chunk).strip().lower()
            for rx in spec["name_res"]:
                if rx.search(normalized_chunk):
                    return top[start:start + span]
    return None

STATEMENT_FUZZY_NAMES = {
    "G3-1": ["statement of cash receipts and disbursements"],
    "G3-2": ["statement of net position", "balance sheet"],
    "G3-3": ["balance sheet of governmental funds",
             "governmental funds balance sheet"],
    "G3-4": ["statement of activities"],
    "G3-5": ["statement of revenues expenditures and changes in fund balances",
             "governmental funds statement of revenues expenditures "
             "and changes in fund balances"],
    "G3-6": ["statement of cash flows"],
}

def _fuzzy_title_match(normalized_chunk: str, step_code: str,
                       threshold: int = 80) -> bool:
    "Fuzzy-compare a normalized title chunk to canonical names (OCR-tolerant)."
    names = STATEMENT_FUZZY_NAMES.get(step_code, [])
    if not names or not normalized_chunk:
        return False
    chunk = re.sub(r"\bfor the (?:fiscal )?year(?:s)? ended.*$", "", normalized_chunk)
    chunk = re.sub(r"[^a-z0-9 ]", " ", chunk)
    chunk = re.sub(r"\s+", " ", chunk).strip()
    for canon in names:
        if fuzz.token_set_ratio(chunk, canon) >= threshold:
            return True
        if fuzz.partial_ratio(chunk, canon) >= threshold + 8:
            return True
    return False

def _fuzzy_title_match_page(doc, page_idx, step_code, threshold: int = 80) -> bool:
    "Slide a 1..G3_TITLE_SPAN_MAX line window over the top lines and fuzzy-test."
    top = page_top_lines(doc, page_idx, G3_HEADING_TOP_LINES)
    if not top:
        return False
    n = len(top)
    for start in range(n):
        for span in range(1, G3_TITLE_SPAN_MAX + 1):
            if start + span > n:
                break
            chunk = " ".join(top[i].text for i in range(start, start + span))
            chunk = chunk.replace("\u00A0", " ").replace("\u2019", "'")
            chunk = re.sub(r"\s+", " ", chunk).strip().lower()
            if _fuzzy_title_match(chunk, step_code, threshold):
                return True
    return False

DEBUG_DUMP_PAGE_LINES: Optional[List[int]] = None    # e.g. [51] for Austin CCD page 51

def _debug_dump_page(doc: fitz.Document, page_idx: int, label: str = ""):
    "Print all extracted lines from a page (0-indexed page_idx)."
    if DEBUG_DUMP_PAGE_LINES is None:
        return
    if page_idx not in DEBUG_DUMP_PAGE_LINES:
        return
    print(f"\n📄 [DEBUG] All lines on page {page_idx+1} {label}:")
    lines = extract_page_lines(doc, page_idx)
    for i, ln in enumerate(lines):
        indicator = "🎯" if i < G3_HEADING_TOP_LINES else "  "
        print(f"    {indicator} [{i}]: {ln.text[:100]!r}")
    print()

def _check_rows(doc: fitz.Document,
                start_page: int,
                spec: Dict[str, Any]) -> Tuple[bool, List[int], List[PageLine]]:
    """Walk downward through up to G3_STMT_MAX_CONT_PAGES pages..."""
    required_groups: List[List] = spec["row_res"]
    if not required_groups:
        return True, [start_page], []

    remaining = list(required_groups)
    matched_lines: List[PageLine] = []
    pages_covered: List[int] = [start_page]

    end_page = min(start_page + G3_STMT_MAX_CONT_PAGES, doc.page_count)
    for pp in range(start_page, end_page):
        if pp != start_page:
            hs = page_heading_status(doc, pp)
            if hs == "EXCLUDED":
                break
            if pp not in pages_covered:
                pages_covered.append(pp)
        skip_lines = G3_HEADING_TOP_LINES if pp == start_page else 0
        lines = extract_page_lines(doc, pp)
        for ln in lines[skip_lines: skip_lines + G3_ROW_SEARCH_WINDOW]:
            normalized = (ln.text
                          .replace("\u00A0", " ")
                          .replace("\u2019", "'")
                          .replace("\u2018", "'")
                          .replace("\u201C", '"')
                          .replace("\u201D", '"')
                          .replace("\u2013", "-")
                          .replace("\u2014", "-"))
            normalized = re.sub(r"\s+", " ", normalized).strip().lower()
            for grp in list(remaining):
                if any(rx.search(normalized) for rx in grp):
                    remaining.remove(grp)
                    matched_lines.append(ln)
                    break
            if not remaining:
                break
        if not remaining:
            break

    return (len(remaining) == 0), pages_covered, matched_lines

def _check_columns(doc, start_page, spec):
    """Column-wise structure check with text normalization."""
    col_res = spec.get("col_res", [])
    require_all = spec.get("col_all", False)
    end_page = min(start_page + G3_STMT_MAX_CONT_PAGES, doc.page_count)

    if not col_res:
        for pp in range(start_page, end_page):
            for ln in extract_page_lines(doc, pp):
                if re.search(r"\$?\s*[\d,]{3,}", ln.text):
                    return True
        return False

    hit_regexes = set()
    for pp in range(start_page, end_page):
        for ln in extract_page_lines(doc, pp):
            normalized = (ln.text
                          .replace("\u00A0", " ")
                          .replace("\u2019", "'")
                          .replace("\u2018", "'"))
            normalized = re.sub(r"\s+", " ", normalized).strip().lower()
            for i, rx in enumerate(col_res):
                if i not in hit_regexes and rx.search(normalized):
                    hit_regexes.add(i)
    return len(hit_regexes) == len(col_res) if require_all else len(hit_regexes) >= 1

def _fye_on_statement_pages(doc: fitz.Document,
                              page_range: List[int],
                              target_fye: Optional[Tuple[int, int]]) -> str:
    """Return 'EXACT' | 'PARTIAL' | 'NONE' for FYE presence on statement pages."""
    if not target_fye or not page_range:
        return "NONE"

    best = "NONE"
    for pp in page_range:
        # ─── Layer 1: Top 5 lines (spec-compliant primary search) ───
        top_candidates = list(page_top_lines(doc, pp, FYE_TOP_LINES))
        for ln in top_candidates:
            for m, y, _raw in extract_dates_from_text(ln.text):
                cat = fye_matches(target_fye, (m, y))
                if cat == "EXACT":
                    return "EXACT"
                if cat in ("MONTH", "YEAR"):
                    best = "PARTIAL"

        # ─── Layer 2: Full page text (catches column header dates) ───
        full_page_lines = extract_page_lines(doc, pp)
        for ln in full_page_lines:
            if ln in top_candidates:
                continue
            for m, y, _raw in extract_dates_from_text(ln.text):
                cat = fye_matches(target_fye, (m, y))
                if cat == "EXACT":
                    return "EXACT"
                if cat in ("MONTH", "YEAR"):
                    best = "PARTIAL"

    return best

def _accumulate_stmt_lines(doc, start_p, title_match_fn, end_marker=None):
    """From the title page, gather this page + continuation pages until a
    DIFFERENT statement starts, the end-marker is captured, or the cap is hit.
    Returns (all_lines: List[PageLine], combined_text: str, pages: List[int])."""
    pages = [start_p]
    all_lines = list(extract_page_lines(doc, start_p))
    combined  = " ".join(ln.text for ln in all_lines)

    for nxt in range(start_p + 1, min(start_p + G3_CONT_MAX_PAGES, doc.page_count)):
        nxt_top = " ".join(ln.text for ln in page_top_lines(doc, nxt, 5)).lower()
        same_title = title_match_fn(nxt_top)
        starts_new = ("statement" in nxt_top and not same_title
                      and "continued" not in nxt_top)
        if starts_new:
            break
        nxt_lines = list(extract_page_lines(doc, nxt))
        all_lines.extend(nxt_lines)
        combined += " " + " ".join(ln.text for ln in nxt_lines)
        pages.append(nxt)
        if end_marker and end_marker.search(combined):
            break
    return all_lines, combined, pages

def _fye_present_in_text(text: str, workbook_fye) -> str:
    "Return 'EXACT' | 'NONE' for workbook FYE month+year presence in accumulated text."
    target = parse_fye_from_workbook(workbook_fye)
    if not target:
        return "NONE"
    tm, ty = target
    months = {1:"jan",2:"feb",3:"mar",4:"apr",5:"may",6:"jun",7:"jul",
              8:"aug",9:"sep",10:"oct",11:"nov",12:"dec"}
    mon = months.get(tm, "")
    low = text.lower()
    month_ok = bool(re.search(rf"\b{mon}[a-z]*\.?\b", low) or
                    re.search(rf"\b{tm:02d}[/.\-]\d", low))
    year_ok  = bool(re.search(rf"\b{ty}\b", low))
    return "EXACT" if (month_ok and year_ok) else "NONE"

def _page_excluded_by_top_lines(top_5_lines: List[PageLine]) -> bool:
    """
    Return True if the page should be SKIPPED because any of the first 5
    lines contains an EXCLUDED term AND does NOT contain an ALLOWED term.
    Per spec: ALLOWED terms rescue a page that would otherwise be excluded.
    """
    if not top_5_lines:
        return False

    normalized = []
    for ln in top_5_lines:
        text = (ln.text
                 .replace("\u00A0", " ")
                 .replace("\u2019", "'")
                 .replace("\u2018", "'"))
        text = re.sub(r"\s+", " ", text).strip().lower()
        normalized.append(text)

    combined_top_text = " | ".join(normalized)

    for rx in ALLOWED_TERM_RES:
        if rx.search(combined_top_text):
            return False

    for text in normalized[:3]:
        if len(text) >= 80:
            continue
        if re.search(r"(?:[,;]|\b(?:and|or|to|of|in|for|the|a|an)\s*)$", text):
            continue
        for rx in EXCLUDED_TERM_RES:
            if rx.search(text):
                return True
    return False

def validate_g3_2_net_position_nonlg(doc: fitz.Document,
                                     sector: str, state: str, audit_year,
                                     workbook_fye, screenshot_out: Path
                                     ) -> Tuple[StepResult, List[int]]:
    """G3-2 (NONLG variant): Statement of Net Position / Balance Sheet — continuation-aware."""
    result = StepResult(code="G3-2")

    if sector.upper() != "NONLG":
        result.status = "NOT APPLICABLE"
        result.notes = "This dedicated G3-2 validator applies only to NONLG entities"
        return result, []

    title_patterns = [
        re.compile(
            r"\b(?:combined\s+|consolidated\s+|group(?:ed)?\s+)?"
            r"statement(?:s)?\s+of\s+"
            r"(?:net\s+|financial\s+|net\s+financial\s+)?"
            r"(?:position(?:s)?|asset(?:s)?)\b",
            re.IGNORECASE
        ),
        re.compile(
            r"\b(?:combined\s+|consolidated\s+|group(?:ed)?\s+)?"
            r"balance\s+sheet(?:s)?\b",
            re.IGNORECASE
        ),
    ]

    row_pattern = re.compile(
        r"\basset(?:s)?\b"
        r"|"
        r"\bliabilit(?:y|ies)\b"
        r"|"
        r"\b(?:net\s+position(?:s)?|net\s+asset(?:s)?|"
        r"member(?:s'?|'?s)?\s+equit(?:y|ies))\b",
        re.IGNORECASE
    )

    target_fye = parse_fye_from_workbook(workbook_fye)

    def _title_matches(text: str) -> bool:
        return any(rx.search(text) for rx in title_patterns)

    def _fye_on_page(page_lines) -> bool:
        if not target_fye:
            return False
        target_month, target_year = target_fye
        for ln in page_lines:
            for m, y, _raw in extract_dates_from_text(ln.text):
                if fye_matches(target_fye, (m, y)) in ("EXACT", "MONTH", "YEAR"):
                    return True
        full_text = " ".join(ln.text for ln in page_lines)
        for m, y, _raw in extract_dates_from_text(full_text):
            if fye_matches(target_fye, (m, y)) in ("EXACT", "MONTH", "YEAR"):
                return True
        month_names = ["january", "february", "march", "april", "may", "june",
                       "july", "august", "september", "october", "november", "december"]
        target_month_name = month_names[target_month - 1]
        month_short = target_month_name[:3]
        page_text_lower = full_text.lower()
        month_present = bool(re.search(rf"\b{month_short}\w*\b", page_text_lower))
        year_present = bool(re.search(rf"\b{target_year}\b", page_text_lower))
        return bool(month_present and year_present)

    end_marker = re.compile(r"total\s+net\s+(?:position|asset(?:s)?)"
                            r"|total\s+liabilities\s+and\s+net", re.IGNORECASE)

    best_review_candidate = None
    best_review_pages = []

    for p in range(doc.page_count):
        lines = extract_page_lines(doc, p)
        if not lines:
            continue
        top_5 = lines[:5]

        title_found = False
        title_line_idx = None
        for i in range(len(top_5)):
            if _title_matches(top_5[i].text):
                title_found = True
                title_line_idx = i
                break
            if i + 1 < len(top_5):
                combined = top_5[i].text + " " + top_5[i + 1].text
                if _title_matches(combined):
                    title_found = True
                    title_line_idx = i
                    break
        if not title_found:
            continue

        if page_needs_ocr(doc, p):
            result.status = OCR_STATUS; result.score = 0.0
            result.page_idx = p; result.page_range = str(p)
            result.notes = ("Statement title found but page is image-only; OCR required.")
            shot = build_screenshot_name(sector, state, audit_year, "G3-2", p)
            result.screenshot = str(save_highlighted_screenshot(
                doc, p, [], CLR_LIGHT_YELLOW, screenshot_out, shot))
            return result, [p]
            
        top_5_low = " ".join(ln.text for ln in top_5).lower()
        if _page_excluded_by_top_lines(top_5) and not (
                "net position" in top_5_low or "financial position" in top_5_low
                or "balance sheet" in top_5_low or "net asset" in top_5_low):
            continue

        acc_lines, acc_text, pages_covered = _accumulate_stmt_lines(
            doc, p, _title_matches, end_marker)

        has_row = bool(row_pattern.search(acc_text))
        if not has_row:
            continue

        has_fye = _fye_on_page(acc_lines)

        page_range_str = (f"{pages_covered[0]+1}-{pages_covered[-1]+1}"
                          if len(pages_covered) > 1 else str(pages_covered[0]+1))

        if has_fye:
            status, score = "PASS", 100.0
            highlight = CLR_LIGHT_GREEN
            notes = (f"Title on page {p+1} (line {title_line_idx+1}); spans "
                     f"[{page_range_str}]. Row structure (Asset/Liability/Equity) ✓, FYE ✓")
        else:
            status, score = "REVIEW", 90.0
            highlight = CLR_LIGHT_YELLOW
            notes = (f"Title on page {p+1} (spans [{page_range_str}]); rows ✓, "
                     f"but target FYE not found across statement pages")

        try:
            shot_name = build_screenshot_name(sector, state, audit_year, "G3-2", p)
            shot_path = save_highlighted_screenshot(
                doc, p, [top_5[title_line_idx].bbox],
                highlight, screenshot_out, shot_name
            )
            screenshot_str = str(shot_path)
        except Exception:
            screenshot_str = None

        candidate = StepResult(
            code="G3-2", status=status, score=score, page_idx=p,
            screenshot=screenshot_str, notes=notes,
            page_range=(page_range_str if len(pages_covered) > 1 else None),
        )

        if status == "PASS":
            return candidate, pages_covered
        elif best_review_candidate is None:
            best_review_candidate = candidate
            best_review_pages = pages_covered

    if best_review_candidate is not None:
        return best_review_candidate, best_review_pages

    result.status = "FAIL"
    result.score = 0.0
    result.notes = ("No Statement of Net Position/Balance Sheet page found: no page "
                    "had a valid title AND row-level structure "
                    "(Asset/Liability/Net Position/Net Assets/Member's Equity)")
    try:
        result.screenshot = take_diagnostic_screenshot(
            doc, 0, sector, state, audit_year, "G3-2", screenshot_out, status="FAIL")
    except Exception:
        pass
    return result, []

def validate_g3_4_activities_nonlg(doc: fitz.Document,
                                   sector: str, state: str, audit_year,
                                   workbook_fye, screenshot_out: Path
                                   ) -> Tuple[StepResult, List[int]]:
    """G3-4 (NONLG variant): Statement of Activities / Operations / Changes — continuation-aware."""
    result = StepResult(code="G3-4")

    if sector.upper() != "NONLG":
        result.status = "NOT APPLICABLE"
        result.notes = "This dedicated G3-4 validator applies only to NONLG entities"
        return result, []

    title_patterns = [
        re.compile(
            r"\b(?:combined\s+|consolidated\s+|group(?:ed)?\s+)?"
            r"statement(?:s)?\s+of\s+"
            r"(?:net\s+|financial\s+|net\s+financial\s+)?"
            r"activit(?:y|ies)\b",
            re.IGNORECASE
        ),
        re.compile(
            r"\b(?:combined\s+|consolidated\s+|group(?:ed)?\s+)?"
            r"statement(?:s)?\s+of\s+(?:financial\s+)?operation(?:s)?\b",
            re.IGNORECASE
        ),
        re.compile(
            r"\b(?:combined\s+|consolidated\s+|group(?:ed)?\s+)?"
            r"statement(?:s)?\s+of\s+"
            r"(?:increase(?:s)?|decrease(?:s)?|change(?:s)?)\s+in\s+"
            r"net\s+(?:asset|position)(?:s)?\b",
            re.IGNORECASE
        ),
        re.compile(
            r"\b(?:combined\s+|consolidated\s+|group(?:ed)?\s+)?"
            r"statement(?:s)?\s+of\s+"
            r"(?:revenue(?:s)?|receipt(?:s)?)"
            r"[\s,]+(?:and|&)?\s*"
            r"(?:expense(?:s)?|expenditure(?:s)?)"
            r"[\s,]+(?:and|&)?\s*"
            r"(?:increase(?:s)?|decrease(?:s)?|change(?:s)?)\s+in\s+"
            r"(?:net\s+position(?:s)?|net\s+asset(?:s)?|"
            r"member(?:s'?|'?s)?\s+equit(?:y|ies))\b",
            re.IGNORECASE
        ),
    ]

    row_pattern = re.compile(
        r"\b(?:receipt(?:s)?|revenue(?:s)?|expense(?:s)?|expenditure(?:s)?)\b"
        r"|"
        r"\b(?:change(?:s)?|increase(?:s)?|decrease(?:s)?)\s+in\s+"
        r"(?:net\s+position(?:s)?|net\s+asset(?:s)?|"
        r"member(?:s'?|'?s)?\s+equit(?:y|ies))\b",
        re.IGNORECASE
    )

    target_fye = parse_fye_from_workbook(workbook_fye)

    def _title_matches(text: str) -> bool:
        return any(rx.search(text) for rx in title_patterns)

    def _fye_on_page(page_lines) -> bool:
        if not target_fye:
            return False
        target_month, target_year = target_fye
        for ln in page_lines:
            for m, y, _raw in extract_dates_from_text(ln.text):
                if fye_matches(target_fye, (m, y)) in ("EXACT", "MONTH", "YEAR"):
                    return True
        full_text = " ".join(ln.text for ln in page_lines)
        for m, y, _raw in extract_dates_from_text(full_text):
            if fye_matches(target_fye, (m, y)) in ("EXACT", "MONTH", "YEAR"):
                return True
        month_names = ["january", "february", "march", "april", "may", "june",
                       "july", "august", "september", "october", "november", "december"]
        target_month_name = month_names[target_month - 1]
        month_short = target_month_name[:3]
        page_text_lower = full_text.lower()
        month_present = bool(re.search(rf"\b{month_short}\w*\b", page_text_lower))
        year_present = bool(re.search(rf"\b{target_year}\b", page_text_lower))
        return bool(month_present and year_present)

    end_marker = re.compile(r"net\s+assets?,?\s+end(?:\s+of\s+year)?"
                            r"|change(?:s)?\s+in\s+net\s+position", re.IGNORECASE)

    best_review_candidate = None
    best_review_pages = []

    for p in range(doc.page_count):
        lines = extract_page_lines(doc, p)
        if not lines:
            continue
        top_5 = lines[:5]

        title_found = False
        title_line_idx = None
        for i in range(len(top_5)):
            if _title_matches(top_5[i].text):
                title_found = True
                title_line_idx = i
                break
            if i + 1 < len(top_5):
                combined = top_5[i].text + " " + top_5[i + 1].text
                if _title_matches(combined):
                    title_found = True
                    title_line_idx = i
                    break
        if not title_found:
            continue

        if page_needs_ocr(doc, p):
            result.status = OCR_STATUS; result.score = 0.0
            result.page_idx = p; result.page_range = str(p)
            result.notes = ("Statement title found but page is image-only; OCR required.")
            shot = build_screenshot_name(sector, state, audit_year, "G3-4", p)
            result.screenshot = str(save_highlighted_screenshot(
                doc, p, [], CLR_LIGHT_YELLOW, screenshot_out, shot))
            return result, [p]
            
        top_5_low = " ".join(ln.text for ln in top_5).lower()
        if _page_excluded_by_top_lines(top_5) and not (
                "activit" in top_5_low or "operation" in top_5_low
                or "revenue" in top_5_low or "net asset" in top_5_low
                or "net position" in top_5_low):
            continue

        acc_lines, acc_text, pages_covered = _accumulate_stmt_lines(
            doc, p, _title_matches, end_marker)

        has_row = bool(row_pattern.search(acc_text))
        if not has_row:
            continue

        has_fye = _fye_on_page(acc_lines)

        page_range_str = (f"{pages_covered[0]+1}-{pages_covered[-1]+1}"
                          if len(pages_covered) > 1 else str(pages_covered[0]+1))

        if has_fye:
            status, score = "PASS", 100.0
            highlight = CLR_LIGHT_GREEN
            notes = (f"Title on page {p+1} (line {title_line_idx+1}); spans "
                     f"[{page_range_str}]. Row structure ✓, FYE ✓")
        else:
            status, score = "REVIEW", 90.0
            highlight = CLR_LIGHT_YELLOW
            notes = (f"Title on page {p+1} (spans [{page_range_str}]); rows ✓, "
                     f"but target FYE not found across statement pages")

        try:
            shot_name = build_screenshot_name(sector, state, audit_year, "G3-4", p)
            shot_path = save_highlighted_screenshot(
                doc, p, [top_5[title_line_idx].bbox],
                highlight, screenshot_out, shot_name
            )
            screenshot_str = str(shot_path)
        except Exception:
            screenshot_str = None

        candidate = StepResult(
            code="G3-4", status=status, score=score, page_idx=p,
            screenshot=screenshot_str, notes=notes,
            page_range=(page_range_str if len(pages_covered) > 1 else None),
        )

        if status == "PASS":
            return candidate, pages_covered
        elif best_review_candidate is None:
            best_review_candidate = candidate
            best_review_pages = pages_covered

    if best_review_candidate is not None:
        return best_review_candidate, best_review_pages

    result.status = "FAIL"
    result.score = 0.0
    result.notes = ("No Statement of Activities page found: no page had a valid "
                    "title AND row-level structure (revenues/expenses or change in "
                    "net position/assets)")
    try:
        result.screenshot = take_diagnostic_screenshot(
            doc, 0, sector, state, audit_year, "G3-4", screenshot_out, status="FAIL")
    except Exception:
        pass
    return result, []

def validate_g3_6_cash_flows(doc: fitz.Document,
                             sector: str, state: str, audit_year,
                             workbook_fye, screenshot_out: Path
                             ) -> Tuple[StepResult, List[int]]:
    """G3-6: Statement of Cash Flows — dedicated validator (continuation-aware)."""
    result = StepResult(code="G3-6")

    if sector.upper() != "NONLG":
        result.status = "NOT APPLICABLE"
        result.notes = "G3-6 not applicable to LG entities"
        return result, []

    title_pattern = re.compile(
        r"\b(?:direct\s+|indirect\s+)?"
        r"(?:combined\s+|consolidated\s+|group(?:ed)?\s+)?"
        r"(?:statement(?:s)?\s+of\s+cash\s+flow(?:s)?"
        r"|cash\s+flow(?:s)?\s+statement(?:s)?)\b",
        re.IGNORECASE
    )

    op_pattern    = re.compile(r"\boperat(?:ion(?:s)?|ing)\b", re.IGNORECASE)
    begin_pattern = re.compile(r"\bbeg(?:in(?:n)?ing)?\b", re.IGNORECASE)
    end_pattern   = re.compile(r"\bend(?:ing)?\b", re.IGNORECASE)
    investing_pattern = re.compile(r"\binvest(?:ing|ment(?:s)?)?\b", re.IGNORECASE)
    financing_pattern = re.compile(r"\bfinanc(?:ing|e(?:s)?)?\b", re.IGNORECASE)

    # ✅ FIX 4: use the unified cap (was a local CF_MAX_CONT_PAGES = 4)
    def _accumulate_cf_text(start_p: int) -> Tuple[str, List[int]]:
        pages_covered = [start_p]
        combined = " ".join(ln.text for ln in extract_page_lines(doc, start_p))
        for nxt in range(start_p + 1, min(start_p + G3_CONT_MAX_PAGES, doc.page_count)):
            nxt_top = page_top_lines(doc, nxt, 5)
            nxt_top_text = " ".join(ln.text for ln in nxt_top).lower()
            same_cf   = bool(title_pattern.search(nxt_top_text))
            starts_new_statement = ("statement" in nxt_top_text and not same_cf)
            if starts_new_statement:
                break
            pages_covered.append(nxt)
            combined += " " + " ".join(ln.text for ln in extract_page_lines(doc, nxt))
            if re.search(r"\bend(?:ing)?\b.{0,40}cash", combined, re.IGNORECASE):
                break
        return combined, pages_covered

    best_result = None
    best_pages_covered = []

    for p in range(doc.page_count):
        lines = extract_page_lines(doc, p)
        if not lines:
            continue
        top_5 = lines[:5]

        title_found = False
        title_line_idx = None
        for i in range(len(top_5)):
            if title_pattern.search(top_5[i].text):
                title_found, title_line_idx = True, i
                break
            if i + 1 < len(top_5):
                combined_lines = top_5[i].text + " " + top_5[i + 1].text
                if title_pattern.search(combined_lines):
                    title_found, title_line_idx = True, i
                    break
        if not title_found:
            continue

        if page_needs_ocr(doc, p):
            result.status = OCR_STATUS; result.score = 0.0
            result.page_idx = p; result.page_range = str(p)
            result.notes = ("Statement title found but page is image-only; OCR required.")
            shot = build_screenshot_name(sector, state, audit_year, "G3-6", p)
            result.screenshot = str(save_highlighted_screenshot(
                doc, p, [], CLR_LIGHT_YELLOW, screenshot_out, shot))
            return result, [p]
            
        top_5_text_low = " ".join(ln.text for ln in top_5).lower()
        if _page_excluded_by_top_lines(top_5) and "cash flow" not in top_5_text_low:
            continue

        page_text, pages_covered = _accumulate_cf_text(p)
        has_operating = bool(op_pattern.search(page_text))
        has_beginning = bool(begin_pattern.search(page_text))
        has_ending    = bool(end_pattern.search(page_text))

        if not (has_operating and has_beginning and has_ending):
            continue

        has_investing = bool(investing_pattern.search(page_text))
        has_financing = bool(financing_pattern.search(page_text))

        page_range_str = (f"{pages_covered[0] + 1}-{pages_covered[-1] + 1}"
                          if len(pages_covered) > 1 else str(pages_covered[0] + 1))

        if has_investing or has_financing:
            status, score = "PASS", 100.0
            highlight = CLR_LIGHT_GREEN
            notes = (f"Title on page {p+1} (line {title_line_idx+1}); "
                     f"statement spans pages [{page_range_str}]. "
                     f"Operating✓, "
                     f"Beginning={'✓' if has_beginning else '✗'}, "
                     f"Ending={'✓' if has_ending else '✗'}, "
                     f"Investing={'✓' if has_investing else '✗'}, "
                     f"Financing={'✓' if has_financing else '✗'}")
        else:
            status, score = "REVIEW", 90.0
            highlight = CLR_LIGHT_YELLOW
            notes = (f"Title on page {p+1} (spans [{page_range_str}]), but neither "
                     f"Investing nor Financing section detected")

        try:
            shot_name = build_screenshot_name(sector, state, audit_year, "G3-6", p)
            shot_path = save_highlighted_screenshot(
                doc, p, [top_5[title_line_idx].bbox],
                highlight, screenshot_out, shot_name
            )
            screenshot_str = str(shot_path)
        except Exception:
            screenshot_str = None

        candidate = StepResult(
            code="G3-6", status=status, score=score, page_idx=p,
            screenshot=screenshot_str, notes=notes,
            page_range=(page_range_str if len(pages_covered) > 1 else None),
        )

        if status == "PASS":
            return candidate, pages_covered
        elif best_result is None:
            best_result = candidate
            best_pages_covered = pages_covered

    if best_result is not None:
        return best_result, best_pages_covered

    result.status = "FAIL"
    result.score = 0.0
    result.notes = ("No Cash Flow Statement page found: no page had a valid title "
                    "AND (Operating + Beginning & Ending) rows across continuation pages")
    try:
        result.screenshot = take_diagnostic_screenshot(
            doc, 0, sector, state, audit_year, "G3-6", screenshot_out, status="FAIL")
    except Exception:
        pass
    return result, []

NJ_FUND_PATTERNS = {
    "CURRENT":              re.compile(r"\bcurrent\s+fund\b", re.IGNORECASE),
    "GENERAL_CAPITAL":      re.compile(r"\bgeneral\s+capital\s+fund\b", re.IGNORECASE),
    "GENERAL_FIXED_ASSETS": re.compile(r"\bgeneral\s+fixed\s+assets?\b", re.IGNORECASE),
}

_NJ_FUND_PRIORITY = ["CURRENT", "GENERAL_CAPITAL", "GENERAL_FIXED_ASSETS"]

def detect_page_fund(doc: fitz.Document, page_idx: int,
                     top_n: int = 6, lookback: int = 3) -> Optional: # ✅ FIX 5
    """Return the NJ fund key for a page, or None if no NJ fund context."""
    def _scan(pi: int) -> Optional[str]:
        try:
            lines = page_top_lines(doc, pi, top_n)
        except Exception:
            return None
        blob = " ".join(
            ln.text.replace("\u2019", "'").replace("\u00A0", " ") for ln in lines
        )
        found = {k for k, rx in NJ_FUND_PATTERNS.items() if rx.search(blob)}
        for key in _NJ_FUND_PRIORITY:
            if key in found:
                return key
        return None

    hit = _scan(page_idx)
    if hit:
        return hit
    for back in range(1, lookback + 1):
        pi = page_idx - back
        if pi < 0:
            break
        hit = _scan(pi)
        if hit:
            return hit
    return None

def validate_single_statement(doc, step_code, sector, state, audit_year,
                              workbook_fye, screenshot_out,
                              candidate_pages: Optional[List[int]] = None):
    result = StepResult(code=step_code)
    spec   = STATEMENT_SPECS[step_code]
    if sector.upper() not in spec["sectors"]:
        result.status = "NOT APPLICABLE"; result.notes = f"{step_code} not applicable to {sector}"
        return result, []
    target_fye = parse_fye_from_workbook(workbook_fye)
    pages_iter = candidate_pages if candidate_pages is not None else range(doc.page_count)
    for p in pages_iter:
        title_lines = match_statement_title(doc, p, spec)
        if not title_lines:
            # ✅ FIX 1: OCR-tolerant fuzzy fallback. Previously a stray `continue`
            #          right here skipped EVERY page the fuzzy matcher accepted,
            #          rendering the fallback dead. Now we fall through to the
            #          row/column structure checks.
            if _fuzzy_title_match_page(doc, p, step_code):
                title_lines = page_top_lines(doc, p, G3_HEADING_TOP_LINES)
            else:
                continue

        # ── PER-PAGE OCR GUARD (LG) ──
        # Title matched, but if the page is a scanned image the row/column
        # text check below would wrongly fail. Flag OCR instead.
        if page_needs_ocr(doc, p):
            result.status = OCR_STATUS; result.score = 0.0
            result.page_idx = p; result.page_range = str(p)
            result.notes = ("Statement title found but page is image-only "
                            "(table body not in text layer); OCR required.")
            shot = build_screenshot_name(sector, state, audit_year, step_code, p)
            result.screenshot = str(save_highlighted_screenshot(
                doc, p, [title_lines[0].bbox] if title_lines else [],
                CLR_LIGHT_YELLOW, screenshot_out, shot))
            return result, [p]

        rows_ok, pages_covered, matched_lines = _check_rows(doc, p, spec)
        cols_ok = _check_columns(doc, p, spec)

        if not (rows_ok and cols_ok):
            continue

        fye_cat = _fye_on_statement_pages(doc, pages_covered, target_fye)

        if fye_cat == "EXACT":
            status, score, highlight = "PASS", 100.0, CLR_LIGHT_GREEN
        elif fye_cat == "PARTIAL":
            status, score, highlight = "REVIEW", 85.0, CLR_LIGHT_YELLOW
        else:
            status, score, highlight = "REVIEW", 80.0, CLR_LIGHT_YELLOW

        bboxes = [ln.bbox for ln in title_lines] + [ln.bbox for ln in matched_lines[:2]]
        shot_name = build_screenshot_name(sector, state, audit_year, step_code, p)
        shot_path = save_highlighted_screenshot(
            doc, p, bboxes, highlight, screenshot_out, shot_name
        )

        page_range_str = (f"{pages_covered[0]+1}-{pages_covered[-1]+1}"
                          if len(pages_covered) > 1 else f"{p+1}")
        result.status     = status
        result.score      = score
        result.page_idx   = p
        result.screenshot = str(shot_path)
        result.page_range = page_range_str
        result.notes      = (f"{spec['label']} found; rows={rows_ok}, cols={cols_ok}, "
                             f"fye={fye_cat}, pages={page_range_str}")
        return result, pages_covered

    result.status = "FAIL"
    result.notes  = f"{spec['label']} not found in report."
    result.screenshot = take_diagnostic_screenshot(
        doc, 0, sector, state, audit_year, step_code, screenshot_out, status="FAIL")
    return result, []

def aggregate_gate3(sector: str, results: Dict[str, StepResult]) -> str:
    """Per spec + BTA-tolerance:
         • G3-1 (Cash Receipts/Disbursements) found → automatic PASS
         • LG-declared entity may satisfy either LG-style OR NONLG-style
         • Any missing → FAIL; Any REVIEW (FYE mismatch) → REVIEW"""
    sector = sector.upper()

    g31 = results.get("G3-1")
    if g31 and g31.status in ("PASS", "REVIEW"):
        return "PASS"

    def _evaluate(step_codes):
        statuses = [results[k].status for k in step_codes if k in results]
        if not statuses:
            return "FAIL"
        if any(s in ("FAIL", "NOT FOUND", "NOT APPLICABLE") for s in statuses):
            return "FAIL"
        if any(s == "REVIEW" for s in statuses):
            return "REVIEW"
        if all(s == "PASS" for s in statuses):
            return "PASS"
        return "REVIEW"

    lg_result    = _evaluate(G3_LG_STEPS[1:])       # G3-2, G3-3, G3-4, G3-5
    nonlg_result = _evaluate(G3_NONLG_STEPS[1:])    # G3-2, G3-4, G3-6

    if sector == "LG":
        if lg_result == "PASS":    return "PASS"
        if nonlg_result == "PASS": return "PASS"    # ← BTA-style LG accepted
        if "REVIEW" in (lg_result, nonlg_result): return "REVIEW"
        return "FAIL"
    else:
        return nonlg_result

def overall_validation_status(g1: str, g2: str, g3: str) -> str:
    """
    Priority:
         1. OCR REQUIRED — any gate needs OCR → issuer can't be text-validated.
         2. Fail-loud    — any crashed/unknown gate → FAIL (never silent REVIEW).
         3. FAIL → REVIEW → PASS otherwise.
    """
    gates = [str(g).upper().strip() for g in (g1, g2, g3)]
    _KNOWN = {"PASS", "REVIEW", "FAIL", OCR_STATUS}
    
    if any(g == OCR_STATUS for g in gates):  return OCR_STATUS
    if any(g not in _KNOWN for g in gates):  return "FAIL"    # fail-loud
    if any(g == "FAIL"   for g in gates):    return "FAIL"
    if any(g == "REVIEW" for g in gates):    return "REVIEW"
    if all(g == "PASS"   for g in gates):    return "PASS"
    return "FAIL"


# ============================================================================
# ============================================================================
# SEAM: DB INTEGRATION  (work queue -> engine -> DB status + parquet log)
# ============================================================================
# ============================================================================

def validate_pdf(pdf_path, *, entity_name, state, fye, audit_year,
                 auditor_name, sector, screenshot_dir):
    """
    Run the full Gate 1/2/3 engine on ONE already-downloaded PDF and return a
    result dict. All engine/PDF errors are caught here so a bad file becomes a
    FAIL result instead of crashing the caller.

    Returns:
        {
          "overall": "PASS"|"REVIEW"|"FAIL"|"OCR REQUIRED",
          "gate1"/"gate2"/"gate3": gate status str,
          "ocr_required": bool,
          "framework": {display_text, confidence, confidence_pct, ...},
          "steps": {"G1-1": StepResult, ...},
          "step_status": {"G1-1": "PASS", ...},
          "error": None | str,
        }
    """
    try:
        doc = fitz.open(str(pdf_path))
    except Exception as ex:
        return {"overall": "FAIL", "gate1": "FAIL", "gate2": "FAIL", "gate3": "FAIL",
                "ocr_required": False,
                "framework": {"display_text": "", "confidence": "", "confidence_pct": 0},
                "steps": {}, "step_status": {}, "error": f"PDF open failed: {ex}"}
    try:
        step_times = {}

        # Statement/heading page map (used to scope Gate 1 searches).
        try:
            heading_cache = {p: page_heading_status(doc, p) for p in range(doc.page_count)}
            stmt_pages = sorted([p for p, s in heading_cache.items() if s != "EXCLUDED"])
        except Exception:
            stmt_pages = list(range(doc.page_count))

        ocr_required = doc_requires_ocr(doc)

        try:
            framework = detect_framework_and_basis(doc)
        except Exception:
            framework = {"display_text": "", "confidence": "", "confidence_pct": 0}

        # ---- Gate 1 ----
        try:
            g1, g1r = run_gate1_timed(doc, entity_name, state, fye, sector,
                                      audit_year, screenshot_dir, stmt_pages, step_times)
        except Exception as ex:
            print(f"   [WARN] Gate 1 crashed: {ex}")
            g1, g1r = "FAIL", {}

        # ---- Gate 2 ----
        try:
            g2, g2r = run_gate2_timed(doc, fye, auditor_name, sector, state,
                                      audit_year, screenshot_dir, step_times)
        except Exception as ex:
            print(f"   [WARN] Gate 2 crashed: {ex}")
            g2, g2r = "FAIL", {}

        # ---- Gate 3 ----
        try:
            g3, g3r, _ = run_gate3_timed(doc, sector, state, audit_year, fye,
                                         screenshot_dir, step_times)
        except Exception as ex:
            print(f"   [WARN] Gate 3 crashed: {ex}")
            g3, g3r = "FAIL", {}

        # OCR overlay: stamp "OCR REQUIRED" at step + gate level BEFORE aggregation.
        g1 = apply_ocr_overlay(ocr_required, g1, g1r)
        g2 = apply_ocr_overlay(ocr_required, g2, g2r)
        g3 = apply_ocr_overlay(ocr_required, g3, g3r)

        overall = overall_validation_status(g1, g2, g3)

        steps = {}
        steps.update(g1r or {})
        steps.update(g2r or {})
        steps.update(g3r or {})
        step_status = {c: getattr(r, "status", str(r)) for c, r in steps.items()}

        return {"overall": overall, "gate1": g1, "gate2": g2, "gate3": g3,
                "ocr_required": ocr_required, "framework": framework,
                "steps": steps, "step_status": step_status, "error": None}
    finally:
        try:
            doc.close()
        except Exception:
            pass


# Column list + joins shared by the batch and single-id queries. Callers append
# their own WHERE clause (mirrors PFG_Sourcing.py's _SOURCING_SELECT).
_VALIDATION_SELECT = """
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
        ps.PdfFilePath,
        ps.COAID,
        co.SegmentId AS SegmentId
    FROM TCompanyMaster cm
    JOIN TProcessStatus ps
        ON ps.CompanyId = cm.CompanyId
    LEFT JOIN TCOAMaster co
        ON co.COAID = ps.COAID
"""


def sector_from_segment(seg_id):
    """SegmentId 1 = LG, 2 = NONLG (business rule; ignore DB_schema.xlsx demo values).
    Default to LG when the segment is unknown/NULL — the LG Gate-3 path also probes
    the NONLG statement combination, so LG is the safe default."""
    try:
        seg_id = int(seg_id)
    except (TypeError, ValueError):
        return "LG"
    return "LG" if seg_id == 1 else ("NONLG" if seg_id == 2 else "LG")


def run_validations_for_row(db, row_id, processing_id, issuer_name, master_state_abbr,
                            uei, fy_end_val, pdf_saved, sector, audit_year):
    """
    Validate ONE already-downloaded PDF with the gate engine and log the result to
    the parquet file. Returns the validate_pdf() result dict (the caller maps
    out["overall"] to the DB status / folder / remark).
    """
    # Auditor firm name from TProcessingAdditionalInfo. This table links 1:1 to the
    # physical TProcessStatus row via its ID column (FK AInfo.ID -> TProcessStatus.Id),
    # so key on row_id — NOT ProcessingId (removed from AInfo in the schema change).
    firm_res = db.fetch_one("""
        SELECT AuditorFirmName
        FROM TProcessingAdditionalInfo
        WHERE ID = ?
    """, row_id)
    firm_row = firm_res.data if firm_res.success else None
    auditor_firm_raw = str(firm_row.AuditorFirmName) if (firm_row and firm_row.AuditorFirmName is not None) else ""

    # Per-issuer proof/screenshot directory: PROOF_ROOT / <date> / <STATE>_<Issuer>.
    screenshot_dir = (PROOF_ROOT / datetime.now().strftime("%Y-%m-%d")
                      / f"{normalize(master_state_abbr)}_{sanitize_filename(issuer_name)}")
    try:
        screenshot_dir.mkdir(parents=True, exist_ok=True)
    except Exception:
        screenshot_dir = PROOF_ROOT

    out = validate_pdf(
        pdf_saved,
        entity_name=issuer_name,
        state=master_state_abbr,
        fye=fy_end_val,
        audit_year=audit_year,
        auditor_name=auditor_firm_raw,
        sector=sector,
        screenshot_dir=screenshot_dir,
    )

    fw = out.get("framework") or {}
    row_out = {
        "ISSUER NAME": issuer_name,
        "UEI": uei,
        "OVERALL": out.get("overall", "FAIL"),
        "GATE1": out.get("gate1", "FAIL"),
        "GATE2": out.get("gate2", "FAIL"),
        "GATE3": out.get("gate3", "FAIL"),
        "OCR_REQUIRED": str(out.get("ocr_required", False)),
        "FRAMEWORK": fw.get("display_text", ""),
        "FRAMEWORK_CONFIDENCE": f'{fw.get("confidence", "")} ({fw.get("confidence_pct", 0)}%)',
    }
    row_out.update(out.get("step_status") or {})   # G1-1 .. G3-6
    write_validation_row(row_id, processing_id, row_out)
    return out


_ALL_STEP_CODES = ["G1-1", "G1-2", "G1-3", "G2-1", "G2-2", "G2-3", "G2-4",
                   "G3-1", "G3-2", "G3-3", "G3-4", "G3-5", "G3-6"]


def _process_validation_row(db, db_row):
    """
    Validate ONE work-list row end to end. Wrapped in try/except so a single bad
    row records its own Remarks and the batch keeps going: Database autocommits
    each write, but an *uncaught* exception would make the `with Database()` block
    roll back every prior row's committed status.

    Called by BOTH the batch loop and the single-id path so the steps, parquet
    write, verdict mapping, PDF move and DB update are identical between the two.
    """
    row_id        = db_row.RowId
    processing_id = db_row.ProcessingId
    issuer_name   = normalize(db_row.IssuerName)
    state_abbr    = normalize(db_row.State)
    uei           = normalize(db_row.UEI)
    fy_end_val    = db_row.FyeDate
    audit_year    = db_row.ProcessYear
    segment_id    = getattr(db_row, "SegmentId", None)
    sector        = sector_from_segment(segment_id)
    pdf_file_path = normalize(db_row.PdfFilePath)

    print(f"\n[ROW] Id={row_id} ProcessingId={processing_id} | {issuer_name} "
          f"({state_abbr}) sector={sector} year={audit_year}")

    # Display log: bind these rows to this document; a new row_id auto-flushes the
    # previous row's buffered rows. `audit_year` is TProcessStatus.ProcessYear and
    # selects the file this row is written to (user_display_log_<year>.parquet).
    udl.set_context(row_id, processing_id, year=audit_year)
    udl.started("Validation")

    # Mark in-progress AND wipe any stale remark from a prior run (Remarks -> NULL).
    update_sourcing_validation_status(db, row_id, None, 'p', clear_remarks=True)

    try:
        # ---- locate the downloaded PDF ----
        pdf_saved = Path(pdf_file_path) if pdf_file_path else None
        if (pdf_saved is None) or (not pdf_saved.exists()):
            print(f"[WARN] Id={row_id}: PDF not found at {pdf_file_path!r}")
            fail_row = {"ISSUER NAME": issuer_name, "UEI": uei,
                        "OVERALL": "FAIL", "GATE1": "FAIL", "GATE2": "FAIL", "GATE3": "FAIL",
                        "OCR_REQUIRED": "False", "FRAMEWORK": "", "FRAMEWORK_CONFIDENCE": ""}
            for c in _ALL_STEP_CODES:
                fail_row[c] = "Fail"
            write_validation_row(row_id, processing_id, fail_row)
            update_sourcing_validation_status(db, row_id, status=0, flag='c',
                                              remarks="PDF not found for validation")
            udl.failure("Validation", "PDF not found for validation")
            return

        # ---- run the gate engine ----
        out = run_validations_for_row(db, row_id, processing_id, issuer_name, state_abbr,
                                      uei, fy_end_val, pdf_saved, sector, audit_year)
        overall = out.get("overall", "FAIL")
        step_status = out.get("step_status") or {}

        # Display log: one check row per gate step that ran, then the verdict.
        for _code in _ALL_STEP_CODES:
            _st = step_status.get(_code)
            if _st:
                udl.check("Validation", _CHECK_LABELS.get(_code, _code), str(_st))
        udl.check("Validation", "Overall", overall)

        # ---- verdict mapping: PASS=1, everything else=0 (+ remark + folder) ----
        if overall == "PASS":
            status, flag, dest_dir, remark = 1, 'c', VALIDATED_PDF_DIR, None
        elif overall == "REVIEW":
            codes = [c for c, s in step_status.items() if s == "REVIEW"]
            status, flag, dest_dir = 0, 'c', REVIEW_PDF_DIR
            remark = "REVIEW: " + (", ".join(sorted(codes)) if codes else "manual review required")
        elif overall == "OCR REQUIRED":
            status, flag, dest_dir, remark = 0, 'c', OCR_REQUIRED_PDF_DIR, "OCR REQUIRED"
        else:  # FAIL or any unknown status
            status, flag, dest_dir = 0, 'c', FAILED_VALIDATION_PDF_DIR
            if out.get("error"):
                remark = f"FAIL: {out['error']}"
            else:
                codes = [c for c, s in step_status.items() if s in ("FAIL", "NOT FOUND")]
                remark = "FAIL: " + (", ".join(sorted(codes)) if codes else "did not pass validation")

        # ---- move the PDF to the verdict folder ----
        final_path = pdf_saved
        new_path, move_err = move_pdf(pdf_saved, dest_dir)
        if new_path is not None:
            final_path = new_path
        else:
            note = f"[move-failed, kept at source: {move_err}]"
            remark = note if not remark else f"{remark} {note}"

        if remark is not None:
            remark = remark[:900]

        update_sourcing_validation_status(db, row_id, status=status, flag=flag,
                                          pdf_path=final_path, remarks=remark)
        print(f"[DONE] Id={row_id}: {overall} -> status={status}, "
              f"moved={new_path is not None} -> {dest_dir.name}")

    except Exception as e:
        msg = f"Validation exception: {type(e).__name__}: {e}"[:900]
        try:
            update_sourcing_validation_status(db, row_id, status=0, flag='c', remarks=msg)
        except Exception as e2:
            print(f"[ERROR] Could not persist exception remark for Id={row_id}: {e2}")
        udl.failure("Validation", msg)
        print(f"[ERROR] Id={row_id}: {msg}")
        # best-effort: move the PDF to the failed folder if it still exists at source
        try:
            if pdf_file_path and Path(pdf_file_path).exists():
                move_pdf(Path(pdf_file_path), FAILED_VALIDATION_PDF_DIR)
        except Exception:
            pass


# =========================================================
# BATCH MODE  (no CLI arg -> validate every pending row)
# =========================================================
def main_validate():
    with Database() as db:
        print("[DB] Connected. Fetching pending sourcing-validation rows...")
        res = db.fetch_all(_VALIDATION_SELECT + """
            WHERE cm.IsActive = 1
              AND ps.IsActive = 1
              AND ps.SourcingStatus = 1
              AND ps.SourcingFlag = 'c'
              AND (ps.SourcingValidationFlag = 'c' OR ps.SourcingValidationFlag IS NULL)
              AND (ps.SourcingValidationStatus = 0 OR ps.SourcingValidationStatus IS NULL)
              AND (ps.CompletionStatus = 0 OR ps.CompletionStatus IS NULL)
              AND ps.COAID IN (1, 2)
        """)
        if not res.success:
            print(f"[ERROR] Work-list query failed: {res.error}")
            return
        master_rows = res.data or []
        print(f"[INFO] {len(master_rows)} row(s) pending validation.")
        if not master_rows:
            return

        # Claim the batch (flag 's') so a concurrent run won't pick the same rows.
        row_ids = [r.RowId for r in master_rows]
        clause, params = build_in_clause("Id", row_ids)
        claim = db.update(f"""
            UPDATE TProcessStatus
            SET SourcingValidationFlag = 's'
            WHERE IsActive = 1
              AND {clause}
        """, params)
        if not claim.success:
            print(f"[ERROR] Could not claim rows: {claim.error}")
            return

        udl.register_stages(_DISPLAY_STAGES)
        try:
            for db_row in master_rows:
                _process_validation_row(db, db_row)
        finally:
            udl.flush_all()   # commit the last row's display-log rows

    print("[DONE] Batch validation complete.")


# =========================================================
# SINGLE-ID MODE  (python PFG_Validation.py <id>)
# =========================================================
def validate_one_id(db, row_id):
    """
    Validate a SINGLE TProcessStatus row by its physical Id, using an already-open
    db handle. Filters ONLY on IsActive (not the SourcingValidation flags): an
    explicit Id is validated on demand regardless of its current flags/status —
    same rationale as PFG_Sourcing.py's source_one_id().
    """
    res = db.fetch_one(_VALIDATION_SELECT + """
        WHERE ps.Id = ? AND cm.IsActive = 1 AND ps.IsActive = 1
    """, [row_id])
    if not res.success or res.data is None:
        print(f"[INFO] No active TProcessStatus row for Id={row_id}.")
        return False
    db.update("UPDATE TProcessStatus SET SourcingValidationFlag='s' WHERE IsActive=1 AND Id=?", [row_id])
    udl.register_stages(_DISPLAY_STAGES)
    try:
        _process_validation_row(db, res.data)   # SAME per-row path as batch
    finally:
        udl.flush_all()   # commit this row's display-log rows
    return True


def main_validate_by_id(row_id):
    with Database() as db:
        print(f"[DB] Connected. Validating single row_id={row_id}.")
        validate_one_id(db, row_id)
    print(f"[DONE] Validation completed for row_id={row_id}.")


if __name__ == "__main__":
    # python PFG_Validation.py        -> validate ALL pending rows (batch)
    # python PFG_Validation.py <id>   -> validate ONE TProcessStatus row by its Id
    if len(sys.argv) > 1:
        main_validate_by_id(int(sys.argv[1]))
    else:
        main_validate()
