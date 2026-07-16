# ============================================================================
# US LOCAL GOVT AUDIT SOURCING & VALIDATION PLATFORM
# ============================================================================

import sys, importlib, subprocess

# ---------- PART 1: DEPENDENCY CHECK ----------
# TWEAK ZONE #1: Update REQUIRED_PKGS if you want stricter/looser versions

REQUIRED_PKGS = {
    "pandas":           "2.0.0",
    "openpyxl":         "3.1.0",
    "requests":         "2.28.0",
    "pymupdf":          "1.23.0",     # imported as `fitz`
    "rapidfuzz":        "3.0.0",
    "pillow":           "9.0.0",      # imported as `PIL`
    "tqdm":             "4.60.0",
    "python-dateutil":  "2.8.0",
    "certifi":          "2023.7.22",  # SSL CA bundle (updated for US-govt S3 endpoints)
    "pip-system-certs": "4.0",        # Auto-injects Windows/Corporate CA trust store
    "pyarrow":          "12.0.0",     # Fast CSV/Parquet engine
}

def _ver_tuple(v):
    return tuple(int(x) for x in v.split(".")[:3] if x.isdigit())

def _check_dependencies():
    """
    Verify all required packages meet minimum versions.   
    Uses importlib.metadata as authoritative source (works for all pip packages),
    with module.__version__ as a fast-path fallback.
    """
    from importlib.metadata import version as _pkg_version, PackageNotFoundError

    missing = []
    for pkg, min_ver in REQUIRED_PKGS.items():
        mod_name = {"pymupdf": "fitz", "pillow": "PIL",
                    "python-dateutil": "dateutil",
                    "pip-system-certs": "pip_system_certs"}.get(pkg, pkg)

        # Step 1: Try importlib.metadata first (authoritative, works for everything)
        inst_ver = None
        try:
            inst_ver = _pkg_version(pkg)
        except PackageNotFoundError:
            pass

        # Step 2: Fast-path fallback — try module.__version__
        if not inst_ver:
            try:
                mod = importlib.import_module(mod_name)
                inst_ver = getattr(mod, "__version__", None) \
                        or getattr(mod, "VERSION",     None) \
                        or getattr(mod, "version",     None)
                if inst_ver and not isinstance(inst_ver, str):
                    inst_ver = str(inst_ver)   # some packages return tuples
            except ImportError:
                missing.append(f"{pkg} (NOT INSTALLED, need >= {min_ver})")
                continue

        # Step 3: Compare versions
        if not inst_ver:
            missing.append(f"{pkg} (version undetectable, need >= {min_ver})")
        elif _ver_tuple(inst_ver) < _ver_tuple(min_ver):
            missing.append(f"{pkg} (installed {inst_ver}, need >= {min_ver})")

    if missing:
        print("❌ ABORTING: Missing / outdated dependencies:")
        for m in missing:
            print(f"   pip install --upgrade '{m.split(' ')[0]}'  →  {m}")
        sys.exit(1)
    print("✅ All required libraries present.")

_check_dependencies()

# ---------- STANDARD IMPORTS ----------
import os, re, io, json, time, shutil, hashlib, warnings, unicodedata
from pathlib import Path
from datetime import datetime, date
from dataclasses import dataclass, field
from typing import Optional, List, Dict, Tuple, Any

import pandas as pd
import requests
import fitz  # PyMuPDF
from rapidfuzz import fuzz
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter
from PIL import Image
from dateutil import parser as dtparser
from tqdm import tqdm

warnings.filterwarnings("ignore")

# ---------- CONSTANTS (TWEAK ZONE #2) ----------
FAAC_SEARCH_URL   = "https://app.fac.gov/dissemination/search/"
FAAC_CSV_URL      = "https://app.fac.gov/dissemination/public-data/gsa/full/general.csv"
FAAC_PDF_PREFIX   = "https://app.fac.gov/dissemination/report/pdf/"   # + report_id
CURRENT_YEAR      = datetime.now().year
LATEST_SEARCH_YEAR = CURRENT_YEAR   # dynamically discovered later

# ---------- Highlight & Font Colors (deep = font, light = fill) ----------
CLR_LIGHT_GREEN  = "C6EFCE"
CLR_LIGHT_YELLOW = "FFEB9C"
CLR_LIGHT_RED    = "FFC7CE"
CLR_DEEP_GREEN   = "006100"
CLR_DEEP_YELLOW  = "9C6500"
CLR_DEEP_RED     = "9C0006"

# ---------- Cell Fills (kept for API compatibility; per user spec: never applied) ----------
FILL_GREEN  = PatternFill("solid", fgColor=CLR_LIGHT_GREEN)
FILL_YELLOW = PatternFill("solid", fgColor=CLR_LIGHT_YELLOW)
FILL_RED    = PatternFill("solid", fgColor=CLR_LIGHT_RED)

# ---------- Font Family (TWEAK ZONE #2A) ----------
WORKBOOK_FONT = "Aptos Narrow"
WORKBOOK_FONT_SIZE = 9

# ---------- Named Fonts for status coloring ----------
FONT_GREEN   = Font(name=WORKBOOK_FONT, size=WORKBOOK_FONT_SIZE, color=CLR_DEEP_GREEN, bold=True)
FONT_YELLOW  = Font(name=WORKBOOK_FONT, size=WORKBOOK_FONT_SIZE, color=CLR_DEEP_YELLOW, bold=True)
FONT_RED     = Font(name=WORKBOOK_FONT, size=WORKBOOK_FONT_SIZE, color=CLR_DEEP_RED, bold=True)
FONT_DEFAULT = Font(name=WORKBOOK_FONT, size=WORKBOOK_FONT_SIZE, color="000000", bold=False)
FONT_TIME = Font(name=WORKBOOK_FONT, size=WORKBOOK_FONT_SIZE, color="757171", bold=False)
FONT_HEADER  = Font(name=WORKBOOK_FONT, size=WORKBOOK_FONT_SIZE, color="000000", bold=True)

# ---------- Validation Step Codes ----------
G1_STEPS = ["G1-1", "G1-2", "G1-3"]
G2_STEPS = ["G2-1", "G2-2", "G2-3", "G2-4"]
G3_LG_STEPS    = ["G3-1", "G3-2", "G3-3", "G3-4", "G3-5"]
G3_NONLG_STEPS = ["G3-1", "G3-2", "G3-4", "G3-6"]

# ---------- STATE CODE ↔ FULL NAME MAP ----------
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

# ---------- ENTITY NAME NORMALIZATION TOKENS ----------
# TWEAK ZONE #3: add / remove suffixes to strip during entity-name matching
STOPWORDS = {
    "the","of","and","for","in","on","at","to","a","an","by","with","or",
    "inc","incorporated","llc","plc","ltd","co","corp","corporation","company",
    "dba","dbaof","aka","akaof","fka","fkaof",
    "d/b/a","d.b.a","a/k/a","a.k.a","f/k/a","f.k.a"
}
SPECIAL_CHARS_RE = re.compile(r"[^\w\s]")   # keeps alphanumerics + underscore + space

# ---------- PART 1: FOLDER ARCHITECTURE ----------
def get_desktop_path() -> Path:
    """
    Return the current user's Desktop path, prioritizing OneDrive-backed
    Desktops (common on corporate Windows machines).
    Falls back to non-OneDrive Desktop last.
    """
    home = Path.home()

    # Priority order: OneDrive-backed locations first (they're the active ones
    # on most modern Windows corporate setups). Only fall back to plain Desktop
    # if no OneDrive location exists.
    candidates = [
        home / "OneDrive - Default Directory" / "Desktop",    # Corp OneDrive
        home / "OneDrive - Personal" / "Desktop",             # Newer personal OneDrive
        home / "OneDrive" / "Desktop",                        # Older OneDrive
        home / "Desktop",                                     # Plain Desktop (last resort)
    ]

    for candidate in candidates:
        if candidate.exists() and candidate.is_dir():
            return candidate

    # Last resort: create plain Desktop
    fallback = home / "Desktop"
    fallback.mkdir(parents=True, exist_ok=True)
    return fallback

def ensure_architecture() -> Dict[str, Path]:
    """Create US GOVT AUDIT structure on Desktop; return dict of key paths."""
    root = get_desktop_path() / "US GOVT AUDIT"
    root.mkdir(parents=True, exist_ok=True)
    paths = {
        "ROOT":              root,
        "DOWNLOADED":        root / "DOWNLOADED REPORTS",
        "FAC_CSV":           root / "FAC_CSV_DATA",
        "VALIDATED":         root / "VALIDATED REPORTS",
        "SAMPLES":           root / "VALIDATION SAMPLES",
        "WORKBOOK":          root / "ENTITY MASTER.xlsx",
    }
    for k, p in paths.items():
        if k not in ("WORKBOOK",):
            p.mkdir(parents=True, exist_ok=True)
    return paths

def wait_for_workbook(wb_path: Path):
    """Halt if workbook missing; wait for user 'Continue'."""
    if wb_path.exists():
        return
    print(f"⚠️  ENTITY MASTER.xlsx not found at {wb_path}")
    print("    Please drop the workbook there, then type 'Continue' and hit Enter.")
    while True:
        resp = input("→ ").strip().lower()
        if resp == "continue" and wb_path.exists():
            print("✅ Workbook detected — resuming."); return
        elif resp == "continue":
            print("   Still missing — try again."); continue

def pick_workbook_via_dialog():
    """
    Open a file-browser (always-on-top) so user can pick the workbook.
    Returns None if user cancels or tkinter is unavailable (headless).
    """
    try:
        import tkinter as tk
        from tkinter import filedialog
        root = tk.Tk(); root.withdraw()
        root.attributes("-topmost", True); root.lift()
        p = filedialog.askopenfilename(
            title="Select ENTITY MASTER.xlsx",
            filetypes=[("Excel workbook", "*.xlsx *.xlsm")])
        root.destroy()
        return Path(p) if p else None
    except Exception:
        return None

def resolve_workbook(default_wb: Path) -> Path:
    """Ask user how to supply workbook: default location, path string, or dialog."""
    print("\n📂 How would you like to provide the ENTITY MASTER workbook?")
    print("   [1] Use default location (Desktop/US GOVT AUDIT/ENTITY MASTER.xlsx)")
    print("   [2] Type a file path")
    print("   [3] Open file-browser dialog")
    choice = (input("Enter 1/2/3 [default=1]: ").strip() or "1")
    if choice == "2":
        p = Path(input("Full path to workbook: ").strip('" '))
    elif choice == "3":
        p = pick_workbook_via_dialog() or default_wb
    else:
        p = default_wb
    wait_for_workbook(p)
    return p

        ## ---------- PART 1: ENSURE WORKBOOK IS NOT OPEN (Excel lock guard) ----------
## Goal: before openpyxl touches ENTITY MASTER.xlsx, make sure Excel isn't
## holding it open (which would block the later save with PermissionError).
##   • OPEN  → close it (optionally Save first), then continue.
##   • CLOSED → continue immediately.
## Windows-first via win32com; degrades gracefully if pywin32/COM unavailable.

def _excel_lock_file(wb_path: Path) -> Path:
    "Return the sibling Excel owner-lock file path (e.g. '~$ENTITY MASTER.xlsx')."
    return wb_path.with_name("~$" + wb_path.name)


def _is_file_locked(wb_path: Path) -> bool:
    """
    True if the workbook appears LOCKED by another process (Excel).
    Uses an atomic self-rename: Windows refuses to rename a file that is
    open with a share-deny lock, so a failure here == 'locked/open'.
    (A plain open('r+b') can succeed even when Excel holds the file, so
    the rename probe is the reliable signal on Windows.)
    """
    if not wb_path.exists():
        return False
    try:
        wb_path.rename(wb_path)          # no-op rename to itself
        return False                     # succeeded → not locked
    except (OSError, PermissionError):
        return True                      # denied → locked/open


def _close_via_excel_com(wb_path: Path, save_before_close: bool, verbose: bool) -> bool:
    """
    Try to close the workbook through a RUNNING Excel instance (COM).
    Returns True if the matching workbook was found and closed.
    Silently returns False when pywin32/COM is unavailable or no instance
    is running (caller then falls back to the lock-file prompt).
    """
    try:
        import win32com.client as _w32           # pywin32 (optional dependency)
        import pythoncom as _pyc
    except Exception:
        return False

    target = str(wb_path).lower()
    target_name = wb_path.name.lower()
    closed_any = False

    # Grab a running Excel instance (GetActiveObject first, then GetObject).
    xl = None
    for _getter in ("GetActiveObject", "GetObject"):
        try:
            xl = (_w32.GetActiveObject("Excel.Application")
                  if _getter == "GetActiveObject"
                  else _w32.GetObject(Class="Excel.Application"))
            if xl is not None:
                break
        except Exception:
            xl = None
    if xl is None:
        return False                              # no live Excel → nothing to close

    try:
        for wb in list(xl.Workbooks):
            try:
                full = str(getattr(wb, "FullName", "")).lower()
                nm   = str(getattr(wb, "Name", "")).lower()
            except Exception:
                continue
            if full == target or nm == target_name:
                if save_before_close:
                    try:
                        wb.Save()
                        if verbose:
                            print(f"   💾 Saved open workbook before closing: {wb_path.name}")
                    except Exception as ex:
                        if verbose:
                            print(f"   ⚠️  Could not Save before close ({ex}); closing without save.")
                        wb.Close(SaveChanges=False)
                        closed_any = True
                        continue
                    wb.Close(SaveChanges=False)   # already saved above
                else:
                    wb.Close(SaveChanges=False)
                closed_any = True
                if verbose:
                    print(f"   ✅ Closed open workbook in Excel: {wb_path.name}")

        # Quit Excel only if we closed something and no workbooks remain.
        try:
            if closed_any and xl.Workbooks.Count == 0:
                xl.Quit()
                if verbose:
                    print("   ✅ Excel had no remaining workbooks — instance quit.")
        except Exception:
            pass
    except Exception as ex:
        if verbose:
            print(f"   ⚠️  Excel COM close attempt failed: {ex}")
        return closed_any

    return closed_any


def ensure_workbook_closed(wb_path: Path,
                           save_before_close: bool = True,
                           verbose: bool = True) -> bool:
    """
    Ensure ENTITY MASTER.xlsx is NOT open before the pipeline uses it.

    Parameters
    ----------
    wb_path : Path                 resolved workbook path (from resolve_workbook)
    save_before_close : bool       True → Save any unsaved edits before closing
    verbose : bool                 print progress lines

    Returns
    -------
    bool  True if the workbook is confirmed free (closed / never open),
          False if it could not be freed and the user chose to skip.
    """
    if not wb_path.exists():
        return True                                   # nothing to guard yet

    # 1) Fast path — not locked and no stale owner-lock file → continue.
    if not _is_file_locked(wb_path) and not _excel_lock_file(wb_path).exists():
        if verbose:
            print(f"🔓 Workbook is closed — continuing: {wb_path.name}")
        return True

    if verbose:
        print(f"🔒 Workbook appears OPEN/locked: {wb_path.name} — attempting to close it...")

    # 2) Preferred path — close through the live Excel instance (COM).
    if _close_via_excel_com(wb_path, save_before_close, verbose):
        time.sleep(0.7)                               # let Windows release the handle
        if not _is_file_locked(wb_path):
            _excel_lock_file(wb_path).unlink(missing_ok=True)  # clear stray ~$ lock
            if verbose:
                print("🔓 Lock released — continuing.")
            return True

    # 3) Fallback — COM unavailable or lock persisted → ask the user (like
    #    wait_for_workbook). Loop until the file is free or user skips.
    while True:
        print(f"⚠️  Please CLOSE '{wb_path.name}' in Excel, then type 'Continue' "
              f"(or 'Skip' to proceed anyway).")
        resp = input("→ ").strip().lower()
        if resp == "continue":
            _excel_lock_file(wb_path).unlink(missing_ok=True)
            if not _is_file_locked(wb_path):
                if verbose:
                    print("🔓 Workbook is now closed — continuing.")
                return True
            print("   Still open/locked — try again.")
        elif resp == "skip":
            if verbose:
                print("   ⏭️  Proceeding despite lock (save may fail).")
            return False
# ============================================================================
# PART 2: FAAC CSV DOWNLOADER + SHARED PDF/TEXT/DATE/NAME HELPERS
# ============================================================================

# ---------- 2.A  FAAC CSV DOWNLOAD & INDEXING ----------
def _file_created_date(p: Path) -> date:
    """Return the file's 'created' date (falls back to mtime on Linux)."""
    try:
        ts = p.stat().st_ctime
    except Exception:
        ts = p.stat().st_mtime
    return datetime.fromtimestamp(ts).date()

def ensure_faac_csv(csv_dir: Path, force: bool = False) -> Path:
    """
    Download the FAAC general.csv if none exists OR the newest one is stale.
    Returns the path of the freshest CSV available.
    """
    csv_dir.mkdir(parents=True, exist_ok=True)
    today = datetime.now().date()

    existing = sorted(csv_dir.glob("*.csv"), key=lambda p: p.stat().st_mtime, reverse=True)
    freshest = existing[0] if existing else None
    if freshest and (_file_created_date(freshest) >= today) and not force:
        print(f"✅ FAAC CSV is current: {freshest.name}")
        return freshest

    # Download fresh copy
    print(f"⬇️  Downloading FAAC general.csv from {FAAC_CSV_URL} ...")
    fname = csv_dir / f"general_{today:%Y%m%d}.csv"
    try:
        with HTTP.get(FAAC_CSV_URL, stream=True, timeout=60, allow_redirects=True) as r:
            r.raise_for_status()
            total = int(r.headers.get("content-length", 0))
            with open(fname, "wb") as f, tqdm(total=total, unit="B", unit_scale=True,
                                              desc="FAAC CSV") as bar:
                for chunk in r.iter_content(chunk_size=1 << 15):
                    if chunk:
                        f.write(chunk); bar.update(len(chunk))
        print(f"✅ Saved to {fname}")
        return fname
    except requests.exceptions.SSLError as ssl_ex:
        print(f"⚠️  SSL verification failed: {ssl_ex}")
        print("   Retrying with SSL verification disabled (US-govt public data)...")
        try:
            import urllib3
            urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
            with HTTP.get(FAAC_CSV_URL, stream=True, timeout=60,
                          allow_redirects=True, verify=False) as r:
                r.raise_for_status()
                total = int(r.headers.get("content-length", 0))
                with open(fname, "wb") as f, tqdm(
                    total=total, unit="B", unit_scale=True,
                    desc="FAAC CSV (no-verify)"
                ) as bar:
                    for chunk in r.iter_content(chunk_size=1 << 15):
                        if chunk:
                            f.write(chunk); bar.update(len(chunk))
            print(f"✅ Saved to {fname} (SSL bypassed)")
            return fname
        except Exception as ex2:
            print(f"❌ Even without SSL verify: {ex2}")
            if freshest:
                print(f"   Falling back to older CSV: {freshest.name}")
                return freshest
            raise
    except Exception as ex:
        print(f"❌ Download failed: {ex}")
        if freshest:
            print(f"   Falling back to older CSV: {freshest.name}")
            return freshest
        raise

# Column names we MUST have (TWEAK ZONE #4: rename here if FAAC changes schema)
FAAC_COLS = {
    "uei":            "auditee_uei",
    "year":           "audit_year",
    "state":          "auditee_state",
    "firm":           "auditor_firm_name",
    "contact":        "auditor_contact_name",
    "ein":            "auditee_ein",
    "report_id":      "report_id",
    "fy_end":         "fy_end_date",
    "entity_name":    "auditee_name",
}

def load_faac_csv(csv_path: Path) -> pd.DataFrame:
    """
    Load the FAAC CSV efficiently using a Parquet cache.
       
       Strategy:
         1. Check if a Parquet cache exists and is newer than the source CSV
         2. If yes → load from Parquet (10-30× faster, ~3-5 seconds)
         3. If no → parse CSV with PyArrow engine, cache as Parquet, then return
       
       Optimizations:
         • PyArrow engine (2-4× faster than C engine for wide files)
         • Column filtering during read (not after)
         • Native string type instead of object dtype
         • Parquet caching for subsequent runs
         • Vectorized string cleanup (single pass)
    """
    parquet_cache = csv_path.with_suffix(".parquet")
    
    # ---- Fast path: use cached Parquet if newer than CSV ----
    if parquet_cache.exists() and parquet_cache.stat().st_mtime >= csv_path.stat().st_mtime:
        print(f"⚡ Loading from Parquet cache: {parquet_cache.name}")
        t0 = time.perf_counter()
        try:
            df = pd.read_parquet(parquet_cache)
            elapsed = time.perf_counter() - t0
            print(f"   ✅ Loaded {len(df):,} rows from cache in {elapsed:.2f}s")
            return df
        except Exception as ex:
            print(f"   ⚠️  Parquet cache corrupted ({ex}); falling back to CSV")
    
    # ---- Slow path: parse CSV and build cache ----
    print(f"📖 Reading FAAC CSV: {csv_path.name}")
    print(f"   File size: {csv_path.stat().st_size / (1024*1024):.1f} MB")
    t0 = time.perf_counter()
    
    # Step 1: Peek header to identify available columns (cheap — 1KB read)
    header = pd.read_csv(csv_path, nrows=0)
    missing = [c for c in FAAC_COLS.values() if c not in header.columns]
    if missing:
        print(f"⚠️  Expected columns missing from CSV: {missing}")
        print(f"    Available columns (first 30): {list(header.columns)[:30]}")
    usecols = [c for c in FAAC_COLS.values() if c in header.columns]
    
    # Step 2: Full parse with PyArrow engine + optimized dtypes
    try:
        df = pd.read_csv(
            csv_path,
            usecols=usecols,           # only read what we need (much less I/O)
            dtype="string",            # nullable string type (faster than object)
            engine="pyarrow",          # use PyArrow parser
        )
        print(f"   Used PyArrow engine")
    except Exception as ex:
        # Fallback if PyArrow isn't available or has issues
        print(f"   PyArrow unavailable ({ex}), falling back to C engine")
        df = pd.read_csv(
            csv_path,
            usecols=usecols,
            dtype=str,
            engine="c",
            low_memory=True,           # streaming parse — much less RAM
            na_filter=False,           # skip NaN detection (all strings)
        )
    
    parse_elapsed = time.perf_counter() - t0
    print(f"   Parsed in {parse_elapsed:.2f}s")
    
    # Step 3: Vectorized cleanup (single pass, no per-column loops)
    df.columns = df.columns.str.strip()
    for c in df.columns:
        # Use pandas' vectorized string ops — much faster than .astype().str.strip()
        df[c] = df[c].fillna("").str.strip()
    df["auditee_uei"] = df["auditee_uei"].str.upper()
    df["auditee_state"] = df["auditee_state"].str.upper()
    
    # Step 4: Save as Parquet for next time (compressed, columnar, blazing fast to reload)
    try:
        cache_start = time.perf_counter()
        df.to_parquet(parquet_cache, compression="snappy", index=False)
        cache_elapsed = time.perf_counter() - cache_start
        print(f"   💾 Cached to {parquet_cache.name} in {cache_elapsed:.2f}s "
              f"({parquet_cache.stat().st_size / (1024*1024):.1f} MB)")
    except ImportError:
        print(f"   ⚠️  pyarrow not installed — no cache created. "
              f"Run: pip install pyarrow")
    except Exception as ex:
        print(f"   ⚠️  Failed to write Parquet cache: {ex}")
    
    total_elapsed = time.perf_counter() - t0
    print(f"   ✅ Rows: {len(df):,}   Columns present: {len(df.columns)}   "
          f"Total time: {total_elapsed:.2f}s")
    return df

def index_faac_by_uei_state(df: pd.DataFrame):
    """
    Pre-index by (UEI, STATE) → slice range (start, end) into sorted DataFrame.
    Returns (sorted_df, index_dict).
    """
    print(f"🔎 Building UEI × STATE index for {len(df):,} rows...")
    t0 = time.perf_counter()
    
    df = df.copy()
    df["_audit_year_num"] = pd.to_numeric(df["audit_year"], errors="coerce")
    df = df.sort_values(
        ["auditee_uei", "auditee_state", "_audit_year_num"],
        ascending=[True, True, False]
    ).reset_index(drop=True)
    df = df.drop(columns=["_audit_year_num"])
    
    uei_arr = df["auditee_uei"].values
    state_arr = df["auditee_state"].values
    
    uei_changes = uei_arr[1:] != uei_arr[:-1]
    state_changes = state_arr[1:] != state_arr[:-1]
    boundaries = uei_changes | state_changes
    
    import numpy as np
    boundary_positions = np.concatenate(([0], np.where(boundaries)[0] + 1, [len(df)]))
    
    idx = {}
    for i in range(len(boundary_positions) - 1):
        start = int(boundary_positions[i])
        end = int(boundary_positions[i + 1])
        key = (uei_arr[start], state_arr[start])
        idx[key] = (start, end)
    
    elapsed = time.perf_counter() - t0
    print(f"   ✅ Indexed {len(idx):,} (UEI, STATE) combinations in {elapsed:.2f}s")
    return df, idx           # ← returns TUPLE

# ---------- 2.B  ENTITY NAME NORMALIZATION ----------
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

# ---------- 2.C  FYE DATE PARSING & MATCHING ----------
# All the FYE spellings from your spec (do NOT hardcode PDF terms elsewhere)
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

def reconcile_fye_with_audit_year(workbook_fye,
                                    downloaded_audit_year,
                                    workbook_fy) -> Tuple[Optional[Tuple[int, int]], str]:
    """
    Reconcile the workbook FYE against the actual audit year downloaded from FAAC.
    
    Handles THREE scenarios:
    
    1. FY matches AY, FYE year matches AY:
       → No reconciliation needed. Use workbook FYE as-is.
    
    2. FY doesn't match AY (advanced/retracted year):
       → Adjust FYE year to match downloaded audit year.
       → Update FY column in workbook too.
    
    3. FY matches AY but FYE year doesn't match AY:  ← THE MISSING CASE
       → Data entry error in workbook. FY says 2025 but FYE says 6/30/2024.
       → Auto-correct FYE year to match FY/AY.
       → FYE cell gets light yellow highlight.
    
    Args:
        workbook_fye: raw value from ENTITY_LIST.FYE cell (serial, datetime, or string)
        downloaded_audit_year: audit_year from FAAC record (int or None)
        workbook_fy: raw value from ENTITY_LIST.FY cell
    
    Returns:
        (effective_fye_tuple, reconciliation_note)
    """
    base_fye = parse_fye_from_workbook(workbook_fye)
    if base_fye is None:
        return None, "workbook FYE unparseable"
    
    base_month, base_year = base_fye
    
    # Parse workbook FY
    try:
        wb_fy = int(workbook_fy) if workbook_fy is not None else base_year
    except (ValueError, TypeError):
        wb_fy = base_year
    
    # Parse downloaded audit year
    try:
        dl_ay = int(downloaded_audit_year) if downloaded_audit_year else None
    except (ValueError, TypeError):
        dl_ay = None
    
    # If we don't know the downloaded audit year, use workbook FYE as-is
    if dl_ay is None:
        return base_fye, f"using workbook FYE as-is (downloaded year unknown)"
    
    # ---- CASE 1: Everything matches — no reconciliation needed ----
    if dl_ay == wb_fy and base_year == dl_ay:
        return base_fye, f"OK - workbook FY={wb_fy}, FYE=(month={base_month},year={base_year}), AY={dl_ay}"
    
    # ---- CASE 2: Downloaded audit year is NEWER than workbook FY ----
    if dl_ay > wb_fy:
        effective_fye = (base_month, dl_ay)
        return effective_fye, (
            f"⚡ ADVANCED FYE: workbook FY={wb_fy} FYE=({base_month},{base_year}) → "
            f"downloaded AY={dl_ay} → effective FYE=({base_month},{dl_ay})"
        )
    
    # ---- CASE 3: Downloaded audit year is OLDER than workbook FY ----
    if dl_ay < wb_fy:
        effective_fye = (base_month, dl_ay)
        return effective_fye, (
            f"⚠ RETRACTED FYE: workbook FY={wb_fy} FYE=({base_month},{base_year}) → "
            f"downloaded AY={dl_ay} → effective FYE=({base_month},{dl_ay})"
        )
    
    # ---- CASE 4 (NEW): FY matches AY but FYE year is wrong ----
    # This means data entry error: workbook has FY=2025 correctly but FYE=6/30/2024
    # Fix: advance the FYE year to match the FY/AY
    if dl_ay == wb_fy and base_year != dl_ay:
        effective_fye = (base_month, dl_ay)
        return effective_fye, (
            f"🔧 FYE YEAR CORRECTED: workbook FY={wb_fy} matches AY={dl_ay}, "
            f"but FYE year=({base_year}) is stale → "
            f"effective FYE=({base_month},{dl_ay})"
        )
    
    return base_fye, "no reconciliation needed"

def _extract_original_fye_day(workbook_fye) -> int:
    """
    Extract the day-of-month from the workbook FYE cell.
    Defaults to 31 (last-day-of-month for typical fiscal year ends).
    Handles Excel serials, datetimes, dates, and date strings.
    """
    if workbook_fye is None:
        return 31
    if isinstance(workbook_fye, datetime):
        return workbook_fye.day
    if isinstance(workbook_fye, date):
        return workbook_fye.day
    if isinstance(workbook_fye, (int, float)) and not pd.isna(workbook_fye):
        try:
            serial = int(workbook_fye)
            if 32874 <= serial <= 55153:
                base = datetime(1899, 12, 30)
                dt = base + pd.Timedelta(days=serial)
                return dt.day
        except Exception:
            pass
    if isinstance(workbook_fye, str) and workbook_fye.strip():
        try:
            dt = dtparser.parse(workbook_fye.strip(),
                                dayfirst=False, fuzzy=True,
                                default=datetime(1900, 1, 1))
            if dt.year != 1900:
                return dt.day
        except Exception:
            pass
    return 31

def _format_fye_as_mmdd(workbook_fye) -> str:
    """Convert any workbook FYE representation to 'mm/dd' string format.
    
    Handles:
      • Excel serials (e.g., 45900 → "08/31")
      • datetime/date objects
      • ISO strings ("2025-08-31" → "08/31")
      • US format strings ("8/31/2025" → "08/31")
      • Long-form strings ("August 31, 2025" → "08/31")
    
    Returns empty string if unparseable.
    """
    if workbook_fye is None:
        return ""
    if isinstance(workbook_fye, float) and pd.isna(workbook_fye):
        return ""
    
    # datetime / date object
    if isinstance(workbook_fye, datetime):
        return f"{workbook_fye.month:02d}/{workbook_fye.day:02d}"
    if isinstance(workbook_fye, date):
        return f"{workbook_fye.month:02d}/{workbook_fye.day:02d}"
    
    # Excel serial
    if isinstance(workbook_fye, (int, float)) and not pd.isna(workbook_fye):
        try:
            serial = int(workbook_fye)
            if 32874 <= serial <= 55153:
                base = datetime(1899, 12, 30)
                dt = base + pd.Timedelta(days=serial)
                return f"{dt.month:02d}/{dt.day:02d}"
        except Exception:
            pass
        return ""
    
    # String — use robust FYE parser first for month/year, then _extract for day
    if isinstance(workbook_fye, str) and workbook_fye.strip():
        parsed = parse_fye_from_workbook(workbook_fye)
        if parsed:
            day = _extract_original_fye_day(workbook_fye)
            return f"{parsed02d}/{day:02d}"
    
    return ""
    
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

# Diagnostic flag — set to True to see every FYE comparison in the console
DEBUG_FYE = False    # TWEAK ZONE: enable for debugging

# ---------- 2.D  PDF TEXT & SENTENCE EXTRACTION ----------
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

    def status_style(self):
        if self.status == "PASS":   return FONT_GREEN
        if self.status == "FAIL":   return FONT_RED
        if self.status == "REVIEW": return FONT_YELLOW
        return FONT_RED

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

# ---------- 2.E  SCREENSHOT / HIGHLIGHT PRIMITIVES ----------
def _sanitize_filename(s: str, max_len: int = 34) -> str:
    """Safe filesystem name; trims to max_len chars."""
    s = re.sub(r"[^A-Za-z0-9_\- ]", "", str(s)).strip()
    s = re.sub(r"\s+", "_", s)
    return s[:max_len] or "UNKNOWN"

def screenshot_dir(paths: Dict[str, Path], state: str, entity: str) -> Path:
    """VALIDATION SAMPLES/YYYY-MM-DD/STATE_ENTITY(34)/"""
    day  = datetime.now().strftime("%Y-%m-%d")
    sub  = f"{state}_{_sanitize_filename(entity, 34)}"
    d    = paths["SAMPLES"] / day / sub
    d.mkdir(parents=True, exist_ok=True)
    return d

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

# ─────────────────────────────────────────────────────────────────────────
# DUAL PROGRESS BAR HELPERS  (outer = batch, inner = per-entity phases)
# Reused by BOTH sourcing and validation drivers.
# ─────────────────────────────────────────────────────────────────────────
def _short(txt: str, width: int = 46) -> str:
    "Clamp a label so postfix never wraps the terminal (keeps bars single-line)."
    txt = str(txt).replace("\n", " ").strip()
    return txt if len(txt) <= width else txt[:width - 1] + "…"

def make_batch_bar(total: int, label: str):
    "Outer bar — one tick per entity. Stays on line 0."
    return tqdm(total=total, position=0, leave=True,
               desc=label, unit="entity",
               bar_format="{desc} {percentage:3.0f}%|{bar:20}| "
                          "{n_fmt}/{total_fmt} [{elapsed}<{remaining}]")

def make_entity_bar(total: int):
    "Inner bar — one tick per phase within an entity. Stays on line 1, "
    "cleared after each entity (leave=False)."
    return tqdm(total=total, position=1, leave=False,
               unit="step",
               bar_format="   ↳ {desc} {percentage:3.0f}%|{bar:15}| "
                          "{n_fmt}/{total_fmt}")

# ============================================================================
# OCR-REQUIREMENT DETECTION
# ============================================================================
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
    
# ============================================================================
# PART 3: SOURCING MODULE — FAAC WEBSITE SEARCH, PDF DOWNLOAD, WORKBOOK UPDATE
# ============================================================================

# ---------- 3.A  HTTP SESSION WITH RETRIES ----------
def build_http_session() -> requests.Session:
    """Configured session with retries, backoff, browser-like User-Agent,
       and Windows-friendly SSL trust store.
       TWEAK ZONE #5: adjust timeouts, retries, proxies, SSL verify here."""
    s = requests.Session()

    # ---- SSL trust: prefer certifi bundle; fall back gracefully ----
    try:
        import certifi
        s.verify = certifi.where()
    except ImportError:
        pass   # Will use system default

    s.headers.update({
        "User-Agent": ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                       "AppleWebKit/537.36 (KHTML, like Gecko) "
                       "Chrome/124.0 Safari/537.36"),
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
        "Accept-Language": "en-US,en;q=0.9",
    })
    try:
        from requests.adapters import HTTPAdapter
        from urllib3.util.retry import Retry
        retry = Retry(total=4, backoff_factor=1.5,
                      status_forcelist=[429, 500, 502, 503, 504],
                      allowed_methods=["GET", "POST"])
        adapter = HTTPAdapter(max_retries=retry, pool_connections=10, pool_maxsize=10)
        s.mount("https://", adapter); s.mount("http://", adapter)
    except Exception:
        pass
    return s

HTTP = build_http_session()

# ---------- 3.B  DISCOVER LATEST AVAILABLE AUDIT YEAR ----------
FAAC_SEARCH_API = "https://app.fac.gov/dissemination/search/"     # HTML search form
# TWEAK ZONE #6: the FAAC public API endpoint if you have credentials; otherwise
# we rely on CSV metadata which is authoritative.

def discover_latest_audit_year(faac_df: pd.DataFrame) -> int:
    """Return the highest audit_year present in the CSV — this is the 'latest'
       audit year FAAC has published, regardless of calendar year."""
    try:
        years = pd.to_numeric(faac_df["audit_year"], errors="coerce").dropna().astype(int)
        latest = int(years.max())
        print(f"🔎 Latest audit year in FAAC CSV: {latest}")
        return latest
    except Exception:
        print(f"⚠️  Could not determine latest audit year; defaulting to {CURRENT_YEAR}")
        return CURRENT_YEAR

# ---------- 3.C  FAAC RECORD LOOKUP ----------
def find_faac_record(sorted_df, faac_index, uei, state, latest_year, min_year):
    """Find the newest FAAC record matching (UEI, STATE) within the year range."""
    if not uei or not state:
        return None
    key = (str(uei).upper().strip(), str(state).upper().strip())
    if key not in faac_index:
        return None
    
    start, end = faac_index[key]           # ← index gives (start, end) tuple
    sub = sorted_df.iloc[start:end]        # ← slice from sorted_df
    
    yrs = pd.to_numeric(sub["audit_year"], errors="coerce")
    sort_cols = [c for c in ("is_public", "fac_accepted_date") if c in sub.columns]
    for yr in range(int(latest_year), int(min_year) - 1, -1):
        hit = sub[yrs == yr]
        if not hit.empty:
            if sort_cols:
                hit = hit.sort_values(sort_cols, ascending=[False] * len(sort_cols))
            return hit.iloc[0]
    return None

def build_pdf_url(report_id: str) -> str:
    """Standard FAAC PDF URL pattern."""
    rid = str(report_id).strip()
    return f"{FAAC_PDF_PREFIX}{rid}" if rid else ""

# ---------- 3.D  PDF DOWNLOAD ----------
def download_pdf(url: str, dest_path: Path, label: str = "") -> bool:
    """Stream-download a PDF WITHOUT inner tqdm bar (keeps parent bar single-line).
       Returns True on success."""
    if not url:
        return False
    tmp = dest_path.with_suffix(dest_path.suffix + ".part")
    try:
        with HTTP.get(url, stream=True, timeout=(15, 90)) as r:
            r.raise_for_status()
            total = int(r.headers.get("content-length", 0))
            size_mb = total / (1024 * 1024) if total else 0
            # Single-line status; parent tqdm bar remains uninterrupted
            tqdm.write(f"   ⬇  Downloading {label[:40]} ({size_mb:.1f} MB) ...")
            with open(tmp, "wb") as f:
                for chunk in r.iter_content(chunk_size=1 << 15):
                    if chunk:
                        f.write(chunk)
        # Sanity check
        with open(tmp, "rb") as f:
            if f.read(5) != b"%PDF-":
                tmp.unlink(missing_ok=True)
                return False
        # Retry .replace up to 3 times for Windows AV lock tolerance
        for attempt in range(3):
            try:
                tmp.replace(dest_path)
                break
            except (OSError, PermissionError) as ex:
                if attempt == 2:
                    tqdm.write(f"   ❌ Cannot rename after 3 attempts: {ex}")
                    return False
                time.sleep(0.5 * (attempt + 1))
        if not dest_path.exists() or dest_path.stat().st_size < 1024:
            tqdm.write(f"   ❌ File missing or too small: {dest_path}")
            return False
        return True
    except requests.exceptions.SSLError:
        try:
            with HTTP.get(url, stream=True, timeout=(15, 90), verify=False) as r:
                r.raise_for_status()
                tqdm.write(f"   ⬇  Downloading {label[:40]} [no-verify] ...")
                with open(tmp, "wb") as f:
                    for chunk in r.iter_content(chunk_size=1 << 15):
                        if chunk:
                            f.write(chunk)
            with open(tmp, "rb") as f:
                if f.read(5) != b"%PDF-":
                    tmp.unlink(missing_ok=True); return False
            tmp.replace(dest_path)
            return True
        except Exception as ex2:
            tqdm.write(f"   ❌ Download error even without SSL verify ({label}): {ex2}")
            tmp.unlink(missing_ok=True)
            return False
    except Exception as ex:
        tqdm.write(f"   ❌ Download error ({label}): {ex}")
        tmp.unlink(missing_ok=True)
        return False
   
def build_pdf_filename(sector: str, state: str, entity: str, audit_year) -> str:
    """<SECTOR>_<STATE>_<ENTITY(34)>_<AUDIT_YEAR>.pdf  (underscore-separated)."""
    parts = [
        str(sector).upper().strip() or "UNK",
        str(state).upper().strip()  or "UNK",
        _sanitize_filename(entity, 34),
        str(int(audit_year)) if str(audit_year).isdigit() else str(audit_year),
    ]
    return "_".join(parts) + ".pdf"

# ---------- 3.E  WORKBOOK COLUMN LOCATOR ----------
# Expected ENTITY_LIST headers (exact match on stripped, upper form).
# TWEAK ZONE #7: if your workbook uses slightly different labels, alias them here.
ENTITY_LIST_HEADER_ALIASES = {
    "STATE":                "STATE",
    "SECTOR":               "SECTOR",
    "ENTITY NAME":          "ENTITY NAME",
    "FY":                   "FY",
    "FYE":                  "FYE",
    "UEI":                  "UEI",
    "EIN":                  "EIN",
    "FRAMEWORK":            "FRAMEWORK",
    "DATE OF DOWNLOAD":     "DATE OF DOWNLOAD",
    "DOWNLOAD STATUS":      "DOWNLOAD STATUS",
    "VALIDATION STATUS":    "VALIDATION STATUS",
    "VALIDATED BY":         "VALIDATED BY",
    "VALIDATED ON":         "VALIDATED ON",
    "PDF LINK":             "PDF LINK",
    "AUDITOR'S NAME":       "AUDITOR'S NAME",
    "AUDITORS NAME":        "AUDITOR'S NAME",   # alias
    "AUDITOR NAME":         "AUDITOR'S NAME",   # alias
}

def locate_headers(ws, header_row: int = 1) -> Dict[str, int]:
    """Return {canonical_header: column_index(1-based)} for the ENTITY_LIST sheet."""
    mapping: Dict[str, int] = {}
    max_col = ws.max_column or 30
    for col in range(1, max_col + 1):
        v = ws.cell(row=header_row, column=col).value
        if v is None:
            continue
        key = str(v).strip().upper().replace("’", "'")
        canonical = ENTITY_LIST_HEADER_ALIASES.get(key)
        if canonical and canonical not in mapping:
            mapping[canonical] = col
    return mapping

# ---------- 3.F  ROW ITERATION HONORING DOWNLOAD CRITERIA ----------
def iter_downloadable_rows(ws, hdr: Dict[str, int]):
    """
    Yield (row_idx, row_data_dict) for rows where:
      - ENTITY NAME is non-blank
      - DOWNLOAD STATUS is blank OR 'NOT FOUND'
    """
    col_name = hdr.get("ENTITY NAME")
    col_dl   = hdr.get("DOWNLOAD STATUS")
    if not col_name:
        print("❌ ENTITY_LIST is missing the 'ENTITY NAME' column."); return
    for r in range(2, ws.max_row + 1):
        entity = ws.cell(row=r, column=col_name).value
        if not entity or not str(entity).strip():
            continue
        dl_status = (ws.cell(row=r, column=col_dl).value if col_dl else None)
        dl_str = (str(dl_status).strip().upper() if dl_status else "")
        if dl_str and dl_str != "NOT FOUND":
            continue    # skip already downloaded / other states
        row_data = {
            "STATE":       (ws.cell(row=r, column=hdr["STATE"]).value       if "STATE"       in hdr else None),
            "SECTOR":      (ws.cell(row=r, column=hdr["SECTOR"]).value      if "SECTOR"      in hdr else None),
            "ENTITY NAME": entity,
            "FY":          (ws.cell(row=r, column=hdr["FY"]).value          if "FY"          in hdr else None),
            "FYE":         (ws.cell(row=r, column=hdr["FYE"]).value         if "FYE"         in hdr else None),
            "UEI":         (ws.cell(row=r, column=hdr["UEI"]).value         if "UEI"         in hdr else None),
            "EIN":         (ws.cell(row=r, column=hdr["EIN"]).value         if "EIN"         in hdr else None),
        }
        yield r, row_data

# ---------- 3.G  CELL WRITE HELPERS ----------
def _write_cell(
    ws, row: int, col: int, value, fill=None, font=None, 
    number_format: Optional[str] = None, force_fill: bool = False, 
    horizontal: Optional[str] = None):
    """
    Write a value to a cell.
       • Font family → Aptos Narrow always
       • Font size   → 9 always (WORKBOOK_FONT_SIZE)
       • Fill        → IGNORED by default (per user spec: no background colors)
                       EXCEPT when force_fill=True (used for FYE auto-advancement)
       • Wrap        → forced OFF
       • Horizontal alignment → preserves existing UNLESS 'horizontal' arg provided
       • Row height / column width → untouched
    """

    c = ws.cell(row=row, column=col)
    c.value = value

    # ---- Force Aptos Narrow @ size 9 (retain caller's color/bold/italic if present) ----
    if font is None:
        try:
            font = FONT_DEFAULT
        except NameError:
            font = Font(name=WORKBOOK_FONT, size=WORKBOOK_FONT_SIZE,
                        color="000000", bold=False)
    else:
        try:
            font = Font(
                name=WORKBOOK_FONT,
                size=WORKBOOK_FONT_SIZE,
                color=font.color, bold=font.bold, italic=font.italic,
                underline=font.underline,
            )
        except Exception:
            font = Font(name=WORKBOOK_FONT, size=WORKBOOK_FONT_SIZE,
                        color="000000", bold=False)
    c.font = font

    # ---- Fill: only applied when explicitly forced ----
    # This is the ONLY sanctioned exception to the "no fills" rule.
    # Used exclusively for the FYE/FY auto-advancement case in ENTITY_LIST.
    if fill is not None and force_fill:
        c.fill = fill

    # ---- Force wrap_text OFF; preserve any existing horizontal/vertical alignment ----
    existing_align = c.alignment
    effective_horizontal = horizontal if horizontal else (existing_align.horizontal if existing_align else None)

    c.alignment = Alignment(
        horizontal=effective_horizontal,
        vertical=existing_align.vertical if existing_align else None,
        wrap_text=False,
        shrink_to_fit=existing_align.shrink_to_fit if existing_align else None,
        indent=existing_align.indent if existing_align else 0
    )

    if number_format:
        c.number_format = number_format
    return c

def _fmt_ddmmmyyyy(dt: date) -> str:
    """Format date as 'dd-mmm-yyyy' (e.g. 06-Jul-2026)."""
    return dt.strftime("%d-%b-%Y")

# ---------- 3.H  SINGLE-ROW SOURCING WORKFLOW ----------
def process_download_row(paths: Dict[str, Path],
                         ws, row_idx: int, row: Dict[str, Any],
                         hdr: Dict[str, int],
                         sorted_df: pd.DataFrame,          # ← 6th param
                         faac_index,                        # ← 7th param
                         latest_year: int) -> Dict[str, Any]:
    """Handles Steps 3-5 of the SOURCING MODULE for one workbook row."""
    state      = str(row.get("STATE",  "") or "").strip().upper()
    sector     = str(row.get("SECTOR", "") or "").strip().upper()
    entity     = str(row.get("ENTITY NAME", "") or "").strip()
    workbook_fy = row.get("FY")
    uei        = str(row.get("UEI", "") or "").strip().upper()

    try:
        min_year = int(workbook_fy) if workbook_fy is not None else latest_year
    except Exception:
        min_year = latest_year

    # ---- Find FAAC record walking backwards from latest_year to min_year ----
    hit = find_faac_record(sorted_df, faac_index, uei, state, latest_year, min_year)
    #                       ↑↑↑↑↑↑↑↑↑ MUST include sorted_df here too

    # ---- Fallback: try without state constraint ----
    if hit is None and uei:
        alt_keys = [k for k in faac_index if k[0] == uei]
        if alt_keys:
            for k in alt_keys:
                cand = find_faac_record(sorted_df, faac_index, uei, k[1],
                                        latest_year, min_year)
                #                       ↑↑↑↑↑↑↑↑↑ here too
                if cand is not None:
                    hit = cand; break

    if hit is None:
        # ---- NOT FOUND path ----
        if "DOWNLOAD STATUS" in hdr:
            _write_cell(ws, row_idx, hdr["DOWNLOAD STATUS"], "NOT FOUND", font=FONT_RED)
        if "DATE OF DOWNLOAD" in hdr:
            _write_cell(ws, row_idx, hdr["DATE OF DOWNLOAD"], "")
        tqdm.write(f"   ❌ NOT FOUND  |  {sector} {state}  {entity}  (UEI={uei})")
        return {"row": row_idx, "status": "NOT FOUND", "entity": entity}

    # ---- Build metadata from CSV hit ----
    audit_year = int(str(hit.get("audit_year", "")).strip() or 0) or None
    report_id  = str(hit.get("report_id", "")).strip()
    firm       = str(hit.get("auditor_firm_name", "") or "").strip()
    contact    = str(hit.get("auditor_contact_name", "") or "").strip()
    auditor_name = f"{firm} ({contact})" if (firm and contact) else (firm or contact or "")

    pdf_url = build_pdf_url(report_id)
    if not pdf_url:
        if "DOWNLOAD STATUS" in hdr:
            _write_cell(ws, row_idx, hdr["DOWNLOAD STATUS"], "NOT FOUND", font=FONT_RED)
        tqdm.write(f"   ❌ No report_id available  |  {sector} {state}  {entity}")
        return {"row": row_idx, "status": "NOT FOUND", "entity": entity}

    # ---- Download PDF ----
    fname = build_pdf_filename(sector, state, entity, audit_year or min_year)
    dest  = paths["DOWNLOADED"] / fname
    label = f"{sector} {state} {_sanitize_filename(entity, 20)}"
    ok = download_pdf(pdf_url, dest, label=label)

    if not ok:
        if "DOWNLOAD STATUS" in hdr:
            _write_cell(ws, row_idx, hdr["DOWNLOAD STATUS"], "NOT FOUND", font=FONT_RED)
        return {"row": row_idx, "status": "NOT FOUND", "entity": entity}

    # ---- Workbook writeback: FY (highlight if newer than workbook FY) ----
    highlight_fy = False
    try:
        if workbook_fy is not None and audit_year and int(audit_year) > int(workbook_fy):
            highlight_fy = True
    except Exception:
        pass
    if audit_year and "FY" in hdr:
        fy_font = FONT_YELLOW if highlight_fy else None
        _write_cell(ws, row_idx, hdr["FY"], int(audit_year), font=fy_font)

    # ---- Write remaining columns ----
    if "DATE OF DOWNLOAD" in hdr:
        _write_cell(ws, row_idx, hdr["DATE OF DOWNLOAD"], _fmt_ddmmmyyyy(date.today()))
    if "DOWNLOAD STATUS" in hdr:
        _write_cell(ws, row_idx, hdr["DOWNLOAD STATUS"], "DOWNLOADED", font=FONT_GREEN)
    if "PDF LINK" in hdr:
        _write_cell(ws, row_idx, hdr["PDF LINK"], pdf_url)
    if "AUDITOR'S NAME" in hdr and auditor_name:
        _write_cell(ws, row_idx, hdr["AUDITOR'S NAME"], auditor_name)

    tqdm.write(f"   ✅ DOWNLOADED  |  {sector} {state}  {entity}"
               f"  →  AY={audit_year}  file={fname}"
               + ("  [FY-YELLOW]" if highlight_fy else ""))
    return {
        "row":        row_idx,
        "status":     "DOWNLOADED",
        "entity":     entity,
        "state":      state,
        "sector":     sector,
        "audit_year": audit_year,
        "pdf_path":   str(dest),
        "pdf_url":    pdf_url,
        "auditor":    auditor_name,
        "highlight_fy": highlight_fy,
    }

# ---------- 3.I  MODULE-LEVEL DRIVER ----------
def run_sourcing_module(paths: Dict[str, Path],
                        wb_path: Path) -> Dict[str, Any]:
    """End-to-end sourcing pipeline (dual progress bar: batch + per-entity)."""
    print("\n" + "═" * 78)
    print("SOURCING MODULE")
    print("═" * 78)
    ensure_workbook_closed(wb_path, save_before_close=True)

    # ── local helper: clamp labels so bars stay single-line ──
    def _short(txt: str, width: int = 44) -> str:
        txt = str(txt).replace("\n", " ").strip()
        return txt if len(txt) <= width else txt[:width - 1] + "…"

    # Step 1 – Ensure FAAC CSV is fresh
    csv_path   = ensure_faac_csv(paths["FAC_CSV"])
    faac_df    = load_faac_csv(csv_path)
    sorted_df, faac_index = index_faac_by_uei_state(faac_df)   # ← unpack tuple
    latest_yr  = discover_latest_audit_year(sorted_df)

    # Step 2 – Open workbook
    wb = load_workbook(wb_path, data_only=False, keep_links=False)
    if "ENTITY_LIST" not in wb.sheetnames:
        raise ValueError("ENTITY_LIST worksheet not found in workbook.")
    ws = wb["ENTITY_LIST"]
    hdr = locate_headers(ws, header_row=1)
    if not hdr.get("ENTITY NAME") or not hdr.get("UEI") or not hdr.get("STATE"):
        raise ValueError("ENTITY_LIST is missing ENTITY NAME / UEI / STATE columns.")

    print(f"\n🔎 Scanning ENTITY_LIST rows ({ws.max_row - 1} data rows)...")
    downloads: List[Dict[str, Any]] = []
    todo = list(iter_downloadable_rows(ws, hdr))
    print(f"   {len(todo)} rows are eligible for download.\n")

    src_timings: List[float] = []

    # ── OUTER bar: batch (one tick per entity) — stays on line 0 ──
    batch = tqdm(total=len(todo), desc="Sourcing", unit="entity",
                 position=0, leave=True, dynamic_ncols=True,
                 bar_format="{desc}: {percentage:3.0f}%|{bar:20}| "
                            "{n_fmt}/{total_fmt} [{elapsed}<{remaining}, "
                            "{rate_fmt}]{postfix}")

    # ── INNER bar: current entity's phases — stays on line 1, self-clears ──
    #    3 phases: Lookup → Download → Workbook
    entity_bar = tqdm(total=3, position=1, leave=False, dynamic_ncols=True,
                      bar_format="   ↳ {desc} {percentage:3.0f}%|{bar:15}| "
                                 "{n_fmt}/{total_fmt}")

    for row_idx, row in todo:
        t_row_start = time.perf_counter()
        entity_preview = str(row.get("ENTITY NAME", ""))[:30]
        state_code  = str(row.get("STATE", "") or "").upper()
        sector_code = str(row.get("SECTOR", "") or "").upper()

        # outer bar shows WHICH entity is in flight
        batch.set_postfix_str(_short(f"{sector_code} {state_code} {entity_preview}"),
                              refresh=True)

        # reset inner bar for this entity
        entity_bar.reset(total=3)
        entity_bar.set_description(_short(f"Lookup   {entity_preview}"))
        entity_bar.refresh()

        try:
            # Phase 1 → Lookup / match handled inside process_download_row.
            # We advance the inner bar around the atomic call so the user sees
            # a live second line per entity (Lookup → Download → Workbook).
            entity_bar.update(1)   # Lookup done → entering Download

            entity_bar.set_description(_short(f"Download {entity_preview}"))
            entity_bar.refresh()

            result = process_download_row(
                paths, ws, row_idx, row, hdr,
                sorted_df,        # ← MUST include this
                faac_index,
                latest_yr
            )
            entity_bar.update(1)   # Download done → entering Workbook

            entity_bar.set_description(_short(f"Workbook {entity_preview}"))
            entity_bar.refresh()

            elapsed = time.perf_counter() - t_row_start
            src_timings.append(elapsed)
            result["sourcing_seconds"] = round(elapsed, 2)
            downloads.append(result)
            entity_bar.update(1)   # Workbook done → entity complete

            if len(downloads) % 25 == 0:
                wb.save(wb_path)

        except Exception as ex:
            elapsed = time.perf_counter() - t_row_start
            tqdm.write(f"   ❌ Row {row_idx} error after {elapsed:.2f}s: {ex}")
            downloads.append({"row": row_idx, "status": "ERROR",
                              "error": str(ex),
                              "sourcing_seconds": round(elapsed, 2)})
        finally:
            batch.update(1)        # advance OUTER bar once per entity

    entity_bar.close()
    batch.close()

    wb.save(wb_path)
    print(f"\n✅ Sourcing complete. Workbook saved: {wb_path}")

    dl   = sum(1 for d in downloads if d.get("status") == "DOWNLOADED")
    nf   = sum(1 for d in downloads if d.get("status") == "NOT FOUND")
    err  = sum(1 for d in downloads if d.get("status") == "ERROR")
    print(f"   Downloaded: {dl}   Not Found: {nf}   Errors: {err}")
    return {"downloads": downloads,
            "sourcing_timings": src_timings,
            "sourcing_wall_seconds": sum(src_timings)}

# ============================================================================
# PART 4: VALIDATION GATE 1 — ISSUER DETAILS VALIDATION (LG + NONLG)
#   G1-1  ENTITY NAME CHECK
#   G1-2  STATE OF THE ENTITY CHECK
#   G1-3  FISCAL YEAR CHECK
# ============================================================================

# ---------- 4.A  GATE 1 CONFIG (TWEAK ZONE #8) ----------
G1_FIRST_N_PAGES  = 30      # spec: "first 30 pages"
G1_REVIEW_MIN_PCT = 90.0    # >= 90 and < 100 → REVIEW   (G1-1 entity-name scale)
G1_PASS_PCT       = 100.0   # ==100 → PASS               (G1-1 entity-name scale)
FYE_TOP_LINES     = 5       # spec: "first five lines from top" of a page
# NOTE: G1-2 (state) uses its OWN distance-based tier scale (100/95/90/…),
#       NOT the two constants above — do not conflate them.

# ---------- 4.B  G1-1  ENTITY NAME CHECK ----------
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

# ---------- 4.C  G1-2  STATE OF THE ENTITY CHECK ----------
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

# ---------- 4.D  G1-3  FISCAL YEAR CHECK ----------
# ✅ FIX 2: the unused `_is_cover_page()` helper has been removed (dead code;
#          not referenced anywhere — the multi-page cover scan below covers it).

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
                    if best_ever_match is None or cand[0] > best_ever_match:
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

# ---------- 4.E  GATE 1 AGGREGATION ----------
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

# ============================================================================
# TIMED GATE WRAPPERS — Capture per-step execution time for summary display
# ✅ Crash-hardened: every step runs under _safe() so a single validator
#      exception becomes a FAIL StepResult (with traceback) instead of an
#      unhandled "ERROR" that blanks the whole gate. Combined with the
#      fail-loud overall_validation_status() (Part 7.E), a crash can never
#      masquerade as a benign REVIEW.
# ============================================================================
import traceback as _traceback   # module-level; used by the _safe helpers below

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

# ---------- 4.F  GATE 1 DRIVER ----------
def run_gate1(doc: fitz.Document,
              workbook_entity: str,
              workbook_state: str,
              workbook_fye,
              sector: str, audit_year,
              screenshot_out: Path,
              financial_stmt_pages: Optional[List[int]] = None
              ) -> Tuple[str, Dict[str, StepResult]]:
    """Run G1-1, G1-2, G1-3 and return (gate_status, {step_code: StepResult})."""
    print("  ↳ G1  ISSUER DETAILS VALIDATION")

    g11 = validate_g1_1_entity_name(
        doc, workbook_entity, sector, workbook_state, audit_year, screenshot_out
    )
    print(f"     G1-1 ENTITY NAME       : {g11.status:<7} {g11.score:>6.2f}%  "
          f"p{('-' if g11.page_idx is None else g11.page_idx+1)}")

    g12 = validate_g1_2_state(
        doc, workbook_state, g11, sector, audit_year, screenshot_out
    )
    print(f"     G1-2 STATE             : {g12.status:<7} {g12.score:>6.2f}%  "
          f"p{('-' if g12.page_idx is None else g12.page_idx+1)}")

    g13 = validate_g1_3_fye(
        doc, workbook_fye, sector, workbook_state, audit_year,
        screenshot_out, financial_stmt_pages
    )
    print(f"     G1-3 FISCAL YEAR       : {g13.status:<7} {g13.score:>6.2f}%  "
          f"p{('-' if g13.page_idx is None else g13.page_idx+1)}")

    gate_status = aggregate_gate([g11, g12, g13])
    for _sr in results.values():
        relabel_step_if_page_ocr(_sr, doc)
        relabel_notfound_stmt_if_ocr(sr, doc)
    print(f"     → GATE 1 : {gate_status}")
    return gate_status, {"G1-1": g11, "G1-2": g12, "G1-3": g13}

# ============================================================================
# PART 5: VALIDATION GATE 2 — AUDIT COMPLIANCE VALIDATION (LG + NONLG)
#   G2-1  AUDITOR'S OPINION
#   G2-2  FYE MENTIONED IN AUDITOR'S REPORT
#   G2-3  AUDITOR'S NAME RECONCILIATION
#   G2-4  AUDITOR'S SIGNATURE  (HITL — never PASS; excluded from finalization)
# ============================================================================

# ---------- 5.A  GATE 2 CONFIG (TWEAK ZONE #9) ----------
# ✅ FIX C: comment now matches the actual value (was "first five lines").
G2_HEADING_TOP_LINES = 15     # top-N lines examined for the report heading
G2_OPINION_WINDOW    = 40    # lines after heading considered "OPINION paragraph area"
G2_NAME_FUZZY_MIN    = 85    # RapidFuzz partial_ratio threshold for firm/contact match
G2_SIG_SECTION_MAX_PAGES = 6 # spec: signature may span multiple pages incl. blanks
G2_SIG_MIN_IMG_AREA  = 5000  # px² — filter out tiny logos when searching for signature


# OCR-tolerant 'auditor' fragment: absorbs 'AADITOR' (U→A), 'AUDIT0R' (O→0)
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

# OPINION section heading
OPINION_HEADING_RE = re.compile(r"^\s*opinion(?:s)?\b[:\s\-–—]*$", re.IGNORECASE)

# Matches when "Opinion" is a bold sub-heading immediately followed by paragraph
OPINION_SUB_HEADING_RE = re.compile(r"^\s*opinion(?:s)?\s*$", re.IGNORECASE)

# "In our opinion" paragraph starter (used when no explicit heading)
OPINION_PARAGRAPH_RE = re.compile(r"\bin\s+our\s+opinion\b", re.IGNORECASE)

# Headings that terminate the OPINION section.
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

# "Other Reporting Required by Government Auditing Standards" (anchor for signature)
OTHER_REPORTING_RE = re.compile(
    r"other\s+report(?:ing|s)?\s+required\s+by\s+(?:the\s+)?"
    r"government\s+auditing\s+standards",
    re.IGNORECASE
)

# Opinion keyword tiers
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

# Sub-opinion headings inside a multi-opinion "Opinions" section
_OPINION_SUBHEAD_RE = re.compile(
    r"^\s*(?:qualified|unmodified|unqualified|adverse|disclaimer)\s+"
    r"opinion(?:s)?\b", re.IGNORECASE)

# The operative basis for LG/NJ regulatory & OCBOA reports
_REGULATORY_BASIS_RE = re.compile(
    r"regulatory\s+basis(?:\s+of\s+accounting)?"
    r"|prescribed\s+or\s+permitted\s+by\s+the\s+division"
    r"|modified\s+cash\s+basis"
    r"|statutory\s+basis\s+of\s+accounting"
    r"|other\s+comprehensive\s+basis\s+of\s+accounting", re.IGNORECASE)

# The sub-opinion to DISCARD for OCBOA/regulatory entities
_GAAP_OPINION_RE = re.compile(
    r"u\.?s\.?\s+generally\s+accepted\s+accounting\s+principles"
    r"|accounting\s+principles\s+generally\s+accepted", re.IGNORECASE)

# ---------- 5.B  AUDITOR-NAME NORMALIZATION ----------
# TWEAK ZONE #10: extend suffix list as needed
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

# ---------- 5.C  LOCATE INDEPENDENT / STATE AUDITOR'S REPORT PAGES ----------
# ── Compliance / Single-Audit markers — pages carrying these are NOT the
#    financial-statement audit report and must be EXCLUDED from G2. ──
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

## ── Non-opinion report titles to EXCLUDE (heading-region only) ──
## The FINANCIAL-STATEMENT opinion is the ONLY report we keep for G2.
## Reject any heading that announces Internal Control, a Federal/State
## PROGRAM, Federal AWARDS, or Compliance/Other-Matters reporting.
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

# ── Positive anchor — the financial-statement audit report always carries
#    one of these nearby. ──
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

# ---------- 5.D  G2-1  AUDITOR'S OPINION ----------
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

# ---------- 5.E  G2-2  FYE MENTIONED IN AUDITOR'S REPORT ----------
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

# ---------- 5.F  G2-3  AUDITOR'S NAME RECONCILIATION ----------
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

# ---------- 5.G  G2-4  AUDITOR'S SIGNATURE ----------
# ✅ FIX B: HITL step — outcome is ALWAYS either REVIEW (signature found, image
#          OR text) or FAIL (nothing found). There is NO PASS branch anywhere in
#          this validator, and it is EXCLUDED from Gate-2 finalization (see 5.H).
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


# ─── Expanded signature font keyword list (TWEAK ZONE #10A) ───
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

# Headings that mark the END of the auditor's report section.
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

# ---------- 5.H  GATE 2 DRIVER ----------
def run_gate2(doc: fitz.Document,
              workbook_fye,
              auditor_name_raw: str,
              sector: str, state: str, audit_year,
              screenshot_out: Path) -> Tuple[str, Dict[str, StepResult]]:
    """Run G2-1 → G2-4 and return (gate_status, {step_code: StepResult}).

       ✅ FIX A — G2-4 (Auditor's Signature) is a HUMAN-IN-THE-LOOP step:
          • It can only ever be REVIEW (found) or FAIL (not found), never PASS.
          • Per user spec it is NOT counted in VALIDATION STATUS finalization,
            so Gate-2 status is aggregated over G2-1, G2-2, G2-3 ONLY.
          • G2-4 is still executed and returned so the VALIDATION worksheet
            shows its REVIEW/FAIL verdict for the human reviewer, but a FAILED
            signature can no longer sink Gate 2 (and a REVIEW no longer needs
            the old REVIEW→PASS promotion hack, which has been removed)."""
    print("  ↳ G2  AUDIT COMPLIANCE VALIDATION")

    g21, report_pages, opinion_lines = validate_g2_1_opinion(
        doc, sector, state, audit_year, screenshot_out
    )
    print(f"     G2-1 AUDITOR OPINION   : {g21.status:<7} {g21.score:>6.2f}%  "
          f"p{('-' if g21.page_idx is None else g21.page_idx+1)}")

    g22 = validate_g2_2_fye_in_opinion(
        doc, workbook_fye, opinion_lines,
        sector, state, audit_year, screenshot_out
    )
    print(f"     G2-2 FYE IN OPINION    : {g22.status:<7} {g22.score:>6.2f}%  "
          f"p{('-' if g22.page_idx is None else g22.page_idx+1)}")

    g23 = validate_g2_3_auditor_name(
        doc, auditor_name_raw, sector, state, audit_year, screenshot_out
    )
    print(f"     G2-3 AUDITOR NAME      : {g23.status:<7} {g23.score:>6.2f}%  "
          f"p{('-' if g23.page_idx is None else g23.page_idx+1)}")

    g24 = validate_g2_4_signature(
        doc, report_pages, sector, state, audit_year, screenshot_out
    )
    print(f"     G2-4 AUDITOR SIGNATURE : {g24.status:<7} {g24.score:>6.2f}%  "
          f"p{('-' if g24.page_idx is None else g24.page_idx+1)}  "
          f"[HITL — excluded from gate finalization]")

    # ---- GATE-2 FINALIZATION (G2-4 EXCLUDED) ----
    # Only G2-1, G2-2, G2-3 drive the gate status. G2-4 is reported but never
    # counted, so its REVIEW/FAIL cannot alter the Gate-2 (or overall) verdict.
    gate_status = aggregate_gate([g21, g22, g23])
    for _sr in results.values():
        relabel_step_if_page_ocr(_sr, doc)
        relabel_notfound_stmt_if_ocr(sr, doc)
    print(f"     → GATE 2 : {gate_status}   (signature is HITL-only)")

    # Return ALL four results (G2-4 with its true REVIEW/FAIL intact) so the
    # VALIDATION worksheet still shows the signature verdict for the reviewer.
    return gate_status, {"G2-1": g21, "G2-2": g22, "G2-3": g23, "G2-4": g24}

# ============================================================================
# PART 5.5: FRAMEWORK & BASIS-OF-ACCOUNTING DETECTION   (unchanged logic)
# ============================================================================

# ---------- 5.5.A  DETECTION CONFIG (TWEAK ZONE #14) ----------
FRAMEWORK_MAX_PAGES  = 150      # scan first N pages (audits typically declare framework early)
FRAMEWORK_SECONDARY_MAX_PAGES = 250  # for extremely long audits (>200 pages)
FRAMEWORK_MIN_SCORE  = 2       # minimum weighted score to classify
FRAMEWORK_HIGH_CONF  = 8       # score ≥ this = high confidence (90-100%)
FRAMEWORK_MED_CONF   = 4       # score ≥ this = medium confidence (70-89%)

# ---------- 5.5.B  FRAMEWORK PATTERNS ----------
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

# ---------- 5.5.C  BASIS-OF-ACCOUNTING PATTERNS ----------
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

# ---------- 5.5.D  DETECTION FUNCTIONS ----------
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
# ============================================================================
# PART 6: VALIDATION GATE 3 — FINANCIAL STATEMENT STRUCTURE VALIDATION
#   G3-1  STATEMENT OF CASH RECEIPTS AND DISBURSEMENTS   [LG + NONLG]
#   G3-2  STATEMENT OF NET POSITION / BALANCE SHEET      [LG + NONLG]
#   G3-3  BALANCE SHEET OF GOVERNMENTAL FUNDS            [LG only]
#   G3-4  STATEMENT OF ACTIVITIES / REV,EXP,ΔNP          [LG + NONLG]
#   G3-5  STATEMENT OF REV, EXP & CHANGE IN FUND BAL.    [LG only]
#   G3-6  STATEMENT OF CASH FLOWS                        [NONLG only]
# ============================================================================

# ---------- 6.A  GATE 3 CONFIG (TWEAK ZONE #11) ----------
G3_HEADING_TOP_LINES  = 8     # top-N lines examined for page heading
G3_TITLE_SPAN_MAX     = 3     # statement name may continue up to 3 consecutive lines
G3_ROW_SEARCH_WINDOW  = 60    # lines below title inspected for row/column labels
# ✅ FIX 4: SINGLE SOURCE OF TRUTH for "how many pages one statement may span".
#          Previously three different caps (6 / 4 / 4) silently truncated
#          statements that legitimately span 5-6 pages. All continuation logic
#          now references G3_STMT_MAX_CONT_PAGES.
G3_STMT_MAX_CONT_PAGES = 6    # a single statement may span up to N pages
G3_CONT_MAX_PAGES      = G3_STMT_MAX_CONT_PAGES   # alias (kept for API compat)

# ---------- EXCLUDED / ALLOWED TERM SETS (spec-exact) ----------
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

# ---------- Helper regex fragments (used across STATEMENT_SPECS) ----------
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

# ── Fund qualifiers that must NOT be matched by the generic bare "Balance Sheet"
#    (NJ funds + governmental funds are routed to their OWN buckets explicitly). ──
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

# ═══ SANITY CHECK: verify STATEMENT_SPECS loaded correctly ═══
if not STATEMENT_SPECS or len(STATEMENT_SPECS) < 6:
    raise RuntimeError(
        f"❌ STATEMENT_SPECS has {len(STATEMENT_SPECS)} entries (expected 6)! "
        f"Check for copy-paste corruption in the STATEMENT_SPECS block above."
    )
_expected_keys = {"G3-1", "G3-2", "G3-3", "G3-4", "G3-5", "G3-6"}
_missing = _expected_keys - set(STATEMENT_SPECS.keys())
if _missing:
    raise RuntimeError(f"❌ STATEMENT_SPECS missing keys: {_missing}")
print(f"✅ STATEMENT_SPECS loaded with {len(STATEMENT_SPECS)} step specs")

# ---------- 6.B  PAGE-HEADING FILTER ----------
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

# ---------- 6.C  STATEMENT TITLE MATCHER (multi-line span) ----------
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

# ── OCR-tolerant fuzzy title fallback (rapidfuzz) ──────────────────────────
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

# ---------- 6.C-DEBUG  ONE-TIME PAGE DUMP for troubleshooting ----------
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

# ---------- 6.D  ROW / COLUMN STRUCTURE CHECKERS ----------
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

# ---------- 6.E  FYE-ON-STATEMENT DETECTOR ----------
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
# ✅ FIX 2: the mis-indented, unreachable duplicate of
#     `_page_excluded_by_top_lines` that used to sit here (after the
#     `return best` above) has been DELETED. The single correct
#     definition lives at module level, just below the shared accumulator.

# ─── SHARED: continuation-aware accumulator (title-fn driven) ─────────────
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

# ═══════════════════════════════════════════════════════════════════════════
# SHARED HELPER — Page-level EXCLUDED / ALLOWED term filter
# (SINGLE authoritative definition — used by G3-2/G3-4/G3-6 dedicated validators)
# ═══════════════════════════════════════════════════════════════════════════
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

# ═══════════════════════════════════════════════════════════════════════════
# G3-2 DEDICATED VALIDATOR (NONLG) — Statement of Net Position / Balance Sheet
# ═══════════════════════════════════════════════════════════════════════════
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

# ═══════════════════════════════════════════════════════════════════════════
# G3-4 DEDICATED VALIDATOR (NONLG) — Statement of Activities / Operations
# ═══════════════════════════════════════════════════════════════════════════
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

# ═══════════════════════════════════════════════════════════════════════════
# G3-6 DEDICATED VALIDATOR — Statement of Cash Flows (NONLG)
# ═══════════════════════════════════════════════════════════════════════════
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

# ── NJ regulatory FUND-CONTEXT detection (the real disambiguator) ──
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
            page_top_lines(doc, pi, top_n)
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

# ---------- 6.F  SINGLE-STATEMENT VALIDATOR ----------
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

# ---------- 6.G  GATE 3 AGGREGATION (sector-specific) ----------
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

# ---------- 6.H  GATE 3 DRIVER ----------
def run_gate3(doc, sector, state, audit_year, workbook_fye, screenshot_out):
    """
    ✅ FIX 3: Sector-aware + crash-hardened.
       • NONLG → dedicated validators (G3-2 / G3-4 / G3-6); G3-3 / G3-5 = NOT APPLICABLE.
       • LG    → generic validate_single_statement for G3-1..G3-5, PLUS G3-6 so a
                 BTA-style LG (community college / water district / transit) can still
                 satisfy the NONLG-style combination via aggregate_gate3().
       • EVERY step is wrapped in _safe(): an unhandled exception NO LONGER escapes
         to the caller as "ERROR" and blanks all six columns. Instead the offending
         step becomes FAIL (with a full traceback printed) while the other five
         steps still run and report normally.
    """
    print("  ↳ G3  FINANCIAL STATEMENT STRUCTURE VALIDATION")

    import traceback   # local import keeps the module header untouched
    sector_u = sector.upper()

    # ─── Per-step crash guard ────────────────────────────────────────────
    def _safe(step_code: str, fn) -> Tuple[StepResult, List[int]]:
        try:
            out = fn()
            if isinstance(out, tuple):
                sr, pages = out
            else:
                sr, pages = out, []
            return sr, (pages or [])
        except Exception as ex:
            print(f"     ⚠️  {step_code} crashed: {type(ex).__name__}: {ex}")
            traceback.print_exc()
            sr = StepResult(code=step_code)
            sr.status = "FAIL"
            sr.score  = 0.0
            sr.notes  = f"Validator crashed ({type(ex).__name__}: {ex}); marked FAIL."
            try:
                sr.screenshot = take_diagnostic_screenshot(
                    doc, 0, sector, state, audit_year, step_code,
                    screenshot_out, status="FAIL")
            except Exception:
                pass
            return sr, []

    # ─── Candidate-page pre-filter (shared by the generic LG path) ────────
    heading_cache   = {p: page_heading_status(doc, p) for p in range(doc.page_count)}
    candidate_pages = [p for p, s in heading_cache.items() if s != "EXCLUDED"]

    def _generic(step_code: str):
        "Run the generic STATEMENT_SPECS validator with both sectors temporarily enabled."
        spec = STATEMENT_SPECS[step_code]
        orig_sectors = spec["sectors"]
        spec["sectors"] = {"LG", "NONLG"}     # prevent early NOT-APPLICABLE exit
        try:
            return validate_single_statement(
                doc, step_code, sector, state, audit_year, workbook_fye,
                screenshot_out, candidate_pages=candidate_pages)
        finally:
            spec["sectors"] = orig_sectors     # always restore

    results: Dict[str, StepResult] = {}
    all_stmt_pages: List[int] = []

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
            sr, pages = _safe(code, fn)
            results[code] = sr
            all_stmt_pages.extend(pages)

        for code in ("G3-3", "G3-5"):
            sr = StepResult(code=code)
            sr.status = "NOT APPLICABLE"
            sr.notes  = f"{code} not applicable to NONLG"
            results[code] = sr

    else:
        for code in ("G3-1", "G3-2", "G3-3", "G3-4", "G3-5"):
            sr, pages = _safe(code, (lambda c=code: _generic(c)))
            results[code] = sr
            all_stmt_pages.extend(pages)

        sr6, pages6 = _safe("G3-6", lambda: validate_g3_6_cash_flows(
            doc, "NONLG", state, audit_year, workbook_fye, screenshot_out))
        results["G3-6"] = sr6
        all_stmt_pages.extend(pages6)

    # ✅ OCR RELABELING — MUST run BEFORE aggregation & console print ───────
    #    Bug-fix 1: iterate with _sr (not the stale sr).
    #    Bug-fix 2: relabel first, THEN aggregate, so OCR surfaces at the gate.
    any_ocr = False
    for _sr in results.values():
        relabel_step_if_page_ocr(_sr, doc)                                       # found-page path
        relabel_notfound_stmt_if_ocr(_sr, doc, candidate_pages=candidate_pages)  # not-found / image-only path
        if _sr is not None and _sr.status == OCR_STATUS:
            any_ocr = True

    # ─── Per-step console line (now shows OCR REQUIRED correctly) ─────────
    for code in ("G3-1", "G3-2", "G3-3", "G3-4", "G3-5", "G3-6"):
        sr = results[code]
        pr = sr.page_range or ("-" if sr.page_idx is None else str(sr.page_idx + 1))
        print(f"     {code} {STATEMENT_SPECS[code]['label'][:38]:<38}: "
              f"{sr.status:<14} {sr.score:>6.2f}%  p{pr}")

    # ✅ Aggregate AFTER relabeling; surface OCR at gate level so
    #    overall_validation_status() (Part 7.E) can prioritize it.
    gate_status = aggregate_gate3(sector, results)
    if any_ocr:
        gate_status = OCR_STATUS

    print(f"     → GATE 3 : {gate_status}")
    return gate_status, results, sorted(set(all_stmt_pages))

# ============================================================================
# PART 7: MAIN ORCHESTRATOR + VALIDATION WORKSHEET WRITER + ENTRY POINT
# ============================================================================

# ---------- 7.A  VALIDATION WORKSHEET SCHEMA ----------
# TWEAK ZONE #13: reorder or rename columns here (KEEP header text stable — the
# sourcing module and workbook writer both key off exact text).

VALIDATION_HEADERS = [
    # Issuer identity (value-pasted from ENTITY_LIST)
    "STATE", "SECTOR", "ENTITY NAME", "FY", "FYE", "UEI", "EIN",
    "DATE OF DOWNLOAD", "DOWNLOAD STATUS", "VALIDATION STATUS",
    "FRAMEWORK", "FRAMEWORK CONFIDENCE",
    # Gate 1
    "VALIDATION GATE 1: ISSUER DETAILS VALIDATION",
    "ENTITY NAME CHECK", "ENTITY NAME CHECK %",
    "STATE OF THE ENTITY CHECK", "STATE OF THE ENTITY CHECK %",
    "FISCAL YEAR CHECK", "FISCAL YEAR CHECK %",
    # Gate 2
    "VALIDATION GATE 2: AUDIT COMPLIANCE VALIDATION",
    "AUDITOR'S OPINION", "AUDITOR'S OPINION %",
    "FYE MENTIONED IN AUDITOR'S REPORT",
    "AUDITOR'S NAME RECONCILIATION",
    "AUDITOR'S SIGNATURE",
    # Gate 3
    "VALIDATION GATE 3: FINANCIAL STATEMENT STRUCTURE VALIDATION",
    "STATEMENTS OF CASH RECEIPTS AND DISBURSEMENTS",
    "STATEMENTS OF NET POSITION",
    "BALANCE SHEET OF GOVERNMENTAL FUNDS",
    "STATEMENT OF ACTIVITIES",
    "STATEMENT OF REVENUE, EXPENDITURE AND CHANGE IN FUND BALANCES",
    "STATEMENT OF CASH FLOWS",
    # Per-entity timing — numeric Excel durations (format hh:mm:ss.0)
    "SOURCING TIME", "VALIDATION TIME", "OVERALL TIME",
]

# Map validation-step code → (status column header, percent column header or None)
STEP_TO_COLUMN = {
    "G1-1": ("ENTITY NAME CHECK",              "ENTITY NAME CHECK %"),
    "G1-2": ("STATE OF THE ENTITY CHECK",      "STATE OF THE ENTITY CHECK %"),
    "G1-3": ("FISCAL YEAR CHECK",              "FISCAL YEAR CHECK %"),
    "G2-1": ("AUDITOR'S OPINION",              "AUDITOR'S OPINION %"),
    "G2-2": ("FYE MENTIONED IN AUDITOR'S REPORT", None),
    "G2-3": ("AUDITOR'S NAME RECONCILIATION",  None),
    "G2-4": ("AUDITOR'S SIGNATURE",            None),
    "G3-1": ("STATEMENTS OF CASH RECEIPTS AND DISBURSEMENTS", None),
    "G3-2": ("STATEMENTS OF NET POSITION",     None),
    "G3-3": ("BALANCE SHEET OF GOVERNMENTAL FUNDS", None),
    "G3-4": ("STATEMENT OF ACTIVITIES",        None),
    "G3-5": ("STATEMENT OF REVENUE, EXPENDITURE AND CHANGE IN FUND BALANCES", None),
    "G3-6": ("STATEMENT OF CASH FLOWS",        None),
}

# ---------- 7.B  ENSURE VALIDATION WORKSHEET EXISTS WITH CORRECT HEADERS ----------
def ensure_validation_ws(wb) -> Any:
    """
    Locate the VALIDATION worksheet and map its existing headers to columns.
    Per user spec: DO NOT modify headers, row heights, column widths, or formatting.
    If the VALIDATION sheet is missing, create it with a bare header row —
    the user can format headers themselves afterwards.
    """
    if "VALIDATION" in wb.sheetnames:
        ws = wb["VALIDATION"]
    else:
        # Sheet missing — create minimally, no styling applied
        ws = wb.create_sheet("VALIDATION")
        for i, h in enumerate(VALIDATION_HEADERS, start=1):
            ws.cell(row=1, column=i, value=h)   # ← no font, no fill, no dims changed

    # Snapshot existing header text (case- & punctuation-insensitive)
    existing = {}
    for col in range(1, (ws.max_column or 0) + 1):
        v = ws.cell(row=1, column=col).value
        if v is None:
            continue
        key = str(v).strip().upper().replace("’", "'")
        existing[key] = col
    return ws, existing

def _validation_col(header_map: Dict[str, int], header_text: str):
    """
    Look up column index for a given header text.
    Case-insensitive and treats curly/straight apostrophes as equivalent.
    """
    key = header_text.upper().replace("’", "'").strip()
    return header_map.get(key)

# ---------- 7.C  LOCATE-OR-CREATE ROW FOR AN ENTITY IN VALIDATION SHEET ----------
def find_or_create_validation_row(ws, hdr_map: Dict[str, int],
                                  uei: str, entity: str, fy) -> int:
    """Match on (UEI, ENTITY NAME, FY) triple.  Append a new row if not present."""
    col_uei    = _validation_col(hdr_map, "UEI")
    col_entity = _validation_col(hdr_map, "ENTITY NAME")
    col_fy     = _validation_col(hdr_map, "FY")

    tgt_uei = str(uei or "").strip().upper()
    tgt_ent = str(entity or "").strip().upper()
    try:
        tgt_fy = int(fy) if fy is not None and str(fy).strip() else None
    except Exception:
        tgt_fy = None

    for r in range(2, ws.max_row + 1):
        u = str(ws.cell(row=r, column=col_uei).value or "").strip().upper()   if col_uei    else ""
        e = str(ws.cell(row=r, column=col_entity).value or "").strip().upper() if col_entity else ""
        f_raw = ws.cell(row=r, column=col_fy).value                          if col_fy     else None
        try:
            f = int(f_raw) if f_raw is not None and str(f_raw).strip() else None
        except Exception:
            f = None
        fy_match = (f == tgt_fy) or (f is None and tgt_fy is None)
        if u == tgt_uei and e == tgt_ent and fy_match:
            return r

    # Not found → append
    return (ws.max_row or 1) + 1

# ---------- 7.D  WRITE ONE STEP RESULT INTO VALIDATION WORKSHEET ----------
def _font_for_status(status: str) -> Font:
    if status == "PASS":      return FONT_GREEN
    if status == "FAIL":      return FONT_RED
    if status == "REVIEW":    return FONT_YELLOW
    if status == "NOT FOUND": return FONT_RED
    if status == OCR_STATUS:  return FONT_YELLOW    # needs OCR → amber/attention
    return FONT_YELLOW    # NOT APPLICABLE etc.

def _fill_for_status(status: str):
    """
    Kept for API compatibility; returns None so no fill is ever applied.
    Per user spec: no background cell colors anywhere in the workbook.
    """
    return None

def write_step_result(ws, row: int, hdr_map: Dict[str, int], step: StepResult):
    """
    Write a StepResult into its target status column (+ % column if applicable).
    Applies deep-colored Aptos Narrow font; NO background fill.
    Appends reconciliation_note (if present) to the display with ' - ' separator.
    """
    col_hdr, pct_hdr = STEP_TO_COLUMN.get(step.code, (None, None))
    if not col_hdr:
        return

    # Build display value: STATUS [classification] [page-range] per spec
    display = step.status

    # append classification label (e.g., PASS [Present Fairly])
    if step.classification:
        display = f"{display} [{step.classification}]"

    # Then append page-range / page-index bracket
    if step.page_range:
        display = f"{display} [{step.page_range}]"
    elif step.page_idx is not None:
        display = f"{display} [{step.page_idx + 1}]"

    # Append reconciliation note with dash separator (e.g., FYE advancement)
    if step.reconciliation_note:
        display = f"{display} - {step.reconciliation_note}"

    col_i = _validation_col(hdr_map, col_hdr)
    if col_i:
        _write_cell(ws, row, col_i, display,
                    font=_font_for_status(step.status))

    if pct_hdr:
        col_p = _validation_col(hdr_map, pct_hdr)
        if col_p:
            # Store as a real fraction (0.95) with percentage format → 95.00%
            _write_cell(ws, row, col_p, round(step.score / 100.0, 4),
                        font=_font_for_status(step.status),
                        number_format="0.00%")

# ---------- 7.E  OVERALL VALIDATION STATUS AGGREGATION ----------
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

def _print_issuer_summary(
    entity, state, sector, ay, g1_status, g1_results, g2_status, g2_results,
    g3_status, g3_results, overall, gate_times, step_times,
    sourcing_time, validation_time):
    """
    Boxed, column-aligned per-issuer summary.

    • EVERY field is CLAMPED (truncate + pad) to an exact width, so
      over-length labels / page ranges can never break the right wall.
    • Emitted as ONE atomic tqdm.write() → the live progress bar cannot
      slice into the middle of the box.
    """
    BOX_W = 79   # printable columns between the two border chars

    # ── robust display width (emoji render as 2 cols in modern terminals) ──
    def _dw(s: str) -> int:
        w = 0
        for ch in s:
            o = ord(ch)
            if o == 0xFE0F:                 # variation selector — zero width
                continue
            w += 2 if (o >= 0x1F000 or o in (0x274C, 0x2705, 0x26A0, 0x23ED)) else 1
        return w

    def _clamp(s: str, width: int, align: str = "<") -> str:
        "Force `s` to EXACTLY `width` display columns (truncate then pad)."
        s = str(s)
        while _dw(s) > width:               # truncate on display width
            s = s[:-1]
        pad = width - _dw(s)
        return (s + " " * pad) if align == "<" else (" " * pad + s)

    def _row(content: str = "") -> str:
        "Wrap content in borders, CLAMPED to BOX_W so the wall always aligns."
        return f"│{_clamp(content, BOX_W)}│"

    def _rule() -> str:
        return f"├{'─' * BOX_W}┤"

    def _icon(status: str) -> str:
        return ("✅" if status == "PASS"
                else "❌" if status in ("FAIL", "NOT FOUND")
                else "⚠️" if status == "REVIEW"
                else "🔍" if status == OCR_STATUS
                else "•")

    def _page_str(s) -> str:
        pr = getattr(s, "page_range", None)
        if pr:
            return f"p{pr}"
        if getattr(s, "page_idx", None) is not None:
            return f"p{s.page_idx + 1}"
        return "p—"

    NAMES = {
        "G1-1": "Entity Name",             "G1-2": "State",
        "G1-3": "Fiscal Year",
        "G2-1": "Auditor Opinion",         "G2-2": "FYE in Opinion",
        "G2-3": "Auditor Name",            "G2-4": "Auditor Signature",
        "G3-1": "Cash Receipts/Disburse",  "G3-2": "Net Position/Bal Sheet",
        "G3-3": "Balance Sheet Gov Funds", "G3-4": "Statement of Activities",
        "G3-5": "Rev/Exp/Change Fund Bal", "G3-6": "Cash Flows",
    }

    def _step_line(code: str, s) -> str:
        st  = (getattr(s, "status", "") or "").upper()
        pct = f"{getattr(s, 'score', 0.0):.2f}%"
        t   = f"{step_times.get(code, 0.0):.2f}s"
        # each field clamped to a fixed width → colons & wall always align
        body = ("  "
                + _clamp(code, 5)
                + _clamp(NAMES.get(code, code), 24)
                + ": "
                + _clamp(st, 14)
                + _clamp(pct, 8, ">")
                + "  " + _clamp(_page_str(s), 9)
                + "→ " + _clamp(t, 8, ">"))
        return _row(body)

    def _gate_header(title: str, secs: float) -> str:
        left, t = f"  {title}", f"{secs:.2f}s  "
        gap = BOX_W - _dw(left) - _dw(t)
        return _row(left + " " * max(1, gap) + t)

    def _gate_verdict(n: int, status: str) -> str:
        txt = f"Gate {n}: {_icon(status)} {status}          "
        gap = BOX_W - _dw(txt)
        return _row(" " * max(0, gap) + txt)

    def _gate_secs(gkey: str, codes) -> float:
        if gate_times and gkey in gate_times:
            return gate_times[gkey]
        return sum(step_times.get(c, 0.0) for c in codes)

    # ── assemble ──
    L = [f"┌{'─' * BOX_W}┐"]

    suffix  = f"  ({state}, {sector})  |  AY {ay}"
    max_ent = BOX_W - _dw(f"  📋 {suffix}")
    L.append(_row(f"  📋 {_clamp(entity, max(4, max_ent))}{suffix}"))
    L.append(_rule())

    g1c = ("G1-1", "G1-2", "G1-3")
    L.append(_gate_header("Gate 1 — Issuer Details", _gate_secs("G1", g1c)))
    L += [_step_line(c, g1_results[c]) for c in g1c if c in g1_results]
    L.append(_gate_verdict(1, g1_status))

    g2c = ("G2-1", "G2-2", "G2-3", "G2-4")
    L.append(_gate_header("Gate 2 — Audit Compliance", _gate_secs("G2", g2c)))
    L += [_step_line(c, g2_results[c]) for c in g2c if c in g2_results]
    L.append(_gate_verdict(2, g2_status))

    g3c = ("G3-1", "G3-2", "G3-3", "G3-4", "G3-5", "G3-6")
    L.append(_gate_header("Gate 3 — Financial Structure", _gate_secs("G3", g3c)))
    L += [_step_line(c, g3_results[c]) for c in g3c if c in g3_results]
    L.append(_gate_verdict(3, g3_status))

    L.append(_rule())
    total = sourcing_time + validation_time
    L.append(_row(f"  OVERALL: {_icon(overall)} {overall}   |   "
                  f"Sourcing: {sourcing_time:.2f}s   "
                  f"Validation: {validation_time:.2f}s   Total: {total:.2f}s"))
    L.append(f"└{'─' * BOX_W}┘")

    tqdm.write("\n".join(L))

def _move_pdf_safely(src: Path, dest: Path, overall_status: str,
                     doc: "Optional[fitz.Document]" = None) -> bool:
    """
    Move a PDF from src to dest, with Windows AV / OneDrive lock tolerance.

    ROOT-CAUSE NOTE (WinError 32):
      The "used by another process" lock is almost always held by THIS
      Python process via the still-open fitz.Document handle. We therefore
      (1) close `doc` if supplied, and (2) force gc.collect() to drop any
      lingering finalizers BEFORE attempting the move.
    """
    import gc  # local import keeps this a true drop-in (no top-of-file edit)

    # ---- 0) Release any PyMuPDF handle THIS process may still hold ----
    if doc is not None:
        try:
            doc.close()
        except Exception:
            pass
    gc.collect()  # force finalizers so orphaned fitz handles are dropped

    if not src.exists():
        tqdm.write(f"     ⚠️  Source PDF missing: {src.name}")
        return False

    if dest.exists():
        # Already moved in a previous run — remove source to avoid duplicate
        try:
            src.unlink()
            tqdm.write(f"     ⏭  {overall_status} — destination already exists; "
                       f"removed duplicate from DOWNLOADED")
            return True
        except Exception as ex:
            tqdm.write(f"     ⚠️  Duplicate found but couldn't remove source: {ex}")
            return False

    dest.parent.mkdir(parents=True, exist_ok=True)

    # ---- 1) Primary move with progressive back-off (5 attempts) ----
    last_err = None
    for attempt in range(1, 6):
        try:
            shutil.move(str(src), str(dest))
            tqdm.write(f"     📁 {overall_status} — moved to "
                       f"{dest.parent.name}/{dest.name}")
            return True
        except (OSError, PermissionError) as ex:
            last_err = ex
            gc.collect()               # retry a handle release each round
            time.sleep(0.7 * attempt)  # 0.7s → 3.5s progressive back-off

    # ---- 2) Fallback: copy + delete (works when move is blocked but
    #         the file is still READABLE — the typical OneDrive case) ----
    try:
        shutil.copy2(str(src), str(dest))
        # Source may stay briefly locked by OneDrive/AV — retry the unlink
        for attempt in range(1, 6):
            try:
                src.unlink()
                break
            except (OSError, PermissionError):
                gc.collect()
                time.sleep(0.7 * attempt)
        tqdm.write(f"     📁 {overall_status} — moved via copy+delete "
                   f"fallback to {dest.parent.name}/{dest.name}")
        return True
    except Exception as ex2:
        tqdm.write(f"     ❌ Move failed after 5 attempts + fallback: "
                   f"{ex2 or last_err}")
        return False

# ---------- 7.F  VALIDATION MODULE DRIVER ----------
def run_validation_module(
    paths: Dict[str, Path], wb_path: Path,
    downloads: List[Dict[str, Any]]) -> Dict[str, int]:
    """
    Iterate over successful downloads whose DATE OF DOWNLOAD is today,
    run Gates 1-3, aggregate, and write VALIDATION worksheet.
    (Dual progress bar: batch + per-entity phases.)
    """
    print("\n" + "═" * 78)
    print("VALIDATION MODULE")
    print("═" * 78)
    ensure_workbook_closed(wb_path, save_before_close=True) 

    # ── local helper: clamp labels so bars stay single-line ──
    def _short(txt: str, width: int = 44) -> str:
        txt = str(txt).replace("\n", " ").strip()
        return txt if len(txt) <= width else txt[:width - 1] + "…"

    wb = load_workbook(wb_path, data_only=False, keep_links=False)
    ws_entity = wb["ENTITY_LIST"]
    hdr_entity = locate_headers(ws_entity, header_row=1)
    ws_val, hdr_val = ensure_validation_ws(wb)

    today_str = _fmt_ddmmmyyyy(date.today())
    tally = {"PASS": 0, "REVIEW": 0, "FAIL": 0, OCR_STATUS: 0, "SKIPPED": 0, "ERROR": 0}

    # Precompute (UEI, ENTITY NAME, FY) → ENTITY_LIST row index for value-pasting metadata
    entity_row_map: Dict[Tuple[str, str, str], int] = {}
    for r in range(2, ws_entity.max_row + 1):
        u = str(ws_entity.cell(row=r, column=hdr_entity["UEI"]).value or "").strip().upper()
        e = str(ws_entity.cell(row=r, column=hdr_entity["ENTITY NAME"]).value or "").strip().upper()
        f = str(ws_entity.cell(row=r, column=hdr_entity["FY"]).value or "").strip()
        entity_row_map[(u, e, f)] = r

    # Filter to today's downloads only
    to_validate = [d for d in downloads if d.get("status") == "DOWNLOADED"]
    print(f"🔎 Validating {len(to_validate)} downloaded PDF(s) from today...\n")

    val_timings: List[float] = []

    # ─────────────────────────────────────────────────────────────────────
    # TIME-CELL WRITER — writes a duration DIRECTLY on the openpyxl cell.
    #   • value  = seconds / 86400  → a REAL number (Excel day-fraction)
    #   • format = "hh:mm:ss.0"     → renders hh:mm:ss.s, STAYS numeric
    #   • font   = FONT_TIME (#757171, Aptos Narrow 9, not bold)
    #   • align  = left, no wrap, NO fill
    # Bypasses _write_cell so the time format can never be overridden/stringified.
    # ─────────────────────────────────────────────────────────────────────
    def _put_time_cell(row_idx: int, col_hdr: str, secs) -> None:
        ci = _validation_col(hdr_val, col_hdr)
        if not ci or row_idx is None:
            return
        try:
            day_fraction = round(float(secs or 0) / 86400.0, 10)
        except (TypeError, ValueError):
            day_fraction = 0.0
        cell = ws_val.cell(row=row_idx, column=ci)
        cell.value = day_fraction                 # NUMBER, not text
        cell.number_format = "hh:mm:ss.0"         # hh:mm:ss.s display
        cell.font = FONT_TIME                      # #757171, size 9, not bold
        cell.alignment = Alignment(horizontal="right", vertical="center",
                                   wrap_text=False)

    # ── OUTER bar: batch (one tick per PDF) — stays on line 0 ──
    batch = tqdm(total=len(to_validate), desc="Validating", unit="pdf",
                 position=0, leave=True, dynamic_ncols=True,
                 bar_format="{desc}: {percentage:3.0f}%|{bar:20}| "
                            "{n_fmt}/{total_fmt} [{elapsed}<{remaining}, "
                            "{rate_fmt}]{postfix}")

    # ── INNER bar: current entity's phases — line 1, self-clears ──
    phase_bar = tqdm(total=5, position=1, leave=False, dynamic_ncols=True,
                     bar_format="   ↳ {desc} {percentage:3.0f}%|{bar:15}| "
                                "{n_fmt}/{total_fmt}")

    for dl in to_validate:
        t_val_start = time.perf_counter()
        entity   = dl.get("entity",  "")
        state    = dl.get("state",   "")
        sector   = dl.get("sector",  "")
        ay       = dl.get("audit_year")
        pdf_path = Path(dl.get("pdf_path", ""))
        auditor  = dl.get("auditor",  "")

        batch.set_postfix_str(_short(f"{sector} {state} {entity}"[:40]), refresh=True)

        phase_bar.reset(total=5)
        phase_bar.set_description(_short(f"Prep     {entity}"))
        phase_bar.refresh()

        # Row lookups
        er = dl.get("row")
        wb_fye_raw = ws_entity.cell(row=er, column=hdr_entity["FYE"]).value if er else None
        wb_fy_raw  = ws_entity.cell(row=er, column=hdr_entity["FY"]).value  if er and "FY" in hdr_entity else None
        wb_uei     = ws_entity.cell(row=er, column=hdr_entity["UEI"]).value if er else ""
        wb_ein     = ws_entity.cell(row=er, column=hdr_entity["EIN"]).value if er and "EIN" in hdr_entity else ""

        # FYE reconciliation
        downloaded_audit_year = dl.get("audit_year")
        effective_fye, reconciliation_note = reconcile_fye_with_audit_year(
            wb_fye_raw, downloaded_audit_year, wb_fy_raw
        )
        if effective_fye:
            _month_str = format(effective_fye[0], "02d")
            wb_fye = f"{effective_fye[1]}-{_month_str}-01"
        else:
            wb_fye = wb_fye_raw

        if not pdf_path.exists():
            tqdm.write(f"   ⚠️  PDF missing on disk: {pdf_path}")
            tally["SKIPPED"] += 1
            batch.update(1)
            continue

        try:
            doc = fitz.open(pdf_path)
        except Exception as ex:
            tqdm.write(f"   ❌ Could not open {pdf_path.name}: {ex}")
            tally["ERROR"] += 1
            batch.update(1)
            continue

        step_times: Dict[str, float] = {}
        gate_times: Dict[str, float] = {}

        g1_status, g1_results = "ERROR", {}
        g2_status, g2_results = "ERROR", {}
        g3_status, g3_results, stmt_pages = "ERROR", {}, []
        overall = "ERROR"
        v_row = None

        try:
            shot_dir = screenshot_dir(paths, state, entity)

            try:
                heading_cache = {p: page_heading_status(doc, p)
                                 for p in range(doc.page_count)}
                stmt_pages = sorted([p for p, s in heading_cache.items()
                                     if s != "EXCLUDED"])
            except Exception:
                stmt_pages = list(range(doc.page_count))
            ocr_required = doc_requires_ocr(doc)

            # ---- PHASE 1: Detect Framework & Basis of Accounting ----
            phase_bar.set_description(_short(f"Framewk  {entity}")); phase_bar.refresh()
            try:
                framework_result = detect_framework_and_basis(doc)
                framework_display = framework_result["display_text"]
                framework_conf_pct = framework_result["confidence_pct"]
                framework_conf = framework_result["confidence"]
                framework_conf_display = f"{framework_conf} ({framework_conf_pct}%)"
                tqdm.write(f"     🏛  Framework: {framework_display}  "
                           f"[{framework_conf_display}]")
            except Exception as ex_fw:
                tqdm.write(f"     ⚠️  Framework detection failed: {ex_fw}")
                framework_display = "UNCLASSIFIED"
                framework_conf_display = "ERROR"
                framework_conf_pct = 0
            phase_bar.update(1)

            # ---- PHASE 2: GATE 1 ----
            phase_bar.set_description(_short(f"Gate 1   {entity}")); phase_bar.refresh()
            t_g1 = time.perf_counter()
            try:
                g1_status, g1_results = run_gate1_timed(
                    doc, entity, state, wb_fye, sector, ay, shot_dir,
                    stmt_pages, step_times
                )
            except Exception as ex:
                tqdm.write(f"   ⚠️  Gate 1 crashed: {ex}")
            gate_times["G1"] = time.perf_counter() - t_g1
            phase_bar.update(1)

            if reconciliation_note and ("ADVANCED" in reconciliation_note
                                        or "RETRACTED" in reconciliation_note):
                if "G1-3" in g1_results and g1_results["G1-3"] is not None:
                    compact = (reconciliation_note
                               .replace("⚡ ADVANCED FYE: ", "FYE ADVANCED ")
                               .replace("⚠ RETRACTED FYE: ", "FYE RETRACTED ")
                               .replace(" → ", " -> ")
                               .strip())
                    g1_results["G1-3"].reconciliation_note = compact

            # ---- PHASE 3: GATE 2 ----
            phase_bar.set_description(_short(f"Gate 2   {entity}")); phase_bar.refresh()
            t_g2 = time.perf_counter()
            try:
                g2_status, g2_results = run_gate2_timed(
                    doc, wb_fye, auditor, sector, state, ay, shot_dir, step_times
                )
            except Exception as ex:
                tqdm.write(f"   ⚠️  Gate 2 crashed: {ex}")
            gate_times["G2"] = time.perf_counter() - t_g2
            phase_bar.update(1)

            # ---- PHASE 4: GATE 3 ----
            phase_bar.set_description(_short(f"Gate 3   {entity}")); phase_bar.refresh()
            t_g3 = time.perf_counter()
            try:
                g3_status, g3_results, stmt_pages = run_gate3_timed(
                    doc, sector, state, ay, wb_fye, shot_dir, step_times
                )
            except Exception as ex:
                tqdm.write(f"   ⚠️  Gate 3 crashed: {ex}")
            gate_times["G3"] = time.perf_counter() - t_g3
            phase_bar.update(1)

            # ---- OCR overlay: stamp OCR REQUIRED at step + gate level ----
            g1_status = apply_ocr_overlay(ocr_required, g1_status, g1_results)
            g2_status = apply_ocr_overlay(ocr_required, g2_status, g2_results)
            g3_status = apply_ocr_overlay(ocr_required, g3_status, g3_results)
            if ocr_required:
                tqdm.write("     🔍 OCR REQUIRED — no usable text layer on content "
                           "pages; text validators cannot run.")

            overall = overall_validation_status(g1_status, g2_status, g3_status)

            # ---- PHASE 5: WRITE WORKSHEET + MOVE PDF ----
            phase_bar.set_description(_short(f"Write    {entity}")); phase_bar.refresh()

            try:
                v_row = find_or_create_validation_row(ws_val, hdr_val, wb_uei, entity, ay)

                def _put(col_hdr, value, fill=None, font=None, force_fill=False, horizontal=None, number_format=None):
                    ci = _validation_col(hdr_val, col_hdr)
                    if ci:
                        _write_cell(ws_val, v_row, ci, value, fill=fill, font=font,
                                    force_fill=force_fill, horizontal=horizontal,
                                    number_format=number_format)

                _put("STATE", state, horizontal="center")
                _put("SECTOR", sector, horizontal="center")
                _put("ENTITY NAME", entity)
                _put("FY", int(ay) if ay else None, horizontal="center", number_format="0")
                _put("FYE", _format_fye_as_mmdd(wb_fye_raw), horizontal="center")
                _put("UEI", wb_uei, horizontal="center")
                _put("EIN", wb_ein, horizontal="center")
                _put("DATE OF DOWNLOAD", today_str, horizontal="center")
                _put("DOWNLOAD STATUS", "DOWNLOADED", font=FONT_GREEN, horizontal="center")
                _put("VALIDATION STATUS", overall, font=_font_for_status(overall), horizontal="center")

                fw_font = (FONT_GREEN if framework_conf_pct >= 90 else
                           FONT_YELLOW if framework_conf_pct >= 40 else
                           FONT_RED)
                _put("FRAMEWORK", framework_display, font=FONT_DEFAULT, horizontal="left")
                _put("FRAMEWORK CONFIDENCE", framework_conf_display, font=fw_font, horizontal="center")

                _put("VALIDATION GATE 1: ISSUER DETAILS VALIDATION", g1_status, font=_font_for_status(g1_status), horizontal="center")
                _put("VALIDATION GATE 2: AUDIT COMPLIANCE VALIDATION", g2_status, font=_font_for_status(g2_status), horizontal="center")
                _put("VALIDATION GATE 3: FINANCIAL STATEMENT STRUCTURE VALIDATION", g3_status, font=_font_for_status(g3_status), horizontal="center")

                for code, res in g1_results.items(): write_step_result(ws_val, v_row, hdr_val, res)
                for code, res in g2_results.items(): write_step_result(ws_val, v_row, hdr_val, res)
                for code, res in g3_results.items(): write_step_result(ws_val, v_row, hdr_val, res)

                if "VALIDATION STATUS" in hdr_entity:
                    _write_cell(ws_entity, er, hdr_entity["VALIDATION STATUS"],
                                overall, font=_font_for_status(overall))
                if "VALIDATED BY" in hdr_entity:
                    _write_cell(ws_entity, er, hdr_entity["VALIDATED BY"], "SYSTEM")
                if "VALIDATED ON" in hdr_entity:
                    _write_cell(ws_entity, er, hdr_entity["VALIDATED ON"], today_str)
                if "FRAMEWORK" in hdr_entity:
                    _write_cell(ws_entity, er, hdr_entity["FRAMEWORK"], framework_display, font=FONT_DEFAULT, horizontal="left")

                if (effective_fye and downloaded_audit_year and reconciliation_note
                    and any(marker in reconciliation_note for marker in ("ADVANCED", "RETRACTED", "CORRECTED"))):

                    original_day = _extract_original_fye_day(wb_fye_raw)

                    import calendar as _cal
                    last_day = _cal.monthrange(effective_fye[1], effective_fye[0])[1]

                    new_day = min(original_day, last_day)
                    new_fye_dt = datetime(effective_fye[1], effective_fye[0], new_day)

                    if "FY" in hdr_entity:
                        current_fy = ws_entity.cell(row=er, column=hdr_entity["FY"]).value
                        try:
                            current_fy_int = int(current_fy) if current_fy else 0
                        except (ValueError, TypeError):
                            current_fy_int = 0
                        if current_fy_int != int(downloaded_audit_year):
                            _write_cell(ws_entity, er, hdr_entity["FY"],
                                        int(downloaded_audit_year),
                                        fill=FILL_YELLOW, font=FONT_DEFAULT,
                                        force_fill=True)
                            tqdm.write(f"     🔄 FY auto-updated: {current_fy_int} → "
                                       f"{downloaded_audit_year} (light yellow)")

                    if "FYE" in hdr_entity:
                        _write_cell(ws_entity, er, hdr_entity["FYE"],
                                    new_fye_dt,
                                    fill=FILL_YELLOW, font=FONT_DEFAULT,
                                    force_fill=True,
                                    number_format="dd/mm/yyyy")
                        tqdm.write(f"     🔄 FYE auto-updated: "
                                   f"{new_fye_dt.strftime('%d/%m/%Y')} (light yellow)")

                if overall == "FAIL":
                    try:
                        if "DATE OF DOWNLOAD" in hdr_entity:
                            _write_cell(ws_entity, er, hdr_entity["DATE OF DOWNLOAD"], "", font=FONT_DEFAULT, horizontal="center")
                        if "DOWNLOAD STATUS" in hdr_entity:
                            _write_cell(ws_entity, er, hdr_entity["DOWNLOAD STATUS"], "NOT FOUND", font=FONT_DEFAULT, horizontal="center")
                    except Exception as ex_reset:
                        tqdm.write(f"     ⚠️  FAIL reset failed: {ex_reset}")

                wb.save(wb_path)
            except Exception as ex_write:
                tqdm.write(f"   ❌ Failed to write VALIDATION sheet: {ex_write}")

            # ---- Move PDF based on overall validation status ----
            try:
                if overall == "PASS":
                    dest_dir = paths["VALIDATED"]
                    dest_dir.mkdir(parents=True, exist_ok=True)
                    dest = dest_dir / pdf_path.name
                    try: doc.close()
                    except Exception: pass
                    _move_pdf_safely(pdf_path, dest, overall)

                elif overall == "REVIEW":
                    dest_dir = paths["VALIDATED"] / "REVIEW"
                    dest_dir.mkdir(parents=True, exist_ok=True)
                    dest = dest_dir / pdf_path.name
                    try: doc.close()
                    except Exception: pass
                    _move_pdf_safely(pdf_path, dest, overall)

                elif overall == OCR_STATUS:
                    dest_dir = paths["VALIDATED"] / "OCR REQUIRED"
                    dest_dir.mkdir(parents=True, exist_ok=True)
                    dest = dest_dir / pdf_path.name
                    try: doc.close()
                    except Exception: pass
                    _move_pdf_safely(pdf_path, dest, overall)

                elif overall == "FAIL":
                    tqdm.write(f"     ⏭  PDF retained in DOWNLOADED REPORTS "
                               f"(overall status: FAIL)")

                else:
                    tqdm.write(f"     ⚠️  Unknown overall status: {overall!r} — "
                               f"PDF left in place")

            except Exception as ex:
                tqdm.write(f"     ⚠️  Could not route PDF for {entity[:30]}: {ex}")

            phase_bar.update(1)

        except Exception as ex:
            tqdm.write(f"   ❌ Unexpected validation error for {entity}: {ex}")
        finally:
            try:
                doc.close()
            except Exception:
                pass

            # Validation time = full per-entity time in THIS module (incl. Phase-5 write).
            val_elapsed = time.perf_counter() - t_val_start
            val_timings.append(val_elapsed)
            src_time = float(dl.get("sourcing_seconds", 0) or 0)

            # OVERALL = sourcing + validation (validation already includes the
            # workbook-write phase, per your definition).
            overall_elapsed = src_time + val_elapsed

            # ---- Write the three TIMING columns (numeric Excel durations) ----
            if v_row is not None:
                try:
                    _put_time_cell(v_row, "SOURCING TIME",   src_time)
                    _put_time_cell(v_row, "VALIDATION TIME", val_elapsed)
                    _put_time_cell(v_row, "OVERALL TIME",    overall_elapsed)
                    wb.save(wb_path)
                except Exception as ex_time:
                    tqdm.write(f"     ⚠️  Time-column write failed: {ex_time}")

            # SINGLE authoritative tally increment per entity.
            tally[overall] = tally.get(overall, 0) + 1

            # ---- BOXED PER-ISSUER SUMMARY ----
            _print_issuer_summary(
                entity=entity, state=state, sector=sector, ay=ay,
                g1_status=g1_status, g1_results=g1_results,
                g2_status=g2_status, g2_results=g2_results,
                g3_status=g3_status, g3_results=g3_results,
                overall=overall,
                gate_times=gate_times, step_times=step_times,
                sourcing_time=src_time, validation_time=val_elapsed
            )

            batch.update(1)

    # ── loop finished — NOW close the bars ──
    phase_bar.close()
    batch.close()

    wb.save(wb_path)
    print(f"\n✅ Validation complete. Workbook saved: {wb_path}")
    return {"tally": tally,
            "validation_timings": val_timings,
            "validation_wall_seconds": sum(val_timings)}

# ---------- 7.G  MAIN ORCHESTRATOR (ENTRY POINT) ----------
def main():
    """Single-cell entry point.  Runs everything end-to-end with full timing."""
    run_start = time.perf_counter()
    run_start_wall = datetime.now()

    print(f"  Run started: {run_start_wall:%Y-%m-%d %H:%M:%S}".ljust(76))

    # Step 1 — Folder architecture
    paths = ensure_architecture()
    print("📁 Folder architecture verified:")
    for k, p in paths.items():
        print(f"   {k:<12}: {p}")

    # Step 2 — Workbook resolution
    wb_path = resolve_workbook(paths["WORKBOOK"])
    print(f"\n📘 Using workbook: {wb_path}\n")

    # Step 3 — Sourcing Module
    t_src_start = time.perf_counter()
    src_result = run_sourcing_module(paths, wb_path)
    downloads = src_result["downloads"] if isinstance(src_result, dict) else src_result
    src_timings = src_result.get("sourcing_timings", []) if isinstance(src_result, dict) else []
    src_wall_secs = time.perf_counter() - t_src_start

    # Step 4 — Validation Module
    t_val_start = time.perf_counter()
    val_result = run_validation_module(paths, wb_path, downloads)
    tally = val_result["tally"] if isinstance(val_result, dict) else val_result
    val_timings = val_result.get("validation_timings", []) if isinstance(val_result, dict) else []
    val_wall_secs = time.perf_counter() - t_val_start

    run_end_wall = datetime.now()
    total_secs = time.perf_counter() - run_start

    # Human-friendly duration formatter
    def _fmt_hms(secs: float) -> str:
        h = int(secs // 3600)
        m = int((secs % 3600) // 60)
        s = secs - (h * 3600) - (m * 60)
        if h > 0:  return f"{h}h {m}m {s:.1f}s"
        if m > 0:  return f"{m}m {s:.1f}s"
        return f"{s:.2f}s"

    # ─── FINAL SUMMARY ───
    print("\n" + "═" * 78)
    print("FINAL RUN SUMMARY".center(78))
    print("═" * 78)
    print(f"\n📊 SOURCING METRICS")
    print(f"   Total attempts        : {len(downloads)}")
    print(f"   Downloaded            : {sum(1 for d in downloads if d.get('status') == 'DOWNLOADED')}")
    print(f"   Not Found             : {sum(1 for d in downloads if d.get('status') == 'NOT FOUND')}")
    print(f"   Errors                : {sum(1 for d in downloads if d.get('status') == 'ERROR')}")
    if src_timings:
        print(f"   Avg time / issuer     : {(sum(src_timings)/len(src_timings)):.2f}s")
        print(f"   Fastest / slowest     : {min(src_timings):.2f}s / {max(src_timings):.2f}s")
    print(f"   Total sourcing time   : {_fmt_hms(src_wall_secs)}")

    print(f"\n🔍 VALIDATION METRICS")
    print(f"   PASS                  : {tally.get('PASS', 0)}")
    print(f"   REVIEW                : {tally.get('REVIEW', 0)}")
    print(f"   FAIL                  : {tally.get('FAIL', 0)}")
    print(f"   OCR REQUIRED          : {tally.get('OCR REQUIRED', 0)}")
    print(f"   Skipped / Errors      : {tally.get('SKIPPED', 0) + tally.get('ERROR', 0)}")
    if val_timings:
        print(f"   Avg time / issuer     : {(sum(val_timings)/len(val_timings)):.2f}s")
        print(f"   Fastest / slowest     : {min(val_timings):.2f}s / {max(val_timings):.2f}s")
    print(f"   Total validation time : {_fmt_hms(val_wall_secs)}")

    print(f"\n⏱  OVERALL")
    print(f"   Run started           : {run_start_wall:%Y-%m-%d %H:%M:%S}")
    print(f"   Run finished          : {run_end_wall:%Y-%m-%d %H:%M:%S}")
    print(f"   Total completion time : {_fmt_hms(total_secs)}")
    print("═" * 78 + "\n")

# ---------- 7.H  KICK IT OFF (Shift+Enter runs this cell → main() executes) ----------
if __name__ == "__main__":
    main()