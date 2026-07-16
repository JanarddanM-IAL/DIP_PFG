"""
Drop this in your project folder and run:
  python debug_dsr.py "path/to/LG_CIT_IA_600024512_2025.pdf"

It prints the full text of pages 68-78 and 160-170, plus runs
_page_has_dsr_table() with verbose output to show exactly which
condition is failing.
"""
import sys, re
from pathlib import Path

PDF = sys.argv[1] if len(sys.argv) > 1 else r"c:\S2\S2_Khushbu\AI Projects\Financial Data Extraction\MDB_multipleAI\03_Validated_Report\LG_CIT_IA_600024512_2025.pdf"

SUSPECT_PAGES = list(range(68, 79)) + [165]   # pages to dump

# ── copy of regexes from llm_page_identifier ──────────────────────────
_DSR_YEAR_ROW_RE = re.compile(
    r"^\s*(20[2-9]\d)(?:\s*[-–]\s*20[2-9]\d)?\s+(?:\$\s*)?[\d,]+",
    re.MULTILINE,
)
_BOND_SERIES_ROW_RE = re.compile(
    r"^\s*20\d\d[A-Z]?\s+\d{1,2}/\d{1,2}/\d{2,4}\b",
    re.MULTILINE,
)
_DSR_PRINCIPAL_COL_RE = re.compile(r"\bPrincipal\b", re.IGNORECASE)
_DSR_INTEREST_COL_RE  = re.compile(r"\bInterest\b(?:\s*\(\d+\))?", re.IGNORECASE)
_SWAP_TERMS_RE = re.compile(
    r"\bNotional\b|\bCounty\s+Pays\b|\bCounty\s+Receives\b"
    r"|\bFair\s+Value\b|\bSwap\s+#\b|\bSwap\s+Description\b"
    r"|\bAssociated\s+Variable\s+Rate\b",
    re.IGNORECASE,
)

def get_text(pdf_path, page_no):
    try:
        import pdfplumber
        with pdfplumber.open(pdf_path) as pdf:
            if 1 <= page_no <= len(pdf.pages):
                return pdf.pages[page_no-1].extract_text() or ""
    except Exception:
        pass
    try:
        from pypdf import PdfReader
        r = PdfReader(pdf_path)
        if 1 <= page_no <= len(r.pages):
            return r.pages[page_no-1].extract_text() or ""
    except Exception:
        pass
    return ""

def diagnose(text, page_no):
    print(f"\n{'='*60}")
    print(f"PAGE {page_no} DIAGNOSIS")
    print(f"{'='*60}")
    
    year_rows = _DSR_YEAR_ROW_RE.findall(text)
    bond_rows = _BOND_SERIES_ROW_RE.findall(text)
    swap_hit  = _SWAP_TERMS_RE.search(text)
    princ_hit = _DSR_PRINCIPAL_COL_RE.search(text)
    int_matches = list(_DSR_INTEREST_COL_RE.finditer(text))
    
    print(f"  year_dollar_rows ({len(year_rows)}): {year_rows[:5]}")
    print(f"  bond_series_rows ({len(bond_rows)}): {bond_rows[:3]}")
    print(f"  swap_terms hit  : {bool(swap_hit)}")
    print(f"  Principal hit   : {bool(princ_hit)}")
    
    has_interest = False
    for m in int_matches:
        after = text[m.end():m.end()+10].strip()
        is_rate = after.lower().startswith("rate")
        print(f"  Interest match: '{m.group()}' → after='{after[:10]}' is_rate={is_rate}")
        if not is_rate:
            has_interest = True
    print(f"  Interest col OK : {has_interest}")
    
    # Gate results
    if swap_hit and (not year_rows or len(bond_rows) >= len(year_rows)):
        print(f"  RESULT: REJECTED — swap table")
        return
    if len(year_rows) < 2:
        print(f"  RESULT: REJECTED — need ≥2 year rows, got {len(year_rows)}")
        # Show all lines with 20xx to help understand format
        print(f"\n  Lines containing '20xx':")
        for ln in text.splitlines():
            if re.search(r'\b20[2-9]\d\b', ln):
                print(f"    {repr(ln)}")
        return
    if not princ_hit:
        print(f"  RESULT: REJECTED — no Principal header")
        return
    if not has_interest:
        print(f"  RESULT: REJECTED — no Interest column (only 'Interest Rate'?)")
        return
    print(f"  RESULT: ACCEPTED ✓")

for pg in SUSPECT_PAGES:
    text = get_text(PDF, pg)
    if not text:
        print(f"\nPage {pg}: NO TEXT EXTRACTED")
        continue
    print(f"\n{'─'*60}")
    print(f"PAGE {pg} RAW TEXT (first 60 lines):")
    for i, ln in enumerate(text.splitlines()[:60]):
        print(f"  {i+1:3d}: {repr(ln)}")
    diagnose(text, pg)