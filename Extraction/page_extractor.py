"""
page_extractor.py  (DSR scoring-based rewrite + DEBT schedule)
=========================
Extracts all 8 financial statement sections + DSR table + DEBT schedule
from government CAFR/ACFR PDFs.

DSR detection uses a SCORING model (threshold >= 8).
DEBT detection uses a SCORING model (threshold >= 4).
"""
import logging
import os
import re
import sys
import datetime
import subprocess
import tempfile
import pdfplumber
from pypdf import PdfReader, PdfWriter
logging.getLogger("pypdf").setLevel(logging.ERROR)
# ── Global scan start (1-based) ───────────────────────────────────────────────
START_PAGE = 4

# ══════════════════════════════════════════════════════════════════════════════
# SHARED EXCLUSION PATTERNS
# ══════════════════════════════════════════════════════════════════════════════
EXCLUDE_PATTERNS = [
    re.compile(r"^\s*table\s+of\s+contents?\b",          re.IGNORECASE),
    re.compile(r"\btable\s+of\s+contents?\b",             re.IGNORECASE),
    re.compile(r"^\s*contents?\b",                        re.IGNORECASE),
    re.compile(r"^\s*foreword\b",                         re.IGNORECASE),
    re.compile(r"\bforeword\b",                           re.IGNORECASE),
    re.compile(r"^\s*introduction\b",                     re.IGNORECASE),
    re.compile(r"\bintroduction\b",                       re.IGNORECASE),
    re.compile(r"\bmd\s*(?:&|and)\s*a\b",                 re.IGNORECASE),
    re.compile(r"\bmanagement[\u2018\u2019']?s?\s+discussion\s+(?:&|and)\s+analysis\b",
               re.IGNORECASE),
    re.compile(r"\bbudget\b",                             re.IGNORECASE),
    re.compile(r"\bstatistical\b",                        re.IGNORECASE),
    re.compile(r"\bcombining\b",                          re.IGNORECASE),
    re.compile(r"\bsupplemental\b",                       re.IGNORECASE),
    re.compile(r"\bschedule\b",                           re.IGNORECASE),
    re.compile(r"\breconciliation\b",                     re.IGNORECASE),
    re.compile(r"\bcustodial\s+fund",                     re.IGNORECASE),
    re.compile(r"\bfiduciary\b",                          re.IGNORECASE),
    re.compile(r"\bnotes?\s+to\b",                        re.IGNORECASE),
]

_PROP_SNP_RECONCILIATION_RE = re.compile(
    r"\breconciliation\b"
    r"|\bAmounts\s+reported\s+for\s+governmental\s+activities\s+in\s+the\s+Statement\s+of\s+Net\s+Position\s+are\s+different\b"
    r"|\bTotal\s+fund\s+balances?\s*[-–—]\s*governmental\s+funds?\b",
    re.IGNORECASE,
)

_FUND_FOOTER_NOTES_REF_RE = re.compile(
    r"\bSee\s+(?:accompanying\s+)?notes?\s+to\s+(?:the\s+)?(?:basic\s+)?(?:financial|required)\b"
    r"|\bThe\s+accompanying\s+notes?\s+are\s+an\s+integral\s+part\b"
    r"|\bNotes?\s+are\s+an\s+integral\s+part\b",
    re.IGNORECASE,
)
SNP_PAGE_N_OF_M_RE = re.compile(
    r"\bpage\s+([2-9]|\d{2,})\s+of\s+\d+\b",
    re.IGNORECASE,
)
_RECONCILIATION_BODY_RE = re.compile(
    r'\bAmounts\s+reported\s+for\s+governmental\s+activities\b'
    r'.*\bare\s+different\s+because\b'
    r'|\bReconciliation\s+of\s+(?:the\s+)?(?:Balance\s+Sheet|Statement)\b'
    r'|\bNet\s+change\s+in\s+fund\s+balances?\s*[-–]\s*governmental\s+funds\b'
    r'|\bTotal\s+(?:fund\s+)?net\s+position\s*[-–]\s*governmental\s+activities\b',
    re.IGNORECASE | re.DOTALL,
)
SNP_EXCLUDE_PATTERNS = [
    re.compile(r"^\s*table\s+of\s+contents?\b", re.IGNORECASE),
    re.compile(r"\btable\s+of\s+contents?\b",   re.IGNORECASE),
    re.compile(r"\bforeword\b",                 re.IGNORECASE),
    re.compile(r"\bintroduction\b",             re.IGNORECASE),
    re.compile(r"\bmd\s*(?:&|and)\s*a\b",       re.IGNORECASE),
    re.compile(r"\bmanagement[\u2018\u2019']?s?\s+discussion\s+(?:&|and)\s+analysis\b",
               re.IGNORECASE),
    re.compile(r"\bbudget\b",                   re.IGNORECASE),
    re.compile(r"\bsupplemental\b",             re.IGNORECASE),
    re.compile(r"\bproprietary\s+fund\b",       re.IGNORECASE),
    re.compile(r"\bstatistical\b",              re.IGNORECASE),
    re.compile(r"\bnon[-\s]?major\b",           re.IGNORECASE),
    re.compile(r"\bcondensed\b",                re.IGNORECASE),
]

_CLAIMS_LIABILITY_MARKERS = [
    re.compile(r"\bclaims?\s+liabilit", re.IGNORECASE),
    re.compile(r"\bself[-\s]?insur", re.IGNORECASE),
    re.compile(r"\bworkers[\u2019']?\s+compensation\b", re.IGNORECASE),
    re.compile(r"\bcurrent[-\s]?year\s+claims\s+and\s+estimates\b", re.IGNORECASE),
    re.compile(r"\bon[-\s]the[-\s]job\s+injury\b", re.IGNORECASE),
    re.compile(r"\brisk\s+management\b", re.IGNORECASE),
    re.compile(r"\brisk\s+financing\b", re.IGNORECASE),
    re.compile(r"\bcommercial\s+insurance\b", re.IGNORECASE),
    re.compile(r"\bstop[/\s]?loss\b", re.IGNORECASE),
]

_DEBT_SPECIFIC_HEADERS_FOR_CLAIMS_CHECK = re.compile(
    r"\bLong[-\s]?[Tt]erm\s+(?:Debt|Obligation|Liabilit)\b"
    r"|\bBonds?\s+[Pp]ayable\b"
    r"|\bNotes?\s+[Pp]ayable\b"
    r"|\bGeneral\s+Obligation\b",
    re.IGNORECASE,
)
"""
extract_all_statements.py  (DSR scoring-based rewrite + DEBT schedule)
=========================
Extracts all 8 financial statement sections + DSR table + DEBT schedule
from government CAFR/ACFR PDFs.

DSR detection uses a SCORING model (threshold >= 8).
DEBT detection uses a SCORING model (threshold >= 4).
"""

import os
import re
import sys
import datetime
import subprocess
import tempfile
import pdfplumber
from pypdf import PdfReader, PdfWriter

# ── Global scan start (1-based) ───────────────────────────────────────────────
START_PAGE = 4

# ══════════════════════════════════════════════════════════════════════════════
# SHARED EXCLUSION PATTERNS
# ══════════════════════════════════════════════════════════════════════════════
EXCLUDE_PATTERNS = [
    re.compile(r"^\s*table\s+of\s+contents?\b",          re.IGNORECASE),
    re.compile(r"\btable\s+of\s+contents?\b",             re.IGNORECASE),
    re.compile(r"^\s*contents?\b",                        re.IGNORECASE),
    re.compile(r"^\s*foreword\b",                         re.IGNORECASE),
    re.compile(r"\bforeword\b",                           re.IGNORECASE),
    re.compile(r"^\s*introduction\b",                     re.IGNORECASE),
    re.compile(r"\bintroduction\b",                       re.IGNORECASE),
    re.compile(r"\bmd\s*(?:&|and)\s*a\b",                 re.IGNORECASE),
    re.compile(r"\bmanagement[\u2018\u2019']?s?\s+discussion\s+(?:&|and)\s+analysis\b",
               re.IGNORECASE),
    re.compile(r"\bbudget\b",                             re.IGNORECASE),
    re.compile(r"\bstatistical\b",                        re.IGNORECASE),
    re.compile(r"\bcombining\b",                          re.IGNORECASE),
    re.compile(r"\bsupplemental\b",                       re.IGNORECASE),
    re.compile(r"\bschedule\b",                           re.IGNORECASE),
    re.compile(r"\breconciliation\b",                     re.IGNORECASE),
    re.compile(r"\bcustodial\s+fund",                     re.IGNORECASE),
    re.compile(r"\bfiduciary\b",                          re.IGNORECASE),
    re.compile(r"\bnotes?\s+to\b",                        re.IGNORECASE),
]
SNP_PAGE_N_OF_M_RE = re.compile(
    r"\bpage\s+([2-9]|\d{2,})\s+of\s+\d+\b",
    re.IGNORECASE,
)
SNP_EXCLUDE_PATTERNS = [
    re.compile(r"^\s*table\s+of\s+contents?\b", re.IGNORECASE),
    re.compile(r"\btable\s+of\s+contents?\b",   re.IGNORECASE),
    re.compile(r"\bforeword\b",                 re.IGNORECASE),
    re.compile(r"\bintroduction\b",             re.IGNORECASE),
    re.compile(r"\bmd\s*(?:&|and)\s*a\b",       re.IGNORECASE),
    re.compile(r"\bmanagement[\u2018\u2019']?s?\s+discussion\s+(?:&|and)\s+analysis\b",
               re.IGNORECASE),
    re.compile(r"\bbudget\b",                   re.IGNORECASE),
    re.compile(r"\bsupplemental\b",             re.IGNORECASE),
    re.compile(r"\bproprietary\s+fund\b",       re.IGNORECASE),
    re.compile(r"\bstatistical\b",              re.IGNORECASE),
    re.compile(r"\bnon[-\s]?major\b",           re.IGNORECASE),
    re.compile(r"\bcondensed\b",                re.IGNORECASE),
]

_CLAIMS_LIABILITY_MARKERS = [
    re.compile(r"\bclaims?\s+liabilit", re.IGNORECASE),
    re.compile(r"\bself[-\s]?insur", re.IGNORECASE),
    re.compile(r"\bworkers[\u2019']?\s+compensation\b", re.IGNORECASE),
    re.compile(r"\bcurrent[-\s]?year\s+claims\s+and\s+estimates\b", re.IGNORECASE),
    re.compile(r"\bon[-\s]the[-\s]job\s+injury\b", re.IGNORECASE),
    re.compile(r"\brisk\s+management\b", re.IGNORECASE),
    re.compile(r"\brisk\s+financing\b", re.IGNORECASE),
    re.compile(r"\bcommercial\s+insurance\b", re.IGNORECASE),
    re.compile(r"\bstop[/\s]?loss\b", re.IGNORECASE),
]

_DEBT_SPECIFIC_HEADERS_FOR_CLAIMS_CHECK = re.compile(
    r"\bLong[-\s]?[Tt]erm\s+(?:Debt|Obligation|Liabilit)\b"
    r"|\bBonds?\s+[Pp]ayable\b"
    r"|\bNotes?\s+[Pp]ayable\b"
    r"|\bGeneral\s+Obligation\b",
    re.IGNORECASE,
)


def _is_claims_liability_page(text: str) -> bool:
    hits = sum(1 for p in _CLAIMS_LIABILITY_MARKERS if p.search(text or ""))
    has_debt_header = bool(_DEBT_SPECIFIC_HEADERS_FOR_CLAIMS_CHECK.search(text or ""))
    return hits >= 2 and not has_debt_header
# ══════════════════════════════════════════════════════════════════════════════
# GOVERNMENT-WIDE SNP
# ══════════════════════════════════════════════════════════════════════════════
SNP_TITLE_RE = re.compile(
    r"(?:exhibit\s+[a-z]?\s*[-–—]?\s*)?"  # Optional "Exhibit A –"
    r"statement\s+of\s+net\s+pos(?:i)?tion\b"
    r"|government[-\s]*wide\s*[-–—]?\s*statement\s+of\s+net\s+pos(?:i)?tion\b",
    re.IGNORECASE | re.MULTILINE,
)

SNP_FINANCIAL_ROW_KEYWORDS = [
    re.compile(r"\btotal\s+assets\b",   re.IGNORECASE),
    re.compile(r"\btotal\s+liabilit",   re.IGNORECASE),
    re.compile(r"\bnet\s+position\b",   re.IGNORECASE),
]

SNP_BODY_CONTINUATION_RE = re.compile(
    r"^\s*(?:"
    r"deferred\s+(?:inflows?|outflows?)\s+of\s+resources"  # Deferred Inflows section
    r"|net\s+position"  # Net Position line
    r"|restricted\s+(?:for|net\s+position)"  # Restricted net position
    r"|unrestricted"  # Unrestricted net position
    r"|net\s+(?:investment|restricted|unrestricted)"  # Net investment categories
    r")\b",
    re.IGNORECASE | re.MULTILINE,
)

SNP_COMPONENT_UNIT_PAGE_START_RE = re.compile(
    r"^\s*component\s+unit",
    re.IGNORECASE,
)
def _first_nonblank_line(text: str) -> str:
    for line in (text or "").splitlines():
        if line.strip():
            return line
    return ""
# ── Outstanding debt summary patterns ────────────────────────────
# Detects tables like:
#   Type | Interest Rate | Final Maturity | Original Amount | Balance
# (debt issued summary, NOT a changes-in-debt schedule)

_OUTSTANDING_INTEREST_RATE_RE = re.compile(
    r"\bInterest\s+Rate\b", re.IGNORECASE,
)
_OUTSTANDING_MATURITY_RE = re.compile(
    r"\bFinal\s+Maturity\b"
    r"|\bMaturity\s+Date\b"
    r"|\bDate\s+of\s+Maturity\b",
    re.IGNORECASE,
)
_OUTSTANDING_AMOUNT_RE = re.compile(
    r"\bOriginal\s+Amount\b"
    r"|\bAmount\s+of\s+Issue\b"
    r"|\bOriginal\s+Issue\s+Amount\b"
    r"|\bIssue\s+Amount\b"
    r"|\bDate\s+of\s+Issue\b",
    re.IGNORECASE,
)
_OUTSTANDING_DEBT_CONTEXT_RE = re.compile(
    r"\blong[-\s]?term\s+(?:debt|obligation|liabilit|bonds?)\b"
    r"|\bgeneral\s+obligation\s+(?:refunding\s+)?(?:bonds?|notes?)\b"
    r"|\brevenue\s+(?:refunding\s+)?bonds?\b"
    r"|\bbonds?\s+payable\b"
    r"|\bnotes?\s+payable\b"
    r"|\bdebt\s+outstanding\b"
    r"|\boutstanding\s+(?:bonds?|notes?|debt|loans?)\b"
    r"|\bdirect\s+(?:borrowing|placement)\b",
    re.IGNORECASE,
)


def _snp_excluded(text: str) -> bool:
    first_lines = "\n".join((text or "").splitlines()[:10])
    return any(p.search(first_lines) for p in SNP_EXCLUDE_PATTERNS)


_NARRATIVE_CONTINUATION_RE = re.compile(
    r"^(?:and|the|is|in|,|of|to|that|for|on|are|was|were|which|where|"
    r"which\s+is|reported|presented|shown|found|described)\b",
    re.IGNORECASE,
)
_MAX_TITLE_LINE_LEN = 70


def _snp_title_in_header(text: str) -> bool:
    for line in (text or "").splitlines()[:25]:
        if SNP_TITLE_RE.match(line.strip()):
            return True
    return False



def _all_snp_financial_rows(text: str) -> bool:
    return all(p.search(text or "") for p in SNP_FINANCIAL_ROW_KEYWORDS)


def _any_snp_financial_row(text: str) -> bool:
    return any(p.search(text or "") for p in SNP_FINANCIAL_ROW_KEYWORDS)


def find_snp_pages(pdf_path: str) -> list:
    collected = []
    with pdfplumber.open(pdf_path) as pdf:
        total = len(pdf.pages)
        for i in range(START_PAGE - 1, total):
            page_no = i + 1
            text = (pdf.pages[i].extract_text() or "").replace(
                "Assets Activities Units", "Assets"
            )
            if _snp_excluded(text):
                continue
            if not (_snp_title_in_header(text) and
                    re.search(r"\bassets\b", text, re.IGNORECASE)):
                continue
 
            collected.append(page_no)
            combined = text
 
            if not _all_snp_financial_rows(combined):
                j = i + 1
                while j < total:
                    nt = pdf.pages[j].extract_text() or ""
                    if _snp_excluded(nt):
                        j += 1; continue
                    if _any_snp_financial_row(nt):
                        collected.append(j + 1)
                        combined += "\n" + nt
                        if _all_snp_financial_rows(combined):
                            break
                    else:
                        break
                    j += 1
 
            # ===== FIXED LOOP 2 =====
            # Start from next page after last collected
            j = collected[-1]  # j = 20 if collected = [20], or 21 if [20,21]
            
            while j < total:
                nt = pdf.pages[j].extract_text() or ""
                
                if _snp_excluded(nt):
                    j += 1; continue
                
                # ← MOVED DOWN: Check continuation BEFORE title
                first_line = _first_nonblank_line(nt)
                first6  = "\n".join(nt.splitlines()[:6])   # widened from 4
                first15 = "\n".join(nt.splitlines()[:15])
                
                # If it looks like SNP continuation, add it regardless of title
                is_continuation = (
                    bool(SNP_COMPONENT_UNIT_PAGE_START_RE.search(first_line)) or  # Fix 1
                    bool(SNP_BODY_CONTINUATION_RE.search(first15)) or
                    bool(SNP_PAGE_N_OF_M_RE.search(first15))                   # Fix 2
                )
                
                if is_continuation:
                    collected.append(j + 1)
                    j += 1
                    continue
                
                # ← NOW check for different statement's title
                # If it's a DIFFERENT statement (e.g., SOA), break
                if _snp_title_in_header(nt):
                    # Only break if it's NOT SNP continuation
                    # Check if this is still the SNP (has deferred/net position markers)
                    if not is_continuation:
                        break
                
                # If no continuation marker and not SNP continuation, stop
                break
 
            break
 
    if not collected:
        raise ValueError(f"Government-wide SNP not found in {os.path.basename(pdf_path)}")
    return sorted(set(collected))
 


# ══════════════════════════════════════════════════════════════════════════════
# GOVERNMENT-WIDE SOA
# ══════════════════════════════════════════════════════════════════════════════
SOA_TITLE_RE = re.compile(
    r"\b(?:"
    r"statement\s+of\s+activities"
    r"|statement\s+of\s+revenues,\s+expenses,\s+and\s+changes\s+in\s+net\s+position"
    r"|statement\s+of\s+revenues\s+expenses\s+and\s+changes\s+in\s+net\s+position"
    r"|statement\s+of\s+changes\s+in\s+net\s+position"
    r")\b",
    re.IGNORECASE,
)

SOA_BODY_CONTINUATION_RE = re.compile(
    r"^\s*(?:"
    r"general\s+revenues"
    r"|change\s+in\s+net\s+position"
    r"|changes\s+in\s+net\s+position"
    r"|net\s+position"
    r"|governmental\s+activities"
    r"|business[-\s]*type\s+activities"
    r"|primary\s+government"
    r"|component\s+units?"
    r"|program\s+revenues"
    r"|charges\s+for\s+services"
    r"|operating\s+grants"
    r"|capital\s+grants"
    r")\b",
    re.IGNORECASE | re.MULTILINE,
)

SOA_CONTINUATION_RE = re.compile(
    r"\bprimary\s+government\b.*\bcomponent\s+units?\b"
    r"|\bcomponent\s+units?\b.*\bprimary\s+government\b"
    r"|\bnet\s*\(?expense\)?\s+revenues?\b"
    r"|\bgovernmental\b\s+\bbusiness[-\s]*type\b",
    re.IGNORECASE | re.DOTALL,
)

_SOA_FALSE_POSITIVE_RE = re.compile(
    r"\bcondensed\b"
    r"|\banalysis\s+of\b"
    r"|\bfollowing\b"
    r"|\bsummary\b"
    # ── NEW: catch MD&A narrative lines that reference the SOA ──
    r"|\bas\s+shown\s+in\b"
    r"|\bpresents\s+information\b"
    r"|\bdesigned\s+to\s+provide\b",
    re.IGNORECASE,
)



def _soa_title_in_header(text: str) -> bool:
    lines = (text or "").splitlines()
    for line in lines[:25]:
        if SOA_TITLE_RE.search(line):
            if not _SOA_FALSE_POSITIVE_RE.search(line):
                return True
    return False


def _soa_start_keywords_ok(text: str) -> bool:
    patterns = [
        r"\bexpenses\b", r"\bprogram\s+revenues\b", r"\bfunctions/programs\b",
        r"\bfunctions\s*/\s*programs\b", r"\bnet\s+\(expense\)\s+revenue\b",
        r"\bnet\s+expense\s+revenue\b", r"\bcharges\s+for\s+services\b",
        r"\boperating\s+grants\s+and\s+contributions\b",
        r"\bcapital\s+grants\s+and\s+contributions\b",
        r"\bgeneral\s+revenues\b", r"\bchange\s+in\s+net\s+position\b",
        r"\bchanges\s+in\s+net\s+position\b", r"\bgovernmental\s+activities\b",
        r"\bbusiness[-\s]*type\s+activities\b",
    ]
    return sum(1 for p in patterns if re.search(p, text or "", re.IGNORECASE)) >= 2


def _is_soa_continuation(text: str) -> bool:
    first = "\n".join((text or "").splitlines()[:30])
    return (bool(SOA_CONTINUATION_RE.search(first)) or
            bool(SOA_BODY_CONTINUATION_RE.search(first)))


def find_soa_pages(pdf_path: str, start_page: int = START_PAGE) -> list:
    collected = []
    with pdfplumber.open(pdf_path) as pdf:
        total = len(pdf.pages)
        scan_from = max(start_page, START_PAGE) - 1

        for i in range(scan_from, total):
            page_no = i + 1
            text = pdf.pages[i].extract_text() or ""

            if _snp_excluded(text):
                continue
            if not (_soa_title_in_header(text) and _soa_start_keywords_ok(text)):
                continue

            collected.append(page_no)
            next_i = i + 1
            while next_i < total:
                next_text = pdf.pages[next_i].extract_text() or ""
                if _snp_title_in_header(next_text):
                    break
                if _is_soa_continuation(next_text):
                    collected.append(next_i + 1)
                    next_i += 1
                    continue
                if _soa_title_in_header(next_text):
                    break
                break

            break

    if not collected:
        raise ValueError(f"Statement of Activities not found in {pdf_path}")
    return sorted(set(collected))


# ══════════════════════════════════════════════════════════════════════════════
# SCRAMBLED-TITLE DETECTION
# ══════════════════════════════════════════════════════════════════════════════
_STATEMENT_TITLE_RE = re.compile(
    r"\bStatement\s+of\b|\bReconciliation\b|\bSchedule\b",
    re.IGNORECASE,
)

# ── GOV_BS ────────────────────────────────────────────────────────────────────
GOV_BS_SCRAMBLED_KW = [
    re.compile(r"t\s*o\s*t\s*a\s*l\s+a\s*s\s*s\s*e\s*t\s*s", re.IGNORECASE),
    re.compile(r"t\s*o\s*t\s*a\s*l\s+l\s*i\s*a\s*b\s*i\s*l", re.IGNORECASE),
    re.compile(r"f\s*u\s*n\s*d\s+b\s*a\s*l\s*a\s*n\s*c\s*e", re.IGNORECASE),
]
_GOV_BS_COMPLETE_RE = re.compile(
    r"\bTotal\s+(?:fund\s+balances?|liabilities[^$\n]{0,80}fund\s+balances?)\b",
    re.IGNORECASE | re.DOTALL,
)
_GOV_BS_TOTAL_COL_RE = re.compile(
    r"\bTotal\b.{0,80}\bGovernmental\b.{0,40}\bFunds?\b",
    re.IGNORECASE,
)

def _gov_bs_is_complete(text: str) -> bool:
    return bool(_GOV_BS_COMPLETE_RE.search(text))

def _gov_bs_is_column_overflow(text: str) -> bool:
    lines = (text or "").splitlines()
    joined_top = " ".join(lines[:8])
    has_total_col = (
        bool(re.search(r"\bTotal\b", joined_top, re.IGNORECASE))
        and bool(re.search(r"\bFunds?\b", joined_top, re.IGNORECASE))
    )
    if has_total_col:
        return True
    first_dollar = next((i for i, l in enumerate(lines) if "$" in l), 999)
    if first_dollar < 2:
        return False
    if _STATEMENT_TITLE_RE.search(joined_top):
        return False
    top12 = " ".join(lines[:12])
    if not re.search(r"\bTotals?\b", top12, re.IGNORECASE):
        return False
    has_numbers = any(re.search(r"\b\d[\d,]{2,}\b", l) for l in lines[2:8])
    return has_numbers

# ── GOV_IS ────────────────────────────────────────────────────────────────────
GOV_IS_SCRAMBLED_KW = [
    re.compile(r"\bexpenditures\b",      re.IGNORECASE),
    re.compile(r"\brevenues\b",          re.IGNORECASE),
    re.compile(r"\bfund\s+balance",      re.IGNORECASE),
    re.compile(r"\bgovernmental\s+fund", re.IGNORECASE),
]
_GOV_IS_COMPLETE_RE = re.compile(
    r"\bFund\s+balances?\s*[-–,]\s*(?:ending|June|July|August|September|October|"
    r"November|December|January|February|March|April|May)\b"
    r"|\bFUND\s+BALANCES\b.{0,30}(?:20\d\d|\bending\b)"
    r"|\bNet\s+change\s+in\s+fund\s+balances?\b",
    re.IGNORECASE,
)
_GOV_IS_TOTAL_COL_RE = re.compile(
    r"\bTotal\b.{0,80}\bGovernmental\b.{0,40}\bFunds?\b",
    re.IGNORECASE,
)

def _gov_is_is_complete(text: str) -> bool:
    return bool(_GOV_IS_COMPLETE_RE.search(text))

def _gov_is_is_column_overflow(text: str) -> bool:
    lines = (text or "").splitlines()
    joined_top = " ".join(lines[:8])

    # FIX: require "Total ... Governmental ... Funds" as a real column-header
    # phrase (already defined as _GOV_IS_TOTAL_COL_RE), not just incidental
    # co-occurrence of "Total" and "Funds" anywhere in the top 8 lines.
    has_total_col = bool(_GOV_IS_TOTAL_COL_RE.search(joined_top))
    if has_total_col:
        return True

    first_dollar = next((i for i, l in enumerate(lines) if "$" in l), 999)
    if first_dollar < 2:
        return False
    if _STATEMENT_TITLE_RE.search(joined_top):
        return False
    top12 = " ".join(lines[:12])
    if not re.search(r"\bTotals?\b", top12, re.IGNORECASE):
        return False
    has_numbers = any(re.search(r"\b\d[\d,]{2,}\b", l) for l in lines[2:8])
    return has_numbers
    
# ── PROP_SNP ──────────────────────────────────────────────────────────────────
PROP_SNP_SCRAMBLED_KW = [
    re.compile(r"\btotal\s+assets\b",    re.IGNORECASE),
    re.compile(r"\btotal\s+liabilit",    re.IGNORECASE),
    re.compile(r"\bnet\s+position",      re.IGNORECASE),
    re.compile(r"\bproprietary\s+fund",  re.IGNORECASE),
]
_PROP_SNP_COMPLETE_RE = re.compile(
    r"\bTotal\s+(?:fund\s+)?net\s+position\b",
    re.IGNORECASE,
)

def _prop_snp_is_complete(text: str) -> bool:
    return bool(_PROP_SNP_COMPLETE_RE.search(text))

def _looks_like_prop_overflow_page(text: str) -> bool:
    lines = (text or "").splitlines()
    if not lines:
        return False
    top12 = " ".join(lines[:12])
    top25 = " ".join(lines[:25])
    if re.search(
        r"\bnotes?\s+to\b|\brequired\s+supplementary\b|\bsupplementary\b|\bstatistical\b|\bfiduciary\b",
        top12, re.IGNORECASE,
    ):
        return False
    if re.search(
        r"\bStatement\s+of\s+Fiduciary\b"
        r"|\bBalance\s+Sheet\b.{0,60}\bGovernmental\s+Funds\b"
        r"|\bGovernmental\s+Funds\b.{0,60}\bBalance\s+Sheet\b",
        top12, re.IGNORECASE
    ):
        return False
    header_hit = bool(re.search(
        r"\b(Totals?|Business[-\s]*type\s+Activities|Enterprise\s+Funds?|"
        r"Internal\s+Service\s+Funds?|Governmental\s+Activities|Proprietary\s+Funds?|"
        r"Food\s+Service|Nonmajor|Adjustment)\b",
        top25, re.IGNORECASE,
    ))
    if not header_hit:
        return False
    numeric_lines = sum(1 for line in lines[:40] if re.search(r"\(?-?\d[\d,]{2,}\)?", line))
    dash_lines = sum(1 for line in lines[:40] if re.search(r"(^|\s)[\-–—](\s|$)", line))
    return numeric_lines >= 2 or dash_lines >= 3


def _prop_snp_is_column_overflow(text: str) -> bool:
    lines = (text or "").splitlines()
    top12 = " ".join(lines[:12])

    # ── HARD STOP: page is actually PROP_IS or PROP_CFS, not an SNP overflow ──
    if re.search(
        r"\bStatement\s+of\s+Revenues[,\s]+(?:and\s+)?(?:Expenses|Expenditures)\b"
        r"|\bStatement\s+of\s+Revenues\s+and\s+Expenses\b"
        r"|\bRevenues,?\s+Expenses,?\s+and\s+Changes\s+in\s+(?:Fund\s+)?Net\s+Position\b"
        r"|\bStatement\s+of\s+Cash\s+Flows?\b",
        top12, re.IGNORECASE,
    ):
        return False

    # ── Extra guard: PROP_IS body keywords in header ──
    if re.search(
        r"\bOperating\s+Revenues?\b"
        r"|\bOperating\s+Expenses?\b"
        r"|\bCharges\s+for\s+services\b.{0,40}\bTotal\b",
        top12, re.IGNORECASE,
    ):
        return False

    return _looks_like_prop_overflow_page(text)

# ── PROP_IS ───────────────────────────────────────────────────────────────────
PROP_IS_SCRAMBLED_KW = [
    re.compile(r"\boperating\s+revenues?\b", re.IGNORECASE),
    re.compile(r"\boperating\s+expenses?\b", re.IGNORECASE),
    re.compile(r"\bnet\s+position",          re.IGNORECASE),
    re.compile(r"\bproprietary\s+fund",      re.IGNORECASE),
]

_PROP_IS_COMPLETE_RE = re.compile(
    r"\bNet\s+(?:position|assets)\s*[-–]\s*(?:end\w*|begin\w*|of\s+year)\b"
    r"|\bNet\s+(?:position|assets)\s*,?\s*(?:at\s+)?end\s+of\s+(?:fiscal\s+)?year\b"
    r"|\bNet\s+(?:position|assets)\s+at\s+(?:end|beginning)\s+of\s+(?:fiscal\s+)?year\b"
    r"|\bChange\s+in\s+net\s+(?:position|assets)\s+of\s+business",
    re.IGNORECASE,
)



def _prop_is_is_complete(text: str) -> bool:
    """
    Return True if the page contains the closing line of a PROP_IS statement.

    Healthcare statements typically end with one of:
      • "Excess of revenues over expenses" (Inova-style nonprofits)
      • "Increase in net assets"
      • "Change in net position - end of year" (GASB)
      • "Net position at end of year"
    """
    if _PROP_IS_COMPLETE_RE.search(text):
        return True
    # ── NEW: nonprofit-style closers ──
    if re.search(
        r"\bExcess\s+of\s+revenues?\s+over\s+expenses\b"
        r"|\bExcess\s+of\s+expenses?\s+over\s+revenues?\b"
        r"|\bIncrease\s+\(?[Dd]ecrease\)?\s+in\s+net\s+(?:assets|position)\b"
        r"|\bAttributable\s+to\b",
        text or "", re.IGNORECASE,
    ):
        return True
    return False


def _prop_is_is_column_overflow(text: str) -> bool:
    lines = (text or "").splitlines()
    top12 = " ".join(lines[:12])
    if re.search(r"\bStatement\s+of\s+Cash\s+Flows?\b", top12, re.IGNORECASE):
        return False
    return _looks_like_prop_overflow_page(text)

# ── PROP_CFS ──────────────────────────────────────────────────────────────────
PROP_CFS_SCRAMBLED_KW = [
    re.compile(r"\bcash\s+(?:flows?|provided|used)\b", re.IGNORECASE),
    re.compile(r"\boperating\s+activit",               re.IGNORECASE),
    re.compile(r"\bproprietary\s+fund",                re.IGNORECASE),
]


_PROP_CFS_COMPLETE_RE = re.compile(
    # ── Closer 1: "Cash equivalents end-of-year" with optional prefixes ──
    # Matches all of these patterns:
    #   "Cash and cash equivalents, end of year"
    #   "Cash and cash equivalents at end of period"
    #   "Cash and cash equivalents - end of year"
    #   "CURRENT AND RESTRICTED CASH AND CASH\nEQUIVALENTS, END OF YEAR"   ← NEW
    #   "RESTRICTED CASH AND CASH EQUIVALENTS, END OF YEAR"                ← NEW
    #   "Unrestricted Cash and Cash Equivalents at End of Fiscal Year"
    r"\b(?:(?:current|unrestricted|restricted)"
    r"(?:\s+and\s+(?:current|unrestricted|restricted))?\s+)?"
    r"Cash\s+(?:and\s+)?(?:[Cc]ash\s+)?[Ee]quivalents?"
    r"(?:\s*[,\-–—]\s*|\s+(?:at\s+)?)"
    r"end\s+of\s+(?:fiscal\s+)?(?:year|period)\b"

    # ── Closer 2: Compound version "Cash, cash equivalents and restricted cash..." ──
    r"|\bCash\s*,\s*cash\s+equivalents?[^\n]{0,80}end\s+of\s+(?:fiscal\s+)?(?:year|period)\b"

    # ── Closer 3: Non-cash investing/capital/financing ──
    r"|\bNon[-\s]?cash\b[^\n]{0,50}(?:investing|capital|financ)"
    r"|\bNon[-\s]?cash\s+investing"

    # ── Closer 4: Net (increase|decrease) in cash and cash equivalents ──
    r"|\bNet\s+(?:increase|decrease)\s+in\s+cash\s+and\s+cash\s+equivalents?\b"

    # ── Closer 5: Supplemental disclosure/schedule of (non)cash ──
    r"|\bSupplemental\s+(?:schedule|disclosure)\s+of\s+(?:non[-\s]?cash|cash\s+flow)\b"

    # ── Closer 6: Reconciliation of operating income/loss to net cash ──
    r"|\bReconciliation\s+of\s+operating\s+(?:income|loss)\s+to\s+net\s+cash\b",
    re.IGNORECASE,
)



def _prop_cfs_is_complete(text: str) -> bool:
    return bool(_PROP_CFS_COMPLETE_RE.search(text))

def _prop_cfs_is_column_overflow(text: str) -> bool:
    lines = (text or "").splitlines()
    top12 = " ".join(lines[:12])
    if re.search(
        r"\bStatement\s+of\s+Fiduciary\b|\bNotes?\s+to\b|\bRequired\s+Supplementary\b",
        top12, re.IGNORECASE,
    ):
        return False
    return _looks_like_prop_overflow_page(text)

def _is_scrambled_anchor(text: str, kw_list: list) -> bool:
    lines = (text or "").splitlines()
    single_char_lines = sum(1 for l in lines[:20] if len(l.strip()) <= 2)
    spaced_word_lines = sum(1 for l in lines[:20] if re.match(r"^([A-Za-z]\s){2,}", l.strip()))
    if not ((single_char_lines >= 3) or (spaced_word_lines >= 1)):
        return False
    return all(kw.search(text) for kw in kw_list)


# ══════════════════════════════════════════════════════════════════════════════
# SHARED COMPLETION HELPER
# ══════════════════════════════════════════════════════════════════════════════
_NUMERIC_AMOUNT_RE = re.compile(r"[0-9]{1,3}(?:,[0-9]{3})+")

def _complete_row_has_numbers(text: str, complete_re: re.Pattern) -> bool:
    for line in (text or "").splitlines():
        if complete_re.search(line):
            if _NUMERIC_AMOUNT_RE.search(line):
                return True
    return False


# ══════════════════════════════════════════════════════════════════════════════
# FUND-LEVEL STATEMENT DEFINITIONS
# ══════════════════════════════════════════════════════════════════════════════
FUND_STATEMENTS = [
    {
        "suffix": "GOV_BS",
    "title_re": re.compile(
    r"(?:exhibit\s+[a-z]?\s*[-–—]?\s*)?"
    r"(?:balance\s+sheets?\b[^.]{0,120}governmental\s+fund"   # ← requires "governmental fund" within 120 chars
    r"|governmental\s+fund[^.]{0,120}balance\s+sheets?\b)",   # ← dropped bare "balance\s+sheets?\b"
    re.IGNORECASE | re.DOTALL,
),
        "require_kw": [
            re.compile(r"\bassets\b",       re.IGNORECASE),
            re.compile(r"\bfund\s+balance", re.IGNORECASE),
        ],
        "scrambled_kw": GOV_BS_SCRAMBLED_KW,
        "stop_res": [
            (re.compile(r"\breconciliation\b", re.IGNORECASE), 12),
            (re.compile(
                r"statement\s+of\s+revenues,?\s+expenditures.*changes\s+in\s+fund\s+balance",
                re.IGNORECASE | re.DOTALL), 25),
        ],
        "cont_re": re.compile(
            r"(?:balance\s+sheets?\b.{0,60}(?:continued|cont\.?)"
            r"|governmental\s+fund.{0,60}balance\s+sheets?\b.{0,60}(?:continued|cont\.?)"
            r"|\bfund\s+balances?\b|\bliabilit|\btotal\s+assets\b)",
            re.IGNORECASE | re.DOTALL,
        ),
        "scrambled_cont_kw": [
            re.compile(r"\bfund\s+balance",              re.IGNORECASE),
            re.compile(r"\btotal\s+(?:assets|liabilit)", re.IGNORECASE),
        ],
        "is_complete_fn":     _gov_bs_is_complete,
        "is_col_overflow_fn": _gov_bs_is_column_overflow,
        "gov_overflow":       True,
    },
    {
    "suffix": "GOV_IS",
    "title_re": re.compile(
        r"(?:exhibit\s+[a-z]?\s*[-–—]?\s*)?"  # Optional "Exhibit E –"
        r"(?:statement\s+of\s+revenues,?\s+expenditures"
        r"|revenues,?\s+expenditures,?\s+and\s+changes\s+in\s+fund\s+balance)"
        r"(?:.{0,200}governmental\s+fund)?"  # ← CHANGED: Make "governmental fund" optional
        r"|governmental\s+fund.{0,200}(?:statement\s+of\s+revenues,?\s+expenditures"
        r"|revenues,?\s+expenditures,?\s+and\s+changes\s+in\s+fund\s+balance)",
        re.IGNORECASE | re.DOTALL,
    ),
        "require_kw": [
            re.compile(r"\bexpenditures\b", re.IGNORECASE),
            re.compile(r"\brevenues\b",     re.IGNORECASE),
        ],
        "scrambled_kw": GOV_IS_SCRAMBLED_KW,
        "stop_res": [
            (re.compile(r"\breconciliation\b", re.IGNORECASE), 12),
            (re.compile(
                r"statement\s+of\s+net\s+position.{0,60}proprietary"
                r"|proprietary.{0,60}statement\s+of\s+net\s+position",
                re.IGNORECASE | re.DOTALL), 25),
        ],
        "cont_re": re.compile(
            r"(?:revenues,?\s+expenditures.*changes\s+in\s+fund\s+balance.*(?:continued|cont\.?)"
            r"|governmental\s+fund.*(?:continued|cont\.?)"
            r"|\bexpenditures\b|\bfund\s+balance\b)",
            re.IGNORECASE | re.DOTALL,
        ),
        "scrambled_cont_kw": [
            re.compile(r"\bfund\s+balance", re.IGNORECASE),
            re.compile(r"\bexpenditures\b", re.IGNORECASE),
        ],
        "is_complete_fn":     _gov_is_is_complete,
        "is_col_overflow_fn": _gov_is_is_column_overflow,
        "gov_overflow":       True,
    },
    {
        "suffix": "PROP_SNP",
        "title_re": re.compile(
    r"(?:exhibit\s+[a-z0-9]+\s*[-–—]?\s*)?"
    r"(?:"
    r"statements?\s+of\s+(?:fund\s+)?net\s+(?:pos(?:i)?tion|assets)[^.]{0,60}"
    r"\b(?:proprietary|enterprise|internal\s+service)\b"
    r"|"
    r"\b(?:proprietary|enterprise|internal\s+service)\s+funds?\b[^.]{0,60}"
    r"statements?\s+of\s+(?:fund\s+)?net\s+(?:pos(?:i)?tion|assets)"
    r")",
    re.IGNORECASE | re.DOTALL,
),
"require_kw": [
    re.compile(r"\bassets\b", re.IGNORECASE),
    re.compile(r"\bnet\s+(?:position|assets)", re.IGNORECASE),  # was: r"\bnet\s+position"
],
        "scrambled_kw": [
    re.compile(r"\btotal\s+assets\b", re.IGNORECASE),
    re.compile(r"\btotal\s+liabilit", re.IGNORECASE),
    re.compile(r"\bnet\s+(?:position|assets)", re.IGNORECASE),  # was: r"\bnet\s+position"
    re.compile(r"\bproprietary\s+fund", re.IGNORECASE),
],
        "stop_res": [
            (re.compile(
                r"statement\s+of\s+revenues,?\s+expenses.*changes.*net\s+position.*proprietary"
                r"|proprietary.*statement\s+of\s+revenues,?\s+expenses.*changes",
                re.IGNORECASE | re.DOTALL), 25),
            (re.compile(r"\bstatement\s+of\s+cash\s+flow", re.IGNORECASE), 25),
        ],
        "cont_re": re.compile(
            r"(?:statement\s+of\s+(?:fund\s+)?net\s+position.*(?:continued|cont\.?)"
            r"|proprietary\s+fund.*(?:continued|cont\.?)"
            r"|\bnet\s+position\b|\bliabilit|\btotal\s+assets\b)",
            re.IGNORECASE | re.DOTALL,
        ),
        "scrambled_cont_kw": [
            re.compile(r"\bnet\s+position",              re.IGNORECASE),
            re.compile(r"\btotal\s+(?:assets|liabilit)", re.IGNORECASE),
        ],
        "is_complete_fn":     _prop_snp_is_complete,
        "is_col_overflow_fn": _prop_snp_is_column_overflow,
        "complete_re":        _PROP_IS_COMPLETE_RE,
    },
    {
        "suffix": "PROP_IS",
        "title_re": re.compile(
    r"(?:exhibit\s+[a-z0-9]+\s*[-–—]?\s*)?"
    r"(?:"
    r"statements?\s+of\s+revenues?,?\s+(?:and\s+)?(?:expens(?:es|e)|expenditures),?\s+and\s+changes\s+in\s+(?:fund\s+)?net\s+(?:pos(?:i)?tion|assets)[^.]{0,60}"
    r"\b(?:proprietary|enterprise|internal\s+service)\b"
    r"|"
    r"\b(?:proprietary|enterprise|internal\s+service)\s+funds?\b[^.]{0,60}"
    r"statements?\s+of\s+revenues?,?\s+(?:and\s+)?(?:expens(?:es|e)|expenditures),?\s+and\s+changes\s+in\s+(?:fund\s+)?net\s+(?:pos(?:i)?tion|assets)"
    r")",
    re.IGNORECASE | re.DOTALL,
),
"require_kw": [
    re.compile(r"\bexpenses\b", re.IGNORECASE),
    re.compile(r"\bnet\s+(?:position|assets)", re.IGNORECASE),  # was: r"\bnet\s+position"
],
        "scrambled_kw": PROP_IS_SCRAMBLED_KW,
        "stop_res": [
            (re.compile(r"\bstatement\s+of\s+cash\s+flow", re.IGNORECASE), 25),
        ],
        "cont_re": re.compile(
            r"(?:revenues,?\s+(?:expenses|expenditures).*changes.*net\s+position.*(?:continued|cont\.?)"
            r"|proprietary\s+fund.*(?:continued|cont\.?)"
            r"|\bnet\s+position\b|\boperating\s+(?:revenue|expense))",
            re.IGNORECASE | re.DOTALL,
        ),
        "scrambled_cont_kw": [
            re.compile(r"\bnet\s+position",                  re.IGNORECASE),
            re.compile(r"\boperating\s+(?:revenue|expense)", re.IGNORECASE),
        ],
        "is_complete_fn":     _prop_is_is_complete,
        "is_col_overflow_fn": _prop_is_is_column_overflow,
        "complete_re":        _PROP_IS_COMPLETE_RE,
        
    

    },
    {
        "suffix": "PROP_CFS",
        "title_re": re.compile(
    r"(?:exhibit\s+[a-z0-9]+\s*[-–—]?\s*)?"
    r"(?:"
    r"statements?\s+of\s+cash\s+flows?[^.]{0,60}"
    r"\b(?:proprietary|enterprise|internal\s+service)\b"
    r"|"
    r"\b(?:proprietary|enterprise|internal\s+service)\s+funds?\b[^.]{0,60}"
    r"statements?\s+of\s+cash\s+flows?"
    r")",
    re.IGNORECASE | re.DOTALL,
),
        "require_kw": [
            re.compile(r"\bcash\b",              re.IGNORECASE),
            re.compile(r"\boperating\s+activit", re.IGNORECASE),
        ],
        "scrambled_kw": PROP_CFS_SCRAMBLED_KW,
        "stop_res": [
            (re.compile(r"\bnotes?\s+to\b",  re.IGNORECASE), 12),
            (re.compile(r"\bfiduciary\b",    re.IGNORECASE), 12),
            (re.compile(
                r"\bbalance\s+sheet\b.{0,40}\bgovernmental\b",
                re.IGNORECASE | re.DOTALL), 25),
        ],
        "cont_re": re.compile(
            r"(?:statement\s+of\s+cash\s+flow.*(?:continued|cont\.?)"
            r"|proprietary\s+fund.*(?:continued|cont\.?)"
            r"|\bcash\s+(?:provided|used|flow)|\bnoncash\b"
            r"|\breconciliation\s+of\s+(?:operating|net))",
            re.IGNORECASE | re.DOTALL,
        ),
        "scrambled_cont_kw": [
            re.compile(r"\bcash\s+(?:provided|used|flows?)\b", re.IGNORECASE),
            re.compile(r"\boperating\s+activit",               re.IGNORECASE),
        ],
        "is_complete_fn":     _prop_cfs_is_complete,
        "is_col_overflow_fn": _prop_cfs_is_column_overflow,
        "complete_re":        _PROP_CFS_COMPLETE_RE,
    },
]

# ══════════════════════════════════════════════════════════════════════════════
# NON-LG FUND STATEMENT DEFINITIONS
# ══════════════════════════════════════════════════════════════════════════════
#
# NON-LG files (enterprise funds, utilities, hospitals, universities, etc.)
# have the SAME 3 proprietary-style statements but WITHOUT "Proprietary Fund"
# prefix in the title.  e.g.:
#   "Statement of Net Position"  (not "Proprietary Fund Statement of ...")
#   "Statement of Revenues, Expenses, and Changes in Net Position"
#   "Statement of Cash Flows"
# ──────────────────────────────────────────────────────────────────────────────

NONLG_SNP_EXCLUDE_PATTERNS = [
    re.compile(r"\btable\s+of\s+contents?\b",   re.IGNORECASE),
    re.compile(r"\bforeword\b",                 re.IGNORECASE),
    re.compile(r"\bintroduction\b",             re.IGNORECASE),
    re.compile(r"\bmd\s*(?:&|and)\s*a\b",       re.IGNORECASE),
    re.compile(r"\bmanagement[\u2018\u2019']?s?\s+discussion", re.IGNORECASE),
    re.compile(r"\bbudget\b",                   re.IGNORECASE),
    re.compile(r"\bstatistical\b",              re.IGNORECASE),
    re.compile(r"\bcombining\b",                re.IGNORECASE),
    re.compile(r"\bsupplemental\b",             re.IGNORECASE),
    re.compile(r"\bnotes?\s+to\s+(?:the\s+)?(?:financial|required)", re.IGNORECASE),
    re.compile(r"\bfiduciary\b",                re.IGNORECASE),
    re.compile(r"\brequired\s+supplementary\b", re.IGNORECASE),
    re.compile(r"\bcondensed\b",                re.IGNORECASE),
]


# ── TOC detection: lines like "Consolidated Balance Sheets ..........  4" ──
_TOC_LINE_RE = re.compile(
    r"^(?!.*\$)"                        # no $ sign anywhere on line
    r"(?!.*\b\d{1,3},\d{3}\b)"         # no comma-grouped number
    r"(?!.*\b(?:liability|payable|assets?|liabilities|"
    r"position|inflows?|outflows?|maturities|absences)\b)"  # no financial keywords
    r".{5,}\.{3,}\s*\d{1,3}"
    r"(?:\s*[-–—]\s*\d{1,3})?\s*$",
    re.MULTILINE | re.IGNORECASE,
)
# Phrases that, when found near the page top, strongly suggest a TOC page
_TOC_HEADER_RE = re.compile(
    r"\b(?:table\s+of\s+contents?"
    r"|audited\s+consolidated\s+financial\s+statements"
    r"|consolidated\s+financial\s+statements"
    r"|index\s+to\s+financial\s+statements)\b",
    re.IGNORECASE,
)


def _is_toc_page(text: str) -> bool:
    if not text:
        return False
    toc_lines = len(_TOC_LINE_RE.findall(text))
    if toc_lines >= 3:
        # ── Safety override: if page has a statement title + financial data,
        #    it's NOT a TOC even if dot-leaders fooled us ──
        has_stmt_title = bool(re.search(
            r"\bstatement\s+of\s+(?:net\s+position|activities|revenues|cash\s+flows?)\b"
            r"|\bbalance\s+sheet\b",
            text, re.IGNORECASE,
        ))
        has_financial_data = bool(re.search(r"\d{1,3}(?:,\d{3}){2,}", text))  # 7+ digit amounts
        if has_stmt_title and has_financial_data:
            return False
        return True
    first15 = "\n".join(text.splitlines()[:15])
    if toc_lines >= 2 and _TOC_HEADER_RE.search(first15):
        return True
    return False

def _has_financial_data(text: str, min_amounts: int = 8) -> bool:
    """
    A real balance sheet / income statement has many formatted dollar
    amounts (e.g. "1,723,433", "$ 557,679"). TOC pages do not.
    """
    amounts = re.findall(r"\b\d{1,3}(?:,\d{3}){1,}\b", text or "")
    return len(amounts) >= min_amounts


# ── MD&A header detection — apostrophe-independent ──
# Catches "Management's Discussion and Analysis", "Management Discussion and
# Analysis", any apostrophe variant (', ', `, ‛, etc.), and also catches
# "‐ Continued" / "- Continued" headers used on MD&A continuation pages.
_MDA_HEADER_RE = re.compile(
    r"\b(?:management['\u2018\u2019\u02BC\u2032`]?s?\s+)?"
    r"discussion\s+(?:and|&)\s+analysis\b",
    re.IGNORECASE,
)
_MDA_CONTINUED_RE = re.compile(
    r"\bdiscussion\s+(?:and|&)\s+analysis\s*[\u2010\u2011\u2012\u2013\u2014\-]\s*continued\b",
    re.IGNORECASE,
)

# ── Narrative-summary phrases that appear in MD&A / Notes, NEVER in a
#    real Statement of Net Position / Financial Position page header. ──
_NARRATIVE_SUMMARY_RE = re.compile(
    r"\b(?:summarized\s+as\s+follows"
    r"|(?:are|is)\s+summarized\s+as"
    r"|(?:are|is)\s+as\s+follows"
    r"|presented\s+as\s+follows"
    r"|shown\s+(?:as\s+follows|below)"
    r"|reported\s+as\s+follows"
    r"|the\s+following\s+table"
    r"|the\s+following\s+(?:summary|schedule)"
    r"|condensed\s+(?:comparative\s+)?(?:summary|statements?)"
    r"|as\s+of\s+the\s+end\s+of\s+each\s+of"
    r"|for\s+(?:each\s+of\s+)?the\s+(?:last|past)\s+three\s+(?:fiscal\s+)?years)"
    r"\b",
    re.IGNORECASE,
)


def _is_narrative_summary_page(text: str) -> bool:
    """
    Detect MD&A / Notes pages that QUOTE a statement title inside narrative
    prose ("...statements of net position as of the end of each of the last
    three years are summarized as follows: ...") followed by a condensed
    3-year summary table. Real statement pages NEVER contain these phrases
    in their first 20 lines.
    """
    if not text:
        return False
    first20 = "\n".join(text.splitlines()[:20])
    return bool(_NARRATIVE_SUMMARY_RE.search(first20))




_FOOTER_NOTES_REF_RE = re.compile(
    r"\bSee\s+(?:accompanying\s+)?notes?\s+to\s+(?:the\s+)?(?:financial|required)\b"
    r"|\bThe\s+accompanying\s+notes?\s+are\s+an\s+integral\s+part\b"
    r"|\bNotes?\s+are\s+an\s+integral\s+part\b",      # ← add this catch-all
    re.IGNORECASE,
)
_pypdf_reader_cache = {}

_NUMERIC_AMOUNT_RE_FALLBACK = re.compile(r"\d{1,3}(?:,\d{3})+")
def _get_cached_pdf_reader(pdf_path: str) -> PdfReader:
    """Get or create cached PdfReader."""
    if pdf_path not in _pypdf_reader_cache:
        _pypdf_reader_cache[pdf_path] = PdfReader(pdf_path)
    return _pypdf_reader_cache[pdf_path]
 
def _line_is_word_shredded(line: str) -> bool:
    """
    True if a line looks like a normal word/phrase that's been shredded
    into mostly single-or-double-character tokens by pdfplumber's
    clustering bug, e.g.:
        "u r d u e U n i v e r s i t y"      -> all 1-char tokens
        "C u r r e n t A s s e ts :"        -> mostly 1-char, one "ts" pair
        "ia b ilitie s a n d D e fe r r"    -> mixed cluster sizes
 
    Unlike a strict alternating-regex, this tolerates pdfplumber
    occasionally clustering 2-3 adjacent glyphs together (the clustering
    is inconsistent even within one shredded line), by requiring most
    tokens be short rather than requiring a perfectly uniform pattern.
    """
    tokens = line.strip().split()
    if len(tokens) < 4:
        return False
    short_tokens = sum(1 for t in tokens if len(t.strip(":.,()")) <= 2)
    return (short_tokens / len(tokens)) >= 0.6
 
 
def _is_char_per_line_garbled(text: str) -> bool:
    """
    Detect pdfplumber's character-per-LINE garbling pattern.
 
    Two related sub-patterns both count as "garbled lines":
      (a) Lines that are themselves <= 2 characters long (isolated single
          letters/digits, e.g. "P", "S", "A").
      (b) Lines where most of the line's whitespace-separated tokens are
          1-2 characters long, i.e. a word/phrase shredded into glyph
          clusters by pdfplumber (e.g. "C u r r e n t A s s e ts :"),
          detected via _line_is_word_shredded().
 
    Both (a) and (b) show up together on the same garbled pages (this
    PDF's title/header text alternates between isolated single letters
    and shredded word fragments line by line), so a real garbled page
    has a high combined ratio of (a)+(b) over the first ~30 lines.
 
    The numeric-blindness check is the decisive second signal: a real
    financial-statement page has a "$" somewhere, but pdfplumber's
    garbling here also shreds the digit groupings, so no comma-grouped
    numeric token (e.g. "431,418") survives intact. Requiring both the
    line-shape signal AND the numeric-blindness signal avoids false
    positives on pages that are merely short-labeled but not garbled.
    """
    if not text:
        return False
 
    lines = text.splitlines()
    first30 = lines[:30]
    if not first30:
        return False
 
    garbled_lines = sum(
        1 for l in first30
        if len(l.strip()) <= 2 or _line_is_word_shredded(l)
    )
    garbled_ratio = garbled_lines / len(first30)
 
    has_dollar = "$" in text
    has_grouped_numbers = bool(_NUMERIC_AMOUNT_RE_FALLBACK.search(text))
 
    # Garbled signature: lots of single-char-or-shredded-word lines AND
    # (page clearly should have financial numbers, given a $ sign,
    #  but pdfplumber extracted none in normal comma-grouped form)
    return garbled_ratio >= 0.3 and has_dollar and not has_grouped_numbers
 
 
# ══════════════════════════════════════════════════════════════════════════
# 2. ROBUST TEXT GETTER: pdfplumber first, pypdf fallback when garbled
# ══════════════════════════════════════════════════════════════════════════
 
def _get_pypdf_reader(pdf_path: str) -> PdfReader:
    """Cache the PdfReader per file path to avoid re-parsing the PDF
    on every page lookup (find_fund_statement_pages scans many pages)."""
    reader = _pypdf_reader_cache.get(pdf_path)
    if reader is None:
        reader = PdfReader(pdf_path)
        _pypdf_reader_cache[pdf_path] = reader
    return reader
 
def _get_robust_page_text(pdf, pdf_path: str, page_index: int) -> str:
    """
    Return the best available text for a page (0-indexed), preferring
    pdfplumber's output but falling back to pypdf when pdfplumber's
    output shows the char-per-line garbling signature.
 
    `pdf` is an already-open pdfplumber.PDF object (so callers already
    inside a `with pdfplumber.open(pdf_path) as pdf:` block don't pay to
    reopen it). `pdf_path` is needed to access/cache the pypdf reader.
    """
    pp_text = pdf.pages[page_index].extract_text() or ""
 
    if _is_char_per_line_garbled(pp_text):
        try:
            reader = _get_pypdf_reader(pdf_path)
            py_text = reader.pages[page_index].extract_text() or ""
            # Only switch if pypdf actually produced something with real
            # numbers -- otherwise keep the (still imperfect) pdfplumber
            # text rather than silently returning blank/garbage from pypdf.
            if py_text.strip() and _NUMERIC_AMOUNT_RE_FALLBACK.search(py_text):
                return py_text
        except Exception:
            pass
 
    return pp_text


def _nonlg_excluded(text: str) -> bool:
    """
    Exclude NON-LG false positives:
      TOC, MD&A (including ‐ Continued pages), condensed summaries,
      narrative-embedded titles, Notes, RSI, etc.

    BUG FIXES applied here:
    -----------------------
    1. (Earlier) Strip "See notes to the financial statements." footer
       so short SNP pages aren't wrongly excluded.
    2. (NEW) Apostrophe-independent MD&A detection — catches U+0027,
       U+2018, U+2019, U+02BC, U+2032, backtick, and also handles the
       "Discussion and Analysis ‐ Continued" headers on pages 2-7
       of HC_KS_900001170_2024.pdf etc.
    3. (NEW) Reject narrative-summary pages where the statement title
       appears inside a sentence followed by a condensed 3-year table
       (typical MD&A pattern).
    """
    if not text:
        return True

    # ── 1. Strip the "See notes to the financial statements." footer ──
    clean_text = "\n".join(
        l for l in text.splitlines() if not _FOOTER_NOTES_REF_RE.search(l)
    )

    lines   = clean_text.splitlines()
    first12 = "\n".join(lines[:12])
    first30 = "\n".join(lines[:30])

    # ── 2. TOC detection ──
    if _is_toc_page(text):
        return True

    # ── 3. MD&A detection — fires on ANY page that carries an MD&A
    #       header or its "‐ Continued" variant in the top 15 lines. ──
    first15 = "\n".join(lines[:15])
    if _MDA_HEADER_RE.search(first15) or _MDA_CONTINUED_RE.search(first15):
        return True

    # ── 4. Narrative-summary page rejection (MD&A condensed tables) ──
    if _is_narrative_summary_page(text):
        return True

    # ── 5. Other hard excludes (first 30 lines) ──
    hard_excludes_30 = [
    r"\btable\s+of\s+contents?\b",
    r"\bforeword\b",
    r"\bintroduction\b",
    r"\bstatistical\b",
    r"\bbudget\b",
    r"\bcombining\b",
    r"\bsupplemental\b(?!\s+(?:disclosures?|schedule\s+of\s+(?:non)?cash|non[-\s]?cash|cash\s+flow|payment\b))",
    r"\brequired\s+supplementary\b",
    r"\bfiduciary\s+(?:fund|net\s+position|trust|component)\b",
]
    if any(re.search(p, first30, re.IGNORECASE) for p in hard_excludes_30):
        return True

    # ── 6. "Notes to financial statements" as a PAGE HEADER (first 12) ──

    # Allow optional descriptors ("Consolidated", "Combined", etc.) between
    # "Notes to" and "Financial". Catches Mayo Clinic style:
    #   "Notes to Consolidated Financial Statements"
    #   "Notes to Combined Financial Statements"
    #   "Notes to the Consolidated Financial Statements"
    if re.search(
        r"\bnotes?\s+to\s+(?:the\s+)?"
        r"(?:consolidated\s+|combined\s+|comparative\s+|audited\s+|interim\s+)?"
        r"(?:financial|required)\b",
        first12, re.IGNORECASE,
    ):
        return True


    # ── 7. Must have enough numeric data to qualify as a real statement ──
    if not _has_financial_data(text, min_amounts=8):
        return True

    return False


NONLG_FUND_STATEMENTS = [
    # ── NON-LG: Statement of Net Position / Financial Position ────────────
    
    {
        "suffix": "PROP_SNP",
"title_re": re.compile(
    r"(?<![A-Za-z])"
    r"(?:statements?\s+of\s+"
    r"(?:(?:fund\s+)?net\s+(?:pos(?:i)?tion|assets)"
    r"|financial\s+position)"
    r"|(?:consolidated\s+|combined\s+)?balance\s+sheets?)"
    r"\b(?!\s+(?:as\s+of\s+the\s+end\s+of|are\s+summarized|is\s+summarized"
    r"|are\s+as\s+follows|is\s+as\s+follows|presented\s+as"
    r"|provides?|presents?|reflects?|reports?|is\s+the\s+"
    r"(?:(?:consolidated\s+|combined\s+|combining\s+)?balance\s+sheets?"
    r"|distinguishes?|provide\s+information)))",   # ← was "))", now ")))"
    re.IGNORECASE,
),
        "require_kw": [
            re.compile(r"\b(?:total\s+)?assets\b", re.IGNORECASE),
            re.compile(
                r"\b(?:net\s+position|net\s+assets|liabilit|"
                r"total\s+liabilities|fund\s+balances?)\b",
                re.IGNORECASE,
            ),
        ],
        "scrambled_kw": PROP_SNP_SCRAMBLED_KW,
            "stop_res": [
                (re.compile(
                    r"statements?\s+of\s+revenues,?\s+(?:expenses|expenditures)"
                    r"|statements?\s+of\s+activities",
                    re.IGNORECASE | re.DOTALL), 25),
                (re.compile(r"\bstatements?\s+of\s+cash\s+flows?", re.IGNORECASE), 25),
            ],


            "cont_re": re.compile(
                r"(?:statement\s+of\s+(?:fund\s+)?net\s+position.*(?:continued|cont.?)"
                r"|statement\s+of\s+financial\s+position.*(?:continued|cont.?)"
                r"|\bnet\s+position\b|\bliabilit|\btotal\s+assets\b"
                r"|\bdeferred\s+(?:inflows?|outflows?)\s+of\s+resources\b)",
                re.IGNORECASE | re.DOTALL,
            ),
            "scrambled_cont_kw": [
                re.compile(r"\bnet\s+position",              re.IGNORECASE),
                re.compile(r"\btotal\s+(?:assets|liabilit)", re.IGNORECASE),
            ],
            "is_complete_fn":     _prop_snp_is_complete,
            "is_col_overflow_fn": _prop_snp_is_column_overflow,
            "complete_re":        _PROP_SNP_COMPLETE_RE,
            "exclude_fn":         _nonlg_excluded,    # NON-LG specific exclusion
        },

    # ── NON-LG: Statement of Revenues, Expenses & Changes in Net Position ─
    {
        "suffix": "PROP_IS",
        "title_re": re.compile(
    r"(?:exhibit\s+[a-z]?\s*[-–—]?\s*)?"
    r"(?:"
    r"statements?\s+of\s+revenues?\s*(?:,|\s+and)\s+(?:expenses|expenditures)"
    r"(?:,?\s+and\s+changes?\s+in\s+(?:fund\s+)?net\s+(?:position|assets))?"
    r"|(?:consolidated\s+|combined\s+)?statements?\s+of\s+operations"
    r"|(?:consolidated\s+|combined\s+)?statements?\s+of\s+activities"
    r"|statements?\s+of\s+activities"
    r"(?:\s+and\s+changes?\s+in\s+net\s+assets)?"
    # ── NEW: FRS 102 / UK GAAP / HE sector style ──
    r"|(?:consolidated\s+(?:and\s+\w+\s+)?)?statements?\s+of\s+comprehensive\s+income"
    r"|(?:consolidated\s+(?:and\s+\w+\s+)?)?income\s+and\s+expenditure\s+(?:account|statement)"
    r"|statements?\s+of\s+financial\s+activities"    # charities
    r")",
    re.IGNORECASE | re.DOTALL,
),
        "require_kw": [
            re.compile(r"\b(?:expenses|expenditures)\b", re.IGNORECASE),
            re.compile(r"\b(?:net\s+position|net\s+assets|revenue)", re.IGNORECASE),
        ],
        "scrambled_kw": PROP_IS_SCRAMBLED_KW,

        "stop_res": [
            (re.compile(r"\bstatements?\s+of\s+cash\s+flows?", re.IGNORECASE), 25),
            (re.compile(r"\bstatements?\s+of\s+functional\s+expenses", re.IGNORECASE), 25),
            (re.compile(r"\bnotes?\s+to\s+(?:the\s+)?(?:financial|required)", re.IGNORECASE), 12),
        ],

            "cont_re": re.compile(
                # Continuation ONLY when the page explicitly says so —
                # not just because it mentions "net assets" somewhere
                r"(?:statements?\s+of\s+revenues?,?\s+(?:expenses|expenditures)"
                r".*(?:continued|cont\.?|cont'd)"
                r"|statements?\s+of\s+operations.*(?:continued|cont\.?|cont'd)"
                r"|\(continued\)"
                r"|\bcontinued\s+from\s+(?:previous|prior)\s+page\b"
                r"|\bcont(?:inued|\.|'d)\b)",
                re.IGNORECASE | re.DOTALL,
            ),

        
                "scrambled_cont_kw": [
                    re.compile(r"\bnet\s+(?:position|assets)",         re.IGNORECASE),
                    re.compile(r"\boperating\s+(?:revenue|expense)",   re.IGNORECASE),
                ],

        
        "stop_res": [
            (re.compile(r"\bstatements?\s+of\s+cash\s+flows?", re.IGNORECASE), 25),
            (re.compile(r"\bstatements?\s+of\s+functional\s+expenses", re.IGNORECASE), 25),
            (re.compile(r"\bnotes?\s+to\s+(?:the\s+)?(?:financial|required)", re.IGNORECASE), 12),
            # ── NEW: stop when next page is a different statement ──
            (re.compile(
                r"(?:consolidated\s+|combined\s+)?statements?\s+of\s+changes\s+in\s+net\s+(?:assets|position)",
                re.IGNORECASE), 25),
            (re.compile(r"(?:consolidated\s+|combined\s+)?balance\s+sheets?\b", re.IGNORECASE), 25),
            (re.compile(r"\bstatements?\s+of\s+(?:fund\s+)?net\s+(?:position|assets)\b", re.IGNORECASE), 25),
        ],

        "is_complete_fn":     _prop_is_is_complete,
        "is_col_overflow_fn": _prop_is_is_column_overflow,
        "complete_re":        _PROP_IS_COMPLETE_RE,
        "exclude_fn":         _nonlg_excluded,
    },


    # ── NON-LG: Statement of Cash Flows ───────────────────────────────────
    {
        "suffix": "PROP_CFS",
        "title_re": re.compile(
            r"(?:exhibit\s+[a-z0-9]+\s*[-–—]?\s*)?"
            r"(?:consolidated\s+|combined\s+)?"
            r"statements?\s+of\s+cash\s+flows?",
            re.IGNORECASE | re.DOTALL,
        ),
        "require_kw": [
            re.compile(r"\bcash\b",              re.IGNORECASE),
            re.compile(r"\boperating\s+activit", re.IGNORECASE),
        ],
        "scrambled_kw": PROP_CFS_SCRAMBLED_KW,
        "stop_res": [
            # ── NEW: Stop when next page starts with "Reconciliation of Operating..." ──
            # Page 43 ends the main CFS; page 44 starts with this header — exclude it
            # ── existing patterns below ──
            (re.compile(r"\bnotes?\s+to\s+(?:the\s+)?(?:financial|required)", re.IGNORECASE), 12),
            (re.compile(r"\bfiduciary\b", re.IGNORECASE), 12),
            (re.compile(r"\brequired\s+supplementary\b", re.IGNORECASE), 12),

            # Stop when next page is a DIFFERENT statement
            (re.compile(
                r"(?:consolidated\s+|combined\s+|nongovernmental\s+)?"
                r"(?:discretely\s+presented\s+component\s+units?[^\n]{0,80})?"
                r"(?:balance\s+sheets?\b"
                r"|statements?\s+of\s+(?:fund\s+)?net\s+position)",
                re.IGNORECASE), 15),
            (re.compile(
                r"(?:consolidated\s+|combined\s+|nongovernmental\s+)?"
                r"(?:discretely\s+presented\s+component\s+units?[^\n]{0,80})?"
                r"statements?\s+of\s+(?:operations|activities|revenues?,?\s+expenses)",
                re.IGNORECASE), 15),
            (re.compile(
                r"statements?\s+of\s+changes\s+in\s+net\s+(?:assets|position)",
                re.IGNORECASE), 15),
            (re.compile(r"\bSchedule\s+\d+\b", re.IGNORECASE), 5),
            (re.compile(
    r"\bReconciliation\s+of\s+(?:net\s+)?operating\s+"
    r"(?:income|loss|revenue|revenues|expenses)"
    r"(?:\s*\([^)]+\))?\s+to\s+net\s+cash\b",
    re.IGNORECASE), 9),
        ],


        # ── TIGHTENED: only treat as continuation if explicit marker OR
        #    body still shows cash-flow activity sections ──
        "cont_re": re.compile(
            r"(?:statements?\s+of\s+cash\s+flows?.*(?:continued|cont\.?|cont'd)"
            r"|\(continued\)"
            r"|\bcash\s+flows?\s+from\s+(?:operating|investing|financing|noncapital|capital)\s+activit"
            r"|\bnet\s+cash\s+(?:provided|used)\s+by\b"
            r"|\breconciliation\s+of\s+operating\s+(?:income|loss)\s+to\s+net\s+cash\b"
            r"|\bnon[-\s]?cash\s+(?:investing|capital|financing)\b"
            r"|\bsupplemental\s+(?:schedule|disclosure)\s+of\s+(?:non)?cash\b)",
            re.IGNORECASE | re.DOTALL,
        ),
        "scrambled_cont_kw": [
            re.compile(r"\bcash\s+(?:provided|used|flows?)\b", re.IGNORECASE),
            re.compile(r"\boperating\s+activit",               re.IGNORECASE),
        ],
        "is_complete_fn":     _prop_cfs_is_complete,
        "is_col_overflow_fn": _prop_cfs_is_column_overflow,
        "complete_re":        _PROP_CFS_COMPLETE_RE,
        "exclude_fn":         _nonlg_excluded,
    },

]



def _fund_excluded(text: str, defn: dict = None) -> bool:
    if defn and "exclude_fn" in defn:
        return defn["exclude_fn"](text)
    lines = (text or "").splitlines()
    single_char = sum(1 for l in lines[:20] if len(l.strip()) <= 2)
    spaced_word = sum(1 for l in lines[:20] if re.match(r"^([A-Za-z]\s){2,}", l.strip()))
    if (single_char >= 3) or (spaced_word >= 1):
        return False

    if _RECONCILIATION_BODY_RE.search(text or ""):
        return True

    # FIX: strip footer before pattern matching
    clean_lines = [l for l in lines if not _FUND_FOOTER_NOTES_REF_RE.search(l)]
    
    # ── WIDENED from 12 → 20 lines ──
    first20 = "\n".join(clean_lines[:20])
    
    for p in EXCLUDE_PATTERNS:
        if p.search(first20):
            if p.pattern == r"\breconciliation\b" and defn is not None:
                title_re = defn.get("title_re")
                req_kw   = defn.get("require_kw", [])
                if title_re and _title_in_header(text, title_re):
                    if all(kw.search(text or "") for kw in req_kw):
                        continue
            return True
    return False



def _title_in_header(text: str, title_re: re.Pattern, n_lines: int = 25) -> bool:
    if not text:
        return False
    all_lines = text.splitlines()

    # Head check (original)
    head = all_lines[:n_lines]
    for line in head:
        if title_re.search(line.strip()):
            return True
    joined_head = " ".join(l.strip() for l in head if l.strip())
    if title_re.search(joined_head):
        return True

    # NEW: also check tail — rotated title banners extract AFTER the table body
    tail = all_lines[-n_lines:]
    for line in tail:
        if title_re.search(line.strip()):
            return True
    joined_tail = " ".join(l.strip() for l in tail if l.strip())
    if title_re.search(joined_tail):
        return True

    return False

def _collapse_spaced_text(text: str) -> str:
    out_lines = []
    for line in text.splitlines():
        tokens = line.split()
        if tokens and len(tokens) >= 4 and all(len(t) <= 2 for t in tokens):
            out_lines.append("".join(tokens))
        else:
            out_lines.append(line)
    return "\n".join(out_lines)

# ══════════════════════════════════════════════════════════════════════════════
# PROP_SNP REVERSE-LOOKBACK FALLBACK (NON-LG only)
# ══════════════════════════════════════════════════════════════════════════════
# When NON-LG SNP title detection fails (e.g., title is fused to entity name
# like "Purdue UniversityStatement of Net Position", or sits below a page
# number/figure block that defeats the regex), but PROP_IS was successfully
# detected, scan BACKWARD up to N pages from PROP_IS's first page looking for
# a Statement of Net Position by body content + require_kw + completion anchor.
#
# Mirrors find_prop_is_sandwich() but in the OPPOSITE direction.
# ──────────────────────────────────────────────────────────────────────────────

_PROP_SNP_BODY_TRIO = [
    re.compile(r"\bTotal\s+(?:current\s+)?assets\b",                    re.IGNORECASE),
    re.compile(r"\bTotal\s+(?:current\s+)?liabilit",                    re.IGNORECASE),
    re.compile(r"\b(?:Total\s+)?net\s+(?:position|assets)\b",           re.IGNORECASE),
]


_PROP_CFS_BODY_TRIO = [
    re.compile(r"\bCash\s+flows?\s+from\s+operating\s+activit", re.IGNORECASE),
    re.compile(r"\bNet\s+(?:cash|change)\s+(?:used|provided|in\s+cash|from)", re.IGNORECASE),
    # ── Relaxed: allow ", and restricted cash" / extra words between
    #     "equivalents" and "end/beginning of (the) (fiscal) year/period"
    re.compile(
        r"\bCash[,\s]+(?:and\s+)?(?:cash\s+)?equivalents?\b"
        r"(?:[^\n]{0,60}?)"
        r"(?:at\s+)?(?:beginning|end)\s+of\s+(?:the\s+)?(?:fiscal\s+)?(?:year|period)",
        re.IGNORECASE,
    ),
]


# Loose title — accepts any "Statements of Cash Flows" line, including
# "Statements of Cash Flows ‐ Discretely Presented Component Unit"
_PROP_CFS_LOOSE_TITLE_RE = re.compile(
    r"statements?\s+of\s+cash\s+flows?",
    re.IGNORECASE,
)

# Same title regex but tolerant of de-scrambled/collapsed text (no spaces)
_PROP_CFS_COLLAPSED_TITLE_RE = re.compile(
    r"statements?\s*of\s*cash\s*flows?",
    re.IGNORECASE,
)

# Component-unit / discrete-presentation hints (strengthens confidence)
_PROP_CFS_COMPONENT_UNIT_HINT_RE = re.compile(
    r"\bdiscretely\s+presented\s+component\s+unit\b"
    r"|\bcomponent\s+unit\b"
    r"|\bfoundation\b"
    r"|\bchange\s+in\s+net\s+assets\b",     # FASB-style (Foundation reports)
    re.IGNORECASE,
)


def _looks_like_prop_cfs_body(text: str) -> bool:
    """
    Confirm a page is a Statement of Cash Flows purely by body content,
    without needing the title regex to match perfectly.
    Requires at least 2 of the 3 trio markers + numeric amounts.
    """
    if not text:
        return False
    trio_hits = sum(1 for p in _PROP_CFS_BODY_TRIO if p.search(text))
    if trio_hits < 2:
        return False
    # Must have real dollar amounts (avoid TOC / narrative pages)
    amounts = re.findall(r"\b\d{1,3}(?:,\d{3})+\b", text or "")
    return len(amounts) >= 5


def find_prop_cfs_forward_lookahead(pdf_path: str, prop_cfs_last_page: int,
                                    defn: dict, lookahead: int = 4) -> list:
    """
    Scan forward from PROP_CFS's last collected page to find a component-unit
    Statement of Cash Flows that the main loop missed (e.g., because the
    primary PROP_CFS already hit its completion anchor).

    Mirrors find_prop_snp_reverse_lookback() in direction-reversed form.

    Returns a list of 1-based page numbers (anchor + any continuation pages),
    or [] if nothing qualifies.
    """
    require_kw     = defn.get("require_kw", [])
    is_complete_fn = defn.get("is_complete_fn")
    cont_re        = defn.get("cont_re")
    stop_res       = defn.get("stop_res", [])
    collected = []

    with pdfplumber.open(pdf_path) as pdf:
        total = len(pdf.pages)
        scan_start = prop_cfs_last_page          # 1-based; start AFTER last page
        scan_end   = min(total, prop_cfs_last_page + lookahead)

        for page_no in range(scan_start + 1, scan_end + 1):
            text = _get_robust_page_text(pdf, pdf_path, page_no - 1)
            if not text:
                continue

            # ── Hard exclusions ──
            if _fund_excluded(text, defn):
                continue
            if _is_toc_page(text):
                continue
            if _is_narrative_summary_page(text):
                continue

            # ── Stop early if we hit a clearly different statement / Notes ──
            stop_hit = False
            for stop_re, n_lines in stop_res:
                header_area = "\n".join(text.splitlines()[:n_lines])
                if stop_re.search(header_area):
                    stop_hit = True
                    break
            if stop_hit:
                break

            collapsed = _collapse_spaced_text(text)

            # ── Title / body detection ──
            has_loose_title = (
                bool(_PROP_CFS_LOOSE_TITLE_RE.search(collapsed)) or
                bool(_PROP_CFS_COLLAPSED_TITLE_RE.search(collapsed))
            )
            body_ok = _looks_like_prop_cfs_body(collapsed)
            require_ok = all(kw.search(collapsed) for kw in require_kw) if require_kw else True
            component_hint = bool(_PROP_CFS_COMPONENT_UNIT_HINT_RE.search(collapsed))

            # Anchor accepted if EITHER:
            #   (a) Loose title + body content, OR
            #   (b) Body content + component-unit hint (catches fused titles)
            if not ((has_loose_title and body_ok and require_ok)
                    or (body_ok and component_hint)):
                continue

            collected.append(page_no)

            # ── Look one page further for continuation (e.g., reconciliation) ──
            if page_no + 1 <= scan_end:
                nxt_text = _get_robust_page_text(pdf, pdf_path, page_no)
                if nxt_text and not _fund_excluded(nxt_text, defn):
                    nxt_collapsed = _collapse_spaced_text(nxt_text)
                    if cont_re and cont_re.search(nxt_collapsed):
                        collected.append(page_no + 1)

            # Stop once anchor found + completion detected on this OR next page
            window_text = collapsed
            if is_complete_fn and is_complete_fn(window_text):
                break

    return collected

# Loose title hint — accepts ANY occurrence of the title phrase anywhere on
# the page, including fused-to-entity-name variants. Used ONLY in fallback.

# Already given earlier — matches normal spaced text
_PROP_SNP_LOOSE_TITLE_RE = re.compile(
    r"statements?\s+of\s+"
    r"(?:(?:fund\s+)?net\s+(?:pos(?:i)?tion|assets)|financial\s+position)",
    re.IGNORECASE,
)

# NEW — matches collapsed/no-space text from de-scrambling
_PROP_SNP_COLLAPSED_TITLE_RE = re.compile(
    r"statements?\s*of\s*"
    r"(?:(?:fund\s*)?net\s*(?:pos(?:i)?tion|assets)|financial\s*position)",
    re.IGNORECASE,
)

# Same for require_kw — use a spaced-flexible variant
_PROP_SNP_COLLAPSED_REQUIRE_KW = [
    re.compile(r"assets",                              re.IGNORECASE),
    re.compile(r"net\s*(?:position|assets)",           re.IGNORECASE),
]




_PROP_SNP_RECONCILIATION_RE = re.compile(
    r"\breconciliation\b"
    r"|\bAmounts\s+reported\s+for\s+governmental\s+activities\s+in\s+the\s+Statement\s+of\s+Net\s+Position\s+are\s+different\b"
    r"|\bTotal\s+fund\s+balances?\s*[-–—]\s*governmental\s+funds?\b",
    re.IGNORECASE,
)


def _looks_like_prop_snp_body(text: str) -> bool:
    """
    Confirm a page is a Statement of Net Position purely by body content,
    without needing the title regex to match.

    NEW: Explicitly reject Reconciliation-of-Balance-Sheet-to-SNP pages,
    which otherwise satisfy the trio check by accident (bonds payable +
    long-term liabilities + total net position all appear as
    reconciling items).
    """
    if not text:
        return False

    # ── Reconciliation guard (NEW) ────────────────────────────────
    first25 = "\n".join(text.splitlines()[:25])
    if _PROP_SNP_RECONCILIATION_RE.search(first25):
        return False

    trio_hits = sum(1 for p in _PROP_SNP_BODY_TRIO if p.search(text))
    if trio_hits < 2:
        return False

    # Must have real dollar amounts (avoid TOC / narrative pages)
    amounts = re.findall(r"\b\d{1,3}(?:,\d{3})+\b", text or "")
    return len(amounts) >= 5






def find_prop_snp_reverse_lookback(pdf_path: str, prop_is_first_page: int,
                                    defn: dict, lookback: int = 4) -> list:
    """
    Scan backward from PROP_IS's first page looking for PROP_SNP anchor page.

    Handles THREE failure modes that defeat normal title detection:
      1. Fused entity names ("Purdue UniversityStatement of Net Position")
      2. Page number prefix on title line ("22 Statement of Net Position")
      3. SCRAMBLED text where pdfplumber extracts each char spaced apart
         ("S t a t e m e n t  o f  N e t  P o s i t i o n")

    NEW guard (PROP_SNP page-41 bug fix):
      Reject "Reconciliation of the Balance Sheet to the Statement of Net
      Position - Governmental Funds" pages. These pages contain the string
      "Statement of Net Position" in their title AND all three body-trio
      markers ("Bonds payable", "Long-term liabilities", "Total net
      position") — they fool BOTH the loose title regex and
      _looks_like_prop_snp_body(). The main forward scan already excludes
      them via EXCLUDE_PATTERNS (\breconciliation\b), but this reverse-
      lookback fallback bypasses that check, so we explicitly skip
      reconciliation pages here.
    """
    require_kw   = defn["require_kw"]
    scrambled_kw = defn.get("scrambled_kw", [])
    cont_re      = defn.get("cont_re")

    candidates = []

    with pdfplumber.open(pdf_path) as pdf:
        total = len(pdf.pages)
        start = max(1, prop_is_first_page - lookback)
        end   = prop_is_first_page - 1

        for pn in range(end, start - 1, -1):       # walk BACKWARD
            raw_text = pdf.pages[pn - 1].extract_text() or ""
            if not raw_text.strip():
                continue

            # ── De-scramble character-spaced text ──
            collapsed = _collapse_spaced_text(raw_text)

            # ── Hard exclusions (use whichever form has more real words) ──
            chk_text = collapsed if len(re.findall(r"\b[A-Za-z]{4,}\b", collapsed)) \
                                 > len(re.findall(r"\b[A-Za-z]{4,}\b", raw_text)) \
                                 else raw_text
            if _nonlg_excluded(chk_text):
                continue

            # ╔═══════════════════════════════════════════════════════════════╗
            # ║  NEW GUARD — Reject Reconciliation-to-SNP pages               ║
            # ║  (INPUT 1.pdf page 41 = "Reconciliation of the Balance Sheet  ║
            # ║   to the Statement of Net Position - Governmental Funds")    ║
            # ╚═══════════════════════════════════════════════════════════════╝
            first25_raw       = "\n".join(raw_text.splitlines()[:25])
            first25_collapsed = "\n".join(collapsed.splitlines()[:25])
            if (_PROP_SNP_RECONCILIATION_RE.search(first25_raw) or
                _PROP_SNP_RECONCILIATION_RE.search(first25_collapsed)):
                print(f"      [PROP_SNP fallback] page {pn} SKIPPED "
                      f"(reconciliation page)")
                continue

            # ╔═══════════════════════════════════════════════════════════════╗
            # ║  PATCH 1 — Title detection (try ALL 3 patterns on both forms) ║
            # ╚═══════════════════════════════════════════════════════════════╝
            has_loose_title = (
                bool(_PROP_SNP_LOOSE_TITLE_RE.search(collapsed))     or
                bool(_PROP_SNP_COLLAPSED_TITLE_RE.search(collapsed)) or
                bool(_PROP_SNP_LOOSE_TITLE_RE.search(raw_text))
            )

            # ╔═══════════════════════════════════════════════════════════════╗
            # ║  PATCH 2 — require_kw (try raw, collapsed, AND flex version)  ║
            # ╚═══════════════════════════════════════════════════════════════╝
            kw_ok = (
                all(p.search(raw_text)  for p in require_kw) or
                all(p.search(collapsed) for p in require_kw) or
                all(p.search(collapsed) for p in _PROP_SNP_COLLAPSED_REQUIRE_KW)
            )
            if not kw_ok:
                continue

            # ── Body / completion detection ──
            has_body     = _looks_like_prop_snp_body(collapsed) or \
                           _looks_like_prop_snp_body(raw_text)
            has_complete = bool(_PROP_SNP_COMPLETE_RE.search(collapsed)) or \
                           bool(_PROP_SNP_COMPLETE_RE.search(raw_text))
            is_scrambled = bool(scrambled_kw) and _is_scrambled_anchor(raw_text, scrambled_kw)

            if has_loose_title or has_body or is_scrambled:
                score = (
                    (3 if has_loose_title else 0) +
                    (2 if has_complete   else 0) +
                    (1 if has_body       else 0) +
                    (2 if is_scrambled   else 0)
                )
                candidates.append((pn, score))
                print(f"      [PROP_SNP fallback] page {pn} candidate "
                      f"(title={has_loose_title}, body={has_body}, "
                      f"complete={has_complete}, scrambled={is_scrambled}, "
                      f"score={score})")

        if not candidates:
            return []

        # ── Pick EARLIEST viable candidate (score >= 3) ──
        viable = [c for c in candidates if c[1] >= 3]
        if not viable:
            return []
        viable.sort(key=lambda x: x[0])      # earliest page first
        anchor_page = viable[0][0]
        collected   = [anchor_page]

        # ── Try to extend forward by 1 page for continuation ──
        next_pn = anchor_page + 1
        if next_pn < prop_is_first_page and next_pn <= total:
            next_raw       = pdf.pages[next_pn - 1].extract_text() or ""
            next_collapsed = _collapse_spaced_text(next_raw)
            next_chk       = next_collapsed if len(next_collapsed) > len(next_raw) else next_raw

            # ── Same reconciliation guard on continuation page ──
            next_first25_raw       = "\n".join(next_raw.splitlines()[:25])
            next_first25_collapsed = "\n".join(next_collapsed.splitlines()[:25])
            is_recon_next = (
                _PROP_SNP_RECONCILIATION_RE.search(next_first25_raw) or
                _PROP_SNP_RECONCILIATION_RE.search(next_first25_collapsed)
            )

            if (next_raw.strip()
                    and not _nonlg_excluded(next_chk)
                    and not is_recon_next):
                is_cont = (
                    (cont_re and (cont_re.search(next_raw) or cont_re.search(next_collapsed))) or
                    re.search(r"\(continued\b", next_raw,       re.IGNORECASE) or
                    re.search(r"\(continued\b", next_collapsed, re.IGNORECASE) or
                    _looks_like_prop_snp_body(next_collapsed) or
                    _looks_like_prop_snp_body(next_raw) or
                    bool(_PROP_SNP_COMPLETE_RE.search(next_collapsed)) or
                    bool(_PROP_SNP_COMPLETE_RE.search(next_raw))
                )
                if is_cont:
                    collected.append(next_pn)
                    print(f"      [PROP_SNP fallback] page {next_pn} added as continuation")
            elif is_recon_next:
                print(f"      [PROP_SNP fallback] page {next_pn} SKIPPED "
                      f"as continuation (reconciliation page)")

    return sorted(collected)


# ══════════════════════════════════════════════════════════════════════════════
# PROP_IS SANDWICH FALLBACK
# ══════════════════════════════════════════════════════════════════════════════
# When PROP_IS title is unreadable (e.g., overlapping text layers in OUC-style
# PDFs), but PROP_SNP and PROP_CFS were both found, any page sandwiched between
# them containing the trio (Operating revenues + Operating expenses + Operating
# income) is definitively PROP_IS.
# ──────────────────────────────────────────────────────────────────────────────

_PROP_IS_BODY_TRIO = [
    re.compile(r"\bOperating\s+revenues?\b",  re.IGNORECASE),
    re.compile(r"\bOperating\s+expenses?\b",  re.IGNORECASE),
    re.compile(r"\bOperating\s+income\b",     re.IGNORECASE),
]

# Secondary markers used when only 2 of the 3 trio members match
# (e.g., page uses "Operating loss" instead of "Operating income")
_PROP_IS_BODY_FALLBACK = [
    re.compile(r"\bOperating\s+(?:income|loss)\b",                          re.IGNORECASE),
    re.compile(r"\bTotal\s+operating\s+revenues?\b",                        re.IGNORECASE),
    re.compile(r"\bTotal\s+operating\s+expenses?\b",                        re.IGNORECASE),
    re.compile(r"\bNon[-\s]?operating\s+(?:revenues?|expenses?|income)\b",   re.IGNORECASE),
    re.compile(r"\bChange\s+in\s+net\s+(?:position|assets)\b",              re.IGNORECASE),
]


def _looks_like_prop_is_body(text: str) -> bool:
    """
    Confirm a page is a Statement of Revenues/Expenses/Changes in Net Position
    purely by body content — without needing a readable title.

    A page qualifies if EITHER:
      • All three trio markers match (Operating revenues + expenses + income), OR
      • At least 3 of the 5 fallback markers match
    """
    trio_hits = sum(1 for p in _PROP_IS_BODY_TRIO if p.search(text or ""))
    if trio_hits == 3:
        return True
    fallback_hits = sum(1 for p in _PROP_IS_BODY_FALLBACK if p.search(text or ""))
    return fallback_hits >= 3
_PROP_IS_BODY_TRIO_FRS102 = [
    re.compile(r"\bTotal\s+income\b",      re.IGNORECASE),
    re.compile(r"\bTotal\s+expenditure\b", re.IGNORECASE),
    re.compile(r"\bSurplus\s+(?:for|before)\s+(?:the\s+)?year\b"
               r"|\bDeficit\s+(?:for|before)\s+(?:the\s+)?year\b"
               r"|\bExcess\s+of\s+(?:income|revenue)\b", re.IGNORECASE),
]
def _looks_like_prop_is_body(text: str) -> bool:
    # existing GASB check
    trio_hits = sum(1 for p in _PROP_IS_BODY_TRIO if p.search(text or ""))
    if trio_hits == 3:
        return True
    fallback_hits = sum(1 for p in _PROP_IS_BODY_FALLBACK if p.search(text or ""))
    if fallback_hits >= 3:
        return True
    # ── NEW: FRS 102 style ──
    frs102_hits = sum(1 for p in _PROP_IS_BODY_TRIO_FRS102 if p.search(text or ""))
    return frs102_hits == 3
def find_prop_is_sandwich(pdf_path: str, snp_last_page: int,
                          cfs_first_page: int) -> list:
    """
    Look for PROP_IS pages sandwiched strictly between PROP_SNP and PROP_CFS.
    Uses BOTH pdfplumber AND pypdf — pypdf bypasses overlapping-text-layer
    bugs that hide the title (e.g., OUC-style PDFs).

    Args:
        snp_last_page  : last page of PROP_SNP  (exclusive lower bound for scan)
        cfs_first_page : first page of PROP_CFS (exclusive upper bound for scan)

    Returns:
        List of 1-based page numbers confirmed as PROP_IS by body content.
    """
    if cfs_first_page - snp_last_page <= 1:
        return []  # no gap → nothing to sandwich

    collected = []
    reader = _get_cached_pdf_reader(pdf_path) 

    with pdfplumber.open(pdf_path) as pdf:
        for page_no in range(snp_last_page + 1, cfs_first_page):
            # Try pdfplumber first
            pp_text = pdf.pages[page_no - 1].extract_text() or ""
            # Always also pull pypdf — it survives overlapping layers
            try:
                py_text = reader.pages[page_no - 1].extract_text() or ""
            except Exception:
                py_text = ""

            # Combined text gives us both layers to match against
            combined = pp_text + "\n" + py_text

            # Skip pages that look like notes / RSI / supplemental
            if _nonlg_excluded(combined):
                continue

            if _looks_like_prop_is_body(combined):
                collected.append(page_no)

    return collected

def find_fund_statement_pages(pdf_path: str, defn: dict, start_page: int,
                               next_anchor_re=None) -> list:
    title_re           = defn["title_re"]
    require_kw         = defn["require_kw"]
    stop_res           = defn["stop_res"]
    cont_re            = defn["cont_re"]
    scrambled_kw       = defn.get("scrambled_kw", [])
    scr_cont_kw        = defn.get("scrambled_cont_kw", [])
    is_complete_fn     = defn.get("is_complete_fn")
    is_col_overflow_fn = defn.get("is_col_overflow_fn")
    complete_re        = defn.get("complete_re")
    gov_overflow       = defn.get("gov_overflow", False)

    def is_anchor(text: str) -> bool:
        if not text:
            return False
        if _title_in_header(text, title_re):
            first40 = "\n".join((text or "").splitlines()[:40])
            kw_matches = sum(1 for kw in require_kw if kw.search(first40))
            has_asset_block = bool(re.search(r"\bassets\b", first40, re.IGNORECASE))
            
            # ── NEW: reject MD&A narrative pages that mention "balance sheet"
            #    inside a sentence rather than as a standalone title ──
            if defn.get("suffix") == "GOV_BS":
                if not _has_financial_data(text, min_amounts=8):
                    return False
            
            if kw_matches >= 1 or has_asset_block:
                return True
        if scrambled_kw and _is_scrambled_anchor(text, scrambled_kw):
            return True
        if defn["suffix"] == "PROP_CFS" and _looks_like_prop_cfs_body(text):
            return True
        return False


    def hits_stop(text: str) -> bool:
        for stop_re, n in stop_res:
            if stop_re.search("\n".join((text or "").splitlines()[:n])):
                return True
        return False

    def is_scrambled_cont(text: str) -> bool:
        if not scr_cont_kw:
            return False
        lines = (text or "").splitlines()
        single_char = sum(1 for l in lines[:20] if len(l.strip()) <= 2)
        spaced_word = sum(1 for l in lines[:20] if re.match(r"^([A-Za-z]\s){2,}", l.strip()))
        dollar_lines = sum(1 for l in lines if re.search(r"\$\s*[\d,]+", l))
        if not (single_char >= 3 or spaced_word >= 1 or dollar_lines >= 3):
            return False
        return all(kw.search(text or "") for kw in scr_cont_kw)

    _CONT_MARKER_RE = re.compile(
    r"\bcont(?:inued\b|\.|'d\b)"
    r"|\bPAGE\s+[2-9]\d*\s+OF\s+\d+\b",
    re.IGNORECASE,
)


    def is_continuation(text: str, prev_text: str = "") -> bool:
        if _fund_excluded(text, defn):
            return False
        if hits_stop(text):
            return False

        first5  = "\n".join((text or "").splitlines()[:5])
        first15 = "\n".join((text or "").splitlines()[:15])
        first30 = "\n".join((text or "").splitlines()[:30])

        # ── NEW: explicit "(continued from previous page)" marker → strong continuation
        if re.search(
            r"\(continued\s+from\s+previous\s+page\)"
            r"|\bcontinued\s+from\s+previous\s+page\b"
            r"|\(continued\)",
            first15,
            re.IGNORECASE,
        ):
            return True

        # If next page has its own title and no continuation marker, stop
        if is_anchor(text) and not _CONT_MARKER_RE.search(first5):
            prev_last_lines = "\n".join((prev_text or "").splitlines()[-5:])
            if not _CONT_MARKER_RE.search(prev_last_lines):
                return False

        if cont_re.search(first30):
            return True
        return is_scrambled_cont(text)


    def is_truly_complete(page_text: str) -> bool:
        if not is_complete_fn:
            return False
        if not is_complete_fn(page_text):
            return False
        if complete_re is not None:
            return _complete_row_has_numbers(page_text, complete_re)
        return True

    collected = []
    scan_from = max(start_page, START_PAGE) - 1

    with pdfplumber.open(pdf_path) as pdf:
        total = len(pdf.pages)

        for i in range(scan_from, total):
            page_no = i + 1
            #text = pdf.pages[i].extract_text() or ""
            text = _get_robust_page_text(pdf, pdf_path, i)  # <-- NEW
            if _fund_excluded(text, defn):
                continue
            if not is_anchor(text):
                continue
            collected.append(page_no)
            last_non_overflow_complete = is_truly_complete(text)
            j = i + 1
 
            while j < total:
                #next_text = pdf.pages[j].extract_text() or ""
                next_text = _get_robust_page_text(pdf, pdf_path, j)    # <-- NEW
                next_first6 = "\n".join(next_text.splitlines()[:6])
                # ── FIX: Check "(continued)" title BEFORE completion break ──
                # This ensures pages like "Statements of Cash Flows, continued"
                # are always included, even if the previous page was "complete"
                

                # ── FIX: NOW check completion — only break if last page was
                #    complete AND next page has no "(continued)" marker ──
                

                if _fund_excluded(next_text, defn):   # ← FIX: pass defn
                    break

                if hits_stop(next_text):
                    break
                if next_anchor_re and _title_in_header(next_text, next_anchor_re):
                    break
                if (_title_in_header(next_text, title_re)
                    and _CONT_MARKER_RE.search(next_first6)
                    and not last_non_overflow_complete):
                    collected.append(j + 1)
                    last_non_overflow_complete = is_truly_complete(next_text)
                    j += 1
                    continue
                if is_col_overflow_fn and is_col_overflow_fn(next_text):
                    collected.append(j + 1)
                    last_non_overflow_complete = is_truly_complete(next_text)  # ← update state
                    j += 1
                    if gov_overflow and last_non_overflow_complete:            # ← only break when done
                        break
                    continue

                if last_non_overflow_complete:
                    break

                #prev_text = pdf.pages[j - 1].extract_text() or ""
                prev_text = _get_robust_page_text(pdf, pdf_path, j - 1)   # <-- NEW
                if is_continuation(next_text, prev_text=prev_text):
                    collected.append(j + 1)
                    last_non_overflow_complete = is_truly_complete(next_text)
                    j += 1
                    continue
                break
            break

    if not collected:
        raise ValueError(f"{defn['suffix']} not found in {os.path.basename(pdf_path)}")
    return sorted(set(collected))




_PROP_PRESENCE_KW = [
    # ── Strong single-keyword signals ──
    re.compile(r"\benterprise\s+funds?\b",                       re.IGNORECASE),
    re.compile(r"\bproprietary\s+funds?\b",                      re.IGNORECASE),
    re.compile(r"\binternal\s+service\s+funds?\b",               re.IGNORECASE),  # ← NEW
    re.compile(r"\bbusiness[-\s]*type\s+activit",                re.IGNORECASE),
    re.compile(r"\bmajor\s+enterprise\b",                        re.IGNORECASE),
    re.compile(r"\bnonmajor\s+enterprise\b",                     re.IGNORECASE),

    # ── Title-pair signals (window widened 120 → 300) ──
    re.compile(r"\bstatement\s+of\s+(?:fund\s+)?net\s+position\b.{0,300}"
               r"\b(?:proprietary|enterprise|internal\s+service)\b",
               re.IGNORECASE | re.DOTALL),
    re.compile(r"\b(?:proprietary|enterprise|internal\s+service)\b.{0,300}"
               r"\bstatement\s+of\s+(?:fund\s+)?net\s+position\b",
               re.IGNORECASE | re.DOTALL),
    re.compile(r"\bstatement\s+of\s+cash\s+flows?\b.{0,300}"
               r"\b(?:proprietary|enterprise|internal\s+service)\b",
               re.IGNORECASE | re.DOTALL),
    re.compile(r"\b(?:proprietary|enterprise|internal\s+service)\b.{0,300}"
               r"\bstatement\s+of\s+cash\s+flows?\b",
               re.IGNORECASE | re.DOTALL),
    re.compile(r"\b(?:proprietary|enterprise|internal\s+service)\b.{0,300}"
               r"\bstatement\s+of\s+revenues",
               re.IGNORECASE | re.DOTALL),

    # ── Near-definitive: Statement of Cash Flows itself ──
    # (Under GASB, governmental funds do NOT produce a CFS — only
    #  proprietary/enterprise/internal service funds do.)
    re.compile(r"\bstatement\s+of\s+cash\s+flows?\b",            re.IGNORECASE),
]


def _has_proprietary_fund_statements(pdf_path: str) -> bool:
    """
    Returns True if PDF contains ANY proprietary-style fund statement,
    including:
      • Proprietary Fund    (full proprietary fund column)
      • Enterprise Fund     (Water, Sewer, Utility, Airport, etc.)
      • Internal Service Fund (Self-Insurance, Fleet, IT, etc.)  ← was being missed
      • Business-type Activities

    Threshold lowered from 2 → 1 hit. A single strong signal is enough;
    false positives just cause empty extraction, which is cheap.
    """
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""

            # Skip Notes / Schedules / MD&A / Statistical pages
            first6 = "\n".join(text.splitlines()[:6])
            if re.search(
                r"\bnotes?\s+to\b"
                r"|\bsupplemental\b"
                r"|\bschedule\s+of\b"
                r"|\bmanagement.{0,30}discussion\b"
                r"|\bmd\s*(?:&|and)\s*a\b"
                r"|\btable\s+of\s+contents?\b"
                r"|\bstatistical\b",
                first6, re.IGNORECASE,
            ):
                continue

            for kw in _PROP_PRESENCE_KW:
                if kw.search(text):
                    return True
    return False




# ══════════════════════════════════════════════════════════════════════════════
# DSR — DEBT SERVICE REQUIREMENT FINDER  (scoring-based)
# ══════════════════════════════════════════════════════════════════════════════

DSR_SCORE_THRESHOLD = 8

DSR_KEYWORDS = [
    "principal", "principle","interest", "maturity", "debt service", "obligation",
    "bonds payable", "notes payable", "loan payable", "amortization",
    "outstanding", "payment", "installment", "annual", "retirement",
    "unamortized", "premium", "discount",
]

_DSR_HARD_EXCLUDE_RE = re.compile(
    r"\btable\s+of\s+contents?\b"
    r"|\bforeword\b"
    r"|\bletter\s+of\s+transmittal\b"
    r"|\bintroductory\b"
    r"|\bmd\s*(?:&|and)\s*a\b"
    r"|\bmanagement.{0,30}discussion\b"
    r"|\bstatistical\b"
    r"|\bbudget\b"
    r"|\bcombining\b"
    r"|\bsupplemental\b(?!\s+information\s+about)"
    r"|\bcomponent\s+unit\b",
    re.IGNORECASE,
)

_DSR_MIN_NUMERIC_AMOUNTS = 10
_DSR_MIN_NUMERIC_AMOUNTS_RELAXED = 5

_DSR_LEASE_SUBTABLE_RE = re.compile(
    r"\bfuture\s+minimum\s+(?:lease|rental)\s+(?:obligations?|payments?)\b"
    r"|\bminimum\s+(?:lease|rental)\s+(?:obligations?|payments?)\b",
    re.IGNORECASE,
)

_DSR_MIN_COLS_FOR_REAL_DSR = 4

_NEW_NOTE_RE = re.compile(
    r"\bNOTE\s+\d+\s*[-\u2013\u2014]\s*(?!long|debt|bond|loan)",
    re.IGNORECASE,
)

_DSR_PENSION_OPEB_RE = re.compile(
    r"\bpension\b"
    r"|\bOPEB\b"
    r"|\bactuarial\b"
    r"|\bnet\s+pension\s+liability\b"
    r"|\bfiduciary\s+net\s+position\b"
    r"|\bdeferred\s+(?:outflows|inflows)\s+of\s+resources\b"
    r"|\bretirement\s+(?:system|plan|fund)\b"
    r"|\bpost[-\s]?employment\b"
    r"|\bsingle[-\s]?employer\b"
    r"|\bcost[-\s]?sharing\b",
    re.IGNORECASE,
)

_DSR_DEBT_SERVICE_RE = re.compile(
    r"\bdebt\s+service\b"
    r"|\bprincipal\s+and\s+interest\b"
    r"|\bfinanced\s+purchase"
    r"|\bbonds?\s+payable\b"
    r"|\bnotes?\s+payable\b"
    r"|\bloan\s+payable\b"
    r"|\blong[-\s]term\s+(?:debt|liabilit)"
    r"|\bnoncurrent\s+liabilit"
    r"|\bsubscription[-\s]based\b"
    r"|\bSBITA\b",
    re.IGNORECASE,
)

_DSR_EXPLICIT_HEADER_RE = re.compile(
    r"\bdebt\s+service\s+requirements?\s+to\s+maturity\b"
    r"|\bannual\s+(?:debt\s+service\s+)?requirements?\s+to\s+maturity\b"
    r"|\bdebt\s+service\s+requirements?\b"
    r"|\bannual\s+debt\s+service\s+(?:for|on)\b[^\n]{0,80}\bto\s+maturity\b"   # ← NEW
    r"|\bdebt\s+service\s+(?:for|on)\b[^\n]{0,80}\bto\s+maturity\b",          # ← NEW
    re.IGNORECASE,
)


def _dsr_max_year_row_cols(text: str) -> int:
    max_cols = 0
    for line in (text or "").splitlines():
        stripped = line.strip()
        if re.match(r"^20[2-9]\d", stripped) or re.match(r"^\$", stripped):
            cols = len(re.findall(r"\d{1,3}(?:,\d{3})+", line))
            max_cols = max(max_cols, cols)
    return max_cols


def _dsr_is_lease_only_page(text: str) -> bool:
    if not _DSR_LEASE_SUBTABLE_RE.search(text):
        return False
    return _dsr_max_year_row_cols(text) < _DSR_MIN_COLS_FOR_REAL_DSR


def _dsr_hard_excluded(text: str) -> bool:
    first8 = "\n".join((text or "").splitlines()[:8])
    return bool(_DSR_HARD_EXCLUDE_RE.search(first8))


def _dsr_is_pension_opeb_page(text: str) -> bool:
    pension_hits = len(_DSR_PENSION_OPEB_RE.findall(text or ""))
    debt_hits    = len(_DSR_DEBT_SERVICE_RE.findall(text or ""))
    if not (pension_hits >= 3 and debt_hits <= 1):
        return False
    top_half = "\n".join((text or "").splitlines()[:15])
    has_year_rows = bool(re.search(r"^\s*20[2-9]\d", top_half, re.MULTILINE))
    has_principal_interest = (
        bool(re.search(r"\bPrincipal\b", top_half, re.IGNORECASE))
        and bool(re.search(r"\bInterest\b", top_half, re.IGNORECASE))
    )
    if has_year_rows and has_principal_interest:
        return False
    return True


def _detect_fiscal_year_for_dsr(pdf_path: str) -> int:
    reader = _get_cached_pdf_reader(pdf_path)
    for page in reader.pages[:15]:
        text = page.extract_text() or ""
        m = re.search(r"[Ff]iscal\s+[Yy]ear\s+[Ee]nded?\b.*?(20\d{2})", text, re.S)
        if m:
            return int(m.group(1))
        m = re.search(r"[Yy]ear\s+[Ee]nded?\s+\w+\s+\d+,\s+(20\d{2})", text)
        if m:
            return int(m.group(1))
    return datetime.date.today().year - 1



def _dsr_score(text: str, fiscal_year: int) -> int:
    score = 0
    for kw in DSR_KEYWORDS:
        if re.search(r"\b" + re.escape(kw) + r"\b", text, re.IGNORECASE):
            score += 1

    # ── NEW: Strong bonus for explicit DSR header ────────────────────
    # "Debt service requirements" is the most definitive DSR signal.
    # Pre-filters already relax min_years/min_amounts for it,
    # but the score itself needs a boost for small DSR tables
    # (e.g., counties with only 2 future year rows).
    if _DSR_EXPLICIT_HEADER_RE.search(text):
        score += 3

    if re.search(r"\byear\s+end(?:ing|ed)\b|\bYears\b", text, re.IGNORECASE):
        score += 3
    row_label_years: set = set()
    for line in text.splitlines():
        m = re.match(r"^\s*(20[2-9]\d)(?:\s*-\s*20[2-9]\d)?\b", line.strip())
        if m:
            yr = int(m.group(1))
            if yr > fiscal_year:
                row_label_years.add(yr)
    sorted_row_years = sorted(row_label_years)
    if len(sorted_row_years) >= 3:
        score += 3
        if sorted_row_years == sorted(sorted_row_years):
            score += 2
    amounts = re.findall(r"\d{1,3}(?:,\d{3})+|\b\d{5,}\b", text)
    if len(amounts) >= 25:
        score += 3
    return score


def _count_row_label_years(text: str, fiscal_year: int) -> int:
    row_label_years: set = set()
    for line in (text or "").splitlines():
        m = re.match(r"^\s*(20[2-9]\d)(?:\s*-\s*20[2-9]\d)?\b", line.strip())
        if m:
            yr = int(m.group(1))
            if yr > fiscal_year:
                row_label_years.add(yr)
    return len(row_label_years)


def _dsr_table_complete(text: str) -> bool:
    lines = (text or "").splitlines()
    data_rows = 0
    saw_data_row = False
    for line in lines:
        stripped = line.strip()
        if re.match(r"^\s*20[2-9]\d", stripped) and re.search(r"\d{1,3}(?:,\d{3})+", line):
            saw_data_row = True
            data_rows += 1
            continue
        if not saw_data_row:
            continue
        if data_rows < 5:
            continue
        if re.match(r"^\s*(?:Totals?|Total\s+Debt\s+Service)", stripped, re.IGNORECASE):
            return True
        if (re.match(r"^\s*\$", stripped)
                and len(re.findall(r"\d{1,3}(?:,\d{3})+", stripped)) >= 2
                and not re.search(r"20[2-9]\d", stripped)):
            return True
    return False


def _dsr_continuation_score(text: str, fiscal_year: int) -> bool:
    if _dsr_hard_excluded(text):
        return False
    first5 = "\n".join((text or "").splitlines()[:5])
    if _NEW_NOTE_RE.search(first5):
        return False
    if _count_row_label_years(text, fiscal_year) < 2:
        return False
    lines = (text or "").splitlines()
    tabular_numeric_lines = sum(
        1 for l in lines
        if re.search(r"(?:^\s*\$?\s*\d{1,3}(?:,\d{3})+|\d{1,3}(?:,\d{3})+\s*$)", l.strip())
    )
    return tabular_numeric_lines >= 4


def _do_scrambled_lookahead(pdf, collected, collected_set, total):
    _POST_LOOK = 5
    _anchor_note_num = None
    _anm = re.search(
        r"\bNOTE\s+(\d+)\b",
        _collapse_spaced_text(pdf.pages[collected[0] - 1].extract_text() or ""),
        re.IGNORECASE,
    )
    if _anm:
        _anchor_note_num = int(_anm.group(1))

    _last_collected_idx = max(collected_set) - 1
    for _lk in range(_last_collected_idx + 1,
                     min(_last_collected_idx + 1 + _POST_LOOK, total)):
        _raw_lk = pdf.pages[_lk].extract_text() or ""
        _lt_lk  = _collapse_spaced_text(_raw_lk)

        if _dsr_hard_excluded(_lt_lk):
            continue

        _lk_top8 = "\n".join(_lt_lk.splitlines()[:8])
        _lk_note_m = re.search(r"\bNOTE\s+(\d+)\b", _lk_top8, re.IGNORECASE)
        if _lk_note_m and _anchor_note_num is not None:
            if int(_lk_note_m.group(1)) != _anchor_note_num:
                break

        _raw_lines = _raw_lk.splitlines()
        _n_single = sum(1 for l in _raw_lines[:50] if len(l.strip()) <= 2)
        _n_spaced = sum(1 for l in _raw_lines[:50]
                        if re.match(r"^([A-Za-z]\s){2,}", l.strip()))
        if _n_single < 5 and _n_spaced < 2:
            continue

        _has_debt_kw = bool(re.search(
            r"\b(?:Interest|Principal|Leases?|debt\s+service"
            r"|bonds?\s+payable|notes?\s+payable)\b",
            _lt_lk, re.IGNORECASE,
        ))

        if _has_debt_kw and "$" in _raw_lk:
            collected.append(_lk + 1)
            collected_set.add(_lk + 1)


def find_dsr_pages(pdf_path: str, start_page: int = START_PAGE) -> list:
    collected = []
    fiscal_year = _detect_fiscal_year_for_dsr(pdf_path)

    with pdfplumber.open(pdf_path) as pdf:
        total = len(pdf.pages)
        collected_set = set()

        i = START_PAGE - 1
        while i < total:
            page_no = i + 1

            if page_no in collected_set:
                i += 1
                continue

            text = pdf.pages[i].extract_text() or ""
            text = _collapse_spaced_text(text)
            if _dsr_hard_excluded(text):
                i += 1
                continue

            if _dsr_is_pension_opeb_page(text):
                i += 1
                continue

            numeric_amounts = re.findall(r"\d{1,3}(?:,\d{3})+|\b\d{5,}\b", text)
            has_explicit_header = bool(_DSR_EXPLICIT_HEADER_RE.search(text))
            min_amounts = (_DSR_MIN_NUMERIC_AMOUNTS_RELAXED
                           if has_explicit_header
                           else _DSR_MIN_NUMERIC_AMOUNTS)
            if len(numeric_amounts) < min_amounts:
                i += 1
                continue

            row_label_years: set = set()
            for line in text.splitlines():
                m = re.match(r"^\s*(20[2-9]\d)(?:\s*-\s*20[2-9]\d)?\b", line.strip())
                if m:
                    yr = int(m.group(1))
                    if yr > fiscal_year:
                        row_label_years.add(yr)
            min_years = 2 if has_explicit_header else 3
            if len(row_label_years) < min_years:
                i += 1
                continue
            score = _dsr_score(text, fiscal_year)
            if score < DSR_SCORE_THRESHOLD:
                i += 1
                continue

            collected.append(page_no)
            collected_set.add(page_no)

            if _dsr_is_lease_only_page(text):
                j = i + 1
                while j < total:
                    next_text = pdf.pages[j].extract_text() or ""
                    next_text = _collapse_spaced_text(next_text)

                    if _dsr_is_pension_opeb_page(next_text):
                        j += 1
                        continue

                    next_score = _dsr_score(next_text, fiscal_year)
                    next_row_years = _count_row_label_years(next_text, fiscal_year)
                    if (next_score >= DSR_SCORE_THRESHOLD
                            and next_row_years >= 3
                            and not _dsr_hard_excluded(next_text)
                            and not _dsr_is_lease_only_page(next_text)):
                        collected.append(j + 1)
                        collected_set.add(j + 1)
                        j += 1
                        if _dsr_table_complete(next_text):
                            _stop = True
                            if j < total:
                                _peek = _collapse_spaced_text(pdf.pages[j].extract_text() or "")
                                if not _dsr_is_pension_opeb_page(_peek):
                                    _peek_score = _dsr_score(_peek, fiscal_year)
                                    _peek_years = _count_row_label_years(_peek, fiscal_year)
                                    if (_peek_score >= DSR_SCORE_THRESHOLD
                                            and _peek_years >= 3
                                            and not _dsr_hard_excluded(_peek)):
                                        _stop = False
                                    elif _dsr_continuation_score(_peek, fiscal_year):
                                        _stop = False
                            if _stop:
                                break
                        continue
                    if _dsr_continuation_score(next_text, fiscal_year):
                        collected.append(j + 1)
                        collected_set.add(j + 1)
                        j += 1
                        if _dsr_table_complete(next_text):
                            _stop = True
                            if j < total:
                                _peek = _collapse_spaced_text(pdf.pages[j].extract_text() or "")
                                if not _dsr_is_pension_opeb_page(_peek):
                                    _peek_score = _dsr_score(_peek, fiscal_year)
                                    _peek_years = _count_row_label_years(_peek, fiscal_year)
                                    if (_peek_score >= DSR_SCORE_THRESHOLD
                                            and _peek_years >= 3
                                            and not _dsr_hard_excluded(_peek)):
                                        _stop = False
                                    elif _dsr_continuation_score(_peek, fiscal_year):
                                        _stop = False
                            if _stop:
                                break
                        continue
                    j += 1

                _do_scrambled_lookahead(pdf, collected, collected_set, total)
                i = max(collected_set)
                continue

            anchor_complete = _dsr_table_complete(text)

            j = i + 1
            while j < total:
                next_text = pdf.pages[j].extract_text() or ""
                next_text = _collapse_spaced_text(next_text)

                if _dsr_is_pension_opeb_page(next_text):
                    j += 1
                    continue

                next_score = _dsr_score(next_text, fiscal_year)
                next_row_years = _count_row_label_years(next_text, fiscal_year)
                if (next_score >= DSR_SCORE_THRESHOLD
                        and next_row_years >= 3
                        and not _dsr_hard_excluded(next_text)):
                    collected.append(j + 1)
                    collected_set.add(j + 1)
                    j += 1
                    if _dsr_table_complete(next_text):
                        _stop = True
                        if j < total:
                            _peek = _collapse_spaced_text(pdf.pages[j].extract_text() or "")
                            if not _dsr_is_pension_opeb_page(_peek):
                                _peek_score = _dsr_score(_peek, fiscal_year)
                                _peek_years = _count_row_label_years(_peek, fiscal_year)
                                if (_peek_score >= DSR_SCORE_THRESHOLD
                                        and _peek_years >= 3
                                        and not _dsr_hard_excluded(_peek)):
                                    _stop = False
                                elif _dsr_continuation_score(_peek, fiscal_year):
                                    _stop = False
                        if _stop:
                            break
                    continue
                if _dsr_continuation_score(next_text, fiscal_year):
                    collected.append(j + 1)
                    collected_set.add(j + 1)
                    j += 1
                    if _dsr_table_complete(next_text):
                        _stop = True
                        if j < total:
                            _peek = _collapse_spaced_text(pdf.pages[j].extract_text() or "")
                            if not _dsr_is_pension_opeb_page(_peek):
                                _peek_score = _dsr_score(_peek, fiscal_year)
                                _peek_years = _count_row_label_years(_peek, fiscal_year)
                                if (_peek_score >= DSR_SCORE_THRESHOLD
                                        and _peek_years >= 3
                                        and not _dsr_hard_excluded(_peek)):
                                    _stop = False
                                elif _dsr_continuation_score(_peek, fiscal_year):
                                    _stop = False
                        if _stop:
                            break
                    continue
                if anchor_complete:
                    break
                break

            _do_scrambled_lookahead(pdf, collected, collected_set, total)
            i = max(collected_set)
            continue

    if not collected:
        raise ValueError(f"DSR table not found in {os.path.basename(pdf_path)}")

    return sorted(set(collected))


# ══════════════════════════════════════════════════════════════════════════════
# DEBT — LONG-TERM & SHORT-TERM DEBT SCHEDULE FINDER  (scoring-based)
# ══════════════════════════════════════════════════════════════════════════════

DEBT_SCORE_THRESHOLD = 4

_DEBT_KEYWORDS = [
    "long-term", "long term", "short-term", "short term",
    "bonds payable", "notes payable", "loan payable",
    "lease agreements", "lease", "compensated absences",
    "net pension liability", "total opeb liability", "opeb",
    "unamortized", "premium", "discount", "revenue notes", "obligation",
]

_DEBT_PAGE_ANCHOR_RES = [
    re.compile(r"\blong[-\s]term\s+(?:debt|liabilit)", re.IGNORECASE),
    re.compile(r"\bshort[-\s]term\s+(?:debt|liabilit)", re.IGNORECASE),
    re.compile(r"\bcomponents?\s+of\s+(?:long[-\s]term|short[-\s]term)?\s*liabilit", re.IGNORECASE),
    re.compile(r"\bchanges?\s+in\s+long[-\s]term\s+liabilit", re.IGNORECASE),
    re.compile(r"\bsummary\s+of\s+(?:changes?\s+in\s+)?long[-\s]term\s+liabilit", re.IGNORECASE),
    re.compile(r"\bNOTE\s+\d+\s*[-\u2013\u2014]\s*(?:LONG[-\s]TERM\s+DEBT|SHORT[-\s]TERM\s+DEBT|LIABILIT)", re.IGNORECASE),
    re.compile(r"\blong[-\s]term\s+debt\b", re.IGNORECASE),
    re.compile(r"\blong[-\s]term\s+obligation", re.IGNORECASE),
    re.compile(r"\bgeneral\s+long[-\s]term\s+obligations?\b", re.IGNORECASE),
]


_DEBT_STRUCTURE_ANCHOR_RES = [
    re.compile(r"\bbeginning\s+balance\b", re.IGNORECASE),
    re.compile(r"\bending\s+balance\b", re.IGNORECASE),
    re.compile(r"\badditions?\b", re.IGNORECASE),
    re.compile(r"\bdeletion?\b", re.IGNORECASE),
    re.compile(r"\bdisposal?\b", re.IGNORECASE),
    re.compile(r"\bretirement?\b", re.IGNORECASE),
    re.compile(r"\bmaturities\b", re.IGNORECASE),
    re.compile(r"\breductions?\b", re.IGNORECASE),
    re.compile(r"\bdue\s+within\s+one\s+year\b", re.IGNORECASE),
    re.compile(r"\bgovernmental\s+activities\b", re.IGNORECASE),
    re.compile(r"\bbusiness[-\s]type\s+activities\b", re.IGNORECASE),
    re.compile(r"\bg.?o.?\s+bonds?\b", re.IGNORECASE),
    re.compile(r"\btotal\s+(?:governmental|business)", re.IGNORECASE),
    re.compile(r"\blease\s+(?:assets?|liabilit)", re.IGNORECASE),
    re.compile(r"\bamounts?\s+due\b", re.IGNORECASE),
    re.compile(r"\bdeductions?\b", re.IGNORECASE),
    re.compile(r"\b(?:non[-\s]?current|noncurrent)\b", re.IGNORECASE),
    re.compile(r"\badjustments?\b", re.IGNORECASE),
    re.compile(r"\bretirements?\b", re.IGNORECASE),
    re.compile(r"\bissuances?\b", re.IGNORECASE),
    re.compile(r"\bincurred\b", re.IGNORECASE),               # "Debt Incurred"
    re.compile(r"\bretired\b", re.IGNORECASE),                 # "Bond Retired"
    # ── NEW: common alternate column headers ──
    re.compile(r"\bbalance\s+at\b", re.IGNORECASE),          # "Balance at June 30, 20XX"
    #re.compile(r"\bincreases?\b", re.IGNORECASE),             # alternate for "Additions"
    #re.compile(r"\bdecreases?\b", re.IGNORECASE),             # alternate for "Reductions"
    re.compile(r"\bdue\s+in\s+less\s+than\b", re.IGNORECASE),     # "Due In Less Than One Year"
    re.compile(r"\bBalance\s+(?:beginning|end)\s+of\s+(?:fiscal\s+)?year\b",
           re.IGNORECASE),
    #re.compile(r"\bas\s+of\b", re.IGNORECASE),                      # "As of June 30, 2025"

]

_DEBT_HEADER_ONLY_ANCHOR_RES = [
    re.compile(r"\bincreases?\b", re.IGNORECASE),
    re.compile(r"\bdecreases?\b", re.IGNORECASE),
    re.compile(r"\bas\s+of\b", re.IGNORECASE),
]


_OPEB_TABLE_RE = re.compile(
    r"\bChanges?\s+in\s+(?:the\s+)?Total\s+OPEB\s+Liability\b"
    r"|\bChanges?\s+in\s+(?:the\s+)?Net\s+OPEB\s+Liability\b"
    r"|\bTotal\s+OPEB\s+Liability\b.{0,30}\bBalance\s+at\b",
    re.IGNORECASE | re.DOTALL,
)



_DEBT_HARD_EXCLUDE_RE = re.compile(
    r"\bstatement\s+of\s+net\s+position\b"
    r"|\bstatement\s+of\s+activities\b"
    # ── NEW: Fund-level financial statement pages ────────────────
    r"|\bbalance\s+sheet\b"                          # BS – Governmental Funds
    r"|\bstatement\s+of\s+revenues\b"                # GOV_IS, PROP_IS, budget schedules
    r"|\bstatement\s+of\s+cash\s+flows?\b"           # PROP_CFS
    r"|\bbudget\s+and\s+actual\b"                     # Budget comparison schedules
    # ── Existing patterns below (unchanged) ──────────────────────
    r"|\bstatistical\b"
    r"|\bcombining\b"
    r"|\btable\s+of\s+contents?\b"
    r"|\bmanagement.{0,30}discussion\b"
    r"|\bmd\s*(?:&|and)\s*a\b"
    r"|\bbudget(?:ary)?\s+comparison\b"
    r"|\brequired\s+supplementary\b"
    r"|\bsupplemental\b"
    r"|\bNOTE\s+\d+\s*[-\u2013\u2014]\s*[^\n]{0,40}OPEB\b"
    r"|\bschedule\s+of\s+(?:changes\s+in\s+)?(?:total\s+)?OPEB\b"
    r"|\bdefined\s+benefit\s+pension\b"
    r"|\bschedule\s+of\s+proportionate\s+share\b"
    r"|\bNOTE\s+\w+\s*[-:]\s*(?:Subsequent\s+Events|Accounting\s+Changes?|Restatement)\b"
    r"|\bschedule\s+of\s+contributions\b",
    re.IGNORECASE,
)

_DEBT_LEASE_ASSET_EXCLUDE_RE = re.compile(
    r"\bNOTE\s+\d+\s*[-\u2013\u2014]\s*[^\n]{0,40}ASSETS?\b"
    r"|\b[A-Z]\.\s+Capital\s+Assets?\b"
    r"|\bLease\s+Assets?\s*:\s*$"
    r"|\bCapital\s+Asset\s+Activity\b"
    r"|\bAccumulated\s+Amortization\b.{0,60}\bLease\b"
    r"|\bAmortization\s+expense\s+was\s+charged\b.{0,200}\blease\b"
    r"|\bchanges\s+in\s+capital\s+assets\b",
    re.IGNORECASE | re.MULTILINE,
)
_DEBT_TABLE_COLUMN_KEYWORD_RES = [
    re.compile(r"\bAdditions?\b", re.IGNORECASE),
    re.compile(r"\bReductions?\b", re.IGNORECASE),
    re.compile(r"\bIncreases?\b", re.IGNORECASE),
    re.compile(r"\bDecreases?\b", re.IGNORECASE),
    re.compile(r"\bDue\s+[Ww]ithin\s+[Oo]ne\s+[Yy]ear\b", re.IGNORECASE),
    re.compile(r"\bDeletions?\b", re.IGNORECASE),
    re.compile(r"\bMaturities\b", re.IGNORECASE),
    re.compile(r"\bRetirements?\b", re.IGNORECASE),   # plural/noun form only, not bare "Retired"
]
_DEBT_FUSED_HEADERS = [
    (re.compile(r"\bBalanceat\b", re.IGNORECASE), "Balance at"),
    (re.compile(r"\bBalan\b", re.IGNORECASE), "Balance"),
    (re.compile(r"\bceat\b", re.IGNORECASE), "at"),
    (re.compile(r"\bDueW\b", re.IGNORECASE), "Due W"),
    (re.compile(r"\bith\s+in\b", re.IGNORECASE), "ithin"),
    (re.compile(r"\bOneY\b", re.IGNORECASE), "One Y"),
    (re.compile(r"\be\s+a\s+r\b", re.IGNORECASE), "ear"),
    (re.compile(r"\bJune30\b", re.IGNORECASE), "June 30"),
]
def _count_debt_column_keywords(text: str) -> int:
    return sum(1 for kw in _DEBT_TABLE_COLUMN_KEYWORD_RES if kw.search(text or ""))

def _normalize_debt_headers(text: str) -> str:
    for pattern, replacement in _DEBT_FUSED_HEADERS:
        text = pattern.sub(replacement, text)
    return text

def _debt_score(text: str) -> int:
    score = 0
    if any(re.search(r"\b" + re.escape(kw) + r"\b", text, re.IGNORECASE)
           for kw in _DEBT_KEYWORDS):
        score += 2
    if any(p.search(text) for p in _DEBT_PAGE_ANCHOR_RES):
        score += 2
    # Full-page structure patterns
    score += sum(1 for p in _DEBT_STRUCTURE_ANCHOR_RES if p.search(text))
    # Header-only patterns — check ONLY first 15 lines
    header_area = "\n".join((text or "").splitlines()[:15])
    score += sum(1 for p in _DEBT_HEADER_ONLY_ANCHOR_RES if p.search(header_area))
    return score

def _page_has_debt_table_structure(text: str) -> bool:
    """
    Return True ONLY if the page contains a DEBT changes-schedule structure.

    Handles three header formats:
      Format 1: "Beginning Balance" / "Ending Balance" (exact phrase)
      Format 2: Multi-line split: "Beginning  Ending" + "Balance  Balance" (first 6 lines)
      Format 3: "Balance  Balance  Due Within" with date lines like
                "July 1, 2024  Additions  Reductions  June 30, 2025  One Year"
                (searches entire page, not just first 6 lines)
    """
    lines = (text or "").splitlines()

    # ── NEW: Hard exclusion for restatement / accounting principle pages ──
    # Prevents false positives on pages like "Note 17: Change in Accounting
    # Principle" with GASB 101 restatement tables that look like DEBT but
    # are NOT debt change schedules.

    _HARD_EXCLUDE_RESTATEMENT_RE = re.compile(
        r"\bRestatement\b.*\bAccounting\s+Principle\b"
        r"|\bAccounting\s+Principle\b"
        # ── TIGHTENED: only block GASB 101 when it's the TOPIC of the page,
        #    not when it's a mere footnote on a debt/compensated-absence row ──
        r"|\bGASB\s+(?:Statement\s+)?(?:No\.?\s*)?101\s*[,:\-\u2013\u2014]?\s*Compensated\s+Absences\b"
        r"|\bImplementation\s+of\s+GASB\s+(?:Statement\s+)?(?:No\.?\s*)?101\b.{0,80}\b(?:Restatement|Change\s+in\s+Accounting)"
        r"|\bChange\s+in\s+Accounting\s+Principle\b"
        r"|\bNet\s+Position\s+Restatement\b"
        r"|\bBalance\s+of\s+Contracts\b",
        re.IGNORECASE | re.DOTALL,
    )

    # Apply to header area only (first 25 lines), NOT entire page
    header_for_restatement = "\n".join((text or "").splitlines()[:25])
    if _HARD_EXCLUDE_RESTATEMENT_RE.search(header_for_restatement):
        return False
    is_transposed_vertical = False

    # ── STEP 1: Try exact phrase match ───────────────────────────────
    # FIX 1: Removed "May" from month list — "balance may be used" is
    #         common English, not a date. Steps 3-9 handle May dates.
    # FIX 2: Removed Date...Balance DOTALL pattern — it crossed entire
    #         pages (1000+ chars). Steps 3-9 handle date+balance precisely.
    # FIX 3: Removed re.DOTALL flag — no longer needed.
    

# ── STEP 1: Try exact phrase match ───────────────────────────────
    has_balance = bool(re.search(
    r"\b(?:Beginning|Ending|Opening|Closing)\s+Balance\b"
    r"|\bBalance,?\s+(?:at\s+)?(?:July|June|January|August|September|October"
    r"|November|December|February|March|April|[0-9]{1,2})\b"
    r"|\bBalance\s+Outstanding\b"
    r"|\bPrior\s+Year'?s?\s+Balance\b"
    r"|\bBalance\s+(?:at\s+)?\d{1,2}[/-]\d{1,2}[/-]\d{2,4}\b"
    # ── NEW: transposed format — row labels like
    #    "Balance beginning of year" / "Balance, end of fiscal year"
    r"|\bBalance,?\s+(?:at\s+the\s+)?(?:beginning|end)\s+of\s+"
    r"(?:the\s+)?(?:fiscal\s+|prior\s+|current\s+)?year\b",
    text or "", re.IGNORECASE))


    # ── STEP 2: Multi-line header fallback (first 6 lines) ──────────
    if not has_balance:
        header_area = " ".join(lines[:6]).lower()
        if ("beginning" in header_area
                and "ending" in header_area
                and "balance" in header_area):
            has_balance = True

    # ── STEP 3: "Balance Balance" pattern (entire page) ──────────────
    _DATE_RE = re.compile(
    r"(?:Jan(?:uary)?|Feb(?:ruary)?|Mar(?:ch)?|Apr(?:il)?"
    r"|May|Jun(?:e)?|Jul(?:y)?|Aug(?:ust)?|Sep(?:t(?:ember)?)?"
    r"|Oct(?:ober)?|Nov(?:ember)?|Dec(?:ember)?)\.?\s+\d{1,2}"
    r"|\d{1,2}[/-]\d{1,2}[/-]\d{2,4}"
    r"|\b\d{1,2}/\d{1,2}/\d{2}\b",
    re.IGNORECASE,
)
    _CHANGE_RE = re.compile(
    r"\b(?:Additions?|Reductions?|Deletions?|Increases?"
    r"|Decreases?|Maturities|Issued|Retired|Retirements?"
    r"|Payments?|Repaid|Repayments?|Proceeds?)\b",
    re.IGNORECASE,
)
    if not has_balance:
        for i, line in enumerate(lines):
            if not re.search(r"\bBalance\b", line, re.IGNORECASE):
                continue
            window = lines[i : min(len(lines), i + 7)]
            window_text = " ".join(window)
            if len(re.findall(r"\bBalance\b", window_text, re.IGNORECASE)) >= 2:
                nearby = " ".join(
                    lines[max(0, i - 1) : min(len(lines), i + 8)]
                )
                if _DATE_RE.search(nearby) and _CHANGE_RE.search(nearby):
                    has_balance = True
                    break

    # ── STEP 4: "Balance" + date on adjacent lines (entire page) ─────
    if not has_balance:
        for i in range(len(lines) - 1):
            pair = lines[i] + " " + lines[i + 1]
            match = bool(re.search(
                r"\bBalance\b.*\b(?:Jan(?:uary)?|Feb(?:ruary)?|Mar(?:ch)?"
                r"|Apr(?:il)?|May|Jun(?:e)?|Jul(?:y)?|Aug(?:ust)?"
                r"|Sep(?:t(?:ember)?)?|Oct(?:ober)?|Nov(?:ember)?"
                r"|Dec(?:ember)?)\.?\s+\d{1,2}",
                pair, re.IGNORECASE))
            if match:
                balance_count = len(re.findall(r"\bBalance\b", pair, re.IGNORECASE))
                if balance_count >= 2:
                    has_balance = True
                    break

    # ── STEP 5: Fiscal-year date label pattern ────────────────────
    if not has_balance:
        _SHORT_DATE_RE = re.compile(
            r"\b\d{1,2}/\d{1,2}/\d{2,4}\b"
        )
        _DEBT_COL_RE = re.compile(
            r"\b(?:Issued|Retired|Retirements?|Additions?|Reductions?"
            r"|Proceeds?|Repaid|Payments?)\b",
            re.IGNORECASE,
        )
        for i, line in enumerate(lines):
            if not re.search(r"\bBalance\b", line, re.IGNORECASE):
                continue
            window = lines[i : min(len(lines), i + 5)]
            window_text = " ".join(window)
            date_hits = len(_SHORT_DATE_RE.findall(window_text))
            balance_hits = len(re.findall(
                r"\bBalance\b", window_text, re.IGNORECASE))
            col_hit = bool(_DEBT_COL_RE.search(window_text))
            if date_hits >= 2 and balance_hits >= 2 and col_hit:
                has_balance = True
                break

    # ── STEP 6: Date-pair column headers (no "Balance" keyword needed) ──
    if not has_balance:
        header_text = " ".join(lines[:25])
        _FULL_DATE_RE = re.compile(
            r"(?:Jan(?:uary)?|Feb(?:ruary)?|Mar(?:ch)?|Apr(?:il)?"
            r"|May|Jun(?:e)?|Jul(?:y)?|Aug(?:ust)?|Sep(?:t(?:ember)?)?"
            r"|Oct(?:ober)?|Nov(?:ember)?|Dec(?:ember)?)"
            r"\.?\s+\d{1,2}[,.]?\s*\d{4}",
            re.IGNORECASE,
        )
        date_matches = _FULL_DATE_RE.findall(header_text)
        if len(date_matches) >= 2:
            # ── CHANGED: require >=2 DISTINCT strict column keywords,
            # not just one loose "Issued"/"Retired" hit ──
            if _count_debt_column_keywords(header_text) >= 2:
                has_balance = True

    # ── STEP 8: "Beginning Ending" on one line + "Balance" on adjacent line ──
    if not has_balance:
        for i in range(len(lines) - 1):
            line_a = lines[i]
            line_b = lines[i + 1] if i + 1 < len(lines) else ""
            has_beg = bool(re.search(r"\bBeginning\b", line_a, re.IGNORECASE))
            has_end = bool(re.search(r"\bEnding\b", line_a, re.IGNORECASE))
            has_bal_next = bool(re.search(r"\bBalance\b", line_b, re.IGNORECASE))
            if has_beg and has_end and has_bal_next:
                has_balance = True
                break

    # ── STEP 9: Date-pair header with change + due/one-year keywords ──
    if not has_balance:
        _DATE_TOKEN_RE = re.compile(
            r"(?:Jan(?:uary)?|Feb(?:ruary)?|Mar(?:ch)?|Apr(?:il)?"
            r"|May|Jun(?:e)?|Jul(?:y)?|Aug(?:ust)?|Sep(?:t(?:ember)?)?"
            r"|Oct(?:ober)?|Nov(?:ember)?|Dec(?:ember)?)\.?\s+\d{1,2}[,.]?\s*\d{2,4}",
            re.IGNORECASE,
        )
        _CHANGE_HDR_RE = re.compile(
            r"\b(?:Additions?|Reductions?|Deletions?|Increases?|Decreases?|Maturities"
            r"|Issued|Retired|Retirements?|Payments?|Repaid|Repayments?|Proceeds?)\b",
            re.IGNORECASE,
        )
        _DUE_HDR_RE = re.compile(
            r"\bDue\s+[Ww]ithin\b|\b[Ww]ithin\s+[Oo]ne\s+[Yy]ear\b"
            r"|\bDue\s+[Ii]n\s+[Ll]ess\s+[Tt]han\b|\b[Ll]ess\s+[Tt]han\s+[Oo]ne\s+[Yy]ear\b"
            r"|\bOne\s+Year\b",
            re.IGNORECASE,
        )
        header30 = " ".join(lines[:30])
        if (len(_DATE_TOKEN_RE.findall(header30)) >= 2
                and _CHANGE_HDR_RE.search(header30)
                and _DUE_HDR_RE.search(header30)):
            has_balance = True

    # ── STEP 10: Transposed schedule — row labels in first column ───────
    # Detects pages where debt types are COLUMNS and the change-steps
    # (Balance beginning / Increases / Decreases / Balance end /
    #  Due within one year) are ROW LABELS at the start of each line.
    if not has_balance:
        _ROW_LABEL_RES = [
            re.compile(r"^\s*Balance,?\s+(?:at\s+the\s+)?beginning\s+of\s+"
                    r"(?:the\s+)?(?:fiscal\s+)?year\b", re.IGNORECASE),
            re.compile(r"^\s*Balance,?\s+(?:at\s+the\s+)?end\s+of\s+"
                    r"(?:the\s+)?(?:fiscal\s+)?year\b", re.IGNORECASE),
            re.compile(r"^\s*(?:Increases?|Additions?|Issued|Issuances?)\b",
                    re.IGNORECASE),
            re.compile(r"^\s*(?:Decreases?|Reductions?|Retired|"
                    r"Retirements?|Repaid|Repayments?|Payments?)\b",
                    re.IGNORECASE),
            re.compile(r"^\s*Due\s+within\s+one\s+year\b", re.IGNORECASE),
        ]
        _row_label_hits = 0
        for line in lines:
            for rx in _ROW_LABEL_RES:
                if rx.search(line):
                    _row_label_hits += 1
                    break
        if _row_label_hits >= 3:
            has_balance = True
            is_transposed_vertical = True


    # ── STEP 11: Outstanding debt summary table ─────────────────────────
    # Detects "debt issued summary" tables that have NO change columns
    # (no Additions/Reductions), only Interest Rate / Final Maturity /
    # Original Amount / Balance columns.
    #
    # Important: pypdf often STACKS multi-word column headers vertically,
    # e.g.
    #     Original
    #     Interest   Final     Amount      Balance
    #     Type       Rate      Maturity    of Issue   6-30-24
    # When flattened, "Interest" appears on a different line than "Rate".
    # So phrase matching like \bInterest\s+Rate\b fails. We use a
    # bag-of-words approach instead: look for the individual column-header
    # tokens anywhere in the first 40 lines, and require multiple co-occurrences.
    is_outstanding_summary = False
    if not has_balance:
        header_area = "\n".join(lines[:40])

        has_interest = bool(re.search(r"\bInterest\b",     header_area, re.IGNORECASE))
        has_rate     = bool(re.search(r"\bRate\b",         header_area, re.IGNORECASE))
        has_final    = bool(re.search(r"\bFinal\b",        header_area, re.IGNORECASE))
        has_maturity = bool(re.search(r"\bMaturity\b",     header_area, re.IGNORECASE))
        has_original = bool(re.search(r"\bOriginal\b",     header_area, re.IGNORECASE))
        has_amount   = bool(re.search(r"\bAmount\b",       header_area, re.IGNORECASE))
        has_issue    = bool(re.search(r"\b(?:of\s+)?Issue\b", header_area, re.IGNORECASE))
        has_bal_hdr  = bool(re.search(r"\bBalance\b",      header_area, re.IGNORECASE))

        # Score column-header signals (each pair OR keyword = 1 signal)
        sig = 0
        if has_interest and has_rate:      sig += 1   # "Interest Rate" column
        if has_final    and has_maturity:  sig += 1   # "Final Maturity" column
        if has_original and has_amount:    sig += 1   # "Original Amount" column
        if has_issue:                      sig += 1   # "of Issue" / "Issue Amount"
        if has_bal_hdr:                    sig += 1   # "Balance" column

        # Debt context (always required — keeps us out of unrelated tables)
        has_debt_ctx = bool(_OUTSTANDING_DEBT_CONTEXT_RE.search(text or ""))

        # Numeric density: rows with currency amounts (e.g. "$ 21,990,000")
        money_rows = sum(
            1 for l in lines
            if len(re.findall(r"\$\s*\d{1,3}(?:,\d{3})+", l)) >= 1
        )

        # Need at least 3 column-header signals + debt context + 2 money rows
        if sig >= 3 and has_debt_ctx and money_rows >= 2:
            is_outstanding_summary = True
            has_balance = True   # let the rest of the pipeline accept it


    # ── ALL balance detection steps exhausted ────────────────────────
    if not has_balance:
        return False

    
    # ── At least one change/due column (NOT required for outstanding summaries) ──
    has_change = _count_debt_column_keywords(text) >= 1 

    has_due = bool(re.search(
        r"\bDue\s+[Ww]ithin\b"
        r"|\b[Ww]ithin\s+[Oo]ne\s+[Yy]ear\b"
        r"|\bCurrent\s+Portion\b"
        r"|\bAmounts?\s+Due\b"
        r"|\b(?:Non[-\s]?current|Noncurrent)\b"
        r"|\bDue\s+in\s+[Mm]ore\s+[Tt]han\b"
        r"|\bDue\s+[Ii]n\s+[Ll]ess\s+[Tt]han\b"
        r"|\b[Ll]ess\s+[Tt]han\s+[Oo]ne\s+[Yy]ear\b"
        r"|\bDue\s+[Aa]fter\b",
        text or "", re.IGNORECASE))

    # Outstanding-summary tables don't have change/due columns — that's OK.
    # Only enforce this requirement for changes-schedule tables.
    if not is_outstanding_summary:
        if not (has_change or has_due):
            return False

    # ── Numeric density check ─────────────────────────────────────────
    # Changes-schedules: ≥ 4 rows with ≥ 3 formatted numbers
    #   (Beginning Bal | Additions | Reductions | Ending Bal | Due Within)
    # Outstanding summaries: ≥ 2 rows with ≥ 2 formatted numbers
    #   (Original Amount | Current Balance) — that's all an outstanding
    #   summary actually has per row.
    if is_outstanding_summary:
        rows_with_2plus_nums = sum(
            1 for l in lines
            if len(re.findall(r"\d{1,3}(?:,\d{3})+", l)) >= 2
        )
        return rows_with_2plus_nums >= 2

    # ADD THIS NEW BRANCH
    if is_transposed_vertical:
        lines_with_1plus_num = sum(
            1 for l in lines
            if len(re.findall(r"\d{1,3}(?:,\d{3})+", l)) >= 1
        )
        return lines_with_1plus_num >= 4

    # ── Standard changes-schedule: either the classic 3-row check,
    #    OR a single dense row (small table with only 1-2 liability types) ──
    lines_with_2plus_nums = sum(
        1 for l in lines
        if len(re.findall(r"\d{1,3}(?:,\d{3})+", l)) >= 2
    )
    if lines_with_2plus_nums >= 3:
        return True

    # Fallback: a single very dense data row (≥6 grouped numbers) is enough
    # to confirm a full Balance/Additions/Retirements/Balance/Due-Within row,
    # even if it's the only line item in this fund's debt schedule.
    max_nums_in_one_line = max(
        (len(re.findall(r"\d{1,3}(?:,\d{3})+", l)) for l in lines),
        default=0,
    )
    return max_nums_in_one_line >= 6


def _is_capital_asset_page(text: str) -> bool:
    """
    Return True if the page is a Capital Assets activity table.
    """
    capital_markers = [
        re.compile(r"\bCapital\s+assets?\s+(?:not\s+being|being)\s+depreciat", re.IGNORECASE),
        re.compile(r"\bCapital\s+asset\s+activity\b", re.IGNORECASE),
        re.compile(r"\bTotal\s+accumulated\s+depreciation\b", re.IGNORECASE),
        re.compile(r"\bLess\s+accumulated\s+depreciation\b", re.IGNORECASE),
        re.compile(r"\bDepreciation\s+expense\s+was\s+charged\b", re.IGNORECASE),
        re.compile(r"\bcapital\s+assets?,?\s+net\b", re.IGNORECASE),
        re.compile(r"\bTotal\s+capital\s+assets\b", re.IGNORECASE),
        re.compile(r"\bConstruction\s+in\s+pr[oc][gc]ess\b", re.IGNORECASE),
        re.compile(r"\bInfrastructure\b", re.IGNORECASE),
        re.compile(r"\bMachinery,?\s+equip", re.IGNORECASE),
        re.compile(r"\bBuildings?\s+and\s+improvements?\b", re.IGNORECASE),
    ]

    capital_hits = sum(1 for m in capital_markers if m.search(text or ""))

    if capital_hits >= 3:
        has_debt_header = bool(re.search(
            r"\bLong[-\s\u2013\u2014]Term\s+(?:Debt\b|Liabilit|Obligations\b)"
            r"|\bChanges?\s+in\s+Long[-\s\u2013\u2014]Term\s+(?:Debt\b|Liabilit)"
            r"|\bNoncurrent\s+Liabilit",
            text or "", re.IGNORECASE
        ))
        if not has_debt_header:
            return True

    return False
def _extract_text_layout(pdf_path: str, page_no: int) -> str:
    """
    Extract one page using pypdf (cross-platform, already a dependency).
    pypdf handles character-spaced fonts far better than pdfplumber
    for wide multi-column financial tables.
    """
    reader = _get_cached_pdf_reader(pdf_path)
    return reader.pages[page_no - 1].extract_text() or ""

def _join_header_lines(text: str, n_lines: int = 15) -> str:
    """
    Join first n lines with spaces to handle split table headers.
    
    Problem: Some PDFs have table headers split across multiple lines:
      Line 1: "Balance  Retirements/  Balance  Due Within"
      Line 2: "July 1, 2024  Additions  Deletions  June 30, 2025  One Year"
    
    When split like this, regex patterns fail to match "Balance July 2024" pattern
    because they're on different lines. This function joins them with spaces
    so the pattern matching works correctly.
    """
    lines = (text or "").splitlines()
    if len(lines) <= n_lines:
        return text
    header = " ".join(lines[:n_lines])
    rest = "\n".join(lines[n_lines:])
    return header + "\n" + rest

def _is_pension_opeb_note_page(text: str) -> bool:
    """
    Return True if the page is a pension / OPEB note page, NOT a DEBT
    changes schedule.

    Problem: Pension deferrals tables and OPEB plan descriptions contain
    words like "Retirement" (from plan names), "premium" (OPEB
    contributions), and fiscal-year dates that trick the DEBT detector
    into thinking the page is a debt changes table.

    Solution: Count pension/OPEB markers vs. debt-specific markers.
    If pension/OPEB dominates and no real debt header exists → exclude.
    """
    _PENSION_OPEB_MARKERS = [
        re.compile(r"\bDeferred\s+(?:Outflows|Inflows)\s+of\s+Resources\b", re.I),
        re.compile(r"\bPension\s+(?:Deferrals?|Expense|Plan)\b", re.I),
        re.compile(r"\bNet\s+Pension\s+(?:Liability|Asset)\b", re.I),
        re.compile(r"\b(?:Total|Net)\s+OPEB\s+Liability\b", re.I),
        re.compile(r"\bPost[-\s]?[Ee]mployment\s+(?:Benefit|Healthcare)\b", re.I),
        re.compile(r"\bActuarial\s+[Aa]ssumptions?\b", re.I),
        re.compile(r"\bFiduciary\s+[Nn]et\s+[Pp]osition\b", re.I),
        re.compile(r"\bDiscount\s+[Rr]ate\b", re.I),
        re.compile(r"\bMortality\s+[Rr]ates?\b", re.I),
        re.compile(r"\bHealthcare\s+(?:Benefits?|Cost)\s+(?:Plan|Trend)\b", re.I),
        re.compile(r"\bContribution\s+[Dd]eficiency\b", re.I),
        re.compile(r"\bProportionate\s+[Ss]hare\b", re.I),
        re.compile(r"\bRetirement\s+System\b", re.I),
        re.compile(r"\bSeparation\s+Allowance\b", re.I),
    ]

    # Debt-specific headers that mean this IS a real debt page
    # (even if it also mentions pension as a line item)
    _DEBT_SPECIFIC_HEADERS = re.compile(
        r"\bLong[-\s]?[Tt]erm\s+Obligation\s+Activity\b"
        r"|\bInstallment\s+(?:Financing|Purchase|Loan)\b"
        r"|\bCertificate[s]?\s+of\s+Participation\b"
        r"|\bLimited\s+Obligation\s+Bond\b"
        r"|\bFuture\s+[Mm]inimum\s+[Pp]ayments?\b"
        r"|\bDebt\s+[Rr]elated\s+to\s+Capital\b"
        r"|\bChanges?\s+in\s+.*Long[-\s]?[Tt]erm\s+(?:Debt|Obligation|Liabilit)"
        r"|\bSummary\s+of\s+Changes?\s+in\b.*\b(?:Debt|Obligation|Liabilit)"
        r"|\bInstallment\s+[Ff]inancing\b"
        r"|\b[Dd]ebt\s+[Ss]ervice\s+[Rr]equirements?\b",
        re.IGNORECASE | re.DOTALL,
    )

    pension_hits = sum(1 for p in _PENSION_OPEB_MARKERS if p.search(text or ""))
    has_debt_header = bool(_DEBT_SPECIFIC_HEADERS.search(text or ""))

    return pension_hits >= 3 and not has_debt_header

def find_debt_pages(pdf_path: str, start_page: int = START_PAGE) -> list:
    """
    Scan the ENTIRE PDF (no early stopping). Every page that independently
    satisfies the debt-table criteria (score + structure check) is collected,
    regardless of gaps caused by intervening narrative/DSR-style amortization
    pages, NOTE-number changes, or section breaks.
    """
    collected = []
    collected_set = set()

    with pdfplumber.open(pdf_path) as pdf:
        total = len(pdf.pages)

    for i in range(max(start_page, START_PAGE) - 1, total):
        page_no = i + 1

        if page_no in collected_set:
            continue

        text = _extract_text_layout(pdf_path, page_no)
        text = _normalize_debt_headers(text)

        first8 = "\n".join(text.splitlines()[:8])

        # ── Hard content exclusions still apply (these are NOT "early stop"
        #    on scanning — they just reject THIS page from being a DEBT
        #    table; the scan keeps going to the next page regardless) ──
        if _DEBT_HARD_EXCLUDE_RE.search(first8):
            continue
        if _OPEB_TABLE_RE.search(text):
            continue
        if re.search(r"\b[A-Z]\.\s+Receivables\b", text, re.IGNORECASE):
            continue
        if re.search(
            r"\bSchedule\s+of\s+Bonds?\s+Payable\b"
            r"|\bSchedule\s+of\s+Bond\s+Interest\b"
            r"|\bSchedule\s+of\s+Future\s+Debt\b"
            r"|\bDebt\s+Service\s+Schedule\b",
            text, re.IGNORECASE,
        ):
            if not re.search(r"\bNotes?\s+to\b.{0,30}\bFinancial\b", text, re.IGNORECASE):
                continue

        if _DEBT_LEASE_ASSET_EXCLUDE_RE.search(text):
            _has_debt_header = bool(re.search(
                r"\blong[-\s]term\s+(?:debt|obligation|liabilit)"
                r"|\blease\s+liabilit",
                text, re.IGNORECASE
            ))
            if not _has_debt_header:
                continue

        if _is_capital_asset_page(text):
            continue
        if _is_pension_opeb_note_page(text):
            continue
        if _is_claims_liability_page(text):
            continue

        score = _debt_score(text)
        if score < DEBT_SCORE_THRESHOLD:
            continue

        if not _page_has_debt_table_structure(text):
            continue

        # ── Year-row sanity check (same as before, just per-page now) ──
        _year_row_count = sum(
            1 for line in text.splitlines()
            if re.match(r"^\s*20[2-5]\d(?:\s*[-–—]\s*20[2-5]\d)?\s+", line.strip())
        )
        if _year_row_count >= 3:
            header_check_text = " ".join(text.splitlines()[:20])  # ✅ Simple join
            _has_debt_balance_header = bool(re.search(
                r"\b(?:Beginning|Ending|Opening|Closing)\s+Balance\b"
                r"|\bBalance,?\s+(?:at\s+)?(?:July|June|January|August|September|October"
                r"|November|December|February|March|April|[0-9]{1,2})\b"
                r"|\bBalance\s+[A-Za-z0-9\s,./\-–—]{0,50}(?:20\d{2}|\d{1,2})"
                r"|\bNet\s+Change\b|\bBalance\s+Outstanding\b"
                r"|\bDebt\s+Incurred\b|\bBond\s+Retired\b"
                r"|\bIncurred\b.{0,30}\bRetired\b"
                r"|\bBeginning\b[^\n]{0,40}\bEnding\b[^\n]{0,5}\n[^\n]{0,5}\bBalance\b",
                header_check_text, re.IGNORECASE,
            ))
            if not _has_debt_balance_header:
                continue

        # ── This page IS a DEBT table — collect it, then keep scanning ──
        collected.append(page_no)
        collected_set.add(page_no)
        print(f"    [DEBT] page {page_no}  (score={score})")

    collected = sorted(set(collected))

    if collected:
        print(f"    [DEBT] Final pages: {collected}")
    else:
        print(f"    [DEBT] ERROR: DEBT schedule not found in "
              f"{os.path.basename(pdf_path)}")

    return collected
# ══════════════════════════════════════════════════════════════════════════════
# SHARED UTILITIES
# ══════════════════════════════════════════════════════════════════════════════

def _pages_suffix(pages: list) -> str:
    if len(pages) == 1:
        return f"_p{pages[0]}"
    return "_p" + "-".join(str(p) for p in pages)


def extract_pages_to_pdf(src: str, page_numbers: list, out: str):
    reader = _get_cached_pdf_reader(src)      # was: PdfReader(src)
    writer = PdfWriter()
    for pn in page_numbers:
        writer.add_page(reader.pages[pn - 1])
    with open(out, "wb") as fh:
        writer.write(fh)


# ══════════════════════════════════════════════════════════════════════════════
# PROCESS ONE PDF
# ══════════════════════════════════════════════════════════════════════════════



# ══════════════════════════════════════════════════════════════════════════════
# PROCESS ONE PDF
# ══════════════════════════════════════════════════════════════════════════════

def process_file(src: str, out_dir: str = None, prompts_folder: str = None,
                 sector: str = "LG",
                 allowed_suffixes: set = None):
    """
    Extract relevant statement pages from one PDF.

    Parameters
    ----------
    src              : full path to the source PDF
    out_dir          : directory to write extracted statement PDFs
                       (defaults to same folder as src)
    prompts_folder   : (optional) path passed downstream; unused here
    sector           : "LG"  → extract 9 statements (SNP, SOA, GOV_BS, GOV_IS,
                              PROP_SNP, PROP_IS, PROP_CFS, DSR, DEBT)
                       "NON-LG" → extract 3 statements (PROP_SNP, PROP_IS, PROP_CFS)
    allowed_suffixes : optional set to restrict which suffixes get emitted
                       (e.g., {"PROP_SNP", "PROP_IS"}); None = all for sector
    """
    if out_dir is None:
        out_dir = os.path.dirname(src)
    os.makedirs(out_dir, exist_ok=True)

    base = os.path.splitext(os.path.basename(src))[0]
    print(f"Page Extraction is Running of [FILE] {os.path.basename(src)}  ||  [SECTOR] {sector}")

    produced = []   # list of output filenames written

    # ──────────────────────────────────────────────────────────────────
    # Helper to emit a single statement PDF
    # ──────────────────────────────────────────────────────────────────
    def _emit(suffix: str, pages: list, sector_tag: str):
        if not pages:
            return
        if allowed_suffixes is not None:
            normalized_allowed = {s.lstrip("_").upper() for s in allowed_suffixes}
            if suffix.upper() not in normalized_allowed:
                return
        out_name = f"{base}_{suffix}{_pages_suffix(pages)}.pdf"
        out_path = os.path.join(out_dir, out_name)
        try:
            extract_pages_to_pdf(src, pages, out_path)
            produced.append(out_path)   # ← FIX: full path, not bare name
            #print(f"      [{suffix:<10}] pages {pages}  →  {out_name}")
        except Exception as e:
            print(f"      [{suffix:<10}] ERROR writing {out_name}: {e}")

    # ══════════════════════════════════════════════════════════════════
    # NON-LG branch: only PROP_SNP, PROP_IS, PROP_CFS
    # ══════════════════════════════════════════════════════════════════
    if sector.upper() == "NON-LG":
        print(f"      [NON-LG] Extracting tables...")

        # Track first/last pages across statements for fallback hooks
        snp_pages_collected = []
        is_pages_collected  = []
        cfs_pages_collected = []

        # ── Pass 1: collect PROP_IS first (anchor for PROP_SNP fallback) ──
        prop_is_defn = next(d for d in NONLG_FUND_STATEMENTS if d["suffix"] == "PROP_IS")
        print(f"      [PROP_IS   ] Searching from page {START_PAGE} ...")

        try:
            is_pages_collected = find_fund_statement_pages(
                src, prop_is_defn, START_PAGE
            )
        except ValueError as e:
            print(f"      [PROP_IS   ] Not found — skipping ({e})")
            is_pages_collected = []


        # ── Pass 2: collect PROP_CFS (anchor for PROP_IS sandwich fallback) ──
        prop_cfs_defn = next(d for d in NONLG_FUND_STATEMENTS if d["suffix"] == "PROP_CFS")
        cfs_start = (is_pages_collected[-1] + 1) if is_pages_collected else START_PAGE
        print(f"      [PROP_CFS  ] Searching from page {cfs_start} ...")
        
        try:
            cfs_pages_collected = find_fund_statement_pages(
                src, prop_cfs_defn, START_PAGE
            )
        except ValueError as e:
            print(f"      [PROP_CFS  ] Not found — skipping ({e})")
            cfs_pages_collected = []


        # ── NEW: PROP_CFS forward-lookahead for component-unit cash flows ──
        if cfs_pages_collected:
            extra_cfs = find_prop_cfs_forward_lookahead(
                src, cfs_pages_collected[-1], prop_cfs_defn, lookahead=4
            )
            if extra_cfs:
                for p in extra_cfs:
                    if p not in cfs_pages_collected:
                        cfs_pages_collected.append(p)
                #print(f"      [PROP_CFS  ] forward-lookahead added pages "
                      #f"{extra_cfs}  →  component unit detected")

        # ── Pass 3: collect PROP_SNP (with reverse-lookback if missing) ──
        prop_snp_defn = next(d for d in NONLG_FUND_STATEMENTS if d["suffix"] == "PROP_SNP")
        print(f"      [PROP_SNP  ] Searching from page {START_PAGE} ...")

        try:
            snp_pages_collected = find_fund_statement_pages(
                src, prop_snp_defn, START_PAGE
            )
        except ValueError as e:
            print(f"      [PROP_SNP  ] Not found — skipping ({e})")
            snp_pages_collected = []


        if not snp_pages_collected and is_pages_collected:
            prop_is_first = is_pages_collected[0]
            print(f"      [PROP_SNP  ] No direct match — running reverse-lookback "
                  f"from PROP_IS p{prop_is_first} ...")
            fallback_snp = find_prop_snp_reverse_lookback(
                src, prop_is_first, prop_snp_defn, lookback=4
            )
            if fallback_snp:
                snp_pages_collected = fallback_snp
                print(f"      [PROP_SNP  ] reverse-lookback → pages {fallback_snp}")
            else:
                print(f"      [PROP_SNP  ] No pages found — skipping")

        # ── PROP_IS sandwich fallback (between SNP and CFS) ──
        if not is_pages_collected and snp_pages_collected and cfs_pages_collected:
            snp_last  = snp_pages_collected[-1]
            cfs_first = cfs_pages_collected[0]
            print(f"      [PROP_IS   ] No direct match — running sandwich "
                  f"between p{snp_last} and p{cfs_first} ...")
            sandwich_is = find_prop_is_sandwich(src, snp_last, cfs_first)
            if sandwich_is:
                is_pages_collected = sandwich_is
                print(f"      [PROP_IS   ] sandwich → pages {sandwich_is}")

        # ── Report any still-missing statements ──
        if not snp_pages_collected:
            print(f"      [PROP_SNP  ] ERROR: PROP_SNP not found in {os.path.basename(src)}")
            print(f"      [PROP_SNP  ] No pages found — skipping")
        if not is_pages_collected:
            print(f"      [PROP_IS   ] ERROR: PROP_IS not found in {os.path.basename(src)}")
            print(f"      [PROP_IS   ] No pages found — skipping")
        if not cfs_pages_collected:
            print(f"      [PROP_CFS  ] ERROR: PROP_CFS not found in {os.path.basename(src)}")
            print(f"      [PROP_CFS  ] No pages found — skipping")

        # ── Emit ──
        _emit("PROP_SNP", snp_pages_collected, "NON-LG")
        _emit("PROP_IS",  is_pages_collected,  "NON-LG")
        _emit("PROP_CFS", cfs_pages_collected, "NON-LG")

        return produced

    # ══════════════════════════════════════════════════════════════════
    # LG branch: full 9-statement extraction
    # ══════════════════════════════════════════════════════════════════
    print(f"      [LG] Extracting tables...")

    # ── 1. Government-Wide SNP ──
    print(f"      [SNP       ] Searching from page {START_PAGE} ...")
    snp_pages = find_snp_pages(src)
    if not snp_pages:
        print(f"      [SNP       ] No pages found — skipping")

    # ── 2. Government-Wide SOA (start scanning after SNP, if any) ──
    soa_start = (snp_pages[-1] + 1) if snp_pages else START_PAGE
    print(f"      [SOA       ] Searching from page {soa_start} ...")
    soa_pages = find_soa_pages(src, start_page=soa_start)
    if not soa_pages:
        print(f"      [SOA       ] No pages found — skipping")

    _emit("SNP", snp_pages, "LG")
    _emit("SOA", soa_pages, "LG")
    has_prop = _has_proprietary_fund_statements(src)
    if not has_prop:
        print(f"      [PROP_*    ] No proprietary/enterprise funds detected — "
          f"skipping PROP_SNP / PROP_IS / PROP_CFS")
    # ── 3-7. Fund-level statements (GOV_BS, GOV_IS, PROP_SNP, PROP_IS, PROP_CFS) ──
    # Track anchors across ALL fund statements for DEBT floor logic
    # Track results for proper chaining
    gov_bs_pages   = []
    gov_is_pages   = []
    prop_snp_pages = []
    prop_is_pages  = []
    prop_cfs_pages = []

    for defn in FUND_STATEMENTS:
        suffix = defn["suffix"]
        
        if suffix in ("PROP_SNP", "PROP_IS", "PROP_CFS") and not has_prop:
            continue
        
        # ✅ FIX: Chain from LAST found statement
        prior_pages = (gov_bs_pages or gov_is_pages or
                       prop_snp_pages or prop_is_pages or prop_cfs_pages)
        stmt_start = (prior_pages[-1] + 1) if prior_pages else START_PAGE
        
        print(f"      [{suffix:<10}] Searching from page {stmt_start} ...")
        try:
            pages = find_fund_statement_pages(src, defn, stmt_start)
        except ValueError:
            pages = []
            print(f"      [{suffix:<10}] Not found — skipping")
        
        # ✅ FIX: Cache results for downstream chaining
        if suffix == "GOV_BS":
            gov_bs_pages = pages
        elif suffix == "GOV_IS":
            gov_is_pages = pages
        elif suffix == "PROP_SNP":
            prop_snp_pages = pages
        elif suffix == "PROP_IS":
            prop_is_pages = pages
        elif suffix == "PROP_CFS":
            prop_cfs_pages = pages
        
        _emit(suffix, pages, "LG")

    # ── 8. DSR ──
    print(f"      [DSR       ] Searching ...")
    dsr_pages = find_dsr_pages(src)
    if not dsr_pages:
        print(f"      [DSR       ] No pages found — skipping")
    _emit("DSR", dsr_pages, "LG")

    # ── 9. DEBT ──
    # Floor = max last-page across EVERY fund-statement anchor found so far,
    # not just the PROP_* ones. This guarantees that if PROP_CFS (or any
    # single statement) fails to detect, the scan still doesn't fall back
    # to START_PAGE and pick up MD&A pages before the financial statements.
    debt_start = max(
        (prop_cfs_pages[-1] if prop_cfs_pages else 0),
        (prop_is_pages[-1]  if prop_is_pages  else 0),
        (prop_snp_pages[-1] if prop_snp_pages else 0),
        (gov_is_pages[-1]   if gov_is_pages   else 0),
        (gov_bs_pages[-1]   if gov_bs_pages   else 0),
        (soa_pages[-1]      if soa_pages      else 0),
        (snp_pages[-1]      if snp_pages      else 0),
    ) + 1
    debt_start = max(debt_start, START_PAGE)

    print(f"      [DEBT      ] Searching from page {debt_start} ...")
    debt_pages = find_debt_pages(src, start_page=debt_start)
    if not debt_pages:
        print(f"      [DEBT      ] No pages found — skipping")
    _emit("DEBT", debt_pages, "LG")

    return produced

# ══════════════════════════════════════════════════════════════════════════════
# FOLDER RUNNER
# ══════════════════════════════════════════════════════════════════════════════
_PRODUCED_RE = re.compile(
    r"_(?:SNP|SOA|GOV_BS|GOV_IS|PROP_SNP|PROP_IS|PROP_CFS|DSR|DEBT)_p\d*(?:-\d+)*\.pdf$",
    re.IGNORECASE,
)
def process_file_llm_guided(
    src: str,
    out_dir: str = None,
    sector: str = "LG",
    allowed_suffixes: set = None,
    id_model: str = "gemini-2.5-flash",
    use_llm_id: bool = True,
    file_index: int = None,   # ← ADD
    file_total: int = None,   # ← ADD
) -> list:
    """
    LLM-guided extraction:
    1. Send full PDF to small model for page identification (one call)
    2. Use identified page numbers to slice PDFs with extract_pages_to_pdf()
    3. Falls back to Python keyword detection per-table if LLM misses any
    """
    from llm_page_identifier import identify_pages_with_fallback

    if out_dir is None:
        out_dir = os.path.dirname(src)
    os.makedirs(out_dir, exist_ok=True)

    base = os.path.splitext(os.path.basename(src))[0]
    if file_index is not None and file_total is not None:
        file_label = f"FILE {file_index}/{file_total}"
    else:
        file_label = f"FILE"

    print(f"Page Extraction is Running of [{file_label}] {os.path.basename(src)}  ||  [SECTOR] {sector} .......")
    print(f"\n{'='*70}")
    produced = []

    def _emit(suffix: str, pages: list):
        if not pages:
            return
        if allowed_suffixes is not None:
            normalized_allowed = {s.lstrip("_").upper() for s in allowed_suffixes}
            if suffix.upper() not in normalized_allowed:
                return
        out_name = f"{base}_{suffix}{_pages_suffix(pages)}.pdf"
        out_path = os.path.join(out_dir, out_name)
        try:
            extract_pages_to_pdf(src, pages, out_path)
            produced.append(out_path)
            #print(f"  [{suffix:<10}] pages {pages}  →  {out_name}")
        except Exception as e:
            print(f"  [{suffix:<10}] ERROR: {e}")

    # ── Determine which suffixes to extract for this sector ──
    if sector.upper() == "NON-LG":
        target_suffixes = ["PROP_SNP", "PROP_IS", "PROP_CFS"]
    else:
        target_suffixes = ["SNP", "SOA", "GOV_BS", "GOV_IS",
                           "PROP_SNP", "PROP_IS", "PROP_CFS", "DSR", "DEBT"]

    if allowed_suffixes:
        normalized_allowed = {s.lstrip("_").upper() for s in allowed_suffixes}
        target_suffixes = [s for s in target_suffixes if s in normalized_allowed]

    # ── Step 1: LLM page identification (one call for whole PDF) ──
    llm_pages: dict[str, list[int]] = {}
    if use_llm_id:
        # Build a fallback that runs Python detection for a specific suffix
        llm_pages = identify_pages_with_fallback(
            pdf_path=src,
            model=id_model,
            fallback_fn=None,  # we handle per-suffix fallback below
            allowed_suffixes=allowed_suffixes,   # ← ADD THIS LINE
        )

    # ── Step 2: For each target suffix, use LLM result or Python fallback ──
    python_fallback_fns = {
        "SNP":      lambda: find_snp_pages(src),
        "SOA":      lambda: find_soa_pages(src),
        "GOV_BS":   lambda: find_fund_statement_pages(
                        src,
                        next(d for d in FUND_STATEMENTS if d["suffix"] == "GOV_BS"),
                        START_PAGE),
        "GOV_IS":   lambda: find_fund_statement_pages(
                        src,
                        next(d for d in FUND_STATEMENTS if d["suffix"] == "GOV_IS"),
                        START_PAGE),
        "PROP_SNP": lambda: find_fund_statement_pages(
                        src,
                        next(d for d in (FUND_STATEMENTS if sector == "LG"
                             else NONLG_FUND_STATEMENTS) if d["suffix"] == "PROP_SNP"),
                        START_PAGE),
        "PROP_IS":  lambda: find_fund_statement_pages(
                        src,
                        next(d for d in (FUND_STATEMENTS if sector == "LG"
                             else NONLG_FUND_STATEMENTS) if d["suffix"] == "PROP_IS"),
                        START_PAGE),
        "PROP_CFS": lambda: find_fund_statement_pages(
                        src,
                        next(d for d in (FUND_STATEMENTS if sector == "LG"
                             else NONLG_FUND_STATEMENTS) if d["suffix"] == "PROP_CFS"),
                        START_PAGE),
        "DSR":      lambda: find_dsr_pages(src),
        "DEBT":     lambda: find_debt_pages(src),
    }

    for suffix in target_suffixes:
        pages = llm_pages.get(suffix, [])

        if pages:
            print(f"  [{suffix:<10}] LLM identified: pages {pages}")
            _emit(suffix, pages)
        else:
            try:
                fb = python_fallback_fns.get(suffix)
                pages = fb() if fb else []
                if pages:
                    print(f"  [{suffix:<10}] Python fallback: pages {pages}")
                    _emit(suffix, pages)
                else:
                    print(f"  [{suffix:<10}] Not found by LLM or Python — skipping")
            except ValueError:
                print(f"  [{suffix:<10}] Not found by LLM or Python — skipping")
            except Exception as e:
                print(f"  [{suffix:<10}] Not found — fallback error: {e}")           

    return produced

def process_folder(folder: str, out_dir=None):
    if out_dir is None:
        out_dir = folder

    for fname in sorted(os.listdir(folder)):
        if not fname.lower().endswith(".pdf"):
            continue
        if _PRODUCED_RE.search(fname):
            continue

        src = os.path.join(folder, fname)

        # ── Detect sector from filename ──
        sector = "LG"  # default
        if re.search(r'NON[-_]?LG', fname, re.IGNORECASE):
            sector = "NON-LG"

        #print(f"\n{'='*70}")
        print(f"  File  : {fname}" + "  Sector: {sector}")
        
        #print(f"{'='*70}")

        process_file(src, out_dir=out_dir, sector=sector)


if __name__ == "__main__":
    
    script_dir     = os.path.dirname(os.path.abspath(__file__))
    default_folder = os.path.join(script_dir, "03_Validated_Report")

    if len(sys.argv) >= 3:
        folder  = os.path.abspath(sys.argv[1])
        out_dir = os.path.abspath(sys.argv[2])
    elif len(sys.argv) == 2:
        folder  = os.path.abspath(sys.argv[1])
        out_dir = folder
    else:
        folder  = default_folder
        out_dir = default_folder

    process_folder(folder, out_dir)
