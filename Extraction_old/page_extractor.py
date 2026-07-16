"""
extract_all_statements.py
=========================
Extracts all 7 financial statement sections from government CAFR/ACFR PDFs
in a single pass:

  GW Statements (Government-Wide):
    1. SNP     – Statement of Net Position (Government-Wide)
    2. SOA     – Statement of Activities

  Fund-Level Statements:
    3. GOV_BS   – Balance Sheet – Governmental Funds
    4. GOV_IS   – Statement of Revenues, Expenditures, and Changes in Fund Balances – Governmental Funds
    5. PROP_SNP – Statement of Net Position – Proprietary Funds
    6. PROP_IS  – Statement of Revenues, Expenses, and Changes in Fund Net Position – Proprietary Funds
    7. PROP_CFS – Statement of Cash Flows – Proprietary Funds

Search order:
  SNP → SOA → GOV_BS → GOV_IS → PROP_SNP → PROP_IS → PROP_CFS
  Each search starts AFTER the last page of the previous anchor (where applicable).

Usage
-----
  python extract_all_statements.py [input_folder] [output_folder]

  If no folder given, defaults to ./03_Validated_Report next to the script.
"""

import os
import re
import sys

import pdfplumber
from pypdf import PdfReader, PdfWriter

# ── Global scan start (1-based) ───────────────────────────────────────────────
START_PAGE = 11


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
    re.compile(r"\bmd\s*(?:&|and)\s*a\b",                re.IGNORECASE),
    re.compile(r"\bmanagement[\u2018\u2019']?s?\s+discussion\s+(?:&|and)\s+analysis\b",
               re.IGNORECASE),
    re.compile(r"\bbudget\b",                             re.IGNORECASE),
    re.compile(r"\bstatistical\b",                        re.IGNORECASE),
    re.compile(r"\bcombining\b",                          re.IGNORECASE),
    re.compile(r"\bsupplemental\b",                       re.IGNORECASE),
    re.compile(r"\bschedule\b",                           re.IGNORECASE),
    re.compile(r"\breconciliation\b",                     re.IGNORECASE),
    # NOTE: "internal service fund" is intentionally NOT excluded here.
    # Some entities report their only proprietary fund statements as internal
    # service fund statements (e.g. school boards with self-insurance funds).
    # Excluding it would block those primary statements.
    # The combining/supplemental/schedule exclusions above already guard
    # against picking up internal-service-fund schedules in later sections.
    re.compile(r"\bcustodial\s+fund",                     re.IGNORECASE),
    re.compile(r"\bfiduciary\b",                          re.IGNORECASE),
    re.compile(r"\bnotes?\s+to\b",                        re.IGNORECASE),
]

EXCLUDE_HEADING_ONLY_PATTERNS = [
    re.compile(r"\bnon[-\s]?major\b", re.IGNORECASE),
    re.compile(r"\bpension\b",        re.IGNORECASE),
]

# GW-SNP/SOA use a tighter exclusion list (no "proprietary fund" allowed for GW)
SNP_EXCLUDE_PATTERNS = [
    re.compile(r"^\s*table\s+of\s+contents?\b", re.IGNORECASE),
    re.compile(r"\btable\s+of\s+contents?\b",   re.IGNORECASE),
    re.compile(r"\bforeword\b",                 re.IGNORECASE),
    re.compile(r"\bintroduction\b",             re.IGNORECASE),
    re.compile(r"\bmd\s*(?:&|and)\s*a\b",      re.IGNORECASE),
    re.compile(r"\bmanagement[\u2018\u2019']?s?\s+discussion\s+(?:&|and)\s+analysis\b",
               re.IGNORECASE),
    re.compile(r"\bbudget\b",                   re.IGNORECASE),
    re.compile(r"\bsupplemental\b",             re.IGNORECASE),
    re.compile(r"\bproprietary\s+fund\b",       re.IGNORECASE),
    re.compile(r"\bstatistical\b",              re.IGNORECASE),
    re.compile(r"\bnon[-\s]?major\b",           re.IGNORECASE),
]


# ══════════════════════════════════════════════════════════════════════════════
# GOVERNMENT-WIDE SNP
# ══════════════════════════════════════════════════════════════════════════════
SNP_TITLE_RE = re.compile(
    r"^\s*(?:"
    r"statement\s+of\s+net\s+pos(?:i)?tion"
    r"|government[-\s]*wide\s*[-–—]?\s*statement\s+of\s+net\s+pos(?:i)?tion"
    r")\b",
    re.IGNORECASE | re.MULTILINE,
)

SNP_FINANCIAL_ROW_KEYWORDS = [
    re.compile(r"\btotal\s+assets\b",   re.IGNORECASE),
    re.compile(r"\btotal\s+liabilit",   re.IGNORECASE),
    re.compile(r"\bnet\s+position\b",   re.IGNORECASE),
]

SNP_BODY_CONTINUATION_RE = re.compile(
    r"^\s*(?:deferred\s+(?:inflows|outflows)|net\s+position\b)",
    re.IGNORECASE | re.MULTILINE,
)

SNP_COMPONENT_UNIT_PAGE_START_RE = re.compile(
    r"^\s*component\s+unit",
    re.IGNORECASE | re.MULTILINE,
)


def _snp_excluded(text: str) -> bool:
    first_lines = "\n".join((text or "").splitlines()[:10])
    return any(p.search(first_lines) for p in SNP_EXCLUDE_PATTERNS)


def _snp_title_in_header(text: str) -> bool:
    for line in (text or "").splitlines()[:25]:
        if SNP_TITLE_RE.match(line.strip()):
            return True
    return False


def _all_snp_financial_rows(text: str) -> bool:
    return all(p.search(text or "") for p in SNP_FINANCIAL_ROW_KEYWORDS)


def _any_snp_financial_row(text: str) -> bool:
    return any(p.search(text or "") for p in SNP_FINANCIAL_ROW_KEYWORDS)


def find_snp_pages(pdf_path: str) -> list[int]:
    """Find government-wide Statement of Net Position pages."""
    collected: list[int] = []

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

            j = collected[-1]
            while j < total:
                nt = pdf.pages[j].extract_text() or ""
                if _snp_excluded(nt):
                    j += 1; continue
                if _snp_title_in_header(nt):
                    break
                first4 = "\n".join(nt.splitlines()[:4])
                if (bool(SNP_COMPONENT_UNIT_PAGE_START_RE.search(first4)) or
                        bool(SNP_BODY_CONTINUATION_RE.search("\n".join(nt.splitlines()[:15])))):
                    collected.append(j + 1)
                else:
                    break
                j += 1

            break  # first occurrence only

    if not collected:
        raise ValueError(f"Government-wide SNP not found in {pdf_path}")
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


def _soa_title_in_header(text: str) -> bool:
    lines = (text or "").splitlines()
    return bool(SOA_TITLE_RE.search("\n".join(lines[:25])))


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


def find_soa_pages(pdf_path: str, start_page: int = START_PAGE) -> list[int]:
    """Find Statement of Activities pages, always starting after SNP."""
    collected: list[int] = []

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

            break  # first occurrence only

    if not collected:
        raise ValueError(f"Statement of Activities not found in {pdf_path}")
    return sorted(set(collected))


# ══════════════════════════════════════════════════════════════════════════════
# SCRAMBLED-TITLE DETECTION (for rotated sidebar PDFs)
# ══════════════════════════════════════════════════════════════════════════════
_STATEMENT_TITLE_RE = re.compile(
    r"\bStatement\s+of\b|\bReconciliation\b|\bSchedule\b",
    re.IGNORECASE,
)

# ── GOV_BS (Doc 1 logic) ──────────────────────────────────────────────────────
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
    """Page contains the closing 'Total fund balances' row — statement ends here."""
    return bool(_GOV_BS_COMPLETE_RE.search(text))

def _gov_bs_is_column_overflow(text: str) -> bool:
    lines = (text or "").splitlines()
    first_dollar = next((i for i, l in enumerate(lines) if "$" in l), 999)
    if first_dollar < 2:
        return False
    joined_top = " ".join(lines[:6])
    if _STATEMENT_TITLE_RE.search(joined_top):   # guard: new-page title → not overflow
        return False
    return (bool(_GOV_BS_TOTAL_COL_RE.search(joined_top))
            and not _gov_bs_is_complete(text))

# ── GOV_IS (Doc 1 logic) ──────────────────────────────────────────────────────
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
    """Page contains 'Fund balances - ending' or 'Net change in fund balances' — statement ends here."""
    return bool(_GOV_IS_COMPLETE_RE.search(text))

def _gov_is_is_column_overflow(text: str) -> bool:
    lines = (text or "").splitlines()
    first_dollar = next((i for i, l in enumerate(lines) if "$" in l), 999)
    if first_dollar < 2:
        return False
    joined_top = " ".join(lines[:6])
    if _STATEMENT_TITLE_RE.search(joined_top):   # guard: new-page title → not overflow
        return False
    return (bool(_GOV_IS_TOTAL_COL_RE.search(joined_top))
            and not _gov_is_is_complete(text))

# ── PROP_SNP (Doc 2 logic) ────────────────────────────────────────────────────
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
_PROP_SNP_OVERFLOW_HEADER_RE = re.compile(
    r"\bBusiness[-\s]*type\s+Activities\b",
    re.IGNORECASE,
)

def _prop_snp_is_complete(text: str) -> bool:
    return bool(_PROP_SNP_COMPLETE_RE.search(text))

def _prop_snp_is_column_overflow(text: str) -> bool:
    lines = (text or "").splitlines()
    first_dollar = next((i for i, l in enumerate(lines) if "$" in l), 999)
    # True overflow pages have only column headers before the first $;
    # real continuation pages (title + body) have $ much further down.
    if not (2 <= first_dollar <= 5):
        return False
    joined_top = " ".join(lines[:6])
    # A page with its own statement title is a new statement, not overflow
    if _STATEMENT_TITLE_RE.search(joined_top):
        return False
    return bool(_PROP_SNP_OVERFLOW_HEADER_RE.search(joined_top))

# ── PROP_IS (Doc 2 logic) ─────────────────────────────────────────────────────
PROP_IS_SCRAMBLED_KW = [
    re.compile(r"\boperating\s+revenues?\b", re.IGNORECASE),
    re.compile(r"\boperating\s+expenses?\b", re.IGNORECASE),
    re.compile(r"\bnet\s+position",          re.IGNORECASE),
    re.compile(r"\bproprietary\s+fund",      re.IGNORECASE),
]
_PROP_IS_COMPLETE_RE = re.compile(
    r"\bNet\s+position\s*[-–]\s*(?:end\w*|begin\w*|of\s+year)\b"
    r"|\bNet\s+position\s*,?\s*end\s+of\s+(?:fiscal\s+)?year\b"
    r"|\bChange\s+in\s+net\s+position\s+of\s+business",
    re.IGNORECASE,
)
_PROP_IS_OVERFLOW_HEADER_RE = re.compile(
    r"\bBusiness[-\s]*type\s+Activities\b",
    re.IGNORECASE,
)

def _prop_is_is_complete(text: str) -> bool:
    return bool(_PROP_IS_COMPLETE_RE.search(text))

def _prop_is_is_column_overflow(text: str) -> bool:
    lines = (text or "").splitlines()
    first_dollar = next((i for i, l in enumerate(lines) if "$" in l), 999)
    if not (2 <= first_dollar <= 5):
        return False
    joined_top = " ".join(lines[:6])
    if _STATEMENT_TITLE_RE.search(joined_top):
        return False
    return bool(_PROP_IS_OVERFLOW_HEADER_RE.search(joined_top))

# ── PROP_CFS (Doc 2 logic) ────────────────────────────────────────────────────
PROP_CFS_SCRAMBLED_KW = [
    re.compile(r"\bcash\s+(?:flows?|provided|used)\b", re.IGNORECASE),
    re.compile(r"\boperating\s+activit",               re.IGNORECASE),
    re.compile(r"\bproprietary\s+fund",                re.IGNORECASE),
]
_PROP_CFS_COMPLETE_RE = re.compile(
    r"\bCash\s+and\s+[Cc]ash\s+[Ee]quivalents?\s+(?:at\s+)?[Ee]nd\s+of\s+[Yy]ear\b"
    r"|\bNon[-\s]?cash\b.{0,50}(?:investing|capital|financial)"
    r"|\bNon[-\s]?cash\s+investing",
    re.IGNORECASE,
)
_PROP_CFS_OVERFLOW_HEADER_RE = re.compile(
    r"\bBusiness[-\s]*type\s+Activities\b",
    re.IGNORECASE,
)

def _prop_cfs_is_complete(text: str) -> bool:
    return bool(_PROP_CFS_COMPLETE_RE.search(text))

def _prop_cfs_is_column_overflow(text: str) -> bool:
    lines = (text or "").splitlines()
    first_dollar = next((i for i, l in enumerate(lines) if "$" in l), 999)
    if not (2 <= first_dollar <= 5):
        return False
    joined_top = " ".join(lines[:6])
    if _STATEMENT_TITLE_RE.search(joined_top):
        return False
    return bool(_PROP_CFS_OVERFLOW_HEADER_RE.search(joined_top))


def _is_scrambled_anchor(text: str, kw_list: list) -> bool:
    lines = (text or "").splitlines()
    single_char_lines = sum(1 for l in lines[:20] if len(l.strip()) <= 2)
    spaced_word_lines = sum(
        1 for l in lines[:20] if re.match(r"^([A-Za-z]\s){2,}", l.strip())
    )
    if not ((single_char_lines >= 3) or (spaced_word_lines >= 1)):
        return False
    return all(kw.search(text) for kw in kw_list)


# ══════════════════════════════════════════════════════════════════════════════
# SHARED COMPLETION HELPER  (Doc 2 — used by PROP_* only)
# ══════════════════════════════════════════════════════════════════════════════

def _complete_row_has_numbers(text: str, complete_re: re.Pattern) -> bool:
    """
    Returns True only when the completion row appears on a line that ALSO
    contains a numeric value — meaning the statement is truly done on this
    page and the numbers are not still sitting on a separate overflow page.

    GOV_BS / GOV_IS:
        Closing label row has no numbers on the left page → False →
        overflow page is collected. ✓  (these statements use Doc 1 logic
        and never call this helper.)

    PROP_IS / PROP_CFS / PROP_SNP:
        Closing row already carries its numbers on the same line
        → True → any following overflow page is skipped. ✓
    """
    for line in (text or "").splitlines():
        if complete_re.search(line):
            if re.search(r"[\d,]{3,}", line):
                return True
    return False


# ══════════════════════════════════════════════════════════════════════════════
# FUND-LEVEL STATEMENT DEFINITIONS
# ══════════════════════════════════════════════════════════════════════════════
FUND_STATEMENTS = [
    # ── GOV_BS: Doc 1 overflow logic (guard keeps _is_complete_fn / overflow
    #    separate; no complete_re so is_truly_complete falls back to plain
    #    is_complete_fn in the finder) ─────────────────────────────────────────
    {
        "suffix": "GOV_BS",
        "title_re": re.compile(
            r"(?:balance\s+sheet\b.{0,120}governmental\s+fund"
            r"|governmental\s+fund.{0,120}balance\s+sheet)",
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
            r"(?:balance\s+sheet\b.{0,60}(?:continued|cont\.?)"
            r"|governmental\s+fund.{0,60}balance\s+sheet.{0,60}(?:continued|cont\.?)"
            r"|\bfund\s+balance\b|\bliabilit|\btotal\s+assets\b)",
            re.IGNORECASE | re.DOTALL,
        ),
        "scrambled_cont_kw": [
            re.compile(r"\bfund\s+balance",              re.IGNORECASE),
            re.compile(r"\btotal\s+(?:assets|liabilit)", re.IGNORECASE),
        ],
        # Doc 1: overflow detection excludes pages that already contain
        # the closing row, so is_complete_fn acts as the true stop signal.
        "is_complete_fn":     _gov_bs_is_complete,
        "is_col_overflow_fn": _gov_bs_is_column_overflow,
        # No complete_re → finder uses plain is_complete_fn (Doc 1 behaviour).
        # gov_overflow=True → overflow collected unconditionally (guard is
        # baked into _gov_bs_is_column_overflow via `and not _is_complete`).
        "gov_overflow":       True,
    },
    # ── GOV_IS: Doc 1 overflow logic ─────────────────────────────────────────
    {
        "suffix": "GOV_IS",
        "title_re": re.compile(
            r"(?:statement\s+of\s+revenues,?\s+expenditures"
            r"|revenues,?\s+expenditures,?\s+and\s+changes\s+in\s+fund\s+balance"
            r").{0,200}governmental\s+fund"
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
        # Doc 1: same pattern as GOV_BS.
        "is_complete_fn":     _gov_is_is_complete,
        "is_col_overflow_fn": _gov_is_is_column_overflow,
        # No complete_re → finder uses plain is_complete_fn (Doc 1 behaviour).
        "gov_overflow":       True,
    },
    # ── PROP_SNP: Doc 2 logic ─────────────────────────────────────────────────
    {
        "suffix": "PROP_SNP",
        "title_re": re.compile(
            r"(?:statement\s+of\s+(?:fund\s+)?net\s+position.{0,120}proprietary"
            r"|proprietary\s+fund.{0,120}statement\s+of\s+(?:fund\s+)?net\s+position)",
            re.IGNORECASE | re.DOTALL,
        ),
        "require_kw": [
            re.compile(r"\bassets\b",       re.IGNORECASE),
            re.compile(r"\bnet\s+position", re.IGNORECASE),
        ],
        "scrambled_kw": PROP_SNP_SCRAMBLED_KW,
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
        # Doc 2: numbers on same line → True → overflow skipped.
        "complete_re":        _PROP_SNP_COMPLETE_RE,
    },
    # ── PROP_IS: Doc 2 logic ──────────────────────────────────────────────────
    {
        "suffix": "PROP_IS",
        "title_re": re.compile(
            r"(?:statement\s+of\s+revenues,?\s+expenses.*changes.*(?:fund\s+)?net\s+position.*proprietary"
            r"|proprietary.*statement\s+of\s+revenues,?\s+expenses.*changes.*(?:fund\s+)?net\s+position)",
            re.IGNORECASE | re.DOTALL,
        ),
        "require_kw": [
            re.compile(r"\bexpenses\b",     re.IGNORECASE),
            re.compile(r"\bnet\s+position", re.IGNORECASE),
        ],
        "scrambled_kw": PROP_IS_SCRAMBLED_KW,
        "stop_res": [
            (re.compile(r"\bstatement\s+of\s+cash\s+flow", re.IGNORECASE), 25),
        ],
        "cont_re": re.compile(
            r"(?:revenues,?\s+expenses.*changes.*net\s+position.*(?:continued|cont\.?)"
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
        # Doc 2: numbers on same line → True → overflow skipped.
        "complete_re":        _PROP_IS_COMPLETE_RE,
    },
    # ── PROP_CFS: Doc 2 logic ─────────────────────────────────────────────────
    {
        "suffix": "PROP_CFS",
        "title_re": re.compile(
            r"(?:statement\s+of\s+cash\s+flows?.{0,120}proprietary"
            r"|proprietary\s+fund.{0,120}statement\s+of\s+cash\s+flow)",
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
        # Doc 2: numbers on same line → True → overflow skipped.
        "complete_re":        _PROP_CFS_COMPLETE_RE,
    },
]


# ══════════════════════════════════════════════════════════════════════════════
# GENERIC FUND-LEVEL FINDER
# ══════════════════════════════════════════════════════════════════════════════

def _fund_excluded(text: str) -> bool:
    """Exclusion check for fund-level statements (scrambled pages bypass this)."""
    lines = (text or "").splitlines()
    single_char = sum(1 for l in lines[:20] if len(l.strip()) <= 2)
    spaced_word = sum(1 for l in lines[:20] if re.match(r"^([A-Za-z]\s){2,}", l.strip()))
    if (single_char >= 3) or (spaced_word >= 1):
        return False  # scrambled page — let keyword check decide
    first12 = "\n".join(lines[:12])
    if any(p.search(first12) for p in EXCLUDE_PATTERNS):
        return True
    first3 = "\n".join(lines[:3])
    return any(p.search(first3) for p in EXCLUDE_HEADING_ONLY_PATTERNS)


def _title_in_header(text: str, title_re: re.Pattern, n_lines: int = 25) -> bool:
    return bool(title_re.search("\n".join((text or "").splitlines()[:n_lines])))


def find_fund_statement_pages(
    pdf_path: str,
    defn: dict,
    start_page: int,
    next_anchor_re: re.Pattern | None = None,
) -> list[int]:
    title_re           = defn["title_re"]
    require_kw         = defn["require_kw"]
    stop_res           = defn["stop_res"]
    cont_re            = defn["cont_re"]
    scrambled_kw       = defn.get("scrambled_kw", [])
    scr_cont_kw        = defn.get("scrambled_cont_kw", [])
    is_complete_fn     = defn.get("is_complete_fn")
    is_col_overflow_fn = defn.get("is_col_overflow_fn")
    complete_re        = defn.get("complete_re")   # present for PROP_*, absent for GOV_*
    gov_overflow       = defn.get("gov_overflow", False)  # True for GOV_BS / GOV_IS

    # ── helpers ───────────────────────────────────────────────────────────────

    def is_anchor(text: str) -> bool:
        if _title_in_header(text, title_re):
            return all(kw.search(text or "") for kw in require_kw)
        if scrambled_kw:
            return _is_scrambled_anchor(text, scrambled_kw)
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
        single_char  = sum(1 for l in lines[:20] if len(l.strip()) <= 2)
        spaced_word  = sum(1 for l in lines[:20] if re.match(r"^([A-Za-z]\s){2,}", l.strip()))
        dollar_lines = sum(1 for l in lines if re.search(r"\$\s*[\d,]+", l))
        if not ((single_char >= 3) or (spaced_word >= 1) or (dollar_lines >= 3)):
            return False
        return all(kw.search(text) for kw in scr_cont_kw)

    _CONT_MARKER_RE = re.compile(r"\bcont(?:inued|\.)", re.IGNORECASE)

    def is_continuation(text: str) -> bool:
        if _fund_excluded(text):
            return False
        if hits_stop(text):
            return False
        if is_anchor(text) and not _CONT_MARKER_RE.search(
                "\n".join((text or "").splitlines()[:5])):
            return False
        if bool(cont_re.search("\n".join((text or "").splitlines()[:30]))):
            return True
        return is_scrambled_cont(text)

    def is_truly_complete(page_text: str) -> bool:
        """
        Determines whether a non-overflow page marks the end of the statement.

        GOV_BS / GOV_IS  (no complete_re in defn):
            Uses plain is_complete_fn — the overflow detector already excludes
            pages with the closing row, so reaching is_complete_fn == True
            means we are past the closing-row page and should stop.

        PROP_SNP / PROP_IS / PROP_CFS  (complete_re present):
            Uses _complete_row_has_numbers — only True when numbers already
            appear on the same line as the closing row, meaning no overflow
            page is needed.
        """
        if not is_complete_fn:
            return False
        if not is_complete_fn(page_text):
            return False
        if complete_re is not None:
            # Doc 2 path: require numbers on the closing-row line
            return _complete_row_has_numbers(page_text, complete_re)
        # Doc 1 path (GOV_BS / GOV_IS): plain completion flag is sufficient
        return True

    # ── main scan ─────────────────────────────────────────────────────────────

    collected: list[int] = []
    scan_from = max(start_page, START_PAGE) - 1

    with pdfplumber.open(pdf_path) as pdf:
        total = len(pdf.pages)

        for i in range(scan_from, total):
            page_no = i + 1
            text    = pdf.pages[i].extract_text() or ""

            if _fund_excluded(text):
                continue
            if not is_anchor(text):
                continue

            collected.append(page_no)

            # Seed completion state from the anchor page itself.
            last_non_overflow_complete = is_truly_complete(text)

            j = i + 1
            while j < total:
                next_text = pdf.pages[j].extract_text() or ""

                # ── column-overflow check (BEFORE exclusion filter) ────────
                # Must run first: overflow pages can contain "Nonmajor" or
                # other exclusion-triggering words in column headers.
                if is_col_overflow_fn and is_col_overflow_fn(next_text):
                    if gov_overflow:
                        # GOV_BS / GOV_IS (Doc 1 path):
                        # _gov_*_is_column_overflow already excludes pages
                        # that contain the closing row via its own internal
                        # `and not _is_complete(text)` guard, so collect
                        # unconditionally here — no external check needed.
                        collected.append(j + 1)
                        j += 1
                        continue
                    else:
                        # PROP_SNP / PROP_IS / PROP_CFS (Doc 2 path):
                        # Overflow detector has no internal exclusion guard,
                        # so stop if the previous non-overflow page already
                        # carried the closing row's numbers (truly complete).
                        if last_non_overflow_complete:
                            break
                        collected.append(j + 1)
                        j += 1
                        # Overflow pages do NOT update last_non_overflow_complete.
                        continue

                # ── standard exclusion / stop / anchor checks ─────────────
                if _fund_excluded(next_text):
                    break
                if hits_stop(next_text):
                    break
                if next_anchor_re and _title_in_header(next_text, next_anchor_re):
                    break

                # If the left-column page was truly complete and this page is
                # not an overflow, the statement is fully collected.
                if last_non_overflow_complete:
                    break

                # ── continuation ──────────────────────────────────────────
                if is_continuation(next_text):
                    collected.append(j + 1)
                    last_non_overflow_complete = is_truly_complete(next_text)
                    j += 1
                    continue

                break  # nothing matched — stop

    if not collected:
        raise ValueError(f"{defn['suffix']} not found in {os.path.basename(pdf_path)}")
    return sorted(set(collected))


# ══════════════════════════════════════════════════════════════════════════════
# PROPRIETARY FUND PRESENCE CHECK
# ══════════════════════════════════════════════════════════════════════════════

_PROP_PRESENCE_KW = [
    re.compile(r"\benterprise\s+fund\b",             re.IGNORECASE),
    re.compile(r"\bstatement\s+of\s+net\s+position\b.{0,120}\bproprietary\b",
               re.IGNORECASE | re.DOTALL),
    re.compile(r"\bproprietary\b.{0,120}\bstatement\s+of\s+net\s+position\b",
               re.IGNORECASE | re.DOTALL),
    re.compile(r"\bstatement\s+of\s+cash\s+flows?\b.{0,120}\bproprietary\b",
               re.IGNORECASE | re.DOTALL),
    re.compile(r"\bproprietary\b.{0,120}\bstatement\s+of\s+cash\s+flows?\b",
               re.IGNORECASE | re.DOTALL),
    re.compile(r"\bproprietary\s+fund\b.{0,120}\bstatement\s+of\s+revenues",
               re.IGNORECASE | re.DOTALL),
]


def _has_proprietary_fund_statements(pdf_path: str) -> bool:
    """
    Quick scan: does this PDF contain actual proprietary fund statement pages
    (not just policy/notes references)?  Returns False for purely governmental
    entities that have no enterprise or internal-service-fund primary statements.
    """
    prop_page_hits = 0
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            first4 = "\n".join(text.splitlines()[:4])
            if re.search(r"\bnotes?\s+to\b|\bsupplemental\b|\bschedule\b",
                         first4, re.IGNORECASE):
                continue
            for kw in _PROP_PRESENCE_KW:
                if kw.search(text):
                    prop_page_hits += 1
                    break
            if prop_page_hits >= 2:
                return True
    return False


# ══════════════════════════════════════════════════════════════════════════════
# SHARED UTILITIES
# ══════════════════════════════════════════════════════════════════════════════

def _pages_suffix(pages: list[int]) -> str:
    if len(pages) == 1:
        return f"_p{pages[0]}"
    return f"_p{pages[0]}-{pages[-1]}"


def extract_pages_to_pdf(src: str, page_numbers: list[int], out: str):
    reader = PdfReader(src)
    writer = PdfWriter()
    for pn in page_numbers:
        writer.add_page(reader.pages[pn - 1])
    with open(out, "wb") as fh:
        writer.write(fh)


# ══════════════════════════════════════════════════════════════════════════════
# PROCESS ONE PDF  (all 7 statements)
# ══════════════════════════════════════════════════════════════════════════════

def process_file(src: str, out_dir: str):
    stem = os.path.splitext(os.path.basename(src))[0]
    os.makedirs(out_dir, exist_ok=True)

    print(f"\n{'─'*60}")
    print(f"  {os.path.basename(src)}")
    print(f"{'─'*60}")

    # ── 1. SNP (Government-Wide) ──────────────────────────────────────────────
    snp_pages = None
    try:
        snp_pages = find_snp_pages(src)
        out = os.path.join(out_dir, f"{stem}_SNP{_pages_suffix(snp_pages)}.pdf")
        extract_pages_to_pdf(src, snp_pages, out)
        print(f"  [SNP]       pages {snp_pages}  →  {os.path.basename(out)}")
    except Exception as e:
        print(f"  [SNP]       ERROR: {e}")

    # ── 2. SOA (Government-Wide) — starts after last SNP page ─────────────────
    soa_pages = None
    soa_search_start = (max(snp_pages) + 1) if snp_pages else START_PAGE
    try:
        soa_pages = find_soa_pages(src, start_page=soa_search_start)
        out = os.path.join(out_dir, f"{stem}_SOA{_pages_suffix(soa_pages)}.pdf")
        extract_pages_to_pdf(src, soa_pages, out)
        print(f"  [SOA]       pages {soa_pages}  →  {os.path.basename(out)}")
    except Exception as e:
        print(f"  [SOA]       ERROR: {e}")

    # ── 3-7. Fund-level statements — start after last SOA page ────────────────
    fund_search_start = (max(soa_pages) + 1) if soa_pages else soa_search_start

    prop_suffixes = {"PROP_SNP", "PROP_IS", "PROP_CFS"}
    has_prop = None

    for idx, defn in enumerate(FUND_STATEMENTS):
        suffix = defn["suffix"]
        next_anchor = FUND_STATEMENTS[idx + 1]["title_re"] if idx + 1 < len(FUND_STATEMENTS) else None

        if suffix in prop_suffixes and has_prop is None:
            has_prop = _has_proprietary_fund_statements(src)
            if not has_prop:
                print(f"  [INFO]      No proprietary fund statements detected — "
                      f"PROP_SNP / PROP_IS / PROP_CFS skipped.")

        if suffix in prop_suffixes and not has_prop:
            continue

        try:
            pages = find_fund_statement_pages(src, defn, fund_search_start, next_anchor)
            out   = os.path.join(out_dir, f"{stem}_{suffix}{_pages_suffix(pages)}.pdf")
            extract_pages_to_pdf(src, pages, out)
            print(f"  [{suffix:8s}] pages {pages}  →  {os.path.basename(out)}")
        except Exception as e:
            print(f"  [{suffix:8s}] ERROR: {e}")


# ══════════════════════════════════════════════════════════════════════════════
# FOLDER RUNNER
# ══════════════════════════════════════════════════════════════════════════════

_PRODUCED_RE = re.compile(
    r"_(?:SNP|SOA|GOV_BS|GOV_IS|PROP_SNP|PROP_IS|PROP_CFS)_p\d+(?:-\d+)?\.pdf$",
    re.IGNORECASE,
)


def process_folder(folder: str, out_dir: str | None = None):
    if out_dir is None:
        out_dir = folder

    pdfs = [
        f for f in os.listdir(folder)
        if f.lower().endswith(".pdf")
        and not f.endswith("_PageExtracted.pdf")
        and not _PRODUCED_RE.search(f)
    ]

    if not pdfs:
        print(f"No eligible PDFs found in: {folder}")
        return

    print(f"Scanning {len(pdfs)} PDF(s) in: {folder}")
    for fname in sorted(pdfs):
        process_file(os.path.join(folder, fname), out_dir)

    print(f"\nDone. Output in: {out_dir}")


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