from decimal import Decimal, InvalidOperation

# ─────────────────────────────────────────────────────────────────────────────
# PATCH NOTES
# ─────────────────────────────────────────────────────────────────────────────
#
# BUG 23 (FIXED): "Total Current Liabilities" in _PROP_SNP, _SNP, _GOV_BS
#   used "dp_sum_by_coa" matching zero rows → false FAIL.
#   Fixed: new rule type "dp_sum_current_liabilities" with Strategy A
#   (Items label starts with "current liabilities:") and Strategy B
#   (COA Datapoint does NOT end with "- noncurrent").
#
# BUG 24 (FIXED): _PROP_IS "Change In Net Position" chain ordering.
#   Fixed: Chain 1 = IBT + Contributions & Transfers (section fallback).
#
# BUG 25 (FIXED): "Total Liabilities, Deferred Inflows, and Net Position"
#   skipped when Deferred Inflows or Net Position had no CP subtotal.
#   Fixed: Chain 3 uses dict-form members with section-sum fallback.
#
# BUG 26 (FIXED): DSR "Total" CP Premium/Discount column causes false FAIL.
#   Fixed: dedicated _DSR block skips Premium/Discount column and adds
#   CP(Premium/Discount) to the Total Debt Service DP sum.
#
# BUG 27 (FIXED): _PROP_IS "Operating Income (Loss)" false FAIL when
#   "Total Operating Revenues" CP row is absent.
#   Fixed: new rule type "cp_or_section_subtract".
#
# BUG 28 (FIXED): _PROP_IS blank-COA-Datapoint CP rows never matched.
#   Fixed: new rule type "section_cp_blank_datapoint" with __BLANK__ sentinel.
#
# BUG 29 (FIXED): _PROP_CFS "Total Balance of Cash and Cash Equivalents"
#   missed extra DP rows in the Balance section.
#   Fixed: new rule type "section_dp_plus_cp".
#
# BUG 30 (FIXED): _SNP / _GOV_BS "Total Noncurrent Liabilities" dp_sum
#   matched "Current Liabilities: Current portion of long-term liabilities"
#   because "long-term liabilities" is a substring of that Items label.
#   Fixed: added "current liabilities:" to the excludes list for all three
#   statement types (_SNP, _GOV_BS, _PROP_SNP).
#
# BUG 31 (FIXED): _SOA "Total General Revenues and Transfer" used
#   "cp_or_section_sum" with a "coa_filter" key that was never implemented
#   in _sum_cp_or_section(). The member with cp=None fell through to summing
#   ALL DPs in "General Revenues" (including non-Transfer rows) → wrong total.
#   Fixed:
#   (a) New rule type "dp_sum_by_coa_in_section" — sums DP rows whose
#       COA Datapoint matches a target string, within a named section.
#   (b) New rule type "cp_sum_with_dp_fallback" — two-member rule:
#       member 0: CP lookup (Total General Revenues),
#       member 1: dp_sum_by_coa_in_section for Transfers DP in General Revenues.
#   This correctly computes: Total GR CP + Transfers DP = Total GR and Transfer CP.
#   Backward-compatible: if no Transfers DP exists in General Revenues,
#   member 1 contributes 0 (same as before).
#
# BUG 32 (FIXED): _PROP_SNP "Total Noncurrent Assets" dp_sum used
#   contains=["noncurrent assets"] but Items labels have a hyphen:
#   "Non-current Assets: ...". The substring "noncurrent assets" (no hyphen)
#   never matched → computed = 0 → false FAIL.
#   Fixed: added "non-current assets" to contains list so both hyphenated
#   and non-hyphenated prefixes match.
#
# BUG 33 (FIXED): _DEBT validation always produced "Skipped - no rule defined"
#   because RULES matched CP COA Datapoints by hardcoded name (e.g.
#   "Total governmental-type activities") which differed from actual CP names
#   in files (e.g. "Total", "Total Business-Type Activities").
#   Fixed: new rule type "dpg_cp_sum" — no COA Datapoint matching at all.
#   The engine scans each section for DPG→DP*→CP triplets in row order and
#   validates each CP as the sum of its immediately preceding DP rows
#   (those between it and the previous DPG or section start).
#   Fully dynamic: works for any CP name, any number of sub-groups,
#   any number of DP rows per group. Zero hardcoded strings.
#
# BUG 34 (FIXED): _apply_check used exact Decimal equality. LLM rounding
#   causes ±1-unit discrepancies (e.g. SOA Charges for Services off by $1,
#   Changes in Net Position off by $1) that are not real errors.
#   Fixed: tolerance of ±1 applied in _apply_check. Any column where
#   |computed − reported| ≤ 1 is treated as PASS. Columns with larger
#   discrepancies still FAIL normally.
# ─────────────────────────────────────────────────────────────────────────────


# ── BUG 18 chain ─────────────────────────────────────────────────────────────
CHAIN_TOTAL_ASSETS_AND_DEFERRED_OUTFLOWS = [
    [
        "Total Assets",
        {"cp": "Total Deferred Outflows of Resources",
         "section": "Deferred Outflows of Resources"},
    ],
    [
        {"cp_multi": ["Total Current Assets", "Total Noncurrent Assets"]},
        {"cp": "Total Deferred Outflows of Resources",
         "section": "Deferred Outflows of Resources"},
    ],
]

# ── BUG 19 / BUG 25 chain ────────────────────────────────────────────────────
CHAIN_TOTAL_LIAB_DEFERRED_INFLOWS_NET_POSITION = [
    # Chain 1 — explicit combined CP row printed.
    [
        "Total Liabilities and Deferred Inflows of Resources",
        "Total Net Position",
    ],
    # Chain 2 — three separate CP rows all present.
    [
        "Total Liabilities",
        "Total Deferred Inflows of Resources",
        "Total Net Position",
    ],
    # Chain 3 — section-sum fallback for Deferred Inflows / Net Position
    # when their CP rows are absent. Dict members never block _chain_usable().
    [
        "Total Liabilities",
        {"cp": "Total Deferred Inflows of Resources",
         "section": "Deferred Inflows of Resources"},
        {"cp": "Total Net Position",
         "section": "Net Position"},
    ],
]

RULES = {

    # ── PROP_SNP ──────────────────────────────────────────────────────────────
    "_PROP_SNP": [
        ("Total Current Assets",      "dp_sum",  "Assets",
            ["current assets"], ["noncurrent", "non-current"]),
        ("Net Capital Assets",         "dp_sum",  "Assets",
            ["capital assets"], []),
        ("Total Capital Assets",      "dp_sum_by_coa", "Assets",
            ["Capital Assets"], []),
        # BUG 32 FIX: added "non-current assets" (hyphenated form).
        ("Total Noncurrent Assets",    "dp_sum",  "Assets",
            ["noncurrent assets", "non-current assets"], []),
        ("Total Assets",               "section",  "Assets",
            [], []),
        ("Total Deferred Outflows of Resources", "section",
            "Deferred Outflows of Resources", [], []),
        ("Total Assets and Deferred Outflows of Resources",
            "cp_sum_first_present", None,
            CHAIN_TOTAL_ASSETS_AND_DEFERRED_OUTFLOWS, []),
        ("Total Current Liabilities",  "dp_sum_current_liabilities",  "Liabilities",
            [], []),
        # BUG 30 FIX: added "current liabilities:" to excludes.
        ("Total Noncurrent Liabilities", "dp_sum", "Liabilities",
            ["noncurrent liabilities", "non-current liabilities", "long-term liabilities"],
            ["current liabilities:"]),
        ("Total Liabilities",          "section",  "Liabilities",
            [], []),
        ("Total Deferred Inflows of Resources", "section",
            "Deferred Inflows of Resources", [], []),
        ("Total Liabilities and Deferred Inflows of Resources", "cp_sum", None,
            ["Total Liabilities", "Total Deferred Inflows of Resources"], []),
        ("Total Net Position",         "section", "Net Position",
            [], ["adjustment"]),
        ("Total Liabilities and Net Position", "cp_sum", None,
            ["Total Liabilities", "Total Net Position"], []),
        ("Total Liabilities, Deferred Inflows of Resources, and Net Position",
            "cp_sum_first_present", None,
            CHAIN_TOTAL_LIAB_DEFERRED_INFLOWS_NET_POSITION, []),
    ],

    # ── PROP_SNP_NONPROFIT ────────────────────────────────────────────────────
    "_PROP_SNP_NONPROFIT": [
        ("Total Assets",       "section", "Assets",      [], []),
        ("Total Capital Assets", "dp_sum_by_coa", "Assets",
            ["Capital Assets"], []),
        ("Total Liabilities",  "section", "Liabilities", [], []),
        ("Total Net Assets",   "section", "Net Assets",  [], []),
        ("Total Net Position", "section", "Net Assets",  [], []),
        ("Total Liabilities and Net Assets", "cp_sum", None,
            ["Total Liabilities", "Total Net Assets"], []),
        ("Total Liabilities and Net Position", "cp_sum", None,
            ["Total Liabilities", "Total Net Position"], []),
        ("Total Liabilities, Deferred inflows of resources, and Net position",
            "cp_sum_first_present", None,
            CHAIN_TOTAL_LIAB_DEFERRED_INFLOWS_NET_POSITION, []),
        ("Total Liabilities, Deferred inflows of resources, and Net assets",
            "cp_sum", None,
            ["Total Liabilities", "Total Net Assets"], []),
    ],

    # ── SNP ───────────────────────────────────────────────────────────────────
    "_SNP": [
        ("Total Current Assets",       "dp_sum",  "Assets",
            ["current assets"], ["noncurrent", "non-current"]),
        ("Net Capital Assets",         "dp_sum",  "Assets",
            ["capital assets"], []),
        ("Total Capital Assets",       "dp_sum_by_coa", "Assets",
            ["Capital Assets"], []),
        ("Total Noncurrent Assets",    "dp_sum",  "Assets",
            ["noncurrent assets", "non-current assets"], []),
        ("Total Assets",               "section",  "Assets",
            [], []),
        ("Total Deferred Outflows of Resources", "section",
            "Deferred Outflows of Resources", [], []),
        ("Total Assets and Deferred Outflows of Resources",
            "cp_sum_first_present", None,
            CHAIN_TOTAL_ASSETS_AND_DEFERRED_OUTFLOWS, []),
        ("Total Current Liabilities",  "dp_sum_current_liabilities",  "Liabilities",
            [], []),
        # BUG 30 FIX: added "current liabilities:" to excludes.
        ("Total Noncurrent Liabilities", "dp_sum", "Liabilities",
            ["noncurrent liabilities", "non-current liabilities", "long-term liabilities"],
            ["current liabilities:"]),
        ("Total Liabilities",          "section",  "Liabilities",
            [], []),
        ("Total Deferred Inflows of Resources", "section",
            "Deferred Inflows of Resources", [], []),
        ("Total Liabilities and Deferred Inflows of Resources", "cp_sum", None,
            ["Total Liabilities", "Total Deferred Inflows of Resources"], []),
        ("Total Restricted",           "dp_sum",  "Net Position",
            ["restricted"], ["unrestricted"]),
        ("Total Net Position",         "dp_cp_sum", "Net Position",
            [
                {"type": "dp",  "coa": "Net Investment in Capital Assets"},
                {"type": "cp_or_dp_sum", "coa": "Total Restricted",
                 "section": "Net Position", "dp_coa": "Restricted Net Assets"},
                {"type": "dp",  "coa": "Unrestricted Net Assets"},
            ], []),
        ("Total Liabilities and Net Position", "cp_sum", None,
            ["Total Liabilities", "Total Net Position"], []),
        ("Total Liabilities, Deferred inflows of resources, and Net position",
            "cp_sum_first_present", None,
            CHAIN_TOTAL_LIAB_DEFERRED_INFLOWS_NET_POSITION, []),
    ],

    # ── GOV_BS ────────────────────────────────────────────────────────────────
    "_GOV_BS": [
        ("Total Current Assets",       "dp_sum",  "Assets",
            ["current assets"], ["noncurrent", "non-current"]),
        ("Total Noncurrent Assets",    "dp_sum",  "Assets",
            ["noncurrent assets", "non-current assets"], []),
        ("Total Capital Assets",       "dp_sum_by_coa", "Assets",
            ["Capital Assets"], []),
        ("Total Assets",               "section",  "Assets",
            [], []),
        ("Total Deferred Outflows of Resources", "section",
            "Deferred Outflows of Resources", [], []),
        ("Total Assets and Deferred Outflows of Resources",
            "cp_sum_first_present", None,
            CHAIN_TOTAL_ASSETS_AND_DEFERRED_OUTFLOWS, []),
        ("Total Current Liabilities",  "dp_sum_current_liabilities",  "Liabilities",
            [], []),
        # BUG 30 FIX: added "current liabilities:" to excludes.
        ("Total Noncurrent Liabilities", "dp_sum", "Liabilities",
            ["noncurrent liabilities", "non-current liabilities", "long-term liabilities"],
            ["current liabilities:"]),
        ("Total Liabilities",          "section",  "Liabilities",
            [], []),
        ("Total Deferred Inflows of Resources", "section",
            "Deferred Inflows of Resources", [], []),
        ("Total Liabilities and Deferred Inflows of Resources", "cp_sum", None,
            ["Total Liabilities", "Total Deferred Inflows of Resources"], []),
        ("Total Fund Balances",        "section", "Fund Balances", [], []),
        ("Total Liabilities, Deferred Inflows of Resources, And Fund Balances",
            "cp_or_section_sum", None,
            [
                {"cp": "Total Liabilities", "section": "Liabilities"},
                {"cp": "Total Deferred Inflows of Resources",
                 "section": "Deferred Inflows of Resources"},
                {"cp": "Total Fund Balances", "section": "Fund Balances"},
            ], []),
    ],

    # ── GOV_IS ────────────────────────────────────────────────────────────────
    "_GOV_IS": [
        ("Total Operating Revenues",   "section", "Revenues",         [], []),
        ("Total Operating Expenses",   "section", "Expenditures",     [], []),
        ("Excess Of Revenues Over/Under Expenditures", "cp_subtract", None,
            ["Total Operating Revenues", "Total Operating Expenses"], []),
        ("Total Other Financing Sources", "section",
            "Other Financing Sources (Uses)", [], []),
        ("Net Change In Fund Balances", "cp_sum", None,
            ["Excess Of Revenues Over/Under Expenditures",
             "Total Other Financing Sources"], []),
    ],

    # ── SOA ───────────────────────────────────────────────────────────────────
    "_SOA": [
        # BUG 31a FIX: exclude Transfers DP from Total General Revenues sum.
        ("Total General Revenues", "section", "General Revenues",
            [], ["transfers"]),
        # BUG 31b FIX: new rule type "cp_sum_with_dp_fallback".
        # Computes: CP("Total General Revenues") + DP(COA="Transfers" in "General Revenues").
        # Handles filings where Transfers DP lives inside "General Revenues" section
        # rather than in a separate "Transfers" section.
        ("Total General Revenues and Transfer", "cp_sum_with_dp_fallback", None,
    [
        {"type": "cp", "coa": "Total General Revenues",
         "section": "General Revenues"},   # <-- fallback section added
        {"type": "dp_coa_in_section",
         "coa": "Transfers",
         "section": "General Revenues"},
    ], []),
        ("Total Program Receipts - Charges for Services", "dp_sum_by_coa",
            "Program Revenue", ["Charges for Services"], []),
        ("Total Program Receipts - Operating Grants and Contributions",
            "dp_sum_by_coa", "Program Revenue",
            ["Operating Grants and Contributions"], []),
        ("Total Program Receipts - Capital Grants and Contributions",
            "dp_sum_by_coa", "Program Revenue",
            ["Capital Grants and Contributions"], []),
        ("Total Program Expenses",     "section", "Program Expenses",  [], []),
        ("Changes in Net Position",    "cp_subtract_last", None,
            ["Total General Revenues and Transfer",
             "Total Program Receipts - Charges for Services",
             "Total Program Receipts - Operating Grants and Contributions",
             "Total Program Receipts - Capital Grants and Contributions",
             "Total Program Expenses"],
            []),
    ],

    # ── PROP_IS ───────────────────────────────────────────────────────────────
    "_PROP_IS": [
        ("Total Operating Revenues",   "section",
            "Operating Revenues",      [], []),
        ("Total Operating Expenses",   "section",
            "Operating Expenses",      [], []),
        ("Operating Income (Loss)",    "cp_or_section_subtract", None,
            ["Total Operating Revenues", "Operating Revenues",
             "Total Operating Expenses"], []),
        ("Total Nonoperating Revenues/Expense", "section",
            "Nonoperating Revenues (Expenses)", [], []),
        ("Income (Loss) before Transfers", "cp_sum", None,
            ["Operating Income (Loss)",
             "Total Nonoperating Revenues/Expense"], []),
        ("Change In Net Position", "cp_sum_first_present", None,
            [
                [
                    "Income (Loss) before Transfers",
                    {"cp": "Total contributions and transfers",
                     "section": "Contributions and Transfers"},
                ],
                ["Income (Loss) before Transfers"],
                ["Operating Income (Loss)",
                 "Total Nonoperating Revenues/Expense"],
                ["Operating Income",
                 "Total Non-operating Income"],
                [
                    "Operating Income (Loss)",
                    {"cp": "Total contributions and transfers",
                     "section": "Contributions and Transfers"},
                ],
            ], []),
        ("__BLANK__", "section_cp_blank_datapoint",
            "Contributions and Transfers", [], []),
    ],

    # ── PROP_CFS ──────────────────────────────────────────────────────────────
    "_PROP_CFS": [
        ("Total Cash Flows from Operating Activities", "section",
            "Cash Flows from Operating Activities", [], []),
        ("Total Cash Flows from Noncapital Financing Activities", "section",
            "Cash Flows from Noncapital Financing Activities", [], []),
        ("Total Cash Flows from Capital and Related Financing Activities",
            "section",
            "Cash Flows from Capital and Related Financing Activities", [], []),
        ("Total Cash Flows from Investing Activities", "section",
            "Cash Flows from Investing Activities", [], []),
        ("Net Change in Cash and Cash Equivalents", "dp_cp_sum", None,
            [
                {"type": "cp", "coa": "Total Cash Flows from Operating Activities"},
                {"type": "cp", "coa": "Total Cash Flows from Noncapital Financing Activities"},
                {"type": "cp", "coa": "Total Cash Flows from Capital and Related Financing Activities"},
                {"type": "cp_or_dp_sum",
                 "coa": "Total Cash Flows from Investing Activities",
                 "section": "Cash Flows from Investing Activities",
                 "dp_coa": "Cash Flows from Investing Activities"},
            ], []),
        ("Total Balance of Cash and Cash Equivalents", "section_dp_plus_cp",
            "Balance of Cash and Cash Equivalents",
            ["Net Change in Cash and Cash Equivalents",
             "Cash and Cash Equivalents"], []),
    ],

    # ── DSR ───────────────────────────────────────────────────────────────────
    "_DSR": [
        ("Total", "section", None, [], []),
    ],

    # ── DEBT ──────────────────────────────────────────────────────────────────
    # BUG 33 FIX: replaced all hardcoded CP name rules with a single
    # "dpg_cp_sum" sentinel rule. The engine's _DEBT special block handles
    # all sections: it scans row-order for DPG→DP*→CP triplets and validates
    # each CP as the sum of its immediately preceding DP rows. Fully dynamic —
    # works regardless of how the LLM names the CP rows or how many sub-groups
    # exist within each section.
    "_DEBT": [
        ("__DPG_CP_SUM__", "dpg_cp_sum", None, [], []),
    ],
}


# ─────────────────────────────────────────────────────────────────────────────
# NONPROFIT _PROP_SNP DETECTION
# ─────────────────────────────────────────────────────────────────────────────
def _is_nonprofit_prop_snp(data: dict) -> bool:
    sections = data.get("Sections", {})
    section_names_lower = {s.strip().lower() for s in sections.keys()}
    has_net_assets   = "net assets" in section_names_lower
    has_net_position = "net position" in section_names_lower
    has_deferred      = any(
        "deferred outflows" in s or "deferred inflows" in s
        for s in section_names_lower
    )
    return has_net_assets and not has_net_position and not has_deferred


def resolve_stmt_type_for_rules(data: dict, stmt_type: str) -> str:
    if stmt_type == "_PROP_SNP" and _is_nonprofit_prop_snp(data):
        return "_PROP_SNP_NONPROFIT"
    return stmt_type


def _table_id(row: dict) -> tuple:
    return (
        row.get("Statement", ""),
        row.get("Issuer Name", ""),
        row.get("FYE", ""),
        row.get("Page No", ""),
    )


# ─────────────────────────────────────────────────────────────────────────────
# CORE HELPERS
# ─────────────────────────────────────────────────────────────────────────────

def _val(v) -> Decimal:
    if v is None or str(v).strip() in ("-", "", "null"):
        return Decimal(0)
    s = str(v).strip().replace(",", "")
    if s.startswith("(") and s.endswith(")"):
        s = "-" + s[1:-1]
    try:
        return Decimal(s)
    except InvalidOperation:
        return Decimal(0)


def _items_key(data):
    for rows in data.get("Sections", {}).values():
        for r in rows:
            for k in r:
                if "Items" in k:
                    return k
    return "Items"


def _get_cp(sections, coa_dp, table_id=None):
    if not coa_dp:
        return None
    coa_dp_lower = coa_dp.strip().lower()
    for rows in sections.values():
        for r in rows:
            if table_id is not None and _table_id(r) != table_id:
                continue
            if (r.get("COA Flag") == "CP"
                    and r.get("COA Datapoint", "").strip().lower() == coa_dp_lower):
                return r
    return None


def _get_dp(sections, coa_dp, section_name=None, table_id=None):
    coa_dp_lower = coa_dp.strip().lower()
    search_sections = (
        {section_name: sections[section_name]}
        if section_name and section_name in sections
        else sections
    )
    for rows in search_sections.values():
        for r in rows:
            if table_id is not None and _table_id(r) != table_id:
                continue
            if (r.get("COA Flag") == "DP"
                    and r.get("COA Datapoint", "").strip().lower() == coa_dp_lower):
                return r
    return None


def _dp_rows_in_section(sections, section_name, table_id=None):
    rows = sections.get(section_name, [])
    if table_id is None:
        return [r for r in rows if r.get("COA Flag") == "DP"]
    return [r for r in rows
            if r.get("COA Flag") == "DP" and _table_id(r) == table_id]


def _dp_rows_by_coa(sections, coa_dp, section_name=None, table_id=None):
    coa_dp_lower = coa_dp.strip().lower()
    pool = (_dp_rows_in_section(sections, section_name, table_id=table_id)
            if section_name else
            [r for sec in sections.values() for r in sec
             if r.get("COA Flag") == "DP"
             and (table_id is None or _table_id(r) == table_id)])
    return [r for r in pool
            if r.get("COA Datapoint", "").strip().lower() == coa_dp_lower]


# BUG 34 FIX: tolerance of ±1 for LLM rounding discrepancies.
def _apply_check(cp_row, member_vals, cols, tolerance=1):
    if cp_row is None:
        return
    fails = []
    for col in cols:
        computed = member_vals.get(col, Decimal(0))
        reported = _val(cp_row.get(col, "-"))
        if abs(computed - reported) > tolerance:
            fails.append(col)
    cp_row["Total Check Status"] = (
        "PASS" if not fails
        else " | ".join(f"{c}: FAIL" for c in fails)
    )


def _sum_cp_vals(sections, cp_datapoints, cols, subtract_last=False, table_id=None):
    totals = {col: Decimal(0) for col in cols}
    for i, dp in enumerate(cp_datapoints):
        row = _get_cp(sections, dp, table_id=table_id)
        if row is None:
            continue
        for col in cols:
            v = _val(row.get(col, "-"))
            if subtract_last and i == len(cp_datapoints) - 1:
                totals[col] -= v
            else:
                totals[col] += v
    return totals


def _sum_cp_or_section(sections, members, cols, ikey, excludes=None, table_id=None):
    excludes = excludes or []
    totals = {col: Decimal(0) for col in cols}
    for m in members:
        cp_name = m.get("cp")
        cp_row = _get_cp(sections, cp_name, table_id=table_id) if cp_name else None
        if cp_row is not None:
            for col in cols:
                totals[col] += _val(cp_row.get(col, "-"))
            continue
        pool = _dp_rows_in_section(sections, m["section"], table_id=table_id)
        pool = [r for r in pool
                if not any(e.lower() in r.get(ikey, "").lower() for e in excludes)]
        for r in pool:
            for col in cols:
                totals[col] += _val(r.get(col, "-"))
    return totals


def _sum_dp_cp_mixed(sections, members, cols, section_name=None, table_id=None):
    totals = {col: Decimal(0) for col in cols}
    for m in members:
        mtype = m.get("type")
        coa   = m.get("coa", "")

        if mtype == "dp":
            row = _get_dp(sections, coa, section_name, table_id=table_id)
            if row is None:
                continue
            for col in cols:
                totals[col] += _val(row.get(col, "-"))

        elif mtype == "cp":
            row = _get_cp(sections, coa, table_id=table_id)
            if row is None:
                continue
            for col in cols:
                totals[col] += _val(row.get(col, "-"))

        elif mtype == "cp_or_dp_sum":
            cp_row = _get_cp(sections, coa, table_id=table_id)
            if cp_row is not None:
                for col in cols:
                    totals[col] += _val(cp_row.get(col, "-"))
                continue
            fallback_section = m.get("section", section_name)
            dp_coa = m.get("dp_coa", coa)
            pool = _dp_rows_by_coa(sections, dp_coa,
                                   section_name=fallback_section,
                                   table_id=table_id)
            for r in pool:
                for col in cols:
                    totals[col] += _val(r.get(col, "-"))
        else:
            continue
    return totals


# ─────────────────────────────────────────────────────────────────────────────
# BUG 31b FIX: cp_sum_with_dp_fallback handler
# ─────────────────────────────────────────────────────────────────────────────
def _sum_cp_with_dp_fallback(sections, members, cols, table_id=None, ikey=None):
    totals = {col: Decimal(0) for col in cols}
    section_fallbacks_used = set()   # track which sections were summed via fallback

    for m in members:
        mtype = m.get("type")
        if mtype == "cp":
            row = _get_cp(sections, m["coa"], table_id=table_id)
            if row is not None:
                for col in cols:
                    totals[col] += _val(row.get(col, "-"))
            else:
                fallback_sec = m.get("section")
                if fallback_sec:
                    pool = _dp_rows_in_section(sections, fallback_sec, table_id=table_id)
                    excl = m.get("excludes", [])
                    if excl and ikey:
                        pool = [r for r in pool
                                if not any(e.lower() in str(r.get(ikey, "")).lower()
                                           for e in excl)]
                    for col in cols:
                        totals[col] += sum(_val(r.get(col, "-")) for r in pool)
                    section_fallbacks_used.add(fallback_sec)

        elif mtype == "dp_coa_in_section":
            # Skip if the parent section was already fully summed via fallback
            if m.get("section") in section_fallbacks_used:
                continue
            coa_target = m["coa"].strip().lower()
            sec_rows = _dp_rows_in_section(sections, m["section"], table_id=table_id)
            matching = [r for r in sec_rows
                        if r.get("COA Datapoint", "").strip().lower() == coa_target]
            for r in matching:
                for col in cols:
                    totals[col] += _val(r.get(col, "-"))

    return totals

# ─────────────────────────────────────────────────────────────────────────────
# cp_sum_first_present helpers
# ─────────────────────────────────────────────────────────────────────────────
def _resolve_chain_member(sections, member, cols, ikey, table_id=None):
    if isinstance(member, dict):
        if "cp_multi" in member:
            totals = {col: Decimal(0) for col in cols}
            for name in member["cp_multi"]:
                row = _get_cp(sections, name, table_id=table_id)
                if row is None:
                    continue
                for col in cols:
                    totals[col] += _val(row.get(col, "-"))
            return (totals, False)
        cp_row = _get_cp(sections, member["cp"], table_id=table_id)
        if cp_row is not None:
            return ({col: _val(cp_row.get(col, "-")) for col in cols}, False)
        pool = _dp_rows_in_section(sections, member["section"], table_id=table_id)
        totals = {col: Decimal(0) for col in cols}
        for r in pool:
            for col in cols:
                totals[col] += _val(r.get(col, "-"))
        return (totals, False)
    else:
        cp_row = _get_cp(sections, member, table_id=table_id)
        if cp_row is None:
            return ({col: Decimal(0) for col in cols}, True)
        return ({col: _val(cp_row.get(col, "-")) for col in cols}, False)


def _chain_usable(sections, chain, table_id=None):
    for member in chain:
        if isinstance(member, dict):
            if "cp_multi" in member:
                for name in member["cp_multi"]:
                    if _get_cp(sections, name, table_id=table_id) is None:
                        return False
                continue
            # {"cp":.., "section":..} always has DP fallback — never blocks.
            continue
        if _get_cp(sections, member, table_id=table_id) is None:
            return False
    return True


def _chain_label(member):
    if isinstance(member, str):
        return member
    if "cp_multi" in member:
        return "(" + " + ".join(member["cp_multi"]) + ")"
    return f"{member['cp']} (or section: {member['section']})"


def _sum_cp_first_present(sections, chains, cols, ikey, table_id=None):
    for chain in chains:
        if _chain_usable(sections, chain, table_id=table_id):
            totals = {col: Decimal(0) for col in cols}
            for member in chain:
                vals, _ = _resolve_chain_member(
                    sections, member, cols, ikey, table_id=table_id
                )
                for col in cols:
                    totals[col] += vals[col]
            label = " + ".join(_chain_label(m) for m in chain)
            return totals, label, True
    return None, None, False


# ─────────────────────────────────────────────────────────────────────────────
# MAIN ENGINE
# ─────────────────────────────────────────────────────────────────────────────

def run_total_check(data: dict, stmt_type: str) -> dict:
    try:
        from column_shift_repair import repair_column_shifts
        repair_column_shifts(data, stmt_type)
    except Exception as e:
        print(f"[WARN] column_shift_repair skipped: {e}")

    sections = data.get("Sections", {})
    cols     = data.get("Reporting Columns", [])

    effective_stmt_type = resolve_stmt_type_for_rules(data, stmt_type)
    rules = RULES.get(effective_stmt_type, [])
    ikey  = _items_key(data)

    # Initialise all CP rows to PENDING; non-CP rows to blank.
    for rows in sections.values():
        for r in rows:
            flag = r.get("COA Flag", "")
            r["Total Check Status"] = "PENDING" if flag == "CP" else ""

    for rule in rules:
        cp_dp, rule_type, section, contains, excludes = rule

        # Sentinel rules handled in dedicated post-processing blocks below.
        if cp_dp in ("__BLANK__", "__DPG_CP_SUM__"):
            continue

        matching_cp_rows = [
            r for rows in sections.values() for r in rows
            if r.get("COA Flag") == "CP"
            and r.get("COA Datapoint", "").strip().lower() == cp_dp.strip().lower()
        ]
        if not matching_cp_rows:
            continue

        for cp_row in matching_cp_rows:
            tid = _table_id(cp_row)

            # ── dp_sum ────────────────────────────────────────────────────
            if rule_type == "dp_sum":
                if section:
                    pool = [r for r in sections.get(section, [])
                            if r.get("COA Flag") == "DP" and _table_id(r) == tid]
                else:
                    pool = [r for sec in sections.values() for r in sec
                            if r.get("COA Flag") == "DP" and _table_id(r) == tid]
                matched = []
                for r in pool:
                    label = r.get(ikey, "").lower()
                    hit = any(c.lower() in label for c in contains) if contains else True
                    if hit and not any(label.startswith(e.lower()) for e in excludes):
                        matched.append(r)
                member_vals = {col: sum(_val(r.get(col, "-")) for r in matched)
                               for col in cols}
                _apply_check(cp_row, member_vals, cols)

            # ── dp_sum_by_coa ─────────────────────────────────────────────
            elif rule_type == "dp_sum_by_coa":
                if section:
                    pool = [r for r in sections.get(section, [])
                            if r.get("COA Flag") == "DP" and _table_id(r) == tid]
                else:
                    pool = [r for sec in sections.values() for r in sec
                            if r.get("COA Flag") == "DP" and _table_id(r) == tid]
                coa_targets = {c.strip().lower() for c in contains}
                matched = [r for r in pool
                           if r.get("COA Datapoint", "").strip().lower() in coa_targets
                           and not any(e.lower() in r.get(ikey, "").lower()
                                       for e in excludes)]
                member_vals = {col: sum(_val(r.get(col, "-")) for r in matched)
                               for col in cols}
                _apply_check(cp_row, member_vals, cols)

            # ── dp_sum_current_liabilities (BUG 23) ───────────────────────
            elif rule_type == "dp_sum_current_liabilities":
                if section:
                    pool = [r for r in sections.get(section, [])
                            if r.get("COA Flag") == "DP" and _table_id(r) == tid]
                else:
                    pool = [r for sec in sections.values() for r in sec
                            if r.get("COA Flag") == "DP" and _table_id(r) == tid]
                matched = []
                for r in pool:
                    item_label = r.get(ikey, "").strip().lower()
                    coa_dp_val = r.get("COA Datapoint", "").strip().lower()
                    # Strategy A — PROP_SNP style: Items label starts with prefix.
                    if item_label.startswith("current liabilities:"):
                        matched.append(r)
                        continue
                    # Strategy B — SNP / GOV_BS style: not a noncurrent COA.
                    if not coa_dp_val.endswith("- noncurrent"):
                        matched.append(r)
                member_vals = {col: sum(_val(r.get(col, "-")) for r in matched)
                               for col in cols}
                _apply_check(cp_row, member_vals, cols)

            # ── section ───────────────────────────────────────────────────
            elif rule_type == "section":
                if section:
                    pool = _dp_rows_in_section(sections, section, table_id=tid)
                else:
                    pool = [r for sec in sections.values() for r in sec
                            if r.get("COA Flag") == "DP" and _table_id(r) == tid]
                pool = [r for r in pool
                        if not any(e.lower() in r.get(ikey, "").lower()
                                   for e in excludes)]
                member_vals = {col: sum(_val(r.get(col, "-")) for r in pool)
                               for col in cols}
                _apply_check(cp_row, member_vals, cols)

            # ── section_multi ─────────────────────────────────────────────
            elif rule_type == "section_multi":
                pool = []
                for sec_name in contains:
                    pool += _dp_rows_in_section(sections, sec_name, table_id=tid)
                pool = [r for r in pool
                        if not any(e.lower() in r.get(ikey, "").lower()
                                   for e in excludes)]
                member_vals = {col: sum(_val(r.get(col, "-")) for r in pool)
                               for col in cols}
                _apply_check(cp_row, member_vals, cols)

            # ── cp_sum ────────────────────────────────────────────────────
            elif rule_type == "cp_sum":
                member_vals = _sum_cp_vals(sections, contains, cols, table_id=tid)
                _apply_check(cp_row, member_vals, cols)

            # ── cp_subtract / cp_subtract_last ────────────────────────────
            elif rule_type in ("cp_subtract", "cp_subtract_last"):
                member_vals = _sum_cp_vals(
                    sections, contains, cols, subtract_last=True, table_id=tid
                )
                _apply_check(cp_row, member_vals, cols)

            # ── cp_or_section_subtract (BUG 27) ───────────────────────────
            elif rule_type == "cp_or_section_subtract":
                rev_cp_name, rev_section_name, exp_cp_name = (
                    contains[0], contains[1], contains[2]
                )
                rev_cp_row = _get_cp(sections, rev_cp_name, table_id=tid)
                if rev_cp_row is not None:
                    rev_vals = {col: _val(rev_cp_row.get(col, "-")) for col in cols}
                else:
                    pool = _dp_rows_in_section(sections, rev_section_name, table_id=tid)
                    rev_vals = {col: sum(_val(r.get(col, "-")) for r in pool)
                                for col in cols}
                exp_cp_row = _get_cp(sections, exp_cp_name, table_id=tid)
                exp_vals = ({col: _val(exp_cp_row.get(col, "-")) for col in cols}
                            if exp_cp_row else {col: Decimal(0) for col in cols})
                member_vals = {col: rev_vals[col] - exp_vals[col] for col in cols}
                _apply_check(cp_row, member_vals, cols)

            # ── section_dp_plus_cp (BUG 29) ───────────────────────────────
            elif rule_type == "section_dp_plus_cp":
                candidate_sections = [section] + list(contains[1:])
                dp_pool = []
                for sec_candidate in candidate_sections:
                    dp_pool = _dp_rows_in_section(sections, sec_candidate, table_id=tid)
                    if dp_pool:
                        break
                dp_sum = {col: sum(_val(r.get(col, "-")) for r in dp_pool)
                          for col in cols}
                net_cp = _get_cp(sections, contains[0], table_id=tid)
                net_vals = ({col: _val(net_cp.get(col, "-")) for col in cols}
                            if net_cp else {col: Decimal(0) for col in cols})
                member_vals = {col: dp_sum[col] + net_vals[col] for col in cols}
                _apply_check(cp_row, member_vals, cols)

            # ── dp_cp_sum ─────────────────────────────────────────────────
            elif rule_type == "dp_cp_sum":
                member_vals = _sum_dp_cp_mixed(
                    sections, contains, cols, section, table_id=tid
                )
                _apply_check(cp_row, member_vals, cols)

            # ── cp_or_section_sum ─────────────────────────────────────────
            elif rule_type == "cp_or_section_sum":
                member_vals = _sum_cp_or_section(
                    sections, contains, cols, ikey,
                    excludes=excludes, table_id=tid
                )
                _apply_check(cp_row, member_vals, cols)

            # ── cp_sum_with_dp_fallback (BUG 31b) ─────────────────────────
            elif rule_type == "cp_sum_with_dp_fallback":
                member_vals = _sum_cp_with_dp_fallback(
                    sections, contains, cols, table_id=tid, ikey=ikey
                )
                _apply_check(cp_row, member_vals, cols)

            # ── cp_sum_first_present ──────────────────────────────────────
            elif rule_type == "cp_sum_first_present":
                member_vals, chosen_label, matched = _sum_cp_first_present(
                    sections, contains, cols, ikey, table_id=tid
                )
                if not matched:
                    cp_row["Total Check Status"] = "Skipped - no chain matched"
                    cp_row["_chain_used"] = None
                else:
                    _apply_check(cp_row, member_vals, cols)
                    cp_row["_chain_used"] = chosen_label

    # ── BUG 28 FIX: blank-datapoint CP rows ──────────────────────────────────
    blank_dp_rules = {
        rule[2]: rule
        for rule in RULES.get(effective_stmt_type, [])
        if rule[1] == "section_cp_blank_datapoint"
    }
    if blank_dp_rules:
        for sec_name, sec_rows in sections.items():
            if sec_name not in blank_dp_rules:
                continue
            tables_in_sec = {}
            for r in sec_rows:
                tables_in_sec.setdefault(_table_id(r), []).append(r)
            for tid, trows in tables_in_sec.items():
                blank_cp_rows = [
                    r for r in trows
                    if r.get("COA Flag") == "CP"
                    and not str(r.get("COA Datapoint", "")).strip()
                ]
                if not blank_cp_rows:
                    continue
                dp_pool = [r for r in trows if r.get("COA Flag") == "DP"]
                member_vals = {col: sum(_val(r.get(col, "-")) for r in dp_pool)
                               for col in cols}
                for bcp in blank_cp_rows:
                    _apply_check(bcp, member_vals, cols)

    # ── BUG 33 FIX: DEBT dpg_cp_sum block ────────────────────────────────────
    # For each section, walk rows in order. Every DPG starts a new group.
    # All DP rows after a DPG and before the next CP belong to that group.
    # The CP closes the group and is validated as the sum of its DP group.
    # Works for any CP name, any number of groups per section.
    if stmt_type == "_DEBT":
        for sec_name, sec_rows in sections.items():
            # Group rows by table_id first (multi-filing sheets).
            tables_in_sec = {}
            for r in sec_rows:
                tables_in_sec.setdefault(_table_id(r), []).append(r)

            for tid, trows in tables_in_sec.items():
                pending_dps = []
                for r in trows:
                    flag = r.get("COA Flag", "")
                    if flag == "DPG":
                        # Start a fresh DP accumulator for this sub-group.
                        pending_dps = []
                    elif flag == "DP":
                        pending_dps.append(r)
                    elif flag == "CP":
                        # Validate this CP against the accumulated DPs.
                        member_vals = {
                            col: sum(_val(dp.get(col, "-")) for dp in pending_dps)
                            for col in cols
                        }
                        _apply_check(r, member_vals, cols)
                        # Reset so subsequent DPs don't double-count.
                        pending_dps = []

    # ── BUG 26 FIX: DSR special block ────────────────────────────────────────
    if stmt_type == "_DSR":
        pd_col = next(
            (c for c in cols if str(c).strip().lower() == "premium/discount"), None
        )
        tds_col = next(
            (c for c in cols if str(c).strip().lower() == "total debt service"), None
        )
        for sec_name, rows in sections.items():
            tables_in_section = {}
            for r in rows:
                tables_in_section.setdefault(_table_id(r), []).append(r)
            for tid, trows in tables_in_section.items():
                dp_rows = [r for r in trows if r.get("COA Flag") == "DP"]
                cp_row  = next(
                    (r for r in trows if r.get("COA Flag") == "CP"), None
                )
                if cp_row is None:
                    continue
                normal_cols = [c for c in cols if c != pd_col]
                member_vals = {col: sum(_val(r.get(col, "-")) for r in dp_rows)
                               for col in normal_cols}
                if tds_col and tds_col in normal_cols and pd_col:
                    pd_cp_val = _val(cp_row.get(pd_col, "-"))
                    member_vals[tds_col] = member_vals[tds_col] + pd_cp_val
                _apply_check(cp_row, member_vals, normal_cols)

    # Any CP row still PENDING had no matching rule.
    for rows in sections.values():
        for r in rows:
            if r.get("Total Check Status") == "PENDING":
                r["Total Check Status"] = "Skipped - no rule defined"

    return data


# ─────────────────────────────────────────────────────────────────────────────
# XLSX DRIVER
# ─────────────────────────────────────────────────────────────────────────────
import sys
from pathlib import Path
from collections import OrderedDict
import pandas as pd

SHEET_TO_STMT_TYPE = {
    "SNP":      "_SNP",
    "SOA":      "_SOA",
    "GOV_BS":   "_GOV_BS",
    "GOV_IS":   "_GOV_IS",
    "PROP_SNP": "_PROP_SNP",
    "PROP_IS":  "_PROP_IS",
    "PROP_CFS": "_PROP_CFS",
    "DEBT":     "_DEBT",
    "DSR":      "_DSR",
}

_META_COLS = {
    "Statement", "Issuer Name", "FYE", "Page No", "Currency reported",
    "Section", "COA Flag", "COA Datapoint",
    "Total Check Status", "_chain_used",
}


def _clean_sheet_name(name: str) -> str:
    return str(name).strip().replace("\\_", "_")


def _detect_reporting_columns(df: pd.DataFrame) -> list:
    cols = list(df.columns)
    if "COA Datapoint" not in cols:
        raise ValueError("Missing required column: 'COA Datapoint'")
    start = cols.index("COA Datapoint") + 1
    reporting = []
    for c in cols[start:]:
        c_str = str(c).strip()
        if c_str == "Total Check Status":
            break
        if c_str.startswith("_"):
            break
        if c_str in _META_COLS:
            continue
        if not c_str or c_str.lower().startswith("unnamed"):
            continue
        if c_str.endswith(" Items") or c_str == "Row Items":
            continue
        reporting.append(c)
    return reporting


def _sheet_df_to_data(df: pd.DataFrame) -> dict:
    df = df.fillna("").copy()
    for req in ("Section", "COA Flag", "COA Datapoint"):
        if req not in df.columns:
            raise ValueError(f"Missing required column: '{req}'")
    reporting_cols = _detect_reporting_columns(df)
    sections = OrderedDict()
    for idx, row in df.iterrows():
        r = row.to_dict()
        r["_row_order"] = idx
        sec = str(r.get("Section", "")).strip() or "__NO_SECTION__"
        sections.setdefault(sec, []).append(r)
    return {"Reporting Columns": reporting_cols, "Sections": sections}


def _data_to_sheet_df(data: dict, original_cols: list) -> pd.DataFrame:
    rows = []
    for section_rows in data.get("Sections", {}).values():
        rows.extend(dict(r) for r in section_rows)
    rows.sort(key=lambda x: x.get("_row_order", 10**12))
    for r in rows:
        r.pop("_row_order", None)
    all_cols = list(original_cols)
    for r in rows:
        for k in r:
            if k not in all_cols:
                all_cols.append(k)
    if "Total Check Status" not in all_cols:
        all_cols.append("Total Check Status")
    return pd.DataFrame(rows, columns=all_cols)


def _summarise_sheet(sheet: str, stmt_type: str, out_df: pd.DataFrame) -> dict:
    cp = out_df[out_df.get("COA Flag", "").astype(str).str.strip().eq("CP")]
    status = cp.get("Total Check Status", pd.Series(dtype=str)).fillna("").astype(str)
    return {
        "Sheet": sheet,
        "Statement Type": stmt_type,
        "Total Rows": len(out_df),
        "CP Rows": len(cp),
        "PASS":    int(status.eq("PASS").sum()),
        "FAIL":    int(status.str.contains("FAIL", case=False).sum()),
        "SKIPPED": int(status.str.startswith("Skipped").sum()),
        "BLANK":   int(status.str.strip().eq("").sum()),
    }


def validate_xlsx(xlsx_path: str, output_path: str = None) -> str:
    xlsx_path = str(xlsx_path)
    in_file = Path(xlsx_path)
    if not in_file.exists():
        raise FileNotFoundError(f"XLSX not found: {xlsx_path}")
    if output_path is None:
        output_path = str(in_file.with_name(in_file.stem + "_VALIDATED.xlsx"))

    print(f"\n[INFO] Input : {xlsx_path}")
    print(f"[INFO] Output: {output_path}\n")

    xl = pd.ExcelFile(xlsx_path, engine="openpyxl")
    out_sheets = OrderedDict()
    summary = []

    for sheet in xl.sheet_names:
        clean = _clean_sheet_name(sheet)
        stmt_type = SHEET_TO_STMT_TYPE.get(clean)
        df = pd.read_excel(
            xlsx_path, sheet_name=sheet, engine="openpyxl",
            dtype=object, keep_default_na=False,
        )
        original_cols = list(df.columns)
        if stmt_type is None:
            print(f"[SKIP ] {sheet:<10} → no stmt_type mapping")
            out_sheets[sheet] = df
            summary.append({
                "Sheet": sheet, "Statement Type": "", "Total Rows": len(df),
                "CP Rows": "", "PASS": "", "FAIL": "", "SKIPPED": "", "BLANK": "",
                "Note": "No RULES mapping",
            })
            continue
        try:
            data = _sheet_df_to_data(df)
            checked = run_total_check(data, stmt_type)
            out_df = _data_to_sheet_df(checked, original_cols)
            out_sheets[sheet] = out_df
            s = _summarise_sheet(sheet, stmt_type, out_df)
            s["Note"] = "Validated"
            summary.append(s)
            print(
                f"[OK   ] {sheet:<10} ({stmt_type:<20}) "
                f"CP={s['CP Rows']:<3} PASS={s['PASS']:<3} "
                f"FAIL={s['FAIL']:<3} SKIPPED={s['SKIPPED']:<3}"
            )
        except Exception as e:
            import traceback
            print(f"[ERROR] {sheet}: {e}")
            traceback.print_exc()
            out_sheets[sheet] = df
            summary.append({
                "Sheet": sheet, "Statement Type": stmt_type,
                "Total Rows": len(df), "CP Rows": "", "PASS": "",
                "FAIL": "", "SKIPPED": "", "BLANK": "",
                "Note": f"ERROR: {e}",
            })

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        for sheet, sdf in out_sheets.items():
            sdf.to_excel(writer, sheet_name=str(sheet)[:31], index=False)
        pd.DataFrame(summary).to_excel(
            writer, sheet_name="VALIDATION_SUMMARY", index=False
        )

    print(f"\n[DONE] Validated workbook written: {output_path}")
    return output_path


def main():
    xlsx_path = r"D:\S2_Khushbu\AI Projects\Financial Data Extraction\MDB_multipleAI\05_Manual_Validation_Required\ROCKINGHAM COUNTY_2025\ROCKINGHAM COUNTY_2025_All_Statements.xlsx"
    if len(sys.argv) >= 2:
        xlsx_path = sys.argv[1]
    validate_xlsx(xlsx_path)


if __name__ == "__main__":
    main()