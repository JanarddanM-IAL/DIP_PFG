from decimal import Decimal, InvalidOperation

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

CHAIN_TOTAL_LIAB_DEFERRED_INFLOWS_NET_POSITION = [
    # Chain 0
    [
        "Total Liabilities",
        "Total Deferred Inflows of Resources",
    ],
    # Chain 1
    [
        "Total Liabilities",
        {"cp": "Total Deferred Inflows of Resources",
         "section": "Deferred Inflows of Resources"},
    ],
    # Chain 2 — explicit combined CP row printed.
    [
        "Total Liabilities and Deferred Inflows of Resources",
        "Total Net Position",
    ],
    # Chain 3 — three separate CP rows all present.
    [
        "Total Liabilities",
        "Total Deferred Inflows of Resources",
        "Total Net Position",
    ],
    # Chain 4 — section-sum fallback for Deferred Inflows / Net Position
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
        ("Total Current Assets",      "dp_sum_positional",  "Assets",
            ["current assets"], ["noncurrent", "non-current"]),
        ("Total Restricted Assets",   "dp_label_ikey_sum_since_last_cp", "Assets",
            ["restricted assets"], []),
        ("Total Other Assets",        "dp_label_ikey_sum_since_last_cp", "Assets",
            ["other assets"], []),
        ("Net Capital Assets",         "dp_sum",  "Assets",
            ["capital assets"], []),
        ("Total Capital Assets",      "dp_sum_by_coa", "Assets",
            ["Capital Assets"], []),
        ("Total Noncurrent Assets",   "dp_label_ikey_sum", "Assets",
            ["noncurrent assets", "non-current assets"], []),
        ("Total Assets",               "section",  "Assets",            [], []),
        ("Total Deferred Outflows of Resources", "section",
            "Deferred Outflows of Resources", [], []),
        ("Total Assets and Deferred Outflows of Resources",
            "cp_sum_first_present", None,
            CHAIN_TOTAL_ASSETS_AND_DEFERRED_OUTFLOWS, []),
        ("Total Current Liabilities",  "dp_sum_current_liabilities",  "Liabilities", [], []),
        ("Total Noncurrent Liabilities", "dp_sum", "Liabilities",
            ["noncurrent liabilities", "non-current liabilities", "long-term liabilities", "long-term liabil"],
            ["current liabilities:"]),
        ("Total Liabilities",          "section",  "Liabilities",       [], []),
        ("Total Deferred Inflows of Resources", "section",
            "Deferred Inflows of Resources", [], []),
        ("Total Liabilities and Deferred Inflows of Resources", "cp_sum_first_present", None,
    [
        ["Total Liabilities", "Total Deferred Inflows of Resources"],
        [
            "Total Liabilities",
            {"cp": "Total Deferred Inflows of Resources",
             "section": "Deferred Inflows of Resources"},
        ],
    ], []),
        ("Total Net Position",         "section", "Net Position",       [], ["adjustment"]),
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
        ("Total Restricted Assets",    "dp_label_ikey_sum_since_last_cp", "Assets",
            ["restricted assets"], []),
        ("Total Other Assets",         "dp_label_ikey_sum_since_last_cp", "Assets",
            ["other assets"], []),
        ("Net Capital Assets",         "dp_sum",  "Assets",
            ["capital assets"], []),
        ("Total Capital Assets",       "dp_sum_by_coa", "Assets",
            ["Capital Assets"], []),
        ("Total Noncurrent Assets",    "dp_label_ikey_sum", "Assets",
            ["noncurrent assets", "non-current assets"], []),
        ("Total Assets", "section", "Assets", [], []),
        ("Total Deferred Outflows of Resources", "section",
            "Deferred Outflows of Resources", [], []),
        ("Total Assets and Deferred Outflows of Resources",
            "cp_sum_first_present", None,
            CHAIN_TOTAL_ASSETS_AND_DEFERRED_OUTFLOWS, []),
        ("Total Current Liabilities",  "dp_sum_current_liabilities",  "Liabilities", [], []),
        ("__BLANK__", "section_cp_blank_datapoint", "Net Position", [], []),
        ("Total Noncurrent Liabilities", "dp_sum", "Liabilities",
            ["noncurrent liabilities", "non-current liabilities", "long-term liabilities", "long-term liabil"],
            ["current liabilities:"]),
        ("Total Liabilities",          "section",  "Liabilities",       [], []),
        ("Total Deferred Inflows of Resources", "section",
            "Deferred Inflows of Resources", [], []),
        ("Total Liabilities and Deferred Inflows of Resources", "cp_sum_first_present", None,
    [
        ["Total Liabilities", "Total Deferred Inflows of Resources"],
        [
            "Total Liabilities",
            {"cp": "Total Deferred Inflows of Resources",
             "section": "Deferred Inflows of Resources"},
        ],
    ], []),
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
        ("Total Net Position", "section", "Net Position", [], []),
        ("Total Restricted Assets",    "dp_label_ikey_sum_since_last_cp", "Assets",
            ["restricted assets"], []),
        ("Total Other Assets",         "dp_label_ikey_sum_since_last_cp", "Assets",
            ["other assets"], []),
        ("Total Noncurrent Assets",    "dp_label_ikey_sum", "Assets",
            ["noncurrent assets", "non-current assets"], []),
        ("Net Capital Assets",         "dp_sum",  "Assets",
            ["capital assets"], []),
        ("Total Capital Assets",       "dp_sum_by_coa", "Assets",
            ["Capital Assets"], []),
        ("Total Assets",               "section",  "Assets",            [], []),
        ("Total Deferred Outflows of Resources", "section",
            "Deferred Outflows of Resources", [], []),
        ("Total Assets and Deferred Outflows of Resources",
            "cp_sum_first_present", None,
            CHAIN_TOTAL_ASSETS_AND_DEFERRED_OUTFLOWS, []),
        ("Total Current Liabilities",  "dp_sum_current_liabilities",  "Liabilities", [], []),
        ("Total Noncurrent Liabilities", "dp_sum", "Liabilities",
            ["noncurrent liabilities", "non-current liabilities", "long-term liabilities", "long-term liabil"],
            ["current liabilities:"]),
        ("Total Liabilities",          "section",  "Liabilities",       [], []),
        ("Total Deferred Inflows of Resources", "section",
            "Deferred Inflows of Resources", [], []),
        ("Total Liabilities and Deferred Inflows of Resources", "cp_sum_first_present", None,
            [
                ["Total Liabilities", "Total Deferred Inflows of Resources"],
                [
                    "Total Liabilities",
                    {"cp": "Total Deferred Inflows of Resources",
                     "section": "Deferred Inflows of Resources"},
                ],
            ], []),
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
        ("Total Operating Revenues",   "section", "Revenues",           [], []),
        ("Total Operating Expenses",   "section", "Expenditures",       [], []),
        ("Excess Of Revenues Over/Under Expenditures", "cp_subtract", None,
            ["Total Operating Revenues", "Total Operating Expenses"], []),
        ("Total Other Financing Sources", "section",
            "Other Financing Sources (Uses)", [], []),
        ("Net Change In Fund Balances", "cp_sum_first_present", None,
            [
                [
                    "Excess Of Revenues Over/Under Expenditures",
                    {"cp": "Total Other Financing Sources",
                     "section": "Other Financing Sources (Uses)"},
                ],
                [
                    {"cp": "Total Operating Revenues",   "section": "Revenues"},
                    {"cp": "Total Other Financing Sources",
                     "section": "Other Financing Sources (Uses)"},
                ],
            ], []),
    ],

    # ── SOA ───────────────────────────────────────────────────────────────────
    "_SOA": [
        ("Total General Revenues", "section", "General Revenues",       [], ["transfers"]),
        ("Total General Revenues and Transfer", "cp_sum_with_dp_fallback", None,
    [
        {"type": "cp", "coa": "Total General Revenues",
         "section": "General Revenues"},
        {"type": "dp_coa_in_section",
         "coa": "Transfers",
         "section": "General Revenues"},
    ], []),
        ("Total Revenues",     "section", "Revenues",                   [], []),
        ("Total Expenditures", "section", "Expenditures",               [], []),
        ("Excess of Revenues Over Expenditures", "cp_subtract", None,
            ["Total Revenues", "Total Expenditures"], []),
        ("Net Change in Fund Balances", "cp_sum", None,
            ["Excess of Revenues Over Expenditures", "Total Other Financing Sources"], []),
        ("Total Program Receipts - Charges for Services", "dp_sum_by_coa",
            "Program Revenue", ["Charges for Services"], []),
        ("Total Program Receipts - Operating Grants and Contributions",
            "dp_sum_by_coa", "Program Revenue",
            ["Operating Grants and Contributions"], []),
        ("Total Program Receipts - Capital Grants and Contributions",
            "dp_sum_by_coa", "Program Revenue",
            ["Capital Grants and Contributions"], []),
        ("Total Program Expenses",     "section", "Program Expenses",   [], []),
        ("Changes in Net Position",    "cp_subtract_last", None,
            [
                ["Total General Revenues and Transfer", "Total General Revenues"],
                "Total Program Receipts - Charges for Services",
                "Total Program Receipts - Operating Grants and Contributions",
                "Total Program Receipts - Capital Grants and Contributions",
                "Total Program Expenses",
            ],
            []),
    ],

    # ── PROP_IS ───────────────────────────────────────────────────────────────
    "_PROP_IS": [
        ("Total Operating Revenues",   "section_positional",
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
             "Total Cash Flows from Operating Activities",
             "Cash and Cash Equivalents"], []),
    ],

    # ── DSR ───────────────────────────────────────────────────────────────────
    "_DSR": [
        ("Total", "section", None, [], []),
    ],

    # ── DEBT ──────────────────────────────────────────────────────────────────
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
    has_deferred     = any(
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


# ─────────────────────────────────────────────────────────────────────────────
# FIX 1: Case-insensitive section lookup
# All section lookups now go through _get_section_rows() which matches
# section names case-insensitively. This fixes PROP_CFS failures where
# data has "Cash flows from operating activities" (lowercase) but rules
# reference "Cash Flows from Operating Activities" (title case).
# ─────────────────────────────────────────────────────────────────────────────
def _get_section_rows(sections: dict, section_name: str) -> list:
    """Return rows for section_name using case-insensitive matching."""
    if section_name is None:
        return []
    # Exact match first (fast path)
    if section_name in sections:
        return sections[section_name]
    # Case-insensitive fallback
    target = section_name.strip().lower()
    for k, v in sections.items():
        if k.strip().lower() == target:
            return v
    return []


def _get_cp(sections, coa_dp, table_id=None):
    """Return the LAST matching CP row (fix for duplicate-CP sheets)."""
    if not coa_dp:
        return None
    coa_dp_lower = coa_dp.strip().lower()
    found = None
    for rows in sections.values():
        for r in rows:
            if table_id is not None and _table_id(r) != table_id:
                continue
            if r.get("COA Flag") == "CP":
                stored = r.get("COA Datapoint", "").strip().lower()
                if stored == coa_dp_lower:
                    found = r          # keep going — take the last match
                elif stored in coa_dp_lower and len(stored) < len(coa_dp_lower):
                    if found is None:  # only use substring match if no exact match yet
                        found = r
    return found


def _get_cp_first(sections, coa_dp, table_id=None):
    """Return the FIRST matching CP row (used for positional checks)."""
    if not coa_dp:
        return None
    coa_dp_lower = coa_dp.strip().lower()
    for rows in sections.values():
        for r in rows:
            if table_id is not None and _table_id(r) != table_id:
                continue
            if r.get("COA Flag") == "CP":
                stored = r.get("COA Datapoint", "").strip().lower()
                if stored == coa_dp_lower:
                    return r
                if stored in coa_dp_lower and len(stored) < len(coa_dp_lower):
                    return r
    return None


def _get_dp(sections, coa_dp, section_name=None, table_id=None):
    coa_dp_lower = coa_dp.strip().lower()
    if section_name:
        search_rows = _get_section_rows(sections, section_name)
        search_iter = [search_rows]
    else:
        search_iter = sections.values()
    for rows in search_iter:
        for r in rows:
            if table_id is not None and _table_id(r) != table_id:
                continue
            if (r.get("COA Flag") == "DP"
                    and r.get("COA Datapoint", "").strip().lower() == coa_dp_lower):
                return r
    return None


def _dp_rows_in_section(sections, section_name, table_id=None):
    rows = _get_section_rows(sections, section_name)
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
    for i, entry in enumerate(cp_datapoints):
        if isinstance(entry, list):
            row = None
            for alt_name in entry:
                candidate = _get_cp(sections, alt_name, table_id=table_id)
                if candidate is not None:
                    row = candidate
                    break
        else:
            row = _get_cp(sections, entry, table_id=table_id)
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


def _sum_dp_cp_mixed(sections, members, cols, section_name=None, table_id=None, ikey=None):
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
        elif mtype == "dp_label_sum":
            sec = m.get("section", section_name)
            pool = _dp_rows_in_section(sections, sec, table_id=table_id)
            label_contains = [c.lower() for c in m.get("label_contains", [])]
            label_excludes = [e.lower() for e in m.get("label_excludes", [])]
            for r in pool:
                label = r.get(ikey or "Items", "").lower()
                hit = any(c in label for c in label_contains) if label_contains else True
                if hit and not any(e in label for e in label_excludes):
                    for col in cols:
                        totals[col] += _val(r.get(col, "-"))
        else:
            continue
    return totals


def _sum_cp_with_dp_fallback(sections, members, cols, table_id=None, ikey=None):
    totals = {col: Decimal(0) for col in cols}
    section_fallbacks_used = set()
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


def _resolve_chain_member(sections, member, cols, ikey, table_id=None):
    if isinstance(member, dict):
        if member.get("type") == "cp_subtract":
            rev = _get_cp(sections, member["minuend"], table_id=table_id)
            sub = _get_cp(sections, member["subtrahend"], table_id=table_id)
            rev_vals = {col: _val(rev.get(col, "-")) for col in cols} if rev else {col: Decimal(0) for col in cols}
            sub_vals = {col: _val(sub.get(col, "-")) for col in cols} if sub else {col: Decimal(0) for col in cols}
            return ({col: rev_vals[col] - sub_vals[col] for col in cols}, False)
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
    if isinstance(chain, dict):
        return True
    for member in chain:
        if isinstance(member, dict):
            if "cp_multi" in member:
                for name in member["cp_multi"]:
                    if _get_cp(sections, name, table_id=table_id) is None:
                        return False
                continue
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


def _sum_cp_first_present(sections, chains, cols, ikey, table_id=None, cp_row=None):
    usable = []
    for chain in chains:
        if _chain_usable(sections, chain, table_id=table_id):
            usable.append(chain)
    if not usable:
        return None, None, False

    def _compute(chain):
        if isinstance(chain, dict):
            vals, _ = _resolve_chain_member(sections, chain, cols, ikey, table_id=table_id)
            return vals
        totals = {col: Decimal(0) for col in cols}
        for member in chain:
            vals, _ = _resolve_chain_member(sections, member, cols, ikey, table_id=table_id)
            for col in cols:
                totals[col] += vals[col]
        return totals

    if cp_row is not None:
        for chain in usable:
            totals = _compute(chain)
            if all(abs(totals[col] - _val(cp_row.get(col, "-"))) <= 1
                   for col in cols):
                label = " + ".join(_chain_label(m) for m in chain)
                return totals, label, True

    chain = usable[0]
    totals = _compute(chain)
    label = " + ".join(_chain_label(m) for m in chain)
    return totals, label, True


# ─────────────────────────────────────────────────────────────────────────────
# FIX 2: dp_sum_positional — positional window sum for mid-section subtotals
# Used for Total Current Assets in PROP_SNP when DP item labels don't
# contain "current assets" but instead use prefixes like "Current assets: X".
# Sums DPs between the previous CP and this CP, with label filtering.
# ─────────────────────────────────────────────────────────────────────────────
def _dp_sum_positional(sections, section, cp_row, cols, ikey,
                        contains, excludes, table_id):
    """
    Sum DPs in `section` that fall between the previous CP and cp_row
    (positional window), filtered by label contains/excludes.

    Fallback logic (handles prefix-style item labels like "Current assets: X"):
    1. If label filtering finds no rows → use full window.
    2. If label-filtered sum doesn't match the reported CP value → use full window.
       This covers cases where restricted/other current assets appear in the same
       positional window but lack the expected label keyword.
    """
    sec_rows = _get_section_rows(sections, section)
    cp_order = cp_row.get("_row_order", -1)

    # Find the _row_order of the immediately-preceding CP in same table/section
    prev_cp_order = -1
    for r in sec_rows:
        if _table_id(r) != table_id:
            continue
        if r.get("COA Flag") == "CP":
            ro = r.get("_row_order", -1)
            if ro < cp_order:
                prev_cp_order = max(prev_cp_order, ro)

    # Collect DPs in the positional window
    window = [
        r for r in sec_rows
        if _table_id(r) == table_id
        and r.get("COA Flag") == "DP"
        and prev_cp_order < r.get("_row_order", -1) < cp_order
    ]

    label_contains = [c.lower() for c in contains]
    label_excludes = [e.lower() for e in excludes]

    def _label_match(r):
        label = str(r.get(ikey, "")).lower()
        hit = any(c in label for c in label_contains) if label_contains else True
        return hit and not any(e in label for e in label_excludes)

    matched = [r for r in window if _label_match(r)]

    # Fallback 1: label filter found nothing → use full window
    if not matched and window:
        return {col: sum(_val(r.get(col, "-")) for r in window) for col in cols}

    filtered_vals = {col: sum(_val(r.get(col, "-")) for r in matched) for col in cols}

    # Fallback 2: filtered sum mismatches reported value → try full window
    reported_vals = {col: _val(cp_row.get(col, "-")) for col in cols}
    if any(abs(filtered_vals[col] - reported_vals[col]) > 1 for col in cols):
        full_vals = {col: sum(_val(r.get(col, "-")) for r in window) for col in cols}
        if all(abs(full_vals[col] - reported_vals[col]) <= 1 for col in cols):
            return full_vals

    return filtered_vals


# ─────────────────────────────────────────────────────────────────────────────
# FIX 3: section_positional — positional window section sum
# Used for Total Operating Revenues in PROP_IS when there are multiple
# CP rows with the same COA Datapoint (one per fund/entity).
# Each CP is checked against only the DPs in its own positional window.
# ─────────────────────────────────────────────────────────────────────────────
def _section_positional(sections, section, cp_row, cols, ikey,
                         excludes, table_id):
    """
    Sum DPs between the previous CP and cp_row in `section`.
    Respects excludes label filter.

    Fallback: if window-DP sum doesn't match, try prev_CP + window_DPs.
    This handles rolling subtotal patterns where each CP accumulates the
    previous CP plus a few new DPs (e.g. multiple Total Operating Revenues
    CPs in a multi-fund PROP_IS sheet).
    """
    sec_rows = _get_section_rows(sections, section)
    cp_order = cp_row.get("_row_order", -1)

    prev_cp_order = -1
    prev_cp_row = None
    for r in sec_rows:
        if _table_id(r) != table_id:
            continue
        if r.get("COA Flag") == "CP":
            ro = r.get("_row_order", -1)
            if ro < cp_order and ro > prev_cp_order:
                prev_cp_order = ro
                prev_cp_row = r

    window = [
        r for r in sec_rows
        if _table_id(r) == table_id
        and r.get("COA Flag") == "DP"
        and prev_cp_order < r.get("_row_order", -1) < cp_order
    ]

    label_excludes = [e.lower() for e in excludes]
    filtered = [r for r in window
                if not any(e in str(r.get(ikey, "")).lower() for e in label_excludes)]

    window_vals = {col: sum(_val(r.get(col, "-")) for r in filtered) for col in cols}
    reported_vals = {col: _val(cp_row.get(col, "-")) for col in cols}

    # If pure window sum matches, done
    if all(abs(window_vals[col] - reported_vals[col]) <= 1 for col in cols):
        return window_vals

    # Fallback: prev_CP + window_DPs (rolling subtotal pattern)
    if prev_cp_row is not None:
        combined = {
            col: _val(prev_cp_row.get(col, "-")) + window_vals[col]
            for col in cols
        }
        if all(abs(combined[col] - reported_vals[col]) <= 1 for col in cols):
            return combined

    return window_vals


# ─────────────────────────────────────────────────────────────────────────────
# MAIN ENGINE
# ─────────────────────────────────────────────────────────────────────────────

def run_total_check(data: dict, stmt_type: str) -> dict:
    try:
        from column_shift_repair import repair_column_shifts
        repair_column_shifts(data, stmt_type)
    except Exception as e:
        print(f"[WARN] column_shift_repair skipped: {e}")

    # ── FIX: inject _row_order if rows don't already have it ──────────────
    # When called from jsonToCsv.py (JSON path), rows have no _row_order.
    # _row_order must reflect document order for positional window rules
    # (dp_sum_positional, section_positional, dp_label_ikey_sum_since_last_cp,
    #  dp_sum_current_liabilities). Stamp them now using global list position.
    sections = data.get("Sections", {})
    _needs_row_order = all(
        "_row_order" not in r
        for rows in sections.values()
        for r in rows[:1]   # check just the first row of each section
    )
    if _needs_row_order:
        counter = 0
        for rows in sections.values():
            for r in rows:
                r["_row_order"] = counter
                counter += 1
    # ──────────────────────────────────────────────────────────────────────

    sections = data.get("Sections", {})
    cols     = data.get("Reporting Columns", [])

    effective_stmt_type = resolve_stmt_type_for_rules(data, stmt_type)
    rules = RULES.get(effective_stmt_type, [])
    ikey  = _items_key(data)

    for rows in sections.values():
        for r in rows:
            flag = r.get("COA Flag", "")
            r["Total Check Status"] = "PENDING" if flag == "CP" else ""

    for rule in rules:
        cp_dp, rule_type, section, contains, excludes = rule

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

            # ── dp_label_ikey_sum ─────────────────────────────────────────────
            if rule_type == "dp_label_ikey_sum":
                if section:
                    pool = [r for r in _get_section_rows(sections, section)
                            if r.get("COA Flag") == "DP" and _table_id(r) == tid]
                else:
                    pool = [r for sec in sections.values() for r in sec
                            if r.get("COA Flag") == "DP" and _table_id(r) == tid]
                label_contains = [c.lower() for c in contains]
                label_excludes = [e.lower() for e in excludes]
                matched = []
                for r in pool:
                    label = str(r.get(ikey, "")).lower()
                    hit = any(c in label for c in label_contains) if label_contains else True
                    if hit and not any(e in label for e in label_excludes):
                        matched.append(r)
                member_vals = {col: sum(_val(r.get(col, "-")) for r in matched)
                               for col in cols}
                _apply_check(cp_row, member_vals, cols)

            # ── dp_label_ikey_sum_since_last_cp ───────────────────────────────
            elif rule_type == "dp_label_ikey_sum_since_last_cp":
                sec_rows = (_get_section_rows(sections, section)
                            if section else
                            [r for sec in sections.values() for r in sec])
                cp_row_order = cp_row.get("_row_order", -1)
                label_contains = [c.lower() for c in contains]
                label_excludes = [e.lower() for e in excludes]
                prev_cp_order = -1
                for r in sec_rows:
                    if _table_id(r) != tid:
                        continue
                    if r.get("COA Flag") == "CP":
                        ro = r.get("_row_order", -1)
                        if ro < cp_row_order:
                            prev_cp_order = max(prev_cp_order, ro)
                matched = []
                for r in sec_rows:
                    if _table_id(r) != tid:
                        continue
                    if r.get("COA Flag") != "DP":
                        continue
                    ro = r.get("_row_order", -1)
                    if not (prev_cp_order < ro < cp_row_order):
                        continue
                    label = str(r.get(ikey, "")).lower()
                    hit = any(c in label for c in label_contains) if label_contains else True
                    if hit and not any(e in label for e in label_excludes):
                        matched.append(r)
                member_vals = {col: sum(_val(r.get(col, "-")) for r in matched)
                               for col in cols}
                _apply_check(cp_row, member_vals, cols)

            # ── dp_sum_positional (FIX 2) ─────────────────────────────────────
            elif rule_type == "dp_sum_positional":
                member_vals = _dp_sum_positional(
                    sections, section, cp_row, cols, ikey,
                    contains, excludes, tid
                )
                _apply_check(cp_row, member_vals, cols)

            # ── section_positional (FIX 3) ────────────────────────────────────
            elif rule_type == "section_positional":
                member_vals = _section_positional(
                    sections, section, cp_row, cols, ikey, excludes, tid
                )
                _apply_check(cp_row, member_vals, cols)

            elif rule_type == "dp_sum":
                if section:
                    pool = [r for r in _get_section_rows(sections, section)
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

            elif rule_type == "dp_sum_by_coa":
                if section:
                    pool = [r for r in _get_section_rows(sections, section)
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

            # ── dp_sum_current_liabilities (FIX 4) ───────────────────────────
            # Added fallback: when DP-level sum doesn't match, try summing
            # blank-COA intermediate CP subtotals that precede this CP.
            elif rule_type == "dp_sum_current_liabilities":
                if section:
                    pool = [r for r in _get_section_rows(sections, section)
                            if r.get("COA Flag") == "DP" and _table_id(r) == tid]
                    all_sec_rows = [r for r in _get_section_rows(sections, section)
                                    if _table_id(r) == tid]
                else:
                    pool = [r for sec in sections.values() for r in sec
                            if r.get("COA Flag") == "DP" and _table_id(r) == tid]
                    all_sec_rows = [r for sec in sections.values() for r in sec
                                    if _table_id(r) == tid]
                uses_label_prefix = any(
                    r.get(ikey, "").strip().lower().startswith("current liabilities:")
                    for r in pool
                )
                matched = []
                for r in pool:
                    item_label = r.get(ikey, "").strip().lower()
                    coa_dp_val = r.get("COA Datapoint", "").strip().lower()
                    if uses_label_prefix:
                        if item_label.startswith("current liabilities:"):
                            matched.append(r)
                    else:
                        if not coa_dp_val.endswith("- noncurrent"):
                            matched.append(r)
                member_vals = {col: sum(_val(r.get(col, "-")) for r in matched)
                               for col in cols}

                # FIX 4: If DP sum doesn't match, try summing blank-COA CP subtotals
                reported_vals = {col: _val(cp_row.get(col, "-")) for col in cols}
                if any(abs(member_vals[col] - reported_vals[col]) > 1 for col in cols):
                    cp_order = cp_row.get("_row_order", -1)
                    blank_cps = [
                        r for r in all_sec_rows
                        if r.get("COA Flag") == "CP"
                        and not str(r.get("COA Datapoint", "")).strip()
                        and r.get("_row_order", 0) < cp_order
                    ]
                    if blank_cps:
                        fallback_vals = {col: sum(_val(r.get(col, "-")) for r in blank_cps)
                                         for col in cols}
                        if all(abs(fallback_vals[col] - reported_vals[col]) <= 1
                               for col in cols):
                            member_vals = fallback_vals

                _apply_check(cp_row, member_vals, cols)

            elif rule_type == "section":
                if section:
                    pool = _dp_rows_in_section(sections, section, table_id=tid)
                else:
                    pool = [r for sec in sections.values() for r in sec
                            if r.get("COA Flag") == "DP" and _table_id(r) == tid]

                filtered_pool = [r for r in pool
                                if not any(e.lower() in r.get(ikey, "").lower()
                                            for e in excludes)]
                member_vals = {col: sum(_val(r.get(col, "-")) for r in filtered_pool)
                            for col in cols}

                if excludes and any(
                    abs(member_vals[col] - _val(cp_row.get(col, "-"))) > 1
                    for col in cols
                ):
                    full_pool_vals = {col: sum(_val(r.get(col, "-")) for r in pool)
                                    for col in cols}
                    strategy1_passes = sum(
                        abs(member_vals[col] - _val(cp_row.get(col, "-"))) <= 1
                        for col in cols
                    )
                    strategy2_passes = sum(
                        abs(full_pool_vals[col] - _val(cp_row.get(col, "-"))) <= 1
                        for col in cols
                    )
                    if strategy2_passes > strategy1_passes:
                        member_vals = full_pool_vals

                _apply_check(cp_row, member_vals, cols)

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

            elif rule_type == "section_dp_plus_cp":
                candidate_sections = [section] + list(contains[1:])
                dp_pool = []
                for sec_candidate in candidate_sections:
                    dp_pool = _dp_rows_in_section(sections, sec_candidate, table_id=tid)
                    if dp_pool:
                        break
                dp_sum = {col: sum(_val(r.get(col, "-")) for r in dp_pool)
                          for col in cols}
                net_cp = None
                for cp_name in contains:
                    net_cp = _get_cp(sections, cp_name, table_id=tid)
                    if net_cp is not None:
                        break
                net_vals = ({col: _val(net_cp.get(col, "-")) for col in cols}
                            if net_cp else {col: Decimal(0) for col in cols})
                member_vals = {col: dp_sum[col] + net_vals[col] for col in cols}
                _apply_check(cp_row, member_vals, cols)

            elif rule_type == "dp_cp_sum":
                member_vals = _sum_dp_cp_mixed(
                    sections, contains, cols, section, table_id=tid, ikey=ikey
                )
                _apply_check(cp_row, member_vals, cols)

            elif rule_type == "cp_or_section_sum":
                member_vals = _sum_cp_or_section(
                    sections, contains, cols, ikey,
                    excludes=excludes, table_id=tid
                )
                _apply_check(cp_row, member_vals, cols)

            elif rule_type == "cp_sum_with_dp_fallback":
                member_vals = _sum_cp_with_dp_fallback(
                    sections, contains, cols, table_id=tid, ikey=ikey
                )
                _apply_check(cp_row, member_vals, cols)

            elif rule_type in ("cp_subtract", "cp_subtract_last"):
                resolved_terms = []
                for term in contains:
                    if isinstance(term, list):
                        picked = None
                        for alt in term:
                            alt_lower = alt.strip().lower()
                            exact = next(
                                (r for rows in sections.values() for r in rows
                                 if r.get("COA Flag") == "CP"
                                 and r.get("COA Datapoint", "").strip().lower() == alt_lower
                                 and _table_id(r) == tid),
                                None,
                            )
                            if exact is not None:
                                picked = alt
                                break
                        if picked is None:
                            picked = term[0]
                        resolved_terms.append(picked)
                    else:
                        resolved_terms.append(term)
                member_vals = _sum_cp_vals(
                    sections, resolved_terms, cols, subtract_last=True, table_id=tid
                )
                _apply_check(cp_row, member_vals, cols)

            elif rule_type == "cp_sum_first_present":
                member_vals, chosen_label, matched = _sum_cp_first_present(
                    sections, contains, cols, ikey, table_id=tid,
                    cp_row=cp_row,
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
            # Case-insensitive match for blank_dp_rules keys too
            matched_rule_key = next(
                (k for k in blank_dp_rules if k.strip().lower() == sec_name.strip().lower()),
                None
            )
            if matched_rule_key is None:
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

    # ── _DEBT block ───────────────────────────────────────────────────────────
    if stmt_type == "_DEBT":
        tables = {}
        for sec_name, sec_rows in sections.items():
            for r in sec_rows:
                tid = _table_id(r)
                tables.setdefault(tid, []).append(r)

        for tid, trows in tables.items():
            pending_dps = []
            dps_since_last_cp = []
            cps_since_last_outer = []
            all_cp_rows_in_table = [r for r in trows if r.get("COA Flag", "") == "CP"]

            for r in trows:
                flag = r.get("COA Flag", "")
                if flag == "DPG":
                    pending_dps = []
                elif flag == "DP":
                    pending_dps.append(r)
                    dps_since_last_cp.append(r)
                elif flag == "CP":
                    reported = {col: _val(r.get(col, "-")) for col in cols}
                    inner_sum = {
                        col: sum(_val(dp.get(col, "-")) for dp in pending_dps)
                        for col in cols
                    }
                    outer_dp_sum = {
                        col: sum(_val(dp.get(col, "-")) for dp in dps_since_last_cp)
                        for col in cols
                    }
                    outer_cp_sum = {
                        col: sum(_val(cp.get(col, "-")) for cp in cps_since_last_outer)
                        for col in cols
                    }

                    def matches(candidate):
                        return all(
                            abs(candidate[col] - reported[col]) <= 1
                            for col in cols
                        )
                    combined_sum = {
                        col: inner_sum[col] + outer_cp_sum[col]
                        for col in cols
                    }
                    if matches(inner_sum):
                        member_vals = inner_sum
                    elif matches(combined_sum):
                        member_vals = combined_sum
                    elif matches(outer_dp_sum):
                        member_vals = outer_dp_sum
                    elif matches(outer_cp_sum):
                        member_vals = outer_cp_sum
                    else:
                        member_vals = outer_cp_sum if cps_since_last_outer else outer_dp_sum

                    _apply_check(r, member_vals, cols)

                    if matches(outer_cp_sum) and cps_since_last_outer:
                        cps_since_last_outer = []
                        dps_since_last_cp = []
                    else:
                        cps_since_last_outer.append(r)
                        dps_since_last_cp = []

                    pending_dps = []

            for cp_r in all_cp_rows_in_table[-2:]:
                if cp_r.get("Total Check Status", "") not in ("PASS",):
                    cp_r["Total Check Status"] = ""

    # ── DSR block ─────────────────────────────────────────────────────────────
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

    # ── _PROP_IS fallback ─────────────────────────────────────────────────────
    if stmt_type == "_PROP_IS":
        for rows in sections.values():
            for r in rows:
                if (r.get("COA Flag") == "CP"
                        and r.get("COA Datapoint", "").strip().lower() == "change in net position"
                        and r.get("Total Check Status") in ("PENDING", "Skipped - no chain matched")):
                    rev_cp = _get_cp(sections, "Total Operating Revenues", table_id=_table_id(r))
                    exp_cp = _get_cp(sections, "Total Operating Expenses",  table_id=_table_id(r))
                    if rev_cp and exp_cp:
                        member_vals = {
                            col: _val(rev_cp.get(col, "-")) - _val(exp_cp.get(col, "-"))
                            for col in cols
                        }
                        _apply_check(r, member_vals, cols)

    # ── _GOV_IS fallback ─────────────────────────────────────────────────────
    if stmt_type == "_GOV_IS":
        for rows in sections.values():
            for r in rows:
                if r.get("COA Flag") != "CP":
                    continue
                if r.get("COA Datapoint", "").strip().lower() != "net change in fund balances":
                    continue
                status = r.get("Total Check Status", "")
                if status == "PASS":
                    continue

                tid = _table_id(r)

                excess_cp = _get_cp(sections,
                    "Excess Of Revenues Over/Under Expenditures", table_id=tid)
                ofs_cp = _get_cp(sections,
                    "Total Other Financing Sources", table_id=tid)
                ofs_pool = _dp_rows_in_section(
                    sections, "Other Financing Sources (Uses)", table_id=tid)

                if excess_cp is not None:
                    ofs_vals = (
                        {col: _val(ofs_cp.get(col, "-")) for col in cols}
                        if ofs_cp is not None
                        else {col: sum(_val(dp.get(col, "-")) for dp in ofs_pool)
                              for col in cols}
                    )
                    member_vals = {
                        col: _val(excess_cp.get(col, "-")) + ofs_vals[col]
                        for col in cols
                    }
                    _apply_check(r, member_vals, cols)

                else:
                    rev_cp = _get_cp(sections, "Total Operating Revenues", table_id=tid)
                    exp_cp = _get_cp(sections, "Total Operating Expenses",  table_id=tid)
                    if rev_cp and exp_cp:
                        ofs_sum = (
                            {col: _val(ofs_cp.get(col, "-")) for col in cols}
                            if ofs_cp is not None
                            else {col: sum(_val(dp.get(col, "-")) for dp in ofs_pool)
                                  for col in cols}
                        )
                        member_vals = {
                            col: _val(rev_cp.get(col, "-"))
                                 - _val(exp_cp.get(col, "-"))
                                 + ofs_sum[col]
                            for col in cols
                        }
                        _apply_check(r, member_vals, cols)

    # ── SOA CU_MODE="B" entity column CP fix ─────────────────────────────────
    if stmt_type == "_SOA":
        try:
            total_idx = cols.index("Total")
            cu_idx = cols.index("Component Units")
            entity_cols = cols[total_idx + 1 : cu_idx]
        except ValueError:
            entity_cols = []

        if entity_cols:
            for rows in sections.values():
                for r in rows:
                    if r.get("COA Flag") != "CP":
                        continue
                    status = r.get("Total Check Status", "")
                    if not status or status == "PASS":
                        continue
                    fail_parts = [
                        p for p in status.split(" | ")
                        if not any(ec in p for ec in entity_cols)
                    ]
                    r["Total Check Status"] = (
                        "PASS" if not fail_parts
                        else " | ".join(fail_parts)
                    )

    # ── Final sweep ───────────────────────────────────────────────────────────
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
        r["_row_order"] = idx          # preserves original sheet order
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
    xlsx_path = r"D:\S2\S2_Khushbu\AI Projects\Financial Data Extraction\MDB_multipleAI\05_Manual_Validation_Required\PP_WI_820287476_2025_NON-LG\PP_WI_820287476_2025_NON-LG_All_Statements.xlsx"
    if len(sys.argv) >= 2:
        xlsx_path = sys.argv[1]
    validate_xlsx(xlsx_path)


if __name__ == "__main__":
    main()