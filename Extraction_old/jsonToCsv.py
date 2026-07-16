import json
import csv
from pathlib import Path
from collections import OrderedDict
from decimal import Decimal, InvalidOperation
import re

CHECK_STATUS_COL = "Total Check Status"

METADATA_COLUMNS = {
    "Statement", "Issuer Name", "FYE", "Page No", "Currency reported", "Section",
}
NON_NUMERIC_ITEM_COLUMNS = {
    "COA Flag", "Line Items", "SNP Items", "SOA Items",
    "Balance Sheet Items", "Row Items", "COA Datapoint",
}


# ─────────────────────────────────────────────────────────────────────────────
# NUMERIC HELPERS
# ─────────────────────────────────────────────────────────────────────────────

def parse_amount(value) -> Decimal:
    if value is None:
        return Decimal("0")
    s = str(value).strip()
    if s in {"", "-", "–", "—"} or set(s) == {"#"}:
        return Decimal("0")
    s = s.replace("$", "").replace(",", "").replace(" ", "")
    negative = False
    if s.startswith("(") and s.endswith(")"):
        negative, s = True, s[1:-1]
    if s.startswith("-"):
        negative, s = True, s[1:]
    if s in {"", "-", "–", "—"}:
        return Decimal("0")
    try:
        num = Decimal(s)
    except InvalidOperation:
        return Decimal("0")
    return -num if negative else num


def format_amount(value: Decimal) -> str:
    if value is None:
        return ""
    value = Decimal(value)
    return f"{int(value):,}" if value == value.to_integral_value() else f"{value:,.2f}"


def is_numeric_reporting_column(col: str, csv_rows: list[dict]) -> bool:
    if col in METADATA_COLUMNS or col in NON_NUMERIC_ITEM_COLUMNS:
        return False
    for row in csv_rows:
        v = str(row.get(col) or "").strip()
        if not v or v in {"-", "–", "—"}:
            continue
        cleaned = re.sub(r"[$,\s()\-]", "", v)
        if cleaned.replace(".", "", 1).isdigit():
            return True
    return False


# ─────────────────────────────────────────────────────────────────────────────
# ROW LABEL HELPER
# ─────────────────────────────────────────────────────────────────────────────

def get_row_label(row: dict) -> str:
    for key in ["Row Items", "Line Items", "SNP Items", "SOA Items", "Balance Sheet Items"]:
        v = row.get(key)
        if v:
            return str(v).strip()
    return ""


# ─────────────────────────────────────────────────────────────────────────────
# COLUMN DISCOVERY
# ─────────────────────────────────────────────────────────────────────────────

def normalize_value(value):
    if value is None:
        return ""
    if isinstance(value, (dict, list)):
        return json.dumps(value, ensure_ascii=False)
    return value


def discover_dynamic_columns(sections: dict) -> list[str]:
    discovered = OrderedDict()
    for _, items in sections.items():
        if not isinstance(items, list):
            continue
        for item in items:
            if not isinstance(item, dict):
                continue
            for key in item.keys():
                if key not in discovered:
                    discovered[key] = True
    return list(discovered.keys())


def order_item_columns(dynamic_item_columns: list[str]) -> list[str]:
    priority = ["COA Flag", "Row Items", "Line Items", "SNP Items", "SOA Items", "Balance Sheet Items"]
    # Exclude CHECK_STATUS_COL — it is always pinned last separately
    filtered = [c for c in dynamic_item_columns if c != CHECK_STATUS_COL]
    ordered  = [c for c in priority if c in filtered]
    ordered += [c for c in filtered if c not in ordered]
    return ordered


# ─────────────────────────────────────────────────────────────────────────────
# MAIN PIPELINE
# ─────────────────────────────────────────────────────────────────────────────

def run_json_to_csv_pipeline(json_path: str, csv_path: str) -> bool:
    """
    Convert normalised JSON to a wide CSV.

    Returns
    -------
    bool
        True  — all CP rows have Total Check Status == "PASS" or "Skipped …"
        False — at least one CP row has a FAIL status (set by the LLM).
    """
    json_path = Path(json_path)
    csv_path  = Path(csv_path)

    if not json_path.exists():
        raise FileNotFoundError(f"JSON file not found: {json_path}")

    with open(json_path, "r", encoding="utf-8") as f:
        data = json.load(f)

    metadata = data.get("Metadata", {})
    sections = data.get("Sections", {})

    if not isinstance(sections, dict) or not sections:
        raise ValueError("Invalid or empty 'Sections' found in JSON")

    statement         = metadata.get("Statement", "")
    issuer_name       = metadata.get("Issuer Name", "")
    fye               = metadata.get("FYE", "")
    page_no           = metadata.get("Page No", "")
    currency_reported = metadata.get("Currency reported", "")

    dynamic_item_columns = discover_dynamic_columns(sections)
    ordered_item_columns = order_item_columns(dynamic_item_columns)

    base_fieldnames = [
        "Statement", "Issuer Name", "FYE", "Page No", "Currency reported", "Section",
    ] + ordered_item_columns + [CHECK_STATUS_COL]

    csv_rows = []
    for section_name, items in sections.items():
        if not isinstance(items, list):
            continue
        for item in items:
            if not isinstance(item, dict):
                continue
            row = {
                "Statement":         statement,
                "Issuer Name":       issuer_name,
                "FYE":               fye,
                "Page No":           page_no,
                "Currency reported": currency_reported,
                "Section":           section_name,
            }
            for col in ordered_item_columns:
                row[col] = normalize_value(item.get(col))
            # Always carry Total Check Status from JSON; default to "" if absent
            row[CHECK_STATUS_COL] = normalize_value(item.get(CHECK_STATUS_COL))
            csv_rows.append(row)

    if not csv_rows:
        raise RuntimeError("No rows generated for CSV")

    csv_path.parent.mkdir(parents=True, exist_ok=True)
    with open(csv_path, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=base_fieldnames, extrasaction="ignore")
        writer.writeheader()
        writer.writerows(csv_rows)

    print(f" Dynamic CSV saved at: {csv_path}", flush=True)

    # ── Derive all_pass from the LLM-set Total Check Status on CP rows ───────
    # "PASS" and "Skipped …" sentinels are both acceptable.
    # Empty string (DP / DPG rows) is neutral — never a failure.
    # Any other value (e.g. "Governmental Activities: FAIL") is a failure.
    ACCEPTABLE = {"", "PASS", "Skipped - no membership rule defined"}
    all_pass = all(
        row.get(CHECK_STATUS_COL, "") in ACCEPTABLE
        for row in csv_rows
        if str(row.get("COA Flag", "")).strip().upper() == "CP"
    )

    print(f" Validation result  : {'ALL PASS' if all_pass else 'FAIL(S) DETECTED'}", flush=True)
    return all_pass