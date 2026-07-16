import json
import csv
from pathlib import Path
from collections import OrderedDict
from decimal import Decimal, InvalidOperation
import re
import openpyxl
from openpyxl.utils import get_column_letter
import shutil
import copy
from total_check_engine import run_total_check  # ← runs CP-vs-DP arithmetic, fills real PASS/FAIL

CHECK_STATUS_COL = "Total Check Status"

# Infers the RULES dict key (e.g. "_GOV_BS") from a JSON filename stem like
# "DYER_COUNTY_2024_GOV_BS_p12-13". Mirrors the statement-type vocabulary
# already used by _STMT_CSV_RE / STMT_ORDER further down this file.
_STMT_JSON_RE = re.compile(
    r"^(.+?)_(SNP|SOA|GOV_BS|GOV_IS|PROP_SNP|PROP_IS|PROP_CFS|DSR|DEBT)"
    r"(?:_Comp)?_p[\d\-]+$",
    re.IGNORECASE,
)


def infer_stmt_type(json_stem: str) -> str:
    """Return the RULES dict key (e.g. '_GOV_BS') for a JSON filename stem,
    or '' if no recognised statement-type token is found."""
    m = _STMT_JSON_RE.match(json_stem)
    if m:
        return "_" + m.group(2).upper()
    return ""

METADATA_COLUMNS = {
    "Statement", "Issuer Name", "FYE", "Page No", "Currency reported", "Section",
}

NON_NUMERIC_ITEM_COLUMNS = {
    "COA Flag", "Line Items", "SNP Items", "SOA Items",
    "Balance Sheet Items", "Row Items", "COA Datapoint",
    "GOV_BS Items", "GOV_IS Items",
    "PROP_SNP Items", "PROP_IS Items",
    "PROP_CFS Items", "DSR Items", "DEBT Items",
    "Fund Items", "Cash Flow Items",
}


def parse_amount(value) -> Decimal:
    if value is None:
        return Decimal("0")
    s = str(value).strip()
    if s in {"", "-", "\u2013", "\u2014"} or set(s) == {"#"}:
        return Decimal("0")
    s = s.replace("$", "").replace(",", "").replace(" ", "")
    negative = False
    if s.startswith("(") and s.endswith(")"):
        negative, s = True, s[1:-1]
    if s.startswith("-"):
        negative, s = True, s[1:]
    if s in {"", "-", "\u2013", "\u2014"}:
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
        if not v or v in {"-", "\u2013", "\u2014"}:
            continue
        cleaned = re.sub(r"[$,\s()\-]", "", v)
        if cleaned.replace(".", "", 1).isdigit():
            return True
    return False


def get_row_label(row: dict) -> str:
    for key in ["Row Items", "Line Items", "SNP Items", "SOA Items", "Balance Sheet Items"]:
        v = row.get(key)
        if v:
            return str(v).strip()
    return ""


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
    filtered = [c for c in dynamic_item_columns if c != CHECK_STATUS_COL]
    ordered  = [c for c in priority if c in filtered]
    ordered += [c for c in filtered if c not in ordered]
    return ordered


import copy  # ← add this import near the top of jsonToCsv.py, with the other imports


def run_json_to_csv_pipeline(json_path: str, csv_path: str) -> bool:
    json_path = Path(json_path)
    csv_path  = Path(csv_path)

    if not json_path.exists():
        raise FileNotFoundError(f"JSON file not found: {json_path}")

    with open(json_path, "r", encoding="utf-8", errors="replace") as f:
        data = json.load(f)

    metadata = data.get("Metadata", {})
    sections = data.get("Sections", {})

    if not isinstance(sections, dict) or not sections:
        raise ValueError("Invalid or empty 'Sections' found in JSON")

    # ── Run the arithmetic engine on a COPY, for CSV purposes only ───────
    # "Total Check Status" is CSV-only output. The on-disk JSON is never
    # mutated or rewritten here — it stays exactly as the LLM/pipeline
    # produced it. run_total_check() is run against a deepcopy so its
    # in-place row mutations can't leak back into `data`.

    # ── Infer statement type up front (needed for BOTH repair + total check) ──
    stmt_type = infer_stmt_type(json_path.stem)
    if not stmt_type:
        print(f" [WARN] Could not infer statement type from filename "
              f"'{json_path.name}' — skipping Total Check arithmetic and "
              f"column-shift repair; all CP rows will read as "
              f"Total Check Status = \"\".",
              flush=True)
        checked_sections = sections
    else:
        # ── 1. Column-shift repair runs on the REAL `data` so fixes
        #      persist back to the on-disk JSON (this is intentional —
        #      it's a data-correctness fix, not a derived metric). ──
        try:
            from column_shift_repair import repair_column_shifts
            repair_column_shifts(data, stmt_type)
        except ImportError:
            print(" [WARN] column_shift_repair.py not found — skipping repair.",
                  flush=True)
        except Exception as e:
            print(f" [WARN] column_shift_repair failed: {e}", flush=True)

        # ── 2. Persist repaired JSON back to disk (only if any repair fired) ──
        if data.get("_column_shift_repairs"):
            json_path.write_text(
                json.dumps(data, ensure_ascii=False, indent=2),
                encoding="utf-8",
            )
            print(f"   [REPAIR] {len(data['_column_shift_repairs'])} "
                  f"column-shift fix(es) written back to {json_path.name}",
                  flush=True)
            # Refresh `sections` reference so downstream CSV build sees the swap
            sections = data.get("Sections", {})

        # ── 3. Total Check still runs on a deepcopy so its CSV-only
        #      "Total Check Status" mutations don't leak into the JSON. ──
        checked_data = run_total_check(copy.deepcopy(data), stmt_type)
        checked_sections = checked_data.get("Sections", {})


    statement         = metadata.get("Statement", "")
    issuer_name       = metadata.get("Issuer Name", "")
    fye               = metadata.get("FYE", "")
    page_no           = metadata.get("Page No", "")
    currency_reported = metadata.get("Currency reported", "")

    # Discover columns from the checked sections (Total Check Status is
    # added by the engine, so this still picks it up correctly for the
    # CSV header) — but fall back to the original sections shape either way.
    dynamic_item_columns = discover_dynamic_columns(checked_sections)
    ordered_item_columns = order_item_columns(dynamic_item_columns)

    base_fieldnames = [
        "Statement", "Issuer Name", "FYE", "Page No", "Currency reported", "Section",
    ] + ordered_item_columns + [CHECK_STATUS_COL]

    csv_rows = []
    for section_name, items in checked_sections.items():
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
            row[CHECK_STATUS_COL] = normalize_value(item.get(CHECK_STATUS_COL))
            csv_rows.append(row)

    if not csv_rows:
        raise RuntimeError("No rows generated for CSV")

    csv_path.parent.mkdir(parents=True, exist_ok=True)
    with open(csv_path, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=base_fieldnames, extrasaction="ignore")
        writer.writeheader()
        writer.writerows(csv_rows)

    #print(f" Dynamic CSV saved at: {csv_path}", flush=True)

    ACCEPTABLE = {"", "PASS", "Skipped - no membership rule defined",
                  "Skipped - no rule defined"}
    all_pass = all(
        row.get(CHECK_STATUS_COL, "") in ACCEPTABLE
        for row in csv_rows
        if str(row.get("COA Flag", "")).strip().upper() == "CP"
    )

    #print(f" Validation result  : {'ALL PASS' if all_pass else 'FAIL(S) DETECTED'}", flush=True)
    return all_pass

# ─────────────────────────────────────────────────────────────────────────────
# SHARED REGEX / CONSTANTS FOR MERGE FUNCTIONS
# ─────────────────────────────────────────────────────────────────────────────

_STMT_CSV_RE = re.compile(
    r"^(.+?)_(SNP|SOA|GOV_BS|GOV_IS|PROP_SNP|PROP_IS|PROP_CFS|DSR|DEBT)"
    r"(?:_Comp)?_p[\d\-]+\.csv$",
    re.IGNORECASE,
)

_ALL_CSV_RE = re.compile(r"_All\.csv$", re.IGNORECASE)

STMT_ORDER = {
    "SNP": 0, "SOA": 1, "GOV_BS": 2, "GOV_IS": 3,
    "PROP_SNP": 4, "PROP_IS": 5, "PROP_CFS": 6,
    "DEBT": 7, "DSR": 8,
}

PRIORITY_COLS = [
    "Statement", "Issuer Name", "FYE", "Page No",
    "Currency reported", "Section",
]


# ─────────────────────────────────────────────────────────────────────────────
# MERGE ALL CSVs PER ISSUER INTO ONE _All.csv
# ─────────────────────────────────────────────────────────────────────────────

def merge_all_csvs(folder: str, output_folder: str | None = None) -> list[str]:
    """
    Merge individual statement CSVs into one _All.csv per issuer.
    Scans subfolders recursively.
    """
    folder = Path(folder)
    if output_folder is None:
        output_folder = folder
    else:
        output_folder = Path(output_folder)
    output_folder.mkdir(parents=True, exist_ok=True)

    groups: dict[str, list[Path]] = {}

    for f in sorted(folder.rglob("*.csv")):
        if not f.is_file():
            continue
        if _ALL_CSV_RE.search(f.name):
            continue
        m = _STMT_CSV_RE.match(f.name)
        if not m:
            continue
        base_name = m.group(1)
        if base_name not in groups:
            groups[base_name] = []
        groups[base_name].append(f)

    if not groups:
        print(f"  No statement CSVs found in: {folder}")
        return []

    for base_name in groups:
        groups[base_name].sort(
            key=lambda p: STMT_ORDER.get(
                (_STMT_CSV_RE.match(p.name).group(2).upper()
                 if _STMT_CSV_RE.match(p.name) else ""),
                99,
            )
        )

    created = []

    for base_name, csv_files in sorted(groups.items()):
        print(f"\n  Merging: {base_name}")

        all_columns = OrderedDict()
        all_rows = []

        for csv_file in csv_files:
            stmt_m = _STMT_CSV_RE.match(csv_file.name)
            stmt_type = stmt_m.group(2).upper() if stmt_m else "?"

            with open(csv_file, "r", encoding="utf-8-sig", errors="replace") as fh:
                reader = csv.DictReader(fh)
                fieldnames = list(reader.fieldnames or [])
                rows = list(reader)

            for col in fieldnames:
                if col not in all_columns:
                    all_columns[col] = True

            all_rows.extend(rows)
            print(f"    + {csv_file.name}  ({len(rows)} rows)  [{stmt_type}]")

        if not all_rows:
            print(f"    ⚠ No rows — skipping")
            continue

        final_columns = []
        for col in PRIORITY_COLS:
            if col in all_columns:
                final_columns.append(col)
        for col in all_columns:
            if col not in final_columns:
                final_columns.append(col)

        out_path = output_folder / f"{base_name}_All.csv"
        with open(out_path, "w", newline="", encoding="utf-8-sig") as fh:
            writer = csv.DictWriter(fh, fieldnames=final_columns, extrasaction="ignore")
            writer.writeheader()
            for row in all_rows:
                writer.writerow({col: row.get(col, "") for col in final_columns})

        print(f"    → {out_path.name}  ({len(all_rows)} total rows)")
        created.append(str(out_path))

    print(f"\n  Done. {len(created)} merged file(s) created.")
    return created
# ─────────────────────────────────────────────────────────────────────────────
# MERGE CSVs INSIDE ONE DEAL FOLDER → EXCEL (sheet per statement type)
# THEN DELETE INDIVIDUAL CSVs
# ─────────────────────────────────────────────────────────────────────────────

def merge_deal_csvs_to_excel(deal_folder: str, deal_name: str) -> str:
    """
    Merge all statement CSVs inside a single deal folder into one Excel
    workbook with one sheet per statement type, then delete individual CSVs.

    Output: <deal_folder>/<deal_name>_All_Statements.xlsx

    Parameters
    ----------
    deal_folder : str
        Path to the deal subfolder (e.g. 04_Validated_output/LG_CIT_AK_600023935_2022/)
    deal_name : str
        Base name of the deal (e.g. LG_CIT_AK_600023935_2022)

    Returns
    -------
    str
        Path to the created Excel file, or "" if no CSVs found.
    """
    deal_folder = Path(deal_folder)
    output_path = deal_folder / f"{deal_name}_All_Statements.xlsx"

    # ── Find all statement CSVs in this folder ───────────────────────────
    csv_files_by_type: dict[str, list[Path]] = {}
    all_csv_files: list[Path] = []

    for f in sorted(deal_folder.glob("*.csv")):
        if not f.is_file():
            continue
        if _ALL_CSV_RE.search(f.name):
            continue
        if f.name.startswith("_"):
            continue

        m = _STMT_CSV_RE.match(f.name)
        if not m:
            continue

        stmt_type = m.group(2).upper()
        csv_files_by_type.setdefault(stmt_type, []).append(f)
        all_csv_files.append(f)

    if not csv_files_by_type:
        print(f"    No statement CSVs found in: {deal_folder}")
        return ""

    # ── Create Excel workbook ────────────────────────────────────────────
    wb = openpyxl.Workbook()
    wb.remove(wb.active)  # remove default empty sheet

    ordered_types = sorted(
        csv_files_by_type.keys(),
        key=lambda t: STMT_ORDER.get(t, 99),
    )

    for stmt_type in ordered_types:
        csv_files = sorted(csv_files_by_type[stmt_type])

        all_columns = OrderedDict()
        all_rows = []

        for csv_file in csv_files:
            with open(csv_file, "r", encoding="utf-8-sig", errors="replace") as fh:
                reader = csv.DictReader(fh)
                fieldnames = list(reader.fieldnames or [])
                rows = list(reader)

            for col in fieldnames:
                if col not in all_columns:
                    all_columns[col] = True

            all_rows.extend(rows)

        if not all_rows:
            continue

        # ── Build final column order ─────────────────────────────────
        final_columns = []
        for col in PRIORITY_COLS:
            if col in all_columns:
                final_columns.append(col)
        for col in all_columns:
            if col not in final_columns:
                final_columns.append(col)

        # ── Write to sheet ───────────────────────────────────────────
        ws = wb.create_sheet(title=stmt_type)

        # Header row (bold)
        for col_idx, col_name in enumerate(final_columns, start=1):
            cell = ws.cell(row=1, column=col_idx, value=col_name)
            cell.font = openpyxl.styles.Font(bold=True)

        # Data rows
        for row_idx, row_data in enumerate(all_rows, start=2):
            for col_idx, col_name in enumerate(final_columns, start=1):
                value = row_data.get(col_name, "")
                ws.cell(row=row_idx, column=col_idx, value=value)

        # Auto-width (sample first 50 rows)
        for col_idx, col_name in enumerate(final_columns, start=1):
            max_len = len(col_name)
            for row_data in all_rows[:50]:
                cell_val = str(row_data.get(col_name, ""))
                max_len = max(max_len, len(cell_val))
            ws.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 2, 40)

    # ── Save ─────────────────────────────────────────────────────────────
    if not wb.sheetnames:
        print(f"    No data sheets created — skipping Excel save for {deal_name}")
        return ""

    wb.save(str(output_path))
    print(f"    ✅ {output_path.name}")

    # ── Delete individual CSVs ───────────────────────────────────────────
    deleted = 0
    for csv_file in all_csv_files:
        try:
            csv_file.unlink()
            deleted += 1
        except Exception as e:
            print(f"    [WARN] Could not delete {csv_file.name}: {e}")

    # if deleted:
    #     print(f"    🗑  Deleted {deleted} individual CSV(s)")

    return str(output_path)

# ─────────────────────────────────────────────────────────────────────────────
# MERGE ALL CSVs INTO ONE EXCEL — SHEET PER STATEMENT TYPE
# ─────────────────────────────────────────────────────────────────────────────

def merge_csvs_to_excel(folder: str, output_path: str | None = None) -> str:
    """
    Merge all statement CSVs into one Excel workbook with one sheet
    per statement type (SNP, SOA, GOV_BS, GOV_IS, PROP_SNP, PROP_IS,
    PROP_CFS, DSR, DEBT).

    Parameters
    ----------
    folder : str
        Root folder containing per-PDF subfolders with individual CSV files.
    output_path : str, optional
        Full path for the output .xlsx file.
        Defaults to <folder>/All_Statements.xlsx

    Returns
    -------
    str
        Path to the created Excel file, or "" if no CSVs found.
    """
    folder = Path(folder)

    if output_path is None:
        output_path = folder / "All_Statements.xlsx"
    else:
        output_path = Path(output_path)

    output_path.parent.mkdir(parents=True, exist_ok=True)

    # ── Group CSVs by statement type ─────────────────────────────────────
    stmt_groups: dict[str, list[Path]] = {}

    for f in sorted(folder.rglob("*.csv")):
        if not f.is_file():
            continue
        if _ALL_CSV_RE.search(f.name):
            continue
        # Skip summary / merged files
        if f.name.startswith("_"):
            continue

        m = _STMT_CSV_RE.match(f.name)
        if not m:
            continue

        stmt_type = m.group(2).upper()
        if stmt_type not in stmt_groups:
            stmt_groups[stmt_type] = []
        stmt_groups[stmt_type].append(f)

    if not stmt_groups:
        print(f"  No statement CSVs found in: {folder}")
        return ""

    # ── Create Excel workbook ────────────────────────────────────────────
    wb = openpyxl.Workbook()
    wb.remove(wb.active)  # remove default empty sheet

    # Process sheets in defined order
    ordered_types = sorted(
        stmt_groups.keys(),
        key=lambda t: STMT_ORDER.get(t, 99)
    )

    for stmt_type in ordered_types:
        csv_files = sorted(stmt_groups[stmt_type])

        print(f"\n  Sheet: {stmt_type}  ({len(csv_files)} CSV file(s))")

        # ── Collect all rows + discover all columns ──────────────────
        all_columns = OrderedDict()
        all_rows = []

        for csv_file in csv_files:
            with open(csv_file, "r", encoding="utf-8-sig", errors="replace") as fh:
                reader = csv.DictReader(fh)
                fieldnames = list(reader.fieldnames or [])
                rows = list(reader)

            for col in fieldnames:
                if col not in all_columns:
                    all_columns[col] = True

            all_rows.extend(rows)
            print(f"    + {csv_file.name}  ({len(rows)} rows)")

        if not all_rows:
            print(f"    ⚠ No rows — skipping sheet")
            continue

        # ── Build final column order ─────────────────────────────────
        final_columns = []
        for col in PRIORITY_COLS:
            if col in all_columns:
                final_columns.append(col)
        for col in all_columns:
            if col not in final_columns:
                final_columns.append(col)

        # ── Write to sheet ───────────────────────────────────────────
        ws = wb.create_sheet(title=stmt_type)

        # Header row (bold)
        for col_idx, col_name in enumerate(final_columns, start=1):
            cell = ws.cell(row=1, column=col_idx, value=col_name)
            cell.font = openpyxl.styles.Font(bold=True)

        # Data rows
        for row_idx, row_data in enumerate(all_rows, start=2):
            for col_idx, col_name in enumerate(final_columns, start=1):
                value = row_data.get(col_name, "")
                ws.cell(row=row_idx, column=col_idx, value=value)

        # Auto-width (approximate, sample first 50 rows)
        for col_idx, col_name in enumerate(final_columns, start=1):
            max_len = len(col_name)
            for row_data in all_rows[:50]:
                cell_val = str(row_data.get(col_name, ""))
                max_len = max(max_len, len(cell_val))
            ws.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 2, 40)

        print(f"    → Sheet '{stmt_type}' : {len(all_rows)} rows, {len(final_columns)} columns")

    # ── Save ─────────────────────────────────────────────────────────────
    if not wb.sheetnames:
        print(f"  No data sheets created — skipping Excel save")
        return ""

    wb.save(str(output_path))
    print(f"\n  ✅ Excel saved: {output_path}")
    print(f"     Sheets: {', '.join(ordered_types)}")

    return str(output_path)




if __name__ == "__main__":
    DEAL_FOLDER = r"D:\S2_Khushbu\AI Projects\Financial Data Extraction\MDB_multipleAI\04_Validated_Output\NP_CA_822421813_2025_NON-LG"

    # Manual-validation root = sibling of 04_Validated_Output
    MANUAL_VALIDATION_ROOT = (
        Path(DEAL_FOLDER).parent.parent / "05_Manual_Validation_Required"
    )

    deal_path = Path(DEAL_FOLDER)
    deal_name = deal_path.name

    if not deal_path.exists():
        print(f"[ERROR] Folder not found: {deal_path}")
        exit(1)

    json_files = sorted(deal_path.glob("*.json"))

    if not json_files:
        print(f"[WARN] No JSON files found in: {deal_path}")
        exit(1)

    print(f"{'='*70}")
    print(f"  Deal   : {deal_name}")
    print(f"  Folder : {deal_path}")
    print(f"  JSONs  : {len(json_files)}")
    print(f"{'='*70}\n")

    success, failed = 0, 0
    failed_tables: list[str] = []   # ← tracks which tables failed validation

    for jf in json_files:
        csv_path = deal_path / (jf.stem + ".csv")
        print(f"  [{jf.name}]")
        try:
            result = run_json_to_csv_pipeline(str(jf), str(csv_path))
            status = "PASS" if result else "FAIL(s) detected"
            print(f"    -> {csv_path.name}  ({status})\n")
            success += 1
            if not result:
                failed_tables.append(jf.stem)
        except Exception as e:
            print(f"    [ERROR] {e}\n")
            failed += 1
            failed_tables.append(jf.stem)  # exceptions = also needs manual check

    print(f"\n  JSON → CSV : {success} succeeded, {failed} failed")
    if failed_tables:
        print(f"  Validation FAIL on : {failed_tables}\n")
    else:
        print(f"  Validation : ALL PASS\n")

    print(f"  Merging CSVs into Excel...")
    xlsx_path = merge_deal_csvs_to_excel(str(deal_path), deal_name)

    if xlsx_path:
        print(f"\n  ✅ Excel : {xlsx_path}")
    else:
        print(f"\n  [WARN] No Excel created")

    # ─────────────────────────────────────────────────────────────────
    # MOVE TO 05_Manual_Validation_Required IF ANY FAIL
    # ─────────────────────────────────────────────────────────────────
    if failed_tables:
        MANUAL_VALIDATION_ROOT.mkdir(parents=True, exist_ok=True)
        target_dir = MANUAL_VALIDATION_ROOT / deal_name

        # If target already exists, remove it so move() doesn't fail
        if target_dir.exists():
            shutil.rmtree(target_dir)

        try:
            shutil.move(str(deal_path), str(target_dir))
            print(f"\n  ⚠  FAIL detected in: {failed_tables}")
            print(f"  📦 Deal moved to   : {target_dir}")
        except Exception as e:
            print(f"\n  [ERROR] Could not move deal to manual validation: {e}")
    else:
        print(f"\n  ✅ All tables passed — deal stays in 04_Validated_Output")

    print(f"\n{'='*70}\n  DONE\n{'='*70}")