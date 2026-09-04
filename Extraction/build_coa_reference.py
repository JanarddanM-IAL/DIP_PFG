# -*- coding: utf-8 -*-
"""
build_coa_reference.py — generate the static COA reference parquet tables from the
standard COA master.

Reads  : Extraction/Master/standard_coa_master.xlsx
Writes : parquet/CoaDetails.parquet      (exploded per statement type)
         parquet/TemplateType.parquet    (15 statement types — every sheet the engine
                                          emits, incl. the 4 narrative ones which have
                                          no COA datapoints and so no CoaDetails rows)
         parquet/DataType.parquet         (unchanged reference: int/bool/float)
         parquet/DisplayNameInfo.parquet  (unchanged reference: DPG/DP/CP)
         parquet/{RawData,Category,UnitMaster,MetaData}.parquet  (empty scaffolds
                 with the correct columns — populated by the later ingestion phase)

CoaDetails is a PRE-CONFIGURED template (no runtime inserts). Each COA-master
datapoint is exploded across every statement type in its pipe-delimited `Statement`
cell, so e.g. "Cash and Investments" (SNP|PROP_SNP|GOV_BS) becomes three rows, one
per TemplateTypeId. ParentID uses the running nearest-preceding-DPG within each
template (DPG -> -1; DP/CP -> the COAHeaderID of the most recent DPG above it).

Idempotent: re-running rebuilds every file from the master. Uses polars (matches
Extraction/log_writer.py).
"""

import sys
from pathlib import Path

import openpyxl
import polars as pl

from parquet_ingest import PARQUET_DIR, RAWDATA_SCHEMA   # single source of truth for
                                                         # the parquet location + schema

ENGINE_DIR = Path(__file__).resolve().parent          # the Extraction/ folder (this file lives here)
MASTER     = ENGINE_DIR / "Master" / "standard_coa_master.xlsx"
OUT_DIR    = PARQUET_DIR

# Defaults per spec
COAID_DEFAULT     = 1
DATATYPE_DEFAULT  = 3            # float
FLAG_TO_DNI       = {"DPG": 1, "DP": 2, "CP": 3}   # -> DisplayNameInfoId

# Every statement/sheet type the extraction engine emits, in canonical output-sheet
# order (jsonToCsv.STMT_ORDER + 1). (code, friendly name) — TemplateTypeID is the
# 1-based position in this list.
#
# APPEND-ONLY: ids 1-9 are already referenced by ingested RawData / CoaDetails rows,
# so new codes go at the END and never in the middle. A code must be listed here for
# its rows to be storable at all, because an UNMAPPED RawData row records
# TemplateTypeId in place of COAHeaderID (see parquet_ingest module docstring) — that
# is why the 4 narrative sheets are included even though they carry no COA datapoints
# and therefore contribute 0 CoaDetails rows.
TEMPLATE_TYPES = [
    ("SNP",            "Statement of Net Position"),
    ("SOA",            "Statement of Activities"),
    ("GOV_BS",         "Balance Sheet - Governmental Funds"),
    ("GOV_IS",         "Statement of Revenues, Expenditures and Changes in Fund Balances - Governmental Funds"),
    ("PROP_SNP",       "Statement of Fund Net Position - Proprietary Funds"),
    ("PROP_IS",        "Statement of Revenues, Expenses and Changes in Fund Net Position - Proprietary Funds"),
    ("PROP_CFS",       "Statement of Cash Flows - Proprietary Funds"),
    ("DEBT",           "Long-term & Short-term Debt Schedule"),
    ("DSR",            "Debt Service Requirements"),
    # --- appended 2026-09-04 (ids 10-15); do not reorder the entries above ---
    ("CAPITAL_ASSETS", "Capital Assets Schedule"),
    ("TAX_BASE",       "Tax Base Schedule"),
    ("OVERVIEW",       "Issuer Overview"),
    ("PEN",            "Pension Disclosures"),
    ("OPEB",           "OPEB Disclosures"),
    ("FAQS",           "Frequently Asked Questions"),
]
CODE_TO_TID = {code: i for i, (code, _name) in enumerate(TEMPLATE_TYPES, start=1)}

# Unchanged reference tables (no changes per spec — carried into parquet as-is).
DATATYPE_ROWS         = [(1, "int", 1), (2, "bool", 1), (3, "float", 1)]
DISPLAY_NAME_INFO_ROWS = [(1, "DPG", 1), (2, "DP", 1), (3, "CP", 1)]


def read_master_rows():
    """Return the master data rows as tuples (sector, statement, section, flag, datapoint)."""
    wb = openpyxl.load_workbook(MASTER, read_only=True, data_only=True)
    ws = wb.worksheets[0]
    out = []
    for i, row in enumerate(ws.iter_rows(values_only=True)):
        if i == 0:
            continue  # header
        if row is None or all(c is None or str(c).strip() == "" for c in row):
            continue
        sector, statement, section, flag, datapoint = (
            (str(c).strip() if c is not None else "") for c in row[:5]
        )
        out.append((sector, statement, section, flag.upper(), datapoint))
    wb.close()
    return out


def build_coa_details(master_rows):
    """Explode master rows per statement token and assign ids/parents/sequences."""
    # Per-token ordered list of (section, flag, datapoint), preserving master order.
    per_template: dict[str, list[tuple]] = {code: [] for code, _ in TEMPLATE_TYPES}
    for _sector, statement, section, flag, datapoint in master_rows:
        for tok in (t.strip() for t in statement.split("|")):
            if tok in per_template:            # DEBT/DSR never appear -> stay empty
                per_template[tok].append((section, flag, datapoint))
            elif tok:
                print(f"[WARN] master statement token '{tok}' is not a known template type "
                      f"— row '{datapoint}' skipped for that token.")

    cols = {k: [] for k in (
        "COAHeaderID", "COAID", "DisplayName", "ParentID", "COADisplaySequence",
        "DataTypeID", "TemplateTypeId", "DisplayNameInfoId",
    )}
    next_id = 1
    for code, _name in TEMPLATE_TYPES:                 # iterate in TemplateTypeId order
        tid = CODE_TO_TID[code]
        seq = 0
        cur_dpg = -1
        for section, flag, datapoint in per_template[code]:
            cid = next_id
            next_id += 1
            seq += 1
            if flag == "DPG":
                parent = -1
                cur_dpg = cid
            else:                                       # DP / CP
                parent = cur_dpg
            cols["COAHeaderID"].append(cid)
            cols["COAID"].append(COAID_DEFAULT)
            cols["DisplayName"].append(datapoint)
            cols["ParentID"].append(parent)
            cols["COADisplaySequence"].append(seq)
            cols["DataTypeID"].append(DATATYPE_DEFAULT)
            cols["TemplateTypeId"].append(tid)
            cols["DisplayNameInfoId"].append(FLAG_TO_DNI.get(flag))
    return cols


def _with_audit(df: pl.DataFrame, active: int = 1) -> pl.DataFrame:
    """Append the 5 standard audit/default columns (timestamps null, IsActive set)."""
    n = df.height
    return df.with_columns(
        pl.Series("InsertedOn", [None] * n, dtype=pl.Datetime("us")),
        pl.Series("InsertedBy", [None] * n, dtype=pl.Utf8),
        pl.Series("ModifiedBy", [None] * n, dtype=pl.Utf8),
        pl.Series("ModifiedOn", [None] * n, dtype=pl.Datetime("us")),
        pl.Series("IsActive",   [active] * n, dtype=pl.Int64),
    )


def main() -> int:
    if not MASTER.is_file():
        sys.exit(f"[ERROR] COA master not found: {MASTER}")
    OUT_DIR.mkdir(parents=True, exist_ok=True)

    master_rows = read_master_rows()
    print(f"[INFO] Master rows read: {len(master_rows)}")

    # ---- TemplateType ----
    template_type = pl.DataFrame(
        {
            "TemplateTypeID": [CODE_TO_TID[c] for c, _ in TEMPLATE_TYPES],
            "TemplateName":   [name for _c, name in TEMPLATE_TYPES],
            "TemplateCode":   [c for c, _ in TEMPLATE_TYPES],
            "IsActive":       [1] * len(TEMPLATE_TYPES),
        },
        schema={"TemplateTypeID": pl.Int64, "TemplateName": pl.Utf8,
                "TemplateCode": pl.Utf8, "IsActive": pl.Int64},
    )

    # ---- CoaDetails ----
    c = build_coa_details(master_rows)
    coa_details = _with_audit(pl.DataFrame(
        c,
        schema={"COAHeaderID": pl.Int64, "COAID": pl.Int64, "DisplayName": pl.Utf8,
                "ParentID": pl.Int64, "COADisplaySequence": pl.Int64,
                "DataTypeID": pl.Int64, "TemplateTypeId": pl.Int64,
                "DisplayNameInfoId": pl.Int64},
    ))

    # ---- Unchanged reference tables ----
    data_type = pl.DataFrame(
        {"DataTypeID": [r[0] for r in DATATYPE_ROWS],
         "DataType":   [r[1] for r in DATATYPE_ROWS],
         "IsActive":   [r[2] for r in DATATYPE_ROWS]},
        schema={"DataTypeID": pl.Int64, "DataType": pl.Utf8, "IsActive": pl.Int64},
    )
    display_name_info = pl.DataFrame(
        {"DisplayNameInfoId": [r[0] for r in DISPLAY_NAME_INFO_ROWS],
         "Display_Name":      [r[1] for r in DISPLAY_NAME_INFO_ROWS],
         "Is_Active":         [r[2] for r in DISPLAY_NAME_INFO_ROWS]},
        schema={"DisplayNameInfoId": pl.Int64, "Display_Name": pl.Utf8, "Is_Active": pl.Int64},
    )

    # ---- Empty scaffolds (populated by the later ingestion phase) ----
    # Schema is owned by parquet_ingest.RAWDATA_SCHEMA — import it rather than
    # restating the columns, so the scaffold can never drift from what ingestion writes.
    raw_data = pl.DataFrame(schema=RAWDATA_SCHEMA)
    category = pl.DataFrame(schema={"CategoryID": pl.Int64, "CategoryName": pl.Utf8})
    unit_master = pl.DataFrame(schema={"UnitID": pl.Int64, "Name": pl.Utf8, "IsActive": pl.Int64})
    meta_data = pl.DataFrame(schema={"MetaDataID": pl.Int64, "MetadataName": pl.Utf8, "Value": pl.Utf8})

    # ---- Write ----
    # Static reference tables: always (re)written from the master.
    always = {
        "CoaDetails": coa_details,
        "TemplateType": template_type,
        "DataType": data_type,
        "DisplayNameInfo": display_name_info,
    }
    # Runtime tables: only scaffolded when ABSENT, so rebuilding the reference
    # never wipes data already ingested by parquet_ingest.py.
    scaffold_if_absent = {
        "RawData": raw_data,
        "Category": category,
        "UnitMaster": unit_master,
        "MetaData": meta_data,
    }
    for name, df in always.items():
        path = OUT_DIR / f"{name}.parquet"
        df.write_parquet(path)
        print(f"[WRITE]    {path.name:<24} rows={df.height:<4} cols={df.width}")
    for name, df in scaffold_if_absent.items():
        path = OUT_DIR / f"{name}.parquet"
        if path.exists():
            print(f"[KEEP]     {path.name:<24} (exists — preserving ingested data)")
            continue
        df.write_parquet(path)
        print(f"[SCAFFOLD] {path.name:<24} rows={df.height:<4} cols={df.width}")

    # ---- Summary ----
    print("\n[SUMMARY] CoaDetails per TemplateTypeId:")
    summary = (coa_details.group_by("TemplateTypeId")
               .len().sort("TemplateTypeId"))
    tid_to_code = {v: k for k, v in CODE_TO_TID.items()}
    for row in summary.iter_rows(named=True):
        print(f"   {tid_to_code[row['TemplateTypeId']]:<10} "
              f"(TemplateTypeId={row['TemplateTypeId']}) -> {row['len']} rows")
    print(f"   TOTAL CoaDetails rows: {coa_details.height}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
