# -*- coding: utf-8 -*-
"""
parquet_ingest.py — ingest one extracted deal's per-statement JSON into the single
normalized `parquet/RawData.parquet`, resolving the reference tables built by
build_coa_reference.py (CoaDetails / TemplateType) and upserting the runtime
reference tables (Category / UnitMaster / MetaData).

This runs IN ADDITION to the existing JSON + Excel output (which is unchanged); it
does not read or modify the Excel. One deal == one TProcessStatus.ProcessingId.

RawData columns (see parquet-schema/RawData.xlsx):
    DataId              unique per row
    CategoryID          -> Category  (a "Reporting Columns" name; upserted)
    COAHeaderID         -> CoaDetails (matched by COA Datapoint + DPG context)
    MetaDataID          -> MetaData  (FYE value; upserted)
    ProcessingId        = TProcessStatus.ProcessingId (parquet<->DB link)
    Quardinate          null (reserved)
    PageNo              Metadata "Page No"
    PdfDataPpointName   the sheet's raw item-column value (Row Items / SOA Items / ...)
    Value               the actual value of that Reporting Column (raw string)
    UnitId              -> UnitMaster (Currency reported; upserted)
    DataDisplaySequence per-ProcessingId display order; SAME for every Reporting
                        Column (CategoryID) of one source datapoint row
    InsertedOn/By, ModifiedBy/On, IsActive   defaults

Rules:
  * DPG rows are NOT stored (they hold no numeric value; their name is recoverable
    via CoaDetails.ParentID).
  * A datapoint that cannot be matched in CoaDetails (e.g. DEBT/DSR — no COA-master
    coverage yet, or an unknown datapoint) is skipped and reported in the stats.
  * Re-ingesting a ProcessingId REPLACES its existing RawData rows (idempotent).
  * SKIP_EMPTY_VALUES (default True): a Reporting Column whose value is blank / "-"
    produces no RawData row. Set False to store the full grid (blank cells as null).
"""

import json
import re
from pathlib import Path

import polars as pl

# Single shared normalized parquet store for ALL deals (pass or fail).
PARQUET_DIR = Path(r"C:\S2\Public Finance") / "04_Validated_output" / "Parquet"

SKIP_EMPTY_VALUES = True
_EMPTY_TOKENS = {"", "-", "–", "—"}

# Canonical statement/sheet order (matches jsonToCsv.STMT_ORDER) for stable
# DataDisplaySequence across a deal's statements.
_STMT_ORDER = ["SNP", "SOA", "GOV_BS", "GOV_IS", "PROP_SNP", "PROP_IS", "PROP_CFS", "DEBT", "DSR"]
# Longest-first so PROP_SNP matches before SNP, etc.
_STMT_CODES = ["PROP_SNP", "PROP_IS", "PROP_CFS", "GOV_BS", "GOV_IS", "SNP", "SOA", "DSR", "DEBT"]
_STMT_RE = re.compile(r"_(" + "|".join(_STMT_CODES) + r")(?:_Comp)?_p[\d\-]+$", re.IGNORECASE)

RAWDATA_SCHEMA = {
    "DataId": pl.Int64, "CategoryID": pl.Int64, "COAHeaderID": pl.Int64,
    "MetaDataID": pl.Int64, "ProcessingId": pl.Int64, "Quardinate": pl.Utf8,
    "PageNo": pl.Utf8, "PdfDataPpointName": pl.Utf8, "Value": pl.Utf8,
    "UnitId": pl.Int64, "DataDisplaySequence": pl.Int64,
    "InsertedOn": pl.Datetime("us"), "InsertedBy": pl.Utf8, "ModifiedBy": pl.Utf8,
    "ModifiedOn": pl.Datetime("us"), "IsActive": pl.Int64,
}


def detect_stmt_code(json_stem: str):
    m = _STMT_RE.search(json_stem)
    return m.group(1).upper() if m else None


def _is_empty(v) -> bool:
    if v is None:
        return True
    s = str(v).strip()
    return s in _EMPTY_TOKENS or (bool(s) and set(s) == {"#"})


def _norm(s) -> str:
    """Normalize a datapoint/DPG name for tolerant matching: underscores->spaces,
    collapsed whitespace, case-insensitive."""
    if s is None:
        return ""
    return re.sub(r"\s+", " ", str(s).replace("_", " ")).strip().lower()


class ParquetStore:
    """Loads the reference + mutable parquet tables once; ingest_deal() accumulates
    RawData rows in memory; flush() writes RawData + the upserted reference tables."""

    def __init__(self, parquet_dir: Path = PARQUET_DIR):
        self.dir = Path(parquet_dir)
        self._load_refs()
        self._load_mutable()

    # ---- reference tables (read-only) ----
    def _load_refs(self):
        coa = pl.read_parquet(self.dir / "CoaDetails.parquet")
        tt = pl.read_parquet(self.dir / "TemplateType.parquet")
        self.code_to_tid = {r["TemplateCode"]: r["TemplateTypeID"]
                            for r in tt.iter_rows(named=True)}
        self.id_to_display = {r["COAHeaderID"]: r["DisplayName"]
                              for r in coa.iter_rows(named=True)}
        # candidate lookups for datapoint -> COAHeaderID, in template display order.
        # DPG rows (DisplayNameInfoId == 1) are never RawData targets, so exclude them.
        # A datapoint may repeat within a template (e.g. SOA "Charges for Services"
        # under Governmental vs Business-type DPGs); the DPG-context triple keys the
        # right one, the pair is the fallback. MANY JSON rows may legitimately share a
        # single COAHeaderID (many raw line items -> one standardized datapoint), so
        # candidates are NOT consumed — resolution just returns the best match.
        self.by_triple: dict = {}    # (tid, datapoint, parent_dpg)        -> [COAHeaderID]
        self.by_triple_n: dict = {}  # normalized (tid, datapoint, parent) -> [COAHeaderID]
        self.by_pair: dict = {}      # (tid, datapoint)                    -> [COAHeaderID]
        self.by_pair_n: dict = {}    # normalized (tid, datapoint)         -> [COAHeaderID]
        for r in coa.sort(["TemplateTypeId", "COADisplaySequence"]).iter_rows(named=True):
            if r["DisplayNameInfoId"] == 1:
                continue
            tid, dn, cid = r["TemplateTypeId"], r["DisplayName"], r["COAHeaderID"]
            parent_dn = self.id_to_display.get(r["ParentID"])
            self.by_triple.setdefault((tid, dn, parent_dn), []).append(cid)
            self.by_triple_n.setdefault((tid, _norm(dn), _norm(parent_dn)), []).append(cid)
            self.by_pair.setdefault((tid, dn), []).append(cid)
            self.by_pair_n.setdefault((tid, _norm(dn)), []).append(cid)

    # ---- mutable tables (upserted) ----
    def _load_mutable(self):
        cat = pl.read_parquet(self.dir / "Category.parquet")
        unit = pl.read_parquet(self.dir / "UnitMaster.parquet")
        meta = pl.read_parquet(self.dir / "MetaData.parquet")
        raw = pl.read_parquet(self.dir / "RawData.parquet")
        self._cat = {r["CategoryName"]: r["CategoryID"] for r in cat.iter_rows(named=True)}
        self._unit = {r["Name"]: r["UnitID"] for r in unit.iter_rows(named=True)}
        self._meta = {(r["MetadataName"], r["Value"]): r["MetaDataID"]
                      for r in meta.iter_rows(named=True)}
        self._next_cat = (max(self._cat.values()) + 1) if self._cat else 1
        self._next_unit = (max(self._unit.values()) + 1) if self._unit else 1
        self._next_meta = (max(self._meta.values()) + 1) if self._meta else 1
        self.raw_rows = raw.to_dicts()
        ids = [d["DataId"] for d in self.raw_rows if d.get("DataId") is not None]
        self._next_data = (max(ids) + 1) if ids else 1
        # ProcessingId is NON-unique (merge-siblings share it). Clear a ProcessingId's
        # stale rows only on its FIRST ingest this session, so several sibling deals
        # ACCUMULATE within one run, while a full re-run still replaces cleanly.
        self._cleared_pids: set = set()

    def _cat_id(self, name: str) -> int:
        if name not in self._cat:
            self._cat[name] = self._next_cat
            self._next_cat += 1
        return self._cat[name]

    def _unit_id(self, name: str) -> int:
        if name not in self._unit:
            self._unit[name] = self._next_unit
            self._next_unit += 1
        return self._unit[name]

    def _meta_id(self, mname: str, value) -> int:
        key = (mname, value)
        if key not in self._meta:
            self._meta[key] = self._next_meta
            self._next_meta += 1
        return self._meta[key]

    def _resolve(self, tid, datapoint, parent_dpg):
        """Match a JSON datapoint to a CoaDetails COAHeaderID. Tries, in order:
        exact (template, datapoint, parent-DPG) triple, normalized triple, exact
        (template, datapoint) pair, normalized pair. Returns the first candidate
        (COADisplaySequence order); reuse is allowed (many raw rows -> one datapoint)."""
        if tid is None or datapoint is None:
            return None
        ndp, npar = _norm(datapoint), _norm(parent_dpg)
        for bucket, key in (
            (self.by_triple,   (tid, datapoint, parent_dpg)),
            (self.by_triple_n, (tid, ndp, npar)),
            (self.by_pair,     (tid, datapoint)),
            (self.by_pair_n,   (tid, ndp)),
        ):
            cands = bucket.get(key)
            if cands:
                return cands[0]
        return None

    # ---- main entry ----
    def ingest_deal(self, processing_id, json_files) -> dict:
        processing_id = int(processing_id)
        # idempotency: drop this pid's prior rows once per session (see _cleared_pids)
        if processing_id not in self._cleared_pids:
            self.raw_rows = [d for d in self.raw_rows if d.get("ProcessingId") != processing_id]
            self._cleared_pids.add(processing_id)

        order = {c: i for i, c in enumerate(_STMT_ORDER)}
        files = sorted((Path(f) for f in json_files),
                       key=lambda p: (order.get(detect_stmt_code(p.stem) or "", 99), p.name))

        stats = {"rows": 0, "mapped_items": 0, "skipped_items": 0,
                 "files": 0, "unmapped": []}
        disp_seq = 0

        for p in files:
            code = detect_stmt_code(p.stem)
            tid = self.code_to_tid.get(code)
            try:
                d = json.loads(p.read_text(encoding="utf-8", errors="replace"))
            except Exception as e:
                print(f"   [INGEST] unreadable {p.name}: {e}")
                continue
            stats["files"] += 1
            meta = d.get("Metadata", {}) or {}
            rc = d.get("Reporting Columns") or []
            secs = d.get("Sections", {}) or {}
            page_no = meta.get("Page No")
            fye = meta.get("FYE")
            currency = meta.get("Currency reported")
            meta_id = self._meta_id("FYE", fye) if fye not in (None, "") else None
            unit_id = self._unit_id(currency) if currency not in (None, "") else None
            excl = set(rc) | {"COA Flag", "COA Datapoint", "Total Check Status"}

            cur_dpg = None
            for _sname, items in secs.items():
                if not isinstance(items, list):
                    continue
                for it in items:
                    if not isinstance(it, dict):
                        continue
                    flag = str(it.get("COA Flag", "")).strip().upper()
                    datapoint = it.get("COA Datapoint")
                    if flag == "DPG":
                        cur_dpg = datapoint
                        continue                      # DPG rows are not stored
                    coa_id = self._resolve(tid, datapoint, cur_dpg)
                    if coa_id is None:
                        stats["skipped_items"] += 1
                        if datapoint:
                            stats["unmapped"].append(f"{code}:{datapoint}")
                        continue
                    item_cols = [k for k in it.keys() if k not in excl]
                    pdf_name = it.get(item_cols[0]) if item_cols else None

                    # Build this datapoint's Reporting-Column rows first so an
                    # all-blank row does not consume a DataDisplaySequence.
                    pending = []
                    for col in rc:
                        val = it.get(col)
                        if SKIP_EMPTY_VALUES and _is_empty(val):
                            continue
                        pending.append((col, val))
                    if not pending:
                        continue

                    disp_seq += 1
                    stats["mapped_items"] += 1
                    for col, val in pending:
                        self.raw_rows.append({
                            "DataId": self._next_data,
                            "CategoryID": self._cat_id(col),
                            "COAHeaderID": coa_id,
                            "MetaDataID": meta_id,
                            "ProcessingId": processing_id,
                            "Quardinate": None,
                            "PageNo": (str(page_no) if page_no is not None else None),
                            "PdfDataPpointName": (str(pdf_name) if pdf_name is not None else None),
                            "Value": str(val).strip(),
                            "UnitId": unit_id,
                            "DataDisplaySequence": disp_seq,
                            "InsertedOn": None, "InsertedBy": None,
                            "ModifiedBy": None, "ModifiedOn": None, "IsActive": 1,
                        })
                        self._next_data += 1
                        stats["rows"] += 1
        return stats

    # ---- persist ----
    def flush(self):
        raw_cols = {k: [d.get(k) for d in self.raw_rows] for k in RAWDATA_SCHEMA}
        pl.DataFrame(raw_cols, schema=RAWDATA_SCHEMA).write_parquet(self.dir / "RawData.parquet")

        pl.DataFrame(
            {"CategoryID": list(self._cat.values()), "CategoryName": list(self._cat.keys())},
            schema={"CategoryID": pl.Int64, "CategoryName": pl.Utf8},
        ).write_parquet(self.dir / "Category.parquet")

        pl.DataFrame(
            {"UnitID": list(self._unit.values()), "Name": list(self._unit.keys()),
             "IsActive": [1] * len(self._unit)},
            schema={"UnitID": pl.Int64, "Name": pl.Utf8, "IsActive": pl.Int64},
        ).write_parquet(self.dir / "UnitMaster.parquet")

        meta_items = list(self._meta.items())   # ((name, value) -> id)
        pl.DataFrame(
            {"MetaDataID": [i for _k, i in meta_items],
             "MetadataName": [k[0] for k, _i in meta_items],
             "Value": [k[1] for k, _i in meta_items]},
            schema={"MetaDataID": pl.Int64, "MetadataName": pl.Utf8, "Value": pl.Utf8},
        ).write_parquet(self.dir / "MetaData.parquet")


# ---------------------------------------------------------------------------
# Standalone helper: ingest one deal folder for a given ProcessingId (no DB).
#   python parquet_ingest.py "<deal_folder_with_jsons>" <processing_id>
# ---------------------------------------------------------------------------
if __name__ == "__main__":
    import sys
    if len(sys.argv) < 3:
        sys.exit("usage: python parquet_ingest.py <deal_folder> <processing_id>")
    folder, pid = sys.argv[1], int(sys.argv[2])
    jsons = [str(p) for p in sorted(Path(folder).glob("*.json"))]
    print(f"[INGEST] {len(jsons)} json file(s) from {folder} for ProcessingId={pid}")
    store = ParquetStore()
    st = store.ingest_deal(pid, jsons)
    store.flush()
    print(f"[INGEST] done: {st['rows']} RawData rows, {st['mapped_items']} items mapped, "
          f"{st['skipped_items']} items skipped, files={st['files']}")
    if st["unmapped"]:
        from collections import Counter
        print("[INGEST] unmapped datapoints (top 15):",
              dict(Counter(st["unmapped"]).most_common(15)))
