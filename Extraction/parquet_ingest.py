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
    Quardinate          compact JSON of this cell's bounding box, taken from the item's
                        "<Reporting Column>_coord" sibling key — so the box belongs to
                        THIS row's CategoryID (e.g. an SNP row whose CategoryID is
                        "Governmental Activities" gets "Governmental Activities_coord").
                        Null when the engine emitted no box for that cell.
    ReportedInThousand  Metadata "Reported in Thousand" ("Yes"/"No") — the scale
                        indicator, per STATEMENT, so every row from one JSON file
                        shares it. Null for a pre-Update-2 JSON that omits the key.
    PageNo              Metadata "Page No"
    PdfDataPpointName   the sheet's raw item-column value (Row Items / SOA Items / ...)
    TemplateTypeId      -> TemplateType, but ONLY for rows with no COA mapping
    GroupName           the JSON section key, ONLY for rows with no COA mapping
    Value               the actual value of that Reporting Column (raw string)
    UnitId              -> UnitMaster (Currency reported; upserted)
    DataDisplaySequence per-ProcessingId display order; SAME for every Reporting
                        Column (CategoryID) of one source datapoint row
    InsertedOn/By, ModifiedBy/On, IsActive   defaults

MAPPED vs UNMAPPED rows (the COAHeaderID / TemplateTypeId+GroupName split):
  * MAPPED   — the datapoint resolved to a CoaDetails row:
               COAHeaderID = that row, TemplateTypeId = null, GroupName = null.
               The statement and group are recoverable by walking
               CoaDetails.TemplateTypeId / .ParentID.
  * UNMAPPED — the datapoint has no (or no unambiguous) CoaDetails row:
               COAHeaderID = null, TemplateTypeId = the sheet's TemplateType id,
               GroupName = the JSON section key. The value is still stored and
               still attributed to the right statement and group; only the COA
               standardization is missing.

Rules:
  * DPG rows are NOT stored (they hold no numeric value; their name is recoverable
    via CoaDetails.ParentID).
  * Re-ingesting a ProcessingId REPLACES its existing RawData rows (idempotent).
  * SKIP_EMPTY_VALUES (default True): a Reporting Column whose value is blank / "-"
    produces no RawData row. Set False to store the full grid (blank cells as null).
  * A row is dropped outright ONLY when the filename matches no known statement code
    (so not even TemplateTypeId could be recorded) or its COA Datapoint is blank.
"""

import json
import os
import re
import threading
from pathlib import Path

import polars as pl
from filelock import FileLock

# Single shared normalized parquet store for ALL deals (pass or fail).
# NOTE: this is the STANDALONE-run default only. The DB/Dagster workflow injects
# the real location via ParquetStore(parquet_dir=PFG_Extraction.PARQUET_DIR), which
# is derived from the single control point PFG_Extraction._DATA_ROOT.
PARQUET_DIR = Path(r"C:\S2\Public Finance") / "04_Validated_output" / "Parquet"

# ── Concurrency (parallel Dagster jobs write the SAME parquet files) ─────────
# Mirror PFG_Sourcing/PFG_Validation: a thread lock (in-process threads) + a
# cross-process FileLock, both held around the whole read-modify-write. The
# FileLock is created per-instance from the store dir (see ParquetStore.__init__)
# so it follows the injected path and every process on that dir shares one lock.
_STORE_THREAD_LOCK = threading.Lock()


def _atomic_write_parquet(df: "pl.DataFrame", path: Path) -> None:
    """Write to a sibling .tmp then os.replace() so a concurrent reader never sees
    a half-written file (os.replace is atomic on Windows within one directory)."""
    tmp = path.with_suffix(".parquet.tmp")
    df.write_parquet(tmp)
    os.replace(tmp, path)

SKIP_EMPTY_VALUES = True
_EMPTY_TOKENS = {"", "-", "–", "—"}

# _resolve() outcomes. MAPPED -> COAHeaderID is set; the other two mean the row is
# stored UNMAPPED (COAHeaderID null + TemplateTypeId + GroupName).
#   AMBIGUOUS = the name exists in CoaDetails but under several groups and no group
#               context picked one -> deliberately not guessed.
#   UNMAPPED  = the name isn't in CoaDetails for this template at all.
R_MAPPED    = "mapped"
R_AMBIGUOUS = "ambiguous"
R_UNMAPPED  = "unmapped"

# Canonical statement/sheet order (matches jsonToCsv.STMT_ORDER) for stable
# DataDisplaySequence across a deal's statements. All 15 sheets the engine emits are
# listed: the 7 COA-mapped core statements, DEBT/DSR/CAPITAL_ASSETS/TAX_BASE, and the
# 4 narrative sheets. A code MUST appear here (and in TemplateType) for its rows to be
# storable — an unmapped row records TemplateTypeId, so an unknown code stores nothing.
_STMT_ORDER = ["SNP", "SOA", "GOV_BS", "GOV_IS", "PROP_SNP", "PROP_IS", "PROP_CFS",
               "DEBT", "DSR", "CAPITAL_ASSETS", "TAX_BASE",
               "OVERVIEW", "PEN", "OPEB", "FAQS"]
# Longest-first so PROP_SNP matches before SNP, etc. This ordering is LOAD-BEARING:
# the codes become a regex alternation, and `_SNP_p54-55` is a suffix of
# `_PROP_SNP_p54-55`, so a shorter code placed first would win and mis-detect the sheet.
_STMT_CODES = ["CAPITAL_ASSETS", "PROP_SNP", "PROP_CFS", "PROP_IS", "GOV_BS", "GOV_IS",
               "TAX_BASE", "OVERVIEW", "SNP", "SOA", "DSR", "DEBT", "OPEB", "PEN", "FAQS"]
_STMT_RE = re.compile(r"_(" + "|".join(_STMT_CODES) + r")(?:_Comp)?_p[\d\-]+$", re.IGNORECASE)

# Column order matches parquet-schema/RawData.xlsx exactly: ReportedInThousand sits
# between Quardinate and PageNo; TemplateTypeId + GroupName between PdfDataPpointName
# and Value. This dict is the SOLE source of the written columns — flush() builds the
# frame by iterating it — so a column absent here is not written blank, it is not
# written at all (and is stripped from any hand-placed file on the next flush).
RAWDATA_SCHEMA = {
    "DataId": pl.Int64, "CategoryID": pl.Int64, "COAHeaderID": pl.Int64,
    "MetaDataID": pl.Int64, "ProcessingId": pl.Int64, "Quardinate": pl.Utf8,
    "ReportedInThousand": pl.Utf8,
    "PageNo": pl.Utf8, "PdfDataPpointName": pl.Utf8,
    "TemplateTypeId": pl.Int64, "GroupName": pl.Utf8,
    "Value": pl.Utf8,
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


# Suffix the engine appends to a Reporting Column name to carry that cell's bounding
# box, e.g. "Governmental Activities" -> "Governmental Activities_coord".
COORD_SUFFIX = "_coord"


def _coord_str(v):
    """Serialize one `<column>_coord` payload for the RawData.Quardinate column.

    The engine emits a dict — {"page": 39, "x0": .., "y0": .., "x1": .., "y1": ..} —
    which is stored as compact JSON so it round-trips exactly (Quardinate is Utf8).
    Anything unexpected is kept as its string form rather than dropped; a missing or
    empty coord yields None so the column stays null for cells the engine could not
    locate (it does not emit a box for every value).
    """
    if v is None or v == "" or v == {} or v == []:
        return None
    if isinstance(v, (dict, list)):
        try:
            return json.dumps(v, separators=(",", ":"), ensure_ascii=False)
        except Exception:
            return str(v)
    return str(v)


class ParquetStore:
    """Loads the read-only reference tables (CoaDetails / TemplateType) once;
    ingest_deal() accumulates NAME-KEYED pending records in memory; flush() does the
    whole read-modify-write of the 4 mutable tables (RawData / Category / UnitMaster /
    MetaData) atomically under a cross-process lock — so parallel jobs writing the same
    store never lose rows or collide on surrogate keys.

    Surrogate-key assignment (DataId / CategoryID / UnitId / MetaDataID) is DEFERRED to
    flush and done against a FRESH on-disk read inside the lock; that is what keeps
    parallel processes from minting the same ids off a stale snapshot."""

    def __init__(self, parquet_dir: Path = PARQUET_DIR):
        self.dir = Path(parquet_dir)
        self._load_refs()
        # Cross-process lock for this store dir (all processes on the same dir share
        # one lock file); paired with the module-level in-process thread lock.
        self._file_lock = FileLock(str(self.dir / "_store.parquet.lock"), timeout=120)
        # Name-keyed pending rows (final ids resolved at flush).
        self._pending: list = []
        # Every ProcessingId handed to ingest_deal this run (even if it mapped 0 rows),
        # so a re-run replaces the pid's old rows even when the new extraction is empty.
        self._ingested_pids: set = set()
        # ProcessingId is NON-unique (merge-siblings share it). A pid's stale on-disk
        # rows are dropped only on its FIRST flush this session (see flush), so sibling
        # deals ACCUMULATE within a run while a full re-run still replaces cleanly.
        self._cleared_pids: set = set()

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

    @staticmethod
    def _read_table(path: Path) -> "pl.DataFrame | None":
        """Read a mutable table, tolerating a missing/unreadable file (-> None)."""
        try:
            if path.exists():
                return pl.read_parquet(path)
        except Exception:
            pass
        return None

    def _resolve(self, tid, datapoint, parent_dpg, section=None):
        """Match a JSON datapoint to a CoaDetails COAHeaderID.

        Returns ``(coa_id, status)`` with status one of R_MAPPED / R_AMBIGUOUS /
        R_UNMAPPED. Tries, in order:

          1. the DPG-context triple (exact, then normalized);
          2. the SECTION-key triple (exact, then normalized) — the JSON section name
             is the group whenever a sheet emits no DPG row (CAPITAL_ASSETS) or emits
             one that isn't a CoaDetails group (DEBT's 'Housing Finance Authority'
             inside section 'Component Units'). Section keys normalize onto the
             master's group names, so this recovers the right group rather than
             guessing;
          3. the bare (template, datapoint) pair — ONLY when it is unique.

        A pair with several candidates is NEVER guessed: it returns
        (None, R_AMBIGUOUS) so ingest_deal stores the row unmapped (COAHeaderID blank
        + TemplateTypeId + GroupName). Silently taking the first candidate is what
        filed every CAPITAL_ASSETS row under 'Governmental Activities'.

        Reuse is still allowed on a hit: many raw PDF rows legitimately share one
        standardized datapoint, so candidates are not consumed.
        """
        if tid is None or datapoint is None or not str(datapoint).strip():
            return None, R_UNMAPPED
        ndp = _norm(datapoint)

        # 1 + 2: group context, DPG first then the section key.
        for group in (parent_dpg, section):
            if group is None or not str(group).strip():
                continue
            for bucket, key in (
                (self.by_triple,   (tid, datapoint, group)),
                (self.by_triple_n, (tid, ndp, _norm(group))),
            ):
                cands = bucket.get(key)
                if cands:
                    return cands[0], R_MAPPED

        # 3: name-only, accepted only when unambiguous.
        ambiguous = False
        for bucket, key in (
            (self.by_pair,   (tid, datapoint)),
            (self.by_pair_n, (tid, ndp)),
        ):
            cands = bucket.get(key)
            if cands:
                if len(cands) == 1:
                    return cands[0], R_MAPPED
                ambiguous = True
        return None, (R_AMBIGUOUS if ambiguous else R_UNMAPPED)

    # ---- main entry ----
    def ingest_deal(self, processing_id, json_files) -> dict:
        """Parse one deal's JSON into NAME-KEYED pending records (category/unit/meta by
        name, COAHeaderID from the read-only reference). Surrogate ids are minted later
        in flush(), under the lock, against the fresh on-disk tables. No disk writes here."""
        processing_id = int(processing_id)
        self._ingested_pids.add(processing_id)

        order = {c: i for i, c in enumerate(_STMT_ORDER)}
        files = sorted((Path(f) for f in json_files),
                       key=lambda p: (order.get(detect_stmt_code(p.stem) or "", 99), p.name))

        stats = {"rows": 0, "mapped_items": 0, "unmapped_items": 0,
                 "ambiguous_items": 0, "skipped_items": 0,
                 "files": 0, "unmapped": [], "ambiguous": []}
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
            fye_val = fye if fye not in (None, "") else None
            unit_name = currency if currency not in (None, "") else None
            # Scale indicator, per STATEMENT — stamped on every row from this file,
            # like PageNo/FYE/Currency. Left NULL when the key is absent rather than
            # defaulting to "No", so a pre-Update-2 JSON is recorded as "unknown"
            # instead of being asserted to be in units.
            rit = meta.get("Reported in Thousand")
            rit = str(rit).strip() if rit not in (None, "") else None
            excl = set(rc) | {"COA Flag", "COA Datapoint", "Total Check Status"}

            cur_dpg = None
            for sname, items in secs.items():
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
                    coa_id, status = self._resolve(tid, datapoint, cur_dpg, sname)

                    # A row with no COA mapping is still STORED (COAHeaderID blank,
                    # TemplateTypeId + GroupName instead). It is dropped ONLY when
                    # there is no template to record it against — an unrecognised
                    # filename — since then nothing about it would be recoverable.
                    if coa_id is None and tid is None:
                        stats["skipped_items"] += 1
                        if datapoint:
                            stats["unmapped"].append(f"{code}:{datapoint}")
                        continue
                    item_cols = [k for k in it.keys()
                                 if k not in excl and not str(k).endswith("_coord")]
                    pdf_name = it.get(item_cols[0]) if item_cols else None

                    # Build this datapoint's Reporting-Column rows first so an
                    # all-blank row does not consume a DataDisplaySequence. Each
                    # column also carries its own bounding box under
                    # "<column>_coord" -> that row's Quardinate.
                    cols = []
                    for col in rc:
                        val = it.get(col)
                        if SKIP_EMPTY_VALUES and _is_empty(val):
                            continue
                        cols.append((col, val, _coord_str(it.get(col + COORD_SUFFIX))))
                    if not cols:
                        continue

                    disp_seq += 1
                    if coa_id is None:
                        stats["unmapped_items" if status == R_UNMAPPED
                              else "ambiguous_items"] += 1
                        bucket = ("unmapped" if status == R_UNMAPPED else "ambiguous")
                        stats[bucket].append(f"{code}:{datapoint}")
                    else:
                        stats["mapped_items"] += 1
                    for col, val, coord in cols:
                        # NAME-keyed pending record; ids resolved in flush().
                        # COAHeaderID and TemplateTypeId+GroupName are mutually
                        # exclusive — see the module docstring.
                        self._pending.append({
                            "processing_id": processing_id,
                            "category_name": col,
                            "coa_id": coa_id,
                            "quardinate": coord,
                            "reported_in_thousand": rit,
                            "template_type_id": (tid if coa_id is None else None),
                            "group_name": (str(sname) if coa_id is None else None),
                            "fye_value": fye_val,
                            "unit_name": unit_name,
                            "page_no": (str(page_no) if page_no is not None else None),
                            "pdf_name": (str(pdf_name) if pdf_name is not None else None),
                            "value": str(val).strip(),
                            "disp_seq": disp_seq,
                        })
                        stats["rows"] += 1
        return stats

    # ---- persist (concurrency-safe: whole read-modify-write under one lock) ----
    def flush(self):
        """Commit this run's pending rows. The FULL read-modify-write of all 4 mutable
        tables runs inside a thread lock + cross-process FileLock, and every surrogate id
        is minted from a FRESH on-disk read taken INSIDE the lock — so N parallel jobs
        writing the same store never lose each other's rows nor collide on ids.
        Writes are atomic (temp + os.replace) so readers never see a partial file."""
        to_clear = self._ingested_pids - self._cleared_pids
        if not self._pending and not to_clear:
            return

        with _STORE_THREAD_LOCK:                 # serialize sibling threads (fast)
            with self._file_lock:                # serialize other processes (slow)
                # 1. FRESH read of the mutable tables (sees peers' committed rows).
                cat_df = self._read_table(self.dir / "Category.parquet")
                unit_df = self._read_table(self.dir / "UnitMaster.parquet")
                meta_df = self._read_table(self.dir / "MetaData.parquet")
                raw_df = self._read_table(self.dir / "RawData.parquet")

                cat = ({r["CategoryName"]: r["CategoryID"] for r in cat_df.iter_rows(named=True)}
                       if cat_df is not None else {})
                unit = ({r["Name"]: r["UnitID"] for r in unit_df.iter_rows(named=True)}
                        if unit_df is not None else {})
                metam = ({(r["MetadataName"], r["Value"]): r["MetaDataID"]
                          for r in meta_df.iter_rows(named=True)} if meta_df is not None else {})
                raw_rows = raw_df.to_dicts() if raw_df is not None else []

                next_cat = (max(cat.values()) + 1) if cat else 1
                next_unit = (max(unit.values()) + 1) if unit else 1
                next_meta = (max(metam.values()) + 1) if metam else 1
                ids = [d["DataId"] for d in raw_rows if d.get("DataId") is not None]
                next_data = (max(ids) + 1) if ids else 1

                # get-or-create against the FRESH maps (a name a peer just committed is
                # reused, not duplicated; a genuinely new name gets the next free id).
                def cat_id(name):
                    nonlocal next_cat
                    if name not in cat:
                        cat[name] = next_cat; next_cat += 1
                    return cat[name]

                def unit_id(name):
                    nonlocal next_unit
                    if name is None:
                        return None
                    if name not in unit:
                        unit[name] = next_unit; next_unit += 1
                    return unit[name]

                def meta_id(value):
                    nonlocal next_meta
                    if value is None:
                        return None
                    key = ("FYE", value)
                    if key not in metam:
                        metam[key] = next_meta; next_meta += 1
                    return metam[key]

                # 2. idempotent replace: drop each owned pid's stale rows ONCE (first
                #    flush of that pid). Under the lock this only removes THIS job's pid.
                if to_clear:
                    raw_rows = [d for d in raw_rows if d.get("ProcessingId") not in to_clear]
                    self._cleared_pids |= to_clear

                # 3. append this run's pending rows with freshly-minted ids.
                for rec in self._pending:
                    raw_rows.append({
                        "DataId": next_data,
                        "CategoryID": cat_id(rec["category_name"]),
                        "COAHeaderID": rec["coa_id"],
                        "MetaDataID": meta_id(rec["fye_value"]),
                        "ProcessingId": rec["processing_id"],
                        "Quardinate": rec["quardinate"],
                        "ReportedInThousand": rec["reported_in_thousand"],
                        "PageNo": rec["page_no"],
                        "PdfDataPpointName": rec["pdf_name"],
                        "TemplateTypeId": rec["template_type_id"],
                        "GroupName": rec["group_name"],
                        "Value": rec["value"],
                        "UnitId": unit_id(rec["unit_name"]),
                        "DataDisplaySequence": rec["disp_seq"],
                        "InsertedOn": None, "InsertedBy": None,
                        "ModifiedBy": None, "ModifiedOn": None, "IsActive": 1,
                    })
                    next_data += 1

                # 4. atomic write of all four tables.
                raw_cols = {k: [d.get(k) for d in raw_rows] for k in RAWDATA_SCHEMA}
                _atomic_write_parquet(
                    pl.DataFrame(raw_cols, schema=RAWDATA_SCHEMA), self.dir / "RawData.parquet")
                _atomic_write_parquet(
                    pl.DataFrame(
                        {"CategoryID": list(cat.values()), "CategoryName": list(cat.keys())},
                        schema={"CategoryID": pl.Int64, "CategoryName": pl.Utf8}),
                    self.dir / "Category.parquet")
                _atomic_write_parquet(
                    pl.DataFrame(
                        {"UnitID": list(unit.values()), "Name": list(unit.keys()),
                         "IsActive": [1] * len(unit)},
                        schema={"UnitID": pl.Int64, "Name": pl.Utf8, "IsActive": pl.Int64}),
                    self.dir / "UnitMaster.parquet")
                meta_items = list(metam.items())   # ((name, value) -> id)
                _atomic_write_parquet(
                    pl.DataFrame(
                        {"MetaDataID": [i for _k, i in meta_items],
                         "MetadataName": [k[0] for k, _i in meta_items],
                         "Value": [k[1] for k, _i in meta_items]},
                        schema={"MetaDataID": pl.Int64, "MetadataName": pl.Utf8, "Value": pl.Utf8}),
                    self.dir / "MetaData.parquet")

                # 5. pending rows are now on disk; keep _cleared_pids so a later flush of
                #    the same pid appends (accumulates) instead of re-dropping.
                self._pending = []


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
    print(f"[INGEST] done: {st['rows']} RawData rows, files={st['files']} | "
          f"mapped={st['mapped_items']} unmapped={st['unmapped_items']} "
          f"ambiguous={st['ambiguous_items']} dropped={st['skipped_items']}")
    from collections import Counter
    for label in ("unmapped", "ambiguous"):
        if st[label]:
            print(f"[INGEST] {label} datapoints (top 15):",
                  dict(Counter(st[label]).most_common(15)))
