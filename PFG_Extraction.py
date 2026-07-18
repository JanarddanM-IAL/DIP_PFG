# -*- coding: utf-8 -*-
"""
PFG_Extraction.py — DB-integrated financial-statement extraction runner.

Third stage of the pipeline, after PFG_Sourcing.py (sourcing) and
PFG_Validation.py (sourcing validation). Polls TProcessStatus for rows whose
validation is complete, opens each already-validated PDF (path stored in
TProcessStatus.PdfFilePath, left in 03_Validated_Report by validation), and runs
the *new* extraction engine that lives in the Extraction/ package
(Extraction/pipeline.py): page slicing (optionally LLM-guided), per-statement LLM
normalization with COA-mapping + compact-schema, multi-table split, JSON→CSV
Total Check, and a merged per-deal .xlsx.

Unlike the old DB extraction pipeline (Extraction_db_version old/pipeline.py) — which
wrote the JSON to a temp dir and then deleted it — this runner PERSISTS the
per-statement JSON files in the deal's output folder and stores that folder path in
TProcessStatus.OutputPath, so a future parquet task can consume the JSON.

Run modes (mirrors PFG_Sourcing.py / PFG_Validation.py):
    python PFG_Extraction.py           # BATCH  — extract every pending row
    python PFG_Extraction.py <id>      # SINGLE — extract one TProcessStatus.Id

Status lifecycle written back to TProcessStatus (keyed on the physical Id):
    ExtractionFlag       NULL -> 's' (claimed) -> 'p' (in progress) -> 'c' (done)
    DataValidationFlag   'p' (in progress) -> 'c' (done)
    ExtractionStatus     1 = JSON produced, 0 = nothing produced / error
    DataValidationStatus 1 = every table + Total Check passed, else 0
    CompletionStatus     1 iff ExtractionStatus == 1 AND DataValidationStatus == 1
    OutputPath           the deal output folder (JSON + CSV + merged xlsx)
    Remarks              failure reason (NULL on success)

Sector rule: TProcessStatus.COAID -> TCOAMaster.SegmentId; SegmentId 1 = LG,
SegmentId 2 = NON-LG (default LG when unknown) — same rule PFG_Validation.py uses.
"""

import sys
import shutil
import tempfile
from pathlib import Path

# ---------- make the Extraction/ engine importable ----------
# PFG_Extraction.py lives at the project root; the extraction engine and its
# helper modules (page_extractor, jsonToCsv, compact_schema, ...) live in
# Extraction/. Put that folder first on sys.path so `import pipeline` resolves to
# the engine while `import db` still resolves to the project-root db.py.
_ENGINE_DIR = Path(__file__).resolve().parent / "Extraction"
sys.path.insert(0, str(_ENGINE_DIR))

# ---------- project DB layer (db.py, project root) ----------
from db import Database, build_in_clause

# ---------- extraction engine (Extraction/pipeline.py) ----------
# NOTE: importing pipeline installs its print() interceptor (builtins.print) as a
# module-level side effect. It is harmless here — with no LogWriter attached it
# just forwards to the real print() with flush=True.
try:
    from pipeline import (
        run_extraction,
        process_one_deal,
        load_xlsx_as_pipe_text,
        get_base_pdf_name,
    )
except Exception as e:  # ImportError, or SystemExit from a missing engine module
    sys.exit(
        f"[ERROR] Could not import the extraction engine from {_ENGINE_DIR}: "
        f"{type(e).__name__}: {e}"
    )

# The engine prints emoji/box-drawing diagnostics; reconfigure the console to UTF-8
# with replacement so a non-UTF-8 pipe can never crash a run.
for _stream in (sys.stdout, sys.stderr):
    try:
        _stream.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass


# =========================================================
# CONFIG (module-level constants — the sibling convention)
# =========================================================
# --- LLM / engine knobs (defaults mirror the project's extraction command line) ---
PROVIDER          = "gemini"            # openai | gemini | claude
MODEL             = "gemini-3.5-flash"
TIER              = "paid"              # free | paid  (claude is always paid)
MAX_TOKENS        = 65536
USE_LLM_ID        = True                # LLM-guided page identification before slicing
ID_MODEL          = "gemini-3.5-flash"  # model used for page identification
REPORTING_COLUMNS = None

# --- engine resource folders (verified to exist inside Extraction/) ---
PROMPTS_FOLDER     = _ENGINE_DIR / "prompts"
XLSX_PATH          = _ENGINE_DIR / "Master" / "standard_coa_master.xlsx"
COA_MAPPING_FOLDER = _ENGINE_DIR / "COA_Mapping"

# --- extraction output roots (same base as the validation stage; adjust if needed) ---
# Inputs come from TProcessStatus.PdfFilePath (validation leaves them in
# 03_Validated_Report), NOT from a scanned folder.
_DATA_ROOT    = Path(r"C:\Test Code\Public Finance")
OUTPUT_BASE   = _DATA_ROOT / "04_Validated_output"            # all-pass deals
MANUAL_OUTPUT = _DATA_ROOT / "05_Manual_Validation_Required"  # any-fail / total-check-fail deals


# =========================================================
# SMALL HELPERS
# =========================================================
def _norm(v) -> str:
    """Trimmed string for any value (None -> '')."""
    return "" if v is None else str(v).strip()


def sector_from_segment(seg_id) -> str:
    """
    SegmentId 1 = LG, 2 = NON-LG (business rule). Default LG when unknown/NULL.
    Returns the exact sector token the engine keys on (SECTOR_TABLE_SUFFIXES uses
    'LG' and 'NON-LG' with a hyphen).
    """
    try:
        seg_id = int(seg_id)
    except (TypeError, ValueError):
        return "LG"
    return "LG" if seg_id == 1 else ("NON-LG" if seg_id == 2 else "LG")


def _make_client():
    """Only OpenAI needs a pre-built client; gemini/claude paths build their own."""
    if PROVIDER == "openai":
        try:
            from openai import OpenAI
            return OpenAI()
        except Exception:
            return None
    return None


# =========================================================
# DB STATUS WRITERS  (all keyed on the physical TProcessStatus.Id)
# =========================================================
def mark_extraction_progress(db, row_id):
    """Row start: ExtractionFlag='p', DataValidationFlag='p', wipe any stale Remarks."""
    res = db.update("""
        UPDATE TProcessStatus
        SET ExtractionFlag = 'p',
            DataValidationFlag = 'p',
            Remarks = NULL,
            ModifiedOn = GETDATE()
        WHERE Id = ?
    """, [row_id])
    if not res.success:
        print(f"[FAILED] mark_extraction_progress row_id={row_id}: {res.error}")
    return res


def finalize_extraction(db, row_id, extraction_status, data_validation_status,
                        completion_status, output_path=None, remarks=None):
    """
    Write the terminal extraction-stage columns for one row (keyed on Id):

      extraction_status      -> ExtractionStatus       (0 fail / 1 produced JSON)
      data_validation_status -> DataValidationStatus   (0 fail / 1 all-pass)
      completion_status      -> CompletionStatus        (1 iff both above are 1)

    Always marks the stage complete: ExtractionFlag='c', DataValidationFlag='c'.
    OutputPath / Remarks are only written when provided (Remarks truncated to 500).
    """
    sets = [
        "ExtractionStatus = ?",
        "ExtractionFlag = 'c'",
        "DataValidationStatus = ?",
        "DataValidationFlag = 'c'",
        "CompletionStatus = ?",
        "ModifiedOn = GETDATE()",
    ]
    params = [extraction_status, data_validation_status, completion_status]

    if output_path is not None:
        sets.append("OutputPath = ?")
        params.append(str(output_path))
    if remarks is not None:
        sets.append("Remarks = ?")
        params.append(str(remarks)[:500])

    params.append(row_id)
    res = db.update(f"UPDATE TProcessStatus SET {', '.join(sets)} WHERE Id = ?", params)
    if not res.success:
        print(f"[FAILED] finalize_extraction row_id={row_id}: {res.error}")
    return res


# Column list + joins shared by the batch and single-id queries (callers append
# their own WHERE clause). Mirrors PFG_Validation.py's _VALIDATION_SELECT so the
# work-item shape (RowId, ProcessingId, PdfFilePath, SegmentId, ...) is identical.
_EXTRACTION_SELECT = """
    SELECT
        cm.CompanyId,
        ps.Id AS RowId,
        ps.ProcessingId,
        ps.ProcessingCode,
        cm.IssuerName,
        cm.State,
        cm.Sector,
        cm.SubSector,
        cm.UEI,
        cm.EIN,
        ps.ProcessYear,
        ps.FyeDate,
        ps.PdfFilePath,
        ps.COAID,
        co.SegmentId AS SegmentId
    FROM TCompanyMaster cm
    JOIN TProcessStatus ps
        ON ps.CompanyId = cm.CompanyId
    LEFT JOIN TCOAMaster co
        ON co.COAID = ps.COAID
"""


# =========================================================
# PER-ROW PROCESSOR
# =========================================================
def process_one_row(db, db_row, coa_text):
    """
    Extract ONE work-list row end to end. Wrapped in try/except so a single bad row
    records its own Remarks and the batch keeps going (Database autocommits each
    write; an *uncaught* exception would roll back every prior row's committed
    status inside the `with Database()` block).

    One DB row == one source PDF == one "deal". The row's PDF is copied into an
    isolated temp dir so the engine's folder-based slicer touches only it; the
    per-statement JSON is persisted under OUTPUT_BASE/<deal> (pass) or
    MANUAL_OUTPUT/<deal> (fail), and that folder is written to OutputPath.
    """
    row_id        = db_row.RowId
    processing_id = db_row.ProcessingId
    issuer_name   = _norm(db_row.IssuerName)
    sector        = sector_from_segment(getattr(db_row, "SegmentId", None))
    pdf_file_path = _norm(db_row.PdfFilePath)

    print(f"\n[ROW] Id={row_id} ProcessingId={processing_id} | {issuer_name} "
          f"sector={sector}")

    # Mark in-progress AND wipe any stale remark from a prior run.
    mark_extraction_progress(db, row_id)

    # ---- locate the validated source PDF ----
    src_pdf = Path(pdf_file_path) if pdf_file_path else None
    if (src_pdf is None) or (not src_pdf.exists()):
        print(f"[WARN] Id={row_id}: source PDF not found at {pdf_file_path!r}")
        finalize_extraction(db, row_id, 0, 0, 0,
                            remarks="source PDF not found for extraction")
        return

    deal_name = get_base_pdf_name(src_pdf.stem)
    tmp_dir = Path(tempfile.mkdtemp(prefix=f"pfg_ext_{row_id}_"))

    try:
        # Copy the source PDF into an isolated working dir so run_extraction slices
        # ONLY this row's PDF (03_Validated_Report holds many other rows' PDFs).
        shutil.copy2(str(src_pdf), str(tmp_dir / src_pdf.name))

        extracted = run_extraction(
            str(tmp_dir),
            prompts_folder = str(PROMPTS_FOLDER),
            use_llm_id     = USE_LLM_ID,
            id_model       = ID_MODEL,
        )
        if not extracted:
            print(f"[WARN] Id={row_id}: no extractable tables produced")
            finalize_extraction(db, row_id, 0, 0, 0,
                                remarks="no extractable tables found")
            return

        outcome = process_one_deal(
            deal_name          = deal_name,
            deal_pdfs          = extracted,
            prompts_folder     = str(PROMPTS_FOLDER),
            coa_text           = coa_text,
            output_folder      = str(OUTPUT_BASE),
            manual_output      = str(MANUAL_OUTPUT),
            raw_folder         = str(tmp_dir),
            provider           = PROVIDER,
            model              = MODEL,
            max_tokens         = MAX_TOKENS,
            reporting_columns  = REPORTING_COLUMNS,
            tier               = TIER,
            coa_mapping_folder = str(COA_MAPPING_FOLDER),
            client             = _make_client(),
            sector             = sector,
        )

        # Unexpected engine failure with nothing produced -> hard fail.
        if outcome.get("error") and not outcome.get("json_files"):
            finalize_extraction(db, row_id, 0, 0, 0,
                                output_path=outcome.get("output_folder"),
                                remarks=f"extraction error: {outcome['error']}")
            return

        extraction_status      = 1 if outcome.get("json_files") else 0
        data_validation_status = 1 if outcome.get("passed") else 0
        completion_status      = 1 if (extraction_status == 1 and
                                       data_validation_status == 1) else 0

        remark = None
        if completion_status != 1:
            if extraction_status == 0:
                remark = "extraction produced no JSON output"
            elif outcome.get("total_check_failed"):
                remark = "Total Check failed — routed to manual validation"
            else:
                remark = "one or more tables failed extraction — routed to manual validation"

        finalize_extraction(db, row_id, extraction_status, data_validation_status,
                            completion_status, output_path=outcome.get("output_folder"),
                            remarks=remark)
        print(f"[DONE] Id={row_id}: extraction={extraction_status} "
              f"data_validation={data_validation_status} completion={completion_status} "
              f"-> {outcome.get('output_folder')}")

    except Exception as e:
        msg = f"Extraction exception: {type(e).__name__}: {e}"[:500]
        print(f"[ERROR] Id={row_id}: {msg}")
        try:
            finalize_extraction(db, row_id, 0, 0, 0, remarks=msg)
        except Exception as e2:
            print(f"[ERROR] Could not persist exception remark for Id={row_id}: {e2}")
    finally:
        # The temp dir (copied source PDF + sliced sub-PDFs) is disposable — the
        # JSON now lives under OUTPUT_BASE/MANUAL_OUTPUT and is intentionally kept.
        shutil.rmtree(tmp_dir, ignore_errors=True)


# =========================================================
# BATCH MODE  (no CLI arg -> extract every pending row)
# =========================================================
def main_extract():
    coa_text = load_xlsx_as_pipe_text(str(XLSX_PATH))
    print(f"[COA] Loaded {len(coa_text.splitlines())} rows from {XLSX_PATH}")

    with Database() as db:
        print("[DB] Connected. Fetching rows pending extraction...")
        res = db.fetch_all(_EXTRACTION_SELECT + """
            WHERE cm.IsActive = 1
              AND ps.IsActive = 1
              AND ps.SourcingStatus = 1
              AND ps.SourcingFlag = 'c'
              AND ps.SourcingValidationStatus = 1
              AND ps.SourcingValidationFlag = 'c'
              AND (ps.ExtractionFlag = 'c' OR ps.ExtractionFlag IS NULL)
              AND (ps.ExtractionStatus = 0 OR ps.ExtractionStatus IS NULL)
              AND (ps.CompletionStatus = 0 OR ps.CompletionStatus IS NULL)
              AND ps.COAID IN (1, 2)
        """)
        if not res.success:
            print(f"[ERROR] Work-list query failed: {res.error}")
            return
        master_rows = res.data or []
        print(f"[INFO] {len(master_rows)} row(s) pending extraction.")
        if not master_rows:
            return

        # Claim the batch (flag 's') so a concurrent run won't pick the same rows.
        row_ids = [r.RowId for r in master_rows]
        clause, params = build_in_clause("Id", row_ids)
        claim = db.update(f"""
            UPDATE TProcessStatus
            SET ExtractionFlag = 's'
            WHERE IsActive = 1
              AND {clause}
        """, params)
        if not claim.success:
            print(f"[ERROR] Could not claim rows: {claim.error}")
            return

        for db_row in master_rows:
            process_one_row(db, db_row, coa_text)

    print("[DONE] Batch extraction complete.")


# =========================================================
# SINGLE-ID MODE  (python PFG_Extraction.py <id>)
# =========================================================
def extract_one_id(db, row_id, coa_text):
    """
    Extract a SINGLE TProcessStatus row by its physical Id, using an already-open db
    handle. Filters ONLY on IsActive (not the extraction flags): an explicit Id is
    extracted on demand regardless of its current flags — same rationale as
    PFG_Sourcing.source_one_id / PFG_Validation.validate_one_id.
    """
    res = db.fetch_one(_EXTRACTION_SELECT + """
        WHERE ps.Id = ? AND cm.IsActive = 1 AND ps.IsActive = 1
    """, [row_id])
    if not res.success or res.data is None:
        print(f"[INFO] No active TProcessStatus row for Id={row_id}.")
        return False
    db.update("UPDATE TProcessStatus SET ExtractionFlag='s' WHERE IsActive=1 AND Id=?", [row_id])
    process_one_row(db, res.data, coa_text)   # SAME per-row path as batch
    return True


def main_extract_by_id(row_id):
    coa_text = load_xlsx_as_pipe_text(str(XLSX_PATH))
    print(f"[COA] Loaded {len(coa_text.splitlines())} rows from {XLSX_PATH}")
    with Database() as db:
        print(f"[DB] Connected. Extracting single row_id={row_id}.")
        extract_one_id(db, row_id, coa_text)
    print(f"[DONE] Extraction completed for row_id={row_id}.")


if __name__ == "__main__":
    # python PFG_Extraction.py        -> extract ALL pending rows (batch)
    # python PFG_Extraction.py <id>   -> extract ONE TProcessStatus row by its Id
    if len(sys.argv) > 1:
        main_extract_by_id(int(sys.argv[1]))
    else:
        main_extract()
