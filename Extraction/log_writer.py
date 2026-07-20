"""
log_writer.py
─────────────
Writes pipeline log lines to a persistent Parquet file using Polars.
No DuckDB / pandas dependency — pure Polars read → concat → write pattern.

Parquet schema (pipeline_logs.parquet):
    ProcessingLogId  Int64     — auto-incremented internally per print() call
    ProcessingId     Int64     — DB ID corresponding to the document being processed
    Stage            Utf8      — s / p / sv / pv  (see STAGE_* constants)
    Remarks          Utf8      — log message text
    Time             Datetime  — auto-set to datetime.now() on insert

Usage:
    # Explicit insert (all fields supplied by caller):
    lw.write_log(log_id, processing_id, STAGE_PROCESSING, "[INFO] ...")

    # Print interceptor (auto-increments ProcessingLogId):
    lw.set_context(processing_id, STAGE_PROCESSING)
    print("[INFO] ...")   # captured automatically via _logging_print
"""

import atexit
import os
import signal
import threading
from datetime import datetime
from pathlib import Path

import polars as pl
from filelock import FileLock

# ── Paths ─────────────────────────────────────────────────────────────────────
# STANDALONE-run default only. The DB/Dagster workflow injects the real location via
# LogWriter(log_dir=PFG_Extraction.LOG_DIR) (see pipeline.start_pipeline_logging), which
# derives from the single control point PFG_Extraction._DATA_ROOT. An operational log,
# kept in its own 999_Log_Trackers folder — separate from the data parquets.
LOG_DIR     = Path(r"C:\S2\Public Finance") / "999_Log_Trackers"
LOG_PARQUET = LOG_DIR / "pipeline_logs.parquet"

# ── Concurrency (parallel jobs append to the SAME log parquet) ───────────────
# Thread lock (in-process) + a per-instance cross-process FileLock (see __init__),
# mirroring PFG_Sourcing/PFG_Validation. Held around the whole read→concat→write.
_LOG_THREAD_LOCK = threading.Lock()


def _atomic_write_parquet(df: "pl.DataFrame", path: Path) -> None:
    """Write to a sibling .tmp then os.replace() so a reader never sees a partial file."""
    tmp = path.with_suffix(".parquet.tmp")
    df.write_parquet(tmp, compression="snappy")
    os.replace(tmp, path)

# ── Stage constants ───────────────────────────────────────────────────────────
STAGE_SOURCING              = "s"
STAGE_PROCESSING            = "p"
STAGE_SOURCING_VALIDATION   = "sv"
STAGE_PROCESSING_VALIDATION = "pv"

# ── Polars schema (single source of truth) ────────────────────────────────────
_SCHEMA: dict[str, pl.DataType] = {
    "ProcessingLogId": pl.Int64,
    "ProcessingId":    pl.Utf8,    # ← CHANGED from pl.Int64 to pl.Utf8
    "Stage":           pl.Utf8,
    "Remarks":         pl.Utf8,
    "Time":            pl.Datetime("us"),
}

def _empty_df() -> pl.DataFrame:
    return pl.DataFrame(
        {col: pl.Series(col, [], dtype=dtype) for col, dtype in _SCHEMA.items()}
    )


def _read_existing(path: Path) -> pl.DataFrame:
    """
    Read the existing Parquet file at `path`.
    Returns an empty DataFrame (correct schema) if the file doesn't exist
    or is unreadable.
    """
    if not Path(path).exists():
        return _empty_df()
    try:
        return pl.read_parquet(path)
    except Exception:
        return _empty_df()


# ─────────────────────────────────────────────────────────────────────────────
# LogWriter
# ─────────────────────────────────────────────────────────────────────────────

class LogWriter:
    """
    Polars-backed log writer. Rows accumulate in an in-memory buffer;
    on flush() they are appended to the Parquet file via:
        read existing → pl.concat → write_parquet

    Typical lifecycle in pipeline.py:

        lw = LogWriter()

        # Option A — explicit insert (you supply all fields):
        lw.write_log(log_id, processing_id, STAGE_PROCESSING,
                     "[INFO] Starting extraction ...")

        # Option B — print interceptor (auto-increments ProcessingLogId):
        lw.set_context(processing_id, STAGE_PROCESSING)
        print("[INFO] ...")   # captured via _logging_print in pipeline.py

        lw.close()
    """

    FLUSH_EVERY = 50  # rows buffered before auto-flush

    def __init__(self, log_dir=None):
        # log_dir lets the DB/Dagster workflow inject its centralized path
        # (PFG_Extraction.LOG_DIR); standalone runs fall back to module LOG_DIR.
        self.log_dir     = Path(log_dir) if log_dir else LOG_DIR
        self.log_parquet = self.log_dir / "pipeline_logs.parquet"
        # Cross-process lock for THIS log file (per-instance so it follows log_dir;
        # all processes on the same file share the one lock file).
        self._file_lock  = FileLock(str(self.log_parquet.with_suffix(".parquet.lock")), timeout=60)

        self.log_dir.mkdir(parents=True, exist_ok=True)

        self._auto_log_id:    int        = 0
        self._current_pid:    str | None = None   # ← NOW a string (filename)
        self._current_stage:  str        = STAGE_PROCESSING
        self._buffer: list[dict]         = []
        self._closed                     = False

        # Create the empty-schema file if missing (guarded so parallel processes
        # don't race to create it).
        if not self.log_parquet.exists():
            with self._file_lock:
                if not self.log_parquet.exists():
                    _atomic_write_parquet(_empty_df(), self.log_parquet)

        atexit.register(self._atexit_handler)
        self._register_signal_handlers()
    # ── Print interceptor path ────────────────────────────────────────────────

    def write(self, line: str) -> None:
        """Called by _logging_print interceptor for every print() call."""
        sub_lines = line.split("\n")
        for sub in sub_lines:
            sub = sub.rstrip("\r")
            if not sub.strip():
                continue
            self._auto_log_id += 1
            self._buffer.append({
                "ProcessingLogId": self._auto_log_id,
                "ProcessingId":    self._current_pid,   # filename string
                "Stage":           self._current_stage,
                "Remarks":         sub,
                "Time":            datetime.now(),
            })
        if len(self._buffer) >= self.FLUSH_EVERY:
            self.flush()

    def set_context(self, processing_id: str, stage: str) -> None:  # ← str now
        """
        Call when starting a new file so subsequent write() calls
        carry the correct ProcessingId (filename) and Stage.

            lw.set_context("LG_CIT_WI_600005841_2024.pdf", "p")
        """
        self._current_pid   = processing_id
        self._current_stage = stage

    # ── Explicit insert ───────────────────────────────────────────────────────

    def write_log(
        self,
        processing_log_id: int,
        processing_id:     int,
        stage:             str,
        remarks:           str,
    ) -> None:
        """
        Insert a log row. All fields are supplied by the caller;
        Time is set automatically.

        Usage:
            lw.write_log(log_id, processing_id, STAGE_PROCESSING,
                         "[INFO] Extracting page 12 ...")
        """
        self._buffer.append({
            "ProcessingLogId": processing_log_id,
            "ProcessingId":    processing_id,
            "Stage":           stage,
            "Remarks":         remarks,
            "Time":            datetime.now(),
        })
        if len(self._buffer) >= self.FLUSH_EVERY:
            self.flush()

    # ── Flush buffer → Parquet ────────────────────────────────────────────────

    def flush(self) -> None:
        if not self._buffer:
            return

        new_df = pl.DataFrame(
            {col: [row[col] for row in self._buffer] for col in _SCHEMA},
            schema=_SCHEMA,
        )

        # Read→concat→write inside BOTH locks so parallel jobs never lose appends and
        # never see a torn file. Append-only ⇒ reading the latest under the lock is enough.
        with _LOG_THREAD_LOCK:
            with self._file_lock:
                existing_df = _read_existing(self.log_parquet)
                updated_df  = pl.concat([existing_df, new_df], how="diagonal_relaxed")
                _atomic_write_parquet(updated_df, self.log_parquet)

        self._buffer.clear()

    # ── Clean shutdown ────────────────────────────────────────────────────────

    def close(self) -> None:
        if self._closed:
            return
        self._closed = True
        try:
            self.flush()
        except Exception:
            pass

    # ── Safety net: atexit + signal handlers ─────────────────────────────────

    def _atexit_handler(self) -> None:
        if not self._closed:
            try:
                self.close()
            except Exception:
                pass

    def _register_signal_handlers(self) -> None:
        import threading
        if threading.current_thread() is not threading.main_thread():
            return

        def _handler(signum, frame):
            if not self._closed:
                try:
                    self.close()
                except Exception:
                    pass
            signal.signal(signum, signal.SIG_DFL)
            signal.raise_signal(signum)

        for sig in (signal.SIGTERM, signal.SIGINT):
            try:
                signal.signal(sig, _handler)
            except (OSError, ValueError):
                pass


# ─────────────────────────────────────────────────────────────────────────────
# Module-level write_log() — drop-in for direct import usage
# ─────────────────────────────────────────────────────────────────────────────

_default_writer: LogWriter | None = None


def _get_default_writer() -> LogWriter:
    global _default_writer
    if _default_writer is None:
        _default_writer = LogWriter()
    return _default_writer


def write_log(
    processing_log_id: int,
    processing_id:     int,
    stage:             str,
    remarks:           str,
) -> None:
    """
    Module-level drop-in — works without instantiating LogWriter manually.

    Usage:
        from log_writer import write_log, STAGE_PROCESSING

        write_log(log_id, processing_id, STAGE_PROCESSING,
                  "[INFO] Starting extraction ...")
    """
    _get_default_writer().write_log(processing_log_id, processing_id, stage, remarks)