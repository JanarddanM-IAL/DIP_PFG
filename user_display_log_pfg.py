"""
user_display_log_pfg.py
═══════════════════════
The PFG (Public Finance) copy of the short, human-readable log: one file, a
handful of rows per document, meant to be read by a person asking "what happened
to this report?".

This is the PFG-owned sibling of the C4F `user_display_log.py`. It is byte-for-byte
the same logic; ONLY the config block below differs (PFG env-var names + a PFG
default destination). Keeping it as a separate file means the two pipelines never
share a maintenance burden and their logs never collide.

It is the companion to the detailed logs (processing_log.parquet /
pipeline_logs.parquet), not a replacement. Those record everything a developer
needs to debug a run. This one records the steps a user cares about:

    Id  ProcessingId  Stage  SubStage          Seq  Remarks                    Time
    216 10            s      FAC Search        1    FAC sourcing started
    216 10            s      FAC Search        2    FAC Excel downloaded
    216 10            s      Report Download   1    PDF saved: <name>
    216 10            sv     Validation        1    Validation started
    216 10            sv     Validation        2    Entity Name: PASSED
    216 10            p      Extraction        1    Extraction started
    216 10            p      Extraction        2    Extraction finished

SCHEMA — 7 columns
──────────────────
    Id            Int64     TProcessStatus.Id
    ProcessingId  Int64     TProcessStatus.ProcessingId (null until resolved)
    Stage         Utf8      's' sourcing | 'sv' sourcing validation | 'p' processing
    SubStage      Utf8      the detailed log's Stage value ("FAC Search", "Extraction", …)
    Seq           Int64     1..N within one (Id, ProcessingId, Stage, SubStage) group
    Remarks       Utf8      free text for a milestone, "Name: Result" for a check
    Time          Datetime  when the row was recorded

No RunId and no ProcessingLogId. They are not needed for merging — there is one
shared file guarded by a cross-process FileLock rather than per-process shards —
and re-running a group REPLACES its rows, so there is never more than one
attempt's worth to tell apart. Seq exists because Time alone cannot guarantee a
stable display order within a group.

RE-RUN = REPLACE
────────────────
Flushing a group deletes every existing row with the same
(Id, ProcessingId, Stage, SubStage) and inserts the new ones. So re-running the
sourcing for a document replaces its sourcing rows and leaves its validation and
extraction rows alone. That is what keeps this log readable after a partial retry
instead of showing two interleaved attempts.

WHY BUFFER, AND WHY PER GROUP
─────────────────────────────
Parquet has no append: every write is read-whole-file → modify → rewrite. Rows
are therefore buffered in memory and flushed once per completed group — roughly
5-7 rewrites per document instead of one per row, while still matching the
delete-then-insert unit exactly. Flushing is cross-process safe: the lock is held
across the read and the write, and the write goes to a temp file that is then
atomically replaced, so a concurrent reader never sees a half-written file.

NOTHING HERE MAY BREAK A RUN
────────────────────────────
Every public function swallows its own exceptions. Losing a display-log row must
never cost the work it was describing. Failures print one line and move on.
"""

from __future__ import annotations

import os
import re
import threading
from datetime import datetime
from pathlib import Path
from typing import Any, Optional

# ═══════════════════════════════════════════════════════════════════════════════
#  CONFIG  (the ONLY difference from the C4F user_display_log.py)
# ═══════════════════════════════════════════════════════════════════════════════
# SINGLE CONTROL POINT for the PFG display log's destination. Change the default
# literal below, or set the PFG_LOG_DIR environment variable. This log lives under
# the Public Finance tree — a DIFFERENT directory than the C4F/ESG display log —
# and uses PFG_* env-var names so the two never collide even in one process.

LOG_DIR = Path(
    os.getenv("PFG_LOG_DIR", r"C:\S2\Public Finance\999_Log_Trackers")
)
LOG_FILENAME = os.getenv("PFG_USER_LOG_FILENAME", "user_display_log.parquet")

#: Seconds to wait for another process's read-modify-write.
LOCK_TIMEOUT = int(os.getenv("PFG_USER_LOG_LOCK_TIMEOUT", "120"))

# ── Stage buckets ─────────────────────────────────────────────────────────────
STAGE_SOURCING   = "s"
STAGE_VALIDATION = "sv"
STAGE_PROCESSING = "p"

VALID_STAGES = (STAGE_SOURCING, STAGE_VALIDATION, STAGE_PROCESSING)

#: Lifecycle order, which is NOT alphabetical order.
#:
#: Sorting the Stage column as text puts 'p' before 's' and 'sv', so a reader sees
#: extraction listed above the sourcing that produced its input — backwards, in
#: the one log whose whole purpose is being readable. Rows are ranked by this
#: instead, both on disk and on read.
_STAGE_RANK = {STAGE_SOURCING: 0, STAGE_VALIDATION: 1, STAGE_PROCESSING: 2}


def _sort_display(frame):
    """Sort CHRONOLOGICALLY within a document: groups in the order they started.

    Deliberately NOT grouped by Stage. A document's real history can revisit the
    same bucket more than once, so ranking purely by s/sv/p would destroy the
    sequence. Stage is a label on each row, not the axis to read along.

    Groups are ordered by the EARLIEST Time in the group, and rows within a group
    by Seq. Both are derived from the data, not from the in-process registration:
    a reader — the `python user_display_log_pfg.py --id 209` CLI, a dashboard,
    anything that did not run the pipeline — has an empty registry, so anything
    depending on registration order would fall back to sorting SubStage as text.

    Grouping by min(Time) rather than sorting every row by Time keeps a group
    contiguous even when two stages overlap in time.
    """
    import polars as pl
    return (
        frame
        .with_columns(
            pl.col("Time").min()
              .over(["Id", "ProcessingId", "Stage", "SubStage"])
              .alias("_started")
        )
        .sort(["Id", "ProcessingId", "_started", "SubStage", "Seq"])
        .drop("_started")
    )

#: sub-stage -> (bucket, human label, display order).
#:
#: Registration is an ALLOW-LIST, not a hint. A sub-stage that is not registered
#: produces no rows at all. That is the difference between this log and the
#: detailed one: feeding it every stage both pipelines report gives a user
#: internal bookkeeping which is noise to them. Each pipeline names the handful of
#: steps worth showing.
#:
#: The bucket cannot be guessed from the name either, so each entry states it.
_STAGES: dict[str, tuple] = {}
_NEXT_ORDER = [0]


def register_stages(stages) -> None:
    """Declare the sub-stages worth showing, in the order they happen.

        register_stages([
            ("FAC Search",      "s",  "FAC sourcing"),
            ("Report Download", "s",  "Report download"),
            ("Validation",      "sv", "Sourcing validation"),
            ("Extraction",      "p",  "Extraction"),
        ])

    Each entry is (sub_stage, bucket, label). `sub_stage` is what lands in the
    SubStage column verbatim; `label` is only used to phrase auto-generated
    remarks, so a user reads "Extraction started" rather than "EXTRACT started".

    LIST ORDER IS DISPLAY ORDER. Sorting SubStage as text would be alphabetical,
    not chronological, in a log whose only job is to be read in sequence.

    Additive across calls, so multiple pipelines can register their own without
    clobbering each other when one process imports both.
    """
    for entry in (stages or []):
        try:
            if len(entry) == 3:
                sub_stage, bucket, label = entry
            else:
                sub_stage, bucket = entry
                label = str(sub_stage)
        except (TypeError, ValueError):
            _warn(f"ignoring malformed stage entry {entry!r}")
            continue
        if bucket not in VALID_STAGES:
            _warn(f"ignoring stage {sub_stage!r}: bucket {bucket!r} is not one of "
                  f"{VALID_STAGES}")
            continue
        key = str(sub_stage)
        if key in _STAGES:                       # keep the original position
            order = _STAGES[key][2]
        else:
            order = _NEXT_ORDER[0]
            _NEXT_ORDER[0] += 1
        _STAGES[key] = (bucket, str(label or key), order)


def register_stage_buckets(mapping: dict[str, str]) -> None:
    """Backwards-compatible shim: a plain {sub_stage: bucket} mapping.

    Dicts preserve insertion order, so the caller still controls display order.
    """
    register_stages([(k, v, k) for k, v in (mapping or {}).items()])


def is_registered(sub_stage: str) -> bool:
    return str(sub_stage) in _STAGES


def bucket_for(sub_stage: str) -> Optional[str]:
    """The bucket, or None when the sub-stage is not on the allow-list."""
    entry = _STAGES.get(str(sub_stage))
    return entry[0] if entry else None


def label_for(sub_stage: str) -> str:
    entry = _STAGES.get(str(sub_stage))
    return entry[1] if entry else str(sub_stage)


def order_for(sub_stage: str) -> int:
    entry = _STAGES.get(str(sub_stage))
    return entry[2] if entry else 999


# ═══════════════════════════════════════════════════════════════════════════════
#  SCHEMA
# ═══════════════════════════════════════════════════════════════════════════════

COLUMNS = ("Id", "ProcessingId", "Stage", "SubStage", "Seq", "Remarks", "Time")

#: The columns that identify a replaceable group.
GROUP_KEY = ("Id", "ProcessingId", "Stage", "SubStage")


def _schema():
    import polars as pl
    return {
        "Id":           pl.Int64,
        "ProcessingId": pl.Int64,
        "Stage":        pl.Utf8,
        "SubStage":     pl.Utf8,
        "Seq":          pl.Int64,
        "Remarks":      pl.Utf8,
        "Time":         pl.Datetime("us"),
    }


# ── Year partitioning ─────────────────────────────────────────────────────────
# One file per TProcessStatus.ProcessYear: user_display_log_2025.parquet. The
# year is a PER-DOCUMENT value, so it is carried on the context alongside the
# ids and chosen when a group is flushed — every row of a document lands in that
# document's year file, whatever calendar year the run happens in.
#
# A row whose ProcessYear is unknown/unusable falls back to the unpartitioned
# LOG_FILENAME, which therefore means "process-year unknown" rather than
# "everything"; pre-partitioning history keeps living there untouched.
_STEM   = Path(LOG_FILENAME).stem
_SUFFIX = Path(LOG_FILENAME).suffix or ".parquet"

#: Matches a per-year file. Exactly four digits, so it can never collide with
#: the unpartitioned name.
_YEAR_RE = re.compile(rf"^{re.escape(_STEM)}_(\d{{4}}){re.escape(_SUFFIX)}$")


def coerce_year(value: Any) -> int | None:
    """A usable 4-digit process year, or None. Accepts the string form the DB
    layer hands over ("2025"), and rejects anything out of range."""
    y = _coerce_int(value)
    return y if y is not None and 1900 <= y <= 2999 else None


def log_path(year: Any = None) -> Path:
    """The file a document with this ProcessYear writes to."""
    y = coerce_year(year)
    return LOG_DIR / (f"{_STEM}_{y}{_SUFFIX}" if y is not None else LOG_FILENAME)


def log_paths() -> list[Path]:
    """Every display-log file at this location: each year file plus the
    unpartitioned one. This is what read_log() spans, so a reader never has to
    know which years exist."""
    if not LOG_DIR.is_dir():
        return []
    out = [p for p in sorted(LOG_DIR.glob(f"{_STEM}*{_SUFFIX}"))
           if p.name == LOG_FILENAME or _YEAR_RE.match(p.name)]
    return out


def _warn(msg: str) -> None:
    print(f"[USER-LOG] {msg}", flush=True)


# ═══════════════════════════════════════════════════════════════════════════════
#  BUFFER
# ═══════════════════════════════════════════════════════════════════════════════
# Rows accumulate per (Id, ProcessingId, Stage, SubStage) and are written when the
# group is flushed. The lock is in-process only: the buffer belongs to this
# process, and the cross-process guard is the FileLock taken inside _flush().

_LOCK = threading.Lock()
_BUFFER: dict[tuple, list[dict[str, Any]]] = {}
_CTX: dict[str, Any] = {"row_id": None, "processing_id": None, "year": None}


def set_context(row_id: Any = None, processing_id: Any = None,
                year: Any = None) -> None:
    """Whose rows these are. Call again when the ProcessingId becomes known.

    Both are remembered independently, so passing only processing_id keeps the
    row_id already set. Rows already buffered pick the resolved id up at flush
    time, so a milestone recorded before the lookup still lands under the right
    document.

    Moving to a DIFFERENT row_id flushes first. Without that, buffered rows from
    the previous document would be stamped with the new document's identity when
    they were finally written — silent mis-attribution.

    `year` is TProcessStatus.ProcessYear and selects the file this document's
    rows are written to (user_display_log_<year>.parquet). Like the ids it is
    remembered independently, so a later two-argument call cannot blank it. Pass
    it as soon as the DB row is known; a document flushed without one lands in
    the unpartitioned file.
    """
    new_row_id = _coerce_int(row_id) if row_id is not None else None
    moved = (new_row_id is not None
             and _CTX["row_id"] is not None
             and new_row_id != _CTX["row_id"])
    if moved:
        flush_all()

    with _LOCK:
        if row_id is not None:
            _CTX["row_id"] = new_row_id
        if processing_id is not None:
            _CTX["processing_id"] = _coerce_int(processing_id)
        # The year belongs to the DOCUMENT, so arriving at a new one drops the
        # previous year FIRST. Without this, "remembered independently" means a
        # document whose ProcessYear is NULL silently inherits the last
        # document's year and is filed under it — the exact mis-attribution the
        # row_id flush above exists to prevent.
        if moved:
            _CTX["year"] = None
        if year is not None:
            _CTX["year"] = coerce_year(year)


def clear_context() -> None:
    """Forget the current document. Flushes first — see set_context()."""
    flush_all()
    with _LOCK:
        _CTX["row_id"] = None
        _CTX["processing_id"] = None
        _CTX["year"] = None


def _coerce_int(value: Any) -> Optional[int]:
    """Int64 or None. A non-numeric id must not take the run down over a log row.

    This is also what keeps a "row-209" placeholder out of a user-facing column:
    it simply becomes null here.
    """
    if value is None or value == "":
        return None
    try:
        return int(value)
    except (TypeError, ValueError):
        return None


def _key(stage: str, sub_stage: str) -> tuple:
    """The buffer key: stage and sub-stage ONLY.

    Deliberately does NOT include Id/ProcessingId. Those are stamped at flush
    time, because ProcessingId may not be known when the first rows of a run are
    recorded. Keying on it here would split one logical group in two. Stamping at
    flush keeps a group whole.
    """
    return (stage, str(sub_stage))


def _add(stage: str, sub_stage: str, remark: str) -> None:
    if not remark:
        return
    with _LOCK:
        _BUFFER.setdefault(_key(stage, sub_stage), []).append({
            "Remarks": str(remark).strip(),
            "Time":    datetime.now(),
        })


# ═══════════════════════════════════════════════════════════════════════════════
#  PUBLIC API
# ═══════════════════════════════════════════════════════════════════════════════

def step(sub_stage: str, remark: str, stage: str | None = None) -> None:
    """A milestone, in plain words — "FAC sourcing started".

    A no-op for an unregistered sub-stage: this log is an allow-list, so a stage
    nobody asked to display simply does not appear.
    """
    try:
        b = stage or bucket_for(sub_stage)
        if b is None:
            return
        _add(b, sub_stage, remark)
    except Exception as exc:
        _warn(f"step() failed: {type(exc).__name__}: {exc}")


def check(sub_stage: str, name: str, result: str, reason: str = "",
          stage: str | None = None) -> None:
    """A checkpoint, rendered as "Name: Result" in one column.

    `reason` is appended only when the result is a failure. A passing check
    explains itself ("Entity Name: PASSED"); a failing one does not, and the
    reason is the only place that says why.
    """
    try:
        b = stage or bucket_for(sub_stage)
        if b is None:
            return
        # A checkpoint with no result has nothing to report.
        if result is None or str(result).strip() == "":
            return
        label = f"{name}: {result}"
        if reason and _looks_failed(result):
            label = f"{label} — {reason}"
        _add(b, sub_stage, label)
    except Exception as exc:
        _warn(f"check() failed: {type(exc).__name__}: {exc}")


def _looks_failed(result: Any) -> bool:
    r = str(result or "").strip().lower()
    return bool(r) and r not in ("ok", "pass", "passed", "skipped", "true", "1")


def failure(sub_stage: str, reason: str, stage: str | None = None) -> None:
    """A failed step. Prefixed so it stands out when skimming the column.

    Unlike step() and check() this is recorded even for an unregistered
    sub-stage — a failure the user is never told about is the one thing worse
    than a noisy log.
    """
    try:
        _add(stage or bucket_for(sub_stage) or STAGE_PROCESSING,
             sub_stage, f"FAILED: {reason}")
    except Exception as exc:
        _warn(f"failure() failed: {type(exc).__name__}: {exc}")


def started(sub_stage: str) -> None:
    """"<Label> started", phrased from the registered human label."""
    step(sub_stage, f"{label_for(sub_stage)} started")


def finished(sub_stage: str) -> None:
    step(sub_stage, f"{label_for(sub_stage)} finished")


def flush_group(sub_stage: str, stage: str | None = None) -> int:
    """Write one completed group, replacing any previous attempt at it.

    Returns the number of rows written, or -1 on error.
    """
    try:
        b = stage or bucket_for(sub_stage)
        if b is None:
            # Nothing was buffered for it, but a failure() may have used the
            # fallback bucket — flush that key too rather than stranding it.
            b = STAGE_PROCESSING
        return _flush([_key(b, sub_stage)])
    except Exception as exc:
        _warn(f"flush_group() failed: {type(exc).__name__}: {exc}")
        return -1


def flush_all() -> int:
    """Write every buffered group. Call at the end of a run, in a finally."""
    try:
        with _LOCK:
            keys = list(_BUFFER.keys())
        return _flush(keys) if keys else 0
    except Exception as exc:
        _warn(f"flush_all() failed: {type(exc).__name__}: {exc}")
        return -1


def buffered_groups() -> list[tuple]:
    """The groups still waiting to be written. For tests and diagnostics."""
    with _LOCK:
        return list(_BUFFER.keys())


# ═══════════════════════════════════════════════════════════════════════════════
#  WRITE
# ═══════════════════════════════════════════════════════════════════════════════

def _flush(keys: list[tuple]) -> int:
    """Replace each named group in the file with its buffered rows.

    The whole read-modify-write happens under one cross-process FileLock. That is
    what makes 3-4 concurrent Dagster workers safe: without it, two processes
    would each read the same baseline and the second write would silently drop
    the first one's rows.
    """
    import polars as pl
    from filelock import FileLock

    with _LOCK:
        payload = {k: _BUFFER.pop(k) for k in keys if k in _BUFFER}
        # Read the identity once, here, so every row in a group agrees on it even
        # if the ProcessingId was resolved partway through recording the group.
        row_id = _CTX["row_id"]
        processing_id = _CTX["processing_id"]
        # Read the year here too, for the same reason: every row of this flush
        # must agree on which year file it belongs to, even if the context was
        # completed partway through recording the group.
        year = _CTX["year"]
    if not payload:
        return 0

    schema = _schema()
    new_rows: list[dict[str, Any]] = []
    groups: list[tuple] = []
    for (stage, sub_stage), rows in payload.items():
        groups.append((row_id, processing_id, stage, sub_stage))
        for seq, row in enumerate(rows, start=1):
            new_rows.append({
                "Id":           row_id,
                "ProcessingId": processing_id,
                "Stage":        stage,
                "SubStage":     sub_stage,
                "Seq":          seq,
                "Remarks":      row["Remarks"],
                "Time":         row["Time"],
            })

    # The document's ProcessYear picks the file; the replace-on-re-run below then
    # operates within that year, which is correct because a document's
    # ProcessYear does not change between runs.
    path = log_path(year)
    try:
        path.parent.mkdir(parents=True, exist_ok=True)
    except OSError as exc:
        _warn(f"cannot create {path.parent}: {exc}")
        return -1

    lock = FileLock(str(path) + ".lock", timeout=LOCK_TIMEOUT)
    try:
        with lock:
            existing = None
            if path.is_file():
                try:
                    existing = pl.read_parquet(path)
                except Exception as exc:
                    # A corrupt file must not block the run forever, but silently
                    # starting over would destroy history — say so loudly and
                    # keep going with just this run's rows.
                    _warn(f"could not read {path.name} ({exc}); rewriting it")
                    existing = None

            frame = pl.DataFrame(new_rows, schema=schema)

            if existing is not None and existing.height:
                existing = _align(existing, schema)
                for g_row_id, g_pid, g_stage, g_sub in groups:
                    # Also drop any rows for this document+group that were left
                    # under a null ProcessingId by an earlier partial run, or the
                    # replace would leave those behind as duplicates.
                    existing = existing.filter(
                        ~(
                            _eq(pl, "Id", g_row_id)
                            & (pl.col("ProcessingId").is_null()
                               | _eq(pl, "ProcessingId", g_pid))
                            & (pl.col("Stage") == g_stage)
                            & (pl.col("SubStage") == g_sub)
                        )
                    )
                frame = pl.concat([existing, frame], how="vertical")

            _atomic_write(_sort_display(frame), path)
        return len(new_rows)
    except Exception as exc:
        _warn(f"flush failed for {len(new_rows)} row(s): "
              f"{type(exc).__name__}: {exc}")
        return -1


def _eq(pl, column: str, value):
    """Equality that also matches NULL, which `col == None` does not.

    ProcessingId is null on rows written before the DB lookup resolved it, and
    those still have to be replaceable — a plain == would never match them, so
    the old rows would survive the delete and the group would end up duplicated.
    """
    if value is None:
        return pl.col(column).is_null()
    return pl.col(column) == value


def _align(frame, schema: dict):
    """Force a frame read from disk onto the current schema.

    An older file missing a column, or carrying it with a different type, would
    otherwise fail the concat. Missing columns become null.
    """
    import polars as pl
    out = frame
    for col, dtype in schema.items():
        if col not in out.columns:
            out = out.with_columns(pl.lit(None).cast(dtype).alias(col))
    return out.select([pl.col(c).cast(t, strict=False) for c, t in schema.items()])


def _atomic_write(frame, path: Path) -> None:
    """Temp file then replace, so a reader never sees a partial file."""
    tmp = path.with_suffix(path.suffix + f".tmp{os.getpid()}")
    frame.write_parquet(tmp, compression="snappy")
    os.replace(tmp, path)


# ═══════════════════════════════════════════════════════════════════════════════
#  READ
# ═══════════════════════════════════════════════════════════════════════════════

def read_log(year: Any = None):
    """The display log, in display order. Empty frame if there is none.

    With no argument this spans EVERY year file plus the unpartitioned one, so
    year partitioning is invisible to readers and read_for()/describe_for() keep
    working unchanged. Pass `year` to read a single year's file.
    """
    import polars as pl
    schema = _schema()
    paths = [log_path(year)] if year is not None else log_paths()
    frames = []
    for p in paths:
        if not p.is_file():
            continue
        try:
            frames.append(_align(pl.read_parquet(p), schema))
        except Exception as exc:
            # One unreadable year must not hide the others.
            _warn(f"could not read {p.name} ({exc}); skipping it")
    if not frames:
        return pl.DataFrame(schema=schema)
    return _sort_display(pl.concat(frames, how="vertical"))


def read_for(row_id: Any = None, processing_id: Any = None):
    """The rows for one document, in display order."""
    import polars as pl
    frame = read_log()
    if row_id is not None:
        frame = frame.filter(pl.col("Id") == _coerce_int(row_id))
    if processing_id is not None:
        frame = frame.filter(pl.col("ProcessingId") == _coerce_int(processing_id))
    return frame


def describe_for(row_id: Any = None, processing_id: Any = None) -> str:
    """The rows for one document as plain text — what a user would be shown."""
    frame = read_for(row_id, processing_id)
    if not frame.height:
        return "(no display-log rows)"
    lines = []
    last = None
    for r in frame.iter_rows(named=True):
        head = (r["Stage"], r["SubStage"])
        if head != last:
            lines.append(f"[{r['Stage']}] {r['SubStage']}")
            last = head
        stamp = str(r["Time"])[11:19]
        lines.append(f"    {stamp}  {r['Remarks']}")
    return "\n".join(lines)


if __name__ == "__main__":
    import argparse

    ap = argparse.ArgumentParser(description="Read the PFG user display log.")
    ap.add_argument("--id", type=int, default=None, help="TProcessStatus.Id")
    ap.add_argument("--processing-id", type=int, default=None)
    ap.add_argument("--raw", action="store_true", help="Print the frame, not text.")
    args = ap.parse_args()

    _files = log_paths()
    print(f"[USER-LOG] {LOG_DIR}")
    for _p in _files:
        print(f"[USER-LOG]   {_p.name}")
    if not _files:
        print(f"[USER-LOG]   (none yet; next write -> {log_path().name})")
    if args.raw:
        print(read_for(args.id, args.processing_id))
    elif args.id or args.processing_id:
        print(describe_for(args.id, args.processing_id))
    else:
        frame = read_log()
        print(f"{frame.height} row(s)")
        print(frame)
