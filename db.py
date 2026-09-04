# -*- coding: utf-8 -*-
"""
db.py — A small, generic, reusable database layer for pyodbc / ODBC databases.

Project-agnostic: this module contains no application- or schema-specific logic.
It is driven entirely by environment variables (.env) and the SQL you pass in,
so it can be dropped into any Python project that talks to an ODBC data source
(SQL Server, PostgreSQL, MySQL, etc. — anything with an ODBC driver).

A single place for:
  * get_db_connection()  -> a configured pyodbc connection (from .env)
  * Database             -> a small class you instantiate once and reuse; you
                            pass a SQL query (+ params) to its methods and get
                            back a uniform DBResult carrying a success flag,
                            any returned rows, the affected/returned row count,
                            and an error message on failure.

Design goals
------------
- "Just pass the query": every method takes a SQL string and optional params
  and returns a DBResult. A bad query never raises out of these methods — it
  comes back as DBResult(success=False, error=...).
- Safe by default: values are always bound with pyodbc "?" placeholders.
- Reusable: import from any file, in any project:

    from db import Database

    with Database() as db:                     # opens once, auto commit+close
        res = db.fetch_all("SELECT id, name FROM users WHERE active = ?", [1])
        if res.success:
            for row in res.data:
                print(row.id, row.name)

        upd = db.update("UPDATE users SET status = ? WHERE id = ?", ["done", 123])
        print(upd.success, upd.rowcount)

Environment (.env)
------------------
    DB_DRIVER=ODBC Driver 17 for SQL Server   # any installed ODBC driver
    DB_SERVER=<host,port>
    DB_DATABASE=<database>
    DB_USERNAME=<user>
    DB_PASSWORD=<password>
    DB_TRUSTED_CONNECTION=no                   # yes -> Windows/integrated auth
"""

from __future__ import annotations

import os
import re
import logging
from dataclasses import dataclass
from typing import Any, Optional, Sequence, Tuple

import pyodbc
from dotenv import load_dotenv


# =========================================================
# LOGGING
# =========================================================
# Library-style logging: attach a NullHandler so importing this module never
# forces logging config on the host application. Callers can configure the
# "db" logger if they want to see messages.
logger = logging.getLogger("db")
logger.addHandler(logging.NullHandler())


# =========================================================
# CONNECTION FACTORY
# =========================================================
def get_db_connection() -> "pyodbc.Connection":
    """
    Create a pyodbc connection using credentials from the .env file.

    .env must contain:
        DB_DRIVER=ODBC Driver 17 for SQL Server
        DB_SERVER=your_server[,port]
        DB_DATABASE=your_database
        DB_USERNAME=your_user
        DB_PASSWORD=your_password
        DB_TRUSTED_CONNECTION=no

    Returns a live pyodbc.Connection. Raises RuntimeError if required
    configuration is missing.
    """
    # override=True makes the .env file authoritative over any value already in the
    # OS environment. Without it, a stale GEMINI_API_KEY / ANTHROPIC_API_KEY left in
    # the machine's user environment (Windows HKCU\Environment) shadows the correct
    # key in .env — the app then sends the stale key and Gemini rejects it as
    # "API key not valid" even though .env has the right one. DB creds happened to
    # work only because they were NOT also present in the OS environment.
    load_dotenv(override=True)

    driver   = os.getenv("DB_DRIVER", "ODBC Driver 17 for SQL Server")
    server   = os.getenv("DB_SERVER")
    database = os.getenv("DB_DATABASE")
    username = os.getenv("DB_USERNAME")
    password = os.getenv("DB_PASSWORD")
    trusted  = os.getenv("DB_TRUSTED_CONNECTION", "no").strip().lower()

    if not server or not database:
        raise RuntimeError(
            "DB_SERVER and DB_DATABASE must be set in .env file.\n"
            "Copy .env.example to .env and fill in your credentials."
        )

    if trusted in ("yes", "true", "1"):
        conn_str = (
            f"DRIVER={{{driver}}};"
            f"SERVER={server};"
            f"DATABASE={database};"
            f"Trusted_Connection=yes;"
            f"TrustServerCertificate=yes;"
        )
    else:
        if not username or not password:
            raise RuntimeError(
                "DB_USERNAME and DB_PASSWORD must be set in .env "
                "(or use DB_TRUSTED_CONNECTION=yes)."
            )
        conn_str = (
            f"DRIVER={{{driver}}};"
            f"SERVER={server};"
            f"DATABASE={database};"
            f"UID={username};"
            f"PWD={password};"
            f"TrustServerCertificate=yes;"
        )

    return pyodbc.connect(conn_str)


# =========================================================
# RESULT TYPE
# =========================================================
@dataclass
class DBResult:
    """
    Uniform return value for every Database method.

    success  : True if the statement ran without error, False otherwise.
    data     : fetch_all -> list[pyodbc.Row]; fetch_one -> pyodbc.Row | None;
               writes    -> None.
    rowcount : rows affected (writes) or rows returned (reads); -1 if unknown.
    error    : error message string when success is False, else None.
    """
    success: bool
    data: Any = None
    rowcount: int = -1
    error: Optional[str] = None

    def __bool__(self) -> bool:
        return self.success

    @property
    def rows(self) -> list:
        """Always return a list, whether data is a list, a single row, or None."""
        if self.data is None:
            return []
        if isinstance(self.data, list):
            return self.data
        return [self.data]

    @property
    def one(self) -> Any:
        """Return a single row: data itself, or the first row of a list, else None."""
        if isinstance(self.data, list):
            return self.data[0] if self.data else None
        return self.data


# =========================================================
# QUERY CHECKS / HELPERS
# =========================================================
_READ_KEYWORDS = {"SELECT", "WITH", "EXEC", "EXECUTE"}
_WRITE_KEYWORDS = {"INSERT", "UPDATE", "DELETE", "MERGE"}

_LEADING_COMMENT_RE = re.compile(r"^\s*(--[^\n]*\n|/\*.*?\*/\s*)", re.DOTALL)
_FIRST_TOKEN_RE = re.compile(r"[A-Za-z_]+")


def _leading_keyword(query: str) -> str:
    """Return the first SQL keyword (uppercased), skipping leading comments/whitespace."""
    q = query or ""
    while True:
        m = _LEADING_COMMENT_RE.match(q)
        if not m:
            break
        q = q[m.end():]
    q = q.lstrip()
    m = _FIRST_TOKEN_RE.match(q)
    return m.group(0).upper() if m else ""


def _normalize_params(params: Any) -> tuple:
    """
    Coerce whatever the caller passed into the sequence pyodbc expects.

      None                 -> ()
      list / tuple         -> tuple(params)
      scalar (incl. str)   -> (params,)
    """
    if params is None:
        return ()
    if isinstance(params, (list, tuple)):
        return tuple(params)
    return (params,)


def build_in_clause(column: str, values: Sequence[Any]) -> Tuple[str, list]:
    """
    Build a safe, dynamic `col IN (?, ?, ...)` clause.

    Returns (clause_string, params_list). With an empty `values` it returns a
    clause that matches nothing ("1 = 0") so callers never emit invalid SQL.

        clause, params = build_in_clause("ProcessingId", [1, 2, 3])
        db.execute(f"UPDATE TProcessStatus SET SourcingFlag='s' WHERE {clause}", params)
    """
    vals = list(values or [])
    if not vals:
        return "1 = 0", []
    placeholders = ", ".join(["?"] * len(vals))
    return f"{column} IN ({placeholders})", vals


# =========================================================
# DATABASE
# =========================================================
class Database:
    """
    Reusable database handle. Instantiate once and reuse across many queries.

    Parameters
    ----------
    conn : optional existing pyodbc connection to reuse. If None, a connection
           is opened lazily on first use via get_db_connection() and is owned
           (and closed) by this instance.
    autocommit_writes : if True (default), write methods commit after a
           successful statement. Set False to batch several writes and call
           commit() yourself.
    strict_checks : if True (default), reject obviously misused queries
           (e.g. a SELECT passed to execute(), or a write passed to fetch_*).

    Usable as a context manager:

        with Database() as db:
            db.execute("UPDATE ...", [...])
        # commits on clean exit, rolls back on exception, then closes.
    """

    def __init__(self, conn: Optional["pyodbc.Connection"] = None, *,
                 autocommit_writes: bool = True, strict_checks: bool = True):
        self._conn = conn
        self._owns_conn = conn is None      # only close what we opened ourselves
        self.autocommit_writes = autocommit_writes
        self.strict_checks = strict_checks

    # ---- connection lifecycle -------------------------------------------------
    def _ensure_conn(self) -> "pyodbc.Connection":
        if self._conn is None:
            self._conn = get_db_connection()
            self._owns_conn = True
            logger.debug("Opened new database connection.")
        return self._conn

    @property
    def connection(self) -> "pyodbc.Connection":
        """The underlying pyodbc connection (opening it if needed)."""
        return self._ensure_conn()

    def commit(self) -> DBResult:
        try:
            if self._conn is not None:
                self._conn.commit()
            return DBResult(success=True)
        except Exception as e:  # noqa: BLE001
            logger.error("Commit failed: %s", e)
            return DBResult(success=False, error=str(e))

    def rollback(self) -> DBResult:
        try:
            if self._conn is not None:
                self._conn.rollback()
            return DBResult(success=True)
        except Exception as e:  # noqa: BLE001
            logger.error("Rollback failed: %s", e)
            return DBResult(success=False, error=str(e))

    def close(self) -> None:
        """Close the connection if this instance owns it."""
        if self._conn is not None and self._owns_conn:
            try:
                self._conn.close()
            except Exception as e:  # noqa: BLE001
                logger.warning("Error while closing connection: %s", e)
        self._conn = None

    def __enter__(self) -> "Database":
        self._ensure_conn()
        return self

    def __exit__(self, exc_type, exc_val, exc_tb) -> bool:
        if self._conn is not None:
            if exc_type is None:
                self.commit()
            else:
                self.rollback()
        self.close()
        return False  # never suppress exceptions

    # ---- internal validation --------------------------------------------------
    def _check_query(self, query: str, mode: str, require: Optional[str] = None) -> Optional[str]:
        """
        Validate a query for the given mode. Returns an error string if the
        query should be rejected, otherwise None.

          mode="write" : reject reads (SELECT/WITH).
          mode="read"  : reject plain writes (INSERT/UPDATE/DELETE/MERGE).
          require      : exact leading keyword required (used by insert/update/delete).
        """
        if not query or not str(query).strip():
            return "Empty query."

        if not self.strict_checks:
            return None

        kw = _leading_keyword(query)

        if require is not None and kw != require:
            return f"Expected a {require} statement but query starts with '{kw or '?'}'."

        if mode == "write" and kw in _READ_KEYWORDS:
            return f"execute() is for writes; '{kw}' looks like a query — use fetch_one()/fetch_all()."

        if mode == "read" and kw in _WRITE_KEYWORDS:
            return f"fetch_*() is for queries; '{kw}' is a write — use execute()/insert()/update()/delete()."

        return None

    # ---- non-query (add / update / delete) ------------------------------------
    def execute(self, query: str, params: Any = None, *,
                commit: Optional[bool] = None, _require: Optional[str] = None) -> DBResult:
        """
        Run a non-query statement (INSERT / UPDATE / DELETE / MERGE / DDL).
        Returns DBResult(success, data=None, rowcount=<affected>).
        """
        err = self._check_query(query, mode="write", require=_require)
        if err:
            logger.error("Query rejected: %s", err)
            return DBResult(success=False, error=err)

        do_commit = self.autocommit_writes if commit is None else commit
        bound = _normalize_params(params)
        cur = None
        try:
            conn = self._ensure_conn()
            cur = conn.cursor()
            cur.execute(query, bound)
            affected = cur.rowcount
            if do_commit:
                conn.commit()
            return DBResult(success=True, data=None, rowcount=affected)
        except Exception as e:  # noqa: BLE001
            logger.error("execute() failed: %s | SQL: %s", e, _snippet(query))
            self.rollback()
            return DBResult(success=False, error=str(e))
        finally:
            _safe_close_cursor(cur)

    def execute_many(self, query: str, seq_of_params: Sequence[Sequence[Any]], *,
                     commit: Optional[bool] = None) -> DBResult:
        """
        Run one statement against many parameter sets via cursor.executemany().
        Returns DBResult(success, data=None, rowcount=<affected or count>).
        """
        err = self._check_query(query, mode="write")
        if err:
            logger.error("Query rejected: %s", err)
            return DBResult(success=False, error=err)

        do_commit = self.autocommit_writes if commit is None else commit
        rows = [_normalize_params(p) for p in (seq_of_params or [])]
        if not rows:
            return DBResult(success=True, data=None, rowcount=0)

        cur = None
        try:
            conn = self._ensure_conn()
            cur = conn.cursor()
            cur.executemany(query, rows)
            affected = cur.rowcount
            if do_commit:
                conn.commit()
            # Some drivers report -1 for executemany; fall back to the batch size.
            if affected is None or affected < 0:
                affected = len(rows)
            return DBResult(success=True, data=None, rowcount=affected)
        except Exception as e:  # noqa: BLE001
            logger.error("execute_many() failed: %s | SQL: %s", e, _snippet(query))
            self.rollback()
            return DBResult(success=False, error=str(e))
        finally:
            _safe_close_cursor(cur)

    # Semantic aliases matching the "add / update / delete" wording -------------
    def insert(self, query: str, params: Any = None, *, commit: Optional[bool] = None) -> DBResult:
        return self.execute(query, params, commit=commit, _require="INSERT")

    def update(self, query: str, params: Any = None, *, commit: Optional[bool] = None) -> DBResult:
        return self.execute(query, params, commit=commit, _require="UPDATE")

    def delete(self, query: str, params: Any = None, *, commit: Optional[bool] = None) -> DBResult:
        return self.execute(query, params, commit=commit, _require="DELETE")

    # ---- queries (read) -------------------------------------------------------
    def fetch_one(self, query: str, params: Any = None) -> DBResult:
        """
        Run a query and return a single row.
        Returns DBResult(success, data=<pyodbc.Row | None>, rowcount=0/1).
        Rows keep native attribute access (row.ColumnName).
        """
        err = self._check_query(query, mode="read")
        if err:
            logger.error("Query rejected: %s", err)
            return DBResult(success=False, error=err)

        bound = _normalize_params(params)
        cur = None
        try:
            conn = self._ensure_conn()
            cur = conn.cursor()
            cur.execute(query, bound)
            row = cur.fetchone()
            return DBResult(success=True, data=row, rowcount=(1 if row is not None else 0))
        except Exception as e:  # noqa: BLE001
            logger.error("fetch_one() failed: %s | SQL: %s", e, _snippet(query))
            return DBResult(success=False, error=str(e))
        finally:
            _safe_close_cursor(cur)

    def fetch_all(self, query: str, params: Any = None) -> DBResult:
        """
        Run a query and return all rows.
        Returns DBResult(success, data=<list[pyodbc.Row]>, rowcount=<len>).
        """
        err = self._check_query(query, mode="read")
        if err:
            logger.error("Query rejected: %s", err)
            return DBResult(success=False, error=err)

        bound = _normalize_params(params)
        cur = None
        try:
            conn = self._ensure_conn()
            cur = conn.cursor()
            cur.execute(query, bound)
            rows = cur.fetchall()
            return DBResult(success=True, data=list(rows), rowcount=len(rows))
        except Exception as e:  # noqa: BLE001
            logger.error("fetch_all() failed: %s | SQL: %s", e, _snippet(query))
            return DBResult(success=False, error=str(e))
        finally:
            _safe_close_cursor(cur)

    # Convenience passthrough for the IN-clause builder -------------------------
    build_in_clause = staticmethod(build_in_clause)


# =========================================================
# SMALL INTERNAL HELPERS
# =========================================================
def _safe_close_cursor(cur) -> None:
    if cur is not None:
        try:
            cur.close()
        except Exception:  # noqa: BLE001
            pass


def _snippet(query: str, limit: int = 200) -> str:
    """Compact one-line SQL snippet for log messages."""
    s = re.sub(r"\s+", " ", (query or "").strip())
    return s if len(s) <= limit else s[:limit] + "…"


# =========================================================
# OPTIONAL MODULE-LEVEL CONVENIENCE
# =========================================================
# For callers who just want `import db; db.fetch_all(sql)` without managing an
# instance. Backed by one lazily-created, process-wide Database. The class API
# above remains the primary, documented interface.
_default_db: Optional[Database] = None


def get_default_db() -> Database:
    """Return the shared module-level Database, creating it on first use."""
    global _default_db
    if _default_db is None:
        _default_db = Database()
    return _default_db


def close_default_db() -> None:
    """Close and drop the shared module-level Database (if any)."""
    global _default_db
    if _default_db is not None:
        _default_db.close()
        _default_db = None


def execute(query: str, params: Any = None, *, commit: Optional[bool] = None) -> DBResult:
    return get_default_db().execute(query, params, commit=commit)


def execute_many(query: str, seq_of_params: Sequence[Sequence[Any]], *,
                 commit: Optional[bool] = None) -> DBResult:
    return get_default_db().execute_many(query, seq_of_params, commit=commit)


def insert(query: str, params: Any = None, *, commit: Optional[bool] = None) -> DBResult:
    return get_default_db().insert(query, params, commit=commit)


def update(query: str, params: Any = None, *, commit: Optional[bool] = None) -> DBResult:
    return get_default_db().update(query, params, commit=commit)


def delete(query: str, params: Any = None, *, commit: Optional[bool] = None) -> DBResult:
    return get_default_db().delete(query, params, commit=commit)


def fetch_one(query: str, params: Any = None) -> DBResult:
    return get_default_db().fetch_one(query, params)


def fetch_all(query: str, params: Any = None) -> DBResult:
    return get_default_db().fetch_all(query, params)


# =========================================================
#  TEST
# =========================================================
if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO, format="%(levelname)s %(name)s: %(message)s")

    print("[db] Running connectivity smoke test...")
    try:
        with Database() as _db:
            r = _db.fetch_one("SELECT 1 AS ok")
            print(f"[db] fetch_one -> success={r.success}, "
                  f"ok={getattr(r.data, 'ok', None)}, error={r.error}")

            # Validation checks should be rejected cleanly (no exception raised):
            bad_empty = _db.execute("   ")
            print(f"[db] empty query   -> success={bad_empty.success}, error={bad_empty.error}")

            wrong_mode = _db.fetch_all("UPDATE some_table SET col='x' WHERE 1=0")
            print(f"[db] write via read -> success={wrong_mode.success}, error={wrong_mode.error}")
    except Exception as exc:  # noqa: BLE001
        print(f"[db] Smoke test could not connect: {exc}")
