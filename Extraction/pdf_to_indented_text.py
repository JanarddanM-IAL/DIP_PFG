"""
pdf_to_indented_text.py
=======================
Replaces pdfplumber's extract_text() for financial statement PDFs.

Problem with extract_text():
  It strips ALL indentation — every row label comes out at 0 leading spaces,
  even when the PDF visually shows 2-3 levels of hierarchy. This causes LLMs
  to guess hierarchy from label semantics, which fails for ambiguous cases like
  "Investments with fiscal agent" (looks like a peer of "Cash and investments"
   but is actually a child of it at one deeper indent level).

Solution:
  1. Extract words with their exact x0 PDF coordinates.
  2. Cluster x0 values into discrete indent levels.
  3. Map each level → N leading spaces (2 per level).
  4. Reconstruct each row as: "<spaces><label>  <numeric values>"

This preserves the true visual hierarchy and gives the LLM unambiguous
indentation signals identical to what a human reader sees on the page.

Usage:
  from pdf_to_indented_text import pdf_to_indented_text

  text = pdf_to_indented_text("statement.pdf")          # all pages joined
  pages = pdf_to_indented_text("statement.pdf", per_page=True)  # list of strings

  # Zero-copy path (pipeline.py) — open PDF once, share words with
  # coordinate_extractor:
  from pdf_to_indented_text import extract_all_pages_words, pdf_to_indented_text_from_words
  all_pages_words = extract_all_pages_words("statement.pdf")
  text = pdf_to_indented_text_from_words(all_pages_words)
"""

import bisect
import re
from collections import defaultdict
from pathlib import Path

import pdfplumber


# ─────────────────────────────────────────────────────────────────────────────
# CONSTANTS
# ─────────────────────────────────────────────────────────────────────────────

# Maximum x0 (points) for a word to be considered part of the row label.
# Words to the right of this threshold are treated as numeric column values.
# 220 pt ≈ 3.1 inches from the left edge, which covers all label text in
# standard government-wide financial statements (landscape or portrait).
LABEL_MAX_X = 220

# Tolerance (points) for clustering x0 values into the same indent level.
# Words within X0_TOLERANCE of each other share the same indent bucket.
X0_TOLERANCE = 1.5
MIN_LEVEL_GAP = 10.0

# Spaces per indent level in the output text.
SPACES_PER_LEVEL = 2

# Shared numeric-token pattern used by all extraction helpers.
_NUM_PAT = re.compile(r'^\(?\d[\d,]*\)?$')


# ─────────────────────────────────────────────────────────────────────────────
# INDENT-LEVEL CLUSTERING
# ─────────────────────────────────────────────────────────────────────────────

def cluster_x0s(raw_x0s: list[float], min_level_gap: float = MIN_LEVEL_GAP) -> list[float]:
    """
    Cluster raw x0 values into discrete indent-level buckets.

    Splits on any gap >= min_level_gap between consecutive unique x0 values.

    The previous median*2 formula was intended to suppress jitter but had a
    fatal flaw: when only 2 indent levels exist, the single gap IS the median,
    making threshold = 2*gap — a value the gap can never reach. This collapsed
    every 2-level hierarchy (e.g., "Restricted assets" + children) into one
    bucket, destroying indentation signals.

    Fix: any gap >= min_level_gap is a real indent boundary. Jitter (sub-pixel
    x0 variation between same-level rows) is always < 2pt; MIN_LEVEL_GAP=10
    safely separates real levels from noise.
    """
    if not raw_x0s:
        return []

    sorted_x0s = sorted(set(round(x, 2) for x in raw_x0s))
    if len(sorted_x0s) == 1:
        return [sorted_x0s[0]]

    gaps = [sorted_x0s[i + 1] - sorted_x0s[i] for i in range(len(sorted_x0s) - 1)]

    # Any gap >= min_level_gap is a true indent-level boundary.
    # Sub-pixel jitter between same-level rows is always < 2pt.
    # MIN_LEVEL_GAP=10 gives an 8pt safety margin above jitter noise.
    split_after = [i for i, g in enumerate(gaps) if g >= min_level_gap]

    buckets = []
    start = 0
    for split_i in split_after:
        buckets.append(sorted_x0s[start: split_i + 1])
        start = split_i + 1
    buckets.append(sorted_x0s[start:])

    return [sum(b) / len(b) for b in buckets]


# ─────────────────────────────────────────────────────────────────────────────
# SHARED LINE-BUILDING LOGIC
# ─────────────────────────────────────────────────────────────────────────────

def _build_lines_from_rows(
    rows: dict[float, list],
    label_max_x: float,
) -> str:
    """
    Core line-builder shared by both the Page-object path and the
    pre-extracted-words path.

    rows  — dict keyed by rounded `top` value; values are lists of word dicts
            (each with keys: text, x0, x1, top, bottom).
    label_max_x — x0 threshold separating label words from numeric values.

    Returns a single string with newline-separated indented rows.
    """

    def _row_has_numeric_values(row_ws: list) -> bool:
        return any(
            _NUM_PAT.match(w["text"].replace("$", "").strip())
            for w in row_ws
            if w["x0"] >= label_max_x
        )

    # ── Collect tops of rows that have numeric values ─────────────────────
    numeric_tops = set()
    for top_key, row_ws in rows.items():
        if _row_has_numeric_values(row_ws):
            numeric_tops.add(top_key)

    numeric_tops_sorted = sorted(numeric_tops)

    def _is_near_numeric_row(top_key: float, tolerance: float = 20.0) -> bool:
        idx = bisect.bisect_left(numeric_tops_sorted, top_key - tolerance)
        return (
            idx < len(numeric_tops_sorted)
            and numeric_tops_sorted[idx] <= top_key + tolerance
        )

    # ── Collect x0 values for indent-level clustering ─────────────────────
    raw_x0s: list[float] = []
    for top_key in rows:
        label_ws = [w for w in rows[top_key] if w["x0"] < label_max_x]
        if label_ws and (
            _row_has_numeric_values(rows[top_key])
            or _is_near_numeric_row(top_key)
        ):
            raw_x0s.append(round(min(w["x0"] for w in label_ws), 2))

    if not raw_x0s:
        # No structured table rows found — fall back to flat word sequence
        all_words = [w for row_ws in rows.values() for w in row_ws]
        return " ".join(
            w["text"] for w in sorted(all_words, key=lambda w: (w["top"], w["x0"]))
        )

    buckets = cluster_x0s(raw_x0s, min_level_gap=MIN_LEVEL_GAP)

    def indent_for_x0(x0: float) -> str:
        closest_idx = min(range(len(buckets)), key=lambda i: abs(buckets[i] - x0))
        return " " * (closest_idx * SPACES_PER_LEVEL)

    # ── Build output lines ────────────────────────────────────────────────
    lines: list[str] = []
    for top_key in sorted(rows.keys()):
        row_ws = rows[top_key]
        label_ws = sorted(
            [w for w in row_ws if w["x0"] < label_max_x], key=lambda w: w["x0"]
        )
        value_ws = sorted(
            [w for w in row_ws if w["x0"] >= label_max_x], key=lambda w: w["x0"]
        )
        if label_ws:
            x0     = label_ws[0]["x0"]
            indent = indent_for_x0(x0)
            label  = " ".join(w["text"] for w in label_ws)
            values = "  " + " ".join(w["text"] for w in value_ws) if value_ws else ""
            lines.append(f"{indent}{label}{values}")
        elif value_ws:
            lines.append("  " + " ".join(w["text"] for w in value_ws))

    return "\n".join(lines)


# ─────────────────────────────────────────────────────────────────────────────
# PATH A — pdfplumber Page object  (used by pdf_to_indented_text)
# ─────────────────────────────────────────────────────────────────────────────

def extract_page_indented(page, label_max_x: float = LABEL_MAX_X) -> str:
    """
    Convert one pdfplumber Page into indented plain text.

    Each row is output as:
        <indent><label text>  <numeric/value text>

    Indent is 2 spaces × indent_level, where indent_level is determined by
    the x0 position of the first character of the row label.
    """
    words = page.extract_words(x_tolerance=3, y_tolerance=3)
    if not words:
        return ""

    rows: dict[float, list] = defaultdict(list)
    for w in words:
        top = round(w["top"], 1)
        rows[top].append(w)

    return _build_lines_from_rows(rows, label_max_x)


# ─────────────────────────────────────────────────────────────────────────────
# PATH B — pre-extracted word list  (used by pdf_to_indented_text_from_words)
# ─────────────────────────────────────────────────────────────────────────────

def _extract_page_indented_from_words(
    words: list[dict],
    label_max_x: float = LABEL_MAX_X,
) -> str:
    """
    Same logic as extract_page_indented() but accepts a pre-extracted
    word list instead of a pdfplumber Page object.

    This is the zero-copy path: pipeline.py calls extract_all_pages_words()
    once, then passes the result to both this function (for LLM context) and
    attach_coordinates_from_words() (for coordinate injection), avoiding a
    second pdfplumber open.
    """
    if not words:
        return ""

    rows: dict[float, list] = defaultdict(list)
    for w in words:
        top = round(w["top"], 1)
        rows[top].append(w)

    return _build_lines_from_rows(rows, label_max_x)


# ─────────────────────────────────────────────────────────────────────────────
# SINGLE-OPEN HELPERS  (zero-copy pipeline path)
# ─────────────────────────────────────────────────────────────────────────────

def extract_all_pages_words(pdf_path: str) -> list[list[dict]]:
    """
    Open PDF once and extract all words from all pages.
    Returns a list of per-page word lists.
    Each word dict has: text, x0, x1, top, bottom, page (1-based).

    This is the SINGLE pdfplumber open call for the entire pipeline.
    Pass the result to both pdf_to_indented_text_from_words() and
    coordinate_extractor.attach_coordinates_from_words().
    """
    path = Path(pdf_path)
    if not path.exists():
        raise FileNotFoundError(f"PDF not found: {pdf_path}")

    all_pages_words: list[list[dict]] = []
    with pdfplumber.open(str(path)) as pdf:
        for page_idx, page in enumerate(pdf.pages, start=1):
            words = page.extract_words(x_tolerance=3, y_tolerance=3)
            # Attach 1-based page number to each word for coordinate_extractor
            for w in words:
                w["page"] = page_idx
            all_pages_words.append(words)

    return all_pages_words


def pdf_to_indented_text_from_words(
    all_pages_words: list[list[dict]],
    label_max_x: float = LABEL_MAX_X,
    page_separator: str = "\n\n",
) -> str:
    """
    Build indented text from a pre-extracted word list (output of
    extract_all_pages_words). Avoids a second pdfplumber open.

    Args:
        all_pages_words: list of per-page word lists from extract_all_pages_words()
        label_max_x:     x0 threshold (points) separating label from value columns
        page_separator:  string inserted between pages in the joined output

    Returns:
        Single joined string with indentation-preserving text.
    """
    page_texts: list[str] = []
    for page_words in all_pages_words:
        page_texts.append(
            _extract_page_indented_from_words(page_words, label_max_x=label_max_x)
        )
    return page_separator.join(page_texts)


# ─────────────────────────────────────────────────────────────────────────────
# LEGACY SINGLE-CALL PATH  (used by fallback paths and standalone tools)
# ─────────────────────────────────────────────────────────────────────────────

def pdf_to_indented_text(
    pdf_path: str,
    per_page: bool = False,
    label_max_x: float = LABEL_MAX_X,
    page_separator: str = "\n\n",
) -> str | list[str]:
    """
    Extract all pages of a PDF as indentation-preserving plain text.

    Args:
        pdf_path:       Path to the PDF file.
        per_page:       If True, return a list of strings (one per page).
                        If False (default), return a single joined string.
        label_max_x:    x0 threshold (points) separating label from value columns.
        page_separator: String inserted between pages when per_page=False.

    Returns:
        str or list[str] depending on per_page.
    """
    path = Path(pdf_path)
    if not path.exists():
        raise FileNotFoundError(f"PDF not found: {pdf_path}")

    page_texts: list[str] = []
    with pdfplumber.open(str(path)) as pdf:
        for page in pdf.pages:
            page_texts.append(extract_page_indented(page, label_max_x=label_max_x))

    if per_page:
        return page_texts
    return page_separator.join(page_texts)


# ─────────────────────────────────────────────────────────────────────────────
# CLI  (python pdf_to_indented_text.py statement.pdf)
# ─────────────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    import sys

    if len(sys.argv) < 2:
        print("Usage: python pdf_to_indented_text.py <pdf_path> [label_max_x]")
        sys.exit(1)

    pdf_path = sys.argv[1]
    label_x = float(sys.argv[2]) if len(sys.argv) > 2 else LABEL_MAX_X
    result = pdf_to_indented_text(pdf_path, label_max_x=label_x)
    print(result)