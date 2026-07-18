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
"""

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

# Spaces per indent level in the output text.
SPACES_PER_LEVEL = 2


# ─────────────────────────────────────────────────────────────────────────────
# CORE FUNCTION
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

    # ── 1. Group words into rows by vertical position ─────────────────────
    rows: dict[float, list] = defaultdict(list)
    for w in words:
        top = round(w["top"], 1)
        rows[top].append(w)

    # ── 2. Discover x0 buckets from label-column words ────────────────────
    raw_x0s: list[float] = []
    for top_key in rows:
        label_ws = [w for w in rows[top_key] if w["x0"] < label_max_x]
        if label_ws:
            raw_x0s.append(round(min(w["x0"] for w in label_ws), 2))

    if not raw_x0s:
        return page.extract_text() or ""

    # Cluster raw x0 values into discrete buckets
    sorted_x0s = sorted(set(raw_x0s))
    buckets: list[float] = []
    for x in sorted_x0s:
        if not buckets or x - buckets[-1] > X0_TOLERANCE:
            buckets.append(x)

    def indent_for_x0(x0: float) -> str:
        closest = min(buckets, key=lambda b: abs(b - x0))
        level = buckets.index(closest)
        return " " * (level * SPACES_PER_LEVEL)

    # ── 3. Build output lines ──────────────────────────────────────────────
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
            x0 = label_ws[0]["x0"]
            indent = indent_for_x0(x0)
            label = " ".join(w["text"] for w in label_ws)
            values = "  " + " ".join(w["text"] for w in value_ws) if value_ws else ""
            lines.append(f"{indent}{label}{values}")
        elif value_ws:
            # Values-only row (e.g. continuation of a wrapped label)
            values = "  " + " ".join(w["text"] for w in value_ws)
            lines.append(values)

    return "\n".join(lines)


def pdf_to_indented_text(
    pdf_path: str,
    per_page: bool = False,
    label_max_x: float = LABEL_MAX_X,
    page_separator: str = "\n\n",
) -> str | list[str]:
    """
    Extract all pages of a PDF as indentation-preserving plain text.

    Args:
        pdf_path:      Path to the PDF file.
        per_page:      If True, return a list of strings (one per page).
                       If False (default), return a single joined string.
        label_max_x:   x0 threshold (points) separating label from value columns.
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