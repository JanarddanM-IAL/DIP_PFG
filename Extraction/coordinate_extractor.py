"""
coordinate_extractor.py
=======================
POST-PROCESSING MODULE — No prompt changes required.

After the LLM produces its JSON output (SNP, SOA, GOV_BS, etc.),
this module:
  1. Reads Metadata.Page No from the JSON to learn which pages of the
     ORIGINAL full PDF the sliced PDF corresponds to.
  2. Extracts ALL numeric words from the sliced PDF with x0,y0,x1,y1,page.
  3. Translates the sliced-PDF page numbers → original full-PDF page numbers.
  4. Injects a "_coord" sibling key for every numeric value in every row.

Example — Metadata.Page No = "71,72,73"
  Sliced PDF page 1 → original page 71
  Sliced PDF page 2 → original page 72
  Sliced PDF page 3 → original page 73

  Before:
    { "Governmental Activities": "12,772,921" }

  After:
    {
      "Governmental Activities": "12,772,921",
      "Governmental Activities_coord": {
          "page": 71,          ← original full-PDF page
          "x0": 312.5, "y0": 187.4,
          "x1": 368.2, "y1": 198.6
      }
    }

Usage:
  from coordinate_extractor import attach_coordinates
  enriched = attach_coordinates(json_data, pdf_path)
"""

from __future__ import annotations

import re
import copy
from pathlib import Path
from typing import Any

import pdfplumber

# ── Constants ──────────────────────────────────────────────────────────────────

_NUM_RE = re.compile(r"^\(?-?[\d,]+\)?$")

_Y_TOLERANCE = 6.0    # PDF points — words within this share the same visual row
_X_TOLERANCE = 20.0   # PDF points — horizontal tie-breaking tolerance

# ── All possible top-level keys that hold the row-list sections ────────────────

_ITEMS_KEYS = {
    "Row Items",
}

_META_KEYS = {
    "Metadata", "Reporting Columns", "Total Check Status",
    "COA Flag", "COA Datapoint",
}

# ── Page-offset helpers ────────────────────────────────────────────────────────

def _parse_page_no(page_no_value: Any) -> list[int]:
    """
    Parse Metadata['Page No'] into a sorted list of original-PDF page numbers.

    Handles all formats the LLM produces:
      "71"           → [71]
      "71,72,73"     → [71, 72, 73]
      "71, 72, 73"   → [71, 72, 73]
      71             → [71]
      [71, 72]       → [71, 72]
    """
    if page_no_value is None:
        return []

    if isinstance(page_no_value, int):
        return [page_no_value]

    if isinstance(page_no_value, list):
        result = []
        for v in page_no_value:
            try:
                result.append(int(str(v).strip()))
            except (ValueError, TypeError):
                pass
        return sorted(result)

    # String — split on commas, strip spaces
    parts = str(page_no_value).split(",")
    result = []
    for p in parts:
        p = p.strip()
        if p.isdigit():
            result.append(int(p))
    return sorted(result)


def _build_page_map(original_pages: list[int]) -> dict[int, int]:
    """
    Build a mapping:  sliced_pdf_page (1-based) → original_pdf_page

    If original_pages = [71, 72, 73]:
      { 1: 71, 2: 72, 3: 73 }

    If original_pages is empty or has one entry:
      { 1: original_pages[0] }   (or empty dict → caller falls back to identity)
    """
    if not original_pages:
        return {}
    return {sliced_p: orig_p
            for sliced_p, orig_p in enumerate(original_pages, start=1)}


def _translate_page(sliced_page: int, page_map: dict[int, int]) -> int:
    """
    Translate a sliced-PDF page number to the original full-PDF page number.
    Falls back to the sliced page number if the map doesn't cover it
    (e.g. the sliced PDF has more pages than Metadata.Page No lists —
    shouldn't happen normally, but safe fallback).
    """
    if not page_map:
        return sliced_page
    if sliced_page in page_map:
        return page_map[sliced_page]
    # Extrapolate: if sliced page > max mapped, extend linearly from last entry
    max_sliced = max(page_map.keys())
    max_orig   = page_map[max_sliced]
    offset     = sliced_page - max_sliced
    return max_orig + offset


# ── Normalisation helpers ──────────────────────────────────────────────────────

def _normalise(value: str) -> str:
    """Strip formatting: '(1,234,567)' → '1234567'."""
    return re.sub(r"[,\(\)\-\s]", "", value)


def _is_numeric_token(value: Any) -> bool:
    """Return True if value is a non-empty numeric cell string."""
    if not isinstance(value, str):
        return False
    stripped = value.strip()
    if stripped in ("-", "–", "—", ""):
        return False
    return bool(_NUM_RE.match(stripped))


# ── PDF word extraction ────────────────────────────────────────────────────────

def _extract_numeric_words(pdf_path: str) -> list[dict]:
    """
    Use pdfplumber to extract every word that looks like a number.
    page numbers here are 1-based SLICED-PDF page numbers.
    They are translated to original-PDF page numbers later in attach_coordinates().
    """
    words = []
    with pdfplumber.open(pdf_path) as pdf:
        for page_idx, page in enumerate(pdf.pages, start=1):
            for w in page.extract_words(x_tolerance=3, y_tolerance=3):
                t = w["text"].strip()
                clean = re.sub(r"[,\(\)\-\s]", "", t)
                if clean.isdigit() and len(clean) >= 1:
                    words.append({
                        "page": page_idx,   # sliced-PDF page (translated later)
                        "text": t,
                        "norm": clean,
                        "x0":   round(w["x0"],    2),
                        "y0":   round(w["top"],    2),
                        "x1":   round(w["x1"],    2),
                        "y1":   round(w["bottom"], 2),
                        "used": False,
                    })
    return words


# ── Row-grouping helper ────────────────────────────────────────────────────────

def _group_words_into_rows(words: list[dict]) -> list[list[dict]]:
    """Group words sharing the same visual row (by page + y0 proximity)."""
    if not words:
        return []

    sorted_words = sorted(words, key=lambda w: (w["page"], w["y0"], w["x0"]))
    rows: list[list[dict]] = []
    current_row:   list[dict] = [sorted_words[0]]
    current_avg_y  = sorted_words[0]["y0"]
    current_page   = sorted_words[0]["page"]

    for w in sorted_words[1:]:
        same_row = (
            w["page"] == current_page
            and abs(w["y0"] - current_avg_y) <= _Y_TOLERANCE
        )
        if same_row:
            current_row.append(w)
            current_avg_y = sum(x["y0"] for x in current_row) / len(current_row)
        else:
            rows.append(sorted(current_row, key=lambda x: x["x0"]))
            current_row   = [w]
            current_avg_y = w["y0"]
            current_page  = w["page"]

    rows.append(sorted(current_row, key=lambda x: x["x0"]))
    return rows


# ── Coordinate index ───────────────────────────────────────────────────────────

class _CoordIndex:
    """Searchable index of all numeric PDF words."""

    def __init__(self, pdf_path: str = None, all_pages_words: list = None):
        """
        Accept EITHER a pdf_path (opens pdfplumber internally)
        OR pre-extracted all_pages_words (no pdfplumber open needed).
        """
        if all_pages_words is not None:
            # Words already extracted — just filter for numerics
            raw = self._words_from_preextracted(all_pages_words)
        elif pdf_path is not None:
            raw = _extract_numeric_words(pdf_path)
        else:
            raise ValueError("Either pdf_path or all_pages_words must be provided")

        self._rows = _group_words_into_rows(raw)
        self._flat = raw

    @staticmethod
    def _words_from_preextracted(all_pages_words: list) -> list[dict]:
        """
        Filter pre-extracted words to numeric tokens only.
        Mirrors _extract_numeric_words() but without opening pdfplumber.
        """
        words = []
        for page_words in all_pages_words:
            for w in page_words:
                t     = w["text"].strip()
                clean = re.sub(r"[,\(\)\-\s]", "", t)
                if clean.isdigit() and len(clean) >= 1:
                    words.append({
                        "page": w["page"],
                        "text": t,
                        "norm": clean,
                        "x0":   round(w["x0"],    2),
                        "y0":   round(w["top"],    2),
                        "x1":   round(w["x1"],    2),
                        "y1":   round(w["bottom"], 2),
                        "used": False,
                    })
        return words
    def find(
        self,
        norm_value: str,
        col_index: int,
        num_cols: int,
        row_label: str = "",
    ) -> dict | None:
        """
        Find the best-matching PDF word for a given normalised numeric value.

        Strategy:
        1. Filter all unused words whose norm matches norm_value.
        2. Among matches, prefer words whose horizontal x0 position falls
            in the expected column zone (col_index / num_cols of page width).
        3. Return the best match's coord dict and mark it as used.
        4. Return None if no match found.
        """
        candidates = [w for w in self._flat if w["norm"] == norm_value and not w["used"]]

        if not candidates:
            return None

        # ── Score each candidate by column-zone proximity ─────────────────────
        # Estimate expected x-center for this column based on page width (~612pt)
        PAGE_WIDTH   = 612.0
        col_fraction = (col_index + 0.5) / num_cols   # 0.0 – 1.0
        expected_x   = PAGE_WIDTH * col_fraction

        # Pick candidate whose x0 is closest to expected_x
        best = min(candidates, key=lambda w: abs(w["x0"] - expected_x))

        best["used"] = True

        return {
            "page": best["page"],
            "x0":   best["x0"],
            "y0":   best["y0"],
            "x1":   best["x1"],
            "y1":   best["y1"],
        }

def attach_coordinates_from_words(
    json_data: dict,
    all_pages_words: list,
) -> dict:
    """
    Same as attach_coordinates() but accepts pre-extracted words
    instead of a pdf_path — no pdfplumber open needed.

    Args:
        json_data       : Parsed LLM output dict.
        all_pages_words : Output of extract_all_pages_words() from
                          pdf_to_indented_text.py — list of per-page
                          word lists, each word having 'page' key.

    Returns:
        Enriched dict with _coord keys injected.
    """
    if not json_data or not all_pages_words:
        return json_data

    result = copy.deepcopy(json_data)

    reporting_columns = _discover_reporting_columns(result)
    if not reporting_columns:
        return result

    raw_page_no    = _discover_page_no(result)
    original_pages = _parse_page_no(raw_page_no)
    page_map       = _build_page_map(original_pages)

    try:
        coord_index = _CoordIndex(all_pages_words=all_pages_words)
    except Exception:
        return result

    section_lists = _collect_section_lists(result)
    if not section_lists:
        return result

    for parent_container, row_list, key in section_lists:
        enriched_rows = []
        for row in row_list:
            if isinstance(row, dict):
                enriched_rows.append(
                    _enrich_row(row, reporting_columns, coord_index, page_map)
                )
            else:
                enriched_rows.append(row)
        parent_container[key] = enriched_rows

    return result

# ── Row enrichment ─────────────────────────────────────────────────────────────

def _enrich_row(
    row: dict,
    reporting_columns: list[str],
    coord_index: _CoordIndex,
    page_map: dict[int, int],
) -> dict:
    """
    Inject _coord keys for every numeric value in a data row dict.
    Translates sliced-PDF page numbers → original-PDF page numbers.
    """
    n        = len(reporting_columns)
    enriched = dict(row)

    for col_idx, col_name in enumerate(reporting_columns):
        value = row.get(col_name)
        if not _is_numeric_token(value):
            continue
        norm  = _normalise(value)
        coord = coord_index.find(
            norm_value=norm,
            col_index=col_idx,
            num_cols=n,
            row_label=str(row.get(list(row.keys())[0], "")),
        )
        if coord:
            # ── Translate page number ─────────────────────────────────────
            translated_coord = dict(coord)
            translated_coord["page"] = _translate_page(coord["page"], page_map)
            enriched[f"{col_name}_coord"] = translated_coord

    return enriched


# ── Section / structure discovery ─────────────────────────────────────────────

def _discover_reporting_columns(json_data: dict) -> list[str]:
    """
    Find reporting columns robustly regardless of nesting depth.
    Checks top-level first, then digs into nested structures.
    """
    cols = json_data.get("Reporting Columns") or json_data.get("reporting_columns")
    if cols and isinstance(cols, list):
        return cols

    for items_key in _ITEMS_KEYS:
        block = json_data.get(items_key)
        if isinstance(block, dict):
            cols = block.get("Reporting Columns") or block.get("reporting_columns")
            if cols and isinstance(cols, list):
                return cols

    for v in json_data.values():
        if isinstance(v, dict):
            cols = v.get("Reporting Columns") or v.get("reporting_columns")
            if cols and isinstance(cols, list):
                return cols

    return []


def _discover_page_no(json_data: dict) -> Any:
    """
    Find Metadata.Page No from wherever it lives in the JSON.
    Returns the raw value (string, int, or list) or None.
    """
    # Standard top-level Metadata
    metadata = json_data.get("Metadata")
    if isinstance(metadata, dict):
        val = metadata.get("Page No") or metadata.get("page_no")
        if val is not None:
            return val

    # Nested inside <TYPE> Items
    for items_key in _ITEMS_KEYS:
        block = json_data.get(items_key)
        if isinstance(block, dict):
            meta = block.get("Metadata")
            if isinstance(meta, dict):
                val = meta.get("Page No") or meta.get("page_no")
                if val is not None:
                    return val

    # One level deeper scan
    for v in json_data.values():
        if isinstance(v, dict):
            meta = v.get("Metadata")
            if isinstance(meta, dict):
                val = meta.get("Page No") or meta.get("page_no")
                if val is not None:
                    return val

    return None


def _collect_section_lists(json_data: dict) -> list[tuple[Any, list, str]]:
    """
    Discover all lists-of-row-dicts in the JSON, regardless of nesting.
    Returns (parent_container, row_list, key_in_parent) tuples.
    """
    found: list[tuple[Any, list, str]] = []
    visited_ids: set[int] = set()

    def _register(parent, lst, key):
        lst_id = id(lst)
        if lst_id in visited_ids:
            return
        visited_ids.add(lst_id)
        found.append((parent, lst, key))

    # 1. Standard "Sections" dict at top level
    sections = json_data.get("Sections")
    if isinstance(sections, dict):
        for sec_key, sec_val in sections.items():
            if isinstance(sec_val, list):
                _register(sections, sec_val, sec_key)

    # 2. "<TYPE> Items" blocks (may contain their own "Sections")
    for items_key in _ITEMS_KEYS:
        block = json_data.get(items_key)
        if isinstance(block, dict):
            inner_sections = block.get("Sections")
            if isinstance(inner_sections, dict):
                for sec_key, sec_val in inner_sections.items():
                    if isinstance(sec_val, list):
                        _register(inner_sections, sec_val, sec_key)
            for k, v in block.items():
                if isinstance(v, list) and k not in ("Reporting Columns",):
                    _register(block, v, k)

    # 3. Flat top-level list (some DSR / DEBT layouts)
    for top_key, top_val in json_data.items():
        if top_key in _META_KEYS or top_key in _ITEMS_KEYS or top_key == "Sections":
            continue
        if isinstance(top_val, list) and top_val and isinstance(top_val[0], dict):
            _register(json_data, top_val, top_key)

    # 4. Fallback: any list-of-dicts one level deep not yet captured
    for top_key, top_val in json_data.items():
        if isinstance(top_val, dict):
            for inner_key, inner_val in top_val.items():
                if isinstance(inner_val, list) and inner_val and isinstance(inner_val[0], dict):
                    _register(top_val, inner_val, inner_key)

    return found


# ── Public API ─────────────────────────────────────────────────────────────────

def attach_coordinates(json_data: dict, pdf_path: str) -> dict:
    """
    Public API. Takes the LLM output dict and the source (sliced) PDF path.
    Returns a deep copy with _coord keys injected for every numeric cell.

    Page numbers in _coord are translated from sliced-PDF page numbers to
    the ORIGINAL full-PDF page numbers using Metadata.Page No.

    Example:
      Metadata.Page No = "71,72,73"
      → sliced page 1 = original page 71
      → sliced page 2 = original page 72
      etc.

    Works for ALL statement types (SNP, SOA, GOV_BS, GOV_IS, PROP_SNP,
    PROP_IS, PROP_CFS, DSR, DEBT) across LG and NON-LG sectors.
    No prompt changes required.

    Args:
        json_data : Parsed JSON dict from any LLM normalisation call.
        pdf_path  : Path to the extracted (sliced) PDF sent to the LLM.

    Returns:
        Enriched dict (same structure + _coord keys).
        Returns original json_data unchanged if pdf_path does not exist
        or json_data is empty.
    """
    if not json_data:
        return json_data

    pdf_path_obj = Path(pdf_path)
    if not pdf_path_obj.exists():
        return json_data

    result = copy.deepcopy(json_data)

    # ── 1. Discover reporting columns ─────────────────────────────────────────
    reporting_columns = _discover_reporting_columns(result)
    if not reporting_columns:
        return result

    # ── 2. Build page translation map from Metadata.Page No ──────────────────
    #
    #    Metadata.Page No tells us which pages of the ORIGINAL full PDF
    #    this sliced PDF represents. We use it to translate sliced-PDF
    #    page numbers (1, 2, 3, …) into original-PDF page numbers (71, 72, …).
    #
    raw_page_no     = _discover_page_no(result)
    original_pages  = _parse_page_no(raw_page_no)
    page_map        = _build_page_map(original_pages)

    # ── 3. Build PDF coordinate index (one pdfplumber pass) ──────────────────
    try:
        coord_index = _CoordIndex(str(pdf_path_obj))
    except Exception:
        return result   # pdfplumber failure — return unenriched, never crash

    # ── 4. Find all row-list sections and enrich in-place ────────────────────
    section_lists = _collect_section_lists(result)
    if not section_lists:
        return result

    for parent_container, row_list, key in section_lists:
        enriched_rows = []
        for row in row_list:
            if isinstance(row, dict):
                enriched_rows.append(
                    _enrich_row(row, reporting_columns, coord_index, page_map)
                )
            else:
                enriched_rows.append(row)
        parent_container[key] = enriched_rows

    return result


# ── File-level convenience helper ──────────────────────────────────────────────

def attach_coordinates_to_file(json_path: str, pdf_path: str) -> None:
    """Load JSON, enrich with coordinates, write back in-place."""
    import json as _json

    json_path = Path(json_path)
    with open(json_path, encoding="utf-8") as f:
        data = _json.load(f)

    enriched = attach_coordinates(data, pdf_path)

    with open(json_path, "w", encoding="utf-8") as f:
        _json.dump(enriched, f, indent=2, ensure_ascii=False)


# ── CLI ────────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    import sys
    import json as _json

    if len(sys.argv) < 3:
        print("Usage: python coordinate_extractor.py <json_file> <pdf_file>")
        sys.exit(1)

    json_file = sys.argv[1]
    pdf_file  = sys.argv[2]

    with open(json_file, encoding="utf-8") as f:
        data = _json.load(f)

    enriched = attach_coordinates(data, pdf_file)

    out_path = json_file.replace(".json", "_with_coords.json")
    with open(out_path, "w", encoding="utf-8") as f:
        _json.dump(enriched, f, indent=2, ensure_ascii=False)

    print(f"Saved enriched JSON → {out_path}")