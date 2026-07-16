"""
compact_schema.py
==================
Single source of truth for the compact (short-key) JSON schema used to cut
Gemini/OpenAI COMPLETION-token cost. This is separate from context caching:
caching discounts INPUT tokens (system prompt + COA), this discounts
OUTPUT tokens (the JSON the LLM writes), which are priced highest.

SCHEMA SOURCE OF TRUTH: confirmed directly from a real production workbook
(DALLAS_COUNTY_2025..._All_Statements.xlsx, all 9 sheets) by reading each
sheet's header row. Confirmed columns per statement type:

    SNP       : Row Items
    DEBT      : Row Items
    SOA       : SOA Items
    GOV_BS    : GOV_BS Items
    GOV_IS    : GOV_IS Items
    PROP_SNP  : PROP_SNP Items
    PROP_IS   : PROP_IS Items
    PROP_CFS  : PROP_CFS Items
    DSR       : DSR Items

IMPORTANT CORRECTION FROM AN EARLIER VERSION OF THIS FILE: an earlier draft
assumed "Row Items" was used generically across ALL statement types, based
on only having seen SNP and DEBT sample output. That was WRONG — 7 of the
9 statement types use a TYPE-SPECIFIC "<TYPE> Items" column name instead.
This version restores correct per-statement-type resolution (matching
SECTION_ITEM_LABEL below), confirmed against real header rows from all 9
sheets, not inferred from a partial sample. If this ever silently reverts
to a single generic label again, COA/CSV mapping for 7 of 9 statement
types will break, since jsonToCsv.py / downstream code expects the
TYPE-SPECIFIC long key name (e.g. "SOA Items"), not "Row Items".

The CSV/Excel "Section" column (denormalized section name per row) is a
flattening artifact of jsonToCsv.py and is NEVER a JSON key in the LLM's
output — it doesn't appear here and needs no mapping.

How it works:
  1. The LLM is instructed to emit JSON using SHORT keys (defined below)
     instead of the normal long, human-readable keys, for every field
     EXCEPT the reporting-column value keys (those stay as their real
     column names — see "REPORTING COLUMN KEYS" note below). The Items
     label itself ("_ri") is generic in the SHORT form but resolves to
     the correct TYPE-SPECIFIC long name on expansion, based on the
     stmt_type passed in.
  2. expand_compact_json() walks that short-keyed JSON and renames every
     known short key back to its full long-form name, recursively, before
     anything else (jsonToCsv.py, validation, file save) ever sees it.

The LLM's REASONING, COA mapping, validation, and row content are 100%
unchanged — only the key LABELS it writes for the fixed fields shrink.
This is a pure wire-format optimization, not a structuring change.
"""

from __future__ import annotations

# ─────────────────────────────────────────────────────────────────────────
# SHORT KEY -> LONG KEY MAP  (fields that do NOT vary by statement type)
# ─────────────────────────────────────────────────────────────────────────
# Do NOT reuse a short key for two different long keys.

SHORT_TO_LONG: dict[str, str] = {
    # ── Top level ──
    "_md":  "Metadata",
    "_rc":  "Reporting Columns",
    "_sec": "Sections",

    # ── Metadata block ──
    "_iss": "Issuer Name",
    "_stm": "Statement",
    "_fye": "FYE",
    "_pg":  "Page No",
    "_cur": "Currency reported",

    # ── Row-level FIXED fields that DON'T vary by statement type ──
    "_cf":  "COA Flag",
    "_cd":  "COA Datapoint",
    "_tcs": "Total Check Status",

    # NOTE: "_ri" (the Items-label field) is intentionally NOT in this
    # table — it's handled separately below via SECTION_ITEM_LABEL because
    # its LONG name depends on stmt_type. See _expand_key() / _compact_key().
}

# ─────────────────────────────────────────────────────────────────────────
# PER-STATEMENT-TYPE "Items" COLUMN — CONFIRMED FROM REAL WORKBOOK HEADERS
# ─────────────────────────────────────────────────────────────────────────
# Maps the stmt_type string (as produced by pipeline.py's detect_type(),
# e.g. "_SNP", "_GOV_BS") to that statement type's actual long-form Items
# column name. Every one of these 9 entries was read directly from a real
# production .xlsx's sheet header row — not inferred.
SECTION_ITEM_LABEL: dict[str, str] = {
    "_SNP":      "Row Items",
    "_SOA":      "SOA Items",
    "_GOV_BS":   "GOV_BS Items",
    "_GOV_IS":   "GOV_IS Items",
    "_PROP_SNP": "PROP_SNP Items",
    "_PROP_IS":  "PROP_IS Items",
    "_PROP_CFS": "PROP_CFS Items",
    "_DSR":      "DSR Items",
    "_DEBT":     "Row Items",
}

# ─────────────────────────────────────────────────────────────────────────
# COA FLAG SHORT FORMS  (compresses every row by ~5 chars)
# ─────────────────────────────────────────────────────────────────────────
FLAG_SHORT_TO_LONG: dict[str, str] = {
    "G": "DPG",
    "D": "DP",
    "C": "CP",
}
FLAG_LONG_TO_SHORT: dict[str, str] = {v: k for k, v in FLAG_SHORT_TO_LONG.items()}

# ─────────────────────────────────────────────────────────────────────────
# TOTAL CHECK STATUS SHORT FORMS
# ─────────────────────────────────────────────────────────────────────────
STATUS_SHORT_TO_LONG: dict[str, str] = {
    "P":  "PASS",
    "":   "",
}

# Short keys for row-array positions (used only inside instruction prompt)
ROW_ARRAY_POSITIONS = ["FLAG", "ITEM", "DATAPOINT", "TCS", "V1", "V2", "..."]

# Used only if stmt_type is missing/unrecognized — should not normally
# happen since pipeline.py always passes a real stmt_type into the
# normalize_one_* call chain, but kept as a safe fallback rather than
# crashing.
DEFAULT_ITEMS_LABEL = "Row Items"

_ITEMS_SHORT_KEY = "_ri"


def _resolve_items_label(stmt_type: str) -> str:
    return SECTION_ITEM_LABEL.get(stmt_type, DEFAULT_ITEMS_LABEL)


def _invert(d: dict[str, str]) -> dict[str, str]:
    inv: dict[str, str] = {}
    for short, long in d.items():
        if long in inv:
            raise ValueError(
                f"compact_schema.py: duplicate long key '{long}' mapped from "
                f"both '{inv[long]}' and '{short}' — fix SHORT_TO_LONG."
            )
        inv[long] = short
    return inv


LONG_TO_SHORT: dict[str, str] = _invert(SHORT_TO_LONG)



def build_short_key_instruction(stmt_type: str) -> str:
    """
    Tell the LLM to emit output in the COMPACT-ARRAY schema:
      • top-level keys as short codes
      • each section's rows as POSITIONAL JSON ARRAYS (no per-row keys)
      • flag codes single-letter, status single-letter
      • numbers without commas
    
    All extraction/COA/validation logic in the main prompt still applies.
    """
    items_label = _resolve_items_label(stmt_type)

    return f"""

╔══════════════════════════════════════════════════════════════════════════════╗
║  ★★★ HIGHEST-PRIORITY OUTPUT FORMAT — OVERRIDES ALL "OUTPUT SHAPE" /         ║
║       "JSON OUTPUT FORMAT" / "FINAL SELF-CHECK" SECTIONS ABOVE ★★★           ║
║                                                                              ║
║  You MUST output the COMPACT-ARRAY schema below.  Do NOT output the          ║
║  long-form keys (\"Metadata\", \"Sections\", \"COA Flag\",                       ║
║  \"COA Datapoint\", \"Total Check Status\", \"{items_label}\")                  ║
║  even though the prompt above describes them.                                ║
╚══════════════════════════════════════════════════════════════════════════════╝

Return VALID JSON ONLY. No markdown fences, no commentary.

────────────── TOP-LEVEL SHAPE (exactly 3 keys) ──────────────
{{
  "_md":  {{ ...metadata... }},
  "_rc":  [ "<col1>", "<col2>", ... ],
  "_sec": {{ "<Section Name>": [ <row>, <row>, ... ], ... }}
}}

────────────── METADATA "_md" (single-letter sub-keys) ──────────────
  "s" -> Statement title
  "i" -> Issuer Name
  "f" -> FYE
  "p" -> Page No
  "c" -> Currency reported  (default "USD ($)")
  "u" -> Unit (only if the original prompt asks for it)

────────────── EACH ROW = JSON ARRAY (positional, NO KEYS) ──────────────

  [ FLAG, ITEM, DATAPOINT, TCS, V1, V2, ..., VN ]

  FLAG       -> "G" for DPG  |  "D" for DP  |  "C" for CP
  ITEM       -> the row-label string (replaces "{items_label}" key)
  DATAPOINT  -> COA Datapoint string (same rules as main prompt)
  TCS        -> Total Check Status, SHORT form:
                  ""                   -> DP / DPG rows
                  "P"                  -> CP PASS
                  "F:<col>|F:<col>"    -> CP FAIL, pipe-separated cols
                  "Skipped - <reason>" -> verbatim
  V1..VN     -> N values, same order as "_rc"

  NEVER output a row as an object with keys.
  ALWAYS as a flat JSON array.

────────────── NUMERIC VALUE FORMAT ──────────────
  • NO commas, NO "$", NO spaces.
        "1,234,567"  ->  "1234567"
        "(123,456)"  ->  "-123456"   (use leading minus, NOT parentheses)
  • Dash cell -> "-"
  • DPG rows: every Vk = null   (JSON null)
  • Output values as STRINGS, except null for DPG rows.

────────────── MINIMAL EXAMPLE (PROP_SNP, 3 columns) ──────────────
{{
  "_md": {{
    "s": "Proprietary Funds Statement of Net Position",
    "i": "City of X",
    "f": "June 30, 2025",
    "p": "47",
    "c": "USD ($)",
    "u": "Units"
  }},
  "_rc": ["Water Fund", "Sewer Fund", "Total Enterprise Funds"],
  "_sec": {{
    "Assets": [
      ["G", "Assets", "Assets", "", null, null, null],
      ["D", "Current assets: Cash and investments",
            "Cash and Investments", "", "1234567", "2345678", "3580245"],
      ["D", "Current assets: Accounts receivable",
            "Accounts Receivable", "", "50000", "-", "50000"],
      ["C", "Total current assets", "Total Current Assets", "P",
            "1284567", "2345678", "3630245"]
    ]
  }}
}}

All other rules (hierarchy flattening, section splitting, COA mapping,
Total Check arithmetic, dash preservation, column assignment, two-page
merging, anti-shift rules, fallback rules) from the prompt above STILL
APPLY UNCHANGED.  Only the OUTPUT SHAPE and NUMBER FORMAT change.

═════════════════════════════════════════════════════════════════════════════
"""


def _expand_status(t):
    """Expand short status to long form."""
    if not isinstance(t, str):
        return ""
    if t in STATUS_SHORT_TO_LONG:
        return STATUS_SHORT_TO_LONG[t]
    if t.startswith("F:"):
        # "F:Maturities|F:Interest" -> "Maturities: FAIL | Interest: FAIL"
        parts = [p[2:] if p.startswith("F:") else p for p in t.split("|")]
        return " | ".join(f"{p}: FAIL" for p in parts)
    return t   # "Skipped - ..." passes through verbatim


def _expand_row_array(arr, items_key, columns):
    """
    Convert a positional row array into the long-form row dict.
      arr = [FLAG, ITEM, DATAPOINT, TCS, V1, V2, ..., VN]
    """
    if not isinstance(arr, list) or len(arr) < 4:
        return None
    
    flag_code = arr[0]
    item      = arr[1]
    dp        = arr[2]
    tcs       = arr[3]
    vals      = list(arr[4:4 + len(columns)])
    
    # Pad missing values with "-"
    while len(vals) < len(columns):
        vals.append("-")
    
    row = {
        "COA Flag":           FLAG_SHORT_TO_LONG.get(flag_code, flag_code),
        items_key:            item,
        "COA Datapoint":      dp,
        "Total Check Status": _expand_status(tcs),
    }
    for col, v in zip(columns, vals):
        row[col] = v
    return row


def expand_compact_json(data, stmt_type: str = "") -> dict:
    """
    Recursively expand the LLM's compact JSON back to the full long-form
    schema that jsonToCsv.py / downstream consumers expect.
    
    AUTO-DETECTS FORMAT — handles 3 cases safely:
      1. FULL FORMAT     (LLM returned "Metadata"/"Sections")     → passthrough
      2. COMPACT-ARRAY   (LLM returned "_md"/"_rc"/"_sec" + arrays) → expand
      3. UNKNOWN/EMPTY                                             → return as-is + warn
    
    Never silently drops data — every code path returns SOMETHING and logs
    a [WARN] if format is unrecognized, so failures are easy to diagnose.
    """
    # ── Reject obviously broken input ──
    if data is None:
        print(f"   [WARN] expand_compact_json: input is None for {stmt_type}")
        return None
    
    if not isinstance(data, dict):
        print(f"   [WARN] expand_compact_json: input is {type(data).__name__} "
              f"(not dict) for {stmt_type}")
        return data
    
    if not data:
        print(f"   [WARN] expand_compact_json: empty dict for {stmt_type}")
        return data

    # ── CASE 1: FULL FORMAT — LLM ignored override, returned long-form ──
    if "Metadata" in data or "Sections" in data or "Reporting Columns" in data:
        print(f"   [INFO] expand_compact_json: full schema detected for "
              f"{stmt_type} — passthrough")
        return data

    # ── CASE 2: COMPACT-ARRAY FORMAT ──
    if "_md" in data or "_sec" in data or "_rc" in data:
        items_key = _resolve_items_label(stmt_type)

        # ── Metadata: expand single-letter sub-keys to full names ──
        md = data.get("_md", {}) or {}
        metadata = {
            "Statement":         md.get("s", ""),
            "Issuer Name":       md.get("i", ""),
            "FYE":               md.get("f", ""),
            "Page No":           md.get("p", ""),
            "Currency reported": md.get("c", "USD ($)"),
        }
        if "u" in md:
            metadata["Unit"] = md["u"]

        # ── Reporting Columns ──
        cols = list(data.get("_rc", []) or [])

        # ── Sections: each row is a positional array ──
        sections_out = {}
        for sec_name, rows in (data.get("_sec", {}) or {}).items():
            out_rows = []
            for r in rows or []:
                # Handle BOTH compact-array AND already-expanded dict rows
                # (LLM sometimes mixes formats — be defensive)
                if isinstance(r, list):
                    expanded = _expand_row_array(r, items_key, cols)
                    if expanded:
                        out_rows.append(expanded)
                elif isinstance(r, dict):
                    # Already expanded — pass through
                    out_rows.append(r)
            sections_out[sec_name] = out_rows

        result = {
            "Metadata":          metadata,
            "Reporting Columns": cols,
            "Sections":          sections_out,
        }

        # Sanity check
        row_count = sum(len(v) for v in sections_out.values() if isinstance(v, list))
        #print(f"   [INFO] expand_compact_json: compact schema expanded for "
              #f"{stmt_type} — {len(sections_out)} sections, {row_count} rows")
        return result

    # ── CASE 3: UNKNOWN FORMAT — log loudly, return as-is ──
    #print(f"   [WARN] expand_compact_json: unknown format for {stmt_type}, "
          #f"keys={list(data.keys())[:10]} — returning as-is")
    return data

def compact_size_estimate(sample_json: dict, stmt_type: str = "") -> tuple[int, int]:
    """
    Debug helper: returns (long_form_char_count, compact_form_char_count)
    for a sample JSON dict, using json.dumps with no extra whitespace, so
    you can sanity-check actual savings on a real output sample.

    `stmt_type` resolves which long Items-column name to treat as
    compactable to "_ri" (must match the statement type the sample
    actually represents, e.g. "_SOA" for an SOA sample).
    """
    import json

    items_label = _resolve_items_label(stmt_type)

    def _compact_key(key: str) -> str:
        if key == items_label:
            return _ITEMS_SHORT_KEY
        return LONG_TO_SHORT.get(key, key)

    def _compact(node):
        if isinstance(node, dict):
            return {_compact_key(k): _compact(v) for k, v in node.items()}
        elif isinstance(node, list):
            return [_compact(item) for item in node]
        else:
            return node

    long_str = json.dumps(sample_json, separators=(",", ":"))
    compact_str = json.dumps(_compact(sample_json), separators=(",", ":"))
    return len(long_str), len(compact_str)