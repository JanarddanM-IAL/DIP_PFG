"""
claude_cache_client.py

PAID-tier Claude calls using Anthropic prompt caching.

Design matches your existing OpenAI/Gemini cache clients:
Return shape:
{
    "data": dict,
    "prop_snp_split": list[dict],   # populated for multi-table PROP_* responses
    "usage": {
        "prompt_tokens": int,
        "completion_tokens": int,
        "total_tokens": int,
        "cached_tokens": int,
        "cache_creation_tokens": int,
        "cache_read_tokens": int,
    },
    "cost_usd": float,
    "cached_hit": bool,
    "model": str,
}

Claude has no free API tier in this project:
- Use this client only for provider="claude", tier="paid", batch=False.
"""

import base64
import json
import os
import re
from pathlib import Path

try:
    from anthropic import Anthropic
except ImportError as e:
    raise ImportError(
        "anthropic package not installed. Run: pip install anthropic"
    ) from e

from compact_schema import build_short_key_instruction, expand_compact_json


# ─────────────────────────────────────────────────────────────────────────────
# Claude model display-name aliasing
# ─────────────────────────────────────────────────────────────────────────────
# UI can pass friendly names or API IDs. This normalizes them.
CLAUDE_MODEL_ALIASES = {
    "Claude Sonnet 4.6": "claude-sonnet-4-6",
    "Claude Sonnet 4.5": "claude-sonnet-4-5",
    "Claude Opus 4.5":   "claude-opus-4-5",
    "Claude Opus 4.8":   "claude-opus-4-8",

    "claude-sonnet-4.6": "claude-sonnet-4-6",
    "claude-sonnet-4.5": "claude-sonnet-4-5",
    "claude-opus-4.5":   "claude-opus-4-5",
    "claude-opus-4.8":   "claude-opus-4-8",
}


def resolve_claude_model(model: str | None) -> str:
    if not model:
        return "claude-sonnet-4-6"
    return CLAUDE_MODEL_ALIASES.get(model, model)


# ─────────────────────────────────────────────────────────────────────────────
# Cost table ($ per 1M tokens)
# ─────────────────────────────────────────────────────────────────────────────
# Current practical pricing family:
# Sonnet: $3 input / $15 output
# Opus:   $5 input / $25 output
#
# Prompt caching:
# - cache write is typically higher than normal input
# - cache read is heavily discounted
# This table uses 5-minute ephemeral cache approximation:
# write = 1.25x normal input
# read  = 0.10x normal input
CLAUDE_NORMAL_RATES = {
    "claude-sonnet-4-6": (3.00, 15.00),
    "claude-sonnet-4-5": (3.00, 15.00),
    "claude-opus-4-8":   (5.00, 25.00),
    "claude-opus-4-5":   (5.00, 25.00),
}

DEFAULT_RATE = (3.00, 15.00)


def _rate_for_model(model: str) -> tuple[float, float]:
    return CLAUDE_NORMAL_RATES.get(model, DEFAULT_RATE)


def _calc_cost(
    model: str,
    input_tokens: int,
    output_tokens: int,
    cache_creation_tokens: int,
    cache_read_tokens: int,
) -> float:
    in_rate, out_rate = _rate_for_model(model)

    # Anthropic usage generally separates:
    # input_tokens                 = fresh normal input
    # cache_creation_input_tokens  = cache write (costs more than normal)
    # cache_read_input_tokens      = cache hit/read (heavily discounted)
    cache_write_rate = in_rate * 1.25
    cache_read_rate  = in_rate * 0.10

    return (
        (input_tokens          / 1_000_000) * in_rate
        + (cache_creation_tokens / 1_000_000) * cache_write_rate
        + (cache_read_tokens     / 1_000_000) * cache_read_rate
        + (output_tokens         / 1_000_000) * out_rate
    )


def _client() -> Anthropic:
    api_key = os.environ.get("ANTHROPIC_API_KEY")
    if not api_key:
        raise RuntimeError("ANTHROPIC_API_KEY not set")
    return Anthropic(api_key=api_key)


# ─────────────────────────────────────────────────────────────────────────────
# Multi-table delimiter constants
# ─────────────────────────────────────────────────────────────────────────────
# These must match exactly what the LLM prompt instructs the model to emit.

_MULTI_TABLE_DELIMITERS = [
    "---PROP_SNP_TABLE_BREAK---",
    "---PROP_IS_TABLE_BREAK---",
    "---PROP_CFS_TABLE_BREAK---",
]


# ─────────────────────────────────────────────────────────────────────────────
# JSON helpers
# ─────────────────────────────────────────────────────────────────────────────

def _parse_json(raw: str) -> dict:
    text = (raw or "").strip()

    # Direct JSON
    try:
        return json.loads(text)
    except Exception:
        pass

    # ```json ... ```
    m = re.search(r"```(?:json)?\s*([\s\S]+?)\s*```", text, re.IGNORECASE)
    if m:
        try:
            return json.loads(m.group(1).strip())
        except Exception:
            pass

    # First {...} block
    try:
        return json.loads(text[text.index("{"): text.rindex("}") + 1])
    except Exception:
        pass

    raise ValueError(f"Cannot parse JSON. First 500 chars:\n{text[:500]}")


def _extract_text_from_message(message) -> str:
    parts = []
    for block in getattr(message, "content", []) or []:
        if getattr(block, "type", None) == "text":
            parts.append(getattr(block, "text", "") or "")
        elif isinstance(block, dict) and block.get("type") == "text":
            parts.append(block.get("text", "") or "")
    return "\n".join(parts).strip()


def _build_page_reference_block(page_info: dict | None) -> str | None:
    if not page_info:
        return None

    pages_str = ",".join(str(p) for p in page_info.get("pages", []))
    label = page_info.get("label") or pages_str

    return (
        f"SOURCE PAGE REFERENCE - MANDATORY:\n"
        f"The attached extracted PDF was created from source page(s): {pages_str}.\n"
        f'For Metadata.Page No, you MUST copy exactly this value: "{label}".\n'
        f"Do NOT infer page number from any printed footer/header inside the PDF."
    )


def _build_cached_system_text(
    prompt_text: str,
    coa_text: str,
    reporting_columns,
    stmt_type: str,
    coa_map_rules: str = "",
) -> str:
    parts = [
        "You are a deterministic financial-statement table parser.",
        "Follow the statement prompt exactly.",
        "",
        "=== SYSTEM PROMPT ===",
        prompt_text,
        "",
        "=== COMPACT OUTPUT INSTRUCTION ===",
        build_short_key_instruction(stmt_type),
        "",
        "=== STANDARD COA MASTER FILTERED FOR THIS STATEMENT TYPE ===",
        coa_text or "",
    ]

    if coa_map_rules:
        parts.extend([
            "",
            "=== COA DATAPOINT MAPPING RULES ===",
            coa_map_rules,
        ])

    if reporting_columns:
        parts.extend([
            "",
            "=== LOCKED REPORTING COLUMNS ===",
            json.dumps(reporting_columns, ensure_ascii=False, indent=2),
        ])

    return "\n".join(parts)


# ─────────────────────────────────────────────────────────────────────────────
# Main Claude cached normalizer
# ─────────────────────────────────────────────────────────────────────────────

def normalize_one_claude_cached(
    pdf_path: str,
    prompt_path: str,
    coa_text: str,
    reporting_columns,
    model: str,
    max_tokens: int,
    page_info: dict | None = None,
    stmt_type: str = "UNKNOWN",
    coa_map_rules: str = "",
    indented_text: str | None = None,       # ← NEW
) -> dict:
    """
    Paid sync Claude normalization with prompt caching.

    Cache strategy:
    - cache prompt + compact instruction + COA + mapping rules + reporting columns
    - send PDF + page note as fresh per-call content

    Return shape:
    {
        "data":           dict | None,   # first (or only) parsed table
        "prop_snp_split": list[dict],    # all parsed tables (len >= 1 if multi-table)
        "usage":          {...},
        "cost_usd":       float,
        "cached_hit":     bool,
        "model":          str,
    }

    Multi-table handling:
    When the LLM returns two JSON tables separated by a delimiter
    (e.g. ---PROP_SNP_TABLE_BREAK---), both tables are parsed, expanded,
    and returned in prop_snp_split. "data" holds the first table so the
    caller's single-table fast-path still works unchanged.
    """

    model  = resolve_claude_model(model)
    client = _client()

    prompt_text        = Path(prompt_path).read_text(encoding="utf-8", errors="replace")
    cached_system_text = _build_cached_system_text(
        prompt_text=prompt_text,
        coa_text=coa_text,
        reporting_columns=reporting_columns,
        stmt_type=stmt_type,
        coa_map_rules=coa_map_rules,
    )

    pdf_bytes = Path(pdf_path).read_bytes()
    pdf_b64   = base64.b64encode(pdf_bytes).decode("utf-8")

    # ── Build FRESH content: page reference + indented text + PDF ────────────
    page_note = _build_page_reference_block(page_info)

    # Use pre-extracted indented_text if provided; fallback only if None
    if indented_text is None:
        try:
            from pdf_to_indented_text import pdf_to_indented_text as _pit
            indented_text = _pit(pdf_path)
        except Exception:
            indented_text = None

    # Assemble user content blocks
    user_content = []

    # 1. PDF document (always first so the model sees the visual layout)
    user_content.append(
        {
            "type": "document",
            "source": {
                "type": "base64",
                "media_type": "application/pdf",
                "data": pdf_b64,
            },
        }
    )

    # 2. Page reference note (mandatory when present)
    user_text = "Extract and normalize the attached financial statement PDF. Return JSON only."
    if page_note:
        user_text += "\n\n" + page_note

    # 3. Indented text block (hierarchy-preserving extraction)
    if indented_text:
        user_text += (
            "\n\n"
            "INDENTED TEXT EXTRACTION OF THE PDF (HIERARCHY-PRESERVING):\n"
            "The following is the financial statement text extracted with visual\n"
            "indentation preserved. Leading spaces indicate hierarchy level:\n"
            "  0 spaces = root level\n"
            "  2 spaces = level-1 child\n"
            "  4 spaces = level-2 grandchild\n"
            "  6 spaces = total/subtotal row\n\n"
            "USE THIS INDENTED TEXT (NOT the raw PDF) for Section E hierarchy "
            "flattening. Trust the leading spaces -- do not override them with "
            "semantic reasoning about label names.\n\n"
            "--- BEGIN INDENTED TEXT ---\n"
            + indented_text
            + "\n--- END INDENTED TEXT ---"
        )

    user_content.append({"type": "text", "text": user_text})

    response = client.messages.create(
        model=model,
        max_tokens=max_tokens,
        system=[
            {
                "type": "text",
                "text": cached_system_text,
                "cache_control": {"type": "ephemeral"},
            }
        ],
        messages=[
            {
                "role": "user",
                "content": user_content,
            }
        ],
    )

    raw = _extract_text_from_message(response)

    # ── Usage & cost ──────────────────────────────────────────────────────────
    usage_obj = getattr(response, "usage", None)

    input_tokens          = int(getattr(usage_obj, "input_tokens",                 0) or 0)
    output_tokens         = int(getattr(usage_obj, "output_tokens",                0) or 0)
    cache_creation_tokens = int(getattr(usage_obj, "cache_creation_input_tokens",  0) or 0)
    cache_read_tokens     = int(getattr(usage_obj, "cache_read_input_tokens",      0) or 0)

    cached_tokens = cache_creation_tokens + cache_read_tokens
    total_tokens  = input_tokens + output_tokens + cached_tokens

    cost_usd = _calc_cost(
        model=model,
        input_tokens=input_tokens,
        output_tokens=output_tokens,
        cache_creation_tokens=cache_creation_tokens,
        cache_read_tokens=cache_read_tokens,
    )

    usage_dict = {
        "prompt_tokens":         input_tokens + cached_tokens,
        "completion_tokens":     output_tokens,
        "total_tokens":          total_tokens,
        "cached_tokens":         cached_tokens,
        "cache_creation_tokens": cache_creation_tokens,
        "cache_read_tokens":     cache_read_tokens,
    }

    # ── Multi-table delimiter detection ───────────────────────────────────────
    # Must happen BEFORE _parse_json() because a delimited response is two
    # JSON documents joined by a literal string — not valid JSON on its own.
    active_delimiter = next((d for d in _MULTI_TABLE_DELIMITERS if d in raw), None)

    if active_delimiter:
        parts          = [p.strip() for p in raw.split(active_delimiter) if p.strip()]
        prop_snp_split = []
        first_data     = None

        for idx, part in enumerate(parts):
            try:
                parsed   = _parse_json(part)
                expanded = expand_compact_json(parsed, stmt_type)
                prop_snp_split.append(expanded)
                if first_data is None:
                    first_data = expanded
            except Exception as e:
                print(
                    f"  [CLAUDE CACHE] Multi-table parse error on sub-table "
                    f"{idx + 1}: {e}"
                )

        return {
            "data":           first_data,
            "prop_snp_split": prop_snp_split,
            "usage":          usage_dict,
            "cost_usd":       cost_usd,
            "cached_hit":     cache_read_tokens > 0,
            "model":          model,
        }

    # ── Normal single-table path ──────────────────────────────────────────────
    data = _parse_json(raw)
    data = expand_compact_json(data, stmt_type)

    return {
        "data":           data,
        "prop_snp_split": [],
        "usage":          usage_dict,
        "cost_usd":       cost_usd,
        "cached_hit":     cache_read_tokens > 0,
        "model":          model,
    }