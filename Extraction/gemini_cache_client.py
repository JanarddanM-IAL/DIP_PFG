"""
gemini_cache_client.py
Context-cached Gemini calls for PAID tier — DEEP-CACHE EDITION.

KEY DESIGN PRINCIPLE:
  Everything that is identical across calls of the same stmt_type goes
  INTO the cache (25% input rate). Only the PDF and per-call page_info
  go fresh (full input rate).

CACHE CONTENTS (constant per stmt_type):
  • System prompt (the full prompt file)
  • Compact-mode instruction (build_short_key_instruction)
  • Filtered COA Master (per-stmt_type subset)
  • Reporting columns lock (if any)
  • A leading role/instruction header

FRESH CONTENT (varies per call):
  • The PDF (file bytes)
  • The page reference block ("Source pages: page 47" etc.)

This maximizes the cached prefix length → bigger savings per call.
"""

import os
import json
import hashlib
from pathlib import Path
from threading import Lock

try:
    from google import genai
    from google.genai import types
    from google.genai.errors import APIError
except ImportError as e:
    raise ImportError(
        "google-genai not installed. Run: pip install google-genai"
    ) from e

from compact_schema import build_short_key_instruction, expand_compact_json

# Per-process registry: { "stmt_type|model|content_hash" : cache_name }
_CACHE_REGISTRY: dict[str, str] = {}
_REGISTRY_LOCK = Lock()
_CACHE_TTL_SECONDS = 21_600   # 6 hours — survives a full batch run

# Cached input ≈ 25% of normal input rate for Gemini
GEMINI_CACHED_RATES = {
    "gemini-3.5-flash":              (0.375,  9.00),
    "gemini-3.1-flash-lite-preview": (0.0625, 1.50),
    "gemini-2.5-flash":              (0.0375, 0.60),
    "gemini-2.5-pro":                (0.3125, 10.00),
}

# Normal input rates (for fresh-token cost component)
GEMINI_NORMAL_RATES = {
    "gemini-3.5-flash":              (1.50,  9.00),
    "gemini-3.1-flash-lite-preview": (0.25,  1.50),
    "gemini-2.5-flash":              (0.15,  0.60),
    "gemini-2.5-pro":                (1.25, 10.00),
}


class CacheUnavailableError(RuntimeError):
    """Raised when context caching cannot be used → pipeline must fall back."""
    pass


def _client():
    api_key = os.environ.get("GEMINI_API_KEY")
    if not api_key:
        raise CacheUnavailableError("GEMINI_API_KEY not set")
    return genai.Client(api_key=api_key)


# ─────────────────────────────────────────────────────────────────────────
# CACHE CONSTRUCTION — Everything constant per stmt_type goes in here
# ─────────────────────────────────────────────────────────────────────────
def _build_cached_text(stmt_type, prompt_text, coa_text,
                       reporting_columns, coa_map_rules: str = "") -> str:
    parts = [
        "=== ROLE ===",
        "You are a deterministic financial-statement table parser. "
        "Follow the system prompt below exactly. The PDF will be attached "
        "in the user message.",
        "",
        "=== SYSTEM PROMPT ===",
        prompt_text,
        "",
        "=== COMPACT OUTPUT INSTRUCTION ===",
        build_short_key_instruction(stmt_type),
        "",
        "=== STANDARD COA MASTER (filtered for this statement type) ===",
        coa_text,
    ]
    if coa_map_rules:
        parts.append("")
        parts.append("=== COA DATAPOINT MAPPING RULES ===")
        parts.append(coa_map_rules)
    if reporting_columns:
        parts.append("")
        parts.append("=== REPORTING COLUMNS LOCK ===")
        parts.append(json.dumps(reporting_columns, indent=2))
    return "\n".join(parts)

def _content_hash(text: str) -> str:
    """Short hash to invalidate cache when prompt/COA content changes."""
    return hashlib.sha256(text.encode("utf-8")).hexdigest()[:12]


def _get_or_create_cache(stmt_type, prompt_text, coa_text,
                          reporting_columns, model,
                          coa_map_rules: str = "") -> str:
    """
    Returns the cache name for this (stmt_type, model, content_hash).
    Creates on first call; reused for subsequent calls.
    """
    cached_text = _build_cached_text(
        stmt_type, prompt_text, coa_text, reporting_columns,
        coa_map_rules=coa_map_rules,   # ← ADD
    )
    chash = _content_hash(cached_text)
    key = f"{stmt_type}|{model}|{chash}"
    
    with _REGISTRY_LOCK:
        if key in _CACHE_REGISTRY:
            print(f"   [CACHE REUSE] {stmt_type} ({model}) → "
                  f"{_CACHE_REGISTRY[key][-20:]}")
            return _CACHE_REGISTRY[key]
    
    # Create new cache
    try:
        client = _client()
        cache = client.caches.create(
            model=model,
            config=types.CreateCachedContentConfig(
                contents=[
                    types.Content(
                        role="user",
                        parts=[types.Part.from_text(text=cached_text)],
                    )
                ],
                ttl=f"{_CACHE_TTL_SECONDS}s",
                display_name=f"{stmt_type}_{model}_{chash}",
            ),
        )
    except APIError as e:
        raise CacheUnavailableError(
            f"Cannot create cache for {stmt_type} on {model}: {e}"
        ) from e
    
    cache_name = cache.name
    print(f"   [CACHE NEW] {stmt_type} ({model}) → "
          f"{cache_name[-20:]}  TTL=6h  content_hash={chash}")
    
    with _REGISTRY_LOCK:
        _CACHE_REGISTRY[key] = cache_name
    return cache_name


# ─────────────────────────────────────────────────────────────────────────
# COST CALCULATION (transparent — shows fresh vs cached breakdown)
# ─────────────────────────────────────────────────────────────────────────
def _calc_cost(model: str,
               prompt_tokens: int,
               cached_tokens: int,
               completion_tokens: int) -> tuple[float, str]:
    """Returns (cost_usd, breakdown_string)."""
    in_rate, out_rate     = GEMINI_NORMAL_RATES.get(model, (1.50, 9.00))
    c_in_rate, _          = GEMINI_CACHED_RATES.get(model, (in_rate * 0.25, out_rate))
    
    fresh = max(0, prompt_tokens - cached_tokens)
    
    fresh_cost  = (fresh             / 1_000_000) * in_rate
    cached_cost = (cached_tokens     / 1_000_000) * c_in_rate
    output_cost = (completion_tokens / 1_000_000) * out_rate
    total       = fresh_cost + cached_cost + output_cost
    
    ratio = (cached_tokens / prompt_tokens * 100) if prompt_tokens else 0
    breakdown = (
        f"cached={cached_tokens:,}({ratio:.0f}%) "
        f"fresh={fresh:,} out={completion_tokens:,} → "
        f"${total:.4f}"
    )
    return total, breakdown


# ─────────────────────────────────────────────────────────────────────────
# MAIN ENTRY POINT — Sends ONLY the PDF + page_info fresh
# ─────────────────────────────────────────────────────────────────────────
def normalize_one_gemini_cached(
    pdf_path, prompt_path, coa_text, reporting_columns,
    model, max_tokens, page_info=None,
    stmt_type="UNKNOWN",
    coa_map_rules: str = "",           # ← ADD
) -> dict:
    """
    PAID-tier Gemini call using context caching.
    
    Sends ONLY the PDF + page reference as fresh content.
    Everything else (system prompt, compact instruction, COA, reporting
    columns) is in the cache → ~75% input discount.
    """
    from gemini_client import (
        _parse_json,
        _check_finish_reason,
        _build_page_reference_block,
    )
    
    pdf_name = Path(pdf_path).name
    
    # Read prompt file
    prompt_text = Path(prompt_path).read_text(encoding="utf-8", errors="replace")
    
    # Get or create the cache (constant content)
    cache_name = _get_or_create_cache(
        stmt_type=stmt_type,
        prompt_text=prompt_text,
        coa_text=coa_text,
        reporting_columns=reporting_columns,
        model=model,
        coa_map_rules=coa_map_rules,   # ← ADD
    )
    
    # ─── Build FRESH content: only PDF + page reference ───
    fresh_parts = []
    
    page_ref = _build_page_reference_block(page_info)
    if page_ref:
        fresh_parts.append(types.Part.from_text(text=page_ref))
    
    fresh_parts.append(
        types.Part.from_bytes(
            data=Path(pdf_path).read_bytes(),
            mime_type="application/pdf",
        )
    )
    
    fresh_parts.append(
        types.Part.from_text(
            text="\nExtract per the cached system prompt. "
                 "Return ONLY the compact JSON described in the "
                 "'COMPACT OUTPUT INSTRUCTION' section of the cache."
        )
    )
    
    # ─── Make the call using the cache ───
    try:
        client = _client()
        response = client.models.generate_content(
            model=model,
            contents=[
                types.Content(role="user", parts=fresh_parts),
            ],
            config=types.GenerateContentConfig(
                cached_content=cache_name,
                max_output_tokens=max_tokens,
                temperature=0.0,
            ),
        )
    except APIError as e:
        msg = str(e).lower()
        if "cache" in msg or "not found" in msg:
            # Cache may have expired or been deleted — invalidate and retry
            with _REGISTRY_LOCK:
                for k, v in list(_CACHE_REGISTRY.items()):
                    if v == cache_name:
                        del _CACHE_REGISTRY[k]
            raise CacheUnavailableError(
                f"Cache expired or invalid for {stmt_type}: {e}"
            ) from e
        raise
    
    # ─── Parse response ───
    raw = _check_finish_reason(response, model)
    data = _parse_json(raw)
    data = expand_compact_json(data, stmt_type)
    
    # ─── Extract usage stats ───
    usage = response.usage_metadata
    prompt_tokens     = usage.prompt_token_count or 0
    cached_tokens     = usage.cached_content_token_count or 0
    completion_tokens = usage.candidates_token_count or 0
    total_tokens      = usage.total_token_count or 0
    
    cost, breakdown = _calc_cost(
        model, prompt_tokens, cached_tokens, completion_tokens
    )
    
    cached_hit = cached_tokens > 0
    tag = "[CACHE HIT ]" if cached_hit else "[CACHE MISS]"
    print(f"   {tag} {pdf_name}: {breakdown}")
    
    return {
        "data":       data,
        "raw":        raw,
        "usage": {
            "prompt_tokens":     prompt_tokens,
            "completion_tokens": completion_tokens,
            "total_tokens":      total_tokens,
            "cached_tokens":     cached_tokens,
        },
        "cost_usd":   cost,
        "cached_hit": cached_hit,
        "model":      model,
    }


# ─────────────────────────────────────────────────────────────────────────
# CACHE CLEANUP
# ─────────────────────────────────────────────────────────────────────────
def cleanup_caches():
    """Delete all caches this process created. Call at end of pipeline."""
    if not _CACHE_REGISTRY:
        return
    try:
        client = _client()
    except CacheUnavailableError:
        return
    
    deleted = 0
    for key, name in list(_CACHE_REGISTRY.items()):
        try:
            client.caches.delete(name=name)
            #print(f"   [CACHE DEL] {key}")
            deleted += 1
        except Exception as e:
            print(f"   [WARN] cache delete failed for {key}: {e}")
    _CACHE_REGISTRY.clear()
    print(f"   [CACHE CLEANUP] Deleted {deleted} cache(s)")