"""
openai_cache_client.py
PAID-tier OpenAI calls using automatic prompt caching.

Mirrors pipeline.py's normalize_one_openai() exactly so it works with
reasoning models like gpt-5.5. OpenAI automatically caches identical
prefixes ≥1024 tokens (5-min sliding TTL) → 50% discount on cached input.

Key differences from a "normal" chat completion:
  - Uses max_completion_tokens (NOT max_tokens) — required by reasoning models
  - No temperature override (reasoning models reject it)
  - No response_format json_object (interacts badly with file inputs on gpt-5.5)
  - PDF sent as {"type":"file", "file":{...}} — the only format gpt-5.5 accepts

Return shape matches normalize_one_openai() in pipeline.py exactly:
    {
        "data": {...},
        "usage": {"prompt_tokens": int, "completion_tokens": int,
                   "total_tokens": int, "cached_tokens": int},
        "cost_usd": float,     # cost with caching discount already applied
        "cached_hit": bool,
        "model": str,
    }
"""

import os
import json
import base64
import time
from pathlib import Path
from openai import OpenAI, APIError, APITimeoutError, RateLimitError, BadRequestError

# FIX: compact_schema.py exports build_short_key_instruction() and
# expand_compact_json() — NOT "COMPACT_OVERRIDE" / "expand_compact". Those
# names never existed in compact_schema.py, so this import would have
# raised ImportError at module load (or silently used a different, broken
# shadow file if one existed elsewhere on the path).
from compact_schema import build_short_key_instruction, expand_compact_json

# Cached input rate is 50% of normal input rate (OpenAI standard discount).
# IMPORTANT: these rates must be kept in sync with OPENAI_COST_TABLE in
# pipeline.py — the normal-rate side in particular.
OPENAI_CACHED_RATES = {
    "gpt-4o-mini":  (0.075,  0.60),
    "gpt-4.1-mini": (0.20,   1.60),
    "gpt-4o":       (1.25,  10.00),
    "gpt-4.1":      (1.00,   8.00),
    "gpt-5.4-mini": (0.20,   1.60),
    "gpt-5.5":      (2.50,  30.00),   # 50% of corrected normal input rate (5.00)
}

# Must match OPENAI_COST_TABLE in pipeline.py.
OPENAI_NORMAL_RATES = {
    "gpt-4o-mini":  (0.15,   0.60),
    "gpt-4.1-mini": (0.40,   1.60),
    "gpt-4o":       (2.50,  10.00),
    "gpt-4.1":      (2.00,   8.00),
    "gpt-5.4-mini": (0.40,   1.60),
    "gpt-5.5":      (5.00,  30.00),
}


def _client():
    api_key = os.environ.get("OPENAI_API_KEY")
    if not api_key:
        raise RuntimeError("OPENAI_API_KEY not set")
    return OpenAI(api_key=api_key)


def _calc_cost(model: str, prompt_tokens: int, cached_tokens: int,
               completion_tokens: int) -> float:
    in_rate, out_rate = OPENAI_NORMAL_RATES.get(model, (0.0, 0.0))
    c_in_rate, _      = OPENAI_CACHED_RATES.get(model, (in_rate * 0.5, out_rate))
    fresh = max(0, prompt_tokens - cached_tokens)
    return (
        (fresh             / 1_000_000) * in_rate +
        (cached_tokens     / 1_000_000) * c_in_rate +
        (completion_tokens / 1_000_000) * out_rate
    )


def normalize_one_openai_cached(
    pdf_path: str,
    prompt_path: str,
    coa_text: str,
    reporting_columns,
    model: str,
    max_tokens: int,
    page_info: dict | None = None,
    stmt_type: str = "UNKNOWN",
    coa_map_rules: str = "",           # ← ADD
) -> dict:
    """
    PAID-tier OpenAI call. Uses automatic prompt caching.
    Mirrors normalize_one_openai() from pipeline.py exactly for compatibility.
    """
    # Lazy imports from pipeline to keep existing helpers
    from pipeline import extract_page_info_from_filename, parse_json_response
    from pdf_to_indented_text import pdf_to_indented_text

    client = _client()
    system_prompt = Path(prompt_path).read_text(encoding="utf-8", errors="replace")
    system_prompt = system_prompt + "\n\n" + build_short_key_instruction(stmt_type)
    if coa_map_rules:                  # ← ADD BLOCK
        system_prompt = system_prompt + "\n\n" + coa_map_rules
    pdf_bytes     = Path(pdf_path).read_bytes()
    b64           = base64.b64encode(pdf_bytes).decode("utf-8")
    pdf_data_uri  = f"data:application/pdf;base64,{b64}"
    pdf_name      = Path(pdf_path).name

    # Use page_info already passed in (from pipeline) or derive
    if page_info is None:
        page_info = extract_page_info_from_filename(pdf_path)

    try:
        indented_text = pdf_to_indented_text(pdf_path)
    except Exception as e:
        print(f"  [WARN] pdf_to_indented_text failed ({e}), falling back to no indented text")
        indented_text = None

    instruction_lines = [
        "Normalize the financial statement from the attached PDF.",
        "Follow all steps in the system prompt exactly.",
        "Use the COA Master table above for COA Datapoint mapping.",
        "Return valid JSON only — no markdown, no extra text.",
    ]
    if reporting_columns:
        instruction_lines += ["", "[[REPORTING COLUMNS]]"] + reporting_columns

    # ── Build user_content with EXACT same order as working version ──
    user_content = [
        {"type": "file", "file": {"filename": pdf_name, "file_data": pdf_data_uri}},
        {
            "type": "text",
            "text": (
                "STANDARD COA MASTER (pipe-delimited):\n"
                "Format: Statement | Section | COA Flag | COA Datapoint\n\n"
                + coa_text
            ),
        },
    ]

    if page_info:
        if page_info["start"] == page_info["end"]:
            page_note = (
                f"SOURCE PAGE REFERENCE:\n"
                f"This extracted PDF corresponds to {page_info['label']} of the "
                f"original full financial report (1-based page numbering).\n"
                f"When populating any 'Source Page', 'Page Number', or 'Page Reference' "
                f"field in the output JSON, use: {page_info['start']}"
            )
        else:
            pages_str = ", ".join(str(p) for p in page_info["pages"])
            page_note = (
                f"SOURCE PAGE REFERENCE:\n"
                f"This extracted PDF corresponds to {page_info['label']} of the "
                f"original full financial report (1-based page numbering).\n"
                f"The content spans pages: {pages_str}.\n"
                f"When populating any 'Source Page', 'Page Number', or 'Page Reference' "
                f"field in the output JSON, use the page where each data point appears."
            )
        user_content.append({"type": "text", "text": page_note})
        print(f"  Source pages : {page_info['label']}")
    else:
        print(f"  Source pages : [WARN] no page tag found in filename — skipping page context")

    if indented_text:
        user_content.append({
            "type": "text",
            "text": (
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
            ),
        })

    user_content.append({
        "type": "text",
        "text": "\n".join(instruction_lines),
    })

    messages = [
        {"role": "system", "content": system_prompt},
        {"role": "user",   "content": user_content},
    ]

    # ── Call with retry on transient errors, fail-fast on 400 ──
    last_err = None
    response = None
    for attempt in range(1, 4):
        try:
            
            response = client.chat.completions.create(
                model=model,
                max_completion_tokens=max_tokens,   # reasoning models need this
                messages=messages,
                # NO temperature   (reasoning models reject override)
                # NO response_format (interacts badly with file inputs on gpt-5.5)
            )
            break
        except BadRequestError as e:
            err_str = str(e)
            print(f"   [ERROR] OpenAI BadRequestError: {err_str[:400]}")
            raise RuntimeError(f"OpenAI 400: {err_str[:200]}") from e
        except (APITimeoutError, RateLimitError, APIError) as e:
            last_err = e
            wait = 10 * attempt
            print(f"   [WARN] OpenAI {type(e).__name__} (attempt {attempt}/3). Retry in {wait}s ...")
            time.sleep(wait)
    if response is None:
        raise RuntimeError(f"OpenAI call failed after retries: {last_err}")

    raw = (response.choices[0].message.content or "").strip()

    # ── Cost calc with cached prefix detection ──
    usage = response.usage
    prompt_tokens     = getattr(usage, "prompt_tokens", 0) or 0
    completion_tokens = getattr(usage, "completion_tokens", 0) or 0
    details = getattr(usage, "prompt_tokens_details", None)
    cached_tokens = 0
    if details is not None:
        cached_tokens = int(getattr(details, "cached_tokens", 0) or 0)

    cost = _calc_cost(model, prompt_tokens, cached_tokens, completion_tokens)

    print(
        f"  Tokens — prompt: {prompt_tokens:,}  completion: {completion_tokens:,}  "
        f"total: {prompt_tokens + completion_tokens:,}"
    )
    print(
        f"   [OPENAI {'HIT ' if cached_tokens else 'MISS'}] "
        f"cached={cached_tokens:,}  fresh={prompt_tokens - cached_tokens:,}  "
        f"out={completion_tokens:,}  ≈ ${cost:.4f}"
    )

    # ── Parse JSON, THEN expand compact keys ─────────────────────────────────
    # FIX: previously this dict had "data" defined TWICE — once correctly as
    # parse_json_response(raw), then immediately overwritten by
    # expand_compact("data", stmt_type), which passed the LITERAL STRING
    # "data" (not the parsed dict) into a function that doesn't even exist
    # under that name. Python dict literals silently keep only the LAST
    # value for a duplicate key, so the real parsed JSON was discarded
    # every single call — this is why every table failed with "No data in
    # response" even though tokens were being spent on real LLM calls.
    parsed_data = parse_json_response(raw)
    expanded_data = expand_compact_json(parsed_data, stmt_type=stmt_type)

    return {
        "data": expanded_data,
        "usage": {
            "prompt_tokens":     prompt_tokens,
            "completion_tokens": completion_tokens,
            "total_tokens":      prompt_tokens + completion_tokens,
            "cached_tokens":     cached_tokens,
        },
        "cost_usd":   cost,
        "cached_hit": cached_tokens > 0,
        "model":      model,
    }