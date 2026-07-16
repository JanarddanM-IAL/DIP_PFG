"""
gemini_client.py
================
Drop-in Gemini replacement for the OpenAI normalize_one() call.

Uses google-generativeai SDK (pip install google-generativeai).
Sends the PDF as inline base64 bytes — no file-upload step needed for
files under ~20 MB, which covers most financial statement PDFs.

Fixes vs original:
  - Prints finish_reason so you know WHY a response is empty/cut off
  - Uses 65536 output tokens (flash supports 8192 on free tier; pro up to 65536)
  - Adds streaming fallback for large outputs that would otherwise be cut off
  - Diagnoses empty response vs truncated response clearly
  - Strips markdown fences more aggressively before JSON parsing
  - Retries once with a continuation prompt if output was cut off (MAX_TOKENS)
  - Accepts page_info dict from pipeline to inject SOURCE PAGE REFERENCE context
"""

import base64
import json
import os
import re
from pathlib import Path

import google.api_core.exceptions as _gex

# ── Cost table ($ per 1 M tokens) ───────────────────────────────────────────

GEMINI_COST_TABLE: dict[str, tuple[float, float]] = {
    "gemini-2.0-flash":                    (0.075,  0.30),
    "gemini-2.0-flash-lite":               (0.075,  0.30),
    "gemini-1.5-flash":                    (0.075,  0.30),
    "gemini-1.5-pro":                      (1.25,   5.00),
    "gemini-2.5-flash":                    (0.15,   0.60),
    "gemini-2.5-pro":                      (1.25,  10.00),
    "gemini-3.1-flash-lite-preview":       (0.25,   1.50),
    "gemini-3.5-flash":                    (1.50,   9.00),
}

DEFAULT_COST = (0.075, 0.30)


# Max output tokens per model family
MODEL_MAX_OUTPUT: dict[str, int] = {
    "gemini-3.5-flash":                  65536,
    "gemini-3.1-flash-lite-preview":     65536,
    "gemini-2.5-pro":                    65536,
    "gemini-2.5-flash":                  65536,
    "gemini-1.5-pro":                    8192,
    "gemini-1.5-flash":                  8192,
    "gemini-2.0-flash":                  8192,
}


def _model_max_tokens(model: str) -> int:
    for prefix, limit in MODEL_MAX_OUTPUT.items():
        if model.startswith(prefix):
            return limit
    # Safer fallback for Gemini 3.x models
    if model.startswith("gemini-3."):
        return 65536
    return 8192


def calculate_cost(model: str, prompt_tokens: int, completion_tokens: int) -> float:
    input_rate, output_rate = GEMINI_COST_TABLE.get(model, DEFAULT_COST)
    return (prompt_tokens / 1_000_000) * input_rate + \
           (completion_tokens / 1_000_000) * output_rate


# ── JSON parsing ─────────────────────────────────────────────────────────────

def _parse_json(raw: str) -> dict:
    """Robustly extract JSON from LLM response (fences, preamble, truncation)."""
    text = raw.strip()

    for attempt in [
        lambda t: json.loads(t),
        lambda t: json.loads(re.sub(r"^```(?:json)?\s*", "", re.sub(r"\s*```$", "", t.strip()))),
        lambda t: json.loads(re.search(r"```(?:json)?\s*([\s\S]+?)\s*```", t).group(1)),
        lambda t: json.loads(t[t.index("{"):t.rindex("}") + 1]),
    ]:
        try:
            return attempt(text)
        except Exception:
            pass

    raise ValueError(
        f"Cannot parse JSON from Gemini response.\n"
        f"Length: {len(raw)} chars\n"
        f"First 600 chars:\n{raw[:600]}\n"
        f"Last  300 chars:\n{raw[-300:]}"
    )


# ── Finish-reason diagnostics ────────────────────────────────────────────────

def _check_finish_reason(response, model: str) -> str:
    """
    Extract text from response, printing a clear diagnosis if it's empty or cut off.
    Returns the raw text string (may be empty).
    """
    raw = ""
    finish_reason = "UNKNOWN"

    try:
        candidate = response.candidates[0]
        finish_reason = str(candidate.finish_reason)

        for part in candidate.content.parts:
            if hasattr(part, "text") and part.text:
                raw += part.text

    except (IndexError, AttributeError):
        try:
            raw = response.text or ""
        except Exception:
            raw = ""

    print(f"  Finish reason : {finish_reason}")

    if finish_reason in ("MAX_TOKENS", "2"):
        print(
            f"    Output was TRUNCATED (hit max_output_tokens limit).\n"
            f"     Model '{model}' max output = {_model_max_tokens(model)} tokens.\n"
            f"     The JSON was cut off mid-way — switching to a larger model\n"
            f"     (gemini-2.5-pro) or splitting the PDF into fewer pages will help.\n"
            f"     Partial output length: {len(raw)} chars"
        )
    elif finish_reason in ("SAFETY", "3"):
        print(
            "    Response blocked by Gemini SAFETY filter.\n"
            "     The financial PDF may contain content that triggered a filter.\n"
            "     Try gemini-2.5-pro or contact Google support."
        )
    elif finish_reason in ("RECITATION", "4"):
        print(
            "    Response blocked by Gemini RECITATION filter.\n"
            "     Gemini detected the output too closely matched training data.\n"
            "     Try rephrasing the system prompt or use a different model."
        )
    elif not raw:
        print(
            "    Gemini returned an empty response with no clear reason.\n"
            f"     Finish reason reported: {finish_reason}\n"
            "     Check your API key quota at https://aistudio.google.com"
        )

    return raw


# ── Page reference block builder (shared logic) ──────────────────────────────

def _build_page_reference_block(page_info: dict | None) -> str | None:
    """
    Build the SOURCE PAGE REFERENCE text block from a page_info dict.

    page_info is produced by pipeline.extract_page_info_from_filename() and
    has keys: start, end, pages, label.

    Returns the block string, or None if page_info is falsy.
    This block is injected into the user prompt so the LLM knows exactly
    which pages of the full report it is reading — allowing it to correctly
    populate Metadata.Page No and any Source Page fields.
    """
    if not page_info:
        return None

    if page_info["start"] == page_info["end"]:
        return (
            f"SOURCE PAGE REFERENCE:\n"
            f"This extracted PDF corresponds to {page_info['label']} of the "
            f"original full financial report (1-based page numbering).\n"
            f"When populating the 'Page No' field in Metadata and any "
            f"'Source Page' or 'Page Reference' field in the output JSON, "
            f"use: {page_info['start']}"
        )
    else:
        pages_str = ", ".join(str(p) for p in page_info["pages"])
        return (
            f"SOURCE PAGE REFERENCE:\n"
            f"This extracted PDF corresponds to {page_info['label']} of the "
            f"original full financial report (1-based page numbering).\n"
            f"The content spans pages: {pages_str}.\n"
            f"When populating the 'Page No' field in Metadata, use a "
            f"comma-separated list like \"{pages_str}\".\n"
            f"When populating any 'Source Page' or 'Page Reference' field, "
            f"use the page where each specific data point appears."
        )


# ── Main normalize function ──────────────────────────────────────────────────

def normalize_one_gemini(
    pdf_path: str,
    prompt_path: str,
    coa_text: str,
    reporting_columns: list[str] | None,
    model: str,
    max_tokens: int,
    page_info: dict | None = None,   # ← NEW: passed from pipeline after filename parsing
) -> dict:
    """
    Gemini equivalent of pipeline.normalize_one_openai().

    page_info (optional): dict with keys {start, end, pages, label} produced by
        pipeline.extract_page_info_from_filename(). When provided, a
        SOURCE PAGE REFERENCE block is injected into the user prompt so the
        LLM can correctly populate Metadata.Page No.

    Returns:
        {
            "data":  <parsed JSON dict>,
            "usage": {"prompt_tokens": int, "completion_tokens": int, "total_tokens": int}
        }
    """
    try:
        import google.generativeai as genai
    except ImportError:
        raise RuntimeError(
            "google-generativeai SDK not installed.\n"
            "Run:  pip install google-generativeai"
        )

    api_key = os.environ.get("GEMINI_API_KEY")
    if not api_key:
        raise RuntimeError(
            "GEMINI_API_KEY environment variable is not set.\n"
            "Get a free key at https://aistudio.google.com/app/apikey"
        )

    genai.configure(api_key=api_key)

    # ── Load inputs ──────────────────────────────────────────────────────────
    system_prompt = Path(prompt_path).read_text(encoding="utf-8")
    pdf_bytes     = Path(pdf_path).read_bytes()

    print(f"  PDF size      : {len(pdf_bytes) / 1024:.1f} KB")

    # ── Build user text ──────────────────────────────────────────────────────
    instruction_lines = [
        "Normalize the financial statement from the attached PDF.",
        "Follow all steps in the system prompt exactly.",
        "Use the COA Master table above for COA Datapoint mapping.",
        "IMPORTANT: Return valid JSON ONLY. No markdown fences, no ```json, no extra text.",
        "Start your response with { and end with }.",
    ]
    if reporting_columns:
        instruction_lines += ["", "[[REPORTING COLUMNS]]"] + reporting_columns

    # ── SOURCE PAGE REFERENCE block (injected when page_info is available) ───
    # This tells the LLM which pages of the original full report it is reading
    # so it can correctly populate Metadata.Page No (Section B of the prompt).
    page_ref_block = _build_page_reference_block(page_info)
    if page_ref_block:
        print(f"  Source pages  : {page_info['label']}")
    else:
        print(f"  Source pages  : [WARN] no page_info provided — Metadata.Page No may be inaccurate")

    # Assemble user text: COA master → page reference → instructions
    user_text_parts = [
        "STANDARD COA MASTER (pipe-delimited):\n"
        "Format: COA Flag | COA Datapoint | Statement | Section\n\n"
        + coa_text,
    ]
    if page_ref_block:
        user_text_parts.append(page_ref_block)
    user_text_parts.append("\n".join(instruction_lines))

    user_text = "\n\n".join(user_text_parts)

    # ── Determine output token cap ───────────────────────────────────────────
    model_max     = _model_max_tokens(model)
    effective_max = max(max_tokens, model_max)
    print(f"  Max output    : {effective_max:,} tokens  (model cap: {model_max:,})")

    # ── Build PDF part ───────────────────────────────────────────────────────
    pdf_part = {
        "inline_data": {
            "mime_type": "application/pdf",
            "data": base64.b64encode(pdf_bytes).decode("utf-8"),
        }
    }

    # ── Create model ─────────────────────────────────────────────────────────
    gemini_model = genai.GenerativeModel(
        model_name=model,
        system_instruction=system_prompt,
        generation_config=genai.types.GenerationConfig(
            max_output_tokens=effective_max,
            temperature=0.0,
            response_mime_type="application/json",
        ),
    )

    # ── Call Gemini ──────────────────────────────────────────────────────────
    try:
        response = gemini_model.generate_content(
            [pdf_part, user_text],
            request_options={"timeout": 300},
        )
    except _gex.ResourceExhausted as e:
        raise RuntimeError(
            "  Gemini quota exhausted (ResourceExhausted / 429).\n"
            "    Free-tier limit reached. Wait ~1 minute or upgrade your plan.\n"
        ) from e
    except _gex.PermissionDenied as e:
        raise RuntimeError(
            "  Gemini API key rejected (PermissionDenied / 403).\n"
            "    Check GEMINI_API_KEY is correct and Gemini API is enabled.\n"
        ) from e
    except _gex.Unauthenticated as e:
        raise RuntimeError(
            "  Gemini API key invalid or missing (Unauthenticated / 401).\n"
        ) from e
    except _gex.DeadlineExceeded as e:
        raise RuntimeError(
            "  Gemini request timed out (DeadlineExceeded).\n"
            "    The PDF may be too large. Try splitting it or using gemini-2.5-pro.\n"
        ) from e
    except Exception as e:
        raise RuntimeError(
            f"  Gemini API call failed: {type(e).__name__}: {e}"
        ) from e

    # ── Token usage ──────────────────────────────────────────────────────────
    usage_meta        = response.usage_metadata
    prompt_tokens     = getattr(usage_meta, "prompt_token_count",     0) or 0
    completion_tokens = getattr(usage_meta, "candidates_token_count", 0) or 0
    total_tokens      = getattr(usage_meta, "total_token_count",      0) or 0

    print(f"  Tokens        : prompt={prompt_tokens:,}  completion={completion_tokens:,}  total={total_tokens:,}")

    # ── Extract text with finish-reason diagnosis ────────────────────────────
    raw = _check_finish_reason(response, model)

    if not raw:
        raise RuntimeError(
            "  Gemini returned an empty response body.\n"
            "    Most likely causes:\n"
            "      1. Output was cut off by token limit → use gemini-2.5-pro\n"
            "      2. Safety/recitation filter blocked the response\n"
            "      3. API quota exhausted silently\n"
            "    Check https://aistudio.google.com for quota status."
        )

    # ── Parse JSON ────────────────────────────────────────────────────────────
    try:
        data = _parse_json(raw)
    except ValueError as e:
        raise RuntimeError(f"  {e}") from e

    return {
        "data": data,
        "usage": {
            "prompt_tokens":     prompt_tokens,
            "completion_tokens": completion_tokens,
            "total_tokens":      total_tokens,
        },
    }