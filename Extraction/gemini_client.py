import json
import os
import re
from pathlib import Path

from google import genai
from google.genai import types
from google.genai.errors import APIError

# FIX: compact_schema.py exports build_short_key_instruction() and
# expand_compact_json() — NOT "COMPACT_OVERRIDE" / "expand_compact". The
# old import names here didn't exist in compact_schema.py at all, which
# either crashes this module on import, or (if some other shadow copy of
# compact_schema.py existed with those names) fed broken data through.
from compact_schema import build_short_key_instruction, expand_compact_json

# ── Model registry (single source of truth for all rates + token caps) ──────
from gemini_model_registry import (
    calculate_cost,
    max_output_tokens as _model_max_tokens,
    standard_rates,
    GEMINI_COST_TABLE,   # kept for any external callers that reference it
    MODEL_MAX_OUTPUT,    # kept for any external callers that reference it
)


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
                                 # ← nothing here inside the loop
    raise ValueError(            # ← CORRECT: outside the loop, after all attempts
        f"Cannot parse JSON from Gemini response.\n"
        f"Length: {len(raw)} chars\n"
        f"First 600 chars:\n{raw[:600]}\n"
        f"Last  300 chars:\n{raw[-300:]}\n"
        f"Raw response:\n{raw}"
    )


# ── Finish-reason diagnostics ────────────────────────────────────────────────

def _check_finish_reason(response, model: str) -> str:
    """
    Extract text from response, printing a clear diagnosis if empty or cut off.
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

    #print(f"  Finish reason : {finish_reason}")

    if finish_reason in ("MAX_TOKENS", "2", "FinishReason.MAX_TOKENS"):
        print(
            f"    Output was TRUNCATED (hit max_output_tokens limit).\n"
            f"     Model '{model}' max output = {_model_max_tokens(model)} tokens.\n"
            f"     The JSON was cut off mid-way — switching to a larger model\n"
            f"     (gemini-2.5-pro) or splitting the PDF into fewer pages will help.\n"
            f"     Partial output length: {len(raw)} chars"
        )
    elif finish_reason in ("SAFETY", "3", "FinishReason.SAFETY"):
        print(
            "    Response blocked by Gemini SAFETY filter.\n"
            "     The financial PDF may contain content that triggered a filter.\n"
            "     Try gemini-2.5-pro or contact Google support."
        )
    elif finish_reason in ("RECITATION", "4", "FinishReason.RECITATION"):
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
 
        Always instructs the model to copy the filename-derived page string
        verbatim into Metadata.Page No — never to infer it from PDF content,
        and never references a per-row "Source Page"/"Page Reference" field
        (no such field exists in the output schema).
 
        Returns the block string, or None if page_info is falsy.
        """
        if not page_info:
            return None
 
        pages_str = ",".join(str(p) for p in page_info["pages"])
 
        if page_info["start"] == page_info["end"]:
            return (
                "SOURCE PAGE REFERENCE:\n"
                f"This extracted PDF corresponds to {page_info['label']} of the "
                "original full financial report (1-based page numbering), as "
                "determined from the source filename — NOT from any page "
                "number printed inside the PDF.\n"
                "Metadata.Page No MUST be set to exactly this string:\n"
                f"  \"{pages_str}\"\n"
                "Do NOT read, infer, or substitute any page number printed in "
                "the PDF body, footer, or header. Use the value above "
                "verbatim, regardless of what page number(s) appear in the "
                "document text."
            )
 
        return (
            "SOURCE PAGE REFERENCE:\n"
            f"This extracted PDF corresponds to {page_info['label']} of the "
            "original full financial report (1-based page numbering), as "
            "determined from the source filename — NOT from any page "
            "number(s) printed inside the PDF.\n"
            f"The content spans these pages, in order: {pages_str}.\n"
            "Metadata.Page No MUST be set to exactly this comma-joined "
            "string (no spaces), covering the FULL extracted range:\n"
            f"  \"{pages_str}\"\n"
            "Do NOT read, infer, or substitute any page number printed in "
            "the PDF body, footer, or header. Do NOT pick just one page "
            "from the range. Do NOT determine page numbers per data "
            "point/row — Metadata.Page No is a single value for the entire "
            "statement. Use the value above verbatim, regardless of what "
            "page number(s) appear in the document text."
        )

def normalize_one_gemini(
    pdf_path: str,
    prompt_path: str,
    coa_text: str,
    reporting_columns: list[str] | None,
    model: str,
    max_tokens: int,
    page_info: dict | None = None,
    stmt_type: str = "UNKNOWN",
    coa_map_rules: str = "",
    indented_text: str | None = None,       # ← NEW
) -> dict:
    """
    Gemini equivalent of pipeline.normalize_one_openai().

    Returns:
        {
            "data":  <parsed JSON dict>,
            "usage": {"prompt_tokens": int, "completion_tokens": int, "total_tokens": int}
        }
    """
    

    api_key = os.environ.get("GEMINI_API_KEY") or os.environ.get("GOOGLE_API_KEY")
    if not api_key:
        raise RuntimeError(
            "GEMINI_API_KEY (or GOOGLE_API_KEY) environment variable is not set.\n"
            "Get a free key at https://aistudio.google.com/app/apikey"
        )

    # ── Create client ────────────────────────────────────────────────────────
    client = genai.Client(api_key=api_key)

    system_prompt = Path(prompt_path).read_text(encoding="utf-8", errors="replace")
    system_prompt = system_prompt + "\n\n" + build_short_key_instruction(stmt_type)
    if coa_map_rules:                  # ← ADD BLOCK
        system_prompt = system_prompt + "\n\n" + coa_map_rules
    pdf_bytes     = Path(pdf_path).read_bytes()

    #print(f"  PDF size      : {len(pdf_bytes) / 1024:.1f} KB")

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

    # ── SOURCE PAGE REFERENCE block ──────────────────────────────────────────
    page_ref_block = _build_page_reference_block(page_info)
    # if page_ref_block:
    #     print(f"  Source pages  : {page_info['label']}")
    # else:
    #     print(f"  Source pages  : [WARN] no page_info provided — Metadata.Page No may be inaccurate")

    if indented_text is None:
        try:
            from pdf_to_indented_text import pdf_to_indented_text as _pit
            indented_text = _pit(pdf_path)
        except Exception:
            indented_text = None

    # Assemble user text: COA master → page reference → indented text → instructions
    user_text_parts = [
        "STANDARD COA MASTER (pipe-delimited):\n"
        "Format: COA Flag | COA Datapoint | Statement | Section\n\n"
        + coa_text,
    ]
    if page_ref_block:
        user_text_parts.append(page_ref_block)
    if indented_text:
        user_text_parts.append(
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
    user_text_parts.append("\n".join(instruction_lines))
    user_text = "\n\n".join(user_text_parts)

    # ── Build PDF part using new SDK ─────────────────────────────────────────
    pdf_part = types.Part.from_bytes(
        data=pdf_bytes,
        mime_type="application/pdf",
    )

    # ── Inner call function (allows model escalation retry) ──────────────────
    def _call_gemini(use_model: str) -> dict:
        import time

        # Determine output token cap for this model
        model_max     = _model_max_tokens(use_model)
        effective_max = model_max
        #print(f"  Model         : {use_model}")
        #print(f"  Max output    : {effective_max:,} tokens  (model cap: {model_max:,})")

        # Build config with correct cap for this model
        call_config = types.GenerateContentConfig(
            system_instruction=system_prompt,
            max_output_tokens=effective_max,
            temperature=0.0,
            response_mime_type="application/json",
        )

        try:
            response = client.models.generate_content(
                model=use_model,
                contents=[pdf_part, user_text],
                config=call_config,
            )
        except APIError as e:
            error_code = getattr(e, "code", None) or getattr(e, "status", "")
            error_msg  = str(e)

            if "429" in str(error_code) or "RESOURCE_EXHAUSTED" in error_msg.upper():
                raise RuntimeError(
                    "  Gemini quota exhausted / rate limit hit (429).\n"
                    "    Check your quota at https://aistudio.google.com\n"
                ) from e
            elif any(kw in error_msg.upper() for kw in [
                "503", "UNAVAILABLE", "OVERLOADED", "HIGH DEMAND",
                "DEADLINE_EXCEEDED", "INTERNAL",
            ]):
                raise RuntimeError(
                    f"  Gemini server busy / unavailable.\n"
                    f"    Error: {error_msg[:200]}\n"
                ) from e
            elif "403" in str(error_code) or "PERMISSION_DENIED" in error_msg.upper():
                raise RuntimeError(
                    "  Gemini API key rejected (PermissionDenied / 403).\n"
                    "    Check GEMINI_API_KEY is correct and Gemini API is enabled.\n"
                ) from e
            elif "401" in str(error_code) or "UNAUTHENTICATED" in error_msg.upper():
                raise RuntimeError(
                    "  Gemini API key invalid or missing (Unauthenticated / 401).\n"
                ) from e
            else:
                raise RuntimeError(
                    f"  Gemini API call failed: {type(e).__name__}: {e}"
                ) from e
        except Exception as e:
            raise RuntimeError(
                f"  Gemini API call failed: {type(e).__name__}: {e}"
            ) from e

        # ── Token usage ──────────────────────────────────────────────────────
        usage_meta        = response.usage_metadata
        prompt_tokens     = getattr(usage_meta, "prompt_token_count",     0) or 0
        completion_tokens = getattr(usage_meta, "candidates_token_count", 0) or 0
        total_tokens      = getattr(usage_meta, "total_token_count",      0) or 0

        in_rate, out_rate = standard_rates(use_model)
        cost = calculate_cost(use_model, prompt_tokens, completion_tokens)

        #print(f"  Tokens        : prompt={prompt_tokens:,}  completion={completion_tokens:,}  total={total_tokens:,}")
        #print(f"  Cost          : ≈ ${cost:.4f}  (${prompt_tokens/1e6*in_rate:.4f} in + ${completion_tokens/1e6*out_rate:.4f} out)")

        # ── Extract text with finish-reason diagnosis ────────────────────────
        raw = _check_finish_reason(response, use_model)

        if not raw:
            raise RuntimeError(
                "  Gemini returned an empty response body.\n"
                "    Most likely causes:\n"
                "      1. Output was cut off by token limit → use gemini-2.5-pro\n"
                "      2. Safety/recitation filter blocked the response\n"
                "      3. API quota exhausted silently\n"
                "    Check https://aistudio.google.com for quota status."
            )

        # ── PROP_SNP multi-table split detection (BEFORE parse attempt) ──────
        # Import here to avoid circular import at module level
# ── Multi-table split detection (PROP_SNP/IS/CFS) BEFORE parse ──────
        try:
            from pipeline import get_table_break, split_multi_table_response
        except ImportError:
            get_table_break = None
            split_multi_table_response = None

        delimiter = get_table_break(stmt_type) if get_table_break else None

        if (delimiter
                and delimiter in raw
                and split_multi_table_response is not None):
            print(f"  [{stmt_type} SPLIT] Delimiter detected — splitting tables")
            parts = split_multi_table_response(raw, stmt_type)
            split_tables = []
            for idx, part in enumerate(parts):
                try:
                    parsed   = _parse_json(part)
                    expanded = expand_compact_json(parsed, stmt_type=stmt_type)
                    split_tables.append(expanded)
                    print(f"  [{stmt_type} SPLIT] Table {idx + 1} parsed OK "
                          f"({len(expanded.get('Sections', {}))} sections)")
                except Exception as e:
                    print(f"  [{stmt_type} SPLIT] Parse error on table {idx + 1}: {e}")
            return {
                "data":           split_tables[0] if split_tables else None,
                "prop_snp_split": split_tables,
                "raw":            raw,
                "usage": {
                    "prompt_tokens":     prompt_tokens,
                    "completion_tokens": completion_tokens,
                    "total_tokens":      total_tokens,
                },
                "cost_usd": cost,
                "model":    use_model,
            }

        # ── Normal single-table parse ─────────────────────────────────────
        data = _parse_json(raw)
        data = expand_compact_json(data, stmt_type=stmt_type)

        return {
            "data":           data,
            "prop_snp_split": [],
            "raw":            raw,
            "usage": {
                "prompt_tokens":     prompt_tokens,
                "completion_tokens": completion_tokens,
                "total_tokens":      total_tokens,
            },
            "cost_usd": cost,
            "model":    use_model,
        }


    # ── Call with primary model; escalate to gemini-2.5-flash on failure ─────
    FALLBACK_MODEL = "gemini-2.5-flash"

    try:
        return _call_gemini(model)

    except (ValueError, RuntimeError) as primary_err:
        # Only escalate if the fallback is actually a different model
        if model == FALLBACK_MODEL:
            raise RuntimeError(f"  {primary_err}") from primary_err

        print(f"\n  [ESCALATE] Primary model '{model}' failed: {str(primary_err)[:120]}")
        print(f"  [ESCALATE] Retrying with fallback model '{FALLBACK_MODEL}' ...\n")

        try:
            return _call_gemini(FALLBACK_MODEL)
        except (ValueError, RuntimeError) as fallback_err:
            raise RuntimeError(
                f"  Both primary model '{model}' and fallback '{FALLBACK_MODEL}' failed.\n"
                f"  Primary error  : {str(primary_err)[:200]}\n"
                f"  Fallback error : {str(fallback_err)[:200]}\n"
            ) from fallback_err


# ─────────────────────────────────────────────────────────────────────────────
# CONSOLIDATED EXTRACTION — whole PDF → all statement types in one call
# ─────────────────────────────────────────────────────────────────────────────

_CONSOLIDATED_STMT_TYPES = [
    "SNP", "SOA", "GOV_BS", "GOV_IS", "PROP_SNP", "PROP_IS", "PROP_CFS",
    "DSR", "DEBT", "OVERVIEW", "CAPITAL_ASSETS", "TAX_BASE", "PEN", "OPEB", "FAQS",
]


def normalize_consolidated_gemini(
    pdf_path: str,
    prompt_path: str,
    coa_text: str,
    model: str,
    max_tokens: int,
) -> dict:
    """
    Send the WHOLE PDF with the consolidated extraction prompt.

    Returns a dict keyed by statement type:
      {
        "SNP":  { "Metadata":{}, "Reporting Columns":[], "Sections":{} },
        "SOA":  { ... },
        ...
        "FAQS": { ... },
        "_usage": { "prompt_tokens":int, "completion_tokens":int, "total_tokens":int },
        "_cost_usd": float,
        "_model": str,
      }
    Statement types not present in the PDF have value {"NOT_FOUND": True}.
    """
    api_key = os.environ.get("GEMINI_API_KEY") or os.environ.get("GOOGLE_API_KEY")
    if not api_key:
        raise RuntimeError(
            "GEMINI_API_KEY (or GOOGLE_API_KEY) environment variable is not set."
        )

    client = genai.Client(api_key=api_key)

    system_prompt = Path(prompt_path).read_text(encoding="utf-8", errors="replace")
    pdf_bytes     = Path(pdf_path).read_bytes()

    user_text = (
        "Extract ALL financial statement tables from the attached PDF.\n"
        "Follow the consolidated prompt exactly.\n"
        "Return VALID JSON ONLY — no markdown, no fences, no commentary.\n"
        "Your response starts with { and ends with }.\n\n"
        "STANDARD COA MASTER (pipe-delimited):\n"
        "Format: COA Flag | COA Datapoint | Statement | Section\n\n"
        + coa_text
    )

    pdf_part = types.Part.from_bytes(data=pdf_bytes, mime_type="application/pdf")

    model_max     = _model_max_tokens(model)
    effective_max = min(max_tokens, model_max)

    call_config = types.GenerateContentConfig(
        system_instruction=system_prompt,
        max_output_tokens=effective_max,
        temperature=0.0,
        response_mime_type="application/json",
    )

    try:
        response = client.models.generate_content(
            model=model,
            contents=[pdf_part, user_text],
            config=call_config,
        )
    except Exception as e:
        raise RuntimeError(
            f"[CONSOLIDATED] Gemini API call failed: {type(e).__name__}: {e}"
        ) from e

    usage_meta        = response.usage_metadata
    prompt_tokens     = getattr(usage_meta, "prompt_token_count",     0) or 0
    completion_tokens = getattr(usage_meta, "candidates_token_count", 0) or 0
    total_tokens      = getattr(usage_meta, "total_token_count",      0) or 0
    cost              = calculate_cost(model, prompt_tokens, completion_tokens)

    raw = _check_finish_reason(response, model)
    if not raw:
        raise RuntimeError("[CONSOLIDATED] Gemini returned an empty response.")

    combined = _parse_json(raw)

    # Ensure every expected key is present (add NOT_FOUND for missing ones)
    for stype in _CONSOLIDATED_STMT_TYPES:
        if stype not in combined:
            combined[stype] = {"NOT_FOUND": True}

    combined["_usage"]    = {
        "prompt_tokens":     prompt_tokens,
        "completion_tokens": completion_tokens,
        "total_tokens":      total_tokens,
    }
    combined["_cost_usd"] = cost
    combined["_model"]    = model

    return combined