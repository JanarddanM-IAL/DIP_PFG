"""
claude_batch_client.py

Submits all PDFs as an Anthropic Claude Message Batch job.

Batch mode:
- async
- approximately 50% discount
- paid tier only
- uses same prompt/content shaping as sync Claude client

Return shape of wait_and_download_claude():
{
    "<custom_id>": {
        "ok":             bool,
        "data":           dict | None,      # first (or only) parsed+expanded table
        "prop_snp_split": list[dict],       # all tables for multi-table responses
        "usage":          {...},
        "cost_usd":       float,
        "cached_hit":     bool,
        "model":          str,
        "error":          str | None,       # present only when ok=False
    },
    ...
}

NOTE: Unlike the OpenAI/Gemini batch clients which return raw "json_text",
this client returns already-parsed "data" dicts.  pipeline.py's Claude batch
branch reads "data" directly and skips the json_text → parse_json step.
"""

import base64
import json
import os
import re
import time
from pathlib import Path

try:
    from anthropic import Anthropic
except ImportError as e:
    raise ImportError(
        "anthropic package not installed. Run: pip install anthropic"
    ) from e

from compact_schema import build_short_key_instruction, expand_compact_json

# Import shared helpers from the sync client so behaviour stays consistent.
from claude_cache_client import (
    resolve_claude_model,
    _calc_cost,             # FIX #3: was called but never imported
    _MULTI_TABLE_DELIMITERS,
    _parse_json,
    _extract_text_from_message,
)


_POLL_SECONDS = 60

# Batch = 50% of normal rates  ($ per 1M tokens)
CLAUDE_BATCH_RATES = {
    "claude-sonnet-4-6": (1.50,  7.50),
    "claude-sonnet-4-5": (1.50,  7.50),
    "claude-opus-4-8":   (2.50, 12.50),
    "claude-opus-4-5":   (2.50, 12.50),
}

_DEFAULT_BATCH_RATE = (1.50, 7.50)


def _batch_rate_for_model(model: str) -> tuple[float, float]:
    """Return batch-discounted (input, output) rate in $/1M tokens."""
    model = resolve_claude_model(model)
    if model in CLAUDE_BATCH_RATES:
        return CLAUDE_BATCH_RATES[model]
    print(
        f"[CLAUDE BATCH] [WARN] No rate entry for model={model!r}. "
        f"Using fallback rate {_DEFAULT_BATCH_RATE} ($/1M in,out)."
    )
    return _DEFAULT_BATCH_RATE


def _calc_batch_cost(
    model: str,
    input_tokens: int,
    output_tokens: int,
    cache_creation_tokens: int,
    cache_read_tokens: int,
) -> float:
    """
    Cost calculation using batch-discounted rates.

    Batch pricing is ~50% of normal; cache multipliers still apply on top:
    - cache write: 1.25x the batch input rate
    - cache read:  0.10x the batch input rate
    """
    in_rate, out_rate = _batch_rate_for_model(model)

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


def _build_page_reference_block(page_info: dict | None) -> str | None:
    if not page_info:
        return None

    pages_str = ",".join(str(p) for p in page_info.get("pages", []))
    label     = page_info.get("label") or pages_str

    return (
        f"SOURCE PAGE REFERENCE - MANDATORY:\n"
        f"The attached extracted PDF was created from source page(s): {pages_str}.\n"
        f'For Metadata.Page No, you MUST copy exactly this value: "{label}".\n'
        f"Do NOT infer page number from any printed footer/header inside the PDF."
    )


def _build_system_text(
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


def _build_request_for_pdf(
    pdf_path,
    prompt_text,
    coa_text,
    reporting_columns,
    stmt_type,
    max_tokens,
    page_info=None,
    coa_map_rules: str = "",
    model: str = "claude-sonnet-4-6",
) -> dict:
    model = resolve_claude_model(model)

    pdf_bytes = Path(pdf_path).read_bytes()
    pdf_b64   = base64.b64encode(pdf_bytes).decode("utf-8")

    system_text = _build_system_text(
        prompt_text=prompt_text,
        coa_text=coa_text,
        reporting_columns=reporting_columns,
        stmt_type=stmt_type,
        coa_map_rules=coa_map_rules,
    )

    page_note = _build_page_reference_block(page_info)
    user_text = "Extract and normalize the attached financial statement PDF. Return JSON only."
    if page_note:
        user_text += "\n\n" + page_note

    return {
        "model":      model,
        "max_tokens": max_tokens,
        "system": [
            {
                "type":          "text",
                "text":          system_text,
                # Cache breakpoint kept here — if Anthropic applies caching
                # in batch mode, usage will reflect it.
                "cache_control": {"type": "ephemeral"},
            }
        ],
        "messages": [
            {
                "role": "user",
                "content": [
                    {
                        "type": "document",
                        "source": {
                            "type":       "base64",
                            "media_type": "application/pdf",
                            "data":       pdf_b64,
                        },
                    },
                    {
                        "type": "text",
                        "text": user_text,
                    },
                ],
            }
        ],
    }


# ─────────────────────────────────────────────────────────────────────────────
# Submit
# ─────────────────────────────────────────────────────────────────────────────

def submit_claude_batch(
    jobs: list[dict],
    model: str,
    display_name: str = "fs-claude-batch",
) -> str:
    """
    Build and submit a Claude Message Batch.

    jobs item shape:
    {
        custom_id, pdf_path, prompt_text, coa_text,
        reporting_columns, stmt_type, max_tokens,
        page_info, coa_map_rules
    }

    Returns batch_id (str).
    """
    client = _client()
    model  = resolve_claude_model(model)

    requests = []
    for j in jobs:
        body = _build_request_for_pdf(
            pdf_path=j["pdf_path"],
            prompt_text=j["prompt_text"],
            coa_text=j["coa_text"],
            reporting_columns=j.get("reporting_columns"),
            stmt_type=j["stmt_type"],
            max_tokens=j["max_tokens"],
            page_info=j.get("page_info"),
            coa_map_rules=j.get("coa_map_rules", ""),
            model=model,
        )
        requests.append({
            "custom_id": j["custom_id"],
            "params":    body,
        })

    print(f"[CLAUDE BATCH] Submitting {len(requests)} request(s), model={model}")
    batch = client.messages.batches.create(requests=requests)

    batch_id = getattr(batch, "id", None)
    if not batch_id:
        raise RuntimeError(
            f"Claude batch created but no batch id returned: {batch}"
        )

    print(f"[CLAUDE BATCH] Submitted batch_id={batch_id}")
    return batch_id


# ─────────────────────────────────────────────────────────────────────────────
# Poll & download
# ─────────────────────────────────────────────────────────────────────────────

def _get_processing_status(batch) -> str:
    """Normalise across SDK versions that expose different attribute names."""
    return (
        getattr(batch, "processing_status", None)
        or getattr(batch, "status", None)
        or ""
    )


def _parse_one_item(
    raw: str,
    stmt_type: str,
    model: str,
    cost_usd: float,
    usage_dict: dict,
    cache_read_tokens: int,
) -> dict:
    """
    Parse the raw text from one batch result item.
    Handles both single-table and multi-table (delimiter-separated) responses.
    Always returns a fully-formed result dict with ok=True.
    Raises on unrecoverable parse failure.
    """
    active_delimiter = next(
        (d for d in _MULTI_TABLE_DELIMITERS if d in raw), None
    )

    if active_delimiter:
        # ── Multi-table path ─────────────────────────────────────────────────
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
                    f"  [CLAUDE BATCH] Multi-table parse error on sub-table "
                    f"{idx + 1}: {e}"
                )

        if not prop_snp_split:
            raise ValueError(
                "Delimiter found but no sub-table parsed successfully"
            )

        return {
            "ok":             True,
            "data":           first_data,
            "prop_snp_split": prop_snp_split,
            "usage":          usage_dict,
            "cost_usd":       cost_usd,
            "cached_hit":     cache_read_tokens > 0,
            "model":          model,
        }

    # ── Normal single-table path ─────────────────────────────────────────────
    # expand_compact_json is intentionally called here with stmt_type so that
    # compact keys are resolved before the result is handed back to pipeline.py.
    # pipeline.py's re-expansion block (after wait_and_download_claude returns)
    # is a no-op on already-expanded data, so double-calling is safe.
    data = _parse_json(raw)
    data = expand_compact_json(data, stmt_type)

    return {
        "ok":             True,
        "data":           data,
        "prop_snp_split": [],
        "usage":          usage_dict,
        "cost_usd":       cost_usd,
        "cached_hit":     cache_read_tokens > 0,
        "model":          model,
    }


def wait_and_download_claude(
    batch_id: str,
    model: str,
    jobs: list[dict] | None = None,
) -> dict[str, dict]:
    """
    Poll until the batch reaches a terminal state, then download every result.

    Parameters
    ----------
    batch_id : str
        ID returned by submit_claude_batch().
    model : str
        Model string used for cost calculation when per-item model is absent.
    jobs : list[dict] | None
        Optional job list (same objects passed to submit_claude_batch).
        Used to resolve stmt_type per custom_id for accurate expand_compact_json
        calls.  If None, stmt_type defaults to "" for all items.

    Returns
    -------
    dict[custom_id, result_dict]
        Every item in the batch has an entry.  Shape documented at top of file.
    """
    client = _client()
    model  = resolve_claude_model(model)

    # Build a lookup table so we can find stmt_type for each custom_id.
    # FIX #2: stmt_type was undefined during result processing.
    job_stmt_type: dict[str, str] = {}
    if jobs:
        for j in jobs:
            job_stmt_type[j["custom_id"]] = j.get("stmt_type", "")

    print(f"[CLAUDE BATCH] Polling batch_id={batch_id}")

    # ── Poll until terminal ──────────────────────────────────────────────────
    status = ""
    while True:
        batch  = client.messages.batches.retrieve(batch_id)
        status = _get_processing_status(batch)
        print(f"[CLAUDE BATCH] status={status}")

        # Anthropic terminal statuses: ended, canceled, expired, errored
        # (older SDK versions may use "completed")
        if status in {"ended", "completed", "canceled", "expired", "errored"}:
            break

        time.sleep(_POLL_SECONDS)

    if status not in {"ended", "completed"}:
        raise RuntimeError(f"Claude batch ended with non-success status={status!r}")

    # ── Download & parse results ─────────────────────────────────────────────
    # FIX #1: the original code had `return` inside the for-loop, meaning only
    # the FIRST item was ever returned.  We now accumulate into `results` dict
    # and return after the loop completes.

    results: dict[str, dict] = {}

    for item in client.messages.batches.results(batch_id):
        custom_id = getattr(item, "custom_id", None)
        result    = getattr(item, "result",    None)

        # ── Failed item ──────────────────────────────────────────────────────
        result_type = getattr(result, "type", None)
        if result_type != "succeeded":
            err = getattr(result, "error", None)
            results[custom_id] = {
                "ok":             False,        # FIX #7: ok key always present
                "error":          f"Claude batch item failed: {err}",
                "data":           None,
                "prop_snp_split": [],
                "usage": {
                    "prompt_tokens":         0,
                    "completion_tokens":     0,
                    "total_tokens":          0,
                    "cached_tokens":         0,
                    "cache_creation_tokens": 0,
                    "cache_read_tokens":     0,
                },
                "cost_usd":   0.0,
                "cached_hit": False,
                "model":      model,
            }
            continue

        # ── Succeeded item ───────────────────────────────────────────────────
        response = getattr(result, "message", None)
        raw      = _extract_text_from_message(response)

        usage_obj             = getattr(response, "usage", None)
        input_tokens          = int(getattr(usage_obj, "input_tokens",                0) or 0)
        output_tokens         = int(getattr(usage_obj, "output_tokens",               0) or 0)
        cache_creation_tokens = int(getattr(usage_obj, "cache_creation_input_tokens", 0) or 0)
        cache_read_tokens     = int(getattr(usage_obj, "cache_read_input_tokens",     0) or 0)
        cached_tokens         = cache_creation_tokens + cache_read_tokens
        total_tokens          = input_tokens + output_tokens + cached_tokens

        # FIX #3 & #4: use _calc_batch_cost (not the imported _calc_cost which
        # uses normal rates) so the 50% batch discount is correctly applied.
        # _rate_for_model / in_rate / out_rate are no longer computed separately
        # (they were computed but never used in the original code).
        cost_usd = _calc_batch_cost(
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

        # Resolve stmt_type for this item so expand_compact_json uses the
        # correct schema.  FIX #2: was undefined / "" in original code.
        stmt_type = job_stmt_type.get(custom_id, "")

        try:
            item_result = _parse_one_item(
                raw=raw,
                stmt_type=stmt_type,
                model=model,
                cost_usd=cost_usd,
                usage_dict=usage_dict,
                cache_read_tokens=cache_read_tokens,
            )
        except Exception as e:
            item_result = {
                "ok":             False,
                "error":          f"Parse error: {e}",
                "data":           None,
                "prop_snp_split": [],
                "usage":          usage_dict,
                "cost_usd":       cost_usd,
                "cached_hit":     cache_read_tokens > 0,
                "model":          model,
            }

        results[custom_id] = item_result

    print(f"[CLAUDE BATCH] Downloaded {len(results)} result(s).")
    return results