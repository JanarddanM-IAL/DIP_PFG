"""
claude_batch_client.py
... (docstring unchanged)
"""

import base64
import hashlib          # ★ FIX 1: moved to top with all other imports
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

from claude_cache_client import (
    _calc_cost,
    _MULTI_TABLE_DELIMITERS,
    _parse_json,
    _extract_text_from_message,
)
# resolve_claude_model and batch cost helpers come from the registry directly.
# To add a new model, edit claude_model_registry.py ONLY.
from claude_model_registry import (
    resolve_claude_model,
    calculate_batch_cost as _registry_batch_cost,
)


_POLL_SECONDS = 60

# ★ FIX 1: _safe_custom_id defined here, before submit_claude_batch uses it
_MAX_CUSTOM_ID_LEN = 64

def _safe_custom_id(raw_id: str) -> str:
    """
    Anthropic enforces a 64-character limit on custom_id.
    Short ids are returned unchanged.
    Long ids: first 32 chars (human-readable) + 32-char SHA-256 hex suffix
    = exactly 64 chars, guaranteed unique.
    """
    if len(raw_id) <= _MAX_CUSTOM_ID_LEN:
        return raw_id
    suffix = hashlib.sha256(raw_id.encode()).hexdigest()[:32]
    return raw_id[:32] + suffix


def _calc_batch_cost(
    model: str,
    input_tokens: int,
    output_tokens: int,
    cache_creation_tokens: int,
    cache_read_tokens: int,
) -> float:
    """Delegates to the registry — no local rate table needed."""
    return _registry_batch_cost(
        model                = resolve_claude_model(model),
        input_tokens         = input_tokens,
        output_tokens        = output_tokens,
        cache_creation_tokens= cache_creation_tokens,
        cache_read_tokens    = cache_read_tokens,
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


# ─────────────────────────────────────────────────────────────────────────────
# _build_request_for_pdf
# ─────────────────────────────────────────────────────────────────────────────

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
    indented_text: str | None = None,      # ← NEW: pre-extracted indented text
) -> dict:
    """
    Build one Anthropic Batch API request dict for a single PDF.

    indented_text is injected into the user text block alongside the page
    reference note, matching the behaviour of normalize_one_claude_cached()
    in the sync path.  The caller passes it from the job dict; this function
    falls back to calling pdf_to_indented_text() only when it is None, so
    the PDF is never opened twice in normal pipeline operation.
    """
    model     = resolve_claude_model(model)
    pdf_bytes = Path(pdf_path).read_bytes()
    pdf_b64   = base64.b64encode(pdf_bytes).decode("utf-8")

    system_text = _build_system_text(
        prompt_text=prompt_text,
        coa_text=coa_text,
        reporting_columns=reporting_columns,
        stmt_type=stmt_type,
        coa_map_rules=coa_map_rules,
    )

    # ── Fallback: extract indented text if not pre-supplied ───────────────
    if indented_text is None:
        try:
            from pdf_to_indented_text import pdf_to_indented_text as _pit
            indented_text = _pit(pdf_path)
        except Exception:
            indented_text = None

    # ── Build user text: page reference + optional indented text ─────────
    page_note = _build_page_reference_block(page_info)
    user_text = (
        "Extract and normalize the attached financial statement PDF. "
        "Return JSON only."
    )
    if page_note:
        user_text += "\n\n" + page_note

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

    return {
        "model":      model,
        "max_tokens": max_tokens,
        "system": [
            {
                "type":          "text",
                "text":          system_text,
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
) -> tuple[str, dict[str, str]]:
    """
    jobs = list of dicts with keys:
        custom_id, pdf_path, prompt_text, coa_text, reporting_columns,
        stmt_type, max_tokens,
        page_info       (optional),
        coa_map_rules   (optional),
        indented_text   (optional) ← NEW: pass pre-extracted text to avoid
                                         a second pdfplumber open per PDF.
    Returns: (batch_id, id_map)  where id_map maps safe_id → original custom_id.
    """
    client = _client()
    model  = resolve_claude_model(model)

    id_map: dict[str, str] = {}
    requests = []
    seen_safe_ids: set[str] = set()

    for j in jobs:
        body = _build_request_for_pdf(
            pdf_path          = j["pdf_path"],
            prompt_text       = j["prompt_text"],
            coa_text          = j["coa_text"],
            reporting_columns = j.get("reporting_columns"),
            stmt_type         = j["stmt_type"],
            max_tokens        = j["max_tokens"],
            page_info         = j.get("page_info"),
            coa_map_rules     = j.get("coa_map_rules", ""),
            model             = model,
            indented_text     = j.get("indented_text"),   # ← NEW
        )
        safe_id = _safe_custom_id(j["custom_id"])

        # ── Guard: skip if this safe_id already seen ──────────────────
        if safe_id in seen_safe_ids:
            print(f"[CLAUDE BATCH] [WARN] Duplicate custom_id skipped: {safe_id}")
            continue
        seen_safe_ids.add(safe_id)
        # ─────────────────────────────────────────────────────────────

        id_map[safe_id] = j["custom_id"]
        requests.append({
            "custom_id": safe_id,
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
    return batch_id, id_map


# ─────────────────────────────────────────────────────────────────────────────
# Poll & download
# ─────────────────────────────────────────────────────────────────────────────

def _get_processing_status(batch) -> str:
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
    active_delimiter = next(
        (d for d in _MULTI_TABLE_DELIMITERS if d in raw), None
    )

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
                    f"  [CLAUDE BATCH] Multi-table parse error on sub-table "
                    f"{idx + 1}: {e}"
                )

        if not prop_snp_split:
            raise ValueError("Delimiter found but no sub-table parsed successfully")

        return {
            "ok":             True,
            "data":           first_data,
            "prop_snp_split": prop_snp_split,
            "usage":          usage_dict,
            "cost_usd":       cost_usd,
            "cached_hit":     cache_read_tokens > 0,
            "model":          model,
        }

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
    id_map: dict[str, str] | None = None,
) -> dict[str, dict]:
    """
    Poll until the batch reaches a terminal state, then download every result.

    id_map (safe_id -> original_id) is used in two places:
      1. To look up stmt_type during parsing (Anthropic echoes safe_ids back).
      2. To restore original custom_id keys in the returned dict so that
         pipeline.py's job_by_id lookups work correctly.
    """
    client = _client()
    model  = resolve_claude_model(model)

    # ★ FIX 3 & 4: build job_stmt_type keyed by SAFE id, not original id.
    # Anthropic returns safe_ids in batch results, so looking up by original
    # id silently returned "" for every truncated filename in the old code.
    job_stmt_type: dict[str, str] = {}
    if jobs and id_map:
        # Reverse id_map: original_id -> safe_id, then map safe_id -> stmt_type
        original_to_safe = {v: k for k, v in id_map.items()}
        for j in jobs:
            safe = original_to_safe.get(j["custom_id"], j["custom_id"])
            job_stmt_type[safe] = j.get("stmt_type", "")
    elif jobs:
        # No id_map means no truncation happened; safe_id == original_id
        for j in jobs:
            job_stmt_type[j["custom_id"]] = j.get("stmt_type", "")

    print(f"[CLAUDE BATCH] Polling batch_id={batch_id}")

    # ── Poll until terminal ──────────────────────────────────────────────────
    status = ""
    while True:
        batch  = client.messages.batches.retrieve(batch_id)
        status = _get_processing_status(batch)
        print(f"[CLAUDE BATCH] status={status}")
        if status in {"ended", "completed", "canceled", "expired", "errored"}:
            break
        time.sleep(_POLL_SECONDS)

    if status not in {"ended", "completed"}:
        raise RuntimeError(f"Claude batch ended with non-success status={status!r}")

    # ── Download & parse results ─────────────────────────────────────────────
    results: dict[str, dict] = {}

    for item in client.messages.batches.results(batch_id):
        custom_id = getattr(item, "custom_id", None)   # this is the safe_id
        result    = getattr(item, "result",    None)

        result_type = getattr(result, "type", None)
        if result_type != "succeeded":
            err = getattr(result, "error", None)
            results[custom_id] = {
                "ok":             False,
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

        response = getattr(result, "message", None)
        raw      = _extract_text_from_message(response)

        usage_obj             = getattr(response, "usage", None)
        input_tokens          = int(getattr(usage_obj, "input_tokens",                0) or 0)
        output_tokens         = int(getattr(usage_obj, "output_tokens",               0) or 0)
        cache_creation_tokens = int(getattr(usage_obj, "cache_creation_input_tokens", 0) or 0)
        cache_read_tokens     = int(getattr(usage_obj, "cache_read_input_tokens",     0) or 0)
        cached_tokens         = cache_creation_tokens + cache_read_tokens
        total_tokens          = input_tokens + output_tokens + cached_tokens

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

        # custom_id here is the safe_id — job_stmt_type is now keyed by safe_id
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

    # Restore original custom_ids so pipeline.py's job_by_id lookups work
    if id_map:
        results = {id_map.get(k, k): v for k, v in results.items()}

    print(f"[CLAUDE BATCH] Downloaded {len(results)} result(s).")
    return results