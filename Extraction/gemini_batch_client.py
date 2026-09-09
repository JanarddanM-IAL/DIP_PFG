"""
gemini_batch_client.py
Submits ALL PDFs as a single Gemini Batch job (50% discount).
Free tier is NOT supported by Batch API — caller must guard.
"""
import os, json, time, base64
from pathlib import Path
from google import genai
from google.genai.errors import APIError
from compact_schema import build_short_key_instruction, expand_compact_json

_POLL_SECONDS = 60
_TERMINAL = {"JOB_STATE_SUCCEEDED", "JOB_STATE_FAILED",
             "JOB_STATE_CANCELLED", "JOB_STATE_EXPIRED"}

# ── All rate data comes from the single registry ──────────────────────────────
from gemini_model_registry import (
    batch_rates,
    standard_rates,
    calculate_batch_cost,
    GEMINI_BATCH_RATES,    # backwards-compat alias (read-only)
    GEMINI_STANDARD_RATES, # backwards-compat alias (read-only)
)


def _rate_for_model(model: str) -> tuple[float, float]:
    """Return batch (in, out) rates for *model* via the central registry."""
    return batch_rates(model)


def _client():
    api_key = os.environ.get("GEMINI_API_KEY")
    if not api_key:
        raise RuntimeError("GEMINI_API_KEY not set")
    return genai.Client(api_key=api_key)


def wait_and_download_gemini(batch_name: str, model: str) -> dict[str, dict]:
    """
    Polls until terminal state, downloads output JSONL, returns:
        { custom_id : parsed_dict_or_error }

    Each result dict has shape:
        {
            "ok":        bool,
            "json_text": str,       # present if ok
            "usage":     dict,      # present if ok
            "cost_usd":  float,     # present if ok — per-job cost at batch rate
            "error":     str,       # present if not ok
            "raw":       str|None,  # present if not ok and we have partial text
        }

    Token field names in the Gemini Batch JSONL output (REST/camelCase):
        promptTokenCount       → input tokens
        candidatesTokenCount   → output tokens  (NOT responseTokenCount)
        totalTokenCount
    The fallback chain below handles both camelCase and snake_case variants
    so the client is robust to SDK version differences.
    """
    client = _client()
    _MAX_404_RETRIES = 5
    _404_retry_count = 0

    while True:
        try:
            b = client.batches.get(name=batch_name)
            _404_retry_count = 0
        except APIError as e:
            if e.code == 404:
                _404_retry_count += 1
                print(f"[GEMINI BATCH] [WARN] 404 on poll attempt "
                      f"({_404_retry_count}/{_MAX_404_RETRIES}) — "
                      f"retrying in {_POLL_SECONDS}s ...")
                if _404_retry_count >= _MAX_404_RETRIES:
                    raise RuntimeError(
                        f"Gemini batch job '{batch_name}' returned 404 "
                        f"on {_MAX_404_RETRIES} consecutive polls."
                    ) from e
                time.sleep(_POLL_SECONDS)
                continue
            raise

        state = b.state.name
        print(f"[GEMINI BATCH] state={state}")
        if state in _TERMINAL:
            break
        time.sleep(_POLL_SECONDS)

    if state != "JOB_STATE_SUCCEEDED":
        raise RuntimeError(f"Gemini batch ended in {state}")

    out_file = b.dest.file_name
    raw = client.files.download(file=out_file).decode("utf-8")

    in_rate, out_rate = _rate_for_model(model)

    results: dict[str, dict] = {}
    total_in = total_out = 0
    parse_failures = 0

    for line_num, line in enumerate(raw.splitlines(), start=1):
        if not line.strip():
            continue

        try:
            rec = json.loads(line)
        except Exception as e:
            print(f"[GEMINI BATCH] [WARN] line {line_num}: could not parse "
                  f"JSONL record ({e}). Skipping.")
            parse_failures += 1
            continue

        cid = rec.get("key")
        if cid is None:
            print(f"[GEMINI BATCH] [WARN] line {line_num}: record has no "
                  f"'key' — skipping.")
            parse_failures += 1
            continue

        try:
            if "response" not in rec:
                err = rec.get("error", "unknown error (no 'response' field)")
                results[cid] = {"ok": False, "error": str(err)}
                continue

            resp = rec["response"]

            candidates = resp.get("candidates") or []
            if not candidates:
                finish_reason = None
                try:
                    finish_reason = resp.get("promptFeedback", {}).get("blockReason")
                except Exception:
                    pass
                results[cid] = {
                    "ok": False,
                    "error": (f"No candidates returned "
                              f"(blockReason={finish_reason!r}). "
                              f"Possible safety block or empty response."),
                    "raw": json.dumps(resp)[:2000],
                }
                continue

            parts = candidates[0].get("content", {}).get("parts") or []
            if not parts or "text" not in parts[0]:
                results[cid] = {
                    "ok": False,
                    "error": ("Candidate had no text part (possibly hit "
                              "MAX_TOKENS with no output, or non-text part)."),
                    "raw": json.dumps(candidates[0])[:2000],
                }
                continue

            text = parts[0]["text"]

            finish_reason = candidates[0].get("finishReason", "")
            if finish_reason == "MAX_TOKENS":
                results[cid] = {
                    "ok":   False,
                    "error": (f"Response truncated (finishReason=MAX_TOKENS). "
                              f"Increase --max-tokens. "
                              f"Last 200 chars: {text[-200:]!r}"),
                    "raw":  text,
                }
                parse_failures += 1
                continue

            usage = resp.get("usageMetadata") or resp.get("usage_metadata") or {}

            # Fallback chain: camelCase (REST/JSONL) then snake_case (SDK objects).
            # candidatesTokenCount is the correct output field in the JSONL;
            # responseTokenCount is an SDK alias that also works.
            prompt_tokens = int(
                usage.get("promptTokenCount",   0)
                or usage.get("prompt_token_count", 0) or 0
            )
            completion_tokens = int(
                usage.get("candidatesTokenCount",   0)
                or usage.get("candidates_token_count", 0)
                or usage.get("responseTokenCount",     0)
                or usage.get("response_token_count",   0) or 0
            )

            # Debug sample: log first 3 records so field-name issues are visible
            # if len(results) < 3:
            #     print(f"[GEMINI BATCH] [DEBUG] key={cid!r}  "
            #           f"usage_fields={list(usage.keys())}  "
            #           f"prompt_tok={prompt_tokens}  "
            #           f"completion_tok={completion_tokens}")

            total_in  += prompt_tokens
            total_out += completion_tokens

            job_cost = calculate_batch_cost(model, prompt_tokens, completion_tokens)

            results[cid] = {
                "ok":        True,
                "json_text": text,
                "usage":     usage,
                "cost_usd":  job_cost,
            }

        except Exception as e:
            parse_failures += 1
            print(f"[GEMINI BATCH] [WARN] line {line_num} (key={cid!r}): "
                  f"unexpected error ({type(e).__name__}: {e}). Marking failed.")
            results[cid] = {
                "ok":   False,
                "error": f"{type(e).__name__}: {e}",
                "raw":  line[:2000],
            }

    if total_in == 0 and results:
        n_ok = sum(1 for r in results.values() if r.get("ok"))
        print(f"[GEMINI BATCH] [WARN] All per-line token counts are 0 "
              f"({n_ok} successful results). The Gemini Batch API may not "
              f"be populating usageMetadata in the output JSONL for this "
              f"model/region. Cost will show as $0.0000. Check the Google "
              f"Cloud console for actual billing.")

    total_cost = calculate_batch_cost(model, total_in, total_out)
    # print(f"[GEMINI BATCH] tokens in={total_in:,}  out={total_out:,}  "
    #       f"cost≈ ${total_cost:.4f}  "
    #       f"(results: {len(results)} ok-or-failed, "
    #       f"{parse_failures} line-level parse failures)")
    return results


def _build_page_reference_block(page_info: dict | None) -> str | None:
    """
    Builds the SOURCE PAGE REFERENCE instruction block from filename-derived
    page info. Always instructs the model to copy the page string verbatim
    into Metadata.Page No — never to infer it from PDF content.
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


def _build_request_for_pdf(pdf_path, prompt_text, coa_text,
                            reporting_columns, stmt_type, max_tokens,
                            page_info=None,
                            coa_map_rules: str = "") -> dict:
    """Build one GenerateContentRequest for a single PDF."""
    pdf_bytes = Path(pdf_path).read_bytes()
    pdf_b64   = base64.b64encode(pdf_bytes).decode("utf-8")
    compact_instr = build_short_key_instruction(stmt_type)

    rc_block = ""
    if reporting_columns:
        rc_block = "\nLOCKED REPORTING COLUMNS:\n" + " | ".join(reporting_columns)

    page_ref_block = _build_page_reference_block(page_info)
    page_ref_text  = ("\n\n" + page_ref_block) if page_ref_block else ""

    sys_text = (
        prompt_text + "\n\n" + compact_instr + rc_block +
        "\n\nCOA MASTER (filtered):\n" + coa_text +
        page_ref_text
    )
    if coa_map_rules:
        sys_text = sys_text + "\n\n" + coa_map_rules

    return {
        "contents": [{
            "role": "user",
            "parts": [
                {"text": sys_text},
                {"inline_data": {"mime_type": "application/pdf", "data": pdf_b64}},
            ],
        }],
        "generationConfig": {
            "temperature": 0,
            "max_output_tokens": max_tokens,
            "response_mime_type": "application/json",
        },
    }


def submit_gemini_batch(jobs: list[dict], model: str,
                        display_name: str = "fs-batch") -> str:
    """
    jobs = list of {custom_id, pdf_path, prompt_text, coa_text,
                    reporting_columns, stmt_type, max_tokens,
                    page_info (optional), coa_map_rules (optional)}
    Returns: batch_job_name (str)
    """
    client = _client()
    jsonl_path = Path(f"_batch_in_{int(time.time())}.jsonl")
    with jsonl_path.open("w", encoding="utf-8") as f:
        for j in jobs:
            req = _build_request_for_pdf(
                j["pdf_path"], j["prompt_text"], j["coa_text"],
                j["reporting_columns"], j["stmt_type"], j["max_tokens"],
                page_info=j.get("page_info"),
                coa_map_rules=j.get("coa_map_rules", ""),
            )
            f.write(json.dumps({"key": j["custom_id"], "request": req}) + "\n")

    print(f"[GEMINI BATCH] Uploading {jsonl_path.name} "
          f"({jsonl_path.stat().st_size/1e6:.1f} MB)")
    uploaded = client.files.upload(
        file=str(jsonl_path),
        config={"mime_type": "application/jsonl"},
    )
    batch = client.batches.create(
        model=model,
        src=uploaded.name,
        config={"display_name": display_name},
    )
    print(f"[GEMINI BATCH] Created job: {batch.name}")
    return batch.name