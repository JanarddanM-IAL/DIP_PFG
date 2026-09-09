"""
openai_batch_client.py
Submits ALL PDFs as a single OpenAI Batch job (50% discount).
"""
import os, json, time, base64
from pathlib import Path
from openai import OpenAI
from compact_schema import build_short_key_instruction

_POLL_SECONDS = 60
_TERMINAL = {"completed", "failed", "expired", "cancelled"}

# ─────────────────────────────────────────────────────────────────────────────
# All batch rates come from openai_model_registry.py (50% of standard by
# default). To add a new model, edit openai_model_registry.py ONLY.
# ─────────────────────────────────────────────────────────────────────────────
from openai_model_registry import calculate_batch_cost as _registry_batch_cost


def _rate_for_model(model: str) -> tuple[float, float]:
    """
    Kept for any legacy callers; returns (batch_in, batch_out) per 1M tokens.
    Prefer calling _calc_cost() directly.
    """
    from openai_model_registry import get_spec
    spec = get_spec(model)
    return spec.effective_batch_in(), spec.effective_batch_out()


def _client():
    return OpenAI(api_key=os.environ["OPENAI_API_KEY"])


def build_page_note(page_info: dict) -> str:
    """
    FIX (root cause of "page number from PDF, not filename" in batch mode):
    pipeline.py's sync path builds this exact instruction via its own
    build_page_note() and appends it to the user message so the LLM is
    explicitly told to copy the filename-derived page number(s) into
    Metadata.Page No verbatim, rather than reading/inferring a page number
    printed inside the PDF body/header/footer. run_normalization_batch()
    in pipeline.py computes page_info per job and puts it in the job dict,
    but the OLD _build_chat_body() here never read it or included any page
    instruction at all — so batch-mode requests had zero page guidance and
    the LLM fell back to whatever number it found printed on the page,
    which is exactly the discrepancy reported.
    Mirrors pipeline.py's build_page_note() so behavior is identical
    between the sync and batch paths.
    """
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
            "the PDF body, footer, or header. Use the value above verbatim, "
            "regardless of what page number(s) appear in the document text."
        )

    return (
        "SOURCE PAGE REFERENCE:\n"
        f"This extracted PDF corresponds to {page_info['label']} of the "
        "original full financial report (1-based page numbering), as "
        "determined from the source filename — NOT from any page number(s) "
        "printed inside the PDF.\n"
        f"The content spans these pages, in order: {pages_str}.\n"
        "Metadata.Page No MUST be set to exactly this comma-joined string "
        "(no spaces), covering the FULL extracted range:\n"
        f"  \"{pages_str}\"\n"
        "Do NOT read, infer, or substitute any page number printed in the "
        "PDF body, footer, or header. Do NOT pick just one page from the "
        "range. Do NOT determine page numbers per data point/row — "
        "Metadata.Page No is a single value for the entire statement. Use "
        "the value above verbatim, regardless of what page number(s) "
        "appear in the document text."
    )


def _is_reasoning_model(model: str) -> bool:
    """
    FIX (root cause of completed=0/N failed=N, tokens in=0/out=0 batches):
    GPT-5-family ("reasoning") models on /v1/chat/completions reject two
    parameters that the old _build_chat_body sent unconditionally for every
    model:
      - "temperature": any value other than the default (1) is rejected
        with 400 "Unsupported value: 'temperature' does not support 0 with
        this model. Only the default (1) value is supported."
      - "max_tokens": no longer accepted at all on these models; the API
        requires "max_completion_tokens" instead and returns 400
        "Unsupported parameter: 'max_tokens' is not supported with this
        model. Use 'max_completion_tokens' instead."
    Both are exactly the kind of request-validation failure that produces
    completed=0/N failed=N with zero token usage in a batch — OpenAI
    rejects the line before any model execution happens, so output_file_id
    ends up empty and only error_file_id has the (previously unread)
    rejection reason. This was silently turning every gpt-5.5 batch job
    into a "MISSING (no result returned from batch)" with no diagnostic.
    Extend this matcher if other non-gpt-5 reasoning models (o1/o3/o4) are
    ever routed through this client.
    """
    m = model.lower()
    return m.startswith("gpt-5") or m.startswith("o1") or m.startswith("o3") or m.startswith("o4")


def _build_chat_body(pdf_path, prompt_text, coa_text, reporting_columns,
                     stmt_type, model, max_tokens, page_info=None,
                     coa_map_rules: str = ""):
    pdf_b64 = base64.b64encode(Path(pdf_path).read_bytes()).decode("utf-8")
    compact_instr = build_short_key_instruction(stmt_type)
    rc_block = ""
    if reporting_columns:
        rc_block = "\nLOCKED REPORTING COLUMNS:\n" + " | ".join(reporting_columns)

    sys_prompt = prompt_text + "\n\n" + compact_instr + rc_block + \
                 "\n\nCOA MASTER (filtered):\n" + coa_text
    if coa_map_rules:                              # ← ADD BLOCK
        sys_prompt = sys_prompt + "\n\n" + coa_map_rules

    user_content = [
        {"type": "file", "file": {
            "filename": Path(pdf_path).name,
            "file_data": f"data:application/pdf;base64,{pdf_b64}",
        }},
    ]

    # FIX: page_info was computed by the caller (run_normalization_batch in
    # pipeline.py) but never actually forwarded into the request body, so
    # the LLM had no filename-derived page instruction in batch mode and
    # fell back to reading/inferring a page number from the PDF itself.
    if page_info:
        user_content.append({"type": "text", "text": build_page_note(page_info)})

    user_content.append(
        {"type": "text", "text": "Extract per the system prompt. Return JSON only."}
    )

    body = {
        "model": model,
        "response_format": {"type": "json_object"},
        "messages": [
            {"role": "system", "content": sys_prompt},
            {"role": "user", "content": user_content},
        ],
    }

    # FIX: reasoning-family models (gpt-5.x, o1/o3/o4) need
    # max_completion_tokens instead of max_tokens, and cannot take a
    # non-default temperature at all. Every other model keeps the exact
    # same body shape as before this fix.
    if _is_reasoning_model(model):
        body["max_completion_tokens"] = max_tokens
    else:
        body["temperature"] = 0
        body["max_tokens"] = max_tokens

    return body


def submit_openai_batch(jobs: list[dict], model: str) -> str:
    """
    jobs same dict shape as Gemini submitter. Returns batch_id.

    Each job dict may include "page_info" (built by
    extract_page_info_from_filename() in pipeline.py) — when present, it is
    forwarded into the request body as an explicit SOURCE PAGE REFERENCE
    instruction (see build_page_note()), so Metadata.Page No is filled from
    the filename-derived page range rather than from whatever page number
    the LLM finds printed inside the PDF.
    """
    cli = _client()
    jsonl_path = Path(f"_oai_batch_in_{int(time.time())}.jsonl")
    with jsonl_path.open("w", encoding="utf-8") as f:
        for j in jobs:
            body = _build_chat_body(
                j["pdf_path"], j["prompt_text"], j["coa_text"],
                j["reporting_columns"], j["stmt_type"],
                model, j["max_tokens"],
                page_info=j.get("page_info"),
                coa_map_rules=j.get("coa_map_rules", ""),   # ← ADD
            )
            f.write(json.dumps({
                "custom_id": j["custom_id"],
                "method": "POST",
                "url": "/v1/chat/completions",
                "body": body,
            }) + "\n")

    print(f"[OPENAI BATCH] Uploading {jsonl_path.name} "
          f"({jsonl_path.stat().st_size/1e6:.1f} MB)")
    up = cli.files.create(file=open(jsonl_path, "rb"), purpose="batch")
    batch = cli.batches.create(
        input_file_id=up.id,
        endpoint="/v1/chat/completions",
        completion_window="24h",
    )
    print(f"[OPENAI BATCH] Created batch: {batch.id}")
    return batch.id


def wait_and_download_openai(batch_id: str, model: str) -> dict[str, dict]:
    """
    Polls until terminal state, downloads output JSONL, returns:
        { custom_id : parsed_dict_or_error }

    Each result dict has shape:
        {
            "ok":        bool,
            "json_text": str,            # present if ok
            "usage":     dict,           # present if ok
            "cost_usd":  float,          # present if ok — per-job cost at batch rate
            "error":     str,            # present if not ok
        }

    FIX (mirrors bug #4 in gemini_batch_client): per-line parsing is now
    wrapped so a single malformed/errored/empty record cannot raise an
    uncaught exception that kills parsing for the ENTIRE batch (i.e. all
    ~40,000 PDFs). Each line is isolated in its own try/except; failures
    are recorded as {"ok": False, "error": ...} for that custom_id only,
    and the loop continues.
    """
    cli = _client()
    while True:
        b = cli.batches.retrieve(batch_id)
        rc = b.request_counts
        print(f"[OPENAI BATCH] status={b.status} "
              f"completed={rc.completed}/{rc.total} failed={rc.failed}")
        if b.status in _TERMINAL:
            break
        time.sleep(_POLL_SECONDS)

    if b.status != "completed":
        raise RuntimeError(f"OpenAI batch ended in {b.status}")

    results: dict[str, dict] = {}
    total_in = total_out = 0
    parse_failures = 0

    # ── FIX (BUG: silent zero-result batches on full request-validation
    # failure): the old code only ever read b.output_file_id. When OpenAI
    # rejects request lines at validation time (e.g. malformed body, wrong
    # endpoint for the model, bad param) — which is exactly what
    # completed=0/N failed=N with tokens in=0/out=0 means — OpenAI puts the
    # per-line rejection reasons in b.error_file_id, and output_file_id is
    # often None. The old code's `if b.output_file_id:` guard then skipped
    # the entire parse loop, leaving results={} with NO indication of why —
    # every PDF in the batch came back "MISSING (no result returned from
    # batch)" downstream with zero diagnostic content, even though OpenAI
    # told us exactly what was wrong with each line.
    # Fixed: fetch error_file_id the same defensive way as output_file_id,
    # and record a real {"ok": False, "error": <reason>} per custom_id from
    # it, so request-validation failures are visible and actionable instead
    # of indistinguishable from "no result returned."
    out_text = ""
    if b.output_file_id:
        out_text = cli.files.content(b.output_file_id).text

    err_text = ""
    if b.error_file_id:
        err_text = cli.files.content(b.error_file_id).text

    if not out_text and not err_text:
        print(f"[OPENAI BATCH] [WARN] batch {batch_id} completed but BOTH "
              f"output_file_id and error_file_id are empty/None — "
              f"{rc.failed} failed / {rc.completed} completed per "
              f"request_counts, but no per-line detail is retrievable. "
              f"This usually means the batch itself was rejected before "
              f"any line was evaluated (e.g. bad input_file_id, wrong "
              f"endpoint for completion_window). Check the OpenAI dashboard "
              f"for batch {batch_id} directly.")

    for line_num, line in enumerate(err_text.strip().splitlines(), start=1):
        if not line.strip():
            continue
        try:
            rec = json.loads(line)
        except Exception as e:
            print(f"[OPENAI BATCH] [WARN] error-file line {line_num}: "
                  f"could not parse JSONL record itself ({e}). Skipping.")
            parse_failures += 1
            continue

        cid = rec.get("custom_id")
        if cid is None:
            print(f"[OPENAI BATCH] [WARN] error-file line {line_num}: "
                  f"record has no 'custom_id' — skipping.")
            parse_failures += 1
            continue

        err_body = rec.get("error") or rec.get("response", {}).get("body", {})
        results[cid] = {
            "ok": False,
            "error": f"[request validation failed] {err_body}",
        }
        print(f"[OPENAI BATCH] [REJECTED] custom_id={cid!r}: {err_body}")

    for line_num, line in enumerate(out_text.strip().splitlines(), start=1):
        if not line.strip():
            continue

        # ── Outer guard: malformed JSONL line itself ───────────────────────
        try:
            rec = json.loads(line)
        except Exception as e:
            print(f"[OPENAI BATCH] [WARN] line {line_num}: could not parse "
                  f"JSONL record itself ({e}). Skipping.")
            parse_failures += 1
            continue

        cid = rec.get("custom_id")
        if cid is None:
            print(f"[OPENAI BATCH] [WARN] line {line_num}: record has no "
                  f"'custom_id' — skipping, result will be MISSING for "
                  f"whichever PDF this was.")
            parse_failures += 1
            continue

        # ── Inner guard: isolate per-record extraction failures ────────────
        try:
            if rec.get("error"):
                results[cid] = {"ok": False, "error": str(rec["error"])}
                continue

            body = rec.get("response", {}).get("body", {})
            if not body:
                results[cid] = {"ok": False, "error": "empty response body"}
                continue

            choices = body.get("choices") or []
            if not choices:
                results[cid] = {
                    "ok": False,
                    "error": "No choices in response body "
                             "(possible content filter or empty completion).",
                }
                continue

            msg = choices[0].get("message", {}).get("content")
            if msg is None:
                finish_reason = choices[0].get("finish_reason")
                results[cid] = {
                    "ok": False,
                    "error": f"No message content (finish_reason="
                             f"{finish_reason!r}). Possibly hit max_tokens "
                             f"with no output or was filtered.",
                }
                continue

            usage = body.get("usage", {}) or {}

            # ── Assign FIRST, then use ────────────────────────────────────
            prompt_tokens     = usage.get("prompt_tokens", 0)
            completion_tokens = usage.get("completion_tokens", 0)
            total_in  += prompt_tokens
            total_out += completion_tokens

            job_cost = _registry_batch_cost(
                model, prompt_tokens, cached_tokens=0,
                completion_tokens=completion_tokens,
            )

            # ── Debug print AFTER assignment (safe) ───────────────────────
            prompt_tokens     = usage.get("prompt_tokens", 0)
            completion_tokens = usage.get("completion_tokens", 0)

            results[cid] = {
                "ok": True,
                "json_text": msg,
                "usage": usage,
                "cost_usd": job_cost,
            }

        except Exception as e:
            parse_failures += 1
            print(f"[OPENAI BATCH] [WARN] line {line_num} (custom_id="
                  f"{cid!r}): unexpected error extracting result "
                  f"({type(e).__name__}: {e}). Marking as failed, "
                  f"continuing with remaining results.")
            results[cid] = {
                "ok": False,
                "error": f"{type(e).__name__}: {e}",
            }

    total_cost = _registry_batch_cost(model, total_in, cached_tokens=0, completion_tokens=total_out)
    print(f"[OPENAI BATCH] tokens in={total_in}  out={total_out}  "
          f"cost≈ ${total_cost:.4f}  "
          f"(results: {len(results)} ok-or-failed, {parse_failures} line-level "
          f"parse failures)")
    return results