"""
llm_page_identifier.py
──────────────────────
Page identification routes to Gemini, Claude, or OpenAI depending on
the model string passed via --id-model.

  gemini-*     → google-genai  (PDF binary + Context Cache for static prompt)
  claude-*     → TEXT-BASED extraction + system-prompt caching (ephemeral)
  gpt-* / o1-* → openai SDK   (PDF binary + automatic prefix caching)

CACHING STRATEGY PER PROVIDER
──────────────────────────────
Claude  : Static instruction prompt → system block with cache_control=ephemeral.
          Per-PDF paged text → user message (never cached, always unique).
          Cache TTL = 5 min (ephemeral). Hits from 2nd PDF onward.

Gemini  : Static instruction prompt → google.genai CachedContent object.
          Cached once per process lifetime; reused for all subsequent PDFs.
          Minimum 32 k tokens required; falls back to plain call if shorter.
          Cache TTL = 60 min (configurable via GEMINI_CACHE_TTL_MINUTES).

OpenAI  : Automatic prefix caching — no explicit API call needed.
          The static system prompt is always the same prefix, so OpenAI
          caches it automatically after the first call.
          Cache hit shows as cached_tokens > 0 in usage.prompt_tokens_details.
"""

import json
import re
import os
import threading
from pathlib import Path


# ════════════════════════════════════════════════════════════════════════
# MODULE-LEVEL CACHE STORES
# ════════════════════════════════════════════════════════════════════════

# Gemini: one CachedContent object per (model, prompt_text) pair.
# Key = (model_str, prompt_hash); Value = CachedContent resource name.
_gemini_cache_lock    = threading.Lock()
_gemini_cache_store: dict[tuple, str] = {}   # (model, hash) → cache_name

# OpenAI: nothing to store — caching is automatic.
# Claude: nothing to store — cache_control is per-request, TTL managed by Anthropic.

GEMINI_CACHE_TTL_MINUTES = 60   # adjust if needed


# ════════════════════════════════════════════════════════════════════════
# STRUCTURAL VALIDATORS
# ════════════════════════════════════════════════════════════════════════

_DSR_YEAR_ROW_RE = re.compile(
    r"^\s*(20[2-9]\d)(?:\s*[-–]\s*20[2-9]\d)?\s+"
    r"(?:\$\s*)?[\d,]+"
    r"|^\s*(20[2-9]\d)(?:\s*[-–]\s*20[2-9]\d)?\s*$",
    re.MULTILINE,
)
_PROP_SNP_VAL_TITLE_RE = re.compile(
    r"\bbalance\s+sheets?\b"
    r"|\bstatement\s+of\s+(?:fund\s+)?net\s+(?:position|assets)\b"
    r"|\bstatement\s+of\s+financial\s+position\b",
    re.IGNORECASE,
)
_PROP_IS_VAL_TITLE_RE = re.compile(
    r"statements?\s+of\s+revenues?\s*(?:,|\s+and)\s+(?:expenses|expenditures)\b"
    r"|statements?\s+of\s+operations\b"
    r"|statements?\s+of\s+activities\b"
    r"|statements?\s+of\s+comprehensive\s+income\b"
    r"|income\s+and\s+expenditure\s+(?:account|statement)\b"
    r"|statements?\s+of\s+financial\s+activities\b"
    r"|statements?\s+of\s+changes?\s+in\s+reserves?\b"
    r"|statements?\s+of\s+changes?\s+in\s+equity\b",
    re.IGNORECASE,
)
_PROP_IS_BODY_TRIO = [
    re.compile(r"\bOperating\s+revenues?\b",  re.IGNORECASE),
    re.compile(r"\bOperating\s+expenses?\b",  re.IGNORECASE),
    re.compile(r"\bOperating\s+income\b",     re.IGNORECASE),
]
_PROP_IS_BODY_TRIO_FRS102 = [
    re.compile(r"\bTotal\s+income\b",                          re.IGNORECASE),
    re.compile(r"\bTotal\s+expenditure\b"
               r"|\bStaff\s+costs\b",                          re.IGNORECASE),
    re.compile(r"\bSurplus\s+(?:for|before)\s+(?:the\s+)?year\b"
               r"|\bDeficit\s+(?:for|before)\s+(?:the\s+)?year\b"
               r"|\bTotal\s+comprehensive\s+income\b",         re.IGNORECASE),
]
_PROP_IS_BODY_FALLBACK = [
    re.compile(r"\bOperating\s+(?:income|loss)\b",                         re.IGNORECASE),
    re.compile(r"\bTotal\s+operating\s+revenues?\b",                       re.IGNORECASE),
    re.compile(r"\bTotal\s+operating\s+expenses?\b",                       re.IGNORECASE),
    re.compile(r"\bNon[-\s]?operating\s+(?:revenues?|expenses?|income)\b", re.IGNORECASE),
    re.compile(r"\bChange\s+in\s+net\s+(?:position|assets)\b",             re.IGNORECASE),
]
_PROP_CFS_VAL_TITLE_RE = re.compile(
    r"statements?\s+of\s+cash\s+flows?\b",
    re.IGNORECASE,
)
_BOND_SERIES_ROW_RE = re.compile(
    r"^\s*20\d\d[A-Z]?\s+\d{1,2}/\d{1,2}/\d{2,4}\b",
    re.MULTILINE,
)
_DSR_PRINCIPAL_COL_RE = re.compile(
    r"\bPrincipal\b"
    r"|P[\s]*r[\s]*i[\s]*n[\s]*c[\s]*i[\s]*p[\s]*a[\s]*l"
    r"|\bPrincip\b",
    re.IGNORECASE,
)
_DSR_INTEREST_COL_RE = re.compile(
    r"\bInterest\b(?:\s*\(\d+\))?"
    r"|I[\s]*n[\s]*t[\s]*e[\s]*r[\s]*e[\s]*s[\s]*t",
    re.IGNORECASE,
)
_SWAP_TERMS_RE = re.compile(
    r"\bNotional\b|\bCounty\s+Pays\b|\bCounty\s+Receives\b"
    r"|\bFair\s+Value\b|\bSwap\s+#\b|\bSwap\s+Description\b"
    r"|\bAssociated\s+Variable\s+Rate\b",
    re.IGNORECASE,
)
_TABLE_KEY_TO_SUFFIX = {
    "SNP":      "_SNP",
    "SOA":      "_SOA",
    "GOV_BS":   "_GOV_BS",
    "GOV_IS":   "_GOV_IS",
    "PROP_SNP": "_PROP_SNP",
    "PROP_IS":  "_PROP_IS",
    "PROP_CFS": "_PROP_CFS",
}
_DEBT_ROW_RE = re.compile(
    r"^\s*(?:General\s+obligation\s+bonds?|Revenue\s+bonds?|Special\s+assessment"
    r"|Notes?\s+payable|Lease\s+(?:liability|financed)|SBITA\s+liabilit"
    r"|Subscription[\s-]+based\s+liabilit"
    r"|Compensated\s+absences|Claims\s+payable|Plus\s+premiums?|Discounts?"
    r"|Net\s+pension|OPEB\s+liabilit|Total\s+(?:bonds?|long.term|liabilit)"
    r"|Revenue\s+notes?|Subscription\b|Workers.{0,15}compensation)",
    re.IGNORECASE | re.MULTILINE,
)
_DEBT_OPENING_RE = re.compile(
    r"\bBeginning\s+Balance\b|\bBalance\s+at\b|\bRestated\s+Balance\b"
    r"|\bBalance\s+(?:July|June|January|August|September|October|November"
    r"|December|February|March|April)\b",
    re.IGNORECASE,
)
_DEBT_CHANGE_RE = re.compile(
    r"\bAdditions?\b|\bReductions?\b|\bIssuances?\b|\bRetirements?\b"
    r"|\bMaturities\b|\bIncreases?\b|\bDecreases?\b|\bIssued\b|\bRetired\b",
    re.IGNORECASE,
)


# ════════════════════════════════════════════════════════════════════════
# PROVIDER DETECTION
# ════════════════════════════════════════════════════════════════════════

def _provider_for_model(model: str) -> str:
    """Return 'gemini', 'claude', or 'openai' based on the model string."""
    m = model.lower()
    if m.startswith("gemini"):
        return "gemini"
    if m.startswith("claude"):
        return "claude"
    return "openai"


# ════════════════════════════════════════════════════════════════════════
# CHAR-SPACING COLLAPSE
# ════════════════════════════════════════════════════════════════════════

def _collapse_char_spacing(text: str) -> str:
    """
    Collapse character-level spacing introduced by some PDF renderers.

    Pass 1 — line-level: collapse lines where ≥50% tokens are 1-2 chars.
    Pass 2 — word-group: collapse spaced-out word runs within mixed lines.
    Pass 3 — cross-line: join a spaced fragment at end of line N with a
             spaced fragment at start of line N+1.
    """
    lines = text.splitlines()
    out = []

    for line in lines:
        tokens = line.split(" ")
        non_empty = [t for t in tokens if t.strip()]

        if not non_empty:
            out.append(line)
            continue

        short = sum(1 for t in non_empty if len(t) <= 2)
        ratio = short / len(non_empty)

        if ratio >= 0.50 and len(non_empty) >= 4:
            leading = len(line) - len(line.lstrip())
            collapsed = line[:leading] + "".join(non_empty)
            out.append(collapsed)
            continue

        result_tokens = []
        i = 0
        while i < len(tokens):
            t = tokens[i]
            if not t.strip():
                result_tokens.append(t)
                i += 1
                continue
            if len(t.strip()) <= 2:
                run = [t.strip()]
                j = i + 1
                while j < len(tokens) and len(tokens[j].strip()) <= 2:
                    if tokens[j].strip():
                        run.append(tokens[j].strip())
                    j += 1
                if len(run) >= 3:
                    result_tokens.append("".join(run))
                    i = j
                else:
                    result_tokens.append(t)
                    i += 1
            else:
                result_tokens.append(t)
                i += 1

        out.append(" ".join(result_tokens))

    # Pass 3: cross-line fragment joining
    def _is_spaced_fragment(s: str) -> bool:
        s = s.strip()
        if not s:
            return False
        toks = s.split()
        if not toks:
            return False
        short = sum(1 for t in toks if len(t) <= 2)
        return len(toks) >= 2 and short / len(toks) >= 0.75

    def _last_word(s: str) -> str:
        parts = s.rstrip().split()
        return parts[-1] if parts else ""

    def _first_word(s: str) -> str:
        parts = s.lstrip().split()
        return parts[0] if parts else ""

    merged = []
    i = 0
    while i < len(out):
        line = out[i]

        if i + 1 < len(out):
            next_line = out[i + 1]
            last  = _last_word(line)
            first = _first_word(next_line)

            if (last and first
                    and len(last) <= 2
                    and len(first) <= 2
                    and _is_spaced_fragment(line.rstrip().rsplit(None, 3)[-1] if line.strip() else "")
                    and _is_spaced_fragment(next_line.lstrip().split(None, 1)[0] if next_line.strip() else "")):

                curr_tokens = line.rstrip().split()
                next_tokens = next_line.lstrip().split()

                frag_start = len(curr_tokens)
                for k in range(len(curr_tokens) - 1, -1, -1):
                    if len(curr_tokens[k]) <= 2:
                        frag_start = k
                    else:
                        break

                frag_end = 0
                for k in range(len(next_tokens)):
                    if len(next_tokens[k]) <= 2:
                        frag_end = k + 1
                    else:
                        break

                left_frag   = "".join(curr_tokens[frag_start:])
                right_frag  = "".join(next_tokens[:frag_end])
                joined_word = left_frag + right_frag

                new_curr = " ".join(curr_tokens[:frag_start] + [joined_word])
                new_next = " ".join(next_tokens[frag_end:])

                merged.append(new_curr)
                out[i + 1] = new_next
                i += 1
                continue

        merged.append(line)
        i += 1

    return "\n".join(merged)


# ════════════════════════════════════════════════════════════════════════
# TEXT EXTRACTION HELPERS  (Claude text-based path)
# ════════════════════════════════════════════════════════════════════════

_MAX_CHARS_PER_PAGE = 2000


def _extract_pdf_as_paged_text(pdf_path: str) -> tuple[str, int]:
    """
    Extract text from every page with pdfplumber (pypdf fallback).
    Returns (paged_text, total_pages).
    Each page wrapped with <<<PAGE N>>> / <<<END PAGE N>>> markers.
    """
    pages_out: list[str] = []

    try:
        import pdfplumber
        with pdfplumber.open(pdf_path) as pdf:
            total = len(pdf.pages)
            for i, page in enumerate(pdf.pages):
                text = _collapse_char_spacing((page.extract_text() or "").strip())

                if not text or len(text) < 30:
                    try:
                        from pypdf import PdfReader
                        reader = PdfReader(pdf_path)
                        text = (reader.pages[i].extract_text() or "").strip()
                    except Exception:
                        pass

                if not text:
                    continue

                if len(text) > _MAX_CHARS_PER_PAGE:
                    text = text[:_MAX_CHARS_PER_PAGE] + "\n...[truncated]"

                pages_out.append(
                    f"<<<PAGE {i + 1}>>>\n{text}\n<<<END PAGE {i + 1}>>>"
                )
    except Exception as e:
        raise RuntimeError(f"Text extraction failed for {pdf_path}: {e}")

    if not pages_out:
        raise RuntimeError(f"No text could be extracted from {pdf_path}")

    return "\n\n".join(pages_out), total


def _build_user_content_from_paged_text(paged_text: str, total_pages: int) -> str:
    """
    Wrap paged text in the standard format note.
    This is the PER-PDF user message (never cached).
    """
    return (
        f"DOCUMENT FORMAT NOTE\n"
        f"{'=' * 60}\n"
        f"The following is text extracted from a {total_pages}-page government\n"
        f"financial report (CAFR/ACFR). Each page is delimited by:\n"
        f"  <<<PAGE N>>>     — start of page N\n"
        f"  <<<END PAGE N>>> — end of page N\n\n"
        f"N is the 1-based physical page number in the original PDF.\n"
        f"Return these N values directly in your JSON — they ARE the correct\n"
        f"physical page numbers. Do not add or subtract any offset.\n"
        f"{'=' * 60}\n\n"
        f"{paged_text}\n\n"
        f"{'=' * 60}\n"
        f"END OF DOCUMENT  ({total_pages} pages total)\n"
        f"{'=' * 60}\n"
    )


# ════════════════════════════════════════════════════════════════════════
# CLAUDE — system-prompt caching (ephemeral, 5-min TTL)
# ════════════════════════════════════════════════════════════════════════

def _call_claude_text(pdf_path: str, model: str, prompt: str,
                      max_tokens: int = 8192) -> str:
    """
    Send extracted page text to Claude using prompt caching.

    Caching layout:
      system  = static instruction prompt  ← cache_control: ephemeral
                                             (same for every PDF → cache HIT from 2nd PDF)
      user    = format_note + paged_text   ← unique per PDF, never cached

    Why text instead of PDF binary:
      • No 100-page hard limit (CAFRs are often 150-400 pages)
      • ~5-10x fewer tokens → ~10x cheaper
      • pdfplumber extraction is already tuned for CAFRs
      • Works with every Claude model, no vision quota consumed
    """
    import anthropic
    import time

    api_key = os.environ.get("ANTHROPIC_API_KEY", "")
    if not api_key:
        try:
            import winreg
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, "Environment") as key:
                api_key, _ = winreg.QueryValueEx(key, "ANTHROPIC_API_KEY")
                os.environ["ANTHROPIC_API_KEY"] = api_key
        except Exception:
            pass
    if not api_key:
        raise RuntimeError("ANTHROPIC_API_KEY not set")

    paged_text, total_pages = _extract_pdf_as_paged_text(pdf_path)
    user_content = _build_user_content_from_paged_text(paged_text, total_pages)

    client = anthropic.Anthropic(api_key=api_key)

    for attempt in range(3):
        try:
            response = client.messages.create(
                model=model,
                max_tokens=max_tokens,
                temperature=0.0,
                # ── STATIC prompt → system block with cache_control ──────
                # Claude caches this after the first call (TTL = 5 min).
                # Every subsequent PDF with the same prompt is a cache HIT,
                # saving ~90% of the prompt-token cost.
                system=[
                    {
                        "type": "text",
                        "text": prompt,
                        "cache_control": {"type": "ephemeral"},
                    }
                ],
                # ── PER-PDF content → user message (never cached) ────────
                messages=[
                    {"role": "user", "content": user_content}
                ],
            )

            # ── Log cache performance ──────────────────────────────────
            usage         = response.usage
            cache_read    = getattr(usage, "cache_read_input_tokens",     0) or 0
            cache_created = getattr(usage, "cache_creation_input_tokens", 0) or 0
            input_tok     = getattr(usage, "input_tokens",  0) or 0
            output_tok    = getattr(usage, "output_tokens", 0) or 0

            if cache_read > 0:
                print(f"  [ID-CACHE] Claude HIT  — "
                      f"cache_read={cache_read:,}  input={input_tok:,}  "
                      f"output={output_tok:,}")
            elif cache_created > 0:
                print(f"  [ID-CACHE] Claude MISS (wrote cache) — "
                      f"cache_created={cache_created:,}  input={input_tok:,}  "
                      f"output={output_tok:,}")
            else:
                print(f"  [ID-CACHE] Claude NO-CACHE — "
                      f"input={input_tok:,}  output={output_tok:,}")

            raw = "".join(
                block.text for block in response.content
                if hasattr(block, "text")
            ).strip()

            if not raw:
                print(f"  [ID] Claude attempt {attempt + 1}: empty — retrying")
                time.sleep(2 ** attempt)
                continue
            return raw

        except Exception as e:
            print(f"  [ID] Claude attempt {attempt + 1} failed: {e}")
            if attempt < 2:
                time.sleep(2 ** attempt)

    raise RuntimeError("Claude text mode: all retries exhausted")


def _call_claude(pdf_bytes: bytes, model: str, prompt: str,
                 max_tokens: int = 8192,
                 pdf_path: str | None = None) -> str:
    """
    Entry point for Claude calls.
    Prefers pdf_path (so pdfplumber works directly).
    Falls back to writing pdf_bytes to a temp file.
    pdf_bytes is kept for signature compatibility but NOT sent to the API.
    """
    import tempfile

    if pdf_path and os.path.isfile(pdf_path):
        return _call_claude_text(pdf_path, model, prompt, max_tokens)

    with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp:
        tmp.write(pdf_bytes)
        tmp_path = tmp.name
    try:
        return _call_claude_text(tmp_path, model, prompt, max_tokens)
    finally:
        try:
            os.unlink(tmp_path)
        except Exception:
            pass


# ════════════════════════════════════════════════════════════════════════
# GEMINI — Context Cache for static prompt
# ════════════════════════════════════════════════════════════════════════

def _get_or_create_gemini_cache(client, model: str, prompt: str) -> str | None:
    """
    Return the name of a CachedContent object for (model, prompt).

    Creates a new cache on first call; returns the existing name on
    subsequent calls within the same process.

    Returns None if the prompt is too short for caching (Gemini requires
    ≥ 32 k tokens; we fall back to plain calls in that case).

    Cache layout:
      contents = [prompt text]        ← static instruction (cached)
      Per-PDF paged text is appended as a normal user message at call time.
    """
    import hashlib

    prompt_hash = hashlib.md5(prompt.encode()).hexdigest()
    cache_key   = (model, prompt_hash)

    with _gemini_cache_lock:
        if cache_key in _gemini_cache_store:
            return _gemini_cache_store[cache_key]

    # ── Estimate token count (rough: 4 chars ≈ 1 token) ──────────────
    estimated_tokens = len(prompt) // 4
    if estimated_tokens < 32_000:
        # Prompt too short — Gemini context cache requires ≥ 32 k tokens.
        # Fall back to plain call; no error, just no caching.
        return None

    try:
        from google.genai import types as gtypes
        import datetime

        ttl_seconds = GEMINI_CACHE_TTL_MINUTES * 60

        cached = client.caches.create(
            model=model,
            config=gtypes.CreateCachedContentConfig(
                contents=[
                    gtypes.Content(
                        role="user",
                        parts=[gtypes.Part.from_text(text=prompt)],
                    )
                ],
                ttl=f"{ttl_seconds}s",
                display_name=f"id-prompt-{prompt_hash[:8]}",
            ),
        )
        cache_name = cached.name
        print(f"  [ID-CACHE] Gemini cache CREATED — "
              f"name={cache_name}  ttl={GEMINI_CACHE_TTL_MINUTES}min  "
              f"~{estimated_tokens:,} tokens")

        with _gemini_cache_lock:
            _gemini_cache_store[cache_key] = cache_name

        return cache_name

    except Exception as e:
        print(f"  [ID-CACHE] Gemini cache creation failed ({e}) — "
              f"falling back to plain call")
        return None


def _call_gemini(client, pdf_bytes: bytes, model: str, prompt: str,
                 max_tokens: int = 8192) -> str:
    """
    Call Gemini with context caching for the static prompt.

    If a CachedContent exists for this prompt, the call uses
    cached_content= so Gemini reads the prompt from cache.
    The per-PDF binary is appended as a live Part.

    Falls back to the plain (non-cached) path if caching is unavailable
    (prompt too short, API error, etc.).
    """
    from google.genai import types as gtypes
    import time

    cache_name = _get_or_create_gemini_cache(client, model, prompt)

    for attempt in range(3):
        try:
            if cache_name:
                # ── CACHED path ──────────────────────────────────────
                # The static prompt lives in the cache; only the PDF
                # binary is sent as a live Part.
                response = client.models.generate_content(
                    model=model,
                    contents=[
                        gtypes.Part.from_bytes(
                            data=pdf_bytes, mime_type="application/pdf"
                        ),
                        gtypes.Part.from_text(
                            text="Identify the financial statement pages as instructed."
                        ),
                    ],
                    config=gtypes.GenerateContentConfig(
                        max_output_tokens=max_tokens,
                        temperature=0.0,
                        cached_content=cache_name,
                    ),
                )
            else:
                # ── PLAIN path (prompt too short or cache creation failed) ──
                response = client.models.generate_content(
                    model=model,
                    contents=[
                        gtypes.Part.from_bytes(
                            data=pdf_bytes, mime_type="application/pdf"
                        ),
                        gtypes.Part.from_text(text=prompt),
                    ],
                    config=gtypes.GenerateContentConfig(
                        max_output_tokens=max_tokens,
                        temperature=0.0,
                    ),
                )

            # ── Log cache usage ───────────────────────────────────────
            try:
                meta = response.usage_metadata
                cached_tok = getattr(meta, "cached_content_token_count", 0) or 0
                total_tok  = getattr(meta, "total_token_count", 0) or 0
                if cache_name:
                    status = "HIT" if cached_tok > 0 else "MISS"
                    print(f"  [ID-CACHE] Gemini {status} — "
                          f"cached_tokens={cached_tok:,}  total={total_tok:,}")
            except Exception:
                pass

            raw = (response.text or "").strip()
            if not raw:
                print(f"  [ID] Gemini attempt {attempt + 1}: empty — retrying")
                time.sleep(2 ** attempt)
                continue
            return raw

        except Exception as e:
            print(f"  [ID] Gemini attempt {attempt + 1} failed: {e}")
            if attempt < 2:
                time.sleep(2 ** attempt)

    raise RuntimeError("Gemini: all retries exhausted")


def cleanup_gemini_id_caches():
    """
    Delete all Gemini CachedContent objects created by this module.
    Call at pipeline shutdown (optional; they expire automatically).
    """
    with _gemini_cache_lock:
        items = list(_gemini_cache_store.items())

    if not items:
        return

    # We need a client to delete — reconstruct from env key.
    try:
        from google import genai
        api_key = os.environ.get("GEMINI_API_KEY", "")
        if not api_key:
            return
        client = genai.Client(api_key=api_key)
        for (model, h), cache_name in items:
            try:
                client.caches.delete(name=cache_name)
                print(f"  [ID-CACHE] Gemini cache DELETED — {cache_name}")
            except Exception as e:
                print(f"  [ID-CACHE] Gemini cache delete failed ({e}): {cache_name}")
    except Exception:
        pass

    with _gemini_cache_lock:
        _gemini_cache_store.clear()


# ════════════════════════════════════════════════════════════════════════
# OPENAI — automatic prefix caching (no explicit API calls needed)
# ════════════════════════════════════════════════════════════════════════

def _call_openai(pdf_bytes: bytes, model: str, prompt: str,
                 max_tokens: int = 8192) -> str:
    """
    Call OpenAI with the PDF encoded as base64.

    Caching: OpenAI caches identical message prefixes automatically.
    The static instruction prompt is always the system message; the
    per-PDF PDF binary is the user message.  After the first call,
    the system prompt prefix is cached and subsequent calls pay ~50%
    of the input-token price for that prefix.

    Cache hits appear as usage.prompt_tokens_details.cached_tokens > 0.
    No extra API calls or objects are needed.
    """
    import openai
    import base64
    import time

    api_key = os.environ.get("OPENAI_API_KEY", "")
    if not api_key:
        try:
            import winreg
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, "Environment") as key:
                api_key, _ = winreg.QueryValueEx(key, "OPENAI_API_KEY")
                os.environ["OPENAI_API_KEY"] = api_key
        except Exception:
            pass
    if not api_key:
        raise RuntimeError("OPENAI_API_KEY not set")

    client  = openai.OpenAI(api_key=api_key)
    pdf_b64 = base64.standard_b64encode(pdf_bytes).decode("utf-8")

    for attempt in range(3):
        try:
            response = client.chat.completions.create(
                model=model,
                max_tokens=max_tokens,
                temperature=0.0,
                messages=[
                    # ── STATIC system prompt (same for all PDFs) ─────────
                    # OpenAI automatically caches identical system prefixes.
                    # No cache_control parameter needed — it's implicit.
                    {
                        "role": "system",
                        "content": prompt,
                    },
                    # ── PER-PDF user message (unique, never cached) ───────
                    {
                        "role": "user",
                        "content": [
                            {
                                "type": "text",
                                "text": (
                                    "The following is a base64-encoded PDF of a government "
                                    "financial report. Read it carefully and answer the task.\n\n"
                                    f"[PDF BASE64 START]\n{pdf_b64}\n[PDF BASE64 END]"
                                ),
                            }
                        ],
                    },
                ],
            )

            # ── Log cache performance ──────────────────────────────────
            try:
                usage       = response.usage
                details     = getattr(usage, "prompt_tokens_details", None)
                cached_tok  = getattr(details, "cached_tokens", 0) or 0
                input_tok   = getattr(usage, "prompt_tokens",     0) or 0
                output_tok  = getattr(usage, "completion_tokens", 0) or 0

                if cached_tok > 0:
                    print(f"  [ID-CACHE] OpenAI HIT  — "
                          f"cached={cached_tok:,}  input={input_tok:,}  "
                          f"output={output_tok:,}")
                else:
                    print(f"  [ID-CACHE] OpenAI MISS — "
                          f"input={input_tok:,}  output={output_tok:,}")
            except Exception:
                pass

            raw = (response.choices[0].message.content or "").strip()
            if not raw:
                print(f"  [ID] OpenAI attempt {attempt + 1}: empty — retrying")
                time.sleep(2 ** attempt)
                continue
            return raw

        except Exception as e:
            print(f"  [ID] OpenAI attempt {attempt + 1} failed: {e}")
            if attempt < 2:
                time.sleep(2 ** attempt)

    raise RuntimeError("OpenAI: all retries exhausted")


# ════════════════════════════════════════════════════════════════════════
# DISPATCH
# ════════════════════════════════════════════════════════════════════════

def _dispatch_llm_call(pdf_bytes: bytes, model: str, prompt: str,
                       gemini_client=None, max_tokens: int = 8192,
                       pdf_path: str | None = None) -> str:
    """Route to the correct provider and return raw response text."""
    provider = _provider_for_model(model)

    if provider == "gemini":
        if gemini_client is None:
            raise RuntimeError("Gemini client not initialised")
        return _call_gemini(gemini_client, pdf_bytes, model, prompt, max_tokens)

    if provider == "claude":
        return _call_claude(pdf_bytes, model, prompt, max_tokens,
                            pdf_path=pdf_path)

    return _call_openai(pdf_bytes, model, prompt, max_tokens)


def _parse_json_from_raw(raw: str) -> dict:
    """Strip markdown fences and parse JSON."""
    m = re.search(r"```(?:json)?\s*([\s\S]+?)\s*```", raw)
    if m:
        raw = m.group(1)
    else:
        raw = re.sub(r"^```(?:json)?\s*", "", raw).strip()
    return json.loads(raw)


# ════════════════════════════════════════════════════════════════════════
# STRUCTURAL BODY VALIDATORS
# ════════════════════════════════════════════════════════════════════════

def _looks_like_prop_is_body(text: str) -> bool:
    if not text:
        return False
    trio_hits = sum(1 for p in _PROP_IS_BODY_TRIO if p.search(text))
    if trio_hits == 3:
        return True
    fallback_hits = sum(1 for p in _PROP_IS_BODY_FALLBACK if p.search(text))
    if fallback_hits >= 3:
        return True
    frs102_hits = sum(1 for p in _PROP_IS_BODY_TRIO_FRS102 if p.search(text))
    if frs102_hits >= 2:
        return True
    return False


def _has_statement_title_and_amounts(text: str, title_re,
                                      min_amounts: int = 5) -> bool:
    if not text:
        return False
    first15 = "\n".join(text.splitlines()[:15])
    if not title_re.search(first15):
        return False
    amounts = re.findall(r"\b\d{1,3}(?:,\d{3})+\b", text)
    return len(amounts) >= min_amounts


# ════════════════════════════════════════════════════════════════════════
# SECTOR PROMPT TRIMMING
# ════════════════════════════════════════════════════════════════════════

def _trim_prompt_for_sector(prompt_text: str,
                             allowed_suffixes: set[str] | None) -> str:
    if allowed_suffixes is None:
        return prompt_text

    needed  = {k for k, s in _TABLE_KEY_TO_SUFFIX.items() if s in allowed_suffixes}
    skipped = {k for k, s in _TABLE_KEY_TO_SUFFIX.items() if s not in allowed_suffixes}

    if not skipped:
        return prompt_text

    needed_list  = ", ".join(sorted(needed))
    skipped_list = ", ".join(sorted(skipped))

    json_lines = ["{"]
    for key in ["SNP", "SOA", "GOV_BS", "GOV_IS", "PROP_SNP", "PROP_IS", "PROP_CFS"]:
        if key in needed:
            json_lines.append(f'  "{key}": [...],')
    json_lines.append('  "DSR":  [],')
    json_lines.append('  "DEBT": []')
    json_lines.append("}")
    trimmed_json = "\n".join(json_lines)

    sector_override = (
        f"\n\n"
        f"════════════════════════════════════════════════════════════════\n"
        f"SECTOR OVERRIDE — ONLY FIND THESE TABLES\n"
        f"════════════════════════════════════════════════════════════════\n"
        f"For this document, you only need to find: {needed_list}\n"
        f"DO NOT search for or return pages for: {skipped_list}\n"
        f"These tables do not exist in this type of document.\n"
        f"Searching for them wastes time — skip them entirely.\n\n"
        f"════════════════════════════════════════════════════════════════\n"
        f"CRITICAL — STATEMENT NAMING VARIES BY ACCOUNTING FRAMEWORK\n"
        f"════════════════════════════════════════════════════════════════\n"
        f"This document may use UK GAAP / FRS 102 / SORP / HE sector\n"
        f"terminology instead of US GASB proprietary-fund names.\n"
        f"You MUST accept these alternative titles:\n\n"
        f"  PROP_SNP (Balance Sheet equivalent) — accept ANY of:\n"
        f"    • 'Balance Sheet'\n"
        f"    • 'Consolidated Balance Sheet'\n"
        f"    • 'Consolidated and [Entity] Balance Sheets'\n"
        f"    • 'Statement of Financial Position'\n"
        f"    • 'Consolidated Statement of Financial Position'\n"
        f"  Confirm by rows: Fixed assets / Current assets / Creditors /\n"
        f"    Net assets / Total reserves / Restricted reserves\n\n"
        f"  PROP_IS (Income Statement equivalent) — accept ANY of:\n"
        f"    • 'Statement of Comprehensive Income'\n"
        f"    • 'Consolidated and [Entity] Statement of Comprehensive Income'\n"
        f"    • 'Consolidated Statement of Comprehensive Income'\n"
        f"    • 'Income and Expenditure Account'\n"
        f"    • 'Consolidated Income and Expenditure Account'\n"
        f"    • 'Statement of Financial Activities'\n"
        f"  Confirm by rows: Total income / Staff costs / Total expenditure /\n"
        f"    Surplus for the year / Total comprehensive income for the year\n\n"
        f"  PROP_CFS (Cash Flow equivalent) — accept ANY of:\n"
        f"    • 'Statement of Cash Flows'\n"
        f"    • 'Consolidated Statement of Cash Flows'\n"
        f"    • 'Cash Flow Statement'\n\n"
        f"DO NOT reject a page just because it lacks the words\n"
        f"'Proprietary Funds', 'Enterprise Funds', 'Net Position',\n"
        f"'Revenues', or 'Expenses' in the title.\n"
        f"UK documents NEVER use these GASB terms.\n\n"
        f"════════════════════════════════════════════════════════════════\n"
        f"MANDATORY RULE — UK PROP_IS ALWAYS SPANS TWO PAGES\n"
        f"════════════════════════════════════════════════════════════════\n"
        f"OVERRIDE: The earlier instruction saying 'Statement of Changes in\n"
        f"Reserves is NOT PROP_IS' applies ONLY to standalone occurrences.\n"
        f"When it appears IMMEDIATELY AFTER the income statement, you MUST\n"
        f"include it. This OVERRIDE takes priority over all earlier rules.\n\n"
        f"In UK/FRS 102 documents, PROP_IS ALWAYS spans exactly two pages:\n"
        f"  Page N   → income statement (Comprehensive Income / I&E Account)\n"
        f"  Page N+1 → 'Statement of Changes in Reserves' or\n"
        f"             'Statement of Changes in Equity'\n\n"
        f"STEP-BY-STEP for PROP_IS:\n"
        f"  1. Find the page titled 'Statement of Comprehensive Income'\n"
        f"     (or equivalent UK title listed above)\n"
        f"     → add page N to PROP_IS\n"
        f"  2. Check page N+1\n"
        f"  3. If page N+1 is titled 'Statement of Changes in Reserves'\n"
        f"     OR 'Statement of Changes in Equity' OR similar\n"
        f"     → add page N+1 to PROP_IS as well\n"
        f"  4. STOP after N+1 — do not add any further pages\n\n"
        f"CORRECT:   PROP_IS: [N, N+1]\n"
        f"INCORRECT: PROP_IS: [N]        ← missing the Changes in Reserves page\n"
        f"INCORRECT: PROP_IS: [N, N+1, N+2] ← too many pages\n"
        f"════════════════════════════════════════════════════════════════\n\n"
        f"Return ONLY this JSON structure (omit the skipped keys):\n"
        f"{trimmed_json}\n"
        f"════════════════════════════════════════════════════════════════\n"
    )

    return prompt_text + sector_override


# ════════════════════════════════════════════════════════════════════════
# PAGE TEXT HELPERS
# ════════════════════════════════════════════════════════════════════════

def _get_page_text(pdf_path: str, page_no: int) -> str:
    """Get page text with character-spacing collapse applied."""
    raw = ""
    try:
        import pdfplumber
        with pdfplumber.open(pdf_path) as pdf:
            if page_no < 1 or page_no > len(pdf.pages):
                return ""
            raw = pdf.pages[page_no - 1].extract_text() or ""
    except Exception:
        pass

    if not raw:
        try:
            from pypdf import PdfReader
            reader = PdfReader(pdf_path)
            if page_no < 1 or page_no > len(reader.pages):
                return ""
            raw = reader.pages[page_no - 1].extract_text() or ""
        except Exception:
            return ""

    return _collapse_char_spacing(raw)


# ════════════════════════════════════════════════════════════════════════
# DSR / DEBT PAGE VALIDATORS
# ════════════════════════════════════════════════════════════════════════

def _page_has_dsr_table(text: str) -> bool:
    if not text:
        return False

    text = _collapse_char_spacing(text)

    if _SWAP_TERMS_RE.search(text):
        year_dollar_rows = _DSR_YEAR_ROW_RE.findall(text)
        bond_series_rows = _BOND_SERIES_ROW_RE.findall(text)
        if not year_dollar_rows or len(bond_series_rows) >= len(year_dollar_rows):
            return False

    year_dollar_rows = _DSR_YEAR_ROW_RE.findall(text)

    _YEAR_ALONE_RE = re.compile(
        r"^\s*(20[2-9]\d(?:\s*[-–]\s*20[2-9]\d)?)\s*$",
        re.MULTILINE,
    )
    standalone_years = _YEAR_ALONE_RE.findall(text)
    has_amounts      = len(re.findall(r"\b\d{1,3}(?:,\d{3})+\b", text)) >= 3

    year_check_passed = (
        len(year_dollar_rows) >= 1
        or (len(standalone_years) >= 2 and has_amounts)
    )

    if not year_check_passed:
        return False

    if not _DSR_PRINCIPAL_COL_RE.search(text):
        print(f"    [DSR-DBG] FAILED: no Principal keyword")
        return False

    interest_matches  = list(_DSR_INTEREST_COL_RE.finditer(text))
    has_interest_col  = False
    for m in interest_matches:
        after = text[m.end():m.end() + 10].strip()
        if not after.lower().startswith("rate"):
            has_interest_col = True
            break
    if not has_interest_col:
        print(f"    [DSR-DBG] FAILED: no Interest column")
        return False

    return True


def _page_has_debt_table(text: str) -> bool:
    if not text:
        return False
    has_liability_rows = bool(_DEBT_ROW_RE.search(text))
    has_change_cols    = bool(_DEBT_CHANGE_RE.search(text))
    balance_count = len(re.findall(r"\bBalance\b", text, re.IGNORECASE))
    has_balance = (
        balance_count >= 2
        or bool(re.search(
            r"\bBeginning\s+Balance\b|\bRestated\s+Balance\b",
            text, re.IGNORECASE
        ))
        or any(
            len(re.findall(r"\bBalance\b", line, re.IGNORECASE)) >= 2
            for line in text.splitlines()
        )
    )
    if has_liability_rows and has_change_cols and has_balance:
        return True
    has_balance_beginning = bool(re.search(
        r"\bBalance\s+beginning\b|\bBalance\s+at\s+beginning\b"
        r"|\bBalance\s+beginning\s+of\s+year\b",
        text, re.IGNORECASE
    ))
    has_balance_end = bool(re.search(
        r"\bBalance\s+end\b|\bBalance\s+at\s+end\b"
        r"|\bBalance\s+end\s+of\s+year\b|\bEnding\s+[Bb]alance\b",
        text, re.IGNORECASE
    ))
    has_due_within = bool(re.search(
        r"\bDue\s+within\s+one\s+year\b|\bCurrent\s+portion\b",
        text, re.IGNORECASE
    ))
    has_liability_names = bool(_DEBT_ROW_RE.search(text))
    if (has_balance_beginning and has_balance_end
            and has_liability_names and has_due_within):
        return True
    return False


def _validate_dsr_pages(pdf_path: str, llm_pages: list) -> list:
    validated = []
    for p in llm_pages:
        text = _get_page_text(pdf_path, p)
        if _page_has_dsr_table(text):
            validated.append(p)
        else:
            print(f"  [VAL] DSR p{p} REMOVED — no genuine DSR table found")
    return validated


def _validate_debt_pages(pdf_path: str, llm_pages: list) -> list:
    validated = []
    for p in llm_pages:
        text = _get_page_text(pdf_path, p)
        if _page_has_debt_table(text):
            validated.append(p)
        else:
            print(f"  [VAL] DEBT p{p} REMOVED — no genuine DEBT table found")
    return validated


def _recover_missed_dsr_pages(pdf_path: str, validated_dsr: list,
                               validated_debt: list,
                               scan_range_extra: int = 25) -> list:
    recovered = set(validated_dsr)
    all_candidate_starts = set(validated_debt)
    if validated_dsr:
        all_candidate_starts.add(max(validated_dsr))

    try:
        import pdfplumber
        with pdfplumber.open(pdf_path) as pdf:
            total_pages = len(pdf.pages)
    except Exception:
        try:
            from pypdf import PdfReader
            total_pages = len(PdfReader(pdf_path).pages)
        except Exception:
            return sorted(recovered)

    for start in sorted(all_candidate_starts):
        for p in range(start, min(start + scan_range_extra + 1, total_pages + 1)):
            if p in recovered:
                continue
            text = _get_page_text(pdf_path, p)
            if _page_has_dsr_table(text):
                recovered.add(p)
                print(f"  [VAL] DSR p{p} RECOVERED — DSR table found by validator scan")

    return sorted(recovered)


def _recover_missed_debt_pages(pdf_path: str, validated_dsr: list,
                                validated_debt: list,
                                scan_range_extra: int = 3) -> list:
    recovered = set(validated_debt)

    try:
        import pdfplumber
        with pdfplumber.open(pdf_path) as pdf:
            total_pages = len(pdf.pages)
    except Exception:
        try:
            from pypdf import PdfReader
            total_pages = len(PdfReader(pdf_path).pages)
        except Exception:
            return sorted(recovered)

    for dsr_page in validated_dsr:
        for p in range(max(1, dsr_page - scan_range_extra), dsr_page + 1):
            if p in recovered:
                continue
            text = _get_page_text(pdf_path, p)
            if _page_has_debt_table(text):
                recovered.add(p)
                print(f"  [VAL] DEBT p{p} RECOVERED — DEBT table found by validator scan")

    return sorted(recovered)


# ════════════════════════════════════════════════════════════════════════
# PROMPT LOADERS
# ════════════════════════════════════════════════════════════════════════

def _load_prompt(prompt_path: str | None = None) -> str:
    candidates = []
    if prompt_path:
        candidates.append(Path(prompt_path))
    script_dir = Path(__file__).parent
    candidates.append(script_dir / "prompts" / "Page_Identification_Prompt.txt")
    candidates.append(Path("prompts") / "Page_Identification_Prompt.txt")
    for p in candidates:
        if p.exists():
            return p.read_text(encoding="utf-8")
    searched = "\n  ".join(str(c) for c in candidates)
    raise FileNotFoundError(
        f"Page identification prompt not found. Searched:\n  {searched}"
    )


def _load_dsr_debt_prompt(prompt_dir: str | None = None) -> str:
    script_dir = Path(__file__).parent
    candidates = [
        script_dir / "prompts" / "DSR_DEBT_Prompt.txt",
        Path("prompts") / "DSR_DEBT_Prompt.txt",
    ]
    if prompt_dir:
        candidates.insert(0, Path(prompt_dir) / "DSR_DEBT_Prompt.txt")
    for p in candidates:
        if p.exists():
            return p.read_text(encoding="utf-8")
    raise FileNotFoundError(f"DSR_DEBT prompt not found. Searched: {candidates}")


def _load_main_tables_prompt(prompt_dir: str | None = None) -> str:
    script_dir = Path(__file__).parent
    candidates = [
        script_dir / "prompts" / "Main_Tables_Prompt.txt",
        Path("prompts") / "Main_Tables_Prompt.txt",
    ]
    if prompt_dir:
        candidates.insert(0, Path(prompt_dir) / "Main_Tables_Prompt.txt")
    for p in candidates:
        if p.exists():
            return p.read_text(encoding="utf-8")
    raise FileNotFoundError(f"Main tables prompt not found. Searched: {candidates}")


# ════════════════════════════════════════════════════════════════════════
# PAGE OFFSET DETECTION & CORRECTION
# ════════════════════════════════════════════════════════════════════════

def _detect_page_offset(pdf_path: str) -> int:
    from collections import Counter
    try:
        import pdfplumber
        with pdfplumber.open(pdf_path) as pdf:
            total   = len(pdf.pages)
            offsets = []
            for phys_idx in range(min(50, total)):
                text  = pdf.pages[phys_idx].extract_text() or ""
                lines = [l.strip() for l in text.splitlines() if l.strip()]
                if not lines:
                    continue
                last = lines[-1]
                if last.isdigit() and 1 <= int(last) <= 50:
                    printed  = int(last)
                    physical = phys_idx + 1
                    offset   = physical - printed
                    if offset >= 0:
                        offsets.append(offset)
            if not offsets:
                return 0
            counter      = Counter(offsets)
            modal_offset = counter.most_common(1)[0][0]
            return modal_offset
    except Exception:
        pass
    return 0


def _correct_page_numbers(normalized: dict, pdf_path: str) -> dict:
    offset = _detect_page_offset(pdf_path)
    if offset == 0:
        return normalized

    anchor_pages = []
    anchor_key   = None
    for key in ["SNP", "SOA", "GOV_BS", "GOV_IS", "PROP_SNP", "PROP_IS", "PROP_CFS"]:
        if normalized.get(key):
            anchor_pages = normalized[key]
            anchor_key   = key
            break

    if not anchor_pages:
        return normalized

    ANCHOR_KEYWORDS = {
        "SNP":      ["net position", "governmental activities", "business-type"],
        "SOA":      ["net (expense)", "program revenues", "general revenues"],
        "GOV_BS":   ["fund balances", "governmental funds", "nonspendable"],
        "GOV_IS":   ["expenditures", "fund balances", "other financing sources"],
        "PROP_SNP": [
            "proprietary", "enterprise", "internal service",
            "net position", "business-type activities",
            "fixed assets", "tangible assets", "total reserves",
            "restricted reserves", "unrestricted reserves",
            "net current assets", "creditors",
            "net assets without donor restrictions",
            "net assets with donor restrictions",
            "total net assets", "balance sheet",
            "total assets", "accounts receivable", "total liabilities",
        ],
        "PROP_IS":  [
            "proprietary", "enterprise", "operating revenues",
            "operating expenses", "operating income",
            "total comprehensive income", "surplus for the year",
            "deficit for the year", "staff costs",
            "total expenditure", "total income",
            "excess of revenue", "excess of revenues",
            "patient service revenue", "statement of operations",
            "total operating expenses", "total revenue", "nonoperating",
        ],
        "PROP_CFS": [
            "proprietary", "cash flows from operating",
            "cash flows from capital", "cash flows from investing",
            "cash flow from operating activities",
            "cash flows from financing activities",
            "increase in cash and cash equivalents",
            "cash flows from financing",
            "net cash provided by operating",
            "net cash used in investing",
            "cash and cash equivalents at end",
            "net increase in cash",
        ],
    }
    check_keywords = ANCHOR_KEYWORDS.get(anchor_key,
                                          ["net position", "assets", "liabilities"])

    raw_text      = _get_page_text(pdf_path, anchor_pages[0])
    adjusted_text = _get_page_text(pdf_path, anchor_pages[0] + offset)

    raw_has_content      = any(kw in raw_text.lower()      for kw in check_keywords)
    adjusted_has_content = any(kw in adjusted_text.lower() for kw in check_keywords)

    if adjusted_has_content and not raw_has_content:
        use_offset = True
    elif raw_has_content and not adjusted_has_content:
        use_offset = False
    elif raw_has_content and adjusted_has_content:
        use_offset = False
    else:
        import re as _re
        raw_amounts      = len(_re.findall(r"\b\d{1,3}(?:,\d{3})+\b", raw_text))
        adjusted_amounts = len(_re.findall(r"\b\d{1,3}(?:,\d{3})+\b", adjusted_text))

        _STMT_SIGNALS = _re.compile(
            r"\btotal\s+(?:assets|liabilities|net\s+(?:assets|position)|revenue|expenses)\b"
            r"|\bbalance\s+sheet\b"
            r"|\bstatement\s+of\s+(?:operations|activities|revenues|cash\s+flows)\b"
            r"|\bnet\s+assets\b|\bnet\s+position\b",
            _re.IGNORECASE,
        )
        raw_signals      = len(_STMT_SIGNALS.findall(raw_text))
        adjusted_signals = len(_STMT_SIGNALS.findall(adjusted_text))

        if raw_signals >= 2 and adjusted_signals < 2:
            use_offset = False
            print(f"  [ID] Heuristic: raw p{anchor_pages[0]} has {raw_signals} stmt signals "
                  f"vs adjusted p{anchor_pages[0]+offset} has {adjusted_signals} — no correction")
        elif adjusted_signals >= 2 and raw_signals < 2:
            use_offset = True
            print(f"  [ID] Heuristic: adjusted p{anchor_pages[0]+offset} has "
                  f"{adjusted_signals} stmt signals vs raw has {raw_signals} — applying offset")
        elif raw_amounts >= 5 and adjusted_amounts < raw_amounts // 2:
            use_offset = False
            print(f"  [ID] Heuristic: raw has {raw_amounts} amounts vs adjusted "
                  f"{adjusted_amounts} — raw is the financial statement, no correction")
        else:
            use_offset = False
            print(f"  [ID] Heuristic: ambiguous (raw_signals={raw_signals}, "
                  f"adj_signals={adjusted_signals}, raw_amt={raw_amounts}, "
                  f"adj_amt={adjusted_amounts}) — defaulting to no correction")

    if not use_offset:
        return normalized

    print(f"  [ID] Applying page offset +{offset} to all statements "
          f"(LLM used printed numbers)")
    corrected = {}
    for key, pages in normalized.items():
        if not pages:
            corrected[key] = pages
        else:
            corrected[key] = [p + offset for p in pages]
            print(f"  [ID] {key}: {pages} → {corrected[key]}")
    return corrected


# ════════════════════════════════════════════════════════════════════════
# ANCHORED PROMPT BUILDER
# ════════════════════════════════════════════════════════════════════════

def _build_anchored_prompt(prompt_text: str, pdf_path: str) -> str:
    """
    Prepend a verified page-numbering anchor hint to the prompt.
    For Claude text mode this is less critical (<<<PAGE N>>> markers
    already give exact physical page numbers), but kept so
    Gemini / OpenAI paths still benefit from it.
    """
    from collections import Counter
    try:
        import pdfplumber
        with pdfplumber.open(pdf_path) as pdf:
            total             = len(pdf.pages)
            anchor_candidates = []

            for phys_idx in range(min(50, total)):
                text  = pdf.pages[phys_idx].extract_text() or ""
                lines = [l.strip() for l in text.splitlines() if l.strip()]
                if not lines:
                    continue
                last = lines[-1]
                if last.isdigit() and 1 <= int(last) <= 50:
                    printed  = int(last)
                    physical = phys_idx + 1
                    offset   = physical - printed
                    if offset >= 0:
                        anchor_candidates.append(
                            (physical, printed, offset, lines[0][:60])
                        )

            if not anchor_candidates:
                return prompt_text

            offsets      = [c[2] for c in anchor_candidates]
            counter      = Counter(offsets)
            modal_offset = counter.most_common(1)[0][0]

            if modal_offset == 0:
                return prompt_text

            modal_anchors                     = [c for c in anchor_candidates
                                                  if c[2] == modal_offset]
            physical, printed, offset, title  = modal_anchors[-1]

            hint = (
                f"\n\n"
                f"════════════════════════════════════════════════════════════════\n"
                f"CRITICAL PAGE NUMBERING ANCHOR — READ THIS FIRST\n"
                f"════════════════════════════════════════════════════════════════\n"
                f"This PDF has {offset} un-numbered front-matter page(s) "
                f"(covers, blank pages, roman-numeral TOC pages) before the "
                f"main content begins.\n\n"
                f"VERIFIED REFERENCE POINT:\n"
                f"  Physical page {physical} (the {physical}th page counting "
                f"from the very first page of the file)\n"
                f"  has printed footer number '{printed}'\n"
                f"  and starts with: \"{title}\"\n\n"
                f"FORMULA FOR THIS PDF:\n"
                f"  physical page number = printed footer number + {offset}\n\n"
                f"EXAMPLES:\n"
                f"  Printed footer '5'  → return physical page {5  + offset}\n"
                f"  Printed footer '10' → return physical page {10 + offset}\n"
                f"  Printed footer '42' → return physical page {42 + offset}\n\n"
                f"YOU MUST return physical page numbers (counted from the first\n"
                f"page of the file), NOT the printed numbers in the footers.\n"
                f"DO NOT return the printed footer number — add {offset} to it first.\n"
                f"════════════════════════════════════════════════════════════════\n"
            )
            return hint + prompt_text
    except Exception:
        pass
    return prompt_text


# ════════════════════════════════════════════════════════════════════════
# NON-LG PAGE VALIDATION
# ════════════════════════════════════════════════════════════════════════

def _validate_nonlg_pages(pdf_path: str, normalized: dict) -> dict:
    from page_extractor import (
        _looks_like_prop_snp_body, _looks_like_prop_is_body,
        _looks_like_prop_cfs_body, _nonlg_excluded,
    )

    _UK_CHANGES_RESERVES_RE = re.compile(
        r"statements?\s+of\s+changes?\s+in\s+reserves?\b"
        r"|statements?\s+of\s+changes?\s+in\s+equity\b"
        r"|consolidated\s+(?:and\s+\w+\s+)?statements?\s+of\s+changes",
        re.IGNORECASE,
    )

    VALIDATORS = {
        "PROP_SNP": (_looks_like_prop_snp_body, _PROP_SNP_VAL_TITLE_RE),
        "PROP_IS":  (_looks_like_prop_is_body,  _PROP_IS_VAL_TITLE_RE),
        "PROP_CFS": (_looks_like_prop_cfs_body, _PROP_CFS_VAL_TITLE_RE),
    }

    validated = {}
    for key, pages in normalized.items():
        entry = VALIDATORS.get(key)
        if not entry or not pages:
            validated[key] = pages
            continue

        body_validator, title_re = entry
        valid_pages = []

        for idx, p in enumerate(pages):
            text = _get_page_text(pdf_path, p)
            if not text or _nonlg_excluded(text):
                print(f"  [VAL] {key} p{p} REMOVED — excluded")
                continue

            if body_validator(text) or _has_statement_title_and_amounts(text, title_re):
                valid_pages.append(p)
                continue

            if key == "PROP_IS" and idx > 0:
                prev_p = pages[idx - 1]
                if p == prev_p + 1 and prev_p in valid_pages:
                    if _UK_CHANGES_RESERVES_RE.search(text):
                        valid_pages.append(p)
                        print(f"  [VAL] {key} p{p} KEPT — UK Changes in "
                              f"Reserves immediately follows p{prev_p}")
                        continue

            print(f"  [VAL] {key} p{p} REMOVED — body content doesn't match")

        validated[key] = valid_pages

    return validated


def _keep_first_cluster(normalized: dict, max_gap: int = 5) -> dict:
    all_pages = sorted(set(
        p for pages in normalized.values() for p in pages
    ))
    if not all_pages:
        return normalized
    cluster = [all_pages[0]]
    for p in all_pages[1:]:
        if p - cluster[-1] <= max_gap:
            cluster.append(p)
        else:
            break
    cluster_set = set(cluster)
    return {key: [p for p in pages if p in cluster_set]
            for key, pages in normalized.items()}


# ════════════════════════════════════════════════════════════════════════
# TWO FOCUSED LLM CALLS
# ════════════════════════════════════════════════════════════════════════

def _identify_dsr_debt_only(gemini_client, pdf_bytes: bytes, model: str,
                             prompt_dir: str | None = None,
                             pdf_path: str | None = None) -> dict:
    """
    Call 1: identify DSR + DEBT pages only.
    The static DSR/DEBT prompt is cached by all three providers.
    """
    prompt_text = _load_dsr_debt_prompt(prompt_dir)
    if pdf_path:
        prompt_text = _build_anchored_prompt(prompt_text, pdf_path)

    try:
        raw = _dispatch_llm_call(pdf_bytes, model, prompt_text,
                                  gemini_client=gemini_client,
                                  max_tokens=4096,
                                  pdf_path=pdf_path)
        result     = _parse_json_from_raw(raw)
        normalized = {}
        for key in ["DSR", "DEBT"]:
            val = result.get(key, [])
            normalized[key] = [
                int(p) for p in val
                if isinstance(p, (int, float))
                or (isinstance(p, str) and str(p).isdigit())
            ]
        return normalized
    except Exception as e:
        print(f"  [ID] DSR/DEBT call failed ({e}) — returning empty")
        return {"DSR": [], "DEBT": []}


def _identify_main_tables(gemini_client, pdf_bytes: bytes, model: str,
                           prompt_dir: str | None = None,
                           pdf_path: str | None = None,
                           allowed_suffixes: set[str] | None = None) -> dict:
    """
    Call 2: identify main financial statement pages.
    The static Main Tables prompt is cached by all three providers.
    """
    prompt_text = _load_main_tables_prompt(prompt_dir)
    prompt_text = _trim_prompt_for_sector(prompt_text, allowed_suffixes)
    if pdf_path:
        prompt_text = _build_anchored_prompt(prompt_text, pdf_path)

    try:
        raw = _dispatch_llm_call(pdf_bytes, model, prompt_text,
                                  gemini_client=gemini_client,
                                  max_tokens=8192,
                                  pdf_path=pdf_path)
        result     = _parse_json_from_raw(raw)
        normalized = {}
        for key, val in result.items():
            key_upper = key.upper().replace("-", "_")
            if key_upper in ("DSR", "DEBT"):
                continue
            if isinstance(val, list):
                normalized[key_upper] = [
                    int(p) for p in val
                    if isinstance(p, (int, float))
                    or (isinstance(p, str) and str(p).isdigit())
                ]
            elif isinstance(val, int):
                normalized[key_upper] = [val]
        return normalized
    except Exception as e:
        print(f"  [ID] Main tables call failed ({e}) — returning empty")
        return {}


def _full_scan_dsr_debt(pdf_path: str):
    try:
        import pdfplumber
        with pdfplumber.open(pdf_path) as pdf:
            total = len(pdf.pages)
    except Exception:
        try:
            from pypdf import PdfReader
            total = len(PdfReader(pdf_path).pages)
        except Exception:
            return [], []

    dsr_pages, debt_pages = [], []
    for p in range(1, total + 1):
        text = _get_page_text(pdf_path, p)
        if _page_has_dsr_table(text):
            dsr_pages.append(p)
        if _page_has_debt_table(text):
            debt_pages.append(p)
    return dsr_pages, debt_pages


# ════════════════════════════════════════════════════════════════════════
# PUBLIC ENTRY POINT
# ════════════════════════════════════════════════════════════════════════

def identify_pages_gemini(
    pdf_path: str,
    model: str = "gemini-2.5-flash",
    prompt_path: str | None = None,
    allowed_suffixes: set[str] | None = None,
) -> dict:
    """
    Identify financial-statement pages in *pdf_path* using *model*.

    Supported models + caching behaviour:
      gemini-*   → Gemini Context Cache (static prompt cached in cloud,
                   TTL = GEMINI_CACHE_TTL_MINUTES, requires ≥ 32 k tokens)
      claude-*   → Claude ephemeral system-prompt cache (TTL = 5 min,
                   text-based — no 100-page limit, ~10x cheaper than binary)
      gpt-* / …  → OpenAI automatic prefix cache (implicit, no extra calls)
    """
    provider = _provider_for_model(model)

    # ── Build Gemini client only when needed ──────────────────────────
    gemini_client = None
    if provider == "gemini":
        try:
            from google import genai
        except ImportError:
            raise RuntimeError("google-genai not installed: pip install google-genai")

        api_key = os.environ.get("GEMINI_API_KEY", "")
        if not api_key:
            try:
                import winreg
                with winreg.OpenKey(winreg.HKEY_CURRENT_USER, "Environment") as key:
                    value, _ = winreg.QueryValueEx(key, "GEMINI_API_KEY")
                    api_key  = value
                    os.environ["GEMINI_API_KEY"] = api_key
            except Exception:
                pass
        if not api_key:
            raise RuntimeError("GEMINI_API_KEY not set")
        gemini_client = genai.Client(api_key=api_key)

    # pdf_bytes always read — Gemini/OpenAI use binary; Claude ignores it
    pdf_bytes = Path(pdf_path).read_bytes()

    # ── CALL 1: DSR + DEBT ────────────────────────────────────────────
    dsr_debt_keys  = {"_DSR", "_DEBT"}
    needs_dsr_debt = (allowed_suffixes is None) or bool(dsr_debt_keys & allowed_suffixes)

    if needs_dsr_debt:
        dsr_debt_result = _identify_dsr_debt_only(
            gemini_client, pdf_bytes, model,
            pdf_path=pdf_path,
        )
    else:
        print("  [ID] Skipping DSR/DEBT call — not needed for this sector")
        dsr_debt_result = {"DSR": [], "DEBT": []}

    # ── CALL 2: Main tables ───────────────────────────────────────────
    main_result = _identify_main_tables(
        gemini_client, pdf_bytes, model,
        pdf_path=pdf_path,
        allowed_suffixes=allowed_suffixes,
    )

    # ── Merge ─────────────────────────────────────────────────────────
    normalized = {**main_result, **dsr_debt_result}

    if allowed_suffixes is not None:
        allowed_keys = {s.lstrip("_").upper() for s in allowed_suffixes}
        normalized   = {k: v for k, v in normalized.items() if k in allowed_keys}

    if allowed_suffixes is not None and not needs_dsr_debt:
        normalized = _validate_nonlg_pages(pdf_path, normalized)
        normalized = _keep_first_cluster(normalized, max_gap=5)

    normalized = _correct_page_numbers(normalized, pdf_path)

    if not needs_dsr_debt:
        print(f"  [ID] DSR/DEBT validation skipped — not needed for this sector")
        normalized["DSR"]  = []
        normalized["DEBT"] = []
        return normalized

    final_dsr  = normalized.get("DSR",  [])
    final_debt = normalized.get("DEBT", [])

    if not final_dsr and not final_debt:
        print(f"  [VAL] LLM returned no DSR/DEBT — running full document scan as fallback")
        final_dsr, final_debt = _full_scan_dsr_debt(pdf_path)

    normalized["DSR"]  = sorted(set(final_dsr))
    normalized["DEBT"] = sorted(set(final_debt))

    return normalized


def identify_pages_with_fallback(
    pdf_path: str,
    model: str,
    fallback_fn=None,
    prompt_path: str | None = None,
    allowed_suffixes: set[str] | None = None,
) -> dict:
    try:
        result = identify_pages_gemini(
            pdf_path, model=model,
            prompt_path=prompt_path,
            allowed_suffixes=allowed_suffixes,
        )
        if result:
            return result
        print("  [ID] LLM returned empty — falling back to Python detection")
    except Exception as e:
        print(f"  [ID] LLM identification failed ({e}) — falling back")

    if fallback_fn:
        return fallback_fn(pdf_path)

    return {}