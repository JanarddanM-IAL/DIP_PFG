"""
claude_model_registry.py
═════════════════════════
SINGLE SOURCE OF TRUTH for every Anthropic Claude model used in this pipeline.

To add a new model, add ONE entry to CLAUDE_MODELS below.
All cost helpers derive from it automatically.

Entry format:
    "model-name": ModelSpec(
        standard_in       = $/1M input tokens
        standard_out      = $/1M output tokens
        cache_write_in    = $/1M tokens written to prompt cache
                            (Anthropic charges slightly more for the first write)
                            set to None to default to 1.25x standard_in
        cached_in         = $/1M tokens read from prompt cache
                            set to None to default to 10% of standard_in
        batch_in          = $/1M input tokens via Batch API
                            set to None to default to 50% of standard_in
        batch_out         = $/1M output tokens via Batch API
                            set to None to default to 50% of standard_out
    )
"""

import re as _re
from dataclasses import dataclass
from typing import Optional


@dataclass(frozen=True)
class ModelSpec:
    standard_in:     float
    standard_out:    float
    display_name:    str             = ""     # human label shown in UI dropdowns
    id_capable:      bool            = True   # True → appears in ID Model dropdown
    extract_capable: bool            = False  # True → appears in extraction Model dropdown
    cache_write_in:  Optional[float] = None   # defaults to 1.25x standard_in
    cached_in:       Optional[float] = None   # defaults to 10% of standard_in
    batch_in:        Optional[float] = None   # defaults to 50% of standard_in
    batch_out:       Optional[float] = None   # defaults to 50% of standard_out

    def effective_cache_write_in(self) -> float:
        return self.cache_write_in if self.cache_write_in is not None else self.standard_in * 1.25

    def effective_cached_in(self) -> float:
        return self.cached_in if self.cached_in is not None else self.standard_in * 0.10

    def effective_batch_in(self) -> float:
        return self.batch_in if self.batch_in is not None else self.standard_in * 0.50

    def effective_batch_out(self) -> float:
        return self.batch_out if self.batch_out is not None else self.standard_out * 0.50


# =============================================================================
# ADD NEW MODELS HERE — this is the ONLY place you need to touch.
# All cost helpers, alias maps, and UI dropdowns derive automatically.
# =============================================================================
#                                         std_in  std_out  display_name         id     extract
CLAUDE_MODELS: dict[str, ModelSpec] = {
    "claude-haiku-4-5-20251001": ModelSpec(1.00,   5.00,  "Claude Haiku 4.5",  True,  False),
    "claude-sonnet-4-5":         ModelSpec(3.00,  15.00,  "Claude Sonnet 4.5", True,  True),
    "claude-sonnet-4-6":         ModelSpec(3.00,  15.00,  "Claude Sonnet 4.6", True,  True),
    "claude-opus-4-5":           ModelSpec(5.00,  25.00,  "Claude Opus 4.5",   True,  True),
    "claude-opus-4-8":           ModelSpec(5.00,  25.00,  "Claude Opus 4.8",   True,  True),

    # Add future models below this line — everything else updates automatically:
    # "claude-sonnet-5": ModelSpec(X.XX, X.XX, "Claude Sonnet 5", True, True),
}
# ▲▲▲  END OF MODEL REGISTRY  ▲▲▲
# ══════════════════════════════════════════════════════════════════════════════

_DEFAULT_SPEC = ModelSpec(standard_in=3.00, standard_out=15.00)

# ---------------------------------------------------------------------------
# Display-name → API-ID alias map (auto-built from registry).
# Covers both "Claude Sonnet 4.6" and dot-style "claude-sonnet-4.6" variants
# so claude_cache_client never needs its own hardcoded CLAUDE_MODEL_ALIASES.
# ---------------------------------------------------------------------------
CLAUDE_MODEL_ALIASES: dict[str, str] = {}

def _init_aliases() -> None:
    """Populate CLAUDE_MODEL_ALIASES from the registry (called once at import)."""
    for api_id, spec in CLAUDE_MODELS.items():
        if spec.display_name:
            CLAUDE_MODEL_ALIASES[spec.display_name] = api_id
        # dot-style: "claude-sonnet-4-6" → "claude-sonnet-4.6"
        dot_variant = _re.sub(r"-(\d+)-(\d+)$", r"-\1.\2", api_id)
        if dot_variant != api_id:
            CLAUDE_MODEL_ALIASES[dot_variant] = api_id

_init_aliases()


def resolve_claude_model(model: str | None) -> str:
    """
    Normalise any model identifier (display name, dot-style alias, or API ID)
    to the canonical API model string.  Falls back to claude-sonnet-4-6.
    """
    if not model:
        return "claude-sonnet-4-6"
    return CLAUDE_MODEL_ALIASES.get(model, model)


def get_spec(model: str) -> ModelSpec:
    if model in CLAUDE_MODELS:
        return CLAUDE_MODELS[model]
    # Fuzzy prefix match (e.g. "claude-sonnet-4-5-20241022" → "claude-sonnet-4-5")
    for key in CLAUDE_MODELS:
        if model.startswith(key):
            return CLAUDE_MODELS[key]
    print(
        f"[claude_model_registry] WARN: model '{model}' not in registry — "
        f"falling back to default rates {(_DEFAULT_SPEC.standard_in, _DEFAULT_SPEC.standard_out)}. "
        f"Add it to CLAUDE_MODELS in claude_model_registry.py."
    )
    return _DEFAULT_SPEC


def standard_rates(model: str) -> tuple[float, float]:
    s = get_spec(model)
    return s.standard_in, s.standard_out


def calculate_cost(model: str, prompt_tokens: int, completion_tokens: int) -> float:
    in_rate, out_rate = standard_rates(model)
    return (prompt_tokens / 1_000_000) * in_rate + (completion_tokens / 1_000_000) * out_rate


def calculate_batch_cost(
    model: str,
    input_tokens: int,
    output_tokens: int,
    cache_creation_tokens: int = 0,
    cache_read_tokens: int = 0,
) -> float:
    """
    Batch-API cost for a Claude call (50% of standard rates by default).
    Cache multipliers still apply on top of the batch base rate.
    """
    spec     = get_spec(resolve_claude_model(model))
    in_rate  = spec.effective_batch_in()
    out_rate = spec.effective_batch_out()
    cache_write_rate = in_rate * 1.25
    cache_read_rate  = in_rate * 0.10
    return (
        (input_tokens            / 1_000_000) * in_rate
        + (cache_creation_tokens / 1_000_000) * cache_write_rate
        + (cache_read_tokens     / 1_000_000) * cache_read_rate
        + (output_tokens         / 1_000_000) * out_rate
    )


def calculate_cached_cost(
    model: str,
    prompt_tokens: int,
    cache_read_tokens: int,
    cache_write_tokens: int,
    completion_tokens: int,
) -> tuple[float, str]:
    """
    Full Claude prompt-caching cost breakdown.
    Returns (total_usd, human_readable_breakdown).
    """
    spec = get_spec(model)
    in_rate    = spec.standard_in
    out_rate   = spec.standard_out
    write_rate = spec.effective_cache_write_in()
    read_rate  = spec.effective_cached_in()

    fresh_tokens = max(0, prompt_tokens - cache_read_tokens - cache_write_tokens)
    fresh_cost   = (fresh_tokens       / 1_000_000) * in_rate
    write_cost   = (cache_write_tokens / 1_000_000) * write_rate
    read_cost    = (cache_read_tokens  / 1_000_000) * read_rate
    output_cost  = (completion_tokens  / 1_000_000) * out_rate
    total        = fresh_cost + write_cost + read_cost + output_cost

    ratio = (cache_read_tokens / prompt_tokens * 100) if prompt_tokens else 0
    breakdown = (
        f"cache_read={cache_read_tokens:,}({ratio:.0f}%) "
        f"cache_write={cache_write_tokens:,} "
        f"fresh={fresh_tokens:,} out={completion_tokens:,} → "
        f"${total:.4f}"
    )
    return total, breakdown


# Backwards-compat dict — read-only view, same shape as old CLAUDE_COST_TABLE
CLAUDE_COST_TABLE: dict[str, tuple[float, float]] = {
    m: (s.standard_in, s.standard_out) for m, s in CLAUDE_MODELS.items()
}

# ── UI helpers — derive dropdown lists from the registry ──────────────────────

def ui_extract_api_models() -> list[str]:
    """API model strings for the extraction model dropdown (provider=claude)."""
    return [m for m, s in CLAUDE_MODELS.items() if s.extract_capable]

def ui_extract_display_names() -> list[str]:
    """Human labels for the extraction model dropdown."""
    return [s.display_name for m, s in CLAUDE_MODELS.items() if s.extract_capable]

def ui_display_to_api() -> dict[str, str]:
    """Map display_name -> api_model_string (for resolving the UI selection)."""
    return {s.display_name: m for m, s in CLAUDE_MODELS.items() if s.display_name}

def ui_id_models() -> list[str]:
    """API model strings that should appear in the ID Model dropdown."""
    return [m for m, s in CLAUDE_MODELS.items() if s.id_capable]