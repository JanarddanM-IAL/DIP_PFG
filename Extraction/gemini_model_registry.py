"""
gemini_model_registry.py
════════════════════════
SINGLE SOURCE OF TRUTH for every Gemini model used in this pipeline.

To add a new model, add ONE entry to GEMINI_MODELS below.
All rate tables (standard, batch, cached) and max-output-token caps
are derived automatically from it — no other file needs editing.

Entry format:
    "model-name": ModelSpec(
        standard_in   = $/1M input tokens  (standard / sync rate)
        standard_out  = $/1M output tokens
        batch_in      = $/1M input tokens  (Batch API, typically 50% of standard)
        batch_out     = $/1M output tokens (Batch API)
        cached_in     = $/1M input tokens  (Context Cache hit, typically 25% of standard)
        max_output    = maximum output tokens this model supports
        supports_batch  = True if model is available via Gemini Batch API
        supports_cache  = True if model supports Context Caching
    )
"""

from dataclasses import dataclass, field


@dataclass(frozen=True)
class ModelSpec:
    standard_in:     float
    standard_out:    float
    batch_in:        float
    batch_out:       float
    cached_in:       float       # cached_out is always == standard_out
    max_output:      int
    supports_batch:  bool = True
    supports_cache:  bool = True
    id_capable:      bool = True  # True -> appears in ID Model dropdown
    extract_capable: bool = True  # True -> appears in extraction Model dropdown


# ══════════════════════════════════════════════════════════════════════════════
# ▼▼▼  ADD NEW MODELS HERE — this is the ONLY place you need to touch  ▼▼▼
# ══════════════════════════════════════════════════════════════════════════════
GEMINI_MODELS: dict[str, ModelSpec] = {
    "gemini-3.8-flash": ModelSpec(
                standard_in=0.75,  standard_out=3.75,
                batch_in=0.38,     batch_out=1.88,
                cached_in=0.19,
                max_output=65536,
            ),
    # ── Legacy models ──────────────────────────────────────────────────────────
    "gemini-2.0-flash": ModelSpec(
        standard_in=0.075, standard_out=0.30,
        batch_in=0.0375,   batch_out=0.15,
        cached_in=0.01875,
        max_output=8192,
    ),
    "gemini-2.0-flash-lite": ModelSpec(
        standard_in=0.075, standard_out=0.30,
        batch_in=0.0375,   batch_out=0.15,
        cached_in=0.01875,
        max_output=8192,
    ),
    "gemini-1.5-flash": ModelSpec(
        standard_in=0.075, standard_out=0.30,
        batch_in=0.0375,   batch_out=0.15,
        cached_in=0.01875,
        max_output=8192,
    ),
    "gemini-1.5-pro": ModelSpec(
        standard_in=1.25,  standard_out=5.00,
        batch_in=0.625,    batch_out=2.50,
        cached_in=0.3125,
        max_output=8192,
    ),

    # ── Current generation ─────────────────────────────────────────────────────
    "gemini-2.5-flash": ModelSpec(
        standard_in=0.30,  standard_out=2.50,
        batch_in=0.15,     batch_out=1.25,
        cached_in=0.0375,
        max_output=65536,
    ),
    "gemini-2.5-pro": ModelSpec(
        standard_in=1.25,  standard_out=10.00,
        batch_in=0.625,    batch_out=5.00,
        cached_in=0.3125,
        max_output=65536,
    ),

    # ── Next generation ────────────────────────────────────────────────────────
    "gemini-3.1-flash-lite-preview": ModelSpec(
        standard_in=0.25,  standard_out=1.50,
        batch_in=0.125,    batch_out=0.75,
        cached_in=0.0625,
        max_output=65536,
    ),
    "gemini-3.5-flash": ModelSpec(
        standard_in=1.50,  standard_out=9.00,
        batch_in=0.75,     batch_out=4.50,
        cached_in=0.375,
        max_output=65536,
    ),
    "gemini-3.6-flash": ModelSpec(
        standard_in=1.50,  standard_out=7.50,
        batch_in=0.75,     batch_out=3.75,
        cached_in=0.375,
        max_output=65536,
    ),
    
    # ── Add future models below this line ─────────────────────────────────────
    # "gemini-4.0-flash": ModelSpec(
    #     standard_in=X.XX, standard_out=X.XX,
    #     batch_in=X.XX,    batch_out=X.XX,
    #     cached_in=X.XX,
    #     max_output=65536,
    # ),
}
# ▲▲▲  END OF MODEL REGISTRY  ▲▲▲
# ══════════════════════════════════════════════════════════════════════════════

_DEFAULT_SPEC = ModelSpec(
    standard_in=0.75,  standard_out=4.50,
    batch_in=0.375,    batch_out=2.25,
    cached_in=0.1875,
    max_output=65536,
)


# ── Public helpers ─────────────────────────────────────────────────────────────

def get_spec(model: str) -> ModelSpec:
    """
    Return the ModelSpec for *model*.
    Falls back to _DEFAULT_SPEC (with a warning) if the model is not registered.
    """
    if model in GEMINI_MODELS:
        return GEMINI_MODELS[model]
    print(
        f"[gemini_model_registry] WARN: model '{model}' not in registry — "
        f"using default rates {(_DEFAULT_SPEC.standard_in, _DEFAULT_SPEC.standard_out)}. "
        f"Add it to GEMINI_MODELS in gemini_model_registry.py."
    )
    return _DEFAULT_SPEC


def max_output_tokens(model: str) -> int:
    """Return the max output token cap for *model*."""
    return get_spec(model).max_output


def standard_rates(model: str) -> tuple[float, float]:
    """Return (input_rate, output_rate) in $/1M tokens — standard (sync) pricing."""
    s = get_spec(model)
    return s.standard_in, s.standard_out


def batch_rates(model: str) -> tuple[float, float]:
    """Return (input_rate, output_rate) in $/1M tokens — Batch API pricing."""
    s = get_spec(model)
    return s.batch_in, s.batch_out


def cached_rates(model: str) -> tuple[float, float]:
    """Return (cached_input_rate, output_rate) in $/1M tokens — Context Cache pricing."""
    s = get_spec(model)
    return s.cached_in, s.standard_out


def calculate_cost(model: str, prompt_tokens: int, completion_tokens: int) -> float:
    """Standard (sync) cost for a call. Used by gemini_client.py."""
    in_rate, out_rate = standard_rates(model)
    return (prompt_tokens / 1_000_000) * in_rate + (completion_tokens / 1_000_000) * out_rate


def calculate_batch_cost(model: str, prompt_tokens: int, completion_tokens: int) -> float:
    """Batch-discounted cost for a call. Used by gemini_batch_client.py."""
    in_rate, out_rate = batch_rates(model)
    return (prompt_tokens / 1_000_000) * in_rate + (completion_tokens / 1_000_000) * out_rate


def calculate_cached_cost(
    model: str,
    prompt_tokens: int,
    cached_tokens: int,
    completion_tokens: int,
) -> tuple[float, str]:
    """
    Context-cache cost breakdown.  Returns (total_usd, human_readable_breakdown).
    Used by gemini_cache_client.py.
    """
    in_rate,  out_rate    = standard_rates(model)
    c_in_rate, _          = cached_rates(model)

    fresh        = max(0, prompt_tokens - cached_tokens)
    fresh_cost   = (fresh             / 1_000_000) * in_rate
    cached_cost  = (cached_tokens     / 1_000_000) * c_in_rate
    output_cost  = (completion_tokens / 1_000_000) * out_rate
    total        = fresh_cost + cached_cost + output_cost

    ratio = (cached_tokens / prompt_tokens * 100) if prompt_tokens else 0
    breakdown = (
        f"cached={cached_tokens:,}({ratio:.0f}%) "
        f"fresh={fresh:,} out={completion_tokens:,} → "
        f"${total:.4f}"
    )
    return total, breakdown


# ── Backwards-compat dicts (so old code referencing the dicts still works) ───
# These are READ-ONLY views — do not edit them; edit GEMINI_MODELS above.

GEMINI_COST_TABLE: dict[str, tuple[float, float]] = {
    m: (s.standard_in, s.standard_out) for m, s in GEMINI_MODELS.items()
}
GEMINI_BATCH_RATES: dict[str, tuple[float, float]] = {
    m: (s.batch_in, s.batch_out) for m, s in GEMINI_MODELS.items()
}
GEMINI_STANDARD_RATES: dict[str, tuple[float, float]] = GEMINI_COST_TABLE   # alias
GEMINI_CACHED_RATES: dict[str, tuple[float, float]] = {
    m: (s.cached_in, s.standard_out) for m, s in GEMINI_MODELS.items()
}
GEMINI_NORMAL_RATES: dict[str, tuple[float, float]] = GEMINI_COST_TABLE     # alias
MODEL_MAX_OUTPUT: dict[str, int] = {
    m: s.max_output for m, s in GEMINI_MODELS.items()
}

# ── UI helpers — derive dropdown lists from the registry ──────────────────────

def ui_extract_models() -> list[str]:
    """Model strings for the extraction Model dropdown (provider=gemini)."""
    return [m for m, s in GEMINI_MODELS.items() if s.extract_capable]

def ui_id_models() -> list[str]:
    """Model strings that should appear in the ID Model dropdown."""
    return [m for m, s in GEMINI_MODELS.items() if s.id_capable]