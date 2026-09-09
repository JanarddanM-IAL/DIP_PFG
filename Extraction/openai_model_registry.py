"""
openai_model_registry.py
═════════════════════════
SINGLE SOURCE OF TRUTH for every OpenAI model used in this pipeline.

To add a new model, add ONE entry to OPENAI_MODELS below.
All cost helpers derive from it automatically.

Entry format:
    "model-name": ModelSpec(
        standard_in   = $/1M input tokens
        standard_out  = $/1M output tokens
        cached_in     = $/1M cached input tokens (OpenAI automatic prefix caching)
                        set to None to default to 50% of standard_in
    )
"""

from dataclasses import dataclass, field
from typing import Optional


@dataclass(frozen=True)
class ModelSpec:
    standard_in:     float
    standard_out:    float
    cached_in:       Optional[float] = None  # defaults to 50% of standard_in
    batch_in:        Optional[float] = None  # defaults to 50% of standard_in
    batch_out:       Optional[float] = None  # defaults to 50% of standard_out
    id_capable:      bool            = True  # True -> appears in ID Model dropdown
    extract_capable: bool            = True  # True -> appears in extraction Model dropdown

    def effective_cached_in(self) -> float:
        return self.cached_in if self.cached_in is not None else self.standard_in * 0.5

    def effective_batch_in(self) -> float:
        return self.batch_in if self.batch_in is not None else self.standard_in * 0.5

    def effective_batch_out(self) -> float:
        return self.batch_out if self.batch_out is not None else self.standard_out * 0.5


# ══════════════════════════════════════════════════════════════════════════════
# ▼▼▼  ADD NEW MODELS HERE — this is the ONLY place you need to touch  ▼▼▼
# ══════════════════════════════════════════════════════════════════════════════
OPENAI_MODELS: dict[str, ModelSpec] = {
    "gpt-4o-mini":  ModelSpec(standard_in=0.15,  standard_out=0.60),
    "gpt-4.1-mini": ModelSpec(standard_in=0.40,  standard_out=1.60),
    "gpt-4o":       ModelSpec(standard_in=2.50,  standard_out=10.00),
    "gpt-4.1":      ModelSpec(standard_in=2.00,  standard_out=8.00),
    "gpt-5.4-mini": ModelSpec(standard_in=0.40,  standard_out=1.60),
    "gpt-5.5":      ModelSpec(standard_in=5.00,  standard_out=30.00),

    # Add future models below this line:
    # "gpt-6-mini": ModelSpec(standard_in=X.XX, standard_out=X.XX),
}
# ▲▲▲  END OF MODEL REGISTRY  ▲▲▲
# ══════════════════════════════════════════════════════════════════════════════

_DEFAULT_SPEC = ModelSpec(standard_in=0.0, standard_out=0.0)


def get_spec(model: str) -> ModelSpec:
    if model in OPENAI_MODELS:
        return OPENAI_MODELS[model]
    # Fuzzy prefix match (e.g. "gpt-4o-mini-2024-07-18" → "gpt-4o-mini")
    for key in OPENAI_MODELS:
        if model.startswith(key):
            return OPENAI_MODELS[key]
    print(
        f"[openai_model_registry] WARN: model '{model}' not in registry — "
        f"cost will show as $0.00. Add it to OPENAI_MODELS in openai_model_registry.py."
    )
    return _DEFAULT_SPEC


def standard_rates(model: str) -> tuple[float, float]:
    s = get_spec(model)
    return s.standard_in, s.standard_out


def calculate_cost(model: str, prompt_tokens: int, completion_tokens: int) -> float:
    in_rate, out_rate = standard_rates(model)
    return (prompt_tokens / 1_000_000) * in_rate + (completion_tokens / 1_000_000) * out_rate


def calculate_cached_cost(
    model: str,
    prompt_tokens: int,
    cached_tokens: int,
    completion_tokens: int,
) -> float:
    """Cost with OpenAI automatic prefix-caching discount applied."""
    spec      = get_spec(model)
    in_rate   = spec.standard_in
    out_rate  = spec.standard_out
    c_in_rate = spec.effective_cached_in()
    fresh     = max(0, prompt_tokens - cached_tokens)
    return (
        (fresh         / 1_000_000) * in_rate
        + (cached_tokens / 1_000_000) * c_in_rate
        + (completion_tokens / 1_000_000) * out_rate
    )


def calculate_batch_cost(
    model: str,
    prompt_tokens: int,
    cached_tokens: int,
    completion_tokens: int,
) -> float:
    """Batch-API cost (50% of standard rates by default)."""
    spec      = get_spec(model)
    in_rate   = spec.effective_batch_in()
    out_rate  = spec.effective_batch_out()
    # Cached discount still applies on top of the batch base rate
    c_in_rate = in_rate * 0.5
    fresh     = max(0, prompt_tokens - cached_tokens)
    return (
        (fresh         / 1_000_000) * in_rate
        + (cached_tokens / 1_000_000) * c_in_rate
        + (completion_tokens / 1_000_000) * out_rate
    )


# Backwards-compat dict — read-only view, same shape as old OPENAI_COST_TABLE
OPENAI_COST_TABLE: dict[str, tuple[float, float]] = {
    m: (s.standard_in, s.standard_out) for m, s in OPENAI_MODELS.items()
}

# ── UI helpers — derive dropdown lists from the registry ──────────────────────

def ui_extract_models() -> list[str]:
    """Model strings for the extraction Model dropdown (provider=openai)."""
    return [m for m, s in OPENAI_MODELS.items() if s.extract_capable]

def ui_id_models() -> list[str]:
    """Model strings that should appear in the ID Model dropdown."""
    return [m for m, s in OPENAI_MODELS.items() if s.id_capable]