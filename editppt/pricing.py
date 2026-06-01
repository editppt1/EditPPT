"""Per-token USD pricing for the LLM models we call.

OpenAI pricing: standard tier, source https://platform.openai.com/docs/pricing
Gemini pricing: standard tier, source https://ai.google.dev/gemini-api/docs/pricing
Last verified: 2026-04-30

Notes
-----
- All rates are USD per 1 token (the listed-per-1M values divided by 1e6).
- "cached_input" is the discounted rate for prompt-cache hits. When a model
  does NOT support prompt caching, the key is omitted and `compute_cost`
  falls back to the standard input rate.
- Models with context-tiered pricing (e.g. Gemini 2.5 Pro >200K, gpt-5.5
  >272K) are listed at the BASE tier; prompts beyond the threshold are
  charged more by the provider but our pipeline never approaches those
  ceilings, so undercounting risk is negligible.
- Legacy models (gpt-4-*, gpt-4-turbo-*, gpt-3.5-*, davinci, babbage) are
  intentionally excluded — we do not target them.
"""

_M = 1_000_000

# USD per token. Each entry: {input, cached_input?, output}
PRICING: dict[str, dict[str, float]] = {
    # ----- OpenAI GPT-5 family -----
    "gpt-5.5":          {"input":  5.00 / _M, "cached_input": 0.50  / _M, "output":  30.00 / _M},
    "gpt-5.5-pro":      {"input": 30.00 / _M,                              "output": 180.00 / _M},
    "gpt-5.4":          {"input":  2.50 / _M, "cached_input": 0.25  / _M, "output":  15.00 / _M},
    "gpt-5.4-mini":     {"input":  0.75 / _M, "cached_input": 0.075 / _M, "output":   4.50 / _M},
    "gpt-5.4-nano":     {"input":  0.20 / _M, "cached_input": 0.02  / _M, "output":   1.25 / _M},
    "gpt-5.4-pro":      {"input": 30.00 / _M,                              "output": 180.00 / _M},
    "gpt-5.2":          {"input":  1.75 / _M, "cached_input": 0.175 / _M, "output":  14.00 / _M},
    "gpt-5.2-pro":      {"input": 21.00 / _M,                              "output": 168.00 / _M},
    "gpt-5.1":          {"input":  1.25 / _M, "cached_input": 0.125 / _M, "output":  10.00 / _M},
    "gpt-5":            {"input":  1.25 / _M, "cached_input": 0.125 / _M, "output":  10.00 / _M},
    "gpt-5-mini":       {"input":  0.25 / _M, "cached_input": 0.025 / _M, "output":   2.00 / _M},
    "gpt-5-nano":       {"input":  0.05 / _M, "cached_input": 0.005 / _M, "output":   0.40 / _M},
    "gpt-5-pro":        {"input": 15.00 / _M,                              "output": 120.00 / _M},

    # ----- OpenAI GPT-4.1 family -----
    "gpt-4.1":          {"input":  2.00 / _M, "cached_input": 0.50  / _M, "output":   8.00 / _M},
    "gpt-4.1-mini":     {"input":  0.40 / _M, "cached_input": 0.10  / _M, "output":   1.60 / _M},
    "gpt-4.1-nano":     {"input":  0.10 / _M, "cached_input": 0.025 / _M, "output":   0.40 / _M},

    # ----- OpenAI GPT-4o family -----
    "gpt-4o":           {"input":  2.50 / _M, "cached_input": 1.25  / _M, "output":  10.00 / _M},
    "gpt-4o-2024-05-13":{"input":  5.00 / _M,                              "output":  15.00 / _M},
    "gpt-4o-mini":      {"input":  0.15 / _M, "cached_input": 0.075 / _M, "output":   0.60 / _M},

    # ----- OpenAI o-series (reasoning) -----
    "o1":               {"input": 15.00 / _M, "cached_input": 7.50  / _M, "output":  60.00 / _M},
    "o1-pro":           {"input":150.00 / _M,                              "output": 600.00 / _M},
    "o1-mini":          {"input":  1.10 / _M, "cached_input": 0.55  / _M, "output":   4.40 / _M},
    "o3":               {"input":  2.00 / _M, "cached_input": 0.50  / _M, "output":   8.00 / _M},
    "o3-pro":           {"input": 20.00 / _M,                              "output":  80.00 / _M},
    "o3-mini":          {"input":  1.10 / _M, "cached_input": 0.55  / _M, "output":   4.40 / _M},
    "o4-mini":          {"input":  1.10 / _M, "cached_input": 0.275 / _M, "output":   4.40 / _M},

    # ----- Anthropic Claude (Messages API) -----
    # Pricing source: https://www.anthropic.com/pricing#api
    # cached_input rate = "cache read" rate (10% of base input). Cache writes
    # (cache_creation_input_tokens) are slightly above base input rate
    # (~1.25x) but we lump them into the base input bucket below; this
    # under-counts cache-write cost by ~25% on those tokens.
    "claude-sonnet-4-6":   {"input": 3.00 / _M, "cached_input": 0.30 / _M, "output": 15.00 / _M},
    "claude-sonnet-4-5":   {"input": 3.00 / _M, "cached_input": 0.30 / _M, "output": 15.00 / _M},
    "claude-opus-4-7":     {"input":15.00 / _M, "cached_input": 1.50 / _M, "output": 75.00 / _M},
    "claude-haiku-4-5-20251001": {"input": 1.00 / _M, "cached_input": 0.10 / _M, "output": 5.00 / _M},
    "claude-haiku-4-5":          {"input": 1.00 / _M, "cached_input": 0.10 / _M, "output": 5.00 / _M},

    # ----- Google Gemini 3.x (Preview) -----
    # Gemini 3.1 Pro: tiered (≤200K = base; >200K input/output ≈ 2×). Base only.
    "gemini-3.1-pro-preview":          {"input": 2.00 / _M, "cached_input": 0.20  / _M, "output": 12.00 / _M},
    "gemini-3.1-flash-lite-preview":   {"input": 0.25 / _M, "cached_input": 0.025 / _M, "output":  1.50 / _M},
    # Live model: text-only rates listed; audio/video input would be more expensive.
    "gemini-3.1-flash-live-preview":   {"input": 0.75 / _M,                              "output":  4.50 / _M},
    "gemini-3-flash-preview":          {"input": 0.50 / _M, "cached_input": 0.05  / _M, "output":  3.00 / _M},

    # ----- Google Gemini 2.5 -----
    # Gemini 2.5 Pro: tiered. ≤200K = base; >200K = ~2× across all rates.
    "gemini-2.5-pro":                                {"input": 1.25 / _M, "cached_input": 0.125 / _M, "output": 10.00 / _M},
    "gemini-2.5-flash":                              {"input": 0.30 / _M, "cached_input": 0.03  / _M, "output":  2.50 / _M},
    "gemini-2.5-flash-lite":                         {"input": 0.10 / _M, "cached_input": 0.01  / _M, "output":  0.40 / _M},
    "gemini-2.5-flash-lite-preview-09-2025":         {"input": 0.10 / _M, "cached_input": 0.01  / _M, "output":  0.40 / _M},
    # Native audio (Live API): text-only rates listed; audio/video input is more expensive.
    "gemini-2.5-flash-native-audio-preview-12-2025": {"input": 0.50 / _M,                              "output":  2.00 / _M},
    # Computer Use: tiered. Caching not available.
    "gemini-2.5-computer-use-preview-10-2025":       {"input": 1.25 / _M,                              "output": 10.00 / _M},

    # ----- Google Gemini Image Generation (🍌) -----
    # These models charge image OUTPUT tokens at a much higher rate than the
    # underlying text model. Output rate below assumes image generation (the
    # dominant use case). Text-only output (rare) would be over-counted.
    # Reference: gemini-3-pro-image $120/M ≈ $0.134/image (1120 tokens, 1K-2K).
    #            gemini-2.5-flash-image $30/M  ≈ $0.039/image (1290 tokens, 1K).
    "gemini-3-pro-image-preview":     {"input": 2.00 / _M, "output": 120.00 / _M},
    "gemini-3.1-flash-image-preview": {"input": 0.50 / _M, "output":  60.00 / _M},
    "gemini-2.5-flash-image":         {"input": 0.30 / _M, "output":  30.00 / _M},
}


def normalize_model(model: str) -> str:
    """Strip provider-specific prefixes so the same model resolves regardless
    of whether the caller passed the bare id or the SDK-prefixed form
    (e.g. Gemini accepts both 'gemini-X' and 'models/gemini-X')."""
    if not model:
        return model
    if model.startswith("models/"):
        return model[len("models/"):]
    return model


def compute_cost(model: str, input_tokens: int, output_tokens: int,
                 cached_input_tokens: int = 0) -> float:
    """USD cost for a single LLM call. Returns 0.0 for unknown models so an
    unrecognized name never crashes the accumulator — a missing-model warning
    is logged elsewhere. Models without prompt caching support fall back to
    the standard input rate for any cached tokens (defensive: should be 0)."""
    p = PRICING.get(normalize_model(model))
    if not p:
        return 0.0
    in_tok = int(input_tokens or 0)
    out_tok = int(output_tokens or 0)
    cached = int(cached_input_tokens or 0)
    cached_rate = p.get("cached_input", p["input"])
    return (
        max(0, in_tok - cached) * p["input"]
        + cached * cached_rate
        + out_tok * p["output"]
    )
