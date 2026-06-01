# llm_client.py
import os
import threading
import time
import json
import contextvars
from pathlib import Path
from types import SimpleNamespace
from loguru import logger
from openai import OpenAI
from anthropic import Anthropic
from dotenv import load_dotenv
from google import genai
from google.genai import types
import base64
from io import BytesIO
from PIL import Image

from editppt.config import PROJECT_ROOT
from editppt.pricing import compute_cost, PRICING, normalize_model

load_dotenv(PROJECT_ROOT / ".env")

UPSTAGE_API_KEY = os.environ.get("UPSTAGE_API_KEY")
ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY")
OPENAI_API_KEY = os.environ.get("OPENAI_API_KEY")
GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY")

# ---- Token accumulator ----
# Monotonic counter; never reset implicitly. Pipeline code should call
# get_token_snapshot() / diff_tokens() to compute per-scope usage instead of
# relying on reset+read. The lock guards both the dict and read-modify-write
# accumulation across threads (e.g. dispatcher prefetch).
_token_accumulator = {
    "input_tokens": 0,
    "output_tokens": 0,
    "cached_input_tokens": 0,
    "cost_usd": 0.0,
}
# Track unknown models seen (so we warn once instead of spamming)
_unknown_models_warned: set[str] = set()
_token_lock = threading.Lock()

# ---- Per-call JSONL persistence ----
# Every successful or failed LLM call appends one line so cost is recoverable
# even if the app crashes before the in-memory record is read.
#
# `_token_log_context` is a ContextVar so background workers (e.g. the
# dispatcher's prefetch ThreadPoolExecutor) can carry their own slide_index
# without stomping on siblings or the main thread. Callers that want
# isolation must wrap submits with `contextvars.copy_context().run(...)`.
_token_log_path: Path | None = None
_token_log_context: contextvars.ContextVar = contextvars.ContextVar(
    "_token_log_context", default={}
)
# Guards file appends only; the context is per-ContextVar copy and needs no lock.
_token_log_lock = threading.Lock()

# ---- Per-request summary JSONL ----
# One line per user request (CLI / web pipeline / web retry). Carries
# wall-clock submit→done time plus delta token totals for fast experiment
# analysis without rolling up every per-call line in token_usage.jsonl.
_request_summary_path: Path | None = None
_request_summary_lock = threading.Lock()


def set_request_summary_path(path: "Path | str | None"):
    """Set the JSONL file for per-request summaries. None disables it."""
    global _request_summary_path
    if path is None:
        _request_summary_path = None
        return
    p = Path(path)
    p.parent.mkdir(parents=True, exist_ok=True)
    _request_summary_path = p


def log_request_summary(*, request_id, user_input: str,
                        submit_ts: float, done_ts: float,
                        status: str, scope: str,
                        task_count: int = 0, slide_count: int = 0,
                        input_tokens: int = 0, output_tokens: int = 0,
                        cached_input_tokens: int = 0, cost_usd: float = 0.0,
                        **extra):
    """Append one summary line per request. Best-effort, never raises.

    elapsed_seconds is computed from submit_ts (when the user submitted —
    Flask receive in web, input() return in CLI) to done_ts (pipeline
    finally), so it includes queue wait, planner, COM ops, and validators.
    """
    if _request_summary_path is None:
        return
    try:
        record = {
            "ts": done_ts,
            "submit_ts": submit_ts,
            "done_ts": done_ts,
            "elapsed_seconds": round(float(done_ts - submit_ts), 4),
            "request_id": request_id,
            "user_input": user_input,
            "status": status,
            "scope": scope,
            "task_count": int(task_count or 0),
            "slide_count": int(slide_count or 0),
            "input_tokens": int(input_tokens or 0),
            "output_tokens": int(output_tokens or 0),
            "cached_input_tokens": int(cached_input_tokens or 0),
            "cost_usd": round(float(cost_usd or 0.0), 8),
        }
        if extra:
            record.update(extra)
        line = json.dumps(record, ensure_ascii=False) + "\n"
        with _request_summary_lock:
            with open(_request_summary_path, "a", encoding="utf-8") as f:
                f.write(line)
    except Exception as e:
        logger.debug(f"[request_summary] write failed: {e}")


def set_token_log_path(path: "Path | str | None"):
    """Set the JSONL file path for per-call token logging. None disables it."""
    global _token_log_path
    if path is None:
        _token_log_path = None
        return
    p = Path(path)
    p.parent.mkdir(parents=True, exist_ok=True)
    _token_log_path = p


def set_token_log_context(**kwargs):
    """Update metadata attached to subsequent token-usage log lines.
    Pass None as a value to remove a key. Each ContextVar copy (e.g. a
    prefetch worker running under copy_context().run) sees only its own
    state."""
    cur = dict(_token_log_context.get())
    for k, v in kwargs.items():
        if v is None:
            cur.pop(k, None)
        else:
            cur[k] = v
    _token_log_context.set(cur)


def reset_token_counter():
    """Reset the monotonic accumulator. Use only at session boundaries (e.g.
    file open). Per-request accounting must use snapshot diffs instead."""
    with _token_lock:
        _token_accumulator["input_tokens"] = 0
        _token_accumulator["output_tokens"] = 0
        _token_accumulator["cached_input_tokens"] = 0
        _token_accumulator["cost_usd"] = 0.0


def get_token_count() -> dict:
    """Snapshot of the accumulated token counts (legacy alias)."""
    return get_token_snapshot()


def get_token_snapshot() -> dict:
    """Snapshot of the monotonic token counters."""
    with _token_lock:
        return dict(_token_accumulator)


def diff_tokens(start: dict, end: dict) -> dict:
    """Return end - start for each token/cost field. Negative deltas (caused by
    an intervening explicit reset) are clamped to 0."""
    keys = set(_token_accumulator.keys()) | set(start.keys()) | set(end.keys())
    return {k: max(0, end.get(k, 0) - start.get(k, 0)) for k in keys}


def _accumulate(model: str, input_tokens: int, output_tokens: int,
                cached_input_tokens: int = 0) -> float:
    """Accumulate token counts and the resulting USD cost. Returns the cost
    for this single call so callers can also log it."""
    cost = compute_cost(model, input_tokens, output_tokens, cached_input_tokens)
    canon = normalize_model(model)
    if canon and canon not in PRICING and canon not in _unknown_models_warned:
        _unknown_models_warned.add(canon)
        logger.warning(f"[token_cost] No pricing entry for model '{model}' — cost will be reported as 0. Update editppt/pricing.py")
    with _token_lock:
        _token_accumulator["input_tokens"] += int(input_tokens or 0)
        _token_accumulator["output_tokens"] += int(output_tokens or 0)
        _token_accumulator["cached_input_tokens"] += int(cached_input_tokens or 0)
        _token_accumulator["cost_usd"] += float(cost)
    return cost


def _extract_openai_usage(response_or_obj) -> tuple[int, int, int]:
    """Pull (input, output, cached_input) from an OpenAI Responses-API response
    or a usage object. Returns zeros on any failure."""
    try:
        usage = getattr(response_or_obj, "usage", None)
        if usage is None:
            usage = response_or_obj
        if usage is None:
            return 0, 0, 0
        input_tokens = getattr(usage, "input_tokens", 0) or 0
        output_tokens = getattr(usage, "output_tokens", 0) or 0
        cached = 0
        details = getattr(usage, "input_tokens_details", None)
        if details is not None:
            cached = getattr(details, "cached_tokens", 0) or 0
        return int(input_tokens), int(output_tokens), int(cached)
    except Exception:
        return 0, 0, 0


def _extract_gemini_usage(meta) -> tuple[int, int, int]:
    try:
        if meta is None:
            return 0, 0, 0
        input_tokens = getattr(meta, "prompt_token_count", 0) or 0
        output_tokens = getattr(meta, "candidates_token_count", 0) or 0
        cached = getattr(meta, "cached_content_token_count", 0) or 0
        return int(input_tokens), int(output_tokens), int(cached)
    except Exception:
        return 0, 0, 0


def _try_extract_usage_from_exception(e: Exception, provider: str) -> tuple[int, int, int]:
    """Best-effort: some SDKs attach a partial response with usage on failure.
    Returns zeros if nothing usable is found."""
    try:
        for attr in ("response", "body"):
            obj = getattr(e, attr, None)
            if obj is None:
                continue
            if provider in ("openai", "upstage"):
                ext = _extract_openai_usage(obj)
                if any(ext):
                    return ext
                try:
                    usage = obj.get("usage") if hasattr(obj, "get") else None
                    if usage is not None:
                        ext = _extract_openai_usage(usage)
                        if any(ext):
                            return ext
                except Exception:
                    pass
            elif provider == "gemini":
                meta = getattr(obj, "usage_metadata", None) or obj
                ext = _extract_gemini_usage(meta)
                if any(ext):
                    return ext
    except Exception:
        pass
    return 0, 0, 0


def _log_token_event(model: str, provider: str, input_tokens: int, output_tokens: int,
                     cached_input_tokens: int = 0, cost_usd: float = 0.0,
                     latency_seconds: float = 0.0,
                     error: str | None = None):
    """Append one JSONL line to the token-usage log. Best-effort, never raises."""
    if _token_log_path is None:
        return
    try:
        # ContextVar.get() is lock-free; each context carries its own dict.
        ctx = dict(_token_log_context.get())
        record = {
            "ts": time.time(),
            "model": model,
            "provider": provider,
            "input_tokens": int(input_tokens or 0),
            "output_tokens": int(output_tokens or 0),
            "cached_input_tokens": int(cached_input_tokens or 0),
            "cost_usd": round(float(cost_usd or 0.0), 8),
            "latency_seconds": round(float(latency_seconds or 0.0), 4),
            **ctx,
        }
        if error is not None:
            record["error"] = error
        line = json.dumps(record, ensure_ascii=False) + "\n"
        # File append must be serialized so concurrent workers don't interleave.
        with _token_log_lock:
            with open(_token_log_path, "a", encoding="utf-8") as f:
                f.write(line)
    except Exception as e:
        logger.debug(f"[token_log] write failed: {e}")


def record_external_llm_usage(model: str, provider: str,
                              response=None, error: Exception | None = None,
                              latency_seconds: float = 0.0) -> float:
    """Account for an LLM call that bypassed call_llm / call_llm_gemini.
    Used by code that must hit an SDK directly (e.g. native image-generation
    calls returning binary content). Pass `response` on success, `error` on
    failure. `latency_seconds` should be the wall-clock time the caller
    measured around the SDK invocation. Returns the per-call cost so the
    caller can log it if desired."""
    in_tok, out_tok, cached = 0, 0, 0
    if error is not None:
        in_tok, out_tok, cached = _try_extract_usage_from_exception(error, provider)
    elif response is not None:
        if provider in ("openai", "upstage"):
            in_tok, out_tok, cached = _extract_openai_usage(response)
        elif provider == "gemini":
            in_tok, out_tok, cached = _extract_gemini_usage(getattr(response, "usage_metadata", None))

    cost = 0.0
    if any((in_tok, out_tok, cached)):
        cost = _accumulate(model, in_tok, out_tok, cached)
    err_name = type(error).__name__ if error is not None else None
    _log_token_event(model, provider, in_tok, out_tok, cached, cost, latency_seconds, error=err_name)
    return cost


def get_api_key_and_provider(model: str):
    """
    Determine which provider and API key to use based on the model name.
    """
    provider = None
    api_key = None

    # OpenAI
    if model.startswith("gpt-"):
        provider = "openai"
        api_key = OPENAI_API_KEY

    # Anthropic
    elif model.startswith("claude-"):
        provider = "anthropic"
        api_key = ANTHROPIC_API_KEY

    # Gemini
    elif model.startswith("gemini-"):
        provider = "gemini"
        api_key = GEMINI_API_KEY

    # Upstage(Solar)
    elif model.startswith("solar-"):
        provider = "upstage"
        api_key = UPSTAGE_API_KEY

    else:
        raise ValueError(f"Unsupported model: {model}")

    if api_key is None:
        raise ValueError(f"API key not set for {provider}. (model={model})")

    return provider, api_key


def is_anthropic_model(model: str) -> bool:
    """Cheap provider check by model-name prefix (no API key required)."""
    return model.startswith("claude-") or model.startswith("anthropic")


def get_client_for_model(model: str):
    """
    Return an OpenAI-style client for the given model.
    """
    provider, api_key = get_api_key_and_provider(model)

    if provider == "openai":
        client = OpenAI(api_key=api_key)
        
    elif provider == "upstage":
        client = OpenAI(api_key=api_key, base_url="https://api.upstage.ai/v1")

    elif provider == "anthropic":
        client = Anthropic(api_key=api_key)

    elif provider == "gemini":
        client = genai.Client(api_key=api_key)
    
    else:
        raise ValueError(f"Unknown provider: {provider}")

    return client, provider


class APIKeyError(Exception):
    """Raised when an API key is invalid, expired, or has insufficient quota."""
    pass


def _classify_api_error(e: Exception, provider: str) -> str | None:
    """Classify an LLM exception as an API key issue and return a user message, or None."""
    err_str = str(e).lower()
    err_type = type(e).__name__

    # OpenAI errors
    if "insufficient_quota" in err_str or "exceeded your current quota" in err_str:
        return f"[{provider.upper()}] API key quota exhausted. Please check your billing at https://platform.openai.com/usage"
    if "invalid_api_key" in err_str or "incorrect api key" in err_str:
        return f"[{provider.upper()}] Invalid API key. Please check your API key in Settings."
    if err_type == "AuthenticationError" or "401" in err_str:
        return f"[{provider.upper()}] Authentication failed. Your API key may be invalid or revoked."
    if "rate_limit" in err_str or "429" in err_str:
        if "quota" in err_str or "billing" in err_str:
            return f"[{provider.upper()}] API quota exceeded. Please check your plan and billing."
        return None  # transient rate limit, not a key issue

    # Gemini errors
    if "api_key_invalid" in err_str or "api key not valid" in err_str:
        return f"[{provider.upper()}] Invalid API key. Please check your Gemini API key."
    if "resource_exhausted" in err_str or "quota" in err_str:
        return f"[{provider.upper()}] API quota exhausted. Please check your Google AI Studio billing."

    # Anthropic errors
    if provider == "anthropic":
        if "invalid x-api-key" in err_str or "invalid api key" in err_str:
            return f"[{provider.upper()}] Invalid API key. Please check ANTHROPIC_API_KEY in .env."
        if "credit balance" in err_str or "billing" in err_str:
            return f"[{provider.upper()}] Anthropic credit/billing issue. Please check your console.anthropic.com balance."

    return None


# ─────────────────────────────────────────────────────────────────────────
# Anthropic adapters
# ─────────────────────────────────────────────────────────────────────────
# The rest of the codebase consumes responses in OpenAI Responses-API shape:
# `response.output_text`, `response.output[i].{type,name,arguments,call_id}`,
# `response.usage.{input_tokens,output_tokens,input_tokens_details.cached_tokens}`,
# and `response.model_dump()` for logging. The functions below convert
# OpenAI-style requests to Anthropic Messages-API requests and wrap the
# Anthropic response back into a SimpleNamespace with those attributes so
# no caller needs to branch on provider.


def _openai_to_anthropic_messages(messages):
    """Extract system text and convert content blocks to Anthropic shape.

    Returns (system_text_or_None, [anthropic_messages]). System messages from
    the OpenAI list are joined and moved to Anthropic's top-level `system`
    parameter. Image and text blocks are remapped from
    `input_text`/`input_image` to Anthropic's `text`/`image`.
    """
    system_parts = []
    new_messages = []
    for m in messages:
        role = m.get("role")
        content = m.get("content")
        if role == "system":
            if isinstance(content, str):
                system_parts.append(content)
            elif isinstance(content, list):
                for block in content:
                    if isinstance(block, dict) and block.get("type") in ("input_text", "text"):
                        system_parts.append(block.get("text", ""))
            continue

        new_content = []
        if isinstance(content, str):
            new_content = [{"type": "text", "text": content}]
        elif isinstance(content, list):
            for block in content:
                if not isinstance(block, dict):
                    continue
                btype = block.get("type")
                if btype in ("input_text", "text"):
                    new_content.append({"type": "text", "text": block.get("text", "")})
                elif btype == "input_image":
                    image_url = block.get("image_url", "") or ""
                    if image_url.startswith("data:"):
                        header, b64 = image_url.split(",", 1)
                        media_type = header[5:].split(";")[0] or "image/png"
                        new_content.append({
                            "type": "image",
                            "source": {
                                "type": "base64",
                                "media_type": media_type,
                                "data": b64,
                            },
                        })
                    else:
                        new_content.append({
                            "type": "image",
                            "source": {"type": "url", "url": image_url},
                        })
                # Any other block type is intentionally skipped; OpenAI-only
                # primitives don't have an Anthropic equivalent.
        new_messages.append({"role": role, "content": new_content})

    return ("\n\n".join(system_parts) if system_parts else None, new_messages)


def _openai_to_anthropic_tools(tools):
    """Convert OpenAI tool schema list to Anthropic `tools` list.

    OpenAI: {"type": "function", "name", "description", "parameters": {schema}}
    Anthropic: {"name", "description", "input_schema": {schema}}
    """
    if not tools:
        return None
    out = []
    for t in tools:
        out.append({
            "name": t["name"],
            "description": t.get("description", "") or "",
            "input_schema": t.get("parameters") or {"type": "object", "properties": {}},
        })
    return out


def _openai_to_anthropic_tool_choice(tool_choice):
    if tool_choice is None:
        return None
    if tool_choice == "auto":
        return {"type": "auto"}
    if tool_choice == "required":
        return {"type": "any"}
    if tool_choice == "none":
        return None  # Anthropic disables tools by omitting them; absent type=none
    if isinstance(tool_choice, dict):
        fn = tool_choice.get("function") or {}
        if fn.get("name"):
            return {"type": "tool", "name": fn["name"]}
        if tool_choice.get("name"):
            return {"type": "tool", "name": tool_choice["name"]}
    return {"type": "auto"}


def _extract_anthropic_usage(usage) -> tuple[int, int, int]:
    """Anthropic usage → (total_input, output, cached_input).

    Anthropic's `input_tokens` excludes cache reads/writes. Total billable
    input = input + cache_creation + cache_read. We report cache_read as
    the "cached" subset so pricing.py applies the discounted rate to it.
    """
    try:
        if usage is None:
            return 0, 0, 0
        base_in = int(getattr(usage, "input_tokens", 0) or 0)
        cache_create = int(getattr(usage, "cache_creation_input_tokens", 0) or 0)
        cache_read = int(getattr(usage, "cache_read_input_tokens", 0) or 0)
        out_tok = int(getattr(usage, "output_tokens", 0) or 0)
        total_in = base_in + cache_create + cache_read
        return total_in, out_tok, cache_read
    except Exception:
        return 0, 0, 0


def _build_anthropic_adapter_response(anthropic_response, structured_schema_name: str | None = None):
    """Wrap an Anthropic Messages response in an OpenAI-Responses-API shape.

    structured_schema_name: when the caller requested structured output, we
    forced a single tool of that name. Its `input` (parsed JSON) is exposed
    as `output_text` and excluded from `output` so the caller's
    `json.loads(response.output_text)` works unchanged.
    """
    output_items = []
    text_parts = []
    structured_output_json = None

    for block in getattr(anthropic_response, "content", []) or []:
        btype = getattr(block, "type", None)
        if btype == "text":
            text_parts.append(getattr(block, "text", "") or "")
        elif btype == "tool_use":
            name = getattr(block, "name", "")
            tool_input = getattr(block, "input", {}) or {}
            try:
                args_str = json.dumps(tool_input, ensure_ascii=False)
            except Exception:
                args_str = "{}"
            if structured_schema_name and name == structured_schema_name:
                structured_output_json = args_str
                continue  # don't surface the forced tool as a function call
            output_items.append(SimpleNamespace(
                type="function_call",
                name=name,
                arguments=args_str,
                call_id=getattr(block, "id", "") or "",
            ))

    if structured_output_json is not None:
        output_text = structured_output_json
    else:
        output_text = "".join(text_parts)

    in_tok, out_tok, cached = _extract_anthropic_usage(getattr(anthropic_response, "usage", None))
    adapter_usage = SimpleNamespace(
        input_tokens=in_tok,
        output_tokens=out_tok,
        input_tokens_details=SimpleNamespace(cached_tokens=cached),
    )

    def _model_dump():
        raw_content = []
        for block in getattr(anthropic_response, "content", []) or []:
            btype = getattr(block, "type", None)
            if btype == "text":
                raw_content.append({"type": "text", "text": getattr(block, "text", "")})
            elif btype == "tool_use":
                raw_content.append({
                    "type": "tool_use",
                    "id": getattr(block, "id", ""),
                    "name": getattr(block, "name", ""),
                    "input": getattr(block, "input", {}),
                })
        return {
            "model": getattr(anthropic_response, "model", None),
            "stop_reason": getattr(anthropic_response, "stop_reason", None),
            "content": raw_content,
            "usage": {
                "input_tokens": in_tok,
                "output_tokens": out_tok,
                "cached_input_tokens": cached,
            },
            "output_text": output_text,
        }

    return SimpleNamespace(
        output=output_items,
        output_text=output_text,
        usage=adapter_usage,
        model_dump=_model_dump,
    )


def _call_anthropic(client, model, messages, tools, tool_choice, **kwargs):
    """Anthropic Messages call with OpenAI-Responses-API-shaped response.

    Translates structured-output requests (kwargs["text"] with json_schema)
    into a forced single-tool call. `prompt_cache_key` is OpenAI-only and
    ignored here; Anthropic prompt caching is wired via cache_control
    markers on system + last tool (acts as a breakpoint covering the static
    prefix).
    """
    system_text, anth_messages = _openai_to_anthropic_messages(messages)

    structured_schema_name = None
    text_param = kwargs.pop("text", None)
    if isinstance(text_param, dict):
        fmt = text_param.get("format") or {}
        if fmt.get("type") == "json_schema":
            structured_schema_name = fmt.get("name") or "structured_output"
            # Build the forced tool in OpenAI shape (`parameters` not
            # `input_schema`) so the shared converter below produces the
            # right Anthropic payload. Building it in Anthropic shape would
            # cause `_openai_to_anthropic_tools` to read a missing
            # `parameters` key and emit an empty schema.
            forced_tool = {
                "type": "function",
                "name": structured_schema_name,
                "description": "Return the response in the required JSON schema.",
                "parameters": fmt.get("schema") or {"type": "object", "properties": {}},
            }
            tools = list(tools or []) + [forced_tool]
            tool_choice = {"type": "tool", "name": structured_schema_name}
            # OpenAI strict json_schema mode forces the model to populate every
            # field in the schema. Anthropic's forced tool call has no such
            # guarantee — without an explicit directive the model may emit an
            # empty `input` object. Append a hard reminder to the system text.
            schema_directive = (
                f"\n\nIMPORTANT: You MUST invoke the `{structured_schema_name}` tool "
                f"and populate every required field of its input_schema with a "
                f"concrete value derived from the user's request. Never return an "
                f"empty tool input."
            )
            system_text = (system_text + schema_directive) if system_text else schema_directive.lstrip()

    anth_tools = _openai_to_anthropic_tools(tools)
    # tool_choice may already be Anthropic-shaped (from our structured-output
    # forcing above) or OpenAI-shaped (from the caller). Pass-through if
    # already native; otherwise convert.
    if isinstance(tool_choice, dict) and tool_choice.get("type") in ("auto", "any", "tool", "none"):
        anth_tool_choice = tool_choice
    else:
        anth_tool_choice = _openai_to_anthropic_tool_choice(tool_choice)

    payload = {
        "model": model,
        "messages": anth_messages,
        "max_tokens": int(kwargs.pop("max_tokens", 8192)),
    }
    # Anthropic prompt caching: cache_control markers act as breakpoints.
    # Mark system (if any) and the last tool (if any) so the static prefix
    # — system prompt + tool schemas — gets cached. Sonnet caches ≥1024 tok
    # automatically; below that the marker is a no-op so cost-of-adding=0.
    if system_text:
        payload["system"] = [{
            "type": "text",
            "text": system_text,
            "cache_control": {"type": "ephemeral"},
        }]
    if anth_tools:
        anth_tools = list(anth_tools)
        last = dict(anth_tools[-1])
        last["cache_control"] = {"type": "ephemeral"}
        anth_tools[-1] = last
        payload["tools"] = anth_tools
    if anth_tool_choice:
        payload["tool_choice"] = anth_tool_choice

    if "temperature" in kwargs:
        payload["temperature"] = kwargs.pop("temperature")
    else:
        payload["temperature"] = 0.1

    # Strip any leftover OpenAI-only kwargs silently (e.g., prompt_cache_key
    # already removed by caller, but defensive).
    for k in ("prompt_cache_key", "metadata", "store", "user"):
        kwargs.pop(k, None)
    payload.update(kwargs)  # any remaining kwargs pass through

    raw = client.messages.create(**payload)
    return _build_anthropic_adapter_response(raw, structured_schema_name=structured_schema_name)


def call_llm(
    model: str,
    messages,
    tools=None,
    tool_choice=None,
    prompt_cache_key: str | None = None,
    **kwargs,
):
    """
    Unified LLM call wrapper.

    prompt_cache_key: Optional routing hint for OpenAI's automatic prompt
    caching. When the same key is reused across requests with a stable
    prefix (system prompt + tools), responses route to the same shard for
    higher cache-hit rates. Ignored by non-OpenAI providers.
    """
    client, provider = get_client_for_model(model)

    t0 = time.monotonic()
    try:
        if provider == "anthropic":
            response = _call_anthropic(
                client, model, messages, tools, tool_choice, **kwargs
            )
            latency = time.monotonic() - t0
            in_tok = response.usage.input_tokens
            out_tok = response.usage.output_tokens
            cached = response.usage.input_tokens_details.cached_tokens
            cost = _accumulate(model, in_tok, out_tok, cached)
            _log_token_event(model, provider, in_tok, out_tok, cached, cost, latency)
            return response

        # OpenAI / Upstage path (Responses API).
        # Some models (e.g. gpt-5) reject the temperature parameter. Callers
        # can pass `temperature=None` to opt out; the default remains 0.1.
        payload = {
            "model": model,
            "input": messages,
        }
        explicit_temp = kwargs.pop("temperature", "__unset__")
        if explicit_temp == "__unset__":
            payload["temperature"] = 0.1
        elif explicit_temp is not None:
            payload["temperature"] = explicit_temp
        if tools is not None:
            payload["tools"] = tools
        if tool_choice is not None:
            payload["tool_choice"] = tool_choice
        if prompt_cache_key is not None and provider in ("openai", "upstage"):
            payload["prompt_cache_key"] = prompt_cache_key
        payload.update(kwargs)

        response = client.responses.create(**payload)
        latency = time.monotonic() - t0
        in_tok, out_tok, cached = _extract_openai_usage(response)
        cost = _accumulate(model, in_tok, out_tok, cached)
        _log_token_event(model, provider, in_tok, out_tok, cached, cost, latency)
        return response
    except Exception as e:
        latency = time.monotonic() - t0
        # Best-effort: capture any usage attached to the error response so a
        # billable failure still increments the accumulator.
        in_tok, out_tok, cached = _try_extract_usage_from_exception(e, provider)
        cost = 0.0
        if any((in_tok, out_tok, cached)):
            cost = _accumulate(model, in_tok, out_tok, cached)
        _log_token_event(model, provider, in_tok, out_tok, cached, cost, latency, error=type(e).__name__)
        key_msg = _classify_api_error(e, provider)
        if key_msg:
            logger.error(key_msg)
            raise APIKeyError(key_msg) from e
        logger.warning(f"[call_llm] LLM call failed ({type(e).__name__}): {e}")
        raise


def call_llm_gemini(
    model: str,
    messages: str,
    image: base64 = None,
    response_schema: dict | None = None,
):
    """
    Gemini-specific call wrapper (google-genai v1.0+).

    When `response_schema` is provided, the response is forced to JSON matching
    the schema (response_mime_type=application/json + response_schema).
    """
    client, provider = get_client_for_model(model)
    config = None
    if response_schema is not None:
        config = types.GenerateContentConfig(
            response_mime_type="application/json",
            response_schema=response_schema,
            temperature=0.1,
        )
    t0 = time.monotonic()
    try:
        response = client.models.generate_content(
            model=model,
            contents=[
                types.Part.from_bytes(data=image, mime_type="image/png"),
                messages
            ],
            config=config,
        )
        latency = time.monotonic() - t0
        in_tok, out_tok, cached = _extract_gemini_usage(getattr(response, "usage_metadata", None))
        cost = _accumulate(model, in_tok, out_tok, cached)
        _log_token_event(model, provider, in_tok, out_tok, cached, cost, latency)
        return response.text

    except Exception as e:
        latency = time.monotonic() - t0
        in_tok, out_tok, cached = _try_extract_usage_from_exception(e, provider)
        cost = 0.0
        if any((in_tok, out_tok, cached)):
            cost = _accumulate(model, in_tok, out_tok, cached)
        _log_token_event(model, provider, in_tok, out_tok, cached, cost, latency, error=type(e).__name__)
        key_msg = _classify_api_error(e, provider)
        if key_msg:
            logger.error(key_msg)
            raise APIKeyError(key_msg) from e
        logger.warning(f"[call_llm_gemini] Gemini call failed ({type(e).__name__}): {e}")
        raise