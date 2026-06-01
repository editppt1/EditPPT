"""Visual fixer agent.

Receives a list of vision defects (each with a concrete numeric ActionableFix)
from the vision validator and translates them into tool calls. Loops up to
max_attempts, re-validating after each round so the fix can chase residual
defects without rolling back the upstream text edit.

Triggered by base_agent.run on vision-validation failure (replaces the prior
rollback path) so semantic intent established by text validation is preserved.
"""

import json

from editppt.tools.agent_tool_registry import TOOL_METADATA
from editppt.tools.tools import FUNCTION_MAP, clamp_shapes_to_slide, clamp_text_to_slide
from editppt.utils.llm_client import (
    call_llm,
    APIKeyError,
    set_token_log_context,
    _token_log_context,
    is_anthropic_model,
)
from editppt.utils.utils import slim_for_vision
from editppt.utils.logger_manual import init_logger
from editppt.prompts import create_visual_fixer_agent_system_prompt

logger = init_logger()


# Tools the fixer is allowed to call. Two fixer-specific bulk tools (no char
# range / no per-run granularity) cover text fixes; layout tools cover
# geometry. set_text_style / manage_bullet_points / apply_visual_style are
# intentionally excluded — too granular or out of scope for visual fixes.
#
# Text-structure tools (edit_text_rewrite, edit_text_replace) are included
# because vision validator frequently reports "missing line breaks / merged
# paragraphs" which font-size and layout adjustments cannot fix. Fixer is
# expected to use these only when defects explicitly call for restoring text
# structure, not for repeated geometry tweaks.
_FIXER_TOOL_NAMES = {
    "set_shape_font_size",
    "set_shape_paragraph_spacing",
    "adjust_layout",
    "align_shapes",
    "distribute_shapes",
    "edit_text_rewrite",
    "edit_text_replace",
}

# Tools whose post-execution clamping must use clamp_text_to_slide (font/text
# changes) rather than clamp_shapes_to_slide (geometric changes).
_TEXT_CLAMP_TOOLS = {
    "set_shape_font_size",
    "set_shape_paragraph_spacing",
    "edit_text_rewrite",
    "edit_text_replace",
}


class VisualFixerAgent:
    def __init__(self, container, model: str, tools_schema: list[dict]):
        self.container = container
        self.model = model
        self.tools_schema = [t for t in tools_schema if t.get("name") in _FIXER_TOOL_NAMES]

    @classmethod
    def create(cls, container, model: str, tools_schema: list[dict]):
        return cls(container, model, tools_schema)

    def run(
        self,
        page_number: int,
        defects: list[dict],
        slide_json: dict | None,
        vision_validator_agent,
        agent_request: str,
        used_tools: list[dict],
        parser=None,
        max_attempts: int = 3,
    ) -> bool:
        """Iteratively repair vision defects.

        Returns True if vision validation passes within max_attempts, False if
        the budget is exhausted or the input is unusable. On False, the caller
        should accept the slide as-is — the text intent is already validated
        and preserved.
        """
        if not defects:
            # Caller invoked us on a vision failure but provided no defects.
            # vision_validator returns issues=[] when its response cannot be
            # parsed; without specific fix plans we have nothing to act on.
            logger.warning(
                "VisualFixer: invoked with empty defects (vision validator response unparsable); aborting"
            )
            return False
        if vision_validator_agent is None:
            logger.warning("VisualFixer: no vision_validator_agent available; cannot iterate")
            return False

        # Stash agent_request so _execute_tool can inject it into tools that
        # require it (edit_text_rewrite, replace_table_text, etc.).
        self._agent_request = agent_request

        # Re-tag token usage as visual_fixer for the duration so log entries
        # don't inherit the upstream specialist's agent_type/attempt_idx.
        # set_token_log_context(value=None) removes a key, so passing the
        # captured prior values (None when missing) restores correctly.
        prior_ctx = dict(_token_log_context.get())
        prior_agent = prior_ctx.get("agent_type")
        prior_attempt = prior_ctx.get("attempt_idx")

        # Accumulate every tool the fixer executes so subsequent vision passes
        # don't compress those shapes via scope_for_used_tools. The original
        # base_agent tools come first; fixer-emitted calls append per attempt.
        accumulated_tools: list[dict] = list(used_tools or [])

        current_defects = defects
        current_slide_json = slide_json
        try:
            for attempt in range(1, max_attempts + 1):
                set_token_log_context(agent_type="visual_fixer", attempt_idx=attempt)

                tool_calls = self._plan_tool_calls(current_defects, current_slide_json)
                if not tool_calls:
                    logger.warning(f"VisualFixer attempt {attempt}: no tool calls produced; aborting")
                    return False

                for tc in tool_calls:
                    logger.info(f"VisualFixer attempt {attempt} tool: {tc['name']}({tc['arguments']})")

                executed_any = False
                for tc in tool_calls:
                    try:
                        self._execute_tool(tc["name"], tc["arguments"], current_slide_json)
                        executed_any = True
                    except Exception as e:
                        logger.warning(f"VisualFixer attempt {attempt}: tool {tc['name']} failed: {e}")
                if not executed_any:
                    logger.warning(f"VisualFixer attempt {attempt}: every tool call raised; aborting")
                    return False

                accumulated_tools.extend(tool_calls)

                # Skip re-parse + re-validation on the final attempt — neither
                # the parser cache update nor the gemini round trip can change
                # the outcome (no more iterations possible). For
                # max_attempts=1 this means a single fix shot with zero
                # validation overhead; for max_attempts>1 it still skips the
                # last (otherwise wasted) validation round.
                if attempt >= max_attempts:
                    logger.info(
                        f"VisualFixer: attempt {attempt}/{max_attempts} executed; "
                        f"skipping re-validation (no remaining attempts)"
                    )
                    return False

                # Re-parse so the next vision call reads post-fix numeric state
                # from the JSON (image is exported fresh by vision_validator).
                # Without this the validator reads stale Position/Size/Font
                # values and keeps reporting the original defect.
                if parser is not None:
                    try:
                        current_slide_json = parser.process(page_number, force=True)
                    except Exception as e:
                        logger.warning(
                            f"VisualFixer attempt {attempt}: re-parse failed ({e}); aborting"
                        )
                        return False

                valid, reason, new_issues = vision_validator_agent.process(
                    page_number=page_number,
                    agent_request=agent_request,
                    parsed_contents=current_slide_json,
                    used_tools=accumulated_tools,
                )
                if valid:
                    logger.info(f"VisualFixer: defects resolved on attempt {attempt}")
                    return True

                logger.info(f"VisualFixer attempt {attempt} did not converge: {reason}")
                current_defects = new_issues

            logger.warning(f"VisualFixer: exhausted {max_attempts} attempts; accepting slide as-is")
            return False
        finally:
            set_token_log_context(agent_type=prior_agent, attempt_idx=prior_attempt)

    def _plan_tool_calls(self, defects: list[dict], slide_json: dict | None) -> list[dict]:
        system_prompt = create_visual_fixer_agent_system_prompt()
        # Slim the JSON to the same compact shape vision_validator emitted the
        # defects against. Saves ~80% of slide-JSON tokens and keeps field
        # names (id, left/top/width/height, max_font_size) consistent with
        # ActionableFix references.
        slim = slim_for_vision(slide_json) if isinstance(slide_json, dict) else slide_json
        user_payload = {
            "defects": defects,
            "slide_json": slim,
        }
        user_prompt = (
            "Vision defects and current slide state:\n"
            f"{json.dumps(user_payload, ensure_ascii=False, indent=2)}\n\n"
            "Emit one tool call per defect to apply each ActionableFix literally."
        )
        messages = [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_prompt},
        ]
        tc = {"type": "any"} if is_anthropic_model(self.model) else "auto"
        try:
            set_token_log_context(component="visual_fixer")
            response = call_llm(
                model=self.model,
                messages=messages,
                tools=self.tools_schema,
                tool_choice=tc,
                prompt_cache_key="visual_fixer",
            )
        except APIKeyError:
            raise
        except Exception as e:
            logger.error(f"VisualFixer LLM call failed: {e}")
            return []

        tool_calls = []
        for item in response.output:
            if item.type == "function_call":
                args = item.arguments
                if isinstance(args, str):
                    try:
                        args = json.loads(args)
                    except json.JSONDecodeError:
                        logger.warning(f"VisualFixer: invalid args JSON for {item.name}: {args[:200]}")
                        continue
                tool_calls.append({"name": item.name, "arguments": args, "call_id": item.call_id})
        return tool_calls

    def _execute_tool(self, name: str, args: dict, slide_json: dict | None = None):
        if name not in FUNCTION_MAP:
            raise ValueError(f"VisualFixer: unknown tool '{name}'")
        meta = TOOL_METADATA.get(name)
        if meta:
            # Mirror base_agent's meta injection so text-structure tools work
            # (edit_text_rewrite needs slide_json AND agent_request to drive
            # the internal style_mapper).
            if meta.needs_slide_json:
                args["slide_json"] = slide_json
            if meta.needs_agent_request:
                args["agent_request"] = getattr(self, "_agent_request", "")
            if meta.needs_container:
                args["container"] = self.container
            if meta.cleanup_false_args:
                for key in meta.cleanup_false_args:
                    if args.get(key) is False:
                        args.pop(key)

        result = FUNCTION_MAP[name](self.container.prs, **args)

        slide_number = args.get("slide_number")
        if slide_number is not None:
            if name in _TEXT_CLAMP_TOOLS:
                shape_id = args.get("shape_id")
                if shape_id is not None:
                    clamp_text_to_slide(self.container.prs, slide_number, shape_id)
            else:
                clamp_shapes_to_slide(self.container.prs, slide_number)
        return result
