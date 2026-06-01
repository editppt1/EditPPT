from dataclasses import dataclass
from typing import Callable, List
import json

from editppt.utils.llm_client import call_llm, call_llm_gemini, is_anthropic_model, set_token_log_context
from editppt.prompts import *
from editppt.utils.logger_manual import *
from editppt.utils.msoffice_map import *
from editppt.utils.utils import parse_active_slide_objects
from pathlib import Path


@dataclass
class ToolMeta:
    needs_slide_json: bool = False
    needs_agent_request: bool = False
    needs_container: bool = False
    cleanup_false_args: List[str] | None = None
    # When True, post-tool validation routes to editppt.agents.slide_validator
    # (rule-based COM checks) instead of text/vision validation.
    slide_rule_validation: bool = False




TOOL_METADATA = {
    "set_text_style": ToolMeta(
        needs_slide_json=True,
        cleanup_false_args=["bold", "italic", "underline", "font_name", "font_size"],
    ),
    # "edit_text_replace": ToolMeta(
    #     needs_slide_json=True,
    #     needs_container=True,
    #     needs_agent_request=True,
    # ),
    "cell_text_style": ToolMeta(
        needs_slide_json=True,
    ),
    "replace_table_text": ToolMeta(
        needs_slide_json=True,
        needs_agent_request=True,
    ),
    "edit_text_rewrite": ToolMeta(
        needs_slide_json=True,
        needs_agent_request=True,
    ),
    "create_textbox": ToolMeta(
        needs_slide_json=True,
    ),
    "create_placeholder": ToolMeta(
        needs_slide_json=True,
    ),
    "insert_image": ToolMeta(),
    "edit_image": ToolMeta(),
    # Slide-level ops: changes don't appear in parsed slide JSON. Routed to
    # rule-based validation (slide_validator.validate_slide_ops) instead of
    # text/vision validation.
    "add_slide": ToolMeta(slide_rule_validation=True),
    "delete_slide": ToolMeta(slide_rule_validation=True),
    "duplicate_slide": ToolMeta(slide_rule_validation=True),
    "set_slide_transition": ToolMeta(slide_rule_validation=True),
    "set_slide_background": ToolMeta(slide_rule_validation=True),
}


def _additional_call_create_shape(function_name, function_args, agent_request, model, tools):
    """
    Re-invoke LLM to enrich an incomplete create_shape tool call,
    especially to fill missing or ambiguous shape number/type.
    """

    # --- Shape guide as JSON string ---
    shape_guide_json = json.dumps(
        AUTOSHAPE_TYPE_MAP, ensure_ascii=False, indent=2
    )

    # --- System prompt ---
    system_prompt = f"""
The current tool call is incomplete and cannot be executed as-is.

User request (raw):
{json.dumps(agent_request, ensure_ascii=False, indent=2)}

Original tool call:
Function name: {function_name}
Arguments:
{json.dumps(function_args, ensure_ascii=False, indent=2)}

Your task:
- Analyze the user request and the original tool arguments
- Identify missing or ambiguous shape number / type
- Generate a corrected and executable tool call
"""

    # --- User prompt ---
    user_prompt = f"""
Available shape types (JSON reference):
{shape_guide_json}

Based on the information above,
generate a corrected tool call with a valid shape number.
"""

    messages = [
        {"role": "system", "content": system_prompt},
        {"role": "user", "content": user_prompt},
    ]

    # --- Call LLM ---
    tc = {"type": "any"} if is_anthropic_model(model) else "auto"
    set_token_log_context(component="shape_resolver")
    response = call_llm(
        model=model,
        messages=messages,
        tools=tools,
        tool_choice=tc,
    )

    # --- Parse tool calls ---
    tool_calls = []
    for item in response.output:
        if item.type == "function_call":
            tool_calls.append({
                "name": item.name,
                "arguments": (
                    json.loads(item.arguments)
                    if isinstance(item.arguments, str)
                    else item.arguments
                ),
                "call_id": item.call_id,
            })

    if not tool_calls:
        raise RuntimeError(
            "additional_call_create_shape: no tool call was generated"
        )

    return tool_calls[0]

ADDITIONAL_CALL_TOOL_DICT = {
    "create_shape":_additional_call_create_shape
}