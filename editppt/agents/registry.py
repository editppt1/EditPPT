from dataclasses import dataclass, field
from typing import Callable, List

from editppt.prompts import (
    create_text_style_agent_system_prompt,
    create_table_agent_system_prompt,
    create_chart_agent_system_prompt,
    create_shape_layout_agent_system_prompt,
    create_slide_agent_system_prompt,
)


@dataclass
class AgentSpec:
    agent_type: str
    tool_names: List[str]
    system_prompt_builder: Callable[[], str]
    description: str


AGENT_REGISTRY: dict[str, AgentSpec] = {}


def register_agent(spec: AgentSpec):
    AGENT_REGISTRY[spec.agent_type] = spec


def get_agent_spec(agent_type: str) -> AgentSpec:
    if agent_type not in AGENT_REGISTRY:
        raise KeyError(f"Unknown agent_type: {agent_type}. Available: {list(AGENT_REGISTRY.keys())}")
    return AGENT_REGISTRY[agent_type]


def get_all_agent_descriptions() -> str:
    lines = []
    for agent_type, spec in AGENT_REGISTRY.items():
        tools_str = ", ".join(spec.tool_names)
        lines.append(f"- **{agent_type}**: {spec.description} (tools: {tools_str})")
    return "\n".join(lines)


# ── Register all 5 specialists ──

register_agent(AgentSpec(
    agent_type="text_style",
    tool_names=[
        "set_text_style",
        "edit_text_insert",
        "edit_text_delete",
        "edit_text_replace",
        "edit_text_rewrite",
        "set_paragraph_alignment",
        "manage_bullet_points",
    ],
    system_prompt_builder=create_text_style_agent_system_prompt,
    description="Text content and formatting: font styles, insert/delete/replace/rewrite text, paragraph alignment, bullet points",
))

register_agent(AgentSpec(
    agent_type="table",
    tool_names=[
        "cell_text_style",
        "replace_table_text",
        "table_layout_style",
    ],
    system_prompt_builder=create_table_agent_system_prompt,
    description="Table operations: cell text styling, cell content replacement, table layout and borders",
))

register_agent(AgentSpec(
    agent_type="chart",
    tool_names=[
        "update_chart_categories",
        "update_chart_series",
        "update_chart_structure",
        "update_chart_axes",
        "update_chart_colors",
    ],
    system_prompt_builder=create_chart_agent_system_prompt,
    description="Chart modifications: data, series, categories, axes, colors, and chart structure",
))

register_agent(AgentSpec(
    agent_type="shape_layout",
    tool_names=[
        "adjust_layout",
        "distribute_shapes",
        "align_shapes",
        "create_textbox",
        "create_placeholder",
        "create_shape",
        "delete_shape",
        "duplicate_shape",
        "duplicate_shape_within_slide",
        "apply_visual_style",
        "apply_gradient_fill",
        "insert_image",
        "edit_image",
    ],
    system_prompt_builder=create_shape_layout_agent_system_prompt,
    description="Shape positioning, sizing, alignment, creation, deletion, duplication, visual effects, image insertion and editing",
))

register_agent(AgentSpec(
    agent_type="slide",
    tool_names=[
        "add_slide",
        "delete_slide",
        "duplicate_slide",
        "set_slide_transition",
        "set_slide_background",
    ],
    system_prompt_builder=create_slide_agent_system_prompt,
    description="Slide-level operations: add, delete, duplicate slides, set slide transitions, and set slide backgrounds",
))
