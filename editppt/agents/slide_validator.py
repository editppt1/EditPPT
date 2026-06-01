"""
Rule-based validation for slide-level tool operations.

Slide-level edits (transition, background, add/delete/duplicate) don't surface
in the parsed slide JSON, so the text/vision validators can't catch failures
on these tools. These checks compare COM state directly against the args the
tool was called with.
"""
from __future__ import annotations
from typing import Tuple


# Mirror of transition_map in tools.set_slide_transition
TRANSITION_VALUES = {
    "none": 0,
    "fade": 1793,
    "fade_smoothly": 3849,
    "push": 3852,
    "wipe": 2819,
    "split": 3585,
    "cover": 1284,
    "uncover": 2052,
    "dissolve": 1537,
    "cut": 257,
    "random_bars": 2305,
    "checkerboard": 1025,
    "morph": 3854,
}

_ADD_LIKE = {"add_slide", "duplicate_slide"}
_DEL_LIKE = {"delete_slide"}


def capture_pre_state(prs) -> dict:
    """Snapshot state needed by post-tool rule checks. Cheap; safe to always call."""
    return {"count": prs.Slides.Count}


def validate_slide_ops(prs, tool_calls: list, pre_state: dict) -> Tuple[bool, str]:
    """
    Rule-check a batch of slide-level tool calls against the post-tool COM
    state. Returns (valid, reason). First failing rule short-circuits.
    """
    expected_delta = sum(
        1 if tc["name"] in _ADD_LIKE else (-1 if tc["name"] in _DEL_LIKE else 0)
        for tc in tool_calls
    )
    actual_delta = prs.Slides.Count - pre_state["count"]
    if actual_delta != expected_delta:
        return False, (
            f"slide count delta mismatch: expected {expected_delta:+d}, "
            f"got {actual_delta:+d} ({pre_state['count']} -> {prs.Slides.Count})"
        )

    for tc in tool_calls:
        name = tc["name"]
        args = tc.get("arguments", {}) or {}

        if name == "set_slide_transition":
            ok, reason = _check_transition(prs, args)
        elif name == "set_slide_background":
            ok, reason = _check_background(prs, args)
        elif name == "duplicate_slide":
            ok, reason = _check_duplicate(prs, args)
        elif name == "add_slide":
            ok, reason = _check_add_slide(prs, args)
        elif name == "delete_slide":
            # Count delta already covers this.
            ok, reason = True, "ok"
        else:
            ok, reason = True, "ok"

        if not ok:
            return False, f"{name}: {reason}"

    return True, "ok"


def _check_transition(prs, args) -> Tuple[bool, str]:
    slide_number = args.get("slide_number")
    ttype = args.get("transition_type", "fade")
    expected = TRANSITION_VALUES.get(ttype)
    if expected is None:
        return False, f"unknown transition_type '{ttype}'"
    try:
        slide = prs.Slides(slide_number)
        actual = slide.SlideShowTransition.EntryEffect
    except Exception as e:
        return False, f"failed to read EntryEffect on slide {slide_number}: {e}"
    if actual != expected:
        return False, (
            f"slide {slide_number} EntryEffect: expected {expected} ({ttype}), "
            f"got {actual}"
        )
    return True, "ok"


def _check_background(prs, args) -> Tuple[bool, str]:
    slide_number = args.get("slide_number")
    fill_type = args.get("fill_type", "solid")
    color_hex = args.get("color_hex", "#FFFFFF")

    try:
        slide = prs.Slides(slide_number)
    except Exception as e:
        return False, f"slide {slide_number} not accessible: {e}"

    if fill_type == "none":
        try:
            follows = bool(slide.FollowMasterBackground)
        except Exception as e:
            return False, f"failed to read FollowMasterBackground: {e}"
        if not follows:
            return False, "fill_type='none' but FollowMasterBackground is False"
        return True, "ok"

    # solid / gradient: master should not be followed, ForeColor should match
    try:
        if bool(slide.FollowMasterBackground):
            return False, f"background still follows master after fill_type='{fill_type}'"
    except Exception:
        # Some COM dispatch versions raise on this property; non-fatal.
        pass

    try:
        actual_rgb = slide.Background.Fill.ForeColor.RGB
    except Exception as e:
        return False, f"failed to read Background.Fill.ForeColor.RGB: {e}"

    expected_rgb = _hex_to_bgr_int(color_hex)
    if actual_rgb != expected_rgb:
        return False, (
            f"slide {slide_number} ForeColor: expected 0x{expected_rgb:06X} "
            f"({color_hex}), got 0x{actual_rgb:06X}"
        )
    return True, "ok"


def _check_duplicate(prs, args) -> Tuple[bool, str]:
    slide_number = args.get("slide_number")
    try:
        total = prs.Slides.Count
    except Exception as e:
        return False, f"failed to read Slides.Count: {e}"

    # The duplicate lands directly after the source.
    if slide_number is None or slide_number < 1 or slide_number + 1 > total:
        return False, (
            f"duplicate target out of range (source={slide_number}, total={total})"
        )
    try:
        src_count = prs.Slides(slide_number).Shapes.Count
        dup_count = prs.Slides(slide_number + 1).Shapes.Count
    except Exception as e:
        return False, f"failed to read Shapes.Count: {e}"
    if src_count != dup_count:
        return False, f"shape count mismatch: source={src_count}, duplicate={dup_count}"
    return True, "ok"


def _check_add_slide(prs, args) -> Tuple[bool, str]:
    position = args.get("position")
    try:
        total = prs.Slides.Count
    except Exception as e:
        return False, f"failed to read Slides.Count: {e}"
    expected_pos = position if position else total
    expected_pos = max(1, min(expected_pos, total))
    try:
        prs.Slides(expected_pos)
    except Exception as e:
        return False, f"new slide not present at position {expected_pos}: {e}"
    return True, "ok"


def _hex_to_bgr_int(hex_str: str) -> int:
    """Match tools._hex_to_rgb_int (BGR ordering used by win32 COM)."""
    hex_str = hex_str.lstrip("#")
    r = int(hex_str[0:2], 16)
    g = int(hex_str[2:4], 16)
    b = int(hex_str[4:6], 16)
    return (b << 16) | (g << 8) | r
