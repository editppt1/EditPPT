import json
import re
import time
import os

from typing import Any, List, Optional
from editppt.utils.utils import codepoint_to_utf16

import logging
logger = logging.getLogger(__name__)

# --- Internal Helper Functions ---

def _hex_to_rgb_int(hex_str):
    """Converts HEX string to win32-compatible BGR integer.

    Accepts the variants the LLM actually produces:
      - '#FFFFFF' / 'FFFFFF'              (standard 6-char)
      - '#FFF'    / 'FFF'                 (CSS shorthand; expanded by doubling)
      - '#FFFFFFFF' / 'FFFFFFFF'          (8-char with alpha; alpha discarded)
    Anything else is rejected.
    """
    h = hex_str.lstrip('#').strip()
    if len(h) == 3:
        h = "".join(c * 2 for c in h)
    elif len(h) == 8:
        h = h[:6]
    if len(h) != 6:
        raise ValueError(f"HEX code must be 3, 6, or 8 hex digits (got '{hex_str}').")
    try:
        r = int(h[0:2], 16)
        g = int(h[2:4], 16)
        b = int(h[4:6], 16)
    except ValueError as e:
        raise ValueError(f"HEX code contains non-hex characters: '{hex_str}'") from e
    return (b << 16) | (g << 8) | r


def clamp_shapes_to_slide(prs, slide_number):
    """
    Ensure all shapes on the given slide stay within the slide boundaries.
    Shapes that overflow are repositioned/resized to fit.
    For layout/shape tools — preserves shape size, adjusts position.
    """
    try:
        slide = prs.Slides(slide_number)
        slide_w = prs.PageSetup.SlideWidth
        slide_h = prs.PageSetup.SlideHeight
    except Exception:
        return

    for shape in slide.Shapes:
        try:
            left = shape.Left
            top = shape.Top
            width = shape.Width
            height = shape.Height

            # Clamp width/height to slide dimensions
            if width > slide_w:
                shape.Width = slide_w
                width = slide_w
            if height > slide_h:
                shape.Height = slide_h
                height = slide_h

            # Clamp position so shape stays within bounds
            if left < 0:
                shape.Left = 0
            elif left + width > slide_w:
                shape.Left = slide_w - width

            if top < 0:
                shape.Top = 0
            elif top + height > slide_h:
                shape.Top = slide_h - height
        except Exception:
            continue


def clamp_text_to_slide(prs, slide_number, shape_id):
    """
    For text tools — preserve position, shrink font size until the shape
    fits within the slide boundaries.  Uses adaptive scaling with run-level
    granularity for speed (runs << characters in COM call count).
    """
    try:
        slide_w = prs.PageSetup.SlideWidth
        slide_h = prs.PageSetup.SlideHeight
        shape = _find_shape_by_id(prs, slide_number, shape_id)
    except Exception:
        return

    if not shape.HasTextFrame:
        return

    tr = shape.TextFrame.TextRange
    if not tr.Text:
        return

    def _overflow():
        """Max overflow in points (0 = fits within slide)."""
        over_r = (shape.Left + shape.Width) - slide_w
        over_b = (shape.Top + shape.Height) - slide_h
        return max(0, over_r, over_b)

    if _overflow() <= 1:
        return

    # --- Collect run-level font sizes (much fewer COM calls than per-char) ---
    run_info = []  # [(COM range object, original_size)]
    for pi in range(1, tr.Paragraphs().Count + 1):
        para = tr.Paragraphs(pi)
        rc = para.Runs().Count
        if rc == 0:
            # Empty paragraph or no distinct runs — use paragraph range
            try:
                size = para.Font.Size
                if not size or size <= 0:
                    size = 12.0
            except Exception:
                size = 12.0
            run_info.append((para, size))
        else:
            for ri in range(1, rc + 1):
                run = para.Runs(ri)
                try:
                    size = run.Font.Size
                    if not size or size <= 0:
                        size = 12.0
                except Exception:
                    size = 12.0
                run_info.append((run, size))

    if not run_info:
        return

    max_orig = max(s for _, s in run_info)
    min_allowed = max(6.0, max_orig * 0.5)

    def _apply_scale(scale_factor):
        """Apply proportional font-size reduction to all runs."""
        for run_obj, orig in run_info:
            new_s = max(min_allowed, round(orig * scale_factor * 2) / 2)  # snap to 0.5pt
            try:
                run_obj.Font.Size = new_s
            except Exception:
                pass

    # --- Adaptive: estimate initial scale from overflow ratio ---
    shape_dim = max(shape.Width, shape.Height, 1)
    scale = max(0.5, 1.0 - _overflow() / shape_dim)
    _apply_scale(scale)

    # Refine (max 4 extra passes) — re-estimate from remaining overflow
    for _ in range(4):
        ovf = _overflow()
        if ovf <= 1:
            break
        scale *= max(0.7, 1.0 - ovf / shape_dim)
        scale = max(0.5, scale)
        _apply_scale(scale)


def _find_shape_by_id(prs, slide_number, shape_id):
    """Finds a specific Shape object by its unique ID on a given slide.

    Parser flattens group members into Objects_Detail, so the LLM can target
    shapes nested inside groups. The lookup must therefore recurse into
    GroupItems too — otherwise group-internal shape_ids raise 'not found'
    even though the parser advertised them.
    """
    try:
        slide = prs.Slides(slide_number)
    except Exception as e:
        raise ValueError(f"Error accessing slide {slide_number}: {e}")
    found = _find_shape_recursive(slide.Shapes, shape_id)
    if found is not None:
        return found
    raise ValueError(f"Shape with ID {shape_id} not found on slide {slide_number}.")

def _find_shape_recursive(shapes, shape_id):
    """
    Traverses the win32com shape collection, searching for shape_id including inside groups.
    """
    for shape in shapes:
        # 1. Check if the current shape's ID matches
        if shape.Id == shape_id:
            return shape
        
        # 2. If this shape is a group (msoGroup = 6)
        # In win32com, shape.Type == 6 means a group shape.
        if shape.Type == 6: 
            # Search again within the group's internal items (GroupItems)
            found = _find_shape_recursive(shape.GroupItems, shape_id)
            if found:
                return found
    return None


ALIGN_MAP = {1: "left",2: "center",3: "right",4: "justify",5: "distributed"}
ALIGN_MAP_REV = {v: k for k, v in ALIGN_MAP.items()}


#########################################################################
######################## [A]  Text Style Editing ########################
#########################################################################
def _get_text_with_offsets(
    prs,
    slide_number: int,
    shape_id: int,
    *,
    container: str = "shape",
    row_index: int = None,
    col_index: int = None,
):
    shape = _find_shape_by_id(prs, slide_number, shape_id)

    if container == "shape":
        if not shape.HasTextFrame or not shape.TextFrame.HasText:
            return "", []
        tr = shape.TextFrame.TextRange

    elif container == "table_cell":
        if row_index is None or col_index is None:
            raise ValueError("row_index / col_index required.")
        cell = shape.Table.Cell(row_index, col_index)
        if not cell.Shape.TextFrame.HasText:
            return "", []
        tr = cell.Shape.TextFrame.TextRange

    else:
        raise ValueError(f"Unknown container: {container}")

    text = tr.Text or ""
    return text, list(range(len(text)))


def _normalize_char_range(
    text: str,
    char_start_index: int,
    target_text: str,
    char_end: int = None,
):
    """Resolve (start, end) of `target_text` inside `text`.

    The LLM's `char_start_index` is treated as a hint, not gospel — UTF-16 vs
    Python str index drift and small off-by-N errors are the dominant failure
    mode in production. So we accept the hint when it's exact, otherwise we
    locate every occurrence of `target_text` and pick the one whose start is
    closest to the hint. If `target_text` is absent entirely, raise.
    """
    expected_len = len(target_text)
    if not target_text:
        raise ValueError("target_text is empty.")

    # 1. Exact match at the LLM-supplied index.
    if 0 <= char_start_index <= len(text) - expected_len:
        if text[char_start_index:char_start_index + expected_len] == target_text:
            return char_start_index, char_start_index + expected_len

    # 2. Honour an explicit (start, end) if the LLM provided both and they match.
    if char_end is not None and 0 <= char_start_index < char_end <= len(text):
        if text[char_start_index:char_end] == target_text:
            return char_start_index, char_end

    # 3. Locate every occurrence; pick the one whose start is closest to the hint.
    occurrences = []
    pos = 0
    while True:
        idx = text.find(target_text, pos)
        if idx == -1:
            break
        occurrences.append(idx)
        pos = idx + 1
    if not occurrences:
        raise ValueError("Unable to resolve exact character range.")

    hint = max(0, min(char_start_index, len(text)))
    best = min(occurrences, key=lambda i: abs(i - hint))
    return best, best + expected_len

def _get_detail_from_json(slide_json: dict, shape_id: int, keys: list):
    objects = slide_json.get("Objects_Detail", [])
    for obj in objects:
        if obj.get("Shape_Id") == shape_id:
            current = obj
            try:
                for key in keys:
                    current = current[key]
                return current
            except (KeyError, TypeError):
                return {}
                
    for obj in objects:
        if obj.get("Type") == "Group":
            found = _find_json_group_recursive(obj.get("More_detail", {}).get("Group", []), shape_id, keys)
            if found is not None:
                return found

    raise ValueError(f"Shape_Id {shape_id} not found in slide JSON.")


def _find_json_group_recursive(items, shape_id, keys):
    for item in items:
        if item.get("Shape_Id") == shape_id:
            current = item
            try:
                for key in keys:
                    if key == "More_detail" and "More_detail" not in current:
                        continue
                    current = current[key]
                return current
            except (KeyError, TypeError):
                return {}

        # If it's a group again, recurse
        if item.get("Type") == "Group":
            children = item.get("More_detail", {}).get("Group", [])
            found = _find_json_group_recursive(children, shape_id, keys) # fix missing keys arg
            if found is not None:
                return found
    return None


def _iter_run_slices_from_shape_json(slide_json, shape_id, start, end):
    runs = _get_detail_from_json(
        slide_json,
        shape_id,
        ["More_detail", "TextFrame", "Runs"]
    )
    yield from _iter_run_slices_from_runs(runs, start, end)

def _iter_run_slices_from_runs(runs, start, end):
    if not runs:
        return
    for run in runs:
        if "Run_Start_Index" not in run:
            rs = 0
        else:
            rs = run["Run_Start_Index"]
        run_end = rs + len(run["Text"])

        s = max(start, rs)
        e = min(end, run_end)

        if s < e:
            yield s, e, run
            
            

def _apply_font_snapshot(font, snap: dict):
    if "Name" in snap: font.Name = snap["Name"]
    if "Size" in snap: font.Size = snap["Size"]

    font.Bold = int(snap.get("Bold", False))
    font.Italic = int(snap.get("Italic", False))
    font.Underline = int(snap.get("Underline", False))

    if "Strikethrough" in snap:
        try: font.Strikethrough = int(snap["Strikethrough"])
        except Exception: pass
    if "Subscript" in snap:
        try: font.Subscript = int(snap["Subscript"])
        except Exception: pass
    if "Superscript" in snap:
        try: font.Superscript = int(snap["Superscript"])
        except Exception: pass

    color = snap.get("Color")
    if color:
        font.Color.RGB = _hex_to_rgb_int(color)

def _apply_overrides(font, *,
                    font_name=None,
                    font_size=None,
                    bold=None,
                    italic=None,
                    underline=None,
                    color_hex=None):
    if font_name is not None:
        font.Name = font_name
    if font_size is not None:
        font.Size = font_size
    if bold is not None:
        font.Bold = int(bold)
    if italic is not None:
        font.Italic = int(italic)
    if underline is not None:
        font.Underline = int(underline)
    if color_hex is not None:
        font.Color.RGB = _hex_to_rgb_int(color_hex)

def set_text_style(
    prs, slide_number: int, shape_id: int, slide_json: dict,
    *, scope: str = "range",
    char_start_index: int = None, target_text: str = None, char_end: int = None,
    container: str = "shape", row_index: int = None, col_index: int = None,
    font_name: str = None, font_size=None, bold=None, italic=None, underline=None, color_hex=None,
):
    # 1. Extract text and normalize range
    text, _ = _get_text_with_offsets(
        prs, slide_number, shape_id,
        container=container, row_index=row_index, col_index=col_index
    )

    if scope == "all":
        start, end = 0, len(text)
    elif scope == "range":
        if char_start_index is None or target_text is None:
            raise ValueError("scope='range' requires char_start_index and target_text.")
        start, end = _normalize_char_range(
            text=text, char_start_index=char_start_index, target_text=target_text, char_end=char_end
        )
    else:
        raise ValueError(f"Unknown scope: {scope!r} (expected 'range' or 'all').")

    if start >= end:
        return {
            "operation": "set_text_style", "scope": scope, "applied_range": [start, end],
            "shape_id": shape_id, "slide": slide_number, "note": "empty range; nothing to do",
        }

    # 2. Obtain TextRange
    shape = _find_shape_by_id(prs, slide_number, shape_id)
    tr = (
        shape.TextFrame.TextRange if container == "shape"
        else shape.Table.Cell(row_index, col_index).Shape.TextFrame.TextRange
    )

    # 3. Determine JSON data source (key fix)
    if container == "table_cell":
        target_json_detail = _get_detail_from_json(
            slide_json,
            shape_id,
            ["More_detail", "Table", "Cells", f"{row_index},{col_index}"]
        )
    else:
        target_json_detail = _get_detail_from_json(
            slide_json,
            shape_id,
            ["More_detail", "TextFrame"]
        )
    # 4. Iterate run-level slices (using corrected JSON data)
    # Modify _iter_run_slices or directly use the Runs from the target area
    runs_data = target_json_detail.get("Runs", [])
    
    for s, e, run in _iter_run_slices_from_runs(runs_data, start, end):
        length = e - s
        if length <= 0:
            continue

        # Convert codepoint offsets to UTF-16 code units for COM
        com_start = codepoint_to_utf16(text, s) + 1   # 1-based
        com_len = codepoint_to_utf16(text, e) - (com_start - 1)
        target = tr.Characters(com_start, com_len)
        font = target.Font

        _apply_font_snapshot(font, run["Font"])
        _apply_overrides(
            font,
            font_name=font_name,
            font_size=font_size,
            bold=bold,
            italic=italic,
            underline=underline,
            color_hex=color_hex
        )

    return {
        "operation": "set_text_style",
        "container": container,
        "applied_range": [start, end],
        "shape_id": shape_id,
        "slide": slide_number,
    }

def _resolve_insert_position(text: str, preceding_text: str, char_start_index: int) -> int:
    """
    Prioritizes preceding_text, but corrects by finding the position closest to char_start_index.
    """
    # 1. Handle text start marker ([SOS])
    if preceding_text == "[SOS]":
        return 0

    # 2. If preceding_text is provided, find all occurrence positions in the text
    if preceding_text:
        # Use re.finditer to find start indices of all (possibly overlapping) occurrences
        all_matches = [m.start() for m in re.finditer(re.escape(preceding_text), text)]
        
        if all_matches:
            # Select the match whose end position (start + len) is closest to the agent's char_start_index
            # (the insertion/deletion point is right after preceding_text)
            best_match_start = min(all_matches, key=lambda x: abs((x + len(preceding_text)) - char_start_index))
            return best_match_start + len(preceding_text)

    # 3. Fallback: if the word was not found or not provided, use clamped index
    return max(0, min(len(text), char_start_index))


def edit_text_insert(
    prs,
    slide_number,
    shape_id,
    preceding_text,
    char_start_index,
    new_text,
    *,
    container="shape",
    row_index=None,
    col_index=None,
    auto_resize: bool = False,
):
    text, _ = _get_text_with_offsets(
        prs, slide_number, shape_id,
        container=container,
        row_index=row_index,
        col_index=col_index,
    )

    # Calculate corrected insertion position
    insert_pos = _resolve_insert_position(
        text=text,
        preceding_text=preceding_text,
        char_start_index=char_start_index,
    )

    shape = _find_shape_by_id(prs, slide_number, shape_id)
    tf = shape.TextFrame if container == "shape" else shape.Table.Cell(row_index, col_index).Shape.TextFrame
    tr = tf.TextRange

    if auto_resize:
        original_height = shape.Height if container == "shape" else shape.Table.Cell(row_index, col_index).Shape.Height

    # PowerPoint TextRange.Characters uses 1-based UTF-16 index
    com_pos = codepoint_to_utf16(text, insert_pos) + 1
    anchor = tr.Characters(com_pos, 0)
    anchor.InsertAfter(new_text)

    # Height-based font shrink when auto_resize is enabled
    if auto_resize:
        new_tr = tf.TextRange
        current_size = new_tr.Font.Size if new_tr.Font.Size else 12.0
        while new_tr.BoundHeight > original_height and current_size > 9.0:
            current_size -= 0.5
            new_tr.Font.Size = current_size

    return {
        "operation": "insert",
        "resolved_insert_pos": insert_pos,
        "new_text": new_text,
        "shape_id": shape_id,
        "slide": slide_number,
    }


def edit_text_delete(
    prs,
    slide_number,
    shape_id,
    preceding_text,   # added
    target_text,
    char_start_index,
    *,
    container="shape",
    row_index=None,
    col_index=None,
):
    text, _ = _get_text_with_offsets(
        prs, slide_number, shape_id,
        container=container,
        row_index=row_index,
        col_index=col_index,
    )

    # 1. Correct the deletion start reference point based on preceding_text
    refined_start = _resolve_insert_position(
        text=text,
        preceding_text=preceding_text,
        char_start_index=char_start_index,
    )

    # 2. Search for the actual start of target_text near the corrected position (tolerance +-5 chars)
    actual_start = text.find(target_text, max(0, refined_start - 5))
    
    if actual_start != -1:
        # If found, set the deletion range from that position for the length of target_text
        start, end = actual_start, actual_start + len(target_text)
    else:
        # If not found, fall back to the forced matching logic
        start, end = _normalize_char_range(
            text=text,
            char_start_index=char_start_index,
            target_text=target_text,
        )

    shape = _find_shape_by_id(prs, slide_number, shape_id)
    tr = (
        shape.TextFrame.TextRange
        if container == "shape"
        else shape.Table.Cell(row_index, col_index).Shape.TextFrame.TextRange
    )

    # Delete via Characters(start, length) — codepoint to UTF-16 conversion
    com_start = codepoint_to_utf16(text, start) + 1
    com_len = codepoint_to_utf16(text, end) - (com_start - 1)
    tr.Characters(com_start, com_len).Delete()

    return {
        "operation": "delete",
        "range": [start, end],
        "deleted_text": target_text,
        "shape_id": shape_id,
        "slide": slide_number,
    }

def edit_text_replace(
    prs,
    slide_number,
    shape_id,
    preceding_text,
    target_text,
    char_start_index,
    new_text,
    auto_resize: bool = False,
):
    """
    Replace target_text with new_text inside a shape.
    Uses preceding_text + char_start_index heuristic for robust positioning.
    """

    # 1. Get full text (shape only)
    text, _ = _get_text_with_offsets(
        prs, slide_number, shape_id,
        container="shape"
    )

    # 2. Correct reference point based on preceding_text
    refined_start = _resolve_insert_position(
        text=text,
        preceding_text=preceding_text,
        char_start_index=char_start_index,
    )

    # 3. Search for target_text near the corrected position (tolerance +-5 chars)
    search_start = max(0, refined_start - 5)
    actual_start = text.find(target_text, search_start)

    if actual_start != -1:
        start = actual_start
        end = actual_start + len(target_text)
    else:
        # fallback
        start, end = _normalize_char_range(
            text=text,
            char_start_index=char_start_index,
            target_text=target_text,
        )

    shape = _find_shape_by_id(prs, slide_number, shape_id)
    tf = shape.TextFrame
    tr = tf.TextRange

    if auto_resize:
        original_height = shape.Height

    # 4. Delete existing text — codepoint to UTF-16 conversion
    com_start = codepoint_to_utf16(text, start) + 1
    com_len = codepoint_to_utf16(text, end) - (com_start - 1)
    tr.Characters(com_start, com_len).Delete()

    # 5. Insert new text at the same position
    anchor = tr.Characters(com_start, 0)
    anchor.InsertAfter(new_text)

    # 6. Height-based font shrink when auto_resize is enabled
    if auto_resize:
        new_tr = tf.TextRange
        current_size = new_tr.Font.Size if new_tr.Font.Size else 12.0
        while new_tr.BoundHeight > original_height and current_size > 9.0:
            current_size -= 0.5
            new_tr.Font.Size = current_size

    return {
        "operation": "replace",
        "range": [start, end],
        "old_text": target_text,
        "new_text": new_text,
        "shape_id": shape_id,
        "slide": slide_number,
    }


#This tool uses Additional LLM
from editppt.utils.llm_client import call_llm, is_anthropic_model, set_token_log_context
from editppt.utils.logger_manual import log_path
from editppt.utils.utils import parse_llm_response, build_paragraph_ir_from_textframe
from editppt.prompts import FLATTEXT_STYLE_MAPPING_PROMPT, PARAGRAPH_STYLE_MAPPING_PROMPT
from editppt.utils.msoffice_map import BULLET_CHAR_MAP, BULLET_STYLE_MAP
from editppt.config import STYLE_MAPPER_MODEL


def _reconcile_paragraph_mapping(parsed, paragraph_ir, new_text):
    """Force `parsed` to have exactly len(paragraph_ir) entries.

    The style_mapper LLM occasionally returns a paragraph count that doesn't
    match the original (most often collapsing many paragraphs into one),
    which leaves later paragraphs untouched or merges content visually.
    Vision validator catches the resulting overflow/merge frequently.

    Fallback: split new_text on \\r, pad/merge to expected count, and
    synthesize one default-font run per paragraph using paragraph_ir's
    primary font. Returns the (possibly rebuilt) parsed list.
    """
    if not isinstance(parsed, list):
        parsed = []
    expected = len(paragraph_ir)
    if expected == 0:
        return parsed
    if len(parsed) == expected:
        return parsed

    logger.warning(
        f"style_mapper paragraph count mismatch: got {len(parsed)}, expected "
        f"{expected}; rebuilding from new_text split."
    )
    new_paragraphs_text = new_text.split('\r')
    # Don't strip trailing empties — original may include intentional blank paragraphs
    if len(new_paragraphs_text) < expected:
        new_paragraphs_text += [""] * (expected - len(new_paragraphs_text))
    elif len(new_paragraphs_text) > expected:
        # Merge the overflow into the last expected slot (preserves all words)
        head = new_paragraphs_text[:expected - 1]
        tail = " ".join(new_paragraphs_text[expected - 1:])
        new_paragraphs_text = head + [tail]

    rebuilt = []
    for i, ptext in enumerate(new_paragraphs_text):
        default_font = {}
        runs_meta = paragraph_ir[i].get("runs", []) if i < len(paragraph_ir) else []
        if runs_meta:
            # First run's font is usually the paragraph's base style
            default_font = runs_meta[0].get("font", {}) or {}
        rebuilt.append({
            "para_id": i + 1,
            "runs": [{"Text": ptext, "Font": default_font}],
        })
    return rebuilt


def _strip_trailing_empty_paragraphs(tf):
    """Delete trailing empty paragraphs (trailing \\r) from a TextFrame."""
    tr = tf.TextRange
    while tr.Length > 0:
        last_char = tr.Characters(tr.Length, 1)
        if last_char.Text == "\r":
            last_char.Delete()
            tr = tf.TextRange  # refresh after mutation
        else:
            break

def edit_text_rewrite(
    prs,
    slide_number: int,
    shape_id: int,
    new_text: str,
    slide_json: dict,
    agent_request,
    auto_resize: bool = False,
):
    """
    Fully rewrite the text content of a shape while preserving the original
    paragraph structure and run-level styles via LLM style mapping.

    Use this ONLY for tasks that require the entire text to be replaced:
    translation, summarization, full restructuring, etc.
    """
    # 0. Normalize text
    if isinstance(new_text, str):
        new_text = new_text.replace("\n", "\r")

    # 1. Load shape detail from JSON
    text_frame_detail = _get_detail_from_json(
        slide_json,
        shape_id,
        ["More_detail", "TextFrame"]
    )
    old_runs = text_frame_detail.get("Runs", []) or []
    old_paragraphs = text_frame_detail.get("Paragraphs")  # may be None

    # 2. Resolve shape and save geometry state
    shape = _find_shape_by_id(prs, slide_number, shape_id)
    if not shape.HasTextFrame:
        raise ValueError(f"Shape {shape_id} does not have a text frame.")

    tf = shape.TextFrame
    tr = tf.TextRange
    # AutoSize: 0=None, 1=ShapeToFitText, 2=TextToFitShape, -2=Mixed
    # Values 1 and 2 mean "auto-adjusting"; anything else means fixed size
    original_auto_size = tf.AutoSize
    is_auto_size = original_auto_size in (1, 2)
    original_height = shape.Height
    original_width = shape.Width

    # 3. Build paragraph IR (LLM path)
    paragraph_ir = []
    if old_paragraphs:
        paragraph_ir = build_paragraph_ir_from_textframe(
            old_runs, old_paragraphs
        )

    payload = []
    for p in paragraph_ir:
        item = {
            "para_id": p["paragraph_index"],
            "text": p["text"],
            "has_bullet": p["has_bullet"],
            "runs": p["runs"],
        }
        if p["has_bullet"]:
            item["bullet_meta"] = p.get("bullet_meta", {})
        if "Alignment" in p:
            item["Alignment"] = p["Alignment"]
        if "IndentLevel" in p:
            item["IndentLevel"] = p["IndentLevel"]
        payload.append(item)

    sizes = [
        run.get("Font", {}).get("Size")
        for run in old_runs
        if run.get("Font", {}).get("Size")
    ]
    old_base_font_size = max(sizes) if sizes else None

    # 4. LLM prompt for style mapping
    if isinstance(agent_request, (tuple, list)) and len(agent_request) >= 3:
        task_description, action_type, slide_contents = agent_request[0], agent_request[1], agent_request[2]
    elif isinstance(agent_request, str):
        task_description, action_type, slide_contents = agent_request, "", ""
    else:
        task_description, action_type, slide_contents = str(agent_request), "", ""
    is_paragraph_mode = len(paragraph_ir) > 1

    if is_paragraph_mode:
        llm_prompt = [
            {"role": "system", "content": PARAGRAPH_STYLE_MAPPING_PROMPT},
            {
                "role": "user",
                "content": json.dumps(
                    {
                        "user_request": {
                            "task_description": task_description,
                            "action_type": action_type,
                            "slide_contents": slide_contents,
                        },
                        "paragraphs": payload,
                        "new_text": new_text,
                    },
                    ensure_ascii=False,
                ),
            },
        ]
    else:
        llm_prompt = [
            {"role": "system", "content": FLATTEXT_STYLE_MAPPING_PROMPT},
            {
                "role": "user",
                "content": json.dumps(
                    {
                        "user_request": {
                            "task_description": task_description,
                            "action_type": action_type,
                            "slide_contents": slide_contents,
                        },
                        "old_runs": old_runs,
                        "new_text": new_text,
                    },
                    ensure_ascii=False,
                ),
            },
        ]

    set_token_log_context(component="style_mapper")
    raw_response = call_llm(model=STYLE_MAPPER_MODEL, messages=llm_prompt)
    if raw_response is None:
        raise RuntimeError("LLM call failed.")

    if is_anthropic_model(STYLE_MAPPER_MODEL):
        response_text = raw_response.output_text
    else:
        response_text = raw_response.output[0].content[0].text

    parsed = parse_llm_response(response_text)
    if isinstance(parsed, tuple):
        parsed = parsed[0]
    if parsed is None:
        raise ValueError(
            "Style mapper LLM produced no parseable JSON; retry will re-prompt."
        )

    # Defensive: ensure parsed paragraph count matches the original; otherwise
    # later paragraphs keep stale text and vision validator catches the merge.
    if is_paragraph_mode:
        parsed = _reconcile_paragraph_mapping(parsed, paragraph_ir, new_text)

    # 5. Prepare TextFrame for rewrite
    if not is_auto_size:
        # AutoSize OFF: fix dimensions so BoundHeight is accurate after rewrite
        tf.WordWrap = -1  # msoTrue
        shape.Width = original_width
    tf.AutoSize = 0  # temporarily disable to prevent shape resizing during writes

    # 6. Apply text (flat mode)
    if not is_paragraph_mode:
        tr.Text = ""
        current_range = tr
        for run in parsed:
            text_seg = run.get("Text", "")
            if not text_seg:
                continue
            inserted = current_range.InsertAfter(text_seg)
            _apply_font_snapshot(inserted.Font, run.get("Font", {}))
            current_range = inserted

    # 7. Apply text (paragraph mode)
    #    Replace text per-paragraph and re-apply run-level styles only.
    #    Paragraph structure (alignment, indent, spacing, bullets) is preserved
    #    because we never destroy the paragraph objects.
    else:
        para_map = {p["para_id"]: p for p in parsed}
        sorted_pids = sorted(para_map.keys())
        para_count = tr.Paragraphs().Count

        for idx, pid in enumerate(sorted_pids):
            if idx >= para_count:
                break
            para_data = para_map[pid]
            curr_para = tr.Paragraphs(idx + 1)

            # Build full text for this paragraph from mapped runs
            # Strip \r and \n to prevent COM from creating extra paragraphs
            runs_clean = []
            for run in para_data.get("runs", []):
                clean_text = run.get("Text", "").replace("\r", "").replace("\n", "")
                runs_clean.append((clean_text, run.get("Font", {})))

            full_text = "".join(t for t, _ in runs_clean)

            # Replace text within the paragraph (preserves paragraph-level formatting)
            curr_para.Text = full_text

            # Apply run-level font styles via _apply_font_snapshot.
            # NOTE: applying Font.Italic=True or Font.Name="Cambria Math" can
            # trigger PowerPoint to auto-convert plain ASCII into math italic
            # supplementary codepoints (e.g. T → 𝑇, D → 𝐷, each 2 UTF-16 units).
            # That mutation changes the paragraph's UTF-16 layout MID-LOOP, so we
            # MUST re-read curr_para.Text on every iteration to keep positions
            # correct.  codepoint count is preserved across the conversion (T and
            # 𝑇 are both 1 codepoint), so char_offset_cp remains a stable cursor.
            char_offset_cp = 0
            for text_seg, font_snap in runs_clean:
                if not text_seg:
                    continue
                cp_len = len(text_seg)
                actual_text = curr_para.Text or ""
                u16_start = codepoint_to_utf16(actual_text, char_offset_cp) + 1
                u16_end = codepoint_to_utf16(actual_text, char_offset_cp + cp_len)
                u16_len = u16_end - (u16_start - 1)
                run_range = curr_para.Characters(u16_start, u16_len)
                _apply_font_snapshot(run_range.Font, font_snap)
                char_offset_cp += cp_len

    # 7.5 Remove trailing empty paragraphs (auto_resize)
    if auto_resize:
        _strip_trailing_empty_paragraphs(tf)

    # 8. Font size scaling (only when AutoSize was OFF and auto_resize is enabled)
    if not is_auto_size and auto_resize:
        new_tr = tf.TextRange
        shape.Width = original_width  # re-enforce width before measuring

        current_size = new_tr.Font.Size if new_tr.Font.Size else 12.0
        if old_base_font_size:
            current_size = min(current_size, old_base_font_size)

        # Shrink uniformly until text fits within the original height
        while new_tr.BoundHeight > original_height and current_size > 9.0:
            current_size -= 0.5
            new_tr.Font.Size = current_size

        # Per-character proportional scaling to preserve mixed font sizes
        if new_tr.Font.Size and current_size != new_tr.Font.Size:
            scale = current_size / new_tr.Font.Size
            if new_tr.Length > 0:
                for i in range(1, new_tr.Length + 1):
                    ch = new_tr.Characters(i, 1)
                    if ch.Font.Size:
                        ch.Font.Size *= scale
    elif is_auto_size:
        # AutoSize was ON: restore original setting so shape auto-adjusts
        try:
            tf.AutoSize = original_auto_size
        except Exception:
            pass  # -2 (Mixed) can't be set back; leave at 0

    return {
        "operation": "edit_text_rewrite",
        "mode": "paragraph" if is_paragraph_mode else "flat",
        "slide": slide_number,
        "shape_id": shape_id,
    }



def set_paragraph_alignment(prs, slide_number, shape_id, alignment="left", 
                           line_spacing=None, space_before=None, space_after=None):
    """
    Adjusts paragraph-level formatting.
    
    Args:
        alignment: 'left', 'center', 'right', 'justify', 'distribute'
        line_spacing: Line spacing multiplier (e.g., 1.5)
        space_before/after: Points before/after paragraph
    """
    shape = _find_shape_by_id(prs, slide_number, shape_id)
    if not shape.HasTextFrame:
        return f"Error: Shape {shape_id} cannot contain text."
    
    tr = shape.TextFrame.TextRange
    alignment_map = {
        'left': 1, 'center': 2, 'right': 3, 'justify': 4, 'distribute': 5
    }
    
    if alignment in alignment_map:
        tr.ParagraphFormat.Alignment = alignment_map[alignment]
    if line_spacing:
        tr.ParagraphFormat.LineRuleWithin = 0  # Multiple spacing
        tr.ParagraphFormat.SpaceWithin = line_spacing
    if space_before is not None:
        tr.ParagraphFormat.SpaceBefore = space_before
    if space_after is not None:
        tr.ParagraphFormat.SpaceAfter = space_after
    
    return f"Paragraph formatting applied to Shape {shape_id}."


def manage_bullet_points(prs, slide_number, shape_id, bullet_type="bullet", 
                     bullet_char=None, start_value=1):
    """
    Adds or modifies bullet/numbering format. 
    
    Args:
        bullet_type: 'bullet', 'number', 'none'
        bullet_char: Custom bullet character
        start_value: Starting number for numbered lists
    """
    shape = _find_shape_by_id(prs, slide_number, shape_id)
    if not shape.HasTextFrame:
        return f"Error: Shape {shape_id} cannot contain text."
    
    tr = shape.TextFrame.TextRange
    
    if bullet_type == "none":
        tr.ParagraphFormat.Bullet.Visible = False
    elif bullet_type == "bullet":
        tr.ParagraphFormat.Bullet.Visible = True
        tr.ParagraphFormat.Bullet.Type = 1  # ppBulletUnnumbered
        if bullet_char:
            tr.ParagraphFormat.Bullet.Character = ord(bullet_char)
    elif bullet_type == "number":
        tr.ParagraphFormat.Bullet.Visible = True
        tr.ParagraphFormat.Bullet.Type = 2  # ppBulletNumbered
        tr.ParagraphFormat.Bullet.StartValue = start_value
    
    return f"Bullet formatting applied to Shape {shape_id}."


from editppt.utils.msoffice_map import CHART_TYPES, LEGEND_POS

REVERSE_CHART_TYPES = {v: k for k, v in CHART_TYPES.items()}


def _get_chart(prs, slide_number: int, shape_id: int):
    slide = prs.Slides(slide_number)
    shape = next((s for s in slide.Shapes if s.Id == shape_id), None)
    if not shape or not shape.HasChart:
        raise ValueError("Chart not found.")
    return shape.Chart


def _to_com_series_index(series_index, count):
    """Normalize a caller-supplied series_index to a 1-based COM index.

    External callers (LLMs) consistently use 0-based indexing — series_index=0
    means "first series". COM SeriesCollection.Item() is 1-based, so we
    translate. Returns None for out-of-range indices.
    """
    if series_index is None:
        return None
    if not isinstance(series_index, int):
        try:
            series_index = int(series_index)
        except Exception:
            return None
    if 0 <= series_index < count:
        return series_index + 1
    # Tolerate 1-based input that already lands in valid COM range
    if 1 <= series_index <= count:
        return series_index
    return None


def _set_series_literal_name(series, name):
    """Set a chart series Name to a literal string. COM Series.Name expects
    either a formula (e.g., =Sheet1!$A$1) or a quoted literal (="My Name"),
    not a raw string — raw strings silently fail to persist as the displayed
    name. This helper wraps the literal with `="..."` and falls back to a
    raw assignment if the formula form is rejected by the host."""
    if not name:
        return
    quoted = f'="{name}"'
    try:
        series.Name = quoted
    except Exception:
        try:
            series.Name = name
        except Exception:
            pass

def update_chart_categories(
    prs,
    slide_number: int,
    shape_id: int,
    new_categories: list[str],
):
    try:
        chart = _get_chart(prs, slide_number, shape_id)
        sc = chart.SeriesCollection()

        cat_tuple = tuple(new_categories)
        for i in range(1, sc.Count + 1):
            try:
                sc.Item(i).XValues = cat_tuple
            except Exception:
                pass

        try:
            chart.ChartData.Workbook.Application.Update()
        except Exception:
            pass

        return {"status": "success", "message": "Chart categories updated."}

    except Exception as e:
        return {"status": "error", "message": str(e)}


def update_chart_series(
    prs,
    slide_number: int,
    shape_id: int,
    rename: list[dict] | None = None,
    add: dict | None = None,
    delete: int | None = None,
):
    try:
        chart = _get_chart(prs, slide_number, shape_id)
        sc = chart.SeriesCollection()

        # 1. Rename series
        if rename:
            for req in rename:
                com_idx = _to_com_series_index(req.get("series_index"), sc.Count)
                name = req.get("new_name")
                if com_idx is not None and name:
                    _set_series_literal_name(sc.Item(com_idx), name)

        # 2. Delete series
        del_idx = _to_com_series_index(delete, sc.Count) if isinstance(delete, int) else None
        if del_idx is not None:
            sc.Item(del_idx).Delete()

        # 3. Add series
        if add:
            values = add.get("values") or []
            name = add.get("name")
            categories = add.get("categories")

            if values:
                new_series = sc.NewSeries()
                new_series.Values = tuple(values)

                if name:
                    _set_series_literal_name(new_series, name)

                if categories:
                    new_series.XValues = tuple(categories)

        try:
            chart.ChartData.Workbook.Application.Update()
        except Exception:
            pass

        return {"status": "success", "message": "Chart series updated."}

    except Exception as e:
        return {"status": "error", "message": str(e)}


def update_chart_structure(
    prs,
    slide_number: int,
    shape_id: int,
    chart_type: str | None = None,
    new_title: str | None = None,
    legend_position: str | None = None,
):
    try:
        chart = _get_chart(prs, slide_number, shape_id)

        # 1. Chart type
        if chart_type and chart_type in CHART_TYPES:
            chart.ChartType = CHART_TYPES[chart_type]

        # 2. Title
        if new_title:
            chart.HasTitle = True
            chart.ChartTitle.Text = new_title

        # 3. Legend
        if legend_position:
            if legend_position == "none":
                chart.HasLegend = False
            elif legend_position in LEGEND_POS:
                chart.HasLegend = True
                chart.Legend.Position = LEGEND_POS[legend_position]

        return {"status": "success", "message": "Chart presentation updated."}

    except Exception as e:
        return {"status": "error", "message": str(e)}

def update_chart_axes(
    prs,
    slide_number: int,
    shape_id: int,
    y_axis: dict | None = None,
):
    try:
        if not y_axis:
            return {"status": "success", "message": "No axis changes requested."}

        chart = _get_chart(prs, slide_number, shape_id)
        ax = chart.Axes(2)  # Y-axis

        if "min" in y_axis:
            ax.MinimumScaleIsAuto = False
            ax.MinimumScale = y_axis["min"]

        if "max" in y_axis:
            ax.MaximumScaleIsAuto = False
            ax.MaximumScale = y_axis["max"]

        if "unit" in y_axis:
            ax.MajorUnitIsAuto = False
            ax.MajorUnit = y_axis["unit"]

        if "title" in y_axis:
            ax.HasTitle = True
            ax.AxisTitle.Text = y_axis["title"]

        return {"status": "success", "message": "Chart axes updated."}

    except Exception as e:
        return {"status": "error", "message": str(e)}

def update_chart_colors(
    prs,
    slide_number: int,
    shape_id: int,
    series_colors: list[dict],
):
    """Set chart colors at either the series level or individual data-point level.

    Each entry of `series_colors` supports an optional 0-based `point_index`:
      - {series_index: 0, color_hex: "#FF0000"}                       — colors the whole series
      - {series_index: 0, point_index: 2, color_hex: "#FF0000"}       — colors only the 3rd column/slice/marker
    Use the per-point form to make each column in a column chart, each slice
    in a pie chart, or each marker in a line/scatter chart a different color.
    """
    try:
        chart = _get_chart(prs, slide_number, shape_id)
        sc = chart.SeriesCollection()

        for req in series_colors:
            com_idx = _to_com_series_index(req.get("series_index"), sc.Count)
            hex_c = req.get("color_hex")
            if com_idx is None or not hex_c:
                continue

            rgb = _hex_to_rgb_int(hex_c)
            series = sc.Item(com_idx)

            # Per-point coloring (column/slice/marker level)
            point_index = req.get("point_index")
            if point_index is not None:
                try:
                    pts = series.Points()
                    pt_count = pts.Count
                except Exception:
                    pts = None
                    pt_count = 0
                # Normalize to 1-based COM (accept 0-based input)
                if 0 <= point_index < pt_count:
                    com_pt = point_index + 1
                elif 1 <= point_index <= pt_count:
                    com_pt = point_index
                else:
                    continue
                try:
                    point = pts(com_pt)
                    if chart.ChartType in (4, 65):  # xlLine, xlLineMarkers
                        point.Format.Line.ForeColor.RGB = rgb
                        # also color the marker on line charts when present
                        try:
                            point.MarkerForegroundColor = rgb
                            point.MarkerBackgroundColor = rgb
                        except Exception:
                            pass
                    else:
                        point.Format.Fill.ForeColor.RGB = rgb
                except Exception:
                    pass
                continue

            # Series-level coloring (default)
            try:
                if chart.ChartType in (4, 65):  # xlLine, xlLineMarkers
                    series.Format.Line.ForeColor.RGB = rgb
                else:
                    series.Format.Fill.ForeColor.RGB = rgb
            except Exception:
                pass

        return {"status": "success", "message": "Chart colors updated."}

    except Exception as e:
        return {"status": "error", "message": str(e)}

    
#Table
def cell_text_style(
    prs,
    slide_number: int,
    shape_id: int,
    slide_json: dict,
    *,
    row_index: int,
    col_index: int,
    char_start_index: int,
    target_text: str,
    char_end: int = None,
    font_name: str = None,
    font_size=None,
    bold=None,
    italic=None,
    underline=None,
    color_hex=None,
):
    """
    Change text style inside a single table cell while preserving run-level styles.
    Text content is NOT modified.
    """

    # ------------------------------------------------------------------
    # 1. Resolve shape & cell
    # ------------------------------------------------------------------
    # slide = prs.Slides(slide_number)
    shape = _find_shape_by_id(prs, slide_number, shape_id)

    if not shape or not shape.HasTable:
        raise ValueError(f"Shape {shape_id} is not a table.")

    cell = shape.Table.Cell(row_index, col_index)
    tr = cell.Shape.TextFrame.TextRange

    # ------------------------------------------------------------------
    # 2. Extract full text & normalize target range
    # ------------------------------------------------------------------
    text, _ = _get_text_with_offsets(
        prs,
        slide_number,
        shape_id,
        container="table_cell",
        row_index=row_index,
        col_index=col_index,
    )

    start, end = _normalize_char_range(
        text=text,
        char_start_index=char_start_index,
        target_text=target_text,
        char_end=char_end,
    )

    if start >= end:
        return {
            "operation": "cell_text_style",
            "status": "noop",
            "reason": "empty_range",
            "shape_id": shape_id,
            "row": row_index,
            "col": col_index,
        }

    # ------------------------------------------------------------------
    # 3. Load JSON Runs for this cell
    # ------------------------------------------------------------------
    cell_detail = _get_detail_from_json(
        slide_json,
        shape_id,
        ["More_detail", "Table", "Cells", f"{row_index},{col_index}"]
    )

    runs_data = cell_detail.get("Runs", [])
    if not runs_data:
        return {
            "operation": "cell_text_style",
            "status": "noop",
            "reason": "no_runs",
            "shape_id": shape_id,
            "row": row_index,
            "col": col_index,
        }

    # ------------------------------------------------------------------
    # 4. Apply style per run slice
    # ------------------------------------------------------------------
    for s, e, run in _iter_run_slices_from_runs(runs_data, start, end):
        length = e - s
        if length <= 0:
            continue

        # Convert codepoint offsets to UTF-16 code units for COM
        com_start = codepoint_to_utf16(text, s) + 1   # 1-based
        com_len = codepoint_to_utf16(text, e) - (com_start - 1)
        target_range = tr.Characters(com_start, com_len)
        font = target_range.Font

        # restore original run style first
        _apply_font_snapshot(font, run.get("Font", {}))

        # apply overrides
        _apply_overrides(
            font,
            font_name=font_name,
            font_size=font_size,
            bold=bold,
            italic=italic,
            underline=underline,
            color_hex=color_hex,
        )

    return {
        "operation": "cell_text_style",
        "slide": slide_number,
        "shape_id": shape_id,
        "row": row_index,
        "col": col_index,
        "applied_range": [start, end],
    }



def replace_table_text(
    prs,
    slide_number: int,
    shape_id: int,
    row_index: int,
    col_index: int,
    new_text: str,
    slide_json: dict,
    agent_request,
    auto_resize: bool = False,
):
    # ------------------------------------------------------------
    # 0. Normalize text
    # ------------------------------------------------------------
    if isinstance(new_text, str):
        new_text = new_text.replace("\n", "\r")

    # ------------------------------------------------------------
    # 1. Load cell detail from JSON
    # ------------------------------------------------------------
    cell_key = f"{row_index},{col_index}"

    cell_detail = _get_detail_from_json(
        slide_json,
        shape_id,
        ["More_detail", "Table", "Cells", cell_key]
    )

    if not isinstance(cell_detail, dict):
        cell_detail = {}

    old_runs = cell_detail.get("Runs", []) or []
    old_paragraphs = cell_detail.get("Paragraphs")  # may be None

    run_level = len(old_runs)

    # ------------------------------------------------------------
    # 2. Resolve slide / table / cell
    # ------------------------------------------------------------
    slide = prs.Slides(slide_number)
    shape = _find_shape_recursive(slide.Shapes, shape_id)

    if not shape or not shape.HasTable:
        raise ValueError(f"Shape {shape_id} is not a table.")

    cell = shape.Table.Cell(row_index, col_index)
    tf = cell.Shape.TextFrame
    tr = tf.TextRange
    original_height = cell.Shape.Height

    # ------------------------------------------------------------
    # 3. SIMPLE CASE (run_level <= 1)
    #    - replace text only
    #    - rely on PPT to preserve style
    #    - keep size compensation logic
    # ------------------------------------------------------------
    if run_level <= 1:
        tr.Text = new_text or ""

        if auto_resize:
            _strip_trailing_empty_paragraphs(tf)

        # --- font size compensation (cell height bound) ---
        if auto_resize:
            current_size = tr.Font.Size
            base_size = current_size

            while tr.BoundHeight > original_height and current_size > 9.0:
                current_size -= 0.5
                tr.Font.Size = current_size

            # character-level fallback scaling (COM occasionally lies)
            if tr.Font.Size and base_size != tr.Font.Size:
                scale = current_size / base_size
                if tr.Length > 0:
                    for i in range(1, tr.Length + 1):
                        ch = tr.Characters(i, 1)
                        if ch.Font.Size:
                            ch.Font.Size *= scale

        return {
            "operation": "replace_table_text",
            "mode": "simple_python",
            "slide": slide_number,
            "shape_id": shape_id,
            "row": row_index,
            "col": col_index,
        }

    # ------------------------------------------------------------
    # 4. Build paragraph IR (LLM path)
    # ------------------------------------------------------------
    paragraph_ir = []
    if old_paragraphs:
        paragraph_ir = build_paragraph_ir_from_textframe(
            old_runs, old_paragraphs
        )

    payload = []
    for p in paragraph_ir:
        item = {
            "para_id": p["paragraph_index"],
            "text": p["text"],
            "has_bullet": p["has_bullet"],
            "runs": p["runs"],
        }

        if p["has_bullet"]:
            item["bullet_meta"] = p.get("bullet_meta", {})

        if "Alignment" in p:
            item["Alignment"] = p["Alignment"]

        if "IndentLevel" in p:
            item["IndentLevel"] = p["IndentLevel"]

        payload.append(item)

    sizes = [
        run.get("Font", {}).get("Size")
        for run in old_runs
        if run.get("Font", {}).get("Size")
    ]
    old_base_font_size = max(sizes) if sizes else None

    # ------------------------------------------------------------
    # 5. LLM prompt
    # ------------------------------------------------------------
    if isinstance(agent_request, (tuple, list)) and len(agent_request) >= 3:
        task_description, action_type, slide_contents = agent_request[0], agent_request[1], agent_request[2]
    elif isinstance(agent_request, str):
        task_description, action_type, slide_contents = agent_request, "", ""
    else:
        task_description, action_type, slide_contents = str(agent_request), "", ""
    is_paragraph_mode = len(paragraph_ir) > 1

    if is_paragraph_mode:
        llm_prompt = [
            {"role": "system", "content": PARAGRAPH_STYLE_MAPPING_PROMPT},
            {
                "role": "user",
                "content": json.dumps(
                    {
                        "user_request": {
                            "task_description": task_description,
                            "action_type": action_type,
                            "slide_contents": slide_contents,
                        },
                        "paragraphs": payload,
                        "new_text": new_text,
                    },
                    ensure_ascii=False,
                ),
            },
        ]
    else:
        llm_prompt = [
            {"role": "system", "content": FLATTEXT_STYLE_MAPPING_PROMPT},
            {
                "role": "user",
                "content": json.dumps(
                    {
                        "user_request": {
                            "task_description": task_description,
                            "action_type": action_type,
                            "slide_contents": slide_contents,
                        },
                        "old_runs": old_runs,
                        "new_text": new_text,
                    },
                    ensure_ascii=False,
                ),
            },
        ]

    set_token_log_context(component="style_mapper")
    raw_response = call_llm(model=STYLE_MAPPER_MODEL, messages=llm_prompt)
    if raw_response is None:
        raise RuntimeError("LLM call failed.")

    if is_anthropic_model(STYLE_MAPPER_MODEL):
        response_text = raw_response.output_text
    else:
        response_text = raw_response.output[0].content[0].text
    parsed = parse_llm_response(response_text)
    if isinstance(parsed, tuple):
        parsed = parsed[0]

    # Defensive: same paragraph-count safety net as edit_text_rewrite.
    if is_paragraph_mode:
        parsed = _reconcile_paragraph_mapping(parsed, paragraph_ir, new_text)

    # ------------------------------------------------------------
    # 6–7. Apply text (flat mode)
    # ------------------------------------------------------------
    if not is_paragraph_mode:
        tr.Text = ""
        current_range = tr

        for run in parsed:
            text_seg = run.get("Text", "")
            if not text_seg:
                continue

            inserted = current_range.InsertAfter(text_seg)
            _apply_font_snapshot(inserted.Font, run.get("Font", {}))
            current_range = inserted

    # ------------------------------------------------------------
    # 8. Apply text (paragraph mode)
    #    Replace text per-paragraph, re-apply run-level styles only.
    #    Paragraph structure (alignment, indent, spacing, bullets)
    #    is preserved because we never destroy the paragraph objects.
    # ------------------------------------------------------------
    else:
        para_map = {p["para_id"]: p for p in parsed}
        sorted_pids = sorted(para_map.keys())
        para_count = tr.Paragraphs().Count

        for idx, pid in enumerate(sorted_pids):
            if idx >= para_count:
                break
            para_data = para_map[pid]
            curr_para = tr.Paragraphs(idx + 1)

            # Build full text for this paragraph from mapped runs
            # Strip \r and \n to prevent COM from creating extra paragraphs
            runs_clean = []
            for run in para_data.get("runs", []):
                clean_text = run.get("Text", "").replace("\r", "").replace("\n", "")
                runs_clean.append((clean_text, run.get("Font", {})))

            full_text = "".join(t for t, _ in runs_clean)

            # Replace text within the paragraph (preserves paragraph-level formatting)
            curr_para.Text = full_text

            # Apply run-level font styles via _apply_font_snapshot.
            # Re-read curr_para.Text every iteration: applying Font.Italic /
            # Font.Name can trigger PowerPoint auto-conversion (T → 𝑇), which
            # mutates paragraph length mid-loop. codepoint count is stable, so
            # we use char_offset_cp as the logical cursor.
            char_offset_cp = 0
            for text_seg, font_snap in runs_clean:
                if not text_seg:
                    continue
                cp_len = len(text_seg)
                actual_text = curr_para.Text or ""
                u16_start = codepoint_to_utf16(actual_text, char_offset_cp) + 1
                u16_end = codepoint_to_utf16(actual_text, char_offset_cp + cp_len)
                u16_len = u16_end - (u16_start - 1)
                run_range = curr_para.Characters(u16_start, u16_len)
                _apply_font_snapshot(run_range.Font, font_snap)
                char_offset_cp += cp_len

    # ------------------------------------------------------------
    # 8.5 Remove trailing empty paragraphs (auto_resize)
    # ------------------------------------------------------------
    if auto_resize:
        _strip_trailing_empty_paragraphs(tf)

    # ------------------------------------------------------------
    # 9. Final size compensation (only when auto_resize is enabled)
    # ------------------------------------------------------------
    if auto_resize:
        new_tr = tf.TextRange
        current_size = new_tr.Font.Size or 12.0

        if old_base_font_size:
            current_size = min(current_size, old_base_font_size)

        while new_tr.BoundHeight > original_height and current_size > 9.0:
            current_size -= 0.5
            new_tr.Font.Size = current_size

    return {
        "operation": "replace_table_text",
        "mode": "llm",
        "slide": slide_number,
        "shape_id": shape_id,
        "row": row_index,
        "col": col_index,
        "paragraph_mode": is_paragraph_mode,
    }


def table_layout_style(
    prs,
    slide_number: int,
    shape_id: int,
    *,
    structure_actions: list = None,
    style_actions: list = None,
):
    try:
        # 1. Resolve table
        slide = prs.Slides(slide_number)
        shape = _find_shape_recursive(slide.Shapes, shape_id)

        if not shape or not shape.HasTable:
            return {"status": "error", "message": "Table not found."}

        table = shape.Table

        # 2. Apply structure actions (same as before)
        if structure_actions:
            for action in structure_actions:
                a_type = action.get("type")
                idx = action.get("index")
                if a_type == "add_row": table.Rows.Add()
                elif a_type == "add_col": table.Columns.Add()
                elif a_type == "delete_row" and idx is not None:
                    if 1 <= idx <= table.Rows.Count: table.Rows(idx).Delete()
                elif a_type == "delete_col" and idx is not None:
                    if 1 <= idx <= table.Columns.Count: table.Columns(idx).Delete()

        # 3. Helper: resolve target cells (same as before)
        def iter_target_cells(target):
            r = target.get("row")
            c = target.get("col")
            rows = range(1, table.Rows.Count + 1) if r == "*" else [r]
            cols = range(1, table.Columns.Count + 1) if c == "*" else [c]
            for rr in rows:
                for cc in cols:
                    if rr is not None and cc is not None:
                        if 1 <= rr <= table.Rows.Count and 1 <= cc <= table.Columns.Count:
                            yield table.Cell(rr, cc)

        # 4. Apply visual style & cell modification actions
        if style_actions:
            for action in style_actions:
                target = action.get("target", {})
                styles = action.get("styles", {})

                for cell in iter_target_cells(target):
                    shape_cell = cell.Shape

                    # ---- [added] split feature ----
                    # e.g.: "split": {"rows": 2, "cols": 1}
                    if "split" in styles:
                        split_data = styles["split"]
                        s_rows = split_data.get("rows", 1)
                        s_cols = split_data.get("cols", 1)
                        # Split the cell into s_rows * s_cols parts
                        cell.Split(s_rows, s_cols)

                    # ---- merge (existing) ----
                    if "merge_with" in styles:
                        other = styles["merge_with"]
                        cell.Merge(table.Cell(other["row"], other["col"]))

                    # ---- background color ----
                    if "bg_color_hex" in styles:
                        shape_cell.Fill.Visible = -1
                        shape_cell.Fill.ForeColor.RGB = _hex_to_rgb_int(styles["bg_color_hex"])

                    # ---- alignment & borders (keep existing logic) ----
                    if "vertical_align" in styles:
                        v_map = {"top": 1, "middle": 3, "bottom": 4}
                        val = v_map.get(styles["vertical_align"])
                        try:
                            if val: 
                                shape_cell.TextFrame.VerticalAnchor = val
                        except Exception:
                            # print("Failed to set vertical align for table cell.")
                            pass

                    if "horizontal_align" in styles:
                        h_map = {"left": 1, "center": 2, "right": 3, "justify": 4}
                        val = h_map.get(styles["horizontal_align"])
                        if val: shape_cell.TextFrame.TextRange.ParagraphFormat.Alignment = val

                    if "border" in styles:
                        border = styles["border"]
                        line = shape_cell.Line
                        if "color_hex" in border: line.ForeColor.RGB = _hex_to_rgb_int(border["color_hex"])
                        if "weight" in border: line.Weight = border["weight"]

        return {"status": "success", "shape_id": shape_id, "slide": slide_number}

    except Exception as e:
        return {"status": "error", "message": str(e)}

# --- [C] Enhanced Layout / Geometry Editing ---

def adjust_layout(prs, slide_number, shape_id, left=None, top=None, 
                 width=None, height=None, rotation=None):
    """Adjusts position, size, and rotation of a shape."""
    shape = _find_shape_by_id(prs, slide_number, shape_id)
    
    if left is not None: shape.Left = left
    if top is not None: shape.Top = top
    if width is not None: shape.Width = width
    if height is not None: shape.Height = height
    if rotation is not None: shape.Rotation = rotation
    
    return f"Successfully adjusted layout for Shape {shape_id}."


def distribute_shapes(prs, slide_number, shape_ids, direction="horizontal",
                     spacing=None, use_slide_bounds=False, margin=0):
    """
    Distributes multiple shapes evenly.

    Args:
        direction: 'horizontal' or 'vertical'
        spacing: Fixed spacing between shapes (if None, distribute evenly)
        use_slide_bounds: If True, distribute across the full slide width/height
                          instead of the existing shape range
        margin: Margin from slide edges when use_slide_bounds=True (points)
    """
    shapes = [_find_shape_by_id(prs, slide_number, sid) for sid in shape_ids]
    if len(shapes) < 2:
        return "Need at least 2 shapes to distribute."

    if direction == "horizontal":
        shapes.sort(key=lambda s: s.Left)
        if spacing:
            for i in range(1, len(shapes)):
                shapes[i].Left = shapes[i-1].Left + shapes[i-1].Width + spacing
        else:
            total_width = sum(s.Width for s in shapes)
            if use_slide_bounds:
                start = margin
                end = prs.PageSetup.SlideWidth - margin
            else:
                start = shapes[0].Left
                end = shapes[-1].Left + shapes[-1].Width
            available = end - start - total_width
            gap = available / (len(shapes) - 1) if len(shapes) > 1 else 0

            current_left = start
            for shape in shapes:
                shape.Left = current_left
                current_left += shape.Width + gap
    else:  # vertical
        shapes.sort(key=lambda s: s.Top)
        if spacing:
            for i in range(1, len(shapes)):
                shapes[i].Top = shapes[i-1].Top + shapes[i-1].Height + spacing
        else:
            total_height = sum(s.Height for s in shapes)
            if use_slide_bounds:
                start = margin
                end = prs.PageSetup.SlideHeight - margin
            else:
                start = shapes[0].Top
                end = shapes[-1].Top + shapes[-1].Height
            available = end - start - total_height
            gap = available / (len(shapes) - 1) if len(shapes) > 1 else 0

            current_top = start
            for shape in shapes:
                shape.Top = current_top
                current_top += shape.Height + gap

    return f"Distributed {len(shapes)} shapes {direction}ly."


def align_shapes(prs, slide_number, shape_ids, align_type="left"):
    """
    Aligns multiple shapes.

    Args:
        align_type:
            Shape-to-shape: 'left', 'right', 'top', 'bottom', 'center_h', 'center_v'
            Slide-based:    'slide_center_h', 'slide_center_v', 'slide_center'
    """
    shapes = [_find_shape_by_id(prs, slide_number, sid) for sid in shape_ids]

    # Slide-based alignment works with 1+ shapes
    if align_type.startswith("slide_center"):
        slide_w = prs.PageSetup.SlideWidth
        slide_h = prs.PageSetup.SlideHeight
        # Treat all shapes as a bounding group
        group_left = min(s.Left for s in shapes)
        group_top = min(s.Top for s in shapes)
        group_right = max(s.Left + s.Width for s in shapes)
        group_bottom = max(s.Top + s.Height for s in shapes)
        group_w = group_right - group_left
        group_h = group_bottom - group_top

        if align_type in ("slide_center_h", "slide_center"):
            offset_x = (slide_w - group_w) / 2 - group_left
            for shape in shapes:
                shape.Left += offset_x
        if align_type in ("slide_center_v", "slide_center"):
            offset_y = (slide_h - group_h) / 2 - group_top
            for shape in shapes:
                shape.Top += offset_y

        return f"Aligned {len(shapes)} shapes to {align_type}."

    if len(shapes) < 2:
        return "Need at least 2 shapes to align."

    if align_type == "left":
        left_most = min(s.Left for s in shapes)
        for shape in shapes:
            shape.Left = left_most
    elif align_type == "right":
        right_most = max(s.Left + s.Width for s in shapes)
        for shape in shapes:
            shape.Left = right_most - shape.Width
    elif align_type == "top":
        top_most = min(s.Top for s in shapes)
        for shape in shapes:
            shape.Top = top_most
    elif align_type == "bottom":
        bottom_most = max(s.Top + s.Height for s in shapes)
        for shape in shapes:
            shape.Top = bottom_most - shape.Height
    elif align_type == "center_h":
        avg_center = sum(s.Left + s.Width / 2 for s in shapes) / len(shapes)
        for shape in shapes:
            shape.Left = avg_center - shape.Width / 2
    elif align_type == "center_v":
        avg_center = sum(s.Top + s.Height / 2 for s in shapes) / len(shapes)
        for shape in shapes:
            shape.Top = avg_center - shape.Height / 2

    return f"Aligned {len(shapes)} shapes by {align_type}."


# --- [D] Enhanced Object Lifecycle ---
def create_shape(
    prs,
    slide_number,
    shape_type=1,
    left=100,
    top=100,
    width=100,
    height=100,
    text=None,

    # ---- Shape Fill ----
    fill_color_hex=None,      # "#FF0000"
    fill_transparency=None,   # 0.0 ~ 1.0

    # ---- Shape Line ----
    line_color_hex=None,      # "#000000"
    line_width=None,          # pt
    line_dash_style=None,     # msoLineDashStyle
):
    """
    Create a shape with detailed shape styling.
    - Text: content only (no text style control)
    - Shape style: fill / line properties
    """

    slide = prs.Slides(slide_number)
    shape = slide.Shapes.AddShape(
        shape_type, left, top, width, height
    )

    # ---- Text (content only) ----
    if text and shape.HasTextFrame:
        shape.TextFrame.TextRange.Text = text

    # ---- Fill Style ----
    if fill_color_hex:
        fill = shape.Fill
        fill.Visible = True
        fill.Solid()
        fill.ForeColor.RGB = _hex_to_rgb_int(fill_color_hex)

        if fill_transparency is not None:
            fill.Transparency = fill_transparency

    # ---- Line Style ----
    line = shape.Line

    if line_color_hex:
        line.Visible = True
        line.ForeColor.RGB =_hex_to_rgb_int(line_color_hex)

    if line_width is not None:
        line.Weight = line_width

    if line_dash_style is not None:
        line.DashStyle = line_dash_style

    return f"Shape created successfully (ID: {shape.Id})."

def delete_shape(prs, slide_number, shape_id):
    """Deletes a shape by ID."""
    
    shape = _find_shape_by_id(prs, slide_number, shape_id)
    shape.Delete()

    return f"Shape {shape_id} deleted successfully."

def duplicate_shape(prs, slide_number, shape_id, left, top):
    """Duplicates a shape and moves it to an explicit position."""

    original = _find_shape_by_id(prs, slide_number, shape_id)
    duplicate = original.Duplicate()

    duplicate.Left = left
    duplicate.Top = top

    return f"Shape {shape_id} duplicated (New ID: {duplicate.Id})."


from collections import Counter

# ROLE_DEFAULT_SIZES = {
#     "title": 32,
#     "subtitle": 22,
#     "body": 16,
#     "caption": 12,
#     "footer": 10
# }

# ROLE_DEFAULT_POSITIONS = {
#     "title":    {"left": 50, "top": 40,  "width": 600, "height": 80},
#     "subtitle": {"left": 50, "top": 80, "width": 600, "height": 60},
#     "body":     {"left": 50, "top": 150, "width": 600, "height": 300},
#     "caption":  {"left": 50, "top": 520, "width": 600, "height": 40},
#     "footer":   {"left": 650,  "top": 520, "width": 700, "height": 30},
# }

def _get_most_common_font(slide_json):
    fonts = []

    for shape in slide_json.get("Objects_Detail", []):
        text_frame = shape.get("More_detail", {}).get("TextFrame", {})
        runs = text_frame.get("Runs", [])
        for run in runs:
            font_name = run.get("Font_Name")
            if font_name:
                fonts.append(font_name)

    if not fonts:
        return "Calibri"

    return Counter(fonts).most_common(1)[0][0]

def create_textbox(
    prs,
    slide_number,
    slide_json,
    runs,
    position,
):
    if not runs:
        raise ValueError("runs must contain at least one item.")

    slide = prs.Slides(slide_number)

    textbox = slide.Shapes.AddTextbox(
        1,
        position["left"],
        position["top"],
        position["width"],
        position["height"],
    )

    # Text processing
    text_range = textbox.TextFrame.TextRange
    text_range.Text = ""

    default_font = _get_most_common_font(slide_json)

    for i, run_data in enumerate(runs):
        if i == 0:
            run_obj = text_range
        else:
            run_obj = text_range.InsertAfter(run_data["text"])

        run_obj.Text = run_data["text"]

        snap = dict(run_data)

        if "Name" not in snap:
            snap["Name"] = default_font

        if "Size" not in snap:
            snap["Size"] = 16

        _apply_font_snapshot(run_obj.Font, snap)

    return f"Textbox created (ID: {textbox.Id})"


def create_placeholder(
    presentation,
    slide_number: int,
    slide_json,
    placeholder_type: str,
    runs: list
):
    """
    Creates a placeholder using MS Office default layout positioning.
    Always expects runs and preserves run-level formatting.
    """

    if not runs:
        raise ValueError("runs must contain at least one item.")

    slide = presentation.Slides(slide_number)

    STRUCTURAL_PLACEHOLDERS = {
        "title": 1,
        "center_title": 3,
        "vertical_title": 5,
        "vertical_title2": 19,
        "body": 2,
        "vertical_body": 6,
        "vertical_body2": 20,
        "subtitle": 4,
        "slide_number": 15,
        "header": 16,
        "footer": 17,
        "date": 18,
    }

    # CONTENT_PLACEHOLDERS = {
    #     "chart": 8,
    #     "table": 9,
    #     "picture": 14,
    #     "object": 7,
    #     "media": 12,
    #     "org_chart": 11,
    # }

    PLACEHOLDER_MAP = {
        **STRUCTURAL_PLACEHOLDERS,
        # **CONTENT_PLACEHOLDERS
    }

    placeholder_type = placeholder_type.lower()

    if placeholder_type not in PLACEHOLDER_MAP:
        raise ValueError(f"Unsupported placeholder_type: {placeholder_type}")

    # Check if placeholder already exists — reuse it instead of adding a duplicate
    pp_type = PLACEHOLDER_MAP[placeholder_type]
    shape = None
    for i in range(1, slide.Shapes.Count + 1):
        s = slide.Shapes(i)
        try:
            if s.PlaceholderFormat.Type == pp_type:
                shape = s
                break
        except Exception:
            # Non-placeholder shapes raise on .PlaceholderFormat.Type access
            continue
    if shape is None:
        try:
            shape = slide.Shapes.AddPlaceholder(pp_type)
        except Exception:
            raise ValueError(f"Cannot add placeholder '{placeholder_type}': slide already has maximum of this type.")

    # Text processing (preserve run-level formatting)
    if not shape.HasTextFrame:
        return {
            "status": "success",
            "shape_id": shape.Id,
            "note": "Placeholder has no text frame."
        }

    text_range = shape.TextFrame.TextRange
    text_range.Text = ""

    default_font = _get_most_common_font(slide_json)

    for i, run_data in enumerate(runs):

        if i == 0:
            run_obj = text_range
        else:
            run_obj = text_range.InsertAfter(run_data["text"])

        run_obj.Text = run_data["text"]

        snap = dict(run_data)

        if "Name" not in snap:
            snap["Name"] = default_font

        if "Size" not in snap:
            snap["Size"] = 18 if placeholder_type == "title" else 16

        _apply_font_snapshot(run_obj.Font, snap)

    return {
        "status": "success",
        "shape_id": shape.Id,
        "slide_number": slide_number,
        "placeholder_type": placeholder_type,
        "is_structural": placeholder_type in STRUCTURAL_PLACEHOLDERS
    }


def add_image(prs, slide_number, image_path, left, top, width=None, height=None, alt_text=""):
    """Inserts an image onto the slide."""
    slide = prs.Slides(slide_number)

    if width and height:
        picture = slide.Shapes.AddPicture(image_path, False, True, left, top, width, height)
    else:
        picture = slide.Shapes.AddPicture(image_path, False, True, left, top)

    # If alt_text is provided, set it as AlternativeText
    if alt_text:
        try:
            picture.AlternativeText = alt_text
        except Exception:
            pass

    return f"Image inserted (ID: {picture.Id})."


def insert_image(prs, slide_number, image_source, query_or_prompt,
                 left=100, top=100, width=None, height=None):
    """
    Search, generate, or load an image and insert it onto a slide.

    Args:
        slide_number: 1-based slide index
        image_source: "search" (Pexels/Unsplash), "generate" (Gemini Imagen), or "file" (local path)
        query_or_prompt: Search query, generation prompt, or file path depending on image_source
        left: X position in points (default 100)
        top: Y position in points (default 100)
        width: Image width in points (optional, keeps original if omitted)
        height: Image height in points (optional, keeps original if omitted)
    """
    from editppt.utils.image_providers import search_image, generate_image_gemini

    if image_source == "search":
        image_path = search_image(query_or_prompt)
    elif image_source == "generate":
        image_path = generate_image_gemini(query_or_prompt)
    elif image_source == "file":
        image_path = query_or_prompt
        if not os.path.isfile(image_path):
            raise FileNotFoundError(f"Image file not found: {image_path}")
    else:
        raise ValueError(f"Invalid image_source: '{image_source}'. Must be 'search', 'generate', or 'file'.")

    return add_image(prs, slide_number, image_path, left, top, width, height,
                     alt_text=query_or_prompt)


def edit_image(prs, slide_number, shape_id, edit_prompt):
    """
    Edit an existing image on a slide using AI (Gemini).
    Exports the image from the shape, sends it to Gemini with the edit prompt,
    and replaces the original shape with the edited image at the same position/size.

    Args:
        slide_number: 1-based slide index
        shape_id: Shape ID of the image to edit
        edit_prompt: Natural language instruction for how to edit the image
    """
    import tempfile
    from editppt.utils.image_providers import edit_image_gemini

    shape = _find_shape_by_id(prs, slide_number, shape_id)

    # Save original position and size
    left = shape.Left
    top = shape.Top
    width = shape.Width
    height = shape.Height

    # Export the shape's image to a temp file
    temp_dir = tempfile.mkdtemp()
    export_path = os.path.join(temp_dir, "source.png")
    shape.Export(export_path, 2)  # 2 = ppShapeFormatPNG

    # Edit via Gemini
    edited_path = edit_image_gemini(export_path, edit_prompt)

    # Delete original shape and insert edited image at same position/size
    slide = prs.Slides(slide_number)
    shape.Delete()
    picture = slide.Shapes.AddPicture(edited_path, False, True, left, top, width, height)
    try:
        picture.AlternativeText = edit_prompt
    except Exception:
        pass

    return f"Image edited and replaced (new ID: {picture.Id})."


def group_shapes(prs, slide_number, shape_ids):
    """Groups multiple shapes together."""
    slide = prs.Slides(slide_number)
    shapes = [_find_shape_by_id(prs, slide_number, sid) for sid in shape_ids]
    
    # Create shape range
    shape_range = slide.Shapes.Range([s.Id for s in shapes])
    grouped = shape_range.Group()
    
    return f"Grouped {len(shape_ids)} shapes (Group ID: {grouped.Id})."


def ungroup_shapes(prs, slide_number, group_id):
    """Ungroups a grouped shape."""
    group = _find_shape_by_id(prs, slide_number, group_id)
    ungrouped = group.Ungroup()
    
    return f"Ungrouped shape {group_id} into {ungrouped.Count} shapes."


# --- [E] Enhanced Visual Style / Theme ---

def apply_visual_style(prs, slide_number, shape_id, bg_color_hex=None, 
                      line_color_hex=None, line_weight=None, line_style=None,
                      transparency=None, shadow=None):
    """
    Sets comprehensive visual styles.
    
    Args:
        line_style: 'solid', 'dash', 'dot', 'dash_dot'
        transparency: 0-1 (0=opaque, 1=transparent)
        shadow: True/False to enable/disable shadow
    """
    shape = _find_shape_by_id(prs, slide_number, shape_id)
    results = []

    if bg_color_hex:
        shape.Fill.Visible = True 
        shape.Fill.ForeColor.RGB = _hex_to_rgb_int(bg_color_hex)
        results.append(f"background({bg_color_hex})")

    if transparency is not None:
        shape.Fill.Transparency = transparency
        results.append(f"transparency({transparency})")

    if line_color_hex:
        shape.Line.Visible = True 
        shape.Line.ForeColor.RGB = _hex_to_rgb_int(line_color_hex)
        results.append(f"line color({line_color_hex})")

    if line_weight is not None:
        shape.Line.Visible = True
        shape.Line.Weight = line_weight
        results.append(f"line weight({line_weight}pt)")
    
    if line_style:
        shape.Line.Visible = True
        style_map = {'solid': 1, 'dash': 2, 'dot': 3, 'dash_dot': 4}
        if line_style in style_map:
            shape.Line.DashStyle = style_map[line_style]
            results.append(f"line style({line_style})")
    
    if shadow is not None:
        shape.Shadow.Visible = shadow
        results.append(f"shadow({'on' if shadow else 'off'})")

    return f"Shape {shape_id} visual style updated: " + ", ".join(results) if results else f"No changes applied to Shape {shape_id}."


def apply_gradient_fill(prs, slide_number, shape_id, color1_hex, color2_hex, 
                       gradient_type="linear", angle=0):
    """
    Applies gradient fill to a shape.
    
    Args:
        gradient_type: 'linear', 'radial', 'rectangular', 'path'
        angle: Gradient angle in degrees (for linear)
    """
    shape = _find_shape_by_id(prs, slide_number, shape_id)
    
    gradient_map = {'linear': 1, 'radial': 3, 'rectangular': 4, 'path': 5}
    
    shape.Fill.TwoColorGradient(gradient_map.get(gradient_type, 1), 1)
    shape.Fill.ForeColor.RGB = _hex_to_rgb_int(color1_hex)
    shape.Fill.BackColor.RGB = _hex_to_rgb_int(color2_hex)
    
    if gradient_type == "linear":
        # Set gradient angle
        shape.Fill.GradientAngle = angle
    
    return f"Gradient applied to Shape {shape_id}."


def set_shape_effect(prs, slide_number, shape_id, effect_type, **kwargs):
    """
    Applies special effects to shapes.
    
    Args:
        effect_type: 'glow', 'soft_edge', 'reflection', '3d'
        kwargs: Effect-specific parameters
    """
    shape = _find_shape_by_id(prs, slide_number, shape_id)
    
    if effect_type == "glow":
        color_hex = kwargs.get('color_hex', 'FFFF00')
        size = kwargs.get('size', 10)
        shape.Glow.Color.RGB = _hex_to_rgb_int(color_hex)
        shape.Glow.Radius = size
        return f"Glow effect applied to Shape {shape_id}."
    
    elif effect_type == "soft_edge":
        radius = kwargs.get('radius', 5)
        shape.SoftEdge.Radius = radius
        return f"Soft edge applied to Shape {shape_id}."
    
    elif effect_type == "reflection":
        shape.Reflection.Type = 1  # Enable reflection
        return f"Reflection applied to Shape {shape_id}."
    
    return f"Unknown effect type: {effect_type}"


# --- [F] Enhanced Consistency / Polishing ---

def align_to_object(prs, slide_number, target_id, base_id, side="right", margin=10):
    """Aligns the target shape relative to a base shape with custom margin."""
    target = _find_shape_by_id(prs, slide_number, target_id)
    base = _find_shape_by_id(prs, slide_number, base_id)
    
    if side == "right":
        target.Left = base.Left + base.Width + margin
        target.Top = base.Top
    elif side == "left":
        target.Left = base.Left - target.Width - margin
        target.Top = base.Top
    elif side == "bottom":
        target.Left = base.Left
        target.Top = base.Top + base.Height + margin
    elif side == "top":
        target.Left = base.Left
        target.Top = base.Top - target.Height - margin
    elif side == "center":
        target.Left = base.Left + (base.Width - target.Width) / 2
        target.Top = base.Top + (base.Height - target.Height) / 2
        
    return f"Aligned {target_id} to the {side} of {base_id}."


def match_formatting(prs, slide_number, source_id, target_ids):
    """Copies formatting from source shape to target shapes."""
    source = _find_shape_by_id(prs, slide_number, source_id)
    targets = [_find_shape_by_id(prs, slide_number, tid) for tid in target_ids]
    
    for target in targets:
        # Copy fill
        if source.Fill.Visible:
            target.Fill.ForeColor.RGB = source.Fill.ForeColor.RGB
            target.Fill.Transparency = source.Fill.Transparency
        
        # Copy line
        if source.Line.Visible:
            target.Line.ForeColor.RGB = source.Line.ForeColor.RGB
            target.Line.Weight = source.Line.Weight
        
        # Copy text format if applicable
        if source.HasTextFrame and target.HasTextFrame:
            src_tr = source.TextFrame.TextRange
            tgt_tr = target.TextFrame.TextRange
            tgt_tr.Font.Name = src_tr.Font.Name
            tgt_tr.Font.Size = src_tr.Font.Size
            tgt_tr.Font.Color.RGB = src_tr.Font.Color.RGB
            tgt_tr.Font.Bold = src_tr.Font.Bold
            tgt_tr.Font.Italic = src_tr.Font.Italic
    
    return f"Formatting copied from {source_id} to {len(target_ids)} shape(s)."


def set_z_order(prs, slide_number, shape_id, order="bring_to_front"):
    """
    Changes the z-order (layering) of a shape.
    
    Args:
        order: 'bring_to_front', 'send_to_back', 'bring_forward', 'send_backward'
    """
    shape = _find_shape_by_id(prs, slide_number, shape_id)
    
    if order == "bring_to_front":
        shape.ZOrder(0)  # msoBringToFront
    elif order == "send_to_back":
        shape.ZOrder(1)  # msoSendToBack
    elif order == "bring_forward":
        shape.ZOrder(2)  # msoBringForward
    elif order == "send_backward":
        shape.ZOrder(3)  # msoSendBackward
    
    return f"Z-order changed for Shape {shape_id}."

def duplicate_shape_within_slide(presentation, shape_id, source_slide_number, target_slide_number):
    source_slide = presentation.Slides(source_slide_number)
    target_slide = presentation.Slides(target_slide_number)

    shape = None
    for s in source_slide.Shapes:
        if s.Id == shape_id:
            shape = s
            break

    if shape is None:
        raise ValueError("Shape not found")

    # Copy
    shape.Copy()

    # Paste
    pasted = target_slide.Shapes.Paste()

    # Paste returns a ShapeRange
    new_shape = pasted.Item(1)

    return new_shape.Id


def set_shape_font_size(prs, slide_number: int, shape_id: int, font_size: float):
    """Bulk-set font_size on every run of a shape's TextFrame.

    Designed for the visual fixer: TEXT_OVERFLOW fixes need a uniform shrink
    of the whole textbox, not character-range precision. Avoids the
    `target_text`/`char_start_index` resolution that set_text_style requires.

    Implementation: PowerPoint COM cascades `TextRange.Font.Size = X` to every
    paragraph and run inside the range, overriding per-run sizes. Verified
    empirically (varied 10/14 → uniform 18 in one call). Other font properties
    (color, bold, italic) are preserved.
    """
    shape = _find_shape_by_id(prs, slide_number, shape_id)
    if not shape.HasTextFrame:
        raise ValueError(f"Shape {shape_id} on slide {slide_number} has no TextFrame.")

    tr = shape.TextFrame.TextRange
    if not tr.Text:
        return f"Shape {shape_id} on slide {slide_number} has empty text; nothing to do."

    new_size = float(font_size)
    if new_size <= 0:
        raise ValueError(f"font_size must be > 0 (got {font_size}).")

    tr.Font.Size = new_size
    return f"Set font_size={new_size} on shape_id={shape_id} (slide {slide_number})."


def set_shape_paragraph_spacing(
    prs, slide_number: int, shape_id: int,
    line_spacing: float = None,
    space_before: float = None,
    space_after: float = None,
):
    """Bulk-set paragraph spacing on every paragraph of a shape's TextFrame.

    Visual fixer's vertical-overflow lever. Each parameter is independent and
    optional — only specified ones are applied. Setting via TextRange.
    ParagraphFormat cascades across every paragraph in the range.

    - line_spacing: line-height multiplier (1.0 = single, 1.5 = 1.5x).
      Implemented as LineRuleWithin=0 (Multiple) + SpaceWithin=multiplier.
    - space_before: padding above each paragraph in points.
    - space_after: padding below each paragraph in points.
    """
    if line_spacing is None and space_before is None and space_after is None:
        return f"No spacing params provided for shape_id={shape_id}; nothing to do."

    shape = _find_shape_by_id(prs, slide_number, shape_id)
    if not shape.HasTextFrame:
        raise ValueError(f"Shape {shape_id} on slide {slide_number} has no TextFrame.")

    tr = shape.TextFrame.TextRange
    if not tr.Text:
        return f"Shape {shape_id} on slide {slide_number} has empty text; nothing to do."

    pf = tr.ParagraphFormat
    applied = []
    if line_spacing is not None:
        if line_spacing <= 0:
            raise ValueError(f"line_spacing must be > 0 (got {line_spacing}).")
        # Clamp to [0.8, 1.5] — outside this range deforms typography (too cramped
        # under 0.8, wastes vertical space and worsens overflow above 1.5).
        clamped = max(0.8, min(1.5, float(line_spacing)))
        pf.LineRuleWithin = 0  # 0 = Multiple (interpret SpaceWithin as multiplier)
        pf.SpaceWithin = clamped
        if clamped != float(line_spacing):
            applied.append(f"line_spacing={clamped} (clamped from {line_spacing})")
        else:
            applied.append(f"line_spacing={clamped}")
    if space_before is not None:
        if space_before < 0:
            raise ValueError(f"space_before must be >= 0 (got {space_before}).")
        pf.SpaceBefore = float(space_before)
        applied.append(f"space_before={space_before}")
    if space_after is not None:
        if space_after < 0:
            raise ValueError(f"space_after must be >= 0 (got {space_after}).")
        pf.SpaceAfter = float(space_after)
        applied.append(f"space_after={space_after}")

    return (
        f"Applied paragraph spacing on shape_id={shape_id} (slide {slide_number}): "
        f"{', '.join(applied)}."
    )


# --- [G] Slide Management ---

def add_slide(prs, layout_index=1, position=None):
    """
    Adds a new slide to the presentation.
    
    Args:
        layout_index: Layout to use (1-based)
        position: Where to insert (None = end)
    """
    # PpSlideLayout enum: 1=Title, 2=Text, 7=Blank, 12=ObjectAndText, etc.
    # CustomLayouts access can fail in dynamic dispatch COM, so use integer layout index
    layout_enum = layout_index if layout_index else 7  # default: Blank
    total = prs.Slides.Count
    target_pos = position if position else total + 1
    target_pos = max(1, min(target_pos, total + 1))  # clamp to 1 ~ Count+1

    # COM Slides.Add valid range is 1~Slides.Count
    # For appending at the end (target_pos > Count), use Add(Count) then MoveTo
    if total == 0:
        new_slide = prs.Slides.Add(1, layout_enum)
    elif target_pos <= total:
        new_slide = prs.Slides.Add(target_pos, layout_enum)
    else:
        new_slide = prs.Slides.Add(total, layout_enum)
        new_slide.MoveTo(prs.Slides.Count)

    return f"Slide added at position {new_slide.SlideIndex}."


def delete_slide(prs, slide_number):
    """Deletes a specific slide."""
    prs.Slides(slide_number).Delete()
    return f"Slide {slide_number} deleted."


def duplicate_slide(prs, slide_number):
    """Duplicates a specific slide."""
    original = prs.Slides(slide_number)
    duplicate = original.Duplicate()
    return f"Slide {slide_number} duplicated to position {duplicate(1).SlideIndex}."


# --- [H] Table Operations ---

def add_table(prs, slide_number, rows, cols, left, top, width, height):
    """Creates a table on the slide."""
    slide = prs.Slides(slide_number)
    table = slide.Shapes.AddTable(rows, cols, left, top, width, height)
    return f"Table created (ID: {table.Id}, {rows}x{cols})."


def update_table_cell(prs, slide_number, table_id, row, col, text, 
                     font_size=None, color_hex=None, bg_color_hex=None):
    """Updates content and style of a specific table cell."""
    table_shape = _find_shape_by_id(prs, slide_number, table_id)
    
    if not table_shape.HasTable:
        return f"Error: Shape {table_id} is not a table."
    
    cell = table_shape.Table.Cell(row, col)
    cell.Shape.TextFrame.TextRange.Text = text
    
    if font_size:
        cell.Shape.TextFrame.TextRange.Font.Size = font_size
    if color_hex:
        cell.Shape.TextFrame.TextRange.Font.Color.RGB = _hex_to_rgb_int(color_hex)
    if bg_color_hex:
        cell.Shape.Fill.ForeColor.RGB = _hex_to_rgb_int(bg_color_hex)
    
    return f"Table cell ({row},{col}) updated."


# --- [I] Animation & Transition ---

def add_animation(prs, slide_number, shape_id, effect_type="appear", 
                 trigger="on_click", duration=1.0):
    """
    Adds animation to a shape.
    
    Args:
        effect_type: 'appear', 'fade', 'fly_in', 'zoom', etc.
        trigger: 'on_click', 'with_previous', 'after_previous'
        duration: Animation duration in seconds
    """
    slide = prs.Slides(slide_number)
    shape = _find_shape_by_id(prs, slide_number, shape_id)
    
    effect_map = {
        'appear': 1,  # msoAnimEffectAppear
        'fade': 10,   # msoAnimEffectFade
        'fly_in': 22, # msoAnimEffectFly
        'zoom': 88,   # msoAnimEffectZoom
    }
    
    effect = slide.TimeLine.MainSequence.AddEffect(
        shape, effect_map.get(effect_type, 1), trigger=1, index=-1
    )
    effect.Timing.Duration = duration
    
    return f"Animation '{effect_type}' added to Shape {shape_id}."


def set_slide_transition(prs, slide_number, transition_type="fade",
                        duration=1.0, advance_on_time=None):
    """
    Sets slide transition effect.

    Args:
        slide_number: 1-based slide index
        transition_type: One of: 'none', 'fade', 'push', 'wipe', 'split', 'cover',
                         'uncover', 'dissolve', 'cut', 'random_bars', 'checkerboard', 'morph'
        duration: Transition duration in seconds (default 1.0)
        advance_on_time: Auto-advance after N seconds (None = click to advance)
    """
    slide = prs.Slides(slide_number)

    transition_map = {
        'none': 0,            # ppEffectNone
        'fade': 1793,         # ppEffectFade
        'fade_smoothly': 3849,# ppEffectFadeSmoothly
        'push': 3852,         # ppEffectPushDown
        'wipe': 2819,         # ppEffectWipeRight
        'split': 3585,        # ppEffectSplitHorizontalOut
        'cover': 1284,        # ppEffectCoverDown
        'uncover': 2052,      # ppEffectUncoverDown
        'dissolve': 1537,     # ppEffectDissolve
        'cut': 257,           # ppEffectCut
        'random_bars': 2305,  # ppEffectRandomBarsHorizontal
        'checkerboard': 1025, # ppEffectCheckerboardAcross
        'morph': 3854,        # ppEffectMorphByObject (PowerPoint 2016+)
    }

    effect_value = transition_map.get(transition_type)
    if effect_value is None:
        available = ", ".join(transition_map.keys())
        raise ValueError(f"Unknown transition_type '{transition_type}'. Available: {available}")

    slide.SlideShowTransition.EntryEffect = effect_value
    slide.SlideShowTransition.Duration = duration

    if advance_on_time:
        slide.SlideShowTransition.AdvanceOnTime = True
        slide.SlideShowTransition.AdvanceTime = advance_on_time
    else:
        slide.SlideShowTransition.AdvanceOnClick = True

    return f"Transition '{transition_type}' applied to slide {slide_number}."


def set_slide_background(prs, slide_number, fill_type="solid", color_hex="#FFFFFF",
                         color_hex_2=None, gradient_style=1):
    """
    Sets the background of a slide.

    Args:
        slide_number: 1-based slide index
        fill_type: 'solid', 'gradient', or 'none' (reset to master)
        color_hex: Primary color hex (e.g. '#ADD8E6')
        color_hex_2: Secondary color hex for gradient
        gradient_style: 1=horizontal, 2=vertical, 3=diagonal_up, 4=diagonal_down
    """
    slide = prs.Slides(slide_number)
    bg = slide.Background
    fill = bg.Fill

    if fill_type == "none":
        slide.FollowMasterBackground = True
        return f"Slide {slide_number} background reset to master."

    slide.FollowMasterBackground = False

    if fill_type == "solid":
        fill.Solid()
        fill.ForeColor.RGB = _hex_to_rgb_int(color_hex)
    elif fill_type == "gradient":
        c2 = color_hex_2 or "#FFFFFF"
        fill.TwoColorGradient(gradient_style, 1)
        fill.ForeColor.RGB = _hex_to_rgb_int(color_hex)
        fill.BackColor.RGB = _hex_to_rgb_int(c2)
    else:
        raise ValueError(f"Unknown fill_type '{fill_type}'. Use 'solid', 'gradient', or 'none'.")

    return f"Slide {slide_number} background set to {fill_type} ({color_hex})."


FUNCTION_MAP = {
    # Text editing
    "set_text_style": set_text_style,
    "edit_text_insert": edit_text_insert,
    "edit_text_delete": edit_text_delete,
    "edit_text_replace": edit_text_replace,
    "edit_text_rewrite": edit_text_rewrite,
    "set_paragraph_alignment": set_paragraph_alignment,
    "manage_bullet_points": manage_bullet_points,

    # Charts
    "update_chart_categories": update_chart_categories,
    "update_chart_series": update_chart_series,
    "update_chart_structure": update_chart_structure,
    "update_chart_axes": update_chart_axes,
    "update_chart_colors": update_chart_colors,

    # Tables
    "cell_text_style": cell_text_style,
    "replace_table_text": replace_table_text,
    "table_layout_style": table_layout_style,

    # Shape & layout
    "adjust_layout": adjust_layout,
    "distribute_shapes": distribute_shapes,
    "align_shapes": align_shapes,
    "create_textbox": create_textbox,
    "create_placeholder": create_placeholder,
    "create_shape": create_shape,
    "delete_shape": delete_shape,
    "duplicate_shape": duplicate_shape,
    "duplicate_shape_within_slide": duplicate_shape_within_slide,
    "apply_visual_style": apply_visual_style,
    "apply_gradient_fill": apply_gradient_fill,
    "set_shape_font_size": set_shape_font_size,
    "set_shape_paragraph_spacing": set_shape_paragraph_spacing,

    # Image
    "insert_image": insert_image,
    "edit_image": edit_image,

    # Slide-level
    "add_slide": add_slide,
    "delete_slide": delete_slide,
    "duplicate_slide": duplicate_slide,
    "set_slide_transition": set_slide_transition,
    "set_slide_background": set_slide_background,
}
