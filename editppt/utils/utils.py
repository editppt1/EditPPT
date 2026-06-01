import pywintypes
import openai
from openai import OpenAI

import re
import json
import ast
import unicodedata

from editppt.utils.logger_manual import *
from editppt.utils.msoffice_map import *


def _com_has(obj, attr: str) -> bool:
    """COM-safe replacement for ``hasattr``.

    Python's ``hasattr`` only swallows ``AttributeError``. PowerPoint COM
    objects raise ``pywintypes.com_error`` ("This member can only be accessed
    for a Chart object." etc.) when an attribute is conditionally available,
    so ``hasattr(shape, "Chart")`` propagates that exception. Always use this
    helper for shape-feature probing on COM objects.
    """
    try:
        getattr(obj, attr)
        return True
    except Exception:
        return False


def codepoint_to_utf16(text: str, cp_offset: int) -> int:
    """Convert a Python codepoint offset to a UTF-16 code-unit offset.

    COM's TextRange.Characters() uses 1-based UTF-16 indices, while Python
    strings are indexed by codepoints.  Supplementary characters (U+10000+)
    occupy 2 UTF-16 code units but only 1 Python codepoint.
    """
    return len(text[:cp_offset].encode("utf-16-le")) // 2


_SMART_QUOTE_MAP = str.maketrans({
    "“": '"', "”": '"', "„": '"', "‟": '"',
    "‘": "'", "’": "'", "‚": "'", "‛": "'",
})


def _sanitize_payload(payload: str) -> str:
    """Best-effort cleanup of common LLM JSON output issues."""
    # Strip BOM
    if payload.startswith("﻿"):
        payload = payload[1:]
    # Smart quotes → straight quotes
    payload = payload.translate(_SMART_QUOTE_MAP)
    # Strip /* */ block comments
    payload = re.sub(r"/\*.*?\*/", "", payload, flags=re.DOTALL)
    # Strip // line comments — only when they begin at line start (optionally
    # indented). This avoids damaging string values that legitimately contain
    # "//" (e.g. URLs, file paths). Comments mid-value are extremely rare in
    # LLM output anyway.
    payload = re.sub(r"^[ \t]*//[^\n]*\n?", "", payload, flags=re.MULTILINE)
    # Trailing commas
    payload = re.sub(r",\s*([\}\]])", r"\1", payload)
    return payload


def parse_llm_response(response):
    """
    Parse JSON or Python-like structures from an LLM response.

    Returns:
    - (parsed_obj, None) on success
    - (None, (exception, payload_or_response)) on failure
    """

    # 1. Input validation
    if not response or not isinstance(response, str):
        e = ValueError("response is empty or not a string")
        return None, (e, response)

    # 2. Remove markdown code fences
    response_clean = re.sub(r'```(?:json)?', '', response).strip()

    # 3. Extract JSON / list
    match = re.search(r'(\{.*\}|\[.*\])', response_clean, re.DOTALL)
    if not match:
        e = ValueError("No JSON object could be decoded")
        return None, (e, response_clean)

    payload = match.group(1)

    # 4. Local sanitization (BOM, smart quotes, comments, trailing commas)
    payload = _sanitize_payload(payload)

    try:
        parsed = json.loads(payload)
        return parsed, None
    except json.JSONDecodeError:
        # 5. Python literal fallback (handles single quotes, True/False/None)
        try:
            parsed = ast.literal_eval(payload)
            return parsed, None
        except Exception as e_ast:
            return None, (e_ast, payload)



def extract_content_after_edit(plan_json):
    result = []
    
    if 'tasks' in plan_json and len(plan_json['tasks']) > 0:
        for task in plan_json['tasks']:
            if 'content after edit' in task and isinstance(task['content after edit'], list):
                result.extend(task['content after edit'])
    
    return result

def extract_last_text_content(plan_json):
    last_text = ""
    
    if 'tasks' in plan_json and len(plan_json['tasks']) > 0:
        for task in plan_json['tasks']:
            if 'contents' in task:
                contents_str = task['contents']
                # Find all 'Text content:' patterns and collect them into a list
                text_contents = re.findall(r'Text content: (.*?)(?=\n\s+Font:|$)', contents_str, re.DOTALL)
                
                # Return the last 'Text content:' value (empty string if none found)
                if text_contents:
                    last_text = text_contents[-1].strip()
    
    return last_text

def create_thinking_queue(plan_json):
    # thinking queue
    temp_tasks = []
    temp_actions = []
    
    print_data_ = ""

    for i in range(len(plan_json['tasks'])):
        temp_tasks.append(plan_json['tasks'][i]['target'])
        temp_actions.append(plan_json['tasks'][i]['action'])
    
    for i in range(len(temp_tasks)):
        print_data_ += f"• Applying '{temp_actions[i]}' to '{temp_tasks[i]}'.\n"
    
    return print_data_


import openai
from openai import OpenAI
import tiktoken

# Per-model token pricing (example: USD per 1K tokens)
PRICING = {
    #"gpt-4.1-2025-04-14":    {"prompt": 0.03/1000, "completion": 0.06/1000},
    "gpt-4.1-mini-2025-04-14":{"prompt": 0.4/1000000, "completion": 1.6/1000000},
    #"gpt-4.1-nano-2025-04-14":{"prompt": 0.001/1000, "completion": 0.001/1000},
    #"o4-mini":               {"prompt": 0.002/1000, "completion": 0.002/1000},
}

def count_tokens(text: str, model: str) -> int:
    """Count tokens using tiktoken"""
    try:
        enc = tiktoken.encoding_for_model(model)
    except KeyError:
        enc = tiktoken.get_encoding("cl100k_base")
    return len(enc.encode(text))

def _call_gpt_api(prompt: str, api_key: str, model: str):
    # --- API key setup and model validation/mapping ---
    openai.api_key = api_key

    allowed = ["gpt-4.1", "gpt-4.1-mini", "gpt-4.1-nano", "o4-mini"]
    if model not in allowed:
        raise ValueError(f"Model must be one of {allowed}")

    if model == "gpt-4.1":
        model = "gpt-4.1-2025-04-14"
    elif model == "gpt-4.1-mini":
        model = "gpt-4.1-mini-2025-04-14"
    elif model == "gpt-4.1-nano":
        model = "gpt-4.1-nano-2025-04-14"
    # o4-mini stays as-is

    # --- API call ---
    client = OpenAI(api_key=api_key)
    response = client.responses.create(
        model=model,
        instructions="You are a coding assistant that edits PowerPoint slides.",
        input=prompt,
    )
    text = response.output_text

    # --- Token counting (use usage field if available, otherwise count_tokens) ---
    if getattr(response, "usage", None):
        inp_toks = response.usage.input_tokens
        out_toks = response.usage.output_tokens
    else:
        inp_toks = count_tokens(prompt, model)
        out_toks = count_tokens(text, model)

    # --- Cost calculation ---
    rates = PRICING.get(model)
    if rates is None:
        total_cost = None
    else:
        total_cost = inp_toks * rates["prompt"] + out_toks * rates["completion"]

    # --- Always return 4 values ---
    return text, inp_toks, out_toks, total_cost



def get_shape_type(shape_type):
    """Convert shape type number to string"""
    return SHAPE_TYPE_MAP.get(shape_type, f"Unknown Type ({shape_type})")

def get_placeholder_type(placeholder_type):
    """Convert placeholder type number to string"""
    return PLACEHOLDER_TYPE_MAP.get(placeholder_type, f"Unknown Placeholder ({placeholder_type})")

def _bgr_int_to_hex(bgr):
    """
    Convert Office COM BGR int (0x00BBGGRR) to '#RRGGBB'.
    """
    if bgr is None:
        return None
    try:
        bgr = int(bgr)
    except Exception:
        return None

    r = bgr & 0xFF
    g = (bgr >> 8) & 0xFF
    b = (bgr >> 16) & 0xFF
    return f"#{r:02X}{g:02X}{b:02X}"

import traceback  # Added for error traceback

def safe(obj, attr, default=None):
    """Safely get an attribute, returning default if an error occurs."""
    try:
        if obj is None:
            return default
        if hasattr(obj, attr):
            val = getattr(obj, attr, default)
            if val is None:
                return default
            return val
        return default
    except Exception:
        return default


def rgb_of(font):
    # if font is None:
    #     return None
    # rgb = None
    # try:
    #     fill = safe(font, "Fill")
    #     fore = safe(fill, "ForeColor")
    #     temp = safe(fore, "RGB")
    #     if temp is not None:
    #         rgb = temp
    # except Exception:
    #     pass
    # if rgb is None:
    #     try:
    #         col = safe(font, "Color")
    #         temp = safe(col, "RGB")
    #         if temp is not None:
    #             rgb = temp
    #     except Exception:
    #         pass
    # return rgb
    try:
        c = safe(font, "Color")
        if not c:
            return None

        bgr = safe(c, "RGB")
        if bgr is None:
            return None

        try:
            bgr = int(bgr)
        except Exception:
            return None

        r = bgr & 0xFF
        g = (bgr >> 8) & 0xFF
        b = (bgr >> 16) & 0xFF

        return f"#{r:02X}{g:02X}{b:02X}"

    except Exception:
        return None


def snap(font):
    if font is None:
        return (None, 0.0, False, False, False, None, False, False, False)
    size_val = safe(font, "Size", 0)
    try:
        size_f = float(size_val)
    except Exception:
        size_f = 0.0
    return (
        safe(font, "Name"),
        round(size_f, 1),
        bool(safe(font, "Bold", 0)),
        bool(safe(font, "Italic", 0)),
        bool(safe(font, "Underline", 0)),
        rgb_of(font),
        bool(safe(font, "Strikethrough", 0)),
        bool(safe(font, "Subscript", 0)),
        bool(safe(font, "Superscript", 0)),
    )


def make_run_dict(text_range_segment):
    if text_range_segment is None:
        return {"Text": ""}

    text = safe(text_range_segment, "Text", "")

    run = {"Text": text}

    f = safe(text_range_segment, "Font")
    if not f:
        return run

    font_dict = {}

    # Name / Size
    name = safe(f, "Name")
    if name is not None:
        font_dict["Name"] = name

    size = safe(f, "Size")
    if size is not None:
        font_dict["Size"] = size

    # Boolean styles (only record when True)
    if safe(f, "Bold", 0):
        font_dict["Bold"] = True
    if safe(f, "Italic", 0):
        font_dict["Italic"] = True
    if safe(f, "Underline", 0):
        font_dict["Underline"] = True
    if safe(f, "Strikethrough", 0):
        font_dict["Strikethrough"] = True
    if safe(f, "Subscript", 0):
        font_dict["Subscript"] = True
    if safe(f, "Superscript", 0):
        font_dict["Superscript"] = True

    # Color
    rgb = rgb_of(f)
    if rgb is not None:
        font_dict["Color"] = rgb


    hyperlink = None
    try:
        act = safe(safe(text_range_segment, "ActionSettings"), 1)
        hyperlink = safe(safe(act, "Hyperlink"), "Address")
    except Exception:
        pass

    if font_dict:
        run["Font"] = font_dict
    if hyperlink is not None:
        run["Hyperlink"] = hyperlink

    return run


def parse_paragraph_bullets(text_frame):
    result = []

    if not text_frame or not safe(text_frame, "HasText", False):
        return result

    tr = text_frame.TextRange
    try:
        para_count = tr.Paragraphs().Count
    except Exception:
        return result

    style_counters = {}
    h_map = {2: "center", 3: "right", 4: "justify"}
    for i in range(1, para_count + 1):
        try:
            current_para = tr.Paragraphs(i)
            pf = current_para.ParagraphFormat

            # 1. Extract basic info
            p_text = getattr(current_para, "Text", "")
            alignment = getattr(pf, "Alignment", 1)
            indent_level = getattr(current_para, "IndentLevel", 1)
            
            # Create required fields first
            para_info = {
                "ParagraphIndex": i,
                "Text": p_text,
            }
            
            # 2. Alignment & Indent (keep existing logic: add only if non-default)
            if alignment != 1:
                para_info["Alignment"] = h_map.get(alignment, alignment)
            if indent_level > 1:
                para_info["IndentLevel"] = indent_level

            # 2-1. Paragraph spacing
            try:
                space_before = getattr(pf, "SpaceBefore", None)
                if space_before is not None and space_before > 0:
                    para_info["SpaceBefore"] = space_before
            except Exception:
                pass
            try:
                space_after = getattr(pf, "SpaceAfter", None)
                if space_after is not None and space_after > 0:
                    para_info["SpaceAfter"] = space_after
            except Exception:
                pass
            try:
                line_rule = getattr(pf, "LineRuleWithin", None)
                line_spacing = getattr(pf, "SpaceWithin", None)
                if line_rule is not None:
                    para_info["LineSpacingRule"] = line_rule
                if line_spacing is not None and line_spacing > 0:
                    para_info["LineSpacing"] = line_spacing
            except Exception:
                pass

            # 3. Check bullet info
            bullet = pf.Bullet
            is_visible = (bullet.Visible != 0)
            b_type = bullet.Type if is_visible else 0

            # Only add fields when bullet is active
            if is_visible and b_type != 0:
                para_info["HasBullet"] = True
                
                actual_label = ""
                style_code = bullet.Style
                style_info = BULLET_STYLE_MAP.get(style_code, ["Standard", "1."])
                
                if b_type == 2:  # Numbered
                    c = style_counters.get(style_code, 0) + 1
                    style_counters[style_code] = c
                    
                    try:
                        actual_label = current_para.TextRange.ListFormat.ListString
                    except Exception:
                        actual_label = ""
                    
                    if not actual_label:
                        fmt = style_info[1]
                        if "1" in fmt: actual_label = fmt.replace("1", str(c))
                        elif "a" in fmt: actual_label = fmt.replace("a", chr(96 + (c % 26 or 26)))
                        elif "A" in fmt: actual_label = fmt.replace("A", chr(64 + (c % 26 or 26)))
                        else: actual_label = fmt
                
                elif b_type == 1:  # Symbol
                    char_code = bullet.Character
                    char_info = BULLET_CHAR_MAP.get(char_code, [None, "•"])
                    actual_label = char_info[1]

                # Update bullet detail info
                para_info.update({
                    "BulletType": b_type,
                    "ActualLabel": actual_label,
                    "BulletCharacter": style_info[0] if b_type == 2 else "Symbol",
                    "CharacterCode": bullet.Character if b_type == 1 else None, # added
                    "BulletStyle": bullet.Style if b_type == 2 else None,
                })

            result.append(para_info)

        except Exception as e:
            continue

    return result

# def build_paragraph_ir_from_textframe(runs, paragraphs):
#     """
#     Generate Paragraph IR (Run-Paragraph mapping with Alignment and IndentLevel)
#     """
#     paragraph_ir = []
#     current_offset = 0

#     # 1. Set paragraph-based regions
#     for p in paragraphs:
#         p_text = p.get("Text", "")
#         p_len = len(p_text)
        
#         start_idx = current_offset
#         end_idx = start_idx + p_len

#         # Map attributes extracted from parse_paragraph_bullets to style_meta
#         p_info = {
#             "paragraph_index": p["ParagraphIndex"],
#             "text": p_text,
#             "has_bullet": p.get("HasBullet", False),
#             "start": start_idx,
#             "end": end_idx,
#             "alignment": p.get("Alignment", "left"),  # default 'left' (1)
#             "indent_level": p.get("IndentLevel", 1),   # default 1
#             "bullet": {
#                 "type": p.get("BulletType"),
#                 "label": p.get("ActualLabel"),
#                 "char": p.get("BulletCharacter"),
#                 "character_code": p.get("CharacterCode"),
#                 "bullet_style": p.get("BulletStyle")
#                 },
#             "runs": []
#         }
#         paragraph_ir.append(p_info)
#         current_offset = end_idx

#     # 2. Run splitting and mapping (keep existing logic)
#     for run in runs:
#         run_text = run.get("Text", "")
#         run_start = run.get("Run_Start_Index", 0)
#         run_end = run_start + len(run_text)
        
#         for p in paragraph_ir:
#             overlap_start = max(run_start, p["start"])
#             overlap_end = min(run_end, p["end"])

#             if overlap_start < overlap_end:
#                 rel_start = overlap_start - run_start
#                 rel_end = overlap_end - run_start
                
#                 fragment_text = run_text[rel_start:rel_end]
                
#                 run_fragment = {
#                     "text": fragment_text,
#                     "font": run.get("Font", {}),
#                     "start": overlap_start,
#                     "end": overlap_end
#                 }
#                 p["runs"].append(run_fragment)

#     return paragraph_ir

from collections import deque

def build_paragraph_ir_from_textframe(runs, paragraphs):
    # 1. Build a full stream of (character, font) tuples from all Runs
    full_stream = []
    for run in runs:
        run_text = run.get("Text", "")
        run_font = run.get("Font", {})
        for char in run_text:
            full_stream.append((char, run_font))
    
    # Convert to a queue for easy front-to-back consumption
    stream_queue = deque(full_stream)
    
    paragraph_ir = []

    # 2. Iterate by paragraph
    for p in paragraphs:
        p_text = p.get("Text", "")
        
        p_info = {
            "paragraph_index": p["ParagraphIndex"],
            "text": p_text,
            "has_bullet": p.get("HasBullet", False),
            "runs": []
        }
        if "Alignment" in p:
            p_info["Alignment"] = p["Alignment"]
        if "IndentLevel" in p:
            p_info["IndentLevel"] = p["IndentLevel"]
        if "SpaceBefore" in p:
            p_info["SpaceBefore"] = p["SpaceBefore"]
        if "SpaceAfter" in p:
            p_info["SpaceAfter"] = p["SpaceAfter"]
        if "LineSpacingRule" in p:
            p_info["LineSpacingRule"] = p["LineSpacingRule"]
        if "LineSpacing" in p:
            p_info["LineSpacing"] = p["LineSpacing"]
        if p_info["has_bullet"]:
            p_info["bullet_meta"] = {
                "type": p.get("BulletType"),
                "label": p.get("ActualLabel"),
                "char": p.get("BulletCharacter"),
                "character_code": p.get("CharacterCode"),
                "bullet_style": p.get("BulletStyle")
            }

        # 3. Match one style per character in the paragraph text
        current_runs = []
        for char in p_text:
            if stream_queue:
                stream_char, stream_font = stream_queue.popleft()
                # If stream_char differs from char (handling data mismatch),
                # prioritize p_text's character and only take the style
                current_runs.append({"char": char, "font": stream_font})
            else:
                # Apply default font if stream is exhausted
                current_runs.append({"char": char, "font": {}})

        # 4. Merge consecutive characters with the same font back into Runs (compression)
        if current_runs:
            merged_runs = []
            if current_runs:
                temp_text = current_runs[0]["char"]
                last_font = current_runs[0]["font"]
                
                for i in range(1, len(current_runs)):
                    if current_runs[i]["font"] == last_font:
                        temp_text += current_runs[i]["char"]
                    else:
                        merged_runs.append({"text": temp_text, "font": last_font})
                        temp_text = current_runs[i]["char"]
                        last_font = current_runs[i]["font"]
                
                merged_runs.append({"text": temp_text, "font": last_font})
            p_info["runs"] = merged_runs

        paragraph_ir.append(p_info)

    return paragraph_ir

def parse_text_frame_debug(text_frame):
    out = {"Has Text": False}
    if not safe(text_frame, "HasText", False):
        return out
    tr = safe(text_frame, "TextRange")
    if not tr:
        return out
    full = safe(tr, "Text", "")
    out.update({"Has Text": True, "Text": full, "Runs": []})

    if not full:
        out["Paragraphs"] = []
        return out

    runs = []
    # Use COM's Length (UTF-16 code units) instead of Python len() (Unicode codepoints).
    # Supplementary characters (U+10000+, e.g. math italic 𝑇) are 2 code units in COM
    # but 1 codepoint in Python. Using len(full) causes the loop to stop early,
    # missing the final runs.
    n = safe(tr, "Length", len(full))
    try:
        cur_idx = 1
        cur_snap = snap(safe(tr.Characters(cur_idx, 1), "Font"))

        for i in range(2, n + 1):
            nxt_snap = snap(safe(tr.Characters(i, 1), "Font"))
            if nxt_snap != cur_snap:
                seg_len = i - cur_idx
                if seg_len > 0:
                    run = make_run_dict(tr.Characters(cur_idx, seg_len))
                    runs.append(run)
                cur_idx = i
                cur_snap = nxt_snap

        last_len = n - cur_idx + 1
        if last_len > 0:
            run = make_run_dict(tr.Characters(cur_idx, last_len))
            runs.append(run)
    except Exception as e:
        print(f"Error parsing runs: {e}")
        traceback.print_exc()
        runs.append(make_run_dict(tr))

    # Assign Run_Start_Index in Python codepoint offsets (not UTF-16 code units).
    # This keeps indices consistent with _normalize_char_range and Python string
    # slicing.  Conversion to UTF-16 happens at the COM call boundary.
    cp_offset = 0
    for run in runs:
        run["Run_Start_Index"] = cp_offset
        cp_offset += len(run.get("Text", ""))

    out["Runs"] = runs

    # Add paragraph/bullet info here
    out["Paragraphs"] = parse_paragraph_bullets(text_frame)

    return out


# def parse_table(table):
#     """Parse table info (returns result as dict)"""
#     result = {}
#     try:
#         rows = getattr(table.Rows, "Count", 0)
#         cols = getattr(table.Columns, "Count", 0)
#         result["Dimensions"] = {"Rows": rows, "Columns": cols}
#         result["FirstRow"]   = getattr(table, "FirstRow", None)
#         result["LastRow"]    = getattr(table, "LastRow", None)
#         result["FirstCol"]   = getattr(table, "FirstCol", None)
#         result["LastCol"]    = getattr(table, "LastCol", None)

#         # Sample cell contents
#         samples = {}
#         max_r = min(3, rows)
#         max_c = min(3, cols)
#         for r in range(1, max_r + 1):
#             for c in range(1, max_c + 1):
#                 key = f"Cell({r},{c})"
#                 try:
#                     txt = table.Cell(r, c).Shape.TextFrame.TextRange.Text
#                     samples[key] = txt[:30] + ("..." if len(txt) > 30 else "")
#                 except Exception:
#                     samples[key] = None
#         result["Sample Cells"] = samples

#     except Exception as e:
#         result["Table Parsing Error"] = str(e)
#     return result

#Table

def clean_cell_detail(cell_detail: dict) -> dict:
    """
    Post-process cell_detail (parse_text_frame_debug result) for brevity:
    - Remove 'Has Text'
    - Remove Runs[*]['Run_Start_Index']
    - Remove Paragraphs key entirely if there is only 1 paragraph with no HasBullet, IndentLevel, or Alignment info
    """
    if not isinstance(cell_detail, dict):
        return cell_detail

    # 1) Remove Has Text
    cell_detail.pop("Has Text", None)

    # 2) If there is only 1 Run and Run_Start_Index == 0, remove it
    runs = cell_detail.get("Runs")
    if isinstance(runs, list) and len(runs) == 1:
        for run in runs:
            if isinstance(run, dict):
                if "Run_Start_Index" in run:
                    if run["Run_Start_Index"] == 0:
                        run.pop("Run_Start_Index", None)

    # 3) Conditionally remove Paragraphs
    paragraphs = cell_detail.get("Paragraphs")
    if isinstance(paragraphs, list) and len(paragraphs) == 1:
        p0 = paragraphs[0]
        if isinstance(p0, dict):
            has_special_style = any(key in p0 for key in ["HasBullet", "IndentLevel", "Alignment"])
        
            if not has_special_style:
                cell_detail.pop("Paragraphs", None)

    return cell_detail

def get_cell_bg_color_hex(cell_shape):
    try:
        fill = cell_shape.Fill
        if fill.Visible and fill.ForeColor:
            rgb = fill.ForeColor.RGB
            # RGB int → hex
            return "#{:06X}".format(rgb & 0xFFFFFF)
    except Exception:
        pass
    return None

def parse_table(table):
    rows = table.Rows.Count
    cols = table.Columns.Count

    result = {
        "Dimensions": {"Rows": rows, "Columns": cols},
        "Cells": {}
    }

    visited = {}  # geom_key -> (anchor_r, anchor_c)
    v_map = {3: "middle", 4: "bottom"}
    
    for r in range(1, rows + 1):
        for c in range(1, cols + 1):
            cell = table.Cell(r, c)
            shape = cell.Shape

            geom_key = (
                round(shape.Left, 2),
                round(shape.Top, 2),
                round(shape.Width, 2),
                round(shape.Height, 2),
            )

            # Handle merged sub-cells
            if geom_key in visited:
                anchor_r, anchor_c = visited[geom_key]
                key = f"{anchor_r},{anchor_c}"
                anchor_cell = result["Cells"][key]

                anchor_cell["_RowSpan"] = max(
                    anchor_cell["_RowSpan"],
                    r - anchor_r + 1
                )
                anchor_cell["_ColSpan"] = max(
                    anchor_cell["_ColSpan"],
                    c - anchor_c + 1
                )
                continue

            # ---- anchor cell (new main cell) ----
            visited[geom_key] = (r, c)

            
              # Skip 1 (top)

            # Parse and clean text frame details
            tf_detail = clean_cell_detail(
                parse_text_frame_debug(shape.TextFrame)
            )

            cell_detail = {
                "Text": tf_detail.get("Text", ""),
                "Runs": tf_detail.get("Runs", [])
            }


            # 1. Parse vertical alignment (Cell.VerticalAnchor)
            tf = shape.TextFrame
            v_anchor = safe(tf, "VerticalAnchor", None)

            # 1 (top) is the default, so skip it
            if v_anchor is not None and v_anchor != 1:
                if v_anchor in v_map:
                    cell_detail["VerticalAlign"] = v_map[v_anchor]
                    
            # 2. Parse horizontal alignment + add paragraph info if present
            if "Paragraphs" in tf_detail:
                cell_detail["Paragraphs"] = tf_detail["Paragraphs"]

            # Background color info
            bg = get_cell_bg_color_hex(shape)
            if bg:
                cell_detail["BgColor"] = bg

            cell_detail["_RowSpan"] = 1
            cell_detail["_ColSpan"] = 1

            result["Cells"][f"{r},{c}"] = cell_detail

    # ---- finalize merge meta ----
    for key, cell in list(result["Cells"].items()):
        rs = cell.pop("_RowSpan")
        cs = cell.pop("_ColSpan")

        if rs > 1 or cs > 1:
            cell["Merged"] = True
            cell["RowSpan"] = rs
            cell["ColSpan"] = cs

    return result



REVERSE_CHART_TYPES = {v: k for k, v in CHART_TYPES.items()}

def parse_chart(chart):
    """
    Parses comprehensive chart information including metadata, data points, 
    visual styles (colors), legend status, and axis configurations.
    """
    result = {}
    try:
        # 1. Basic Metadata & Type
        ct = getattr(chart, "ChartType", None)
        result["ChartType_Raw"] = REVERSE_CHART_TYPES.get(ct, f"Unknown({ct})")
        if result["ChartType_Raw"] is None:
            print(f"Unknown ChartType: {ct}")
        result["HasTitle"] = bool(getattr(chart, "HasTitle", False))
        result["Title"] = getattr(chart.ChartTitle, "Text", "") if result["HasTitle"] else None

        # 2. Legend Information
        result["HasLegend"] = bool(getattr(chart, "HasLegend", False))
        if result["HasLegend"]:
            # Legend positions: -4160(Top), -4107(Bottom), -4131(Left), -4152(Right), 2(TopRight)
            result["LegendPosition"] = getattr(chart.Legend, "Position", None)

        # 3. Series & Data Point Extraction (including Styles and Labels)
        series_data = []
        try:
            sc = chart.SeriesCollection()
            for i in range(1, sc.Count + 1):
                series = sc.Item(i)
                categories = list(series.XValues) if _com_has(series, "XValues") else []
                values = list(series.Values) if _com_has(series, "Values") else []
                
                # Extract Color (RGB)
                color_hex = None
                try:
                    # Format.Fill or Border.Color depending on chart type
                    color_int = series.Format.Fill.ForeColor.RGB
                    color_hex = _bgr_int_to_hex(color_int)
                    # color_hex = {"R": color_int & 0xFF, "G": (color_int >> 8) & 0xFF, "B": (color_int >> 16) & 0xFF}
                except Exception:
                    pass

                series_info = {
                    "SeriesName": getattr(series, "Name", f"Series {i}"),
                    "SeriesIndex": i,
                    "HasDataLabels": bool(getattr(series, "HasDataLabels", False)),
                    "Color": color_hex,
                    "Points": []
                }

                for idx in range(len(values)):
                    cat_val = categories[idx] if idx < len(categories) else f"Point {idx+1}"
                    series_info["Points"].append({
                        "Category": str(cat_val),
                        "Value": values[idx]
                    })
                
                series_data.append(series_info)
        except Exception as se:
            result["SeriesError"] = f"Could not access SeriesCollection: {str(se)}"

        result["FullData"] = series_data

        # 4. Axis Information (Value & Category Scales)
        try:
            axes = {}
            for axis_type in [1, 2]: # 1: xlCategory, 2: xlValue
                ax = chart.Axes(axis_type)
                type_name = "CategoryAxis" if axis_type == 1 else "ValueAxis"
                
                axis_detail = {
                    "HasTitle": bool(ax.HasTitle),
                    "Title": ax.AxisTitle.Text if ax.HasTitle else None,
                }
                
                # For Value Axis, extract scale information for "Axis Settings" requests
                if axis_type == 2: # xlValue
                    axis_detail["MinimumScale"] = getattr(ax, "MinimumScale", None)
                    axis_detail["MaximumScale"] = getattr(ax, "MaximumScale", None)
                    axis_detail["MajorUnit"] = getattr(ax, "MajorUnit", None)

                axes[type_name] = axis_detail
            result["Axes"] = axes
        except Exception:
            pass

    except Exception as e:
        result["ParsingError"] = str(e)
    
    return result

def parse_group_shapes(group_shape):
    """
    Recursively parse all shapes inside a Group
    """
    result = []

    try:
        group_items = group_shape.GroupItems
        count = group_items.Count

        for i in range(1, count + 1):
            sub = group_items.Item(i)

            sid = sub.Id
            name = sub.Name
            stype = sub.Type
            left = sub.Left
            top = sub.Top
            width = sub.Width
            height = sub.Height

            item_info = {
                "Shape_Id": sid,
                "Name": name,
                "Type": SHAPE_TYPE_MAP.get(stype, stype),
                "Position_Left": left,
                "Position_Top": top,
                "Size_Width": width,
                "Size_Height": height,
            }

            # ---- Text ----
            if sub.HasTextFrame: # text
                tf = sub.TextFrame
                if tf.HasText:
                    item_info["Text"] = extract_text_from_shape(sub)

            if stype == 6:  # msoGroup
                item_info["Group"] = parse_group_shapes(sub)

            elif stype in (11, 13):  # Picture / LinkedPicture
                alt_text = getattr(sub, "AlternativeText", "") or ""
                meta: dict | None = pop_picture_meta(getattr(sub, "Id", 0))
                if meta and meta.get("caption"):
                    alt_text = meta["caption"]
                elif not _is_meaningful_alt_text(alt_text):
                    meta = generate_image_caption(sub)
                    if meta and meta.get("caption"):
                        alt_text = meta["caption"]
                        try:
                            sub.AlternativeText = alt_text
                        except Exception:
                            pass
                pic_info = {"AlternativeText": alt_text}
                if meta:
                    pic_info["IsTextful"] = meta["is_textful"]
                    pic_info["TextLanguages"] = meta["text_languages"]
                    pic_info["TextSample"] = meta["text_sample"]
                item_info["Picture"] = pic_info

            elif stype == 3:  # Chart
                item_info["Chart"] = parse_chart(sub.Chart)

            elif stype == 19:  # Table
                item_info["Table"] = parse_table(sub.Table)

            result.append(item_info)

    except Exception as e:
        return {"Group Parsing Error": str(e)}

    return result



# PowerPoint automatically sets the filename as AlternativeText when inserting images,
# so if only a filename is present, treat it as "no description" and include it for captioning.
_IMAGE_FILENAME_RE = re.compile(
    r'^.+\.(png|jpe?g|gif|bmp|tiff?|svg|webp|emf|wmf|ico)$', re.IGNORECASE
)


def _is_meaningful_alt_text(text: str) -> bool:
    """Return True if alt text is a real description, not just a filename."""
    s = text.strip()
    if not s:
        return False
    if _IMAGE_FILENAME_RE.match(s):
        return False
    return True


def _export_shape_png_bytes(shape) -> bytes | None:
    """COM-only: export shape to PNG and return bytes. None on failure."""
    import tempfile
    import os as _os
    tmp_path = _os.path.join(tempfile.gettempdir(), f"caption_{id(shape)}.png")
    try:
        shape.Export(tmp_path, 2)  # 2 = ppShapeFormatPNG
        with open(tmp_path, "rb") as f:
            return f.read()
    except Exception as e:
        logger.warning(f"[_export_shape_png_bytes] Failed: {e}")
        return None
    finally:
        if _os.path.exists(tmp_path):
            try:
                _os.remove(tmp_path)
            except Exception:
                pass


# Caption schema: a single LLM call returns a one-line description plus
# enough metadata for the dispatcher to decide whether the picture's
# embedded text should be rewritten (e.g. translation tasks that need
# edit_image instead of leaving the image untouched).
_CAPTION_PROMPT = (
    "Describe this image in one concise factual sentence. "
    "Then judge whether it visibly contains text and, if so, which languages.\n"
    "Return JSON ONLY with this exact shape:\n"
    "  caption: <one short sentence>\n"
    "  is_textful: true/false — does the image contain any human-readable text?\n"
    "  text_languages: array of ISO 639-1 codes for languages visible in the image, "
    "or null when is_textful is false. Use multiple entries when more than one language is present "
    "(e.g. [\"ko\", \"en\"]).\n"
    "  text_sample: up to ~10 words copied verbatim from the image, or null when is_textful is false."
)

_CAPTION_GEMINI_SCHEMA = {
    "type": "object",
    "properties": {
        "caption": {"type": "string"},
        "is_textful": {"type": "boolean"},
        "text_languages": {
            "type": "array",
            "items": {"type": "string"},
            "nullable": True,
        },
        "text_sample": {"type": "string", "nullable": True},
    },
    "required": ["caption", "is_textful", "text_languages", "text_sample"],
}

_CAPTION_OPENAI_SCHEMA = {
    "type": "object",
    "properties": {
        "caption": {"type": "string"},
        "is_textful": {"type": "boolean"},
        # OpenAI strict mode disallows "nullable"; allow ["array","null"] / ["string","null"]
        "text_languages": {
            "anyOf": [
                {"type": "array", "items": {"type": "string"}},
                {"type": "null"},
            ]
        },
        "text_sample": {"anyOf": [{"type": "string"}, {"type": "null"}]},
    },
    "required": ["caption", "is_textful", "text_languages", "text_sample"],
    "additionalProperties": False,
}


_EMPTY_CAPTION_META: dict = {
    "caption": "",
    "is_textful": False,
    "text_languages": None,
    "text_sample": None,
}


# Prewarm cache: parser.process_batch fires caption LLM calls in a worker
# pool, then parse_active_slide_objects re-walks the slide and looks at each
# picture's AlternativeText. To avoid losing the structured fields
# (is_textful, text_languages, text_sample) on the second pass, prewarm stores
# the full meta dict here keyed by shape.Id; parsing pops it. shape.Id is
# unique within a slide and we process one slide at a time, so collisions
# across slides are not a concern.
_PICTURE_META_CACHE: dict[int, dict] = {}


def cache_picture_meta(shape_id: int, meta: dict) -> None:
    try:
        _PICTURE_META_CACHE[int(shape_id)] = meta
    except Exception:
        pass


def pop_picture_meta(shape_id: int) -> dict | None:
    try:
        return _PICTURE_META_CACHE.pop(int(shape_id), None)
    except Exception:
        return None


def _normalize_caption_meta(data: dict | None) -> dict:
    """Coerce a raw LLM response into the strict caption-meta shape.

    Guarantees: caption is str; is_textful is bool; text_languages is None or
    a list of ISO codes; text_sample is None or non-empty str. When is_textful
    is False, text_languages and text_sample are forced to None.
    """
    if not isinstance(data, dict):
        return dict(_EMPTY_CAPTION_META)
    cap = data.get("caption")
    caption = cap.strip() if isinstance(cap, str) else ""
    is_textful = bool(data.get("is_textful"))
    langs = data.get("text_languages")
    if isinstance(langs, list):
        langs_clean = [str(x).strip().lower() for x in langs if isinstance(x, str) and x.strip()]
        text_languages = langs_clean or None
    else:
        text_languages = None
    sample = data.get("text_sample")
    text_sample = sample.strip() if isinstance(sample, str) and sample.strip() else None
    if not is_textful:
        text_languages = None
        text_sample = None
    return {
        "caption": caption,
        "is_textful": is_textful,
        "text_languages": text_languages,
        "text_sample": text_sample,
    }


def _caption_from_image_bytes(image_bytes: bytes) -> dict:
    """Network-only: caption PNG bytes via LLM, return structured meta dict.

    Keys: caption (str), is_textful (bool), text_languages (list[str] | None),
    text_sample (str | None). Returns _EMPTY_CAPTION_META on failure.
    """
    import base64
    import json as _json
    from editppt.config import IMAGE_CAPTION_MODEL
    from editppt.utils.llm_client import call_llm_gemini, call_llm, GEMINI_API_KEY, set_token_log_context

    if not image_bytes:
        return dict(_EMPTY_CAPTION_META)
    set_token_log_context(component="image_caption")
    try:
        if GEMINI_API_KEY:
            raw = call_llm_gemini(
                IMAGE_CAPTION_MODEL,
                _CAPTION_PROMPT,
                image=image_bytes,
                response_schema=_CAPTION_GEMINI_SCHEMA,
            )
            if not raw or (isinstance(raw, str) and raw.startswith("[Gemini Error]")):
                return dict(_EMPTY_CAPTION_META)
            try:
                parsed = _json.loads(raw)
            except Exception:
                logger.warning(f"[_caption_from_image_bytes] non-JSON gemini reply: {str(raw)[:200]}")
                return dict(_EMPTY_CAPTION_META)
            return _normalize_caption_meta(parsed)
        else:
            b64 = base64.b64encode(image_bytes).decode("utf-8")
            messages = [{
                "role": "user",
                "content": [
                    {"type": "input_image", "image_url": f"data:image/png;base64,{b64}"},
                    {"type": "input_text", "text": _CAPTION_PROMPT},
                ],
            }]
            response = call_llm(
                model="gpt-4.1",
                messages=messages,
                text={
                    "format": {
                        "type": "json_schema",
                        "name": "image_caption_meta",
                        "schema": _CAPTION_OPENAI_SCHEMA,
                        "strict": True,
                    }
                },
            )
            if not response:
                return dict(_EMPTY_CAPTION_META)
            raw = (response.output_text or "").strip()
            try:
                parsed = _json.loads(raw)
            except Exception:
                logger.warning(f"[_caption_from_image_bytes] non-JSON gpt reply: {raw[:200]}")
                return dict(_EMPTY_CAPTION_META)
            return _normalize_caption_meta(parsed)
    except Exception as e:
        logger.warning(f"[_caption_from_image_bytes] Failed: {e}")
        return dict(_EMPTY_CAPTION_META)


def generate_image_caption(shape) -> dict:
    """Export a shape to PNG and generate a structured caption-meta dict via LLM.

    Uses Gemini (IMAGE_CAPTION_MODEL) when available, else GPT-4.1 with
    OpenAI vision format. Returns _EMPTY_CAPTION_META on any failure so
    parsing is never interrupted.
    """
    image_bytes = _export_shape_png_bytes(shape)
    if image_bytes is None:
        return dict(_EMPTY_CAPTION_META)
    return _caption_from_image_bytes(image_bytes)


def collect_pending_captions(slide) -> list[tuple[object, bytes]]:
    """COM-only: walk slide shapes (incl. groups) and export PNG bytes for
    every picture whose AlternativeText is empty/filename. Returns list of
    (shape, png_bytes). Caller decides when to caption them.
    """
    pending: list[tuple[object, bytes]] = []

    def _walk(shapes):
        n = getattr(shapes, "Count", 0)
        for i in range(1, n + 1):
            try:
                sh = shapes(i)
                stype = sh.Type
                if stype == 6:  # msoGroup
                    try:
                        _walk(sh.GroupItems)
                    except Exception:
                        pass
                    continue
                if stype in (11, 13):  # Picture / LinkedPicture
                    alt = getattr(sh, "AlternativeText", "") or ""
                    if _is_meaningful_alt_text(alt):
                        continue
                    img = _export_shape_png_bytes(sh)
                    if img:
                        pending.append((sh, img))
            except Exception as e:
                logger.warning(f"[collect_pending_captions] walk skip: {e}")
                continue

    _walk(slide.Shapes)
    return pending


def warm_captions_parallel(slide, max_workers: int = 6) -> int:
    """Pre-export and caption all pictures on `slide` whose AlternativeText
    is empty or just a filename. Sets shape.AlternativeText in place so the
    subsequent parser pass can read it back without spawning more LLM calls.

    Phase 1 (COM sequential): export PNG for every needing picture.
    Phase 2 (network parallel): caption all collected images.
    Phase 3 (COM sequential): write AlternativeText back to each shape.

    Returns the number of captions actually written.
    """
    from concurrent.futures import ThreadPoolExecutor

    pending = collect_pending_captions(slide)
    if not pending:
        return 0

    def _job(item):
        sh, img = item
        return sh, _caption_from_image_bytes(img)

    results: list[tuple[object, dict]] = []
    with ThreadPoolExecutor(max_workers=max_workers) as ex:
        for sh, meta in ex.map(_job, pending):
            if meta and isinstance(meta, dict) and meta.get("caption"):
                results.append((sh, meta))

    applied = 0
    for sh, meta in results:
        try:
            sh.AlternativeText = meta["caption"]
            applied += 1
        except Exception:
            pass
        try:
            sid = int(getattr(sh, "Id", 0))
            if sid:
                cache_picture_meta(sid, meta)
        except Exception:
            pass
    return applied


def parse_picture(picture):
    """Parse image info (returns result as dict)"""
    result = {}
    try:
        result["Type"] = getattr(picture, "Type", None)
        result["Scale"] = {
            "Width %": getattr(picture, "ScaleWidth", None),
            "Height %": getattr(picture, "ScaleHeight", None)
        }
        pf = getattr(picture, "PictureFormat", None)
        if pf:
            pic_fmt = {}
            for attr in ("Brightness", "Contrast"):
                if hasattr(pf, attr):
                    pic_fmt[attr] = getattr(pf, attr)
            crop = getattr(pf, "Crop", None)
            if crop:
                pic_fmt["Crop"] = {
                    "Left": getattr(crop, "ShapeLeft", None),
                    "Top": getattr(crop, "ShapeTop", None),
                    "Width": getattr(crop, "ShapeWidth", None),
                    "Height": getattr(crop, "ShapeHeight", None)
                }
            result["PictureFormat"] = pic_fmt

    except Exception as e:
        result["Picture Parsing Error"] = str(e)
    return result


def parse_placeholder_details(placeholder):
    """Parse placeholder details (returns result as dict)"""
    result = {}
    try:
        pf = placeholder.PlaceholderFormat
        ptype = getattr(pf, "Type", None)
        result["Placeholder Type"]  = ptype
        result["Placeholder Type Name"] = get_placeholder_type(ptype)
        # result["Placeholder Index"] = getattr(pf, "Index", None)
        if hasattr(pf, "ContainedType"):
            result["Contained Type"] = getattr(pf, "ContainedType", None)

    except Exception as e:
        result["Placeholder Parsing Error"] = str(e)
    return result


def parse_slide_notes(slide):
    """Parse slide notes (returns result as dict)"""
    result = {}
    try:
        # Check if notes page exists
        has_notes = getattr(slide, "HasNotesPage", False)
        result["Has Notes Page"] = bool(has_notes)

        if has_notes:
            notes_page = slide.NotesPage
            shapes = notes_page.Shapes
            count = getattr(shapes, "Count", 0)
            result["Notes Shapes Count"] = count

            # Collect notes text
            texts = []
            for i in range(1, count + 1):
                shape = shapes(i)
                ph = getattr(shape, "PlaceholderFormat", None)
                if ph and getattr(ph, "Type", None) == 2:
                    tf = getattr(shape, "TextFrame", None)
                    if tf and getattr(tf, "HasText", False):
                        texts.append(shape.TextFrame.TextRange.Text)

            # Set based on whether content exists
            if texts:
                result["Notes Content"] = "".join(texts)
            else:
                result["Notes Content"] = None
        else:
            result["Notes Content"] = None

    except Exception as e:
        result["Error parsing notes"] = str(e)

    return result


def parse_slide_properties(slide):
    """Parse slide properties (returns result as dict)"""
    result = {}
    try:
        # Layout is a simple enum (int) in COM, so calling
        # .Type/.Name on it raises an error on the int.
        # Store the code value only; use the CustomLayout object instead.

        layout_code = getattr(slide, "Layout", None)
        if layout_code is not None:
            result["Slide Layout Code"] = layout_code

        # CustomLayout is an object, so we can get its name/index
        custom = getattr(slide, "CustomLayout", None)
        if custom is not None:
            result["CustomLayout Name"]  = getattr(custom, "Name", None)
            result["CustomLayout Index"] = getattr(custom, "Index", None)

        # Background fill info
        bg = getattr(slide, "Background", None)
        if bg is not None:
            fill = getattr(bg, "Fill", None)
            if fill is not None:
                # Safely access fill.Type via getattr
                t = getattr(fill, "Type", None)
                fill_types = {1: "Solid", 2: "Pattern", 3: "Gradient", 4: "Texture", 5: "Picture"}
                result["Background Fill Type"] = fill_types.get(t, f"Unknown ({t})")
                # Extract actual color(s) when present so SLIDE-level background
                # changes are verifiable post-hoc.
                try:
                    if t == 1:  # Solid
                        rgb_int = fill.ForeColor.RGB
                        result["Background Color"] = _bgr_int_to_hex(rgb_int)
                    elif t == 3:  # Gradient — two stops
                        result["Background ForeColor"] = _bgr_int_to_hex(fill.ForeColor.RGB)
                        result["Background BackColor"] = _bgr_int_to_hex(fill.BackColor.RGB)
                except Exception:
                    pass

        # Transition effects
        trans = getattr(slide, "SlideShowTransition", None)
        if trans is not None:
            result["Transition Effect"]   = getattr(trans, "EntryEffect", "None")
            result["Advance Time (s)"]    = getattr(trans, "AdvanceTime", "Manual")
            result["Advance On Click"]    = bool(getattr(trans, "AdvanceOnClick", False))
            result["Advance On Time"]     = bool(getattr(trans, "AdvanceOnTime", False))

    except Exception as e:
        result["error"] = str(e)

    return result


def parse_slide_transition(slide):
    """Return transition info dict, or None if the slide uses the default
    (no) transition. Emitted as the top-level 'Slide_Transition' key only when
    non-default, to keep JSON small for unanimated decks."""
    try:
        tr = getattr(slide, "SlideShowTransition", None)
        if tr is None:
            return None
        entry = getattr(tr, "EntryEffect", 0)
        duration = getattr(tr, "Duration", 0)
        sound = ""
        try:
            sound = getattr(getattr(tr, "SoundEffect", None), "Name", "") or ""
        except Exception:
            sound = ""
        # Treat (entry=None, duration=0, no sound) as default → skip emit.
        if (entry == 0) and (duration in (0, 0.0)) and (not sound):
            return None
        advance_on_click = bool(getattr(tr, "AdvanceOnClick", False))
        advance_on_time  = bool(getattr(tr, "AdvanceOnTime", False))
        advance_time     = getattr(tr, "AdvanceTime", 0) if advance_on_time else None
        return {
            "Entry_Effect":      PP_ENTRY_EFFECT_NAME.get(entry, f"raw_{entry}"),
            "Entry_Effect_Raw":  entry,
            "Duration_Sec":      duration,
            "Advance_On_Click":  advance_on_click,
            "Advance_On_Time":   advance_on_time,
            "Advance_Time_Sec":  advance_time,
            "Sound":             sound if sound else None,
        }
    except pywintypes.com_error as e:
        logger.warning(f"[parse_slide_transition] COM error: {e}")
        return None
    except Exception as e:
        logger.warning(f"[parse_slide_transition] failed: {e}")
        return None


def parse_slide_animations(slide):
    """Return list of effects from slide.TimeLine.MainSequence, or empty list
    if the slide has no animations. Emitted as the top-level 'Slide_Animations'
    key only when non-empty."""
    effects = []
    try:
        timeline = getattr(slide, "TimeLine", None)
        if timeline is None:
            return effects
        seq = getattr(timeline, "MainSequence", None)
        if seq is None:
            return effects
        count = getattr(seq, "Count", 0)
        if count == 0:
            return effects
        for i in range(1, count + 1):
            try:
                eff = seq(i)
                # Shape may not always be present (e.g., interactive triggers).
                shape_id = None
                try:
                    shape_id = getattr(eff.Shape, "Id", None)
                except Exception:
                    shape_id = None
                effect_type     = getattr(eff, "EffectType", 0)
                timing          = getattr(eff, "Timing", None)
                duration        = getattr(timing, "Duration", 0) if timing else 0
                delay           = getattr(timing, "TriggerDelayTime", 0) if timing else 0
                trigger_type    = getattr(timing, "TriggerType", 0) if timing else 0
                is_exit         = bool(getattr(eff, "Exit", False))
                effects.append({
                    "Sequence_Index":  i,
                    "Shape_Id":        shape_id,
                    "Effect_Type":     MSO_ANIM_EFFECT_NAME.get(effect_type, f"raw_{effect_type}"),
                    "Effect_Type_Raw": effect_type,
                    "Trigger":         MSO_ANIM_TRIGGER_NAME.get(trigger_type, f"raw_{trigger_type}"),
                    "Duration_Sec":    duration,
                    "Delay_Sec":       delay,
                    "Is_Exit":         is_exit,
                })
            except Exception as e:
                logger.warning(f"[parse_slide_animations] effect {i} failed: {e}")
        return effects
    except pywintypes.com_error as e:
        logger.warning(f"[parse_slide_animations] COM error: {e}")
        return effects
    except Exception as e:
        logger.warning(f"[parse_slide_animations] failed: {e}")
        return effects


def parse_active_slide_objects(slide_num: int, prs_obj, ppt_app):
    """
    Parse Every Object Information from a Slide.
    Args:
        slide_num (int): Slide number to parse (1-based)
        prs_obj: PPTContainer.prs or win32com Presentation object
    """

    output = {}

    try:
        presentation = prs_obj

        if not presentation:
            return {"status": "No active presentation found."}

        slides = presentation.Slides
        slide_count = slides.Count

        output["Presentation_Name"] = presentation.Name
        output["Total_Slide_Number"] = slide_count
        output["Current_Slide_Number"] = slide_num

        output["Slide Width"]  = presentation.PageSetup.SlideWidth
        output["Slide Height"] = presentation.PageSetup.SlideHeight
        
        # if slide_num < 1 or slide_num > slide_count:
        #     return {"status": f"Invalid slide number (1~{slide_count})"}

        slide = slides(slide_num)

        # Pre-warm picture captions in parallel so the per-shape walk below
        # finds AlternativeText already populated and skips the LLM call.
        try:
            warm_captions_parallel(slide)
        except Exception as e:
            logger.warning(f"[parse_active_slide_objects] warm_captions_parallel failed: {e}")

        output["Slide_Properties"] = parse_slide_properties(slide)

        # Optional slide-level animation / transition keys — conditional emission
        # so JSON stays small on unanimated slides.
        trans = parse_slide_transition(slide)
        if trans:
            output["Slide_Transition"] = trans

        anims = parse_slide_animations(slide)
        if anims:
            output["Slide_Animations"] = anims

        shapes = slide.Shapes
        shape_count = shapes.Count

        output["Objects_Overview"] = f"Found {shape_count} objects"
        output["Objects_Detail"] = []

        for i in range(1, shape_count + 1):
            shape = shapes(i)
            if shape.HasTextFrame and shape.TextFrame2.HasText:
                tr2_ismath = shape.TextFrame2.TextRange
                if tr2_ismath.MathZones.Count > 0:
                    ppt_app.ActiveWindow.View.GotoSlide(slide_num)
                    shape.Select()
                    ppt_app.CommandBars.ExecuteMso("EquationLinearFormat")
        
            # ---- COM property caching ----
            sid = shape.Id
            name = shape.Name
            stype = shape.Type
            left = shape.Left
            top = shape.Top
            width = shape.Width
            height = shape.Height

            shape_info = {
                "Shape_Id": sid,
                "Name": name,
                "Type": SHAPE_TYPE_MAP.get(stype),
                "Position_Left": left,
                "Position_Top": top,
                "Size_Width": width,
                "Size_Height": height,
                "More_detail": parse_shape_details(shape),
            }

            output["Objects_Detail"].append(shape_info)
            if shape.HasTextFrame and shape.TextFrame2.HasText:
                if tr2_ismath.MathZones.Count > 0:
                    ppt_app.CommandBars.ExecuteMso("EquationProfessional")
                
        output["Slide_Notes"] = parse_slide_notes(slide)

    except pywintypes.com_error as e:
        output["Error"] = f"COM error: {e}"

    return output


def extract_text_from_shape(shape, indent_level=1):
    """
    Extract text from all shape types (TextFrame2 checked first)
    """
    result = {}
    tf_to_parse = None
    target_key = ""

    if getattr(shape, "HasTextFrame2", False):
        tf2 = shape.TextFrame2
        if safe(tf2, "HasText", False):
            tf_to_parse = tf2
            target_key = "TextFrame2"

    if not tf_to_parse and getattr(shape, "HasTextFrame", False):
        tf = shape.TextFrame
        if safe(tf, "HasText", False):
            tf_to_parse = tf
            target_key = "TextFrame"

    if tf_to_parse:
        result[target_key] = parse_text_frame_debug(tf_to_parse)

    # 3) Text inside Table
    elif getattr(shape, "Type", None) == 19 and _com_has(shape, "Table"):
        tbl = shape.Table
        rows, cols = getattr(tbl.Rows, "Count", 0), getattr(tbl.Columns, "Count", 0)
        cells = {}
        for r in range(1, rows + 1):
            for c in range(1, cols + 1):
                key = f"Cell({r},{c})"
                try:
                    txt = tbl.Cell(r, c).Shape.TextFrame.TextRange.Text
                except Exception:
                    txt = None
                cells[key] = txt
        result["TableText"] = {"Rows": rows, "Columns": cols, "Cells": cells}

    # 4) Text inside Chart
    elif getattr(shape, "Type", None) == 3 and _com_has(shape, "Chart"):
        chart = shape.Chart
        chart_info = {
            "Title": getattr(chart.ChartTitle, "Text", None) if getattr(chart, "HasTitle", False) else None,
            "Axes": {}
        }
        if _com_has(chart, "Axes"):
            for grp in (1, 2, 3):
                for typ in (1, 2):
                    try:
                        ax = chart.Axes(grp, typ)
                        if getattr(ax, "HasTitle", False):
                            chart_info["Axes"][f"{grp},{typ}"] = ax.AxisTitle.Text
                    except Exception:
                        pass
        result["ChartText"] = chart_info

    # 5) Text inside SmartArt
    elif getattr(shape, "Type", None) == 24 and _com_has(shape, "SmartArt"):
        nodes = getattr(shape.SmartArt, "AllNodes", None)
        smart = {}
        if nodes:
            for i in range(1, getattr(nodes, "Count", 0) + 1):
                try:
                    txt = nodes.Item(i).TextFrame2.TextRange.Text
                except Exception:
                    txt = None
                smart[f"Node {i}"] = txt
        result["SmartArtText"] = smart

    else:
        result["Text"] = None


    return result



def get_alignment_type(alignment_val):
    # Convert paragraph alignment value to text
    alignment_types = {
        1: "Left",
        2: "Center",
        3: "Right",
        4: "Justify",
        5: "Distributed"
    }
    return alignment_types.get(alignment_val, f"Unknown Alignment ({alignment_val})")



def parse_shape_details(shape):
    """Parse all details by shape type into a unified dict (returns Dict)"""
    result = {}
    
    # 1. Extract common properties (Z-Order, Rotation, ID, etc.)
    try:
        # Z-Order
        result["Z-Order"] = getattr(shape, "ZOrderPosition", None)
        
        # Rotation
        if getattr(shape, "Rotation", None) != 0:
            result["Rotation (°)"] = getattr(shape, "Rotation", 0)

        # Transparency
        fill = getattr(shape, "Fill", None)
        if fill and hasattr(fill, "Transparency"):
            if fill.Transparency != 0:
                result["Fill Transparency (%)"] = fill.Transparency * 100

        # Line
        line = getattr(shape, "Line", None)
        if line and getattr(line, "Visible", False):
            rgb = getattr(getattr(line, "ForeColor", None), "RGB", None) #BGR int
            result["Line"] = {
                "Width (pt)": line.Weight,
                "Color": _bgr_int_to_hex(rgb)
                # "Color": {"R": rgb & 0xFF, "G": (rgb >> 8) & 0xFF, "B": (rgb >> 16) & 0xFF}
            }
    except Exception as e:
        result["Basic_Props_Error"] = str(e)

    # 2. Parse text info
    try:
        tf = None
        if getattr(shape, "HasTextFrame", False):
            tf = shape.TextFrame
        elif getattr(shape, "HasTextFrame2", False):
            tf = shape.TextFrame2

        if tf and getattr(tf, "HasText", False):
            parsed = parse_text_frame_debug(tf)
        
            result["TextFrame"] = {
                "FullText": parsed.get("Text", ""),
                "Paragraphs": parsed.get("Paragraphs", []),
                "Runs": parsed.get("Runs", [])
            }
    except Exception as e:
        result["Text_Parsing_Error"] = str(e)

    # 3. Type-specific details (recursive and dedicated parser calls)
    t = getattr(shape, "Type", None)
    try:
        # Group (msoGroup = 6)
        if t == 6:
            # Important: the list returned here goes directly into the JSON
            result["Group"] = parse_group_shapes(shape)

        # Table (msoTable = 19)
        elif t == 19:
            result["Table"] = parse_table(shape.Table) if _com_has(shape, "Table") else "No Table Object"

        # Chart (msoChart = 3)
        elif t == 3:
            result["Chart"] = parse_chart(shape.Chart)

        # Picture (msoPicture = 13, msoLinkedPicture = 11)
        elif t in (11, 13):
            alt_text = getattr(shape, "AlternativeText", "") or ""
            meta: dict | None = pop_picture_meta(getattr(shape, "Id", 0))
            if meta and meta.get("caption"):
                alt_text = meta["caption"]
            elif not _is_meaningful_alt_text(alt_text):
                meta = generate_image_caption(shape)
                if meta and meta.get("caption"):
                    alt_text = meta["caption"]
                    try:
                        shape.AlternativeText = alt_text
                    except Exception:
                        pass
            pic_info = {"AlternativeText": alt_text, "Name": shape.Name}
            if meta:
                pic_info["IsTextful"] = meta["is_textful"]
                pic_info["TextLanguages"] = meta["text_languages"]
                pic_info["TextSample"] = meta["text_sample"]
            result["Picture"] = pic_info

        # Placeholder (msoPlaceholder = 14)
        elif t == 14:
            result["Placeholder"] = parse_placeholder_details(shape)

        # SmartArt (msoSmartArt = 24)
        elif t == 24:
            result["SmartArt_Nodes_Count"] = getattr(shape.SmartArt.AllNodes, "Count", 0)

    except Exception as e:
        result["Type_Specific_Error"] = str(e)

    return result


def _max_font_size_from_runs(runs):
    sizes = [r.get("Font", {}).get("Size") for r in runs or []]
    sizes = [s for s in sizes if isinstance(s, (int, float))]
    return max(sizes) if sizes else None


def _normalize_text(s):
    """Flatten \\r/\\n into spaces. No truncation."""
    if not s:
        return ""
    return s.replace("\r", " ").replace("\n", " ")


def slim_for_vision(parsed: dict) -> dict:
    """Strip parsed_contents to only what the vision validator needs.

    Drops per-run fonts, paragraphs, z-order, fill, placeholder meta and other
    fields that the MLLM can read off the image. Keeps id, type, name,
    position/size, plain text (\\r/\\n flattened), max font size, and (for
    tables) rows/cols + flat cell text list.
    """
    def _shape(obj):
        more = obj.get("More_detail", {}) or {}
        out = {
            "id": obj.get("Shape_Id"),
            "type": obj.get("Type"),
            "name": obj.get("Name"),
            "left": round(obj.get("Position_Left", 0) or 0, 2),
            "top": round(obj.get("Position_Top", 0) or 0, 2),
            "width": round(obj.get("Size_Width", 0) or 0, 2),
            "height": round(obj.get("Size_Height", 0) or 0, 2),
        }
        tf = more.get("TextFrame")
        if tf:
            out["text"] = _normalize_text(tf.get("FullText", ""))
            mfs = _max_font_size_from_runs(tf.get("Runs"))
            if mfs is not None:
                out["max_font_size"] = mfs
        tbl = more.get("Table")
        if tbl:
            dims = tbl.get("Dimensions", {})
            cells = tbl.get("Cells", {}) or {}
            cell_texts = []
            mfs_all = []
            for v in cells.values():
                cell_texts.append(_normalize_text(v.get("Text", "")))
                mfs_all.append(_max_font_size_from_runs(v.get("Runs")))
            out["table"] = {
                "rows": dims.get("Rows"),
                "cols": dims.get("Columns"),
                "cells": cell_texts,
            }
            mfs_clean = [s for s in mfs_all if s is not None]
            if mfs_clean:
                out["max_font_size"] = max(mfs_clean)
        return out

    return {
        "slide_w": parsed.get("Slide Width"),
        "slide_h": parsed.get("Slide Height"),
        "current_slide": parsed.get("Current_Slide_Number"),
        "objects": [_shape(o) for o in parsed.get("Objects_Detail", [])],
    }


# Keys in tool arguments that point to a shape id we should keep "full".
_SHAPE_ID_ARG_KEYS = (
    "shape_id", "target_shape_id", "source_shape_id",
    "shape_ids", "target_shape_ids",
)


def _impacted_shape_ids(used_tools) -> set:
    impacted = set()
    for call in used_tools or []:
        args = call.get("arguments", {}) or {}
        for k in _SHAPE_ID_ARG_KEYS:
            v = args.get(k)
            if isinstance(v, list):
                for item in v:
                    if isinstance(item, int):
                        impacted.add(item)
            elif isinstance(v, int):
                impacted.add(v)
    return impacted


def scope_for_used_tools(slim: dict, used_tools) -> dict:
    """Compress non-impacted shapes to id+name+position. Impacted shapes stay full.

    Falls back to no scoping when used_tools yields no impacted IDs (e.g. tools
    without shape_id like add_slide); in that case the slim payload is returned
    unchanged.
    """
    impacted = _impacted_shape_ids(used_tools)
    if not impacted:
        return slim

    scoped_objects = []
    for o in slim.get("objects", []):
        if o.get("id") in impacted:
            scoped_objects.append(o)
        else:
            scoped_objects.append({
                "id": o.get("id"),
                "type": o.get("type"),
                "name": o.get("name"),
                "left": o.get("left"),
                "top": o.get("top"),
                "width": o.get("width"),
                "height": o.get("height"),
            })
    return {**slim, "objects": scoped_objects}


# Tool names that meaningfully read or modify z-order; when any of these is
# used we must keep Z-Order in the trimmed payload.
_Z_ORDER_TOOL_NAMES = {
    "set_shape_zorder", "bring_to_front", "send_to_back",
    "bring_forward", "send_backward",
}


def _uses_z_order(used_tools) -> bool:
    for call in used_tools or []:
        if call.get("name") in _Z_ORDER_TOOL_NAMES:
            return True
    return False


def _strip_run_indices(runs):
    """Drop Run_Start_Index from every run; preserve other run fields."""
    if not runs:
        return runs
    out = []
    for r in runs:
        rr = {k: v for k, v in r.items() if k != "Run_Start_Index"}
        out.append(rr)
    return out


def _strip_paragraph_meta(paragraphs):
    """Drop redundant LineSpacingRule (LineSpacing remains)."""
    if not paragraphs:
        return paragraphs
    out = []
    for p in paragraphs:
        pp = {k: v for k, v in p.items() if k != "LineSpacingRule"}
        out.append(pp)
    return out


_SLIDE_PROPS_DROP = {
    "Transition Effect",
    "Advance Time (s)",
    "Advance On Click",
    "Advance On Time",
    "Slide Layout Code",
    "CustomLayout Name",
    "CustomLayout Index",
    "Background Fill Type",
}


def trim_meta_for_text(parsed: dict, used_tools=None) -> dict:
    """Remove fields the text validator never reads.

    Keeps every Font, Position, Text, and Alignment field intact so that style
    edits remain verifiable. Only drops:
      - Slide_Properties keys for transition / layout / background
      - Basic_Props_Error (debug noise)
      - Run_Start_Index
      - LineSpacingRule (LineSpacing kept)
      - per-cell BgColor in tables
      - Z-Order — UNLESS used_tools contains a z-order tool, then kept
    """
    keep_z = _uses_z_order(used_tools)

    # Slide-level
    out = {k: v for k, v in parsed.items() if k != "Slide_Properties"}
    sp = parsed.get("Slide_Properties") or {}
    if sp:
        sp_trim = {k: v for k, v in sp.items() if k not in _SLIDE_PROPS_DROP}
        if sp_trim:
            out["Slide_Properties"] = sp_trim

    # Per-shape
    new_objs = []
    for obj in parsed.get("Objects_Detail", []):
        new_obj = dict(obj)
        more = dict(new_obj.get("More_detail", {}) or {})

        more.pop("Basic_Props_Error", None)
        if not keep_z:
            more.pop("Z-Order", None)

        tf = more.get("TextFrame")
        if isinstance(tf, dict):
            tf = dict(tf)
            tf["Runs"] = _strip_run_indices(tf.get("Runs"))
            tf["Paragraphs"] = _strip_paragraph_meta(tf.get("Paragraphs"))
            more["TextFrame"] = tf

        tbl = more.get("Table")
        if isinstance(tbl, dict):
            tbl = dict(tbl)
            cells = dict(tbl.get("Cells") or {})
            new_cells = {}
            for k, cell in cells.items():
                if not isinstance(cell, dict):
                    new_cells[k] = cell
                    continue
                cc = {kk: vv for kk, vv in cell.items() if kk != "BgColor"}
                cc["Runs"] = _strip_run_indices(cc.get("Runs"))
                cc["Paragraphs"] = _strip_paragraph_meta(cc.get("Paragraphs"))
                new_cells[k] = cc
            tbl["Cells"] = new_cells
            more["Table"] = tbl

        new_obj["More_detail"] = more
        new_objs.append(new_obj)
    out["Objects_Detail"] = new_objs
    return out


def diff_for_text(old: dict, new: dict) -> dict:
    """Build a compact diff for the text validator.

    Output:
      {
        "slide_dims": {"w": ..., "h": ...},
        "current_slide": int,
        "changed_shapes":   [{"id": int, "before": <full shape>, "after": <full shape>}],
        "added_shapes":     [{"id": int, "after":  <full shape>}],
        "removed_shapes":   [{"id": int, "before": <full shape>}],
        "unchanged_shape_ids": [int, ...]
      }

    Matching is by Shape_Id. Equality uses a deep dict compare so that any
    field difference (text, font, position, etc.) classifies the shape as
    changed. Tables are kept whole on the changed side; nested-cell diffs are
    intentionally not done here to preserve correctness.
    """
    def _index(parsed):
        return {o.get("Shape_Id"): o for o in parsed.get("Objects_Detail", []) if o.get("Shape_Id") is not None}

    old_idx = _index(old)
    new_idx = _index(new)
    old_ids = set(old_idx)
    new_ids = set(new_idx)

    changed = []
    unchanged = []
    for sid in sorted(old_ids & new_ids):
        if old_idx[sid] != new_idx[sid]:
            changed.append({"id": sid, "before": old_idx[sid], "after": new_idx[sid]})
        else:
            unchanged.append(sid)

    added = [{"id": sid, "after": new_idx[sid]} for sid in sorted(new_ids - old_ids)]
    removed = [{"id": sid, "before": old_idx[sid]} for sid in sorted(old_ids - new_ids)]

    return {
        "slide_dims": {
            "w": new.get("Slide Width") or old.get("Slide Width"),
            "h": new.get("Slide Height") or old.get("Slide Height"),
        },
        "current_slide": new.get("Current_Slide_Number") or old.get("Current_Slide_Number"),
        "changed_shapes": changed,
        "added_shapes": added,
        "removed_shapes": removed,
        "unchanged_shape_ids": unchanged,
    }

    return result