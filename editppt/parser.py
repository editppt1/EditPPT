from editppt.utils.utils import (
    parse_active_slide_objects,
    trim_meta_for_text,
    diff_for_text,
    collect_pending_captions,
    _caption_from_image_bytes,
)
import json
import re
from concurrent.futures import ThreadPoolExecutor
from copy import deepcopy

from loguru import logger

from editppt.utils.logger_manual import log_path
from editppt.utils.llm_client import call_llm, set_token_log_context
from editppt.prompts import *


TEXT_VALIDATOR_SCHEMA = {
    "type": "object",
    "properties": {
        "valid": {"type": "boolean"},
        "strategy": {
            "type": ["string", "null"],
            "enum": ["INCREMENTAL", "ROLLBACK", None],
        },
        "reason": {"type": "string"},
    },
    "required": ["valid", "strategy", "reason"],
    "additionalProperties": False,
}


def _enrich_with_placeholder_index(slide_json: dict) -> dict:
    """Add a top-level ``Placeholder_Index`` mapping role -> [shape_id, ...].

    The role lookup is buried under ``Objects_Detail[i].More_detail.Placeholder``
    (and inside ``More_detail.Group`` for nested shapes). Surfacing it at the
    top lets the dispatcher/specialist LLM pick a target shape without walking
    the whole slide JSON. No-op if ``slide_json`` is empty or already enriched.
    """
    if not isinstance(slide_json, dict) or "Objects_Detail" not in slide_json:
        return slide_json

    index: dict[str, list[int]] = {}

    def _walk(objects):
        for obj in objects or []:
            md = obj.get("More_detail", {}) or {}
            ph = md.get("Placeholder") or {}
            role = ph.get("Placeholder Type Name")
            sid = obj.get("Shape_Id")
            if role and sid is not None:
                index.setdefault(role, []).append(sid)
            # Recurse into groups so group-internal placeholders are addressable
            inner = md.get("Group") or []
            if inner:
                _walk(inner)

    _walk(slide_json.get("Objects_Detail", []))
    slide_json["Placeholder_Index"] = index
    return slide_json


# ── Visual_Role_Index ─────────────────────────────────────────────────────
# Decks that don't bind shapes to real placeholder slots (very common in
# Korean business/research decks) have ``Placeholder_Index`` either sparse or
# all "Body". Visual_Role_Index is a heuristic supplement: combines
#   S1: real placeholder type (deterministic)
#   S2: shape Name regex covering major languages
#   S3: shape positioned in the upper-25%-band with font size ≥ 90% of the
#       slide's largest run font (fallback for title only)
# Emits {role: [shape_id, ...]} with no confidence labels; ordering reflects
# the signal precedence above (highest-confidence shape_ids come first).

_TITLE_NAME_RE = re.compile(
    r"title|heading|header"                                              # English
    r"|제목|표제"                                                          # Korean
    r"|タイトル|表題|見出し"                                                # Japanese
    r"|标题|題目|標題"                                                      # Chinese (S/T)
    r"|título|titre|titel|titolo|tytuł"                                  # ES/FR/DE+NL+SV/IT/PL
    r"|başlık"                                                             # Turkish
    r"|заголовок|название"                                                # Russian
    r"|عنوان"                                                              # Arabic
    r"|tiêu đề|judul|शीर्षक|ชื่อเรื่อง"                                 # VI/ID-MS/HI/TH
    , re.IGNORECASE
)
_SUBTITLE_NAME_RE = re.compile(
    r"sub[\- ]?title|subhead"                                            # English
    r"|부제목|부제|소제목"                                                  # Korean
    r"|サブタイトル|副題"                                                   # Japanese
    r"|副标题|副標題"                                                       # Chinese (S/T)
    r"|subtítulo|sous[\- ]titre|untertitel|sottotitolo"                  # ES/FR/DE/IT
    r"|podtytuł|alt başlık|ondertitel|undertitel"                        # PL/TR/NL/SV
    r"|подзаголовок"                                                      # Russian
    r"|عنوان فرعي"                                                        # Arabic
    r"|उपशीर्षक"                                                          # Hindi
    , re.IGNORECASE
)


def _flatten_shapes(objects):
    """Yield each shape dict, recursing into Group children."""
    for obj in objects or []:
        if not isinstance(obj, dict):
            continue
        yield obj
        for inner in (obj.get("More_detail") or {}).get("Group") or []:
            yield from _flatten_shapes([inner])


def _enrich_with_visual_role_index(slide_json: dict) -> dict:
    """Add a top-level ``Visual_Role_Index`` mapping inferred visual role to
    shape_ids. Combines placeholder type, shape Name regex (multi-language),
    and a top-of-slide + large-font fallback for the title role.

    Designed to be additive — does not modify ``Placeholder_Index``. Safe on
    empty / malformed input.
    """
    if not isinstance(slide_json, dict) or "Objects_Detail" not in slide_json:
        return slide_json

    roles: dict[str, list[int]] = {"title": [], "subtitle": [], "body": [], "footer": []}

    def _add(role: str, sid):
        """Add shape to a role bucket if not already present in THAT bucket.
        A shape can legitimately appear in multiple buckets when signals
        disagree (e.g., a Body placeholder that visually serves as title)."""
        if sid is None:
            return
        if sid not in roles[role]:
            roles[role].append(sid)

    def _promote_to_title(sid):
        """Visual title trumps placeholder type: remove from body/subtitle
        before adding to title so the LLM sees a single, unambiguous role."""
        if sid is None:
            return
        for r in ("body", "subtitle", "footer"):
            if sid in roles[r]:
                roles[r].remove(sid)
        if sid not in roles["title"]:
            roles["title"].append(sid)

    # S1: from Placeholder_Index (deterministic; assumes _enrich_with_placeholder_index ran first)
    ph = slide_json.get("Placeholder_Index") or {}
    s1_map = [
        ("title",    ["Title", "CenterTitle", "VerticalTitle", "VerticalTitle2"]),
        ("subtitle", ["SubTitle"]),
        ("body",     ["Body", "VerticalBody", "VerticalBody2"]),
        ("footer",   ["Footer", "Header", "Slide Number", "Date"]),
    ]
    for role, ph_keys in s1_map:
        for k in ph_keys:
            for sid in ph.get(k, []) or []:
                _add(role, sid)

    # S2: Shape Name regex (multi-language).
    # Subtitle MUST be tested before title — "subtitle" / "부제목" / "副标题"
    # contain the title substring ("title" / "제목" / "标题") and would
    # otherwise be misclassified as title.
    for obj in _flatten_shapes(slide_json.get("Objects_Detail", [])):
        sid = obj.get("Shape_Id")
        name = obj.get("Name") or ""
        if _SUBTITLE_NAME_RE.search(name):
            _add("subtitle", sid)
        elif _TITLE_NAME_RE.search(name):
            _add("title", sid)

    # S3: top-band + large-font fallback (title only; only when S1+S2 found none).
    # Decks where every text shape is a Body placeholder (very common) require
    # this fallback. Shapes inferred as visual title are promoted away from
    # their body classification so the LLM sees them as title only.
    if not roles["title"]:
        slide_h = slide_json.get("Slide Height") or 540
        cands: list[tuple] = []  # (sid, max_font, top)
        for obj in _flatten_shapes(slide_json.get("Objects_Detail", [])):
            tf = (obj.get("More_detail") or {}).get("TextFrame") or {}
            # `Has Text` is not consistently populated by the parser, so
            # detect "text-bearing" by looking at Runs directly.
            sizes = [
                r.get("Font", {}).get("Size")
                for r in (tf.get("Runs") or [])
                if r.get("Font", {}).get("Size") and (r.get("Text") or "").strip()
            ]
            if not sizes:
                continue
            cands.append((obj.get("Shape_Id"), max(sizes), obj.get("Position_Top", 1e9)))
        if cands:
            max_size = max(f for _, f, _ in cands)
            for sid, fsize, top in cands:
                if top < slide_h * 0.25 and fsize >= max_size * 0.9:
                    _promote_to_title(sid)

    slide_json["Visual_Role_Index"] = roles
    return slide_json


def _enrich_slide_json(slide_json: dict) -> dict:
    """Apply both enrichment passes in dependency order."""
    slide_json = _enrich_with_placeholder_index(slide_json)
    slide_json = _enrich_with_visual_role_index(slide_json)
    return slide_json


class Parser:
    def __init__(self, container: object, total_slides: int):
        """
        Args:
            container (object): PPTContainer instance containing the prs.
            total_slides (int): Total number of slides.
        """
        self.database = {}
        self.edit_history = {}
        self.container = container  
        self.total_slides = total_slides

        # for page_num in range(1, min(10, total_slides + 1)):
        #     self.database[page_num] = parse_active_slide_objects(page_num, self.container.prs)
        #     print(f"Parsed slide {page_num}/{total_slides}")


    def process(self, page_number: int, force: bool = False):
        if not force and page_number in self.database and self.database[page_number]:
            return self.database[page_number]

        print(f'Parsing Page {page_number}...')
        print('='*40)
        self.database[page_number] = _enrich_slide_json(
            parse_active_slide_objects(page_number, self.container.prs, self.container.ppt_app)
        )
        with open(log_path("parser_Database.json"), "w", encoding="utf-8") as f:
            json.dump(self.database, f, ensure_ascii=False, indent=4)

        return self.database[page_number]

    def process_batch(self, slide_indices: list[int], max_workers: int = 8) -> dict[int, dict]:
        """Parse multiple slides with caption pipelining across slides.

        Phase A (COM, sequential per slide): export PNG for every needs-caption
          picture and immediately submit its caption job to a shared executor.
          COM moves on to the next slide without waiting for captions.
        Phase B (sync): wait for all caption Futures, set AlternativeText.
        Phase C (COM, sequential): parse each slide into JSON; the per-shape
          walk now finds AlternativeText populated and skips inline captioning.

        Returns dict[slide_index -> slide_json], also populated into
        self.database. Slides already cached are skipped.
        """
        to_parse = [
            idx for idx in slide_indices
            if not (idx in self.database and self.database[idx])
        ]
        if not to_parse:
            return {idx: self.database[idx] for idx in slide_indices if idx in self.database}

        executor = ThreadPoolExecutor(max_workers=max_workers, thread_name_prefix="caption")
        pending_futures: list[tuple[int, object, object]] = []  # (slide_idx, shape, future)

        # Phase A: sequential COM export per slide, async caption submit.
        for idx in to_parse:
            try:
                slide = self.container.prs.Slides(idx)
            except Exception as e:
                logger.warning(f"process_batch: slide {idx} access failed: {e}")
                continue
            pending = collect_pending_captions(slide)
            for shape, img_bytes in pending:
                fut = executor.submit(_caption_from_image_bytes, img_bytes)
                pending_futures.append((idx, shape, fut))

        # Phase B: collect caption-meta dicts, persist caption back to
        # PowerPoint, and cache the structured fields keyed by shape.Id so
        # parse_active_slide_objects can pick them up on the second pass.
        from editppt.utils.utils import cache_picture_meta as _cache_picture_meta
        for idx, shape, fut in pending_futures:
            try:
                meta = fut.result()
            except Exception as e:
                logger.warning(f"process_batch: caption future failed (slide {idx}): {e}")
                continue
            if not meta:
                continue
            caption = meta.get("caption") if isinstance(meta, dict) else None
            if caption:
                try:
                    shape.AlternativeText = caption
                except Exception:
                    pass
            try:
                sid = int(getattr(shape, "Id", 0))
                if sid:
                    _cache_picture_meta(sid, meta)
            except Exception:
                pass
        executor.shutdown(wait=True)

        # Phase C: parse JSON. parse_active_slide_objects internally calls
        # warm_captions_parallel as a safety net, but with AlternativeText
        # already populated it's a no-op for the prewarmed pictures.
        for idx in to_parse:
            print(f"Parsing Page {idx}...")
            print("=" * 40)
            try:
                self.database[idx] = _enrich_slide_json(
                    parse_active_slide_objects(
                        idx, self.container.prs, self.container.ppt_app
                    )
                )
            except Exception as e:
                logger.warning(f"process_batch: parse failed (slide {idx}): {e}")
                self.database[idx] = None

        with open(log_path("parser_Database.json"), "w", encoding="utf-8") as f:
            json.dump(self.database, f, ensure_ascii=False, indent=4)

        return {idx: self.database[idx] for idx in slide_indices if idx in self.database}

    def update_after_edit(self,
                        text_validation: bool,
                        model: str,
                        page_number: int,
                        description: str,
                        action: str,
                        detailed_contents: str,
                        used_tools: list):
        """
        LLM, VLM Validates revision completion.
        Returns (valid, reason, strategy, new_parse).
        """

        old_parse = self.database.get(page_number, None)
        if old_parse is None:
            raise RuntimeError("Slide not parsed by parser.process()")
        new_parse = _enrich_slide_json(
            parse_active_slide_objects(page_number, self.container.prs, self.container.ppt_app)
        )

        # Record pre-edit data (append mode)
        with open(log_path(f"oldparse_{page_number}.txt"), "a", encoding="utf-8") as f:
            f.write(f"\n{'='*50}\n")
            f.write(json.dumps(old_parse, ensure_ascii=False, indent=4))
            f.write("\n")

        # Record post-edit data (append mode)
        with open(log_path(f"newparse_{page_number}.txt"), "a", encoding="utf-8") as f:
            f.write(f"\n{'='*50}\n")
            f.write(json.dumps(new_parse, ensure_ascii=False, indent=4))
            f.write("\n")

        # Text Validation
        if text_validation:
            old_trim = trim_meta_for_text(old_parse, used_tools)
            new_trim = trim_meta_for_text(new_parse, used_tools)
            diff_payload = diff_for_text(old_trim, new_trim)
            # D4 gate (same env var as D5): EDITPPT_DISABLE_D45=1 → omit indices.
            import os as _os
            _d4_off = _os.environ.get("EDITPPT_DISABLE_D45") == "1"
            _ph_idx = None if _d4_off else (new_parse.get("Placeholder_Index") if isinstance(new_parse, dict) else None)
            _vr_idx = None if _d4_off else (new_parse.get("Visual_Role_Index") if isinstance(new_parse, dict) else None)
            messages = [
                {"role": "system",
                "content": create_text_validator_agent_system_prompt(
                    page_number, description, action, detailed_contents)},
                {"role": "user",
                "content": create_text_validator_agent_user_prompt(
                    diff_payload, used_tools,
                    placeholder_index=_ph_idx,
                    visual_role_index=_vr_idx,
                )}
            ]
            set_token_log_context(component="text_validator")
            response = call_llm(
                model=model,
                messages=messages,
                text={
                    "format": {
                        "type": "json_schema",
                        "name": "text_validation",
                        "schema": TEXT_VALIDATOR_SCHEMA,
                        "strict": True,
                    }
                },
                prompt_cache_key="text_validator",
            )
            if response is None:
                return False, "LLM call in Validation Failed.", "rollback", new_parse
            response_text = (response.output_text or "").strip()

            try:
                data = json.loads(response_text)
            except json.JSONDecodeError as e:
                return False, f"Validator JSON decode error: {e}", "rollback", new_parse

            valid = bool(data.get("valid"))
            reason = data.get("reason", "") or ""
            strategy_raw = data.get("strategy")
            if valid:
                return True, reason, None, new_parse
            strategy = (strategy_raw or "ROLLBACK").lower()
            if strategy not in ("incremental", "rollback"):
                strategy = "rollback"
            return False, reason, strategy, new_parse

        else:
            self.edit_history.setdefault(page_number, []).append(deepcopy(old_parse))
            self.database[page_number] = new_parse
            return True, None, None, new_parse


       # Removed. Command may be mis-ordered to edit an already accomplish task, so the result could be identical to the original.
        # #validation 1
        # if new_parse == old_parse:
        #     if used_tools:
        #         return False, "Tool used, but no differences. Check the target slide number or shape id carefully."
        #     else:
        #         return False, "Unknown Error."
        # else:
        #     if text_validation:
        #         ...
