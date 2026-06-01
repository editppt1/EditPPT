import json
import re
import contextvars
from concurrent.futures import ThreadPoolExecutor

from editppt.utils.llm_client import call_llm, APIKeyError, set_token_log_context
from editppt.utils.logger_manual import init_logger
from editppt.agents.registry import get_all_agent_descriptions, AGENT_REGISTRY
from editppt.prompts import (
    create_dispatcher_system_prompt,
    create_dispatcher_user_prompt,
)

logger = init_logger()

FALLBACK_AGENT = "text_style"


class DispatcherAgent:
    def __init__(self, model: str):
        self.model = model

    def dispatch(self, task: dict, slide_objects: list[dict] | None = None) -> list[dict]:
        """Dispatch a task to one or more specialist agents.

        Args:
            task: Task dict with description, action, target, contents, page_number.
            slide_objects: Optional list of shape dicts from parser (Objects_Detail).
                           When provided, the dispatcher groups shapes by agent type.

        Returns:
            List of sub-task dicts, each with:
                - agent_type: str
                - description: str
                - shape_ids: list[int] (optional, absent for slide-level tasks)
        """
        description = task.get("description", "")
        action = task.get("action", "")
        target = task.get("target", "")
        contents = task.get("contents", "")

        # If no slide objects provided (e.g. slide-level tasks), fall back to single-agent dispatch
        if not slide_objects:
            return self._dispatch_single(task)

        # Build shape summary for the LLM. For Pictures we additionally surface
        # AlternativeText (an auto-generated one-line caption stored by the
        # parser) so the LLM can reason about what the picture *depicts* —
        # critical when an image functionally stands in for a table, chart,
        # or text block.
        shape_summary = []
        for obj in slide_objects:
            entry = {
                "Shape_Id": obj.get("Shape_Id"),
                "Name": obj.get("Name", ""),
                "Type": obj.get("Type", ""),
            }
            if entry["Type"] == "Picture":
                pic = ((obj.get("More_detail") or {}).get("Picture") or {})
                alt = pic.get("AlternativeText") or ""
                if alt:
                    entry["AlternativeText"] = alt
                # Caption-meta fields surfaced by parser. The dispatcher uses
                # these to decide whether a picture needs edit_image-class
                # handling (text-on-image targeted by the user's request).
                if "IsTextful" in pic:
                    entry["IsTextful"] = pic["IsTextful"]
                if pic.get("TextLanguages"):
                    entry["TextLanguages"] = pic["TextLanguages"]
                if pic.get("TextSample"):
                    entry["TextSample"] = pic["TextSample"]
            shape_summary.append(entry)

        system_prompt = create_dispatcher_system_prompt(
            agent_descriptions=get_all_agent_descriptions(),
        )
        user_prompt = create_dispatcher_user_prompt(
            description=description,
            action=action,
            target=target,
            contents=contents,
            shape_summary=shape_summary,
        )

        messages = [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_prompt},
        ]

        try:
            set_token_log_context(component="dispatcher")
            response = call_llm(model=self.model, messages=messages, prompt_cache_key="dispatcher:multi")
            raw = (response.output_text or "").strip()

            sub_tasks = self._parse_multi_response(raw)
            if sub_tasks:
                logger.info(f"Dispatcher multi-route: {[(s['agent_type'], s.get('shape_ids', [])) for s in sub_tasks]}")
                return sub_tasks

            logger.warning(f"Dispatcher multi-route parse failed, falling back to single dispatch")
            return self._dispatch_single(task)

        except APIKeyError:
            raise
        except Exception as e:
            logger.error(f"Dispatcher multi-agent LLM call failed: {e}. Falling back to single dispatch.")
            return self._dispatch_single(task)

    def prefetch_dispatch_decisions(self, tasks: list[dict], parser, max_workers: int = 4) -> dict[int, list[dict]]:
        """Pre-compute dispatcher decisions for a list of tasks in parallel.

        Parses unique slides serially (COM is single-threaded), then issues
        the dispatcher LLM calls concurrently. Results may go stale if a
        prior task modifies a slide that a later task targets — callers are
        expected to re-dispatch in that case (track touched slides).

        Returns:
            dict mapping task index → sub_tasks list. Tasks whose slide does
            not yet exist or whose dispatch errored are omitted; callers
            should fall back to per-task dispatch for those.
        """
        try:
            total_slides = parser.container.prs.Slides.Count
        except Exception:
            return {}

        # Step 1: parse unique slides as a single batch. process_batch pipelines
        # caption LLM calls across slides — COM exports for slide N submit
        # captions to a shared executor and immediately move on to N+1, instead
        # of waiting per-slide. JSON build happens after all captions resolve.
        unique_slides = []
        seen = set()
        for task in tasks:
            slide_idx = task.get("page_number")
            if not slide_idx or slide_idx > total_slides or slide_idx in seen:
                continue
            seen.add(slide_idx)
            unique_slides.append(slide_idx)

        parsed_by_slide: dict[int, dict | None] = {}
        if unique_slides:
            try:
                set_token_log_context(slide_index=None, scope="parser_batch")
                parsed_by_slide = parser.process_batch(unique_slides)
            except Exception as e:
                logger.warning(f"prefetch process_batch failed: {e}")
                # Fall back to per-slide parsing for any slides we couldn't batch.
                for idx in unique_slides:
                    if idx in parsed_by_slide and parsed_by_slide[idx]:
                        continue
                    set_token_log_context(slide_index=idx)
                    try:
                        parsed_by_slide[idx] = parser.process(idx)
                    except Exception as ee:
                        logger.warning(f"prefetch parse failed for slide {idx}: {ee}")
                        parsed_by_slide[idx] = None
            finally:
                set_token_log_context(slide_index=None, scope=None)

        # Step 2: build per-task jobs.
        jobs: list[tuple[int, int, dict, list]] = []
        for i, task in enumerate(tasks):
            slide_idx = task.get("page_number")
            if not slide_idx or slide_idx > total_slides:
                continue
            slide_contents = parsed_by_slide.get(slide_idx)
            objects_detail = (slide_contents or {}).get("Objects_Detail", []) if slide_contents else []
            jobs.append((i, slide_idx, task, objects_detail))

        if not jobs:
            return {}

        # Step 3: parallel LLM dispatch. Each worker runs in a copied context
        # so set_token_log_context(slide_index=...) attributes the LLM call to
        # the right slide in token_usage.jsonl without races between siblings.
        def _job(item):
            i, slide_idx, task, objects_detail = item
            set_token_log_context(slide_index=slide_idx, scope="dispatch_prefetch")
            try:
                return i, self.dispatch(task, objects_detail)
            except APIKeyError:
                raise
            except Exception as e:
                logger.warning(f"prefetch dispatch failed for task #{i}: {e}")
                return i, None

        results: dict[int, list[dict]] = {}
        with ThreadPoolExecutor(max_workers=max_workers) as ex:
            futures = []
            for item in jobs:
                ctx = contextvars.copy_context()
                futures.append(ex.submit(ctx.run, _job, item))
            for fut in futures:
                i, sub_tasks = fut.result()
                if sub_tasks is not None:
                    results[i] = sub_tasks
        return results

    def _dispatch_single(self, task: dict) -> list[dict]:
        """Original single-agent dispatch logic."""
        agent_descriptions = get_all_agent_descriptions()

        system_prompt = f"""You are a task dispatcher for a PowerPoint editing system.
Given a task description, determine which specialist agent should handle it.

Available agents:
{agent_descriptions}

Rules:
- Return ONLY the agent_type string (e.g. "text_style", "table", "chart", "shape_layout", "slide").
- Choose the single best-matching agent for the task.
- If the task involves text content or font styling, choose "text_style".
- If the task involves table cells or table formatting, choose "table".
- If the task involves chart data, axes, or chart visuals, choose "chart".
- If the task involves moving, resizing, creating, deleting, or duplicating shapes, choose "shape_layout".
- If the task involves adding, deleting, or duplicating entire slides, choose "slide".
- When in doubt, default to "text_style".
"""

        user_prompt = f"""Task to dispatch:
- Description: {task.get("description", "")}
- Action: {task.get("action", "")}
- Target: {task.get("target", "")}
- Contents: {task.get("contents", "")}

Which agent_type should handle this task? Reply with ONLY the agent_type string."""

        messages = [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_prompt},
        ]

        try:
            set_token_log_context(component="dispatcher")
            response = call_llm(model=self.model, messages=messages, prompt_cache_key="dispatcher:single")
            raw = (response.output_text or "").strip().strip('"').strip("'").lower()

            for agent_type in AGENT_REGISTRY:
                if agent_type in raw:
                    return [{"agent_type": agent_type, "description": task.get("description", "")}]

            logger.warning(f"Dispatcher could not parse agent_type from: '{raw}'. Falling back to '{FALLBACK_AGENT}'")
        except APIKeyError:
            raise
        except Exception as e:
            logger.error(f"Dispatcher LLM call failed: {e}. Falling back to '{FALLBACK_AGENT}'")

        return [{"agent_type": FALLBACK_AGENT, "description": task.get("description", "")}]

    def _parse_multi_response(self, raw: str) -> list[dict] | None:
        """Parse LLM response as JSON array of sub-tasks. Returns None on failure."""
        # Strip markdown code fences if present
        cleaned = raw.strip()
        if cleaned.startswith("```"):
            cleaned = re.sub(r"^```(?:json)?\s*", "", cleaned)
            cleaned = re.sub(r"\s*```$", "", cleaned)

        try:
            parsed = json.loads(cleaned)
        except json.JSONDecodeError:
            # Try to extract JSON array from the response
            match = re.search(r"\[.*\]", cleaned, re.DOTALL)
            if match:
                try:
                    parsed = json.loads(match.group())
                except json.JSONDecodeError:
                    return None
            else:
                return None

        if not isinstance(parsed, list) or len(parsed) == 0:
            return None

        # Validate each sub-task
        result = []
        for item in parsed:
            if not isinstance(item, dict):
                return None
            agent_type = item.get("agent_type", "")
            if agent_type not in AGENT_REGISTRY:
                return None
            sub = {
                "agent_type": agent_type,
                "description": item.get("description", ""),
            }
            if "shape_ids" in item:
                sub["shape_ids"] = item["shape_ids"]
            result.append(sub)

        return result if result else None
