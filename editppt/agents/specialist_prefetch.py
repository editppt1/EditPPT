"""Parallel pre-generation of specialist first-attempt tool_calls.

After the dispatcher has decided which sub-agent handles each shape, every
sub-task's *first* LLM round can be issued in parallel (parser cache hits,
no COM mutation). The sequential consumer in main.py then executes those
cached tool_calls one slide at a time. Retries still go through the normal
sequential path because they need validator feedback.
"""
from __future__ import annotations

import contextvars
import traceback
from concurrent.futures import ThreadPoolExecutor

from loguru import logger

from editppt.utils.llm_client import APIKeyError


def prefetch_specialist_tool_calls(
    tasks: list[dict],
    dispatcher_results: dict[int, list[dict]],
    specialist_agents: dict[str, object],
    parser: object,
    max_workers: int = 8,
) -> dict[tuple[int, int], list[dict] | None]:
    """For every (task, sub_task) pair, generate first-attempt tool_calls.

    Returns a dict keyed by (task_index, sub_index) → tool_calls list (may be
    empty) or None on per-job failure. Callers should treat None as "no cache,
    fall back to normal path".

    Pre-conditions:
    - dispatcher_results was produced by dispatcher.prefetch_dispatch_decisions
      so per-slide parses are already cached in parser.database.
    """
    if not dispatcher_results:
        return {}

    jobs: list[tuple[int, int, object, dict, list]] = []
    for i, task in enumerate(tasks):
        sub_tasks = dispatcher_results.get(i)
        if not sub_tasks:
            continue
        for sub_idx, sub in enumerate(sub_tasks):
            agent_type = sub.get("agent_type")
            agent = specialist_agents.get(agent_type)
            if agent is None:
                continue
            sub_task = {**task, "description": sub.get("description", "")}
            shape_ids = sub.get("shape_ids")
            jobs.append((i, sub_idx, agent, sub_task, shape_ids))

    if not jobs:
        return {}
    logger.info(f"specialist prefetch submitting {len(jobs)} jobs")

    def _job(item):
        i, sub_idx, agent, sub_task, shape_ids = item
        # Run each worker in a copied context so per-thread token-log metadata
        # (slide_index, agent_type, attempt_idx, scope) doesn't race siblings.
        ctx = contextvars.copy_context()
        try:
            tcs = ctx.run(agent.prefetch_first_attempt, sub_task, parser, shape_ids)
        except APIKeyError:
            raise
        except Exception as e:
            slide_idx = sub_task.get("page_number")
            logger.warning(
                f"specialist prefetch failed (task #{i}, sub #{sub_idx}, slide {slide_idx}): "
                f"{type(e).__name__}: {e}\n{traceback.format_exc()}"
            )
            return (i, sub_idx, None)
        return (i, sub_idx, tcs)

    results: dict[tuple[int, int], list[dict] | None] = {}
    with ThreadPoolExecutor(max_workers=max_workers) as ex:
        for i, sub_idx, tcs in ex.map(_job, jobs):
            results[(i, sub_idx)] = tcs
    return results
