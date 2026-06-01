import json
import time
import datetime

from editppt.utils.llm_client import call_llm, APIKeyError, set_token_log_context
from editppt.prompts import *
from editppt.utils.utils import parse_llm_response
from editppt.utils.logger_manual import log_path


def _coerce_page_number(pn, total_slides: int | None = None):
    """Return pn as a 1..total_slides int, or raise ValueError.

    LLM occasionally yields page_number as a non-integer token (e.g. the
    letter 'A' from "All slides" or 'a' from "every chart slide"), or an
    out-of-range int. Catch both so the planner-level retry loop can re-ask.
    """
    if isinstance(pn, bool):
        raise ValueError(f"page_number must be an int, got bool: {pn!r}")
    if isinstance(pn, int):
        n = pn
    elif isinstance(pn, str) and pn.strip().lstrip("-").isdigit():
        n = int(pn.strip())
    else:
        raise ValueError(f"page_number must be an int 1..N, got {pn!r}")
    if total_slides is not None and not (1 <= n <= total_slides):
        raise ValueError(f"page_number {n} out of range 1..{total_slides}")
    return n


def validate_plan_tasks(plan: dict, total_slides: int | None = None) -> dict:
    """Normalize every task.page_number to a valid int 1..total_slides.

    Raises ValueError on any invalid value — handled by the planner's retry
    loop, which will re-prompt the LLM with the error feedback.
    """
    tasks = plan.get("tasks", [])
    for i, t in enumerate(tasks):
        try:
            t["page_number"] = _coerce_page_number(t.get("page_number"), total_slides)
        except ValueError as e:
            raise ValueError(f"tasks[{i}]: {e}")
    return plan


def expand_pattern_plan(plan: dict, total_slides: int | None = None) -> dict:
    """
    Convert pattern-based plan into explicit per-slide tasks.

    Safety net: if a pattern omits `target_page_numbers` or supplies an
    empty list, fall back to ALL slides 1..total_slides — the specialist
    agent will skip non-matching ones. This guards against the LLM emitting
    placeholder characters for conditional patterns it cannot evaluate.
    """
    if plan.get("task_mode") != "pattern":

        return plan

    explicit_tasks = []

    pattern_tasks = plan.get("pattern_tasks", [])
    if not pattern_tasks:
        raise ValueError("task_mode is 'pattern' but no pattern_tasks found.")

    for pattern in pattern_tasks:
        target_pages = pattern.get("target_page_numbers")
        if not target_pages and total_slides is not None:
            target_pages = list(range(1, total_slides + 1))
        elif not target_pages:
            raise ValueError("pattern has empty target_page_numbers and no total_slides given.")

        for i in target_pages:
            task = {
                "page_number": i,
                "description": pattern["description"].format(i=i),
                "target": pattern["target"].format(i=i),
                "action": pattern["action"].format(i=i),
                "contents": pattern.get("contents", {}).copy(),
            }

            # Handle reference (if present)
            if "reference" in pattern:
                ref_page = pattern["reference"]["page_number"]
                if isinstance(ref_page, str):
                    ref_page = ref_page.format(i=i)
                task["reference"] = {
                    "page_number": int(ref_page)
                }

            explicit_tasks.append(task)

    return {
        "understanding": plan.get("understanding", ""),
        "tasks": explicit_tasks,
        "task_mode": "pattern",
    }


class Planner:
    def __init__(self, model, slide_name):
        self.slide_name = slide_name
        self.model = model

    def __call__(self, user_input: str, total_slide_numbers: int):
        planner_system_prompt = create_plan_prompt(
            self.slide_name,
            total_slide_numbers,
        )

        last_error_feedback = ""
        MAX_RETRIES = 3

        logs: list[str] = []
        had_error = False

        def append_log(text: str):
            logs.append(text + "\n")

        for attempt in range(1, MAX_RETRIES + 1):
            try:
                set_token_log_context(component="planner")
                call_llm_response = call_llm(
                    model=self.model,
                    messages=[
                        {"role": "system", "content": planner_system_prompt},
                        {
                            "role": "user",
                            "content": "Now, please create a plan for the following request:\n"
                            + user_input
                            + last_error_feedback,
                        },
                    ],
                    prompt_cache_key="planner",
                )

                response = call_llm_response.output_text

                append_log(
                    f"""[PLANNER: ATTEMPT {attempt}] RAW LLM RESPONSE
[USER INPUT]
{user_input}

[ERROR FEEDBACK]
{last_error_feedback}

[RAW LLM RESPONSE]
{response}
"""
                )

                plan, error = parse_llm_response(response)

                # Semantic validation (page_number int + range) — surfaces as
                # error so the retry loop reprompts the LLM with feedback.
                if error is None:
                    try:
                        plan = expand_pattern_plan(plan, total_slide_numbers)
                        validate_plan_tasks(plan, total_slide_numbers)
                    except ValueError as ve:
                        error = (ve, json.dumps(plan, ensure_ascii=False)[:1200])

                # First attempt succeeded with no errors; return immediately
                if error is None and not had_error:
                    return plan

                # Previous attempts had errors, but this one succeeded
                if error is None:

                    append_log(
                        f"""[PLANNER: ATTEMPT {attempt}] PARSED SUCCESS
{json.dumps(plan, indent=2, ensure_ascii=False)}
"""
                    )
                    break

                # parsing error
                had_error = True
                e, state = error

                append_log(
                    f"""[PLANNER: ATTEMPT {attempt}] PARSE ERROR
[EXCEPTION]
{type(e).__name__}: {e}

[FAILED PAYLOAD / STATE]
{state}
"""
                )

                # When the failure is page_number-shape, add a one-line hint:
                # the LLM needs a clear int constraint, not generic parse advice.
                page_hint = ""
                if "page_number" in str(e):
                    page_hint = (
                        f"\nNOTE: `page_number` must be an integer "
                        f"(1..{total_slide_numbers}), nothing else."
                    )

                last_error_feedback = f"""
#### Additional Information for Correction ####
The previous response could not be parsed.

Error:
{type(e).__name__}: {e}
{page_hint}

Invalid output:
{state}

Please fix the errors above and return ONLY valid JSON.
Do NOT include comments, explanations, or placeholders.
"""

                print(f"[PLANNER: Attempt {attempt}/{MAX_RETRIES}] Parsing failed. Retrying...")
                print(f"  ├─ Error: {type(e).__name__}: {e}")
                print("  └─ RAW LLM RESPONSE:")
                print(response)
                time.sleep(1)

            except APIKeyError:
                raise  # API key issue — no point retrying

            except Exception as e:
                had_error = True

                append_log(
                    f"""[PLANNER: ATTEMPT {attempt}] EXCEPTION DURING LLM CALL
{type(e).__name__}: {e}
"""
                )

                print(f"[PLANNER: Attempt {attempt}/{MAX_RETRIES}] Exception during LLM call: {e}")
                time.sleep(1)

        # Save debug logs only if there were errors
        if had_error:
            ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")

            debug_dir = log_path("llm_debug_logs").parent / "llm_debug_logs"
            debug_dir.mkdir(parents=True, exist_ok=True)

            log_file_path = debug_dir / f"planner_{ts}.log"
            log_file_path.write_text("".join(logs), encoding="utf-8")

        if "plan" in locals():
            # If we exhausted retries with `had_error` still True, the last
            # `plan` may still be semantically invalid. Validate one more time
            # so the cell fails cleanly (and is logged + retried by the
            # batch runner) instead of silently emitting a degenerate plan.
            if had_error:
                try:
                    validate_plan_tasks(plan, total_slide_numbers)
                except ValueError as ve:
                    raise RuntimeError(
                        f"Planner exhausted {MAX_RETRIES} retries and the final "
                        f"plan is still invalid: {ve}"
                    )
            return plan

        raise RuntimeError("Failed to obtain a valid plan from the LLM response.")
