import json
import time
import sys
import argparse
from pathlib import Path

from loguru import logger

from editppt.ppt_core import PPTContainer, kill_powerpoint_processes, initialize_ppt, init_backup
from editppt.utils.logger_manual import *
from editppt.utils.utils import *
from editppt.utils.llm_client import (
    get_token_snapshot,
    diff_tokens,
    set_token_log_path,
    set_token_log_context,
    reset_token_counter,
    set_request_summary_path,
    log_request_summary,
)
from editppt.agents import DispatcherAgent, create_specialist_agents, VisionValidatorAgent, VisualFixerAgent
from editppt.agents.base_agent import ALL_TOOLS_SCHEMA
from editppt.parser import Parser
from editppt.planner import Planner
from editppt.config import *


logger = logger

def parse_args():
    parser = argparse.ArgumentParser(
        description="PPT Editing Agent System"
    )
    
    parser.add_argument(
        "--file_path",
        type=str,
        required=True,
        help="Path to the PPTX file (absolute or relative)"
    )
    
    parser.add_argument(
        "--prompt",
        type=str,
        default=None,
        help="User instruction text (if set, bypasses interactive input())"
    )
    parser.add_argument(
        "--no-vision-validation",
        action="store_true",
        help="Disable the vision validator + visual fixer agents for this run."
    )
    parser.add_argument(
        "--main-model",
        type=str,
        default=None,
        help="Override CURRENT_MODEL_NAME and STYLE_MAPPER_MODEL for this run "
             "(e.g., claude-haiku-4-5, gpt-4.1). Useful for cross-model benchmarks.",
    )
    return parser.parse_args()

def main():
    logger = init_logger()
    args = parse_args()

    # Apply --main-model override before any agent/tool reads model names.
    # Patches both the local binding (from `from editppt.config import *`) and
    # tools.py's local STYLE_MAPPER_MODEL binding so style_mapper calls follow.
    global CURRENT_MODEL_NAME, STYLE_MAPPER_MODEL
    if args.main_model:
        CURRENT_MODEL_NAME = args.main_model
        STYLE_MAPPER_MODEL = args.main_model
        import editppt.config as _cfg
        _cfg.CURRENT_MODEL_NAME = args.main_model
        _cfg.STYLE_MAPPER_MODEL = args.main_model
        import editppt.tools.tools as _ttools
        _ttools.STYLE_MAPPER_MODEL = args.main_model

    ppt_path = Path(args.file_path).expanduser().resolve()

    logger.info("=== PPT Editing Agent System ===")
    logger.info(f"Resolved PPT path: {ppt_path}")

    kill_powerpoint_processes()
    time.sleep(1)

    try:
        prs, ppt_app = initialize_ppt(ppt_path)
        container = PPTContainer(prs, ppt_app)
        logger = init_logger(container)
        
        current_log_root = get_dynamic_log_dir(container)
        logger.info(f"Log directory initialized: {current_log_root}")
        
        logger.info(
            f"Presentation [{container.prs.Name}] loaded "
            f"with {len(container.prs.Slides)} slides."
        )
    except Exception as e:
        logger.error(f"Failed to initialize PowerPoint: {e}")
        sys.exit(1)

    planner = Planner(
        model=CURRENT_MODEL_NAME,
        slide_name=container.prs.Name
    )
    logger.info("Planner initialized")

    parser = Parser(
        container=container,
        total_slides=len(container.prs.Slides),
    )
    logger.info("Parser initialized")

    log_root = current_log_root
    log_root.mkdir(parents=True, exist_ok=True)

    # Per-call token usage persistence — survives crashes and Ctrl-C
    set_token_log_path(log_root / "token_usage.jsonl")
    set_request_summary_path(log_root / "request_summary.jsonl")
    reset_token_counter()

    (log_root / "parser_Database.json").write_text(
        json.dumps(parser.database, ensure_ascii=False, indent=4),
        encoding="utf-8",
    )

    # Backup sharing: initialize backup before creating agents
    init_backup(container)

    dispatcher = DispatcherAgent(model=CURRENT_MODEL_NAME)
    specialist_agents = create_specialist_agents(
        container=container,
        model=CURRENT_MODEL_NAME,
    )

    vision_validator_agent = VisionValidatorAgent.create(
        activate_valid=not args.no_vision_validation,
        container=container,
        model=CURRENT_VISION_MODEL_NAME,
    )
    visual_fixer_agent = VisualFixerAgent.create(
        container=container,
        model=CURRENT_MODEL_NAME,
        tools_schema=ALL_TOOLS_SCHEMA,
    ) if vision_validator_agent is not None else None
    logger.info("Agents initialized (dispatcher + specialists)")

    logger.info(f"[System] Agent is ready using model: {CURRENT_MODEL_NAME}")

    # Agent loop
    cli_request_id = 0
    while True:
        if args.prompt is not None:
            user_input = args.prompt.strip()
        else:
            user_input = input("[User]: ").strip()
        if not user_input:
            continue

        if user_input == "eee":
            kill_powerpoint_processes()
            time.sleep(1)
            sys.exit(0)

        cli_request_id += 1
        submit_ts = time.time()
        token_snapshot_start = get_token_snapshot()
        set_token_log_context(request_id=cli_request_id, scope="cli")

        req_status = "done"
        tasks: list = []
        plan_json: dict = {}
        try:
            plan_json = planner(
                user_input=user_input,
                total_slide_numbers=len(container.prs.Slides)
                )
            logger.info(f"Planner output received")

            (log_root / "planner.json").write_text(
                json.dumps(plan_json, ensure_ascii=False, indent=4),
                encoding="utf-8",
            )
            tasks = plan_json.get("tasks", [])

            # Pre-compute dispatcher decisions in parallel. Slides are parsed
            # serially (COM-bound) but dispatcher LLM calls run concurrently.
            # Results for slides modified by an earlier task are discarded below.
            prefetched = dispatcher.prefetch_dispatch_decisions(tasks, parser) if len(tasks) > 1 else {}
            modified_slides: set[int] = set()

            # Pre-compute specialist first-attempt tool_calls in parallel.
            # Uses the cached parse from dispatcher prefetch (no extra COM).
            # Keyed by (task_index, sub_index); consumed inside the loop below
            # only when the slide hasn't been touched by a prior task.
            from editppt.agents.specialist_prefetch import prefetch_specialist_tool_calls
            specialist_prefetched = prefetch_specialist_tool_calls(
                tasks, prefetched, specialist_agents, parser, max_workers=8
            ) if len(tasks) > 1 else {}

            # Vision validation pipelining: each agent submits an async Future
            # here instead of waiting; we drain after all slides are edited.
            vision_queue: list[dict] = [] if vision_validator_agent is not None else None

            skipped_slides: list[tuple[int, str]] = []
            for i, task in enumerate(tasks):
                slide_idx = task.get("page_number")
                set_token_log_context(slide_index=slide_idx)

                # Per-task safety net: a single slide's failure must never abort
                # the entire request. Worst case we skip the slide and continue
                # — the user wants the deck partially edited rather than not
                # at all.
                try:
                    cached = prefetched.get(i)
                    if cached is not None and slide_idx not in modified_slides:
                        sub_tasks = cached
                    else:
                        slide_contents = parser.process(slide_idx) if slide_idx else None
                        objects_detail = slide_contents.get("Objects_Detail", []) if slide_contents else []
                        sub_tasks = dispatcher.dispatch(task, objects_detail)

                    logger.info(f"Dispatcher routed to: {[(s['agent_type'], s.get('shape_ids', [])) for s in sub_tasks]}")

                    for sub_idx, sub in enumerate(sub_tasks):
                        agent_type = sub["agent_type"]
                        sub_task = {**task, "description": sub["description"]}
                        # Use prefetched tool_calls only if slide is still pristine.
                        prefetched_tcs = None
                        if slide_idx not in modified_slides:
                            prefetched_tcs = specialist_prefetched.get((i, sub_idx))
                        try:
                            specialist_agents[agent_type].run(
                                task=sub_task,
                                parser=parser,
                                vision_validator_agent=vision_validator_agent,
                                visual_fixer_agent=visual_fixer_agent,
                                auto_resize=True,
                                shape_ids=sub.get("shape_ids"),
                                prefetched_tool_calls=prefetched_tcs,
                                vision_queue=vision_queue,
                            )
                        except KeyboardInterrupt:
                            raise
                        except Exception as sub_err:
                            logger.warning(
                                f"Slide {slide_idx} sub-task {sub_idx} ({agent_type}) "
                                f"failed: {type(sub_err).__name__}: {sub_err}. "
                                f"Continuing with next sub-task."
                            )
                            skipped_slides.append((slide_idx, f"{agent_type}: {sub_err}"))

                    if slide_idx:
                        modified_slides.add(slide_idx)
                except KeyboardInterrupt:
                    raise
                except Exception as task_err:
                    logger.warning(
                        f"Slide {slide_idx} (task {i}) errored: {type(task_err).__name__}: "
                        f"{task_err}. Skipping this slide and continuing."
                    )
                    skipped_slides.append((slide_idx, f"task-level: {task_err}"))
                    # Try to recover COM state so the next slide can proceed.
                    try:
                        # If a specialist agent left COM corrupted, force-reopen
                        # from the backup the agent itself just saved on the
                        # previous successful slide.
                        sample_agent = next(iter(specialist_agents.values()))
                        sample_agent._rollback_ppt("task-skip", str(task_err))
                    except Exception as rb_err:
                        logger.error(f"Skip-recovery failed: {rb_err}")

            if skipped_slides:
                logger.warning(
                    f"Skipped {len(skipped_slides)} slide(s): "
                    f"{[(s, r[:60]) for s, r in skipped_slides]}"
                )

            # Drain pending vision validations. Edits are already applied and
            # saved; for any slide that fails validation, hand off to the
            # visual fixer now (sequential, COM-bound).
            if vision_queue:
                logger.info(f"Draining {len(vision_queue)} pending vision validations...")
                for entry in vision_queue:
                    pn = entry.get("page_number", "?")
                    # Per-entry guard so one bad drain doesn't kill the rest.
                    try:
                        try:
                            valid, reason, issues = entry["future"].result()
                        except Exception as e:
                            logger.warning(f"Vision Future for slide {pn} errored: {e}")
                            continue
                        if valid:
                            continue
                        vfa = entry["visual_fixer_agent"]
                        if vfa is None:
                            logger.warning(f"Slide {pn} failed vision but no visual_fixer; leaving as-is.")
                            continue
                        logger.info(f"Slide {pn}: visual_fixer running post-hoc fixes...")
                        # Single fix shot, no re-validation. Trial 3 logs showed
                        # the iterative loop oscillating (whack-a-mole) — converging
                        # rarely justifies the 3x gemini-pro + LLM cost. Accept
                        # best-effort fix and let the user inspect the deck.
                        try:
                            vfa.run(
                                page_number=pn,
                                defects=issues,
                                slide_json=entry["new_parse"],
                                vision_validator_agent=entry["vision_validator_agent"],
                                agent_request=entry["agent_request"],
                                used_tools=entry["tool_calls"],
                                parser=entry["parser"],
                                max_attempts=1,
                            )
                        except Exception as e:
                            logger.warning(f"Slide {pn}: visual_fixer raised ({e}); continuing.")
                        # Refresh parser cache + save after fixes
                        try:
                            new_parse = entry["parser"].process(pn, force=True)
                            entry["parser"].database[pn] = new_parse
                        except Exception as e:
                            logger.warning(f"Slide {pn}: post-fix re-parse failed: {e}")
                        try:
                            entry["container"].prs.SaveCopyAs(entry["backup_path"])
                            entry["container"].prs.SaveCopyAs(entry["original_path"])
                        except Exception as e:
                            logger.warning(f"Slide {pn}: post-fix save failed: {e}")
                    except KeyboardInterrupt:
                        raise
                    except Exception as drain_err:
                        logger.warning(
                            f"Slide {pn} drain entry failed: "
                            f"{type(drain_err).__name__}: {drain_err}. Continuing drain."
                        )
        except KeyboardInterrupt:
            req_status = "aborted"
            raise
        except Exception:
            req_status = "error"
            raise
        finally:
            delta = diff_tokens(token_snapshot_start, get_token_snapshot())
            done_ts = time.time()
            elapsed = done_ts - submit_ts
            slide_count = len({t.get("page_number") for t in tasks if t.get("page_number")})
            log_request_summary(
                request_id=cli_request_id,
                user_input=user_input,
                submit_ts=submit_ts,
                done_ts=done_ts,
                status=req_status,
                scope="cli",
                task_count=len(tasks),
                slide_count=slide_count,
                input_tokens=delta["input_tokens"],
                output_tokens=delta["output_tokens"],
                cached_input_tokens=delta.get("cached_input_tokens", 0),
                cost_usd=delta.get("cost_usd", 0.0),
                planner_task_mode=plan_json.get("task_mode"),
            )
            logger.info(
                f"(Request={user_input}) | elapsed={elapsed:.3f}s | "
                f"tokens: {delta['input_tokens']}in / {delta['output_tokens']}out, "
                f"cached_in: {delta.get('cached_input_tokens', 0)} | "
                f"cost: ${delta.get('cost_usd', 0.0):.6f}"
            )
            set_token_log_context(request_id=None, slide_index=None, scope=None)

        if args.prompt is not None:
            break


if __name__ == "__main__":
    main()
    