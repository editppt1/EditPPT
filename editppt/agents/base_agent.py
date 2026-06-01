import os
import json
import time
import traceback
from copy import deepcopy
from pathlib import Path

from editppt.tools.agent_tool_registry import TOOL_METADATA, ADDITIONAL_CALL_TOOL_DICT
from editppt.tools.tools import FUNCTION_MAP, clamp_shapes_to_slide, clamp_text_to_slide

# Tools where overflow should be handled by shrinking font size (preserve position)
_TEXT_CLAMP_TOOLS = {
    "edit_text_rewrite", "edit_text_replace", "edit_text_insert",
    "replace_table_text", "set_text_style", "manage_bullet_points",
}
from editppt.prompts import create_edit_agent_user_prompt
from editppt.utils.llm_client import call_llm, APIKeyError, set_token_log_context, is_anthropic_model
from editppt.utils.logger_manual import init_logger, log_path
from editppt.agents.registry import AgentSpec, AGENT_REGISTRY
from editppt.agents.slide_validator import capture_pre_state, validate_slide_ops
from editppt.config import RESOURCE_ROOT

logger = init_logger()

SCHEMA_PATH = RESOURCE_ROOT / "editppt" / "tools" / "tools_schema.json"


def _collect_shape_ids(objects):
    """Recursively collect every Shape_Id in Objects_Detail, including those
    nested under group shapes (More_detail.Group). Mirrors the parser's
    Objects_Detail layout."""
    ids = set()
    for obj in objects or []:
        sid = obj.get("Shape_Id")
        if sid is not None:
            ids.add(sid)
        inner = obj.get("More_detail", {}).get("Group", [])
        if inner:
            ids.update(_collect_shape_ids(inner))
    return ids


def _build_shape_inventory(slide_data) -> str:
    """One-line-per-shape compact summary, designed to be appended to retry
    feedback so the LLM has a concrete pick-list instead of re-inferring from
    the full slide JSON. Includes placeholder role index when available."""
    # Gate for A/B trials. Set EDITPPT_DISABLE_D45=1 to revert to pre-D5 feedback.
    if os.environ.get("EDITPPT_DISABLE_D45") == "1":
        return ""
    if not isinstance(slide_data, dict):
        return ""
    lines = []
    # Visual_Role_Index supersedes Placeholder_Index (it folds in the placeholder
    # signal plus name regex and visual-fallback). Surface it first when
    # available — falls back to raw Placeholder_Index for older snapshots.
    vr_idx = slide_data.get("Visual_Role_Index") or {}
    if any(vr_idx.values()):
        vr_str = ", ".join(f"{role}={ids}" for role, ids in vr_idx.items() if ids)
        if vr_str:
            lines.append(f"Visual roles (inferred): {vr_str}")
    ph_idx = slide_data.get("Placeholder_Index") or {}
    if ph_idx:
        idx_str = ", ".join(f"{role}={ids}" for role, ids in ph_idx.items() if ids)
        if idx_str:
            lines.append(f"Placeholder types: {idx_str}")

    def _walk(objects, depth=0):
        for obj in objects or []:
            sid = obj.get("Shape_Id")
            if sid is None:
                continue
            md = obj.get("More_detail", {}) or {}
            ph_role = (md.get("Placeholder") or {}).get("Placeholder Type Name") or "-"
            tf = md.get("TextFrame") or {}
            runs = tf.get("Runs") or []
            text = (runs[0].get("Text", "")[:40].replace("\n", " ").replace("\r", " ")) if runs else ""
            prefix = "  " * depth + "  - "
            lines.append(f"{prefix}id={sid}  type={obj.get('Type','?')}  role={ph_role}  text={text!r}")
            inner = md.get("Group") or []
            if inner:
                _walk(inner, depth + 1)

    _walk(slide_data.get("Objects_Detail", []))
    if not lines:
        return ""
    return "Shape inventory on this slide:\n" + "\n".join(lines)

with open(SCHEMA_PATH, "r", encoding="utf-8") as f:
    ALL_TOOLS_SCHEMA = json.load(f)


class BaseEditAgent:
    def __init__(self, container, model: str, agent_spec: AgentSpec):
        self.container = container
        self.model = model
        self.agent_spec = agent_spec
        self.backup_path = self.container.backup_path
        self.original_path = self.container.original_path

        # Filter tools schema to only this agent's tools
        self.tools_schema = [
            t for t in ALL_TOOLS_SCHEMA
            if t["name"] in agent_spec.tool_names
        ]

        self.messages = []

    _AUTO_RESIZE_TOOLS = {
        "edit_text_rewrite", "edit_text_replace", "edit_text_insert",
        "replace_table_text",
    }

    def prefetch_first_attempt(self, task: dict, parser: object,
                                shape_ids: list[int] | None = None) -> list[dict] | None:
        """Generate first-attempt tool_calls without executing them.

        Safe to call from a worker thread: only parses (via parser cache) and
        does a single LLM call. No COM mutation, no validator, no retry.

        Returns the parsed tool_calls list (may be empty) or None on failure.
        """
        page_number = task.get("page_number")
        if not page_number:
            return None

        # Tag token-usage entries to the right slide for cost rollups.
        set_token_log_context(
            agent_type=self.agent_spec.agent_type,
            attempt_idx=1,
            slide_index=page_number,
            scope="specialist_prefetch",
        )

        description = task.get("description", "")
        action = task.get("action", "")
        reference_page_number = task.get("reference_page_number", "")

        try:
            # Worker thread has no COM context, so only the parser cache is
            # safe to read. If a slide isn't already parsed (dispatcher prefetch
            # should have done that), skip — main thread will fall back to the
            # normal sequential path.
            contents = parser.database.get(page_number)
            if contents is None:
                return None

            if shape_ids:
                contents = deepcopy(contents)
                contents["Objects_Detail"] = [
                    obj for obj in contents.get("Objects_Detail", [])
                    if obj.get("Shape_Id") in shape_ids
                ]

            reference_slide_contents = []
            if reference_page_number:
                idx = reference_page_number
                cached = parser.database.get(idx)
                if cached:
                    reference_slide_contents.append({"page_number": idx, "contents": cached})
                # else: skip reference — re-parsing in worker would touch COM

            payload_message = [
                {"role": "system", "content": self.agent_spec.system_prompt_builder()},
                {"role": "user", "content": create_edit_agent_user_prompt(
                    page_number=page_number,
                    description=description,
                    action=action,
                    contents=contents,
                    reference_slide_contents=reference_slide_contents if reference_page_number else None,
                    feedback=[],
                )},
            ]

            tc = {"type": "any"} if is_anthropic_model(self.model) else "auto"
            set_token_log_context(component="specialist")
            response = call_llm(
                model=self.model,
                messages=payload_message,
                tools=self.tools_schema,
                tool_choice=tc,
                prompt_cache_key=f"specialist:{self.agent_spec.agent_type}",
            )

            tool_calls = []
            for item in response.output:
                if item.type == "function_call":
                    tool_calls.append({
                        "name": item.name,
                        "arguments": json.loads(item.arguments) if isinstance(item.arguments, str) else item.arguments,
                        "call_id": item.call_id,
                    })
            return tool_calls
        except APIKeyError:
            raise
        except Exception as e:
            logger.warning(f"prefetch_first_attempt failed (slide {page_number}): {e}")
            return None
        finally:
            # Don't leak per-prefetch context to subsequent code on the main thread.
            set_token_log_context(scope=None, attempt_idx=None)

    def run(self, task: dict, parser: object, vision_validator_agent: object,
            visual_fixer_agent: object | None = None,
            auto_resize: bool = False, shape_ids: list[int] | None = None,
            prefetched_tool_calls: list[dict] | None = None,
            vision_queue: list | None = None):
        self._auto_resize = auto_resize
        self._shape_ids = shape_ids
        self._visual_fixer_agent = visual_fixer_agent
        self._vision_queue = vision_queue  # None = sync mode, list = async pipelining
        self._parser = parser  # for _validate_tool_args shape_id lookup
        feedback = []
        tool_calls = []
        max_retries = 3
        retry_count = 0

        print()

        set_token_log_context(agent_type=self.agent_spec.agent_type)
        while retry_count < max_retries:
            retry_count += 1
            set_token_log_context(attempt_idx=retry_count)
            payload_message = []
            page_number = task.get("page_number")
            if not page_number:
                print('wrong page number.')
                continue

            description = task.get("description", "")
            action = task.get("action", "")
            detailed_contents = task.get("contents", "")
            reference_page_number = task.get("reference_page_number", "")

            # Skip parsing if slide doesn't exist yet (slide agent will create via add_slide)
            if page_number > parser.container.prs.Slides.Count:
                contents = None
            else:
                contents = parser.process(page_number, force=(retry_count > 1))

            # Filter to only the shapes this agent is responsible for
            if self._shape_ids and contents:
                contents = deepcopy(contents)
                contents["Objects_Detail"] = [
                    obj for obj in contents.get("Objects_Detail", [])
                    if obj.get("Shape_Id") in self._shape_ids
                ]

            reference_slide_contents = []
            if reference_page_number:
                idx = reference_page_number
                if idx in parser.database and parser.database[idx]:
                    reference_slide_contents.append({
                        "page_number": idx,
                        "contents": parser.database[idx],
                    })
                else:
                    cur_page = parser.process(idx)
                    reference_slide_contents.append({
                        "page_number": idx,
                        "contents": cur_page,
                    })

            # Use specialist system prompt instead of generic one.
            # Slide JSON is no longer embedded in the system prompt — it lives
            # in the user prompt so the system prefix stays static across calls
            # (enables prompt caching and avoids duplication).
            payload_message.append({
                "role": "system",
                "content": self.agent_spec.system_prompt_builder()
            })

            payload_message.append({
                "role": "user",
                "content": create_edit_agent_user_prompt(
                    page_number=page_number,
                    description=description,
                    action=action,
                    contents=contents,
                    reference_slide_contents=reference_slide_contents if reference_page_number else None,
                    feedback=feedback,
                )
            })
            with open(
                log_path(f"agent_payload_message_{page_number}.txt"),
                "w",
                encoding="utf-8"
            ) as f:
                for i, msg in enumerate(payload_message, 1):
                    f.write(f"[{i}]\n")
                    f.write(f"role: {msg.get('role')}\n")

                    content = msg.get("content")
                    if isinstance(content, list):
                        for block in content:
                            if block.get("type") == "text":
                                f.write(block.get("text", ""))
                                f.write("\n")
                    else:
                        f.write(str(content) + "\n")
                    f.write("\n" + "-" * 50 + "\n\n")

            # Use filtered tools schema instead of full TOOLS_SCHEMA.
            # prompt_cache_key routes specialist calls to the same cache shard
            # so the (now-static) system prompt + tools schema benefit from
            # OpenAI's automatic prefix caching across calls and retries.
            # First iteration may consume a prefetched tool_calls list (parallel
            # generation done before this slide was reached). Subsequent
            # iterations always call the LLM fresh because retry needs feedback.
            if retry_count == 1 and prefetched_tool_calls is not None:
                tool_calls = prefetched_tool_calls
                with open(
                    log_path(f"agent_toolcall_response_{page_number}.json"),
                    "w",
                    encoding="utf-8"
                ) as f:
                    json.dump({"source": "prefetched", "tool_calls": tool_calls}, f, ensure_ascii=False, indent=4)
            else:
                tc = {"type": "any"} if is_anthropic_model(self.model) else "auto"
                set_token_log_context(component="specialist")
                response = call_llm(
                    model=self.model,
                    messages=payload_message,
                    tools=self.tools_schema,
                    tool_choice=tc,
                    prompt_cache_key=f"specialist:{self.agent_spec.agent_type}",
                )

                response_dict = response.model_dump()

                with open(
                    log_path(f"agent_toolcall_response_{page_number}.json"),
                    "w",
                    encoding="utf-8"
                ) as f:
                    json.dump(response_dict, f, ensure_ascii=False, indent=4)

                # Tool Call Parsing
                tool_calls = []
                for item in response.output:
                    if item.type == "function_call":
                        tool_calls.append({
                            "name": item.name,
                            "arguments": json.loads(item.arguments) if isinstance(item.arguments, str) else item.arguments,
                            "call_id": item.call_id
                        })

            failed_tool_name = None
            failed_tool_args = None
            tool_error_reason = None

            agent_request = (description, action, detailed_contents)

            # Snapshot state for slide-level rule validation (cheap; always capture)
            pre_slide_state = capture_pre_state(self.container.prs)

            for tool_call in tool_calls:
                function_name = tool_call["name"]
                function_args = tool_call["arguments"]

                if function_name in ADDITIONAL_CALL_TOOL_DICT:
                    new_tool_call = ADDITIONAL_CALL_TOOL_DICT[function_name](
                        function_name=function_name,
                        function_args=function_args,
                        agent_request=agent_request,
                        model=self.model,
                        tools=self.tools_schema
                    )

                    function_name = new_tool_call["name"]
                    function_args = new_tool_call["arguments"]

                logger.info(f"Tool Call: {function_name}({function_args})")

                meta = TOOL_METADATA.get(function_name)

                if meta:
                    if meta.needs_slide_json:
                        function_args["slide_json"] = contents

                    if meta.needs_agent_request:
                        function_args["agent_request"] = agent_request

                    if meta.needs_container:
                        function_args["container"] = self.container

                    if meta.cleanup_false_args:
                        for key in meta.cleanup_false_args:
                            if function_args.get(key) is False:
                                function_args.pop(key)

                try:
                    self._validate_tool_args(function_name, function_args)
                    self._execute_tool(function_name, function_args)

                except Exception as e:
                    failed_tool_name = function_name
                    failed_tool_args = deepcopy(function_args)
                    tool_error_reason = f"{function_name} failed: {e}"
                    break

            if tool_error_reason:
                executed_calls_str = ", ".join(
                    [f"{tc['name']}({tc['arguments']})" for tc in tool_calls]
                )
                inventory = _build_shape_inventory(contents)

                feedback.append(
                    f"""Tool Execution Failed
                    - Failed Tool: {failed_tool_name}
                    - Failed Args: {json.dumps(failed_tool_args, ensure_ascii=False)}
                    - Error: {tool_error_reason}
                    - Planned Tools: [{executed_calls_str}]
                    {inventory}
                    """
                )
                self._rollback_ppt("tool", tool_error_reason)
                continue

            # Route validation:
            #   - Slide-level tools (transition/background/add/delete/duplicate)
            #     don't appear in parsed JSON, so text/vision validation can't
            #     check them. Run rule-based COM checks instead.
            #   - When parse data is unavailable (e.g. new slide just created),
            #     fall through to the same path; the rule check covers add_slide.
            all_slide_rule = bool(tool_calls) and all(
                getattr(TOOL_METADATA.get(tc["name"]), "slide_rule_validation", False)
                for tc in tool_calls
            )
            if contents is None or all_slide_rule:
                if all_slide_rule:
                    valid, reason = validate_slide_ops(
                        self.container.prs, tool_calls, pre_slide_state
                    )
                    if not valid:
                        executed_calls_str = ", ".join(
                            [f"{tc['name']}({tc['arguments']})" for tc in tool_calls]
                        )
                        feedback.append(
                            f"Retry {retry_count} Slide Rule Fail: {reason} | "
                            f"Tools: [{executed_calls_str}]"
                        )
                        self._rollback_ppt("slide_rule", reason)
                        continue
                # Re-parse so parser.database reflects the current slide state
                if page_number <= parser.container.prs.Slides.Count:
                    parser.process(page_number, force=True)
                self.container.prs.SaveCopyAs(self.backup_path)
                self.container.prs.SaveCopyAs(self.original_path)
                break

            # Skip validation when the agent decided no tool was needed.
            # An empty tool_calls means the model judged the slide already
            # satisfies the request — there's nothing to compare/validate.
            if not tool_calls:
                logger.info(f"No tool calls produced for page {page_number}; skipping validation.")
                break

            # Validation 1 (Text/Logic)
            valid, reason, strategy, new_parse = parser.update_after_edit(
                text_validation=True,
                model=self.model,
                page_number=page_number,
                description=description,
                action=action,
                detailed_contents=detailed_contents,
                used_tools=tool_calls
            )

            # Validation 2 (Vision)
            if valid and vision_validator_agent is not None and self._vision_queue is not None:
                # Async pipelining: export PNG now (slide state captured),
                # submit LLM call to the executor, optimistically commit parser
                # DB + save, and let main.py drain the queue at the end to
                # decide whether to invoke visual_fixer.
                future = vision_validator_agent.submit_async(
                    page_number=page_number,
                    agent_request=action,
                    parsed_contents=new_parse,
                    used_tools=tool_calls,
                )
                self._vision_queue.append({
                    "future": future,
                    "page_number": page_number,
                    "agent_request": action,
                    "new_parse": new_parse,
                    "tool_calls": tool_calls,
                    "vision_validator_agent": vision_validator_agent,
                    "visual_fixer_agent": self._visual_fixer_agent,
                    "parser": parser,
                    "container": self.container,
                    "backup_path": self.backup_path,
                    "original_path": self.original_path,
                })
                parser.edit_history.setdefault(page_number, []).append(deepcopy(parser.database.get(page_number, None)))
                parser.database[page_number] = new_parse
                with open(log_path("parser_Edithistory.json"), "w", encoding="utf-8") as f:
                    json.dump(parser.edit_history, f, ensure_ascii=False, indent=4)
                with open(log_path("parser_Database.json"), "w", encoding="utf-8") as f:
                    json.dump(parser.database, f, ensure_ascii=False, indent=4)
                self.container.prs.SaveCopyAs(self.backup_path)
                self.container.prs.SaveCopyAs(self.original_path)
                break

            if valid and vision_validator_agent is not None:
                valid, reason, issues = vision_validator_agent.process(
                    page_number=page_number,
                    agent_request=action,
                    parsed_contents=new_parse,
                    used_tools=tool_calls)

                if valid:
                    parser.edit_history.setdefault(page_number, []).append(deepcopy(parser.database.get(page_number, None)))
                    parser.database[page_number] = new_parse

                    with open(log_path("parser_Edithistory.json"), "w", encoding="utf-8") as f:
                        json.dump(parser.edit_history, f, ensure_ascii=False, indent=4)
                    with open(log_path("parser_Database.json"), "w", encoding="utf-8") as f:
                        json.dump(parser.database, f, ensure_ascii=False, indent=4)

                    self.container.prs.SaveCopyAs(self.backup_path)
                    self.container.prs.SaveCopyAs(self.original_path)
                    break
                else:
                    # Text intent already validated. Hand vision defects to the
                    # visual fixer for incremental, concrete repair instead of
                    # rolling back to pre-edit state.
                    fix_succeeded = False
                    if self._visual_fixer_agent is not None:
                        # Single fix shot; see main.py drain comment for rationale.
                        fix_succeeded = self._visual_fixer_agent.run(
                            page_number=page_number,
                            defects=issues,
                            slide_json=new_parse,
                            vision_validator_agent=vision_validator_agent,
                            agent_request=action,
                            used_tools=tool_calls,
                            parser=parser,
                            max_attempts=1,
                        )

                    # Refresh parser cache so downstream consumers see post-fix state
                    if page_number <= parser.container.prs.Slides.Count:
                        new_parse = parser.process(page_number, force=True)

                    parser.edit_history.setdefault(page_number, []).append(deepcopy(parser.database.get(page_number, None)))
                    parser.database[page_number] = new_parse

                    with open(log_path("parser_Edithistory.json"), "w", encoding="utf-8") as f:
                        json.dump(parser.edit_history, f, ensure_ascii=False, indent=4)
                    with open(log_path("parser_Database.json"), "w", encoding="utf-8") as f:
                        json.dump(parser.database, f, ensure_ascii=False, indent=4)

                    self.container.prs.SaveCopyAs(self.backup_path)
                    self.container.prs.SaveCopyAs(self.original_path)

                    if not fix_succeeded:
                        executed_calls_str = ", ".join([f"{tc['name']}({tc['arguments']})" for tc in tool_calls])
                        feedback.append(
                            f"Retry {retry_count} Vision Fix Exhausted: {reason} | "
                            f"Original Tools: [{executed_calls_str}]"
                        )
                        logger.warning(
                            f"Vision fix did not converge for page {page_number}; accepting current state."
                        )
                    break

            else:
                if valid:
                    parser.edit_history.setdefault(page_number, []).append(deepcopy(parser.database.get(page_number, None)))
                    parser.database[page_number] = new_parse

                    with open(log_path("parser_Edithistory.json"), "w", encoding="utf-8") as f:
                        json.dump(parser.edit_history, f, ensure_ascii=False, indent=4)
                    with open(log_path("parser_Database.json"), "w", encoding="utf-8") as f:
                        json.dump(parser.database, f, ensure_ascii=False, indent=4)

                    self.container.prs.SaveCopyAs(self.backup_path)
                    self.container.prs.SaveCopyAs(self.original_path)
                    break
                else:
                    executed_calls_str = ", ".join([f"{tc['name']}({tc['arguments']})" for tc in tool_calls])
                    # D5: append shape inventory so the LLM gets a concrete
                    # pick-list (placeholder roles + per-shape role/text) instead
                    # of re-inferring from the full JSON.
                    # On INCREMENTAL the state is the post-edit slide → use
                    # new_parse so the inventory reflects what the LLM will see
                    # on the next attempt. On ROLLBACK the agent re-parses
                    # backup state next iteration, so the pre-edit `contents`
                    # is the right snapshot.
                    inventory = _build_shape_inventory(
                        new_parse if strategy == "incremental" else contents
                    )
                    if strategy == "incremental":
                        # Keep current state, agent will re-parse and make targeted fixes
                        logger.info(f"Incremental retry: keeping current state for page {page_number}")
                        feedback.append(
                            f"Retry {retry_count} Text Fail (INCREMENTAL): {reason} | "
                            f"Tools: [{executed_calls_str}]\n{inventory}"
                        )
                    else:
                        # Rollback to last good state and redo from scratch
                        feedback.append(
                            f"Retry {retry_count} Text Fail (ROLLBACK): {reason} | "
                            f"Tools: [{executed_calls_str}]\n{inventory}"
                        )
                        self._rollback_ppt("text", reason)
                    continue

        # Feedback file
        with open(log_path(f"agent_Feedback_{page_number}.json"), "w", encoding="utf-8") as f:
            json.dump(feedback, f, ensure_ascii=False, indent=4)

        # Clear per-agent log context so subsequent calls (validator, dispatcher,
        # other specialist agents) don't inherit a stale attempt_idx.
        set_token_log_context(attempt_idx=None, agent_type=None)
        return bool(tool_calls)

    def _validate_tool_args(self, name, args):
        """Reject obvious LLM hallucinations before they reach COM.

        Slide ranges are deck-specific so we can't encode them in tool schemas;
        without this guard a bad slide_number raises a COM exception that often
        leaves the PowerPoint Application in a state where the subsequent
        rollback (Close+Open) also fails, taking the whole request down.
        """
        slide_number = args.get("slide_number")
        if slide_number is None:
            return
        try:
            total_slides = int(self.container.prs.Slides.Count)
        except Exception:
            return  # Can't introspect — best effort, let the tool try.
        if not isinstance(slide_number, int):
            try:
                slide_number = int(slide_number)
            except Exception:
                raise ValueError(
                    f"slide_number must be an integer, got {type(slide_number).__name__}"
                )
        # Tools that add/duplicate slides are allowed to target Count+1.
        upper = total_slides + 1 if name in {"add_slide", "duplicate_slide"} else total_slides
        if not (1 <= slide_number <= upper):
            raise ValueError(
                f"slide_number={slide_number} out of valid range [1, {upper}] "
                f"(deck has {total_slides} slides). LLM likely confused with shape_id."
            )

        # shape_id existence check — uses parser cache (populated during prefetch
        # or run()). Trial 3 showed gpt-4.1 occasionally emits shape_id that
        # belongs to a different slide; without this guard the COM error
        # cascades through rollback. Best-effort: skip if no parse available.
        shape_id = args.get("shape_id")
        parser = getattr(self, "_parser", None)
        if shape_id is not None and parser is not None:
            slide_data = parser.database.get(slide_number) if hasattr(parser, "database") else None
            if slide_data:
                valid_ids = _collect_shape_ids(slide_data.get("Objects_Detail", []))
                if valid_ids and shape_id not in valid_ids:
                    raise ValueError(
                        f"shape_id={shape_id} not found on slide {slide_number}. "
                        f"Available shape_ids: {sorted(valid_ids)[:15]}"
                        f"{'...' if len(valid_ids) > 15 else ''}. LLM likely targeted "
                        f"the wrong slide."
                    )

    def _rollback_ppt(self, stage, reason):
        logger.warning(f"{stage} Feedback: {reason}")
        # Layer 1: ordinary Close + Open.
        try:
            ppt_app = self.container.prs.Application
            self.container.prs.Close()
            time.sleep(0.5)
            self.container.prs = ppt_app.Presentations.Open(os.path.abspath(self.backup_path))
            return
        except Exception as e:
            logger.warning(f"Rollback (simple) failed during {stage}: {e}; trying emergency recovery")

        # Layer 2: kill PowerPoint and re-initialize from backup.
        try:
            from pathlib import Path as _Path
            from editppt.ppt_core import kill_powerpoint_processes, initialize_ppt
            kill_powerpoint_processes()
            time.sleep(1)
            new_prs, new_ppt_app = initialize_ppt(_Path(self.backup_path))
            self.container.prs = new_prs
            self.container.ppt_app = new_ppt_app
            logger.info(f"Rollback emergency recovery succeeded for {stage}")
        except Exception as ee:
            logger.error(
                f"Rollback emergency recovery also failed during {stage}: {ee}. "
                f"Subsequent COM operations may fail until the request aborts."
            )

    def _execute_tool(self, name, args):
        if name not in FUNCTION_MAP:
            raise ValueError(f"Tool '{name}' not found.")
        if name in self._AUTO_RESIZE_TOOLS:
            args["auto_resize"] = self._auto_resize
        result = FUNCTION_MAP[name](self.container.prs, **args)
        # Clamp to slide bounds after tool execution
        slide_number = args.get("slide_number")
        if slide_number is not None:
            if name in _TEXT_CLAMP_TOOLS:
                shape_id = args.get("shape_id")
                if shape_id is not None:
                    clamp_text_to_slide(self.container.prs, slide_number, shape_id)
            else:
                clamp_shapes_to_slide(self.container.prs, slide_number)
        return result


def create_specialist_agents(container, model: str) -> dict[str, "BaseEditAgent"]:
    """Factory that creates one BaseEditAgent per registered specialist."""
    agents = {}
    for agent_type, spec in AGENT_REGISTRY.items():
        agents[agent_type] = BaseEditAgent(
            container=container,
            model=model,
            agent_spec=spec,
        )
    return agents
