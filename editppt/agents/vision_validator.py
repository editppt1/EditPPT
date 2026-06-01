import json
import base64
import shutil
from concurrent.futures import ThreadPoolExecutor, Future

from editppt.utils.llm_client import call_llm, call_llm_gemini, set_token_log_context
from editppt.prompts import create_vision_validator_agent_system_prompt
from editppt.utils.logger_manual import init_logger
from editppt.utils.utils import slim_for_vision, scope_for_used_tools
from editppt.config import PROJECT_ROOT

logger = init_logger()


_ISSUE_PROPERTIES = {
    "IssueType": {
        "type": "string",
        "enum": ["TEXT_OVERFLOW", "ELEMENT_COLLISION", "ALIGNMENT_INCONSISTENCY"],
    },
    "AffectedShapeIDs": {"type": "array", "items": {"type": "integer"}},
    "TechnicalDiagnosis": {"type": "string"},
    "ActionableFix": {"type": "string"},
}
_ISSUE_REQUIRED = ["IssueType", "AffectedShapeIDs", "TechnicalDiagnosis", "ActionableFix"]

# Gemini accepts a plain JSON-schema-like dict; additionalProperties is unsupported.
GEMINI_VISION_SCHEMA = {
    "type": "object",
    "properties": {
        "HasCriticalIssues": {"type": "string", "enum": ["Yes", "No"]},
        "Issues": {
            "type": "array",
            "items": {
                "type": "object",
                "properties": _ISSUE_PROPERTIES,
                "required": _ISSUE_REQUIRED,
            },
        },
    },
    "required": ["HasCriticalIssues", "Issues"],
}

# OpenAI strict mode requires additionalProperties:false on every object.
OPENAI_VISION_SCHEMA = {
    "type": "object",
    "properties": {
        "HasCriticalIssues": {"type": "string", "enum": ["Yes", "No"]},
        "Issues": {
            "type": "array",
            "items": {
                "type": "object",
                "properties": _ISSUE_PROPERTIES,
                "required": _ISSUE_REQUIRED,
                "additionalProperties": False,
            },
        },
    },
    "required": ["HasCriticalIssues", "Issues"],
    "additionalProperties": False,
}


class VisionValidatorAgent:
    def __init__(self, container, model: str, executor: ThreadPoolExecutor | None = None):
        self.container = container
        self.model = model

        self.output_dir = PROJECT_ROOT / ".SlideScreenshots"
        if self.output_dir.exists():
            shutil.rmtree(self.output_dir)
        self.output_dir.mkdir(parents=True, exist_ok=True)

        # Shared executor for async LLM calls (Option A pipelining).
        # COM Export stays on the main thread; only the network call runs here.
        self._executor = executor or ThreadPoolExecutor(max_workers=4, thread_name_prefix="vision")

    @classmethod
    def create(cls, activate_valid: bool, container, model: str):
        if not activate_valid:
            return None
        return cls(container, model)

    def _export_slide_png(self, page_number: int) -> bytes:
        """Sync COM Export. Must run on the main thread that owns the COM apartment."""
        slide = self.container.prs.Slides(page_number)
        image_path = self.output_dir / f"slide_{page_number}.png"
        slide.Export(str(image_path), "PNG")
        with open(image_path, "rb") as f:
            return f.read()

    def _call_vision_llm(self, img_bytes: bytes, page_number: int,
                         agent_request, parsed_contents, used_tools):
        """Network-only: image bytes → (valid, reason, issues). Safe in a worker thread."""
        set_token_log_context(component="vision_validator", slide_index=page_number)
        compact_contents = scope_for_used_tools(
            slim_for_vision(parsed_contents),
            used_tools,
        )
        system_prompt = create_vision_validator_agent_system_prompt(
            agent_request,
            compact_contents,
            used_tools,
        )

        if self.model.startswith("gpt") or self.model.startswith("claude"):
            encoded_image = base64.b64encode(img_bytes).decode("utf-8")
            messages_for_validation = [{
                "role": "user",
                "content": [
                    {"type": "input_text", "text": system_prompt},
                    {"type": "input_image",
                     "image_url": f"data:image/png;base64,{encoded_image}"},
                ],
            }]
            response = call_llm(
                model=self.model,
                messages=messages_for_validation,
                text={
                    "format": {
                        "type": "json_schema",
                        "name": "vision_validation",
                        "schema": OPENAI_VISION_SCHEMA,
                        "strict": True,
                    }
                },
            )
            response_text = (response.output_text or "").strip()
        elif self.model.startswith("gemini"):
            response_text = call_llm_gemini(
                model=self.model,
                messages=system_prompt,
                image=img_bytes,
                response_schema=GEMINI_VISION_SCHEMA,
            )
            response_text = (response_text or "").strip()
        else:
            raise ValueError(f"Unsupported model for vision validation: {self.model}")

        try:
            data = json.loads(response_text)
        except json.JSONDecodeError as e:
            logger.error(f"Vision response not valid JSON despite schema: {response_text[:300]}")
            return False, f"Vision response JSON decode error: {e}", []

        if data.get("HasCriticalIssues") == "No":
            logger.info(f"Vision validation passed (slide {page_number}).")
            return True, None, []

        issues = data.get("Issues", []) or []
        if not issues:
            logger.info(f"Vision validation passed (slide {page_number}, empty Issues).")
            return True, None, []

        reason = " | ".join(
            f"Issue {i + 1}: {issue.get('TechnicalDiagnosis', 'Unknown issue')} "
            f"(Suggest: {issue.get('ActionableFix', '')})"
            for i, issue in enumerate(issues)
        )
        logger.warning(f"Vision validation failed (slide {page_number}): {reason}")
        return False, reason, issues

    def submit_async(self, page_number, agent_request, parsed_contents, used_tools) -> Future:
        """Export PNG synchronously, queue the LLM call.

        Returns a Future whose result is (valid, reason, issues).
        """
        img_bytes = self._export_slide_png(page_number)
        return self._executor.submit(
            self._call_vision_llm,
            img_bytes, page_number, agent_request, parsed_contents, used_tools,
        )

    def process(self, page_number, agent_request, parsed_contents, used_tools):
        """
        1. Export slide PNG
        2. Call Vision LLM with structured output (JSON schema enforced)
        3. Decide pass/fail from the parsed JSON

        Sync API kept for callers that don't pipeline. New code should prefer
        submit_async() so the LLM round trip overlaps with subsequent edits.
        """

        # --- Slide Export ---
        slide = self.container.prs.Slides(page_number)
        image_path = self.output_dir / f"slide_{page_number}.png"
        slide.Export(str(image_path), "PNG")

        with open(image_path, "rb") as image_file:
            img_bytes = image_file.read()

        set_token_log_context(component="vision_validator", slide_index=page_number)

        compact_contents = scope_for_used_tools(
            slim_for_vision(parsed_contents),
            used_tools,
        )

        system_prompt = create_vision_validator_agent_system_prompt(
            agent_request,
            compact_contents,
            used_tools,
        )

        # --- LLM Call (provider-specific structured output) ---
        # OpenAI and Anthropic both flow through call_llm: the Anthropic
        # adapter in llm_client.py converts `input_text`/`input_image` blocks
        # and turns the `text` json_schema spec into a forced tool call. The
        # schema (OPENAI_VISION_SCHEMA) uses standard JSON Schema features so
        # it works for both providers.
        if self.model.startswith("gpt") or self.model.startswith("claude"):
            encoded_image = base64.b64encode(img_bytes).decode("utf-8")
            messages_for_validation = [
                {
                    "role": "user",
                    "content": [
                        {"type": "input_text", "text": system_prompt},
                        {
                            "type": "input_image",
                            "image_url": f"data:image/png;base64,{encoded_image}",
                        },
                    ],
                }
            ]
            response = call_llm(
                model=self.model,
                messages=messages_for_validation,
                text={
                    "format": {
                        "type": "json_schema",
                        "name": "vision_validation",
                        "schema": OPENAI_VISION_SCHEMA,
                        "strict": True,
                    }
                },
            )
            response_text = (response.output_text or "").strip()

        elif self.model.startswith("gemini"):
            response_text = call_llm_gemini(
                model=self.model,
                messages=system_prompt,
                image=img_bytes,
                response_schema=GEMINI_VISION_SCHEMA,
            )
            response_text = (response_text or "").strip()

        else:
            raise ValueError(f"Unsupported model for vision validation: {self.model}")

        # --- Parse (schema is enforced, but guard against empty/edge cases) ---
        try:
            data = json.loads(response_text)
        except json.JSONDecodeError as e:
            logger.error(f"Vision response not valid JSON despite schema: {response_text[:300]}")
            return False, f"Vision response JSON decode error: {e}", []

        if data.get("HasCriticalIssues") == "No":
            logger.info("Vision validation passed.")
            return True, None, []

        issues = data.get("Issues", []) or []
        if not issues:
            logger.info("Vision validation passed (HasCriticalIssues=Yes but empty Issues).")
            return True, None, []

        reason = " | ".join(
            f"Issue {i + 1}: {issue.get('TechnicalDiagnosis', 'Unknown issue')} "
            f"(Suggest: {issue.get('ActionableFix', '')})"
            for i, issue in enumerate(issues)
        )
        logger.warning(f"Vision validation failed: {reason}")
        return False, reason, issues
