"""Mixpanel analytics for EditPPT.

Anonymous usage tracking — no personal information is collected.
Set MIXPANEL_TOKEN in .env to enable. If unset, all calls are no-ops.
"""

import os
import platform
import threading
import uuid
from pathlib import Path

from editppt.config import APP_VERSION, PROJECT_ROOT

# Telemetry is opt-in: set MIXPANEL_TOKEN in .env to enable. Unset → all no-ops.
_MIXPANEL_TOKEN = os.environ.get("MIXPANEL_TOKEN", "")

_mp = None
_distinct_id: str | None = None
_ID_FILE = PROJECT_ROOT / ".analytics_id"


def _get_or_create_id() -> str:
    """Persistent anonymous user ID (random UUID, stored locally)."""
    try:
        if _ID_FILE.exists():
            stored = _ID_FILE.read_text(encoding="utf-8").strip()
            if stored:
                return stored
    except Exception:
        pass
    new_id = uuid.uuid4().hex
    try:
        _ID_FILE.write_text(new_id, encoding="utf-8")
    except Exception:
        pass
    return new_id


def init_analytics():
    """Initialize Mixpanel client. Call once at app startup."""
    global _mp, _distinct_id
    if not _MIXPANEL_TOKEN:
        _mp = None
        return
    try:
        from mixpanel import Mixpanel
        _mp = Mixpanel(_MIXPANEL_TOKEN)
        _distinct_id = _get_or_create_id()
    except Exception:
        _mp = None


def _track(event: str, properties: dict | None = None):
    """Send event in background thread. No-op if analytics disabled."""
    if _mp is None or _distinct_id is None:
        return

    props = {
        "app_version": APP_VERSION,
        "os": platform.system(),
        "os_version": platform.version(),
    }
    if properties:
        props.update(properties)

    def _send():
        try:
            _mp.track(_distinct_id, event, props)
        except Exception:
            pass

    threading.Thread(target=_send, daemon=True).start()


def flush():
    """No-op. Kept for API compatibility (default Consumer sends immediately)."""
    pass


# --- Event helpers ---

def track_app_launched():
    _track("app_launched")


def track_file_uploaded(filename: str, slide_count: int):
    _track("file_uploaded", {"filename": filename, "slide_count": slide_count})


def track_edit_requested(user_input: str, task_count: int, auto_resize: bool):
    _track("edit_requested", {
        "user_input": user_input,
        "task_count": task_count,
        "auto_resize": auto_resize,
    })


def track_edit_completed(status: str, task_count: int,
                         input_tokens: int, output_tokens: int,
                         duration_seconds: float):
    _track("edit_completed", {
        "status": status,
        "task_count": task_count,
        "input_tokens": input_tokens,
        "output_tokens": output_tokens,
        "duration_seconds": round(duration_seconds, 1),
    })


def track_update_applied(from_version: str, to_version: str):
    _track("update_applied", {
        "from_version": from_version,
        "to_version": to_version,
    })
