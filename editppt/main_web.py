import json
import os
import shutil
import time
import sys
import argparse
import threading
import queue
from pathlib import Path
from dataclasses import dataclass, field

import pythoncom
import requests as http_requests
from loguru import logger
from flask import Flask, render_template, request, jsonify, Response, send_from_directory


from editppt.ppt_core import (
    PPTContainer, kill_powerpoint_processes, initialize_ppt,
    export_slide_image, init_backup,
)
from editppt.utils.logger_manual import init_logger, get_dynamic_log_dir, log_path
from editppt.agents import DispatcherAgent, create_specialist_agents, VisionValidatorAgent, VisualFixerAgent
from editppt.agents.base_agent import ALL_TOOLS_SCHEMA
from editppt.parser import Parser
from editppt.planner import Planner
from editppt.config import CURRENT_MODEL_NAME, CURRENT_VISION_MODEL_NAME, UPLOADS_DIR, PROJECT_ROOT, RESOURCE_ROOT, APP_VERSION, LOG_BASE, UPDATE_REPO
from editppt.utils.llm_client import (
    reset_token_counter,
    get_token_count,
    get_token_snapshot,
    diff_tokens,
    set_token_log_path,
    set_token_log_context,
    set_request_summary_path,
    log_request_summary,
)
from editppt.updater import (
    check_for_update, download_and_prepare_update, launch_update_script,
    cleanup_stale_update_artifacts, acquire_single_instance_mutex,
)
from editppt import analytics
from dotenv import load_dotenv


def _cleanup_logfiles():
    """Remove all contents inside the logfiles directory (exe builds only)."""
    if not getattr(sys, "frozen", False):
        return
    if not LOG_BASE.exists():
        return
    for child in LOG_BASE.iterdir():
        try:
            if child.is_dir():
                shutil.rmtree(child, ignore_errors=True)
            else:
                child.unlink(missing_ok=True)
        except Exception:
            pass


# ===============================
# Data model
# ===============================
@dataclass
class SlideEdit:
    slide_index: int
    before_image: str | None = None
    after_image: str | None = None
    status: str = "pending"
    checkpoint_path: str | None = None
    task_data: list = field(default_factory=list)   # list of planner tasks for this slide
    changed: bool = True


@dataclass
class RequestRecord:
    request_id: int
    user_input: str
    timestamp: float
    slide_edits: list[SlideEdit] = field(default_factory=list)
    status: str = "pending"
    total_input_tokens: int = 0
    total_output_tokens: int = 0
    total_cached_input_tokens: int = 0
    total_cost_usd: float = 0.0
    auto_resize: bool = False


# ===============================
# Global state
# ===============================
if getattr(sys, 'frozen', False):
    _template_dir = str(RESOURCE_ROOT / "editppt" / "templates")
    app = Flask("editppt.main_web", template_folder=_template_dir)
else:
    app = Flask("editppt.main_web")

history: list[RequestRecord] = []
request_counter = 0
is_processing = False
processing_lock = threading.Lock()

# Abort support
abort_event = threading.Event()

# SSE broadcasting
sse_clients: list[queue.Queue] = []
sse_clients_lock = threading.Lock()

# Queue for sending work to the COM thread
work_queue: queue.Queue = queue.Queue()

# Info populated by the COM thread after init, used read-only by Flask
prs_name: str = ""
slide_count: int = 0
image_root: Path | None = None
log_root: Path | None = None
file_loaded: bool = False
checkpoint_dir: Path | None = None

# Event signalling COM thread finished init
com_ready = threading.Event()

# Upload directory
uploads_dir: Path = UPLOADS_DIR

# Current slide context for log filtering (written by COM thread only)
_current_context: dict = {"request_id": None, "slide_index": None}

# Per-request log buffer for bug reports
_request_logs: dict[int, list[str]] = {}


# ===============================
# SSE helpers
# ===============================
def broadcast_sse(event: str, data: dict | str):
    """Send an SSE event to all connected clients."""
    if isinstance(data, dict):
        data = json.dumps(data, ensure_ascii=False)
    msg = f"event: {event}\ndata: {data}\n\n"
    with sse_clients_lock:
        dead = []
        for q in sse_clients:
            try:
                q.put_nowait(msg)
            except queue.Full:
                dead.append(q)
        for q in dead:
            sse_clients.remove(q)


def _loguru_sse_sink(message):
    """Custom loguru sink that broadcasts log lines via SSE."""
    record = message.record
    ts = record["time"].strftime("%H:%M:%S")
    level = record["level"].name
    text = record["message"]
    line = f"{ts} {level:>8} {text}"
    rid = _current_context["request_id"]
    broadcast_sse("log", {
        "text": line,
        "request_id": rid,
        "slide_index": _current_context["slide_index"],
    })
    # Buffer logs per request for bug reports
    if rid is not None:
        _request_logs.setdefault(rid, []).append(line)


# ===============================
# API endpoints
# ===============================
@app.route("/")
def index():
    return render_template("index.html", prs_name=prs_name, file_loaded=file_loaded, app_version=APP_VERSION)


@app.route("/api/upload", methods=["POST"])
def api_upload():
    global file_loaded

    if file_loaded:
        return jsonify({"error": "A file is already loaded"}), 409

    if "file" not in request.files:
        return jsonify({"error": "No file provided"}), 400

    f = request.files["file"]
    if not f.filename or not f.filename.lower().endswith(".pptx"):
        return jsonify({"error": "Only .pptx files are accepted"}), 400

    # secure_filename strips non-ASCII (Korean etc.) — sanitize manually
    raw_name = f.filename.replace("/", "_").replace("\\", "_").replace("\0", "")
    filename = raw_name.strip(". ")
    if not filename:
        filename = "upload.pptx"
    uploads_dir.mkdir(parents=True, exist_ok=True)
    save_path = (uploads_dir / filename).resolve()
    try:
        f.save(str(save_path))
    except PermissionError:
        logger.warning("Upload PermissionError — killing lingering PowerPoint and retrying")
        kill_powerpoint_processes()
        time.sleep(1)
        f.stream.seek(0)
        f.save(str(save_path))

    try:
        _start_com_worker(save_path)
    except Exception as e:
        logger.error(f"Upload failed: {e}")
        return jsonify({"error": str(e)}), 500

    analytics.track_file_uploaded(prs_name, slide_count)

    return jsonify({"prs_name": prs_name})


@app.route("/api/save_as", methods=["POST"])
def api_save_as():
    """Save a copy of the current presentation with a new filename."""
    with processing_lock:
        if is_processing:
            return jsonify({"error": "Cannot save while a request is being processed"}), 409

    if not file_loaded:
        return jsonify({"error": "No file is loaded"}), 400

    body = request.get_json(silent=True) or {}
    new_filename = body.get("filename", "").strip()
    if not new_filename:
        return jsonify({"error": "Filename is required"}), 400

    if not new_filename.lower().endswith(".pptx"):
        new_filename += ".pptx"

    new_filename = new_filename.replace("/", "_").replace("\\", "_").replace("\0", "").strip(". ")
    if not new_filename:
        return jsonify({"error": "Invalid filename"}), 400

    result_event = threading.Event()
    result_holder = {}

    work_queue.put({
        "_type": "save_as",
        "filename": new_filename,
        "_result_event": result_event,
        "_result_holder": result_holder,
    })

    result_event.wait(timeout=15)
    if result_holder.get("error"):
        return jsonify({"error": result_holder["error"]}), 500

    return jsonify({"ok": True, "saved_path": result_holder.get("saved_path", "")})


@app.route("/api/close", methods=["POST"])
def api_close():
    global file_loaded, prs_name, slide_count, image_root, log_root
    global history, request_counter, is_processing

    with processing_lock:
        if is_processing:
            return jsonify({"error": "Cannot close while a request is being processed"}), 409

    if not file_loaded:
        return jsonify({"error": "No file is loaded"}), 400

    body = request.get_json(silent=True) or {}
    save = body.get("save", True)

    if save:
        # Save mode: keep edits, just close
        work_queue.put(None)
    else:
        # Discard mode: restore pristine snapshot before closing
        work_queue.put({"_type": "close_discard"})

    time.sleep(0.5)
    kill_powerpoint_processes()

    # Clean up logfiles (exe only)
    _cleanup_logfiles()

    # Reset state
    prs_name = ""
    slide_count = 0
    image_root = None
    log_root = None
    file_loaded = False
    history = []
    request_counter = 0

    abort_event.clear()

    logger.info("File closed (save=%s), ready for new upload", save)
    return jsonify({"ok": True})


@app.route("/api/shutdown", methods=["POST"])
def api_shutdown():
    """Gracefully shut down the server."""
    global file_loaded

    with processing_lock:
        if is_processing:
            return jsonify({"error": "Cannot shut down while a request is being processed"}), 409

    logger.info("Server shutdown requested by user")

    # Close file if loaded
    if file_loaded:
        work_queue.put(None)  # signal COM worker to close
        time.sleep(0.5)
        kill_powerpoint_processes()
        file_loaded = False

    # Clean up logfiles (exe only)
    _cleanup_logfiles()

    # Flush analytics before exit
    analytics.flush()

    # Schedule server shutdown
    def _shutdown():
        time.sleep(0.5)
        os._exit(0)

    threading.Thread(target=_shutdown, daemon=True).start()
    return jsonify({"ok": True})


@app.route("/api/release-notes")
def api_release_notes():
    """Return release notes for the current version from GitHub."""
    try:
        resp = http_requests.get(
            f"https://api.github.com/repos/{UPDATE_REPO}/releases/tags/{APP_VERSION}",
            timeout=5,
        )
        resp.raise_for_status()
        data = resp.json()
        return jsonify({
            "version": APP_VERSION,
            "release_notes": data.get("body", "") or "",
        })
    except Exception:
        return jsonify({
            "version": APP_VERSION,
            "release_notes": "",
        })


@app.route("/api/check-update")
def api_check_update():
    """Check if a new version is available on GitHub."""
    info = check_for_update()
    return jsonify({
        "has_update": info["has_update"],
        "current_version": info["current"],
        "latest_version": info["latest"],
        "release_notes": info["release_notes"],
    })


@app.route("/api/apply-update", methods=["POST"])
def api_apply_update():
    """Download and apply an update, then restart."""
    global file_loaded

    if not getattr(sys, 'frozen', False):
        return jsonify({"error": "Auto-update only works in packaged mode"}), 400

    with processing_lock:
        if is_processing:
            return jsonify({"error": "Cannot update while a request is being processed"}), 409

    def _run_update():
        global file_loaded
        try:
            def _progress(stage: str, pct: int):
                broadcast_sse("update_progress", {"stage": stage, "percent": pct})

            # Read current port so the new exe restarts on the same port
            try:
                _cur_port = int((PROJECT_ROOT / ".server_port").read_text().strip())
            except Exception:
                _cur_port = 5000

            # Download and extract (bat script is NOT launched yet)
            bat_path = download_and_prepare_update(progress_callback=_progress, port=_cur_port)

            # Save release notes for post-update display
            try:
                info = check_for_update()
                if info.get("release_notes"):
                    marker = PROJECT_ROOT.parent / "_update_complete.md"
                    marker.write_text(
                        f"{info['latest']}\n{info['release_notes']}",
                        encoding="utf-8",
                    )
            except Exception:
                pass

            # Close PowerPoint before launching the swap script
            if file_loaded:
                work_queue.put(None)
                time.sleep(0.5)
                kill_powerpoint_processes()
                file_loaded = False

            broadcast_sse("update_progress", {"stage": "restarting", "percent": 100})

            # Launch the bat script — it will taskkill this process,
            # swap directories, and restart the new exe.
            # subprocess.run blocks until wscript finishes launching the bat.
            launch_update_script(bat_path)
            # bat will kill us; fallback exit if it doesn't within 10s
            time.sleep(10)
            os._exit(0)
        except Exception as e:
            logger.error(f"Update failed: {e}")
            broadcast_sse("update_error", {"error": str(e)})
            return

    threading.Thread(target=_run_update, daemon=True).start()
    return jsonify({"ok": True})


@app.route("/api/abort", methods=["POST"])
def api_abort():
    if not is_processing:
        return jsonify({"error": "Nothing is running"}), 400
    abort_event.set()
    logger.info("Abort requested by user")
    return jsonify({"ok": True})


@app.route("/api/rollback", methods=["POST"])
def api_rollback():
    global is_processing

    body = request.get_json(silent=True) or {}
    request_id = body.get("request_id")
    step_index = body.get("step_index")

    if request_id is None or step_index is None:
        return jsonify({"error": "request_id and step_index required"}), 400

    with processing_lock:
        if is_processing:
            return jsonify({"error": "Cannot rollback while processing"}), 409

    # Find the target record and edit
    target_rec = None
    for r in history:
        if r.request_id == request_id:
            target_rec = r
            break
    if not target_rec:
        return jsonify({"error": "Request not found"}), 404
    if step_index < 0 or step_index >= len(target_rec.slide_edits):
        return jsonify({"error": "Invalid step_index"}), 400

    target_edit = target_rec.slide_edits[step_index]
    if not target_edit.checkpoint_path:
        return jsonify({"error": "No checkpoint for this slide"}), 400

    with processing_lock:
        is_processing = True

    # Send rollback job to COM worker
    work_queue.put({
        "_type": "rollback",
        "request_id": request_id,
        "step_index": step_index,
        "checkpoint_path": target_edit.checkpoint_path,
    })
    return jsonify({"ok": True})


@app.route("/api/retry", methods=["POST"])
def api_retry():
    global is_processing

    body = request.get_json(silent=True) or {}
    request_id = body.get("request_id")
    step_index = body.get("step_index")

    if request_id is None or step_index is None:
        return jsonify({"error": "request_id and step_index required"}), 400

    with processing_lock:
        if is_processing:
            return jsonify({"error": "Cannot retry while processing"}), 409

    target_rec = None
    for r in history:
        if r.request_id == request_id:
            target_rec = r
            break
    if not target_rec:
        return jsonify({"error": "Request not found"}), 404
    if step_index < 0 or step_index >= len(target_rec.slide_edits):
        return jsonify({"error": "Invalid step_index"}), 400

    target_edit = target_rec.slide_edits[step_index]
    if not target_edit.task_data:
        return jsonify({"error": "No task data for this slide"}), 400

    with processing_lock:
        is_processing = True

    auto_resize = bool(body.get("auto_resize", False))

    work_queue.put({
        "_type": "retry",
        "request_id": request_id,
        "step_index": step_index,
        "auto_resize": auto_resize,
        "submit_ts": time.time(),
    })
    return jsonify({"ok": True})


@app.route("/api/rollback-request", methods=["POST"])
def api_rollback_request():
    global is_processing

    body = request.get_json(silent=True) or {}
    request_id = body.get("request_id")

    if request_id is None:
        return jsonify({"error": "request_id required"}), 400

    with processing_lock:
        if is_processing:
            return jsonify({"error": "Cannot rollback while processing"}), 409

    # Find the target request
    target_rec = None
    for r in history:
        if r.request_id == request_id:
            target_rec = r
            break
    if not target_rec:
        return jsonify({"error": "Request not found"}), 404
    if not target_rec.slide_edits:
        return jsonify({"error": "No slide edits in this request"}), 400

    # Use the first slide edit's checkpoint (step_index=0)
    first_edit = target_rec.slide_edits[0]
    if not first_edit.checkpoint_path:
        return jsonify({"error": "No checkpoint for this request"}), 400

    with processing_lock:
        is_processing = True

    work_queue.put({
        "_type": "rollback",
        "request_id": request_id,
        "step_index": 0,
        "checkpoint_path": first_edit.checkpoint_path,
    })
    return jsonify({"ok": True})


@app.route("/api/retry-request", methods=["POST"])
def api_retry_request():
    global is_processing

    body = request.get_json(silent=True) or {}
    request_id = body.get("request_id")

    if request_id is None:
        return jsonify({"error": "request_id required"}), 400

    with processing_lock:
        if is_processing:
            return jsonify({"error": "Cannot retry while processing"}), 409

    # Find the target request
    target_rec = None
    for r in history:
        if r.request_id == request_id:
            target_rec = r
            break
    if not target_rec:
        return jsonify({"error": "Request not found"}), 404
    if not target_rec.slide_edits:
        return jsonify({"error": "No slide edits in this request"}), 400

    first_edit = target_rec.slide_edits[0]
    if not first_edit.checkpoint_path:
        return jsonify({"error": "No checkpoint for this request"}), 400

    with processing_lock:
        is_processing = True

    auto_resize = bool(body.get("auto_resize", False))

    work_queue.put({
        "_type": "retry_request",
        "original_request_id": request_id,
        "checkpoint_path": first_edit.checkpoint_path,
        "auto_resize": auto_resize,
        "submit_ts": time.time(),
    })
    return jsonify({"ok": True})


@app.route("/api/submit", methods=["POST"])
def api_submit():
    global request_counter, is_processing

    if not file_loaded:
        return jsonify({"error": "No file loaded yet"}), 400

    body = request.get_json(silent=True) or {}
    user_input = (body.get("user_input") or "").strip()
    if not user_input:
        return jsonify({"error": "Empty input"}), 400

    auto_resize = bool(body.get("auto_resize", False))

    with processing_lock:
        if is_processing:
            return jsonify({"error": "A request is already being processed"}), 409
        is_processing = True

    request_counter += 1
    rec = RequestRecord(
        request_id=request_counter,
        user_input=user_input,
        timestamp=time.time(),
        status="running",
        auto_resize=auto_resize,
    )
    history.append(rec)
    broadcast_sse("request_update", _serialize_record(rec))

    # Send work to the dedicated COM thread
    work_queue.put(rec)
    return jsonify({"request_id": rec.request_id})


@app.route("/api/history")
def api_history():
    return jsonify([_serialize_record(r) for r in history])


@app.route("/api/request/<int:rid>")
def api_request(rid):
    for r in history:
        if r.request_id == rid:
            return jsonify(_serialize_record(r))
    return jsonify({"error": "Not found"}), 404


@app.route("/api/stream")
def api_stream():
    q = queue.Queue(maxsize=500)
    with sse_clients_lock:
        sse_clients.append(q)

    def generate():
        try:
            while True:
                try:
                    msg = q.get(timeout=30)
                    yield msg
                except queue.Empty:
                    yield ": keepalive\n\n"
        except GeneratorExit:
            pass
        finally:
            with sse_clients_lock:
                if q in sse_clients:
                    sse_clients.remove(q)

    return Response(generate(), mimetype="text/event-stream",
                    headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"})


# ===============================
# Bug report / feedback
# ===============================
_CATEGORY_COLORS = {
    "Bug - Error occurred": 0xFF4444,
    "Bug - Incorrect edit": 0xFF4444,
    "Bug - UI issue": 0xFF4444,
    "Improvement": 0x4488FF,
    "Other": 0x888888,
}


def _send_discord_bug_report(
    webhook_url: str,
    title: str,
    category: str,
    description: str,
    user_request: str | None,
    options_text: str | None,
    image_files: list[Path],
    log_text: str | None,
    feedback_texts: list[tuple[str, str]],
):
    """Send a bug report to Discord via webhook with embed + file attachments.

    Discord limits: 10 files and 25 MB per message. If images exceed the
    first message's capacity, additional messages are sent automatically.
    """
    _DISCORD_MAX_FILES = 10
    _DISCORD_MAX_BYTES = 25 * 1024 * 1024  # 25 MB

    embed = {
        "title": title[:256],
        "description": description[:4096] if description else f"[{category}]",
        "color": _CATEGORY_COLORS.get(category, 0x888888),
        "fields": [{"name": "Category", "value": category or "Other", "inline": True}],
    }
    if user_request:
        embed["fields"].append({"name": "User Request", "value": user_request[:1024]})
    if options_text:
        embed["fields"].append({"name": "Options", "value": options_text, "inline": True})

    # Discord embed total char limit: 6000
    total = len(embed.get("title", "")) + len(embed.get("description", ""))
    for f in embed.get("fields", []):
        total += len(f.get("name", "")) + len(f.get("value", ""))
    if total > 5900:
        embed["fields"] = embed["fields"][:2]

    # --- Build non-image attachments (log + feedback) ---
    text_files: list[tuple[str, str, bytes, str]] = []  # (key, filename, data, mime)
    if log_text:
        text_files.append(("log", "logs.txt", log_text.encode("utf-8"), "text/plain"))
    for fname, content in feedback_texts:
        text_files.append((f"feedback_{fname}", f"{fname}.json",
                           content.encode("utf-8"), "application/json"))

    # --- Read all image data upfront ---
    image_data: list[tuple[str, bytes]] = []  # (filename, data)
    for img_path in image_files:
        try:
            image_data.append((img_path.name, img_path.read_bytes()))
        except Exception:
            pass

    # --- Chunk images respecting file count and size limits ---
    # First message: embed + text_files + as many images as fit
    first_msg_file_slots = _DISCORD_MAX_FILES - len(text_files)
    first_msg_size = sum(len(d) for _, _, d, _ in text_files)

    first_batch: list[tuple[str, bytes]] = []
    remaining: list[tuple[str, bytes]] = []
    for img_name, img_bytes in image_data:
        fits_count = len(first_batch) < first_msg_file_slots
        fits_size = first_msg_size + len(img_bytes) < _DISCORD_MAX_BYTES
        if fits_count and fits_size:
            first_batch.append((img_name, img_bytes))
            first_msg_size += len(img_bytes)
        else:
            remaining.append((img_name, img_bytes))

    # --- Send first message: embed + text attachments + first batch of images ---
    multipart: dict[str, tuple] = {}
    if remaining:
        embed["footer"] = {"text": f"📎 {len(image_data)} images attached across multiple messages"}
    payload = {"embeds": [embed]}
    multipart["payload_json"] = (None, json.dumps(payload), "application/json")
    for key, filename, data, mime in text_files:
        multipart[key] = (filename, data, mime)
    for i, (img_name, img_bytes) in enumerate(first_batch):
        multipart[f"image{i}"] = (img_name, img_bytes, "image/png")

    r = http_requests.post(webhook_url, files=multipart)
    r.raise_for_status()

    # --- Send follow-up messages for remaining images ---
    while remaining:
        chunk: list[tuple[str, bytes]] = []
        chunk_size = 0
        while remaining and len(chunk) < _DISCORD_MAX_FILES:
            img_name, img_bytes = remaining[0]
            if chunk_size + len(img_bytes) >= _DISCORD_MAX_BYTES and chunk:
                break
            chunk.append(remaining.pop(0))
            chunk_size += len(img_bytes)

        follow_embed = {
            "title": title[:256],
            "description": "*(continued)*",
            "color": _CATEGORY_COLORS.get(category, 0x888888),
        }
        follow_multipart: dict[str, tuple] = {
            "payload_json": (None, json.dumps({"embeds": [follow_embed]}),
                             "application/json"),
        }
        for i, (img_name, img_bytes) in enumerate(chunk):
            follow_multipart[f"image{i}"] = (img_name, img_bytes, "image/png")

        r = http_requests.post(webhook_url, files=follow_multipart)
        r.raise_for_status()


@app.route("/api/bug-report", methods=["POST"])
def api_bug_report():
    # Bug reporting is opt-in: set BUG_REPORT_WEBHOOK_URL in .env to enable.
    webhook_url = os.environ.get("BUG_REPORT_WEBHOOK_URL", "")
    if not webhook_url:
        return jsonify({"error": "Bug reporting is not configured"}), 503

    body = request.get_json(silent=True) or {}
    title = (body.get("title") or "").strip()
    if not title:
        return jsonify({"error": "Title is required"}), 400

    category = body.get("category", "Other")
    description = (body.get("description") or "").strip()
    include_context = body.get("include_context", False)
    request_id = body.get("request_id")

    user_request = None
    options_text = None
    image_files: list[Path] = []
    log_text = None
    feedback_texts: list[tuple[str, str]] = []

    if include_context and request_id is not None:
        target_rec = None
        for r in history:
            if r.request_id == request_id:
                target_rec = r
                break

        if target_rec:
            user_request = target_rec.user_input
            options_text = f"Auto-resize: {'ON' if target_rec.auto_resize else 'OFF'}"

            # Before/after images (all slides)
            if image_root:
                for se in target_rec.slide_edits:
                    for img_key in ("before_image", "after_image"):
                        rel = getattr(se, img_key, None)
                        if not rel:
                            continue
                        img_path = image_root / rel
                        if img_path.exists():
                            image_files.append(img_path)

            # Log output
            log_lines = _request_logs.get(request_id, [])
            if log_lines:
                log_text = "\n".join(log_lines)

            # Agent feedback (only non-empty)
            if log_root and log_root.exists():
                for f in sorted(log_root.glob("agent_Feedback_*.json")):
                    try:
                        content = f.read_text(encoding="utf-8").strip()
                        parsed = json.loads(content)
                        if not parsed:
                            continue
                        feedback_texts.append((f.stem, content))
                    except Exception:
                        pass

    try:
        _send_discord_bug_report(
            webhook_url, title, category, description,
            user_request, options_text, image_files, log_text, feedback_texts,
        )
        return jsonify({"success": True})
    except http_requests.HTTPError as e:
        msg = str(e)
        try:
            msg = e.response.text[:300]
        except Exception:
            pass
        return jsonify({"error": f"Discord API error: {msg}"}), 500
    except Exception as e:
        return jsonify({"error": str(e)}), 500


# ===============================
# API Key Settings
# ===============================
_ENV_KEY_NAMES = [
    "OPENAI_API_KEY",
    "GEMINI_API_KEY",
    "PEXELS_API_KEY",
    "UNSPLASH_ACCESS_KEY",
]

def _mask_key(value: str) -> str:
    """Return masked preview of an API key, showing only last 4 chars."""
    if not value:
        return ""
    if len(value) <= 8:
        return "****" + value[-2:]
    return value[:4] + "****" + value[-4:]


@app.route("/api/env-status")
def env_status():
    keys = {}
    for name in _ENV_KEY_NAMES:
        val = os.environ.get(name, "")
        keys[name] = {
            "set": bool(val),
            "preview": _mask_key(val) if val else "",
        }
    return jsonify({"keys": keys})


@app.route("/api/env-keys", methods=["POST"])
def env_keys():
    data = request.get_json(force=True)
    new_keys = data.get("keys", {})
    if not new_keys:
        return jsonify({"error": "No keys provided"}), 400

    # --- 1. Read existing .env, merge new keys, write back ---
    from editppt.config import PROJECT_ROOT
    env_path = PROJECT_ROOT / ".env"

    existing_lines = []
    if env_path.exists():
        existing_lines = env_path.read_text(encoding="utf-8").splitlines()

    # Parse existing key=value pairs preserving order
    updated_keys = set()
    out_lines = []
    for line in existing_lines:
        stripped = line.strip()
        if stripped and not stripped.startswith("#") and "=" in stripped:
            key = stripped.split("=", 1)[0].strip()
            if key in new_keys and new_keys[key]:
                out_lines.append(f'{key}={new_keys[key]}')
                updated_keys.add(key)
                continue
        out_lines.append(line)

    # Append new keys not already in the file
    for key, val in new_keys.items():
        if key not in updated_keys and val:
            out_lines.append(f'{key}={val}')

    env_path.write_text("\n".join(out_lines) + "\n", encoding="utf-8")

    # --- 2. Update os.environ ---
    for key, val in new_keys.items():
        if val:
            os.environ[key] = val

    # --- 3. Reload module-level cached globals in llm_client ---
    import editppt.utils.llm_client as _llm_mod
    _llm_mod.OPENAI_API_KEY = os.environ.get("OPENAI_API_KEY")
    _llm_mod.GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY")

    # Also reload dotenv for any modules that read at import time
    load_dotenv(PROJECT_ROOT / ".env", override=True)

    return jsonify({"status": "ok"})


@app.route("/images/<path:filepath>")
def serve_image(filepath):
    return send_from_directory(str(image_root), filepath)


# ===============================
# Temp file cleanup
# ===============================
def _cleanup_temp_files(container, ckpt_dir):
    """Remove backup file, pristine snapshot, and checkpoint directory on close."""
    # Remove backup file (e.g. test_.pptx)
    try:
        bp = getattr(container, "backup_path", None)
        if bp and os.path.exists(bp):
            os.remove(bp)
            logger.info(f"Removed backup file: {bp}")
    except Exception as e:
        logger.warning(f"Failed to remove backup file: {e}")

    # Remove pristine snapshot (e.g. test_pristine.pptx)
    try:
        pp = getattr(container, "pristine_path", None)
        if pp and os.path.exists(pp):
            os.remove(pp)
            logger.info(f"Removed pristine file: {pp}")
    except Exception as e:
        logger.warning(f"Failed to remove pristine file: {e}")

    # Remove checkpoint directory
    try:
        if ckpt_dir and ckpt_dir.exists():
            shutil.rmtree(ckpt_dir)
            logger.info(f"Removed checkpoint directory: {ckpt_dir}")
    except Exception as e:
        logger.warning(f"Failed to remove checkpoint directory: {e}")


# ===============================
# COM worker thread
# ===============================
def _com_worker(ppt_path: Path):
    """
    Dedicated thread that owns all COM objects.
    Initialises PowerPoint, pipeline components, then loops processing requests.
    """
    global prs_name, slide_count, image_root, log_root, checkpoint_dir

    pythoncom.CoInitialize()
    try:
        # --- Initialise COM & pipeline on THIS thread ---
        prs, ppt_app = initialize_ppt(ppt_path)
        container = PPTContainer(prs, ppt_app)
        init_logger(container)

        log_root = get_dynamic_log_dir(container)
        log_root.mkdir(parents=True, exist_ok=True)

        # Per-call token usage persistence — survives crashes, abort, kill
        set_token_log_path(log_root / "token_usage.jsonl")
        set_request_summary_path(log_root / "request_summary.jsonl")
        reset_token_counter()

        image_root = log_root / "slide_images"
        image_root.mkdir(parents=True, exist_ok=True)

        checkpoint_dir = log_root / "checkpoints"
        checkpoint_dir.mkdir(parents=True, exist_ok=True)

        prs_name = container.prs.Name
        slide_count = len(container.prs.Slides)

        logger.info(f"Loaded [{prs_name}] with {slide_count} slides")

        planner = Planner(model=CURRENT_MODEL_NAME, slide_name=prs_name)
        logger.info("Planner initialized")

        parser_obj = Parser(container=container, total_slides=slide_count)
        logger.info("Parser initialized")

        (log_root / "parser_Database.json").write_text(
            json.dumps(parser_obj.database, ensure_ascii=False, indent=4),
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
            activate_valid=False,
            container=container,
            model=CURRENT_VISION_MODEL_NAME,
        )
        visual_fixer_agent = VisualFixerAgent.create(
            container=container,
            model=CURRENT_MODEL_NAME,
            tools_schema=ALL_TOOLS_SCHEMA,
        ) if vision_validator_agent is not None else None
        logger.info("Agents initialized (dispatcher + specialists)")

        # Add SSE log sink
        logger.add(_loguru_sse_sink, level="DEBUG", format="{message}")

        # Signal Flask that we're ready
        com_ready.set()

        # --- Event loop: process requests forever ---
        while True:
            item = work_queue.get()
            if item is None:
                # Cleanup temp files before exit
                _cleanup_temp_files(container, checkpoint_dir)
                break  # poison pill

            # Close-discard: restore pristine snapshot then exit
            if isinstance(item, dict) and item.get("_type") == "close_discard":
                logger.info("Close with discard: restoring pristine snapshot")
                try:
                    original_path = getattr(container, "original_path", None)
                    pristine_path = getattr(container, "pristine_path", None)
                    if pristine_path and original_path and os.path.exists(pristine_path):
                        try:
                            container.prs.Close()
                        except Exception:
                            pass
                        shutil.copy2(pristine_path, original_path)
                        logger.info("Original file restored from pristine snapshot")
                    else:
                        logger.warning("Pristine snapshot not available, closing without restore")
                except Exception as e:
                    logger.error(f"Failed to restore pristine snapshot: {e}")
                _cleanup_temp_files(container, checkpoint_dir)
                break

            # Save-as: save a copy with a new filename
            if isinstance(item, dict) and item.get("_type") == "save_as":
                result_event = item["_result_event"]
                result_holder = item["_result_holder"]
                try:
                    original_dir = os.path.dirname(container.original_path)
                    save_path = os.path.join(original_dir, item["filename"])
                    container.prs.SaveCopyAs(save_path)
                    result_holder["saved_path"] = save_path
                    logger.info(f"Save As completed: {save_path}")
                except Exception as e:
                    result_holder["error"] = str(e)
                    logger.error(f"Save As failed: {e}")
                finally:
                    result_event.set()
                continue

            # Special jobs (dict) vs normal request (RequestRecord)
            if isinstance(item, dict) and item.get("_type") == "rollback":
                try:
                    _execute_rollback(item, container, parser_obj)
                except Exception as e:
                    logger.error(f"Rollback failed: {e}")
                finally:
                    with processing_lock:
                        global is_processing
                        is_processing = False
                continue

            if isinstance(item, dict) and item.get("_type") == "retry":
                try:
                    _execute_retry(item, container, parser_obj,
                                   dispatcher, specialist_agents, vision_validator_agent,
                                   visual_fixer_agent)
                except Exception as e:
                    logger.error(f"Retry failed: {e}")
                finally:
                    with processing_lock:
                        is_processing = False
                continue

            if isinstance(item, dict) and item.get("_type") == "retry_request":
                try:
                    _execute_retry_request(item, container, parser_obj, planner,
                                           dispatcher, specialist_agents,
                                           vision_validator_agent, visual_fixer_agent, log_root)
                except Exception as e:
                    logger.error(f"Retry-request failed: {e}")
                finally:
                    with processing_lock:
                        is_processing = False
                continue

            rec = item
            try:
                _run_pipeline(rec, container, planner, parser_obj,
                              dispatcher, specialist_agents, vision_validator_agent,
                              visual_fixer_agent, log_root)
            except Exception as e:
                logger.error(f"Request #{rec.request_id} failed: {e}")
            finally:
                with processing_lock:
                    is_processing = False

    except BaseException as e:
        logger.error(f"COM worker init failed: {e}")
        com_ready.set()  # unblock main even on failure
    finally:
        pythoncom.CoUninitialize()


def _run_pipeline(rec, container, planner, parser_obj,
                  dispatcher, specialist_agents, vision_validator_agent,
                  visual_fixer_agent, log_root):
    global _current_context
    abort_event.clear()
    # Snapshot the monotonic counter so per-request usage is computed as a
    # delta. This is robust to abort, exception, or background dispatcher
    # workers finishing after this function returns.
    token_snapshot_start = get_token_snapshot()
    _pipeline_start = time.time()
    _current_context = {"request_id": rec.request_id, "slide_index": None}
    set_token_log_context(request_id=rec.request_id, slide_index=None, scope="pipeline")
    logger.info(f"Processing request #{rec.request_id}: {rec.user_input}")
    plan_json: dict = {}
    try:
        plan_json = planner(
            user_input=rec.user_input,
            total_slide_numbers=container.prs.Slides.Count,
        )
        logger.info("Planner output received")

        (log_root / "planner.json").write_text(
            json.dumps(plan_json, ensure_ascii=False, indent=4), encoding="utf-8"
        )

        tasks = plan_json.get("tasks", [])
        if not tasks:
            logger.warning("Planner returned no tasks")
            rec.status = "done"
            return

        analytics.track_edit_requested(rec.user_input, len(tasks), rec.auto_resize)

        # Build slide edit entries (one per unique slide)
        slide_edit_map = {}  # slide_idx -> index in rec.slide_edits
        for task in tasks:
            slide_idx = task.get("page_number")
            if not slide_idx:
                continue
            if slide_idx not in slide_edit_map:
                se = SlideEdit(slide_index=slide_idx, status="pending")
                rec.slide_edits.append(se)
                slide_edit_map[slide_idx] = len(rec.slide_edits) - 1
        broadcast_sse("request_update", _serialize_record(rec))

        # Pre-compute dispatcher decisions for tasks whose target slide
        # already exists. Slides are parsed serially (COM-bound) but the
        # dispatcher LLM calls run concurrently. Cached decisions for slides
        # touched by an earlier task are skipped below in favor of a fresh
        # dispatch.
        prefetched_dispatch: dict[int, list[dict]] = (
            dispatcher.prefetch_dispatch_decisions(tasks, parser_obj) if len(tasks) > 1 else {}
        )
        modified_slides_for_dispatch: set[int] = set()

        # Execute each task
        for i, task in enumerate(tasks):
            # Abort check
            if abort_event.is_set():
                logger.warning(f"Abort: skipping remaining slides for request #{rec.request_id}")
                for se in rec.slide_edits:
                    if se.status not in ("done", "error"):
                        se.status = "skipped"
                rec.status = "aborted"
                return

            slide_idx = task.get("page_number")
            if not slide_idx:
                continue

            _current_context = {"request_id": rec.request_id, "slide_index": slide_idx}
            set_token_log_context(slide_index=slide_idx)

            se = rec.slide_edits[slide_edit_map[slide_idx]]
            se.task_data.append(task)
            is_first_task = (se.before_image is None and se.checkpoint_path is None)
            se.status = "running"
            broadcast_sse("request_update", _serialize_record(rec))

            try:
                slide_exists = slide_idx <= container.prs.Slides.Count
                img_dir = image_root / f"slide_{slide_idx}"

                # Before screenshot + checkpoint only on first task for this slide
                if is_first_task and slide_exists:
                    slide = container.prs.Slides(slide_idx)

                    before_path = export_slide_image(slide, img_dir, f"req{rec.request_id}_before")
                    se.before_image = f"slide_{slide_idx}/{before_path.name}"
                    broadcast_sse("slide_update", {
                        "request_id": rec.request_id,
                        "slide_index": slide_idx,
                        "before_image": se.before_image,
                    })

                    if checkpoint_dir:
                        se_idx = slide_edit_map[slide_idx]
                        ckpt_path = checkpoint_dir / f"checkpoint_req{rec.request_id}_step{se_idx}_slide{slide_idx}.pptx"
                        try:
                            container.prs.SaveCopyAs(str(ckpt_path))
                            se.checkpoint_path = str(ckpt_path)
                        except Exception as e:
                            logger.warning(f"Checkpoint save failed: {e}")

                # Run agent(s) via dispatcher — multi-agent routing by shape type
                if not slide_exists:
                    # Route directly to slide agent if slide doesn't exist
                    slide_contents = None
                    objects_detail = []
                    sub_tasks = [{"agent_type": "slide", "description": task.get("description", "")}]
                else:
                    cached_dispatch = prefetched_dispatch.get(i)
                    if cached_dispatch is not None and slide_idx not in modified_slides_for_dispatch:
                        sub_tasks = cached_dispatch
                    else:
                        slide_contents = parser_obj.process(slide_idx)
                        objects_detail = slide_contents.get("Objects_Detail", []) if slide_contents else []
                        sub_tasks = dispatcher.dispatch(task, objects_detail)
                logger.info(f"Dispatcher routed to: {[(s['agent_type'], s.get('shape_ids', [])) for s in sub_tasks]}")

                had_any_tool_calls = False
                for sub in sub_tasks:
                    agent_type = sub["agent_type"]
                    sub_task = {**task, "description": sub["description"]}
                    had = specialist_agents[agent_type].run(
                        task=sub_task,
                        parser=parser_obj,
                        vision_validator_agent=vision_validator_agent,
                        visual_fixer_agent=visual_fixer_agent,
                        auto_resize=rec.auto_resize,
                        shape_ids=sub.get("shape_ids"),
                    )
                    if had:
                        had_any_tool_calls = True
                if had_any_tool_calls:
                    se.changed = True
                modified_slides_for_dispatch.add(slide_idx)

                # After screenshot (always refresh — last task result is final)
                if slide_idx <= container.prs.Slides.Count:
                    slide = container.prs.Slides(slide_idx)
                    after_path = export_slide_image(slide, img_dir, f"req{rec.request_id}_after")
                    se.after_image = f"slide_{slide_idx}/{after_path.name}"
                broadcast_sse("slide_update", {
                    "request_id": rec.request_id,
                    "slide_index": slide_idx,
                    "before_image": se.before_image,
                    "after_image": se.after_image,
                })

            except Exception as e:
                logger.error(f"Slide {slide_idx} failed: {e}")
                se.changed = False
                se.status = "error"

            broadcast_sse("request_update", _serialize_record(rec))

        # Mark all non-error slide edits as done
        for se in rec.slide_edits:
            if se.status == "running":
                se.status = "done"
        rec.status = "done"

    except Exception as e:
        import traceback
        logger.error(f"Pipeline error: {e}\n{traceback.format_exc()}")
        rec.status = "error"

    finally:
        # Always collect token usage — abort, error, or success.
        # Delta against the start snapshot is robust to:
        #  - early returns / exceptions before this point
        #  - background dispatcher prefetch threads finishing late
        #  - any other LLM call invoked transitively during the request
        delta = diff_tokens(token_snapshot_start, get_token_snapshot())
        rec.total_input_tokens = delta["input_tokens"]
        rec.total_output_tokens = delta["output_tokens"]
        rec.total_cached_input_tokens = delta.get("cached_input_tokens", 0)
        rec.total_cost_usd = delta.get("cost_usd", 0.0)
        broadcast_sse("request_update", _serialize_record(rec))
        done_ts = time.time()
        log_request_summary(
            request_id=rec.request_id,
            user_input=rec.user_input,
            submit_ts=rec.timestamp,
            done_ts=done_ts,
            status=rec.status,
            scope="pipeline",
            task_count=sum(len(se.task_data) for se in rec.slide_edits),
            slide_count=len(rec.slide_edits),
            input_tokens=delta["input_tokens"],
            output_tokens=delta["output_tokens"],
            cached_input_tokens=delta.get("cached_input_tokens", 0),
            cost_usd=delta.get("cost_usd", 0.0),
            planner_task_mode=plan_json.get("task_mode"),
        )
        logger.info(
            f"Request #{rec.request_id} finished (status={rec.status}, "
            f"elapsed={done_ts - rec.timestamp:.3f}s, "
            f"tokens: {delta['input_tokens']}in / {delta['output_tokens']}out, "
            f"cached_in: {delta.get('cached_input_tokens', 0)}, "
            f"cost: ${delta.get('cost_usd', 0.0):.6f})"
        )
        analytics.track_edit_completed(
            status=rec.status,
            task_count=len(rec.slide_edits),
            input_tokens=rec.total_input_tokens,
            output_tokens=rec.total_output_tokens,
            duration_seconds=done_ts - _pipeline_start,
        )
        _current_context = {"request_id": None, "slide_index": None}
        set_token_log_context(request_id=None, slide_index=None, scope=None)


# ===============================
# Retry execution (runs on COM thread)
# ===============================
def _execute_retry(job: dict, container, parser_obj,
                   dispatcher, specialist_agents, vision_validator_agent,
                   visual_fixer_agent=None):
    """Re-run the agent on a specific slide, optionally restoring from checkpoint first."""
    global _current_context
    target_rid = job["request_id"]
    target_step = job["step_index"]
    submit_ts = job.get("submit_ts", time.time())

    # Snapshot taken before any work so token attribution survives early
    # returns, checkpoint-restore failures, or any unhandled exception below.
    token_snapshot_start = get_token_snapshot()
    target_rec: RequestRecord | None = None
    slide_idx: int | None = None
    set_token_log_context(request_id=target_rid, scope="retry")

    try:
        # Find record and edit
        for r in history:
            if r.request_id == target_rid:
                target_rec = r
                break
        if not target_rec:
            logger.error(f"Retry: request #{target_rid} not found")
            return

        se = target_rec.slide_edits[target_step]
        task_list = se.task_data
        if not task_list:
            logger.error(f"Retry: no task data for step {target_step}")
            return

        slide_idx = se.slide_index
        _current_context = {"request_id": target_rid, "slide_index": slide_idx}
        set_token_log_context(slide_index=slide_idx)
        logger.info(f"Retrying request #{target_rid} step #{target_step} (slide {slide_idx}, {len(task_list)} tasks)")

        # If checkpoint exists, restore to pre-edit state first
        if se.checkpoint_path and os.path.exists(se.checkpoint_path):
            logger.info("Restoring checkpoint before retry...")
            original_path = os.path.abspath(container.prs.FullName)
            backup_path = container.backup_path

            try:
                container.prs.Close()
            except Exception as e:
                logger.warning(f"Error closing presentation: {e}")

            shutil.copy2(se.checkpoint_path, backup_path)
            shutil.copy2(se.checkpoint_path, original_path)
            container.prs = container.ppt_app.Presentations.Open(str(original_path))

            # Re-init parser
            global slide_count
            slide_count = len(container.prs.Slides)
            parser_obj.container = container
            parser_obj.total_slides = slide_count
            parser_obj.database = {}


            # Mark all edits after this one as rolled_back (undo-stack)
            found_after = False
            for rec in history:
                for i, edit in enumerate(rec.slide_edits):
                    if found_after:
                        edit.status = "rolled_back"
                        edit.checkpoint_path = None
                    if rec.request_id == target_rid and i == target_step:
                        found_after = True
                if found_after and rec.request_id != target_rid:
                    rec.status = "partial"

        # Mark as running
        se.status = "running"
        broadcast_sse("request_update", _serialize_record(target_rec))

        slide = container.prs.Slides(slide_idx)
        img_dir = image_root / f"slide_{slide_idx}"

        # Before screenshot
        before_path = export_slide_image(slide, img_dir, f"req{target_rid}_retry_before")
        se.before_image = f"slide_{slide_idx}/{before_path.name}"

        # Save new checkpoint
        if checkpoint_dir:
            ckpt_path = checkpoint_dir / f"checkpoint_req{target_rid}_step{target_step}_slide{slide_idx}_retry.pptx"
            try:
                container.prs.SaveCopyAs(str(ckpt_path))
                se.checkpoint_path = str(ckpt_path)
            except Exception as e:
                logger.warning(f"Checkpoint save failed: {e}")

        # Re-run all tasks for this slide via dispatcher
        try:
            had_any_tool_calls = False
            for task in task_list:
                slide_contents = parser_obj.process(slide_idx)
                objects_detail = slide_contents.get("Objects_Detail", []) if slide_contents else []

                sub_tasks = dispatcher.dispatch(task, objects_detail)
                logger.info(f"Retry dispatcher routed to: {[(s['agent_type'], s.get('shape_ids', [])) for s in sub_tasks]}")
                for sub in sub_tasks:
                    agent_type = sub["agent_type"]
                    sub_task = {**task, "description": sub["description"]}
                    had = specialist_agents[agent_type].run(
                        task=sub_task,
                        parser=parser_obj,
                        vision_validator_agent=vision_validator_agent,
                        visual_fixer_agent=visual_fixer_agent,
                        auto_resize=job.get("auto_resize", target_rec.auto_resize),
                        shape_ids=sub.get("shape_ids"),
                    )
                    if had:
                        had_any_tool_calls = True
            se.changed = bool(had_any_tool_calls)
            se.status = "done"
        except Exception as e:
            logger.error(f"Retry slide {slide_idx} failed: {e}")
            se.changed = False
            se.status = "error"

        # After screenshot
        try:
            after_path = export_slide_image(slide, img_dir, f"req{target_rid}_retry_after")
            se.after_image = f"slide_{slide_idx}/{after_path.name}"
        except Exception as e:
            logger.warning(f"Retry after-screenshot failed: {e}")

        broadcast_sse("slide_update", {
            "request_id": target_rid,
            "slide_index": slide_idx,
            "before_image": se.before_image,
            "after_image": se.after_image,
        })

    finally:
        # Always settle token usage — even when target_rec/se lookup failed,
        # checkpoint restore raised, or any other exception escaped above.
        delta = diff_tokens(token_snapshot_start, get_token_snapshot())
        if target_rec is not None:
            target_rec.total_input_tokens += delta["input_tokens"]
            target_rec.total_output_tokens += delta["output_tokens"]
            target_rec.total_cached_input_tokens += delta.get("cached_input_tokens", 0)
            target_rec.total_cost_usd += delta.get("cost_usd", 0.0)
            broadcast_sse("request_update", _serialize_record(target_rec))
        done_ts = time.time()
        log_request_summary(
            request_id=target_rid,
            user_input=(target_rec.user_input if target_rec is not None else ""),
            submit_ts=submit_ts,
            done_ts=done_ts,
            status=("done" if target_rec is not None and target_rec.slide_edits[target_step].status == "done" else "error"),
            scope="retry",
            task_count=(len(target_rec.slide_edits[target_step].task_data) if target_rec is not None else 0),
            slide_count=1,
            input_tokens=delta["input_tokens"],
            output_tokens=delta["output_tokens"],
            cached_input_tokens=delta.get("cached_input_tokens", 0),
            cost_usd=delta.get("cost_usd", 0.0),
            step_index=target_step,
            slide_index=slide_idx,
        )
        logger.info(
            f"Retry settled for request #{target_rid} step #{target_step} "
            f"(elapsed={done_ts - submit_ts:.3f}s, "
            f"tokens: {delta['input_tokens']}in / {delta['output_tokens']}out, "
            f"cached_in: {delta.get('cached_input_tokens', 0)}, "
            f"cost: ${delta.get('cost_usd', 0.0):.6f})"
        )
        _current_context = {"request_id": None, "slide_index": None}
        set_token_log_context(request_id=None, slide_index=None, scope=None)


# ===============================
# Retry-request execution (runs on COM thread)
# ===============================
def _execute_retry_request(job: dict, container, parser_obj, planner,
                           dispatcher, specialist_agents,
                           vision_validator_agent, visual_fixer_agent, log_root):
    """Rollback to the start of a request, then re-run the full pipeline with the same user_input."""
    global request_counter

    original_rid = job["original_request_id"]
    ckpt_path = job["checkpoint_path"]

    # Find original request
    original_rec = None
    for r in history:
        if r.request_id == original_rid:
            original_rec = r
            break
    if not original_rec:
        logger.error(f"Retry-request: request #{original_rid} not found")
        return

    user_input = original_rec.user_input

    # Step 1: Rollback to checkpoint (reuses _execute_rollback logic inline)
    logger.info(f"Retry-request: rolling back request #{original_rid} first...")
    _execute_rollback({
        "_type": "rollback",
        "request_id": original_rid,
        "step_index": 0,
        "checkpoint_path": ckpt_path,
    }, container, parser_obj)

    # Ensure the original request is marked as rolled_back
    original_rec.status = "rolled_back"
    broadcast_sse("request_update", _serialize_record(original_rec))

    # Step 2: Create a new RequestRecord and run pipeline.
    # Carry the original click-time submit_ts so request_summary.jsonl
    # measures wall-clock from the user's retry click, not from the moment
    # rollback finished.
    request_counter += 1
    new_rec = RequestRecord(
        request_id=request_counter,
        user_input=user_input,
        timestamp=job.get("submit_ts", time.time()),
        status="running",
        auto_resize=job.get("auto_resize", original_rec.auto_resize),
    )
    history.append(new_rec)
    broadcast_sse("request_update", _serialize_record(new_rec))

    logger.info(f"Retry-request: re-running as new request #{new_rec.request_id}")

    try:
        _run_pipeline(new_rec, container, planner, parser_obj,
                      dispatcher, specialist_agents, vision_validator_agent,
                      visual_fixer_agent, log_root)
    except Exception as e:
        logger.error(f"Retry-request pipeline failed: {e}")
        new_rec.status = "error"
        broadcast_sse("request_update", _serialize_record(new_rec))


# ===============================
# Rollback execution (runs on COM thread)
# ===============================
def _execute_rollback(job: dict, container, parser_obj):
    """Reopen presentation from checkpoint. Undo-stack: all edits after target are lost."""
    target_rid = job["request_id"]
    target_step = job["step_index"]
    ckpt_path = job["checkpoint_path"]

    logger.info(f"Rolling back to checkpoint: req#{target_rid} step#{target_step}")

    if not os.path.exists(ckpt_path):
        logger.error(f"Checkpoint file not found: {ckpt_path}")
        return

    original_path = os.path.abspath(container.prs.FullName)
    backup_path = container.backup_path

    # Close current presentation
    try:
        container.prs.Close()
    except Exception as e:
        logger.warning(f"Error closing presentation: {e}")

    # Copy checkpoint to backup and original locations
    shutil.copy2(ckpt_path, backup_path)
    shutil.copy2(ckpt_path, original_path)

    # Reopen
    container.prs = container.ppt_app.Presentations.Open(str(original_path))
    logger.info("Presentation reopened from checkpoint")

    # Re-init parser
    global slide_count
    slide_count = len(container.prs.Slides)
    parser_obj.container = container
    parser_obj.total_slides = slide_count
    parser_obj.database = {}

    # Mark target step and all subsequent edits as rolled_back
    found = False
    for rec in history:
        for i, se in enumerate(rec.slide_edits):
            if rec.request_id == target_rid and i == target_step:
                found = True
            if found:
                se.status = "rolled_back"
                se.checkpoint_path = None
        if found and rec.request_id != target_rid:
            rec.status = "partial"
        elif found and rec.request_id == target_rid:
            # Check if any edits before target_step are still done
            has_done = any(se.status == "done" for se in rec.slide_edits)
            rec.status = "partial" if has_done else "rolled_back"
        broadcast_sse("request_update", _serialize_record(rec))

    logger.info("Rollback completed")


# ===============================
# Serialization
# ===============================
def _serialize_record(rec: RequestRecord) -> dict:
    return {
        "request_id": rec.request_id,
        "user_input": rec.user_input,
        "timestamp": rec.timestamp,
        "status": rec.status,
        "total_input_tokens": rec.total_input_tokens,
        "total_output_tokens": rec.total_output_tokens,
        "total_cached_input_tokens": rec.total_cached_input_tokens,
        "total_cost_usd": rec.total_cost_usd,
        "slide_edits": [
            {
                "slide_index": se.slide_index,
                "before_image": se.before_image,
                "after_image": se.after_image,
                "status": se.status,
                "changed": se.changed,
                "has_checkpoint": se.checkpoint_path is not None,
                "has_task": bool(se.task_data),
            }
            for se in rec.slide_edits
        ],
    }


# ===============================
# COM startup helper
# ===============================
def _start_com_worker(ppt_path: Path):
    """Kill PowerPoint, start COM worker thread, wait until ready."""
    global file_loaded

    kill_powerpoint_processes()
    time.sleep(1)

    com_ready.clear()
    com_thread = threading.Thread(target=_com_worker, args=(ppt_path,), daemon=True)
    com_thread.start()
    com_ready.wait()

    if image_root is None:
        raise RuntimeError("COM worker failed to initialise")

    file_loaded = True


# ===============================
# Startup
# ===============================
def parse_args():
    p = argparse.ArgumentParser(description="EditPPT Web UI")
    p.add_argument("--file_path", type=str, required=False, default=None,
                   help="Path to .pptx file (optional — upload via browser if omitted)")
    p.add_argument("--port", type=int, default=5000, help="Flask port (default 5000)")
    p.add_argument("--host", type=str, default="127.0.0.1", help="Flask host")
    return p.parse_args()


def find_free_port(start=5000, end=5100):
    """Find a free port in the given range."""
    import socket
    for port in range(start, end):
        try:
            with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
                s.bind(("127.0.0.1", port))
                return port
        except OSError:
            continue
    raise RuntimeError(f"No free port found in range {start}-{end}")


def _open_app_window(url: str):
    """Open URL in Chrome app mode (no tabs/address bar). Falls back to default browser."""
    import subprocess, webbrowser
    try:
        import winreg
        key = winreg.OpenKey(
            winreg.HKEY_LOCAL_MACHINE,
            r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\chrome.exe",
        )
        chrome = winreg.QueryValue(key, None)
        winreg.CloseKey(key)
        subprocess.Popen([chrome, "--new-window", f"--app={url}"])
    except Exception:
        webbrowser.open(url)


def main():
    init_logger()
    analytics.init_analytics()
    analytics.track_app_launched()

    # Single-instance guard (exe builds only)
    _port_file = PROJECT_ROOT / ".server_port"
    if getattr(sys, "frozen", False):
        mutex = acquire_single_instance_mutex()
        if mutex is None:
            # Already running — open the existing instance's browser window
            try:
                port = int(_port_file.read_text(encoding="utf-8").strip())
            except Exception:
                port = 5000
            _open_app_window(f"http://127.0.0.1:{port}")
            sys.exit(0)

    # Clean up leftover files from a previous update / session (background —
    # shutil.rmtree on large backup dirs can take minutes on Windows)
    threading.Thread(target=cleanup_stale_update_artifacts, daemon=True).start()
    _cleanup_logfiles()

    args = parse_args()
    logger.info("=== EditPPT Web UI ===")

    if args.file_path:
        ppt_path = Path(args.file_path).expanduser().resolve()
        logger.info(f"PPT path: {ppt_path}")
        _start_com_worker(ppt_path)

    # Auto-find free port if default is occupied
    port = args.port
    import socket
    try:
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            s.bind((args.host, port))
    except OSError:
        port = find_free_port(port + 1)
        logger.info(f"Port {args.port} is in use, using port {port}")

    logger.info(f"Starting Flask server on {args.host}:{port}")
    _port_file.write_text(str(port), encoding="utf-8")

    # Check for updates in background (non-blocking)
    def _startup_checks():
        try:
            time.sleep(2)  # wait for SSE clients to connect

            # Check for successful update — show release notes
            complete_marker = PROJECT_ROOT.parent / "_update_complete.md"
            if complete_marker.exists():
                try:
                    content = complete_marker.read_text(encoding="utf-8")
                    first_nl = content.index("\n")
                    version = content[:first_nl].strip()
                    notes = content[first_nl + 1:]
                    broadcast_sse("update_complete", {
                        "version": version,
                        "release_notes": notes,
                    })
                except Exception:
                    pass
                complete_marker.unlink(missing_ok=True)

            # Check for previous update failure
            fail_marker = PROJECT_ROOT / "_update_failed.txt"
            if fail_marker.exists():
                try:
                    reason = fail_marker.read_text(encoding="utf-8").strip()
                except Exception:
                    reason = "Unknown error"
                logger.warning(f"Previous update failed: {reason}")
                broadcast_sse("update_failed_previous", {"reason": reason})
                fail_marker.unlink(missing_ok=True)

            # Check for new updates
            info = check_for_update()
            if info["has_update"]:
                logger.info(f"Update available: {info['current']} → {info['latest']}")
                broadcast_sse("update_available", {
                    "current": info["current"],
                    "latest": info["latest"],
                    "release_notes": info["release_notes"],
                })
        except Exception:
            pass

    threading.Thread(target=_startup_checks, daemon=True).start()

    # Skip opening a new window after update — the old browser will auto-reconnect
    _update_marker = PROJECT_ROOT.parent / "_update_complete.md"
    if not _update_marker.exists():
        threading.Timer(1.0, _open_app_window, args=[f"http://{args.host}:{port}"]).start()

    app.run(host=args.host, port=port, debug=False, threaded=True)


if __name__ == "__main__":
    main()
