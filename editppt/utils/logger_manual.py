# editppt/utils/logger_manual.py

from __future__ import annotations

import sys
import atexit
from datetime import datetime
from pathlib import Path
from loguru import logger

from editppt.config import LOG_BASE

# Current log folder, determined at runtime
_current_log_dir: Path | None = None


def _safe_filename(name: str) -> str:
    if not name:
        return "unknown"
    # Windows forbidden: \ / : * ? " < > |
    for ch in ['\\', '/', ':', '*', '?', '"', '<', '>', '|']:
        name = name.replace(ch, "_")
    return name.strip()


def get_dynamic_log_dir(container=None) -> Path | None:
    """
    Create logfiles/YYYYMMDD/TIMESTAMP_Name_SlideCount folder based on container info.
    """
    global _current_log_dir

    if container:
        now = datetime.now()
        date_folder = now.strftime("%Y%m%d")
        timestamp = now.strftime("%Y%m%d_%H%M%S")

        # Extract file name and slide count
        try:
            file_name = _safe_filename(container.prs.Name)
        except Exception:
            file_name = "unknown.pptx"

        try:
            slide_count = len(container.prs.Slides)
        except Exception:
            slide_count = 0

        # Final folder name: 20260126_190129_example.pptx_16
        folder_name = f"{timestamp}_{file_name}_{slide_count}"

        _current_log_dir = LOG_BASE / date_folder / folder_name
        _current_log_dir.mkdir(parents=True, exist_ok=True)

    return _current_log_dir


def log_path(filename: str) -> Path:
    """
    Return a file path within the current log directory.
    Falls back to LOG_BASE if container has not been initialized yet.
    """
    log_dir = get_dynamic_log_dir()
    if log_dir is None:
        log_dir = LOG_BASE
        log_dir.mkdir(parents=True, exist_ok=True)
    return log_dir / filename


class StreamToLoguru:
    """
    Wrapper that redirects print/stdout/stderr to loguru.
    Lines without newline stay in buffer; flush on exit.
    """

    def __init__(self, level: str = "INFO"):
        self.level = level
        self._buffer = ""

    def write(self, message: str):
        if not message:
            return

        self._buffer += message

        while "\n" in self._buffer:
            line, self._buffer = self._buffer.split("\n", 1)
            line = line.rstrip()
            if line:
                logger.log(self.level, line)

    def flush(self):
        if self._buffer.strip():
            logger.log(self.level, self._buffer.rstrip())
        self._buffer = ""


def init_logger(container=None):
    """
    Initialize logger.
    - Without container: console logging only (no file sinks)
    - With container: file logging into TIMESTAMP_Name_SlideCount folder
    """
    global _current_log_dir

    # If already initialized, only allow directory update when container is provided
    if getattr(init_logger, "_initialized", False):
        if container is not None:
            # Switch to container-based folder and reconfigure sinks
            init_logger._initialized = False
        else:
            return logger

    # Remove existing sinks
    logger.remove()

    # Console output (stderr is None in windowed mode)
    if sys.stderr is not None:
        logger.add(
            sys.stderr,
            level="DEBUG",
            format="<green>{time:HH:mm:ss}</green> | <level>{level}</level> | {message}",
        )

    # File output: only when container exists (disabled in frozen builds)
    log_dir = get_dynamic_log_dir(container)
    if log_dir is not None and not getattr(sys, "frozen", False):
        logger.add(
            log_dir / "error.log",
            level="ERROR",
            encoding="utf-8",
            rotation="10 MB",
            enqueue=True,
            backtrace=True,
            diagnose=False,
        )

        logger.add(
            log_dir / "app.log",
            level="DEBUG",
            encoding="utf-8",
            rotation="10 MB",
            enqueue=True,
            backtrace=True,
            diagnose=False,
        )

        logger.add(
            log_dir / "shell_all.log",
            level="DEBUG",
            encoding="utf-8",
            rotation="10 MB",
            enqueue=True,
            backtrace=True,
            diagnose=False,
        )

    # stdout/stderr redirect
    stdout_proxy = StreamToLoguru("INFO")
    stderr_proxy = StreamToLoguru("ERROR")

    # Flush buffer on exit (prevent losing output without trailing newline)
    atexit.register(stdout_proxy.flush)
    atexit.register(stderr_proxy.flush)

    init_logger._initialized = True
    return logger
