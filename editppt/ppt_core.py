import os
import shutil
import subprocess
from pathlib import Path

from loguru import logger
import pythoncom
import win32com.client
import win32com.client.dynamic


def _clear_gen_py_cache():
    """Remove stale win32com gen_py cache that causes CLSIDToClassMap errors.

    Do NOT recreate the directory after rmtree: an empty dir has no
    __init__.py and makes the next `import win32com.gen_py` raise
    ModuleNotFoundError, defeating the retry. gencache rebuilds the layout
    itself on the next EnsureDispatch.
    """
    import sys

    try:
        cache_dir = Path(win32com.__gen_path__)
        if cache_dir.exists():
            shutil.rmtree(cache_dir)
    except Exception as e:
        logger.warning(f"Failed to clear gen_py cache dir: {e}")

    # Purge any corrupt gen_py modules from sys.modules
    stale = [k for k in sys.modules if k.startswith("win32com.gen_py")]
    for k in stale:
        del sys.modules[k]

    # Reset gencache internal state so it re-scans from scratch
    try:
        from win32com.client import gencache
        gencache.__init__()
    except Exception:
        pass


class PPTContainer:
    def __init__(self, prs, ppt_app):
        self.prs = prs
        self.ppt_app = ppt_app
        self.original_path: str | None = None
        self.backup_path: str | None = None
        self.pristine_path: str | None = None


def kill_powerpoint_processes():
    try:
        subprocess.run(
            ["taskkill", "/F", "/IM", "powerpnt.exe", "/T"],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )
        logger.info("Cleared all existing PowerPoint processes.")
    except Exception as e:
        logger.warning(f"Error while clearing PowerPoint processes (may not be running): {e}")


def initialize_ppt(ppt_path: Path):
    """Connects to PowerPoint and opens a specific file.

    Uses early binding (gencache.EnsureDispatch) for ~3-5x faster COM property
    access. Falls back to a one-shot cache rebuild if the cached typelib is
    stale, then to dynamic dispatch if early binding still fails.
    """
    ppt_path = ppt_path.resolve()
    if not ppt_path.exists():
        raise FileNotFoundError(f"PPT file not found: {ppt_path}")

    logger.info("PowerPoint is not running. Launching new instance...")

    from win32com.client import gencache
    try:
        ppt_app = gencache.EnsureDispatch("PowerPoint.Application")
    except (AttributeError, ImportError, Exception) as e:
        logger.warning(f"Early-binding init failed ({type(e).__name__}: {e}); rebuilding cache and retrying.")
        _clear_gen_py_cache()
        try:
            ppt_app = gencache.EnsureDispatch("PowerPoint.Application")
        except Exception as e2:
            logger.warning(f"Early binding still fails ({type(e2).__name__}: {e2}); falling back to dynamic dispatch.")
            ppt_app = win32com.client.dynamic.Dispatch("PowerPoint.Application")
    ppt_app.Visible = True

    logger.info(f"Opening PPT file: {ppt_path}")
    prs = ppt_app.Presentations.Open(str(ppt_path))
    return prs, ppt_app


def export_slide_image(slide, out_dir: Path, tag: str) -> Path:
    out_dir.mkdir(parents=True, exist_ok=True)
    image_path = out_dir / f"slide_{slide.SlideIndex}_{tag}.png"
    slide.Export(str(image_path), "PNG")
    return image_path


def init_backup(container: PPTContainer) -> None:
    """Create initial backup of the presentation."""
    full_name = os.path.abspath(container.prs.FullName)
    container.original_path = full_name
    container.backup_path = full_name.replace(".pptx", "_.pptx")
    if os.path.exists(container.backup_path):
        os.remove(container.backup_path)
    container.prs.SaveCopyAs(container.backup_path)

    # Pristine snapshot: immutable copy of the original state at session start
    container.pristine_path = full_name.replace(".pptx", "_pristine.pptx")
    if os.path.exists(container.pristine_path):
        os.remove(container.pristine_path)
    container.prs.SaveCopyAs(container.pristine_path)
    logger.info(f"Pristine snapshot saved: {container.pristine_path}")
