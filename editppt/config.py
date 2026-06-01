import sys
from pathlib import Path

# --- LLM Models ---
CURRENT_MODEL_NAME = "gpt-4.1"
STYLE_MAPPER_MODEL = "gpt-4.1"
CURRENT_VISION_MODEL_NAME = "gemini-2.5-pro"
IMAGE_GENERATION_MODEL = "models/gemini-3-pro-image-preview"
IMAGE_CAPTION_MODEL = "gemini-2.5-flash-lite"

# --- App Version & Update ---
APP_VERSION = "v0.1.1"
UPDATE_REPO = "anonymous/EditPPT-release"

# --- Directory Layout ---
if getattr(sys, 'frozen', False):
    # PyInstaller bundle: read-only resources in _MEIPASS, writable files next to exe
    RESOURCE_ROOT = Path(sys._MEIPASS)
    PROJECT_ROOT = Path(sys.executable).parent
else:
    RESOURCE_ROOT = Path(__file__).resolve().parent.parent
    PROJECT_ROOT = RESOURCE_ROOT

LOG_BASE = PROJECT_ROOT / "logfiles"
UPLOADS_DIR = PROJECT_ROOT / "uploads"
