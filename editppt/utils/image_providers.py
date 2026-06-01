"""
Image search (Pexels / Unsplash) and generation (Gemini Imagen) utilities.

Downloaded / generated images are saved to a temp directory under logfiles/.
"""

import os
import time
import uuid
import logging
from pathlib import Path

import requests
from dotenv import load_dotenv

from editppt.config import LOG_BASE, IMAGE_GENERATION_MODEL, PROJECT_ROOT
from editppt.utils.llm_client import record_external_llm_usage

load_dotenv(PROJECT_ROOT / ".env")
logger = logging.getLogger(__name__)

# ── Temp directory for downloaded / generated images ──

_IMAGE_TEMP_DIR = LOG_BASE / "temp_images"
_IMAGE_TEMP_DIR.mkdir(parents=True, exist_ok=True)


def _save_bytes(data: bytes, ext: str = "png") -> str:
    """Save raw bytes to a temp file and return the absolute path."""
    filename = f"{uuid.uuid4().hex[:12]}.{ext}"
    filepath = _IMAGE_TEMP_DIR / filename
    filepath.write_bytes(data)
    return str(filepath)


# ─────────────────────────────────────────────
# 1. Pexels
# ─────────────────────────────────────────────

def search_image_pexels(query: str, count: int = 1) -> str:
    """Search Pexels for an image. Returns the local file path of the downloaded image."""
    api_key = os.getenv("PEXELS_API_KEY")
    if not api_key:
        raise EnvironmentError("PEXELS_API_KEY is not set in .env")

    resp = requests.get(
        "https://api.pexels.com/v1/search",
        headers={"Authorization": api_key},
        params={"query": query, "per_page": count, "orientation": "landscape"},
        timeout=15,
    )
    resp.raise_for_status()
    data = resp.json()

    photos = data.get("photos", [])
    if not photos:
        raise RuntimeError(f"Pexels returned no results for '{query}'")

    image_url = photos[0]["src"]["large"]
    img_resp = requests.get(image_url, timeout=30)
    img_resp.raise_for_status()

    ext = "jpg"
    return _save_bytes(img_resp.content, ext)


# ─────────────────────────────────────────────
# 2. Unsplash
# ─────────────────────────────────────────────

def search_image_unsplash(query: str, count: int = 1) -> str:
    """Search Unsplash for an image. Returns the local file path of the downloaded image."""
    access_key = os.getenv("UNSPLASH_ACCESS_KEY")
    if not access_key:
        raise EnvironmentError("UNSPLASH_ACCESS_KEY is not set in .env")

    resp = requests.get(
        "https://api.unsplash.com/search/photos",
        headers={"Authorization": f"Client-ID {access_key}"},
        params={"query": query, "per_page": count, "orientation": "landscape"},
        timeout=15,
    )
    resp.raise_for_status()
    data = resp.json()

    results = data.get("results", [])
    if not results:
        raise RuntimeError(f"Unsplash returned no results for '{query}'")

    image_url = results[0]["urls"]["regular"]
    img_resp = requests.get(image_url, timeout=30)
    img_resp.raise_for_status()

    ext = "jpg"
    return _save_bytes(img_resp.content, ext)


# ─────────────────────────────────────────────
# 3. Unified search (Pexels → Unsplash fallback)
# ─────────────────────────────────────────────

def search_image(query: str) -> str:
    """Search for an image using Pexels first, falling back to Unsplash on error."""
    try:
        return search_image_pexels(query)
    except Exception as e:
        logger.warning(f"Pexels search failed ({e}), falling back to Unsplash")
        return search_image_unsplash(query)


# ─────────────────────────────────────────────
# 4. Gemini image generation
# ─────────────────────────────────────────────

def generate_image_gemini(prompt: str) -> str:
    """Generate an image using Gemini 3 Pro. Returns the local file path."""
    api_key = os.getenv("GEMINI_API_KEY")
    if not api_key:
        raise EnvironmentError("GEMINI_API_KEY is not set in .env")

    from google import genai
    from google.genai import types

    client = genai.Client(api_key=api_key)
    t0 = time.monotonic()
    try:
        response = client.models.generate_content(
            model=IMAGE_GENERATION_MODEL,
            contents=prompt,
            config=types.GenerateContentConfig(
                response_modalities=["IMAGE"],
            ),
        )
    except Exception as e:
        record_external_llm_usage(IMAGE_GENERATION_MODEL, "gemini", error=e,
                                  latency_seconds=time.monotonic() - t0)
        raise
    record_external_llm_usage(IMAGE_GENERATION_MODEL, "gemini", response=response,
                              latency_seconds=time.monotonic() - t0)

    parts = response.candidates[0].content.parts
    image_part = next((p for p in parts if p.inline_data and p.inline_data.mime_type.startswith("image/")), None)
    if not image_part:
        raise RuntimeError(f"Gemini returned no image for prompt: {prompt}")

    return _save_bytes(image_part.inline_data.data, "png")


# ─────────────────────────────────────────────
# 5. Gemini image editing (reference image + prompt)
# ─────────────────────────────────────────────

def edit_image_gemini(image_path: str, prompt: str) -> str:
    """Edit an existing image using Gemini 3 Pro. Returns the local file path of the result."""
    api_key = os.getenv("GEMINI_API_KEY")
    if not api_key:
        raise EnvironmentError("GEMINI_API_KEY is not set in .env")

    import PIL.Image
    from google import genai
    from google.genai import types

    img = PIL.Image.open(image_path)

    client = genai.Client(api_key=api_key)
    t0 = time.monotonic()
    try:
        response = client.models.generate_content(
            model=IMAGE_GENERATION_MODEL,
            contents=[img, prompt],
            config=types.GenerateContentConfig(
                response_modalities=["IMAGE"],
            ),
        )
    except Exception as e:
        record_external_llm_usage(IMAGE_GENERATION_MODEL, "gemini", error=e,
                                  latency_seconds=time.monotonic() - t0)
        raise
    record_external_llm_usage(IMAGE_GENERATION_MODEL, "gemini", response=response,
                              latency_seconds=time.monotonic() - t0)

    candidates = response.candidates or []
    parts = (candidates[0].content.parts if candidates and candidates[0].content else None) or []
    image_part = next((p for p in parts if p.inline_data and p.inline_data.mime_type.startswith("image/")), None)
    if not image_part:
        raise RuntimeError(f"Gemini returned no edited image for prompt: {prompt}")

    return _save_bytes(image_part.inline_data.data, "png")
