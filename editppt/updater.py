"""Auto-update module for EditPPT.

Checks GitHub Releases for new versions, downloads, extracts,
and generates a .bat script to swap files while the app restarts.
"""

import ctypes
import fnmatch
import hashlib
import os
import shutil
import subprocess
import sys
import textwrap
import zipfile
from pathlib import Path
from typing import Callable

import requests

from editppt.config import APP_VERSION, UPDATE_REPO, PROJECT_ROOT


_GITHUB_API = f"https://api.github.com/repos/{UPDATE_REPO}/releases/latest"
_ASSET_NAME = "EditPPT.zip"
_SHA256_ASSET_NAME = "EditPPT.zip.sha256"
_STAGING_DIR_NAME = "_update_staging"


# ---------------------------------------------------------------------------
# Phase 1A: Cleanup stale update artifacts on startup
# ---------------------------------------------------------------------------

def cleanup_stale_update_artifacts():
    """Remove leftover files from a previous (possibly interrupted) update."""
    dirs_to_remove = [
        PROJECT_ROOT / _STAGING_DIR_NAME,           # old staging
        PROJECT_ROOT.parent / "_editppt_backup",     # new-flow backup
        PROJECT_ROOT.parent / "_editppt_new",        # new-flow staging
    ]
    files_to_remove = [
        PROJECT_ROOT / "_update.bat",                # old script
        PROJECT_ROOT.parent / "_update_editppt.bat", # new-flow script
        PROJECT_ROOT.parent / "rollback.bat",         # rollback script
        PROJECT_ROOT.parent / "_launch_update.vbs",   # VBS launcher
    ]
    for d in dirs_to_remove:
        if d.exists():
            shutil.rmtree(d, ignore_errors=True)
    for f in files_to_remove:
        f.unlink(missing_ok=True)


# ---------------------------------------------------------------------------
# Phase 1B: Single-instance mutex
# ---------------------------------------------------------------------------

def acquire_single_instance_mutex():
    """Acquire a system-wide mutex to prevent multiple instances.

    Returns the mutex handle on success, or None if another instance is running.
    """
    kernel32 = ctypes.windll.kernel32
    ERROR_ALREADY_EXISTS = 183
    handle = kernel32.CreateMutexW(None, False, "Global\\EditPPT_SingleInstance")
    if kernel32.GetLastError() == ERROR_ALREADY_EXISTS:
        kernel32.CloseHandle(handle)
        return None
    return handle


# ---------------------------------------------------------------------------
# Version parsing / update check
# ---------------------------------------------------------------------------

def parse_version(tag: str) -> tuple[int, ...]:
    """Parse a version tag like 'v0.1' into a comparable tuple (0, 1)."""
    return tuple(int(x) for x in tag.lstrip("v").split("."))


def check_for_update() -> dict:
    """Check GitHub for a newer release.

    Returns dict with keys: has_update, current, latest, download_url,
    release_notes, sha256_url.
    """
    result = {
        "has_update": False,
        "current": APP_VERSION,
        "latest": APP_VERSION,
        "download_url": None,
        "release_notes": "",
        "sha256_url": None,
    }
    try:
        resp = requests.get(_GITHUB_API, timeout=5)
        resp.raise_for_status()
        data = resp.json()
    except Exception:
        return result

    latest_tag = data.get("tag_name", APP_VERSION)
    result["latest"] = latest_tag
    result["release_notes"] = data.get("body", "") or ""

    if parse_version(latest_tag) <= parse_version(APP_VERSION):
        return result

    # Find the zip asset and optional sha256 asset
    for asset in data.get("assets", []):
        name = asset.get("name", "")
        if name == _ASSET_NAME:
            result["has_update"] = True
            result["download_url"] = asset["browser_download_url"]
        elif name == _SHA256_ASSET_NAME:
            result["sha256_url"] = asset["browser_download_url"]

    return result


# ---------------------------------------------------------------------------
# Download
# ---------------------------------------------------------------------------

def download_update(
    url: str,
    dest_dir: Path,
    progress_callback: Callable[[int, int], None] | None = None,
) -> Path:
    """Download the update zip to dest_dir. Returns path to the downloaded zip."""
    dest_dir.mkdir(parents=True, exist_ok=True)
    zip_path = dest_dir / _ASSET_NAME

    try:
        resp = requests.get(url, stream=True, timeout=(10, 30))
        resp.raise_for_status()
    except requests.ConnectionError:
        raise RuntimeError("Network connection failed. Please check your internet connection.")
    except requests.Timeout:
        raise RuntimeError("Server response timed out. Please try again later.")
    except requests.RequestException as e:
        raise RuntimeError(f"Download request failed: {e}")

    total = int(resp.headers.get("Content-Length", 0))
    downloaded = 0

    try:
        with open(zip_path, "wb") as f:
            for chunk in resp.iter_content(chunk_size=65536):
                f.write(chunk)
                downloaded += len(chunk)
                if progress_callback and total > 0:
                    progress_callback(downloaded, total)
    except Exception:
        zip_path.unlink(missing_ok=True)
        raise RuntimeError("Download interrupted. Please check your network and try again.")

    # Verify downloaded size matches Content-Length
    if total > 0 and downloaded != total:
        zip_path.unlink(missing_ok=True)
        raise RuntimeError(
            f"Incomplete download: {downloaded}/{total} bytes. Please try again."
        )

    return zip_path


# ---------------------------------------------------------------------------
# Phase 2: SHA-256 integrity verification
# ---------------------------------------------------------------------------

def verify_sha256(zip_path: Path, expected_hash: str) -> bool:
    """Return True if the SHA-256 hash of *zip_path* matches *expected_hash*."""
    sha = hashlib.sha256()
    with open(zip_path, "rb") as f:
        for chunk in iter(lambda: f.read(65536), b""):
            sha.update(chunk)
    return sha.hexdigest().lower() == expected_hash.strip().lower()


# ---------------------------------------------------------------------------
# Phase 1C: Enhanced extraction with validation
# ---------------------------------------------------------------------------

def extract_update(zip_path: Path, staging_dir: Path) -> Path:
    """Extract zip and return path to the inner EditPPT folder."""
    if staging_dir.exists():
        shutil.rmtree(staging_dir)
    staging_dir.mkdir(parents=True, exist_ok=True)

    with zipfile.ZipFile(zip_path, "r") as zf:
        zf.extractall(staging_dir)

    # The zip should contain an EditPPT/ folder at root
    inner = staging_dir / "EditPPT"
    if inner.is_dir() and (inner / "EditPPT.exe").exists():
        _validate_extracted_contents(inner)
        return inner

    # Fallback: maybe files are at root level
    if (staging_dir / "EditPPT.exe").exists():
        _validate_extracted_contents(staging_dir)
        return staging_dir

    raise FileNotFoundError("EditPPT.exe not found in the update package")


def _validate_extracted_contents(folder: Path):
    """Verify critical files exist in the extracted folder."""
    has_python_dll = any(
        fnmatch.fnmatch(p.name, "python*.dll")
        for p in folder.rglob("*.dll")
    )
    if not has_python_dll:
        raise FileNotFoundError(
            "Update package is incomplete: no python*.dll found alongside EditPPT.exe"
        )


# ---------------------------------------------------------------------------
# Phase 3: Atomic directory swap script generation
# ---------------------------------------------------------------------------

def generate_update_script(new_dir: Path, target_dir: Path, port: int = 5000) -> Path:
    """Generate a .bat script that atomically swaps directories and restarts.

    The script is placed in the *parent* of target_dir so it is not inside
    the directory being renamed.
    """
    bat_path = target_dir.parent / "_update_editppt.bat"
    pid = os.getpid()
    backup_dir = target_dir.parent / "_editppt_backup"
    target_name = target_dir.name  # e.g. "EditPPT"

    fail_marker = target_dir / "_update_failed.txt"

    script = textwrap.dedent(f"""\
        @echo off
        chcp 65001 >nul 2>&1
        echo Stopping EditPPT...
        taskkill /PID {pid} /F >nul 2>&1

        :waitloop
        ping -n 2 127.0.0.1 >nul
        tasklist /FI "PID eq {pid}" 2>NUL | find /I "{pid}" >NUL
        if not errorlevel 1 goto waitloop

        echo Waiting for port {port} to be released...
        :portloop
        netstat -ano | findstr ":{port} " | findstr "LISTENING" >nul 2>&1
        if not errorlevel 1 (
            timeout /t 1 /nobreak >nul
            goto portloop
        )

        echo Backing up current installation...
        if exist "{backup_dir}" rmdir /S /Q "{backup_dir}"
        rename "{target_dir}" "_editppt_backup"
        if errorlevel 1 (
            echo ERROR: Failed to rename current installation for backup.
            echo rename_failed: Could not rename current installation for backup. > "{fail_marker}"
            exit /b 1
        )

        echo Applying update...
        rename "{new_dir}" "{target_name}"
        if errorlevel 1 (
            echo ERROR: Failed to move new version into place. Rolling back...
            rename "{backup_dir}" "{target_name}"
            echo apply_failed: Could not move new version into place. Rolled back to previous version. > "{fail_marker}"
            exit /b 1
        )

        echo Verifying update...
        if not exist "{target_dir}\\EditPPT.exe" (
            echo ERROR: EditPPT.exe missing after update. Rolling back...
            rmdir /S /Q "{target_dir}"
            rename "{backup_dir}" "{target_name}"
            echo exe_missing: EditPPT.exe not found after update. Rolled back to previous version. > "{fail_marker}"
            exit /b 1
        )

        echo Restoring user data...
        if exist "{backup_dir}\\.env" copy /Y "{backup_dir}\\.env" "{target_dir}\\.env" >nul
        if exist "{backup_dir}\\.analytics_id" copy /Y "{backup_dir}\\.analytics_id" "{target_dir}\\.analytics_id" >nul

        echo Cleaning up...
        rmdir /S /Q "{backup_dir}" 2>nul
        rmdir /S /Q "{target_dir}\\{_STAGING_DIR_NAME}" 2>nul

        echo Starting EditPPT on port {port}...
        start "" "{target_dir}\\EditPPT.exe" --port {port}

        del "%~f0"
    """)

    bat_path.write_text(script, encoding="utf-8")
    return bat_path


def generate_rollback_script(target_dir: Path) -> Path:
    """Generate a manual rollback.bat in the parent directory."""
    bat_path = target_dir.parent / "rollback.bat"
    backup_name = "_editppt_backup"
    target_name = target_dir.name

    script = textwrap.dedent(f"""\
        @echo off
        chcp 65001 >nul 2>&1
        cd /d "%~dp0"

        if not exist "{backup_name}" (
            echo No backup found. Cannot rollback.
            pause
            exit /b 1
        )

        echo Rolling back to previous version...
        if exist "{target_name}" rmdir /S /Q "{target_name}"
        rename "{backup_name}" "{target_name}"
        if errorlevel 1 (
            echo ERROR: Rollback failed.
            pause
            exit /b 1
        )

        echo Rollback complete.
        pause
    """)

    bat_path.write_text(script, encoding="utf-8")
    return bat_path


# ---------------------------------------------------------------------------
# Main orchestration
# ---------------------------------------------------------------------------

def download_and_prepare_update(
    progress_callback: Callable[[str, int], None] | None = None,
    port: int = 5000,
) -> Path:
    """Download, extract, verify, and generate the update script. Returns bat_path.

    Does NOT launch the script -- caller should close resources first,
    then call launch_update_script().
    On failure, all staging artifacts are cleaned up so the next attempt
    starts fresh.
    """
    staging_root = PROJECT_ROOT / _STAGING_DIR_NAME
    new_dir = PROJECT_ROOT.parent / "_editppt_new"

    # Clean up any leftovers from a previous failed attempt
    if staging_root.exists():
        shutil.rmtree(staging_root, ignore_errors=True)
    if new_dir.exists():
        shutil.rmtree(new_dir, ignore_errors=True)

    try:
        info = check_for_update()
        if not info["has_update"] or not info["download_url"]:
            raise RuntimeError("No update available")

        # Download
        def _dl_progress(downloaded: int, total: int):
            pct = min(int(downloaded / total * 100), 100)
            if progress_callback:
                progress_callback("downloading", pct)

        if progress_callback:
            progress_callback("downloading", 0)

        zip_path = download_update(info["download_url"], staging_root, _dl_progress)

        # SHA-256 verification (Phase 2)
        if info.get("sha256_url"):
            if progress_callback:
                progress_callback("verifying", 0)
            try:
                sha_resp = requests.get(info["sha256_url"], timeout=10)
                sha_resp.raise_for_status()
                expected_hash = sha_resp.text.strip().split()[0]
                if not verify_sha256(zip_path, expected_hash):
                    zip_path.unlink(missing_ok=True)
                    raise RuntimeError("SHA-256 verification failed: download may be corrupted")
            except requests.RequestException:
                # If we can't download the hash file, warn but continue
                pass

        # Extract
        if progress_callback:
            progress_callback("extracting", 0)

        inner_dir = extract_update(zip_path, staging_root / "extracted")

        # Move extracted dir to parent-level staging (_editppt_new)
        if new_dir.exists():
            shutil.rmtree(new_dir)
        shutil.move(str(inner_dir), str(new_dir))

        # Generate bat script + rollback script
        if progress_callback:
            progress_callback("installing", 0)

        generate_rollback_script(PROJECT_ROOT)
        bat_path = generate_update_script(new_dir, PROJECT_ROOT, port=port)
        return bat_path

    except Exception:
        # Clean up all artifacts so next retry starts clean
        shutil.rmtree(staging_root, ignore_errors=True)
        shutil.rmtree(new_dir, ignore_errors=True)
        bat_path = PROJECT_ROOT.parent / "_update_editppt.bat"
        bat_path.unlink(missing_ok=True)
        raise


def launch_update_script(bat_path: Path):
    """Launch the update .bat script independently (survives parent exit).

    Uses CreateProcessW via ctypes to bypass Python's subprocess module,
    which can fail in frozen PyInstaller executables. The bat script will
    taskkill this process, then swap directories and restart.
    """

    class STARTUPINFOW(ctypes.Structure):
        _fields_ = [
            ("cb", ctypes.c_ulong), ("lpReserved", ctypes.c_wchar_p),
            ("lpDesktop", ctypes.c_wchar_p), ("lpTitle", ctypes.c_wchar_p),
            ("dwX", ctypes.c_ulong), ("dwY", ctypes.c_ulong),
            ("dwXSize", ctypes.c_ulong), ("dwYSize", ctypes.c_ulong),
            ("dwXCountChars", ctypes.c_ulong), ("dwYCountChars", ctypes.c_ulong),
            ("dwFillAttribute", ctypes.c_ulong), ("dwFlags", ctypes.c_ulong),
            ("wShowWindow", ctypes.c_ushort), ("cbReserved2", ctypes.c_ushort),
            ("lpReserved2", ctypes.c_void_p),
            ("hStdInput", ctypes.c_void_p), ("hStdOutput", ctypes.c_void_p),
            ("hStdError", ctypes.c_void_p),
        ]

    class PROCESS_INFORMATION(ctypes.Structure):
        _fields_ = [
            ("hProcess", ctypes.c_void_p), ("hThread", ctypes.c_void_p),
            ("dwProcessId", ctypes.c_ulong), ("dwThreadId", ctypes.c_ulong),
        ]

    si = STARTUPINFOW()
    si.cb = ctypes.sizeof(STARTUPINFOW)
    pi = PROCESS_INFORMATION()

    comspec = os.environ.get("COMSPEC", r"C:\Windows\System32\cmd.exe")
    cmd_line = ctypes.create_unicode_buffer(f'"{comspec}" /c "{bat_path}"')

    CREATE_NEW_PROCESS_GROUP = 0x00000200
    CREATE_NO_WINDOW = 0x08000000

    ok = ctypes.windll.kernel32.CreateProcessW(
        None, cmd_line, None, None, False,
        CREATE_NEW_PROCESS_GROUP | CREATE_NO_WINDOW,
        None, str(bat_path.parent), ctypes.byref(si), ctypes.byref(pi),
    )
    if ok:
        ctypes.windll.kernel32.CloseHandle(pi.hProcess)
        ctypes.windll.kernel32.CloseHandle(pi.hThread)
    else:
        err = ctypes.windll.kernel32.GetLastError()
        raise RuntimeError(f"CreateProcessW failed with error {err}")


def apply_update(progress_callback: Callable[[str, int], None] | None = None):
    """Convenience: download + prepare + launch + exit. For simple callers."""
    bat_path = download_and_prepare_update(progress_callback)
    launch_update_script(bat_path)
    if progress_callback:
        progress_callback("restarting", 100)
