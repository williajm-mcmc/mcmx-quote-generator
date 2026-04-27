# updater.py
"""
Auto-update helper for MCMX Quote Generator.

Public API used by main.py
──────────────────────────
  check_for_update(current_version)   → dict | None
  download_zip(url, dest, progress_cb=None)
  apply_update_windows(zip_path, app_exe)
  is_bundled() → bool
"""

import json
import os
import platform
import re
import sys
import tempfile
import urllib.request
import zipfile

# ── Constants ─────────────────────────────────────────────────────────────────
GITHUB_REPO   = "williajm-mcmc/mcmx-quote-generator"
RELEASES_API  = f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest"
WINDOWS_ASSET = "QuoteGenerator-Windows.zip"
MAC_ASSET     = "QuoteGenerator-Mac.zip"

_HEADERS = {
    "User-Agent": "MCMX-QuoteGenerator-Updater/1.0",
    "Accept":     "application/vnd.github+json",
}


# ── Version helpers ───────────────────────────────────────────────────────────

def _parse_version(tag: str) -> tuple:
    """'v2.0.8'  →  (2, 0, 8).  Non-numeric parts become 0."""
    tag = tag.lstrip("vV")
    parts = re.split(r"[.\-]", tag)
    result = []
    for p in parts:
        try:
            result.append(int(p))
        except ValueError:
            result.append(0)
    return tuple(result) or (0,)


# ── Network calls ─────────────────────────────────────────────────────────────

def check_for_update(current_version: str) -> dict | None:
    """
    Query GitHub for the latest release.

    Returns a dict  {tag, download_url, body}  when a newer version exists,
    or None when the app is already current or no matching asset is found.
    Raises on network / JSON errors — callers should catch Exception.
    """
    req = urllib.request.Request(RELEASES_API, headers=_HEADERS)
    with urllib.request.urlopen(req, timeout=12) as resp:
        data = json.loads(resp.read())

    tag  = data.get("tag_name", "")
    body = data.get("body", "")

    latest_ver  = _parse_version(tag)
    current_ver = _parse_version(current_version)

    if latest_ver <= current_ver:
        return None  # already up to date

    # Locate the right asset for this platform
    is_mac   = platform.system() == "Darwin"
    target   = MAC_ASSET if is_mac else WINDOWS_ASSET
    url      = None
    for asset in data.get("assets", []):
        if asset.get("name", "").lower() == target.lower():
            url = asset["browser_download_url"]
            break

    if not url:
        return None  # no matching asset — can't auto-update

    return {"tag": tag, "download_url": url, "body": body}


def download_zip(url: str, dest_path: str, progress_cb=None) -> None:
    """
    Stream-download *url* to *dest_path*.

    *progress_cb* is called with (bytes_done: int, total_bytes: int).
    total_bytes may be 0 if the server does not send Content-Length.
    """
    req = urllib.request.Request(url, headers=_HEADERS)
    with urllib.request.urlopen(req, timeout=90) as resp:
        total  = int(resp.headers.get("Content-Length") or 0)
        done   = 0
        chunk  = 65_536
        with open(dest_path, "wb") as f:
            while True:
                buf = resp.read(chunk)
                if not buf:
                    break
                f.write(buf)
                done += len(buf)
                if progress_cb:
                    progress_cb(done, total)


# ── Update application ────────────────────────────────────────────────────────

def apply_update_windows(zip_path: str, app_exe: str) -> None:
    """
    Extract *zip_path* to a temp folder, write a .bat that swaps the files
    while the app is not running, re-launches the EXE, then self-deletes.

    The CI produces  Compress-Archive -Path dist\\QuoteGenerator\\* ...
    so the zip root already contains QuoteGenerator.exe directly.

    Call this immediately before quitting the application.
    """
    app_dir    = os.path.dirname(os.path.abspath(app_exe))
    update_dir = os.path.join(tempfile.gettempdir(), "mcmx_update")

    # Clear any leftover temp from a previous attempt
    if os.path.isdir(update_dir):
        import shutil
        shutil.rmtree(update_dir, ignore_errors=True)

    with zipfile.ZipFile(zip_path, "r") as zf:
        zf.extractall(update_dir)

    bat_path = os.path.join(tempfile.gettempdir(), "mcmx_updater.bat")
    # Use short-name-safe quoting; robocopy handles sub-directories recursively
    bat = (
        "@echo off\n"
        "REM MCMX Quote Generator auto-updater — do not close\n"
        "timeout /t 2 /nobreak >nul\n"
        f'robocopy "{update_dir}" "{app_dir}" /E /IS /IT /NJH /NJS /NDL /NC /NS >nul 2>&1\n'
        f'start "" "{app_exe}"\n'
        f'rd /s /q "{update_dir}"\n'
        f'del /f /q "{zip_path}"\n'
        'del "%~f0"\n'
    )
    with open(bat_path, "w", encoding="ascii", errors="replace") as f:
        f.write(bat)

    import subprocess
    subprocess.Popen(
        ["cmd.exe", "/c", bat_path],
        creationflags=subprocess.CREATE_NO_WINDOW,
        close_fds=True,
    )


# ── Utility ───────────────────────────────────────────────────────────────────

def is_bundled() -> bool:
    """True when the process is running inside a PyInstaller bundle."""
    return getattr(sys, "_MEIPASS", None) is not None
