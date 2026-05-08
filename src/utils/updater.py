from __future__ import annotations

import hashlib
import json
import os
import subprocess
import sys
import tempfile
import time
import urllib.error
import urllib.request
from dataclasses import dataclass


DEFAULT_REPO = "h4k33my/PDF-EXCEL"
DEFAULT_EXE_ASSET = "GAC-PDF-EXCEL-CONVERTER.exe"
DEFAULT_SHA_ASSET = "GAC-PDF-EXCEL-CONVERTER.exe.sha256"


def _strip_v(v: str) -> str:
    v = (v or "").strip()
    return v[1:] if v.lower().startswith("v") else v


def parse_version_tuple(v: str) -> tuple[int, ...]:
    """
    Very small semver-ish parser: '1.2.3' -> (1,2,3). Non-numeric parts are ignored.
    """
    v = _strip_v(v)
    parts = []
    for tok in (v or "").replace("-", ".").split("."):
        tok = tok.strip()
        if not tok:
            continue
        num = ""
        for ch in tok:
            if ch.isdigit():
                num += ch
            else:
                break
        if num:
            parts.append(int(num))
    return tuple(parts) if parts else (0,)


def is_newer_version(current: str, latest: str) -> bool:
    return parse_version_tuple(latest) > parse_version_tuple(current)


def _http_get_json(url: str, *, timeout_s: int = 12) -> dict:
    req = urllib.request.Request(
        url,
        headers={
            "Accept": "application/vnd.github+json",
            "User-Agent": "GAC-PDF-EXCEL-CONVERTER",
        },
    )
    with urllib.request.urlopen(req, timeout=timeout_s) as resp:
        data = resp.read()
    return json.loads(data.decode("utf-8", errors="replace"))


def _http_download(url: str, dest_path: str, *, timeout_s: int = 30) -> None:
    os.makedirs(os.path.dirname(dest_path), exist_ok=True)
    req = urllib.request.Request(
        url,
        headers={"User-Agent": "GAC-PDF-EXCEL-CONVERTER"},
    )
    with urllib.request.urlopen(req, timeout=timeout_s) as resp:
        with open(dest_path, "wb") as f:
            while True:
                chunk = resp.read(1024 * 128)
                if not chunk:
                    break
                f.write(chunk)


def sha256_file(path: str) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1024 * 128), b""):
            h.update(chunk)
    return h.hexdigest()


def _parse_sha256_file(text: str) -> str | None:
    # Support common formats: "<hash>  filename" or just "<hash>"
    for line in (text or "").splitlines():
        line = line.strip()
        if not line:
            continue
        token = line.split()[0].strip()
        if len(token) == 64 and all(c in "0123456789abcdefABCDEF" for c in token):
            return token.lower()
    return None


@dataclass
class ReleaseInfo:
    tag_name: str
    html_url: str
    exe_download_url: str
    sha_download_url: str | None


def get_latest_release_info(
    *,
    repo: str = DEFAULT_REPO,
    exe_asset_name: str = DEFAULT_EXE_ASSET,
    sha_asset_name: str = DEFAULT_SHA_ASSET,
) -> ReleaseInfo:
    url = f"https://api.github.com/repos/{repo}/releases/latest"
    data = _http_get_json(url)
    tag = str(data.get("tag_name") or "").strip()
    html = str(data.get("html_url") or "").strip()
    assets = data.get("assets") or []
    exe_url = ""
    sha_url = None
    for a in assets:
        name = str(a.get("name") or "").strip()
        dl = str(a.get("browser_download_url") or "").strip()
        if not dl:
            continue
        if name == exe_asset_name:
            exe_url = dl
        elif name == sha_asset_name:
            sha_url = dl

    if not exe_url:
        for a in assets:
            name = str(a.get("name") or "").strip().lower()
            dl = str(a.get("browser_download_url") or "").strip()
            if not dl:
                continue
            if exe_asset_name.lower() in name or name.endswith(".exe"):
                exe_url = dl
                break

    if not sha_url:
        for a in assets:
            name = str(a.get("name") or "").strip().lower()
            dl = str(a.get("browser_download_url") or "").strip()
            if not dl:
                continue
            if sha_asset_name.lower() in name or name.endswith(".sha256"):
                sha_url = dl
                break

    if not tag or not exe_url:
        raise RuntimeError("Latest release missing required tag or exe asset.")
    return ReleaseInfo(tag_name=tag, html_url=html, exe_download_url=exe_url, sha_download_url=sha_url)


def download_release_exe_to_temp(release: ReleaseInfo) -> tuple[str, str | None]:
    """
    Returns (exe_path, sha_path_or_none).
    """
    root = os.path.join(tempfile.gettempdir(), "GAC_PDF_EXCEL_UPDATER")
    os.makedirs(root, exist_ok=True)
    exe_path = os.path.join(root, DEFAULT_EXE_ASSET)
    sha_path = os.path.join(root, DEFAULT_SHA_ASSET) if release.sha_download_url else None
    _http_download(release.exe_download_url, exe_path)
    if release.sha_download_url and sha_path:
        _http_download(release.sha_download_url, sha_path)
    return exe_path, sha_path


def verify_download(exe_path: str, sha_path: str | None) -> tuple[bool, str]:
    if not os.path.isfile(exe_path):
        return False, "Downloaded exe not found."
    actual = sha256_file(exe_path)
    if not sha_path or not os.path.isfile(sha_path):
        return True, f"Downloaded (no sha256 file). sha256={actual}"
    with open(sha_path, "r", encoding="utf-8", errors="replace") as f:
        expected = _parse_sha256_file(f.read())
    if not expected:
        return False, "sha256 file could not be parsed."
    if expected != actual:
        return False, f"sha256 mismatch. expected={expected} actual={actual}"
    return True, "sha256 OK"


def _write_replace_script(
    *,
    new_exe: str,
    target_exe: str,
    start_args: list[str] | None = None,
) -> str:
    """
    Writes a .bat script that waits briefly, replaces target exe, restarts it.
    Returns path to script.
    """
    root = os.path.dirname(new_exe)
    script_path = os.path.join(root, "apply_update.bat")
    backup = target_exe + ".old"
    args = start_args or []
    arg_str = " ".join(f"\"{a}\"" for a in args)
    # Use ping for a small delay (works without PowerShell dependencies).
    content = "\n".join(
        [
            "@echo off",
            "setlocal",
            "REM wait for app to exit",
            "ping 127.0.0.1 -n 3 >nul",
            f"if exist \"{backup}\" del /f /q \"{backup}\"",
            f"if exist \"{target_exe}\" move /y \"{target_exe}\" \"{backup}\" >nul",
            f"move /y \"{new_exe}\" \"{target_exe}\" >nul",
            "REM restart",
            f"start \"\" \"{target_exe}\" {arg_str}",
            "REM cleanup backup after a moment (best effort)",
            "ping 127.0.0.1 -n 3 >nul",
            f"if exist \"{backup}\" del /f /q \"{backup}\"",
            "endlocal",
        ]
    )
    with open(script_path, "w", encoding="utf-8") as f:
        f.write(content)
    return script_path


def apply_update_and_restart(downloaded_exe: str) -> None:
    """
    For frozen builds: replace the currently running exe and restart.
    """
    if not getattr(sys, "frozen", False):
        raise RuntimeError("apply_update_and_restart is only supported for frozen builds.")
    target_exe = os.path.abspath(sys.executable)
    script = _write_replace_script(new_exe=downloaded_exe, target_exe=target_exe)
    # Detach cmd so current process can exit.
    subprocess.Popen(
        ["cmd", "/c", script],
        creationflags=subprocess.CREATE_NEW_PROCESS_GROUP | subprocess.DETACHED_PROCESS,
        close_fds=True,
        cwd=os.path.dirname(script),
    )


def safe_check_latest(
    *,
    repo: str,
    current_version: str,
) -> tuple[ReleaseInfo | None, str | None]:
    try:
        rel = get_latest_release_info(repo=repo)
        if not is_newer_version(current_version, rel.tag_name):
            return None, None
        return rel, None
    except urllib.error.HTTPError as e:
        return None, f"HTTP error checking updates: {e.code}"
    except urllib.error.URLError:
        return None, "Network error checking updates."
    except Exception as e:
        return None, str(e)
