from __future__ import annotations

import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import time
import urllib.request
from dataclasses import dataclass
from pathlib import Path
from typing import Callable


LATEST_RELEASE_URL = "https://api.github.com/repos/1-Choose-Goose/CorrectionsIQ/releases/latest"


@dataclass(frozen=True)
class UpdateInfo:
    version: str
    page_url: str
    asset_url: str


ProgressCallback = Callable[[int, int], None]


def normalize_version(value: str) -> tuple[int, ...]:
    numbers = re.findall(r"\d+", value)
    return tuple(int(number) for number in numbers) if numbers else (0,)


def is_newer_version(latest: str, current: str) -> bool:
    latest_parts = normalize_version(latest)
    current_parts = normalize_version(current)
    max_len = max(len(latest_parts), len(current_parts))
    latest_parts += (0,) * (max_len - len(latest_parts))
    current_parts += (0,) * (max_len - len(current_parts))
    return latest_parts > current_parts


def check_latest_release(current_version: str, timeout: float = 5.0) -> UpdateInfo | None:
    request = urllib.request.Request(
        LATEST_RELEASE_URL,
        headers={
            "Accept": "application/vnd.github+json",
            "User-Agent": "CorrectionsIQ",
        },
    )
    with urllib.request.urlopen(request, timeout=timeout) as response:
        payload = json.loads(response.read().decode("utf-8"))

    tag_name = str(payload.get("tag_name") or "").strip()
    latest_version = tag_name.removeprefix("v").strip()
    if not latest_version or not is_newer_version(latest_version, current_version):
        return None

    asset_url = ""
    for asset in payload.get("assets") or []:
        name = str(asset.get("name") or "")
        if name.lower().endswith(".zip"):
            asset_url = str(asset.get("browser_download_url") or "")
            break

    return UpdateInfo(
        version=latest_version,
        page_url=str(payload.get("html_url") or ""),
        asset_url=asset_url,
    )


def application_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def application_exe() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve()
    return Path(sys.executable).resolve()


def download_update(update: UpdateInfo, progress: ProgressCallback | None = None) -> Path:
    if not update.asset_url:
        raise RuntimeError("В релизе не найден архив обновления.")

    updates_dir = Path(tempfile.gettempdir()) / "CorrectionsIQ" / "updates"
    updates_dir.mkdir(parents=True, exist_ok=True)
    target = updates_dir / f"CorrectionsIQ-v{update.version}.zip"
    request = urllib.request.Request(update.asset_url, headers={"User-Agent": "CorrectionsIQ"})

    with urllib.request.urlopen(request, timeout=30) as response:
        total = int(response.headers.get("Content-Length") or 0)
        received = 0
        with target.open("wb") as stream:
            while True:
                chunk = response.read(1024 * 256)
                if not chunk:
                    break
                stream.write(chunk)
                received += len(chunk)
                if progress:
                    progress(received, total)

    if progress:
        progress(target.stat().st_size, target.stat().st_size)
    return target


def _copy_helper_to_temp(helper_path: Path) -> Path:
    temp_helper_dir = Path(tempfile.gettempdir()) / "CorrectionsIQ" / "installer"
    temp_helper_dir.mkdir(parents=True, exist_ok=True)
    target = temp_helper_dir / f"CorrectionsIQUpdater-{int(time.time())}.exe"
    shutil.copy2(helper_path, target)
    return target


def launch_update_installer(zip_path: Path) -> None:
    root = application_dir()
    helper = root / "_internal" / "CorrectionsIQUpdater.exe"
    if not helper.exists():
        helper = root / "CorrectionsIQUpdater.exe"
    if not helper.exists():
        raise RuntimeError("Не найден установщик обновлений CorrectionsIQUpdater.exe.")

    helper_to_run = _copy_helper_to_temp(helper)
    creationflags = 0
    if os.name == "nt":
        creationflags = subprocess.CREATE_NEW_PROCESS_GROUP | subprocess.DETACHED_PROCESS

    subprocess.Popen(
        [
            str(helper_to_run),
            "--zip",
            str(zip_path),
            "--target",
            str(root),
            "--app",
            str(application_exe()),
            "--pid",
            str(os.getpid()),
        ],
        close_fds=True,
        creationflags=creationflags,
    )
