from __future__ import annotations

import hashlib
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import time
import urllib.request
from collections.abc import Callable
from dataclasses import dataclass
from pathlib import Path
from typing import Literal
from urllib.parse import urlparse

LATEST_RELEASE_URL = "https://api.github.com/repos/1-Choose-Goose/CorrectionsIQ/releases/latest"
RELEASE_DOWNLOAD_PREFIX = "/1-Choose-Goose/CorrectionsIQ/releases/download/"
MAX_DOWNLOAD_SIZE = 500_000_000


@dataclass(frozen=True)
class UpdateInfo:
    version: str
    page_url: str
    asset_url: str
    asset_digest: str = ""
    asset_size: int = 0


UpdatePhase = Literal["idle", "available", "downloading", "installing", "error"]


@dataclass
class UpdateState:
    phase: UpdatePhase = "idle"
    info: UpdateInfo | None = None
    received: int = 0
    total: int = 0
    error: str = ""

    @property
    def busy(self) -> bool:
        return self.phase in {"downloading", "installing"}

    @property
    def progress(self) -> float:
        if self.phase == "installing":
            return 1.0
        if self.total <= 0:
            return 0.0
        return min(1.0, max(0.0, self.received / self.total))

    @property
    def progress_visible(self) -> bool:
        return self.phase in {"downloading", "installing"}

    @property
    def status_text(self) -> str:
        version = self.info.version if self.info else ""
        if self.phase == "available":
            return f"Доступно обновление v{version}"
        if self.phase == "downloading":
            if self.total > 0:
                return f"Скачивание обновления... {int(self.progress * 100)}%"
            return f"Скачивание обновления v{version}..."
        if self.phase == "installing":
            return "Подготовка установки обновления..."
        if self.phase == "error":
            return f"Не удалось установить обновление: {self.error}"
        return ""

    @property
    def button_text(self) -> str:
        if self.phase == "downloading":
            return "Загрузка..."
        if self.phase == "installing":
            return "Установка..."
        if self.phase == "error":
            return "Повторить"
        return "Обновить"

    def mark_available(self, info: UpdateInfo) -> None:
        self.phase = "available"
        self.info = info
        self.received = 0
        self.total = 0
        self.error = ""

    def start_download(self) -> None:
        if self.info is None:
            raise RuntimeError("Сведения об обновлении отсутствуют.")
        self.phase = "downloading"
        self.received = 0
        self.total = 0
        self.error = ""

    def update_download(self, received: int, total: int) -> None:
        self.phase = "downloading"
        self.received = max(0, received)
        self.total = max(0, total)

    def mark_installing(self) -> None:
        self.phase = "installing"

    def mark_error(self, message: str) -> None:
        self.phase = "error"
        self.error = message.strip() or "Неизвестная ошибка"


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
    asset_digest = ""
    asset_size = 0
    expected_name = f"CorrectionsIQ-v{latest_version}.zip".lower()
    for asset in payload.get("assets") or []:
        name = str(asset.get("name") or "")
        if name.lower() == expected_name:
            asset_url = str(asset.get("browser_download_url") or "")
            asset_digest = str(asset.get("digest") or "")
            asset_size = int(asset.get("size") or 0)
            break

    if not asset_url:
        return None
    validate_release_asset_url(asset_url)

    return UpdateInfo(
        version=latest_version,
        page_url=str(payload.get("html_url") or ""),
        asset_url=asset_url,
        asset_digest=asset_digest,
        asset_size=asset_size,
    )


def application_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def application_exe() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve()
    return Path(sys.executable).resolve()


def validate_release_asset_url(url: str) -> None:
    parsed = urlparse(url)
    if (
        parsed.scheme != "https"
        or parsed.hostname != "github.com"
        or parsed.port not in (None, 443)
        or not parsed.path.startswith(RELEASE_DOWNLOAD_PREFIX)
        or parsed.username is not None
        or parsed.password is not None
    ):
        raise RuntimeError("Релиз содержит недопустимый адрес файла обновления.")


def _validate_download_response_url(url: str) -> None:
    parsed = urlparse(url)
    allowed_hosts = {
        "github.com",
        "objects.githubusercontent.com",
        "release-assets.githubusercontent.com",
    }
    if parsed.scheme != "https" or parsed.hostname not in allowed_hosts:
        raise RuntimeError("Сервер перенаправил загрузку на недопустимый адрес.")


def verify_download(path: Path, update: UpdateInfo) -> None:
    actual_size = path.stat().st_size
    if update.asset_size > 0 and actual_size != update.asset_size:
        raise RuntimeError("Размер скачанного обновления не совпадает с данными GitHub.")
    if not update.asset_digest.lower().startswith("sha256:"):
        raise RuntimeError("GitHub не предоставил контрольную сумму обновления.")
    expected_digest = update.asset_digest.split(":", 1)[1].strip().lower()
    hasher = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            hasher.update(chunk)
    if hasher.hexdigest() != expected_digest:
        raise RuntimeError("Не совпадает контрольная сумма скачанного обновления.")


def download_update(update: UpdateInfo, progress: ProgressCallback | None = None) -> Path:
    if not update.asset_url:
        raise RuntimeError("В релизе не найден архив обновления.")
    validate_release_asset_url(update.asset_url)
    if update.asset_size > MAX_DOWNLOAD_SIZE:
        raise RuntimeError("Файл обновления превышает допустимый размер.")

    updates_dir = Path(tempfile.gettempdir()) / "CorrectionsIQ" / "updates"
    updates_dir.mkdir(parents=True, exist_ok=True)
    target = updates_dir / f"CorrectionsIQ-v{update.version}.zip"
    partial = target.with_suffix(".zip.part")
    request = urllib.request.Request(update.asset_url, headers={"User-Agent": "CorrectionsIQ"})

    try:
        with urllib.request.urlopen(request, timeout=30) as response:
            _validate_download_response_url(response.geturl())
            total = int(response.headers.get("Content-Length") or update.asset_size or 0)
            if total > MAX_DOWNLOAD_SIZE:
                raise RuntimeError("Файл обновления превышает допустимый размер.")
            received = 0
            with partial.open("wb") as stream:
                while True:
                    chunk = response.read(1024 * 256)
                    if not chunk:
                        break
                    received += len(chunk)
                    if received > MAX_DOWNLOAD_SIZE:
                        raise RuntimeError("Файл обновления превышает допустимый размер.")
                    stream.write(chunk)
                    if progress:
                        progress(received, total)
        verify_download(partial, update)
        partial.replace(target)
    except Exception:
        partial.unlink(missing_ok=True)
        raise

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
