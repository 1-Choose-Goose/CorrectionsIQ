from __future__ import annotations

import json
import re
import urllib.request
from dataclasses import dataclass


LATEST_RELEASE_URL = "https://api.github.com/repos/1-Choose-Goose/CorrectionsIQ/releases/latest"


@dataclass(frozen=True)
class UpdateInfo:
    version: str
    page_url: str
    asset_url: str


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
