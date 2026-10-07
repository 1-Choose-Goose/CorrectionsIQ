from __future__ import annotations

import argparse
import ctypes
import shutil
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path
from tkinter import BOTH, X, Tk, ttk
import tkinter as tk


BG = "#0B1220"
PANEL = "#111827"
TEXT = "#F8FAFC"
MUTED = "#AAB4C3"
BLUE = "#2B7BBB"


def wait_for_process(pid: int, timeout_ms: int = 120_000) -> None:
    if pid <= 0 or sys.platform != "win32":
        return
    synchronize = 0x00100000
    handle = ctypes.windll.kernel32.OpenProcess(synchronize, False, pid)
    if not handle:
        return
    try:
        ctypes.windll.kernel32.WaitForSingleObject(handle, timeout_ms)
    finally:
        ctypes.windll.kernel32.CloseHandle(handle)


class InstallerWindow:
    def __init__(self) -> None:
        self.root = Tk()
        self.root.title("CorrectionsIQ")
        self.root.geometry("520x190")
        self.root.configure(bg=BG)
        self.root.resizable(False, False)
        self._center()

        frame = tk.Frame(self.root, bg=PANEL, padx=22, pady=20)
        frame.pack(fill=BOTH, expand=True, padx=14, pady=14)

        self.title = tk.Label(
            frame,
            text="Установка обновления",
            fg=TEXT,
            bg=PANEL,
            font=("Segoe UI", 16, "bold"),
            anchor="w",
        )
        self.title.pack(fill=X)

        self.status = tk.Label(
            frame,
            text="Подготовка...",
            fg=MUTED,
            bg=PANEL,
            font=("Segoe UI", 10),
            anchor="w",
        )
        self.status.pack(fill=X, pady=(12, 12))

        style = ttk.Style(self.root)
        style.theme_use("clam")
        style.configure(
            "CorrectionsIQ.Horizontal.TProgressbar",
            troughcolor="#1F2A3A",
            background=BLUE,
            bordercolor="#1F2A3A",
            lightcolor=BLUE,
            darkcolor=BLUE,
        )
        self.progress = ttk.Progressbar(
            frame,
            mode="determinate",
            maximum=100,
            style="CorrectionsIQ.Horizontal.TProgressbar",
        )
        self.progress.pack(fill=X)

    def _center(self) -> None:
        self.root.update_idletasks()
        width = 520
        height = 190
        left = max(0, (self.root.winfo_screenwidth() - width) // 2)
        top = max(0, (self.root.winfo_screenheight() - height) // 2)
        self.root.geometry(f"{width}x{height}+{left}+{top}")

    def update(self, text: str, percent: int) -> None:
        self.status.configure(text=text)
        self.progress.configure(value=max(0, min(100, percent)))
        self.root.update_idletasks()


def source_root(extracted: Path) -> Path:
    wrapped = extracted / "CorrectionsIQ"
    if (wrapped / "CorrectionsIQ.exe").exists():
        return wrapped
    return extracted


def install(zip_path: Path, target: Path, app: Path, window: InstallerWindow) -> None:
    extract_dir = Path(tempfile.gettempdir()) / "CorrectionsIQ" / "extracted_update"
    if extract_dir.exists():
        shutil.rmtree(extract_dir, ignore_errors=True)
    extract_dir.mkdir(parents=True, exist_ok=True)

    window.update("Распаковка обновления...", 15)
    with zipfile.ZipFile(zip_path) as archive:
        archive.extractall(extract_dir)

    src = source_root(extract_dir)
    window.update("Замена файлов программы...", 45)
    internal = target / "_internal"
    if internal.exists():
        shutil.rmtree(internal, ignore_errors=True)

    items = list(src.iterdir())
    total = max(1, len(items))
    for index, item in enumerate(items, start=1):
        destination = target / item.name
        if item.is_dir():
            shutil.copytree(item, destination, dirs_exist_ok=True)
        else:
            shutil.copy2(item, destination)
        window.update("Установка файлов обновления...", 45 + int(index / total * 40))

    window.update("Запуск обновленной программы...", 95)
    subprocess.Popen([str(app)], cwd=str(target), close_fds=True)
    window.update("Готово", 100)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser()
    parser.add_argument("--zip", required=True, type=Path)
    parser.add_argument("--target", required=True, type=Path)
    parser.add_argument("--app", required=True, type=Path)
    parser.add_argument("--pid", required=True, type=int)
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    window = InstallerWindow()
    try:
        window.update("Ожидание закрытия CorrectionsIQ...", 5)
        wait_for_process(args.pid)
        install(args.zip, args.target, args.app, window)
        window.root.after(800, window.root.destroy)
        window.root.mainloop()
        return 0
    except Exception as exc:
        window.update(f"Ошибка установки: {exc}", 100)
        window.root.mainloop()
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
