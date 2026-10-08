from __future__ import annotations

import argparse
import ctypes
import os
import shutil
import stat
import subprocess
import sys
import tempfile
import zipfile
from contextlib import suppress
from pathlib import Path
from tkinter import TclError
from typing import Protocol
from uuid import uuid4

import customtkinter as ctk

BG = "#0B1220"
PANEL = "#111827"
BORDER = "#1F2A3A"
TEXT = "#F8FAFC"
MUTED = "#AAB4C3"
BLUE = "#2B7BBB"
ERROR = "#EF4444"
MAX_ARCHIVE_FILES = 20_000
MAX_UNCOMPRESSED_SIZE = 1_000_000_000


class ProgressWindow(Protocol):
    def update(self, text: str, percent: int) -> None: ...


def resource_path(relative_path: str) -> Path:
    bundle_root = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
    return bundle_root / relative_path


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
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")
        self.root = ctk.CTk(fg_color=BG)
        self.root.title("CorrectionsIQ")
        self.root.geometry("560x220")
        self.root.resizable(False, False)
        self.root.protocol("WM_DELETE_WINDOW", lambda: None)
        self._set_icon()
        self._center(560, 220)

        frame = ctk.CTkFrame(
            self.root,
            fg_color=PANEL,
            corner_radius=12,
            border_width=1,
            border_color=BORDER,
        )
        frame.pack(fill="both", expand=True, padx=18, pady=18)

        self.title = ctk.CTkLabel(
            frame,
            text="Установка обновления",
            text_color=TEXT,
            font=ctk.CTkFont(family="Segoe UI", size=20, weight="bold"),
            anchor="w",
        )
        self.title.pack(fill="x", padx=24, pady=(22, 0))

        self.status = ctk.CTkLabel(
            frame,
            text="Подготовка...",
            text_color=MUTED,
            font=ctk.CTkFont(family="Segoe UI", size=13),
            anchor="w",
        )
        self.status.pack(fill="x", padx=24, pady=(18, 12))

        self.progress = ctk.CTkProgressBar(
            frame,
            height=9,
            corner_radius=5,
            fg_color=BORDER,
            progress_color=BLUE,
        )
        self.progress.pack(fill="x", padx=24, pady=(0, 10))
        self.progress.set(0)

        self.percent = ctk.CTkLabel(
            frame,
            text="0%",
            text_color=MUTED,
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
            anchor="e",
        )
        self.percent.pack(fill="x", padx=24, pady=(0, 18))

        self.error_text = ctk.CTkTextbox(
            frame,
            fg_color=BG,
            border_width=1,
            border_color=BORDER,
            corner_radius=8,
            text_color=TEXT,
            font=ctk.CTkFont(family="Segoe UI", size=13),
            wrap="word",
        )
        self.close_button = ctk.CTkButton(
            frame,
            text="Закрыть",
            width=120,
            height=34,
            corner_radius=8,
            fg_color=BLUE,
            hover_color="#236AA3",
            font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"),
            command=self.root.destroy,
        )

    def _set_icon(self) -> None:
        icon_path = resource_path("assets/app_icon.ico")
        if not icon_path.is_file():
            return

        def apply_icon() -> None:
            with suppress(OSError, TclError):
                self.root.iconbitmap(str(icon_path))

        apply_icon()
        self.root.after(200, apply_icon)

    def _center(self, width: int, height: int) -> None:
        self.root.update_idletasks()
        left = max(0, (self.root.winfo_screenwidth() - width) // 2)
        top = max(0, (self.root.winfo_screenheight() - height) // 2)
        self.root.geometry(f"{width}x{height}+{left}+{top}")

    def update(self, text: str, percent: int) -> None:
        value = max(0, min(100, percent))
        self.status.configure(text=text)
        self.progress.set(value / 100)
        self.percent.configure(text=f"{value}%")
        self.root.update()

    def show_error(self, text: str) -> None:
        self.title.configure(text="Не удалось установить обновление", text_color=ERROR)
        self.status.pack_forget()
        self.progress.pack_forget()
        self.percent.pack_forget()
        self.error_text.configure(state="normal")
        self.error_text.delete("1.0", "end")
        self.error_text.insert("1.0", text.strip() or "Неизвестная ошибка")
        self.error_text.configure(state="disabled")
        self.error_text.pack(fill="both", expand=True, padx=24, pady=(16, 12))
        self.close_button.pack(anchor="e", padx=24, pady=(0, 20))
        self.root.resizable(True, True)
        self.root.minsize(560, 320)
        self._center(640, 360)
        self.root.protocol("WM_DELETE_WINDOW", self.root.destroy)
        self.root.update()


def source_root(extracted: Path) -> Path:
    wrapped = extracted / "CorrectionsIQ"
    if (wrapped / "CorrectionsIQ.exe").exists():
        source = wrapped
    elif (extracted / "CorrectionsIQ.exe").exists():
        source = extracted
    else:
        raise RuntimeError("Архив обновления не содержит CorrectionsIQ.exe.")
    if not (source / "_internal").is_dir():
        raise RuntimeError("Архив обновления не содержит папку _internal.")
    return source


def validate_install_paths(zip_path: Path, target: Path, app: Path) -> tuple[Path, Path, Path]:
    archive = zip_path.resolve(strict=True)
    install_dir = target.resolve(strict=True)
    executable = app.resolve(strict=False)
    if not install_dir.is_dir():
        raise RuntimeError("Папка программы не найдена.")
    if executable.parent != install_dir or executable.name.lower() != "correctionsiq.exe":
        raise RuntimeError("Файл запуска должен находиться в папке программы.")
    if not executable.is_file():
        raise RuntimeError("Не найден файл CorrectionsIQ.exe в папке программы.")
    return archive, install_dir, executable


def safe_extract_archive(zip_path: Path, destination: Path) -> None:
    destination.mkdir(parents=True, exist_ok=True)
    root = destination.resolve()
    with zipfile.ZipFile(zip_path) as archive:
        members = archive.infolist()
        if len(members) > MAX_ARCHIVE_FILES:
            raise RuntimeError("Архив обновления содержит слишком много файлов.")
        if sum(member.file_size for member in members) > MAX_UNCOMPRESSED_SIZE:
            raise RuntimeError("Архив обновления слишком большой.")

        for member in members:
            normalized_name = member.filename.replace("\\", "/")
            destination_path = (root / normalized_name).resolve()
            if not destination_path.is_relative_to(root):
                raise RuntimeError("Архив обновления содержит недопустимый путь.")
            file_type = (member.external_attr >> 16) & 0o170000
            if file_type == stat.S_IFLNK:
                raise RuntimeError("Архив обновления содержит недопустимую ссылку.")

        archive.extractall(root)


def prepare_staged_install(source: Path, target: Path, staging: Path) -> None:
    if not (source / "CorrectionsIQ.exe").is_file() or not (source / "_internal").is_dir():
        raise RuntimeError("Файлы обновления имеют неверную структуру.")
    if staging.exists():
        shutil.rmtree(staging)

    shutil.copytree(target, staging)
    staged_internal = staging / "_internal"
    if staged_internal.exists():
        shutil.rmtree(staged_internal)
    shutil.copytree(source / "_internal", staged_internal)

    for item in source.iterdir():
        if item.name == "_internal":
            continue
        destination = staging / item.name
        if item.is_dir():
            shutil.copytree(item, destination, dirs_exist_ok=True)
        else:
            shutil.copy2(item, destination)

    legacy_helper = staging / "CorrectionsIQUpdater.exe"
    if legacy_helper.exists():
        legacy_helper.unlink()


def install(zip_path: Path, target: Path, app: Path, window: ProgressWindow) -> None:
    zip_path, target, app = validate_install_paths(zip_path, target, app)
    operation_id = uuid4().hex
    work_root = Path(tempfile.gettempdir()) / "CorrectionsIQ" / f"update-{operation_id}"
    extract_dir = work_root / "extracted"
    staging = target.parent / f".{target.name}-update-{operation_id}"
    backup = target.parent / f".{target.name}-backup-{operation_id}"
    work_root.mkdir(parents=True, exist_ok=False)

    try:
        window.update("Проверка и распаковка обновления...", 15)
        safe_extract_archive(zip_path, extract_dir)
        source = source_root(extract_dir)

        window.update("Подготовка файлов обновления...", 40)
        prepare_staged_install(source, target, staging)

        window.update("Установка обновления...", 75)
        os.replace(target, backup)
        try:
            os.replace(staging, target)
        except Exception:
            os.replace(backup, target)
            raise

        new_app = target / app.name
        try:
            window.update("Запуск обновленной программы...", 95)
            subprocess.Popen([str(new_app)], cwd=str(target), close_fds=True)
        except Exception:
            shutil.rmtree(target, ignore_errors=True)
            os.replace(backup, target)
            raise

        shutil.rmtree(backup, ignore_errors=True)
        window.update("Обновление установлено", 100)
    finally:
        shutil.rmtree(work_root, ignore_errors=True)
        if staging.exists():
            shutil.rmtree(staging, ignore_errors=True)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser()
    parser.add_argument("--zip", required=True, type=Path)
    parser.add_argument("--target", required=True, type=Path)
    parser.add_argument("--app", required=True, type=Path)
    parser.add_argument("--pid", required=True, type=int)
    return parser.parse_args()


def main() -> int:
    os.chdir(tempfile.gettempdir())
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
        window.show_error(str(exc))
        window.root.mainloop()
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
