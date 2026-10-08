import platform
import subprocess
import sys
import traceback
from importlib.util import find_spec

REQUIRED_PACKAGES = {
    "docx": "python-docx==1.2.0",
    "customtkinter": "customtkinter==6.0.0",
    "openpyxl": "openpyxl==3.1.5",
}


def configure_windows_app_id() -> None:
    if sys.platform != "win32":
        return
    try:
        import ctypes

        app_id = "correctionsiq.app"
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(app_id)
    except Exception:
        pass


def check_system_compatibility() -> None:
    if sys.platform != "win32":
        return

    version = sys.getwindowsversion()
    if version.major < 10:
        raise RuntimeError(
            "CorrectionsIQ рассчитан на Windows 10/11. "
            f"Обнаружена система: {platform.platform()}. "
            "На Windows XP/7/8 современный графический интерфейс может не запускаться из-за отсутствующих системных компонентов."
        )


def ensure_dependencies() -> None:
    if getattr(sys, "frozen", False):
        return

    missing = [package for module, package in REQUIRED_PACKAGES.items() if find_spec(module) is None]
    if not missing:
        return

    startupinfo = None
    if sys.platform == "win32":
        startupinfo = subprocess.STARTUPINFO()
        startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW

    def run_pip(args: list[str]) -> None:
        subprocess.check_call(
            [sys.executable, "-m", *args],
            startupinfo=startupinfo,
        )

    try:
        run_pip(["pip", "--version"])
    except Exception:
        run_pip(["ensurepip", "--upgrade"])

    install_args = [
        "pip",
        "install",
        "--disable-pip-version-check",
        "--quiet",
        *missing,
    ]
    try:
        run_pip(install_args)
    except subprocess.CalledProcessError:
        run_pip([*install_args[:4], "--user", *install_args[4:]])


def run() -> int:
    configure_windows_app_id()
    check_system_compatibility()
    ensure_dependencies()

    if len(sys.argv) > 1:
        from average_headcount import main as cli_main

        return cli_main(sys.argv[1:])

    from app_gui import run_app

    run_app()
    return 0


def show_startup_error(exc: BaseException) -> None:
    if not getattr(sys, "frozen", False):
        raise exc

    log_path = sys.executable.rsplit("\\", 1)[0] + "\\CorrectionsIQ_error.log"
    error_text = "".join(traceback.format_exception(type(exc), exc, exc.__traceback__))
    try:
        with open(log_path, "w", encoding="utf-8") as stream:
            stream.write(error_text)
    except Exception:
        log_path = ""

    try:
        import tkinter.messagebox as messagebox

        details = f"\n\nПодробности сохранены: {log_path}" if log_path else ""
        messagebox.showerror(
            "CorrectionsIQ",
            "Программа не смогла запуститься на этом компьютере." + details,
        )
    except Exception:
        pass


if __name__ == "__main__":
    try:
        raise SystemExit(run())
    except Exception as startup_error:
        show_startup_error(startup_error)
        raise SystemExit(1) from startup_error
