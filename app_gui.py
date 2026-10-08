from __future__ import annotations

import queue
import re
import threading
import tkinter as tk
from contextlib import suppress
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from tkinter import filedialog, messagebox
from typing import Any

import customtkinter as ctk

from average_headcount import (
    Segment,
    build_segments,
    parse_ru_date,
    read_records,
    write_xlsx,
)
from updater import (
    UpdateInfo,
    UpdateState,
    check_latest_release,
    download_update,
    launch_update_installer,
)

NAVY = "#111827"
NAVY_DARK = "#0F172A"
BLUE = "#2B7BBB"
PAGE_BG = "#0B1220"
PANEL_BG = "#111827"
BORDER = "#1F2A3A"
TEXT = "#F8FAFC"
MUTED = "#AAB4C3"
FIELD_BG = "#30363A"
APP_VERSION = "1.4"


@dataclass(frozen=True)
class CalculationResult:
    segments: list[Segment]
    total_days: int
    total_weighted: int
    average: float
    institution: str
    start: date
    end: date
    category: str | None


class HeadcountApp(ctk.CTk):
    def __init__(self) -> None:
        super().__init__()
        self.title("CorrectionsIQ")
        self.project_dir = Path(__file__).resolve().parent
        self._set_app_icon()
        self.minsize(1080, 680)

        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")

        self.result: CalculationResult | None = None
        self.worker_queue: queue.Queue[tuple[str, object]] = queue.Queue()
        self.update_queue: queue.Queue[tuple[str, object]] = queue.Queue()
        self.is_running = False
        self.update_state = UpdateState()
        self.calculation_error: str | None = None

        self.folder_var = ctk.StringVar(value="")
        self.institution_var = ctk.StringVar(value="")
        self.start_var = ctk.StringVar(value="")
        self.end_var = ctk.StringVar(value="")
        self.settlement_var = ctk.BooleanVar(value=False)
        self.epkt_var = ctk.BooleanVar(value=False)
        self.active_module = "headcount"

        self._build_layout()
        self.after(0, lambda: self._center_window(1240, 780))
        self.after(120, self._poll_worker)
        self.after(800, self._start_update_check)

    def _center_window(self, width: int, height: int) -> None:
        self.update_idletasks()
        left = max(0, (self.winfo_screenwidth() - width) // 2)
        top = max(0, (self.winfo_screenheight() - height) // 2)
        self.geometry(f"{width}x{height}+{left}+{top}")

    def _set_app_icon(self) -> None:
        icon_path = self.project_dir / "assets" / "app_icon.ico"
        png_path = self.project_dir / "assets" / "app_icon.png"
        if icon_path.exists():
            with suppress(OSError, tk.TclError):
                self.iconbitmap(str(icon_path))
        if png_path.exists():
            try:
                self._icon_image = tk.PhotoImage(file=str(png_path))
                self.iconphoto(True, self._icon_image)
            except (OSError, tk.TclError):
                pass

    def _build_layout(self) -> None:
        self.grid_columnconfigure(0, weight=0)
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        self.sidebar = ctk.CTkFrame(self, width=240, corner_radius=0, fg_color=NAVY_DARK)
        self.sidebar.grid(row=0, column=0, sticky="nsew")
        self.sidebar.grid_propagate(False)
        self.sidebar.grid_columnconfigure(0, weight=1)
        self.sidebar.grid_rowconfigure(8, weight=1)
        self._build_sidebar()

        self.content = ctk.CTkFrame(self, corner_radius=0, fg_color=PAGE_BG)
        self.content.grid(row=0, column=1, sticky="nsew")
        self.content.grid_columnconfigure(0, weight=1)
        self._render_headcount_module()

    def _build_sidebar(self) -> None:
        logo = ctk.CTkFrame(self.sidebar, fg_color=PANEL_BG, corner_radius=8)
        logo.grid(row=0, column=0, sticky="ew", padx=14, pady=(18, 16))
        logo.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(
            logo,
            text="CorrectionsIQ",
            text_color=TEXT,
            anchor="w",
            font=ctk.CTkFont(size=22, weight="bold"),
        ).grid(row=0, column=0, sticky="ew", padx=12, pady=(12, 0))
        ctk.CTkFrame(logo, height=10, fg_color="transparent").grid(row=1, column=0, sticky="ew")

        self.nav_buttons: dict[str, ctk.CTkButton] = {}
        self._nav_button(1, "headcount", "Среднесписочная")
        self._nav_button(2, "about", "О программе")

    def _nav_button(self, row: int, key: str, text: str) -> None:
        button = ctk.CTkButton(
            self.sidebar,
            text=text,
            height=40,
            corner_radius=7,
            fg_color=BLUE if key == self.active_module else NAVY_DARK,
            hover_color="#1C52BF" if key == self.active_module else "#173A85",
            text_color="#FFFFFF",
            anchor="w",
            command=lambda item=key: self._switch_module(item),
        )
        button.grid(row=row, column=0, sticky="ew", padx=16, pady=(0, 8))
        self.nav_buttons[key] = button

    def _switch_module(self, key: str) -> None:
        if key not in {"headcount", "about"}:
            key = "headcount"
        self.active_module = key
        for name, button in self.nav_buttons.items():
            button.configure(
                fg_color=BLUE if name == key else NAVY_DARK,
                hover_color="#1C52BF" if name == key else "#173A85",
            )

        if key == "headcount":
            self._render_headcount_module()
        elif key == "about":
            self._render_about_module()

    def _clear_content(self) -> None:
        for child in self.content.winfo_children():
            child.destroy()

    def _render_headcount_module(self) -> None:
        self._clear_content()
        for row_index in range(8):
            self.content.grid_rowconfigure(row_index, weight=1 if row_index == 5 else 0)
        self._build_header("Среднесписочная")
        self._build_update_panel()
        self._build_parameter_panel()
        self._build_actions()
        self._build_summary()
        self._build_table()
        if self.is_running:
            self._show_empty_state("Расчет выполняется...")
        elif self.result is not None:
            self._show_result(self.result)
        elif self.calculation_error:
            self._render_error(self.calculation_error)
        else:
            self._show_empty_state()
        self._sync_calculation_controls()

    def _render_about_module(self) -> None:
        self._clear_content()
        for row_index in range(8):
            self.content.grid_rowconfigure(row_index, weight=1 if row_index == 1 else 0)
        self._build_header("О программе")

        panel = ctk.CTkFrame(self.content, fg_color=PANEL_BG, corner_radius=12, border_width=1, border_color=BORDER)
        panel.grid(row=1, column=0, sticky="nsew", padx=22, pady=(0, 18))
        panel.grid_columnconfigure(0, weight=1)

        rows = (
            ("Название", "CorrectionsIQ"),
            ("Версия", APP_VERSION),
            ("Разработчик", "Куц Олег Олегович"),
            ("Автор идеи", "Казначеев Андрей Юрьевич"),
        )
        for row_index, (label, value) in enumerate(rows):
            item = ctk.CTkFrame(panel, fg_color="#0B1220", corner_radius=8, border_width=1, border_color=BORDER)
            item.grid(row=row_index, column=0, sticky="ew", padx=22, pady=(18 if row_index == 0 else 0, 10))
            item.grid_columnconfigure(0, weight=1)
            ctk.CTkLabel(
                item,
                text=label,
                anchor="w",
                text_color=MUTED,
                font=ctk.CTkFont(size=13),
            ).grid(row=0, column=0, sticky="ew", padx=16, pady=(12, 2))
            ctk.CTkLabel(
                item,
                text=value,
                anchor="w",
                text_color=TEXT,
                font=ctk.CTkFont(size=18, weight="bold"),
            ).grid(row=1, column=0, sticky="ew", padx=16, pady=(0, 12))

    def _build_header(self, title: str) -> None:
        header = ctk.CTkFrame(self.content, fg_color="transparent")
        header.grid(row=0, column=0, sticky="ew", padx=22, pady=(22, 10))
        header.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(
            header,
            text=title,
            anchor="w",
            text_color=TEXT,
            font=ctk.CTkFont(size=24, weight="bold"),
        ).grid(row=0, column=0, sticky="ew")
        ctk.CTkLabel(
            header,
            text=f"v{APP_VERSION}",
            text_color="#86EFAC",
            fg_color="#123C2B",
            corner_radius=10,
            width=70,
            height=34,
            font=ctk.CTkFont(size=12, weight="bold"),
        ).grid(row=0, column=1, sticky="e", padx=(12, 0))

    def _build_update_panel(self) -> None:
        self.update_panel = ctk.CTkFrame(
            self.content,
            fg_color="#123C2B",
            corner_radius=12,
            border_width=1,
            border_color="#1E6B48",
        )
        self.update_panel.grid(row=1, column=0, sticky="ew", padx=22, pady=(0, 12))
        self.update_panel.grid_columnconfigure(0, weight=1)
        self.update_panel.grid_remove()

        self.update_label = ctk.CTkLabel(
            self.update_panel,
            text="",
            anchor="w",
            text_color="#D1FAE5",
            font=ctk.CTkFont(size=13, weight="bold"),
        )
        self.update_label.grid(row=0, column=0, sticky="ew", padx=16, pady=12)
        self.update_button = ctk.CTkButton(
            self.update_panel,
            text="Обновить",
            width=110,
            height=32,
            corner_radius=6,
            fg_color="#16A34A",
            hover_color="#15803D",
            command=self._start_update_install,
        )
        self.update_button.grid(row=0, column=1, sticky="e", padx=(8, 16), pady=10)
        self.update_progress = ctk.CTkProgressBar(self.update_panel, height=6, corner_radius=4)
        self.update_progress.grid(row=1, column=0, columnspan=2, sticky="ew", padx=16, pady=(0, 12))
        self.update_progress.set(0)
        self.update_progress.grid_remove()
        self._sync_update_panel()

    def _build_parameter_panel(self) -> None:
        panel = ctk.CTkFrame(self.content, fg_color=PANEL_BG, corner_radius=12, border_width=1, border_color=BORDER)
        panel.grid(row=2, column=0, sticky="ew", padx=22, pady=(0, 12))
        panel.grid_columnconfigure(0, weight=2)
        panel.grid_columnconfigure(1, weight=2)
        panel.grid_columnconfigure(2, weight=1)
        panel.grid_columnconfigure(3, weight=1)
        panel.grid_columnconfigure(4, weight=1)

        self._folder_field(panel, 0)
        self._field(panel, 1, "Учреждение", self.institution_var)
        self._field(panel, 2, "Начало", self.start_var)
        self._field(panel, 3, "Конец", self.end_var)

        options = ctk.CTkFrame(panel, fg_color="transparent")
        options.grid(row=0, column=4, rowspan=2, sticky="nsew", padx=(8, 14), pady=(14, 14))
        ctk.CTkCheckBox(
            options,
            text="Колония-поселение",
            variable=self.settlement_var,
            command=self._select_settlement,
            corner_radius=4,
            checkbox_width=18,
            checkbox_height=18,
            text_color=TEXT,
        ).grid(row=0, column=0, sticky="w", pady=(22, 8))
        ctk.CTkCheckBox(
            options,
            text="ЕПКТ",
            variable=self.epkt_var,
            command=self._select_epkt,
            corner_radius=4,
            checkbox_width=18,
            checkbox_height=18,
            text_color=TEXT,
        ).grid(row=1, column=0, sticky="w")

    def _folder_field(self, parent: ctk.CTkFrame, column: int) -> None:
        ctk.CTkLabel(
            parent,
            text="Папка со сводками",
            anchor="w",
            text_color=MUTED,
            font=ctk.CTkFont(size=12),
        ).grid(row=0, column=column, sticky="ew", padx=(14, 8), pady=(14, 5))

        wrapper = ctk.CTkFrame(parent, fg_color="transparent")
        wrapper.grid(row=1, column=column, sticky="ew", padx=(14, 8), pady=(0, 14))
        wrapper.grid_columnconfigure(0, weight=1)
        ctk.CTkEntry(
            wrapper,
            textvariable=self.folder_var,
            height=36,
            corner_radius=6,
            border_color="#5F6A75",
            fg_color=FIELD_BG,
            text_color=TEXT,
        ).grid(row=0, column=0, sticky="ew", padx=(0, 8))
        ctk.CTkButton(
            wrapper,
            text="Выбрать",
            command=self._choose_folder,
            height=34,
            width=92,
            corner_radius=6,
            fg_color=BLUE,
            hover_color="#1C52BF",
        ).grid(row=0, column=1, sticky="e")

    def _field(self, parent: ctk.CTkFrame, column: int, label: str, variable: ctk.StringVar) -> None:
        ctk.CTkLabel(
            parent,
            text=label,
            anchor="w",
            text_color=MUTED,
            font=ctk.CTkFont(size=12),
        ).grid(row=0, column=column, sticky="ew", padx=(14, 8), pady=(14, 5))
        ctk.CTkEntry(
            parent,
            textvariable=variable,
            height=36,
            corner_radius=6,
            border_color="#5F6A75",
            fg_color=FIELD_BG,
            text_color=TEXT,
        ).grid(row=1, column=column, sticky="ew", padx=(14, 8), pady=(0, 14))

    def _build_actions(self) -> None:
        actions = ctk.CTkFrame(self.content, fg_color="transparent")
        actions.grid(row=3, column=0, sticky="ew", padx=22)
        actions.grid_columnconfigure(0, weight=1)
        self.calculate_button = ctk.CTkButton(
            actions,
            text="Рассчитать",
            command=self._start_calculation,
            width=130,
            height=34,
            corner_radius=6,
            fg_color=BLUE,
            hover_color="#1C52BF",
            font=ctk.CTkFont(size=13, weight="bold"),
        )
        self.calculate_button.grid(row=0, column=1, sticky="e", padx=(8, 0), pady=(0, 10))
        self.export_button = ctk.CTkButton(
            actions,
            text="Сохранить Excel",
            command=self._export_xlsx,
            width=140,
            height=34,
            corner_radius=6,
            state="disabled",
            fg_color="#374151",
            text_color=TEXT,
            hover_color="#4B5563",
        )
        self.export_button.grid(row=0, column=2, sticky="e", padx=(8, 0), pady=(0, 10))

    def _build_summary(self) -> None:
        self.summary_frame = ctk.CTkFrame(self.content, fg_color="transparent")
        self.summary_frame.grid(row=4, column=0, sticky="ew", padx=22, pady=(10, 10))
        self.summary_frame.grid_columnconfigure((0, 1, 2), weight=1, uniform="summary")
        self.avg_value = self._summary_card(0, "Среднесписочное", "—")
        self.days_value = self._summary_card(1, "Дней в периоде", "—")
        self.weighted_value = self._summary_card(2, "Человеко-дней", "—")

    def _summary_card(self, column: int, label: str, value: str) -> ctk.CTkLabel:
        card = ctk.CTkFrame(self.summary_frame, corner_radius=12, fg_color=PANEL_BG, border_width=1, border_color=BORDER)
        card.grid(row=0, column=column, sticky="ew", padx=(0, 10) if column < 2 else 0)
        card.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(
            card,
            text=label,
            anchor="w",
            text_color=MUTED,
            font=ctk.CTkFont(size=12, weight="bold"),
        ).grid(row=0, column=0, sticky="ew", padx=16, pady=(13, 3))
        value_label = ctk.CTkLabel(
            card,
            text=value,
            anchor="w",
            text_color=TEXT,
            font=ctk.CTkFont(size=22, weight="bold"),
        )
        value_label.grid(row=1, column=0, sticky="ew", padx=16, pady=(0, 13))
        return value_label

    def _build_table(self) -> None:
        table_panel = ctk.CTkFrame(self.content, corner_radius=12, fg_color=PANEL_BG, border_width=1, border_color=BORDER)
        table_panel.grid(row=5, column=0, sticky="nsew", padx=22, pady=(0, 18))
        table_panel.grid_columnconfigure(0, weight=1)
        table_panel.grid_rowconfigure(1, weight=1)

        table_header = ctk.CTkFrame(table_panel, fg_color="#1B2533", corner_radius=0)
        table_header.grid(row=0, column=0, sticky="ew")
        table_header.grid_columnconfigure((0, 1), weight=1, uniform="table")
        for col, text in enumerate(("Дата", "Количество")):
            ctk.CTkLabel(
                table_header,
                text=text,
                anchor="w",
                text_color=TEXT,
                font=ctk.CTkFont(size=13),
            ).grid(row=0, column=col, sticky="ew", padx=14, pady=12)

        self.rows_frame = ctk.CTkScrollableFrame(
            table_panel,
            corner_radius=0,
            fg_color="#0B1220",
            border_width=1,
            border_color=BORDER,
        )
        self.rows_frame.grid(row=1, column=0, sticky="nsew")
        self.rows_frame.grid_columnconfigure((0, 1), weight=1, uniform="table")

        self.progress = ctk.CTkProgressBar(self.content, height=5, corner_radius=0)
        self.progress.grid(row=6, column=0, sticky="ew", padx=22, pady=(0, 8))
        self.progress.set(0)
        self.status_label = ctk.CTkLabel(
            self.content,
            text="Готово к расчету",
            anchor="w",
            justify="left",
            text_color=MUTED,
            font=ctk.CTkFont(size=12),
        )
        self.status_label.grid(row=7, column=0, sticky="ew", padx=22, pady=(0, 14))

    def _start_update_check(self) -> None:
        worker = threading.Thread(target=self._check_update_in_worker, daemon=True)
        worker.start()
        self.after(250, self._poll_update)

    def _check_update_in_worker(self) -> None:
        try:
            update = check_latest_release(APP_VERSION)
            if update:
                self.update_queue.put(("update", update))
        except Exception:
            self.update_queue.put(("none", None))

    def _poll_update(self) -> None:
        try:
            kind, payload = self.update_queue.get_nowait()
        except queue.Empty:
            self.after(250, self._poll_update)
            return
        if kind == "update":
            if isinstance(payload, UpdateInfo):
                self._show_update(payload)
        elif kind == "download_progress":
            if isinstance(payload, tuple) and len(payload) == 2:
                received, total = payload
                if isinstance(received, int) and isinstance(total, int):
                    self.update_state.update_download(received, total)
                    self._sync_update_panel()
            self.after(100, self._poll_update)
            return
        elif kind == "downloaded":
            if isinstance(payload, Path):
                self._install_downloaded_update(payload)
        elif kind == "update_error":
            self.update_state.mark_error(str(payload))
            self._sync_update_panel()
        self.after(250, self._poll_update)

    def _show_update(self, update: UpdateInfo) -> None:
        self.update_state.mark_available(update)
        self._sync_update_panel()

    @staticmethod
    def _is_live_widget(widget: Any | None) -> bool:
        if widget is None or not hasattr(widget, "winfo_exists"):
            return False
        try:
            return bool(widget.winfo_exists())
        except tk.TclError:
            return False

    def _sync_update_panel(self) -> None:
        panel: Any = getattr(self, "update_panel", None)
        label: Any = getattr(self, "update_label", None)
        button: Any = getattr(self, "update_button", None)
        progress: Any = getattr(self, "update_progress", None)
        if not all(self._is_live_widget(widget) for widget in (panel, label, button, progress)):
            return

        if self.update_state.phase == "idle":
            panel.grid_remove()
            return

        label.configure(text=self.update_state.status_text)
        button.configure(
            text=self.update_state.button_text,
            state="disabled" if self.update_state.busy else "normal",
        )
        progress.set(self.update_state.progress)
        if self.update_state.progress_visible:
            progress.grid()
        else:
            progress.grid_remove()
        panel.grid()

    def _start_update_install(self) -> None:
        update = self.update_state.info
        if update is None or self.update_state.busy:
            return
        self.update_state.start_download()
        self._sync_update_panel()
        worker = threading.Thread(target=self._download_update_in_worker, args=(update,), daemon=True)
        worker.start()

    def _download_update_in_worker(self, update: UpdateInfo) -> None:
        def progress(received: int, total: int) -> None:
            self.update_queue.put(("download_progress", (received, total)))

        try:
            zip_path = download_update(update, progress)
            self.update_queue.put(("downloaded", zip_path))
        except Exception as exc:
            self.update_queue.put(("update_error", str(exc)))

    def _install_downloaded_update(self, zip_path: Path) -> None:
        self.update_state.mark_installing()
        self._sync_update_panel()
        try:
            launch_update_installer(zip_path)
        except Exception as exc:
            self.update_state.mark_error(str(exc))
            self._sync_update_panel()
            messagebox.showerror("CorrectionsIQ", str(exc))
            return
        self.after(350, self.destroy)

    def _choose_folder(self) -> None:
        folder = filedialog.askdirectory(initialdir=self.folder_var.get().strip() or str(self.project_dir))
        if folder:
            self.folder_var.set(folder)

    def _select_settlement(self) -> None:
        if self.settlement_var.get():
            self.epkt_var.set(False)

    def _select_epkt(self) -> None:
        if self.epkt_var.get():
            self.settlement_var.set(False)

    def _selected_category(self) -> str | None:
        if self.settlement_var.get():
            return "Колония - поселение"
        if self.epkt_var.get():
            return "ЕПКТ"
        return None

    def _start_calculation(self) -> None:
        if self.is_running:
            return
        try:
            folder = Path(self.folder_var.get().strip())
            institution = self.institution_var.get().strip()
            category = self._selected_category()
            if not self.folder_var.get().strip():
                raise ValueError("Выберите папку со сводками.")
            if not folder.exists():
                raise ValueError("Папка со сводками не найдена.")
            if not folder.is_dir():
                raise ValueError("Указанный путь не является папкой со сводками.")
            if not institution:
                raise ValueError("Укажите учреждение.")
            if not self.start_var.get().strip():
                raise ValueError("Укажите дату начала.")
            if not self.end_var.get().strip():
                raise ValueError("Укажите дату окончания.")
            start = parse_ru_date(self.start_var.get())
            end = parse_ru_date(self.end_var.get())
            if start > end:
                raise ValueError("Дата начала не может быть позже даты окончания.")
        except Exception as exc:
            self._show_error(str(exc))
            return

        self.is_running = True
        self.calculation_error = None
        self.calculate_button.configure(state="disabled", text="Расчет...")
        self.export_button.configure(state="disabled")
        self.status_label.configure(text="Чтение Word-файлов и сбор дневных значений...")
        self.progress.configure(mode="indeterminate")
        self.progress.start()
        self._clear_rows()

        worker = threading.Thread(
            target=self._calculate_in_worker,
            args=(folder, institution, category, start, end),
            daemon=True,
        )
        worker.start()

    def _calculate_in_worker(self, folder: Path, institution: str, category: str | None, start: date, end: date) -> None:
        try:
            records = read_records([folder], institution, start, end, category, False, False)
            segments = build_segments(records, start, end)
            total_days = (end - start).days + 1
            total_weighted = sum(segment.weighted for segment in segments)
            self.worker_queue.put(
                (
                    "success",
                    CalculationResult(
                        segments=segments,
                        total_days=total_days,
                        total_weighted=total_weighted,
                        average=total_weighted / total_days,
                        institution=institution,
                        start=start,
                        end=end,
                        category=category,
                    ),
                )
            )
        except Exception as exc:
            self.worker_queue.put(("error", str(exc)))

    def _poll_worker(self) -> None:
        try:
            kind, payload = self.worker_queue.get_nowait()
        except queue.Empty:
            self.after(120, self._poll_worker)
            return

        self.is_running = False
        self._sync_calculation_controls()
        if kind == "success" and isinstance(payload, CalculationResult):
            self._accept_calculation_result(payload)
        else:
            self._accept_calculation_error(str(payload))
        self.after(120, self._poll_worker)

    def _sync_calculation_controls(self) -> None:
        progress: Any = getattr(self, "progress", None)
        calculate_button: Any = getattr(self, "calculate_button", None)
        export_button: Any = getattr(self, "export_button", None)
        status_label: Any = getattr(self, "status_label", None)
        if not all(
            self._is_live_widget(widget)
            for widget in (progress, calculate_button, export_button, status_label)
        ):
            return

        if self.is_running:
            calculate_button.configure(state="disabled", text="Расчет...")
            export_button.configure(state="disabled")
            status_label.configure(text="Чтение Word-файлов и сбор дневных значений...")
            progress.configure(mode="indeterminate")
            progress.start()
            return

        progress.stop()
        progress.configure(mode="determinate")
        progress.set(1 if self.result is not None else 0)
        calculate_button.configure(state="normal", text="Рассчитать")
        export_button.configure(state="normal" if self.result is not None else "disabled")

    def _accept_calculation_result(self, result: CalculationResult) -> None:
        self.result = result
        self.calculation_error = None
        if self.active_module == "headcount":
            self._show_result(result)

    def _accept_calculation_error(self, message: str) -> None:
        self.result = None
        self.calculation_error = self._friendly_error_message(message)
        if self.active_module == "headcount":
            self._render_error(self.calculation_error)

    def _show_result(self, result: CalculationResult) -> None:
        self.result = result
        self.calculation_error = None
        if not all(
            self._is_live_widget(widget)
            for widget in (
                getattr(self, "avg_value", None),
                getattr(self, "days_value", None),
                getattr(self, "weighted_value", None),
                getattr(self, "status_label", None),
                getattr(self, "export_button", None),
                getattr(self, "rows_frame", None),
            )
        ):
            return
        self.avg_value.configure(text=f"{result.average:.2f}")
        self.days_value.configure(text=str(result.total_days))
        self.weighted_value.configure(text=f"{result.total_weighted:,}".replace(",", " "))
        self.status_label.configure(text="Расчет выполнен. Проверьте дневные значения в таблице.")
        self.export_button.configure(state="normal")
        self._clear_rows()
        for row_index, segment in enumerate(result.segments):
            bg = "#0B1220" if row_index % 2 == 0 else "#101A2A"
            for col_index, value in enumerate((segment.report_date.strftime("%d.%m.%Y"), str(segment.value))):
                ctk.CTkLabel(
                    self.rows_frame,
                    text=value,
                    anchor="w",
                    fg_color=bg,
                    text_color=TEXT,
                    height=42,
                ).grid(row=row_index, column=col_index, sticky="ew", padx=0, pady=0)

    def _show_error(self, message: str) -> None:
        self._accept_calculation_error(message)

    def _friendly_error_message(self, message: str) -> str:
        if "Не найдено ни одной подходящей сводки" in message:
            institution = self.institution_var.get().strip()
            return (
                f"Не найдено учреждение «{institution}» за выбранный период.\n\n"
                "Проверьте название учреждения. Например: ИК-1, ИК-10, ИЦ-1, УИЦ-1, СИЗО-1, ЛПУ-3."
            )
        return message

    def _render_error(self, message: str) -> None:
        if not all(
            self._is_live_widget(widget)
            for widget in (
                getattr(self, "avg_value", None),
                getattr(self, "days_value", None),
                getattr(self, "weighted_value", None),
                getattr(self, "status_label", None),
                getattr(self, "export_button", None),
                getattr(self, "rows_frame", None),
            )
        ):
            return
        self.export_button.configure(state="disabled")
        self.status_label.configure(text=f"Ошибка: {message}")
        self.avg_value.configure(text="—")
        self.days_value.configure(text="—")
        self.weighted_value.configure(text="—")
        self._clear_rows()
        ctk.CTkLabel(
            self.rows_frame,
            text=message,
            anchor="w",
            justify="left",
            wraplength=840,
            text_color="#B42318",
            font=ctk.CTkFont(size=14, weight="bold"),
        ).grid(row=0, column=0, columnspan=2, sticky="ew", padx=16, pady=18)

    def _show_empty_state(self, message: str = "Заполните параметры и нажмите «Рассчитать».") -> None:
        self._clear_rows()
        ctk.CTkLabel(
            self.rows_frame,
            text=message,
            anchor="center",
            text_color=MUTED,
            font=ctk.CTkFont(size=15),
        ).grid(row=0, column=0, columnspan=2, sticky="ew", padx=12, pady=34)

    def _clear_rows(self) -> None:
        for child in self.rows_frame.winfo_children():
            child.destroy()

    def _export_xlsx(self) -> None:
        if not self.result:
            return
        target = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx"), ("Все файлы", "*.*")],
            initialfile=self._default_excel_filename(self.result),
        )
        if not target:
            return
        try:
            write_xlsx(Path(target), self.result.segments, self.result.average)
        except Exception as exc:
            self.status_label.configure(text=f"Не удалось сохранить Excel: {exc}")
            messagebox.showerror("CorrectionsIQ", f"Не удалось сохранить Excel-файл.\n\n{exc}")
            return
        self.status_label.configure(text=f"Excel сохранен: {target}")

    def _default_excel_filename(self, result: CalculationResult) -> str:
        created_at = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        period = f"{result.start:%d.%m.%Y}-{result.end:%d.%m.%Y}"
        institution = result.institution
        if result.category:
            institution = f"{institution}_{result.category}"
        raw_name = f"{created_at}__{period}__{institution}.xlsx"
        return re.sub(r'[<>:"/\\\\|?*]', "_", raw_name)


def run_app() -> None:
    app = HeadcountApp()
    app.mainloop()
