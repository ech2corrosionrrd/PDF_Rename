"""Tk-інтерфейс PDF Rename Expert.

Tk-частина оркеструє чисті модулі: `theme`, `naming`, `file_builder`, `report`,
`pdf_preview`, `suffix_history`, `excel_db`. Уся «бізнес-логіка» (формування
імені, експорт, рендеринг) живе поза цим файлом і покрита pytest-ами.
"""

from __future__ import annotations

import logging
import os
import sys
import threading
from datetime import date
from pathlib import Path
from queue import Empty, Queue
from typing import Any, List, Optional, Tuple

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

try:
    from tkcalendar import DateEntry  # type: ignore
except Exception:  # pragma: no cover
    DateEntry = None  # type: ignore

try:
    from PIL import ImageTk  # type: ignore
except ImportError:  # pragma: no cover
    ImageTk = None  # type: ignore

from excel_db import ConsumerRecord, ExcelConsumerDB
from meters_form import MetersWindow
from file_builder import FilenameInputs, build_pdf_filename
from naming import (
    MAX_FILENAME_LEN,
    ensure_unique_path,
    init_naming_rules,
    sanitize_component,
)
from user_manual import (
    APP_USER_MANUAL_FILENAME,
    load_app_user_manual_text,
    resolve_app_user_manual_path,
)
from user_settings import get_str, load_settings, save_settings
from pdf_preview import (
    MAX_PREVIEW_PAGES,
    PreviewUnavailable,
    render_pdf_pages,
)
from report import build_report_frames, has_any_rows, write_report
from suffix_history import (
    MAX_SUFFIX_HISTORY_ITEMS,
    load_suffix_history,
    save_suffix_history,
)
from theme import (
    FONT_MONO,
    FONT_MONO_SM,
    PALETTE,
    listbox_options,
    preview_len_style,
    setup_theme,
    status_dot_style,
    text_widget_options,
)
from version import __version__

APP_TITLE = "PDF_Rename_Expert"
WINDOW_TITLE = f"PDF Rename Expert v{__version__}"
SOURCES = ("Населення", "Промислові")
WORK_TYPES = (
    "Монтаж АСКОЕ",
    "Монтаж лічильника",
    "Монтаж АСКОЕ + монтаж лічильника",
    "Технічна перевірка",
)
CHECK_TYPES = (
    "Планова",
    "Позапланова",
    "Виконання припису",
    "Заміна лічильника",
)
PHASES = ("1Ф", "3Ф", "3Ф-Ктт-XXX")
MATCHES_COMBO_HEIGHT = 30


class App(ttk.Frame):
    def __init__(self, master: tk.Tk, app_dir: Path) -> None:
        super().__init__(master)
        self.master = master
        self.app_dir = app_dir

        self.db = ExcelConsumerDB(app_dir)
        self.matches: List[ConsumerRecord] = []
        self._eic_search_after_id: Optional[str] = None
        self._db_load_thread: Optional[threading.Thread] = None
        self._db_load_queue: "Queue[tuple]" = Queue()
        self._suppress_search: bool = False
        self._db_load_silent: bool = False

        self._preview_queue: "Queue[Tuple[int, str, Any, Any, Any]]" = Queue()
        self._preview_generation: int = 0
        self._export_queue: "Queue[tuple]" = Queue()
        self._rename_queue: "Queue[tuple]" = Queue()
        self._busy_depth: int = 0
        self._shutdown_scheduler: bool = False
        self._last_preview_had_errors: bool = False

        self._init_state_vars()
        self._suffix_history: List[str] = load_suffix_history(app_dir)

        self._current_photos: List["ImageTk.PhotoImage"] = []
        self._preview_timer: Optional[str] = None

        self._build_ui()
        self._apply_saved_user_settings()
        self._wire_events()
        self._refresh_file_list()
        self._update_preview()

        self.after(200, lambda: self._load_db_into_memory(silent=True))
        self.after(60, self._poll_preview_queue)
        self._toggle_askoe_fields()

    # ---------------- State ----------------
    def _init_state_vars(self) -> None:
        self.var_source = tk.StringVar(value=SOURCES[0])
        self.var_eic_suffix = tk.StringVar()
        self.var_eic_full = tk.StringVar()
        self.var_name = tk.StringVar()
        self.var_address = tk.StringVar()
        self.var_meter_no = tk.StringVar()
        self.var_invest = tk.BooleanVar(value=False)
        self.var_phase = tk.StringVar(value="3Ф")
        self.var_phase_ktt = tk.StringVar(value="XXX")
        self.var_work_type = tk.StringVar(value=WORK_TYPES[0])
        self.var_work_custom = tk.StringVar()
        self.var_ip = tk.StringVar()
        self.var_modem = tk.StringVar()
        self.var_check_type = tk.StringVar(value=CHECK_TYPES[0])
        self.var_pdf_name = tk.StringVar()
        self.var_preview_len = tk.StringVar(value="0")
        self.var_status = tk.StringVar(value="Готово.")
        self.var_name_suffix = tk.StringVar()
        self.var_technical = tk.BooleanVar(value=False)
        self.var_match_pick = tk.StringVar()

    def _apply_saved_user_settings(self) -> None:
        data = load_settings()
        src = get_str(data, "last_source")
        if src in SOURCES:
            self.var_source.set(src)

    def _persist_user_settings(self, **kwargs: str) -> None:
        data = load_settings()
        for k, v in kwargs.items():
            if v is not None and str(v).strip():
                data[k] = str(v).strip()
        save_settings(data)

    def _set_busy(self, busy: bool) -> None:
        self._busy_depth += 1 if busy else -1
        if self._busy_depth < 0:
            self._busy_depth = 0
        on = self._busy_depth > 0
        try:
            for w in (self.btn_process, self.btn_export):
                w.configure(state=("disabled" if on else "normal"))
        except Exception:
            pass

    def _on_user_close(self) -> None:
        self._shutdown_scheduler = True
        try:
            self._persist_user_settings(last_source=self.var_source.get())
        except Exception:
            pass
        self.master.destroy()

    def _append_log(self, line: str) -> None:
        if self._shutdown_scheduler:
            return
        try:
            self.txt_log.insert(tk.END, line.rstrip() + "\n")
            self.txt_log.see(tk.END)
        except tk.TclError:
            pass
        except Exception:
            pass

    def _on_log_key(self, event: tk.Event) -> Optional[str]:
        if event.state & 0x4 and event.keysym.lower() in ("c", "a", "x"):
            return None
        if event.keysym in (
            "Left", "Right", "Up", "Down", "Home", "End",
            "Prior", "Next", "Shift_L", "Shift_R", "Control_L", "Control_R",
        ):
            return None
        return "break"

    # ---------------- UI ----------------
    def _build_ui(self) -> None:
        self.master.title(WINDOW_TITLE)
        self.master.minsize(1280, 800)
        self.master.configure(background=PALETTE["bg"])
        self.pack(fill="both", expand=True, padx=16, pady=12)

        self.columnconfigure(0, weight=1)
        self.rowconfigure(4, weight=1)

        self._accent_bar = tk.Frame(
            self, height=3, bg=PALETTE["accent"], highlightthickness=0,
        )
        self._accent_bar.grid(row=0, column=0, sticky="ew")
        self._accent_bar.grid_propagate(False)

        self._build_header()
        self._build_toolbar()

        ttk.Separator(self, orient="horizontal").grid(
            row=3, column=0, sticky="ew", pady=(0, 8)
        )

        body = ttk.PanedWindow(self, orient="horizontal")
        body.grid(row=4, column=0, sticky="nsew")

        left = ttk.Frame(body)
        center = ttk.Frame(body)
        right = ttk.Frame(body)
        body.add(left, weight=2)
        body.add(center, weight=5)
        body.add(right, weight=3)

        self._build_file_list(left)
        self._build_preview(center)
        self._build_form(right)
        self._build_status_bar()

    def _build_header(self) -> None:
        header = ttk.Frame(self)
        header.grid(row=1, column=0, sticky="ew", pady=(12, 8))
        header.columnconfigure(1, weight=1)

        badge = tk.Label(
            header,
            text="PDF",
            bg=PALETTE["accent"],
            fg=PALETTE["accent_fg"],
            font=("Segoe UI", 9, "bold"),
            padx=10,
            pady=4,
            bd=0,
            highlightthickness=0,
        )
        badge.grid(row=0, column=0, rowspan=2, sticky="nw", padx=(0, 14))

        ttk.Label(header, text="PDF Rename Expert", style="Header.TLabel").grid(
            row=0, column=1, sticky="w"
        )
        ttk.Label(
            header,
            text="Автоматичне перейменування сканованих PDF за базою Excel",
            style="SubHeader.TLabel",
        ).grid(row=1, column=1, sticky="w", pady=(4, 0))

    def _build_toolbar(self) -> None:
        strip = tk.Frame(self, bg=PALETTE["surface_alt"], highlightthickness=0)
        strip.grid(row=2, column=0, sticky="ew", pady=(0, 4))
        strip.columnconfigure(0, weight=1)

        toolbar = ttk.Frame(strip, style="Toolbar.TFrame", padding=(12, 10))
        toolbar.grid(row=0, column=0, sticky="ew")
        toolbar.columnconfigure(2, weight=1)

        ttk.Label(toolbar, text="Джерело:", style="ToolbarSection.TLabel").grid(
            row=0, column=0, sticky="w", padx=(0, 10)
        )
        src_frame = ttk.Frame(toolbar, style="Toolbar.TFrame")
        src_frame.grid(row=0, column=1, sticky="w")
        for i, src in enumerate(SOURCES):
            ttk.Radiobutton(
                src_frame,
                text=src,
                value=src,
                variable=self.var_source,
                style="Toolbar.TRadiobutton",
            ).grid(
                row=0,
                column=i,
                sticky="w",
                padx=(0, 0) if i == 0 else (18, 0),
            )

        ttk.Button(
            toolbar,
            text="Лічильники по підстанціям",
            command=self._open_meters,
            style="Toolbar.TButton",
        ).grid(row=0, column=2, sticky="e", padx=(0, 8))
        ttk.Button(
            toolbar,
            text="Оновити список (F5)",
            command=self._refresh_file_list,
            style="Toolbar.TButton",
        ).grid(row=0, column=3, sticky="e", padx=(0, 8))
        ttk.Button(
            toolbar,
            text="Інструкція (F1)",
            command=self._show_user_manual,
            style="Toolbar.TButton",
        ).grid(row=0, column=4, sticky="e")

    def _open_meters(self) -> None:
        win = getattr(self, "_meters_win", None)
        try:
            if win is not None and win.winfo_exists():
                win.deiconify()
                win.lift()
                win.focus_force()
                return
        except tk.TclError:
            pass
        self._meters_win = MetersWindow(self.master, self.app_dir)

    def _build_file_list(self, parent: ttk.Frame) -> None:
        lf = ttk.LabelFrame(parent, text="  Скан-файли (PDF)  ", style="Card.TLabelframe")
        lf.pack(fill="both", expand=True, padx=(0, 4))
        lf.columnconfigure(0, weight=1)
        lf.rowconfigure(2, weight=1)

        self.lbl_folder = ttk.Label(lf, text="", style="CardMuted.TLabel")
        self.lbl_folder.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 8))

        tools_left = ttk.Frame(lf, style="Card.TFrame")
        tools_left.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(0, 8))
        ttk.Button(
            tools_left,
            text="Відкрити",
            command=self._open_selected_pdf,
            style="Secondary.TButton",
        ).grid(row=0, column=0, padx=(0, 8))
        ttk.Button(
            tools_left,
            text="Відкрити папку",
            command=self._open_source_folder,
            style="Secondary.TButton",
        ).grid(row=0, column=1)

        list_wrap = ttk.Frame(lf, style="Card.TFrame")
        list_wrap.grid(row=2, column=0, columnspan=2, sticky="nsew")
        list_wrap.columnconfigure(0, weight=1)
        list_wrap.rowconfigure(0, weight=1)

        self.list_files = self._make_listbox(list_wrap, height=18)
        self.list_files.grid(row=0, column=0, sticky="nsew")
        sb = ttk.Scrollbar(list_wrap, orient="vertical", command=self.list_files.yview)
        sb.grid(row=0, column=1, sticky="ns")
        self.list_files.configure(yscrollcommand=sb.set)

        log_lf = ttk.LabelFrame(lf, text="  Журнал  ", style="Card.TLabelframe")
        log_lf.grid(row=3, column=0, columnspan=2, sticky="ew", pady=(8, 0))
        log_lf.columnconfigure(0, weight=1)
        self.txt_log = tk.Text(
            log_lf,
            height=5,
            wrap="word",
            font=FONT_MONO_SM,
            **text_widget_options(background=PALETTE["surface_alt"]),
        )
        self.txt_log.grid(row=0, column=0, sticky="ew")
        log_sb = ttk.Scrollbar(log_lf, orient="vertical", command=self.txt_log.yview)
        log_sb.grid(row=0, column=1, sticky="ns")
        self.txt_log.configure(yscrollcommand=log_sb.set)
        self.txt_log.bind("<Key>", self._on_log_key)

    def _build_preview(self, parent: ttk.Frame) -> None:
        cf = ttk.LabelFrame(parent, text="  Перегляд вмісту   ", style="Card.TLabelframe")
        cf.pack(fill="both", expand=True, padx=4)
        cf.columnconfigure(0, weight=1)
        cf.rowconfigure(0, weight=1)

        self.preview_canvas = tk.Canvas(
            cf,
            background=PALETTE["surface_alt"],
            highlightthickness=1,
            highlightbackground=PALETTE["border"],
            highlightcolor=PALETTE["accent"],
            borderwidth=0,
        )
        self.preview_canvas.grid(row=0, column=0, sticky="nsew")

        self.vsb_preview = ttk.Scrollbar(
            cf, orient="vertical", command=self.preview_canvas.yview
        )
        self.vsb_preview.grid(row=0, column=1, sticky="ns")
        self.hsb_preview = ttk.Scrollbar(
            cf, orient="horizontal", command=self.preview_canvas.xview
        )
        self.hsb_preview.grid(row=1, column=0, sticky="ew")

        self.preview_canvas.configure(
            yscrollcommand=self.vsb_preview.set,
            xscrollcommand=self.hsb_preview.set,
        )

        self.lbl_preview_msg = ttk.Label(
            cf, text="Оберіть файл для перегляду", style="CardMuted.TLabel"
        )
        self.lbl_preview_msg.place(relx=0.5, rely=0.5, anchor="center")

    def _build_form(self, parent: ttk.Frame) -> None:
        rf = ttk.Frame(parent)
        rf.pack(fill="both", expand=True, padx=(4, 0))
        rf.columnconfigure(0, weight=1)
        rf.rowconfigure(0, weight=1)

        scroll_host = ttk.Frame(rf)
        scroll_host.grid(row=0, column=0, sticky="nsew")
        scroll_host.columnconfigure(0, weight=1)
        scroll_host.rowconfigure(0, weight=1)

        self._form_canvas = tk.Canvas(
            scroll_host,
            highlightthickness=0,
            borderwidth=0,
            background=PALETTE["bg"],
        )
        form_vsb = ttk.Scrollbar(
            scroll_host, orient="vertical", command=self._form_canvas.yview
        )
        self._form_canvas.grid(row=0, column=0, sticky="nsew")
        form_vsb.grid(row=0, column=1, sticky="ns")
        self._form_canvas.configure(yscrollcommand=form_vsb.set)

        self._form_inner = ttk.Frame(self._form_canvas)
        self._form_canvas_window = self._form_canvas.create_window(
            (0, 0), window=self._form_inner, anchor="nw"
        )
        self._form_inner.columnconfigure(0, weight=1)
        self._form_inner.columnconfigure(1, weight=1)

        def _on_form_area_configure(_event: Optional[tk.Event] = None) -> None:
            self._sync_form_scroll()

        self._form_inner.bind("<Configure>", _on_form_area_configure)
        self._form_canvas.bind("<Configure>", _on_form_area_configure)

        self._build_search_card(self._form_inner)
        self._build_doc_card(self._form_inner)
        self._build_work_check_cards(self._form_inner)
        self._build_askoe_card(self._form_inner)
        self._build_suffix_card(self._form_inner)
        self._build_preview_card(self._form_inner)
        self._build_actions(rf)

    def _build_search_card(self, parent: ttk.Frame) -> None:
        grp = ttk.LabelFrame(parent, text="  Пошук споживача  ", style="Card.TLabelframe")
        grp.grid(row=0, column=0, columnspan=2, sticky="ew")
        grp.columnconfigure(1, weight=1)

        ttk.Label(
            grp, text="EIC (останні 5+ символів):", style="Card.TLabel"
        ).grid(row=0, column=0, sticky="w", pady=(0, 6), padx=(0, 10))
        self.ent_eic = ttk.Entry(grp, textvariable=self.var_eic_suffix)
        self.ent_eic.grid(row=0, column=1, sticky="ew", pady=(0, 6))
        self.btn_load_db = ttk.Button(
            grp,
            text="Завантажити базу",
            command=lambda: self._load_db_into_memory(silent=False),
            style="Secondary.TButton",
        )
        self.btn_load_db.grid(row=0, column=2, sticky="e", padx=(10, 0), pady=(0, 6))

        ttk.Label(
            grp, text="EIC (повний, 16 символів):", style="Card.TLabel"
        ).grid(row=1, column=0, sticky="w", pady=(0, 6), padx=(0, 10))
        self.ent_eic_full = ttk.Entry(grp, textvariable=self.var_eic_full)
        self.ent_eic_full.grid(row=1, column=1, sticky="ew", pady=(0, 6))
        self.pb_db = ttk.Progressbar(grp, mode="indeterminate", length=160)
        self.pb_db.grid(row=1, column=2, sticky="e", padx=(10, 0), pady=(0, 6))
        self.pb_db.grid_remove()

        ttk.Label(grp, text="Збіги:", style="Card.TLabel").grid(
            row=2, column=0, sticky="nw", padx=(0, 10), pady=(2, 0)
        )
        self.matches_wrap = ttk.Frame(grp)
        self.matches_wrap.grid(row=2, column=1, columnspan=2, sticky="ew")
        self.matches_wrap.columnconfigure(0, weight=1)

        self.list_matches = self._make_listbox(self.matches_wrap, height=2)
        self.list_matches.grid(row=0, column=0, sticky="ew")
        self.matches_sb = ttk.Scrollbar(
            self.matches_wrap,
            orient="vertical",
            command=self.list_matches.yview,
        )
        self.matches_sb.grid(row=0, column=1, sticky="ns")
        self.list_matches.configure(yscrollcommand=self.matches_sb.set)

        self.cmb_matches = ttk.Combobox(
            self.matches_wrap,
            textvariable=self.var_match_pick,
            state="readonly",
            height=MATCHES_COMBO_HEIGHT,
        )
        self.cmb_matches.grid(row=0, column=0, sticky="ew")
        self.cmb_matches.grid_remove()
        self.cmb_matches.bind(
            "<Button-1>", lambda _e: self._configure_matches_combobox_height(),
            add="+",
        )

        self.chk_technical = ttk.Checkbutton(
            grp,
            text="Технічний облік",
            variable=self.var_technical,
            style="Card.TCheckbutton",
            command=self._on_technical_toggled,
        )
        self.chk_technical.grid(
            row=3, column=0, columnspan=3, sticky="w", pady=(8, 0)
        )

    def _build_doc_card(self, parent: ttk.Frame) -> None:
        grp = ttk.LabelFrame(parent, text="  Дані документа  ", style="Card.TLabelframe")
        grp.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(10, 0))
        grp.columnconfigure(1, weight=1)

        self.ent_name = self._labeled_entry(grp, "Назва (T-AD):", self.var_name, 0)
        self.ent_addr = self._labeled_entry(grp, "Адреса (AN):", self.var_address, 1)

        ttk.Label(grp, text="Дата:", style="Card.TLabel").grid(
            row=2, column=0, sticky="w", pady=(0, 6), padx=(0, 10)
        )
        date_frame = ttk.Frame(grp, style="Card.TFrame")
        date_frame.grid(row=2, column=1, sticky="ew", pady=(0, 6))
        date_frame.columnconfigure(0, weight=0)
        self._build_date_widget(date_frame)
        ttk.Checkbutton(
            date_frame, text="Інвест", variable=self.var_invest,
            style="Card.TCheckbutton",
        ).grid(row=0, column=1, sticky="w", padx=(18, 0))

        self.ent_meter = self._labeled_entry(grp, "№ Лічильника:", self.var_meter_no, 3)

        ttk.Label(grp, text="Фаза:", style="Card.TLabel").grid(
            row=4, column=0, sticky="w", padx=(0, 10)
        )
        phase_frame = ttk.Frame(grp, style="Card.TFrame")
        phase_frame.grid(row=4, column=1, sticky="ew")
        self.cmb_phase = ttk.Combobox(
            phase_frame, textvariable=self.var_phase, values=list(PHASES), width=12,
        )
        self.cmb_phase.grid(row=0, column=0, sticky="w")
        self.ent_phase_ktt = ttk.Entry(
            phase_frame, textvariable=self.var_phase_ktt, width=10
        )
        self.ent_phase_ktt.grid(row=0, column=1, sticky="w", padx=(10, 0))
        ttk.Label(phase_frame, text="(для 3Ф-Ктт)", style="CardMuted.TLabel").grid(
            row=0, column=2, sticky="w", padx=(10, 0)
        )

    def _build_date_widget(self, parent: ttk.Frame) -> None:
        if DateEntry is not None:
            self.date_entry = DateEntry(
                parent,
                date_pattern="yyyy-mm-dd",
                width=12,
                background=PALETTE["accent"],
                foreground=PALETTE["accent_fg"],
                bordercolor=PALETTE["border"],
                headersbackground=PALETTE["accent"],
                headersforeground=PALETTE["accent_fg"],
                normalbackground=PALETTE["surface"],
                normalforeground=PALETTE["text"],
                weekendbackground=PALETTE["surface"],
                weekendforeground=PALETTE["muted"],
                othermonthbackground=PALETTE["surface_alt"],
                othermonthforeground=PALETTE["muted"],
                selectbackground=PALETTE["accent"],
                selectforeground=PALETTE["accent_fg"],
            )
            self.date_entry.set_date(date.today())
        else:
            self.date_entry = ttk.Entry(parent)
            self.date_entry.insert(0, date.today().isoformat())
        self.date_entry.grid(row=0, column=0, sticky="w")

    def _build_work_check_cards(self, parent: ttk.Frame) -> None:
        wc_row = ttk.Frame(parent)
        wc_row.grid(row=2, column=0, columnspan=2, sticky="ew", pady=(10, 0))
        wc_row.columnconfigure(0, weight=1)
        wc_row.columnconfigure(1, weight=1)

        grp_work = ttk.LabelFrame(wc_row, text="  Вид роботи  ", style="Card.TLabelframe")
        grp_work.grid(row=0, column=0, sticky="nsew", padx=(0, 6))
        grp_work.columnconfigure(0, weight=1)
        grp_work.columnconfigure(1, weight=1)
        self._fill_radio_grid(grp_work, WORK_TYPES, self.var_work_type)

        row_other = (len(WORK_TYPES) + 1) // 2
        other_row = ttk.Frame(grp_work, style="Card.TFrame")
        other_row.grid(row=row_other, column=0, columnspan=2, sticky="ew", pady=(4, 0))
        other_row.columnconfigure(1, weight=1)
        ttk.Radiobutton(
            other_row, text="Інше", value="Інше", variable=self.var_work_type,
            style="Card.TRadiobutton",
        ).grid(row=0, column=0, sticky="w")
        self.ent_work_custom = ttk.Entry(other_row, textvariable=self.var_work_custom)
        self.ent_work_custom.grid(row=0, column=1, sticky="ew", padx=(10, 0))

        grp_check = ttk.LabelFrame(wc_row, text="  Вид перевірки  ", style="Card.TLabelframe")
        grp_check.grid(row=0, column=1, sticky="nsew", padx=(6, 0))
        grp_check.columnconfigure(0, weight=1)
        grp_check.columnconfigure(1, weight=1)
        self._fill_radio_grid(grp_check, CHECK_TYPES, self.var_check_type)

    def _build_askoe_card(self, parent: ttk.Frame) -> None:
        self.grp_askoe = ttk.LabelFrame(parent, text="  Дані АСКОЕ  ", style="Card.TLabelframe")
        self.grp_askoe.grid(row=3, column=0, columnspan=2, sticky="ew", pady=(10, 0))
        self.grp_askoe.columnconfigure(1, weight=1)
        self.grp_askoe.columnconfigure(3, weight=1)

        ttk.Label(self.grp_askoe, text="IP адреса:", style="Card.TLabel").grid(
            row=0, column=0, sticky="w", padx=(0, 10)
        )
        self.ent_ip = ttk.Entry(self.grp_askoe, textvariable=self.var_ip)
        self.ent_ip.grid(row=0, column=1, sticky="ew", padx=(0, 20))

        ttk.Label(self.grp_askoe, text="№ Модему:", style="Card.TLabel").grid(
            row=0, column=2, sticky="w", padx=(0, 10)
        )
        self.ent_modem = ttk.Entry(self.grp_askoe, textvariable=self.var_modem)
        self.ent_modem.grid(row=0, column=3, sticky="ew")

    def _build_suffix_card(self, parent: ttk.Frame) -> None:
        grp = ttk.LabelFrame(parent, text="  Належність лічильника  ", style="Card.TLabelframe")
        grp.grid(row=4, column=0, columnspan=2, sticky="ew", pady=(10, 0))
        grp.columnconfigure(0, weight=1)
        self.cmb_name_suffix = ttk.Combobox(
            grp, textvariable=self.var_name_suffix, values=self._suffix_history,
        )
        self.cmb_name_suffix.grid(row=0, column=0, sticky="ew")

    def _build_preview_card(self, parent: ttk.Frame) -> None:
        grp = ttk.LabelFrame(parent, text="  Ім'я PDF (прев'ю)  ", style="Card.TLabelframe")
        grp.grid(row=5, column=0, columnspan=2, sticky="ew", pady=(10, 0))
        grp.columnconfigure(0, weight=1)

        header = ttk.Frame(grp, style="Card.TFrame")
        header.grid(row=0, column=0, sticky="ew", pady=(0, 4))
        header.columnconfigure(1, weight=1)
        ttk.Label(header, text="Довжина ім'я:", style="CardMuted.TLabel").grid(
            row=0, column=0
        )
        self.lbl_preview_len = ttk.Label(
            header,
            textvariable=self.var_preview_len,
            style="PreviewLenOk.TLabel",
        )
        self.lbl_preview_len.grid(row=0, column=1, sticky="w", padx=(6, 0))
        ttk.Label(
            header, text=f"(макс. {MAX_FILENAME_LEN})", style="CardMuted.TLabel"
        ).grid(row=0, column=2, sticky="e")

        body = ttk.Frame(grp, style="Card.TFrame")
        body.grid(row=1, column=0, sticky="ew")
        body.columnconfigure(0, weight=1)

        self.txt_preview = tk.Text(
            body,
            height=2,
            wrap="none",
            font=FONT_MONO,
            **text_widget_options(background=PALETTE["surface_alt"]),
        )
        self.txt_preview.grid(row=0, column=0, sticky="ew")
        xsb = ttk.Scrollbar(body, orient="horizontal", command=self.txt_preview.xview)
        xsb.grid(row=1, column=0, sticky="ew", pady=(4, 0))
        self.txt_preview.configure(xscrollcommand=xsb.set)

        ttk.Button(
            body,
            text="Копіювати",
            command=self._copy_preview,
            style="Secondary.TButton",
        ).grid(row=0, column=1, sticky="ns", padx=(10, 0))

    def _build_actions(self, parent: ttk.Frame) -> None:
        actions = ttk.Frame(parent)
        actions.grid(row=1, column=0, sticky="ew", pady=(10, 0))
        actions.columnconfigure(1, weight=1)

        self.btn_export = ttk.Button(
            actions,
            text="Експорт в Excel",
            command=self._export_monthly_excel,
            style="Secondary.TButton",
        )
        self.btn_export.grid(row=0, column=0, sticky="w")
        self.btn_process = ttk.Button(
            actions,
            text="Обробити   (Ctrl+Enter)",
            command=self._process_selected,
            style="Accent.TButton",
        )
        self.btn_process.grid(row=0, column=2, sticky="e")

    def _build_status_bar(self) -> None:
        ttk.Separator(self, orient="horizontal").grid(
            row=5, column=0, sticky="ew", pady=(8, 0)
        )
        status = ttk.Frame(self, style="Status.TFrame")
        status.grid(row=6, column=0, sticky="ew")
        status.columnconfigure(1, weight=1)
        self.lbl_status_dot = ttk.Label(
            status, text="\u25CF", style=status_dot_style("info")
        )
        self.lbl_status_dot.grid(row=0, column=0, sticky="w")
        ttk.Label(
            status, textvariable=self.var_status, style="Status.TLabel"
        ).grid(row=0, column=1, sticky="w")

    # ---------------- UI helpers ----------------
    def _make_listbox(self, parent: tk.Widget, *, height: int) -> tk.Listbox:
        return tk.Listbox(parent, height=height, **listbox_options())

    def _labeled_entry(
        self, parent: ttk.LabelFrame, label: str, var: tk.StringVar, row: int
    ) -> ttk.Entry:
        ttk.Label(parent, text=label, style="Card.TLabel").grid(
            row=row, column=0, sticky="w", pady=(0, 6), padx=(0, 10)
        )
        ent = ttk.Entry(parent, textvariable=var)
        ent.grid(row=row, column=1, sticky="ew", pady=(0, 6))
        return ent

    def _fill_radio_grid(
        self, parent: ttk.LabelFrame, items: tuple, var: tk.StringVar
    ) -> None:
        for i, label in enumerate(items):
            rr, cc = i // 2, i % 2
            ttk.Radiobutton(
                parent, text=label, value=label, variable=var,
                style="Card.TRadiobutton",
            ).grid(
                row=rr, column=cc, sticky="w",
                padx=(0, 10) if cc == 0 else (10, 0),
                pady=(0 if rr == 0 else 4, 0),
            )

    # ---------------- Events ----------------
    def _wire_events(self) -> None:
        self.var_source.trace_add("write", lambda *_: self._refresh_file_list())

        self.ent_eic.bind("<KeyRelease>", lambda _e: self._schedule_consumer_search())
        self.var_eic_suffix.trace_add(
            "write", lambda *_: self._schedule_consumer_search()
        )
        self.list_matches.bind(
            "<<ListboxSelect>>", lambda _e: self._apply_selected_match()
        )
        self.cmb_matches.bind(
            "<<ComboboxSelected>>", lambda _e: self._apply_selected_match()
        )
        self.var_work_type.trace_add(
            "write", lambda *_: self._toggle_askoe_fields()
        )

        for v in (
            self.var_eic_full, self.var_name, self.var_address, self.var_meter_no,
            self.var_invest, self.var_phase, self.var_phase_ktt,
            self.var_work_type, self.var_work_custom,
            self.var_ip, self.var_modem,
            self.var_check_type, self.var_name_suffix,
        ):
            v.trace_add("write", lambda *_: self._update_preview())

        self.list_files.bind("<<ListboxSelect>>", lambda _e: self._on_file_selected())
        self.list_files.bind("<Double-Button-1>", lambda _e: self._open_selected_pdf())

        self.preview_canvas.bind("<MouseWheel>", self._on_mousewheel)
        self.preview_canvas.bind("<Button-4>", self._on_mousewheel)
        self.preview_canvas.bind("<Button-5>", self._on_mousewheel)
        self._bind_form_mousewheel_recursive(self._form_canvas)
        self._bind_form_mousewheel_recursive(self._form_inner)

        if DateEntry is not None:
            self.date_entry.bind(
                "<<DateEntrySelected>>", lambda _e: self._update_preview()
            )
        else:
            self.date_entry.bind("<KeyRelease>", lambda _e: self._update_preview())

        self.master.bind("<F5>", lambda _e: self._refresh_file_list())
        self.master.bind("<F1>", lambda _e: self._show_user_manual())
        self.master.bind("<Control-Return>", lambda _e: self._process_selected())
        self.ent_eic.bind("<Return>", lambda _e: self._accept_first_match())

        self.master.protocol("WM_DELETE_WINDOW", self._on_user_close)

        self._wire_editing_shortcuts()

    def _wire_editing_shortcuts(self) -> None:
        """Ctrl+A/C/V/X у полях вводу; Ctrl+C — у списках і журналі."""

        def select_all_entry(event: tk.Event) -> str:
            try:
                event.widget.select_range(0, tk.END)
                event.widget.icursor(tk.END)
            except Exception:
                pass
            return "break"

        def entry_copy(event: tk.Event) -> str:
            w = event.widget
            try:
                if w.selection_present():
                    self.clipboard_clear()
                    self.clipboard_append(w.selection_get())
            except tk.TclError:
                pass
            return "break"

        def entry_cut(event: tk.Event) -> str:
            w = event.widget
            try:
                if w.selection_present():
                    self.clipboard_clear()
                    self.clipboard_append(w.selection_get())
                    w.delete(w.index("sel.first"), w.index("sel.last"))
            except tk.TclError:
                pass
            return "break"

        def entry_paste(event: tk.Event) -> str:
            w = event.widget
            try:
                text = self.clipboard_get()
                if w.selection_present():
                    w.delete(w.index("sel.first"), w.index("sel.last"))
                w.insert(w.index("insert"), text)
            except tk.TclError:
                pass
            return "break"

        def text_clipboard(event: tk.Event, action: str) -> str:
            try:
                event.widget.event_generate(f"<<{action}>>")
            except tk.TclError:
                pass
            return "break"

        def select_all_text(event: tk.Event) -> str:
            try:
                event.widget.tag_remove("sel", "1.0", tk.END)
                event.widget.tag_add("sel", "1.0", "end-1c")
            except Exception:
                pass
            return "break"

        def select_all_text_log(_event: tk.Event) -> str:
            try:
                self.txt_log.tag_remove("sel", "1.0", tk.END)
                self.txt_log.tag_add("sel", "1.0", "end-1c")
            except Exception:
                pass
            return "break"

        def listbox_copy(event: tk.Event) -> str:
            try:
                sel = event.widget.curselection()
                if sel:
                    self.clipboard_clear()
                    self.clipboard_append(event.widget.get(sel[0]))
            except Exception:
                pass
            return "break"

        entry_like = (
            self.ent_eic, self.ent_eic_full, self.ent_name, self.ent_addr,
            self.ent_meter, self.cmb_phase, self.ent_phase_ktt,
            self.ent_work_custom, self.ent_ip, self.ent_modem,
            self.cmb_name_suffix, self.date_entry,
        )
        for w in entry_like:
            w.bind("<Control-a>", select_all_entry)
            w.bind("<Control-A>", select_all_entry)
            w.bind("<Control-c>", entry_copy)
            w.bind("<Control-C>", entry_copy)
            w.bind("<Control-v>", entry_paste)
            w.bind("<Control-V>", entry_paste)
            w.bind("<Control-x>", entry_cut)
            w.bind("<Control-X>", entry_cut)

        for seq, handler in (
            ("<Control-a>", select_all_text),
            ("<Control-A>", select_all_text),
            ("<Control-c>", lambda e: text_clipboard(e, "Copy")),
            ("<Control-C>", lambda e: text_clipboard(e, "Copy")),
            ("<Control-v>", lambda e: text_clipboard(e, "Paste")),
            ("<Control-V>", lambda e: text_clipboard(e, "Paste")),
            ("<Control-x>", lambda e: text_clipboard(e, "Cut")),
            ("<Control-X>", lambda e: text_clipboard(e, "Cut")),
        ):
            self.txt_preview.bind(seq, handler)

        def combobox_copy(event: tk.Event) -> None:
            try:
                val = str(event.widget.get()).strip()
                if val:
                    self.clipboard_clear()
                    self.clipboard_append(val)
            except Exception:
                pass

        for w in (self.list_files, self.list_matches):
            w.bind("<Control-c>", listbox_copy)
            w.bind("<Control-C>", listbox_copy)
        self.cmb_matches.bind("<Control-c>", combobox_copy)
        self.cmb_matches.bind("<Control-C>", combobox_copy)

        self.txt_log.bind("<Control-a>", select_all_text_log)
        self.txt_log.bind("<Control-A>", select_all_text_log)
        self.txt_log.bind("<Control-c>", lambda _e: self._log_copy_selection())
        self.txt_log.bind("<Control-C>", lambda _e: self._log_copy_selection())

    def _log_copy_selection(self) -> str:
        try:
            if self.txt_log.tag_ranges("sel"):
                t = self.txt_log.get("sel.first", "sel.last")
            else:
                t = self.txt_log.get("1.0", "end-1c")
            if t:
                self.clipboard_clear()
                self.clipboard_append(t)
        except Exception:
            pass
        return "break"

    # ---------------- Search ----------------
    def _schedule_consumer_search(self) -> None:
        if self._suppress_search or self.var_technical.get():
            return
        if self._eic_search_after_id is not None:
            try:
                self.after_cancel(self._eic_search_after_id)
            except Exception:
                pass
            self._eic_search_after_id = None
        self._eic_search_after_id = self.after(250, self._run_consumer_search)

    def _on_technical_toggled(self) -> None:
        technical = self.var_technical.get()
        state = "disabled" if technical else "normal"
        for w in (self.ent_eic, self.ent_eic_full):
            try:
                w.configure(state=state)
            except tk.TclError:
                pass
        if technical:
            self.var_eic_suffix.set("")
            self.var_eic_full.set("")
        self._set_matches_widget(technical)
        self._run_consumer_search()

    def _set_matches_widget(self, technical: bool) -> None:
        if technical:
            self.list_matches.grid_remove()
            self.matches_sb.grid_remove()
            self.cmb_matches.grid(row=0, column=0, sticky="ew")
        else:
            self.cmb_matches.grid_remove()
            self.list_matches.grid(row=0, column=0, sticky="ew")
            self.matches_sb.grid(row=0, column=1, sticky="ns")

    def _configure_matches_combobox_height(self) -> None:
        """Розгорнутий список combobox — біля 30 видимих рядків (Windows/ttk)."""
        try:
            popdown = self.cmb_matches.tk.call(
                "ttk::combobox::PopdownWindow", self.cmb_matches
            )
            lb = f"{popdown}.f.l"
            self.cmb_matches.tk.call(
                lb, "configure", "-height", MATCHES_COMBO_HEIGHT
            )
        except tk.TclError:
            pass

    def _selected_match_index(self) -> Optional[int]:
        if self.var_technical.get():
            pick = self.var_match_pick.get().strip()
            displays = [m.display() for m in self.matches]
            if pick:
                try:
                    return displays.index(pick)
                except ValueError:
                    return None
            if len(self.matches) == 1:
                return 0
            return None
        sel = self.list_matches.curselection()
        if sel:
            i = int(sel[0])
            if 0 <= i < len(self.matches):
                return i
        if len(self.matches) == 1:
            return 0
        return None

    def _run_consumer_search(self) -> None:
        self._eic_search_after_id = None

        try:
            if self.var_technical.get():
                self.matches = self.db.search_technical()
            else:
                suffix = self.var_eic_suffix.get().strip()
                if len(suffix) < 5:
                    self.matches = []
                    self._clear_matches_list()
                    self._set_status(
                        "Введіть мінімум 5 символів EIC для пошуку.", "info"
                    )
                    self._update_preview()
                    return
                self.matches = self.db.search_by_eic_suffix(suffix)
        except FileNotFoundError as e:
            self.matches = []
            self._set_status(str(e), "warning")
            self._clear_matches_list()
            return
        except PermissionError:
            self.matches = []
            self._set_status("Файл Excel зайнятий іншим процесом.", "error")
            self._clear_matches_list()
            return
        except Exception as e:
            self.matches = []
            self._set_status(f"Помилка читання Excel: {e}", "error")
            self._clear_matches_list()
            return

        self._fill_matches_list()
        level = "success" if self.matches else "warning"
        if self.var_technical.get():
            msg = f"Технічний облік: знайдено {len(self.matches)}"
        else:
            msg = f"Знайдено записів: {len(self.matches)}"
        self._set_status(msg, level)

        if len(self.matches) == 1:
            if self.var_technical.get():
                self.var_match_pick.set(self.matches[0].display())
            else:
                self.list_matches.selection_clear(0, tk.END)
                self.list_matches.selection_set(0)
                self.list_matches.activate(0)
            self._apply_selected_match()
        else:
            if self.var_technical.get():
                self.var_match_pick.set("")
            else:
                self.list_matches.selection_clear(0, tk.END)
            self._update_preview()

    def _clear_matches_list(self) -> None:
        if self.var_technical.get():
            self.cmb_matches["values"] = ()
            self.var_match_pick.set("")
        else:
            self.list_matches.delete(0, tk.END)
            self.list_matches.selection_clear(0, tk.END)

    def _fill_matches_list(self) -> None:
        displays = [m.display() for m in self.matches]
        if self.var_technical.get():
            self.cmb_matches["values"] = displays
            if len(displays) == 1:
                self.var_match_pick.set(displays[0])
            else:
                self.var_match_pick.set("")
        else:
            self.list_matches.delete(0, tk.END)
            for line in displays:
                self.list_matches.insert(tk.END, line)
            self.list_matches.selection_clear(0, tk.END)

    def _apply_selected_match(self) -> None:
        idx = self._selected_match_index()
        if idx is None:
            self._update_preview()
            return
        rec = self.matches[idx]

        if rec.is_technical:
            self.var_eic_full.set(rec.tech_code)
            self.var_eic_suffix.set(rec.tech_code)
        else:
            self.var_eic_full.set(rec.eic_full)
        self.var_name.set(rec.name)
        self.var_address.set(rec.address)
        self._update_preview()

    def _accept_first_match(self) -> None:
        if not self.matches:
            return
        if self.var_technical.get():
            self.var_match_pick.set(self.matches[0].display())
        else:
            self.list_matches.selection_clear(0, tk.END)
            self.list_matches.selection_set(0)
            self.list_matches.activate(0)
        self._apply_selected_match()

    def _toggle_askoe_fields(self) -> None:
        is_askoe = "АСКОЕ" in self.var_work_type.get()
        if is_askoe:
            try:
                self.grp_askoe.grid()
            except Exception:
                pass
        else:
            try:
                self.grp_askoe.grid_remove()
            except Exception:
                pass
            self.var_ip.set("")
            self.var_modem.set("")
        self._sync_form_scroll()
        self._update_preview()

    # ---------------- Preview ----------------
    def _on_file_selected(self) -> None:
        self._update_preview()
        pdf = self._selected_pdf_path()
        if pdf and pdf.exists():
            if self._preview_timer:
                self.after_cancel(self._preview_timer)
            try:
                self.preview_canvas.update_idletasks()
                cw = self.preview_canvas.winfo_width()
            except Exception:
                cw = 0
            target_w = max(50, cw - 40) if cw > 0 else 650
            self._preview_timer = self.after(
                150, lambda p=pdf, tw=target_w: self._start_pdf_preview(p, tw)
            )

    def _start_pdf_preview(self, pdf_path: Path, target_width: int) -> None:
        self._preview_timer = None
        if ImageTk is None:
            self.lbl_preview_msg.configure(
                text="Помилка: PyMuPDF/Pillow не встановлено."
            )
            self.lbl_preview_msg.place(relx=0.5, rely=0.5, anchor="center")
            return

        self._preview_generation += 1
        gen = self._preview_generation
        self.lbl_preview_msg.place(relx=0.5, rely=0.5, anchor="center")
        self.lbl_preview_msg.configure(text="Завантаження перегляду…")

        def worker() -> None:
            try:
                images, total_pages = render_pdf_pages(pdf_path, target_width)
                self._preview_queue.put((gen, "ok", images, total_pages, None))
            except PreviewUnavailable as e:
                self._preview_queue.put((gen, "unavail", None, None, str(e)))
            except Exception as e:
                logging.exception("PDF preview failed")
                self._preview_queue.put((gen, "err", None, None, str(e)))

        threading.Thread(target=worker, daemon=True).start()

    def _poll_preview_queue(self) -> None:
        if self._shutdown_scheduler:
            return
        try:
            if not self.winfo_exists():
                return
        except tk.TclError:
            return
        try:
            latest: Optional[Tuple[int, str, Any, Any, Any]] = None
            while True:
                try:
                    msg = self._preview_queue.get_nowait()
                except Empty:
                    break
                if msg[0] == self._preview_generation:
                    latest = msg
            if latest is not None:
                self._apply_preview_queue_msg(latest)
        except tk.TclError:
            return
        except Exception:
            logging.exception("preview queue poll")
        if self._shutdown_scheduler:
            return
        try:
            self.after(50, self._poll_preview_queue)
        except tk.TclError:
            pass

    def _apply_preview_queue_msg(self, msg: Tuple[int, str, Any, Any, Any]) -> None:
        gen, kind, images, total_pages, detail = msg
        if gen != self._preview_generation:
            return
        if ImageTk is None:
            return
        if kind == "ok":
            if total_pages is None:
                return
            tp = int(total_pages)
            if tp == 0:
                self._clear_pdf_preview("Порожній PDF")
                return
            self._paint_pdf_preview_images(images or [], tp)
        elif kind == "unavail":
            self.lbl_preview_msg.configure(text=f"Помилка: {detail}")
            self.lbl_preview_msg.place(relx=0.5, rely=0.5, anchor="center")
        else:
            self._clear_pdf_preview(f"Помилка завантаження:\n{detail}")

    def _paint_pdf_preview_images(self, images: List[Any], total_pages: int) -> None:
        self.preview_canvas.delete("all")
        self._current_photos = []
        self.preview_canvas.update_idletasks()
        cw = max(1, self.preview_canvas.winfo_width())
        gap = 15
        current_y = 0
        for img in images:
            photo = ImageTk.PhotoImage(img)
            self._current_photos.append(photo)
            self.preview_canvas.create_image(
                cw // 2, current_y, image=photo, anchor="n"
            )
            current_y += img.height + gap

        self.preview_canvas.config(scrollregion=(0, 0, cw, current_y))
        self.preview_canvas.yview_moveto(0)

        if total_pages > MAX_PREVIEW_PAGES:
            self.lbl_preview_msg.configure(
                text=f"Відображено перші {MAX_PREVIEW_PAGES} сторінок з {total_pages}"
            )
            self.lbl_preview_msg.place(relx=0.5, rely=0.02, anchor="n")
        else:
            self.lbl_preview_msg.place_forget()

    def _clear_pdf_preview(self, msg: str = "Оберіть файл для перегляду") -> None:
        self.preview_canvas.delete("all")
        self._current_photos = []
        self.preview_canvas.config(scrollregion=(0, 0, 0, 0))
        self.lbl_preview_msg.configure(text=msg)
        self.lbl_preview_msg.place(relx=0.5, rely=0.5, anchor="center")

    def _on_mousewheel(self, event: tk.Event) -> None:
        if event.delta:
            self.preview_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        elif event.num == 4:
            self.preview_canvas.yview_scroll(-1, "units")
        elif event.num == 5:
            self.preview_canvas.yview_scroll(1, "units")

    def _on_form_mousewheel(self, event: tk.Event) -> Optional[str]:
        if event.delta:
            self._form_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        elif event.num == 4:
            self._form_canvas.yview_scroll(-1, "units")
        elif event.num == 5:
            self._form_canvas.yview_scroll(1, "units")
        return "break"

    def _bind_form_mousewheel_recursive(self, widget: tk.Widget) -> None:
        for seq in ("<MouseWheel>", "<Button-4>", "<Button-5>"):
            widget.bind(seq, self._on_form_mousewheel)
        for child in widget.winfo_children():
            self._bind_form_mousewheel_recursive(child)

    def _sync_form_scroll(self) -> None:
        canvas = self._form_canvas
        canvas.update_idletasks()
        bbox = canvas.bbox("all")
        if bbox:
            canvas.configure(scrollregion=bbox)
        width = canvas.winfo_width()
        if width > 1:
            canvas.itemconfigure(self._form_canvas_window, width=width)

    def _show_user_manual(self) -> None:
        if getattr(self, "_manual_window", None) is not None:
            try:
                if self._manual_window.winfo_exists():
                    self._manual_window.lift()
                    self._manual_window.focus_force()
                    return
            except tk.TclError:
                pass

        text, path = load_app_user_manual_text(self.app_dir)

        win = tk.Toplevel(self.master)
        win.title("Довідка — PDF Rename Expert")
        win.minsize(640, 480)
        win.geometry("920x680")
        win.configure(background=PALETTE["bg"])
        win.transient(self.master)
        self._manual_window = win

        top = ttk.Frame(win, padding=(10, 10, 10, 6))
        top.pack(fill="x")
        ttk.Label(
            top,
            text="Довідка для оператора. F1 — відкрити знову. Прокручуйте текст нижче.",
            style="Muted.TLabel",
        ).pack(anchor="w")

        body = ttk.Frame(win, padding=(10, 0, 10, 10))
        body.pack(fill="both", expand=True)
        body.columnconfigure(0, weight=1)
        body.rowconfigure(0, weight=1)

        txt = tk.Text(
            body,
            wrap="word",
            font=("Segoe UI", 10),
            state="normal",
            **text_widget_options(),
        )
        txt.grid(row=0, column=0, sticky="nsew")
        vsb = ttk.Scrollbar(body, orient="vertical", command=txt.yview)
        vsb.grid(row=0, column=1, sticky="ns")
        txt.configure(yscrollcommand=vsb.set)
        txt.insert("1.0", text)
        txt.configure(state="disabled")

        def _manual_key(event: tk.Event) -> Optional[str]:
            if event.state & 0x4 and event.keysym.lower() in ("c", "a"):
                return None
            return "break"

        txt.bind("<Key>", _manual_key)

        foot = ttk.Frame(win, padding=(10, 0, 10, 10))
        foot.pack(fill="x")

        def _open_external() -> None:
            p = path or resolve_app_user_manual_path(self.app_dir)
            if p is None:
                messagebox.showwarning(
                    APP_TITLE,
                    f"Файл {APP_USER_MANUAL_FILENAME} не знайдено.",
                    parent=win,
                )
                return
            try:
                os.startfile(str(p))  # type: ignore[attr-defined]
            except Exception as e:
                messagebox.showerror(
                    APP_TITLE,
                    f"Не вдалося відкрити файл:\n{p}\n\n{e}",
                    parent=win,
                )

        def _copy_to_app_dir() -> None:
            src = path or resolve_app_user_manual_path(self.app_dir)
            if src is None:
                messagebox.showwarning(APP_TITLE, "Немає файлу для копіювання.", parent=win)
                return
            dest = self.app_dir / APP_USER_MANUAL_FILENAME
            if dest.resolve() == src.resolve():
                messagebox.showinfo(
                    APP_TITLE,
                    f"Довідка вже тут:\n{dest}",
                    parent=win,
                )
                return
            try:
                dest.write_text(src.read_text(encoding="utf-8"), encoding="utf-8")
                messagebox.showinfo(
                    APP_TITLE,
                    f"Довідку збережено:\n{dest}",
                    parent=win,
                )
            except OSError as e:
                messagebox.showerror(APP_TITLE, f"Помилка збереження:\n{e}", parent=win)

        ttk.Button(foot, text="Закрити", command=win.destroy).pack(side="right")
        if path is not None:
            ttk.Button(
                foot, text="Відкрити у блокноті", command=_open_external
            ).pack(side="right", padx=(0, 8))
            ttk.Button(
                foot,
                text="Зберегти довідку поруч із програмою",
                command=_copy_to_app_dir,
            ).pack(side="right", padx=(0, 8))

        win.protocol("WM_DELETE_WINDOW", win.destroy)

    # ---------------- Files ----------------
    def _source_dir(self) -> Path:
        return self.app_dir / "Scan" / self.var_source.get()

    def _refresh_file_list(self) -> None:
        src = self._source_dir()
        self.lbl_folder.configure(text=str(src))
        self.list_files.delete(0, tk.END)

        if not src.exists():
            try:
                src.mkdir(parents=True, exist_ok=True)
            except Exception as e:
                messagebox.showerror(
                    APP_TITLE,
                    f"Не вдалося створити папку джерела:\n{src}\n\n{e}",
                )
                self._set_status("Помилка створення папки джерела.", "error")
                return

        pdfs = sorted(
            (p for p in src.iterdir() if p.is_file() and p.suffix.lower() == ".pdf"),
            reverse=True,
        )
        for p in pdfs:
            self.list_files.insert(tk.END, p.name)
        self._set_status(f"PDF у списку: {len(pdfs)}", "info")
        self._update_preview()
        try:
            self._persist_user_settings(last_source=self.var_source.get())
        except Exception:
            pass

    def _selected_pdf_path(self) -> Optional[Path]:
        sel = self.list_files.curselection()
        if not sel:
            return None
        name = self.list_files.get(sel[0])
        return self._source_dir() / name

    def _open_selected_pdf(self) -> None:
        pdf = self._selected_pdf_path()
        if pdf is None:
            messagebox.showwarning(APP_TITLE, "Оберіть PDF-файл у списку.")
            return
        if not pdf.exists():
            messagebox.showerror(APP_TITLE, "Файл не знайдено.")
            return
        try:
            os.startfile(str(pdf))  # type: ignore[attr-defined]
        except Exception as e:
            messagebox.showerror(APP_TITLE, f"Не вдалося відкрити PDF:\n{e}")

    def _open_source_folder(self) -> None:
        folder = self._source_dir()
        try:
            folder.mkdir(parents=True, exist_ok=True)
        except Exception:
            pass
        try:
            os.startfile(str(folder))  # type: ignore[attr-defined]
        except Exception as e:
            messagebox.showerror(APP_TITLE, f"Не вдалося відкрити папку:\n{e}")

    # ---------------- Build & process filename ----------------
    def _get_date_str(self) -> str:
        if DateEntry is not None:
            return self.date_entry.get_date().isoformat()
        return str(self.date_entry.get()).strip()

    def _collect_inputs(self) -> FilenameInputs:
        pdf = self._selected_pdf_path()
        return FilenameInputs(
            pdf_filename=pdf.name if pdf else None,
            eic_full_raw=self.var_eic_full.get(),
            eic_suffix_raw=self.var_eic_suffix.get(),
            has_matches=bool(self.matches),
            date_str=self._get_date_str(),
            meter_raw=self.var_meter_no.get(),
            name_raw=self.var_name.get(),
            address_raw=self.var_address.get(),
            invest=bool(self.var_invest.get()),
            phase_raw=self.var_phase.get(),
            phase_ktt_raw=self.var_phase_ktt.get(),
            work_type=self.var_work_type.get(),
            work_custom_raw=self.var_work_custom.get(),
            check_type=self.var_check_type.get(),
            ip_raw=self.var_ip.get(),
            modem_raw=self.var_modem.get(),
            name_suffix_raw=self.var_name_suffix.get(),
            technical_account=bool(self.var_technical.get()),
        )

    def _update_preview(self) -> None:
        fn, errs = build_pdf_filename(self._collect_inputs())
        self.var_pdf_name.set(fn)
        fn_len = len(fn)
        self.var_preview_len.set(str(fn_len))
        try:
            self.lbl_preview_len.configure(
                style=preview_len_style(fn_len, MAX_FILENAME_LEN)
            )
        except Exception:
            pass
        try:
            self.txt_preview.delete("1.0", tk.END)
            self.txt_preview.insert("1.0", fn)
        except Exception:
            pass
        had_errs = bool(errs)
        if errs:
            self._set_status(
                " / ".join(errs[:2]) + (" ..." if len(errs) > 2 else ""),
                "warning",
            )
        elif self._last_preview_had_errors:
            self._set_status("Готово.", "info")
        self._last_preview_had_errors = had_errs

    def _copy_preview(self) -> None:
        try:
            if self.txt_preview.tag_ranges("sel"):
                text = self.txt_preview.get("sel.first", "sel.last")
            else:
                text = self.var_pdf_name.get()
        except Exception:
            text = self.var_pdf_name.get()
        if not text:
            return
        try:
            self.clipboard_clear()
            self.clipboard_append(text)
            self._set_status("Скопійовано ім'я PDF в буфер обміну.", "success")
        except Exception:
            pass

    def _remember_suffix(self, raw: str) -> None:
        s = sanitize_component(raw, spaces_to_dash=True, underscores_to_dash=True)
        if not s:
            return
        rest = [x for x in self._suffix_history if x != s]
        self._suffix_history = ([s] + rest)[:MAX_SUFFIX_HISTORY_ITEMS]
        save_suffix_history(self.app_dir, self._suffix_history)
        try:
            self.cmb_name_suffix["values"] = tuple(self._suffix_history)
        except Exception:
            pass

    def _clear_for_next(self) -> None:
        # Залишаємо: дата, фаза, вид роботи (і текст «Інше»), вид перевірки, належність лічильника.
        self._suppress_search = True
        try:
            self.var_eic_suffix.set("")
            self.var_eic_full.set("")
            self.matches = []
            self._clear_matches_list()
            self.var_name.set("")
            self.var_address.set("")
            self.var_meter_no.set("")
            self.var_invest.set(False)
            self.var_ip.set("")
            self.var_modem.set("")
        finally:
            self._suppress_search = False
        self._update_preview()

    def _process_selected(self) -> None:
        pdf = self._selected_pdf_path()
        if pdf is None:
            messagebox.showwarning(APP_TITLE, "Оберіть PDF-файл у списку.")
            return

        final_name, errs = build_pdf_filename(self._collect_inputs())
        if errs:
            messagebox.showwarning(
                APP_TITLE, "Не можна обробити:\n- " + "\n- ".join(errs)
            )
            return

        pdf_src = pdf
        self._append_log(f"Перейменування: {pdf_src.name} → {final_name}")

        def worker() -> None:
            try:
                dst = ensure_unique_path(pdf_src.parent / final_name)
                os.replace(str(pdf_src), str(dst))
                self._rename_queue.put(("ok", str(dst)))
            except PermissionError:
                self._rename_queue.put(("perm", ""))
            except OSError as e:
                if getattr(e, "winerror", None) == 32:
                    self._rename_queue.put(("perm", ""))
                else:
                    logging.exception("Rename failed")
                    self._rename_queue.put(("err", str(e)))
            except Exception as e:
                logging.exception("Rename failed")
                self._rename_queue.put(("err", str(e)))

        self._set_busy(True)
        threading.Thread(target=worker, daemon=True).start()
        self.after(120, self._poll_rename_queue)

    def _poll_rename_queue(self) -> None:
        try:
            kind, payload = self._rename_queue.get_nowait()
        except Empty:
            self.after(120, self._poll_rename_queue)
            return
        self._set_busy(False)
        if kind == "ok":
            dst = Path(str(payload))
            self._set_status(f"Перейменовано: {dst.name}", "success")
            self._append_log(f"OK: {dst.name}")
            self._remember_suffix(self.var_name_suffix.get())
            self._refresh_file_list()
            self._clear_for_next()
            self._clear_pdf_preview()
        elif kind == "perm":
            messagebox.showerror(APP_TITLE, "PDF-файл заблокований іншою програмою.")
            self._append_log("Помилка: файл заблокований.")
        else:
            messagebox.showerror(APP_TITLE, f"Помилка перейменування:\n{payload}")
            self._append_log(f"Помилка: {payload}")

    # ---------------- Status & DB ----------------
    def _set_status(self, text: str, level: str = "info") -> None:
        """level: info | success | warning | error"""
        self.var_status.set(text)
        try:
            self.lbl_status_dot.configure(style=status_dot_style(level))
        except Exception:
            pass

    def _load_db_into_memory(self, *, silent: bool = False) -> None:
        if self._db_load_thread is not None and self._db_load_thread.is_alive():
            return

        self._db_load_silent = silent
        self.btn_load_db.configure(state="disabled")
        self.pb_db.grid()
        self.pb_db.start(10)
        self._set_status("Завантаження бази в пам'ять...", "info")

        def worker() -> None:
            try:
                self.db.load()
                self._db_load_queue.put(("ok", self.db.record_count))
            except FileNotFoundError as e:
                self._db_load_queue.put(("warn", str(e)))
            except PermissionError:
                self._db_load_queue.put(("err", "Файл Excel зайнятий іншим процесом."))
            except Exception as e:
                self._db_load_queue.put(("err", f"Помилка читання Excel: {e}"))

        self._db_load_thread = threading.Thread(target=worker, daemon=True)
        self._db_load_thread.start()
        self.after(100, self._poll_db_load_queue)

    def _poll_db_load_queue(self) -> None:
        try:
            kind, payload = self._db_load_queue.get_nowait()
        except Empty:
            self.after(100, self._poll_db_load_queue)
            return

        self.pb_db.stop()
        self.pb_db.grid_remove()
        self.btn_load_db.configure(state="normal")

        silent = self._db_load_silent
        if kind == "ok":
            n = int(payload)
            self._set_status(f"База завантажена в пам'ять. Записів: {n}", "success")
            if self.var_technical.get():
                self._run_consumer_search()
            if not silent:
                messagebox.showinfo(
                    APP_TITLE, f"База успішно завантажена.\nЗаписів: {n}"
                )
        elif kind == "warn":
            self._set_status(str(payload), "warning")
            messagebox.showwarning(APP_TITLE, str(payload))
        else:
            self._set_status(str(payload), "error")
            messagebox.showerror(APP_TITLE, str(payload))

    # ---------------- Export ----------------
    def _export_monthly_excel(self) -> None:
        data = load_settings()
        initial_dir = get_str(data, "last_export_dir") or str(self.app_dir)
        report_name = f"Звіт_всі_дані_{date.today().isoformat()}.xlsx"
        report_path_str = filedialog.asksaveasfilename(
            parent=self.master,
            title="Експорт звіту в Excel",
            initialdir=initial_dir,
            initialfile=report_name,
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx"), ("Усі файли", "*.*")],
        )
        if not report_path_str:
            return
        report_path = Path(report_path_str)

        if report_path.exists():
            if not messagebox.askyesno(
                APP_TITLE,
                f"Файл уже існує:\n{report_path.name}\n\nПерезаписати?",
            ):
                return

        scan_dir = self.app_dir / "Scan"
        for sub in ("Населення", "Промислові"):
            try:
                (scan_dir / sub).mkdir(parents=True, exist_ok=True)
            except Exception:
                pass

        self._append_log(f"Експорт Excel: {report_path}")
        self._set_busy(True)

        def worker() -> None:
            try:
                frames = build_report_frames(scan_dir)
                if not has_any_rows(frames):
                    self._export_queue.put(("empty", None, None, None))
                    return
                n_pop = len(frames["Населення"])
                n_ind = len(frames["Промислові"])
                write_report(report_path, frames)
                self._export_queue.put(("ok", str(report_path), n_pop, n_ind))
            except PermissionError:
                self._export_queue.put(("perm", None, None, None))
            except Exception as e:
                logging.exception("Excel export failed")
                self._export_queue.put(("err", str(e), None, None))

        threading.Thread(target=worker, daemon=True).start()
        self.after(150, self._poll_export_queue)

    def _poll_export_queue(self) -> None:
        try:
            kind = self._export_queue.get_nowait()
        except Empty:
            self.after(150, self._poll_export_queue)
            return
        self._set_busy(False)
        if not isinstance(kind, tuple) or not kind:
            return
        tag = kind[0]
        if tag == "ok":
            path_s = str(kind[1])
            n_pop = int(kind[2] or 0)
            n_ind = int(kind[3] or 0)
            rp = Path(path_s)
            try:
                self._persist_user_settings(last_export_dir=str(rp.parent))
            except Exception:
                pass
            self._set_status(f"Експорт завершено: {rp.name}", "success")
            self._append_log(f"Експорт OK: {rp.name} (Населення: {n_pop}, Промислові: {n_ind})")
            summary = (
                f"Звіт: {rp.name}\n"
                f"Населення: {n_pop}\n"
                f"Промислові: {n_ind}"
            )
            messagebox.showinfo(APP_TITLE, f"Готово.\n\n{summary}")
        elif tag == "empty":
            messagebox.showinfo(
                APP_TITLE, "Немає оброблених файлів у папці Scan для експорту."
            )
            self._append_log("Експорт: немає рядків для звіту.")
        elif tag == "perm":
            messagebox.showerror(
                APP_TITLE, "Excel-звіт зайнятий іншим процесом (закрийте файл)."
            )
            self._append_log("Помилка: Excel зайнятий.")
        else:
            err = str(kind[1])
            messagebox.showerror(APP_TITLE, f"Не вдалося записати Excel-звіт:\n{err}")
            self._append_log(f"Помилка експорту: {err}")


# ---------------- App bootstrap ----------------
def _setup_logging(app_dir: Path) -> None:
    try:
        log_file = str(app_dir / "pdf_rename_expert.log")
        handler = logging.FileHandler(log_file, encoding="utf-8")
        handler.setFormatter(
            logging.Formatter("%(asctime)s %(levelname)s %(message)s")
        )
        root_logger = logging.getLogger()
        root_logger.setLevel(logging.INFO)
        if not root_logger.handlers:
            root_logger.addHandler(handler)
    except Exception:
        pass


def _resolve_app_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def run_app() -> int:
    app_dir = _resolve_app_dir()
    try:
        os.chdir(str(app_dir))
    except Exception:
        pass

    _setup_logging(app_dir)
    logging.info("Start PDF_Rename_Expert from %s", app_dir)

    init_naming_rules(app_dir)

    root = tk.Tk()
    try:
        setup_theme(root)
    except Exception:
        logging.exception("Theme setup failed")
    App(root, app_dir=app_dir)
    root.mainloop()
    return 0
