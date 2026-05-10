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
from queue import Queue
from typing import List, Optional

import tkinter as tk
from tkinter import messagebox, ttk

try:
    from tkcalendar import DateEntry  # type: ignore
except Exception:  # pragma: no cover
    DateEntry = None  # type: ignore

try:
    from PIL import ImageTk  # type: ignore
except ImportError:  # pragma: no cover
    ImageTk = None  # type: ignore

from excel_db import ConsumerRecord, ExcelConsumerDB
from file_builder import FilenameInputs, build_pdf_filename
from naming import (
    MAX_FILENAME_LEN,
    ensure_unique_path,
    sanitize_component,
)
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
from theme import PALETTE, setup_theme

APP_TITLE = "PDF_Rename_Expert"
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

        self._init_state_vars()
        self._suffix_history: List[str] = load_suffix_history(app_dir)

        self._current_photos: List["ImageTk.PhotoImage"] = []
        self._preview_timer: Optional[str] = None

        self._build_ui()
        self._wire_events()
        self._refresh_file_list()
        self._update_preview()

        self.after(200, lambda: self._load_db_into_memory(silent=True))
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

    # ---------------- UI ----------------
    def _build_ui(self) -> None:
        self.master.title("PDF Rename Expert")
        self.master.minsize(1280, 800)
        self.master.configure(background=PALETTE["bg"])
        self.pack(fill="both", expand=True, padx=14, pady=10)

        self.columnconfigure(0, weight=1)
        self.rowconfigure(2, weight=1)

        self._build_header()
        self._build_toolbar()

        body = ttk.PanedWindow(self, orient="horizontal")
        body.grid(row=2, column=0, sticky="nsew")

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
        header.grid(row=0, column=0, sticky="ew")
        header.columnconfigure(0, weight=1)
        ttk.Label(header, text="PDF Rename Expert", style="Header.TLabel").grid(
            row=0, column=0, sticky="w"
        )
        ttk.Label(
            header,
            text="Автоматичне перейменування сканованих PDF за базою Excel",
            style="SubHeader.TLabel",
        ).grid(row=1, column=0, sticky="w", pady=(2, 0))

    def _build_toolbar(self) -> None:
        toolbar = ttk.Frame(self)
        toolbar.grid(row=1, column=0, sticky="ew", pady=(8, 8))
        toolbar.columnconfigure(2, weight=1)

        ttk.Label(toolbar, text="Джерело:", style="Section.TLabel").grid(
            row=0, column=0, sticky="w", padx=(0, 10)
        )
        src_frame = ttk.Frame(toolbar)
        src_frame.grid(row=0, column=1, sticky="w")
        for i, src in enumerate(SOURCES):
            ttk.Radiobutton(
                src_frame, text=src, value=src, variable=self.var_source
            ).grid(row=0, column=i, sticky="w", padx=(0, 0) if i == 0 else (18, 0))

        ttk.Button(
            toolbar, text="Оновити список (F5)", command=self._refresh_file_list
        ).grid(row=0, column=3, sticky="e")

    def _build_file_list(self, parent: ttk.Frame) -> None:
        lf = ttk.LabelFrame(parent, text="  Скан-файли (PDF)  ", style="Card.TLabelframe")
        lf.pack(fill="both", expand=True, padx=(0, 4))
        lf.columnconfigure(0, weight=1)
        lf.rowconfigure(2, weight=1)

        self.lbl_folder = ttk.Label(lf, text="", style="CardMuted.TLabel")
        self.lbl_folder.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 8))

        tools_left = ttk.Frame(lf, style="Card.TFrame")
        tools_left.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(0, 8))
        ttk.Button(tools_left, text="Відкрити", command=self._open_selected_pdf).grid(
            row=0, column=0, padx=(0, 8)
        )
        ttk.Button(
            tools_left, text="Відкрити папку", command=self._open_source_folder
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

    def _build_preview(self, parent: ttk.Frame) -> None:
        cf = ttk.LabelFrame(parent, text="  Перегляд вмісту   ", style="Card.TLabelframe")
        cf.pack(fill="both", expand=True, padx=4)
        cf.columnconfigure(0, weight=1)
        cf.rowconfigure(0, weight=1)

        self.preview_canvas = tk.Canvas(
            cf,
            background=PALETTE["surface_alt"],
            highlightthickness=0,
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
        rf.columnconfigure(1, weight=1)

        self._build_search_card(rf)
        self._build_doc_card(rf)
        self._build_work_check_cards(rf)
        self._build_askoe_card(rf)
        self._build_suffix_card(rf)
        self._build_preview_card(rf)
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
        matches_wrap = ttk.Frame(grp)
        matches_wrap.grid(row=2, column=1, columnspan=2, sticky="ew")
        matches_wrap.columnconfigure(0, weight=1)
        self.list_matches = self._make_listbox(matches_wrap, height=2)
        self.list_matches.grid(row=0, column=0, sticky="ew")
        sb = ttk.Scrollbar(
            matches_wrap, orient="vertical", command=self.list_matches.yview
        )
        sb.grid(row=0, column=1, sticky="ns")
        self.list_matches.configure(yscrollcommand=sb.set)

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
        header.grid(row=0, column=0, sticky="ew", padx=10, pady=(0, 4))
        header.columnconfigure(1, weight=1)
        ttk.Label(header, text="Довжина ім'я:", style="CardMuted.TLabel").grid(
            row=0, column=0
        )
        ttk.Label(
            header, textvariable=self.var_preview_len, style="Card.TLabel",
            font=("Segoe UI", 10, "bold"),
        ).grid(row=0, column=1, sticky="w", padx=(5, 0))
        ttk.Label(
            header, text=f"(макс. {MAX_FILENAME_LEN})", style="CardMuted.TLabel"
        ).grid(row=0, column=2, sticky="e")

        body = ttk.Frame(grp, style="Card.TFrame")
        body.grid(row=1, column=0, sticky="ew")
        body.columnconfigure(0, weight=1)

        self.txt_preview = tk.Text(
            body,
            height=2, wrap="none", font=("Consolas", 10),
            background=PALETTE["surface_alt"], foreground=PALETTE["text"],
            relief="flat", borderwidth=1,
            highlightthickness=1, highlightbackground=PALETTE["border"],
            highlightcolor=PALETTE["accent"],
            padx=12, pady=10,
        )
        self.txt_preview.grid(row=0, column=0, sticky="ew")
        xsb = ttk.Scrollbar(body, orient="horizontal", command=self.txt_preview.xview)
        xsb.grid(row=1, column=0, sticky="ew", pady=(4, 0))
        self.txt_preview.configure(xscrollcommand=xsb.set)

        ttk.Button(body, text="Копіювати", command=self._copy_preview).grid(
            row=0, column=1, sticky="ns", padx=(10, 0)
        )

    def _build_actions(self, parent: ttk.Frame) -> None:
        actions = ttk.Frame(parent)
        actions.grid(row=6, column=0, columnspan=2, sticky="ew", pady=(10, 0))
        actions.columnconfigure(1, weight=1)

        self.btn_export = ttk.Button(
            actions, text="Експорт в Excel", command=self._export_monthly_excel
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
        status = ttk.Frame(self, style="Status.TFrame")
        status.grid(row=3, column=0, sticky="ew", pady=(10, 0))
        status.columnconfigure(1, weight=1)
        self.lbl_status_dot = ttk.Label(status, text="\u25CF", style="StatusDot.TLabel")
        self.lbl_status_dot.grid(row=0, column=0, sticky="w")
        ttk.Label(
            status, textvariable=self.var_status, style="Status.TLabel"
        ).grid(row=0, column=1, sticky="w")

    # ---------------- UI helpers ----------------
    def _make_listbox(self, parent: tk.Widget, *, height: int) -> tk.Listbox:
        return tk.Listbox(
            parent,
            height=height,
            exportselection=False,
            activestyle="none",
            borderwidth=1,
            highlightthickness=1,
            highlightbackground=PALETTE["border"],
            highlightcolor=PALETTE["accent"],
            background=PALETTE["surface"],
            foreground=PALETTE["text"],
            selectbackground=PALETTE["accent"],
            selectforeground=PALETTE["accent_fg"],
            font=("Segoe UI", 10),
            relief="flat",
        )

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

        self.ent_eic.bind("<KeyRelease>", lambda _e: self._schedule_eic_search())
        self.var_eic_suffix.trace_add(
            "write", lambda *_: self._schedule_eic_search()
        )
        self.list_matches.bind(
            "<<ListboxSelect>>", lambda _e: self._apply_selected_match()
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

        if DateEntry is not None:
            self.date_entry.bind(
                "<<DateEntrySelected>>", lambda _e: self._update_preview()
            )
        else:
            self.date_entry.bind("<KeyRelease>", lambda _e: self._update_preview())

        self.master.bind("<F5>", lambda _e: self._refresh_file_list())
        self.master.bind("<Control-Return>", lambda _e: self._process_selected())
        self.ent_eic.bind("<Return>", lambda _e: self._accept_first_match())

        self._wire_editing_shortcuts()

    def _wire_editing_shortcuts(self) -> None:
        """Ctrl+A — виділити все у полях; Ctrl+C — копіювати рядок зі списків."""

        def select_all_entry(event: tk.Event) -> str:
            try:
                event.widget.select_range(0, tk.END)
                event.widget.icursor(tk.END)
            except Exception:
                pass
            return "break"

        def select_all_text(_event: tk.Event) -> str:
            try:
                self.txt_preview.tag_remove("sel", "1.0", tk.END)
                self.txt_preview.tag_add("sel", "1.0", "end-1c")
            except Exception:
                pass
            return "break"

        def listbox_copy(event: tk.Event) -> None:
            try:
                sel = event.widget.curselection()
                if sel:
                    self.clipboard_clear()
                    self.clipboard_append(event.widget.get(sel[0]))
            except Exception:
                pass

        entry_like = (
            self.ent_eic, self.ent_eic_full, self.ent_name, self.ent_addr,
            self.ent_meter, self.cmb_phase, self.ent_phase_ktt,
            self.ent_work_custom, self.ent_ip, self.ent_modem,
            self.cmb_name_suffix,
        )
        for w in entry_like:
            w.bind("<Control-a>", select_all_entry)
            w.bind("<Control-A>", select_all_entry)
        if DateEntry is not None:
            self.date_entry.bind("<Control-a>", select_all_entry)
            self.date_entry.bind("<Control-A>", select_all_entry)

        self.txt_preview.bind("<Control-a>", select_all_text)
        self.txt_preview.bind("<Control-A>", select_all_text)
        for w in (self.list_files, self.list_matches):
            w.bind("<Control-c>", listbox_copy)
            w.bind("<Control-C>", listbox_copy)

    # ---------------- Search ----------------
    def _schedule_eic_search(self) -> None:
        if self._suppress_search:
            return
        if self._eic_search_after_id is not None:
            try:
                self.after_cancel(self._eic_search_after_id)
            except Exception:
                pass
            self._eic_search_after_id = None
        self._eic_search_after_id = self.after(250, self._on_eic_changed)

    def _on_eic_changed(self) -> None:
        self._eic_search_after_id = None
        suffix = self.var_eic_suffix.get().strip()
        if len(suffix) < 5:
            self.matches = []
            self._clear_matches_list()
            self._set_status("Введіть мінімум 5 символів EIC для пошуку.", "info")
            self._update_preview()
            return

        try:
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
        self._set_status(f"Знайдено записів: {len(self.matches)}", level)

        if len(self.matches) == 1:
            self.list_matches.selection_clear(0, tk.END)
            self.list_matches.selection_set(0)
            self.list_matches.activate(0)
            self._apply_selected_match()
        else:
            self.list_matches.selection_clear(0, tk.END)
            self._update_preview()

    def _clear_matches_list(self) -> None:
        self.list_matches.delete(0, tk.END)
        self.list_matches.selection_clear(0, tk.END)

    def _fill_matches_list(self) -> None:
        self.list_matches.delete(0, tk.END)
        for m in self.matches:
            self.list_matches.insert(tk.END, m.display())
        self.list_matches.selection_clear(0, tk.END)

    def _apply_selected_match(self) -> None:
        sel = self.list_matches.curselection()
        rec: Optional[ConsumerRecord] = None
        if sel:
            i = int(sel[0])
            if 0 <= i < len(self.matches):
                rec = self.matches[i]
        elif len(self.matches) == 1:
            rec = self.matches[0]
        if rec is None:
            self._update_preview()
            return

        self.var_eic_full.set(rec.eic_full)
        self.var_name.set(rec.name)
        self.var_address.set(rec.address)
        self._update_preview()

    def _accept_first_match(self) -> None:
        if not self.matches:
            return
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
        self._update_preview()

    # ---------------- Preview ----------------
    def _on_file_selected(self) -> None:
        self._update_preview()
        pdf = self._selected_pdf_path()
        if pdf and pdf.exists():
            if self._preview_timer:
                self.after_cancel(self._preview_timer)
            self._preview_timer = self.after(150, lambda: self._display_pdf(pdf))

    def _display_pdf(self, pdf_path: Path) -> None:
        if ImageTk is None:
            self.lbl_preview_msg.configure(text="Помилка: PyMuPDF/Pillow не встановлено.")
            return

        try:
            self.master.update_idletasks()
            cw = self.preview_canvas.winfo_width()
            target_w = max(50, cw - 40)
            images, total_pages = render_pdf_pages(pdf_path, target_w)
        except PreviewUnavailable as e:
            self.lbl_preview_msg.configure(text=f"Помилка: {e}")
            return
        except Exception as e:
            logging.exception("PDF preview failed")
            self._clear_pdf_preview(f"Помилка завантаження:\n{e}")
            return

        if total_pages == 0:
            self._clear_pdf_preview("Порожній PDF")
            return

        self.preview_canvas.delete("all")
        self._current_photos = []
        cw = self.preview_canvas.winfo_width() or (target_w + 40)
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
        )

    def _update_preview(self) -> None:
        fn, errs = build_pdf_filename(self._collect_inputs())
        self.var_pdf_name.set(fn)
        self.var_preview_len.set(str(len(fn)))
        try:
            self.txt_preview.delete("1.0", tk.END)
            self.txt_preview.insert("1.0", fn)
        except Exception:
            pass
        if errs:
            self._set_status(
                " / ".join(errs[:2]) + (" ..." if len(errs) > 2 else ""),
                "warning",
            )

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

        dst_path = ensure_unique_path(pdf.parent / final_name)
        try:
            os.replace(str(pdf), str(dst_path))
        except PermissionError:
            messagebox.showerror(APP_TITLE, "PDF-файл заблокований іншою програмою.")
            return
        except Exception as e:
            logging.exception("Rename failed")
            messagebox.showerror(APP_TITLE, f"Помилка перейменування:\n{e}")
            return

        self._set_status(f"Перейменовано: {dst_path.name}", "success")
        self._remember_suffix(self.var_name_suffix.get())
        self._refresh_file_list()
        self._clear_for_next()
        self._clear_pdf_preview()

    # ---------------- Status & DB ----------------
    def _set_status(self, text: str, level: str = "info") -> None:
        """level: info | success | warning | error"""
        self.var_status.set(text)
        color_map = {
            "info": PALETTE["muted"],
            "success": PALETTE["success"],
            "warning": PALETTE["warning"],
            "error": PALETTE["danger"],
        }
        try:
            self.lbl_status_dot.configure(
                foreground=color_map.get(level, PALETTE["muted"])
            )
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
        except Exception:
            self.after(100, self._poll_db_load_queue)
            return

        self.pb_db.stop()
        self.pb_db.grid_remove()
        self.btn_load_db.configure(state="normal")

        silent = self._db_load_silent
        if kind == "ok":
            n = int(payload)
            self._set_status(f"База завантажена в пам'ять. Записів: {n}", "success")
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
        report_name = f"Звіт_всі_дані_{date.today().isoformat()}.xlsx"
        report_path = self.app_dir / report_name

        if report_path.exists():
            if not messagebox.askyesno(
                APP_TITLE,
                f"Звіт уже існує:\n{report_path.name}\n\nПерезаписати?",
            ):
                return

        scan_dir = self.app_dir / "Scan"
        for sub in ("Населення", "Промислові"):
            try:
                (scan_dir / sub).mkdir(parents=True, exist_ok=True)
            except Exception:
                pass

        frames = build_report_frames(scan_dir)
        if not has_any_rows(frames):
            messagebox.showinfo(
                APP_TITLE, "Немає оброблених файлів у папці Scan для експорту."
            )
            return

        try:
            write_report(report_path, frames)
        except PermissionError:
            messagebox.showerror(
                APP_TITLE, "Excel-звіт зайнятий іншим процесом (закрийте файл)."
            )
            return
        except Exception as e:
            logging.exception("Excel export failed")
            messagebox.showerror(APP_TITLE, f"Не вдалося записати Excel-звіт:\n{e}")
            return

        self._set_status(f"Експорт завершено: {report_path.name}", "success")
        summary = (
            f"Звіт: {report_path.name}\n"
            f"Населення: {len(frames['Населення'])}\n"
            f"Промислові: {len(frames['Промислові'])}"
        )
        messagebox.showinfo(APP_TITLE, f"Готово.\n\n{summary}")


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

    root = tk.Tk()
    try:
        setup_theme(root)
    except Exception:
        logging.exception("Theme setup failed")
    App(root, app_dir=app_dir)
    root.mainloop()
    return 0
