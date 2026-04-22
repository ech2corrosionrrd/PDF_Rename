import json
import logging
import os
import re
import sys
import threading
from queue import Queue
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import pandas as pd

try:
    from tkcalendar import DateEntry  # type: ignore
except Exception:  # pragma: no cover
    DateEntry = None  # type: ignore

import tkinter as tk
from tkinter import font as tkfont
from tkinter import messagebox
from tkinter import ttk

try:
    import fitz  # pymupdf
    from PIL import Image, ImageTk
except ImportError:
    fitz = None
    Image = None
    ImageTk = None


# ---------------- Theme / palette ----------------
PALETTE = {
    "bg":            "#f4f6fb",  # app background
    "surface":       "#ffffff",  # panels / entries
    "surface_alt":   "#eef2f9",  # subtle zebra / preview bg
    "border":        "#d7dde8",  # subtle borders
    "text":          "#1e293b",  # primary text
    "muted":         "#64748b",  # secondary text
    "accent":        "#2563eb",  # primary action
    "accent_hover":  "#1d4ed8",
    "accent_active": "#1e40af",
    "accent_fg":     "#ffffff",
    "success":       "#16a34a",
    "warning":       "#d97706",
    "danger":        "#dc2626",
}


def _setup_theme(root: tk.Tk) -> ttk.Style:
    style = ttk.Style(root)
    try:
        style.theme_use("clam")
    except Exception:
        pass

    # Base fonts (Segoe UI is default on Windows 7+)
    base_family = "Segoe UI"
    try:
        default_font = tkfont.nametofont("TkDefaultFont")
        default_font.configure(family=base_family, size=10)
        text_font = tkfont.nametofont("TkTextFont")
        text_font.configure(family=base_family, size=10)
        heading_font = tkfont.nametofont("TkHeadingFont")
        heading_font.configure(family=base_family, size=10, weight="bold")
    except Exception:
        pass

    bg = PALETTE["bg"]
    surface = PALETTE["surface"]
    border = PALETTE["border"]
    text = PALETTE["text"]
    muted = PALETTE["muted"]
    accent = PALETTE["accent"]

    root.configure(background=bg)

    # Generic widgets
    style.configure(".", background=bg, foreground=text, font=(base_family, 10))
    style.configure("TFrame", background=bg)
    style.configure("TLabel", background=bg, foreground=text)
    style.configure("Muted.TLabel", background=bg, foreground=muted)
    style.configure(
        "Header.TLabel", background=bg, foreground=text,
        font=(base_family, 16, "bold"),
    )
    style.configure(
        "SubHeader.TLabel", background=bg, foreground=muted,
        font=(base_family, 10),
    )
    style.configure(
        "Section.TLabel", background=bg, foreground=text,
        font=(base_family, 10, "bold"),
    )

    # LabelFrame as light card
    style.configure(
        "Card.TLabelframe",
        background=surface,
        bordercolor=border,
        lightcolor=border,
        darkcolor=border,
        relief="solid",
        borderwidth=1,
        padding=10,
    )
    style.configure(
        "Card.TLabelframe.Label",
        background=surface,
        foreground=text,
        font=(base_family, 10, "bold"),
        padding=(4, 0),
    )
    # Panels living inside a card get surface bg so they blend in
    style.configure("Card.TFrame", background=surface)
    style.configure("Card.TLabel", background=surface, foreground=text)
    style.configure("CardMuted.TLabel", background=surface, foreground=muted)

    # Entries / comboboxes
    style.configure(
        "TEntry",
        fieldbackground=surface,
        background=surface,
        foreground=text,
        bordercolor=border,
        lightcolor=border,
        darkcolor=border,
        insertcolor=text,
        padding=4,
    )
    style.map(
        "TEntry",
        bordercolor=[("focus", accent)],
        lightcolor=[("focus", accent)],
        darkcolor=[("focus", accent)],
    )
    style.configure(
        "TCombobox",
        fieldbackground=surface,
        background=surface,
        foreground=text,
        bordercolor=border,
        arrowcolor=text,
        padding=4,
    )
    style.map(
        "TCombobox",
        fieldbackground=[("readonly", surface)],
        foreground=[("readonly", text)],
        bordercolor=[("focus", accent)],
    )
    try:
        root.option_add("*TCombobox*Listbox.background", surface)
        root.option_add("*TCombobox*Listbox.foreground", text)
        root.option_add("*TCombobox*Listbox.selectBackground", accent)
        root.option_add("*TCombobox*Listbox.selectForeground", PALETTE["accent_fg"])
        root.option_add("*TCombobox*Listbox.font", (base_family, 10))
    except Exception:
        pass

    # Buttons
    style.configure(
        "TButton",
        background=surface,
        foreground=text,
        bordercolor=border,
        lightcolor=border,
        darkcolor=border,
        focusthickness=0,
        padding=(12, 6),
    )
    style.map(
        "TButton",
        background=[("active", PALETTE["surface_alt"]), ("pressed", PALETTE["surface_alt"])],
        bordercolor=[("active", accent), ("focus", accent)],
    )
    style.configure(
        "Accent.TButton",
        background=accent,
        foreground=PALETTE["accent_fg"],
        bordercolor=accent,
        lightcolor=accent,
        darkcolor=accent,
        focusthickness=0,
        padding=(16, 8),
        font=(base_family, 10, "bold"),
    )
    style.map(
        "Accent.TButton",
        background=[
            ("pressed", PALETTE["accent_active"]),
            ("active", PALETTE["accent_hover"]),
            ("disabled", "#93c5fd"),
        ],
        foreground=[("disabled", "#e5e7eb")],
        bordercolor=[("active", PALETTE["accent_hover"])],
    )

    # Radios / checkbuttons (on card surface)
    style.configure("Card.TRadiobutton", background=surface, foreground=text)
    style.map(
        "Card.TRadiobutton",
        background=[("active", surface)],
        foreground=[("disabled", muted)],
    )
    style.configure("Card.TCheckbutton", background=surface, foreground=text)
    style.map("Card.TCheckbutton", background=[("active", surface)])

    # Radios on app bg (e.g. Джерело)
    style.configure("TRadiobutton", background=bg, foreground=text)
    style.map("TRadiobutton", background=[("active", bg)])
    style.configure("TCheckbutton", background=bg, foreground=text)
    style.map("TCheckbutton", background=[("active", bg)])

    # Progressbar
    style.configure(
        "Horizontal.TProgressbar",
        background=accent,
        troughcolor=PALETTE["surface_alt"],
        bordercolor=border,
        lightcolor=accent,
        darkcolor=accent,
    )

    # Separator + scrollbar
    style.configure("TSeparator", background=border)
    style.configure(
        "Vertical.TScrollbar",
        background=PALETTE["surface_alt"],
        troughcolor=bg,
        bordercolor=bg,
        arrowcolor=muted,
    )
    style.configure(
        "Horizontal.TScrollbar",
        background=PALETTE["surface_alt"],
        troughcolor=bg,
        bordercolor=bg,
        arrowcolor=muted,
    )

    # Status bar
    style.configure("Status.TFrame", background=PALETTE["surface_alt"])
    style.configure(
        "Status.TLabel",
        background=PALETTE["surface_alt"],
        foreground=text,
        padding=(10, 6),
    )
    style.configure(
        "StatusDot.TLabel",
        background=PALETTE["surface_alt"],
        foreground=PALETTE["success"],
        font=(base_family, 12, "bold"),
        padding=(10, 4, 0, 4),
    )

    return style


INVALID_WIN_CHARS_RE = re.compile(r'[\\/:*?"<>|]+')
EXTRA_BAD_CHARS_RE = re.compile(r"[,;()\[\]{}!@#$%^&+=~`'\u00ab\u00bb\u201c\u201d\u2018\u2019]+")
WS_RE = re.compile(r"\s+")

MAX_FILENAME_LEN = 200  # keep total name+ext well below Windows MAX_PATH
MAX_NAME_PART_LEN = 60
MAX_ADDR_PART_LEN = 80

SUFFIX_HISTORY_FILENAME = "pdf_rename_suffix_history.json"
MAX_SUFFIX_HISTORY_ITEMS = 40

WORK_CODES: Dict[str, str] = {
    "Монтаж АСКОЕ": "Монт-АСКОЕ",
    "Монтаж лічильника": "Монт-Ліч",
    "Монтаж АСКОЕ + монтаж лічильника": "АСКОЕ-Ліч",
    "Технічна перевірка": "Тех-Перев",
}
CHECK_CODES: Dict[str, str] = {
    "Планова": "План",
    "Позапланова": "Позаплан",
    "Виконання припису": "Припис",
    "Заміна лічильника": "Зам-Ліч",
}

EXPORT_COLUMNS: List[str] = [
    "№",
    "№Скану",
    "EIC",
    "Дата",
    "№Лічильника",
    "Назва",
    "Адреса",
    "Інвест",
    "Фаза",
    "Вид перевірки",
    "Вид роботи",
    "IP адреса",
    "№ Модему",
    "Належність лічильника",
]


def load_suffix_history(app_dir: Path) -> List[str]:
    path = app_dir / SUFFIX_HISTORY_FILENAME
    if not path.is_file():
        return []
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        if not isinstance(data, list):
            return []
        out: List[str] = []
        for item in data:
            if isinstance(item, str):
                s = item.strip()
                if s and s not in out:
                    out.append(s)
        return out[:MAX_SUFFIX_HISTORY_ITEMS]
    except Exception:
        return []


def save_suffix_history(app_dir: Path, items: List[str]) -> None:
    path = app_dir / SUFFIX_HISTORY_FILENAME
    try:
        seen: set = set()
        ordered: List[str] = []
        for item in items:
            s = (item or "").strip()
            if not s or s in seen:
                continue
            seen.add(s)
            ordered.append(s)
            if len(ordered) >= MAX_SUFFIX_HISTORY_ITEMS:
                break
        with open(path, "w", encoding="utf-8") as f:
            json.dump(ordered, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


def sanitize_component(value: str, *, spaces_to_dash: bool = False) -> str:
    s = (value or "").strip()
    s = INVALID_WIN_CHARS_RE.sub("-", s)
    s = EXTRA_BAD_CHARS_RE.sub("-", s)
    if spaces_to_dash:
        s = WS_RE.sub("-", s)
    else:
        s = WS_RE.sub(" ", s)
    s = re.sub(r"-{2,}", "-", s)
    s = s.strip(" -._")
    return s


def truncate_part(value: str, max_len: int) -> str:
    if len(value) <= max_len:
        return value
    return value[:max_len].rstrip(" -._")


def ensure_unique_path(path: Path) -> Path:
    if not path.exists():
        return path
    stem = path.stem
    suffix = path.suffix
    parent = path.parent
    i = 1
    while True:
        candidate = parent / f"{stem} ({i}){suffix}"
        if not candidate.exists():
            return candidate
        i += 1


def extract_scan_number(filename: str) -> str:
    m = re.search(r"(\d{4})", filename)
    return m.group(1) if m else "0000"


def parse_renamed_pdf_filename(filename: str) -> Optional[dict]:
    """
    Parse filename produced by this app:
      scan_eic_date_meter_name_address_[Інвест]_phase_check_work_[IP-xxx]_[Mod-yyy]_[належність].pdf
    """
    if not filename.lower().endswith(".pdf"):
        return None
    base = filename[:-4]
    parts = base.split("_")
    if len(parts) < 9:
        return None

    # Base fixed parts
    scan = parts[0]
    eic = parts[1]
    dt = parts[2]
    meter = parts[3]
    nm = parts[4]
    addr = parts[5]

    invest = ""
    idx = 6
    if idx < len(parts) and parts[idx] == "Інвест":
        invest = "Так"
        idx += 1
    else:
        invest = "Ні"

    if idx + 2 >= len(parts):
        return None

    phase = parts[idx]
    check = parts[idx + 1]
    work = parts[idx + 2]
    
    idx += 3
    ip = ""
    modem = ""
    extra_chunks: List[str] = []

    # Optional ASKOE parts + належність лічильника у кінці імені
    while idx < len(parts):
        p = parts[idx]
        if p.startswith("IP-"):
            ip = p[3:]
        elif p.startswith("Mod-"):
            modem = p[4:]
        else:
            extra_chunks.append(p)
        idx += 1

    extra = "_".join(extra_chunks) if extra_chunks else ""

    return {
        "№Скану": scan,
        "EIC": eic,
        "Дата": dt,
        "№Лічильника": meter,
        "Назва": nm,
        "Адреса": addr,
        "Інвест": invest,
        "Фаза": phase,
        "Вид перевірки": check,
        "Вид роботи": work,
        "IP адреса": ip,
        "№ Модему": modem,
        "Належність лічильника": extra,
    }


@dataclass
class ConsumerRecord:
    eic_full: str
    name: str
    address: str

    def display(self) -> str:
        # Keep EIC first for quick scanning in dropdown.
        return f"{self.eic_full} — {self.name} — {self.address}"


class ExcelConsumerDB:
    """
    Reads the Excel workbook:
      - file: next to the app/exe (any *.xlsx, prefer expected name)
      - sheet: 'база'
      - columns by letter:
          AB (EIC), T + AD (Name via dash), AN (Address)
    """

    # Column letters -> 0-based indices
    _IDX_T = 19
    _IDX_AB = 27
    _IDX_AD = 29
    _IDX_AN = 39

    def __init__(self, app_dir: Path) -> None:
        self.app_dir = app_dir
        self.excel_path = self._resolve_excel_path()
        self._df: Optional[pd.DataFrame] = None
        self._eic_col: Optional[List[str]] = None
        self._col_t: Optional[List[str]] = None
        self._col_ad: Optional[List[str]] = None
        self._col_an: Optional[List[str]] = None
        self._suffix5_index: Optional[Dict[str, List[int]]] = None

    def _resolve_excel_path(self) -> Optional[Path]:
        expected = "Споживачі ЕЕЛ-2 (ЕЧ-2).xlsx"
        candidate = self.app_dir / expected
        if candidate.exists():
            return candidate
        xlsx = sorted(self.app_dir.glob("*.xlsx"))
        return xlsx[0] if xlsx else None

    def load(self) -> None:
        if self.excel_path is None or not self.excel_path.exists():
            raise FileNotFoundError(
                "Не знайдено Excel-базу поруч із програмою (*.xlsx)."
            )

        # Read columns up to AN so indices stay stable.
        df = pd.read_excel(
            self.excel_path,
            sheet_name="база",
            engine="openpyxl",
            usecols="A:AN",
            dtype=str,
        )
        self._df = df

        # Preload columns and build an index by last 5 chars of EIC for fast lookup.
        # This makes interactive search much faster on large workbooks.
        if df.shape[1] <= self._IDX_AN:
            raise ValueError("Excel-таблиця має недостатньо стовпчиків (очікується до AN).")

        eic = (
            df.iloc[:, self._IDX_AB]
            .fillna("")
            .astype(str)
            .str.strip()
            .tolist()
        )
        col_t = df.iloc[:, self._IDX_T].fillna("").astype(str).str.strip().tolist()
        col_ad = df.iloc[:, self._IDX_AD].fillna("").astype(str).str.strip().tolist()
        col_an = df.iloc[:, self._IDX_AN].fillna("").astype(str).str.strip().tolist()

        idx: Dict[str, List[int]] = {}
        for i, v in enumerate(eic):
            if not v:
                continue
            key = v[-5:] if len(v) >= 5 else v
            idx.setdefault(key, []).append(i)

        self._eic_col = eic
        self._col_t = col_t
        self._col_ad = col_ad
        self._col_an = col_an
        self._suffix5_index = idx

    def _require_loaded(self) -> pd.DataFrame:
        if self._df is None:
            self.load()
        assert self._df is not None
        return self._df

    def search_by_eic_suffix(self, suffix: str) -> List[ConsumerRecord]:
        suf = (suffix or "").strip()
        if len(suf) < 5:
            return []

        self._require_loaded()
        assert self._eic_col is not None
        assert self._col_t is not None
        assert self._col_ad is not None
        assert self._col_an is not None
        assert self._suffix5_index is not None

        suf5 = suf[-5:]
        candidates = self._suffix5_index.get(suf5, [])
        if not candidates:
            return []

        # If user typed more than 5 chars, narrow down within candidates only.
        if len(suf) > 5:
            candidates = [i for i in candidates if self._eic_col[i].endswith(suf)]

        out: List[ConsumerRecord] = []
        for i in candidates:
            eic_full = self._eic_col[i]
            t = self._col_t[i]
            ad = self._col_ad[i]
            an = self._col_an[i]
            name = f"{t}-{ad}".strip("-").strip()
            out.append(ConsumerRecord(eic_full=eic_full, name=name, address=an))
        return out


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

        self.var_source = tk.StringVar(value="Населення")
        self.var_eic_suffix = tk.StringVar()
        self.var_eic_full = tk.StringVar()
        self.var_name = tk.StringVar()
        self.var_address = tk.StringVar()
        self.var_meter_no = tk.StringVar()
        self.var_invest = tk.BooleanVar(value=False)
        self.var_phase = tk.StringVar(value="3Ф")
        self.var_phase_ktt = tk.StringVar(value="XXX")
        self.var_work_type = tk.StringVar(value="Монтаж АСКОЕ")
        self.var_work_custom = tk.StringVar()
        self.var_ip = tk.StringVar()
        self.var_modem = tk.StringVar()
        self.var_check_type = tk.StringVar(value="Планова")
        self.var_pdf_name = tk.StringVar()
        self.var_preview_len = tk.StringVar(value="0")
        self.var_status = tk.StringVar(value="Готово.")
        self.var_name_suffix = tk.StringVar()
        self._suffix_history: List[str] = load_suffix_history(app_dir)

        self._current_photos = []  # List to keep references to all page images
        self._preview_timer = None

        self._build_ui()
        self._wire_events()
        self._refresh_file_list()
        self._update_preview()

        # Auto-load database in background on startup.
        self.after(200, self._load_db_into_memory)
        self._toggle_askoe_fields()

    # ---------------- UI ----------------
    def _build_ui(self) -> None:
        self.master.title("PDF Rename Expert")
        self.master.minsize(1280, 800)
        self.master.configure(background=PALETTE["bg"])
        self.pack(fill="both", expand=True, padx=14, pady=10)

        self.columnconfigure(0, weight=1)
        self.rowconfigure(2, weight=1)

        # --- Header ---
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

        # --- Toolbar (source + refresh) ---
        toolbar = ttk.Frame(self)
        toolbar.grid(row=1, column=0, sticky="ew", pady=(8, 8))
        toolbar.columnconfigure(2, weight=1)

        ttk.Label(toolbar, text="Джерело:", style="Section.TLabel").grid(
            row=0, column=0, sticky="w", padx=(0, 10)
        )
        src_frame = ttk.Frame(toolbar)
        src_frame.grid(row=0, column=1, sticky="w")
        ttk.Radiobutton(
            src_frame, text="Населення", value="Населення", variable=self.var_source
        ).grid(row=0, column=0, sticky="w")
        ttk.Radiobutton(
            src_frame, text="Промислові", value="Промислові", variable=self.var_source
        ).grid(row=0, column=1, sticky="w", padx=(18, 0))

        ttk.Button(
            toolbar, text="Оновити список (F5)", command=self._refresh_file_list
        ).grid(row=0, column=3, sticky="e")

        # --- Body split ---
        body = ttk.PanedWindow(self, orient="horizontal")
        body.grid(row=2, column=0, sticky="nsew")

        left = ttk.Frame(body)
        center = ttk.Frame(body)
        right = ttk.Frame(body)
        body.add(left, weight=2)
        body.add(center, weight=5)
        body.add(right, weight=3)

        # ---- Left card: file list ----
        lf = ttk.LabelFrame(left, text="  Скан-файли (PDF)  ", style="Card.TLabelframe")
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

        self.list_files = tk.Listbox(
            list_wrap,
            height=18,
            exportselection=False,
            activestyle="dotbox",
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
        self.list_files.grid(row=0, column=0, sticky="nsew")
        sb = ttk.Scrollbar(list_wrap, orient="vertical", command=self.list_files.yview)
        sb.grid(row=0, column=1, sticky="ns")
        self.list_files.configure(yscrollcommand=sb.set)

        # ---- Center: PDF Preview ----
        cf = ttk.LabelFrame(center, text="  Перегляд вмісту   ", style="Card.TLabelframe")
        cf.pack(fill="both", expand=True, padx=4)
        cf.columnconfigure(0, weight=1)
        cf.rowconfigure(0, weight=1)
        cf.columnconfigure(0, weight=1)

        self.preview_canvas = tk.Canvas(
            cf,
            background=PALETTE["surface_alt"],
            highlightthickness=0,
            borderwidth=0
        )
        self.preview_canvas.grid(row=0, column=0, sticky="nsew")

        # Scrollbars for canvas
        self.vsb_preview = ttk.Scrollbar(cf, orient="vertical", command=self.preview_canvas.yview)
        self.vsb_preview.grid(row=0, column=1, sticky="ns")
        self.hsb_preview = ttk.Scrollbar(cf, orient="horizontal", command=self.preview_canvas.xview)
        self.hsb_preview.grid(row=1, column=0, sticky="ew")

        self.preview_canvas.configure(
            yscrollcommand=self.vsb_preview.set,
            xscrollcommand=self.hsb_preview.set
        )

        # Overlay label for messages
        self.lbl_preview_msg = ttk.Label(cf, text="Оберіть файл для перегляду", style="CardMuted.TLabel")
        self.lbl_preview_msg.place(relx=0.5, rely=0.5, anchor="center")

        # ---- Right: inputs ----
        rf = ttk.Frame(right)
        rf.pack(fill="both", expand=True, padx=(4, 0))
        rf.columnconfigure(0, weight=1)
        rf.columnconfigure(1, weight=1)

        # Consumer search card
        grp_search = ttk.LabelFrame(
            rf, text="  Пошук споживача  ", style="Card.TLabelframe"
        )
        grp_search.grid(row=0, column=0, columnspan=2, sticky="ew")
        grp_search.columnconfigure(1, weight=1)

        ttk.Label(
            grp_search, text="EIC (останні 5+ символів):", style="Card.TLabel"
        ).grid(row=0, column=0, sticky="w", pady=(0, 6), padx=(0, 10))
        self.ent_eic = ttk.Entry(grp_search, textvariable=self.var_eic_suffix)
        self.ent_eic.grid(row=0, column=1, sticky="ew", pady=(0, 6))
        self.btn_load_db = ttk.Button(
            grp_search, text="Завантажити базу", command=self._load_db_into_memory
        )
        self.btn_load_db.grid(row=0, column=2, sticky="e", padx=(10, 0), pady=(0, 6))

        ttk.Label(
            grp_search, text="EIC (повний, 16 символів):", style="Card.TLabel"
        ).grid(row=1, column=0, sticky="w", pady=(0, 6), padx=(0, 10))
        self.ent_eic_full = ttk.Entry(grp_search, textvariable=self.var_eic_full)
        self.ent_eic_full.grid(row=1, column=1, sticky="ew", pady=(0, 6))
        self.pb_db = ttk.Progressbar(grp_search, mode="indeterminate", length=160)
        self.pb_db.grid(row=1, column=2, sticky="e", padx=(10, 0), pady=(0, 6))
        self.pb_db.grid_remove()

        ttk.Label(grp_search, text="Збіги:", style="Card.TLabel").grid(
            row=2, column=0, sticky="nw", padx=(0, 10), pady=(2, 0)
        )
        matches_wrap = ttk.Frame(grp_search)
        matches_wrap.grid(row=2, column=1, columnspan=2, sticky="ew")
        matches_wrap.columnconfigure(0, weight=1)
        self.list_matches = tk.Listbox(
            matches_wrap,
            height=2,
            exportselection=False,
            activestyle="dotbox",
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
        self.list_matches.grid(row=0, column=0, sticky="ew")
        sb_matches = ttk.Scrollbar(
            matches_wrap, orient="vertical", command=self.list_matches.yview
        )
        sb_matches.grid(row=0, column=1, sticky="ns")
        self.list_matches.configure(yscrollcommand=sb_matches.set)

        # Document data card
        grp_doc = ttk.LabelFrame(
            rf, text="  Дані документа  ", style="Card.TLabelframe"
        )
        grp_doc.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(10, 0))
        grp_doc.columnconfigure(1, weight=1)

        r = 0
        ttk.Label(grp_doc, text="Назва (T-AD):", style="Card.TLabel").grid(
            row=r, column=0, sticky="w", pady=(0, 6), padx=(0, 10)
        )
        self.ent_name = ttk.Entry(grp_doc, textvariable=self.var_name)
        self.ent_name.grid(row=r, column=1, sticky="ew", pady=(0, 6))

        r += 1
        ttk.Label(grp_doc, text="Адреса (AN):", style="Card.TLabel").grid(
            row=r, column=0, sticky="w", pady=(0, 6), padx=(0, 10)
        )
        self.ent_addr = ttk.Entry(grp_doc, textvariable=self.var_address)
        self.ent_addr.grid(row=r, column=1, sticky="ew", pady=(0, 6))

        r += 1
        ttk.Label(grp_doc, text="Дата:", style="Card.TLabel").grid(
            row=r, column=0, sticky="w", pady=(0, 6), padx=(0, 10)
        )
        date_frame = ttk.Frame(grp_doc, style="Card.TFrame")
        date_frame.grid(row=r, column=1, sticky="ew", pady=(0, 6))
        date_frame.columnconfigure(0, weight=0)

        if DateEntry is not None:
            self.date_entry = DateEntry(
                date_frame,
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
            self.date_entry.grid(row=0, column=0, sticky="w")
        else:
            self.date_entry = ttk.Entry(date_frame)
            self.date_entry.insert(0, date.today().isoformat())
            self.date_entry.grid(row=0, column=0, sticky="w")

        ttk.Checkbutton(
            date_frame,
            text="Інвест",
            variable=self.var_invest,
            style="Card.TCheckbutton",
        ).grid(row=0, column=1, sticky="w", padx=(18, 0))

        r += 1
        ttk.Label(grp_doc, text="№ Лічильника:", style="Card.TLabel").grid(
            row=r, column=0, sticky="w", pady=(0, 6), padx=(0, 10)
        )
        self.ent_meter = ttk.Entry(grp_doc, textvariable=self.var_meter_no)
        self.ent_meter.grid(row=r, column=1, sticky="ew", pady=(0, 6))

        r += 1
        ttk.Label(grp_doc, text="Фаза:", style="Card.TLabel").grid(
            row=r, column=0, sticky="w", padx=(0, 10)
        )
        phase_frame = ttk.Frame(grp_doc, style="Card.TFrame")
        phase_frame.grid(row=r, column=1, sticky="ew")

        self.cmb_phase = ttk.Combobox(
            phase_frame,
            textvariable=self.var_phase,
            values=["1Ф", "3Ф", "3Ф-Ктт-XXX"],
            width=12,
        )
        self.cmb_phase.grid(row=0, column=0, sticky="w")
        self.ent_phase_ktt = ttk.Entry(
            phase_frame, textvariable=self.var_phase_ktt, width=10
        )
        self.ent_phase_ktt.grid(row=0, column=1, sticky="w", padx=(10, 0))
        ttk.Label(
            phase_frame, text="(для 3Ф-Ктт)", style="CardMuted.TLabel"
        ).grid(row=0, column=2, sticky="w", padx=(10, 0))

        # Work + check in one row (2 columns) to reduce window height
        wc_row = ttk.Frame(rf)
        wc_row.grid(row=2, column=0, columnspan=2, sticky="ew", pady=(10, 0))
        wc_row.columnconfigure(0, weight=1)
        wc_row.columnconfigure(1, weight=1)

        # Work type card
        grp_work = ttk.LabelFrame(
            wc_row, text="  Вид роботи  ", style="Card.TLabelframe"
        )
        grp_work.grid(row=0, column=0, sticky="nsew", padx=(0, 6))
        grp_work.columnconfigure(0, weight=1)
        grp_work.columnconfigure(1, weight=1)

        work_types = [
            "Монтаж АСКОЕ",
            "Монтаж лічильника",
            "Монтаж АСКОЕ + монтаж лічильника",
            "Технічна перевірка",
        ]
        for i, wt in enumerate(work_types):
            rr = i // 2
            cc = i % 2
            ttk.Radiobutton(
                grp_work,
                text=wt,
                value=wt,
                variable=self.var_work_type,
                style="Card.TRadiobutton",
            ).grid(
                row=rr,
                column=cc,
                sticky="w",
                padx=(0, 10) if cc == 0 else (10, 0),
                pady=(0 if rr == 0 else 4, 0),
            )
        row_other = (len(work_types) + 1) // 2
        other_row = ttk.Frame(grp_work, style="Card.TFrame")
        other_row.grid(row=row_other, column=0, columnspan=2, sticky="ew", pady=(4, 0))
        other_row.columnconfigure(1, weight=1)
        ttk.Radiobutton(
            other_row,
            text="Інше",
            value="Інше",
            variable=self.var_work_type,
            style="Card.TRadiobutton",
        ).grid(row=0, column=0, sticky="w")
        self.ent_work_custom = ttk.Entry(other_row, textvariable=self.var_work_custom)
        self.ent_work_custom.grid(row=0, column=1, sticky="ew", padx=(10, 0))

        # Check type card
        grp_check = ttk.LabelFrame(
            wc_row, text="  Вид перевірки  ", style="Card.TLabelframe"
        )
        grp_check.grid(row=0, column=1, sticky="nsew", padx=(6, 0))
        grp_check.columnconfigure(0, weight=1)
        grp_check.columnconfigure(1, weight=1)

        check_types = [
            "Планова",
            "Позапланова",
            "Виконання припису",
            "Заміна лічильника",
        ]
        for i, ct in enumerate(check_types):
            rr = i // 2
            cc = i % 2
            ttk.Radiobutton(
                grp_check,
                text=ct,
                value=ct,
                variable=self.var_check_type,
                style="Card.TRadiobutton",
            ).grid(
                row=rr,
                column=cc,
                sticky="w",
                padx=(0, 10) if cc == 0 else (10, 0),
                pady=(0 if rr == 0 else 4, 0),
            )

        # ASKOE extra fields (initially hidden)
        self.grp_askoe = ttk.LabelFrame(
            rf, text="  Дані АСКОЕ  ", style="Card.TLabelframe"
        )
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

        grp_suffix = ttk.LabelFrame(
            rf, text="  Належність лічильника  ", style="Card.TLabelframe"
        )
        grp_suffix.grid(row=4, column=0, columnspan=2, sticky="ew", pady=(10, 0))
        grp_suffix.columnconfigure(0, weight=1)
        self.cmb_name_suffix = ttk.Combobox(
            grp_suffix,
            textvariable=self.var_name_suffix,
            values=self._suffix_history,
        )
        self.cmb_name_suffix.grid(row=0, column=0, sticky="ew")

        # Preview card (monospace)
        grp_preview = ttk.LabelFrame(
            rf, text="  Ім'я PDF (прев'ю)  ", style="Card.TLabelframe"
        )
        grp_preview.grid(row=5, column=0, columnspan=2, sticky="ew", pady=(10, 0))
        grp_preview.columnconfigure(0, weight=1)

        # Header for preview (to hold length label)
        preview_header = ttk.Frame(grp_preview, style="Card.TFrame")
        preview_header.grid(row=0, column=0, sticky="ew", padx=10, pady=(0, 4))
        preview_header.columnconfigure(1, weight=1)

        ttk.Label(preview_header, text="Довжина ім'я:", style="CardMuted.TLabel").grid(row=0, column=0)
        ttk.Label(preview_header, textvariable=self.var_preview_len, style="Card.TLabel", font=("Segoe UI", 10, "bold")).grid(row=0, column=1, sticky="w", padx=(5, 0))
        ttk.Label(preview_header, text=f"(макс. {MAX_FILENAME_LEN})", style="CardMuted.TLabel").grid(row=0, column=2, sticky="e")

        preview_frame = ttk.Frame(grp_preview, style="Card.TFrame")
        preview_frame.grid(row=1, column=0, sticky="ew")
        preview_frame.columnconfigure(0, weight=1)

        self.txt_preview = tk.Text(
            preview_frame,
            height=2,
            wrap="none",
            font=("Consolas", 10),
            background=PALETTE["surface_alt"],
            foreground=PALETTE["text"],
            relief="flat",
            borderwidth=1,
            highlightthickness=1,
            highlightbackground=PALETTE["border"],
            highlightcolor=PALETTE["accent"],
            padx=10,
            pady=8,
        )
        self.txt_preview.grid(row=0, column=0, sticky="ew")
        xsb = ttk.Scrollbar(
            preview_frame, orient="horizontal", command=self.txt_preview.xview
        )
        xsb.grid(row=1, column=0, sticky="ew", pady=(4, 0))
        self.txt_preview.configure(xscrollcommand=xsb.set)

        ttk.Button(preview_frame, text="Копіювати", command=self._copy_preview).grid(
            row=0, column=1, sticky="ns", padx=(10, 0)
        )

        # Actions
        actions = ttk.Frame(rf)
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

        # --- Status bar ---
        status = ttk.Frame(self, style="Status.TFrame")
        status.grid(row=3, column=0, sticky="ew", pady=(10, 0))
        status.columnconfigure(1, weight=1)
        self.lbl_status_dot = ttk.Label(status, text="\u25CF", style="StatusDot.TLabel")
        self.lbl_status_dot.grid(row=0, column=0, sticky="w")
        ttk.Label(
            status, textvariable=self.var_status, style="Status.TLabel"
        ).grid(row=0, column=1, sticky="w")

    def _wire_events(self) -> None:
        self.var_source.trace_add("write", lambda *_: self._refresh_file_list())

        # Use both key events and variable trace so paste/mouse input also triggers search.
        self.ent_eic.bind("<KeyRelease>", lambda _e: self._schedule_eic_search())
        self.var_eic_suffix.trace_add("write", lambda *_: self._schedule_eic_search())
        self.list_matches.bind("<<ListboxSelect>>", lambda _e: self._apply_selected_match())
        self.var_work_type.trace_add("write", lambda *_: self._toggle_askoe_fields())

        for v in [
            self.var_eic_full,
            self.var_name,
            self.var_address,
            self.var_meter_no,
            self.var_invest,
            self.var_phase,
            self.var_phase_ktt,
            self.var_work_type,
            self.var_work_custom,
            self.var_ip,
            self.var_modem,
            self.var_check_type,
            self.var_name_suffix,
        ]:
            v.trace_add("write", lambda *_: self._update_preview())

        self.list_files.bind("<<ListboxSelect>>", lambda _e: self._on_file_selected())
        self.list_files.bind("<Double-Button-1>", lambda _e: self._open_selected_pdf())
        
        # Mouse wheel for PDF preview
        self.preview_canvas.bind("<MouseWheel>", self._on_mousewheel)

        if DateEntry is not None:
            self.date_entry.bind("<<DateEntrySelected>>", lambda _e: self._update_preview())
        else:
            self.date_entry.bind("<KeyRelease>", lambda _e: self._update_preview())

        # Global hotkeys
        self.master.bind("<F5>", lambda _e: self._refresh_file_list())
        self.master.bind("<Control-Return>", lambda _e: self._process_selected())
        self.ent_eic.bind("<Return>", lambda _e: self._accept_first_match())

        self._wire_editing_shortcuts()

    def _wire_editing_shortcuts(self) -> None:
        """Ctrl+A у полях; Ctrl+C у списках PDF/збігів — копіювання рядка."""

        def select_all_entry(event: tk.Event) -> str:
            w = event.widget
            try:
                w.select_range(0, tk.END)
                w.icursor(tk.END)
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
            w = event.widget
            try:
                sel = w.curselection()
                if sel:
                    self.clipboard_clear()
                    self.clipboard_append(w.get(sel[0]))
            except Exception:
                pass

        entry_like = (
            self.ent_eic,
            self.ent_eic_full,
            self.ent_name,
            self.ent_addr,
            self.ent_meter,
            self.cmb_phase,
            self.ent_phase_ktt,
            self.ent_work_custom,
            self.ent_ip,
            self.ent_modem,
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

        self.list_files.bind("<Control-c>", listbox_copy)
        self.list_files.bind("<Control-C>", listbox_copy)
        self.list_matches.bind("<Control-c>", listbox_copy)
        self.list_matches.bind("<Control-C>", listbox_copy)

    # ---------------- Data / matching ----------------
    def _schedule_eic_search(self) -> None:
        # Debounce search to avoid scanning on every keystroke.
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
        wt = self.var_work_type.get()
        is_askoe = "АСКОЕ" in wt
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

    def _on_file_selected(self) -> None:
        self._update_preview()
        pdf = self._selected_pdf_path()
        if pdf and pdf.exists():
            # Debounce rendering slightly if user scrolls fast
            if self._preview_timer:
                self.after_cancel(self._preview_timer)
            self._preview_timer = self.after(150, lambda: self._display_pdf(pdf))

    def _display_pdf(self, pdf_path: Path) -> None:
        if not fitz or not Image or not ImageTk:
            self.lbl_preview_msg.configure(text="Помилка: PyMuPDF/Pillow не встановлено.")
            return

        try:
            doc = fitz.open(str(pdf_path))
            if doc.page_count == 0:
                self._clear_pdf_preview("Порожній PDF")
                doc.close()
                return

            # Prepare canvas and state
            self.preview_canvas.delete("all")
            self._current_photos = []
            
            # Mode: Fit to Width (with scroll)
            self.master.update_idletasks()
            cw = self.preview_canvas.winfo_width()
            if cw < 50: cw = 650 # Fallback
            nw = cw - 40 # Padding for scrollbar
            
            current_y = 0
            gap = 15 # Gap between pages
            
            for i in range(doc.page_count):
                page = doc.load_page(i)
                pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))
                img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                
                # Proportional scale
                r_img = img.width / img.height
                nh = int(nw / r_img)
                img = img.resize((nw, nh), Image.Resampling.LANCZOS)
                
                photo = ImageTk.PhotoImage(img)
                self._current_photos.append(photo)
                
                # Place page on canvas
                self.preview_canvas.create_image(cw // 2, current_y, image=photo, anchor="n")
                current_y += nh + gap

            doc.close()

            # Set scrollregion
            self.preview_canvas.config(scrollregion=(0, 0, cw, current_y))
            self.preview_canvas.yview_moveto(0)
            self.lbl_preview_msg.place_forget()

        except Exception as e:
            logging.exception("PDF preview failed")
            self._clear_pdf_preview(f"Помилка завантаження:\n{e}")

    def _clear_pdf_preview(self, msg: str = "Оберіть файл для перегляду") -> None:
        self.preview_canvas.delete("all")
        self._current_photos = []
        self.preview_canvas.config(scrollregion=(0, 0, 0, 0))
        self.lbl_preview_msg.configure(text=msg)
        self.lbl_preview_msg.place(relx=0.5, rely=0.5, anchor="center")

    def _on_mousewheel(self, event: tk.Event) -> None:
        # Standard cross-platform or Windows-centric scrolling
        if event.delta:
            self.preview_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        elif event.num == 4: # Linux
            self.preview_canvas.yview_scroll(-1, "units")
        elif event.num == 5: # Linux
            self.preview_canvas.yview_scroll(1, "units")

    # ---------------- Files ----------------
    def _source_dir(self) -> Path:
        base = self.app_dir / "Scan"
        return base / self.var_source.get()

    def _report_path(self, year: str, month: str) -> Path:
        return self.app_dir / f"Звіт_Переіменовані_PDF_{year}-{month}.xlsx"

    def _refresh_file_list(self) -> None:
        src = self._source_dir()
        self.lbl_folder.configure(text=str(src))
        self.list_files.delete(0, tk.END)

        if not src.exists():
            try:
                src.mkdir(parents=True, exist_ok=True)
            except Exception as e:
                messagebox.showerror(
                    "PDF_Rename_Expert",
                    f"Не вдалося створити папку джерела:\n{src}\n\n{e}",
                )
                self._set_status("Помилка створення папки джерела.", "error")
                return

        pdfs = sorted(
            [p for p in src.iterdir() if p.is_file() and p.suffix.lower() == ".pdf"],
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
            messagebox.showwarning("PDF_Rename_Expert", "Оберіть PDF-файл у списку.")
            return
        if not pdf.exists():
            messagebox.showerror("PDF_Rename_Expert", "Файл не знайдено.")
            return
        try:
            os.startfile(str(pdf))  # type: ignore[attr-defined]
        except Exception as e:
            messagebox.showerror("PDF_Rename_Expert", f"Не вдалося відкрити PDF:\n{e}")

    def _open_source_folder(self) -> None:
        folder = self._source_dir()
        try:
            folder.mkdir(parents=True, exist_ok=True)
        except Exception:
            pass
        try:
            os.startfile(str(folder))  # type: ignore[attr-defined]
        except Exception as e:
            messagebox.showerror("PDF_Rename_Expert", f"Не вдалося відкрити папку:\n{e}")

    # ---------------- Preview / naming ----------------
    def _get_date_str(self) -> str:
        if DateEntry is not None:
            d = self.date_entry.get_date()
            return d.isoformat()
        raw = str(self.date_entry.get()).strip()
        return raw

    def _phase_value(self) -> Tuple[str, Optional[str]]:
        """Return (value, error_or_None)."""
        p = self.var_phase.get().strip()
        if p == "3Ф-Ктт-XXX":
            raw = self.var_phase_ktt.get().strip()
            xxx = sanitize_component(raw, spaces_to_dash=True)
            if not xxx or xxx.upper() == "XXX":
                return "3Ф-Ктт-XXX", "Заповніть коефіцієнт для 3Ф-Ктт (замість XXX)."
            return f"3Ф-Ктт-{xxx}", None
        return (p or "3Ф"), None

    def _work_value(self) -> str:
        wt = self.var_work_type.get().strip()
        if wt == "Інше":
            custom = sanitize_component(self.var_work_custom.get(), spaces_to_dash=True)
            return f"ІН-{custom}" if custom else "ІН"
        return WORK_CODES.get(wt, "ВР")

    def _check_value(self) -> str:
        ct = self.var_check_type.get().strip()
        return CHECK_CODES.get(ct, "ВП")

    def _build_final_filename(self) -> Tuple[str, List[str]]:
        errors: List[str] = []
        pdf = self._selected_pdf_path()
        scan_no = extract_scan_number(pdf.name) if pdf else "0000"

        eic_full = sanitize_component(self.var_eic_full.get(), spaces_to_dash=False).replace(" ", "")
        if not eic_full:
            suffix = self.var_eic_suffix.get().strip()
            if len(suffix) >= 5 and self.matches:
                errors.append("Оберіть запис зі списку або введіть повний EIC.")
            else:
                errors.append("Введіть повний EIC (16 символів) або виконайте пошук по базі.")
        elif len(eic_full) != 16:
            errors.append("EIC має містити 16 символів.")

        dt = self._get_date_str()
        if not re.fullmatch(r"\d{4}-\d{2}-\d{2}", dt or ""):
            errors.append("Некоректна дата (очікується YYYY-MM-DD).")
        else:
            try:
                if date.fromisoformat(dt) > date.today():
                    errors.append("Дата в майбутньому — перевірте календар.")
            except Exception:
                pass

        meter = sanitize_component(self.var_meter_no.get(), spaces_to_dash=True)
        if not meter:
            errors.append("Заповніть № Лічильника.")

        name = sanitize_component(self.var_name.get(), spaces_to_dash=True)
        if not name:
            errors.append("Не заповнена Назва.")
        name = truncate_part(name, MAX_NAME_PART_LEN)

        addr = sanitize_component(self.var_address.get(), spaces_to_dash=True)
        if not addr:
            errors.append("Не заповнена Адреса.")
        addr = truncate_part(addr, MAX_ADDR_PART_LEN)

        invest = "Інвест" if self.var_invest.get() else ""
        phase_val, phase_err = self._phase_value()
        if phase_err:
            errors.append(phase_err)
        phase = sanitize_component(phase_val, spaces_to_dash=True)
        check = sanitize_component(self._check_value(), spaces_to_dash=True)
        work = sanitize_component(self._work_value(), spaces_to_dash=True)

        parts = [scan_no, eic_full, dt, meter, name, addr]
        if invest:
            parts.append(invest)
        parts.extend([phase, check, work])

        if "АСКОЕ" in self.var_work_type.get():
            ip_val = sanitize_component(self.var_ip.get(), spaces_to_dash=True)
            mod_val = sanitize_component(self.var_modem.get(), spaces_to_dash=True)
            if ip_val:
                parts.append(f"IP-{ip_val}")
            if mod_val:
                parts.append(f"Mod-{mod_val}")

        final = "_".join([p for p in parts if p])
        final = sanitize_component(final, spaces_to_dash=False)
        final = final.replace(" ", "-")
        final = re.sub(r"-{2,}", "-", final).strip("-")

        name_suffix = sanitize_component(self.var_name_suffix.get(), spaces_to_dash=True)
        if name_suffix:
            final = f"{final}_{name_suffix}" if final else name_suffix
            final = re.sub(r"-{2,}", "-", final).strip("-")

        if len(final) + 4 > MAX_FILENAME_LEN:
            allowed = MAX_FILENAME_LEN - 4
            final = final[:allowed].rstrip(" -._")
        return f"{final}.pdf", errors

    def _update_preview(self) -> None:
        fn, errs = self._build_final_filename()
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
        s = sanitize_component(raw, spaces_to_dash=True)
        if not s:
            return
        rest = [x for x in self._suffix_history if x != s]
        self._suffix_history = ([s] + rest)[:MAX_SUFFIX_HISTORY_ITEMS]
        save_suffix_history(self.app_dir, self._suffix_history)
        try:
            self.cmb_name_suffix["values"] = tuple(self._suffix_history)
        except Exception:
            pass

    # ---------------- Process ----------------
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
            messagebox.showwarning("PDF_Rename_Expert", "Оберіть PDF-файл у списку.")
            return

        final_name, errs = self._build_final_filename()
        if errs:
            messagebox.showwarning("PDF_Rename_Expert", "Не можна обробити:\n- " + "\n- ".join(errs))
            return

        # Rename in-place inside Scan\<Джерело>\ (no archive folders).
        dst_path = ensure_unique_path(pdf.parent / final_name)

        try:
            os.replace(str(pdf), str(dst_path))
        except PermissionError:
            messagebox.showerror("PDF_Rename_Expert", "PDF-файл заблокований іншою програмою.")
            return
        except Exception as e:
            logging.exception("Rename failed")
            messagebox.showerror("PDF_Rename_Expert", f"Помилка перейменування:\n{e}")
            return

        self._set_status(f"Перейменовано: {dst_path.name}", "success")
        self._remember_suffix(self.var_name_suffix.get())
        self._refresh_file_list()
        self._clear_for_next()
        self._clear_pdf_preview()

    def _set_status(self, text: str, level: str = "info") -> None:
        """level: info | success | warning | error"""
        self.var_status.set(text)
        color_map = {
            "info": PALETTE["muted"],
            "success": PALETTE["success"],
            "warning": PALETTE["warning"],
            "error": PALETTE["danger"],
        }
        color = color_map.get(level, PALETTE["muted"])
        try:
            self.lbl_status_dot.configure(foreground=color)
        except Exception:
            pass

    def _load_db_into_memory(self) -> None:
        # Async load so UI doesn't freeze.
        if self._db_load_thread is not None and self._db_load_thread.is_alive():
            return

        self.btn_load_db.configure(state="disabled")
        self.pb_db.grid()
        self.pb_db.start(10)
        self._set_status("Завантаження бази в пам'ять...", "info")

        def worker() -> None:
            try:
                self.db.load()
                n = len(self.db._eic_col or [])
                self._db_load_queue.put(("ok", n))
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

        if kind == "ok":
            n = int(payload)
            self._set_status(f"База завантажена в пам'ять. Записів: {n}", "success")
            messagebox.showinfo("PDF_Rename_Expert", f"База успішно завантажена.\nЗаписів: {n}")
        elif kind == "warn":
            self._set_status(str(payload), "warning")
            messagebox.showwarning("PDF_Rename_Expert", str(payload))
        else:
            self._set_status(str(payload), "error")
            messagebox.showerror("PDF_Rename_Expert", str(payload))

    # ---------------- Monthly export ----------------
    def _export_monthly_excel(self) -> None:
        # Export is built ONLY from Scan folder contents by parsing renamed filenames.
        # Now exporting ALL files regardless of the selected date.
        
        cols = EXPORT_COLUMNS
        report_name = f"Звіт_всі_дані_{date.today().isoformat()}.xlsx"
        report_path = self.app_dir / report_name

        if report_path.exists():
            if not messagebox.askyesno(
                "PDF_Rename_Expert",
                f"Звіт уже існує:\n{report_path.name}\n\nПерезаписати?",
            ):
                return

        src_base = self.app_dir / "Scan"
        updated: Dict[str, pd.DataFrame] = {}

        any_rows = False
        for sheet in ["Населення", "Промислові"]:
            folder = src_base / sheet
            if not folder.exists():
                try:
                    folder.mkdir(parents=True, exist_ok=True)
                except Exception:
                    pass

            rows: List[dict] = []
            if folder.exists():
                for p in sorted(folder.iterdir()):
                    if not (p.is_file() and p.suffix.lower() == ".pdf"):
                        continue
                    meta = parse_renamed_pdf_filename(p.name)
                    if meta:
                        rows.append(meta)

            df = pd.DataFrame(rows)
            if not df.empty:
                # Add numbering starting from 1
                df.insert(0, "№", range(1, len(df) + 1))
                
                # Ensure all columns exist and are in correct order
                for c in cols:
                    if c not in df.columns:
                        df[c] = ""
                df = df[cols]
                any_rows = True
            updated[sheet] = df if not df.empty else pd.DataFrame(columns=cols)

        if not any_rows:
            messagebox.showinfo(
                "PDF_Rename_Expert", "Немає оброблених файлів у папці Scan для експорту."
            )
            return

        try:
            with pd.ExcelWriter(report_path, engine="openpyxl", mode="w") as w:
                for sheet, df in updated.items():
                    df.to_excel(w, sheet_name=sheet, index=False)
                    
                    # Auto-adjust column width
                    ws = w.sheets[sheet]
                    for idx, col in enumerate(df.columns):
                        # Find max length in column
                        series = df[col]
                        max_len = max(
                            series.astype(str).map(len).max(),
                            len(str(col))
                        ) + 2 # Add some padding
                        
                        # Convert 0-index to Excel column letter
                        col_letter = chr(65 + idx)
                        if idx >= 26: # Handle AA, AB... if needed
                            col_letter = chr(64 + idx // 26) + chr(65 + idx % 26)
                        
                        ws.column_dimensions[col_letter].width = min(max_len, 60) # Cap at 60
        except PermissionError:
            messagebox.showerror(
                "PDF_Rename_Expert", "Excel-звіт зайнятий іншим процесом (закрийте файл)."
            )
            return
        except Exception as e:
            logging.exception("Excel export failed")
            messagebox.showerror("PDF_Rename_Expert", f"Не вдалося записати Excel-звіт:\n{e}")
            return

        self._set_status(f"Експорт завершено: {report_path.name}", "success")
        summary = (
            f"Звіт: {report_path.name}\n"
            f"Населення: {len(updated['Населення'])}\n"
            f"Промислові: {len(updated['Промислові'])}"
        )
        messagebox.showinfo("PDF_Rename_Expert", f"Готово.\n\n{summary}")


def _setup_logging(app_dir: Path) -> None:
    try:
        log_file = str(app_dir / "pdf_rename_expert.log")
        file_handler = logging.FileHandler(log_file, encoding="utf-8")
        file_handler.setFormatter(logging.Formatter("%(asctime)s %(levelname)s %(message)s"))
        
        root_logger = logging.getLogger()
        root_logger.setLevel(logging.INFO)
        if not root_logger.handlers:
            root_logger.addHandler(file_handler)
    except Exception:
        pass


def main() -> int:
    # For one-file PyInstaller builds, Excel/Scan live next to the .exe,
    # not inside the temporary _MEIPASS extraction directory.
    if getattr(sys, "frozen", False):
        app_dir = Path(sys.executable).resolve().parent
    else:
        app_dir = Path(__file__).resolve().parent
    try:
        os.chdir(str(app_dir))
    except Exception:
        # If we can't change cwd, we still use absolute paths via app_dir.
        pass

    _setup_logging(app_dir)
    logging.info("Start PDF_Rename_Expert from %s", app_dir)

    root = tk.Tk()
    try:
        _setup_theme(root)
    except Exception:
        logging.exception("Theme setup failed")
    App(root, app_dir=app_dir)
    root.mainloop()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

