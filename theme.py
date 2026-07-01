"""Стилізація ttk: палітра кольорів та налаштування `ttk.Style`."""

from __future__ import annotations

import tkinter as tk
from tkinter import font as tkfont
from tkinter import ttk
from typing import Any, Dict

PALETTE = {
    "bg":            "#f4f6fb",
    "surface":       "#ffffff",
    "surface_alt":   "#eef2f9",
    "border":        "#d7dde8",
    "text":          "#1e293b",
    "muted":         "#64748b",
    "accent":        "#2563eb",
    "accent_hover":  "#1d4ed8",
    "accent_active": "#1e40af",
    "accent_fg":     "#ffffff",
    "success":       "#16a34a",
    "warning":       "#d97706",
    "danger":        "#dc2626",
}

FONT_UI = ("Segoe UI", 10)
FONT_MONO = ("Consolas", 10)
FONT_MONO_SM = ("Consolas", 9)

PREVIEW_LEN_WARN_RATIO = 0.85


def preview_len_style(length: int, max_len: int = 200) -> str:
    """Стиль мітки довжини імені PDF залежно від ліміту."""
    if length > max_len:
        return "PreviewLenError.TLabel"
    if length >= int(max_len * PREVIEW_LEN_WARN_RATIO):
        return "PreviewLenWarn.TLabel"
    return "PreviewLenOk.TLabel"


def status_dot_style(level: str) -> str:
    """Стиль індикатора рядка стану (info | success | warning | error)."""
    key = level if level in ("info", "success", "warning", "error") else "info"
    return f"StatusDot{key.capitalize()}.TLabel"


def text_widget_options(**extra: Any) -> Dict[str, Any]:
    """Спільні параметри для `tk.Text` у додатку."""
    opts: Dict[str, Any] = {
        "background": PALETTE["surface"],
        "foreground": PALETTE["text"],
        "insertbackground": PALETTE["text"],
        "selectbackground": PALETTE["accent"],
        "selectforeground": PALETTE["accent_fg"],
        "relief": "flat",
        "borderwidth": 1,
        "highlightthickness": 1,
        "highlightbackground": PALETTE["border"],
        "highlightcolor": PALETTE["accent"],
        "padx": 8,
        "pady": 6,
    }
    opts.update(extra)
    return opts


def listbox_options(**extra: Any) -> Dict[str, Any]:
    """Спільні параметри для `tk.Listbox`."""
    opts: Dict[str, Any] = {
        "exportselection": False,
        "activestyle": "none",
        "borderwidth": 1,
        "highlightthickness": 1,
        "highlightbackground": PALETTE["border"],
        "highlightcolor": PALETTE["accent"],
        "background": PALETTE["surface"],
        "foreground": PALETTE["text"],
        "selectbackground": PALETTE["accent"],
        "selectforeground": PALETTE["accent_fg"],
        "font": FONT_UI,
        "relief": "flat",
    }
    opts.update(extra)
    return opts


def setup_theme(root: tk.Tk) -> ttk.Style:
    style = ttk.Style(root)
    try:
        style.theme_use("clam")
    except Exception:
        pass

    base_family = "Segoe UI"
    try:
        for name in ("TkDefaultFont", "TkTextFont"):
            tkfont.nametofont(name).configure(family=base_family, size=10)
        tkfont.nametofont("TkHeadingFont").configure(
            family=base_family, size=10, weight="bold"
        )
    except Exception:
        pass

    bg = PALETTE["bg"]
    surface = PALETTE["surface"]
    surface_alt = PALETTE["surface_alt"]
    border = PALETTE["border"]
    text = PALETTE["text"]
    muted = PALETTE["muted"]
    accent = PALETTE["accent"]

    root.configure(background=bg)

    style.configure(".", background=bg, foreground=text, font=FONT_UI)
    style.configure("TFrame", background=bg)
    style.configure("TLabel", background=bg, foreground=text)
    style.configure("Muted.TLabel", background=bg, foreground=muted)
    style.configure(
        "Header.TLabel",
        background=bg,
        foreground=text,
        font=(base_family, 16, "bold"),
    )
    style.configure(
        "SubHeader.TLabel",
        background=bg,
        foreground=muted,
        font=(base_family, 10),
    )
    style.configure(
        "Section.TLabel",
        background=bg,
        foreground=text,
        font=(base_family, 10, "bold"),
    )

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
    style.configure("Card.TFrame", background=surface)
    style.configure("Card.TLabel", background=surface, foreground=text)
    style.configure("CardMuted.TLabel", background=surface, foreground=muted)
    style.configure(
        "CardStrong.TLabel",
        background=surface,
        foreground=text,
        font=(base_family, 10, "bold"),
    )

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
        root.option_add("*TCombobox*Listbox.font", FONT_UI)
    except Exception:
        pass

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
        background=[
            ("active", PALETTE["surface_alt"]),
            ("pressed", PALETTE["surface_alt"]),
        ],
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
    style.configure(
        "Secondary.TButton",
        background=PALETTE["surface_alt"],
        foreground=text,
        bordercolor=border,
        lightcolor=border,
        darkcolor=border,
        focusthickness=0,
        padding=(12, 6),
    )
    style.map(
        "Secondary.TButton",
        background=[
            ("active", surface),
            ("pressed", surface),
        ],
        bordercolor=[("active", accent), ("focus", accent)],
    )

    style.configure("Toolbar.TFrame", background=surface_alt)
    style.configure("Toolbar.TLabel", background=surface_alt, foreground=text)
    style.configure(
        "ToolbarSection.TLabel",
        background=surface_alt,
        foreground=text,
        font=(base_family, 10, "bold"),
    )
    style.configure("Toolbar.TRadiobutton", background=surface_alt, foreground=text)
    style.map(
        "Toolbar.TRadiobutton",
        background=[("active", surface_alt)],
        foreground=[("disabled", muted)],
    )
    style.configure(
        "Toolbar.TButton",
        background=surface,
        foreground=text,
        bordercolor=border,
        lightcolor=border,
        darkcolor=border,
        focusthickness=0,
        padding=(12, 6),
    )
    style.map(
        "Toolbar.TButton",
        background=[
            ("active", PALETTE["surface_alt"]),
            ("pressed", PALETTE["surface_alt"]),
        ],
        bordercolor=[("active", accent), ("focus", accent)],
    )

    for name, color in (
        ("PreviewLenOk", PALETTE["success"]),
        ("PreviewLenWarn", PALETTE["warning"]),
        ("PreviewLenError", PALETTE["danger"]),
    ):
        style.configure(
            f"{name}.TLabel",
            background=surface,
            foreground=color,
            font=(base_family, 10, "bold"),
        )

    style.configure("Card.TRadiobutton", background=surface, foreground=text)
    style.map(
        "Card.TRadiobutton",
        background=[("active", surface)],
        foreground=[("disabled", muted)],
    )
    style.configure("Card.TCheckbutton", background=surface, foreground=text)
    style.map("Card.TCheckbutton", background=[("active", surface)])

    style.configure("TRadiobutton", background=bg, foreground=text)
    style.map("TRadiobutton", background=[("active", bg)])
    style.configure("TCheckbutton", background=bg, foreground=text)
    style.map("TCheckbutton", background=[("active", bg)])

    style.configure(
        "Horizontal.TProgressbar",
        background=accent,
        troughcolor=PALETTE["surface_alt"],
        bordercolor=border,
        lightcolor=accent,
        darkcolor=accent,
    )

    style.configure("TSeparator", background=border)
    style.configure(
        "Vertical.TScrollbar",
        background=PALETTE["surface_alt"],
        troughcolor=bg,
        bordercolor=bg,
        arrowcolor=muted,
    )
    style.map(
        "Vertical.TScrollbar",
        background=[("active", border)],
    )
    style.configure(
        "Horizontal.TScrollbar",
        background=PALETTE["surface_alt"],
        troughcolor=bg,
        bordercolor=bg,
        arrowcolor=muted,
    )
    style.map(
        "Horizontal.TScrollbar",
        background=[("active", border)],
    )

    style.configure("TPanedwindow", background=bg)
    style.configure(
        "TPanedwindow.Sash",
        sashthickness=6,
        background=border,
    )
    style.map("TPanedwindow.Sash", background=[("active", accent)])

    style.configure("Status.TFrame", background=surface_alt)
    style.configure(
        "Status.TLabel",
        background=surface_alt,
        foreground=text,
        padding=(10, 6),
    )
    for dot_name, dot_color in (
        ("StatusDotInfo", muted),
        ("StatusDotSuccess", PALETTE["success"]),
        ("StatusDotWarning", PALETTE["warning"]),
        ("StatusDotError", PALETTE["danger"]),
    ):
        style.configure(
            f"{dot_name}.TLabel",
            background=surface_alt,
            foreground=dot_color,
            font=(base_family, 12, "bold"),
            padding=(10, 4, 0, 4),
        )

    return style
