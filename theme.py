"""Стилізація ttk: палітра кольорів та налаштування `ttk.Style`."""

from __future__ import annotations

import tkinter as tk
from tkinter import font as tkfont
from tkinter import ttk

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
    border = PALETTE["border"]
    text = PALETTE["text"]
    muted = PALETTE["muted"]
    accent = PALETTE["accent"]

    root.configure(background=bg)

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

    style.configure(
        "Card.TLabelframe",
        background=surface, bordercolor=border, lightcolor=border, darkcolor=border,
        relief="solid", borderwidth=1, padding=10,
    )
    style.configure(
        "Card.TLabelframe.Label",
        background=surface, foreground=text,
        font=(base_family, 10, "bold"), padding=(4, 0),
    )
    style.configure("Card.TFrame", background=surface)
    style.configure("Card.TLabel", background=surface, foreground=text)
    style.configure("CardMuted.TLabel", background=surface, foreground=muted)

    style.configure(
        "TEntry",
        fieldbackground=surface, background=surface, foreground=text,
        bordercolor=border, lightcolor=border, darkcolor=border,
        insertcolor=text, padding=4,
    )
    style.map(
        "TEntry",
        bordercolor=[("focus", accent)],
        lightcolor=[("focus", accent)],
        darkcolor=[("focus", accent)],
    )
    style.configure(
        "TCombobox",
        fieldbackground=surface, background=surface, foreground=text,
        bordercolor=border, arrowcolor=text, padding=4,
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

    style.configure(
        "TButton",
        background=surface, foreground=text,
        bordercolor=border, lightcolor=border, darkcolor=border,
        focusthickness=0, padding=(12, 6),
    )
    style.map(
        "TButton",
        background=[("active", PALETTE["surface_alt"]), ("pressed", PALETTE["surface_alt"])],
        bordercolor=[("active", accent), ("focus", accent)],
    )
    style.configure(
        "Accent.TButton",
        background=accent, foreground=PALETTE["accent_fg"],
        bordercolor=accent, lightcolor=accent, darkcolor=accent,
        focusthickness=0, padding=(16, 8),
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
        background=accent, troughcolor=PALETTE["surface_alt"],
        bordercolor=border, lightcolor=accent, darkcolor=accent,
    )

    style.configure("TSeparator", background=border)
    style.configure(
        "Vertical.TScrollbar",
        background=PALETTE["surface_alt"], troughcolor=bg,
        bordercolor=bg, arrowcolor=muted,
    )
    style.configure(
        "Horizontal.TScrollbar",
        background=PALETTE["surface_alt"], troughcolor=bg,
        bordercolor=bg, arrowcolor=muted,
    )

    style.configure("Status.TFrame", background=PALETTE["surface_alt"])
    style.configure(
        "Status.TLabel",
        background=PALETTE["surface_alt"], foreground=text, padding=(10, 6),
    )
    style.configure(
        "StatusDot.TLabel",
        background=PALETTE["surface_alt"], foreground=PALETTE["success"],
        font=(base_family, 12, "bold"), padding=(10, 4, 0, 4),
    )

    return style
