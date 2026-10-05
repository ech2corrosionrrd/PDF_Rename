"""Вікно «Лічильники по підстанціям»: аналог форми `frmBase` з книги Excel.

Усі поля будуються за заголовками аркуша «база». Читання — з файлу (openpyxl),
запис — через Excel (`meters_excel`), тому макроси й форматування книги не страждають.
"""

from __future__ import annotations

import logging
import os
import threading
from datetime import datetime
from pathlib import Path
from queue import Empty, Queue
from typing import Any, Callable, Dict, List, Optional

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import meters_excel as mx
from meters_model import (
    ACT_HEADERS,
    ALL_SUBSTATIONS,
    COL_CONN,
    Group,
    LAST_COL,
    Navigator,
    Snapshot,
    build_card,
    build_layout,
    build_value_lists,
    card_pdf_name,
    changed_fields,
    col_letter,
    list_for,
    parse_id_query,
    roles_by_col,
)
from theme import PALETTE
from user_settings import get_str, load_settings, save_settings

WINDOW_TITLE = "Лічильники по підстанціям"
SETTINGS_KEY = "meters_workbook"
DIRTY_BG = "#fff3b0"
LABEL_W = 34


class MetersWindow(tk.Toplevel):
    def __init__(self, master: tk.Misc, app_dir: Path) -> None:
        super().__init__(master)
        self.app_dir = app_dir
        self.title(WINDOW_TITLE)
        self.geometry("1040x720")
        self.minsize(820, 520)
        self.configure(background=PALETTE["bg"])

        self.path: Optional[Path] = None
        self.snap: Optional[Snapshot] = None
        self.nav: Optional[Navigator] = None
        self.layout: List[Group] = []
        self.roles: Dict[int, str] = {}
        self._vars: Dict[int, tk.StringVar] = {}
        self._widgets: Dict[int, ttk.Widget] = {}
        self._dirty: set = set()
        self._loading = False
        self._busy = False
        self._built = False
        self._after_load_msg = ""

        self.var_subst = tk.StringVar(value=ALL_SUBSTATIONS)
        self.var_find = tk.StringVar()
        self.var_id = tk.StringVar()
        self.var_info = tk.StringVar(value="Завантаження…")
        self.var_mark = tk.StringVar()
        self.var_status = tk.StringVar()

        style = ttk.Style(self)
        style.configure("Dirty.TEntry", fieldbackground=DIRTY_BG)
        style.configure("Dirty.TCombobox", fieldbackground=DIRTY_BG)
        style.map("Dirty.TCombobox", fieldbackground=[("readonly", DIRTY_BG)])

        self._build_chrome()
        self.protocol("WM_DELETE_WINDOW", self._on_close)
        self.bind("<Control-s>", lambda _e: self._save())
        self.bind("<Escape>", lambda _e: self._on_close())

        self.after(50, self._resolve_path_and_load)

    # ------------------------------------------------------------ каркас
    def _build_chrome(self) -> None:
        self.columnconfigure(0, weight=1)
        self.rowconfigure(1, weight=1)

        top = ttk.Frame(self, padding=(10, 8, 10, 2))
        top.grid(row=0, column=0, sticky="ew")
        top.columnconfigure(4, weight=1)

        ttk.Label(top, text="Підстанція:").grid(row=0, column=0, sticky="w")
        self.cbo_subst = ttk.Combobox(
            top, textvariable=self.var_subst, state="readonly", width=34
        )
        self.cbo_subst.grid(row=0, column=1, sticky="w", padx=(6, 14))
        self.cbo_subst.bind("<<ComboboxSelected>>", lambda _e: self._apply_filters())

        ttk.Label(top, text="Пошук:").grid(row=0, column=2, sticky="w")
        ent_find = ttk.Entry(top, textvariable=self.var_find, width=24)
        ent_find.grid(row=0, column=3, sticky="w", padx=(6, 6))
        ent_find.bind("<Return>", lambda _e: self._apply_filters())
        btns = ttk.Frame(top)
        btns.grid(row=0, column=4, sticky="w")
        ttk.Button(btns, text="Знайти", command=self._apply_filters).pack(side="left")
        ttk.Button(btns, text="Скинути", command=self._reset_filters).pack(
            side="left", padx=4
        )
        ttk.Button(btns, text="Оновити з файлу", command=lambda: self._reload()).pack(
            side="left"
        )
        ttk.Button(btns, text="Файл…", command=self._choose_file).pack(
            side="left", padx=(4, 0)
        )

        ttk.Label(top, text="EIC / № лічильника:").grid(
            row=1, column=0, columnspan=2, sticky="w", pady=(6, 0)
        )
        row2 = ttk.Frame(top)
        row2.grid(row=1, column=1, columnspan=4, sticky="w", pady=(6, 0), padx=(130, 0))
        ent_id = ttk.Entry(row2, textvariable=self.var_id, width=34)
        ent_id.grid(row=0, column=0)
        ent_id.bind("<Return>", lambda _e: self._apply_id_filter())
        ttk.Button(row2, text="Знайти", command=self._apply_id_filter).grid(
            row=0, column=1, padx=6
        )
        ttk.Label(
            row2,
            text="останні 5+ символів; кілька значень — через кому",
            style="Muted.TLabel",
        ).grid(row=0, column=2)

        nav = ttk.Frame(top)
        nav.grid(row=2, column=0, columnspan=5, sticky="ew", pady=(8, 2))
        for text, where in (("|<", "first"), ("<", "prev"), (">", "next"), (">|", "last")):
            ttk.Button(
                nav, text=text, width=4, command=lambda w=where: self._move(w)
            ).pack(side="left", padx=(0, 2))
        ttk.Label(nav, textvariable=self.var_info, font=("Segoe UI", 10, "bold")).pack(
            side="left", padx=12
        )
        tk.Label(
            nav,
            textvariable=self.var_mark,
            fg=PALETTE["danger"],
            bg=PALETTE["bg"],
            font=("Segoe UI", 10, "bold"),
        ).pack(side="left", padx=8)

        # прокручувана область полів
        body = ttk.Frame(self)
        body.grid(row=1, column=0, sticky="nsew", padx=10, pady=4)
        body.columnconfigure(0, weight=1)
        body.rowconfigure(0, weight=1)
        self.canvas = tk.Canvas(
            body, background=PALETTE["surface"], highlightthickness=1,
            highlightbackground=PALETTE["border"],
        )
        vsb = ttk.Scrollbar(body, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=vsb.set)
        self.canvas.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        self.inner = tk.Frame(self.canvas, background=PALETTE["surface"])
        self._inner_id = self.canvas.create_window((0, 0), window=self.inner, anchor="nw")
        self.inner.bind(
            "<Configure>",
            lambda _e: self.canvas.configure(scrollregion=self.canvas.bbox("all")),
        )
        self.canvas.bind(
            "<Configure>", lambda e: self.canvas.itemconfigure(self._inner_id, width=e.width)
        )
        self.canvas.bind("<Enter>", lambda _e: self.canvas.bind_all("<MouseWheel>", self._wheel))
        self.canvas.bind("<Leave>", lambda _e: self.canvas.unbind_all("<MouseWheel>"))

        bottom = ttk.Frame(self, padding=(10, 4, 10, 8))
        bottom.grid(row=2, column=0, sticky="ew")
        self.btn_save = ttk.Button(bottom, text="Зберегти запис (Ctrl+S)", command=self._save)
        self.btn_new = ttk.Button(bottom, text="Додати рядок", command=self._new)
        self.btn_del = ttk.Button(bottom, text="Видалити", command=self._delete)
        self.btn_print = ttk.Button(bottom, text="Друк картки (PDF)", command=lambda: self._print(False))
        self.btn_print_all = ttk.Button(
            bottom, text="Друк підстанції (PDF)", command=lambda: self._print(True)
        )
        self.btn_hist = ttk.Button(bottom, text="Історія змін", command=self._show_history)
        self.btn_acts = ttk.Button(bottom, text="Журнал актів", command=self._show_acts)
        for b in (self.btn_save, self.btn_new, self.btn_del, self.btn_print,
                  self.btn_print_all, self.btn_hist, self.btn_acts):
            b.pack(side="left", padx=(0, 6))
        ttk.Button(bottom, text="Закрити", command=self._on_close).pack(side="right")
        ttk.Label(self, textvariable=self.var_status, style="Muted.TLabel",
                  padding=(12, 0, 12, 6)).grid(row=3, column=0, sticky="ew")

    def _wheel(self, event: tk.Event) -> None:
        self.canvas.yview_scroll(int(-event.delta / 120), "units")

    # ------------------------------------------------------------ завантаження
    def _resolve_path_and_load(self) -> None:
        saved = get_str(load_settings(), SETTINGS_KEY)
        cand = Path(saved) if saved else None
        if cand is None or not cand.is_file():
            cand = mx.find_workbook(self.app_dir)
        if cand is None:
            self._choose_file(initial=True)
            return
        self.path = cand
        self._reload()

    def _choose_file(self, initial: bool = False) -> None:
        if not initial and not self._confirm_leave():
            return
        p = filedialog.askopenfilename(
            parent=self,
            title="Книга «Лічильники по підстанціям»",
            filetypes=[("Excel з макросами", "*.xlsm"), ("Усі файли", "*.*")],
        )
        if not p:
            if initial:
                self.var_info.set("Книгу не вибрано")
            return
        self.path = Path(p)
        data = load_settings()
        data[SETTINGS_KEY] = str(self.path)
        save_settings(data)
        self._reload()

    def _reload(self, keep_row: Optional[int] = None) -> None:
        if self.path is None:
            return
        path = self.path
        keep = keep_row if keep_row is not None else (self.nav.current if self.nav else None)
        self._run_bg(
            lambda: mx.load_snapshot(path),
            lambda snap: self._on_loaded(snap, keep),
            "Читання книги…",
        )

    def _on_loaded(self, snap: Snapshot, keep_row: Optional[int]) -> None:
        self.snap = snap
        if self.nav is None:
            self.nav = Navigator(snap)
        else:
            self.nav.set_snapshot(snap)
        if not self._built:
            self.layout = build_layout(snap)
            self.roles = roles_by_col(self.layout)
            self._build_fields()
            self._built = True
        self._refresh_lists()
        self.cbo_subst.configure(values=[ALL_SUBSTATIONS] + snap.substations())
        if self.var_subst.get() not in self.cbo_subst.cget("values"):
            self.var_subst.set(ALL_SUBSTATIONS)
        if keep_row is not None:
            self.nav.goto(keep_row)
        self._show_current()
        msg, self._after_load_msg = self._after_load_msg, ""
        self._status(msg or "Книгу прочитано: записів {}.".format(len(snap.values)))

    # ------------------------------------------------------------ побудова полів
    def _build_fields(self) -> None:
        for w in self.inner.winfo_children():
            w.destroy()
        self._vars.clear()
        self._widgets.clear()
        for gi, g in enumerate(self.layout):
            tk.Label(
                self.inner, text=g.name, anchor="w", bg="#ddebf7", fg=PALETTE["text"],
                font=("Segoe UI", 10, "bold"), padx=6, pady=2,
            ).grid(row=gi * 2, column=0, sticky="ew", padx=6, pady=(8, 2))
            frame = tk.Frame(self.inner, background=PALETTE["surface"])
            frame.grid(row=gi * 2 + 1, column=0, sticky="ew", padx=6)
            self.inner.columnconfigure(0, weight=1)
            n = len(g.blocks)
            per_col = max(1, (n + 1) // 2)
            for k, blk in enumerate(g.blocks):
                if g.column_major:
                    gr, gc = k % per_col, k // per_col
                else:
                    gr, gc = k // 2, k % 2
                cell = tk.Frame(frame, background=PALETTE["surface"])
                cell.grid(row=gr, column=gc, sticky="w", padx=(0, 18), pady=1)
                tip = "  ".join(
                    "{} [стовпець {}]".format(cap, col_letter(c))
                    for cap, c in zip(blk.captions, blk.cols)
                )
                lb = tk.Label(
                    cell, text=blk.label, width=LABEL_W, anchor="w",
                    bg=PALETTE["surface"], fg=PALETTE["text"],
                )
                lb.grid(row=0, column=0)
                width = {1: 26, 2: 11, 3: 7}[len(blk.cols)]
                for j, c in enumerate(blk.cols):
                    var = tk.StringVar()
                    self._vars[c] = var
                    w = self._make_widget(cell, c, var, width)
                    w.grid(row=0, column=j + 1, padx=(0, 3))
                    self._widgets[c] = w
                    var.trace_add("write", lambda *_a, col=c: self._on_edit(col))
        for c in range(1, LAST_COL + 1):
            if c not in self._vars:  # захист від втрати стовпців
                self._vars[c] = tk.StringVar()

    def _make_widget(self, parent: tk.Misc, c: int, var: tk.StringVar, width: int) -> ttk.Widget:
        assert self.snap is not None
        values = list_for(c, self.roles, self.snap.ref_lists, build_value_lists(self.snap))
        if c == COL_CONN or values is not None:
            cb = ttk.Combobox(parent, textvariable=var, width=width, values=values or [])
            if c == COL_CONN:
                cb.bind("<<ComboboxSelected>>", lambda _e: self._on_conn_selected())
            return cb
        return ttk.Entry(parent, textvariable=var, width=width)

    def _refresh_lists(self) -> None:
        """Оновлює значення випадаючих списків після перечитування книги."""
        assert self.snap is not None
        auto = build_value_lists(self.snap)
        for c, w in self._widgets.items():
            if c == COL_CONN or not isinstance(w, ttk.Combobox):
                continue
            vals = list_for(c, self.roles, self.snap.ref_lists, auto)
            w.configure(values=vals or [])

    # ------------------------------------------------------------ показ запису
    def _show_current(self) -> None:
        assert self.nav is not None and self.snap is not None
        r = self.nav.current
        self._loading = True
        try:
            for c, var in self._vars.items():
                var.set(self.snap.text(r, c) if r else "")
            for c, w in self._widgets.items():
                locked = r is None or self.snap.is_locked(r, c)
                state = "disabled" if locked else "normal"
                w.configure(state=state, style=self._style_for(w, False))
            conn = self._widgets.get(COL_CONN)
            if isinstance(conn, ttk.Combobox):
                subst = self.snap.subst_of(r) if r else ""
                conn.configure(values=self.snap.connections(subst) if r else [])
        finally:
            self._loading = False
        self._dirty.clear()
        self._show_dirty()
        n = len(self.nav.rows)
        if r is None:
            self.var_info.set("Немає записів за поточним фільтром")
        else:
            self.var_info.set(
                "Запис {} з {}     (рядок листа {})".format(self.nav.idx + 1, n, r)
            )
        self._set_buttons()

    @staticmethod
    def _style_for(w: ttk.Widget, dirty: bool) -> str:
        base = "TCombobox" if isinstance(w, ttk.Combobox) else "TEntry"
        return ("Dirty." + base) if dirty else base

    def _on_edit(self, c: int) -> None:
        if self._loading or self.nav is None or self.snap is None:
            return
        r = self.nav.current
        if r is None:
            return
        w = self._widgets.get(c)
        if w is None or self.snap.is_locked(r, c):
            return
        if self._vars[c].get() != self.snap.text(r, c):
            self._dirty.add(c)
        else:
            self._dirty.discard(c)
        w.configure(style=self._style_for(w, c in self._dirty))
        self._show_dirty()

    def _show_dirty(self) -> None:
        n = len(self._dirty)
        self.var_mark.set("ЗМІНЕНО, НЕ ЗБЕРЕЖЕНО (полів: {})".format(n) if n else "")

    def _set_buttons(self) -> None:
        state = "disabled" if self._busy else "normal"
        has = self.nav is not None and self.nav.current is not None
        for b in (self.btn_save, self.btn_del, self.btn_print, self.btn_print_all, self.btn_hist,
                  self.btn_acts):
            b.configure(state=state if (has or b in (self.btn_hist, self.btn_acts)) else "disabled")
        self.btn_new.configure(state=state)

    def _collect(self) -> Dict[int, str]:
        return {c: v.get().strip() for c, v in self._vars.items()}

    def _pending(self) -> Dict[int, str]:
        if self.nav is None or self.snap is None or self.nav.current is None:
            return {}
        return changed_fields(self.snap, self.nav.current, self._collect())

    # ------------------------------------------------------------ навігація й фільтри
    def _confirm_leave(self) -> bool:
        if not self._pending():
            return True
        ans = messagebox.askyesnocancel(
            WINDOW_TITLE, "У поточному записі є незбережені зміни. Зберегти їх?", parent=self
        )
        if ans is None:
            return False
        if ans:
            self._save()
            return False  # продовжити можна після завершення запису
        return True

    def _move(self, where: str) -> None:
        if self.nav is None or self._busy or not self._confirm_leave():
            return
        self.nav.move(where)
        self._show_current()

    def _apply_filters(self) -> None:
        if self.nav is None or self._busy or not self._confirm_leave():
            return
        self.nav.subst = self.var_subst.get()
        self.nav.text = self.var_find.get()
        self.nav.refresh()
        self._show_current()

    def _apply_id_filter(self) -> None:
        if self.nav is None or self._busy:
            return
        text = self.var_id.get()
        if not text.strip():
            if self.nav.id_tokens is None:
                return
            if not self._confirm_leave():
                return
            self.nav.id_tokens = None
            self.nav.refresh()
            self._show_current()
            self._status("Пошук за EIC / № лічильника скинуто.")
            return
        tokens, err = parse_id_query(text)
        if err:
            self._status(err)
            return
        if not tokens or not self._confirm_leave():
            return
        self.nav.id_tokens = tokens
        self.nav.refresh()
        self._show_current()
        if not self.nav.rows:
            self._status("За EIC / № лічильника нічого не знайдено.")
        else:
            self._status(
                "Знайдено записів: {} (перехід між ними — кнопки < >).".format(len(self.nav.rows))
            )

    def _reset_filters(self) -> None:
        if self.nav is None or self._busy or not self._confirm_leave():
            return
        self.var_subst.set(ALL_SUBSTATIONS)
        self.var_find.set("")
        self.var_id.set("")
        self.nav.reset_filters()
        self._show_current()

    def _on_conn_selected(self) -> None:
        if self.nav is None or self.snap is None or self._busy:
            return
        r = self.nav.current
        if r is None:
            return
        name = self._vars[COL_CONN].get().strip()
        stored = self.snap.text(r, COL_CONN).strip()
        self._loading = True
        self._vars[COL_CONN].set(self.snap.text(r, COL_CONN))
        self._loading = False
        if not name or name == stored:
            return
        target = self.nav.find_connection(self.snap.subst_of(r), name)
        if target is None or not self._confirm_leave():
            return
        self.nav.goto(target)
        self._show_current()

    # ------------------------------------------------------------ дії з записом
    def _save(self) -> None:
        if self.nav is None or self.snap is None or self.path is None or self._busy:
            return
        r = self.nav.current
        changes = self._pending()
        if r is None:
            return
        if not changes:
            self._status("Змін не було.")
            return
        path, snap = self.path, self.snap
        self._run_bg(
            lambda: mx.save_record(path, snap, r, changes),
            lambda n: self._after_write(
                r, "Збережено у рядку {}: змінено полів — {}.".format(r, n)
            ),
            "Збереження в Excel…",
        )

    def _new(self) -> None:
        if self.nav is None or self.snap is None or self.path is None or self._busy:
            return
        if not self._confirm_leave():
            return
        subst = self.var_subst.get()
        cur = self.nav.current
        if subst == ALL_SUBSTATIONS:
            subst = self.snap.subst_of(cur) if cur else ""
        path = self.path
        self._run_bg(
            lambda: mx.insert_record(path, subst),
            lambda row: self._after_write(
                row, "Додано рядок {}. Заповніть поля й збережіть запис.".format(row)
            ),
            "Додавання рядка…",
        )

    def _delete(self) -> None:
        if self.nav is None or self.snap is None or self.path is None or self._busy:
            return
        r = self.nav.current
        if r is None:
            return
        msg = (
            "Видалити рядок {}?\nПідстанція: {}\nПриєднання: {}\nЛічильник: {}\n\n"
            "Рядок буде видалено з аркуша."
        ).format(
            r,
            self.snap.subst_of(r),
            self.snap.text(r, COL_CONN),
            self.snap.text(r, 7),
        )
        if not messagebox.askyesno(WINDOW_TITLE, msg, icon="warning", default="no", parent=self):
            return
        path = self.path
        nxt = self.nav.rows[self.nav.idx + 1] if self.nav.idx + 1 < len(self.nav.rows) else None
        prv = self.nav.rows[self.nav.idx - 1] if self.nav.idx > 0 else None
        # після видалення рядки нижче зсуваються на 1
        keep = (nxt - 1) if nxt else prv
        self._run_bg(
            lambda: mx.delete_record(path, r),
            lambda _x: self._after_write(keep, "Рядок {} видалено.".format(r)),
            "Видалення рядка…",
        )

    def _after_write(self, row: Optional[int], message: str) -> None:
        self._after_load_msg = message
        self._reload(keep_row=row)

    def _show_history(self) -> None:
        self._show_log(
            mx.iter_history,
            "Історія змін",
            ("Дата Час", "Користувач", "Рядок", "Поле", "Старе", "Нове"),
            (140, 90, 60, 220, 180, 180),
        )

    def _show_acts(self) -> None:
        widths = (140, 90, 130, 90, 80, 90, 160, 160, 130, 100, 110, 110, 220, 60, 160)
        self._show_log(mx.iter_acts, "Журнал актів", tuple(ACT_HEADERS), widths)

    def _show_log(
        self,
        reader: Callable[[Path], Any],
        title: str,
        cols: tuple,
        widths: tuple,
    ) -> None:
        if self.path is None:
            return
        path = self.path
        self._run_bg(
            lambda: list(reader(path)),
            lambda rows: self._open_log_window(rows, title, cols, widths),
            "Читання журналу…",
        )

    def _open_log_window(self, rows: List[tuple], title: str, cols: tuple, widths: tuple) -> None:
        win = tk.Toplevel(self)
        win.title(title if rows else title + " (порожньо)")
        win.geometry("980x420")
        tree = ttk.Treeview(win, columns=cols, show="headings")
        for c, w in zip(cols, widths):
            tree.heading(c, text=c)
            tree.column(c, width=w, anchor="w")
        sb = ttk.Scrollbar(win, orient="vertical", command=tree.yview)
        hb = ttk.Scrollbar(win, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=sb.set, xscrollcommand=hb.set)
        hb.pack(side="bottom", fill="x")
        tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        for row in reversed(rows):
            tree.insert("", "end", values=row)

    # ------------------------------------------------------------ друк
    def _print(self, whole_substation: bool) -> None:
        if self.nav is None or self.snap is None or self.path is None or self._busy:
            return
        r = self.nav.current
        if r is None:
            return
        snap, layout, path = self.snap, self.layout, self.path
        subst = snap.subst_of(r)
        if whole_substation:
            rows = [x for x in snap.rows() if snap.subst_of(x) == subst]
            cards = [build_card(snap, layout, x) for x in rows]
            conn = ""
        else:
            cards = [build_card(snap, layout, r, self._collect())]
            conn = snap.text(r, COL_CONN)
        pdf = path.parent / card_pdf_name(subst, conn, datetime.now())
        self._run_bg(
            lambda: (mx.export_cards_pdf(cards, pdf), pdf)[1],
            self._after_print,
            "Створення PDF…",
        )

    def _after_print(self, pdf: Path) -> None:
        self._status("PDF створено: {}".format(pdf.name))
        if messagebox.askyesno(WINDOW_TITLE, "PDF створено:\n{}\n\nВідкрити?".format(pdf), parent=self):
            try:
                os.startfile(str(pdf))  # type: ignore[attr-defined]
            except Exception:
                logging.exception("open pdf")

    # ------------------------------------------------------------ службове
    def _status(self, text: str) -> None:
        if self.winfo_exists():
            self.var_status.set(text)

    def _run_bg(
        self, fn: Callable[[], Any], done: Callable[[Any], None], message: str
    ) -> None:
        self._busy = True
        self._set_buttons()
        self._status(message)
        q: "Queue[tuple]" = Queue()

        def worker() -> None:
            try:
                q.put(("ok", fn()))
            except mx.MetersError as e:
                q.put(("err", str(e)))
            except Exception as e:  # noqa: BLE001
                logging.exception("meters background task")
                q.put(("err", "Неочікувана помилка: {}".format(e)))

        threading.Thread(target=worker, daemon=True).start()

        def poll() -> None:
            if not self.winfo_exists():
                return
            try:
                kind, payload = q.get_nowait()
            except Empty:
                self.after(100, poll)
                return
            self._busy = False
            self._set_buttons()
            if kind == "ok":
                done(payload)
            else:
                self._status("Помилка.")
                messagebox.showerror(WINDOW_TITLE, str(payload), parent=self)

        self.after(100, poll)

    def _on_close(self) -> None:
        if self._busy:
            if not messagebox.askyesno(
                WINDOW_TITLE, "Триває операція з книгою. Закрити вікно все одно?", parent=self
            ):
                return
        elif self._pending():
            ans = messagebox.askyesnocancel(
                WINDOW_TITLE, "Є незбережені зміни. Закрити без збереження?", parent=self
            )
            if not ans:
                return
        self.destroy()
