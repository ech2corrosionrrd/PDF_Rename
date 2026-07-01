"""Генерує мінімальний Excel-файл аркуша «база» для інтеграційних тестів."""

from __future__ import annotations

from pathlib import Path

from excel_db import EXCEL_SHEET_NAME

# Індекси стовпців (0-based), як у ExcelConsumerDB
_COL_B = 1
_COL_T = 19
_COL_AB = 27
_COL_AD = 29
_COL_AN = 39


def _set(ws, row: int, col_idx: int, value: str) -> None:
    ws.cell(row=row, column=col_idx + 1, value=value)


def write_minimal_baza(xlsx_path: Path) -> None:
    """Записує xlsx з двома рядками EIC і одним рядком техобліку."""
    import openpyxl

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = EXCEL_SHEET_NAME

    # Рядок 2: звичайний споживач (EIC закінчується на 01234)
    _set(ws, 2, _COL_B, "")
    _set(ws, 2, _COL_T, "ТОВ")
    _set(ws, 2, _COL_AB, "UA12345678901234")
    _set(ws, 2, _COL_AD, "Компанія")
    _set(ws, 2, _COL_AN, "вул. Тестова 1")

    # Рядок 3: другий EIC з тим самим suffix5 (для перевірки уточнення суфікса)
    _set(ws, 3, _COL_B, "")
    _set(ws, 3, _COL_T, "ПП")
    _set(ws, 3, _COL_AB, "UA9999901234")
    _set(ws, 3, _COL_AD, "Фірма")
    _set(ws, 3, _COL_AN, "вул. Інша 2")

    # Рядок 4: технічний облік
    _set(ws, 4, _COL_B, "Технічний")
    _set(ws, 4, _COL_T, "T-999")
    _set(ws, 4, _COL_AB, "")
    _set(ws, 4, _COL_AD, "Об'єкт")
    _set(ws, 4, _COL_AN, "Адреса тех")

    xlsx_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(xlsx_path)
