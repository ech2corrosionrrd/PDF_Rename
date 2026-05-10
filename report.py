"""Формування Excel-звіту з імен PDF, що лежать у Scan/<розділ>/."""

from __future__ import annotations

from pathlib import Path
from typing import Dict, Iterable

import pandas as pd
from openpyxl.utils import get_column_letter

from naming import EXPORT_COLUMNS, parse_renamed_pdf_filename

DEFAULT_SHEETS: Iterable[str] = ("Населення", "Промислові")
COLUMN_WIDTH_CAP = 60


def build_report_frames(
    scan_dir: Path,
    sheets: Iterable[str] = DEFAULT_SHEETS,
) -> Dict[str, pd.DataFrame]:
    """Збирає таблиці звіту з PDF-файлів у `scan_dir/<sheet>/`.

    Гарантує однаковий набір та порядок стовпців `EXPORT_COLUMNS`. Порожні аркуші
    повертаються як `DataFrame(columns=EXPORT_COLUMNS)`.
    """
    frames: Dict[str, pd.DataFrame] = {}
    for sheet in sheets:
        folder = scan_dir / sheet
        rows = []
        if folder.exists():
            for p in sorted(folder.iterdir()):
                if not (p.is_file() and p.suffix.lower() == ".pdf"):
                    continue
                meta = parse_renamed_pdf_filename(p.name)
                if meta:
                    rows.append(meta)

        df = pd.DataFrame(rows)
        if not df.empty:
            df.insert(0, "№", range(1, len(df) + 1))
            for c in EXPORT_COLUMNS:
                if c not in df.columns:
                    df[c] = ""
            df = df[EXPORT_COLUMNS]
        else:
            df = pd.DataFrame(columns=EXPORT_COLUMNS)
        frames[sheet] = df
    return frames


def has_any_rows(frames: Dict[str, pd.DataFrame]) -> bool:
    return any(not df.empty for df in frames.values())


def write_report(report_path: Path, frames: Dict[str, pd.DataFrame]) -> None:
    """Зберігає звіт у Excel із автопідбором ширини стовпців.

    Може кинути `PermissionError` (файл відкритий у Excel) або інший виняток від pandas.
    """
    with pd.ExcelWriter(report_path, engine="openpyxl", mode="w") as w:
        for sheet, df in frames.items():
            df.to_excel(w, sheet_name=sheet, index=False)
            ws = w.sheets[sheet]
            for idx, col in enumerate(df.columns):
                series = df[col]
                if series.empty:
                    max_len = len(str(col)) + 2
                else:
                    smax = series.astype(str).map(len).max()
                    max_len = (
                        max(0 if pd.isna(smax) else int(smax), len(str(col))) + 2
                    )
                col_letter = get_column_letter(idx + 1)
                ws.column_dimensions[col_letter].width = min(max_len, COLUMN_WIDTH_CAP)
