"""Читання Excel-бази споживачів."""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Optional

import pandas as pd

EXCEL_SHEET_NAME = "база"


@dataclass
class ConsumerRecord:
    eic_full: str
    name: str
    address: str

    def display(self) -> str:
        return f"{self.eic_full} — {self.name} — {self.address}"


class ExcelConsumerDB:
    """
    Читає книгу Excel:
      - файл: поруч із програмою (будь-який *.xlsx, пріоритет очікуваній назві)
      - аркуш: «база»
      - стовпці за літерами: AB (EIC), T + AD (ім'я через дефіс), AN (адреса)
    """

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

        try:
            df = pd.read_excel(
                self.excel_path,
                sheet_name=EXCEL_SHEET_NAME,
                engine="openpyxl",
                usecols="A:AN",
                dtype=str,
            )
        except ValueError as e:
            raise ValueError(
                "Не вдалося відкрити аркуш «база» у Excel-файлі. "
                "Переконайтеся, що книга містить аркуш з такою назвою (як у шаблону)."
            ) from e
        self._df = df

        need_cols = self._IDX_AN + 1
        ncols = int(df.shape[1])
        if ncols < need_cols:
            raise ValueError(
                f"У таблиці замало стовпців: знайдено {ncols}, потрібно щонайменше {need_cols} "
                f"(до стовпця AN включно). Очікується структура: EIC у стовпці AB, "
                f"ім'я — T та AD, адреса — AN."
            )

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

    @property
    def record_count(self) -> int:
        """Кількість записів EIC у завантаженій базі (0, якщо ще не завантажено)."""
        return len(self._eic_col) if self._eic_col else 0

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
