"""Збереження/читання історії значень поля «Належність лічильника»."""

from __future__ import annotations

import json
import logging
from pathlib import Path
from typing import List

SUFFIX_HISTORY_FILENAME = "pdf_rename_suffix_history.json"
MAX_SUFFIX_HISTORY_ITEMS = 40


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
    except Exception as e:
        logging.error("Failed to save suffix history: %s", e)
