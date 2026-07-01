"""Інструкція для вбудованої довідки (кінцевий користувач)."""

from __future__ import annotations

import sys
from pathlib import Path
from typing import Optional

# Довідка в програмі (F1) — лише для оператора, без Python/збірки.
APP_USER_MANUAL_FILENAME = "INSTRUKTSIYA_KORYSTUVACHA_APP.md"

# Повна інструкція (репозиторій, IT) — не показується в F1 за замовчуванням.
FULL_USER_MANUAL_FILENAME = "INSTRUKTSIYA_KORYSTUVACHA.md"


def _search_paths(app_dir: Path, filename: str) -> list[Path]:
    paths: list[Path] = [app_dir / filename]
    if getattr(sys, "frozen", False):
        paths.append(Path(sys._MEIPASS) / filename)  # type: ignore[attr-defined]
    paths.append(Path(__file__).resolve().parent / filename)
    return paths


def resolve_app_user_manual_path(app_dir: Path) -> Optional[Path]:
    """Файл довідки для вікна F1 (пріоритет — поруч із exe)."""
    for p in _search_paths(app_dir, APP_USER_MANUAL_FILENAME):
        if p.is_file():
            return p
    return None


def resolve_full_user_manual_path(app_dir: Path) -> Optional[Path]:
    """Повна інструкція (для IT), якщо є поруч із програмою."""
    for p in _search_paths(app_dir, FULL_USER_MANUAL_FILENAME):
        if p.is_file():
            return p
    return None


def load_app_user_manual_text(app_dir: Path) -> tuple[str, Optional[Path]]:
    """Текст довідки для вбудованого вікна (кінцевий користувач)."""
    path = resolve_app_user_manual_path(app_dir)
    if path is None:
        return (
            "Довідку не знайдено.\n\n"
            f"Очікується файл: {APP_USER_MANUAL_FILENAME}\n"
            "поруч із програмою або в комплекті поставки.",
            None,
        )
    try:
        return path.read_text(encoding="utf-8"), path
    except OSError as e:
        return (f"Не вдалося прочитати довідку:\n{e}", path)
