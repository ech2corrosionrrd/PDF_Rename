"""Збереження користувацьких налаштувань у %APPDATA%\\PDF_Rename_Expert\\."""

from __future__ import annotations

import json
import os
from pathlib import Path
from typing import Any, Dict

CONFIG_DIR_NAME = "PDF_Rename_Expert"
SETTINGS_FILENAME = "settings.json"


def user_config_dir() -> Path:
    base = os.environ.get("APPDATA") or str(Path.home())
    d = Path(base) / CONFIG_DIR_NAME
    d.mkdir(parents=True, exist_ok=True)
    return d


def settings_file_path() -> Path:
    return user_config_dir() / SETTINGS_FILENAME


def load_settings() -> Dict[str, Any]:
    path = settings_file_path()
    if not path.is_file():
        return {}
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
        return data if isinstance(data, dict) else {}
    except Exception:
        return {}


def save_settings(data: Dict[str, Any]) -> None:
    path = settings_file_path()
    tmp = path.with_suffix(".tmp")
    text = json.dumps(data, ensure_ascii=False, indent=2)
    tmp.write_text(text, encoding="utf-8")
    # os.replace атомарно перезаписує на Windows (у т.ч. Win7); pathlib.replace там часто падає.
    os.replace(str(tmp), str(path))


def get_str(settings: Dict[str, Any], key: str, default: str = "") -> str:
    v = settings.get(key)
    if v is None:
        return default
    return str(v).strip()
