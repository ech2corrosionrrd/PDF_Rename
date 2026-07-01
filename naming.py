"""Чисті функції побудови та парсингу імен PDF.

Іменно-залежна логіка (санітизація, обрізка, унікальний шлях, парсинг) зібрана тут.
Tk- і Excel-залежна логіка винесена в `app_ui.py`, `report.py`, `file_builder.py`.

Для зворотної сумісності тут реекспортуються імена з `suffix_history`:
`SUFFIX_HISTORY_FILENAME`, `MAX_SUFFIX_HISTORY_ITEMS`, `load_suffix_history`,
`save_suffix_history`.
"""

from __future__ import annotations

import json
import re
from pathlib import Path
from typing import Dict, List, Optional

from suffix_history import (  # noqa: F401  (back-compat re-export)
    MAX_SUFFIX_HISTORY_ITEMS,
    SUFFIX_HISTORY_FILENAME,
    load_suffix_history,
    save_suffix_history,
)
from user_settings import user_config_dir

INVALID_WIN_CHARS_RE = re.compile(r'[\\/:*?"<>|]+')
EXTRA_BAD_CHARS_RE = re.compile(
    r"[,;()\[\]{}!@#$%^&+=~`'\u00ab\u00bb\u201c\u201d\u2018\u2019]+"
)
WS_RE = re.compile(r"\s+")
# Невидимі / bidi-символи, які часто «крадуться» з Excel або PDF
_ZERO_WIDTH_RE = re.compile(r"[\u200b-\u200f\ufeff\u202a-\u202e]+")

# keep total name+ext well below Windows MAX_PATH
MAX_FILENAME_LEN = 200
MAX_NAME_PART_LEN = 60
MAX_ADDR_PART_LEN = 80

_DEFAULT_WORK_CODES: Dict[str, str] = {
    "Монтаж АСКОЕ": "Монт-АСКОЕ",
    "Монтаж лічильника": "Монт-Ліч",
    "Монтаж АСКОЕ + монтаж лічильника": "АСКОЕ-Ліч",
    "Технічна перевірка": "Тех-Перев",
}
_DEFAULT_CHECK_CODES: Dict[str, str] = {
    "Планова": "План",
    "Позапланова": "Позаплан",
    "Виконання припису": "Припис",
    "Заміна лічильника": "Зам-Ліч",
}

WORK_CODES: Dict[str, str] = {}
WORK_CODES.update(_DEFAULT_WORK_CODES)
CHECK_CODES: Dict[str, str] = {}
CHECK_CODES.update(_DEFAULT_CHECK_CODES)

_scan_number_re = re.compile(r"(\d{4})")

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


def sanitize_component(
    value: str,
    *,
    spaces_to_dash: bool = False,
    underscores_to_dash: bool = False,
) -> str:
    """Нормалізує фрагмент імені файлу.

    `underscores_to_dash`: замінює «_» на «-», щоб не ламати поділ полів при експорті
    (поля з'єднуються символом «_»).
    """
    s = (value or "").strip()
    s = _ZERO_WIDTH_RE.sub("", s)
    if underscores_to_dash:
        s = s.replace("_", "-")
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
    m = _scan_number_re.search(filename or "")
    return m.group(1) if m else "0000"


def init_naming_rules(app_dir: Path) -> None:
    """Завантажує `rules.json` з каталогу програми та з %APPDATA%\\PDF_Rename_Expert\\.

    Дозволяє змінювати `work_codes`, `check_codes` і `scan_number_pattern` без перезбірки exe.
    """
    global _scan_number_re

    WORK_CODES.clear()
    WORK_CODES.update(dict(_DEFAULT_WORK_CODES))
    CHECK_CODES.clear()
    CHECK_CODES.update(dict(_DEFAULT_CHECK_CODES))
    pattern = r"(\d{4})"

    for rules_path in (app_dir / "rules.json", user_config_dir() / "rules.json"):
        if not rules_path.is_file():
            continue
        try:
            data = json.loads(rules_path.read_text(encoding="utf-8"))
        except Exception:
            continue
        if not isinstance(data, dict):
            continue
        wc = data.get("work_codes")
        if isinstance(wc, dict):
            for k, v in wc.items():
                if (
                    isinstance(k, str)
                    and isinstance(v, str)
                    and k.strip()
                    and v.strip()
                ):
                    WORK_CODES[k.strip()] = v.strip()
        cc = data.get("check_codes")
        if isinstance(cc, dict):
            for k, v in cc.items():
                if (
                    isinstance(k, str)
                    and isinstance(v, str)
                    and k.strip()
                    and v.strip()
                ):
                    CHECK_CODES[k.strip()] = v.strip()
        sp = data.get("scan_number_pattern")
        if isinstance(sp, str) and sp.strip():
            pattern = sp.strip()

    try:
        _scan_number_re = re.compile(pattern)
    except re.error:
        _scan_number_re = re.compile(r"(\d{4})")


def parse_renamed_pdf_filename(filename: str) -> Optional[dict]:
    """
    Розбір імені, яке формує програма:
      scan_eic_date_meter_name_address_[Інвест]_phase_check_work_[IP-xxx]_[Mod-yyy]_[належність].pdf

    У полях name/address/meter тощо символ «_» не повинен зустрічатися (замінюється на «-» при
    перейменуванні). Старі файли з «_» у середині полів можуть розбиратися некоректно.
    """
    if not filename.lower().endswith(".pdf"):
        return None
    base = filename[:-4]
    parts = base.split("_")
    if len(parts) < 9:
        return None

    scan = parts[0]
    eic = parts[1]
    dt = parts[2]
    meter = parts[3]
    nm = parts[4]
    addr = parts[5]

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
