"""Чиста логіка побудови імені перейменованого PDF.

Функції цього модуля не залежать від Tkinter і легко тестуються.
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import date
from typing import List, Optional, Tuple

from naming import (
    CHECK_CODES,
    MAX_ADDR_PART_LEN,
    MAX_FILENAME_LEN,
    MAX_NAME_PART_LEN,
    WORK_CODES,
    extract_scan_number,
    sanitize_component,
    truncate_part,
)


@dataclass
class FilenameInputs:
    """Сирі значення з форми: жодних мутацій, лише дані."""

    pdf_filename: Optional[str]
    eic_full_raw: str
    eic_suffix_raw: str
    has_matches: bool
    date_str: str
    meter_raw: str
    name_raw: str
    address_raw: str
    invest: bool
    phase_raw: str
    phase_ktt_raw: str
    work_type: str
    work_custom_raw: str
    check_type: str
    ip_raw: str
    modem_raw: str
    name_suffix_raw: str


def derive_phase(phase: str, phase_ktt_raw: str) -> Tuple[str, Optional[str]]:
    """Нормалізує значення фази; повертає (значення, помилка|None)."""
    p = (phase or "").strip()
    if p == "3Ф-Ктт-XXX":
        xxx = sanitize_component(
            phase_ktt_raw, spaces_to_dash=True, underscores_to_dash=True
        )
        if not xxx or xxx.upper() == "XXX":
            return "3Ф-Ктт-XXX", "Заповніть коефіцієнт для 3Ф-Ктт (замість XXX)."
        return f"3Ф-Ктт-{xxx}", None
    return (p or "3Ф"), None


def derive_work_code(work_type: str, custom_raw: str) -> str:
    wt = (work_type or "").strip()
    if wt == "Інше":
        custom = sanitize_component(
            custom_raw, spaces_to_dash=True, underscores_to_dash=True
        )
        return f"ІН-{custom}" if custom else "ІН"
    return WORK_CODES.get(wt, "ВР")


def derive_check_code(check_type: str) -> str:
    return CHECK_CODES.get((check_type or "").strip(), "ВП")


_DATE_RE = re.compile(r"\d{4}-\d{2}-\d{2}")


def build_pdf_filename(inp: FilenameInputs) -> Tuple[str, List[str]]:
    """Будує фінальне ім'я PDF та повертає список повідомлень про помилки."""
    errors: List[str] = []
    scan_no = extract_scan_number(inp.pdf_filename) if inp.pdf_filename else "0000"

    eic_full = sanitize_component(
        inp.eic_full_raw, spaces_to_dash=False, underscores_to_dash=True
    ).replace(" ", "")
    if not eic_full:
        if len((inp.eic_suffix_raw or "").strip()) >= 5 and inp.has_matches:
            errors.append("Оберіть запис зі списку або введіть повний EIC.")
        else:
            errors.append(
                "Введіть повний EIC (16 символів) або виконайте пошук по базі."
            )
    elif len(eic_full) != 16:
        errors.append("EIC має містити 16 символів.")

    dt = (inp.date_str or "").strip()
    if not _DATE_RE.fullmatch(dt):
        errors.append("Некоректна дата (очікується YYYY-MM-DD).")
    else:
        try:
            if date.fromisoformat(dt) > date.today():
                errors.append("Дата в майбутньому — перевірте календар.")
        except Exception:
            pass

    meter = sanitize_component(
        inp.meter_raw, spaces_to_dash=True, underscores_to_dash=True
    )
    if not meter:
        errors.append("Заповніть № Лічильника.")

    name = sanitize_component(
        inp.name_raw, spaces_to_dash=True, underscores_to_dash=True
    )
    if not name:
        errors.append("Не заповнена Назва.")
    name = truncate_part(name, MAX_NAME_PART_LEN)

    addr = sanitize_component(
        inp.address_raw, spaces_to_dash=True, underscores_to_dash=True
    )
    if not addr:
        errors.append("Не заповнена Адреса.")
    addr = truncate_part(addr, MAX_ADDR_PART_LEN)

    invest = "Інвест" if inp.invest else ""

    phase_val, phase_err = derive_phase(inp.phase_raw, inp.phase_ktt_raw)
    if phase_err:
        errors.append(phase_err)
    phase = sanitize_component(
        phase_val, spaces_to_dash=True, underscores_to_dash=True
    )
    check = sanitize_component(
        derive_check_code(inp.check_type),
        spaces_to_dash=True,
        underscores_to_dash=True,
    )
    work = sanitize_component(
        derive_work_code(inp.work_type, inp.work_custom_raw),
        spaces_to_dash=True,
        underscores_to_dash=True,
    )

    parts = [scan_no, eic_full, dt, meter, name, addr]
    if invest:
        parts.append(invest)
    parts.extend([phase, check, work])

    if "АСКОЕ" in (inp.work_type or ""):
        ip_val = sanitize_component(
            inp.ip_raw, spaces_to_dash=True, underscores_to_dash=True
        )
        mod_val = sanitize_component(
            inp.modem_raw, spaces_to_dash=True, underscores_to_dash=True
        )
        if ip_val:
            parts.append(f"IP-{ip_val}")
        if mod_val:
            parts.append(f"Mod-{mod_val}")

    final = "_".join([p for p in parts if p])
    final = sanitize_component(final, spaces_to_dash=False)
    final = final.replace(" ", "-")
    final = re.sub(r"-{2,}", "-", final).strip("-")

    name_suffix = sanitize_component(
        inp.name_suffix_raw, spaces_to_dash=True, underscores_to_dash=True
    )
    if name_suffix:
        final = f"{final}_{name_suffix}" if final else name_suffix
        final = re.sub(r"-{2,}", "-", final).strip("-")

    if len(final) + 4 > MAX_FILENAME_LEN:
        allowed = MAX_FILENAME_LEN - 4
        final = final[:allowed].rstrip(" -._")
    return f"{final}.pdf", errors
