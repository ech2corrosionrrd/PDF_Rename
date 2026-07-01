"""Тести читання Excel-бази."""

from __future__ import annotations

from pathlib import Path

import pytest

from excel_db import ConsumerRecord, ExcelConsumerDB, _is_technical_marker


def test_is_technical_marker():
    assert _is_technical_marker("Технічний")
    assert _is_technical_marker("(Технічний)")
    assert _is_technical_marker("  технічний  ")
    assert not _is_technical_marker("Населення")
    assert not _is_technical_marker("")


def test_consumer_record_display_technical():
    rec = ConsumerRecord(
        eic_full="",
        name="Ім'я",
        address="Адреса 1",
        tech_code="T-999",
        is_technical=True,
    )
    assert rec.display().startswith("T-999")


def test_load_and_record_count(excel_app_dir: Path) -> None:
    db = ExcelConsumerDB(excel_app_dir)
    db.load()
    assert db.record_count == 3


def test_search_by_eic_suffix_five_digits(excel_app_dir: Path) -> None:
    db = ExcelConsumerDB(excel_app_dir)
    matches = db.search_by_eic_suffix("01234")
    assert len(matches) == 2
    eics = {m.eic_full for m in matches}
    assert eics == {"UA12345678901234", "UA9999901234"}


def test_search_by_eic_suffix_full(excel_app_dir: Path) -> None:
    db = ExcelConsumerDB(excel_app_dir)
    matches = db.search_by_eic_suffix("UA12345678901234")
    assert len(matches) == 1
    assert matches[0].name == "ТОВ-Компанія"
    assert matches[0].address == "вул. Тестова 1"


def test_search_by_eic_suffix_too_short(excel_app_dir: Path) -> None:
    db = ExcelConsumerDB(excel_app_dir)
    assert db.search_by_eic_suffix("1234") == []


def test_search_technical(excel_app_dir: Path) -> None:
    db = ExcelConsumerDB(excel_app_dir)
    matches = db.search_technical()
    assert len(matches) == 1
    rec = matches[0]
    assert rec.is_technical
    assert rec.tech_code == "T-999"
    assert rec.name == "T-999-Об'єкт"
    assert rec.address == "Адреса тех"


def test_load_missing_excel_raises(tmp_path: Path) -> None:
    db = ExcelConsumerDB(tmp_path)
    with pytest.raises(FileNotFoundError, match="Не знайдено Excel-базу"):
        db.load()


def test_load_wrong_sheet_raises(tmp_path: Path) -> None:
    import openpyxl

    path = tmp_path / "bad.xlsx"
    wb = openpyxl.Workbook()
    wb.active.title = "інший"
    wb.save(path)

    db = ExcelConsumerDB(tmp_path)
    with pytest.raises(ValueError, match="база"):
        db.load()
