"""Спільні pytest-фікстури."""

from __future__ import annotations

from pathlib import Path

import pytest

from tests.fixtures.baza_workbook import write_minimal_baza


@pytest.fixture
def excel_app_dir(tmp_path: Path) -> Path:
    """Тимчасовий каталог з мінімальним xlsx для ExcelConsumerDB."""
    write_minimal_baza(tmp_path / "test_baza.xlsx")
    return tmp_path


@pytest.fixture(scope="session", autouse=True)
def _init_default_naming_rules() -> None:
    """Після тестів з кастомним `rules.json` глобальні словники мають повертатися до типових."""
    from naming import init_naming_rules

    repo = Path(__file__).resolve().parent.parent
    init_naming_rules(repo)
    yield
    init_naming_rules(repo)
