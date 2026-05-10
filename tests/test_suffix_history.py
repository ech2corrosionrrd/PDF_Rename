"""Тести збереження/читання історії суфіксів."""

from pathlib import Path

from suffix_history import (
    MAX_SUFFIX_HISTORY_ITEMS,
    SUFFIX_HISTORY_FILENAME,
    load_suffix_history,
    save_suffix_history,
)


def test_load_when_missing_returns_empty(tmp_path: Path):
    assert load_suffix_history(tmp_path) == []


def test_save_then_load_roundtrip(tmp_path: Path):
    save_suffix_history(tmp_path, ["a", "b", "c"])
    assert (tmp_path / SUFFIX_HISTORY_FILENAME).is_file()
    assert load_suffix_history(tmp_path) == ["a", "b", "c"]


def test_save_dedupes_and_strips(tmp_path: Path):
    save_suffix_history(tmp_path, [" a ", "a", "b", "", "  ", "b"])
    assert load_suffix_history(tmp_path) == ["a", "b"]


def test_save_truncates_to_max(tmp_path: Path):
    items = [f"item-{i}" for i in range(MAX_SUFFIX_HISTORY_ITEMS + 5)]
    save_suffix_history(tmp_path, items)
    loaded = load_suffix_history(tmp_path)
    assert len(loaded) == MAX_SUFFIX_HISTORY_ITEMS
    assert loaded[0] == "item-0"


def test_load_ignores_invalid_payload(tmp_path: Path):
    (tmp_path / SUFFIX_HISTORY_FILENAME).write_text("not a json", encoding="utf-8")
    assert load_suffix_history(tmp_path) == []
