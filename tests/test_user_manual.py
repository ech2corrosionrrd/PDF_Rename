"""Тести довідки користувача в програмі."""

from pathlib import Path

from user_manual import (
    APP_USER_MANUAL_FILENAME,
    FULL_USER_MANUAL_FILENAME,
    load_app_user_manual_text,
    resolve_app_user_manual_path,
    resolve_full_user_manual_path,
)


def test_resolve_app_manual_in_repo():
    app_dir = Path(__file__).resolve().parents[1]
    p = resolve_app_user_manual_path(app_dir)
    assert p is not None
    assert p.name == APP_USER_MANUAL_FILENAME


def test_app_manual_has_no_developer_sections():
    app_dir = Path(__file__).resolve().parents[1]
    text, _ = load_app_user_manual_text(app_dir)
    assert "pip install" not in text
    assert "build_win7.bat" not in text
    assert "pdf_rename_expert.py" not in text
    assert "README.md" not in text
    assert "Технічний облік" in text
    assert "F1" in text


def test_full_manual_exists_separately():
    app_dir = Path(__file__).resolve().parents[1]
    p = resolve_full_user_manual_path(app_dir)
    assert p is not None
    assert p.name == FULL_USER_MANUAL_FILENAME
    content = p.read_text(encoding="utf-8")
    assert "Встановлення Python" in content
