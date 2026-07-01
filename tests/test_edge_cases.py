"""Граничні випадки: порожні файли, унікальні шляхи, звіт."""

from __future__ import annotations

import json
from pathlib import Path

import pytest

from naming import ensure_unique_path, init_naming_rules, sanitize_component
from report import build_report_frames
from theme import preview_len_style, status_dot_style


def test_ensure_unique_path_increments_twice(tmp_path: Path) -> None:
    p = tmp_path / "a.pdf"
    p.write_bytes(b"x")
    (tmp_path / "a (1).pdf").write_bytes(b"y")
    u = ensure_unique_path(p)
    assert u.name == "a (2).pdf"


def test_sanitize_strips_zero_width_chars() -> None:
    raw = "\u200bFoo\u200c/Bar\u200b"
    assert sanitize_component(raw) == "Foo-Bar"


def test_zero_byte_pdf_skipped_in_report(tmp_path: Path) -> None:
    scan = tmp_path / "Scan" / "Населення"
    scan.mkdir(parents=True)
    (scan / "empty.pdf").write_bytes(b"")
    frames = build_report_frames(tmp_path / "Scan")
    assert frames["Населення"].empty


def test_init_naming_rules_custom_scan_pattern(tmp_path: Path) -> None:
    rules = {"scan_number_pattern": r"SCN(\d{3})"}
    (tmp_path / "rules.json").write_text(json.dumps(rules), encoding="utf-8")
    init_naming_rules(tmp_path)
    from naming import extract_scan_number

    assert extract_scan_number("prefixSCN007suffix.pdf") == "007"
    repo_root = Path(__file__).resolve().parent.parent
    init_naming_rules(repo_root)


def test_preview_len_style_thresholds() -> None:
    assert preview_len_style(50, 200) == "PreviewLenOk.TLabel"
    assert preview_len_style(170, 200) == "PreviewLenWarn.TLabel"
    assert preview_len_style(201, 200) == "PreviewLenError.TLabel"


def test_status_dot_style_levels() -> None:
    assert status_dot_style("success") == "StatusDotSuccess.TLabel"
    assert status_dot_style("unknown") == "StatusDotInfo.TLabel"


def test_render_corrupt_pdf_raises(tmp_path: Path) -> None:
    pytest.importorskip("fitz", reason="PyMuPDF optional")
    from pdf_preview import render_pdf_pages

    bad = tmp_path / "bad.pdf"
    bad.write_bytes(b"not a pdf")
    with pytest.raises(RuntimeError):
        render_pdf_pages(bad, 400)
