"""Тести чистої логіки імен (pytest)."""

from pathlib import Path

from naming import (
    ensure_unique_path,
    extract_scan_number,
    parse_renamed_pdf_filename,
    sanitize_component,
    truncate_part,
)


def test_sanitize_basic():
    assert sanitize_component("Test/File:Name*") == "Test-File-Name"
    assert sanitize_component("Spaces   to   Single") == "Spaces to Single"


def test_sanitize_underscores_to_dash():
    assert sanitize_component("foo_bar", underscores_to_dash=True) == "foo-bar"
    assert sanitize_component("foo_bar") == "foo_bar"


def test_truncate_part():
    assert truncate_part("abc", 10) == "abc"
    assert truncate_part("abcdefghijklmnop", 5) == "abcde"


def test_ensure_unique_path(tmp_path: Path):
    p = tmp_path / "a.pdf"
    p.write_bytes(b"x")
    u = ensure_unique_path(p)
    assert u.name == "a (1).pdf"


def test_extract_scan_number():
    assert extract_scan_number("abc1234def.pdf") == "1234"
    assert extract_scan_number("nope.pdf") == "0000"


def test_parse_renamed_roundtrip_shape():
    filename = (
        "1234_32X1234567890ABC_2023-10-27_MET123_NAME_ADDR_Інвест_3Ф_План_Монт-АСКОЕ.pdf"
    )
    res = parse_renamed_pdf_filename(filename)
    assert res is not None
    assert res["EIC"] == "32X1234567890ABC"
    assert res["Інвест"] == "Так"


def test_parse_with_suffix_and_underscores_normalized_in_field():
    with_suffix = (
        "1234_32X1234567890ABC_2023-10-27_MET123_NAME_ADDR_Інвест_3Ф_План_"
        "Монт-АСКОЕ_IP-1.1.1.1_Mod-99_Мій-додаток.pdf"
    )
    r2 = parse_renamed_pdf_filename(with_suffix)
    assert r2 is not None
    assert r2["Належність лічильника"] == "Мій-додаток"
    assert r2["IP адреса"] == "1.1.1.1"
    assert r2["№ Модему"] == "99"


def test_parse_invalid_returns_none():
    assert parse_renamed_pdf_filename("foo.txt") is None
    assert parse_renamed_pdf_filename("too_few_parts.pdf") is None
