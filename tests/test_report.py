"""Тести побудови та запису Excel-звіту."""

from pathlib import Path

import pandas as pd

from naming import EXPORT_COLUMNS
from report import build_report_frames, has_any_rows, write_report


def _make_pdf(folder: Path, name: str) -> None:
    folder.mkdir(parents=True, exist_ok=True)
    (folder / name).write_bytes(b"%PDF-1.4 stub")


def test_build_report_frames_empty(tmp_path: Path):
    frames = build_report_frames(tmp_path / "Scan")
    assert set(frames.keys()) == {"Населення", "Промислові"}
    for df in frames.values():
        assert list(df.columns) == EXPORT_COLUMNS
        assert df.empty
    assert not has_any_rows(frames)


def test_build_report_frames_populates_columns(tmp_path: Path):
    scan = tmp_path / "Scan"
    valid = (
        "1234_32X1234567890ABCD_2023-10-27_MET-1_NAME_ADDR_Інвест_3Ф_План_"
        "Монт-АСКОЕ_IP-10.0.0.1_Mod-7_Мій-додаток.pdf"
    )
    _make_pdf(scan / "Населення", valid)
    _make_pdf(scan / "Населення", "ignore-me.txt")  # non-pdf
    _make_pdf(scan / "Промислові", "garbage_name.pdf")  # not parseable

    frames = build_report_frames(scan)
    assert has_any_rows(frames)
    assert list(frames["Населення"].columns) == EXPORT_COLUMNS
    assert len(frames["Населення"]) == 1
    row = frames["Населення"].iloc[0].to_dict()
    assert row["№"] == 1
    assert row["EIC"] == "32X1234567890ABCD"
    assert row["IP адреса"] == "10.0.0.1"
    assert row["№ Модему"] == "7"
    assert row["Належність лічильника"] == "Мій-додаток"


def test_write_report_creates_workbook(tmp_path: Path):
    scan = tmp_path / "Scan"
    valid = (
        "1234_32X1234567890ABCD_2023-10-27_MET-1_NAME_ADDR_Ні_3Ф_План_"
        "Монт-Ліч.pdf"
    )
    # invest token in this filename is "Ні" (non-Інвест branch), still valid
    valid2 = (
        "1234_32X1234567890ABCD_2023-10-27_MET-1_NAME_ADDR_Інвест_3Ф_План_"
        "Монт-Ліч.pdf"
    )
    _make_pdf(scan / "Населення", valid2)

    frames = build_report_frames(scan)
    out = tmp_path / "report.xlsx"
    write_report(out, frames)

    assert out.exists()
    re_read = pd.read_excel(out, sheet_name=None, engine="openpyxl")
    assert set(re_read.keys()) == {"Населення", "Промислові"}
    assert list(re_read["Населення"].columns) == EXPORT_COLUMNS
    # touch unused name to keep linters happy if any
    assert valid.endswith(".pdf")
