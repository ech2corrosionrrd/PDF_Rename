"""Тести чистої побудови імені PDF."""

from datetime import date, timedelta

from file_builder import (
    FilenameInputs,
    build_pdf_filename,
    derive_check_code,
    derive_phase,
    derive_work_code,
)


def _base_inputs(**overrides) -> FilenameInputs:
    defaults = dict(
        pdf_filename="scan1234.pdf",
        eic_full_raw="32X1234567890ABC",  # 16 chars exactly
        eic_suffix_raw="ABCD0",
        has_matches=True,
        date_str=date.today().isoformat(),
        meter_raw="MET-1",
        name_raw="Іванов Іван",
        address_raw="вул. Тестова 1",
        invest=False,
        phase_raw="3Ф",
        phase_ktt_raw="XXX",
        work_type="Монтаж лічильника",
        work_custom_raw="",
        check_type="Планова",
        ip_raw="",
        modem_raw="",
        name_suffix_raw="",
    )
    defaults.update(overrides)
    return FilenameInputs(**defaults)


def test_derive_phase_default():
    assert derive_phase("", "") == ("3Ф", None)
    assert derive_phase("1Ф", "") == ("1Ф", None)


def test_derive_phase_ktt_requires_value():
    val, err = derive_phase("3Ф-Ктт-XXX", "XXX")
    assert val == "3Ф-Ктт-XXX" and err is not None
    val, err = derive_phase("3Ф-Ктт-XXX", "60-5")
    assert val == "3Ф-Ктт-60-5" and err is None


def test_derive_work_code_known_and_other():
    assert derive_work_code("Монтаж лічильника", "") == "Монт-Ліч"
    assert derive_work_code("Інше", "Ремонт") == "ІН-Ремонт"
    assert derive_work_code("Інше", "") == "ІН"


def test_derive_check_code():
    assert derive_check_code("Планова") == "План"
    assert derive_check_code("невідомо") == "ВП"


def test_build_basic_no_errors_and_extension():
    name, errs = build_pdf_filename(_base_inputs())
    assert errs == []
    assert name.endswith(".pdf")
    assert name.startswith("1234_32X1234567890ABC_")


def test_build_invest_segment_present():
    name, errs = build_pdf_filename(_base_inputs(invest=True))
    assert errs == []
    assert "_Інвест_" in name


def test_build_askoe_appends_ip_and_mod():
    name, errs = build_pdf_filename(_base_inputs(
        work_type="Монтаж АСКОЕ", ip_raw="192.168.1.1", modem_raw="42"
    ))
    assert errs == []
    assert "_IP-192.168.1.1" in name
    assert "_Mod-42" in name


def test_build_askoe_without_optional_fields():
    name, errs = build_pdf_filename(_base_inputs(work_type="Монтаж АСКОЕ"))
    assert errs == []
    assert "IP-" not in name and "Mod-" not in name


def test_build_underscore_in_field_becomes_dash_keeps_segments_count():
    name, _ = build_pdf_filename(_base_inputs(
        name_raw="Foo_Bar", address_raw="A_B_C"
    ))
    # Назва й адреса мають перетворитися на Foo-Bar / A-B-C, тобто залишатися 1 токен у кожному
    parts = name[:-4].split("_")
    # 1234, EIC, date, meter, name, addr, phase, check, work => 9
    assert len(parts) == 9


def test_build_eic_too_short_emits_error():
    name, errs = build_pdf_filename(_base_inputs(eic_full_raw="ABC"))
    assert any("EIC має містити 16" in e for e in errs)
    assert name.endswith(".pdf")


def test_build_missing_eic_with_matches_message():
    inp = _base_inputs(eic_full_raw="", eic_suffix_raw="ABCDE", has_matches=True)
    _, errs = build_pdf_filename(inp)
    assert any("Оберіть запис" in e for e in errs)


def test_build_missing_eic_without_matches_message():
    inp = _base_inputs(eic_full_raw="", eic_suffix_raw="", has_matches=False)
    _, errs = build_pdf_filename(inp)
    assert any("Введіть повний EIC" in e for e in errs)


def test_build_invalid_date_emits_error():
    _, errs = build_pdf_filename(_base_inputs(date_str="not-a-date"))
    assert any("Некоректна дата" in e for e in errs)


def test_build_future_date_warns():
    future = (date.today() + timedelta(days=2)).isoformat()
    _, errs = build_pdf_filename(_base_inputs(date_str=future))
    assert any("майбутньому" in e for e in errs)


def test_build_with_name_suffix_appended():
    name, _ = build_pdf_filename(_base_inputs(name_suffix_raw="Мій додаток"))
    assert name.endswith("_Мій-додаток.pdf")


def test_build_filename_length_capped():
    long_name = "Назва " * 50
    long_addr = "Адреса " * 50
    name, _ = build_pdf_filename(_base_inputs(name_raw=long_name, address_raw=long_addr))
    assert len(name) <= 200
