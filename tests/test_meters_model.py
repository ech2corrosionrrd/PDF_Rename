"""Тести моделі аркуша «база» (без Tk і Excel)."""

from datetime import date, datetime

import pytest

import meters_model as mm
from meters_model import FIRST_DATA_ROW, LAST_COL, Snapshot


def _snapshot(rows):
    hdr = [[""] * LAST_COL for _ in range(3)]
    hdr[0][0:6] = ["№", "EIC", "Підстанція", "Стан", "Приєднання", "Лічильник"]
    hdr[0][6] = "№ лічильника"
    wide = [False] * LAST_COL
    # група «Лічильник»: стовпці 12-15 з Квартал/Рік
    for c in range(12, 16):
        wide[c - 1] = True
        hdr[0][c - 1] = "Лічильник"
    hdr[1][12 - 1] = "Держповірка"
    hdr[2][12 - 1] = "Квартал"
    hdr[1][13 - 1] = "Держповірка"
    hdr[2][13 - 1] = "Рік"
    values = []
    for r in rows:
        row = [""] * LAST_COL
        for c, v in r.items():
            row[c - 1] = v
        values.append(row)
    return Snapshot(hdr=hdr, row1_wide=wide, values=values)


ROWS = [
    {1: "1", 2: "62Z100938201573H", 3: "ЕЧЕ-1", 5: "Ввод-1", 7: "84760419"},
    {1: "2", 2: "62Z231508073446X", 3: "ЕЧЕ-1", 5: "Ввод-2", 7: "13613541 (прием)"},
    {1: "3", 2: "", 3: "ЕЧЕ-2", 5: "Ввод-1", 7: "00543"},
    {1: "4", 2: "", 3: "ЕЧЕ-2", 5: "Ввод-2", 7: "1111111 (2222222)"},
]


def test_col_letter():
    assert [mm.col_letter(c) for c in (1, 26, 27, 52, 60)] == ["A", "Z", "AA", "AZ", "BH"]


def test_group_and_caption():
    s = _snapshot(ROWS)
    assert mm.group_of(s, 12) == "Лічильник"
    assert mm.group_of(s, 1) == "Загальні"
    assert mm.caption_of(s, 12) == "Держповірка / Квартал"
    assert mm.caption_of(s, 56) == "Код виконавця"
    assert mm.caption_of(s, 30) == "Стовпець AD"


def test_layout_joins_quarter_year_block():
    s = _snapshot(ROWS)
    blocks = {b.cols: b for g in mm.build_layout(s) for b in g.blocks}
    blk = blocks[(12, 13)]
    assert blk.roles == ("Q", "Y")
    assert blk.label == "Держповірка (квартал / рік)"
    assert all(13 not in b.cols for b in blocks.values() if b.cols != (12, 13))
    # кожен стовпець потрапляє у форму рівно один раз
    cols = sorted(c for b in blocks.values() for c in b.cols)
    assert cols == list(range(1, LAST_COL + 1))


@pytest.mark.parametrize(
    "text,expected",
    [
        ("", (None, False)),
        ("  ", (None, False)),
        ("0999", ("0999", True)),
        ("=1+1", ("=1+1", True)),
        ("2021", (2021, False)),
        ("5,5", (5.5, False)),
        ("0,2S", ("0,2S", False)),
        ("35000/100", ("35000/100", False)),
        ("14.03.2022", (date(2022, 3, 14), False)),
        ("31.02.2022", ("31.02.2022", False)),
        ("0", (0, False)),
    ],
)
def test_parse_cell_text(text, expected):
    assert mm.parse_cell_text(text) == expected


def test_format_value():
    assert mm.format_value(None) == ""
    assert mm.format_value(2021.0) == "2021"
    assert mm.format_value(5.5) == "5,5"
    assert mm.format_value(datetime(2024, 8, 14)) == "14.08.2024"
    assert mm.format_value(" ТОЛ-35 ") == "ТОЛ-35"


def test_id_query_requires_min_length():
    assert mm.parse_id_query("84760419, 3446X") == (["84760419", "3446X"], None)
    tokens, err = mm.parse_id_query("1234")
    assert tokens == [] and err


def test_filter_by_id_ignores_notes_and_matches_tail():
    s = _snapshot(ROWS)
    rows = lambda q: mm.filter_rows(s, id_tokens=mm.parse_id_query(q)[0])
    assert rows("3446X") == [5]
    assert rows("760419") == [4]
    assert rows("13613541") == [5]  # примітка «(прием)» поруч не заважає
    assert rows("2222222") == []  # значення в дужках — це примітка
    assert rows("1111111") == [7]
    assert rows("00543") == [6]


def test_filter_by_substation_and_text():
    s = _snapshot(ROWS)
    assert mm.filter_rows(s, "ЕЧЕ-2") == [6, 7]
    assert mm.filter_rows(s, mm.ALL_SUBSTATIONS, "ввод-2") == [5, 7]
    assert mm.filter_rows(s) == [4, 5, 6, 7]


def test_navigator_keeps_position_and_resets_filters():
    s = _snapshot(ROWS)
    nav = mm.Navigator(s)
    nav.move("last")
    assert nav.current == 7
    nav.goto(6)
    nav.subst = "ЕЧЕ-1"
    nav.refresh()
    assert nav.rows == [4, 5]
    assert nav.goto(6) and nav.current == 6 and nav.subst == mm.ALL_SUBSTATIONS
    assert nav.find_connection("ЕЧЕ-1", "Ввод-2") == 5
    assert nav.find_connection("ЕЧЕ-2", "Ввод-2") == 7
    assert nav.find_connection("ЕЧЕ-2", "Немає") is None


def test_locked_cells_and_changed_fields():
    s = _snapshot(ROWS)
    s.merge_anchor[(5, 5)] = 4  # об'єднання E4:E5
    assert s.is_locked(5, 5) and not s.is_locked(4, 5)
    assert s.is_locked(4, mm.COL_SUBST)
    new = {5: "інше", 7: "NEW", 3: "хак"}
    assert mm.changed_fields(s, 5, new) == {7: "NEW"}


def test_set_text_updates_whole_merge_block():
    s = _snapshot(ROWS)
    s.merge_anchor[(5, 5)] = 4
    s.set_text(4, 5, "X")
    assert s.text(4, 5) == "X" and s.text(5, 5) == "X" and s.text(6, 5) == "Ввод-1"


def test_list_for_priority():
    roles = {12: "Q"}
    ref = {"Квартал": ["I", "II"], "G": ["a"]}
    auto = {7: ["x"], 8: ["y"]}
    assert mm.list_for(12, roles, ref, auto) == ["I", "II"]
    assert mm.list_for(12, roles, {}, auto) == ["I", "II", "III", "IV"]
    assert mm.list_for(7, roles, ref, auto) == ["a"]  # довідник за літерою G
    assert mm.list_for(8, roles, ref, auto) == ["y"]
    assert mm.list_for(9, roles, ref, auto) is None


def test_block_text_and_card_name():
    b = mm.Block(cols=(1, 2, 3), label="Дата", sep=".", roles=("D", "M", "Y"))
    assert mm.block_text(b, ["7", "3", "2022"]) == "07.03.2022"
    assert mm.block_text(b, ["", "", ""]) == ""
    ph = mm.Block(cols=(1, 2, 3), label="№", sep="#")
    assert mm.block_text(ph, ["1", "", "3"]) == "1 / – / 3"
    name = mm.card_pdf_name("ЕЧЕ-1 (Новоселівка)", 'Ввод: 1/2', datetime(2026, 9, 27, 0, 34))
    assert name == "Дані по лічильнику - ЕЧЕ-1 (Новоселівка) - Ввод-1-2 - 27.09.2026 00-34.pdf"


def test_act_date_and_changed_detection():
    assert mm.act_date_text("7", "3", "2022") == "07.03.2022"
    assert mm.act_date_text("", "", "") == ""
    assert mm.act_fields_changed({57: "0543"}) and mm.act_fields_changed({60: "2026"})
    assert not mm.act_fields_changed({7: "x", 9: "y"})


def test_act_entry_uses_new_values_over_saved():
    s = _snapshot(ROWS)
    s.values[0][57 - 1] = "OLD"
    when = datetime(2026, 10, 5, 12, 0)
    e = mm.act_entry_for_record(
        s, 4, {57: "0543", 58: "7", 59: "3", 60: "2022", 56: "00563"},
        user="op", when=when, scan="0001_x.pdf", work_type="Монтаж АСКОЕ",
    )
    assert (e.act_no, e.act_date, e.executor) == ("0543", "07.03.2022", "00563")
    assert (e.subst, e.conn, e.eic, e.meter) == ("ЕЧЕ-1", "Ввод-1", "62Z100938201573H", "84760419")
    row = e.as_row()
    assert len(row) == len(mm.ACT_HEADERS) and row[0] == when and row[13] == 4
    assert e.event == mm.EVENT_RECORD_EDIT
