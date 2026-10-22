"""Разбор Excel: группы, дни, пары, объединения и отбор групп по курсу."""
from __future__ import annotations

import datetime as dt

import pytest

from conftest import SAMPLE_FILES, sample
from schedule_app.excel_parser import (
    DAY_NAMES, Column, SheetSchedule, _clean, _Grid, _normalize_day, _parse_grid, load_excel,
)

# Ожидания по каждому примеру: курс, число столбцов, группы
EXPECTED = {
    "ФКЯиК": (2, 12, ["ТянУИЯ", "201", "202", "203", "204", "205"]),
    "ФРЯ": (5, 7, ["501", "502", "503"]),
    "перев": (4, 19, ["401", "402", "403", "404", "405", "406", "407", "408", "409", "41"]),
    "ф_т_ан_яз,": (2, 32, [str(g) for g in range(201, 217)]),
    "нем_яз": (3, 8, ["301", "302", "303", "304"]),
    "зимний": (2, 32, [str(g) for g in range(201, 217)]),
}


def test_every_sample_has_expectation():
    """Новый пример в testing/ — добавьте его в EXPECTED."""
    for path in SAMPLE_FILES:
        assert sum(key in path.name for key in EXPECTED) == 1, path.name


@pytest.mark.parametrize("key", EXPECTED)
def test_sample_structure(load_sheet, key):
    course, ncols, groups = EXPECTED[key]
    sheet = load_sheet(sample(key))
    assert sheet.course == course
    assert len(sheet.columns) == ncols
    assert sheet.groups == groups
    assert sheet.days, "не найдено ни одного дня"
    for day in sheet.days:
        assert day.name in DAY_NAMES
        assert day.blocks, f"{day.name}: нет пар"
        for blk in day.blocks:
            assert blk.rows
            assert blk.time or blk.pair
    assert sheet.cells, "не прочитано ни одного занятия"


@pytest.mark.parametrize("key", EXPECTED)
def test_sample_columns_are_unique_and_ordered(load_sheet, key):
    sheet = load_sheet(sample(key))
    names = [c.name for c in sheet.columns]
    assert len(names) == len(set(names))
    assert [c.index for c in sheet.columns] == sorted(c.index for c in sheet.columns)
    for c in sheet.columns:
        assert " " not in c.name
        assert c.group == c.name.split("/")[0]


def test_non_numeric_group_is_parsed(load_sheet):
    sheet = load_sheet(sample("ФКЯиК"))
    tyan = [c for c in sheet.columns if c.group == "ТянУИЯ"]
    assert [c.name for c in tyan] == ["ТянУИЯ"]
    assert any(col == tyan[0].index for (_, col) in sheet.cells), "у ТянУИЯ нет ни одного занятия"


def test_load_excel_without_schedule_raises(tmp_path):
    import openpyxl
    path = tmp_path / "empty.xlsx"
    wb = openpyxl.Workbook()
    wb.active["A1"] = "просто текст"
    wb.save(path)
    with pytest.raises(ValueError, match="День недели"):
        load_excel(path)


# --------------------------------------------------------- columns_for_course

def _sheet(*names: str) -> SheetSchedule:
    cols = [Column(index=i, name=n, group=n.split("/")[0]) for i, n in enumerate(names)]
    return SheetSchedule("Лист", "", cols, [])


def _names(cols):
    return [c.name for c in cols]


def test_columns_for_course_keeps_non_numeric_group():
    sheet = _sheet("ТянУИЯ", "201/1", "201/2", "202/1")
    assert _names(sheet.columns_for_course(2)) == ["ТянУИЯ", "201/1", "201/2", "202/1"]


def test_columns_for_course_filters_other_courses():
    sheet = _sheet("201/1", "301/1", "ТянУИЯ", "302/1")
    assert _names(sheet.columns_for_course(3)) == ["301/1", "ТянУИЯ", "302/1"]


def test_columns_for_course_missing_course_returns_all():
    sheet = _sheet("ТянУИЯ", "201/1", "202/1")
    assert _names(sheet.columns_for_course(3)) == ["ТянУИЯ", "201/1", "202/1"]


def test_columns_for_course_on_sample(load_sheet):
    sheet = load_sheet(sample("ФКЯиК"))
    assert "ТянУИЯ" in _names(sheet.columns_for_course(2))
    assert len(sheet.columns_for_course(2)) == len(sheet.columns)


# ------------------------------------------------------------ course property

def test_course_from_title():
    sheet = _sheet("101/1")
    sheet.title = "Расписание, 4-й курс"
    assert sheet.course == 4


def test_course_from_groups_ignores_non_numeric():
    assert _sheet("ТянУИЯ", "301/1", "302/1").course == 3


def test_course_ambiguous():
    assert _sheet("201/1", "301/1").course is None
    assert _sheet("ТянУИЯ").course is None


# ------------------------------------------------------ синтетический лист

def _grid():
    """Лист: шапка в строке 1, понедельник на 4 строки (пара 1 — 1 строка, пара 2 — 3 строки)."""
    values = [
        ["Расписание занятий: ФКЯиК, 2-й курс", None, None, None, None, None],
        ["День недели", "Пара", "Время пары", "ТянУИЯ", "201 / 1", "201/2"],
        ["Понедельник", "1", "8:30-9:50", "Китайский", "Лекция для всех", None],
        [None, "2", "10:05-11:25", "Курс по выбору 1", "Англ", "Нем"],
        [None, None, None, "Курс по выбору 2", None, None],
        [None, None, None, "Курс по выбору 3", None, None],
        ["Вторник", "1", "8:30-9:50", None, None, None],
    ]
    merges = [
        (2, 0, 5, 0),     # понедельник
        (3, 1, 5, 1),     # номер пары 2
        (3, 2, 5, 2),     # время пары 2
        (2, 4, 2, 5),     # лекция на обе подгруппы 201
        (3, 4, 5, 4),     # Англ на три строки
    ]
    return _Grid("Лист1", values, merges)


def test_synthetic_columns_and_title():
    sheet = _parse_grid(_grid())
    assert [c.name for c in sheet.columns] == ["ТянУИЯ", "201/1", "201/2"]
    assert sheet.groups == ["ТянУИЯ", "201"]
    assert sheet.title.startswith("Расписание занятий")
    assert sheet.course == 2


def test_synthetic_blocks_follow_time_merge():
    sheet = _parse_grid(_grid())
    assert [d.name for d in sheet.days] == ["Понедельник", "Вторник"]
    mon = sheet.days[0]
    assert [(b.pair, b.time, b.rows) for b in mon.blocks] == [
        ("1", "8:30-9:50", [2]),
        ("2", "10:05-11:25", [3, 4, 5]),
    ]


def test_synthetic_merges_kept():
    sheet = _parse_grid(_grid())
    lecture = sheet.lesson(2, 4)
    assert lecture.text == "Лекция для всех"
    assert lecture.merge == (2, 4, 2, 5)
    assert sheet.lesson(2, 5).merge == (2, 4, 2, 5)
    assert sheet.lesson(5, 4).merge == (3, 4, 5, 4)
    assert sheet.lesson(4, 3).text == "Курс по выбору 2"


def test_grid_without_header_is_skipped():
    assert _parse_grid(_Grid("Пусто", [["нет шапки"]], [])) is None


# ------------------------------------------------------------------- утилиты

@pytest.mark.parametrize("value, expected", [
    (None, ""),
    (201.0, "201"),
    (2.5, "2.5"),
    (dt.time(8, 30), "08:30"),
    ("  a\xa0\xa0b  \n  c ", "a b\nc"),
])
def test_clean(value, expected):
    assert _clean(value) == expected


@pytest.mark.parametrize("text, day", [
    ("Понедельник", "Понедельник"),
    ("  ВТОРНИК ", "Вторник"),
    ("пт", None),
    ("Суб.", "Суббота"),
    ("Пара", None),
])
def test_normalize_day(text, day):
    assert _normalize_day(text) == day
