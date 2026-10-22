"""Автоопределение факультета, курса, недели и дат по имени файла / заголовку."""
from __future__ import annotations

import datetime as dt
import unicodedata

import pytest

from schedule_app.filename_meta import (
    FileMeta, parse_filename, parse_text, week_dates, week_number,
)


@pytest.mark.parametrize("name, faculty, course, week", [
    ("Расписание занятий   ф-т нем.яз, специалист, очная (дневная), 3-й курс, 7-я неделя.xlsx",
     "нем.яз", 3, 7),
    ("Расписание_занятий_ф_т_нем_яз,_специалист,_очная_дневная,_3_й_курс.xlsx", "нем яз", 3, None),
    ("Расписание_занятий_перев_ф_т,_специалист,_очная_дневная,_4_й_курс.xlsx", "перев", 4, None),
    ("Расписание_занятий_ФРЯ,_специалист,_очная_дневная,_5_й_курс,_8_я.xlsx", None, 5, 8),
    ("Расписание_занятий_ф_т_ан_яз,_специалист,_очная_дневная,_2_й_курс.xlsx", "ан яз", 2, None),
    ("ф_т_ан_яз_спе_дн_2к_зимний_10_11_2025_16_11_2025.xlsx", "ан яз", 2, None),
])
def test_parse_filename(name, faculty, course, week):
    meta = parse_filename(name)
    assert (meta.faculty_text, meta.course, meta.week) == (faculty, course, week)


def test_parse_filename_course_without_week_suffix():
    meta = parse_filename("Расписание_занятий_ФКЯиК,_специалист,_очная_дневная,_2_й_курс,_9.xlsx")
    assert meta.course == 2


@pytest.mark.xfail(reason="«…, 9» в конце имени без «я»/«неделя» пока не распознаётся как неделя",
                   strict=True)
def test_parse_filename_bare_week_number_at_end():
    meta = parse_filename("Расписание_занятий_ФКЯиК,_специалист,_очная_дневная,_2_й_курс,_9.xlsx")
    assert meta.week == 9


def test_parse_filename_dates():
    meta = parse_filename("ф_т_ан_яз_спе_дн_2к_зимний_10_11_2025_16_11_2025.xlsx")
    assert meta.date_start == dt.date(2025, 11, 10)
    assert meta.date_end == dt.date(2025, 11, 16)


def test_single_date_gives_week_range():
    meta = parse_text("неделя с 10.11.2025")
    assert meta.date_start == dt.date(2025, 11, 10)
    assert meta.date_end == dt.date(2025, 11, 16)


def test_invalid_date_ignored():
    assert parse_text("32.13.2025").date_start is None


def test_macos_nfd_filename():
    """macOS отдаёт имена в NFD — «й» разложена на «и» + знак краткой."""
    name = unicodedata.normalize("NFD", "Расписание_занятий_ФРЯ,_5_й_курс,_8_я.xlsx")
    meta = parse_filename(name)
    assert (meta.course, meta.week) == (5, 8)


def test_sheet_title():
    meta = parse_text("Расписание занятий: ФКЯиК, специалист, очная (дневная), 2-й курс, 9-я неделя")
    assert (meta.course, meta.week) == (2, 9)


def test_fill_missing_keeps_own_values():
    own = FileMeta(course=3)
    own.fill_missing(FileMeta(course=5, week=7))
    assert (own.course, own.week) == (3, 7)


def test_week_dates_and_number_roundtrip():
    monday = dt.date(2026, 8, 31)
    assert week_dates(monday, 1) == (monday, dt.date(2026, 9, 6))
    start, end = week_dates(monday, 9)
    assert start == dt.date(2026, 10, 26)
    assert week_number(monday, start) == 9
    assert week_number(monday, end) == 9
