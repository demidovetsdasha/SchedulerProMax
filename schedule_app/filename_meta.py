"""Извлечение факультета, курса, номера недели и дат из имени файла / заголовка листа.

Поддерживаемые варианты:
    "Расписание занятий   ф-т нем.яз, специалист, очная (дневная), 3-й курс, 7-я неделя.xlsx"
    "Расписание_занятий_ф_т_нем_яз,_специалист,_очная_дневная,_3_й_курс.xlsx"
    "Расписание_занятий_перев_ф_т,_специалист,_очная_дневная,_4_й_курс.xlsx"
    "Расписание_занятий_ФРЯ,_специалист,_очная_дневная,_5_й_курс,_8_я.xlsx"
    "ф_т_ан_яз_спе_дн_2к_зимний_10_11_2025_16_11_2025.xlsx"
"""
from __future__ import annotations

import datetime as dt
import re
import unicodedata
from dataclasses import dataclass
from pathlib import Path

_FT = r"ф\s*[-\s]\s*т\.?"          # «ф-т», «ф т» (после замены «_» на пробел)

# «ф-т нем.яз,» → «нем.яз»  (текст после «ф-т » до запятой)
RE_FACULTY = re.compile(_FT + r"\s+([^,]+?)\s*(?:,|$)", re.I)
# «перев.ф-т,» → «перев»  (название перед «ф-т»)
RE_FACULTY_BEFORE = re.compile(r"(?:^|[\s,:])([а-яё]+(?:\.[а-яё]+)?)\.?\s*" + _FT + r"(?=[\s,]|$)", re.I)
# «ф т ан яз спе дн …» (старый формат без запятых) → «ан яз»
RE_FACULTY_OLD = re.compile(_FT + r"\s+(.+?)\s+(?:спе|бак|маг|дн|оч|заоч|веч|\d)", re.I)
# «3-й курс», «3 й курс» → 3
RE_COURSE = re.compile(r"(\d)\s*[-\s]?\s*(?:й|ой|ый)\s*курс", re.I)
# «… 2к …» → 2
RE_COURSE_SHORT = re.compile(r"(?:^|\s)(\d)\s*к(?:\s|$)", re.I)
# «7-я неделя», «7 я неделя» → 7
RE_WEEK = re.compile(r"(\d{1,2})\s*[-\s]?\s*(?:я|ая)?\s*недел", re.I)
# «…, 8 я» в конце имени (слово «неделя» обрезано) → 8
RE_WEEK_TAIL = re.compile(r"(?:^|[\s,])(\d{1,2})\s*[-\s]?\s*(?:я|ая)\s*$", re.I)
# «10_11_2025» / «10.11.2025»
RE_DATE = re.compile(r"(\d{1,2})[_.\-/](\d{1,2})[_.\-/](\d{4})")


@dataclass
class FileMeta:
    faculty_text: str | None = None   # как в источнике, напр. «нем.яз»
    course: int | None = None
    week: int | None = None
    date_start: dt.date | None = None
    date_end: dt.date | None = None

    def fill_missing(self, other: "FileMeta") -> "FileMeta":
        """Дополняет пустые поля значениями из другого источника (напр. заголовка листа)."""
        for name in ("faculty_text", "course", "week", "date_start", "date_end"):
            if getattr(self, name) is None:
                setattr(self, name, getattr(other, name))
        return self


def parse_text(raw: str) -> FileMeta:
    # macOS хранит имена файлов в NFD («й» = «и» + знак краткой) — приводим к NFC
    raw = unicodedata.normalize("NFC", raw)
    meta = FileMeta()
    dates = []
    for d, mo, y in RE_DATE.findall(raw):
        try:
            dates.append(dt.date(int(y), int(mo), int(d)))
        except ValueError:
            pass
    if dates:
        meta.date_start = dates[0]
        meta.date_end = dates[1] if len(dates) > 1 else dates[0] + dt.timedelta(days=6)

    # без дат; «_» → пробел, лишние пробелы схлопываем
    text = RE_DATE.sub(" ", raw).replace("\xa0", " ").replace("_", " ")
    text = re.sub(r"\s+", " ", text).strip()

    m = RE_FACULTY_OLD.search(text) or RE_FACULTY.search(text) or RE_FACULTY_BEFORE.search(text)
    if m:
        meta.faculty_text = m.group(1).strip(" .,")

    m = RE_COURSE.search(text) or RE_COURSE_SHORT.search(text)
    if m:
        meta.course = int(m.group(1))

    m = RE_WEEK.search(text) or RE_WEEK_TAIL.search(text)
    if m:
        meta.week = int(m.group(1))
    return meta


def parse_filename(name: str | Path) -> FileMeta:
    stem = Path(name).name if isinstance(name, Path) else str(name)
    stem = re.sub(r"\.(xlsx|xlsm|xls)$", "", stem, flags=re.I)
    return parse_text(stem)


def week_dates(week1_monday: dt.date, week: int) -> tuple[dt.date, dt.date]:
    start = week1_monday + dt.timedelta(weeks=week - 1)
    return start, start + dt.timedelta(days=6)


def week_number(week1_monday: dt.date, day: dt.date) -> int:
    return (day - week1_monday).days // 7 + 1
