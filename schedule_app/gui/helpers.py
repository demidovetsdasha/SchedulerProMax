"""Вспомогательные функции для QDate ↔ datetime.date."""
from __future__ import annotations

import datetime as dt

from PySide6.QtCore import QDate


def qdate(d: dt.date) -> QDate:
    return QDate(d.year, d.month, d.day)


def pydate(q: QDate) -> dt.date:
    return dt.date(q.year(), q.month(), q.day())