"""Smoke-тест окна без показа на экране: файл → автоопределение → список групп и страницы."""
from __future__ import annotations

import os

import pytest

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
pytest.importorskip("PySide6")

from PySide6.QtCore import QEvent, Qt  # noqa: E402
from PySide6.QtWidgets import QApplication  # noqa: E402

from conftest import sample  # noqa: E402


@pytest.fixture(scope="module")
def qapp():
    return QApplication.instance() or QApplication([])


@pytest.fixture
def window(qapp, tmp_path, monkeypatch):
    from schedule_app import logging_setup
    from schedule_app.gui import main_window

    # не пишем в настоящий config/app.log
    real_setup = logging_setup.setup_file_logging
    monkeypatch.setattr(main_window, "setup_file_logging",
                        lambda: real_setup(tmp_path / "app.log"))
    win = main_window.MainWindow()
    yield win
    # удаляем окно через цикл событий Qt, иначе при выходе Qt ругается на порядок удаления
    win.close()
    win.deleteLater()
    QApplication.sendPostedEvents(None, QEvent.DeferredDelete)


def _group_items(win) -> list[str]:
    return [win.groups_list.item(i).data(Qt.UserRole) for i in range(win.groups_list.count())]


def test_window_opens(window):
    assert window.faculty_combo.count() == len(window.cfg.faculties)


def test_load_sample_detects_meta_and_shows_non_numeric_group(window):
    window.load_excel(sample("ФКЯиК"))
    assert window.current_faculty().code == "ФКЯиК"
    assert window.course_spin.value() == 2
    assert "ТянУИЯ" in _group_items(window)
    assert len(_group_items(window)) == 12


def test_auto_split_includes_non_numeric_group(window):
    window.load_excel(sample("ФКЯиК"))
    window.auto_split()
    all_names = [n for page in window.pages for n in page]
    assert "ТянУИЯ" in all_names
    assert len(all_names) == len(set(all_names)) == 12
    assert all(len(p) <= window.cfg.max_columns_per_page for p in window.pages)


def test_toggle_group_adds_it_to_page(window):
    window.load_excel(sample("ФКЯиК"))
    item = next(window.groups_list.item(i) for i in range(window.groups_list.count())
                if window.groups_list.item(i).data(Qt.UserRole) == "ТянУИЯ")
    item.setCheckState(Qt.Checked)
    assert window.pages[window.pages_list.currentRow()] == ["ТянУИЯ"]


def test_load_sample_with_dates(window):
    window.load_excel(sample("зимний"))
    assert window.current_faculty().code == "ФАЯ"
    assert window.course_spin.value() == 2
    assert window.date_start.date().toString("dd.MM.yyyy") == "10.11.2025"
    assert window.date_end.date().toString("dd.MM.yyyy") == "16.11.2025"
