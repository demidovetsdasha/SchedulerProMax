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
    from schedule_app.gui.theme import apply_light_theme
    app = QApplication.instance() or QApplication([])
    apply_light_theme(app)
    return app


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


# -------------------------------------------------------------------- тема

def test_theme_ignores_os_dark_mode():
    """Тёмная палитра ОС не должна попадать в окно (белый текст на белом фоне)."""
    from PySide6.QtGui import QColor, QPalette

    from schedule_app.gui.theme import apply_light_theme
    app = QApplication.instance() or QApplication([])
    dark = QPalette()
    for role in (QPalette.Window, QPalette.Base, QPalette.Button):
        dark.setColor(role, QColor("#202020"))
    for role in (QPalette.WindowText, QPalette.Text, QPalette.ButtonText):
        dark.setColor(role, QColor("#FFFFFF"))
    app.setPalette(dark)

    apply_light_theme(app)

    pal = app.palette()
    assert app.style().name().lower() == "fusion"
    for role in (QPalette.Window, QPalette.Base, QPalette.Button):
        assert pal.color(role).lightness() > 200, role
    for role in (QPalette.WindowText, QPalette.Text, QPalette.ButtonText):
        assert pal.color(role).lightness() < 100, role


def test_stylesheet_icons_resolved(qapp):
    import re
    from pathlib import Path

    from schedule_app.gui.styles import MODERN_STYLESHEET
    from schedule_app.gui.theme import themed_stylesheet
    qss = themed_stylesheet(MODERN_STYLESHEET)
    assert "%" not in re.sub(r"\d+%", "", qss), "в QSS осталась неподставленная иконка"
    urls = re.findall(r"url\(([^)]+)\)", qss)
    assert urls
    for url in urls:
        assert Path(url).is_file(), url


def test_checked_group_indicator_is_visible(window):
    """Отмеченный чекбокс должен заметно отличаться от неотмеченного (на Win11 был белым на белом)."""
    window.resize(900, 850)
    window.show()
    window.load_excel(sample("ФКЯиК"))
    lst = window.groups_list
    lst.item(0).setCheckState(Qt.Checked)
    QApplication.processEvents()
    img = lst.viewport().grab().toImage()

    def indicator_color(row):
        rect = lst.visualItemRect(lst.item(row))
        # индикатор — слева в элементе; ищем самый насыщенный пиксель в левой части
        best = None
        for x in range(rect.left(), rect.left() + 30):
            for y in range(rect.top(), rect.bottom()):
                c = img.pixelColor(x, y)
                if best is None or c.saturation() > best.saturation():
                    best = c
        return best

    checked, unchecked = indicator_color(0), indicator_color(1)
    assert checked.green() > checked.red() + 40, "отмеченный чекбокс не закрашен зелёным"
    assert unchecked.saturation() < 60, "неотмеченный чекбокс не должен быть закрашен"
