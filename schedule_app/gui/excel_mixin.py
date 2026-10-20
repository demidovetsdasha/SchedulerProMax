"""Логика работы с Excel: загрузка, разбор метаданных, даты недели."""
from __future__ import annotations

import datetime as dt
from pathlib import Path

from PySide6.QtWidgets import QApplication, QFileDialog, QMessageBox

from ..docx_builder import academic_year
from ..excel_parser import SheetSchedule, load_excel
from ..filename_meta import parse_filename, parse_text, week_dates, week_number
from .helpers import pydate, qdate
from .loading_dialog import LoadingDialog


class ExcelMixin:
    """Методы, связанные с Excel-файлом и автоопределением параметров."""

    def choose_excel(self) -> None:
        path, _ = QFileDialog.getOpenFileName(
            self, "Выберите файл расписания", "", "Excel (*.xlsx *.xlsm *.xls)")
        if path:
            self.load_excel(Path(path))

    def load_excel(self, path: Path) -> None:
        self.log_msg("⏳ Запуск парсера Excel... Проверка данных.")
        dialog = LoadingDialog(self, "Чтение данных из Excel листа...")
        try:
            path = path.resolve()
            QApplication.processEvents()
            if self.excel is None:
                dialog.show()
                QApplication.processEvents()
            self.excel = load_excel(path)
        except Exception as e:  # noqa: BLE001
            dialog.close()
            QMessageBox.critical(self, "Ошибка чтения Excel", str(e))
            return
        dialog.close()

        self.log_msg("✅ Файл успешно обработан и загружен!")
        self.file_edit.setText(path.name)

        if not self.last_out_dir:
            self.last_out_dir = path.parent

        self.sheet_combo.blockSignals(True)
        self.sheet_combo.clear()
        for s in self.excel.sheets:
            self.sheet_combo.addItem(s.sheet_name)
        self.sheet_combo.blockSignals(False)
        self.on_sheet_changed()

    def on_sheet_changed(self) -> None:
        sheet = self.current_sheet()
        if not sheet:
            return
        self.log_msg(f"Лист «{sheet.sheet_name}»: {sheet.title or 'без заголовка'}")
        self.meta = parse_filename(self.excel.path.name).fill_missing(
            parse_text(sheet.title))
        self.apply_file_meta(sheet)
        self.pages = [[]]
        self.refresh_pages(select=0)

    def apply_file_meta(self, sheet: SheetSchedule) -> None:
        m = self.meta
        found = []
        self.week_spin.blockSignals(True)
        self.week_spin.setValue(0)
        self.week_spin.blockSignals(False)

        fac = ((m.faculty_text and self.cfg.detect_faculty(m.faculty_text))
               or self.cfg.detect_faculty(self.excel.path.name, sheet.title))
        if fac:
            self.faculty_combo.setCurrentIndex(self.faculty_combo.findData(fac.code))
            found.append(f"факультет {fac.code}")
        else:
            self.log_msg("Факультет не определён — выберите вручную.")
        course = m.course or sheet.course
        if course:
            self.course_spin.blockSignals(True)
            self.course_spin.setValue(course)
            self.course_spin.blockSignals(False)
            found.append(f"курс {course}")
        if m.week:
            self.week_spin.setValue(m.week)
            found.append(f"неделя {m.week}")
        if m.date_start:
            self.set_dates(m.date_start, m.date_end)
            found.append(f"даты {m.date_start:%d.%m.%Y} – {m.date_end:%d.%m.%Y}")
            if not m.week and self.cfg.week1_monday:
                n = week_number(self.cfg.week1_monday, m.date_start)
                if 1 <= n <= self.week_spin.maximum():
                    self.week_spin.blockSignals(True)
                    self.week_spin.setValue(n)
                    self.week_spin.blockSignals(False)
        if found:
            self.log_msg("Определено: " + ", ".join(found))
        if not m.week and not m.date_start:
            self.log_msg("Неделя не найдена в имени файла — укажите даты вручную.")

    def set_dates(self, start: dt.date, end: dt.date) -> None:
        self.date_start.blockSignals(True)
        self.date_start.setDate(qdate(start))
        self.date_start.blockSignals(False)
        self.date_end.setDate(qdate(end))
        self.sync_academic_year()

    def sync_academic_year(self, *_):
        start = int(academic_year(pydate(self.date_start.date())).split("-")[0])
        self.year_spin.setValue(start)
        self.year_end_label.setText(f"– {start + 1}")

    def on_week_changed(self, week: int) -> None:
        if week and self.cfg.week1_monday:
            self.set_dates(*week_dates(self.cfg.week1_monday, week))