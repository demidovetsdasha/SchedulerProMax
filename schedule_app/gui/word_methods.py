"""Формирование итогового документа Word."""
from __future__ import annotations

import datetime as dt
import traceback
from pathlib import Path

from PySide6.QtCore import QUrl
from PySide6.QtGui import QDesktopServices
from PySide6.QtWidgets import QFileDialog, QMessageBox

from ..config import load_config
from ..docx_builder import DocumentParams, build_document
from .helpers import pydate


class BuildMixin:
    def build(self) -> None:
        # Автообновление конфига перед генерацией
        try:
            self.cfg = load_config(self.cfg.path)
        except Exception:  # noqa: BLE001
            pass

        sheet = self.current_sheet()
        fac = self.current_faculty()

        pages = [names for names in self.pages if names]
        if not sheet:
            QMessageBox.warning(self, "Нет данных", "Сначала откройте Excel файл.")
            return
        if not fac:
            QMessageBox.warning(self, "Факультет", "Выберите факультет.")
            return
        if not pages:
            QMessageBox.warning(self, "Подгруппы",
                                "Отметьте подгруппы хотя бы для одной страницы.")
            return
        if not self.cfg.template.exists():
            QMessageBox.critical(self, "Шаблон", f"Не найден шаблон: {self.cfg.template}")
            return

        start, end = pydate(self.date_start.date()), pydate(self.date_end.date())
        course = self.course_spin.value()

        if self.last_out_dir and self.last_out_dir.exists():
            start_dir = self.last_out_dir
        else:
            start_dir = self.excel.path.parent if self.excel else Path.home()

        default_name = f"{fac.code} {course} курс с {start:%d.%m} - {end:%d.%m}.docx"
        path, _ = QFileDialog.getSaveFileName(
            self, "Сохранить расписание", str(start_dir / default_name), "Word (*.docx)"
        )
        if not path:
            return

        out = Path(path)
        if out.suffix.lower() != ".docx":
            out = out.with_suffix(".docx")

        self.last_out_dir = out.parent

        params = DocumentParams(
            course=course,
            faculty_header=fac.header_name,
            date_start=start,
            date_end=end,
            head_ckop_title=self.cfg.head_ckop_title,
            head_ckop_name=self.cfg.head_ckop_name,
            deputy_dean=fac.deputy_dean,
            max_columns_per_page=self.cfg.max_columns_per_page,
            academic_year=f"{self.year_spin.value()}-{self.year_spin.value() + 1}",
            approval_year=dt.date.today().year,
            table_font_max=self.cfg.table_font_max,
            table_font_min=self.cfg.table_font_min,
        )
        try:
            sizes = build_document(self.cfg.template, sheet, pages, params, out)
        except PermissionError:
            QMessageBox.critical(
                self, "Ошибка",
                f"Не удалось записать файл (возможно, он открыт в Word):\n{out}")
            return
        except Exception as e:  # noqa: BLE001
            self.log_msg(traceback.format_exc())
            QMessageBox.critical(self, "Ошибка формирования", str(e))
            return

        self.log_msg(f"Готово: {out}")
        for i, (names, size) in enumerate(zip(pages, sizes), 1):
            self.log_msg(f"  стр. {i}: {', '.join(names)} — шрифт таблицы {size:g} pt")
        if any(s <= self.cfg.table_font_min for s in sizes):
            self.log_msg(
                "Внимание: достигнут минимальный шрифт — проверьте, помещается ли таблица.")

        msg_box = QMessageBox(self)
        msg_box.setWindowTitle("Готово")
        msg_box.setText(f"Файл успешно сохранён:\n{out.name}\n\nОткрыть созданный документ?")
        msg_box.setIcon(QMessageBox.Question)

        btn_yes = msg_box.addButton("Да", QMessageBox.YesRole)
        msg_box.addButton("Нет", QMessageBox.NoRole)
        msg_box.setDefaultButton(btn_yes)
        msg_box.exec()

        if msg_box.clickedButton() == btn_yes:
            QDesktopServices.openUrl(QUrl.fromLocalFile(str(out)))