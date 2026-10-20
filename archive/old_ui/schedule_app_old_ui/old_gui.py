"""Графический интерфейс (PySide6 / Qt) — работает на macOS и Windows."""
from __future__ import annotations

import datetime as dt
import traceback
from pathlib import Path

from PySide6.QtCore import QDate, Qt, QUrl
from PySide6.QtGui import QBrush, QColor, QDesktopServices
from PySide6.QtWidgets import (
    QAbstractItemView, QApplication, QComboBox, QDateEdit, QFileDialog, QFormLayout, QGroupBox,
    QHBoxLayout, QLabel, QLineEdit, QListWidget, QListWidgetItem, QMainWindow, QMessageBox,
    QPlainTextEdit, QPushButton, QSpinBox, QVBoxLayout, QWidget,
)

from .config import AppConfig, load_config
from .docx_builder import DocumentParams, academic_year, build_document, paginate
from .excel_parser import ExcelSchedule, SheetSchedule, load_excel
from .filename_meta import FileMeta, parse_filename, parse_text, week_dates, week_number


def _qdate(d: dt.date) -> QDate:
    return QDate(d.year, d.month, d.day)


def _pydate(q: QDate) -> dt.date:
    return dt.date(q.year(), q.month(), q.day())


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Парсер расписания Excel → Word")
        self.resize(820, 780)
        self.cfg: AppConfig = load_config()
        self.excel: ExcelSchedule | None = None
        self.meta = FileMeta()
        self.pages: list[list[str]] = [[]]   # наборы подгрупп по страницам документа

        central = QWidget()
        self.setCentralWidget(central)
        root = QVBoxLayout(central)

        # 1. Файл Excel
        file_box = QGroupBox("1. Файл расписания Excel")
        fl = QHBoxLayout(file_box)
        self.file_edit = QLineEdit()
        self.file_edit.setReadOnly(True)
        self.file_edit.setPlaceholderText("Выберите .xlsx или .xls файл")
        btn_open = QPushButton("Открыть…")
        btn_open.clicked.connect(self.choose_excel)
        fl.addWidget(self.file_edit)
        fl.addWidget(btn_open)
        root.addWidget(file_box)

        # 2. Параметры
        params_box = QGroupBox("2. Параметры")
        form = QFormLayout(params_box)
        self.sheet_combo = QComboBox()
        self.sheet_combo.currentIndexChanged.connect(self.on_sheet_changed)
        form.addRow("Лист:", self.sheet_combo)
        form.setRowVisible(self.sheet_combo, False)   # используется первый лист с расписанием

        self.faculty_combo = QComboBox()
        self.faculty_combo.currentIndexChanged.connect(self.on_faculty_changed)
        form.addRow("Факультет:", self.faculty_combo)
        self.deputy_label = QLabel()
        form.addRow("Зам. декана:", self.deputy_label)

        self.course_spin = QSpinBox()
        self.course_spin.setRange(1, 6)
        self.course_spin.valueChanged.connect(self.refresh_groups)
        form.addRow("Курс:", self.course_spin)

        self.week_spin = QSpinBox()
        self.week_spin.setRange(0, 60)
        self.week_spin.setSpecialValueText("—")
        self.week_spin.valueChanged.connect(self.on_week_changed)
        form.addRow("Неделя №:", self.week_spin)
        form.setRowVisible(self.week_spin, False)     # номер недели нужен только для расчёта дат

        week = QHBoxLayout()
        self.date_start = QDateEdit(calendarPopup=True)
        self.date_end = QDateEdit(calendarPopup=True)
        for w in (self.date_start, self.date_end):
            w.setDisplayFormat("dd.MM.yyyy")
        today = dt.date.today()
        monday = today - dt.timedelta(days=today.weekday())
        self.date_start.setDate(_qdate(monday))
        self.date_end.setDate(_qdate(monday + dt.timedelta(days=6)))
        self.date_start.dateChanged.connect(
            lambda d: self.date_end.setDate(d.addDays(6)))
        week.addWidget(QLabel("с"))
        week.addWidget(self.date_start)
        week.addWidget(QLabel("по"))
        week.addWidget(self.date_end)
        week.addStretch()
        form.addRow("Даты недели:", week)

        year_row = QHBoxLayout()
        self.year_spin = QSpinBox()
        self.year_spin.setRange(2000, 2100)
        self.year_end_label = QLabel()
        self.year_spin.valueChanged.connect(
            lambda y: self.year_end_label.setText(f"– {y + 1}"))
        year_row.addWidget(self.year_spin)
        year_row.addWidget(self.year_end_label)
        year_row.addStretch()
        form.addRow("Учебный год:", year_row)
        self.date_start.dateChanged.connect(self.sync_academic_year)
        self.sync_academic_year()
        root.addWidget(params_box)

        # 3. Страницы документа: на каждой странице — свой набор подгрупп
        pages_box = QGroupBox("3. Страницы документа и подгруппы")
        pl = QHBoxLayout(pages_box)

        left = QVBoxLayout()
        left.addWidget(QLabel("Страницы:"))
        self.pages_list = QListWidget()
        self.pages_list.setSelectionMode(QAbstractItemView.SingleSelection)
        self.pages_list.currentRowChanged.connect(self.on_page_selected)
        left.addWidget(self.pages_list)
        pbtns = QHBoxLayout()
        btn_add = QPushButton("+ Страница")
        btn_add.clicked.connect(self.add_page)
        btn_del = QPushButton("Удалить")
        btn_del.clicked.connect(self.remove_page)
        btn_auto = QPushButton("Авторазбивка")
        btn_auto.setToolTip(f"Разложить все подгруппы курса по страницам "
                            f"(до {self.cfg.max_columns_per_page} на страницу)")
        btn_auto.clicked.connect(self.auto_split)
        pbtns.addWidget(btn_add)
        pbtns.addWidget(btn_del)
        pbtns.addWidget(btn_auto)
        left.addLayout(pbtns)
        pl.addLayout(left, stretch=2)

        right = QVBoxLayout()
        self.page_groups_label = QLabel("Подгруппы на странице:")
        right.addWidget(self.page_groups_label)
        self.groups_list = QListWidget()
        self.groups_list.setFlow(QListWidget.LeftToRight)
        self.groups_list.setWrapping(True)
        self.groups_list.setResizeMode(QListWidget.Adjust)
        self.groups_list.setSpacing(4)
        self.groups_list.itemChanged.connect(self.on_group_toggled)
        right.addWidget(self.groups_list)
        gb = QHBoxLayout()
        btn_all = QPushButton("Отметить все")
        btn_none = QPushButton("Снять все")
        btn_all.clicked.connect(lambda: self.set_all_groups(True))
        btn_none.clicked.connect(lambda: self.set_all_groups(False))
        gb.addWidget(btn_all)
        gb.addWidget(btn_none)
        gb.addStretch()
        right.addLayout(gb)
        pl.addLayout(right, stretch=5)
        root.addWidget(pages_box, stretch=1)

        # 4. Сохранение
        out_box = QGroupBox("4. Результат")
        ol = QHBoxLayout(out_box)
        self.out_edit = QLineEdit()
        self.out_edit.setPlaceholderText("Папка для сохранения (по умолчанию — папка Excel файла)")
        btn_out = QPushButton("Папка…")
        btn_out.clicked.connect(self.choose_out_dir)
        ol.addWidget(self.out_edit)
        ol.addWidget(btn_out)
        root.addWidget(out_box)

        actions = QHBoxLayout()
        btn_cfg = QPushButton("Открыть конфиг")
        btn_cfg.clicked.connect(lambda: QDesktopServices.openUrl(QUrl.fromLocalFile(str(self.cfg.path))))
        btn_reload = QPushButton("Перечитать конфиг")
        btn_reload.clicked.connect(self.reload_config)
        self.btn_build = QPushButton("Сформировать Word")
        self.btn_build.setDefault(True)
        self.btn_build.setMinimumHeight(36)
        self.btn_build.clicked.connect(self.build)
        actions.addWidget(btn_cfg)
        actions.addWidget(btn_reload)
        actions.addStretch()
        actions.addWidget(self.btn_build)
        root.addLayout(actions)

        self.log = QPlainTextEdit(readOnly=True)
        self.log.setMaximumHeight(120)
        root.addWidget(self.log)

        self.fill_faculties()
        self.refresh_pages(select=0)

    # ---------------------------------------------------------------- helpers
    def log_msg(self, text: str) -> None:
        self.log.appendPlainText(text)

    def fill_faculties(self, select_code: str | None = None) -> None:
        self.faculty_combo.blockSignals(True)
        self.faculty_combo.clear()
        for fac in self.cfg.faculties:
            self.faculty_combo.addItem(f"{fac.code} — {fac.name}", fac.code)
        self.faculty_combo.blockSignals(False)
        if select_code:
            i = self.faculty_combo.findData(select_code)
            if i >= 0:
                self.faculty_combo.setCurrentIndex(i)
        self.on_faculty_changed()

    def current_faculty(self):
        return self.cfg.faculty_by_code(self.faculty_combo.currentData())

    def current_sheet(self) -> SheetSchedule | None:
        if not self.excel:
            return None
        i = self.sheet_combo.currentIndex()
        return self.excel.sheets[i] if 0 <= i < len(self.excel.sheets) else None

    def course_columns(self) -> list[str]:
        """Подгруппы листа для выбранного курса (в порядке Excel)."""
        sheet = self.current_sheet()
        if not sheet:
            return []
        course = str(self.course_spin.value())
        cols = [c for c in sheet.columns if c.group.startswith(course)] or sheet.columns
        return [c.name for c in cols]

    def sheet_order(self, names) -> list[str]:
        sheet = self.current_sheet()
        order = [c.name for c in sheet.columns] if sheet else []
        return sorted(set(names), key=lambda n: order.index(n) if n in order else len(order))

    # ------------------------------------------------------------- страницы
    @staticmethod
    def _page_title(i: int, names: list[str]) -> str:
        if not names:
            return f"Стр. {i + 1}: (нет подгрупп)"
        groups: list[str] = []
        for n in names:
            g = n.split("/")[0]
            if g not in groups:
                groups.append(g)
        return f"Стр. {i + 1}: {', '.join(groups)}  ({len(names)} подгр.)"

    def refresh_pages(self, select: int | None = None) -> None:
        row = self.pages_list.currentRow() if select is None else select
        self.pages_list.blockSignals(True)
        self.pages_list.clear()
        for i, names in enumerate(self.pages):
            self.pages_list.addItem(self._page_title(i, names))
        self.pages_list.blockSignals(False)
        self.pages_list.setCurrentRow(max(0, min(row, len(self.pages) - 1)))
        self.refresh_groups()

    def update_page_titles(self) -> None:
        for i, names in enumerate(self.pages):
            self.pages_list.item(i).setText(self._page_title(i, names))

    def on_page_selected(self, row: int) -> None:
        self.refresh_groups()

    def add_page(self) -> None:
        self.pages.append([])
        self.refresh_pages(select=len(self.pages) - 1)

    def remove_page(self) -> None:
        row = self.pages_list.currentRow()
        if row < 0:
            return
        del self.pages[row]
        if not self.pages:
            self.pages = [[]]
        self.refresh_pages(select=min(row, len(self.pages) - 1))

    def auto_split(self) -> None:
        sheet = self.current_sheet()
        if not sheet:
            return
        names = set(self.course_columns())
        cols = [c for c in sheet.columns if c.name in names]
        self.pages = [[c.name for c in page]
                      for page in paginate(cols, max(1, self.cfg.max_columns_per_page))] or [[]]
        self.refresh_pages(select=0)

    # -------------------------------------------------------------- подгруппы
    def refresh_groups(self) -> None:
        """Чекбоксы подгрупп для текущей страницы; занятые на других страницах помечены."""
        page = self.pages_list.currentRow()
        self.page_groups_label.setText(
            f"Подгруппы на странице {page + 1}:" if page >= 0 else "Подгруппы на странице:")
        used_elsewhere = {
            n: i + 1 for i, names in enumerate(self.pages) if i != page for n in names}
        current = set(self.pages[page]) if 0 <= page < len(self.pages) else set()
        self.groups_list.blockSignals(True)
        self.groups_list.clear()
        for name in self.course_columns():
            other = used_elsewhere.get(name)
            item = QListWidgetItem(f"{name}  (стр. {other})" if other else name)
            item.setData(Qt.UserRole, name)
            item.setFlags(item.flags() | Qt.ItemIsUserCheckable)
            item.setCheckState(Qt.Checked if name in current else Qt.Unchecked)
            if other:
                item.setForeground(QBrush(QColor("#888888")))
            self.groups_list.addItem(item)
        self.groups_list.blockSignals(False)

    def on_group_toggled(self, item: QListWidgetItem) -> None:
        page = self.pages_list.currentRow()
        if not 0 <= page < len(self.pages):
            return
        name = item.data(Qt.UserRole)
        names = set(self.pages[page])
        if item.checkState() == Qt.Checked:
            names.add(name)
        else:
            names.discard(name)
        self.pages[page] = self.sheet_order(names)
        self.update_page_titles()

    def set_all_groups(self, checked: bool) -> None:
        page = self.pages_list.currentRow()
        if not 0 <= page < len(self.pages):
            return
        visible = set(self.course_columns())
        names = set(self.pages[page])
        names = names | visible if checked else names - visible
        self.pages[page] = self.sheet_order(names)
        self.update_page_titles()
        self.refresh_groups()

    # ---------------------------------------------------------------- slots
    def choose_excel(self) -> None:
        path, _ = QFileDialog.getOpenFileName(
            self, "Выберите файл расписания", "", "Excel (*.xlsx *.xlsm *.xls)")
        if path:
            self.load_excel(Path(path))

    def load_excel(self, path: Path) -> None:
        try:
            path = path.resolve()
            self.excel = load_excel(path)
        except Exception as e:  # noqa: BLE001
            QMessageBox.critical(self, "Ошибка чтения Excel", str(e))
            return
        self.file_edit.setText(str(path))
        if not self.out_edit.text():
            self.out_edit.setText(str(path.parent))
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
        # имя файла — основной источник; недостающее берём из заголовка листа
        self.meta = parse_filename(self.excel.path.name).fill_missing(parse_text(sheet.title))
        self.apply_file_meta(sheet)
        self.pages = [[]]
        self.refresh_pages(select=0)

    def apply_file_meta(self, sheet: SheetSchedule) -> None:
        """Заполняет факультет, курс, неделю и даты (поля остаются редактируемыми)."""
        m = self.meta
        found = []
        self.week_spin.blockSignals(True)
        self.week_spin.setValue(0)          # не оставляем неделю от предыдущего файла
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
            self.week_spin.setValue(m.week)       # пересчитает даты, если задано week1_monday
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
        self.date_start.setDate(_qdate(start))
        self.date_start.blockSignals(False)
        self.date_end.setDate(_qdate(end))
        self.sync_academic_year()

    def sync_academic_year(self, *_):
        """Учебный год по дате начала недели (сентябрь–август); можно изменить вручную."""
        start = int(academic_year(_pydate(self.date_start.date())).split("-")[0])
        self.year_spin.setValue(start)
        self.year_end_label.setText(f"– {start + 1}")

    def on_week_changed(self, week: int) -> None:
        if week and self.cfg.week1_monday:
            self.set_dates(*week_dates(self.cfg.week1_monday, week))

    def on_faculty_changed(self) -> None:
        fac = self.current_faculty()
        self.deputy_label.setText(fac.deputy_dean if fac else "—")

    def choose_out_dir(self) -> None:
        d = QFileDialog.getExistingDirectory(self, "Папка для сохранения", self.out_edit.text())
        if d:
            self.out_edit.setText(d)

    def reload_config(self) -> None:
        try:
            self.cfg = load_config(self.cfg.path)
        except Exception as e:  # noqa: BLE001
            QMessageBox.critical(self, "Ошибка конфигурации", str(e))
            return
        self.fill_faculties(self.faculty_combo.currentData())
        self.log_msg(f"Конфиг перечитан: {self.cfg.path}")

    def build(self) -> None:
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

        start, end = _pydate(self.date_start.date()), _pydate(self.date_end.date())
        course = self.course_spin.value()
        out_dir = Path(self.out_edit.text() or self.excel.path.parent)
        default_name = f"{fac.code} {course} курс с {start:%d.%m} - {end:%d.%m}.docx"
        path, _ = QFileDialog.getSaveFileName(
            self, "Сохранить расписание", str(out_dir / default_name), "Word (*.docx)")
        if not path:
            return
        out = Path(path)
        if out.suffix.lower() != ".docx":
            out = out.with_suffix(".docx")

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
            QMessageBox.critical(self, "Ошибка", f"Не удалось записать файл (возможно, он открыт в Word):\n{out}")
            return
        except Exception as e:  # noqa: BLE001
            self.log_msg(traceback.format_exc())
            QMessageBox.critical(self, "Ошибка формирования", str(e))
            return

        self.log_msg(f"Готово: {out}")
        for i, (names, size) in enumerate(zip(pages, sizes), 1):
            self.log_msg(f"  стр. {i}: {', '.join(names)} — шрифт таблицы {size:g} pt")
        if any(s <= self.cfg.table_font_min for s in sizes):
            self.log_msg("Внимание: достигнут минимальный шрифт — проверьте, помещается ли таблица.")
        answer = QMessageBox.question(self, "Готово", f"Файл сохранён:\n{out}\n\nОткрыть его?")
        if answer == QMessageBox.Yes:
            QDesktopServices.openUrl(QUrl.fromLocalFile(str(out)))


def run() -> int:
    import sys

    app = QApplication(sys.argv)
    app.setApplicationName("ScheduleParser")
    win = MainWindow()
    win.show()
    return app.exec()
