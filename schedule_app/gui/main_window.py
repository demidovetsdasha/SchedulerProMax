"""Главное окно приложения."""
from __future__ import annotations

import datetime as dt
import os
from pathlib import Path

from PySide6.QtGui import QKeySequence, QShortcut
from PySide6.QtGui import QFont, QIcon
from PySide6.QtWidgets import (
    QAbstractItemView, QComboBox, QDateEdit, QFormLayout, QGroupBox,
    QHBoxLayout, QLabel, QLineEdit, QListWidget, QMainWindow, QMessageBox,
    QPlainTextEdit, QPushButton, QSpinBox, QVBoxLayout, QWidget,
)

from ..config import CONFIG_PATH, AppConfig, load_config
from ..excel_parser import ExcelSchedule, SheetSchedule
from ..filename_meta import FileMeta
from .build_mixin import BuildMixin
from .excel_mixin import ExcelMixin
from .helpers import qdate
from .pages_mixin import PagesMixin
from .styles import MODERN_STYLESHEET
import logging
from logging.handlers import RotatingFileHandler


class MainWindow(ExcelMixin, PagesMixin, BuildMixin, QMainWindow):
    def __init__(self):
        super().__init__()
        self._setup_file_logging() 
        self._setup_window()
        self._setup_state()
        self._build_ui()
        self._init_data()

        # Горячие клавиши для администратора
        QShortcut(QKeySequence("Ctrl+Shift+L"), self).activated.connect(self._toggle_console)
        QShortcut(QKeySequence("Ctrl+Shift+C"), self).activated.connect(self._open_config)


    # --------------------------------------------------------- инициализация
    def _setup_window(self) -> None:
        self.setWindowTitle("Нарезка расписания")
        self.resize(850, 850)
        self.setFont(QFont("Inter", 10))
        self.setStyleSheet(MODERN_STYLESHEET)
        self.setup_app_icon()

    def _setup_state(self) -> None:
        self.cfg: AppConfig = load_config()
        self.excel: ExcelSchedule | None = None
        self.meta = FileMeta()
        self.pages: list[list[str]] = [[]]
        self.last_out_dir: Path | None = None

    def _build_ui(self) -> None:
        central = QWidget()
        self.setCentralWidget(central)
        root = QVBoxLayout(central)
        root.setSpacing(10)
        root.setContentsMargins(15, 15, 15, 15)

        root.addWidget(self._build_file_box())
        root.addWidget(self._build_params_box())
        root.addWidget(self._build_pages_box(), stretch=1)
        root.addWidget(self._build_console())
        root.addLayout(self._build_actions_bar())
        root.addLayout(self._build_copyright_bar())

    def _init_data(self) -> None:
        self.fill_faculties()
        self.refresh_pages(select=0)
    
    def _setup_file_logging(self) -> None:
        """Пишем лог в config/app.log рядом с faculties.json (с ротацией)."""
        log_file = CONFIG_PATH.parent / "app.log"
        try:
            log_file.parent.mkdir(parents=True, exist_ok=True)
            handler = RotatingFileHandler(
                log_file, maxBytes=2_000_000, backupCount=3, encoding="utf-8"
            )
        except OSError:
            # Папка приложения недоступна на запись (например, Program Files)
            log_file = Path.home() / ".schedule_parser" / "app.log"
            log_file.parent.mkdir(parents=True, exist_ok=True)
            handler = RotatingFileHandler(
                log_file, maxBytes=2_000_000, backupCount=3, encoding="utf-8"
            )
        handler.setFormatter(logging.Formatter(
            "%(asctime)s [%(levelname)s] %(message)s",
            datefmt="%Y-%m-%d %H:%M:%S",
        ))

        self._file_logger = logging.getLogger("schedule_app")
        self._file_logger.setLevel(logging.DEBUG)
        # Не дублируем хэндлеры при повторном создании окна
        if not self._file_logger.handlers:
            self._file_logger.addHandler(handler)

        self._file_logger.info(f"=== Запуск приложения (файл: {log_file}) ===")

    # -------------------------------------------------- построение виджетов
    def _build_file_box(self) -> QGroupBox:
        box = QGroupBox("1. Файл расписания Excel")
        fl = QHBoxLayout(box)
        self.file_edit = QLineEdit()
        self.file_edit.setReadOnly(True)
        self.file_edit.setPlaceholderText("Загрузите файл .xlsx, .xlsm или .xls")
        btn_open = QPushButton("Открыть…")
        btn_open.clicked.connect(self.choose_excel)
        fl.addWidget(self.file_edit)
        fl.addWidget(btn_open)
        return box

    def _build_params_box(self) -> QGroupBox:
        box = QGroupBox("2. Параметры")
        form = QFormLayout(box)
        form.setVerticalSpacing(8)
        form.setHorizontalSpacing(15)

        self.sheet_combo = QComboBox()
        self.sheet_combo.currentIndexChanged.connect(self.on_sheet_changed)
        form.addRow("Лист:", self.sheet_combo)
        form.setRowVisible(self.sheet_combo, False)

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
        form.setRowVisible(self.week_spin, False)

        week = QHBoxLayout()
        self.date_start = QDateEdit(calendarPopup=True)
        self.date_end = QDateEdit(calendarPopup=True)
        for w in (self.date_start, self.date_end):
            w.setDisplayFormat("dd.MM.yyyy")
        today = dt.date.today()
        monday = today - dt.timedelta(days=today.weekday())
        self.date_start.setDate(qdate(monday))
        self.date_end.setDate(qdate(monday + dt.timedelta(days=6)))
        self.date_start.dateChanged.connect(
            lambda d: self.date_end.setDate(d.addDays(6)))

        week.addWidget(QLabel("с"))
        week.addSpacing(6)
        week.addWidget(self.date_start)
        week.addSpacing(12)
        week.addWidget(QLabel("по"))
        week.addSpacing(6)
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
        return box

    def _build_pages_box(self) -> QGroupBox:
        box = QGroupBox("3. Страницы документа и группы")
        pl = QHBoxLayout(box)

        left = QVBoxLayout()
        left.addWidget(QLabel("Страницы:"))
        self.pages_list = QListWidget()
        self.pages_list.setSelectionMode(QAbstractItemView.SingleSelection)
        self.pages_list.currentRowChanged.connect(self.on_page_selected)
        left.addWidget(self.pages_list)

        pbtns = QHBoxLayout()
        btn_add = QPushButton("Добавить ➕")
        btn_add.clicked.connect(self.add_page)
        btn_del = QPushButton("Удалить ❌")
        btn_del.clicked.connect(self.remove_page)
        btn_auto = QPushButton("Авторазбивка ⚡")
        btn_auto.setToolTip(
            f"Разложить все группы курса по страницам "
            f"(до {self.cfg.max_columns_per_page} на страницу)")
        btn_auto.clicked.connect(self.auto_split)
        pbtns.addWidget(btn_add)
        pbtns.addWidget(btn_del)
        pbtns.addWidget(btn_auto)
        left.addLayout(pbtns)
        pl.addLayout(left, stretch=3)

        right = QVBoxLayout()
        self.page_groups_label = QLabel("Группы на странице:")
        right.addWidget(self.page_groups_label)
        self.groups_list = QListWidget()
        self.groups_list.setFlow(QListWidget.LeftToRight)
        self.groups_list.setWrapping(True)
        self.groups_list.setResizeMode(QListWidget.Adjust)
        self.groups_list.setSpacing(4)
        self.groups_list.itemChanged.connect(self.on_group_toggled)
        right.addWidget(self.groups_list)

        gb = QHBoxLayout()
        gb.addStretch()
        btn_all = QPushButton("Выбрать все")
        btn_none = QPushButton("Отменить все")
        btn_all.clicked.connect(lambda: self.set_all_groups(True))
        btn_none.clicked.connect(lambda: self.set_all_groups(False))
        gb.addWidget(btn_all)
        gb.addWidget(btn_none)
        right.addLayout(gb)
        pl.addLayout(right, stretch=5)
        return box

    def _build_console(self) -> QPlainTextEdit:
        self.log = QPlainTextEdit()
        self.log.setObjectName("mini_console")
        self.log.setReadOnly(True)
        self.log.setPlaceholderText("Системные логи и сообщения будут выводиться здесь...")
        self.log.setMaximumHeight(120)
        self.log.setVisible(False)
        return self.log

    def _build_actions_bar(self) -> QHBoxLayout:
        actions = QHBoxLayout()

        btn_cfg = QPushButton("⚙️ Настройка факультетов и деканов")
        btn_cfg.setToolTip("Открыть файл настроек для редактирования списков и ФИО")
        btn_cfg.clicked.connect(self.open_and_auto_reload_config)
        btn_cfg.setVisible(False)

        self.btn_build = QPushButton("Сформировать Word ")
        self.btn_build.setObjectName("btn_build_word")
        self.btn_build.setDefault(True)
        self.btn_build.setMinimumHeight(38)
        self.btn_build.clicked.connect(self.build)

        actions.addWidget(btn_cfg)
        actions.addStretch()
        actions.addWidget(self.btn_build)
        return actions

    def _build_copyright_bar(self) -> QHBoxLayout:
        layout = QHBoxLayout()
        layout.addStretch()
        self.copyright_label = QLabel("© 2026 ЦИТ БГУИЯ | Отдел СИСиТП")
        self.copyright_label.setObjectName("copyright_label")
        layout.addWidget(self.copyright_label)
        return layout

    # ---------------------------------------------------------- вспомогательное
    def setup_app_icon(self) -> None:
        import sys

        if hasattr(sys, "_MEIPASS"):
            icon_path = Path(sys._MEIPASS) / "app_icon.ico"
        else:
            icon_path = Path(__file__).parent.parent.parent / "app_icon.ico"
            if not icon_path.exists():
                icon_path = Path(__file__).parent.parent / "app_icon.ico"
        if not icon_path.exists():
            exe_dir = Path(sys.executable).parent
            icon_path = exe_dir / "app_icon.ico"
        if icon_path.exists():
            self.setWindowIcon(QIcon(str(icon_path)))

    def log_msg(self, text: str) -> None:
        self.log.appendPlainText(text)
        if hasattr(self, "_file_logger"):
            # traceback идёт как ERROR, остальное как INFO
            if "Traceback" in text or "❌" in text:
                self._file_logger.error(text)
            else:
                self._file_logger.info(text)

    # ------------------------------------------------------- конфиг/факультеты
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

    def on_faculty_changed(self) -> None:
        fac = self.current_faculty()
        self.deputy_label.setText(fac.deputy_dean if fac else "—")

    def open_and_auto_reload_config(self) -> None:
        try:
            native_path = os.path.abspath(self.cfg.path)
            if os.name == "nt":
                os.startfile(native_path)
            else:
                import subprocess
                subprocess.call(("open", native_path))
            self.log_msg(
                "⚙️ Файл настроек открыт. После изменения и сохранения файла, "
                "откройте приложение заново.")
            self.cfg = load_config(self.cfg.path)
            self.fill_faculties(self.faculty_combo.currentData())
        except Exception as e:  # noqa: BLE001
            QMessageBox.critical(self, "Ошибка конфигурации",
                                 f"Не удалось открыть файл:\n{e}")

    def reload_config(self) -> None:
        try:
            self.cfg = load_config(self.cfg.path)
        except Exception as e:  # noqa: BLE001
            QMessageBox.critical(self, "Ошибка конфигурации", str(e))
            return
        self.fill_faculties(self.faculty_combo.currentData())
        self.log_msg(f"Конфиг перечитан: {self.cfg.path}")

    # ---------------------------------------------------------------- Excel data
    def current_sheet(self) -> SheetSchedule | None:
        if not self.excel:
            return None
        i = self.sheet_combo.currentIndex()
        return self.excel.sheets[i] if 0 <= i < len(self.excel.sheets) else None

    def course_columns(self) -> list[str]:
        sheet = self.current_sheet()
        if not sheet:
            return []
        course = str(self.course_spin.value())
        if not any(c.group.startswith(course) for c in sheet.columns):
            return [c.name for c in sheet.columns]
        # Группы с нецифровым названием (например, «ТянУИЯ») к курсу не привязаны — оставляем всегда
        cols = [c for c in sheet.columns
                if c.group.startswith(course) or not c.group[:1].isdigit()]
        return [c.name for c in cols]

    def sheet_order(self, names) -> list[str]:
        sheet = self.current_sheet()
        order = [c.name for c in sheet.columns] if sheet else []
        return sorted(set(names),
                      key=lambda n: order.index(n) if n in order else len(order))

    def _toggle_console(self) -> None:
        """Ctrl+Shift+L — показать/скрыть консоль (для администратора)."""
        self.log.setVisible(not self.log.isVisible())

    def _open_config(self) -> None:
        """Ctrl+Shift+C — открыть config/faculties.json (для администратора)."""
        self.open_and_auto_reload_config()