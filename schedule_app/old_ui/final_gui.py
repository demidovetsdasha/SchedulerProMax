"""Графический интерфейс (PySide6 / Qt) — работает на macOS и Windows."""
from __future__ import annotations

import datetime as dt
import traceback
from pathlib import Path
import time
import os



from PySide6.QtCore import QDate, Qt, QUrl
from PySide6.QtGui import QBrush, QColor, QDesktopServices, QFont, QIcon
from PySide6.QtWidgets import (
    QAbstractItemView, QApplication, QComboBox, QDateEdit, QFileDialog, QFormLayout, QGroupBox,
    QHBoxLayout, QLabel, QLineEdit, QListWidget, QListWidgetItem, QMainWindow, QMessageBox,
    QPlainTextEdit, QPushButton, QSpinBox, QVBoxLayout, QWidget, QDialog, QProgressBar
)

from .config import AppConfig, load_config
from .docx_builder import DocumentParams, academic_year, build_document, paginate
from .excel_parser import ExcelSchedule, SheetSchedule, load_excel
from .filename_meta import FileMeta, parse_filename, parse_text, week_dates, week_number


def _qdate(d: dt.date) -> QDate:
    return QDate(d.year, d.month, d.day)


def _pydate(q: QDate) -> dt.date:
    return dt.date(q.year(), q.month(), q.day())


# Современные QSS Стили (Белый + Зеленый + Темная консоль)
MODERN_STYLESHEET = """
    QMainWindow {
        background-color: #F8F9FA;
    }
    
    /* Стилизация контейнеров Групп-боксов */
    QGroupBox {
        font-weight: bold;
        border: 1px solid #EAEAEA;
        border-radius: 8px;
        margin-top: 12px;
        background-color: #FFFFFF;
        padding: 12px;
    }
    QGroupBox::title {
        subcontrol-origin: margin;
        subcontrol-position: top left;
        left: 10px;
        padding: 0 5px;
        color: #2D3748;
    }

    /* Поля ввода, выпадающие списки и числовые поля (SpinBox) */
    QLineEdit, QComboBox, QAbstractSpinBox, QDateEdit {
        border: 1px solid #CBD5E0;
        border-radius: 6px;
        padding: 6px 12px;       /* Увеличиваем внутренний отступ для всех полей */
        background-color: #FFFFFF;
        color: #2D3748;
        selection-background-color: #2F855A;
    }
    
    /* Принудительный внутренний отступ для текста внутри числовых полей */
    QAbstractSpinBox::up-down-key-pressed, 
    QAbstractSpinBox QLineEdit {
        padding: 0px 4px;        /* Дополнительный отступ слева для цифр курса и года */
        background: transparent;
        border: none;
    }

    QLineEdit:focus, QComboBox:focus, QAbstractSpinBox:focus, QDateEdit:focus {
        border: 1px solid #48BB78;
    }

    
    /* Выпадающие календари */
    QDateEdit {
        padding-right: 20px; /* Оставляем место справа для системной стрелочки */
    }

    
    /* Заменяем иконку текстовым символом стрелочки */
    QDateEdit::arrow {
        image: none;
    }

    /* Стиль всплывающего виджета календаря */
    QCalendarWidget QWidget {
        background-color: #FFFFFF;
    }
    QCalendarWidget QNodeWidget {
        background-color: #FFFFFF;
        color: #2D3748;
    }
    QCalendarWidget QAbstractItemView:enabled {
        color: #2D3748;
        selection-background-color: #48BB78; /* Зеленое выделение выбранного дня */
        selection-color: white;
    }
    QCalendarWidget QToolButton {
        color: #2D3748;
        background-color: transparent;
    }
    QCalendarWidget QToolButton:hover {
        background-color: #EDF2F7;
        border-radius: 4px;
    }


    /* Кнопки */
    QPushButton {
        background-color: #EDF2F7;
        color: #2D3748;
        border: 1px solid #E2E8F0;
        border-radius: 6px;
        padding: 6px 16px;
        font-weight: 500;
    }
    QPushButton:hover {
        background-color: #E2E8F0;
        border-color: #CBD5E0;
    }
    QPushButton:pressed {
        background-color: #CBD5E0;
    }
    
    /* Главная акцентная кнопка (Сформировать Word) */
    QPushButton#btn_build_word {
        background-color: #38A169;
        color: #FFFFFF;
        border: none;
        font-weight: bold;
    }
    QPushButton#btn_build_word:hover {
        background-color: #2F855A;
    }
    QPushButton#btn_build_word:pressed {
        background-color: #22543D;
    }

    /* Списки (Страницы и группы) */
    QListWidget {
        border: 1px solid #E2E8F0;
        border-radius: 6px;
        background-color: #FFFFFF;
        padding: 4px;
    }
    QListWidget::item {
        padding: 6px;
        border-radius: 4px;
        margin: 2px 0;
    }
    QListWidget::item:hover {
        background-color: #F0FFF4;
        color: #22543D;
    }
    QListWidget::item:selected {
        background-color: #C6F6D5;
        color: #22543D;
    }

    /* Подписи */
    QLabel {
        color: #4A5568;
    }

    /* ТЕМНАЯ МИНИ-КОНСОЛЬ ВНИЗУ */
    QPlainTextEdit#mini_console {
        background-color: #1A202C;
        color: #A0AEC0;
        font-family: 'Consolas', 'Courier New', monospace;
        font-size: 12px;
        border: 1px solid #2D3748;
        border-radius: 6px;
        padding: 8px;
    }

    /* Окно загрузки (Сплеш-скрин) */
    QDialog#loading_dialog {
        background-color: #FFFFFF;
        border: 2px solid #48BB78;
        border-radius: 12px;
    }
    QLabel#loading_title {
        color: #2D3748;
        font-size: 16px;
        font-weight: bold;
    }
    QLabel#loading_subtitle {
        color: #718096;
        font-size: 13px;
    }

        /* Стилизация всплывающих окон QMessageBox */
    QMessageBox {
        background-color: #FFFFFF;
    }
    QMessageBox QLabel {
        color: #2D3748;
        font-size: 13px;
    }
    QMessageBox QPushButton {
        background-color: #EDF2F7;
        color: #2D3748;
        border: 1px solid #E2E8F0;
        border-radius: 6px;
        padding: 5px 14px;
        min-width: 65px;
    }
    QMessageBox QPushButton:hover {
        background-color: #E2E8F0;
    }
    /* Делаем кнопку подтверждения (Да) слегка зеленоватой */
    QMessageBox QPushButton:default {
        background-color: #E6FFFA;
        border: 1px solid #48BB78;
        color: #22543D;
    }
    QMessageBox QPushButton:default:hover {
        background-color: #C6F6D5;
    }

    /* Стиль для текста копирайта в самом низу */
    QLabel#copyright_label {
        color: #718096;          /* Мягкий серый цвет текста */
        font-size: 11px;         /* Уменьшенный размер шрифта */
        font-weight: normal;
        padding-top: 5px;        /* Легкий отступ от верхних кнопок */
    }



"""

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        
        # 1. Изменяем название приложения
        self.setWindowTitle("Нарезка расписания")
        self.resize(850, 850)
        
        # 2. Настраиваем шрифт Inter для всего приложения
        app_font = QFont("Inter", 10)
        self.setFont(app_font)
        
        # 3. Применяем кастомные стили
        self.setStyleSheet(MODERN_STYLESHEET)

        # 4. Заготовка для иконки приложения
        self.setup_app_icon()

        self.cfg: AppConfig = load_config()
        self.excel: ExcelSchedule | None = None
        self.meta = FileMeta()
        self.pages: list[list[str]] = [[]]   
        self.last_out_dir: Path | None = None 

        central = QWidget()
        self.setCentralWidget(central)
        root = QVBoxLayout(central)
        root.setSpacing(10)
        root.setContentsMargins(15, 15, 15, 15)

        # 1. Файл Excel
        file_box = QGroupBox("1. Файл расписания Excel")
        fl = QHBoxLayout(file_box)
        
        self.file_edit = QLineEdit()
        self.file_edit.setReadOnly(True)
        # Меняем текст подсказки по умолчанию
        self.file_edit.setPlaceholderText("Загрузите файл .xlsx, .xlsm или .xls") 
        
        btn_open = QPushButton("Открыть…")
        btn_open.clicked.connect(self.choose_excel)
        fl.addWidget(self.file_edit)
        fl.addWidget(btn_open)
        root.addWidget(file_box)


        # 2. Параметры
        params_box = QGroupBox("2. Параметры")
        form = QFormLayout(params_box)
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
        self.date_start.setDate(_qdate(monday))
        self.date_end.setDate(_qdate(monday + dt.timedelta(days=6)))
        self.date_start.dateChanged.connect(
            lambda d: self.date_end.setDate(d.addDays(6)))
        
        week.addWidget(QLabel("с"))
        week.addSpacing(6)                # Добавляем 6 пикселей пустого пространства после "с"
        week.addWidget(self.date_start)
        week.addSpacing(12)               # Добавляем 12 пикселей пространства перед "по"
        week.addWidget(QLabel("по"))
        week.addSpacing(6)                # Добавляем 6 пикселей пустого пространства после "по"
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

        # 3. Страницы документа
        pages_box = QGroupBox("3. Страницы документа и подгруппы")
        pl = QHBoxLayout(pages_box)

        # Левая колонка: Страницы
        left = QVBoxLayout()
        left.addWidget(QLabel("Страницы:"))
        self.pages_list = QListWidget()
        self.pages_list.setSelectionMode(QAbstractItemView.SingleSelection)
        self.pages_list.currentRowChanged.connect(self.on_page_selected)
        left.addWidget(self.pages_list)
        
        # Кнопки для страниц с эмодзи-иконками
        pbtns = QHBoxLayout()
        btn_add = QPushButton("Добавить ➕")
        btn_add.clicked.connect(self.add_page)
        btn_del = QPushButton("Удалить ❌")
        btn_del.clicked.connect(self.remove_page)
        btn_auto = QPushButton("Авторазбивка ⚡")
        btn_auto.setToolTip(f"Разложить все подгруппы курса по страницам "
                            f"(до {self.cfg.max_columns_per_page} на страницу)")
        btn_auto.clicked.connect(self.auto_split)
        pbtns.addWidget(btn_add)
        pbtns.addWidget(btn_del)
        pbtns.addWidget(btn_auto)
        left.addLayout(pbtns)
        pl.addLayout(left, stretch=3)

        # Правая колонка: Подгруппы
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
        
        # Кнопки для подгрупп: меняем названия, добавляем эмодзи и выравниваем вправо
        gb = QHBoxLayout()
        gb.addStretch() # Этот сдвиг толкает кнопки к правому краю, отделяя от левого блока
        btn_all = QPushButton("Выбрать все")
        btn_none = QPushButton("Отменить все")
        btn_all.clicked.connect(lambda: self.set_all_groups(True))
        btn_none.clicked.connect(lambda: self.set_all_groups(False))
        gb.addWidget(btn_all)
        gb.addWidget(btn_none)
        right.addLayout(gb)
        pl.addLayout(right, stretch=5)
        
        root.addWidget(pages_box, stretch=1)


        # 5. МИНИ-КОНСОЛЬ (Вынесена отдельно, визуально отличается от всего интерфейса)
        self.log = QPlainTextEdit()
        self.log.setObjectName("mini_console")
        self.log.setReadOnly(True)
        self.log.setPlaceholderText("Системные логи и сообщения будут выводиться здесь...")
        self.log.setMaximumHeight(120)
        root.addWidget(self.log)

        # 6. Кнопки действий
        actions = QHBoxLayout()
        
        # Оставляем одну понятную кнопку вместо двух старых
        btn_cfg = QPushButton("⚙️ Настройка факультетов и деканов")
        btn_cfg.setToolTip("Открыть файл настроек для редактирования списков и ФИО")
        btn_cfg.clicked.connect(self.open_and_auto_reload_config)
        
        self.btn_build = QPushButton("Сформировать Word ")
        self.btn_build.setObjectName("btn_build_word")  # Привязка к акцентному стилю
        self.btn_build.setDefault(True)
        self.btn_build.setMinimumHeight(38)
        self.btn_build.clicked.connect(self.build)
        
        actions.addWidget(btn_cfg)
        actions.addStretch()
        actions.addWidget(self.btn_build)
        root.addLayout(actions)

        # ДОБАВЛЯЕМ БЛОК КОПИРАЙТА:
        copyright_layout = QHBoxLayout()
        copyright_layout.addStretch()  # Сдвигаем текст к правому краю (или уберите, если нужно по центру)
        
        self.copyright_label = QLabel("© 2026 ЦИТ БГУИЯ | Отдел СИСиТП")
        self.copyright_label.setObjectName("copyright_label")  # Привязываем к стилю CSS
        
        copyright_layout.addWidget(self.copyright_label)
        root.addLayout(copyright_layout)



        # Первичная инициализация данных
        self.fill_faculties()
        self.refresh_pages(select=0)

    # ---------------------------------------------------------------- helpers
    def setup_app_icon(self):
        """Подключает иконку приложения к окну, учитывая сборку PyInstaller."""
        import sys
        
        # 1. Если запущено как скомпилированный .exe
        if hasattr(sys, '_MEIPASS'):
            icon_path = Path(sys._MEIPASS) / "app_icon.ico"
        else:
            # 2. Если запускаем в VS Code / Python
            icon_path = Path(__file__).parent.parent / "app_icon.ico"
            if not icon_path.exists():
                icon_path = Path(__file__).parent / "app_icon.ico"
                
        # 3. Резервный вариант: ищем иконку в одной папке РЯДОМ с файлом .exe
        if not icon_path.exists():
            exe_dir = Path(sys.executable).parent
            icon_path = exe_dir / "app_icon.ico"

        if icon_path.exists():
            self.setWindowIcon(QIcon(str(icon_path)))



    def create_loading_screen(self, text: str = "Обработка файла..."):
        """Создает красивое модальное окно загрузки, блокирующее интерфейс."""
        self.loading_dialog = QDialog(self)
        self.loading_dialog.setObjectName("loading_dialog")
        # Убираем стандартную рамку Windows, оставляя чистый прямоугольник
        self.loading_dialog.setWindowFlags(Qt.Window | Qt.CustomizeWindowHint | Qt.WindowTitleHint)
        self.loading_dialog.setModal(True)
        self.loading_dialog.setFixedSize(320, 140)

        layout = QVBoxLayout(self.loading_dialog)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(10)

        title = QLabel("Пожалуйста, подождите ⏳")
        title.setObjectName("loading_title")
        title.setAlignment(Qt.AlignCenter)
        
        self.loading_label = QLabel(text)
        self.loading_label.setObjectName("loading_subtitle")
        self.loading_label.setAlignment(Qt.AlignCenter)

        # Создаем бесконечный индикатор загрузки (бегающая линия)
        progress = QProgressBar()
        progress.setRange(0, 0) 
        progress.setTextVisible(False)
        progress.setStyleSheet("""
            QProgressBar {
                border: 1px solid #E2E8F0;
                border-radius: 4px;
                background-color: #EDF2F7;
                height: 8px;
            }
            QProgressBar::chunk {
                background-color: #48BB78;
                border-radius: 4px;
            }
        """)

        layout.addWidget(title)
        layout.addWidget(self.loading_label)
        layout.addWidget(progress)


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
        self.log_msg("⏳ Запуск парсера Excel... Проверка данных.")
        
        # 1. Замеряем точное время начала обработки файла
        start_time = time.time()
        
        try:
            path = path.resolve()
            
            # Если это первая (тяжелая) загрузка, процесс займет около 7 секунд.
            # Чтобы интерфейс успел отреагировать, мы запускаем чтение,
            # но параллельно проверяем, сколько времени прошло.
            
            # Для этого принудительно даем интерфейсу "вздохнуть"
            QApplication.processEvents()
            
            # Если мы боимся зависания на первой секунде, мы можем показать окно
            # только если предыдущий файл показал, что система работает медленно.
            # Но проще сделать так: создаем окно, но НЕ показываем его сразу.
            self.create_loading_screen("Чтение данных из Excel листа...")
            
            # Небольшой лайфхак: если мы знаем, что это ПЕРВЫЙ запуск (self.excel еще None),
            # то принудительно сразу выводим окно, так как оно точно будет долгим.
            if self.excel is None:
                self.loading_dialog.show()
                QApplication.processEvents()
            
            # Запускаем само чтение Excel-файла
            self.excel = load_excel(path)
            
        except Exception as e:  # noqa: BLE001
            if hasattr(self, 'loading_dialog'):
                self.loading_dialog.close()
            QMessageBox.critical(self, "Ошибка чтения Excel", str(e))
            return
            
        # 2. Закрываем окно, если оно было выведено на экран
        if hasattr(self, 'loading_dialog'):
            self.loading_dialog.close()
            
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
        self.meta = parse_filename(self.excel.path.name).fill_missing(parse_text(sheet.title))
        self.apply_file_meta(sheet)
        self.pages = [[]]
        self.refresh_pages(select=0)

    def apply_file_meta(self, sheet: SheetSchedule) -> None:
        """Заполняет факультет, курс, неделю и даты (поля остаются редактируемыми)."""
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

    def open_and_auto_reload_config(self) -> None:
        """Открывает json файл настроек в системном редакторе и обновляет данные."""
        try:
            # Превращаем путь в стандартный для текущей ОС вид (с правильными слэшами)
            native_path = os.path.abspath(self.cfg.path)
            
            # Безопасный запуск редактора по умолчанию в зависимости от операционной системы
            if os.name == 'nt':  # Если это Windows
                os.startfile(native_path)
            else:                # Если это macOS или Linux
                import subprocess
                subprocess.call(('open', native_path))
                
            self.log_msg("⚙️ Файл настроек открыт. После изменения и сохранения файла, откройте приложение заново.")
            
            # Перечитываем состояние файла для интерфейса
            self.cfg = load_config(self.cfg.path)
            self.fill_faculties(self.faculty_combo.currentData())
        except Exception as e:
            QMessageBox.critical(self, "Ошибка конфигурации", f"Не удалось открыть файл:\n{e}")



    def reload_config(self) -> None:
        try:
            self.cfg = load_config(self.cfg.path)
        except Exception as e:  # noqa: BLE001
            QMessageBox.critical(self, "Ошибка конфигурации", str(e))
            return
        self.fill_faculties(self.faculty_combo.currentData())
        self.log_msg(f"Конфиг перечитан: {self.cfg.path}")


    def build(self) -> None:
        # АВТООБНОВЛЕНИЕ: перед генерацией Word всегда свежие данные из json
        try:
            self.cfg = load_config(self.cfg.path)
        except Exception:
            pass # Если файл занят или открыт, пропускаем, чтобы не вызвать сбой
            
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
        
        # Определяем, какую папку открыть в диалоговом окне
        if self.last_out_dir and self.last_out_dir.exists():
            start_dir = self.last_out_dir
        else:
            start_dir = self.excel.path.parent if self.excel else Path.home()
            
        default_name = f"{fac.code} {course} курс с {start:%d.%m} - {end:%d.%m}.docx"
        
        # Открываем окно сохранения, подставляя актуальный путь
        path, _ = QFileDialog.getSaveFileName(
            self, "Сохранить расписание", str(start_dir / default_name), "Word (*.docx)"
        )
        if not path:
            return
            
        out = Path(path)
        if out.suffix.lower() != ".docx":
            out = out.with_suffix(".docx")

        # ЗАПОМИНАЕМ выбранную папку для следующих сохранений в этой сессии
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

        # русифицированный код:
        msg_box = QMessageBox(self)
        msg_box.setWindowTitle("Готово")
        msg_box.setText(f"Файл успешно сохранён:\n{out.name}\n\nОткрыть созданный документ?")
        msg_box.setIcon(QMessageBox.Question)
        
        # Добавляем кнопки с русским текстом
        btn_yes = msg_box.addButton("Да", QMessageBox.YesRole)
        btn_no = msg_box.addButton("Нет", QMessageBox.NoRole)
        msg_box.setDefaultButton(btn_yes)
        
        msg_box.exec()
        
        # Проверяем, какую кнопку нажал пользователь
        if msg_box.clickedButton() == btn_yes:
            QDesktopServices.openUrl(QUrl.fromLocalFile(str(out)))



def run() -> int:
    import sys

    app = QApplication(sys.argv)
    app.setApplicationName("ScheduleParser")
    win = MainWindow()
    win.show()
    return app.exec()
