"""Единый QSS-стиль приложения.

Пути к иконкам (%CHECK_ICON%, %DOWN_ICON%, %UP_ICON%) подставляет theme.themed_stylesheet.
"""

MODERN_STYLESHEET = """
    QMainWindow {
        background-color: #F8F9FA;
    }

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

    QLineEdit, QComboBox, QAbstractSpinBox, QDateEdit {
        border: 1px solid #CBD5E0;
        border-radius: 6px;
        padding: 6px 12px;
        background-color: #FFFFFF;
        color: #2D3748;
        selection-background-color: #2F855A;
    }

    QComboBox::drop-down, QDateEdit::drop-down {
        subcontrol-origin: padding;
        subcontrol-position: top right;
        width: 24px;
        border: none;
        background: transparent;
    }
    QAbstractSpinBox::up-button, QAbstractSpinBox::down-button {
        width: 18px;
        border: none;
        background: transparent;
    }
    QAbstractSpinBox::up-button:hover, QAbstractSpinBox::down-button:hover {
        background-color: #EDF2F7;
    }
    QComboBox::down-arrow, QDateEdit::down-arrow {
        image: url(%DOWN_ICON%);
        width: 10px;
        height: 10px;
    }
    QAbstractSpinBox::down-arrow {
        image: url(%DOWN_ICON%);
        width: 9px;
        height: 9px;
    }
    QAbstractSpinBox::up-arrow {
        image: url(%UP_ICON%);
        width: 9px;
        height: 9px;
    }

    QAbstractSpinBox::up-down-key-pressed,
    QAbstractSpinBox QLineEdit {
        padding: 0px 4px;
        background: transparent;
        border: none;
    }

    QLineEdit:focus, QComboBox:focus, QAbstractSpinBox:focus, QDateEdit:focus {
        border: 1px solid #48BB78;
    }

    QDateEdit {
        padding-right: 20px;
    }

    QDateEdit::arrow {
        image: none;
    }

    QCalendarWidget QWidget {
        background-color: #FFFFFF;
    }
    QCalendarWidget QNodeWidget {
        background-color: #FFFFFF;
        color: #2D3748;
    }
    QCalendarWidget QAbstractItemView:enabled {
        color: #2D3748;
        selection-background-color: #48BB78;
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

    QListWidget {
        border: 1px solid #E2E8F0;
        border-radius: 6px;
        background-color: #FFFFFF;
        padding: 4px;
        outline: 0;   /* без пунктирной рамки фокуса вокруг текущего элемента */
        color: #2D3748;
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

    QListWidget::indicator {
        width: 14px;
        height: 14px;
        border: 1px solid #A0AEC0;
        border-radius: 3px;
        background-color: #FFFFFF;
    }
    QListWidget::indicator:hover {
        border-color: #48BB78;
    }
    QListWidget::indicator:checked {
        border-color: #38A169;
        background-color: #48BB78;
        image: url(%CHECK_ICON%);
    }

    QLabel {
        color: #4A5568;
    }

    QPlainTextEdit#mini_console {
        background-color: #1A202C;
        color: #A0AEC0;
        font-family: 'Consolas', 'Courier New', monospace;
        font-size: 12px;
        border: 1px solid #2D3748;
        border-radius: 6px;
        padding: 8px;
    }

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
    QMessageBox QPushButton:default {
        background-color: #E6FFFA;
        border: 1px solid #48BB78;
        color: #22543D;
    }
    QMessageBox QPushButton:default:hover {
        background-color: #C6F6D5;
    }

    QLabel#copyright_label {
        color: #718096;
        font-size: 11px;
        font-weight: normal;
        padding-top: 5px;
    }
"""