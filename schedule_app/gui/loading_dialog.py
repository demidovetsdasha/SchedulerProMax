"""Модальное окно-заглушка с индикатором выполнения."""
from __future__ import annotations

from PySide6.QtCore import Qt
from PySide6.QtWidgets import QDialog, QLabel, QProgressBar, QVBoxLayout


class LoadingDialog(QDialog):
    def __init__(self, parent=None, text: str = "Обработка файла..."):
        super().__init__(parent)
        self.setObjectName("loading_dialog")
        self.setWindowFlags(Qt.Window | Qt.CustomizeWindowHint | Qt.WindowTitleHint)
        self.setModal(True)
        self.setFixedSize(320, 140)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(10)

        title = QLabel("Пожалуйста, подождите ⏳")
        title.setObjectName("loading_title")
        title.setAlignment(Qt.AlignCenter)

        self._subtitle = QLabel(text)
        self._subtitle.setObjectName("loading_subtitle")
        self._subtitle.setAlignment(Qt.AlignCenter)

        progress = QProgressBar()
        progress.setRange(0, 0)
        progress.setTextVisible(False)
        progress.setStyleSheet(
            "QProgressBar {"
            "  border: 1px solid #E2E8F0;"
            "  border-radius: 4px;"
            "  background-color: #EDF2F7;"
            "  height: 8px;"
            "}"
            "QProgressBar::chunk {"
            "  background-color: #48BB78;"
            "  border-radius: 4px;"
            "}"
        )

        layout.addWidget(title)
        layout.addWidget(self._subtitle)
        layout.addWidget(progress)

    def set_text(self, text: str) -> None:
        self._subtitle.setText(text)