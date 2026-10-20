"""Графический интерфейс приложения (PySide6 / Qt)."""
from __future__ import annotations

from .main_window import MainWindow

__all__ = ["MainWindow", "run"]


def run() -> int:
    import sys
    from PySide6.QtWidgets import QApplication

    app = QApplication(sys.argv)
    app.setApplicationName("ScheduleParser")
    win = MainWindow()
    win.show()
    return app.exec()