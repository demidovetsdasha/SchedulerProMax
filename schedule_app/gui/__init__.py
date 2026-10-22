"""Графический интерфейс приложения (PySide6 / Qt)."""
from __future__ import annotations

from .main_window import MainWindow

__all__ = ["MainWindow", "run"]


def run() -> int:
    import sys
    from PySide6.QtWidgets import QApplication

    from .theme import apply_light_theme

    app = QApplication(sys.argv)
    app.setApplicationName("ScheduleParser")
    apply_light_theme(app)
    win = MainWindow()
    win.show()
    return app.exec()