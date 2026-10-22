"""Фиксированная светлая тема: не зависит от стиля и тёмного режима ОС.

На Windows 11 Qt по умолчанию берёт стиль «windows11», который плохо сочетается с QSS
(чекбоксы в списках рисуются белыми на белом), а в тёмном режиме подставляет тёмную
палитру — элементы, которым QSS не задаёт цвет, получают белый текст на белом фоне.
"""
from __future__ import annotations

import tempfile
from pathlib import Path

from PySide6.QtCore import QPointF, Qt
from PySide6.QtGui import QColor, QImage, QPainter, QPalette, QPen
from PySide6.QtWidgets import QApplication

TEXT = "#2D3748"
TEXT_DISABLED = "#A0AEC0"
ACCENT = "#48BB78"

_PALETTE = {
    QPalette.Window: "#F8F9FA",
    QPalette.WindowText: TEXT,
    QPalette.Base: "#FFFFFF",
    QPalette.AlternateBase: "#F7FAFC",
    QPalette.ToolTipBase: "#FFFFFF",
    QPalette.ToolTipText: TEXT,
    QPalette.PlaceholderText: "#A0AEC0",
    QPalette.Text: TEXT,
    QPalette.Button: "#EDF2F7",
    QPalette.ButtonText: TEXT,
    QPalette.BrightText: "#FFFFFF",
    QPalette.Light: "#FFFFFF",
    QPalette.Midlight: "#EDF2F7",
    QPalette.Mid: "#CBD5E0",
    QPalette.Dark: "#A0AEC0",
    QPalette.Shadow: "#718096",
    QPalette.Highlight: ACCENT,
    QPalette.HighlightedText: "#FFFFFF",
    QPalette.Link: "#2F855A",
    QPalette.LinkVisited: "#22543D",
}


def light_palette() -> QPalette:
    pal = QPalette()
    for role, color in _PALETTE.items():
        pal.setColor(role, QColor(color))
    for role in (QPalette.WindowText, QPalette.Text, QPalette.ButtonText):
        pal.setColor(QPalette.Disabled, role, QColor(TEXT_DISABLED))
    return pal


def apply_light_theme(app: QApplication) -> None:
    app.setStyle("Fusion")
    hints = app.styleHints()
    if hasattr(hints, "setColorScheme"):   # Qt ≥ 6.8: не реагировать на тёмный режим ОС
        hints.setColorScheme(Qt.ColorScheme.Light)
    app.setPalette(light_palette())


def _icon(name: str, color: str, width: float, points: list[tuple[float, float]]) -> Path:
    """Рисует ломаную в PNG 28×28 (для QSS нужен файл) и возвращает путь к нему.

    Рисуем с запасом — QSS уменьшит до размера элемента, на HiDPI будет чётко.
    """
    path = Path(tempfile.gettempdir()) / f"schedule_parser_{name}.png"
    img = QImage(28, 28, QImage.Format_ARGB32)
    img.fill(Qt.transparent)
    p = QPainter(img)
    p.setRenderHint(QPainter.Antialiasing)
    p.setPen(QPen(QColor(color), width, Qt.SolidLine, Qt.RoundCap, Qt.RoundJoin))
    p.drawPolyline([QPointF(x, y) for x, y in points])
    p.end()
    img.save(str(path), "PNG")
    return path


def themed_stylesheet(qss: str) -> str:
    """Подставляет в QSS пути к нарисованным иконкам (%CHECK_ICON%, %DOWN_ICON%, %UP_ICON%)."""
    icons = {
        "%CHECK_ICON%": _icon("check", "#FFFFFF", 3.5, [(6, 14.5), (11.5, 20), (22, 8)]),
        "%DOWN_ICON%": _icon("down", "#4A5568", 3, [(7, 10), (14, 17), (21, 10)]),
        "%UP_ICON%": _icon("up", "#4A5568", 3, [(7, 18), (14, 11), (21, 18)]),
    }
    for key, path in icons.items():
        qss = qss.replace(key, path.as_posix())
    return qss
