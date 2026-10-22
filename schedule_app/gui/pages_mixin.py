"""Управление страницами документа и подгруппами."""
from __future__ import annotations

from PySide6.QtCore import Qt
from PySide6.QtGui import QBrush, QColor
from PySide6.QtWidgets import QListWidgetItem
from PySide6.QtCore import QSize, Qt
from PySide6.QtGui import QBrush, QColor, QFontMetrics

from ..docx_builder import paginate


class PagesMixin:
    @staticmethod
    def _page_title(i: int, names: list[str]) -> str:
        if not names:
            return f"Стр. {i + 1}: (нет групп)"
        groups: list[str] = []
        for n in names:
            g = n.split("/")[0]
            if g not in groups:
                groups.append(g)
        return f"Стр. {i + 1}: {', '.join(groups)}  ({len(names)} гр.)"

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

    def refresh_groups(self) -> None:
        page = self.pages_list.currentRow()
        self.page_groups_label.setText(
            f"Группы на странице {page + 1}:" if page >= 0 else "Группы на странице:")
        used_elsewhere = {
            n: i + 1 for i, names in enumerate(self.pages) if i != page for n in names}
        current = set(self.pages[page]) if 0 <= page < len(self.pages) else set()
        self.groups_list.blockSignals(True)
        self.groups_list.clear()

        fm = QFontMetrics(self.groups_list.font())        # ← для расчёта ширины текста

        for name in self.course_columns():
            other = used_elsewhere.get(name)
            text = f"{name}  (стр. {other})" if other else name
            item = QListWidgetItem(text)
            item.setData(Qt.UserRole, name)
            item.setFlags(item.flags() | Qt.ItemIsUserCheckable)
            item.setCheckState(Qt.Checked if name in current else Qt.Unchecked)
            if other:
                item.setForeground(QBrush(QColor("#888888")))

            # Явный sizeHint: иначе кастомный QSS-indicator ломает расчёт ширины
            text_w = fm.horizontalAdvance(text)
            item.setSizeHint(QSize(text_w + 44, 26))       # 44 = indicator + padding + запас
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