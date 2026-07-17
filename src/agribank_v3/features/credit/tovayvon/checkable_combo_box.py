from __future__ import annotations

from typing import Any

from PySide6.QtCore import QEvent, QTimer, Qt
from PySide6.QtGui import QStandardItem, QStandardItemModel
from PySide6.QtWidgets import QComboBox


SELECT_ALL_VALUE = "__checkable_combo_select_all__"


class CheckableComboBox(QComboBox):
    """Small multi-select combobox backed by checkable items."""

    def __init__(
        self,
        parent=None,
        placeholder: str = "Chọn tổ vay vốn...",
        all_text: str = "Tất cả",
        all_selected_text: str = "Tất cả tổ vay vốn",
    ) -> None:
        super().__init__(parent)
        self.placeholder = placeholder
        self.all_text = all_text
        self.all_selected_text = all_selected_text
        self._skip_next_hide = False
        self._updating_all_state = False
        self.setModel(QStandardItemModel(self))
        self.setEditable(True)
        self.lineEdit().setReadOnly(True)
        self.lineEdit().setPlaceholderText(placeholder)
        self.lineEdit().installEventFilter(self)
        self.view().pressed.connect(self._toggle_item)
        self._add_select_all_item()
        self._update_display_text()

    def add_check_item(self, text: str, data: Any) -> None:
        item = QStandardItem(text)
        item.setData(data, Qt.ItemDataRole.UserRole)
        item.setFlags(
            Qt.ItemFlag.ItemIsEnabled
            | Qt.ItemFlag.ItemIsUserCheckable
            | Qt.ItemFlag.ItemIsSelectable
        )
        item.setCheckState(Qt.CheckState.Unchecked)
        self.model().appendRow(item)
        self._sync_select_all_state()
        self._update_display_text()

    def get_selected_values(self) -> list[Any]:
        values: list[Any] = []
        for row in range(1, self.model().rowCount()):
            item = self.model().item(row)
            if item.checkState() == Qt.CheckState.Checked:
                values.append(item.data(Qt.ItemDataRole.UserRole))
        return values

    def checked_data(self) -> list[Any]:
        return self.get_selected_values()

    def checked_texts(self) -> list[str]:
        texts: list[str] = []
        for row in range(1, self.model().rowCount()):
            item = self.model().item(row)
            if item.checkState() == Qt.CheckState.Checked:
                texts.append(item.text())
        return texts

    def set_selected_values(self, values: list[Any] | tuple[Any, ...]) -> None:
        selected = {str(value) for value in values}
        for row in range(1, self.model().rowCount()):
            item = self.model().item(row)
            item.setCheckState(
                Qt.CheckState.Checked
                if str(item.data(Qt.ItemDataRole.UserRole)) in selected
                else Qt.CheckState.Unchecked
            )
        self._sync_select_all_state()
        self._update_display_text()

    def set_checked_data(self, values: list[Any] | tuple[Any, ...]) -> None:
        self.set_selected_values(values)

    def clear_selection(self) -> None:
        self.deselect_all()

    def clear_checked(self) -> None:
        self.deselect_all()

    def select_all(self) -> None:
        for row in range(1, self.model().rowCount()):
            self.model().item(row).setCheckState(Qt.CheckState.Checked)
        self._sync_select_all_state()
        self._update_display_text()

    def deselect_all(self) -> None:
        for row in range(1, self.model().rowCount()):
            self.model().item(row).setCheckState(Qt.CheckState.Unchecked)
        self._sync_select_all_state()
        self._update_display_text()

    def selected_count(self) -> int:
        return len(self.get_selected_values())

    def clear(self) -> None:
        super().clear()
        self._add_select_all_item()
        self._update_display_text()

    def hidePopup(self) -> None:
        if self._skip_next_hide:
            self._skip_next_hide = False
            return
        super().hidePopup()

    def mouseReleaseEvent(self, event) -> None:
        QTimer.singleShot(0, self.showPopup)
        event.accept()

    def eventFilter(self, watched, event) -> bool:
        if watched is self.lineEdit() and event.type() == QEvent.Type.MouseButtonRelease:
            QTimer.singleShot(0, self.showPopup)
            return True
        return super().eventFilter(watched, event)

    def _toggle_item(self, index) -> None:
        item = self.model().itemFromIndex(index)
        if item is None:
            return
        self._skip_next_hide = True
        QTimer.singleShot(150, self._reset_skip_next_hide)
        if item.data(Qt.ItemDataRole.UserRole) == SELECT_ALL_VALUE:
            if item.checkState() == Qt.CheckState.Checked:
                self.deselect_all()
            else:
                self.select_all()
            return
        item.setCheckState(
            Qt.CheckState.Unchecked
            if item.checkState() == Qt.CheckState.Checked
            else Qt.CheckState.Checked
        )
        self._sync_select_all_state()
        self._update_display_text()

    def _update_display_text(self) -> None:
        checked = self.checked_texts()
        if not checked:
            text = self.placeholder
        elif len(checked) == self._item_count():
            text = self.all_selected_text
        elif len(checked) == 1:
            text = checked[0]
        else:
            text = f"Đã chọn {len(checked)} tổ"
        self.lineEdit().setText(text)

    def _add_select_all_item(self) -> None:
        item = QStandardItem(self.all_text)
        item.setData(SELECT_ALL_VALUE, Qt.ItemDataRole.UserRole)
        item.setFlags(
            Qt.ItemFlag.ItemIsEnabled
            | Qt.ItemFlag.ItemIsUserCheckable
            | Qt.ItemFlag.ItemIsSelectable
        )
        item.setCheckState(Qt.CheckState.Unchecked)
        self.model().appendRow(item)

    def _sync_select_all_state(self) -> None:
        if self._updating_all_state or self.model().rowCount() == 0:
            return
        self._updating_all_state = True
        try:
            all_item = self.model().item(0)
            total = self._item_count()
            checked = self.selected_count()
            if total == 0 or checked == 0:
                all_item.setCheckState(Qt.CheckState.Unchecked)
            elif checked == total:
                all_item.setCheckState(Qt.CheckState.Checked)
            else:
                all_item.setCheckState(Qt.CheckState.PartiallyChecked)
        finally:
            self._updating_all_state = False

    def _item_count(self) -> int:
        return max(0, self.model().rowCount() - 1)

    def _reset_skip_next_hide(self) -> None:
        self._skip_next_hide = False
