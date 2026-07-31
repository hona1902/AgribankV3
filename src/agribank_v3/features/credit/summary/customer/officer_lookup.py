from __future__ import annotations

from PySide6.QtCore import Signal, Qt, QTimer
from PySide6.QtWidgets import QComboBox, QDialog, QHBoxLayout, QLabel, QSizePolicy, QVBoxLayout, QWidget

from agribank_v3.features.credit.summary.customer.formatters import normalize_officer_code, normalize_officer_name
from agribank_v3.features.credit.summary.customer.repository import CustomerRepository
from agribank_v3.features.credit.summary.customer.widgets import (
    configure_combo_popup_width,
    configure_searchable_combo,
    secondary_button,
)


OFFICER_ROW_ROLE = Qt.ItemDataRole.UserRole + 20


class OfficerLookupWidget(QWidget):
    officerResolved = Signal(dict)

    def __init__(
        self,
        repository: CustomerRepository,
        *,
        branch_code: str = "",
        parent=None,
    ) -> None:
        super().__init__(parent)
        self.repository = repository
        self.branch_code = str(branch_code or "").strip()
        self._selected_officer: dict[str, object] | None = None
        self._pending_officer_code = ""
        self._updating = False
        self._last_query = ""
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(6)
        field_row = QHBoxLayout()
        field_row.setContentsMargins(0, 0, 0, 0)
        field_row.setSpacing(10)
        self.name_combo = _lookup_combo("Nhập mã hoặc tên cán bộ")
        self.code_combo = self.name_combo
        self.status_label = QLabel("")
        self.status_label.setObjectName("MutedText")
        self.status_label.setWordWrap(True)
        self.add_button = secondary_button("Thêm vào danh mục cán bộ")
        self.add_button.hide()
        field_row.addWidget(QLabel("Tên cán bộ mới"))
        field_row.addWidget(self.name_combo, stretch=1)
        layout.addLayout(field_row)
        layout.addWidget(self.status_label)
        layout.addWidget(self.add_button, alignment=Qt.AlignmentFlag.AlignLeft)
        self.lookup_timer = QTimer(self)
        self.lookup_timer.setSingleShot(True)
        self.lookup_timer.setInterval(300)
        self.code_timer = self.lookup_timer
        self.name_timer = self.lookup_timer
        self.lookup_timer.timeout.connect(self.lookup)
        self.name_combo.editTextChanged.connect(self._lookup_edited)
        self.name_combo.activated.connect(lambda _index: self._select_from_combo(self.name_combo))
        self.add_button.clicked.connect(self.add_new_officer)
        self.name_combo.setFocus()

    def code(self) -> str:
        selected = self.selected_officer()
        if selected is not None:
            return normalize_officer_code(selected.get("officer_code"))
        return ""

    def name(self) -> str:
        selected = self.selected_officer()
        if selected is not None:
            return normalize_officer_name(selected.get("officer_name"))
        return normalize_officer_name(self.name_combo.currentText())

    def selected_officer(self) -> dict[str, object] | None:
        if self._selected_officer is None:
            return None
        current = normalize_officer_name(self.name_combo.currentText())
        selected_name = normalize_officer_name(self._selected_officer.get("officer_name"))
        if current.casefold() == selected_name.casefold():
            return dict(self._selected_officer)
        return None

    def set_pending_identity(self, officer_code: str = "", officer_name: str = "") -> None:
        self._pending_officer_code = normalize_officer_code(officer_code)
        self._updating = True
        try:
            self.name_combo.setEditText(normalize_officer_name(officer_name) or normalize_officer_code(officer_code))
        finally:
            self._updating = False
        self._selected_officer = None

    def validate_selected(self) -> tuple[bool, str]:
        selected = self.selected_officer()
        if selected is None:
            return False, "Vui lòng chọn cán bộ trong danh sách gợi ý."
        row = self.repository.get_officer_by_identity(
            normalize_officer_code(selected.get("officer_code")),
            normalize_officer_name(selected.get("officer_name")),
            active_only=True,
        )
        if row is None:
            return False, "Vui lòng chọn cán bộ trong danh sách gợi ý."
        self._select_officer(row)
        return True, ""

    def lookup(self) -> None:
        query = normalize_officer_name(self.name_combo.currentText())
        self._last_query = query
        if not query:
            self._clear_selection("")
            return
        exact_code = self.repository.get_officer_by_code(query, active_only=True)
        if exact_code is not None:
            self._select_officer(exact_code)
            return
        rows_by_code = self.repository.find_officers_by_code_prefix(
            query,
            branch_code=self.branch_code,
            active_only=True,
            limit=25,
        )
        rows_by_name = self.repository.find_officers_by_name(
            query,
            branch_code=self.branch_code,
            active_only=True,
            limit=25,
        )
        rows = _merge_officer_rows(rows_by_code, rows_by_name, limit=25)
        exact_name = [
            row
            for row in rows
            if normalize_officer_name(row.get("officer_name")).casefold() == query.casefold()
        ]
        self._populate_suggestions(rows)
        if len(exact_name) == 1:
            self._select_officer(exact_name[0])
            return
        if len(exact_name) > 1:
            self._clear_selection("Có nhiều cán bộ trùng tên. Vui lòng chọn đúng mã trong danh sách.")
            self.add_button.hide()
            return
        self._clear_selection("Không tìm thấy cán bộ trong danh mục." if not rows else "Chọn cán bộ trong danh sách gợi ý.")
        self.add_button.setVisible(not rows)

    def lookup_by_code(self) -> None:
        self.lookup()

    def lookup_by_name(self) -> None:
        self.lookup()

    def add_new_officer(self) -> None:
        from agribank_v3.features.credit.summary.customer.officer_management_tab import OfficerDirectoryDialog

        text = normalize_officer_name(self.name_combo.currentText())
        pending_code = normalize_officer_code(self._pending_officer_code)
        input_is_code = _looks_like_officer_code(text)
        row = {
            "officer_code": pending_code or (text if input_is_code else ""),
            "officer_name": "" if input_is_code else text,
            "branch_code": self.branch_code,
            "transaction_office": "",
            "is_active": 1,
        }
        dialog = OfficerDirectoryDialog(self.repository, row=row, parent=self)
        if dialog.exec() != QDialog.DialogCode.Accepted:
            return
        officer = self.repository.get_officer_by_code(normalize_officer_code(dialog.code_input.text()), active_only=True)
        if officer is not None:
            self._select_officer(officer)

    def closeEvent(self, event) -> None:
        self.lookup_timer.stop()
        super().closeEvent(event)

    def _lookup_edited(self, _text: str) -> None:
        if self._updating:
            return
        self._selected_officer = None
        self._pending_officer_code = ""
        self.add_button.hide()
        self.lookup_timer.start()

    def _select_from_combo(self, combo: QComboBox) -> None:
        row = combo.currentData(OFFICER_ROW_ROLE)
        if isinstance(row, dict):
            self._select_officer(row)

    def _select_officer(self, row: dict[str, object]) -> None:
        self.lookup_timer.stop()
        self._selected_officer = dict(row)
        name = normalize_officer_name(row.get("officer_name"))
        tooltip = _officer_tooltip(row)
        self._updating = True
        try:
            self._populate_suggestions([row])
            self.name_combo.setCurrentIndex(0)
            self.name_combo.setEditText(name)
            self.name_combo.setToolTip(tooltip)
        finally:
            self._updating = False
        self._pending_officer_code = normalize_officer_code(row.get("officer_code"))
        self.status_label.setText(tooltip.replace("\n", " | "))
        self.add_button.hide()
        self.officerResolved.emit(dict(row))

    def _clear_selection(self, message: str) -> None:
        self._selected_officer = None
        self._pending_officer_code = ""
        self.status_label.setText(message)

    def _populate_suggestions(self, rows: list[dict[str, object]]) -> None:
        current_text = self.name_combo.currentText()
        self.name_combo.blockSignals(True)
        self.name_combo.clear()
        for row in rows:
            self.name_combo.addItem(_officer_display(row), normalize_officer_code(row.get("officer_code")))
            index = self.name_combo.count() - 1
            self.name_combo.setItemData(index, dict(row), OFFICER_ROW_ROLE)
            self.name_combo.setItemData(index, _officer_tooltip(row), Qt.ItemDataRole.ToolTipRole)
        self.name_combo.setEditText(current_text)
        configure_combo_popup_width(self.name_combo, minimum_popup_width=440, maximum_screen_ratio=0.55)
        self.name_combo.blockSignals(False)


def _lookup_combo(placeholder: str) -> QComboBox:
    combo = QComboBox()
    combo.setObjectName("AgribankComboBox")
    combo.setMinimumWidth(340)
    combo.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
    configure_searchable_combo(combo)
    if combo.lineEdit() is not None:
        combo.lineEdit().setPlaceholderText(placeholder)
    return combo


def _merge_officer_rows(*groups: list[dict[str, object]], limit: int) -> list[dict[str, object]]:
    output: list[dict[str, object]] = []
    seen: set[str] = set()
    for group in groups:
        for row in group:
            code = normalize_officer_code(row.get("officer_code"))
            if not code or code.casefold() in seen:
                continue
            seen.add(code.casefold())
            output.append(dict(row))
            if len(output) >= limit:
                return output
    return output


def _looks_like_officer_code(value: str) -> bool:
    text = normalize_officer_code(value)
    return bool(text) and any(char.isdigit() for char in text) and not any(char.isspace() for char in text)


def _officer_display(row: dict[str, object]) -> str:
    code = normalize_officer_code(row.get("officer_code"))
    name = normalize_officer_name(row.get("officer_name"))
    branch = str(row.get("branch_code") or "").strip()
    office = str(row.get("transaction_office") or "").strip()
    suffix = " - ".join(item for item in (branch, office) if item)
    text = f"[{code}] {name}" if code else name
    return f"{text} - {suffix}" if suffix else text


def _officer_tooltip(row: dict[str, object]) -> str:
    active = "Đang sử dụng" if int(row.get("is_active", 1) or 0) == 1 else "Ngừng sử dụng"
    return "\n".join(
        (
            f"Mã: {normalize_officer_code(row.get('officer_code'))}",
            f"Tên: {normalize_officer_name(row.get('officer_name'))}",
            f"Chi nhánh: {str(row.get('branch_code') or '').strip() or 'N/A'}",
            f"Phòng GD: {str(row.get('transaction_office') or '').strip() or 'N/A'}",
            f"Trạng thái: {active}",
        )
    )
