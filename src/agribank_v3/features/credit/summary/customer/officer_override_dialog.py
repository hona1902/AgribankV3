from __future__ import annotations

import getpass
import os

from PySide6.QtWidgets import (
    QComboBox,
    QDialog,
    QFormLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QSizePolicy,
    QVBoxLayout,
)

from agribank_v3.features.credit.summary.customer.officer_lookup import OfficerLookupWidget
from agribank_v3.features.credit.summary.customer.repository import CustomerRepository
from agribank_v3.features.credit.summary.customer.widgets import fit_window_to_screen, primary_button, secondary_button


class OfficerOverrideDialog(QDialog):
    def __init__(
        self,
        repository: CustomerRepository,
        *,
        customer_code: str,
        customer_name: str,
        period: str,
        imported_officer_code: str = "",
        imported_officer_name: str = "",
        parent=None,
    ) -> None:
        super().__init__(parent)
        self.repository = repository
        self.customer_code = str(customer_code or "").strip()
        self.period = str(period or "").strip()
        self.imported_officer_code = str(imported_officer_code or "").strip()
        self.imported_officer_name = str(imported_officer_name or "").strip()
        self.setWindowTitle("Cập nhật cán bộ quản lý - AgribankV3")
        fit_window_to_screen(
            self,
            width_ratio=0.44,
            height_ratio=0.50,
            max_width=680,
            max_height=440,
            min_width=540,
            min_height=340,
        )
        layout = QVBoxLayout(self)
        layout.addWidget(QLabel(f"Mã KH: {self.customer_code}"))
        layout.addWidget(QLabel(f"Tên KH: {customer_name}"))
        layout.addWidget(QLabel(f"Kỳ đang chọn: {self.period}"))
        layout.addWidget(QLabel(f"Cán bộ chính từ dữ liệu import: {self.imported_officer_name or self.imported_officer_code}"))
        form = QFormLayout()
        self.officer_lookup = OfficerLookupWidget(self.repository, parent=self)
        self.officer_code_input = self.officer_lookup.code_combo
        self.officer_name_input = self.officer_lookup.name_combo
        self.reason_input = QLineEdit()
        self.scope_combo = QComboBox()
        self.scope_combo.setObjectName("AgribankComboBox")
        self.scope_combo.addItem("Chỉ kỳ đang chọn", "one_period")
        self.scope_combo.addItem("Từ kỳ đang chọn trở đi", "from_period")
        self.scope_combo.addItem("Từ kỳ đang chọn đến kỳ kết thúc", "range")
        self.scope_combo.addItem("Thiết lập cán bộ hiện hành không giới hạn kỳ kết thúc", "open_ended")
        self.end_period_input = QLineEdit()
        self.end_period_input.setPlaceholderText("YYYY-MM")
        for field in (self.reason_input, self.end_period_input):
            field.setMinimumWidth(340)
            field.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        self.scope_combo.setMinimumWidth(340)
        self.scope_combo.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        form.addRow(self.officer_lookup)
        form.addRow("Lý do cập nhật", self.reason_input)
        form.addRow("Phạm vi áp dụng", self.scope_combo)
        form.addRow("Kỳ kết thúc", self.end_period_input)
        layout.addLayout(form)
        actions = QHBoxLayout()
        save_button = primary_button("Lưu")
        cancel_button = secondary_button("Hủy")
        save_button.clicked.connect(self.save_override)
        cancel_button.clicked.connect(self.reject)
        actions.addStretch()
        actions.addWidget(save_button)
        actions.addWidget(cancel_button)
        layout.addLayout(actions)
        self.setTabOrder(self.officer_name_input, self.reason_input)
        self.setTabOrder(self.reason_input, self.scope_combo)
        self.setTabOrder(self.scope_combo, self.end_period_input)
        self.setTabOrder(self.end_period_input, save_button)
        self.setTabOrder(save_button, cancel_button)
        self.officer_name_input.setFocus()

    def save_override(self) -> None:
        is_valid, message = self.officer_lookup.validate_selected()
        if not is_valid:
            self.officer_lookup.status_label.setText(message)
            QMessageBox.warning(self, "Cập nhật cán bộ quản lý", message)
            return
        officer = self.officer_lookup.selected_officer() or {}
        officer_code = str(officer.get("officer_code") or self.officer_lookup.code()).strip()
        officer_name = str(officer.get("officer_name") or self.officer_lookup.name()).strip()
        scope = str(self.scope_combo.currentData() or "one_period")
        if scope == "one_period":
            to_period = self.period
        elif scope == "range":
            to_period = self.end_period_input.text().strip()
        else:
            to_period = ""
        try:
            self.repository.create_officer_override(
                customer_code=self.customer_code,
                effective_from_period=self.period,
                effective_to_period=to_period,
                officer_code=officer_code,
                officer_name=officer_name,
                reason=self.reason_input.text(),
                created_by=_current_user(),
                computer_name=os.environ.get("COMPUTERNAME", ""),
            )
        except Exception as exc:
            QMessageBox.warning(self, "Cập nhật cán bộ quản lý", str(exc))
            return
        self.accept()


def _current_user() -> str:
    try:
        return getpass.getuser()
    except Exception:
        return ""
