from __future__ import annotations

import getpass
import os

from PySide6.QtWidgets import QComboBox, QDialog, QFormLayout, QHBoxLayout, QLabel, QMessageBox, QVBoxLayout

from agribank_v3.features.credit.summary.customer.formatters import format_money_vn
from agribank_v3.features.credit.summary.customer.repository import CustomerRepository
from agribank_v3.features.credit.summary.customer.widgets import danger_button, secondary_button


class DeleteCustomerPeriodDialog(QDialog):
    def __init__(self, repository: CustomerRepository, parent=None) -> None:
        super().__init__(parent)
        self.repository = repository
        self.deleted_info: dict[str, object] | None = None
        self.setWindowTitle("Xóa dữ liệu khách hàng theo kỳ")
        layout = QVBoxLayout(self)
        form = QFormLayout()
        self.period_combo = QComboBox()
        self.period_combo.setObjectName("AgribankComboBox")
        for period in self.repository.distinct_periods():
            self.period_combo.addItem(period, period)
        self.info_label = QLabel("")
        self.period_combo.currentIndexChanged.connect(self.refresh_info)
        form.addRow("Kỳ cần xóa", self.period_combo)
        form.addRow("Thông tin", self.info_label)
        layout.addLayout(form)
        actions = QHBoxLayout()
        delete_button = danger_button("Xóa dữ liệu")
        cancel_button = secondary_button("Hủy")
        delete_button.clicked.connect(self.delete_period)
        cancel_button.clicked.connect(self.reject)
        actions.addStretch()
        actions.addWidget(delete_button)
        actions.addWidget(cancel_button)
        layout.addLayout(actions)
        self.refresh_info()

    def selected_period(self) -> str:
        return str(self.period_combo.currentData() or "")

    def refresh_info(self) -> None:
        period = self.selected_period()
        if not period:
            self.info_label.setText("Không có kỳ dữ liệu.")
            return
        info = self.repository.customer_period_info(period)
        self.info_label.setText(
            "Số khách hàng: {customer_count:,}\n"
            "Tổng dư nợ: {total_balance}\n"
            "Số quan hệ cán bộ: {officer_relation_count:,}\n"
            "Số import run: {import_run_count:,}\n"
            "Số file import: {import_file_count:,}".format(
                customer_count=int(info.get("customer_count") or 0),
                total_balance=format_money_vn(info.get("total_balance", 0)),
                officer_relation_count=int(info.get("officer_relation_count") or 0),
                import_run_count=int(info.get("import_run_count") or 0),
                import_file_count=int(info.get("import_file_count") or 0),
            ).replace(",", ".")
        )

    def delete_period(self) -> None:
        period = self.selected_period()
        if not period:
            return
        answer = QMessageBox.question(
            self,
            "Xóa dữ liệu khách hàng theo kỳ",
            f"Bạn có chắc muốn xóa toàn bộ dữ liệu khách hàng kỳ {period} không?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        if answer != QMessageBox.StandardButton.Yes:
            return
        self.deleted_info = self.repository.delete_customer_period(
            period,
            user_name=_current_user(),
            computer_name=os.environ.get("COMPUTERNAME", ""),
        )
        self.accept()


def _current_user() -> str:
    try:
        return getpass.getuser()
    except Exception:
        return ""
