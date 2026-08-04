from __future__ import annotations

import getpass
import os

from PySide6.QtWidgets import QCheckBox, QComboBox, QDialog, QFormLayout, QHBoxLayout, QLabel, QMessageBox, QVBoxLayout

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
        self.last_period_label = QLabel(
            "Đây là kỳ dữ liệu cuối cùng.\n"
            "Dữ liệu khách hàng, dư nợ, cán bộ theo kỳ và đơn vị theo kỳ sẽ bị xóa.\n"
            "Danh mục CBTD được lưu độc lập và mặc định sẽ được giữ lại."
        )
        self.last_period_label.setObjectName("ValidationWarningLabel")
        self.last_period_label.setWordWrap(True)
        self.delete_directory_check = QCheckBox("Xóa luôn danh mục CBTD")
        self.delete_override_check = QCheckBox("Xóa các ghi đè cán bộ")
        self.delete_action_log_check = QCheckBox("Xóa nhật ký thao tác")
        self.delete_directory_check.toggled.connect(self._directory_delete_toggled)
        layout.addWidget(self.last_period_label)
        layout.addWidget(self.delete_directory_check)
        layout.addWidget(self.delete_override_check)
        layout.addWidget(self.delete_action_log_check)
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
            self._set_last_period_options_visible(False)
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
        self._set_last_period_options_visible(self._is_last_period(period))

    def delete_period(self) -> None:
        period = self.selected_period()
        if not period:
            return
        if self.delete_directory_check.isChecked() and not self._confirm_delete_directory():
            return
        try:
            self.deleted_info = self.repository.delete_customer_period(
                period,
                user_name=_current_user(),
                computer_name=os.environ.get("COMPUTERNAME", ""),
                delete_officer_directory=self.delete_directory_check.isChecked(),
                delete_officer_overrides=self.delete_override_check.isChecked(),
                delete_action_log=self.delete_action_log_check.isChecked(),
            )
        except Exception as exc:
            QMessageBox.warning(self, "Xóa dữ liệu khách hàng theo kỳ", str(exc))
            return
        self.accept()

    def _is_last_period(self, period: str) -> bool:
        return self.repository.distinct_periods() == [period]

    def _set_last_period_options_visible(self, visible: bool) -> None:
        self.last_period_label.setVisible(visible)
        for checkbox in (
            self.delete_directory_check,
            self.delete_override_check,
            self.delete_action_log_check,
        ):
            checkbox.setVisible(visible)
            checkbox.setEnabled(visible)
            if not visible:
                checkbox.setChecked(False)

    def _directory_delete_toggled(self, checked: bool) -> None:
        if checked:
            self.delete_override_check.setChecked(True)
            self.delete_override_check.setEnabled(False)
        else:
            self.delete_override_check.setEnabled(self.delete_override_check.isVisible())

    def _confirm_delete_directory(self) -> bool:
        answer = QMessageBox.question(
            self,
            "Xóa danh mục CBTD",
            "Danh mục CBTD có thể chứa các thông tin đã chỉnh sửa thủ công. "
            "Thao tác này không thể hoàn tác nếu chưa có bản sao lưu.\n\n"
            "Bạn có chắc muốn xóa danh mục CBTD và các ghi đè cán bộ không?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        return answer == QMessageBox.StandardButton.Yes


def _current_user() -> str:
    try:
        return getpass.getuser()
    except Exception:
        return ""
