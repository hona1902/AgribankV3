from __future__ import annotations

from pathlib import Path

from PySide6.QtWidgets import (
    QCheckBox,
    QDialog,
    QDialogButtonBox,
    QFileDialog,
    QFrame,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QRadioButton,
    QVBoxLayout,
    QWidget,
)

from agribank_v3.settings import BranchProfile
from agribank_v3.settlement.models import SettlementOptions, SettlementSpec
from agribank_v3.ui.dialogs.settlement_period import (
    load_output_prefix,
    save_output_prefix,
)


class Mau1516SettlementDialog(QDialog):
    def __init__(
        self,
        spec: SettlementSpec,
        profile: BranchProfile,
        parent: QWidget | None = None,
    ) -> None:
        super().__init__(parent)
        self.spec = spec
        self.profile = profile
        self.source_path: Path | None = None

        self.setWindowTitle(f"Tạo Mẫu biểu Quyết toán {spec.report_code}/QT")
        self.setModal(True)
        self.setMinimumWidth(690)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(14, 12, 14, 14)
        layout.setSpacing(12)

        source_hint = self._source_hint_text()
        source_label = QLabel(
            "Tên File nguồn dùng xử lý để tạo ra Mẫu biểu "
            f"{spec.report_code}QT là: {source_hint}"
        )
        source_label.setStyleSheet("color: #0000ff; font-weight: 700;")
        source_label.setWordWrap(True)
        layout.addWidget(source_label)

        source_row = QHBoxLayout()
        self.source_edit = QLineEdit()
        self.source_edit.setReadOnly(True)
        source_row.addWidget(self.source_edit, 1)
        choose_button = QPushButton("Chọn File")
        choose_button.clicked.connect(self.choose_source_file)
        source_row.addWidget(choose_button)
        layout.addLayout(source_row)

        options_label = QLabel("Chọn lựa tiêu chí để tạo file:")
        options_label.setStyleSheet("color: #0000ff; font-weight: 700;")
        layout.addWidget(options_label)

        options_row = QHBoxLayout()
        options_row.addWidget(self._customer_code_group())
        options_row.addWidget(self._control_sheet_group())
        layout.addLayout(options_row)

        second_options_row = QHBoxLayout()
        second_options_row.addWidget(self._date_format_group())
        second_options_row.addWidget(self._cleanup_group())
        layout.addLayout(second_options_row)

        self.lds_checkbox = QCheckBox("Sử dụng số LDS thay cho số LAV")
        self.lds_checkbox.setVisible(spec.report_code.casefold() in {"15a", "15b"})
        layout.addWidget(self.lds_checkbox)

        self.output_prefix_group = self._output_prefix_group()
        layout.addWidget(self.output_prefix_group)

        output_label = QLabel(f"Tên File quyết toán Mẫu {spec.report_code}/QT sẽ được tạo ra:")
        output_label.setStyleSheet("color: #0000ff; font-weight: 700;")
        layout.addWidget(output_label)
        self.output_edit = QLineEdit()
        self.output_edit.setReadOnly(True)
        layout.addWidget(self.output_edit)

        buttons = QDialogButtonBox()
        create_button = QPushButton("Tạo Mẫu biểu")
        create_button.setObjectName("PrimaryButton")
        buttons.addButton(create_button, QDialogButtonBox.ButtonRole.AcceptRole)
        cancel_button = QPushButton("Cancel")
        buttons.addButton(cancel_button, QDialogButtonBox.ButtonRole.RejectRole)
        create_button.clicked.connect(self.accept)
        cancel_button.clicked.connect(self.reject)
        layout.addWidget(buttons)

        self._apply_saved_output_prefix()

    def _customer_code_group(self) -> QGroupBox:
        group = QGroupBox("Thêm mã Chi nhánh vào đầu mỗi Mã số KH")
        layout = QVBoxLayout(group)
        self.add_branch_radio = QRadioButton("Thêm mã chi nhánh")
        self.no_branch_radio = QRadioButton("Không thêm mã chi nhánh")
        self.no_branch_radio.setChecked(True)
        layout.addWidget(self.add_branch_radio)
        layout.addWidget(self.no_branch_radio)
        return group

    def _output_prefix_group(self) -> QGroupBox:
        group = QGroupBox("Loại kỳ báo cáo")
        layout = QHBoxLayout(group)
        self.qt_prefix_radio = QRadioButton("Quyết toán năm (QT)")
        self.bn_prefix_radio = QRadioButton("Bán niên (BN)")
        self.qt_prefix_radio.toggled.connect(self._sync_output_prefix)
        self.bn_prefix_radio.toggled.connect(self._sync_output_prefix)
        layout.addWidget(self.qt_prefix_radio)
        layout.addWidget(self.bn_prefix_radio)
        return group

    def _control_sheet_group(self) -> QGroupBox:
        group = QGroupBox("Tạo thêm bảng số liệu phụ để kiểm tra")
        layout = QVBoxLayout(group)
        self.create_control_radio = QRadioButton("Tạo thêm bảng số liệu phụ")
        self.no_control_radio = QRadioButton("Không tạo thêm bảng số liệu phụ")
        self.create_control_radio.setChecked(True)
        layout.addWidget(self.create_control_radio)

        self.accrual_frame = QFrame()
        accrual_layout = QHBoxLayout(self.accrual_frame)
        accrual_layout.setContentsMargins(12, 4, 0, 4)
        self.add_accrual_radio = QRadioButton("Thêm TK dự thu/dự chi")
        self.default_accrual_radio = QRadioButton("Mặc định")
        self.add_accrual_radio.setChecked(True)
        accrual_layout.addWidget(self.add_accrual_radio)
        accrual_layout.addWidget(self.default_accrual_radio)
        self.accrual_frame.setVisible(
            self.spec.report_code.casefold() in {"13", "15a", "15b"}
        )
        layout.addWidget(self.accrual_frame)

        layout.addWidget(self.no_control_radio)
        self.create_control_radio.toggled.connect(self._sync_accrual_options)
        self._sync_accrual_options(self.create_control_radio.isChecked())
        return group

    def _sync_accrual_options(self, enabled: bool) -> None:
        supports_accrual = self.spec.report_code.casefold() in {"13", "15a", "15b"}
        self.accrual_frame.setEnabled(enabled and supports_accrual)
        self.add_accrual_radio.setEnabled(enabled and supports_accrual)
        self.default_accrual_radio.setEnabled(enabled and supports_accrual)

    def _date_format_group(self) -> QGroupBox:
        group = QGroupBox("Định dạng cột ngày tháng năm theo dạng:")
        layout = QVBoxLayout(group)
        self.two_digit_year_radio = QRadioButton("dd/mm/yy (năm thể hiện 2 con số)")
        self.four_digit_year_radio = QRadioButton("dd/mm/yyyy (năm thể hiện 4 con số)")
        self.four_digit_year_radio.setChecked(True)
        layout.addWidget(self.two_digit_year_radio)
        layout.addWidget(self.four_digit_year_radio)
        return group

    def _cleanup_group(self) -> QGroupBox:
        group = QGroupBox("Xóa bỏ các cột thừa sau khi tạo xong mẫu")
        layout = QVBoxLayout(group)
        self.remove_unused_radio = QRadioButton("Xóa bỏ")
        self.keep_unused_radio = QRadioButton("Không xóa bỏ")
        self.remove_unused_radio.setChecked(True)
        layout.addWidget(self.remove_unused_radio)
        layout.addWidget(self.keep_unused_radio)
        return group

    def choose_source_file(self) -> None:
        initial = str(self.source_path.parent) if self.source_path else ""
        file_filter = self._source_file_filter()
        file_name, _ = QFileDialog.getOpenFileName(
            self,
            f"Chọn file nguồn Mẫu {self.spec.report_code}/QT",
            initial,
            file_filter,
        )
        if not file_name:
            return
        self.source_path = Path(file_name)
        self.source_edit.setText(str(self.source_path))
        self._refresh_output_path()

    def accept(self) -> None:
        if self.source_path is None:
            QMessageBox.warning(
                self,
                "Chưa chọn file nguồn",
                f"Hãy chọn file nguồn trước khi tạo Mẫu biểu {self.spec.report_code}/QT.",
            )
            return
        super().accept()

    def output_path(self) -> Path | None:
        if self.source_path is None:
            return None
        return self.source_path.with_name(
            f"{self.profile.branch_code.strip()}{self.output_prefix()}{self.output_report_code()}.xlsx"
        )

    def output_prefix(self) -> str:
        return "BN" if self.bn_prefix_radio.isChecked() else "QT"

    def output_report_code(self) -> str:
        special_codes = {
            "accounting.09a": "9a",
            "accounting.09b": "9b",
            "accounting.09c": "9c",
            "accounting.24": "24a",
        }
        return special_codes.get(self.spec.key, self.spec.report_code)

    def options(self) -> SettlementOptions:
        return SettlementOptions(
            convert_tcvn3_to_unicode=False,
            include_branch_in_customer_id=self.add_branch_radio.isChecked(),
            four_digit_year=self.four_digit_year_radio.isChecked(),
            create_control_sheet=self.create_control_radio.isChecked(),
            remove_unused_columns=self.remove_unused_radio.isChecked(),
            include_accrual_accounts=(
                self.create_control_radio.isChecked()
                and self.add_accrual_radio.isChecked()
            ),
            use_default_accrual_accounts=(
                self.create_control_radio.isChecked()
                and self.default_accrual_radio.isChecked()
            ),
            include_loan_deposit_schedule=self.lds_checkbox.isChecked(),
            output_prefix=self.output_prefix(),
        )

    def _refresh_output_path(self) -> None:
        output_path = self.output_path()
        self.output_edit.setText(str(output_path) if output_path else "")

    def _apply_saved_output_prefix(self) -> None:
        if load_output_prefix() == "BN":
            self.bn_prefix_radio.setChecked(True)
        else:
            self.qt_prefix_radio.setChecked(True)

    def _sync_output_prefix(self) -> None:
        save_output_prefix(self.output_prefix())
        self._refresh_output_path()

    def _source_hint_text(self) -> str:
        return self.spec.source_hint.replace(
            "{MaCN}",
            self.profile.branch_code.strip() or "MaCN",
        )

    def _source_file_filter(self) -> str:
        source_hint = self.spec.source_hint.casefold()
        if ".xls" in source_hint or self.spec.key in {
            "accounting.04",
            "accounting.07a",
            "accounting.08",
            "accounting.09a",
            "accounting.09b",
            "accounting.09c",
        }:
            return (
                "File Excel nguồn (*.xls *.xlsx *.xlsm);;"
                "CSV nguồn quyết toán (*.csv);;"
                "Tất cả file (*.*)"
            )
        return "CSV nguồn quyết toán (*.csv);;Tất cả file (*.*)"
