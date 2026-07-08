from __future__ import annotations

from pathlib import Path

from PySide6.QtWidgets import (
    QCheckBox,
    QDialog,
    QDialogButtonBox,
    QFileDialog,
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
from agribank_v3.settlement.models import SettlementOptions
from agribank_v3.ui.dialogs.settlement_period import (
    load_output_prefix,
    save_output_prefix,
)


class Mau05SettlementDialog(QDialog):
    def __init__(
        self,
        profile: BranchProfile,
        parent: QWidget | None = None,
    ) -> None:
        super().__init__(parent)
        self.profile = profile
        self.source_path: Path | None = None

        self.setWindowTitle("Tạo Mẫu biểu Quyết toán 05/QT")
        self.setModal(True)
        self.setMinimumWidth(690)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(14, 12, 14, 14)
        layout.setSpacing(12)

        source_label = QLabel(
            "Tên File nguồn CSV dùng xử lý để tạo ra Mẫu biểu 05QT là: "
            f"{profile.branch_code.strip()}_rt05.csv"
        )
        source_label.setStyleSheet("color: #0000ff; font-weight: 700;")
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

        layout.addWidget(self._output_prefix_group())

        output_label = QLabel("Tên File quyết toán Mẫu 05/QT sẽ được tạo ra:")
        output_label.setStyleSheet("color: #0000ff; font-weight: 700;")
        layout.addWidget(output_label)
        self.output_edit = QLineEdit()
        self.output_edit.setReadOnly(True)
        layout.addWidget(self.output_edit)

        note = QLabel(
            "(Chỉ xem xét các file do TW xuất gửi về; mã CN và tên tắt CN "
            "được lấy từ Cài đặt/Thông tin chi nhánh.)"
        )
        note.setStyleSheet("color: #b000a0;")
        note.setWordWrap(True)
        layout.addWidget(note)

        self.convert_checkbox = QCheckBox("Chuyển mã TCVN3 sang Unicode")
        self.convert_checkbox.setEnabled(False)
        self.convert_checkbox.setToolTip(
            "Nguồn rt05 hiện được xử lý theo CSV Unicode. Tùy chọn TCVN3 sẽ "
            "được bật khi có mẫu nguồn cũ cần chuyển mã."
        )
        layout.addWidget(self.convert_checkbox)

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

    def _control_sheet_group(self) -> QGroupBox:
        group = QGroupBox("Tạo thêm bảng số liệu phụ để kiểm tra")
        layout = QVBoxLayout(group)
        self.create_control_radio = QRadioButton("Tạo thêm bảng số liệu phụ")
        self.no_control_radio = QRadioButton("Không tạo thêm bảng số liệu phụ")
        self.create_control_radio.setChecked(True)
        layout.addWidget(self.create_control_radio)
        layout.addWidget(self.no_control_radio)
        return group

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

    def choose_source_file(self) -> None:
        initial = str(self.source_path.parent) if self.source_path else ""
        file_name, _ = QFileDialog.getOpenFileName(
            self,
            "Chọn file nguồn Mẫu 05/QT",
            initial,
            "CSV nguồn Mẫu 05 (*.csv)",
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
                "Hãy chọn file CSV nguồn trước khi tạo Mẫu biểu 05/QT.",
            )
            return
        super().accept()

    def output_path(self) -> Path | None:
        if self.source_path is None:
            return None
        return self.source_path.with_name(
            f"{self.profile.branch_code.strip()}{self.output_prefix()}05.xlsx"
        )

    def output_prefix(self) -> str:
        return "BN" if self.bn_prefix_radio.isChecked() else "QT"

    def options(self) -> SettlementOptions:
        return SettlementOptions(
            convert_tcvn3_to_unicode=False,
            include_branch_in_customer_id=self.add_branch_radio.isChecked(),
            four_digit_year=self.four_digit_year_radio.isChecked(),
            create_control_sheet=self.create_control_radio.isChecked(),
            remove_unused_columns=self.remove_unused_radio.isChecked(),
            include_customer_totals=False,
            remove_customer_total_rows=True,
            output_prefix=self.output_prefix(),
        )

    def _apply_saved_output_prefix(self) -> None:
        if load_output_prefix() == "BN":
            self.bn_prefix_radio.setChecked(True)
        else:
            self.qt_prefix_radio.setChecked(True)

    def _sync_output_prefix(self) -> None:
        save_output_prefix(self.output_prefix())
        self._refresh_output_path()

    def _refresh_output_path(self) -> None:
        output_path = self.output_path()
        self.output_edit.setText(str(output_path) if output_path else "")
