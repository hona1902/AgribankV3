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
    QListWidget,
    QMessageBox,
    QPushButton,
    QRadioButton,
    QVBoxLayout,
    QWidget,
)

from agribank_v3.settings import BranchProfile
from agribank_v3.settlement.models import SettlementOptions, SettlementSpec
from agribank_v3.ui.dialogs.settlement_period import load_output_prefix, save_output_prefix


class ConsolidationCsvDialog(QDialog):
    def __init__(
        self,
        spec: SettlementSpec,
        profile: BranchProfile,
        parent: QWidget | None = None,
    ) -> None:
        super().__init__(parent)
        self.spec = spec
        self.profile = profile
        self.source_paths: list[Path] = []

        self.setWindowTitle(f"Tổng hợp Mẫu biểu Quyết toán {spec.report_code}/QT")
        self.setModal(True)
        self.setMinimumWidth(760)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(14, 12, 14, 14)
        layout.setSpacing(8)

        source_label = QLabel(
            "Chọn các file CSV cùng cấu trúc để nối thành một file Excel, "
            f"sau đó tạo tổng hợp Mẫu {spec.report_code}/QT."
        )
        source_label.setStyleSheet("color: #0000ff; font-weight: 700;")
        source_label.setWordWrap(True)
        layout.addWidget(source_label)

        self.source_list = QListWidget()
        self.source_list.setFixedHeight(96)
        layout.addWidget(self.source_list)

        choose_row = QHBoxLayout()
        choose_button = QPushButton("Chọn File")
        choose_button.clicked.connect(self.choose_source_files)
        clear_button = QPushButton("Xóa danh sách")
        clear_button.clicked.connect(self.clear_sources)
        choose_row.addWidget(choose_button)
        choose_row.addWidget(clear_button)
        choose_row.addStretch()
        layout.addLayout(choose_row)

        output_label = QLabel("File Excel sau khi nối dữ liệu CSV:")
        output_label.setStyleSheet("color: #0000ff; font-weight: 700;")
        layout.addWidget(output_label)
        output_row = QHBoxLayout()
        self.output_edit = QLineEdit()
        self.output_edit.setReadOnly(True)
        output_row.addWidget(self.output_edit, 1)
        choose_output_button = QPushButton("Lưu tại...")
        choose_output_button.clicked.connect(self.choose_output_file)
        output_row.addWidget(choose_output_button)
        layout.addLayout(output_row)

        layout.addWidget(self._output_prefix_group())
        if self.spec.key == "consolidation.05":
            layout.addWidget(self._mau05_options_group())

        buttons = QDialogButtonBox()
        create_button = QPushButton("Nối File và Tổng hợp")
        create_button.setObjectName("PrimaryButton")
        buttons.addButton(create_button, QDialogButtonBox.ButtonRole.AcceptRole)
        cancel_button = QPushButton("Cancel")
        buttons.addButton(cancel_button, QDialogButtonBox.ButtonRole.RejectRole)
        create_button.clicked.connect(self.accept)
        cancel_button.clicked.connect(self.reject)
        layout.addWidget(buttons)
        self._apply_saved_output_prefix()

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

    def _mau05_options_group(self) -> QGroupBox:
        group = QGroupBox("Tùy chọn tạo báo cáo chi tiết Mẫu 05")
        layout = QVBoxLayout(group)
        self.guarantee_owner_checkbox = QCheckBox(
            "Đối với TSĐB là bảo lãnh, chạy sao kê theo họ và tên chính chủ TSĐB"
        )
        self.guarantee_owner_checkbox.setChecked(True)
        self.customer_total_checkbox = QCheckBox("Cộng tổng theo từng khách hàng")
        self.bold_customer_total_checkbox = QCheckBox(
            "Tô đậm dòng tổng cộng theo từng tên khách hàng"
        )
        layout.addWidget(self.guarantee_owner_checkbox)
        layout.addWidget(self.customer_total_checkbox)
        layout.addWidget(self.bold_customer_total_checkbox)
        return group

    def choose_source_files(self) -> None:
        initial = str(self.source_paths[-1].parent) if self.source_paths else ""
        file_names, _ = QFileDialog.getOpenFileNames(
            self,
            f"Chọn các file CSV Mẫu {self.spec.report_code}/QT",
            initial,
            "File CSV (*.csv);;Tất cả file (*.*)",
        )
        if not file_names:
            return
        for file_name in file_names:
            path = Path(file_name)
            if path not in self.source_paths:
                self.source_paths.append(path)
                self.source_list.addItem(str(path))
        self._refresh_output_path()

    def choose_output_file(self) -> None:
        initial = str(self.output_path() or Path.home())
        file_name, _ = QFileDialog.getSaveFileName(
            self,
            "Chọn file Excel sau khi nối",
            initial,
            "Excel Workbook (*.xlsx)",
        )
        if not file_name:
            return
        path = Path(file_name)
        if path.suffix.casefold() != ".xlsx":
            path = path.with_suffix(".xlsx")
        self.output_edit.setText(str(path))

    def clear_sources(self) -> None:
        self.source_paths.clear()
        self.source_list.clear()
        self.output_edit.clear()

    def output_path(self) -> Path | None:
        explicit = self.output_edit.text().strip()
        if explicit:
            return Path(explicit)
        if not self.source_paths:
            return None
        branch_code = self.profile.branch_code.strip() or "MaCN"
        return self.source_paths[0].with_name(
            f"{branch_code}{self.output_prefix()}Mau{self.spec.report_code}_TongHop.xlsx"
        )

    def output_prefix(self) -> str:
        return "BN" if self.bn_prefix_radio.isChecked() else "QT"

    def options(self) -> SettlementOptions:
        branch_in_customer_id_specs = {
            "consolidation.05",
            "consolidation.13",
            "consolidation.14",
            "consolidation.15a",
            "consolidation.15b",
            "consolidation.16",
            "consolidation.18",
        }
        return SettlementOptions(
            convert_tcvn3_to_unicode=False,
            include_branch_in_customer_id=self.spec.key in branch_in_customer_id_specs,
            four_digit_year=True,
            create_control_sheet=self.spec.key == "consolidation.13",
            remove_unused_columns=True,
            use_collateral_owner_for_guarantee=(
                getattr(self, "guarantee_owner_checkbox", None) is None
                or self.guarantee_owner_checkbox.isChecked()
            ),
            include_customer_totals=(
                getattr(self, "customer_total_checkbox", None) is not None
                and self.customer_total_checkbox.isChecked()
            ),
            bold_customer_rows=(
                getattr(self, "bold_customer_total_checkbox", None) is not None
                and self.bold_customer_total_checkbox.isChecked()
            ),
            output_prefix=self.output_prefix(),
        )

    def _apply_saved_output_prefix(self) -> None:
        if load_output_prefix() == "BN":
            self.bn_prefix_radio.setChecked(True)
        else:
            self.qt_prefix_radio.setChecked(True)

    def _sync_output_prefix(self) -> None:
        save_output_prefix(self.output_prefix())
        if self.source_paths:
            self.output_edit.clear()
            self._refresh_output_path()

    def accept(self) -> None:
        if not self.source_paths:
            QMessageBox.warning(
                self,
                "Chưa chọn file nguồn",
                "Hãy chọn ít nhất một file CSV trước khi chạy tổng hợp.",
            )
            return
        output_path = self.output_path()
        if output_path is None:
            QMessageBox.warning(
                self,
                "Chưa có file kết quả",
                "Không xác định được file Excel sau khi nối dữ liệu.",
            )
            return
        self.output_edit.setText(str(output_path))
        super().accept()

    def _refresh_output_path(self) -> None:
        if not self.output_edit.text().strip():
            output_path = self.output_path()
            self.output_edit.setText(str(output_path) if output_path else "")
