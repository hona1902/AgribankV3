from __future__ import annotations

from pathlib import Path

from PySide6.QtWidgets import (
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
    QSizePolicy,
    QVBoxLayout,
    QWidget,
)

from agribank_v3.settings import BranchProfile
from agribank_v3.settlement.models import SettlementOptions, SettlementSpec
from agribank_v3.ui.dialogs.settlement_period import load_output_prefix, save_output_prefix


class MultiSourceSettlementDialog(QDialog):
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

        self.setWindowTitle(f"Tạo Mẫu biểu Quyết toán {spec.report_code}/QT")
        self.setModal(True)
        self.setMinimumWidth(720)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(14, 12, 14, 14)
        layout.setSpacing(8)

        source_label = QLabel(
            "Danh sách các file nguồn dùng xử lý để tạo ra Mẫu biểu "
            f"{spec.report_code}QT là: {spec.source_hint}"
        )
        source_label.setStyleSheet("color: #0000ff; font-weight: 700;")
        source_label.setWordWrap(True)
        source_label.setContentsMargins(0, 0, 0, 0)
        layout.addWidget(source_label)

        guide_text = self._guide_text()
        if guide_text:
            guide_label = QLabel(guide_text)
            guide_label.setWordWrap(True)
            guide_label.setStyleSheet("color: #0000ff;")
            guide_label.setContentsMargins(0, 0, 0, 0)
            guide_label.setSizePolicy(QSizePolicy.Policy.Preferred, QSizePolicy.Policy.Fixed)
            layout.addWidget(guide_label)

        self.source_list = QListWidget()
        self.source_list.setFixedHeight(78)
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

        layout.addWidget(self._output_prefix_group())

        output_label = QLabel(
            f"Tên File quyết toán Mẫu {spec.report_code}/QT sẽ được tạo ra:"
        )
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

    def choose_source_files(self) -> None:
        initial = str(self.source_paths[-1].parent) if self.source_paths else ""
        file_names, _ = QFileDialog.getOpenFileNames(
            self,
            f"Chọn file nguồn Mẫu {self.spec.report_code}/QT",
            initial,
            "File Excel nguồn (*.xls *.xlsx *.xlsm);;Tất cả file (*.*)",
        )
        if not file_names:
            return
        for file_name in file_names:
            path = Path(file_name)
            if path not in self.source_paths:
                self.source_paths.append(path)
                self.source_list.addItem(str(path))
        self._refresh_output_path()

    def clear_sources(self) -> None:
        self.source_paths.clear()
        self.source_list.clear()
        self._refresh_output_path()

    def accept(self) -> None:
        if not self.source_paths:
            QMessageBox.warning(
                self,
                "Chưa chọn file nguồn",
                f"Hãy chọn ít nhất một file nguồn trước khi tạo Mẫu biểu {self.spec.report_code}/QT.",
            )
            return
        super().accept()

    def output_path(self) -> Path | None:
        if not self.source_paths:
            return None
        return self.source_paths[0].with_name(
            f"{self.profile.branch_code.strip()}{self.output_prefix()}{self.spec.report_code}.xlsx"
        )

    def options(self) -> SettlementOptions:
        return SettlementOptions(output_prefix=self.output_prefix())

    def output_prefix(self) -> str:
        return "BN" if self.bn_prefix_radio.isChecked() else "QT"

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

    def _guide_text(self) -> str:
        if self.spec.key == "accounting.22":
            return (
                "- Chỉ xem xét các File được xuất ra tại màn hình glst34:\n"
                "- Màn hình glst34:\n"
                "   + Kind of Adj: Chọn Prepaid Expenses; Date: Chọn 31/12/xxxx hoặc 30/6/xxxx; "
                "Kind of term: Chọn Year -> Xuất ra file excel: MaCN_rt22a\n"
                "   + Kind of Adj: Chọn Unearned Income; Date: Chọn 31/12/xxxx hoặc 30/6/xxxx; "
                "Kind of term: Chọn Year -> Xuất ra file excel: MaCN_rt22b\n"
                "- Chọn file: sau đó chọn 2 file MaCN_rt22a và MaCN_rt22b trên"
            )
        if self.spec.key == "accounting.23":
            return (
                "- Chỉ xem xét các File được xuất ra tại màn hình glcb06:\n"
                "- Màn hình glcb06:\n"
                "   + Tr.Date: Từ ngày 05/01/xxxx đến ngày 31/12/xxxx; Account: 790008 "
                "-> Xuất ra file Excel: MaCN_rt23-790008\n"
                "   + Tr.Date: Từ ngày 05/01/xxxx đến ngày 31/12/xxxx; Account: 790009 "
                "-> Xuất ra file Excel: MaCN_rt23-790009\n"
                "   + Tr.Date: Từ ngày 05/01/xxxx đến ngày 31/12/xxxx; Account: 899001 "
                "-> Xuất ra file Excel: MaCN_rt23-899001\n"
                "- Chọn file: sau đó chọn 3 file MaCN_rt23-790008; MaCN_rt23-790009; MaCN_rt23-899001\n"
                "- Lưu ý: Nếu không có dữ liệu không cần xuất file Excel"
            )
        return ""
