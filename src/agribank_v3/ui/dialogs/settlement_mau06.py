from __future__ import annotations

from pathlib import Path

from PySide6.QtWidgets import (
    QDialog,
    QDialogButtonBox,
    QFileDialog,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QRadioButton,
    QVBoxLayout,
    QWidget,
    QGroupBox,
)

from agribank_v3.settings import BranchProfile
from agribank_v3.settlement.models import SettlementOptions
from agribank_v3.ui.dialogs.settlement_period import (
    load_output_prefix,
    save_output_prefix,
)


class Mau06SettlementDialog(QDialog):
    def __init__(
        self,
        profile: BranchProfile,
        parent: QWidget | None = None,
        *,
        window_title: str = "Tạo Mẫu biểu Quyết toán 06/QT",
        source_label_text: str | None = None,
        output_label_text: str = "Tên File quyết toán Mẫu 06/QT sẽ được tạo ra:",
        consolidation_output: bool = False,
    ) -> None:
        super().__init__(parent)
        self.profile = profile
        self.source_path: Path | None = None
        self.consolidation_output = consolidation_output

        self.setWindowTitle(window_title)
        self.setModal(True)
        self.setMinimumWidth(690)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(14, 12, 14, 14)
        layout.setSpacing(12)

        source_label = QLabel(
            source_label_text
            or (
                "Tên File nguồn Mẫu 05/QT dùng xử lý để tạo ra Mẫu biểu 06QT là: "
                f"{profile.branch_code.strip()}QT05.xlsx"
            )
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

        layout.addWidget(self._output_prefix_group())

        output_label = QLabel(output_label_text)
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

    def choose_source_file(self) -> None:
        initial = str(self.source_path.parent) if self.source_path else ""
        file_name, _ = QFileDialog.getOpenFileName(
            self,
            "Chọn file Mẫu 05/QT",
            initial,
            "File Mẫu 05/QT (*.xlsx *.xlsm *.xls)",
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
                "Hãy chọn file Mẫu 05/QT trước khi tạo Mẫu biểu 06/QT.",
            )
            return
        super().accept()

    def output_path(self) -> Path | None:
        if self.source_path is None:
            return None
        if self.consolidation_output:
            return self.source_path.with_name(
                f"{self.profile.branch_code.strip()}{self.output_prefix()}Mau06_TongHop.xlsx"
            )
        return self.source_path.with_name(
            f"{self.profile.branch_code.strip()}{self.output_prefix()}06.xlsx"
        )

    def options(self) -> SettlementOptions:
        return SettlementOptions(
            output_prefix=self.output_prefix(),
            source_report_code="consolidation" if self.consolidation_output else "",
        )

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
