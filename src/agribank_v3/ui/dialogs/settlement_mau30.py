from __future__ import annotations

from pathlib import Path

from PySide6.QtWidgets import (
    QComboBox,
    QDialog,
    QDialogButtonBox,
    QFileDialog,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QCheckBox,
    QVBoxLayout,
    QWidget,
)

from agribank_v3.settings import BranchProfile
from agribank_v3.settlement.models import SettlementOptions


class Mau30SettlementDialog(QDialog):
    def __init__(
        self,
        profile: BranchProfile,
        last_balance_path: Path | None = None,
        parent: QWidget | None = None,
    ) -> None:
        super().__init__(parent)
        self.profile = profile
        self.source_path: Path | None = None
        self.balance_path: Path | None = last_balance_path if last_balance_path and last_balance_path.exists() else None

        self.setWindowTitle("Tạo Mẫu biểu Quyết toán 30/QT")
        self.setModal(True)
        self.setMinimumWidth(690)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(14, 12, 14, 14)
        layout.setSpacing(12)

        model_row = QHBoxLayout()
        model_label = QLabel("Chọn mẫu quyết toán nguồn:")
        self.model_combo = QComboBox()
        self.model_combo.addItems(("05", "15a", "15b", "18", "20a"))
        self.model_combo.currentTextChanged.connect(self._sync_options)
        model_row.addWidget(model_label)
        model_row.addWidget(self.model_combo, 1)
        layout.addLayout(model_row)

        self.include_accrual_checkbox = QCheckBox("Thêm TK dự thu/dự chi")
        self.include_accrual_checkbox.setChecked(True)
        layout.addWidget(self.include_accrual_checkbox)

        source_label = QLabel(
            "Chọn file quyết toán đã tạo để sinh Mẫu biểu 30/QT."
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

        balance_label = QLabel(
            "Chọn file cân đối để tự điền số dư cột E (có thể bỏ trống):"
        )
        balance_label.setStyleSheet("color: #0000ff; font-weight: 700;")
        balance_label.setWordWrap(True)
        layout.addWidget(balance_label)

        balance_row = QHBoxLayout()
        self.balance_edit = QLineEdit()
        self.balance_edit.setReadOnly(True)
        if self.balance_path is not None:
            self.balance_edit.setText(str(self.balance_path))
        balance_row.addWidget(self.balance_edit, 1)
        choose_balance_button = QPushButton("Chọn File")
        choose_balance_button.clicked.connect(self.choose_balance_file)
        balance_row.addWidget(choose_balance_button)
        layout.addLayout(balance_row)

        output_label = QLabel("Mẫu 30/QT sẽ được thêm trực tiếp vào file quyết toán đã chọn:")
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
        self._sync_options(self.model_combo.currentText())

    @property
    def selected_model(self) -> str:
        return self.model_combo.currentText()

    def choose_source_file(self) -> None:
        initial = str(self.source_path.parent) if self.source_path else ""
        file_name, _ = QFileDialog.getOpenFileName(
            self,
            f"Chọn file Mẫu {self.selected_model}/QT",
            initial,
            "File quyết toán (*.xlsx *.xlsm)",
        )
        if not file_name:
            return
        self.source_path = Path(file_name)
        self.source_edit.setText(str(self.source_path))
        self.output_edit.setText(str(self.output_path()))

    def choose_balance_file(self) -> None:
        initial = str(self.balance_path.parent) if self.balance_path else ""
        file_name, _ = QFileDialog.getOpenFileName(
            self,
            "Chọn file cân đối",
            initial,
            "File cân đối (*.xls *.xlsx *.xlsm)",
        )
        if not file_name:
            return
        self.balance_path = Path(file_name)
        self.balance_edit.setText(str(self.balance_path))

    def accept(self) -> None:
        if self.source_path is None:
            QMessageBox.warning(
                self,
                "Chưa chọn file nguồn",
                "Hãy chọn file quyết toán nguồn trước khi tạo Mẫu biểu 30/QT.",
            )
            return
        super().accept()

    def output_path(self) -> Path | None:
        return self.source_path

    def options(self) -> SettlementOptions:
        return SettlementOptions(
            include_accrual_accounts=self.include_accrual_checkbox.isChecked(),
            source_report_code=self.selected_model,
        )

    def _sync_options(self, model: str) -> None:
        enabled = model.casefold() in {"15a", "15b"}
        self.include_accrual_checkbox.setEnabled(enabled)
        self.include_accrual_checkbox.setVisible(enabled)
        if self.source_path is not None:
            self.output_edit.setText(str(self.output_path()))
