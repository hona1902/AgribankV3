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
from agribank_v3.settlement.models import SettlementOptions, SettlementSpec


class Mau30SettlementDialog(QDialog):
    def __init__(
        self,
        profile: BranchProfile,
        spec: SettlementSpec,
        last_balance_path: Path | None = None,
        parent: QWidget | None = None,
    ) -> None:
        super().__init__(parent)
        self.profile = profile
        self.spec = spec
        self.source_path: Path | None = None
        self.uses_balance_folder = self.spec.key == "consolidation.30a"
        if last_balance_path and last_balance_path.exists():
            if self.uses_balance_folder:
                self.balance_path = last_balance_path if last_balance_path.is_dir() else None
            else:
                self.balance_path = last_balance_path if last_balance_path.is_file() else None
        else:
            self.balance_path = None

        self.setWindowTitle("Tạo Mẫu biểu Quyết toán 30/QT")
        self.setModal(True)
        self.setMinimumWidth(690)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(14, 12, 14, 14)
        layout.setSpacing(12)

        model_row = QHBoxLayout()
        model_label = QLabel("Chọn mẫu quyết toán nguồn:")
        self.model_combo = QComboBox()
        self.model_combo.addItems(self._available_models())
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

        balance_label = QLabel(self._balance_source_label())
        balance_label.setStyleSheet("color: #0000ff; font-weight: 700;")
        balance_label.setWordWrap(True)
        layout.addWidget(balance_label)
        self.balance_status_label = QLabel("")
        self.balance_status_label.setStyleSheet("color: #c00000; font-weight: 700;")
        self.balance_status_label.setWordWrap(True)
        self.balance_status_label.setVisible(self.uses_balance_folder)
        layout.addWidget(self.balance_status_label)

        balance_row = QHBoxLayout()
        self.balance_edit = QLineEdit()
        self.balance_edit.setReadOnly(True)
        if self.balance_path is not None:
            self.balance_edit.setText(str(self.balance_path))
        balance_row.addWidget(self.balance_edit, 1)
        choose_balance_button = QPushButton(
            "Chọn thư mục" if self.uses_balance_folder else "Chọn File"
        )
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
        self._refresh_balance_status()

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
        if self.uses_balance_folder:
            initial = str(self.balance_path) if self.balance_path else ""
            folder_name = QFileDialog.getExistingDirectory(
                self,
                "Chọn thư mục chứa các file cân đối",
                initial,
            )
            if not folder_name:
                return
            self.balance_path = Path(folder_name)
            self.balance_edit.setText(str(self.balance_path))
            self._refresh_balance_status()
            return
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
        enabled = model.casefold() in {"13", "15a", "15b"}
        self.include_accrual_checkbox.setEnabled(enabled)
        self.include_accrual_checkbox.setVisible(
            self.spec.key == "consolidation.30a" or enabled
        )
        if self.source_path is not None:
            self.output_edit.setText(str(self.output_path()))

    def _available_models(self) -> tuple[str, ...]:
        if self.spec.key == "consolidation.30a":
            return ("05", "13", "15a", "15b", "18", "22", "23", "24", "20a")
        if self.spec.key == "accounting.30a":
            return ("13", "22", "23", "24")
        return ("05", "15a", "15b", "18", "20a")

    def _balance_source_label(self) -> str:
        if self.uses_balance_folder:
            return (
                "Chọn thư mục chứa các file cân đối để tự điền số dư cột E "
                "(file có dạng MaCN.xls hoặc MaCN.xlsx, có thể bỏ trống):"
            )
        return "Chọn file cân đối để tự điền số dư cột E (có thể bỏ trống):"

    def _refresh_balance_status(self) -> None:
        if not self.uses_balance_folder:
            return
        branch_codes = self._detected_balance_branch_codes()
        if branch_codes:
            self.balance_status_label.setText(
                f"Đã nhận diện {len(branch_codes):02d} file cân đối của các chi nhánh: "
                + ", ".join(branch_codes)
            )
        else:
            self.balance_status_label.setText(
                "Chưa nhận diện được file cân đối hợp lệ trong thư mục đã chọn."
            )

    def _detected_balance_branch_codes(self) -> tuple[str, ...]:
        if self.balance_path is None or not self.balance_path.is_dir():
            return ()
        branch_codes = []
        for path in sorted(self.balance_path.iterdir(), key=lambda item: item.name.casefold()):
            if (
                path.is_file()
                and not path.name.startswith("~$")
                and path.suffix.casefold() in {".xls", ".xlsx", ".xlsm"}
                and path.stem.strip().isdigit()
            ):
                branch_codes.append(path.stem.strip())
        return tuple(branch_codes)
