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
    QVBoxLayout,
    QWidget,
)

from agribank_v3.settings import BranchProfile


class Mau06SettlementDialog(QDialog):
    def __init__(
        self,
        profile: BranchProfile,
        parent: QWidget | None = None,
    ) -> None:
        super().__init__(parent)
        self.profile = profile
        self.source_path: Path | None = None

        self.setWindowTitle("Tạo Mẫu biểu Quyết toán 06/QT")
        self.setModal(True)
        self.setMinimumWidth(690)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(14, 12, 14, 14)
        layout.setSpacing(12)

        source_label = QLabel(
            "Tên File nguồn Mẫu 05/QT dùng xử lý để tạo ra Mẫu biểu 06QT là: "
            f"{profile.branch_code.strip()}QT05.xlsx"
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

        output_label = QLabel("Tên File quyết toán Mẫu 06/QT sẽ được tạo ra:")
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
        self.output_edit.setText(str(self.output_path()))

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
        return self.source_path.with_name(
            f"{self.profile.branch_code.strip()}QT06.xlsx"
        )
