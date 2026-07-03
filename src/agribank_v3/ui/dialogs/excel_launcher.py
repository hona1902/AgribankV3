from __future__ import annotations

import time

from PySide6.QtCore import QTimer
from PySide6.QtWidgets import (
    QComboBox,
    QDialog,
    QDialogButtonBox,
    QLabel,
    QMessageBox,
    QPushButton,
    QVBoxLayout,
    QWidget,
)

from agribank_v3.excel import (
    ExcelConnectionError,
    ExcelContext,
    ExcelInstallation,
    ExcelLaunchHandle,
    ExcelService,
    discover_excel_installations,
    launch_excel,
)


class ExcelLauncherDialog(QDialog):
    def __init__(
        self,
        excel_service: ExcelService,
        parent: QWidget | None = None,
    ) -> None:
        super().__init__(parent)
        self.excel_service = excel_service
        self.context: ExcelContext | None = None
        self.launch_handle: ExcelLaunchHandle | None = None
        self.installations = discover_excel_installations()
        self.launched_installation: ExcelInstallation | None = None
        self.deadline = 0.0

        self.setWindowTitle("Mở Microsoft Excel")
        self.setModal(True)
        self.setMinimumWidth(590)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(22, 20, 22, 18)
        layout.setSpacing(12)

        title = QLabel("Chọn phiên bản Excel")
        title.setObjectName("PageTitle")
        layout.addWidget(title)

        description = QLabel(
            "AgribankV3 chưa tìm thấy phiên Excel đang chạy. Chọn một phiên bản "
            "đã cài; ứng dụng sẽ mở Excel và tạo workbook mới."
        )
        description.setObjectName("MutedText")
        description.setWordWrap(True)
        layout.addWidget(description)

        self.version_combo = QComboBox()
        for installation in self.installations:
            self.version_combo.addItem(installation.display_name, installation)
        self.version_combo.currentIndexChanged.connect(self.update_path)
        layout.addWidget(self.version_combo)

        self.path_label = QLabel()
        self.path_label.setObjectName("MutedText")
        self.path_label.setWordWrap(True)
        layout.addWidget(self.path_label)

        self.status_label = QLabel("")
        self.status_label.setWordWrap(True)
        layout.addWidget(self.status_label)

        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Cancel)
        buttons.rejected.connect(self.reject)
        self.open_button = QPushButton("Mở Excel và tạo workbook")
        self.open_button.setObjectName("PrimaryButton")
        self.open_button.clicked.connect(self.open_excel)
        buttons.addButton(self.open_button, QDialogButtonBox.ButtonRole.AcceptRole)
        layout.addWidget(buttons)

        self.timer = QTimer(self)
        self.timer.setInterval(400)
        self.timer.timeout.connect(self.try_connect)

        self.open_button.setEnabled(bool(self.installations))
        if self.installations:
            self.update_path()
        else:
            self.path_label.setText(
                "Không tìm thấy EXCEL.EXE trong Registry hoặc thư mục Microsoft Office."
            )

    def selected_installation(self) -> ExcelInstallation | None:
        return self.version_combo.currentData()

    def update_path(self) -> None:
        installation = self.selected_installation()
        self.path_label.setText(str(installation.path) if installation else "")

    def open_excel(self) -> None:
        installation = self.selected_installation()
        if installation is None:
            return
        try:
            self.launch_handle = launch_excel(installation)
        except OSError as exc:
            QMessageBox.warning(
                self,
                "Không thể mở Excel",
                f"Không thể chạy:\n{installation.path}\n\n{exc}",
            )
            return

        self.launched_installation = installation
        self.open_button.setEnabled(False)
        self.version_combo.setEnabled(False)
        self.status_label.setText(
            "Đang chờ Excel khởi động và tạo workbook mới…"
        )
        self.status_label.setStyleSheet("color: #8b5c08; font-weight: 600;")
        self.deadline = time.monotonic() + 20
        self.timer.start()

    def try_connect(self) -> None:
        try:
            self.context = self.excel_service.connect(
                retry_attempts=1,
                create_workbook_if_missing=True,
                required_major_version=(
                    self.launched_installation.major_version
                    if self.launched_installation
                    else None
                ),
            )
        except ExcelConnectionError as exc:
            if time.monotonic() < self.deadline:
                return
            self.timer.stop()
            self.open_button.setEnabled(True)
            self.version_combo.setEnabled(True)
            self.status_label.setStyleSheet("color: #9b2c2c; font-weight: 600;")
            self.status_label.setText(
                f"Excel đã mở nhưng chưa sẵn sàng: {exc}\n"
                "Hãy đóng các hộp thoại trong Excel rồi thử kết nối lại."
            )
            return

        self.timer.stop()
        self.status_label.setStyleSheet("color: #257047; font-weight: 600;")
        self.status_label.setText(
            f"Đã kết nối {self.context.excel_name}: {self.context.workbook}"
        )
        self.accept()
