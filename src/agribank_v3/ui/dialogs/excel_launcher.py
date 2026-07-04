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

        self.setWindowTitle("Chọn Microsoft Excel")
        self.setModal(True)
        self.setMinimumWidth(590)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(22, 20, 22, 18)
        layout.setSpacing(12)

        title = QLabel("Chọn phiên bản Excel")
        title.setObjectName("PageTitle")
        layout.addWidget(title)

        description = QLabel(
            "Chọn phiên bản Excel muốn dùng. Nếu phiên bản đó đang mở, ứng dụng "
            "sẽ kết nối vào phiên hiện có; nếu chưa mở, ứng dụng sẽ mở Excel và "
            "tạo workbook mới."
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
        self.open_button = QPushButton("Kết nối / mở Excel")
        self.open_button.setObjectName("PrimaryButton")
        self.open_button.clicked.connect(self.connect_or_open_excel)
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

    def connect_or_open_excel(self) -> None:
        installation = self.selected_installation()
        if installation is None:
            return
        self.open_button.setEnabled(False)
        self.version_combo.setEnabled(False)
        self.status_label.setText("Đang kiểm tra phiên Excel đã mở…")
        self.status_label.setStyleSheet("color: #8b5c08; font-weight: 600;")
        major_version = installation.major_version or None
        if self.launch_handle and self.launched_installation == installation:
            self.try_connect_once()
            return
        try:
            self.context = self.excel_service.connect(
                retry_attempts=1,
                create_workbook_if_missing=True,
                required_major_version=major_version,
            )
        except ExcelConnectionError:
            self.context = None
        if self.context is not None:
            self.status_label.setStyleSheet("color: #257047; font-weight: 600;")
            self.status_label.setText(
                f"Đã kết nối {self.context.excel_name}: {self.context.workbook}"
            )
            self.accept()
            return

        try:
            xlstart_report = None
            if installation.major_version and installation.major_version <= 14:
                xlstart_report = self.excel_service.install_tool_addins_to_xlstart()
            self.launch_handle = launch_excel(installation)
        except OSError as exc:
            self.open_button.setEnabled(True)
            self.version_combo.setEnabled(True)
            QMessageBox.warning(
                self,
                "Không thể mở Excel",
                f"Không thể chạy:\n{installation.path}\n\n{exc}",
            )
            return

        self.launched_installation = installation
        if xlstart_report and xlstart_report.failed:
            failed = "; ".join(
                f"{name}: {reason}" for name, reason in xlstart_report.failed
            )
            self.status_label.setText(
                f"Không cài được add-in vào XLSTART: {failed}"
            )
        elif xlstart_report and xlstart_report.discovered:
            self.status_label.setText("Đang mở Excel…")
        else:
            self.status_label.setText("Đang mở Excel…")
        self.status_label.setStyleSheet("color: #8b5c08; font-weight: 600;")
        if installation.major_version <= 14:
            self.open_button.setText("Kết nối lại Excel 2010")
            self.open_button.setEnabled(True)
            self.version_combo.setEnabled(False)
            self.status_label.setText(
                "Excel 2010 đã mở. Bấm “Kết nối lại Excel 2010” khi Excel sẵn sàng."
            )
            return
        timeout_seconds = 45 if installation.major_version <= 14 else 25
        self.deadline = time.monotonic() + timeout_seconds
        self.timer.start()

    def try_connect_once(self) -> None:
        try:
            self.context = self.excel_service.connect(
                retry_attempts=1,
                create_workbook_if_missing=True,
                required_major_version=(
                    self.launched_installation.major_version
                    if self.launched_installation
                    else None
                ),
                preferred_workbook_path=(
                    self.launch_handle.bootstrap_workbook
                    if self.launch_handle
                    else None
                ),
            )
        except ExcelConnectionError as exc:
            self.open_button.setEnabled(True)
            self.version_combo.setEnabled(False)
            self.status_label.setStyleSheet("color: #9b2c2c; font-weight: 600;")
            self.status_label.setText(
                f"Excel đã mở nhưng chưa sẵn sàng: {exc}\n"
                "Hãy đóng hộp thoại trong Excel rồi bấm “Kết nối lại Excel 2010”. "
                "Nếu Excel 2010 đang báo Product Activation Failed hoặc First Run, "
                "cần xử lý thông báo đó trước."
            )
            return

        self.status_label.setStyleSheet("color: #257047; font-weight: 600;")
        self.status_label.setText("Kết nối thành công.")
        self.accept()

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
                preferred_workbook_path=(
                    self.launch_handle.bootstrap_workbook
                    if self.launch_handle
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
                "Hãy đóng các hộp thoại trong Excel rồi thử kết nối lại. "
                "Với Excel 2010, cần xử lý Product Activation/First Run nếu "
                "Excel đang hiện thông báo kích hoạt hoặc cấu hình ban đầu."
            )
            return

        self.timer.stop()
        self.status_label.setStyleSheet("color: #257047; font-weight: 600;")
        self.status_label.setText("Kết nối thành công.")
        self.accept()
