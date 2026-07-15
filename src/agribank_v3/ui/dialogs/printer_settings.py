from __future__ import annotations

from PySide6.QtCore import Qt
from PySide6.QtGui import QPixmap
from PySide6.QtWidgets import (
    QDialog,
    QFrame,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QMessageBox,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from agribank_v3.ui.icons import app_icon, icon_path
from agribank_v3.printer_manager import (
    PrinterInfo,
    PrinterManagerError,
    get_installed_printers,
    is_windows,
    open_printer_properties,
    open_printer_queue,
    open_windows_printer_settings,
    print_test_page,
    remove_printer,
    set_default_printer,
)
from agribank_v3.ui.workers import run_in_thread


class PrinterSettingsDialog(QDialog):
    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.printers: tuple[PrinterInfo, ...] = ()
        self._threads: list[object] = []

        self.setObjectName("PrinterSettingsDialog")
        self.setWindowTitle("Cài đặt và quản lý máy in")
        self.setWindowIcon(app_icon())
        self.setModal(False)
        self.setMinimumSize(980, 620)

        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(0)

        header = QFrame()
        header.setObjectName("PrinterHeader")
        header.setFixedHeight(100)
        header_layout = QHBoxLayout(header)
        header_layout.setContentsMargins(24, 16, 24, 16)
        header_layout.setSpacing(16)

        logo = QLabel()
        logo.setObjectName("PrinterLogo")
        pixmap = QPixmap(icon_path("logoagri.png"))
        logo.setPixmap(
            pixmap.scaled(
                62,
                62,
                Qt.AspectRatioMode.KeepAspectRatio,
                Qt.TransformationMode.SmoothTransformation,
            )
        )
        logo.setFixedSize(64, 64)

        brand = QVBoxLayout()
        brand.setSpacing(2)
        title = QLabel("AgribankV3")
        title.setObjectName("PrinterBrandTitle")
        subtitle = QLabel("Cài đặt và quản lý máy in")
        subtitle.setObjectName("PrinterBrandSubtitle")
        brand.addStretch()
        brand.addWidget(title)
        brand.addWidget(subtitle)
        brand.addStretch()

        header_layout.addWidget(logo)
        header_layout.addLayout(brand, stretch=1)
        root.addWidget(header)

        content = QWidget()
        content.setObjectName("PrinterContent")
        layout = QVBoxLayout(content)
        layout.setContentsMargins(28, 18, 28, 20)
        layout.setSpacing(12)

        section_row = QHBoxLayout()
        section_icon = QLabel("▣")
        section_icon.setObjectName("PrinterSectionIcon")
        section_icon.setAlignment(Qt.AlignmentFlag.AlignCenter)
        section_title = QLabel("CÀI ĐẶT VÀ QUẢN LÝ MÁY IN")
        section_title.setObjectName("PrinterSectionTitle")
        self.status_label = QLabel("")
        self.status_label.setObjectName("PrinterStatusLabel")
        self.status_label.setTextInteractionFlags(
            Qt.TextInteractionFlag.TextSelectableByMouse
        )
        refresh_button = QPushButton("Làm mới")
        refresh_button.setObjectName("PrinterRefreshButton")
        refresh_button.setCursor(Qt.CursorShape.PointingHandCursor)
        refresh_button.clicked.connect(self.refresh_printers)
        section_row.addWidget(section_icon)
        section_row.addWidget(section_title)
        section_row.addStretch()
        section_row.addWidget(refresh_button)
        layout.addLayout(section_row)
        layout.addWidget(self.status_label)

        self.table = QTableWidget(0, 6)
        self.table.setHorizontalHeaderLabels(
            ["Tên máy in", "Trạng thái", "Mặc định", "Loại kết nối", "Cổng", "Driver"]
        )
        self.table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.table.setSelectionMode(QTableWidget.SelectionMode.SingleSelection)
        self.table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.table.verticalHeader().setVisible(False)
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        for column in range(1, 6):
            self.table.horizontalHeader().setSectionResizeMode(
                column,
                QHeaderView.ResizeMode.ResizeToContents,
            )
        layout.addWidget(self.table, stretch=1)

        action_bar = QFrame()
        action_bar.setObjectName("PrinterActionBar")
        action_row = QHBoxLayout(action_bar)
        action_row.setContentsMargins(12, 10, 12, 10)
        action_row.setSpacing(8)
        action_label = QLabel("Thao tác")
        action_label.setObjectName("PrinterActionLabel")
        default_button = QPushButton("Đặt làm mặc định")
        default_button.setObjectName("PrinterPrimaryAction")
        default_button.setCursor(Qt.CursorShape.PointingHandCursor)
        default_button.clicked.connect(self.set_selected_default)
        queue_button = QPushButton("Mở hàng đợi in")
        queue_button.setObjectName("PrinterActionButton")
        queue_button.setCursor(Qt.CursorShape.PointingHandCursor)
        queue_button.clicked.connect(self.open_selected_queue)
        properties_button = QPushButton("Thuộc tính máy in")
        properties_button.setObjectName("PrinterActionButton")
        properties_button.setCursor(Qt.CursorShape.PointingHandCursor)
        properties_button.clicked.connect(self.open_selected_properties)
        test_button = QPushButton("In trang kiểm tra")
        test_button.setObjectName("PrinterActionButton")
        test_button.setCursor(Qt.CursorShape.PointingHandCursor)
        test_button.clicked.connect(self.print_selected_test_page)
        remove_button = QPushButton("Xóa máy in")
        remove_button.setObjectName("DangerButton")
        remove_button.setCursor(Qt.CursorShape.PointingHandCursor)
        remove_button.clicked.connect(self.remove_selected_printer)
        action_row.addWidget(action_label)
        action_row.addWidget(default_button)
        action_row.addWidget(queue_button)
        action_row.addWidget(properties_button)
        action_row.addWidget(test_button)
        action_row.addWidget(remove_button)
        action_row.addStretch()
        layout.addWidget(action_bar)

        bottom_row = QHBoxLayout()
        windows_button = QPushButton("Cài đặt Windows")
        windows_button.setObjectName("PrinterSecondaryButton")
        windows_button.setCursor(Qt.CursorShape.PointingHandCursor)
        windows_button.clicked.connect(self.open_windows_settings)
        close_button = QPushButton("Đóng")
        close_button.setObjectName("PrinterCloseButton")
        close_button.setCursor(Qt.CursorShape.PointingHandCursor)
        close_button.setDefault(True)
        close_button.clicked.connect(self.close)
        bottom_row.addWidget(windows_button)
        bottom_row.addStretch()
        bottom_row.addWidget(close_button)
        layout.addLayout(bottom_row)
        root.addWidget(content, stretch=1)

        if not is_windows():
            self._set_status("Chức năng quản lý máy in hiện chỉ hỗ trợ Windows.")
        else:
            self.refresh_printers()
        self._apply_style()

    def refresh_printers(self) -> None:
        self._run_background(
            "Đang tải danh sách máy in...",
            get_installed_printers,
            self._apply_printers,
        )

    def _apply_printers(self, printers: tuple[PrinterInfo, ...]) -> None:
        self.printers = printers
        self.table.setRowCount(0)
        for printer in printers:
            row = self.table.rowCount()
            self.table.insertRow(row)
            values = [
                printer.name,
                printer.status,
                "✓ Mặc định" if printer.is_default else "",
                printer.connection_type or "",
                printer.port_name or "",
                printer.driver_name or "",
            ]
            for column, value in enumerate(values):
                item = QTableWidgetItem(value)
                if printer.is_default:
                    item.setBackground(Qt.GlobalColor.lightGray)
                self.table.setItem(row, column, item)
        self._set_status(
            f"Đã tải {len(printers)} máy in."
            if printers
            else "Windows chưa có máy in nào hoặc không đọc được danh sách."
        )

    def selected_printer(self) -> PrinterInfo | None:
        row = self.table.currentRow()
        if row < 0 or row >= len(self.printers):
            return None
        return self.printers[row]

    def set_selected_default(self) -> None:
        printer = self._require_selected_printer()
        if printer is None:
            return
        answer = QMessageBox.question(
            self,
            "Đặt máy in mặc định",
            f"Đặt máy in “{printer.name}” làm mặc định?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        if answer != QMessageBox.StandardButton.Yes:
            return
        self._run_background(
            "Đang đặt máy in mặc định...",
            lambda: set_default_printer(printer.name),
            lambda _: self._after_success("Đã đặt máy in mặc định.", refresh=True),
        )

    def open_selected_queue(self) -> None:
        printer = self._require_selected_printer()
        if printer is not None:
            self._run_user_action(lambda: open_printer_queue(printer.name))

    def open_selected_properties(self) -> None:
        printer = self._require_selected_printer()
        if printer is not None:
            self._run_user_action(lambda: open_printer_properties(printer.name))

    def print_selected_test_page(self) -> None:
        printer = self._require_selected_printer()
        if printer is not None:
            self._run_user_action(
                lambda: print_test_page(printer.name),
                success="Đã gửi lệnh in trang kiểm tra tới Windows.",
            )

    def remove_selected_printer(self) -> None:
        printer = self._require_selected_printer()
        if printer is None:
            return
        message = f"Bạn chắc chắn muốn xóa máy in “{printer.name}”?"
        if printer.is_default:
            message += "\n\nMáy in này đang là máy in mặc định."
        answer = QMessageBox.warning(
            self,
            "Xóa máy in",
            message,
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        if answer != QMessageBox.StandardButton.Yes:
            return
        self._run_background(
            "Đang xóa máy in...",
            lambda: remove_printer(printer.name),
            lambda _: self._after_success("Đã xóa máy in.", refresh=True),
        )

    def open_windows_settings(self) -> None:
        self._run_user_action(open_windows_printer_settings)

    def _require_selected_printer(self) -> PrinterInfo | None:
        printer = self.selected_printer()
        if printer is None:
            QMessageBox.warning(self, "Chưa chọn máy in", "Hãy chọn một máy in trong danh sách.")
        return printer

    def _run_user_action(self, action, success: str | None = None) -> None:
        try:
            action()
        except PrinterManagerError as exc:
            QMessageBox.warning(self, "Không thực hiện được", str(exc))
            return
        if success:
            QMessageBox.information(self, "Hoàn thành", success)

    def _run_background(self, message: str, function, on_finished) -> None:
        self._set_status(message)

        def finish(payload) -> None:
            self._remove_thread(thread)
            on_finished(payload)

        def fail(exc: Exception) -> None:
            self._remove_thread(thread)
            self._set_status("Thao tác không thành công.")
            QMessageBox.warning(self, "Không thực hiện được", str(exc))

        thread = run_in_thread(self, function, finish, fail)
        self._threads.append(thread)

    def _after_success(self, message: str, *, refresh: bool = False) -> None:
        QMessageBox.information(self, "Hoàn thành", message)
        self._set_status(message)
        if refresh:
            self.refresh_printers()

    def _remove_thread(self, thread) -> None:
        if thread in self._threads:
            self._threads.remove(thread)

    def _set_status(self, message: str) -> None:
        self.status_label.setText(message)

    def _apply_style(self) -> None:
        self.setStyleSheet(
            """
            QDialog#PrinterSettingsDialog {
                background: #fff1f5;
                border: none;
            }
            QFrame#PrinterHeader { background: #831f41; border: none; }
            QLabel#PrinterLogo {
                background: rgba(255, 255, 255, 0.96);
                border: 1px solid rgba(255, 255, 255, 0.50);
                border-radius: 10px;
                padding: 2px;
            }
            QLabel#PrinterBrandTitle {
                color: white; font-size: 25px; font-weight: 750;
            }
            QLabel#PrinterBrandSubtitle { color: #f6dce5; font-size: 13px; }
            QWidget#PrinterContent { background: #fff1f5; }
            QLabel#PrinterSectionIcon {
                color: white; background: #a72a53; border-radius: 11px;
                min-width: 22px; min-height: 22px;
                max-width: 22px; max-height: 22px; font-weight: 700;
            }
            QLabel#PrinterSectionTitle {
                color: #3d1626; font-size: 15px; font-weight: 750;
            }
            QLabel#PrinterStatusLabel {
                color: #5c3745; font-size: 12px;
            }
            QFrame#PrinterActionBar {
                background: #ffffff;
                border: 1px solid #e1e5ea;
                border-radius: 10px;
            }
            QTableWidget {
                background: white;
                alternate-background-color: #fafbfc;
                border: 1px solid #e1e5ea;
                border-radius: 10px;
                gridline-color: #e6eaee;
            }
            QHeaderView::section {
                color: #38434d;
                background: #f3f5f7;
                border: none;
                border-right: 1px solid #dde3e8;
                border-bottom: 1px solid #d7dee5;
                padding: 8px 7px;
                font-weight: 700;
            }
            QLabel#PrinterActionLabel {
                color: #3d1626;
                background: #fff1f5;
                border: 1px solid #efc6d5;
                border-radius: 7px;
                padding: 8px 12px;
                font-weight: 750;
            }
            QPushButton#PrinterRefreshButton,
            QPushButton#PrinterSecondaryButton,
            QPushButton#PrinterActionButton {
                color: #831f41;
                background: #ffffff;
                border: 1px solid #d9dfe5;
                border-radius: 7px;
                padding: 9px 14px;
                font-weight: 650;
            }
            QPushButton#PrinterRefreshButton:hover,
            QPushButton#PrinterSecondaryButton:hover,
            QPushButton#PrinterActionButton:hover {
                background: #fff4f7;
                border-color: #bd6b87;
            }
            QPushButton#PrinterPrimaryAction {
                color: white;
                background: #931f49;
                border: none;
                border-radius: 7px;
                padding: 10px 16px;
                font-weight: 700;
            }
            QPushButton#PrinterPrimaryAction:hover {
                background: #ad2c57;
            }
            QPushButton#PrinterCloseButton {
                color: white;
                background: #931f49;
                border: none;
                border-radius: 7px;
                min-width: 78px;
                padding: 10px 18px;
                font-weight: 650;
            }
            QPushButton#PrinterCloseButton:hover { background: #ad2c57; }
            QPushButton#DangerButton {
                color: #9b1c1c;
                background: #fff0f0;
                border: 1px solid #efc8c8;
                border-radius: 7px;
                padding: 9px 15px;
                font-weight: 650;
            }
            QPushButton#DangerButton:hover {
                color: #ffffff;
                background: #b42323;
                border-color: #b42323;
            }
            """
        )
