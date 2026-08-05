from __future__ import annotations

from pathlib import Path

from PySide6.QtCore import Qt, QUrl
from PySide6.QtGui import QDesktopServices
from PySide6.QtWidgets import (
    QApplication,
    QFileDialog,
    QGridLayout,
    QLabel,
    QMenu,
    QMessageBox,
    QProgressBar,
    QSizePolicy,
    QTextEdit,
    QVBoxLayout,
    QDialog,
    QWidget,
)

from agribank_v3.features.credit.summary.customer.repository import CustomerRepository
from agribank_v3.features.credit.summary.customer.widgets import (
    CompactToolbar,
    fit_window_to_screen,
    primary_button,
    secondary_button,
)
from agribank_v3.ui.workers import run_in_thread


class CustomerMaintenanceDialog(QDialog):
    def __init__(self, repository: CustomerRepository, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.repository = repository
        self.last_result: dict[str, object] = {}
        self._worker_thread = None
        self._action_buttons: list[QWidget] = []
        self._database_path_text = ""
        self.setWindowTitle("Bảo trì dữ liệu khách hàng - AgribankV3")
        self.setModal(False)
        self.setSizeGripEnabled(True)
        fit_window_to_screen(
            self,
            width_ratio=0.62,
            height_ratio=0.70,
            max_width=900,
            max_height=720,
            min_width=700,
            min_height=520,
        )
        self._build_ui()
        self.reload()

    def _build_ui(self) -> None:
        layout = QVBoxLayout(self)
        layout.setContentsMargins(14, 14, 14, 14)
        layout.setSpacing(8)

        self.status_grid = QGridLayout()
        self.status_grid.setHorizontalSpacing(10)
        self.status_grid.setVerticalSpacing(5)
        self.status_labels: dict[str, QLabel] = {}
        labels = (
            ("database_path", "Đường dẫn Customer.db"),
            ("size_bytes", "Dung lượng hiện tại"),
            ("period_count", "Số kỳ dữ liệu"),
            ("first_period", "Kỳ đầu tiên"),
            ("last_period", "Kỳ mới nhất"),
            ("master_count", "Số khách hàng master"),
            ("period_summary_count", "Số dòng customer_period_summary"),
            ("officer_period_count", "Số quan hệ khách hàng - cán bộ"),
            ("import_run_count", "Số import run"),
            ("import_file_count", "Số file import"),
            ("override_count", "Số override cán bộ"),
            ("action_log_count", "Số action log"),
            ("officer_directory_count", "Số cán bộ trong danh mục"),
            ("page_count", "SQLite page count"),
            ("freelist_count", "SQLite freelist count"),
            ("reclaimable_bytes", "Dung lượng có khả năng thu hồi"),
            ("last_optimized_at", "Thời điểm tối ưu gần nhất"),
        )
        for index, (key, label_text) in enumerate(labels):
            row = index // 2
            column = 0 if index % 2 == 0 else 2
            label = QLabel(label_text)
            value = QLabel("")
            value.setTextInteractionFlags(
                value.textInteractionFlags()
                | Qt.TextInteractionFlag.TextSelectableByMouse
                | Qt.TextInteractionFlag.TextSelectableByKeyboard
            )
            value.setWordWrap(False)
            value.setMinimumWidth(90)
            value.setSizePolicy(QSizePolicy.Policy.Ignored, QSizePolicy.Policy.Fixed)
            self.status_grid.addWidget(label, row, column)
            self.status_grid.addWidget(value, row, column + 1)
            self.status_labels[key] = value
            if key == "database_path":
                self.database_path_label = value
                value.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
                value.customContextMenuRequested.connect(self._show_database_path_menu)
        self.status_grid.setColumnStretch(1, 1)
        self.status_grid.setColumnStretch(3, 1)
        layout.addLayout(self.status_grid)

        self.result_box = QTextEdit()
        self.result_box.setReadOnly(True)
        self.result_box.setMinimumHeight(140)
        self.result_box.setMaximumHeight(16777215)
        self.result_box.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        self.result_box.setPlaceholderText("Kết quả kiểm tra/tối ưu sẽ hiển thị tại đây.")
        layout.addWidget(self.result_box, stretch=1)

        self.progress = QProgressBar()
        self.progress.setVisible(False)
        layout.addWidget(self.progress)

        self.primary_toolbar = CompactToolbar()
        self.primary_toolbar.setObjectName("CustomerMaintenancePrimaryToolbar")
        self.file_toolbar = CompactToolbar()
        self.file_toolbar.setObjectName("CustomerMaintenanceFileToolbar")
        self.check_button = secondary_button("Kiểm tra cơ sở dữ liệu")
        self.optimize_button = secondary_button("Tối ưu nhanh")
        self.vacuum_button = secondary_button("Thu hồi dung lượng")
        self.backup_button = secondary_button("Sao lưu")
        self.restore_button = secondary_button("Khôi phục")
        self.folder_button = secondary_button("Mở thư mục database")
        self.close_button = primary_button("Đóng")
        self.check_button.clicked.connect(self.quick_check)
        self.optimize_button.clicked.connect(self.optimize_quick)
        self.vacuum_button.clicked.connect(self.vacuum_database)
        self.backup_button.clicked.connect(self.backup_database)
        self.restore_button.clicked.connect(self.restore_database)
        self.folder_button.clicked.connect(self.open_database_folder)
        self.close_button.clicked.connect(self.accept)
        for button in (self.check_button, self.optimize_button, self.vacuum_button):
            self.primary_toolbar.addWidget(button)
        for button in (self.backup_button, self.restore_button, self.folder_button, self.close_button):
            self.file_toolbar.addWidget(button)
        self._action_buttons = [
            self.check_button,
            self.optimize_button,
            self.vacuum_button,
            self.backup_button,
            self.restore_button,
        ]
        layout.addWidget(self.primary_toolbar)
        layout.addWidget(self.file_toolbar)

    def reload(self) -> None:
        status = self.repository.maintenance_status()
        values = {
            "database_path": status.database_path,
            "size_bytes": _format_size(status.size_bytes),
            "period_count": _format_integer_vn(status.period_count),
            "first_period": status.first_period or "N/A",
            "last_period": status.last_period or "N/A",
            "master_count": _format_integer_vn(status.master_count),
            "period_summary_count": _format_integer_vn(status.period_summary_count),
            "officer_period_count": _format_integer_vn(status.officer_period_count),
            "import_run_count": _format_integer_vn(status.import_run_count),
            "import_file_count": _format_integer_vn(status.import_file_count),
            "override_count": _format_integer_vn(status.override_count),
            "action_log_count": _format_integer_vn(status.action_log_count),
            "officer_directory_count": _format_integer_vn(status.officer_directory_count),
            "page_count": _format_integer_vn(status.page_count),
            "freelist_count": _format_integer_vn(status.freelist_count),
            "reclaimable_bytes": _format_size(status.reclaimable_bytes),
            "last_optimized_at": status.last_optimized_at or "Chưa có log",
        }
        for key, value in values.items():
            self._set_status_label(key, value)

    def resizeEvent(self, event) -> None:
        super().resizeEvent(event)
        self._update_database_path_label()

    def _set_status_label(self, key: str, value: object) -> None:
        text = str(value or "")
        label = self.status_labels[key]
        label.setToolTip(text)
        if key == "database_path":
            self._database_path_text = text
            self._update_database_path_label()
            return
        label.setText(text)

    def _update_database_path_label(self) -> None:
        if not self._database_path_text or not hasattr(self, "database_path_label"):
            return
        width = self.database_path_label.width()
        if width <= 0:
            width = max(280, min(520, self.width() - 260))
        self.database_path_label.setText(
            self.database_path_label.fontMetrics().elidedText(
                self._database_path_text,
                Qt.TextElideMode.ElideMiddle,
                max(180, int(width)),
            )
        )

    def _show_database_path_menu(self, position) -> None:
        menu = QMenu(self)
        copy_action = menu.addAction("Sao chép đường dẫn đầy đủ")
        selected = menu.exec(self.database_path_label.mapToGlobal(position))
        if selected == copy_action:
            QApplication.clipboard().setText(self._database_path_text)

    def quick_check(self) -> None:
        self._run_background(
            "Đang kiểm tra...",
            lambda progress: self.repository.check_database(full=False),
            self._check_finished,
            "Kiểm tra cơ sở dữ liệu",
        )

    def optimize_quick(self) -> None:
        self._run_background(
            "Đang tối ưu...",
            lambda progress: self.repository.optimize_database(vacuum=False),
            self._optimize_finished,
            "Tối ưu nhanh",
        )

    def vacuum_database(self) -> None:
        answer = QMessageBox.question(
            self,
            "Thu hồi dung lượng Customer.db",
            "Thu hồi dung lượng có thể mất một khoảng thời gian. "
            "Không được tắt ứng dụng trong quá trình thực hiện. Tiếp tục?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        if answer != QMessageBox.StandardButton.Yes:
            return
        self._run_background(
            "Đang thu hồi dung lượng...",
            lambda progress: self.repository.optimize_database(vacuum=True),
            self._vacuum_finished,
            "Thu hồi dung lượng",
        )

    def backup_database(self) -> None:
        path, _ = QFileDialog.getSaveFileName(
            self,
            "Sao lưu dữ liệu khách hàng",
            "Customer-backup.zip",
            "Backup (*.zip)",
        )
        if not path:
            return
        self._run_background(
            "Đang sao lưu...",
            lambda progress: self.repository.backup_database(Path(path)),
            self._backup_finished,
            "Sao lưu dữ liệu khách hàng",
        )

    def restore_database(self) -> None:
        path, _ = QFileDialog.getOpenFileName(
            self,
            "Khôi phục dữ liệu khách hàng",
            "",
            "Backup (*.zip)",
        )
        if not path:
            return
        answer = QMessageBox.question(
            self,
            "Khôi phục dữ liệu khách hàng",
            "Khôi phục sẽ thay thế Customer.db hiện tại sau khi tạo một bản backup an toàn. Tiếp tục?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        if answer != QMessageBox.StandardButton.Yes:
            return
        parent = self.parent()
        if hasattr(parent, "cancel_period_data_queries"):
            parent.cancel_period_data_queries()
        self._run_background(
            "Đang khôi phục...",
            lambda progress: self.repository.restore_database(Path(path)),
            self._restore_finished,
            "Khôi phục dữ liệu khách hàng",
        )

    def open_database_folder(self) -> None:
        QDesktopServices.openUrl(QUrl.fromLocalFile(str(self.repository.database_path.parent)))

    def _run_background(self, message: str, task, finished, title: str) -> None:
        self.progress.setVisible(True)
        self.progress.setRange(0, 0)
        self.progress.setFormat(message)
        self.result_box.setPlainText(message)
        for button in self._action_buttons:
            button.setEnabled(False)

        def done(result: object) -> None:
            self._finish_background()
            finished(result)
            self.reload()

        def failed(exc: Exception) -> None:
            self._finish_background()
            self.result_box.setPlainText(str(exc))
            QMessageBox.warning(self, title, str(exc))
            self.reload()

        self._worker_thread = run_in_thread(self, task, done, failed, lambda text: self.progress.setFormat(text))

    def _finish_background(self) -> None:
        self.progress.setVisible(False)
        for button in self._action_buttons:
            button.setEnabled(True)

    def _check_finished(self, result: object) -> None:
        data = result if isinstance(result, dict) else {}
        self.last_result = dict(data)
        messages = [str(item) for item in data.get("messages", [])]
        state = "Hợp lệ" if data.get("ok") else "Có lỗi"
        text = "{state}\nChế độ: {mode}\nThời gian: {duration} ms\n{messages}".format(
            state=state,
            mode=data.get("mode", ""),
            duration=_format_integer_vn(data.get("duration_ms", 0)),
            messages="\n".join(messages[:5]),
        )
        self.result_box.setPlainText(text)
        QMessageBox.information(self, "Kiểm tra cơ sở dữ liệu", state)

    def _optimize_finished(self, result: object) -> None:
        data = result if isinstance(result, dict) else {}
        self.last_result = dict(data)
        text = "Đã tối ưu nhanh.\nDung lượng: {before} -> {after}\nThời gian: {duration} ms".format(
            before=_format_size(int(data.get("before_size_bytes", 0))),
            after=_format_size(int(data.get("after_size_bytes", 0))),
            duration=_format_integer_vn(data.get("duration_ms", 0)),
        )
        self.result_box.setPlainText(text)
        QMessageBox.information(self, "Tối ưu nhanh", text)

    def _vacuum_finished(self, result: object) -> None:
        data = result if isinstance(result, dict) else {}
        self.last_result = dict(data)
        text = (
            "Đã thu hồi dung lượng.\n"
            "Dung lượng trước: {before}\n"
            "Dung lượng sau: {after}\n"
            "Đã giảm: {recovered}\n"
            "Thời gian: {duration} ms\n"
            "Backup: {backup}"
        ).format(
            before=_format_size(int(data.get("before_size_bytes", 0))),
            after=_format_size(int(data.get("after_size_bytes", 0))),
            recovered=_format_size(int(data.get("recovered_bytes", 0))),
            duration=_format_integer_vn(data.get("duration_ms", 0)),
            backup=data.get("backup_path") or "N/A",
        )
        self.result_box.setPlainText(text)
        QMessageBox.information(self, "Thu hồi dung lượng", text)

    def _backup_finished(self, result: object) -> None:
        self.last_result = {"backup_path": str(result)}
        self.result_box.setPlainText(f"Đã sao lưu: {result}")
        QMessageBox.information(self, "Sao lưu dữ liệu khách hàng", f"Đã sao lưu: {result}")

    def _restore_finished(self, result: object) -> None:
        self.last_result = {"safety_backup": str(result)}
        self.result_box.setPlainText(f"Đã khôi phục. Backup an toàn: {result}")
        parent = self.parent()
        while parent is not None:
            if hasattr(parent, "refresh_all"):
                parent.refresh_all()
                break
            parent = parent.parent()
        QMessageBox.information(self, "Khôi phục dữ liệu khách hàng", f"Đã khôi phục. Backup an toàn: {result}")


def _format_size(value: int) -> str:
    number = max(0, int(value or 0))
    if number >= 1024 * 1024:
        return f"{number / (1024 * 1024):,.2f} MB".replace(",", "_").replace(".", ",").replace("_", ".")
    if number >= 1024:
        return f"{number / 1024:,.2f} KB".replace(",", "_").replace(".", ",").replace("_", ".")
    return f"{number} bytes"


def _format_integer_vn(value: object) -> str:
    try:
        number = int(value or 0)
    except (TypeError, ValueError):
        number = 0
    return f"{number:,}".replace(",", ".")
