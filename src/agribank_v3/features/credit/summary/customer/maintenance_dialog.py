from __future__ import annotations

from pathlib import Path

from PySide6.QtCore import Qt, QUrl
from PySide6.QtGui import QDesktopServices
from PySide6.QtWidgets import (
    QFileDialog,
    QGridLayout,
    QLabel,
    QMessageBox,
    QProgressBar,
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
        self.setWindowTitle("Bảo trì dữ liệu khách hàng - AgribankV3")
        self.setModal(False)
        fit_window_to_screen(
            self,
            width_ratio=0.58,
            height_ratio=0.70,
            max_width=820,
            max_height=720,
            min_width=620,
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
        for row, (key, label_text) in enumerate(labels):
            label = QLabel(label_text)
            value = QLabel("")
            value.setTextInteractionFlags(value.textInteractionFlags() | Qt.TextInteractionFlag.TextSelectableByMouse)
            value.setWordWrap(True)
            self.status_grid.addWidget(label, row, 0)
            self.status_grid.addWidget(value, row, 1)
            self.status_labels[key] = value
        self.status_grid.setColumnStretch(1, 1)
        layout.addLayout(self.status_grid)

        self.result_box = QTextEdit()
        self.result_box.setReadOnly(True)
        self.result_box.setMaximumHeight(120)
        self.result_box.setPlaceholderText("Kết quả kiểm tra/tối ưu sẽ hiển thị tại đây.")
        layout.addWidget(self.result_box)

        self.progress = QProgressBar()
        self.progress.setVisible(False)
        layout.addWidget(self.progress)

        toolbar = CompactToolbar()
        check_button = secondary_button("Kiểm tra cơ sở dữ liệu")
        optimize_button = secondary_button("Tối ưu nhanh")
        vacuum_button = secondary_button("Thu hồi dung lượng")
        backup_button = secondary_button("Sao lưu")
        restore_button = secondary_button("Khôi phục")
        folder_button = secondary_button("Mở thư mục database")
        close_button = primary_button("Đóng")
        check_button.clicked.connect(self.quick_check)
        optimize_button.clicked.connect(self.optimize_quick)
        vacuum_button.clicked.connect(self.vacuum_database)
        backup_button.clicked.connect(self.backup_database)
        restore_button.clicked.connect(self.restore_database)
        folder_button.clicked.connect(self.open_database_folder)
        close_button.clicked.connect(self.accept)
        for button in (
            check_button,
            optimize_button,
            vacuum_button,
            backup_button,
            restore_button,
            folder_button,
            close_button,
        ):
            toolbar.addWidget(button)
        self._action_buttons = [check_button, optimize_button, vacuum_button, backup_button, restore_button]
        layout.addWidget(toolbar)

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
            self.status_labels[key].setText(value)

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
