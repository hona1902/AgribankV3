from __future__ import annotations

from datetime import date
from pathlib import Path
from typing import Callable

from PySide6.QtCore import QObject, QThread, QTimer, Qt, Signal, Slot
from PySide6.QtWidgets import (
    QApplication,
    QCheckBox,
    QComboBox,
    QFileDialog,
    QFormLayout,
    QGroupBox,
    QHBoxLayout,
    QInputDialog,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QScrollArea,
    QTabWidget,
    QTableWidget,
    QTableWidgetItem,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from update_builder_core import (
    DEFAULT_UPDATE_PATH,
    AutoBuildConfig,
    BuildUpdateConfig,
    MigrationItem,
    UpdateBuilder,
    UpdateBuilderError,
    auto_build_update,
    compare_databases_for_migration,
    create_package_plan,
    create_suggested_python_migration,
    detect_previous_release_version,
    format_package_report,
    find_dev_database,
    insert_python_migration_into_source,
    package_name,
    read_current_version,
    save_schema_snapshot,
    test_generated_python_migration,
    write_package_file_list,
)


class UpdateTaskWorker(QObject):
    progress = Signal(str)
    progress_percent = Signal(int)
    finished = Signal(object)
    failed = Signal(str)

    def __init__(self, task: Callable[[Callable[[str], None], Callable[[int], None]], object]) -> None:
        super().__init__()
        self.task = task

    def run(self) -> None:
        try:
            result = self.task(self.progress.emit, self.progress_percent.emit)
        except Exception as exc:
            self.failed.emit(str(exc))
            return
        self.finished.emit(result)


class UpdateBuilderWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("AgribankV3 Update Builder")
        self.resize(980, 700)
        self.generated_migration_path: Path | None = None
        self.generated_migration_version = ""
        self.detected_dangerous_schema_changes: list[str] = []
        self.log_edit: QTextEdit | None = None
        self.dev_db_path: Path | None = None
        self.previous_release_version: str | None = None
        self.allow_rebuild_same_version = False
        self._worker_thread: QThread | None = None
        self._worker: UpdateTaskWorker | None = None
        self._worker_on_success: Callable[[object], None] | None = None
        self._worker_on_failure: Callable[[str], None] | None = None
        self._worker_title = ""
        self._worker_result: object | None = None
        self._worker_error: str | None = None
        self._worker_succeeded = False
        self._is_running = False
        self._busy_widgets: list[QWidget] = []
        self._build_ui()

    def _build_ui(self) -> None:
        root = QWidget()
        layout = QVBoxLayout(root)
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(8)
        tabs = QTabWidget()

        quick_group = QGroupBox("Chế độ tạo nhanh")
        quick_layout = QVBoxLayout(quick_group)
        quick_layout.setContentsMargins(10, 10, 10, 10)
        quick_layout.setSpacing(8)
        quick_hint = QLabel(
            "Chọn source, nhập version/ghi chú rồi bấm tạo tự động. Builder sẽ tự kiểm tra database và tạo migration khi an toàn."
        )
        quick_hint.setWordWrap(True)
        self.dev_db_label = QLabel("Database dev hiện tại: chưa kiểm tra")
        quick_buttons = QHBoxLayout()
        quick_buttons.setSpacing(6)
        self.choose_dev_db_button = QPushButton("Chọn database dev")
        self.choose_dev_db_button.clicked.connect(self.choose_dev_database)
        self.baseline_button = QPushButton("Thiết lập baseline database")
        self.baseline_button.clicked.connect(self.setup_baseline_database)
        self.auto_build_button = QPushButton("Tạo bản cập nhật tự động")
        self.auto_build_button.setObjectName("PrimaryButton")
        self.auto_build_button.clicked.connect(self.build_update_auto)
        quick_buttons.addWidget(self.choose_dev_db_button)
        quick_buttons.addWidget(self.baseline_button)
        quick_buttons.addStretch()
        quick_buttons.addWidget(self.auto_build_button)
        quick_layout.addWidget(quick_hint)
        quick_layout.addWidget(self.dev_db_label)
        quick_layout.addLayout(quick_buttons)

        source_group = QGroupBox("Thông tin source")
        source_form = QFormLayout(source_group)
        source_form.setContentsMargins(10, 10, 10, 10)
        source_form.setSpacing(8)
        self.source_edit = QLineEdit(str(Path.cwd()))
        source_row = QHBoxLayout()
        source_row.setSpacing(6)
        source_button = QPushButton("Chọn...")
        source_button.clicked.connect(self.choose_source)
        read_button = QPushButton("Đọc phiên bản")
        read_button.clicked.connect(self.read_version)
        source_row.addWidget(self.source_edit, stretch=1)
        source_row.addWidget(source_button)
        source_row.addWidget(read_button)
        self.current_version_label = QLabel("Chưa đọc")
        self.previous_release_label = QLabel("Chưa xác định")
        source_form.addRow("Thư mục source AgribankV3", source_row)
        source_form.addRow("Phiên bản trong source", self.current_version_label)
        source_form.addRow("Phiên bản phát hành trước", self.previous_release_label)

        update_group = QGroupBox("Thông tin bản cập nhật")
        update_form = QFormLayout(update_group)
        update_form.setContentsMargins(10, 10, 10, 10)
        update_form.setSpacing(8)
        self.new_version_edit = QLineEdit()
        self.release_date_edit = QLineEdit(date.today().isoformat())
        self.restart_checkbox = QCheckBox("Yêu cầu khởi động lại")
        self.restart_checkbox.setChecked(True)
        self.auto_version_checkbox = QCheckBox(
            "Tự cập nhật version trong source trước khi đóng gói"
        )
        self.package_mode_combo = QComboBox()
        self.package_mode_combo.addItem("Gói runtime tối thiểu", "runtime")
        self.package_mode_combo.addItem("Gói app đã build EXE", "app")
        self.package_mode_combo.addItem("Gói source đầy đủ", "source")
        self.package_mode_combo.addItem("Gói delta - chỉ file thay đổi", "delta")
        self.package_mode_combo.setCurrentIndex(0)
        self.notes_edit = QTextEdit()
        self.notes_edit.setPlaceholderText("Mỗi dòng là một ghi chú trong manifest.json")
        self.notes_edit.setMinimumHeight(90)
        self.notes_edit.setMaximumHeight(120)
        update_form.addRow("Phiên bản mới", self.new_version_edit)
        update_form.addRow("Ngày phát hành", self.release_date_edit)
        update_form.addRow("", self.restart_checkbox)
        update_form.addRow("", self.auto_version_checkbox)
        update_form.addRow("Kiểu gói cập nhật", self.package_mode_combo)
        update_form.addRow("Ghi chú cập nhật", self.notes_edit)

        migration_group = QGroupBox("Database migration")
        migration_layout = QVBoxLayout(migration_group)
        migration_layout.setContentsMargins(10, 10, 10, 10)
        migration_layout.setSpacing(8)
        self.has_migration_checkbox = QCheckBox("Có thay đổi database")
        warning = QLabel(
            "Migration phải an toàn: không DROP bảng, không ghi đè dữ liệu người dùng."
        )
        warning.setWordWrap(True)
        warning.setStyleSheet("color: #9a3412; font-weight: 600;")
        self.migration_table = QTableWidget(0, 4)
        self.migration_table.setHorizontalHeaderLabels(
            ["Version", "File SQL", "Description", "Python migration"]
        )
        self.migration_table.setEnabled(False)
        self.migration_table.setMinimumHeight(150)
        migration_buttons = QHBoxLayout()
        migration_buttons.setSpacing(6)
        add_migration_button = QPushButton("Thêm migration")
        add_migration_button.clicked.connect(self.add_migration_row)
        remove_migration_button = QPushButton("Xóa migration")
        remove_migration_button.clicked.connect(self.remove_migration_row)
        choose_sql_button = QPushButton("Chọn file .sql")
        choose_sql_button.clicked.connect(self.choose_migration_sql)
        migration_buttons.addWidget(add_migration_button)
        migration_buttons.addWidget(remove_migration_button)
        migration_buttons.addWidget(choose_sql_button)
        migration_buttons.addStretch()
        self.has_migration_checkbox.toggled.connect(self.migration_table.setEnabled)
        migration_layout.addWidget(self.has_migration_checkbox)
        migration_layout.addWidget(warning)
        migration_layout.addWidget(self.migration_table)
        migration_layout.addLayout(migration_buttons)

        generated_group = QGroupBox("Tạo migration database từ database mẫu")
        generated_layout = QVBoxLayout(generated_group)
        generated_layout.setContentsMargins(10, 10, 10, 10)
        generated_layout.setSpacing(8)
        generated_form = QFormLayout()
        generated_form.setSpacing(8)
        self.old_db_edit = QLineEdit()
        self.old_db_edit.setPlaceholderText("Chọn DuLieuV3.db của phiên bản cũ hoặc schema cũ")
        self.new_db_edit = QLineEdit()
        self.new_db_edit.setPlaceholderText("Chọn DuLieuV3.db của phiên bản mới/dev")
        old_db_row = QHBoxLayout()
        old_db_button = QPushButton("Chọn...")
        old_db_button.clicked.connect(lambda: self.choose_database(self.old_db_edit))
        old_db_row.addWidget(self.old_db_edit, stretch=1)
        old_db_row.addWidget(old_db_button)
        new_db_row = QHBoxLayout()
        new_db_button = QPushButton("Chọn...")
        new_db_button.clicked.connect(lambda: self.choose_database(self.new_db_edit))
        new_db_row.addWidget(self.new_db_edit, stretch=1)
        new_db_row.addWidget(new_db_button)
        self.generated_version_edit = QLineEdit()
        self.generated_description_edit = QLineEdit()
        self.auto_insert_migration_checkbox = QCheckBox(
            "Tự thêm migration Python vào src/agribank_v3/update/db_migrations.py"
        )
        generated_form.addRow("Database phiên bản cũ", old_db_row)
        generated_form.addRow("Database phiên bản mới", new_db_row)
        generated_form.addRow("Version migration", self.generated_version_edit)
        generated_form.addRow("Description", self.generated_description_edit)
        generated_form.addRow("", self.auto_insert_migration_checkbox)
        generated_buttons = QHBoxLayout()
        self.compare_db_button = QPushButton("So sánh database")
        self.compare_db_button.clicked.connect(self.compare_databases)
        self.generate_button = QPushButton("Tạo migration gợi ý")
        self.generate_button.clicked.connect(self.generate_suggested_migration)
        self.test_button = QPushButton("Thử chạy migration trên bản copy")
        self.test_button.clicked.connect(self.test_generated_migration)
        self.add_generated_button = QPushButton("Thêm migration vào danh sách")
        self.add_generated_button.clicked.connect(self.add_generated_migration_to_table)
        generated_buttons.addWidget(self.compare_db_button)
        generated_buttons.addWidget(self.generate_button)
        generated_buttons.addWidget(self.test_button)
        generated_buttons.addWidget(self.add_generated_button)
        generated_buttons.addStretch()
        self.diff_result_edit = QTextEdit()
        self.diff_result_edit.setReadOnly(True)
        self.diff_result_edit.setMinimumHeight(130)
        self.diff_result_edit.setMaximumHeight(160)
        generated_layout.addLayout(generated_form)
        generated_layout.addLayout(generated_buttons)
        generated_layout.addWidget(self.diff_result_edit)
        migration_layout.addWidget(generated_group)

        target_group = QGroupBox("Thư mục xuất bản cập nhật")
        target_form = QFormLayout(target_group)
        target_form.setContentsMargins(10, 10, 10, 10)
        target_form.setSpacing(8)
        self.update_path_edit = QLineEdit(DEFAULT_UPDATE_PATH)
        target_row = QHBoxLayout()
        target_row.setSpacing(6)
        target_button = QPushButton("Chọn...")
        target_button.clicked.connect(self.choose_update_path)
        target_row.addWidget(self.update_path_edit, stretch=1)
        target_row.addWidget(target_button)
        target_form.addRow("Thư mục Update", target_row)

        self.log_edit = QTextEdit()
        self.log_edit.setReadOnly(True)
        self.log_edit.setMinimumHeight(420)

        action_row = QHBoxLayout()
        action_row.setSpacing(6)
        self.validate_button = QPushButton("Kiểm tra dữ liệu")
        self.validate_button.clicked.connect(self.validate_config)
        self.preview_files_button = QPushButton("Xem file sẽ đóng gói")
        self.preview_files_button.clicked.connect(self.preview_package_files)
        self.build_button = QPushButton("Tạo bản cập nhật")
        self.build_button.setObjectName("PrimaryButton")
        self.build_button.clicked.connect(self.build_update)
        open_button = QPushButton("Mở thư mục Update")
        open_button.clicked.connect(self.open_update_folder)
        clear_log_button = QPushButton("Xóa log")
        clear_log_button.clicked.connect(self.clear_log)
        close_button = QPushButton("Đóng")
        close_button.clicked.connect(self.close)
        action_row.addWidget(self.validate_button)
        action_row.addWidget(self.preview_files_button)
        action_row.addWidget(self.build_button)
        action_row.addWidget(open_button)
        action_row.addWidget(clear_log_button)
        action_row.addStretch()
        action_row.addWidget(close_button)

        info_tab = QWidget()
        info_layout = QVBoxLayout(info_tab)
        info_layout.setContentsMargins(10, 10, 10, 10)
        info_layout.setSpacing(8)
        info_layout.addWidget(quick_group)
        info_layout.addWidget(source_group)
        info_layout.addWidget(update_group)
        info_layout.addWidget(target_group)
        info_layout.addStretch()

        migration_content = QWidget()
        migration_content_layout = QVBoxLayout(migration_content)
        migration_content_layout.setContentsMargins(0, 0, 0, 0)
        migration_content_layout.setSpacing(8)
        migration_content_layout.addWidget(migration_group)
        migration_content_layout.addStretch()
        migration_scroll = QScrollArea()
        migration_scroll.setWidgetResizable(True)
        migration_scroll.setFrameShape(QScrollArea.Shape.NoFrame)
        migration_scroll.setWidget(migration_content)

        log_tab = QWidget()
        log_layout = QVBoxLayout(log_tab)
        log_layout.setContentsMargins(10, 10, 10, 10)
        log_layout.setSpacing(8)
        log_layout.addLayout(action_row)
        log_layout.addWidget(QLabel("Log"))
        log_layout.addWidget(self.log_edit, stretch=1)

        self.tabs = tabs
        self.tabs.addTab(info_tab, "Thông tin cập nhật")
        self.tabs.addTab(migration_scroll, "Database migration")
        self.tabs.addTab(log_tab, "Tạo gói & Log")
        self.log_tab_index = 2
        self.status_label = QLabel("Sẵn sàng")
        log_layout.insertWidget(0, self.status_label)
        self._busy_widgets = [
            self.choose_dev_db_button,
            self.baseline_button,
            self.auto_build_button,
            self.compare_db_button,
            self.generate_button,
            self.test_button,
            self.add_generated_button,
            self.validate_button,
            self.preview_files_button,
            self.build_button,
        ]

        layout.addWidget(self.tabs)
        self.setCentralWidget(root)
        self.refresh_dev_database()

    def _start_worker(
        self,
        title: str,
        task: Callable[[Callable[[str], None], Callable[[int], None]], object],
        on_success: Callable[[object], None],
        on_failure: Callable[[str], None] | None = None,
    ) -> None:
        if self._is_running or self._worker_thread is not None:
            self._set_status("Đang xử lý tác vụ khác. Vui lòng chờ.")
            self.log("Một tác vụ khác đang chạy. Đã bỏ qua yêu cầu mới.")
            return
        self._is_running = True
        self._worker_title = title
        self._worker_result = None
        self._worker_error = None
        self._worker_succeeded = False
        self._worker_on_success = on_success
        self._worker_on_failure = on_failure
        self._set_busy(True, title)
        if hasattr(self, "tabs"):
            self.tabs.setCurrentIndex(self.log_tab_index)
        thread = QThread(self)
        worker = UpdateTaskWorker(task)
        worker.moveToThread(thread)
        self._worker_thread = thread
        self._worker = worker
        thread.started.connect(worker.run)
        worker.progress.connect(self.log)
        worker.progress_percent.connect(self._set_progress_percent)
        worker.finished.connect(self._worker_finished)
        worker.failed.connect(self._worker_failed)
        worker.finished.connect(thread.quit)
        worker.failed.connect(thread.quit)
        worker.finished.connect(worker.deleteLater)
        worker.failed.connect(worker.deleteLater)
        thread.finished.connect(self._worker_thread_finished)
        thread.finished.connect(thread.deleteLater)
        thread.start()

    @Slot(object)
    def _worker_finished(self, result: object) -> None:
        self._worker_succeeded = True
        self._worker_result = result
        if self._worker_thread is None:
            self._worker_thread_finished()

    @Slot(str)
    def _worker_failed(self, message: str) -> None:
        self._worker_succeeded = False
        self._worker_error = message
        if self._worker_thread is None:
            self._worker_thread_finished()

    @Slot()
    def _worker_thread_finished(self) -> None:
        title = self._worker_title or "Tác vụ"
        succeeded = self._worker_succeeded
        result = self._worker_result
        message = self._worker_error or "Tác vụ không thành công."
        on_success = self._worker_on_success
        on_failure = self._worker_on_failure
        self._cleanup_worker("Hoàn thành." if succeeded else "Có lỗi.")
        if succeeded:
            if on_success is not None:
                QTimer.singleShot(
                    0,
                    lambda result=result, on_success=on_success, title=title: (
                        self._run_success_callback(title, on_success, result)
                    ),
                )
            return
        handler = on_failure or (lambda text: self._show_error("Lỗi", text))
        QTimer.singleShot(
            0,
            lambda message=message, handler=handler, title=title: (
                self._run_failure_callback(title, handler, message)
            ),
        )

    def _run_success_callback(
        self,
        title: str,
        callback: Callable[[object], None],
        result: object,
    ) -> None:
        try:
            callback(result)
        except Exception as exc:
            self._show_task_error(title, str(exc))

    def _run_failure_callback(
        self,
        title: str,
        callback: Callable[[str], None],
        message: str,
    ) -> None:
        try:
            callback(message)
        except Exception as exc:
            self._show_task_error(title, str(exc))

    def _cleanup_worker(self, status: str) -> None:
        self._set_busy(False, status)
        self._is_running = False
        self._worker_thread = None
        self._worker = None
        self._worker_on_success = None
        self._worker_on_failure = None
        self._worker_title = ""
        self._worker_result = None
        self._worker_error = None
        self._worker_succeeded = False

    def _set_busy(self, busy: bool, status: str) -> None:
        for widget in self._busy_widgets:
            widget.setEnabled(not busy)
        self._set_status(status)

    def _set_status(self, status: str) -> None:
        if hasattr(self, "status_label"):
            self.status_label.setText(status)

    @Slot(int)
    def _set_progress_percent(self, value: int) -> None:
        self._set_status(f"Đang xử lý... {value}%")

    def _validate_form_basic(self) -> str | None:
        source_text = self.source_edit.text().strip()
        if not source_text:
            self.source_edit.setFocus()
            return "Thư mục source AgribankV3 không được để trống."
        source_path = Path(source_text)
        if not source_path.is_dir():
            self.source_edit.setFocus()
            return f"Không tìm thấy thư mục source: {source_path}"
        if not (source_path / "src" / "agribank_v3").is_dir():
            self.source_edit.setFocus()
            return (
                "Thư mục source không hợp lệ. Vui lòng chọn thư mục gốc "
                "AgribankV3 có src/agribank_v3."
            )

        new_version = self.new_version_edit.text().strip()
        if not new_version:
            self.new_version_edit.setFocus()
            return "Phiên bản mới không được để trống."
        try:
            package_name(new_version)
        except UpdateBuilderError as exc:
            self.new_version_edit.setFocus()
            return str(exc)

        notes_text = self.notes_edit.toPlainText().strip()
        if not notes_text:
            self.tabs.setCurrentIndex(0)
            self.notes_edit.setFocus()
            return "Release notes không được để trống."

        update_path = self.update_path_edit.text().strip()
        if not update_path:
            self.update_path_edit.setFocus()
            return "Thư mục Update không được để trống."
        try:
            update_path_obj = Path(update_path)
        except Exception as exc:
            self.update_path_edit.setFocus()
            return f"Thư mục Update không hợp lệ: {exc}"
        drive_error = self._missing_drive_error(update_path_obj)
        if drive_error:
            self.update_path_edit.setFocus()
            return drive_error
        return None

    def _show_basic_validation_error(self, title: str, message: str) -> None:
        del title
        self._set_busy(False, "Có lỗi.")
        self.log(f"Lỗi kiểm tra dữ liệu: {message}")

    def _show_warning(self, title: str, message: str) -> None:
        self._set_busy(False, "Có cảnh báo.")
        self._is_running = False
        self.log(f"Cảnh báo: {message}")
        QMessageBox.warning(self, title, message)

    def _show_error(self, title: str, message: str) -> None:
        self._set_busy(False, "Có lỗi.")
        self._is_running = False
        self.log(f"Lỗi: {message}")
        QMessageBox.critical(self, title, message)

    def _show_info(self, title: str, message: str) -> None:
        self._set_busy(False, "Hoàn thành.")
        self.log(message)
        QMessageBox.information(self, title, message)

    def _missing_drive_error(self, path: Path) -> str | None:
        drive = path.drive
        if not drive:
            return None
        try:
            exists = Path(drive + "\\").exists()
        except OSError:
            exists = False
        if exists:
            return None
        return (
            f"Không tìm thấy ổ đĩa {drive}. "
            "Vui lòng kiểm tra kết nối mạng hoặc chọn thư mục Update khác."
        )

    def _prepare_basic_form_for_action(self, title: str) -> bool:
        error = self._validate_form_basic()
        if error:
            self._show_basic_validation_error(title, error)
            return False
        return self._ensure_update_folder_ready(title)

    def _ensure_update_folder_ready(self, title: str) -> bool:
        update_path = Path(self.update_path_edit.text().strip())
        drive_error = self._missing_drive_error(update_path)
        if drive_error:
            self.update_path_edit.setFocus()
            self._show_basic_validation_error(title, drive_error)
            return False
        try:
            if update_path.exists():
                if not update_path.is_dir():
                    self.update_path_edit.setFocus()
                    self._show_warning(
                        title,
                        f"Đường dẫn Update không phải thư mục: {update_path}",
                    )
                    return False
                return True
        except OSError:
            self.update_path_edit.setFocus()
            self._show_warning(
                title,
                f"Không thể kiểm tra thư mục Update: {update_path}",
            )
            return False
        answer = QMessageBox.question(
            self,
            title,
            (
                f"Không tìm thấy thư mục Update:\n{update_path}\n\n"
                "Bạn có muốn tạo thư mục này không?"
            ),
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.Yes,
        )
        if answer != QMessageBox.StandardButton.Yes:
            self.update_path_edit.setFocus()
            self._set_status("Đã hủy do thư mục Update chưa tồn tại.")
            self.log(f"Chưa tạo thư mục Update: {update_path}")
            return False
        try:
            update_path.mkdir(parents=True, exist_ok=True)
        except OSError as exc:
            self.update_path_edit.setFocus()
            self.log(f"Chi tiết lỗi tạo thư mục Update: {exc}")
            self._show_error(
                title,
                (
                    "Không thể tạo thư mục Update. "
                    "Vui lòng chọn thư mục khác.\n"
                    f"{update_path}"
                ),
            )
            return False
        self.log(f"Đã tạo thư mục Update: {update_path}")
        return True

    def choose_source(self) -> None:
        selected = QFileDialog.getExistingDirectory(
            self,
            "Chọn thư mục source AgribankV3",
            self.source_edit.text(),
        )
        if selected:
            self.source_edit.setText(selected)
            self.refresh_dev_database()
            self.read_version()

    def choose_update_path(self) -> None:
        selected = QFileDialog.getExistingDirectory(
            self,
            "Chọn thư mục Update",
            self.update_path_edit.text(),
        )
        if selected:
            self.update_path_edit.setText(selected)
            self.refresh_previous_release_version()

    def read_version(self) -> None:
        try:
            version = read_current_version(Path(self.source_edit.text()))
        except UpdateBuilderError as exc:
            self.current_version_label.setText("Không đọc được")
            self.log(str(exc))
            return
        self.current_version_label.setText(version)
        self.refresh_previous_release_version()
        if not self.new_version_edit.text().strip():
            self.new_version_edit.setText(version)
        if not self.generated_version_edit.text().strip():
            self.generated_version_edit.setText(self.new_version_edit.text().strip())
        if not self.generated_description_edit.text().strip():
            self.generated_description_edit.setText(
                f"Migration database cho phiên bản {self.generated_version_edit.text().strip()}"
            )
        self.log(f"Đọc version trong source: {version}")
        self.log(
            "Version phát hành trước: "
            + (self.previous_release_version or "không xác định")
        )

    def choose_database(self, target_edit: QLineEdit) -> None:
        selected, _ = QFileDialog.getOpenFileName(
            self,
            "Chọn SQLite database",
            self.source_edit.text(),
            "SQLite database (*.db *.sqlite *.sqlite3);;Tất cả tệp (*)",
        )
        if selected:
            target_edit.setText(selected)

    def choose_dev_database(self) -> None:
        selected, _ = QFileDialog.getOpenFileName(
            self,
            "Chọn database dev hiện tại",
            self.source_edit.text(),
            "SQLite database (*.db *.sqlite *.sqlite3);;Tất cả tệp (*)",
        )
        if selected:
            self.dev_db_path = Path(selected)
            self.dev_db_label.setText(f"Database dev hiện tại: {self.dev_db_path}")
            self.new_db_edit.setText(selected)
            self.log(f"Đã chọn database dev: {self.dev_db_path}")

    def refresh_dev_database(self) -> None:
        found = find_dev_database(Path(self.source_edit.text()))
        self.dev_db_path = found
        if found is None:
            self.dev_db_label.setText("Database dev hiện tại: chưa tìm thấy")
            self.log(
                "Không tìm thấy database dev để so sánh. Bản cập nhật sẽ được coi là "
                "chỉ đổi code, trừ khi người dùng chọn database thủ công."
            )
            return
        self.dev_db_label.setText(f"Database dev hiện tại: {found}")
        self.new_db_edit.setText(str(found))

    def refresh_previous_release_version(self) -> None:
        new_version = self.new_version_edit.text().strip() or None
        self.previous_release_version = detect_previous_release_version(
            Path(self.update_path_edit.text().strip() or DEFAULT_UPDATE_PATH),
            Path(self.source_edit.text()),
            new_version,
        )
        self.previous_release_label.setText(
            self.previous_release_version or "Chưa xác định"
        )

    def setup_baseline_database(self) -> None:
        selected, _ = QFileDialog.getOpenFileName(
            self,
            "Chọn database của bản đang phát hành",
            self.source_edit.text(),
            "SQLite database (*.db *.sqlite *.sqlite3);;Tất cả tệp (*)",
        )
        if not selected:
            return
        default_version = self.current_version_label.text()
        if default_version == "Chưa đọc":
            try:
                default_version = read_current_version(Path(self.source_edit.text()))
            except UpdateBuilderError:
                default_version = ""
        version, ok = QInputDialog.getText(
            self,
            "Thiết lập baseline database",
            "Version baseline:",
            text=default_version,
        )
        if not ok:
            return
        try:
            snapshot = save_schema_snapshot(selected, version.strip())
        except Exception as exc:
            self.log(f"Lỗi thiết lập baseline database: {exc}")
            self._show_warning(
                "Thiết lập baseline database",
                "Không thể lưu schema snapshot. Vui lòng kiểm tra lại database baseline.",
            )
            return
        self.old_db_edit.setText(selected)
        self.log(f"Đã lưu baseline database {version.strip()}: {snapshot}")
        self._show_info(
            "Thiết lập baseline database",
            f"Đã lưu schema snapshot:\n{snapshot}",
        )

    def compare_databases(self) -> None:
        old_db = self.old_db_edit.text().strip()
        new_db = self.new_db_edit.text().strip()

        def task(log, progress):
            del progress
            log("Đang so sánh database...")
            return compare_databases_for_migration(old_db, new_db)

        self._start_worker(
            "Đang so sánh database, vui lòng chờ...",
            task,
            self._on_compare_databases_success,
            lambda message: self._show_task_error("So sánh database", message),
        )

    def _on_compare_databases_success(self, diff: object) -> None:
        self.detected_dangerous_schema_changes = [
            f"{item.kind}: {item.name} - {item.detail}"
            for item in diff.dangerous_changes
        ]
        self.diff_result_edit.setPlainText(self._format_diff(diff))
        self.log("Đã so sánh database.")
        if not self.generated_version_edit.text().strip():
            self.generated_version_edit.setText(self.new_version_edit.text().strip())
        if not self.generated_description_edit.text().strip():
            self.generated_description_edit.setText(
                f"Migration database cho phiên bản {self.generated_version_edit.text().strip()}"
            )

    def generate_suggested_migration(self) -> None:
        version = self.generated_version_edit.text().strip() or self.new_version_edit.text().strip()
        if not version:
            self._show_warning("Tạo migration gợi ý", "Chưa nhập version migration.")
            return
        old_db = self.old_db_edit.text().strip()
        new_db = self.new_db_edit.text().strip()
        source_path = Path(self.source_edit.text().strip())
        auto_insert = self.auto_insert_migration_checkbox.isChecked()

        def task(log, progress):
            del progress
            log("Đang tạo migration gợi ý...")
            path, diff = create_suggested_python_migration(
                old_db_path=old_db,
                new_db_path=new_db,
                version=version,
                output_dir=source_path
                / "tools"
                / "update_builder"
                / "generated_migrations",
            )
            backup = None
            if auto_insert:
                backup = insert_python_migration_into_source(
                    source_path=source_path,
                    generated_migration_file=path,
                    version=version,
                )
            return path, diff, backup

        self._start_worker(
            "Đang tạo migration gợi ý, vui lòng chờ...",
            task,
            lambda result: self._on_generate_suggested_migration_success(version, result),
            lambda message: self._show_task_error("Tạo migration gợi ý", message),
        )

    def _on_generate_suggested_migration_success(
        self,
        version: str,
        result: object,
    ) -> None:
        path, diff, backup = result
        if backup is not None:
            self.log(f"Đã chèn migration vào db_migrations.py. Backup: {backup}")
        self.generated_migration_path = path
        self.generated_migration_version = version
        self.detected_dangerous_schema_changes = [
            f"{item.kind}: {item.name} - {item.detail}"
            for item in diff.dangerous_changes
        ]
        self.diff_result_edit.setPlainText(self._format_diff(diff) + f"\n\nĐã tạo file:\n{path}")
        self.log(f"Đã tạo migration gợi ý: {path}")

    def test_generated_migration(self) -> None:
        if self.generated_migration_path is None:
            self._show_warning(
                "Thử migration",
                "Chưa có file migration gợi ý. Hãy bấm 'Tạo migration gợi ý' trước.",
            )
            return
        old_db = self.old_db_edit.text().strip()
        new_db = self.new_db_edit.text().strip()
        migration_file = self.generated_migration_path
        version = self.generated_migration_version

        def task(log, progress):
            del progress
            log("Đang thử migration trên bản copy...")
            result = test_generated_python_migration(
                old_db_path=old_db,
                new_db_path=new_db,
                migration_file=migration_file,
                version=version,
            )
            return result

        self._start_worker(
            "Đang thử migration trên bản copy, vui lòng chờ...",
            task,
            self._on_test_generated_migration_success,
            lambda message: self._show_task_error("Thử migration", message),
        )

    def _on_test_generated_migration_success(self, result: object) -> None:
        if result.success:
            message = f"Thử migration thành công trên bản copy: {result.test_database_path}"
        else:
            message = (
                f"Thử migration chưa đạt. Bản copy: {result.test_database_path}\n"
                + (f"Lỗi: {result.error}\n" if result.error else "")
                + "\n".join(result.missing_items)
            )
        self.diff_result_edit.append("\n" + message)
        self.log(message)

    def add_generated_migration_to_table(self) -> None:
        version = self.generated_migration_version or self.generated_version_edit.text().strip()
        if not version:
            self._show_warning("Thêm migration", "Chưa có version migration.")
            return
        row = self.migration_table.rowCount()
        self.migration_table.insertRow(row)
        self.migration_table.setItem(row, 0, QTableWidgetItem(version))
        self.migration_table.setItem(row, 1, QTableWidgetItem(""))
        self.migration_table.setItem(
            row,
            2,
            QTableWidgetItem(
                self.generated_description_edit.text().strip()
                or f"Migration database cho phiên bản {version}"
            ),
        )
        python_item = QTableWidgetItem()
        python_item.setFlags(python_item.flags() | Qt.ItemFlag.ItemIsUserCheckable)
        python_item.setCheckState(Qt.CheckState.Checked)
        self.migration_table.setItem(row, 3, python_item)
        self.has_migration_checkbox.setChecked(True)
        self.log(f"Đã thêm Python migration {version} vào manifest.")

    def add_migration_row(self) -> None:
        row = self.migration_table.rowCount()
        self.migration_table.insertRow(row)
        self.migration_table.setItem(row, 0, QTableWidgetItem(self.new_version_edit.text().strip()))
        self.migration_table.setItem(row, 1, QTableWidgetItem(""))
        self.migration_table.setItem(row, 2, QTableWidgetItem(""))
        python_item = QTableWidgetItem()
        python_item.setFlags(python_item.flags() | Qt.ItemFlag.ItemIsUserCheckable)
        python_item.setCheckState(Qt.CheckState.Unchecked)
        self.migration_table.setItem(row, 3, python_item)
        self.has_migration_checkbox.setChecked(True)

    def remove_migration_row(self) -> None:
        row = self.migration_table.currentRow()
        if row >= 0:
            self.migration_table.removeRow(row)

    def choose_migration_sql(self) -> None:
        row = self.migration_table.currentRow()
        if row < 0:
            self.add_migration_row()
            row = self.migration_table.rowCount() - 1
        selected, _ = QFileDialog.getOpenFileName(
            self,
            "Chọn file migration SQL",
            self.source_edit.text(),
            "SQL migration (*.sql);;Tất cả tệp (*)",
        )
        if selected:
            self.migration_table.setItem(row, 1, QTableWidgetItem(selected))

    def preview_package_files(self) -> None:
        if not self._prepare_basic_form_for_action("Xem file sẽ đóng gói"):
            return
        self.refresh_previous_release_version()
        config = self._config()

        def task(log, progress):
            del progress
            log("Đang lập danh sách file sẽ đóng gói...")
            plan = create_package_plan(config, self.previous_release_version)
            output_path = (
                Path(self.update_path_edit.text().strip())
                / f"package_files_{config.new_version}.tsv"
            )
            file_list_path = write_package_file_list(output_path, plan.included)
            return plan, file_list_path

        self._start_worker(
            "Đang lập danh sách file sẽ đóng gói...",
            task,
            self._on_preview_package_files_success,
            lambda message: self._show_task_error("Xem file sẽ đóng gói", message),
        )

    def _on_preview_package_files_success(self, result: object) -> None:
        plan, file_list_path = result
        self.log(format_package_report(plan.report))
        self.log(f"Danh sách file đã lưu: {file_list_path}")
        self.log("relative_path | size_bytes | include_reason")
        for item in list(plan.included)[:200]:
            try:
                size = item.path.stat().st_size
            except OSError:
                size = 0
            self.log(
                f"{item.relative_path.as_posix()} | {size} | {item.include_reason}"
            )
        if len(plan.included) > 200:
            self.log(f"... còn {len(plan.included) - 200} file trong file TSV.")
        if plan.report.total_size > 100 * 1024 * 1024:
            self._show_warning(
                "Xem file sẽ đóng gói",
                "Gói cập nhật có dung lượng lớn. Vui lòng kiểm tra danh sách file trước khi phát hành.",
            )

    def validate_config(self) -> None:
        if not self._prepare_basic_form_for_action("Kiểm tra dữ liệu"):
            return
        if not self._confirm_rebuild_same_version_if_needed():
            return
        config = self._config()

        def task(log, progress):
            del progress
            log("Đang kiểm tra dữ liệu...")
            return UpdateBuilder(log).validate_build_deep(config)

        self._start_worker(
            "Đang kiểm tra dữ liệu, vui lòng chờ...",
            task,
            self._on_validate_success,
            lambda message: self._show_task_error("Kiểm tra dữ liệu", message),
        )

    def _on_validate_success(self, result: object) -> None:
        self.log(f"Kiểm tra dữ liệu OK. Version trong source: {result.current_version}")
        self.log(
            "Version phát hành trước: "
            + (result.previous_release_version or "không xác định")
        )
        for warning in result.warnings:
            self.log(f"Cảnh báo: {warning}")
        self._show_info("Kiểm tra dữ liệu", "Dữ liệu hợp lệ.")

    def build_update_auto(self) -> None:
        if not self._prepare_basic_form_for_action("Tạo bản cập nhật tự động"):
            return
        if not self._confirm_rebuild_same_version_if_needed():
            return
        self.clear_log()
        self._start_auto_build_worker(
            code_only_if_missing_baseline=False,
            baseline_db_path=None,
        )

    def _start_auto_build_worker(
        self,
        *,
        code_only_if_missing_baseline: bool,
        baseline_db_path: Path | None,
    ) -> None:
        config = self._auto_build_config(
            code_only_if_missing_baseline=code_only_if_missing_baseline,
            baseline_db_path=baseline_db_path,
        )

        def task(log, progress):
            return auto_build_update(config, log, progress)

        self._start_worker(
            "Đang tạo bản cập nhật tự động, vui lòng chờ...",
            task,
            self._on_auto_build_success,
            self._on_auto_build_failed,
        )

    def _on_auto_build_success(self, result: object) -> None:
        self._show_info(
            "Tạo bản cập nhật tự động",
            f"Đã tạo bản cập nhật:\n{result.build_result.package_path}\n\nManifest:\n{result.build_result.manifest_path}",
        )

    def _on_auto_build_failed(self, message: str) -> None:
        if "Không đủ dữ liệu để tự kiểm tra thay đổi database" not in message:
            self._show_task_error("Tạo bản cập nhật tự động", message)
            return
        action = self._ask_missing_database_action()
        if action == "cancel":
            self.log("Đã hủy tạo bản cập nhật tự động.")
            return
        if action == "code_only":
            self._start_auto_build_worker(
                code_only_if_missing_baseline=True,
                baseline_db_path=None,
            )
            return
        baseline_db = self._choose_baseline_for_auto_build()
        if baseline_db is None:
            self.log("Chưa chọn database baseline. Đã hủy tạo tự động.")
            return
        if self.dev_db_path is None:
            self.choose_dev_database()
        if self.dev_db_path is None:
            self.log("Chưa chọn database dev. Đã hủy tạo tự động.")
            return
        self._start_auto_build_worker(
            code_only_if_missing_baseline=False,
            baseline_db_path=baseline_db,
        )

    def _auto_build_config(
        self,
        *,
        code_only_if_missing_baseline: bool,
        baseline_db_path: Path | None = None,
    ) -> AutoBuildConfig:
        notes = tuple(
            line.strip()
            for line in self.notes_edit.toPlainText().splitlines()
            if line.strip()
        )
        source_path = Path(self.source_edit.text().strip())
        return AutoBuildConfig(
            source_path=source_path,
            update_path=Path(self.update_path_edit.text().strip() or DEFAULT_UPDATE_PATH),
            new_version=self.new_version_edit.text().strip(),
            release_date=self.release_date_edit.text().strip(),
            required_app_restart=self.restart_checkbox.isChecked(),
            notes=notes,
            auto_update_source_version=self.auto_version_checkbox.isChecked(),
            dev_db_path=self.dev_db_path,
            baseline_db_path=baseline_db_path,
            generated_migration_dir=source_path
            / "tools"
            / "update_builder"
            / "generated_migrations",
            code_only_if_missing_baseline=code_only_if_missing_baseline,
            previous_release_version=self.previous_release_version,
            allow_rebuild_same_version=self.allow_rebuild_same_version,
            package_mode=self._package_mode(),
        )

    def _package_mode(self) -> str:
        return str(self.package_mode_combo.currentData() or "runtime")

    def _confirm_rebuild_same_version_if_needed(self) -> bool:
        self.refresh_previous_release_version()
        new_version = self.new_version_edit.text().strip()
        if (
            not new_version
            or not self.previous_release_version
            or new_version != self.previous_release_version
        ):
            self.allow_rebuild_same_version = False
            return True
        answer = QMessageBox.question(
            self,
            "Tạo lại cùng version",
            (
                f"Bản cập nhật {new_version} đã tồn tại trong version phát hành trước. "
                "Bạn có muốn backup bản cũ và tạo lại cùng version để test không?"
            ),
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        self.allow_rebuild_same_version = answer == QMessageBox.StandardButton.Yes
        if not self.allow_rebuild_same_version:
            self.log("Đã hủy tạo lại bản cập nhật cùng version.")
        return self.allow_rebuild_same_version

    def _ask_missing_database_action(self) -> str:
        message = QMessageBox(self)
        message.setWindowTitle("Tạo bản cập nhật tự động")
        message.setText(
            "Không đủ dữ liệu để tự kiểm tra thay đổi database. Bạn muốn làm gì?"
        )
        choose_button = message.addButton(
            "Chọn database bản cũ và bản mới",
            QMessageBox.ButtonRole.AcceptRole,
        )
        code_only_button = message.addButton(
            "Tiếp tục chỉ đổi code",
            QMessageBox.ButtonRole.DestructiveRole,
        )
        cancel_button = message.addButton(
            "Hủy",
            QMessageBox.ButtonRole.RejectRole,
        )
        message.exec()
        clicked = message.clickedButton()
        if clicked == choose_button:
            return "choose"
        if clicked == code_only_button:
            return "code_only"
        if clicked == cancel_button:
            return "cancel"
        return "cancel"

    def _choose_baseline_for_auto_build(self) -> Path | None:
        selected, _ = QFileDialog.getOpenFileName(
            self,
            "Chọn database bản cũ để tạo baseline",
            self.source_edit.text(),
            "SQLite database (*.db *.sqlite *.sqlite3);;Tất cả tệp (*)",
        )
        if not selected:
            return None
        self.old_db_edit.setText(selected)
        return Path(selected)

    def build_update(self) -> None:
        if not self._prepare_basic_form_for_action("Tạo bản cập nhật"):
            return
        if not self._confirm_rebuild_same_version_if_needed():
            return
        config = self._config()

        def task(log, progress):
            del progress
            log("Đang kiểm tra dữ liệu trước khi tạo...")
            return UpdateBuilder(log).validate_build_deep(config)

        self._start_worker(
            "Đang kiểm tra dữ liệu trước khi tạo, vui lòng chờ...",
            task,
            lambda validation: self._confirm_and_start_manual_build(config, validation),
            lambda message: self._show_task_error("Tạo bản cập nhật", message),
        )

    def _confirm_and_start_manual_build(
        self,
        config: BuildUpdateConfig,
        validation,
    ) -> None:
        warnings = list(validation.warnings)
        if self.detected_dangerous_schema_changes:
            warnings.append(
                "Có thay đổi schema nguy hiểm cần xử lý thủ công. "
                "Vui lòng kiểm tra trước khi phát hành."
            )
            warnings.extend(self.detected_dangerous_schema_changes)
        if warnings:
            answer = QMessageBox.warning(
                self,
                "Cảnh báo trước khi tạo",
                "\n".join(warnings) + "\n\nTiếp tục tạo bản cập nhật?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No,
            )
            if answer != QMessageBox.StandardButton.Yes:
                return

        def task(log, progress):
            return UpdateBuilder(log, progress).build(config)

        self._start_worker(
            "Đang tạo bản cập nhật, vui lòng chờ...",
            task,
            self._on_manual_build_success,
            lambda message: self._show_task_error("Tạo bản cập nhật", message),
        )

    def _on_manual_build_success(self, result: object) -> None:
        self.log(f"Hoàn thành: {result.package_path}")
        self._show_info(
            "Tạo bản cập nhật",
            f"Đã tạo bản cập nhật:\n{result.package_path}\n\nManifest:\n{result.manifest_path}",
        )

    def _show_task_error(self, title: str, message: str) -> None:
        self._show_error(title, message)

    def open_update_folder(self) -> None:
        from PySide6.QtCore import QUrl
        from PySide6.QtGui import QDesktopServices

        path = Path(self.update_path_edit.text().strip() or DEFAULT_UPDATE_PATH)
        self.update_path_edit.setText(str(path))
        if not self._ensure_update_folder_ready("Mở thư mục Update"):
            return
        QDesktopServices.openUrl(QUrl.fromLocalFile(str(path)))

    def _config(self) -> BuildUpdateConfig:
        notes = tuple(
            line.strip()
            for line in self.notes_edit.toPlainText().splitlines()
            if line.strip()
        )
        migrations: list[MigrationItem] = []
        if self.has_migration_checkbox.isChecked():
            for row in range(self.migration_table.rowCount()):
                version = self._cell(row, 0)
                file_text = self._cell(row, 1)
                description = self._cell(row, 2)
                python_item = self.migration_table.item(row, 3)
                use_python = (
                    python_item is not None
                    and python_item.checkState() == Qt.CheckState.Checked
                )
                migrations.append(
                    MigrationItem(
                        version=version,
                        description=description,
                        source_file=Path(file_text) if file_text and not use_python else None,
                        use_python_migration=use_python,
                    )
                )
        return BuildUpdateConfig(
            source_path=Path(self.source_edit.text().strip()),
            update_path=Path(self.update_path_edit.text().strip() or DEFAULT_UPDATE_PATH),
            new_version=self.new_version_edit.text().strip(),
            release_date=self.release_date_edit.text().strip(),
            required_app_restart=self.restart_checkbox.isChecked(),
            notes=notes,
            migrations=tuple(migrations),
            auto_update_source_version=self.auto_version_checkbox.isChecked(),
            database_changed=self.has_migration_checkbox.isChecked()
            or bool(self.old_db_edit.text().strip())
            or bool(self.new_db_edit.text().strip()),
            previous_release_version=self.previous_release_version,
            allow_rebuild_same_version=self.allow_rebuild_same_version,
            package_mode=self._package_mode(),
        )

    def _cell(self, row: int, column: int) -> str:
        item = self.migration_table.item(row, column)
        return item.text().strip() if item is not None else ""

    def clear_log(self) -> None:
        log_edit = getattr(self, "log_edit", None)
        if log_edit is not None:
            log_edit.clear()

    @Slot(str)
    def log(self, message: str) -> None:
        text = str(message)
        log_edit = getattr(self, "log_edit", None)
        if log_edit is None:
            print(text)
            return
        log_edit.append(text)
        QApplication.processEvents()

    @staticmethod
    def _format_diff(diff) -> str:
        lines: list[str] = []
        lines.append("Bảng mới:")
        lines.extend(f"- {table.name}" for table in diff.new_tables)
        lines.append("")
        lines.append("Cột mới:")
        lines.extend(f"- {column.table}.{column.name}" for column in diff.new_columns)
        lines.append("")
        lines.append("Index mới:")
        lines.extend(f"- {index.name} ({index.table})" for index in diff.new_indexes)
        lines.append("")
        lines.append("Dữ liệu mặc định mới:")
        lines.extend(f"- app_preferences.{item.key}" for item in diff.new_app_preferences)
        lines.append("")
        lines.append("Cần xử lý thủ công:")
        lines.extend(
            f"- {item.kind}: {item.name} - {item.detail}"
            for item in diff.dangerous_changes
        )
        return "\n".join(lines)
