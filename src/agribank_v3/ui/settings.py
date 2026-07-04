from __future__ import annotations

from pathlib import Path

from PySide6.QtCore import Qt, Signal
from PySide6.QtWidgets import (
    QFileDialog,
    QCheckBox,
    QFormLayout,
    QFrame,
    QGridLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QRadioButton,
    QScrollArea,
    QTabWidget,
    QVBoxLayout,
    QWidget,
)

from agribank_v3.settings import (
    AddinMode,
    AppSettingsDatabase,
    BranchProfile,
    SettingsDatabaseError,
)


FIELD_DEFINITIONS = (
    ("branch_code", "Mã chi nhánh"),
    ("transaction_office_code", "Mã PGD"),
    ("branch_name", "Tên chi nhánh"),
    ("reporting_branch_name", "Chi nhánh báo cáo"),
    ("department_name", "Phòng ban / PGD"),
    ("address", "Địa chỉ"),
    ("tax_code", "Mã số thuế"),
    ("phone", "Điện thoại"),
    ("fax", "Fax"),
    ("parent_branch_name", "Chi nhánh cấp trên"),
    ("parent_branch_code", "Mã CN cấp trên"),
    ("report_location", "Địa điểm báo cáo"),
    ("report_preparer", "Người lập báo cáo"),
)


class SettingsWidget(QWidget):
    connect_excel_requested = Signal()
    show_excel_requested = Signal()
    addin_mode_changed = Signal(str)
    addin_enabled_changed = Signal(str, bool)

    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setObjectName("SettingsPage")
        self.database = AppSettingsDatabase()
        self.fields: dict[str, QLineEdit] = {}
        self.addin_mode = self.database.load_addin_mode()
        self.addin_checkboxes: dict[str, QCheckBox] = {}
        addin_names = tuple(path.name for path in self._addin_files())
        self.addin_states = self.database.load_addin_states(addin_names)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(28, 24, 28, 28)
        layout.setSpacing(16)

        title = QLabel("Cài đặt")
        title.setObjectName("PageTitle")
        subtitle = QLabel(
            "Quản lý kết nối Excel, thông tin chi nhánh và dữ liệu dùng chung "
            "của ứng dụng. Dữ liệu cài đặt không được lưu trên sheet Excel."
        )
        subtitle.setObjectName("MutedText")
        subtitle.setWordWrap(True)
        layout.addWidget(title)
        layout.addWidget(subtitle)

        self.tabs = QTabWidget()
        self.tabs.addTab(self._build_excel_tab(), "Kết nối Excel")
        self.tabs.addTab(self._build_branch_tab(), "Thông tin chi nhánh")
        self.tabs.addTab(self._build_database_tab(), "Cơ sở dữ liệu & sao lưu")
        layout.addWidget(self.tabs, stretch=1)
        self.load_profile()
        self.refresh_database_status()

    def _build_excel_tab(self) -> QWidget:
        page, layout = self._scroll_tab()
        group = QGroupBox("Kết nối phiên Excel đang làm việc")
        group_layout = QVBoxLayout(group)
        info = QLabel(
            "Ứng dụng chỉ làm việc với workbook người dùng đang mở. Các sheet "
            "hệ thống nội bộ không được đưa vào giao diện. Khi kết nối, app tự "
            "nạp các tệp .xla/.xlam trong thư mục tools/addins."
        )
        info.setWordWrap(True)
        info.setObjectName("MutedText")
        self.excel_status = QLabel("Excel chưa kết nối")
        self.excel_status.setObjectName("SectionTitle")
        self.addin_status = QLabel(
            f"Thư mục add-in: {self._addin_directory()} • XLSTART: "
            f"{self._xlstart_directory()}"
        )
        self.addin_status.setObjectName("MutedText")
        self.addin_status.setWordWrap(True)
        self.addin_status.setTextInteractionFlags(
            Qt.TextInteractionFlag.TextSelectableByMouse
        )
        mode_group = QGroupBox("Phạm vi sử dụng add-in")
        mode_layout = QVBoxLayout(mode_group)
        self.permanent_addin_radio = QRadioButton(
            "Cài add-in thường trực cho Excel"
        )
        permanent_note = QLabel(
            "Excel tự nạp add-in kể cả khi AgribankV3 không chạy."
        )
        permanent_note.setObjectName("MutedText")
        self.session_addin_radio = QRadioButton(
            "Chỉ dùng trong phiên AgribankV3"
        )
        session_note = QLabel(
            "Add-in được nạp khi kết nối và được gỡ khi thoát AgribankV3."
        )
        session_note.setObjectName("MutedText")
        self.permanent_addin_radio.setChecked(
            self.addin_mode is AddinMode.PERMANENT
        )
        self.session_addin_radio.setChecked(
            self.addin_mode is AddinMode.SESSION
        )
        self.permanent_addin_radio.toggled.connect(
            self._save_addin_mode_from_controls
        )
        self.session_addin_radio.toggled.connect(
            self._save_addin_mode_from_controls
        )
        mode_layout.addWidget(self.permanent_addin_radio)
        mode_layout.addWidget(permanent_note)
        mode_layout.addWidget(self.session_addin_radio)
        mode_layout.addWidget(session_note)

        addin_control_group = QGroupBox("Cài đặt hoặc gỡ từng add-in")
        addin_control_layout = QVBoxLayout(addin_control_group)
        addin_control_note = QLabel(
            "Tích để cài/nạp; bỏ tích để gỡ add-in khỏi Excel và XLSTART."
        )
        addin_control_note.setObjectName("MutedText")
        addin_control_layout.addWidget(addin_control_note)
        if self.addin_states:
            for file_name, enabled in self.addin_states.items():
                checkbox = QCheckBox(file_name)
                checkbox.setChecked(enabled)
                checkbox.toggled.connect(
                    lambda checked, name=file_name: self._save_addin_enabled(
                        name, checked
                    )
                )
                self.addin_checkboxes[file_name] = checkbox
                addin_control_layout.addWidget(checkbox)
            addin_buttons = QHBoxLayout()
            install_all_button = QPushButton("Cài tất cả")
            install_all_button.clicked.connect(
                lambda checked=False: self._set_all_addins_enabled(True)
            )
            remove_all_button = QPushButton("Gỡ tất cả")
            remove_all_button.clicked.connect(
                lambda checked=False: self._set_all_addins_enabled(False)
            )
            addin_buttons.addWidget(install_all_button)
            addin_buttons.addWidget(remove_all_button)
            addin_buttons.addStretch()
            addin_control_layout.addLayout(addin_buttons)
        else:
            empty_addins = QLabel(
                "Chưa có tệp .xla hoặc .xlam trong thư mục tools/addins."
            )
            empty_addins.setObjectName("MutedText")
            addin_control_layout.addWidget(empty_addins)
        actions = QHBoxLayout()
        connect_button = QPushButton("Kết nối Excel")
        connect_button.setObjectName("PrimaryButton")
        connect_button.clicked.connect(
            lambda checked=False: self.connect_excel_requested.emit()
        )
        show_button = QPushButton("Hiện cửa sổ Excel")
        show_button.clicked.connect(
            lambda checked=False: self.show_excel_requested.emit()
        )
        open_addin_folder_button = QPushButton("Mở thư mục add-in")
        open_addin_folder_button.clicked.connect(self.open_addin_folder)
        actions.addWidget(connect_button)
        actions.addWidget(show_button)
        actions.addWidget(open_addin_folder_button)
        actions.addStretch()
        group_layout.addWidget(self.excel_status)
        group_layout.addWidget(self.addin_status)
        group_layout.addWidget(info)
        group_layout.addWidget(mode_group)
        group_layout.addWidget(addin_control_group)
        group_layout.addLayout(actions)
        layout.addWidget(group)
        layout.addStretch()
        return page

    def _build_branch_tab(self) -> QWidget:
        page, layout = self._scroll_tab()
        group = QGroupBox("Thông tin dùng chung")
        form_grid = QGridLayout(group)
        form_grid.setHorizontalSpacing(14)
        form_grid.setVerticalSpacing(12)
        form_grid.setColumnMinimumWidth(0, 130)
        form_grid.setColumnMinimumWidth(2, 145)
        form_grid.setColumnStretch(1, 1)
        form_grid.setColumnStretch(3, 1)
        for index, (field_name, label) in enumerate(FIELD_DEFINITIONS):
            editor = QLineEdit()
            editor.setMinimumWidth(250)
            editor.setMaxLength(255)
            editor.setClearButtonEnabled(True)
            if field_name in {"branch_code", "transaction_office_code", "tax_code",
                              "phone", "fax", "parent_branch_code"}:
                editor.setMaxLength(50)
            self.fields[field_name] = editor
            column_group = index % 2
            row = index // 2
            label_widget = QLabel(label)
            label_widget.setAlignment(
                Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter
            )
            label_column = column_group * 2
            form_grid.addWidget(label_widget, row, label_column)
            form_grid.addWidget(editor, row, label_column + 1)

        note = QLabel(
            "Ứng dụng luôn sử dụng bộ thông tin mới nhất. Mỗi lần lưu, database "
            "giữ thêm một bản lịch sử nội bộ để có thể đối chiếu hoặc phục hồi "
            "khi cần."
        )
        note.setObjectName("MutedText")
        note.setWordWrap(True)
        actions = QHBoxLayout()
        self.save_button = QPushButton("Lưu thông tin")
        self.save_button.setObjectName("PrimaryButton")
        self.save_button.clicked.connect(self.save_profile)
        reload_button = QPushButton("Tải lại")
        reload_button.clicked.connect(self.load_profile)
        self.saved_label = QLabel()
        self.saved_label.setObjectName("MutedText")
        actions.addWidget(self.save_button)
        actions.addWidget(reload_button)
        actions.addWidget(self.saved_label)
        actions.addStretch()
        layout.addWidget(group)
        layout.addWidget(note)
        layout.addLayout(actions)
        layout.addStretch()
        return page

    def _build_database_tab(self) -> QWidget:
        page, layout = self._scroll_tab()
        status_group = QGroupBox("Trạng thái dữ liệu")
        status_layout = QFormLayout(status_group)
        self.database_path_label = QLabel()
        self.database_path_label.setTextInteractionFlags(
            Qt.TextInteractionFlag.TextSelectableByMouse
        )
        self.database_integrity_label = QLabel()
        self.quiz_database_path_label = QLabel()
        self.quiz_database_path_label.setTextInteractionFlags(
            Qt.TextInteractionFlag.TextSelectableByMouse
        )
        self.quiz_database_integrity_label = QLabel()
        self.database_size_label = QLabel()
        self.database_revision_label = QLabel()
        status_layout.addRow("Database cài đặt", self.database_path_label)
        status_layout.addRow(
            "Toàn vẹn database cài đặt",
            self.database_integrity_label,
        )
        status_layout.addRow(
            "Database trắc nghiệm",
            self.quiz_database_path_label,
        )
        status_layout.addRow(
            "Toàn vẹn database trắc nghiệm",
            self.quiz_database_integrity_label,
        )
        status_layout.addRow("Tổng dung lượng", self.database_size_label)
        status_layout.addRow("Số lần cập nhật thông tin CN", self.database_revision_label)

        actions_group = QGroupBox("Sao lưu và phục hồi")
        actions_layout = QVBoxLayout(actions_group)
        backup_location = QFormLayout()
        self.backup_directory_label = QLabel(
            str(self.database.backup_directory)
        )
        self.backup_directory_label.setTextInteractionFlags(
            Qt.TextInteractionFlag.TextSelectableByMouse
        )
        self.backup_file_label = QLabel(
            "AgribankV3-[ngày giờ]-[mã].zip"
        )
        self.backup_file_label.setTextInteractionFlags(
            Qt.TextInteractionFlag.TextSelectableByMouse
        )
        backup_location.addRow(
            "Thư mục lưu bản sao",
            self.backup_directory_label,
        )
        backup_location.addRow(
            "Tên tệp sao lưu",
            self.backup_file_label,
        )
        buttons = QHBoxLayout()
        verify_button = QPushButton("Kiểm tra database")
        verify_button.clicked.connect(
            lambda checked=False: self.refresh_database_status()
        )
        backup_button = QPushButton("Sao lưu ngay")
        backup_button.setObjectName("PrimaryButton")
        backup_button.clicked.connect(self.create_backup)
        restore_button = QPushButton("Phục hồi từ tệp…")
        restore_button.clicked.connect(self.restore_backup)
        open_folder_button = QPushButton("Mở thư mục sao lưu")
        open_folder_button.clicked.connect(self.open_backup_folder)
        buttons.addWidget(verify_button)
        buttons.addWidget(backup_button)
        buttons.addWidget(restore_button)
        buttons.addWidget(open_folder_button)
        buttons.addStretch()
        actions_layout.addLayout(backup_location)
        actions_layout.addLayout(buttons)
        self.backup_result_label = QLabel()
        self.backup_result_label.setObjectName("MutedText")
        actions_layout.addWidget(self.backup_result_label)

        layout.addWidget(status_group)
        layout.addWidget(actions_group)
        layout.addStretch()
        return page

    @staticmethod
    def _scroll_tab() -> tuple[QScrollArea, QVBoxLayout]:
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.Shape.NoFrame)
        body = QWidget()
        body.setObjectName("SettingsTabBody")
        layout = QVBoxLayout(body)
        layout.setContentsMargins(22, 20, 22, 22)
        layout.setSpacing(16)
        scroll.setWidget(body)
        return scroll, layout

    def load_profile(self) -> None:
        try:
            profile = self.database.load_branch_profile()
        except SettingsDatabaseError as exc:
            self._show_error(str(exc))
            return
        for field_name, editor in self.fields.items():
            editor.setText(str(getattr(profile, field_name)))
        self.saved_label.setText(
            (f"Cập nhật gần nhất: {profile.updated_at}" if profile.updated_at
             else "Chưa lưu thông tin")
        )

    def save_profile(self) -> None:
        profile = BranchProfile(
            **{
                field_name: editor.text()
                for field_name, editor in self.fields.items()
            }
        )
        try:
            saved = self.database.save_branch_profile(profile)
        except SettingsDatabaseError as exc:
            self._show_error(str(exc))
            return
        self.saved_label.setText(
            f"Đã lưu • {saved.updated_at}"
        )
        self.refresh_database_status(show_message=False)

    def refresh_database_status(self, show_message: bool = True) -> None:
        try:
            status = self.database.managed_status()
        except SettingsDatabaseError as exc:
            self._show_error(str(exc))
            return
        self.database_path_label.setText(str(status.settings.path))
        self.database_integrity_label.setText(
            "Tốt (OK)"
            if status.settings.integrity.casefold() == "ok"
            else status.settings.integrity
        )
        self.quiz_database_path_label.setText(str(status.quiz_path))
        self.quiz_database_integrity_label.setText(
            "Tốt (OK)"
            if status.quiz_integrity.casefold() == "ok"
            else status.quiz_integrity
        )
        self.database_size_label.setText(
            self._format_size(status.total_size_bytes)
        )
        self.database_revision_label.setText(
            str(status.settings.branch_revision)
            + (
                f" • {status.settings.last_updated_at}"
                if status.settings.last_updated_at
                else " • chưa có dữ liệu"
            )
        )
        if show_message:
            self.backup_result_label.setText("Đã kiểm tra database thành công.")

    def create_backup(self) -> None:
        try:
            path = self.database.create_backup()
        except SettingsDatabaseError as exc:
            self._show_error(str(exc))
            return
        self.backup_file_label.setText(path.name)
        self.backup_result_label.setText(
            f"Đã sao lưu thành công: {path}"
        )

    def restore_backup(self) -> None:
        selected, _ = QFileDialog.getOpenFileName(
            self,
            "Chọn bản sao lưu AgribankV3",
            str(self.database.backup_directory),
            "Gói sao lưu AgribankV3 (*.zip);;"
            "Bản sao database cũ (*.sqlite3 *.db);;Tất cả tệp (*)",
        )
        if not selected:
            return
        answer = QMessageBox.question(
            self,
            "Xác nhận phục hồi",
            "Các database trong gói sẽ thay thế dữ liệu hiện tại. Ứng dụng sẽ "
            "tự tạo một gói sao lưu an toàn trước khi thực hiện. Tiếp tục?",
        )
        if answer != QMessageBox.StandardButton.Yes:
            return
        try:
            safety_backup = self.database.restore_backup(Path(selected))
        except SettingsDatabaseError as exc:
            self._show_error(str(exc))
            return
        self.load_profile()
        self.refresh_database_status(show_message=False)
        self.backup_result_label.setText(
            f"Đã phục hồi. Bản sao trước phục hồi: {safety_backup}"
        )
        self.backup_file_label.setText(safety_backup.name)

    def open_backup_folder(self) -> None:
        from PySide6.QtCore import QUrl
        from PySide6.QtGui import QDesktopServices

        QDesktopServices.openUrl(
            QUrl.fromLocalFile(str(self.database.backup_directory))
        )

    def open_addin_folder(self) -> None:
        from PySide6.QtCore import QUrl
        from PySide6.QtGui import QDesktopServices

        directory = self._addin_directory()
        directory.mkdir(parents=True, exist_ok=True)
        QDesktopServices.openUrl(QUrl.fromLocalFile(str(directory)))

    def set_excel_status(self, text: str) -> None:
        self.excel_status.setText(text)

    def set_addin_status(self, text: str) -> None:
        self.addin_status.setText(text)

    def _save_addin_mode_from_controls(self, checked: bool) -> None:
        if not checked:
            return
        requested = (
            AddinMode.PERMANENT
            if self.permanent_addin_radio.isChecked()
            else AddinMode.SESSION
        )
        previous = self.addin_mode
        try:
            self.addin_mode = self.database.save_addin_mode(requested)
        except SettingsDatabaseError as exc:
            self.permanent_addin_radio.blockSignals(True)
            self.session_addin_radio.blockSignals(True)
            self.permanent_addin_radio.setChecked(
                previous is AddinMode.PERMANENT
            )
            self.session_addin_radio.setChecked(previous is AddinMode.SESSION)
            self.permanent_addin_radio.blockSignals(False)
            self.session_addin_radio.blockSignals(False)
            self._show_error(str(exc))
            return
        self.addin_mode_changed.emit(self.addin_mode.value)

    def _save_addin_enabled(self, file_name: str, enabled: bool) -> None:
        previous = self.addin_states.get(file_name, True)
        try:
            saved = self.database.save_addin_enabled(file_name, enabled)
        except SettingsDatabaseError as exc:
            checkbox = self.addin_checkboxes[file_name]
            checkbox.blockSignals(True)
            checkbox.setChecked(previous)
            checkbox.blockSignals(False)
            self._show_error(str(exc))
            return
        self.addin_states[file_name] = saved
        self.addin_enabled_changed.emit(file_name, saved)

    def _set_all_addins_enabled(self, enabled: bool) -> None:
        for checkbox in self.addin_checkboxes.values():
            checkbox.setChecked(enabled)

    @staticmethod
    def _addin_directory() -> Path:
        from agribank_v3.excel.service import ExcelService

        return ExcelService.tool_addin_directory()

    @classmethod
    def _addin_files(cls) -> tuple[Path, ...]:
        directory = cls._addin_directory()
        try:
            return tuple(
                sorted(
                    (
                        path
                        for path in directory.iterdir()
                        if path.is_file()
                        and not path.name.startswith("~$")
                        and path.suffix.casefold() in {".xla", ".xlam"}
                    ),
                    key=lambda path: path.name.casefold(),
                )
            )
        except OSError:
            return ()

    @staticmethod
    def _xlstart_directory() -> Path:
        from agribank_v3.excel.service import ExcelService

        return ExcelService.excel_xlstart_directory()

    def show_tab_for_feature(self, title: str) -> None:
        index = {
            "Kết nối Excel": 0,
            "Thông tin chi nhánh": 1,
            "Cơ sở dữ liệu": 2,
            "Sao lưu dữ liệu": 2,
        }.get(title, 0)
        self.tabs.setCurrentIndex(index)

    def _show_error(self, message: str) -> None:
        QMessageBox.warning(self, "Cài đặt", message)

    @staticmethod
    def _format_size(size: int) -> str:
        value = float(size)
        for unit in ("B", "KB", "MB", "GB"):
            if value < 1024 or unit == "GB":
                return f"{value:.1f} {unit}"
            value /= 1024
        return f"{size} B"
