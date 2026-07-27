from __future__ import annotations

from datetime import date
from pathlib import Path

from PySide6.QtCore import QDate, QUrl
from PySide6.QtGui import QDesktopServices
from PySide6.QtWidgets import (
    QButtonGroup,
    QComboBox,
    QDateEdit,
    QDialog,
    QFileDialog,
    QGridLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QRadioButton,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from agribank_v3.features.credit.auto_interest.menu import (
    AUTO_INTEREST_TITLE,
    CREATE_INTEREST_FILE_TITLE,
    CREATE_REPORT_FILE_TITLE,
    REPORT_FOLDER_SETTINGS_TITLE,
)
from agribank_v3.features.credit.auto_interest.processor import (
    AutoInterestCreateRequest,
    AutoInterestError,
    AutoInterestReportRequest,
    COLLECT_ALL_INTEREST,
    COLLECTION_MODE_LABELS,
    NOT_DUE_AND_OVERDUE_INTEREST,
    NOT_DUE_INTEREST,
    OVERDUE_CENTER_INTEREST,
    create_auto_interest_file,
    create_auto_interest_report,
    validate_auto_interest_inputs,
)
from agribank_v3.features.credit.auto_interest.settings import (
    AutoInterestSettings,
    default_auto_interest_settings,
    load_auto_interest_settings,
    save_auto_interest_settings,
)
from agribank_v3.runtime_paths import application_root
from agribank_v3.settings import SettingsDatabaseError


AUTO_INTEREST_PLACEHOLDER_TITLE = AUTO_INTEREST_TITLE


class AutoInterestPlaceholderDialog(QDialog):
    """Backward-compatible placeholder for old routes."""

    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setWindowTitle(AUTO_INTEREST_TITLE)
        self.setModal(True)
        self.setMinimumWidth(560)
        layout = QVBoxLayout(self)
        layout.addWidget(QLabel("Thu lãi bán tự động đã được tách thành màn hình con."))
        close_button = QPushButton("Đóng")
        close_button.clicked.connect(self.reject)
        layout.addWidget(close_button)


class CreateAutoInterestFileWindow(QDialog):
    """UI for the VBA flow ThuLaiBanTuDong.frm -> Loading.frm."""

    def __init__(
        self,
        parent: QWidget | None = None,
        database_path: Path | None = None,
    ) -> None:
        super().__init__(parent)
        self.setWindowTitle(f"{CREATE_INTEREST_FILE_TITLE} - AgribankV3")
        self.setModal(True)
        self.setMinimumSize(860, 640)
        self.database_path = _resolve_database_path(parent, database_path)
        self.settings = load_auto_interest_settings(self.database_path)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(18, 16, 18, 14)
        layout.setSpacing(12)
        title = QLabel(CREATE_INTEREST_FILE_TITLE)
        title.setObjectName("SectionTitle")
        layout.addWidget(title)

        form = QGridLayout()
        form.setHorizontalSpacing(8)
        form.setVerticalSpacing(8)
        form.setColumnMinimumWidth(0, 180)
        form.setColumnStretch(1, 1)

        self.loan_file_edit = QLineEdit()
        self.loan_file_edit.setPlaceholderText("File Loan/lnlr23 - sao kê lãi dự kiến")
        form.addWidget(QLabel("File lãi dự kiến"), 0, 0)
        form.addWidget(self.loan_file_edit, 0, 1)
        loan_button = QPushButton("Chọn...")
        loan_button.clicked.connect(self._choose_loan_file)
        form.addWidget(loan_button, 0, 2)

        self.deposit_file_edit = QLineEdit()
        self.deposit_file_edit.setPlaceholderText("File sao kê tiền gửi MSIT/DPDA")
        form.addWidget(QLabel("File tiền gửi"), 1, 0)
        form.addWidget(self.deposit_file_edit, 1, 1)
        deposit_button = QPushButton("Chọn...")
        deposit_button.clicked.connect(self._choose_deposit_file)
        form.addWidget(deposit_button, 1, 2)

        self.collateral_file_edit = QLineEdit()
        self.collateral_file_edit.setPlaceholderText("File sao kê tài sản bảo đảm")
        form.addWidget(QLabel("File TSBĐ"), 2, 0)
        form.addWidget(self.collateral_file_edit, 2, 1)
        collateral_button = QPushButton("Chọn...")
        collateral_button.clicked.connect(self._choose_collateral_file)
        form.addWidget(collateral_button, 2, 2)

        self.deposit_type_combo = QComboBox()
        self.deposit_type_combo.addItem("MSIT - Msit81 Deposit List Report", "msit")
        self.deposit_type_combo.addItem("DPDA - Dpda08 Account List Report", "dpda")
        dpda_index = self.deposit_type_combo.findData("dpda")
        if dpda_index >= 0:
            self.deposit_type_combo.setCurrentIndex(dpda_index)
        form.addWidget(QLabel("Loại sao kê tiền gửi"), 3, 0)
        form.addWidget(self.deposit_type_combo, 3, 1, 1, 2)

        self.collection_mode_combo = QComboBox()
        for mode in (
            COLLECT_ALL_INTEREST,
            NOT_DUE_INTEREST,
            OVERDUE_CENTER_INTEREST,
            NOT_DUE_AND_OVERDUE_INTEREST,
        ):
            self.collection_mode_combo.addItem(COLLECTION_MODE_LABELS[mode], mode)
        form.addWidget(QLabel("Hình thức thu lãi"), 4, 0)
        form.addWidget(self.collection_mode_combo, 4, 1, 1, 2)

        self.collection_date_edit = QDateEdit()
        self.collection_date_edit.setCalendarPopup(True)
        self.collection_date_edit.setDisplayFormat("dd/MM/yyyy")
        self.collection_date_edit.setDate(_last_day_of_current_month())
        form.addWidget(QLabel("Ngày thu lãi"), 5, 0)
        form.addWidget(self.collection_date_edit, 5, 1, 1, 2)

        self.weekend_interest_group = QButtonGroup(self)
        self.include_weekend_interest_radio = QRadioButton("Tính lãi ngày nghỉ (thứ 7, chủ nhật)")
        self.include_weekend_interest_radio.setChecked(True)
        self.exclude_weekend_interest_radio = QRadioButton("Không tính lãi ngày nghỉ")
        self.weekend_interest_group.addButton(self.include_weekend_interest_radio)
        self.weekend_interest_group.addButton(self.exclude_weekend_interest_radio)
        weekend_options = QHBoxLayout()
        weekend_options.setContentsMargins(0, 0, 0, 0)
        weekend_options.setSpacing(18)
        weekend_options.addWidget(self.include_weekend_interest_radio)
        weekend_options.addWidget(self.exclude_weekend_interest_radio)
        weekend_options.addStretch()
        form.addWidget(QLabel("Ngày nghỉ"), 6, 0)
        form.addLayout(weekend_options, 6, 1, 1, 2)

        self.output_folder_edit = QLineEdit(str(self.settings.output_folder))
        form.addWidget(QLabel("Thư mục kết quả"), 7, 0)
        form.addWidget(self.output_folder_edit, 7, 1)
        output_button = QPushButton("Chọn...")
        output_button.clicked.connect(self._choose_output_folder)
        form.addWidget(output_button, 7, 2)

        self.report_option_group = QButtonGroup(self)
        self.create_report_radio = QRadioButton("Tạo báo cáo thu lãi bán tự động")
        self.create_report_radio.setChecked(True)
        self.skip_report_radio = QRadioButton("Không tạo báo cáo")
        self.report_option_group.addButton(self.create_report_radio)
        self.report_option_group.addButton(self.skip_report_radio)
        report_options = QHBoxLayout()
        report_options.setContentsMargins(0, 0, 0, 0)
        report_options.setSpacing(18)
        report_options.addWidget(self.create_report_radio)
        report_options.addWidget(self.skip_report_radio)
        report_options.addStretch()
        form.addWidget(QLabel("Báo cáo"), 8, 0)
        form.addLayout(report_options, 8, 1, 1, 2)
        layout.addLayout(form)
        self.collection_mode_combo.currentIndexChanged.connect(
            self._update_weekend_interest_options
        )
        self._update_weekend_interest_options()

        log_header = QHBoxLayout()
        log_header.addWidget(QLabel("Nhật ký xử lý"))
        log_header.addStretch()
        clear_button = QPushButton("Xóa log")
        clear_button.setObjectName("SecondaryButton")
        clear_button.setToolTip("Xóa nội dung nhật ký xử lý hiện tại.")
        clear_button.clicked.connect(self.clear_log)
        log_header.addWidget(clear_button)
        layout.addLayout(log_header)

        self.log_edit = QTextEdit()
        self.log_edit.setReadOnly(True)
        self.log_edit.setPlaceholderText("Log kiểm tra và tạo file thu lãi...")
        layout.addWidget(self.log_edit, stretch=1)

        button_row = QHBoxLayout()
        check_button = QPushButton("Kiểm tra dữ liệu")
        check_button.setObjectName("SecondaryButton")
        check_button.clicked.connect(self._check_data)
        button_row.addWidget(check_button)
        button_row.addStretch()
        open_button = QPushButton("Mở thư mục kết quả")
        open_button.setObjectName("SecondaryButton")
        open_button.clicked.connect(lambda: _open_folder(Path(self.output_folder_edit.text().strip()), self))
        button_row.addWidget(open_button)
        create_button = QPushButton("Tạo file thu lãi")
        create_button.setObjectName("PrimaryButton")
        create_button.clicked.connect(self._create_file)
        button_row.addWidget(create_button)
        close_button = QPushButton("Đóng")
        close_button.setObjectName("SecondaryButton")
        close_button.clicked.connect(self.reject)
        button_row.addWidget(close_button)
        layout.addLayout(button_row)

    def clear_log(self) -> None:
        if getattr(self, "log_edit", None) is not None:
            self.log_edit.clear()

    def _update_weekend_interest_options(self) -> None:
        enabled = str(self.collection_mode_combo.currentData()) == OVERDUE_CENTER_INTEREST
        self.include_weekend_interest_radio.setEnabled(enabled)
        self.exclude_weekend_interest_radio.setEnabled(enabled)

    def _choose_loan_file(self) -> None:
        path, _ = QFileDialog.getOpenFileName(
            self,
            "Chọn file Loan/lnlr23",
            str(self.settings.loan_folder or ""),
            "Data files (*.xlsx *.xlsm *.xls *.csv);;All files (*.*)",
        )
        if path:
            self.loan_file_edit.setText(path)

    def _choose_deposit_file(self) -> None:
        path, _ = QFileDialog.getOpenFileName(
            self,
            "Chọn file sao kê tiền gửi",
            str(self.settings.deposit_folder or ""),
            "Data files (*.xlsx *.xlsm *.xls *.csv);;All files (*.*)",
        )
        if path:
            self.deposit_file_edit.setText(path)

    def _choose_collateral_file(self) -> None:
        path, _ = QFileDialog.getOpenFileName(
            self,
            "Chọn file sao kê tài sản bảo đảm",
            str(self.settings.loan_folder or self.settings.deposit_folder or ""),
            "Data files (*.xlsx *.xlsm *.xls *.csv);;All files (*.*)",
        )
        if path:
            self.collateral_file_edit.setText(path)

    def _choose_output_folder(self) -> None:
        path = QFileDialog.getExistingDirectory(
            self,
            "Chọn thư mục lưu file thu lãi",
            self.output_folder_edit.text().strip(),
        )
        if path:
            self.output_folder_edit.setText(path)

    def _check_data(self, show_message: bool = True) -> bool:
        self._append_log("Kiểm tra dữ liệu:")
        try:
            collateral_file = _path_from_edit(self.collateral_file_edit)
            if collateral_file is None:
                raise AutoInterestError("Vui lòng chọn file sao kê tài sản bảo đảm.")
            warnings = validate_auto_interest_inputs(
                Path(self.loan_file_edit.text().strip()),
                Path(self.deposit_file_edit.text().strip()),
                self.deposit_type_combo.currentData(),
                collateral_file,
            )
        except AutoInterestError as exc:
            self._append_log(f"- Lỗi: {exc}")
            if show_message:
                QMessageBox.warning(self, CREATE_INTEREST_FILE_TITLE, str(exc))
            return False
        self._append_log("- Header bắt buộc: hợp lệ")
        for warning in warnings:
            self._append_log(f"- {warning}")
        if show_message:
            QMessageBox.information(self, CREATE_INTEREST_FILE_TITLE, "Dữ liệu hợp lệ.")
        return True

    def _create_file(self) -> None:
        if not self._check_data(show_message=False):
            return
        output_folder = Path(self.output_folder_edit.text().strip())
        report_result = None
        report_error: Exception | None = None
        try:
            request = AutoInterestCreateRequest(
                loan_file=Path(self.loan_file_edit.text().strip()),
                deposit_file=Path(self.deposit_file_edit.text().strip()),
                collection_date=_qdate_to_date(self.collection_date_edit.date()),
                deposit_statement_type=str(self.deposit_type_combo.currentData()),
                collection_mode=str(self.collection_mode_combo.currentData()),
                output_folder=output_folder,
                include_weekend_interest=self.include_weekend_interest_radio.isChecked(),
                collateral_file=_required_path_from_edit(
                    self.collateral_file_edit,
                    "file sao kê tài sản bảo đảm",
                ),
            )
            result = create_auto_interest_file(request)
            self.settings = AutoInterestSettings(
                report_folder=self.settings.report_folder,
                output_folder=output_folder,
                loan_folder=self.settings.loan_folder,
                deposit_folder=self.settings.deposit_folder,
                backup_folder=self.settings.backup_folder,
            )
            save_auto_interest_settings(self.settings, self.database_path)
        except (AutoInterestError, OSError, SettingsDatabaseError) as exc:
            self._append_log(f"- Lỗi tạo file: {exc}")
            QMessageBox.warning(self, CREATE_INTEREST_FILE_TITLE, str(exc))
            return
        if self.create_report_radio.isChecked():
            self._append_log("- Đang tạo báo cáo thu lãi bán tự động...")
            try:
                report_result = create_auto_interest_report(
                    AutoInterestReportRequest(
                        source_file=result.output_file,
                        settings=self.settings,
                        report_date=date.today(),
                        collection_mode=request.collection_mode,
                    )
                )
            except (AutoInterestError, OSError, SettingsDatabaseError) as exc:
                report_error = exc
        self._append_log(f"- Hình thức: {result.summary.get('mode_label', '')}")
        self._append_log(f"- Procedure VBA: {result.summary.get('vba_procedure', '')}")
        self._append_log(f"- Đã tạo {result.row_count} dòng.")
        self._append_log(f"- Đã bỏ qua {result.skipped_count} dòng không đủ điều kiện.")
        pledged_rows = int(result.summary.get("pledged_collateral_rows", 0) or 0)
        if pledged_rows:
            self._append_log(f"- Đã loại {pledged_rows} dòng do khách hàng có tài sản cầm cố 994003.")
        self._append_log(f"- Tổng lãi: {result.summary.get('total_interest', 0):,.0f}")
        for warning in result.warnings:
            self._append_log(f"- Cảnh báo: {warning}")
        self._append_log(f"- File kết quả: {result.output_file}")
        if report_result is not None:
            self._append_log(f"- Đã tạo báo cáo {report_result.row_count} dòng.")
            self._append_log(f"- File báo cáo: {report_result.output_file}")
        elif report_error is not None:
            self._append_log(f"- Lỗi tạo báo cáo: {report_error}")
            QMessageBox.warning(
                self,
                CREATE_INTEREST_FILE_TITLE,
                "File thu lãi đã tạo nhưng không tạo được báo cáo:\n"
                f"{report_error}",
            )
        else:
            self._append_log("- Không tạo báo cáo theo lựa chọn.")
        message = f"Đã tạo file thu lãi:\n{result.output_file}"
        if report_result is not None:
            message += f"\n\nĐã tạo báo cáo:\n{report_result.output_file}"
        QMessageBox.information(
            self,
            CREATE_INTEREST_FILE_TITLE,
            message,
        )

    def _append_log(self, text: str) -> None:
        if getattr(self, "log_edit", None) is not None:
            self.log_edit.append(str(text))
        else:
            print(text)


class AutoInterestFolderSettingsWindow(QDialog):
    """Settings window replacing VBA ChonThuMucBaoCao."""

    def __init__(
        self,
        parent: QWidget | None = None,
        database_path: Path | None = None,
    ) -> None:
        super().__init__(parent)
        self.setWindowTitle(f"{REPORT_FOLDER_SETTINGS_TITLE} thu lãi - AgribankV3")
        self.setModal(True)
        self.setMinimumWidth(1080)
        self.resize(1120, 235)
        self.database_path = _resolve_database_path(parent, database_path)
        self.settings = load_auto_interest_settings(self.database_path)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(18, 16, 18, 14)
        layout.setSpacing(12)
        title = QLabel(REPORT_FOLDER_SETTINGS_TITLE)
        title.setObjectName("SectionTitle")
        layout.addWidget(title)

        form = QGridLayout()
        form.setHorizontalSpacing(8)
        form.setVerticalSpacing(8)
        form.setColumnMinimumWidth(0, 190)
        form.setColumnMinimumWidth(1, 560)
        form.setColumnStretch(1, 1)
        form.setColumnMinimumWidth(2, 310)

        self.report_folder_edit = QLineEdit(str(self.settings.report_folder))
        self._add_folder_row(form, 0, "Thư mục lưu báo cáo", self.report_folder_edit)

        self.output_folder_edit = QLineEdit(str(self.settings.output_folder))
        self._add_folder_row(
            form,
            1,
            "Thư mục kết quả file ThuLaiBanTuDong",
            self.output_folder_edit,
        )
        layout.addLayout(form)

        button_row = QHBoxLayout()
        reset_button = QPushButton("Khôi phục mặc định")
        reset_button.setObjectName("SecondaryButton")
        reset_button.clicked.connect(self._restore_defaults)
        button_row.addWidget(reset_button)
        button_row.addStretch()
        save_button = QPushButton("Lưu cài đặt")
        save_button.setObjectName("PrimaryButton")
        save_button.clicked.connect(self._save)
        button_row.addWidget(save_button)
        close_button = QPushButton("Đóng")
        close_button.setObjectName("SecondaryButton")
        close_button.clicked.connect(self.reject)
        button_row.addWidget(close_button)
        layout.addLayout(button_row)

    def _add_folder_row(
        self,
        form: QGridLayout,
        row: int,
        label: str,
        edit: QLineEdit,
    ) -> None:
        form.addWidget(QLabel(label), row, 0)
        edit.setMinimumWidth(520)
        edit.setMaximumWidth(620)
        form.addWidget(edit, row, 1)
        controls = QHBoxLayout()
        controls.setContentsMargins(10, 0, 0, 0)
        controls.setSpacing(8)
        choose_button = QPushButton("Chọn...")
        choose_button.setFixedWidth(78)
        choose_button.clicked.connect(lambda: self._choose_folder(edit))
        controls.addWidget(choose_button)
        check_button = QPushButton("Kiểm tra")
        check_button.setObjectName("SecondaryButton")
        check_button.setFixedWidth(82)
        check_button.clicked.connect(lambda: self._check_folder(edit))
        controls.addWidget(check_button)
        create_button = QPushButton("Tạo nếu chưa có")
        create_button.setObjectName("SecondaryButton")
        create_button.setFixedWidth(126)
        create_button.clicked.connect(lambda: self._create_folder(edit))
        controls.addWidget(create_button)
        form.addLayout(controls, row, 2)

    def _choose_folder(self, edit: QLineEdit) -> None:
        path = QFileDialog.getExistingDirectory(self, "Chọn thư mục", edit.text().strip())
        if path:
            edit.setText(path)

    def _check_folder(self, edit: QLineEdit) -> None:
        path = _path_from_edit(edit)
        if path is None:
            QMessageBox.warning(self, REPORT_FOLDER_SETTINGS_TITLE, "Vui lòng nhập thư mục.")
            return
        if path.exists():
            QMessageBox.information(self, REPORT_FOLDER_SETTINGS_TITLE, f"Thư mục đã tồn tại:\n{path}")
            return
        QMessageBox.warning(self, REPORT_FOLDER_SETTINGS_TITLE, f"Thư mục chưa tồn tại:\n{path}")

    def _create_folder(self, edit: QLineEdit) -> None:
        path = _path_from_edit(edit)
        if path is None:
            QMessageBox.warning(self, REPORT_FOLDER_SETTINGS_TITLE, "Vui lòng nhập thư mục.")
            return
        try:
            path.mkdir(parents=True, exist_ok=True)
        except OSError as exc:
            QMessageBox.warning(self, REPORT_FOLDER_SETTINGS_TITLE, f"Không thể tạo thư mục: {exc}")
            return
        QMessageBox.information(self, REPORT_FOLDER_SETTINGS_TITLE, f"Đã tạo/kiểm tra thư mục:\n{path}")

    def _restore_defaults(self) -> None:
        defaults = default_auto_interest_settings()
        self.report_folder_edit.setText(str(defaults.report_folder))
        self.output_folder_edit.setText(str(defaults.output_folder))

    def _save(self) -> None:
        try:
            settings = AutoInterestSettings(
                report_folder=_required_path_from_edit(self.report_folder_edit, "thư mục lưu báo cáo"),
                output_folder=_required_path_from_edit(
                    self.output_folder_edit,
                    "thư mục kết quả file ThuLaiBanTuDong",
                ),
                loan_folder=self.settings.loan_folder,
                deposit_folder=self.settings.deposit_folder,
                backup_folder=self.settings.backup_folder,
            )
            settings.report_folder.mkdir(parents=True, exist_ok=True)
            settings.output_folder.mkdir(parents=True, exist_ok=True)
            save_auto_interest_settings(settings, self.database_path)
        except (OSError, SettingsDatabaseError) as exc:
            QMessageBox.warning(self, REPORT_FOLDER_SETTINGS_TITLE, str(exc))
            return
        QMessageBox.information(self, REPORT_FOLDER_SETTINGS_TITLE, "Đã lưu cài đặt.")
        self.accept()


class CreateAutoInterestReportWindow(QDialog):
    """UI for VBA TaoBaoCaoThuBTD."""

    def __init__(
        self,
        parent: QWidget | None = None,
        database_path: Path | None = None,
    ) -> None:
        super().__init__(parent)
        self.setWindowTitle(f"{CREATE_REPORT_FILE_TITLE} - AgribankV3")
        self.setModal(True)
        self.setMinimumSize(820, 540)
        self.database_path = _resolve_database_path(parent, database_path)
        self.settings = load_auto_interest_settings(self.database_path)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(18, 16, 18, 14)
        layout.setSpacing(12)
        title = QLabel(CREATE_REPORT_FILE_TITLE)
        title.setObjectName("SectionTitle")
        layout.addWidget(title)

        form = QGridLayout()
        form.setHorizontalSpacing(8)
        form.setVerticalSpacing(8)
        form.setColumnMinimumWidth(0, 170)
        form.setColumnStretch(1, 1)

        self.source_file_edit = QLineEdit()
        self.source_file_edit.setPlaceholderText("File thu lãi đã tạo, có sheet SaoKeTrichLai")
        form.addWidget(QLabel("File thu lãi"), 0, 0)
        form.addWidget(self.source_file_edit, 0, 1)
        source_button = QPushButton("Chọn...")
        source_button.clicked.connect(self._choose_source_file)
        form.addWidget(source_button, 0, 2)

        self.report_date_edit = QDateEdit()
        self.report_date_edit.setCalendarPopup(True)
        self.report_date_edit.setDisplayFormat("dd/MM/yyyy")
        self.report_date_edit.setDate(QDate.currentDate())
        form.addWidget(QLabel("Ngày thu nợ"), 1, 0)
        form.addWidget(self.report_date_edit, 1, 1, 1, 2)

        self.report_folder_edit = QLineEdit(str(self.settings.report_folder))
        form.addWidget(QLabel("Thư mục báo cáo"), 2, 0)
        form.addWidget(self.report_folder_edit, 2, 1)
        folder_button = QPushButton("Chọn...")
        folder_button.clicked.connect(self._choose_report_folder)
        form.addWidget(folder_button, 2, 2)
        layout.addLayout(form)

        log_header = QHBoxLayout()
        log_header.addWidget(QLabel("Nhật ký xử lý"))
        log_header.addStretch()
        clear_button = QPushButton("Xóa log")
        clear_button.setObjectName("SecondaryButton")
        clear_button.clicked.connect(self.clear_log)
        log_header.addWidget(clear_button)
        layout.addLayout(log_header)

        self.log_edit = QTextEdit()
        self.log_edit.setReadOnly(True)
        self.log_edit.setPlaceholderText("Log tạo file báo cáo thu nợ bán tự động...")
        layout.addWidget(self.log_edit, stretch=1)

        button_row = QHBoxLayout()
        button_row.addStretch()
        open_button = QPushButton("Mở thư mục báo cáo")
        open_button.setObjectName("SecondaryButton")
        open_button.clicked.connect(lambda: _open_folder(Path(self.report_folder_edit.text().strip()), self))
        button_row.addWidget(open_button)
        create_button = QPushButton("Tạo file báo cáo")
        create_button.setObjectName("PrimaryButton")
        create_button.clicked.connect(self._create_report)
        button_row.addWidget(create_button)
        close_button = QPushButton("Đóng")
        close_button.setObjectName("SecondaryButton")
        close_button.clicked.connect(self.reject)
        button_row.addWidget(close_button)
        layout.addLayout(button_row)

    def clear_log(self) -> None:
        if getattr(self, "log_edit", None) is not None:
            self.log_edit.clear()

    def _choose_source_file(self) -> None:
        path, _ = QFileDialog.getOpenFileName(
            self,
            "Chọn file thu lãi bán tự động",
            "",
            "Excel files (*.xlsx *.xlsm);;All files (*.*)",
        )
        if path:
            self.source_file_edit.setText(path)

    def _choose_report_folder(self) -> None:
        path = QFileDialog.getExistingDirectory(
            self,
            "Chọn thư mục lưu báo cáo",
            self.report_folder_edit.text().strip(),
        )
        if path:
            self.report_folder_edit.setText(path)

    def _create_report(self) -> None:
        try:
            settings = AutoInterestSettings(
                report_folder=Path(self.report_folder_edit.text().strip()),
                output_folder=self.settings.output_folder,
                loan_folder=self.settings.loan_folder,
                deposit_folder=self.settings.deposit_folder,
                backup_folder=self.settings.backup_folder,
            )
            save_auto_interest_settings(settings, self.database_path)
            result = create_auto_interest_report(
                AutoInterestReportRequest(
                    source_file=Path(self.source_file_edit.text().strip()),
                    settings=settings,
                    report_date=_qdate_to_date(self.report_date_edit.date()),
                )
            )
        except (AutoInterestError, OSError, SettingsDatabaseError) as exc:
            self._append_log(f"- Lỗi: {exc}")
            QMessageBox.warning(self, CREATE_REPORT_FILE_TITLE, str(exc))
            return
        self._append_log(f"- Đã tạo {result.row_count} dòng báo cáo.")
        self._append_log(f"- File báo cáo: {result.output_file}")
        QMessageBox.information(
            self,
            CREATE_REPORT_FILE_TITLE,
            f"Đã tạo file báo cáo:\n{result.output_file}",
        )

    def _append_log(self, text: str) -> None:
        if getattr(self, "log_edit", None) is not None:
            self.log_edit.append(str(text))
        else:
            print(text)


def _resolve_database_path(parent: QWidget | None, database_path: Path | None) -> Path:
    if database_path is not None:
        return Path(database_path)
    parent_database = getattr(parent, "settings_database", None)
    parent_path = getattr(parent_database, "database_path", None)
    if parent_path is not None:
        return Path(parent_path)
    return application_root() / "data" / "DuLieuV3.db"


def _qdate_to_date(value: QDate) -> date:
    return date(value.year(), value.month(), value.day())


def _path_from_edit(edit: QLineEdit) -> Path | None:
    text = edit.text().strip()
    if not text:
        return None
    return Path(text)


def _required_path_from_edit(edit: QLineEdit, label: str) -> Path:
    path = _path_from_edit(edit)
    if path is None:
        raise OSError(f"Vui lòng nhập {label}.")
    return path


def _last_day_of_current_month() -> QDate:
    today = QDate.currentDate()
    return QDate(today.year(), today.month(), today.daysInMonth())


def _open_folder(folder: Path, parent: QWidget) -> None:
    if not str(folder):
        QMessageBox.warning(parent, AUTO_INTEREST_TITLE, "Vui lòng chọn thư mục trước.")
        return
    try:
        folder.mkdir(parents=True, exist_ok=True)
    except OSError as exc:
        QMessageBox.warning(parent, AUTO_INTEREST_TITLE, f"Không thể tạo/mở thư mục: {exc}")
        return
    if not QDesktopServices.openUrl(QUrl.fromLocalFile(str(folder))):
        QMessageBox.warning(parent, AUTO_INTEREST_TITLE, f"Không thể mở thư mục:\n{folder}")
