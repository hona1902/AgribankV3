from __future__ import annotations

from datetime import date
from pathlib import Path

from PySide6.QtCore import QDate, QUrl
from PySide6.QtGui import QDesktopServices
from PySide6.QtWidgets import (
    QCheckBox,
    QDateEdit,
    QDialog,
    QFileDialog,
    QGridLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from agribank_v3.features.credit.tovayvon.checkable_combo_box import CheckableComboBox
from agribank_v3.features.credit.tovayvon.interest_report import (
    INTEREST_REPORT_TITLE,
    InterestReportError,
    InterestReportRequest,
    create_interest_report,
    default_interest_report_output_path,
    detect_skck_columns,
    detect_sktl_columns,
)
from agribank_v3.features.credit.tovayvon.repository import (
    CreditGroupRepository,
    CreditGroupRepositoryError,
)
from agribank_v3.runtime_paths import application_root


class InterestReportWindow(QDialog):
    """UI for creating Bảng kê thu lãi tổ vay vốn."""

    def __init__(
        self,
        parent: QWidget | None = None,
        database_path: Path | None = None,
    ) -> None:
        super().__init__(parent)
        self.setWindowTitle(INTEREST_REPORT_TITLE)
        self.setModal(True)
        self.setMinimumSize(860, 620)
        self.resize(900, 680)

        self.database_path = self._resolve_database_path(parent, database_path)
        self.repository = CreditGroupRepository(self.database_path)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(18, 16, 18, 14)
        layout.setSpacing(12)

        title = QLabel(INTEREST_REPORT_TITLE)
        title.setObjectName("SectionTitle")
        layout.addWidget(title)

        form = QGridLayout()
        form.setHorizontalSpacing(8)
        form.setVerticalSpacing(8)
        form.setColumnMinimumWidth(0, 190)
        form.setColumnStretch(1, 1)

        self.interest_file_edit = QLineEdit()
        self.interest_file_edit.setPlaceholderText("File lnlr13 sao kê thu lãi trong kỳ")
        self.interest_file_edit.setToolTip(
            "Chọn file lnlr13 sao kê thu lãi trong kỳ, dùng để lấy thông tin lãi đã thu."
        )
        form.addWidget(QLabel("File sao kê thu lãi trong kỳ (SKTL)"), 0, 0)
        form.addWidget(self.interest_file_edit, 0, 1)
        interest_button = QPushButton("Chọn SKTL")
        interest_button.setToolTip(self.interest_file_edit.toolTip())
        interest_button.clicked.connect(self._choose_interest_file)
        form.addWidget(interest_button, 0, 2)

        self.debt_file_edit = QLineEdit()
        self.debt_file_edit.setPlaceholderText("File lnlr13 sao kê cuối kỳ / dư nợ cuối kỳ")
        self.debt_file_edit.setToolTip(
            "Chọn file lnlr13 sao kê cuối kỳ, dùng để lấy dư nợ, nhóm nợ, lãi tồn và thông tin khoản vay cuối kỳ."
        )
        form.addWidget(QLabel("File sao kê cuối kỳ (SKCK)"), 1, 0)
        form.addWidget(self.debt_file_edit, 1, 1)
        debt_button = QPushButton("Chọn SKCK")
        debt_button.setToolTip(self.debt_file_edit.toolTip())
        debt_button.clicked.connect(self._choose_debt_file)
        form.addWidget(debt_button, 1, 2)

        today = QDate.currentDate()
        self.from_date_edit = QDateEdit()
        self.from_date_edit.setCalendarPopup(True)
        self.from_date_edit.setDisplayFormat("dd/MM/yyyy")
        self.from_date_edit.setDate(QDate(today.year(), max(1, today.month() - 2), 1))
        form.addWidget(QLabel("Từ ngày"), 2, 0)
        form.addWidget(self.from_date_edit, 2, 1)

        self.to_date_edit = QDateEdit()
        self.to_date_edit.setCalendarPopup(True)
        self.to_date_edit.setDisplayFormat("dd/MM/yyyy")
        self.to_date_edit.setDate(today)
        form.addWidget(QLabel("Đến ngày"), 3, 0)
        form.addWidget(self.to_date_edit, 3, 1)

        self.all_groups_check = QCheckBox("Tạo bảng kê cho toàn bộ các tổ")
        self.all_groups_check.setChecked(True)
        self.all_groups_check.toggled.connect(self._toggle_group_selector)
        form.addWidget(QLabel("Phạm vi tạo"), 4, 0)
        form.addWidget(self.all_groups_check, 4, 1, 1, 2)

        self.group_combo = CheckableComboBox(placeholder="Chọn tổ vay vốn...")
        self._load_group_options()
        form.addWidget(QLabel("Tổ vay vốn"), 5, 0)
        form.addWidget(self.group_combo, 5, 1, 1, 2)

        self.include_overdue_check = QCheckBox("Tính lãi quá hạn vào tổng lãi thu được")
        form.addWidget(QLabel("Tùy chọn"), 6, 0)
        form.addWidget(self.include_overdue_check, 6, 1, 1, 2)

        self.output_file_edit = QLineEdit()
        self.output_file_edit.setText(str(default_interest_report_output_path(application_root() / "KetQua")))
        form.addWidget(QLabel("File kết quả"), 7, 0)
        form.addWidget(self.output_file_edit, 7, 1)
        output_actions = QHBoxLayout()
        output_actions.setContentsMargins(0, 0, 0, 0)
        output_actions.setSpacing(6)
        output_button = QPushButton("Chọn...")
        output_button.clicked.connect(self._choose_output_file)
        output_actions.addWidget(output_button)
        form.addLayout(output_actions, 7, 2)

        layout.addLayout(form)

        log_header = QHBoxLayout()
        log_header.addWidget(QLabel("Nhật ký xử lý"))
        log_header.addStretch()
        clear_log_button = QPushButton("Xóa log")
        clear_log_button.setObjectName("SecondaryButton")
        clear_log_button.setToolTip("Xóa nội dung nhật ký xử lý hiện tại.")
        clear_log_button.clicked.connect(self.clear_log)
        log_header.addWidget(clear_log_button)
        layout.addLayout(log_header)

        self.log_edit = QTextEdit()
        self.log_edit.setReadOnly(True)
        self.log_edit.setPlaceholderText("Log tiến trình tạo bảng kê...")
        layout.addWidget(self.log_edit, stretch=1)

        button_row = QHBoxLayout()
        check_button = QPushButton("Kiểm tra dữ liệu")
        check_button.setObjectName("SecondaryButton")
        check_button.clicked.connect(self._check_data)
        button_row.addWidget(check_button)
        button_row.addStretch()
        open_output_folder_button = QPushButton("Mở thư mục kết quả")
        open_output_folder_button.setObjectName("SecondaryButton")
        open_output_folder_button.setToolTip("Mở thư mục chứa file kết quả.")
        open_output_folder_button.clicked.connect(self._open_output_folder)
        button_row.addWidget(open_output_folder_button)
        create_button = QPushButton("Tạo bảng kê")
        create_button.setObjectName("PrimaryButton")
        create_button.clicked.connect(self._create_report)
        button_row.addWidget(create_button)

        close_button = QPushButton("Đóng")
        close_button.setObjectName("SecondaryButton")
        close_button.clicked.connect(self.reject)
        button_row.addWidget(close_button)
        layout.addLayout(button_row)
        self._toggle_group_selector(True)

    @staticmethod
    def _resolve_database_path(parent: QWidget | None, database_path: Path | None) -> Path:
        if database_path is not None:
            return Path(database_path)
        parent_database = getattr(parent, "settings_database", None)
        parent_path = getattr(parent_database, "database_path", None)
        if parent_path is not None:
            return Path(parent_path)
        return application_root() / "data" / "DuLieuV3.db"

    def _load_group_options(self) -> None:
        self.group_combo.clear()
        try:
            groups = self.repository.list_groups()
        except CreditGroupRepositoryError as exc:
            self._append_log(f"Không đọc được danh sách tổ vay vốn: {exc}")
            return
        for group in groups:
            label_parts = [group.ma_to, group.ten_to or group.ten_tvv_day_du]
            if group.ten_to_truong:
                label_parts.append(group.ten_to_truong)
            if group.xa:
                label_parts.append(group.xa)
            label = " - ".join(part for part in label_parts if part)
            self.group_combo.add_check_item(label, group.ma_to)

    def _choose_interest_file(self) -> None:
        path, _ = QFileDialog.getOpenFileName(
            self,
            "Chọn file sao kê thu lãi trong kỳ (SKTL)",
            "",
            "Excel files (*.xlsx *.xlsm);;All files (*.*)",
        )
        if path:
            self.interest_file_edit.setText(path)

    def _choose_debt_file(self) -> None:
        path, _ = QFileDialog.getOpenFileName(
            self,
            "Chọn file sao kê cuối kỳ (SKCK)",
            "",
            "Excel files (*.xlsx *.xlsm);;All files (*.*)",
        )
        if path:
            self.debt_file_edit.setText(path)

    def _choose_output_file(self) -> None:
        path, _ = QFileDialog.getSaveFileName(
            self,
            "Chọn nơi lưu bảng kê",
            self.output_file_edit.text().strip(),
            "Excel files (*.xlsx)",
        )
        if path:
            output_path = Path(path)
            if output_path.suffix.lower() != ".xlsx":
                output_path = output_path.with_suffix(".xlsx")
            self.output_file_edit.setText(str(output_path))

    def _open_output_folder(self) -> None:
        output_text = self.output_file_edit.text().strip()
        if not output_text:
            QMessageBox.warning(
                self,
                INTEREST_REPORT_TITLE,
                "Vui lòng chọn file kết quả trước khi mở thư mục.",
            )
            return
        output_path = Path(output_text)
        folder = output_path.parent if output_path.suffix else output_path
        try:
            folder.mkdir(parents=True, exist_ok=True)
        except OSError as exc:
            QMessageBox.warning(
                self,
                INTEREST_REPORT_TITLE,
                f"Không thể tạo/mở thư mục kết quả: {exc}",
            )
            return
        if not QDesktopServices.openUrl(QUrl.fromLocalFile(str(folder))):
            QMessageBox.warning(
                self,
                INTEREST_REPORT_TITLE,
                f"Không thể mở thư mục kết quả:\n{folder}",
            )

    def _create_report(self) -> None:
        try:
            request = self._collect_request()
            if not self._check_data(show_message=False):
                return
            sktl = detect_sktl_columns(request.interest_file)
            skck = detect_skck_columns(request.debt_file)
            optional_missing = sktl.missing_optional + skck.missing_optional
            if optional_missing:
                answer = QMessageBox.question(
                    self,
                    INTEREST_REPORT_TITLE,
                    "Dữ liệu thiếu một số cột tùy chọn: "
                    + ", ".join(optional_missing)
                    + ".\nBạn có muốn tiếp tục không?",
                    QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                    QMessageBox.StandardButton.Yes,
                )
                if answer != QMessageBox.StandardButton.Yes:
                    return
        except InterestReportError as exc:
            QMessageBox.warning(self, INTEREST_REPORT_TITLE, str(exc))
            return

        self._append_log("Đang đọc dữ liệu SKTL/SKCK...")
        self._append_report_mode(request)
        try:
            result = create_interest_report(request, self.repository)
        except (InterestReportError, CreditGroupRepositoryError, OSError) as exc:
            self._append_log(f"Lỗi: {exc}")
            QMessageBox.warning(self, INTEREST_REPORT_TITLE, str(exc))
            return

        self._append_log(
            f"Đã tạo {result.group_count} bảng kê tổ, {result.detail_count} dòng chi tiết."
        )
        for message in result.info_messages:
            self._append_log(message)
        for warning in result.warnings:
            self._append_log(f"Cảnh báo: {warning}")
        self._append_log(f"File kết quả: {result.output_path}")
        QMessageBox.information(
            self,
            INTEREST_REPORT_TITLE,
            "Đã tạo bảng kê thu lãi tổ vay vốn:\n" + str(result.output_path),
        )

    def _collect_request(self) -> InterestReportRequest:
        interest_file = Path(self.interest_file_edit.text().strip())
        debt_file = Path(self.debt_file_edit.text().strip())
        output_file = Path(self.output_file_edit.text().strip())
        if not interest_file.is_file():
            raise InterestReportError("Vui lòng chọn file SKTL hợp lệ.")
        if not debt_file.is_file():
            raise InterestReportError("Vui lòng chọn file SKCK hợp lệ.")
        if not output_file.name:
            raise InterestReportError("Vui lòng chọn nơi lưu file kết quả.")

        from_qdate = self.from_date_edit.date()
        to_qdate = self.to_date_edit.date()
        if not from_qdate.isValid() or not to_qdate.isValid():
            raise InterestReportError("Vui lòng nhập kỳ thu lãi từ ngày và đến ngày.")
        from_date = date(from_qdate.year(), from_qdate.month(), from_qdate.day())
        to_date = date(to_qdate.year(), to_qdate.month(), to_qdate.day())
        if from_date > to_date:
            raise InterestReportError("Từ ngày không được lớn hơn đến ngày.")
        return InterestReportRequest(
            interest_file=interest_file,
            debt_file=debt_file,
            output_path=output_file.with_suffix(".xlsx"),
            from_date=from_date,
            to_date=to_date,
            selected_group_codes=self._selected_group_codes(),
            include_overdue_interest=self.include_overdue_check.isChecked(),
        )

    def _selected_group_codes(self) -> tuple[str, ...]:
        if self.all_groups_check.isChecked():
            return ()
        selected = tuple(str(value) for value in self.group_combo.get_selected_values() if value)
        if not selected:
            raise InterestReportError(
                "Vui lòng chọn ít nhất một tổ vay vốn hoặc tích Tạo bảng kê cho toàn bộ các tổ."
            )
        return selected

    def _check_data(self, *, show_message: bool = True) -> bool:
        try:
            request = self._collect_request()
            sktl = detect_sktl_columns(request.interest_file)
            skck = detect_skck_columns(request.debt_file)
        except InterestReportError as exc:
            self._append_log(f"Kiểm tra dữ liệu lỗi: {exc}")
            if show_message:
                QMessageBox.warning(self, INTEREST_REPORT_TITLE, str(exc))
            return False
        lines = [
            "Kiểm tra dữ liệu:",
            f"SKTL ({sktl.sheet_name}):",
            "- Đã nhận diện: " + (", ".join(sorted(sktl.field_to_header)) or "không"),
            "- Thiếu bắt buộc: " + (", ".join(sktl.missing_required) or "không"),
            "- Thiếu tùy chọn: " + (", ".join(sktl.missing_optional) or "không"),
            f"SKCK ({skck.sheet_name}):",
            "- Đã nhận diện: " + (", ".join(sorted(skck.field_to_header)) or "không"),
            "- Thiếu bắt buộc: " + (", ".join(skck.missing_required) or "không"),
            "- Thiếu tùy chọn: " + (", ".join(skck.missing_optional) or "không"),
        ]
        message = "\n".join(lines)
        self._append_log(message)
        if sktl.missing_required or skck.missing_required:
            if show_message:
                QMessageBox.warning(self, INTEREST_REPORT_TITLE, message)
            return False
        if show_message:
            QMessageBox.information(self, INTEREST_REPORT_TITLE, message + "\n\nCó thể tạo bảng kê.")
        return True

    def _append_log(self, message: str) -> None:
        self.log_edit.append(message)

    def clear_log(self) -> None:
        self.log_edit.clear()

    def _toggle_group_selector(self, all_groups: bool) -> None:
        self.group_combo.setEnabled(not all_groups)
        if all_groups:
            self.group_combo.select_all()

    def _append_report_mode(self, request: InterestReportRequest) -> None:
        if request.selected_group_codes:
            self._append_log(
                f"Chế độ: Tạo bảng kê cho {len(request.selected_group_codes)} tổ được chọn."
            )
            self._append_log("Danh sách tổ: " + ", ".join(request.selected_group_codes))
        else:
            self._append_log("Chế độ: Tạo bảng kê cho toàn bộ tổ.")
