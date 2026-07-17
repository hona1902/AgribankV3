from __future__ import annotations

from datetime import date
from pathlib import Path
import os
import platform
import subprocess

from PySide6.QtCore import QDate
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
from agribank_v3.features.credit.tovayvon.debt_reconciliation import (
    DEBT_RECONCILIATION_TITLE,
    DebtReconciliationError,
    DebtReconciliationRequest,
    create_debt_reconciliation,
    default_debt_reconciliation_output_path,
    detect_debt_columns,
)
from agribank_v3.features.credit.tovayvon.repository import (
    CreditGroupRepository,
    CreditGroupRepositoryError,
)
from agribank_v3.runtime_paths import application_root


class DebtReconciliationWindow(QDialog):
    """UI for creating dư nợ reconciliation workbook by credit group."""

    def __init__(
        self,
        parent: QWidget | None = None,
        database_path: Path | None = None,
    ) -> None:
        super().__init__(parent)
        self.setWindowTitle(DEBT_RECONCILIATION_TITLE)
        self.setModal(True)
        self.setMinimumSize(860, 620)
        self.resize(900, 680)

        self.database_path = self._resolve_database_path(parent, database_path)
        self.repository = CreditGroupRepository(self.database_path)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(18, 16, 18, 14)
        layout.setSpacing(12)

        title = QLabel(DEBT_RECONCILIATION_TITLE)
        title.setObjectName("SectionTitle")
        layout.addWidget(title)

        form = QGridLayout()
        form.setHorizontalSpacing(8)
        form.setVerticalSpacing(8)
        form.setColumnMinimumWidth(0, 190)
        form.setColumnStretch(1, 1)

        self.input_file_edit = QLineEdit()
        self.input_file_edit.setPlaceholderText("File Loan/IPCAS/SKCK có MaTo hoặc GRPNO")
        self.input_file_edit.setToolTip(
            "Chọn file sao kê dư nợ/Loan để đối chiếu theo mã tổ vay vốn."
        )
        form.addWidget(QLabel("File sao kê dư nợ"), 0, 0)
        form.addWidget(self.input_file_edit, 0, 1)
        choose_input_button = QPushButton("Chọn...")
        choose_input_button.setToolTip(self.input_file_edit.toolTip())
        choose_input_button.clicked.connect(self._choose_input_file)
        form.addWidget(choose_input_button, 0, 2)

        today = QDate.currentDate()
        self.reconciliation_date_edit = QDateEdit()
        self.reconciliation_date_edit.setCalendarPopup(True)
        self.reconciliation_date_edit.setDisplayFormat("dd/MM/yyyy")
        self.reconciliation_date_edit.setDate(today)
        form.addWidget(QLabel("Ngày đối chiếu"), 1, 0)
        form.addWidget(self.reconciliation_date_edit, 1, 1)

        self.all_groups_check = QCheckBox("Đối chiếu toàn bộ tổ")
        self.all_groups_check.setChecked(True)
        self.all_groups_check.toggled.connect(self._toggle_group_selector)
        form.addWidget(QLabel("Phạm vi đối chiếu"), 2, 0)
        form.addWidget(self.all_groups_check, 2, 1, 1, 2)

        self.group_combo = CheckableComboBox(placeholder="Chọn tổ vay vốn...")
        self._load_group_options()
        form.addWidget(QLabel("Tổ vay vốn"), 3, 0)
        form.addWidget(self.group_combo, 3, 1, 1, 2)

        self.output_file_edit = QLineEdit()
        self.output_file_edit.setText(
            str(default_debt_reconciliation_output_path(application_root() / "KetQua"))
        )
        form.addWidget(QLabel("File kết quả"), 4, 0)
        form.addWidget(self.output_file_edit, 4, 1)
        choose_output_button = QPushButton("Chọn...")
        choose_output_button.clicked.connect(self._choose_output_file)
        form.addWidget(choose_output_button, 4, 2)

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
        self.log_edit.setPlaceholderText("Log kiểm tra và tạo bảng đối chiếu dư nợ...")
        layout.addWidget(self.log_edit, stretch=1)

        button_row = QHBoxLayout()
        check_button = QPushButton("Kiểm tra dữ liệu")
        check_button.setObjectName("SecondaryButton")
        check_button.clicked.connect(self._check_data)
        button_row.addWidget(check_button)

        open_folder_button = QPushButton("Mở thư mục kết quả")
        open_folder_button.setObjectName("SecondaryButton")
        open_folder_button.clicked.connect(self._open_output_folder)
        button_row.addWidget(open_folder_button)

        button_row.addStretch()

        create_button = QPushButton("Tạo bảng đối chiếu")
        create_button.setObjectName("PrimaryButton")
        create_button.clicked.connect(self._create_reconciliation)
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

    def _choose_input_file(self) -> None:
        path, _ = QFileDialog.getOpenFileName(
            self,
            "Chọn file sao kê dư nợ/Loan",
            "",
            "Data files (*.xlsx *.xlsm *.xls *.csv);;All files (*.*)",
        )
        if path:
            self.input_file_edit.setText(path)

    def _choose_output_file(self) -> None:
        path, _ = QFileDialog.getSaveFileName(
            self,
            "Chọn nơi lưu bảng đối chiếu dư nợ",
            self.output_file_edit.text().strip(),
            "Excel files (*.xlsx)",
        )
        if path:
            output_path = Path(path)
            if output_path.suffix.lower() != ".xlsx":
                output_path = output_path.with_suffix(".xlsx")
            self.output_file_edit.setText(str(output_path))

    def _create_reconciliation(self) -> None:
        try:
            request = self._collect_request()
            if not self._check_data(show_message=False):
                return
        except DebtReconciliationError as exc:
            QMessageBox.warning(self, DEBT_RECONCILIATION_TITLE, str(exc))
            return

        self._append_reconciliation_mode(request)
        try:
            result = create_debt_reconciliation(request, self.repository)
        except (DebtReconciliationError, CreditGroupRepositoryError, OSError) as exc:
            self._append_log(f"Lỗi: {exc}")
            QMessageBox.warning(self, DEBT_RECONCILIATION_TITLE, str(exc))
            return

        self._append_log(
            f"Đã tạo bảng đối chiếu: {result.group_count} tổ, {result.detail_count} dòng chi tiết."
        )
        for warning in result.warnings:
            self._append_log(f"Cảnh báo: {warning}")
        self._append_log(f"File kết quả: {result.output_path}")
        QMessageBox.information(
            self,
            DEBT_RECONCILIATION_TITLE,
            "Đã tạo bảng đối chiếu dư nợ:\n" + str(result.output_path),
        )

    def _collect_request(self) -> DebtReconciliationRequest:
        input_file = Path(self.input_file_edit.text().strip())
        output_file = Path(self.output_file_edit.text().strip())
        if not input_file.is_file():
            raise DebtReconciliationError("Vui lòng chọn file sao kê dư nợ hợp lệ.")
        if not output_file.name:
            raise DebtReconciliationError("Vui lòng chọn nơi lưu file kết quả.")

        qdate = self.reconciliation_date_edit.date()
        if not qdate.isValid():
            raise DebtReconciliationError("Vui lòng nhập ngày đối chiếu.")
        reconciliation_date = date(qdate.year(), qdate.month(), qdate.day())
        return DebtReconciliationRequest(
            input_file=input_file,
            output_path=output_file.with_suffix(".xlsx"),
            reconciliation_date=reconciliation_date,
            selected_group_codes=self._selected_group_codes(),
        )

    def _selected_group_codes(self) -> tuple[str, ...]:
        if self.all_groups_check.isChecked():
            return ()
        selected = tuple(str(value) for value in self.group_combo.get_selected_values() if value)
        if not selected:
            raise DebtReconciliationError(
                "Vui lòng chọn ít nhất một tổ vay vốn hoặc tích Đối chiếu toàn bộ tổ."
            )
        return selected

    def _check_data(self, *, show_message: bool = True) -> bool:
        try:
            request = self._collect_request()
            detection = detect_debt_columns(request.input_file, self.repository)
        except DebtReconciliationError as exc:
            self._append_log(f"Kiểm tra dữ liệu lỗi: {exc}")
            if show_message:
                QMessageBox.warning(self, DEBT_RECONCILIATION_TITLE, str(exc))
            return False

        lines = [
            "Kiểm tra dữ liệu:",
            f"Sao kê dư nợ ({detection.sheet_name}):",
            "- Đã nhận diện: " + (", ".join(sorted(detection.field_to_header)) or "không"),
            "- Thiếu bắt buộc: " + (", ".join(detection.missing_required) or "không"),
            "- Thiếu tùy chọn: " + (", ".join(detection.missing_optional) or "không"),
            f"- Số dòng dữ liệu: {detection.row_count}",
            f"- Số dòng có MaTo: {detection.rows_with_group}",
            f"- Số dòng thiếu MaTo: {detection.rows_missing_group}",
            f"- Số MaTo nhận diện: {detection.group_count}",
            f"- MaTo có trong SQLite: {detection.known_group_count}",
            f"- MaTo chưa có trong SQLite: {detection.unknown_group_count}",
        ]
        message = "\n".join(lines)
        self._append_log(message)
        if detection.missing_required:
            if show_message:
                QMessageBox.warning(self, DEBT_RECONCILIATION_TITLE, message)
            return False
        if show_message:
            QMessageBox.information(
                self,
                DEBT_RECONCILIATION_TITLE,
                message + "\n\nCó thể tạo bảng đối chiếu.",
            )
        return True

    def _toggle_group_selector(self, all_groups: bool) -> None:
        self.group_combo.setEnabled(not all_groups)
        if all_groups:
            self.group_combo.select_all()

    def _open_output_folder(self) -> None:
        output_path = Path(self.output_file_edit.text().strip())
        folder = output_path.parent if output_path.name else application_root() / "KetQua"
        folder.mkdir(parents=True, exist_ok=True)
        try:
            _open_path(folder)
        except OSError as exc:
            QMessageBox.warning(self, DEBT_RECONCILIATION_TITLE, str(exc))

    def _append_reconciliation_mode(self, request: DebtReconciliationRequest) -> None:
        if request.selected_group_codes:
            self._append_log(
                f"Chế độ: Đối chiếu {len(request.selected_group_codes)} tổ được chọn."
            )
            self._append_log("Danh sách tổ: " + ", ".join(request.selected_group_codes))
        else:
            self._append_log("Chế độ: Đối chiếu toàn bộ tổ.")

    def clear_log(self) -> None:
        self.log_edit.clear()

    def _append_log(self, message: str) -> None:
        self.log_edit.append(message)


def _open_path(path: Path) -> None:
    system = platform.system().lower()
    if system == "windows" and hasattr(os, "startfile"):
        os.startfile(path)  # type: ignore[attr-defined]
    elif system == "darwin":
        subprocess.Popen(["open", str(path)])
    else:
        subprocess.Popen(["xdg-open", str(path)])
