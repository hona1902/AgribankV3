from __future__ import annotations

from pathlib import Path
import os
import platform
import subprocess

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QCheckBox,
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
from agribank_v3.features.credit.tovayvon.payment_request import (
    PAYMENT_REQUEST_TITLE,
    PaymentReportData,
    PaymentRequestError,
    analyze_payment_rows,
    default_payment_output_folder,
    default_payment_template_path,
    export_payment_requests,
    load_payment_report_data,
    open_payment_template_for_edit,
)
from agribank_v3.features.credit.tovayvon.repository import CreditGroupRepository
from agribank_v3.runtime_paths import application_root


class PaymentRequestWindow(QDialog):
    """UI for creating Word payment requests from TongHopTheoTo."""

    def __init__(
        self,
        parent: QWidget | None = None,
        database_path: Path | None = None,
    ) -> None:
        super().__init__(parent)
        self.setWindowTitle(PAYMENT_REQUEST_TITLE)
        self.setModal(True)
        self.setMinimumSize(880, 620)
        self.resize(920, 680)

        self.database_path = self._resolve_database_path(parent, database_path)
        self.repository = CreditGroupRepository(self.database_path)
        self.report_data: PaymentReportData | None = None

        layout = QVBoxLayout(self)
        layout.setContentsMargins(18, 16, 18, 14)
        layout.setSpacing(12)

        title = QLabel(PAYMENT_REQUEST_TITLE)
        title.setObjectName("SectionTitle")
        layout.addWidget(title)

        form = QGridLayout()
        form.setHorizontalSpacing(8)
        form.setVerticalSpacing(8)
        form.setColumnMinimumWidth(0, 180)
        form.setColumnStretch(1, 1)

        self.report_file_edit = QLineEdit()
        self.report_file_edit.setPlaceholderText("File bảng kê có sheet TongHopTheoTo")
        form.addWidget(QLabel("File bảng kê"), 0, 0)
        form.addWidget(self.report_file_edit, 0, 1)
        choose_report_button = QPushButton("Chọn...")
        choose_report_button.clicked.connect(self._choose_report_file)
        form.addWidget(choose_report_button, 0, 2)

        self.search_edit = QLineEdit()
        self.search_edit.setPlaceholderText("Tìm theo mã tổ, tên tổ hoặc tổ trưởng")
        self.search_edit.textChanged.connect(self._refresh_group_combo)
        form.addWidget(QLabel("Tìm tổ"), 1, 0)
        form.addWidget(self.search_edit, 1, 1, 1, 2)

        self.export_all_check = QCheckBox("Xuất tất cả tổ trong sheet TongHopTheoTo")
        self.export_all_check.setChecked(True)
        self.export_all_check.toggled.connect(self._toggle_group_selector)
        form.addWidget(QLabel("Phạm vi xuất"), 2, 0)
        form.addWidget(self.export_all_check, 2, 1, 1, 2)

        self.group_combo = CheckableComboBox(placeholder="Chọn tổ vay vốn...")
        form.addWidget(QLabel("Tổ vay vốn"), 3, 0)
        form.addWidget(self.group_combo, 3, 1, 1, 2)

        self.template_file_edit = QLineEdit()
        self.template_file_edit.setText(str(default_payment_template_path()))
        form.addWidget(QLabel("Mẫu Word"), 4, 0)
        form.addWidget(self.template_file_edit, 4, 1)
        choose_template_button = QPushButton("Chọn mẫu khác")
        choose_template_button.clicked.connect(self._choose_template_file)
        form.addWidget(choose_template_button, 4, 2)

        self.output_folder_edit = QLineEdit()
        self.output_folder_edit.setText(str(default_payment_output_folder()))
        form.addWidget(QLabel("Thư mục kết quả"), 5, 0)
        form.addWidget(self.output_folder_edit, 5, 1)
        choose_output_button = QPushButton("Chọn...")
        choose_output_button.clicked.connect(self._choose_output_folder)
        form.addWidget(choose_output_button, 5, 2)

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
        self.log_edit.setPlaceholderText("Log tạo đề nghị thanh toán...")
        layout.addWidget(self.log_edit, stretch=1)

        button_row = QHBoxLayout()
        check_button = QPushButton("Kiểm tra dữ liệu")
        check_button.setObjectName("SecondaryButton")
        check_button.clicked.connect(self._check_data)
        button_row.addWidget(check_button)

        edit_template_button = QPushButton("Mở mẫu để chỉnh sửa")
        edit_template_button.setObjectName("SecondaryButton")
        edit_template_button.clicked.connect(self._open_template)
        button_row.addWidget(edit_template_button)

        open_folder_button = QPushButton("Mở thư mục kết quả")
        open_folder_button.setObjectName("SecondaryButton")
        open_folder_button.clicked.connect(self._open_output_folder)
        button_row.addWidget(open_folder_button)

        button_row.addStretch()

        export_button = QPushButton("Xuất đề nghị thanh toán")
        export_button.setObjectName("PrimaryButton")
        export_button.clicked.connect(self._export)
        button_row.addWidget(export_button)

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

    def _choose_report_file(self) -> None:
        path, _ = QFileDialog.getOpenFileName(
            self,
            "Chọn file bảng kê thu lãi tổ vay vốn",
            "",
            "Excel files (*.xlsx *.xlsm);;All files (*.*)",
        )
        if path:
            self.report_file_edit.setText(path)
            self.report_data = None
            self._check_data(show_message=False)

    def _choose_template_file(self) -> None:
        path, _ = QFileDialog.getOpenFileName(
            self,
            "Chọn mẫu DeNghiThanhToan.docx",
            self.template_file_edit.text().strip(),
            "Word files (*.docx);;All files (*.*)",
        )
        if path:
            self.template_file_edit.setText(path)

    def _choose_output_folder(self) -> None:
        path = QFileDialog.getExistingDirectory(
            self,
            "Chọn thư mục kết quả",
            self.output_folder_edit.text().strip(),
        )
        if path:
            self.output_folder_edit.setText(path)

    def _check_data(self, show_message: bool = True) -> bool:
        try:
            report_path = Path(self.report_file_edit.text().strip())
            self.report_data = load_payment_report_data(report_path)
        except PaymentRequestError as exc:
            self._append_log(f"Lỗi: {exc}")
            if show_message:
                QMessageBox.warning(self, PAYMENT_REQUEST_TITLE, str(exc))
            return False

        self._refresh_group_combo()
        self._append_log("Kiểm tra dữ liệu:")
        self._append_log(f"- File bảng kê: {report_path}")
        self._append_log(f"- Sheet: TongHopTheoTo")
        scoped_rows = self._rows_for_current_scope()
        eligibility = analyze_payment_rows(scoped_rows)
        self._append_log(f"- Tổng số tổ: {eligibility.total}")
        self._append_log(f"- Đủ điều kiện chi: {len(eligibility.eligible)}")
        self._append_log(f"- Không đủ điều kiện chi: {len(eligibility.ineligible)}")
        if eligibility.ineligible:
            skipped = ", ".join(
                f"{row.ma_to} - {row.ten_to or row.ten_to_truong}"
                for row in eligibility.ineligible
            )
            self._append_log(f"- Không đủ điều kiện: {skipped}")
        if self.report_data.period_from or self.report_data.period_to:
            self._append_log(
                f"- Kỳ thu lãi: {self.report_data.period_from} - {self.report_data.period_to}"
            )
        for warning in self.report_data.warnings:
            self._append_log(f"Cảnh báo: {warning}")
        if show_message:
            QMessageBox.information(
                self,
                PAYMENT_REQUEST_TITLE,
                "Kiểm tra xong dữ liệu TongHopTheoTo.\n"
                f"Tổng số tổ: {eligibility.total}\n"
                f"Đủ điều kiện chi: {len(eligibility.eligible)}\n"
                f"Không đủ điều kiện chi: {len(eligibility.ineligible)}",
            )
        return True

    def _export(self) -> None:
        if self.report_data is None and not self._check_data(show_message=False):
            return
        try:
            selected_group_codes = self._selected_group_codes()
            result = export_payment_requests(
                report_path=Path(self.report_file_edit.text().strip()),
                template_path=Path(self.template_file_edit.text().strip()),
                output_folder=Path(self.output_folder_edit.text().strip()),
                repository=self.repository,
                selected_group_codes=selected_group_codes,
                export_all=self.export_all_check.isChecked(),
            )
        except (PaymentRequestError, OSError) as exc:
            self._append_log(f"Lỗi: {exc}")
            QMessageBox.warning(self, PAYMENT_REQUEST_TITLE, str(exc))
            return

        self._append_export_mode()
        self._append_log(f"Đã xuất {len(result.output_paths)} file đề nghị thanh toán.")
        for message in result.logs:
            self._append_log(message)
        for path in result.output_paths:
            self._append_log(f"Đã tạo: {path.name}")
        for warning in result.warnings:
            self._append_log(f"Cảnh báo: {warning}")
        QMessageBox.information(
            self,
            PAYMENT_REQUEST_TITLE,
            f"Đã xuất {len(result.output_paths)} file đề nghị thanh toán.",
        )

    def _open_template(self) -> None:
        try:
            open_payment_template_for_edit(Path(self.template_file_edit.text().strip()))
        except (PaymentRequestError, OSError) as exc:
            QMessageBox.warning(self, PAYMENT_REQUEST_TITLE, str(exc))

    def _open_output_folder(self) -> None:
        folder = Path(self.output_folder_edit.text().strip())
        folder.mkdir(parents=True, exist_ok=True)
        try:
            _open_path(folder)
        except OSError as exc:
            QMessageBox.warning(self, PAYMENT_REQUEST_TITLE, str(exc))

    def _toggle_group_selector(self, export_all: bool) -> None:
        self.group_combo.setEnabled(not export_all)
        self.search_edit.setEnabled(not export_all)
        if export_all:
            self.group_combo.select_all()

    def _refresh_group_combo(self) -> None:
        selected = self.group_combo.get_selected_values()
        self.group_combo.clear()
        if self.report_data is None:
            return
        keyword = self.search_edit.text().strip().casefold()
        for row in self.report_data.rows:
            label = f"{row.ma_to} - {row.ten_to or row.ten_to_truong}"
            if row.ten_to_truong and row.ten_to_truong not in label:
                label = f"{label} - {row.ten_to_truong}"
            haystack = f"{row.ma_to} {row.ten_to} {row.ten_to_truong}".casefold()
            if not keyword or keyword in haystack:
                self.group_combo.add_check_item(label, row.ma_to)
        self.group_combo.set_checked_data(selected)

    def clear_log(self) -> None:
        self.log_edit.clear()

    def _append_log(self, message: str) -> None:
        self.log_edit.append(message)

    def _selected_group_codes(self) -> tuple[str, ...]:
        if self.export_all_check.isChecked():
            return ()
        selected = tuple(str(value) for value in self.group_combo.get_selected_values() if value)
        if not selected:
            raise PaymentRequestError(
                "Vui lòng chọn ít nhất một tổ vay vốn hoặc tích Xuất tất cả các tổ trong sheet TongHopTheoTo."
            )
        return selected

    def _rows_for_current_scope(self) -> tuple:
        if self.report_data is None:
            return ()
        if self.export_all_check.isChecked():
            return self.report_data.rows
        selected = {
            str(value)
            for value in self.group_combo.get_selected_values()
            if value
        }
        if not selected:
            return self.report_data.rows
        return tuple(row for row in self.report_data.rows if row.ma_to in selected)

    def _append_export_mode(self) -> None:
        selected = tuple(str(value) for value in self.group_combo.get_selected_values() if value)
        if self.export_all_check.isChecked():
            self._append_log("Chế độ: Xuất đề nghị thanh toán cho toàn bộ tổ trong TongHopTheoTo.")
        else:
            self._append_log(f"Chế độ: Xuất đề nghị thanh toán cho {len(selected)} tổ được chọn.")


def _open_path(path: Path) -> None:
    system = platform.system().lower()
    if system == "windows" and hasattr(os, "startfile"):
        os.startfile(path)  # type: ignore[attr-defined]
    elif system == "darwin":
        subprocess.Popen(["open", str(path)])
    else:
        subprocess.Popen(["xdg-open", str(path)])
