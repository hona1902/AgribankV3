from __future__ import annotations

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QButtonGroup,
    QDialog,
    QDialogButtonBox,
    QFrame,
    QHBoxLayout,
    QLabel,
    QMessageBox,
    QPushButton,
    QRadioButton,
    QVBoxLayout,
    QWidget,
)

from agribank_v3.excel import CaseMode, ExcelConnectionError, ExcelContext, ExcelService


class CaseConversionDialog(QDialog):
    def __init__(
        self,
        excel_service: ExcelService,
        context: ExcelContext,
        parent: QWidget | None = None,
    ) -> None:
        super().__init__(parent)
        self.excel_service = excel_service
        self.setWindowTitle("Chuyển kiểu chữ")
        self.setModal(True)
        self.setMinimumWidth(510)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(22, 20, 22, 18)
        layout.setSpacing(14)

        title = QLabel("Chuyển kiểu chữ trong vùng đang chọn")
        title.setObjectName("PageTitle")
        layout.addWidget(title)

        context_frame = QFrame()
        context_frame.setObjectName("MetricCard")
        context_layout = QVBoxLayout(context_frame)
        context_layout.setContentsMargins(14, 10, 14, 10)
        if excel_service.is_system_worksheet(context):
            context_text = (
                f"Phiên bản: <b>{context.excel_name}</b> "
                "(workbook hệ thống, thông tin sheet đã ẩn)"
            )
        else:
            context_text = (
                f"Phiên bản: <b>{context.excel_name}</b> "
                f"(COM {context.excel_version})<br>"
                f"Workbook: <b>{context.workbook}</b><br>"
                f"Sheet: <b>{context.worksheet}</b><br>"
                f"Vùng chọn: <b>{context.selection}</b> "
                f"({context.cell_count:,} ô)"
            )
        context_layout.addWidget(QLabel(context_text))
        layout.addWidget(context_frame)

        instruction = QLabel(
            "Chọn kiểu chuyển đổi. Ô công thức, ô số và ô trống sẽ không bị thay đổi."
        )
        instruction.setObjectName("MutedText")
        instruction.setWordWrap(True)
        layout.addWidget(instruction)

        self.mode_group = QButtonGroup(self)
        options = (
            ("CHỮ HOA", CaseMode.UPPER),
            ("chữ thường", CaseMode.LOWER),
            ("Viết Hoa Tên", CaseMode.TITLE),
        )
        for index, (label, mode) in enumerate(options):
            radio = QRadioButton(label)
            radio.setProperty("caseMode", mode.value)
            radio.setMinimumHeight(30)
            self.mode_group.addButton(radio)
            layout.addWidget(radio)
            if index == 0:
                radio.setChecked(True)

        self.result_label = QLabel("")
        self.result_label.setWordWrap(True)
        self.result_label.setStyleSheet("color: #257047; font-weight: 600;")
        layout.addWidget(self.result_label)

        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Close)
        buttons.rejected.connect(self.reject)
        self.undo_button = QPushButton("Hoàn tác lần vừa rồi")
        self.undo_button.setObjectName("SecondaryButton")
        self.undo_button.setEnabled(excel_service.can_undo)
        self.undo_button.clicked.connect(self.undo)
        self.apply_button = QPushButton("Áp dụng")
        self.apply_button.setObjectName("PrimaryButton")
        self.apply_button.setDefault(True)
        self.apply_button.clicked.connect(self.apply_conversion)
        buttons.addButton(self.undo_button, QDialogButtonBox.ButtonRole.ActionRole)
        buttons.addButton(self.apply_button, QDialogButtonBox.ButtonRole.AcceptRole)
        layout.addWidget(buttons)

    def selected_mode(self) -> CaseMode:
        button = self.mode_group.checkedButton()
        return CaseMode(button.property("caseMode"))

    def apply_conversion(self) -> None:
        try:
            result = self.excel_service.convert_selection_case(self.selected_mode())
        except ExcelConnectionError as exc:
            QMessageBox.warning(self, "Không thể xử lý", str(exc))
            return

        self.result_label.setText(
            f"Đã thay đổi {result.changed_cells:,} ô; "
            f"bỏ qua {result.skipped_formulas:,} ô công thức và "
            f"{result.skipped_non_text:,} ô không phải văn bản."
        )
        self.undo_button.setEnabled(self.excel_service.can_undo)

    def undo(self) -> None:
        try:
            restored_areas = self.excel_service.undo_last_change()
        except ExcelConnectionError as exc:
            QMessageBox.warning(self, "Không thể hoàn tác", str(exc))
            return
        self.undo_button.setEnabled(False)
        self.result_label.setText(
            f"Đã hoàn tác dữ liệu trong {restored_areas:,} vùng chọn."
        )
