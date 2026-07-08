from __future__ import annotations

from pathlib import Path

from PySide6.QtWidgets import (
    QDialog,
    QDialogButtonBox,
    QFileDialog,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QRadioButton,
    QVBoxLayout,
    QWidget,
)

from agribank_v3.settings import BranchProfile
from agribank_v3.settlement.models import SettlementOptions, SettlementSpec
from agribank_v3.ui.dialogs.settlement_period import (
    load_output_prefix,
    save_output_prefix,
)


class SimpleSourceSettlementDialog(QDialog):
    """Compact one-source settlement dialog used by accounting fixed templates."""

    def __init__(
        self,
        spec: SettlementSpec,
        profile: BranchProfile,
        parent: QWidget | None = None,
    ) -> None:
        super().__init__(parent)
        self.spec = spec
        self.profile = profile
        self.source_path: Path | None = None

        self.setWindowTitle(f"Tạo Mẫu biểu Quyết toán {spec.report_code}/QT")
        self.setModal(True)
        self.setMinimumWidth(690)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(14, 12, 14, 14)
        layout.setSpacing(12)

        source_label = QLabel(
            "Tên File nguồn dùng xử lý để tạo ra Mẫu biểu "
            f"{spec.report_code}QT là: {self._source_hint_text()}"
        )
        source_label.setStyleSheet("color: #0000ff; font-weight: 700;")
        source_label.setWordWrap(True)
        layout.addWidget(source_label)

        source_row = QHBoxLayout()
        self.source_edit = QLineEdit()
        self.source_edit.setReadOnly(True)
        source_row.addWidget(self.source_edit, 1)
        choose_button = QPushButton("Chọn File")
        choose_button.clicked.connect(self.choose_source_file)
        source_row.addWidget(choose_button)
        layout.addLayout(source_row)

        layout.addWidget(self._output_prefix_group())

        output_label = QLabel(
            f"Tên File quyết toán Mẫu {spec.report_code}/QT sẽ được tạo ra:"
        )
        output_label.setStyleSheet("color: #0000ff; font-weight: 700;")
        layout.addWidget(output_label)
        self.output_edit = QLineEdit()
        self.output_edit.setReadOnly(True)
        layout.addWidget(self.output_edit)

        buttons = QDialogButtonBox()
        create_button = QPushButton("Tạo Mẫu biểu")
        create_button.setObjectName("PrimaryButton")
        buttons.addButton(create_button, QDialogButtonBox.ButtonRole.AcceptRole)
        cancel_button = QPushButton("Cancel")
        buttons.addButton(cancel_button, QDialogButtonBox.ButtonRole.RejectRole)
        create_button.clicked.connect(self.accept)
        cancel_button.clicked.connect(self.reject)
        layout.addWidget(buttons)

        self._apply_saved_output_prefix()

    def _output_prefix_group(self) -> QGroupBox:
        group = QGroupBox("Loại kỳ báo cáo")
        layout = QHBoxLayout(group)
        self.qt_prefix_radio = QRadioButton("Quyết toán năm (QT)")
        self.bn_prefix_radio = QRadioButton("Bán niên (BN)")
        self.qt_prefix_radio.toggled.connect(self._sync_output_prefix)
        self.bn_prefix_radio.toggled.connect(self._sync_output_prefix)
        layout.addWidget(self.qt_prefix_radio)
        layout.addWidget(self.bn_prefix_radio)
        return group

    def choose_source_file(self) -> None:
        initial = str(self.source_path.parent) if self.source_path else ""
        file_name, _ = QFileDialog.getOpenFileName(
            self,
            f"Chọn file nguồn Mẫu {self.spec.report_code}/QT",
            initial,
            self._source_file_filter(),
        )
        if not file_name:
            return
        self.source_path = Path(file_name)
        self.source_edit.setText(str(self.source_path))
        self._refresh_output_path()

    def accept(self) -> None:
        if self.source_path is None:
            QMessageBox.warning(
                self,
                "Chưa chọn file nguồn",
                f"Hãy chọn file nguồn trước khi tạo Mẫu biểu {self.spec.report_code}/QT.",
            )
            return
        super().accept()

    def output_path(self) -> Path | None:
        if self.source_path is None:
            return None
        return self.source_path.with_name(
            f"{self.profile.branch_code.strip()}{self.output_prefix()}{self.output_report_code()}.xlsx"
        )

    def output_prefix(self) -> str:
        return "BN" if self.bn_prefix_radio.isChecked() else "QT"

    def output_report_code(self) -> str:
        special_codes = {
            "accounting.07a": "07A",
            "accounting.09a": "9a",
            "accounting.09b": "9b",
            "accounting.09c": "9c",
        }
        return special_codes.get(self.spec.key, self.spec.report_code)

    def options(self) -> SettlementOptions:
        return SettlementOptions(output_prefix=self.output_prefix())

    def _refresh_output_path(self) -> None:
        output_path = self.output_path()
        self.output_edit.setText(str(output_path) if output_path else "")

    def _apply_saved_output_prefix(self) -> None:
        if load_output_prefix() == "BN":
            self.bn_prefix_radio.setChecked(True)
        else:
            self.qt_prefix_radio.setChecked(True)

    def _sync_output_prefix(self) -> None:
        save_output_prefix(self.output_prefix())
        self._refresh_output_path()

    def _source_hint_text(self) -> str:
        return self.spec.source_hint.replace(
            "{MaCN}",
            self.profile.branch_code.strip() or "MaCN",
        )

    def _source_file_filter(self) -> str:
        return (
            "File Excel nguồn (*.xls *.xlsx *.xlsm);;"
            "CSV nguồn quyết toán (*.csv);;"
            "Tất cả file (*.*)"
        )
