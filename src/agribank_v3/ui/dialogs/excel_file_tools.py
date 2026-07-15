from __future__ import annotations

from pathlib import Path

from PySide6.QtCore import QSize, Qt
from PySide6.QtWidgets import (
    QCheckBox,
    QDialog,
    QFileDialog,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QMessageBox,
    QPushButton,
    QRadioButton,
    QVBoxLayout,
    QWidget,
)

from agribank_v3.excel_tools import list_workbook_sheet_names


class CsvToExcelDialog(QDialog):
    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.mode: str | None = None
        self.source_path: Path | None = None
        self.source_paths: tuple[Path, ...] = ()
        self.source_folder: Path | None = None
        self.output_directory: Path | None = None
        self.single_output_path: Path | None = None

        self.setWindowTitle("Chuyển CSV sang Excel")
        self.setModal(True)
        self.setMinimumWidth(760)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(14, 12, 14, 14)
        layout.setSpacing(8)

        guide = QLabel("Chọn chế độ chuyển CSV sang Excel.")
        guide.setStyleSheet("color: #0000ff; font-weight: 700;")
        layout.addWidget(guide)

        mode_row = QHBoxLayout()
        one_button = QPushButton("Chuyển 01 File")
        one_button.clicked.connect(self.choose_one_file)
        many_button = QPushButton("Chuyển nhiều File")
        many_button.clicked.connect(self.choose_many_files)
        folder_button = QPushButton("Chuyển 01 Folder")
        folder_button.clicked.connect(self.choose_folder)
        for button in (one_button, many_button, folder_button):
            button.setObjectName("SecondaryButton")
            mode_row.addWidget(button)
        mode_row.addStretch()
        layout.addLayout(mode_row)

        format_group = QGroupBox("Chọn định dạng file được xuất ra")
        format_row = QHBoxLayout(format_group)
        self.xls_radio = QRadioButton("XLS (Excel 2003 - 2007)")
        self.xlsx_radio = QRadioButton("XLSX (Excel Workbooks)")
        self.xlsx_radio.setChecked(True)
        self.xls_radio.toggled.connect(self._format_changed)
        self.xlsx_radio.toggled.connect(self._format_changed)
        format_row.addWidget(self.xls_radio)
        format_row.addWidget(self.xlsx_radio)
        format_row.addStretch()
        layout.addWidget(format_group)

        self.summary_label = QLabel("Chưa chọn dữ liệu CSV.")
        self.summary_label.setWordWrap(True)
        layout.addWidget(self.summary_label)

        self.source_list = QListWidget()
        self.source_list.setFixedHeight(120)
        layout.addWidget(self.source_list)

        output_row = QHBoxLayout()
        self.output_edit = QLineEdit()
        self.output_edit.setReadOnly(True)
        output_button = QPushButton("Chọn nơi lưu...")
        output_button.clicked.connect(self.choose_output_location)
        output_row.addWidget(self.output_edit, 1)
        output_row.addWidget(output_button)
        layout.addLayout(output_row)

        button_row = QHBoxLayout()
        convert_button = QPushButton("Chuyển sang Excel")
        convert_button.setObjectName("PrimaryButton")
        convert_button.clicked.connect(self.accept)
        cancel_button = QPushButton("Cancel")
        cancel_button.clicked.connect(self.reject)
        button_row.addStretch()
        button_row.addWidget(convert_button)
        button_row.addWidget(cancel_button)
        layout.addLayout(button_row)

    def choose_one_file(self) -> None:
        file_name, _ = QFileDialog.getOpenFileName(
            self,
            "Chọn file CSV",
            str(self.source_path.parent if self.source_path else Path.home()),
            "CSV (*.csv);;Tất cả file (*.*)",
        )
        if not file_name:
            return
        self.mode = "single"
        self.source_path = Path(file_name)
        self.source_paths = (self.source_path,)
        self.source_folder = None
        self.single_output_path = self.source_path.with_suffix(self.output_suffix())
        self.output_directory = None
        self._refresh_summary()

    def choose_many_files(self) -> None:
        file_names, _ = QFileDialog.getOpenFileNames(
            self,
            "Chọn các file CSV",
            str(self.source_paths[-1].parent if self.source_paths else Path.home()),
            "CSV (*.csv);;Tất cả file (*.*)",
        )
        if not file_names:
            return
        self.mode = "multiple"
        self.source_paths = tuple(Path(name) for name in file_names)
        self.source_path = None
        self.source_folder = None
        self.output_directory = self.source_paths[0].parent
        self.single_output_path = None
        self._refresh_summary()

    def choose_folder(self) -> None:
        folder = QFileDialog.getExistingDirectory(
            self,
            "Chọn folder chứa CSV",
            str(self.source_folder or Path.home()),
        )
        if not folder:
            return
        self.mode = "folder"
        self.source_folder = Path(folder)
        self.source_paths = tuple(sorted(self.source_folder.glob("*.csv")))
        self.source_path = None
        self.output_directory = self.source_folder
        self.single_output_path = None
        self._refresh_summary()

    def choose_output_location(self) -> None:
        if self.mode == "single":
            initial = str(self.output_path() or Path.home() / f"converted{self.output_suffix()}")
            file_name, _ = QFileDialog.getSaveFileName(
                self,
                "Chọn file Excel kết quả",
                initial,
                f"Excel (*{self.output_suffix()})",
            )
            if not file_name:
                return
            path = Path(file_name)
            if path.suffix.casefold() != self.output_suffix():
                path = path.with_suffix(self.output_suffix())
            self.single_output_path = path
        else:
            folder = QFileDialog.getExistingDirectory(
                self,
                "Chọn thư mục lưu file Excel",
                str(self.output_directory or Path.home()),
            )
            if not folder:
                return
            self.output_directory = Path(folder)
        self._refresh_summary()

    def output_path(self) -> Path | None:
        if self.mode != "single":
            return None
        return self.single_output_path

    def output_suffix(self) -> str:
        return ".xls" if self.xls_radio.isChecked() else ".xlsx"

    def output_format(self) -> str:
        return self.output_suffix().lstrip(".")

    def batch_outputs(self) -> tuple[tuple[Path, Path], ...]:
        if self.mode == "single":
            output = self.output_path()
            return ((self.source_paths[0], output),) if output is not None and self.source_paths else ()
        if self.output_directory is None:
            return ()
        suffix = self.output_suffix()
        return tuple(
            (source, self.output_directory / f"{source.stem}{suffix}")
            for source in self.source_paths
        )

    def accept(self) -> None:
        if self.mode is None or not self.source_paths:
            QMessageBox.warning(self, "Chưa chọn CSV", "Hãy chọn file CSV hoặc folder CSV trước khi chuyển.")
            return
        if self.mode == "folder" and not self.source_paths:
            QMessageBox.warning(self, "Không có CSV", "Folder đã chọn không có file CSV.")
            return
        if not self.batch_outputs():
            QMessageBox.warning(self, "Chưa có nơi lưu", "Hãy chọn nơi lưu file Excel kết quả.")
            return
        super().accept()

    def _refresh_summary(self) -> None:
        self.source_list.clear()
        for path in self.source_paths:
            self.source_list.addItem(str(path))
        if self.mode == "single":
            self.summary_label.setText("Chế độ: Chuyển 01 file CSV.")
            if self.single_output_path is None and self.source_paths:
                self.single_output_path = self.source_paths[0].with_suffix(self.output_suffix())
            self.output_edit.setText(str(self.single_output_path or ""))
        elif self.mode == "multiple":
            self.summary_label.setText(f"Chế độ: Chuyển nhiều file CSV ({len(self.source_paths)} file).")
            self.output_edit.setText(str(self.output_directory or ""))
        elif self.mode == "folder":
            self.summary_label.setText(f"Chế độ: Chuyển 01 folder ({len(self.source_paths)} file CSV).")
            self.output_edit.setText(str(self.output_directory or ""))

    def _format_changed(self) -> None:
        if self.mode == "single" and self.single_output_path is not None:
            self.single_output_path = self.single_output_path.with_suffix(self.output_suffix())
        self._refresh_summary()


class SplitSheetsDialog(QDialog):
    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.source_path: Path | None = None
        self.output_directory: Path | None = None
        self.sheet_names: tuple[str, ...] = ()

        self.setWindowTitle("Tách các sheet thành từng file")
        self.setModal(True)
        self.setMinimumWidth(760)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(14, 12, 14, 14)
        layout.setSpacing(8)

        guide = QLabel("Chọn workbook Excel, mỗi sheet sẽ được lưu thành một file .xlsx riêng.")
        guide.setStyleSheet("color: #0000ff; font-weight: 700;")
        guide.setWordWrap(True)
        layout.addWidget(guide)

        source_row = QHBoxLayout()
        self.source_edit = QLineEdit()
        self.source_edit.setReadOnly(True)
        source_button = QPushButton("Chọn Excel")
        source_button.clicked.connect(self.choose_source)
        source_row.addWidget(self.source_edit, 1)
        source_row.addWidget(source_button)
        layout.addLayout(source_row)

        output_row = QHBoxLayout()
        self.output_edit = QLineEdit()
        self.output_edit.setReadOnly(True)
        output_button = QPushButton("Chọn thư mục")
        output_button.clicked.connect(self.choose_output_directory)
        output_row.addWidget(self.output_edit, 1)
        output_row.addWidget(output_button)
        layout.addLayout(output_row)

        sheet_label = QLabel("Chọn sheet cần tách:")
        sheet_label.setStyleSheet("color: #0000ff; font-weight: 700;")
        layout.addWidget(sheet_label)
        self.sheet_list = QListWidget()
        self.sheet_list.setFixedHeight(150)
        self.sheet_list.itemClicked.connect(self.toggle_sheet_item)
        self.sheet_list.setStyleSheet(
            """
            QListWidget {
                padding: 6px;
            }
            QListWidget::item {
                min-height: 34px;
                padding: 4px 8px;
            }
            QListWidget::item:selected {
                background: #fce8ef;
                color: #17212b;
            }
            """
        )
        layout.addWidget(self.sheet_list)

        select_row = QHBoxLayout()
        select_all_button = QPushButton("Tích tất cả")
        select_all_button.clicked.connect(lambda: self.set_all_sheets_checked(True))
        clear_all_button = QPushButton("Bỏ tích tất cả")
        clear_all_button.clicked.connect(lambda: self.set_all_sheets_checked(False))
        select_row.addWidget(select_all_button)
        select_row.addWidget(clear_all_button)
        select_row.addStretch()
        layout.addLayout(select_row)

        button_row = QHBoxLayout()
        split_button = QPushButton("Tách sheet")
        split_button.setObjectName("PrimaryButton")
        split_button.clicked.connect(self.accept)
        cancel_button = QPushButton("Cancel")
        cancel_button.clicked.connect(self.reject)
        button_row.addStretch()
        button_row.addWidget(split_button)
        button_row.addWidget(cancel_button)
        layout.addLayout(button_row)

    def choose_source(self) -> None:
        file_name, _ = QFileDialog.getOpenFileName(
            self,
            "Chọn workbook Excel",
            str(self.source_path.parent if self.source_path else Path.home()),
            "Excel Workbook (*.xls *.xlsx *.xlsm *.xltx *.xltm);;Tất cả file (*.*)",
        )
        if not file_name:
            return
        self.source_path = Path(file_name)
        self.source_edit.setText(str(self.source_path))
        self.load_sheet_list()
        if self.output_directory is None:
            self.output_directory = self.source_path.with_name(
                f"{self.source_path.stem}_tach_sheet"
            )
            self.output_edit.setText(str(self.output_directory))

    def choose_output_directory(self) -> None:
        initial = str(self.output_directory or Path.home())
        folder = QFileDialog.getExistingDirectory(self, "Chọn thư mục lưu file", initial)
        if not folder:
            return
        self.output_directory = Path(folder)
        self.output_edit.setText(str(self.output_directory))

    def accept(self) -> None:
        if self.source_path is None:
            QMessageBox.warning(self, "Chưa chọn file Excel", "Hãy chọn workbook trước khi tách sheet.")
            return
        if self.output_directory is None:
            QMessageBox.warning(self, "Chưa chọn thư mục", "Hãy chọn thư mục lưu file kết quả.")
            return
        if not self.selected_sheet_names():
            QMessageBox.warning(self, "Chưa chọn sheet", "Hãy tích ít nhất một sheet để tách.")
            return
        super().accept()

    def load_sheet_list(self) -> None:
        self.sheet_list.clear()
        self.sheet_names = ()
        if self.source_path is None:
            return
        try:
            self.sheet_names = list_workbook_sheet_names(self.source_path)
        except Exception as exc:
            QMessageBox.warning(self, "Không đọc được workbook", f"Không tải được danh sách sheet:\n{exc}")
            return
        for sheet_name in self.sheet_names:
            item = QListWidgetItem()
            item.setData(Qt.ItemDataRole.UserRole, sheet_name)
            checkbox = QCheckBox(sheet_name)
            checkbox.setStyleSheet("font-size: 14px; padding: 7px 4px;")
            checkbox.setChecked(True)
            checkbox.setCursor(Qt.CursorShape.PointingHandCursor)
            item.setSizeHint(QSize(0, 38))
            self.sheet_list.addItem(item)
            self.sheet_list.setItemWidget(item, checkbox)

    def toggle_sheet_item(self, item: QListWidgetItem) -> None:
        checkbox = self.sheet_list.itemWidget(item)
        if isinstance(checkbox, QCheckBox):
            checkbox.setChecked(not checkbox.isChecked())

    def set_all_sheets_checked(self, checked: bool) -> None:
        for row in range(self.sheet_list.count()):
            checkbox = self.sheet_list.itemWidget(self.sheet_list.item(row))
            if isinstance(checkbox, QCheckBox):
                checkbox.setChecked(checked)

    def selected_sheet_names(self) -> tuple[str, ...]:
        selected: list[str] = []
        for row in range(self.sheet_list.count()):
            item = self.sheet_list.item(row)
            checkbox = self.sheet_list.itemWidget(item)
            if isinstance(checkbox, QCheckBox) and checkbox.isChecked():
                selected.append(str(item.data(Qt.ItemDataRole.UserRole)))
        return tuple(selected)
