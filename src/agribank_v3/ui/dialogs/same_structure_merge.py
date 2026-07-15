from __future__ import annotations

from pathlib import Path

from PySide6.QtWidgets import (
    QAbstractItemView,
    QCheckBox,
    QDialog,
    QDialogButtonBox,
    QFileDialog,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QListWidget,
    QMessageBox,
    QPushButton,
    QVBoxLayout,
    QWidget,
)


class SameStructureMergeDialog(QDialog):
    EXCEL_SUFFIXES = {".xls", ".xlsx", ".xlsm", ".xltx", ".xltm"}

    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.source_paths: list[Path] = []

        self.setWindowTitle("Nối file cùng cấu trúc")
        self.setModal(True)
        self.setMinimumWidth(760)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(14, 12, 14, 14)
        layout.setSpacing(8)

        source_label = QLabel(
            "Chọn các file Excel hoặc CSV cùng cấu trúc. Ứng dụng lấy dòng 1 "
            "làm tiêu đề và nối dữ liệu từ dòng 2 của các file sau vào một file .xlsx."
        )
        source_label.setStyleSheet("color: #0000ff; font-weight: 700;")
        source_label.setWordWrap(True)
        layout.addWidget(source_label)

        self.source_list = QListWidget()
        self.source_list.setSelectionMode(
            QAbstractItemView.SelectionMode.ExtendedSelection
        )
        self.source_list.setFixedHeight(118)
        layout.addWidget(self.source_list)

        choose_row = QHBoxLayout()
        choose_button = QPushButton("Chọn File")
        choose_button.clicked.connect(self.choose_source_files)
        remove_button = QPushButton("Xóa file")
        remove_button.clicked.connect(self.remove_selected_sources)
        clear_button = QPushButton("Xóa danh sách")
        clear_button.clicked.connect(self.clear_sources)
        move_up_button = QPushButton("Di chuyển lên")
        move_up_button.clicked.connect(self.move_selected_up)
        move_down_button = QPushButton("Di chuyển xuống")
        move_down_button.clicked.connect(self.move_selected_down)
        choose_row.addWidget(choose_button)
        choose_row.addWidget(remove_button)
        choose_row.addWidget(clear_button)
        choose_row.addWidget(move_up_button)
        choose_row.addWidget(move_down_button)
        choose_row.addStretch()
        layout.addLayout(choose_row)

        output_label = QLabel("Đường dẫn lưu file kết quả:")
        output_label.setStyleSheet("color: #0000ff; font-weight: 700;")
        layout.addWidget(output_label)
        output_row = QHBoxLayout()
        self.output_directory_edit = QLineEdit()
        self.output_directory_edit.setReadOnly(True)
        output_row.addWidget(self.output_directory_edit, 1)
        choose_output_button = QPushButton("Lưu tại...")
        choose_output_button.clicked.connect(self.choose_output_file)
        output_row.addWidget(choose_output_button)
        layout.addLayout(output_row)

        output_name_label = QLabel("Tên file xuất ra:")
        output_name_label.setStyleSheet("color: #0000ff; font-weight: 700;")
        layout.addWidget(output_name_label)
        self.output_name_edit = QLineEdit()
        self.output_name_edit.setPlaceholderText("Nhập tên file kết quả")
        layout.addWidget(self.output_name_edit)

        self.source_filename_checkbox = QCheckBox(
            "Thêm 1 cột ghi chú tên file gốc sau cùng của bảng tính"
        )
        layout.addWidget(self.source_filename_checkbox)

        buttons = QDialogButtonBox()
        merge_button = QPushButton("Nối File")
        merge_button.setObjectName("PrimaryButton")
        buttons.addButton(merge_button, QDialogButtonBox.ButtonRole.AcceptRole)
        cancel_button = QPushButton("Cancel")
        buttons.addButton(cancel_button, QDialogButtonBox.ButtonRole.RejectRole)
        merge_button.clicked.connect(self.accept)
        cancel_button.clicked.connect(self.reject)
        layout.addWidget(buttons)

    def choose_source_files(self) -> None:
        initial = str(self.source_paths[-1].parent) if self.source_paths else ""
        file_names, _ = QFileDialog.getOpenFileNames(
            self,
            "Chọn các file cùng cấu trúc",
            initial,
            "Excel và CSV (*.xls *.xlsx *.xlsm *.csv);;Excel (*.xls *.xlsx *.xlsm);;CSV (*.csv);;Tất cả file (*.*)",
        )
        if not file_names:
            return
        for file_name in file_names:
            path = Path(file_name)
            if path not in self.source_paths:
                self.source_paths.append(path)
                self.source_list.addItem(str(path))
        self._refresh_output_path()

    def choose_output_file(self) -> None:
        initial = str(self.output_path() or Path.home() / "NoiFileCungCauTruc.xlsx")
        file_name, _ = QFileDialog.getSaveFileName(
            self,
            "Chọn file kết quả",
            initial,
            "Excel Workbook (*.xlsx)",
        )
        if not file_name:
            return
        path = Path(file_name)
        if path.suffix.casefold() != ".xlsx":
            path = path.with_suffix(".xlsx")
        self.output_directory_edit.setText(str(path.parent))
        self.output_name_edit.setText(path.name)

    def clear_sources(self) -> None:
        self.source_paths.clear()
        self.source_list.clear()
        self.output_directory_edit.clear()
        self.output_name_edit.clear()

    def remove_selected_sources(self) -> None:
        selected_rows = self._selected_rows()
        if not selected_rows:
            return
        for row in reversed(selected_rows):
            del self.source_paths[row]
        self._rebuild_source_list()
        if self.source_paths:
            next_row = min(selected_rows[0], len(self.source_paths) - 1)
            self.source_list.setCurrentRow(next_row)
        else:
            self.output_directory_edit.clear()
            self.output_name_edit.clear()

    def move_selected_up(self) -> None:
        selected_rows = self._selected_rows()
        if not selected_rows or selected_rows[0] == 0:
            return
        selected_set = set(selected_rows)
        for row in selected_rows:
            if row - 1 not in selected_set:
                self.source_paths[row - 1], self.source_paths[row] = (
                    self.source_paths[row],
                    self.source_paths[row - 1],
                )
        self._rebuild_source_list([row - 1 for row in selected_rows])

    def move_selected_down(self) -> None:
        selected_rows = self._selected_rows()
        if not selected_rows or selected_rows[-1] == len(self.source_paths) - 1:
            return
        selected_set = set(selected_rows)
        for row in reversed(selected_rows):
            if row + 1 not in selected_set:
                self.source_paths[row + 1], self.source_paths[row] = (
                    self.source_paths[row],
                    self.source_paths[row + 1],
                )
        self._rebuild_source_list([row + 1 for row in selected_rows])

    def include_source_filename(self) -> bool:
        return self.source_filename_checkbox.isChecked()

    def output_path(self) -> Path | None:
        directory_text = self.output_directory_edit.text().strip()
        name_text = self.output_name_edit.text().strip()
        if not directory_text and not self.source_paths:
            return None
        directory = Path(directory_text) if directory_text else self.source_paths[0].parent
        name = name_text or "NoiFileCungCauTruc.xlsx"
        path = Path(name)
        if path.name != name:
            path = Path(path.name)
        if path.suffix.casefold() != ".xlsx":
            path = path.with_suffix(".xlsx")
        return directory / path.name

    def source_kind(self) -> str:
        suffixes = {path.suffix.casefold() for path in self.source_paths}
        if suffixes == {".csv"}:
            return "csv"
        if suffixes and suffixes <= self.EXCEL_SUFFIXES:
            return "excel"
        return "mixed"

    def accept(self) -> None:
        if not self.source_paths:
            QMessageBox.warning(
                self,
                "Chưa chọn file nguồn",
                "Hãy chọn ít nhất một file trước khi nối.",
            )
            return
        kind = self.source_kind()
        if kind == "mixed":
            QMessageBox.warning(
                self,
                "Không cùng loại file",
                "Chỉ nối một nhóm toàn file Excel hoặc toàn file CSV trong một lần chạy.",
            )
            return
        output_path = self.output_path()
        if output_path is None:
            QMessageBox.warning(
                self,
                "Chưa có file kết quả",
                "Không xác định được file kết quả sau khi nối.",
            )
            return
        if any(_same_path(output_path, source_path) for source_path in self.source_paths):
            QMessageBox.warning(
                self,
                "File kết quả trùng file nguồn",
                "Hãy chọn file kết quả khác với các file nguồn.",
            )
            return
        self.output_directory_edit.setText(str(output_path.parent))
        self.output_name_edit.setText(output_path.name)
        super().accept()

    def _refresh_output_path(self) -> None:
        if not self.source_paths:
            return
        if not self.output_directory_edit.text().strip():
            self.output_directory_edit.setText(str(self.source_paths[0].parent))
        if not self.output_name_edit.text().strip():
            self.output_name_edit.setText("NoiFileCungCauTruc.xlsx")

    def _selected_rows(self) -> list[int]:
        return sorted(index.row() for index in self.source_list.selectedIndexes())

    def _rebuild_source_list(self, selected_rows: list[int] | None = None) -> None:
        self.source_list.clear()
        for path in self.source_paths:
            self.source_list.addItem(str(path))
        if selected_rows:
            self.source_list.setCurrentRow(selected_rows[0])
            for row in selected_rows:
                if 0 <= row < self.source_list.count():
                    self.source_list.item(row).setSelected(True)
        if self.source_paths:
            self._refresh_output_path()


def _same_path(left: Path, right: Path) -> bool:
    try:
        return left.resolve() == right.resolve()
    except OSError:
        return left.absolute() == right.absolute()
