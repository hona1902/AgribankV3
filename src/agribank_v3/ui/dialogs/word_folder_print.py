from __future__ import annotations

from pathlib import Path

from PySide6.QtCore import Qt, Signal
from PySide6.QtWidgets import (
    QDialog,
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

from agribank_v3.word_folder_print import WordPrintError, list_word_files
from agribank_v3.ui.dialogs.printer_settings import PrinterSettingsDialog


class WordFolderPrintDialog(QDialog):
    print_requested = Signal(tuple)

    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.folder_path: Path | None = None
        self.word_files: tuple[Path, ...] = ()

        self.setWindowTitle("In tất cả file Word trong 1 folder")
        self.setModal(True)
        self.setMinimumWidth(760)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(14, 12, 14, 14)
        layout.setSpacing(8)

        guide = QLabel(
            "Chọn thư mục chứa các file Word. Ứng dụng sẽ in các file .doc, "
            ".docx, .docm, .rtf trong thư mục theo thứ tự tên file."
        )
        guide.setStyleSheet("color: #0000ff; font-weight: 700;")
        guide.setWordWrap(True)
        layout.addWidget(guide)

        folder_row = QHBoxLayout()
        self.folder_edit = QLineEdit()
        self.folder_edit.setReadOnly(True)
        folder_row.addWidget(self.folder_edit, 1)
        choose_button = QPushButton("Chọn folder")
        choose_button.clicked.connect(self.choose_folder)
        folder_row.addWidget(choose_button)
        layout.addLayout(folder_row)

        self.file_list = QListWidget()
        self.file_list.setFixedHeight(220)
        layout.addWidget(self.file_list)

        self.status_label = QLabel("Chưa chọn thư mục.")
        self.status_label.setTextInteractionFlags(
            Qt.TextInteractionFlag.TextSelectableByMouse
        )
        self.status_label.setWordWrap(True)
        layout.addWidget(self.status_label)

        button_row = QHBoxLayout()
        self.print_button = QPushButton("In tất cả")
        self.print_button.setObjectName("PrimaryButton")
        self.print_button.clicked.connect(self.accept)
        self.print_button.setEnabled(False)
        printer_settings_button = QPushButton("Cài đặt máy in")
        printer_settings_button.clicked.connect(self.show_printer_settings)
        cancel_button = QPushButton("Cancel")
        cancel_button.clicked.connect(self.reject)
        button_row.addStretch()
        button_row.addWidget(printer_settings_button)
        button_row.addWidget(self.print_button)
        button_row.addWidget(cancel_button)
        layout.addLayout(button_row)

    def choose_folder(self) -> None:
        initial = str(self.folder_path or Path.home())
        folder_name = QFileDialog.getExistingDirectory(
            self,
            "Chọn folder chứa file Word",
            initial,
        )
        if not folder_name:
            return
        self.load_folder(Path(folder_name))

    def load_folder(self, folder_path: Path) -> None:
        try:
            files = list_word_files(folder_path)
        except WordPrintError as exc:
            QMessageBox.warning(self, "Không thể đọc folder", str(exc))
            return
        self.folder_path = folder_path
        self.word_files = files
        self.folder_edit.setText(str(folder_path))
        self.file_list.clear()
        for path in files:
            self.file_list.addItem(path.name)
        self.print_button.setEnabled(bool(files))
        if files:
            self.status_label.setText(f"Đã tìm thấy {len(files)} file Word.")
        else:
            self.status_label.setText("Không tìm thấy file Word trong thư mục đã chọn.")

    def show_printer_settings(self) -> None:
        dialog = PrinterSettingsDialog(self)
        dialog.exec()

    def accept(self) -> None:
        if not self.word_files:
            QMessageBox.warning(
                self,
                "Chưa có file Word",
                "Hãy chọn thư mục có ít nhất một file Word trước khi in.",
            )
            return
        self.print_requested.emit(self.word_files)
        super().accept()
