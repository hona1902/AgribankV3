from __future__ import annotations

from PySide6.QtCore import Qt
from PySide6.QtWidgets import QDialog, QDialogButtonBox, QLabel, QVBoxLayout, QWidget


AUTO_INTEREST_PLACEHOLDER_TITLE = "Thu lãi bán tự động"


class AutoInterestPlaceholderDialog(QDialog):
    """Temporary window for the auto-interest collection migration group."""

    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setWindowTitle(AUTO_INTEREST_PLACEHOLDER_TITLE)
        self.setModal(True)
        self.setMinimumWidth(560)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(18, 16, 18, 14)
        layout.setSpacing(12)

        title_label = QLabel(AUTO_INTEREST_PLACEHOLDER_TITLE)
        title_label.setObjectName("SectionTitle")
        layout.addWidget(title_label)

        message = QLabel(
            "Chức năng đang được chuyển đổi từ agribank-tool.xlam. "
            "Nhóm này tách riêng với Tổ vay vốn và hiện chưa chuyển logic nghiệp vụ."
        )
        message.setObjectName("MutedText")
        message.setWordWrap(True)
        message.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignTop)
        layout.addWidget(message)

        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Close)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)
