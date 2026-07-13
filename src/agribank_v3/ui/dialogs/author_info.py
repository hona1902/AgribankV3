from __future__ import annotations

from PySide6.QtCore import QPoint, Qt, QTimer
from PySide6.QtGui import QMouseEvent, QPixmap
from PySide6.QtWidgets import (
    QApplication,
    QDialog,
    QFrame,
    QGridLayout,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QVBoxLayout,
    QWidget,
)

from agribank_v3 import __version__
from agribank_v3.ui.icons import app_icon, icon_path


AUTHOR_NAME = "Nguyễn Hoài Nam"
ORGANIZATION = "Agribank Chi nhánh Lộc Phát Lâm Đồng"
CONTACT = "0972.173.064 (ZALO)"
DESCRIPTION = (
    "Ứng dụng hỗ trợ Excel nghiệp vụ, tra cứu dữ liệu, tạo nhanh các loại "
    "báo cáo hàng ngày, báo cáo quyết toán và ôn tập kiến thức nghiệp vụ."
)
PURPOSE = (
    "Hỗ trợ cán bộ trong quá trình xử lý công việc, học tập, tra cứu và "
    "chuẩn hóa biểu mẫu nghiệp vụ."
)
COPYRIGHT = "© 2026 Nguyễn Hoài Nam. Phát triển phục vụ nội bộ."
SECURITY_NOTICE = (
    "Nội dung và dữ liệu trong ứng dụng chỉ phục vụ công việc nội bộ. "
    "Người dùng cần tuân thủ quy định bảo mật thông tin và quy trình "
    "nghiệp vụ của Agribank."
)


def author_information_text() -> str:
    return "\n".join(
        (
            "AgribankV3 - Công cụ Excel nghiệp vụ",
            "Tên ứng dụng: AgribankV3",
            f"Tác giả: {AUTHOR_NAME}",
            f"Đơn vị: {ORGANIZATION}",
            f"Phiên bản: {__version__} - Bản thử nghiệm",
            f"Liên hệ: {CONTACT}",
            f"Mô tả: {DESCRIPTION}",
            f"Mục đích: {PURPOSE}",
            f"Bản quyền: {COPYRIGHT}",
            f"Lưu ý: {SECURITY_NOTICE}",
        )
    )


class AuthorInfoDialog(QDialog):
    """Modal presenting application, author, and internal-use information."""

    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setObjectName("AuthorInfoDialog")
        self.setWindowTitle("Thông tin tác giả - AgribankV3")
        self.setWindowIcon(app_icon())
        self.setModal(True)
        self.setWindowModality(Qt.WindowModality.WindowModal)
        self.setFixedSize(620, 570)
        self.setWindowFlag(Qt.WindowType.FramelessWindowHint, True)
        self.setWindowFlag(Qt.WindowType.WindowContextHelpButtonHint, False)
        self._drag_start: QPoint | None = None
        self._build_ui()
        self._apply_style()

    def _build_ui(self) -> None:
        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(0)

        header = QFrame()
        header.setObjectName("AuthorHeader")
        header.setFixedHeight(100)
        header_layout = QHBoxLayout(header)
        header_layout.setContentsMargins(24, 16, 16, 16)
        header_layout.setSpacing(16)

        logo = QLabel()
        logo.setObjectName("AuthorLogo")
        pixmap = QPixmap(icon_path("logoagri.png"))
        logo.setPixmap(
            pixmap.scaled(
                62,
                62,
                Qt.AspectRatioMode.KeepAspectRatio,
                Qt.TransformationMode.SmoothTransformation,
            )
        )
        logo.setFixedSize(64, 64)

        brand = QVBoxLayout()
        brand.setSpacing(2)
        title = QLabel("AgribankV3")
        title.setObjectName("AuthorBrandTitle")
        subtitle = QLabel("Công cụ Excel nghiệp vụ")
        subtitle.setObjectName("AuthorBrandSubtitle")
        brand.addStretch()
        brand.addWidget(title)
        brand.addWidget(subtitle)
        brand.addStretch()

        header_layout.addWidget(logo)
        header_layout.addLayout(brand, stretch=1)
        root.addWidget(header)

        content = QWidget()
        content.setObjectName("AuthorContent")
        content_layout = QVBoxLayout(content)
        content_layout.setContentsMargins(28, 20, 28, 20)
        content_layout.setSpacing(12)

        section_row = QHBoxLayout()
        section_icon = QLabel("ⓘ")
        section_icon.setObjectName("AuthorSectionIcon")
        section_icon.setAlignment(Qt.AlignmentFlag.AlignCenter)
        section_title = QLabel("THÔNG TIN TÁC GIẢ")
        section_title.setObjectName("AuthorSectionTitle")
        section_row.addWidget(section_icon)
        section_row.addWidget(section_title)
        section_row.addStretch()
        content_layout.addLayout(section_row)

        details = QFrame()
        details.setObjectName("AuthorDetailsCard")
        details_grid = QGridLayout(details)
        details_grid.setContentsMargins(18, 14, 18, 14)
        details_grid.setHorizontalSpacing(16)
        details_grid.setVerticalSpacing(8)
        details_grid.setColumnMinimumWidth(0, 105)
        details_grid.setColumnStretch(1, 1)

        rows = (
            ("Tên ứng dụng", "AgribankV3"),
            ("Tác giả", AUTHOR_NAME),
            ("Đơn vị", ORGANIZATION),
            ("Phiên bản", f"{__version__} - Bản thử nghiệm"),
            ("Liên hệ", CONTACT),
            ("Mô tả", DESCRIPTION),
            ("Mục đích", PURPOSE),
            ("Bản quyền", COPYRIGHT),
        )
        for row, (label_text, value_text) in enumerate(rows):
            label = QLabel(label_text)
            label.setObjectName("AuthorFieldLabel")
            label.setAlignment(
                Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignLeft
            )
            value = QLabel(value_text)
            value.setObjectName("AuthorFieldValue")
            value.setWordWrap(True)
            value.setTextInteractionFlags(
                Qt.TextInteractionFlag.TextSelectableByMouse
            )
            details_grid.addWidget(label, row, 0)
            details_grid.addWidget(value, row, 1)
        content_layout.addWidget(details)

        notice = QFrame()
        notice.setObjectName("AuthorNotice")
        notice_layout = QHBoxLayout(notice)
        notice_layout.setContentsMargins(14, 10, 14, 10)
        notice_layout.setSpacing(10)
        shield = QLabel("◆")
        shield.setObjectName("AuthorNoticeIcon")
        shield.setAlignment(Qt.AlignmentFlag.AlignTop)
        notice_text = QLabel(SECURITY_NOTICE)
        notice_text.setObjectName("AuthorNoticeText")
        notice_text.setWordWrap(True)
        notice_layout.addWidget(shield)
        notice_layout.addWidget(notice_text, stretch=1)
        content_layout.addWidget(notice)

        buttons = QHBoxLayout()
        buttons.setSpacing(10)
        copy_button = QPushButton("Sao chép thông tin")
        copy_button.setObjectName("AuthorCopyButton")
        copy_button.setCursor(Qt.CursorShape.PointingHandCursor)
        copy_button.clicked.connect(
            lambda: self._copy_information(copy_button)
        )
        close_button = QPushButton("Đóng")
        close_button.setObjectName("AuthorCloseButton")
        close_button.setCursor(Qt.CursorShape.PointingHandCursor)
        close_button.setDefault(True)
        close_button.clicked.connect(self.accept)
        buttons.addStretch()
        buttons.addWidget(copy_button)
        buttons.addWidget(close_button)
        content_layout.addLayout(buttons)
        root.addWidget(content, stretch=1)

    @staticmethod
    def _copy_information(button: QPushButton) -> None:
        QApplication.clipboard().setText(author_information_text())
        button.setText("Đã sao chép")
        button.setEnabled(False)

        def restore_button() -> None:
            if button.isVisible():
                button.setText("Sao chép thông tin")
                button.setEnabled(True)

        QTimer.singleShot(1400, restore_button)

    def _apply_style(self) -> None:
        self.setStyleSheet(
            """
            QDialog#AuthorInfoDialog {
                background: #fff1f5;
                border: 2px solid #AE1C3F;
            }
            QFrame#AuthorHeader { background: #831f41; border: none; }
            QLabel#AuthorLogo {
                background: rgba(255, 255, 255, 0.96);
                border: 1px solid rgba(255, 255, 255, 0.50);
                border-radius: 10px;
                padding: 2px;
            }
            QLabel#AuthorBrandTitle {
                color: white; font-size: 25px; font-weight: 750;
            }
            QLabel#AuthorBrandSubtitle { color: #f6dce5; font-size: 13px; }
            QPushButton#AuthorHeaderClose {
                color: white; background: transparent; border: none;
                border-radius: 17px; font-size: 24px; font-weight: 500;
                padding: 0;
            }
            QPushButton#AuthorHeaderClose:hover {
                background: rgba(255, 255, 255, 0.16);
            }
            QWidget#AuthorContent { background: #fff1f5; }
            QLabel#AuthorSectionIcon {
                color: white; background: #a72a53; border-radius: 11px;
                min-width: 22px; min-height: 22px;
                max-width: 22px; max-height: 22px; font-weight: 700;
            }
            QLabel#AuthorSectionTitle {
                color: #3d1626; font-size: 15px; font-weight: 750;
            }
            QFrame#AuthorDetailsCard {
                background: white; border: 1px solid #e1e5ea;
                border-radius: 10px;
            }
            QLabel#AuthorFieldLabel {
                color: #6b7280; font-size: 12px; font-weight: 600;
            }
            QLabel#AuthorFieldValue {
                color: #1f2937; font-size: 12px; font-weight: 550;
            }
            QFrame#AuthorNotice {
                background: #fff2f6; border: 1px solid #efcbd8;
                border-radius: 8px;
            }
            QLabel#AuthorNoticeIcon { color: #9b1f4d; font-size: 13px; }
            QLabel#AuthorNoticeText { color: #5c3745; font-size: 11px; }
            QPushButton#AuthorCopyButton {
                color: #831f41; background: white;
                border: 1px solid #d9dfe5; border-radius: 7px;
                padding: 9px 15px; font-weight: 600;
            }
            QPushButton#AuthorCopyButton:hover {
                background: #fff4f7; border-color: #bd6b87;
            }
            QPushButton#AuthorCopyButton:disabled {
                color: #00843d; background: #f0faf5;
            }
            QPushButton#AuthorCloseButton {
                color: white; background: #931f49; border: none;
                border-radius: 7px; min-width: 78px;
                padding: 10px 18px; font-weight: 650;
            }
            QPushButton#AuthorCloseButton:hover { background: #ad2c57; }
            """
        )

    def mousePressEvent(self, event: QMouseEvent) -> None:
        if (
            event.button() == Qt.MouseButton.LeftButton
            and event.position().y() <= 100
        ):
            self._drag_start = event.globalPosition().toPoint() - self.frameGeometry().topLeft()
            event.accept()
            return
        super().mousePressEvent(event)

    def mouseMoveEvent(self, event: QMouseEvent) -> None:
        if self._drag_start is not None and event.buttons() & Qt.MouseButton.LeftButton:
            self.move(event.globalPosition().toPoint() - self._drag_start)
            event.accept()
            return
        super().mouseMoveEvent(event)

    def mouseReleaseEvent(self, event: QMouseEvent) -> None:
        self._drag_start = None
        super().mouseReleaseEvent(event)
