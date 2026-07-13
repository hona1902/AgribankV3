from __future__ import annotations

from PySide6.QtCore import QPoint, Qt
from PySide6.QtGui import QMouseEvent, QPixmap
from PySide6.QtWidgets import (
    QDialog,
    QFrame,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QScrollArea,
    QTabWidget,
    QVBoxLayout,
    QWidget,
)

from agribank_v3.ui.icons import app_icon, icon_path


class SettlementGuidanceDialog(QDialog):
    """Tabbed settlement guidance dialog."""

    CREATE_30A_TAB = 0
    CONSOLIDATION_TAB = 1

    def __init__(
        self,
        parent: QWidget | None = None,
        *,
        active_tab: int = CREATE_30A_TAB,
    ) -> None:
        super().__init__(parent)
        self.setObjectName("SettlementGuidanceDialog")
        self.setWindowTitle("Hướng dẫn quyết toán - AgribankV3")
        self.setWindowIcon(app_icon())
        self.setModal(True)
        self.setWindowModality(Qt.WindowModality.WindowModal)
        self.setFixedSize(700, 560)
        self.setWindowFlag(Qt.WindowType.FramelessWindowHint, True)
        self.setWindowFlag(Qt.WindowType.WindowContextHelpButtonHint, False)
        self._drag_start: QPoint | None = None
        self._build_ui()
        self.tabs.setCurrentIndex(active_tab)
        self._apply_style()

    def _build_ui(self) -> None:
        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(0)

        header = QFrame()
        header.setObjectName("GuideHeader")
        header.setFixedHeight(100)
        header_layout = QHBoxLayout(header)
        header_layout.setContentsMargins(24, 16, 16, 16)
        header_layout.setSpacing(16)

        logo = QLabel()
        logo.setObjectName("GuideLogo")
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
        title.setObjectName("GuideBrandTitle")
        subtitle = QLabel("Hướng dẫn quyết toán")
        subtitle.setObjectName("GuideBrandSubtitle")
        brand.addStretch()
        brand.addWidget(title)
        brand.addWidget(subtitle)
        brand.addStretch()

        header_layout.addWidget(logo)
        header_layout.addLayout(brand, stretch=1)
        root.addWidget(header)

        content = QWidget()
        content.setObjectName("GuideContent")
        content_layout = QVBoxLayout(content)
        content_layout.setContentsMargins(28, 18, 28, 20)
        content_layout.setSpacing(12)

        section_row = QHBoxLayout()
        section_icon = QLabel("?")
        section_icon.setObjectName("GuideSectionIcon")
        section_icon.setAlignment(Qt.AlignmentFlag.AlignCenter)
        section_title = QLabel("HƯỚNG DẪN QUYẾT TOÁN")
        section_title.setObjectName("GuideSectionTitle")
        section_row.addWidget(section_icon)
        section_row.addWidget(section_title)
        section_row.addStretch()
        content_layout.addLayout(section_row)

        self.tabs = QTabWidget()
        self.tabs.setObjectName("GuideTabs")
        self.tabs.addTab(
            self._scroll_tab(
                self._create_30a_content(),
            ),
            "Hướng dẫn tạo mẫu 30a",
        )
        self.tabs.addTab(
            self._scroll_tab(
                self._consolidation_content(),
            ),
            "Hướng dẫn tổng hợp quyết toán",
        )
        content_layout.addWidget(self.tabs, stretch=1)

        buttons = QHBoxLayout()
        close_button = QPushButton("Đóng")
        close_button.setObjectName("GuideCloseButton")
        close_button.setCursor(Qt.CursorShape.PointingHandCursor)
        close_button.setDefault(True)
        close_button.clicked.connect(self.accept)
        buttons.addStretch()
        buttons.addWidget(close_button)
        content_layout.addLayout(buttons)
        root.addWidget(content, stretch=1)

    def _scroll_tab(self, body: QWidget) -> QScrollArea:
        scroll = QScrollArea()
        scroll.setObjectName("GuideScrollArea")
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.Shape.NoFrame)
        scroll.setWidget(body)
        return scroll

    def _create_30a_content(self) -> QWidget:
        body = QWidget()
        layout = QVBoxLayout(body)
        layout.setContentsMargins(0, 12, 0, 12)
        layout.setSpacing(10)
        self._add_step(layout, "Bước 1", "Chọn mã mẫu quyết toán trong menu sổ xuống.")
        self._add_step(
            layout,
            "Bước 2",
            "Chọn File quyết toán đã được Chương trình tạo ra (ví dụ 5491QT05.xlsx).",
        )
        self._add_step(
            layout,
            "Bước 3",
            (
                "+ Đối với tổng hợp quyết toán chi nhánh loại I:\n"
                "- Xuất các file cân đối sau dồn tích được xuất ra từ màn hình GLCB41 theo từng chi nhánh vào cùng 01 thư mục.\n"
                "- Tên file cân đối để định dạng MãCN.xls hoặc MãCN.xlsx (ví dụ: 5491.xls, 5404.xls...).\n"
                "- Chọn thư mục chứa các file cân đối sau dồn tích của các chi nhánh.\n\n"
                "+ Đối với file quyết toán chi nhánh loại II:\n"
                "- Chọn file cân đối sau dồn tích, được xuất ra từ màn hình GLCB41."
            ),
        )
        layout.addStretch()
        return body

    def _consolidation_content(self) -> QWidget:
        body = QWidget()
        layout = QVBoxLayout(body)
        layout.setContentsMargins(0, 12, 0, 12)
        layout.setSpacing(10)
        self._add_step(
            layout,
            "Bước 1",
            "Chép các file dữ liệu quyết toán (Ví dụ: Mã CN_rt05.csv) của cùng 1 mẫu biểu vào 01 thư mục.",
        )
        self._add_step(
            layout,
            "Bước 2",
            "Vào Menu Quyết toán tổng hợp -> Chọn Mẫu quyết toán theo mã mẫu báo cáo vừa copy.",
        )
        self._add_step(
            layout,
            "Bước 3",
            "Chọn thư mục vừa copy các file csv.",
        )
        layout.addStretch()
        return body

    @staticmethod
    def _add_step(layout: QVBoxLayout, title: str, text: str) -> None:
        card = QFrame()
        card.setObjectName("GuideStepCard")
        card_layout = QVBoxLayout(card)
        card_layout.setContentsMargins(16, 12, 16, 12)
        card_layout.setSpacing(6)
        title_label = QLabel(title)
        title_label.setObjectName("GuideStepTitle")
        body_label = QLabel(text)
        body_label.setObjectName("GuideStepText")
        body_label.setWordWrap(True)
        body_label.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
        card_layout.addWidget(title_label)
        card_layout.addWidget(body_label)
        layout.addWidget(card)

    def _apply_style(self) -> None:
        self.setStyleSheet(
            """
            QDialog#SettlementGuidanceDialog {
                background: #fff1f5;
                border: 2px solid #AE1C3F;
            }
            QFrame#GuideHeader { background: #831f41; border: none; }
            QLabel#GuideLogo {
                background: rgba(255, 255, 255, 0.96);
                border: 1px solid rgba(255, 255, 255, 0.50);
                border-radius: 10px;
                padding: 2px;
            }
            QLabel#GuideBrandTitle {
                color: white; font-size: 25px; font-weight: 750;
            }
            QLabel#GuideBrandSubtitle { color: #f6dce5; font-size: 13px; }
            QPushButton#GuideHeaderClose {
                color: white; background: transparent; border: none;
                border-radius: 17px; font-size: 24px; font-weight: 500;
                padding: 0;
            }
            QPushButton#GuideHeaderClose:hover {
                background: rgba(255, 255, 255, 0.16);
            }
            QWidget#GuideContent { background: #fff1f5; }
            QLabel#GuideSectionIcon {
                color: white; background: #a72a53; border-radius: 11px;
                min-width: 22px; min-height: 22px;
                max-width: 22px; max-height: 22px; font-weight: 700;
            }
            QLabel#GuideSectionTitle {
                color: #3d1626; font-size: 15px; font-weight: 750;
            }
            QTabWidget#GuideTabs::pane {
                background: white; border: 1px solid #e1e5ea;
                border-radius: 10px; top: -1px;
            }
            QTabWidget#GuideTabs QTabBar::tab {
                color: #4b5563; background: #eef1f5;
                border: 1px solid #dce2ea; border-bottom: none;
                padding: 9px 14px; margin-right: 4px;
                border-top-left-radius: 8px; border-top-right-radius: 8px;
                font-weight: 650;
            }
            QTabWidget#GuideTabs QTabBar::tab:selected {
                color: #831f41; background: white;
            }
            QScrollArea#GuideScrollArea { background: white; border: none; }
            QFrame#GuideStepCard {
                background: #ffffff; border: 1px solid #e1e5ea;
                border-radius: 8px;
            }
            QLabel#GuideStepTitle {
                color: #831f41; font-size: 13px; font-weight: 750;
            }
            QLabel#GuideStepText {
                color: #1f2937; font-size: 12px; line-height: 1.35;
            }
            QPushButton#GuideCloseButton {
                color: white; background: #931f49; border: none;
                border-radius: 7px; min-width: 78px;
                padding: 10px 18px; font-weight: 650;
            }
            QPushButton#GuideCloseButton:hover { background: #ad2c57; }
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
