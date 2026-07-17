from __future__ import annotations

from PySide6.QtCore import Qt
from PySide6.QtGui import QPixmap
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


TOVAYVON_HELP_TITLE = "Hướng dẫn tổ vay vốn"


class ToVayVonHelpWindow(QDialog):
    """Tabbed guidance dialog for the Tổ vay vốn workflow."""

    INTEREST_REPORT_TAB = 0
    PAYMENT_REQUEST_TAB = 1
    DEBT_RECONCILIATION_TAB = 2

    def __init__(
        self,
        parent: QWidget | None = None,
        *,
        active_tab: int = INTEREST_REPORT_TAB,
    ) -> None:
        super().__init__(parent)
        self.setObjectName("ToVayVonHelpWindow")
        self.setWindowTitle("Hướng dẫn tổ vay vốn - AgribankV3")
        self.setWindowIcon(app_icon())
        self.setModal(True)
        self.setWindowModality(Qt.WindowModality.WindowModal)
        self.setFixedSize(780, 640)
        self.setWindowFlag(Qt.WindowType.WindowContextHelpButtonHint, False)
        self._build_ui()
        self.tabs.setCurrentIndex(active_tab)
        self._apply_style()

    def _build_ui(self) -> None:
        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(0)

        header = QFrame()
        header.setObjectName("GuideHeader")
        header.setFixedHeight(112)
        header_layout = QHBoxLayout(header)
        header_layout.setContentsMargins(24, 16, 20, 16)
        header_layout.setSpacing(16)

        logo = QLabel()
        logo.setObjectName("GuideLogo")
        pixmap = QPixmap(icon_path("logoagri.png"))
        logo.setPixmap(
            pixmap.scaled(
                66,
                66,
                Qt.AspectRatioMode.KeepAspectRatio,
                Qt.TransformationMode.SmoothTransformation,
            )
        )
        logo.setFixedSize(68, 68)

        brand = QVBoxLayout()
        brand.setSpacing(4)
        title = QLabel("AgribankV3")
        title.setObjectName("GuideBrandTitle")
        subtitle = QLabel(TOVAYVON_HELP_TITLE)
        subtitle.setObjectName("GuideBrandSubtitle")
        description = QLabel(
            "Hướng dẫn thực hiện các chức năng: tạo bảng kê thu lãi, lập đề nghị "
            "thanh toán hoa hồng và đối chiếu dư nợ theo tổ vay vốn."
        )
        description.setObjectName("GuideBrandDescription")
        description.setWordWrap(True)
        brand.addStretch()
        brand.addWidget(title)
        brand.addWidget(subtitle)
        brand.addWidget(description)
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
        section_icon = QLabel("i")
        section_icon.setObjectName("GuideSectionIcon")
        section_icon.setAlignment(Qt.AlignmentFlag.AlignCenter)
        section_title = QLabel("HƯỚNG DẪN TỔ VAY VỐN")
        section_title.setObjectName("GuideSectionTitle")
        section_row.addWidget(section_icon)
        section_row.addWidget(section_title)
        section_row.addStretch()
        content_layout.addLayout(section_row)

        self.tabs = QTabWidget()
        self.tabs.setObjectName("GuideTabs")
        self.tabs.addTab(
            self._scroll_tab(self._interest_report_content()),
            "Hướng dẫn tạo bảng kê thu lãi",
        )
        self.tabs.addTab(
            self._scroll_tab(self._payment_request_content()),
            "Đề nghị thanh toán",
        )
        self.tabs.addTab(
            self._scroll_tab(self._debt_reconciliation_content()),
            "Đối chiếu tổ vay vốn",
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
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        scroll.setWidget(body)
        return scroll

    def _interest_report_content(self) -> QWidget:
        body = QWidget()
        layout = QVBoxLayout(body)
        layout.setContentsMargins(0, 12, 0, 12)
        layout.setSpacing(10)
        self._add_step(
            layout,
            "Bước 01 - Cập nhật thông tin tổ vay vốn",
            (
                "Cập nhật thông tin bằng chức năng:\n"
                "Tổ vay vốn -> Quản lý tổ vay vốn\n\n"
                "- Có thể thêm/sửa/xóa tổ vay vốn.\n"
                "- Có thể nhập thông tin tổ bằng file Excel.\n"
                "- Cần kiểm tra đầy đủ thông tin tổ trưởng, tài khoản, tổ hội, "
                "tỷ lệ hoa hồng và điều kiện chi trước khi tạo bảng kê."
            ),
            "Lưu ý: Cần cập nhật đầy đủ tỷ lệ hoa hồng trước khi tạo bảng kê.",
        )
        self._add_step(
            layout,
            "Bước 02 - Xuất Sao kê thu lãi trong kỳ",
            (
                "Trong IPCAS:\n"
                "Loan -> lnlr (Báo cáo tín dụng) -> lnlr13 (Sao kê dư nợ tiền vay)\n\n"
                "1. Nhập Từ ngày và Đến ngày.\n"
                "- Nếu chi hoa hồng theo tháng: từ ngày đầu tháng đến ngày cuối tháng.\n"
                "- Nếu chi hoa hồng theo quý: từ ngày đầu quý đến ngày cuối quý.\n\n"
                "2. Tích chọn:\n"
                "- TC\n"
                "- One Currency\n\n"
                "3. Nhấn Tìm Kiếm.\n\n"
                "Ví dụ quý 2 năm 2026:\n"
                "- Từ ngày: 01/04/2026\n"
                "- Đến ngày: 30/06/2026"
            ),
        )
        self._add_step(
            layout,
            "Bước 03 - Xuất Sao kê dư nợ cuối kỳ",
            (
                "Trong IPCAS:\n"
                "Loan -> lnlr (Báo cáo tín dụng) -> lnlr13 (Sao kê dư nợ tiền vay)\n\n"
                "1. Nhập Từ ngày và Đến ngày theo tháng hoặc quý cần chi hoa hồng.\n\n"
                "2. Tích chọn:\n"
                "- Chưa TT\n"
                "- One Currency\n\n"
                "3. Nhấn Tìm Kiếm.\n\n"
                "Ví dụ quý 2 năm 2026:\n"
                "- Từ ngày: 01/04/2026\n"
                "- Đến ngày: 30/06/2026"
            ),
        )
        self._add_step(
            layout,
            "Bước 04 - Tạo bảng kê thu lãi tổ vay vốn",
            (
                "Sử dụng chức năng:\n"
                "Tổ vay vốn -> Bảng kê thu lãi tổ vay vốn\n\n"
                "- Chọn file Sao kê thu lãi trong kỳ ở Bước 02.\n"
                "- Chọn file Sao kê dư nợ cuối kỳ ở Bước 03.\n"
                "- Nhập kỳ thu lãi từ ngày/đến ngày.\n"
                "- Có thể tạo cho toàn bộ các tổ, một tổ hoặc nhiều tổ.\n"
                "- Nhấn Tạo bảng kê."
            ),
            "File kết quả dùng cho bước lập Đề nghị thanh toán hoa hồng tổ vay vốn. Cần kiểm tra sheet TongHopTheoTo sau khi tạo xong.",
        )
        layout.addStretch()
        return body

    def _payment_request_content(self) -> QWidget:
        body = QWidget()
        layout = QVBoxLayout(body)
        layout.setContentsMargins(0, 12, 0, 12)
        layout.setSpacing(10)
        self._add_step(
            layout,
            "Bước 01 - Chọn file bảng kê thu lãi tổ vay vốn",
            (
                "Sử dụng chức năng:\n"
                "Tổ vay vốn -> Đề nghị thanh toán hoa hồng tổ vay vốn\n\n"
                "- Chọn file Bảng kê thu lãi tổ vay vốn đã tạo.\n"
                "- File bảng kê phải có sheet TongHopTheoTo.\n"
                "- Chọn mẫu Word đề nghị thanh toán nếu cần.\n"
                "- Có thể dùng mẫu mặc định DeNghiThanhToan.docx."
            ),
        )
        self._add_step(
            layout,
            "Bước 02 - Chọn tổ cần tạo đề nghị thanh toán",
            (
                "Có thể lựa chọn:\n"
                "- Xuất đề nghị thanh toán cho toàn bộ các tổ trong sheet TongHopTheoTo.\n"
                "- Xuất cho một tổ.\n"
                "- Xuất cho nhiều tổ.\n\n"
                "Thực hiện:\n"
                "- Nếu xuất toàn bộ tổ, tích Xuất tất cả các tổ trong sheet TongHopTheoTo.\n"
                "- Nếu chỉ xuất một số tổ, bỏ tích tùy chọn trên và chọn các tổ cần xuất.\n"
                "- Nhấn Xuất đề nghị thanh toán.\n\n"
                "Kết quả:\n"
                "- Chương trình tạo file Word đề nghị thanh toán cho từng tổ.\n"
                "- Tên file xuất theo dạng MaTo_NgayTao.docx."
            ),
            (
                "Có thể chỉnh sửa mẫu biểu bằng chức năng: "
                "Tổ vay vốn -> Chỉnh sửa mẫu biểu Đề nghị thanh toán."
            ),
        )
        layout.addStretch()
        return body

    def _debt_reconciliation_content(self) -> QWidget:
        body = QWidget()
        layout = QVBoxLayout(body)
        layout.setContentsMargins(0, 12, 0, 12)
        layout.setSpacing(10)
        self._add_step(
            layout,
            "Bước 01 - Xuất sao kê dư nợ ngày cần đối chiếu",
            (
                "Trong màn hình mssr98:\n\n"
                "1. Chọn báo cáo LN01 - Sao Kê Loan.\n"
                "2. Chọn ngày cần đối chiếu, ví dụ 30/06/2026.\n"
                "3. Nhập mã xác thực.\n"
                "4. Nhấn Xuất báo cáo."
            ),
        )
        self._add_step(
            layout,
            "Bước 02 - Lấy file báo cáo đã xuất",
            (
                "Truy cập máy 21 theo đường dẫn:\n"
                "10.0.43.21/mis/ChiNhanh/cnXXXX/MSSR08/XXXX/\n\n"
                "Trong đó:\n"
                "- cnXXXX là mã chi nhánh.\n"
                "- Thư mục ngày tạo báo cáo có định dạng yyyymmdd, ví dụ 20260630.\n\n"
                "Thực hiện:\n"
                "- Copy file MaCN_ln01_yyyymmdd.zip về máy.\n"
                "- Giải nén file zip.\n"
                "- Password giải nén lấy bằng nút Lấy mật khẩu giải nén trong mssr98.\n\n"
                "Ví dụ file sau giải nén: 5491_ln01_20260630.csv"
            ),
        )
        self._add_step(
            layout,
            "Bước 03 - Tạo đối chiếu tổ vay vốn",
            (
                "Sử dụng chức năng:\n"
                "Tổ vay vốn -> Đối chiếu dư nợ theo tổ vay vốn\n\n"
                "- Chọn file CSV vừa giải nén ở Bước 02.\n"
                "- Chọn ngày đối chiếu.\n"
                "- Có thể đối chiếu toàn bộ các tổ, một tổ hoặc nhiều tổ.\n"
                "- Nhấn Tạo bảng đối chiếu.\n\n"
                "Kết quả:\n"
                "- Chương trình tạo file Excel đối chiếu dư nợ theo tổ vay vốn.\n"
                "- Kiểm tra các sheet cảnh báo nếu có."
            ),
        )
        layout.addStretch()
        return body

    @staticmethod
    def _add_step(
        layout: QVBoxLayout,
        title: str,
        text: str,
        note: str | None = None,
    ) -> None:
        card = QFrame()
        card.setObjectName("GuideStepCard")
        card_layout = QVBoxLayout(card)
        card_layout.setContentsMargins(16, 12, 16, 12)
        card_layout.setSpacing(8)
        title_label = QLabel(title)
        title_label.setObjectName("GuideStepTitle")
        body_label = QLabel(text)
        body_label.setObjectName("GuideStepText")
        body_label.setWordWrap(True)
        body_label.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
        card_layout.addWidget(title_label)
        card_layout.addWidget(body_label)
        if note:
            note_label = QLabel(note)
            note_label.setObjectName("GuideStepNote")
            note_label.setWordWrap(True)
            note_label.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
            card_layout.addWidget(note_label)
        layout.addWidget(card)

    def _apply_style(self) -> None:
        self.setStyleSheet(
            """
            QDialog#ToVayVonHelpWindow {
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
            QLabel#GuideBrandSubtitle {
                color: #f6dce5; font-size: 15px; font-weight: 700;
            }
            QLabel#GuideBrandDescription {
                color: #fbe8ef; font-size: 12px;
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
                padding: 9px 13px; margin-right: 4px;
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
            QLabel#GuideStepNote {
                color: #5b2336; background: #fff4d6;
                border: 1px solid #f2d38b; border-radius: 6px;
                padding: 8px; font-size: 12px;
            }
            QPushButton#GuideCloseButton {
                color: white; background: #931f49; border: none;
                border-radius: 7px; min-width: 78px;
                padding: 10px 18px; font-weight: 650;
            }
            QPushButton#GuideCloseButton:hover { background: #ad2c57; }
            """
        )
