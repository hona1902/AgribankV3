from __future__ import annotations

from PySide6.QtCore import QSize, Qt, QTimer, Signal
from PySide6.QtGui import QCloseEvent, QIcon, QMouseEvent, QPixmap
from PySide6.QtWidgets import (
    QFrame,
    QDialog,
    QGridLayout,
    QHBoxLayout,
    QLabel,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QScrollArea,
    QSizePolicy,
    QStackedWidget,
    QStatusBar,
    QVBoxLayout,
    QWidget,
)
import win32con
import win32gui
from pywintypes import error as win32_error

from agribank_v3.features.catalog import Feature, SECTIONS
from agribank_v3.excel import (
    ExcelCompatibility,
    ExcelConnectionError,
    ExcelContext,
    ExcelService,
)
from agribank_v3.ui.dialogs.case_conversion import CaseConversionDialog
from agribank_v3.ui.dialogs.excel_launcher import ExcelLauncherDialog
from agribank_v3.ui.dialogs.quiz import QuizWidget
from agribank_v3.ui.icons import app_icon, icon_path
from agribank_v3.quiz import QuizDatabaseError


NAVIGATION = ["Tổng quan", *SECTIONS.keys()]
NAVIGATION_ICONS = (
    "Logo-HNA.png",
    "caidat.png",
    "case.png",
    "access.png",
    "m09a.png",
    "qtkt.png",
    "qt.png",
    "tracnghiem.png",
    "BamChuot.png",
)


class FeatureCard(QFrame):
    requested = Signal(str)

    def __init__(self, feature: Feature, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.feature = feature
        self.setObjectName("FeatureCard")
        self.setMinimumHeight(164)
        self.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(18, 17, 18, 16)
        layout.setSpacing(8)

        icon = QLabel()
        pixmap = QPixmap(icon_path(feature.icon))
        icon.setPixmap(
            pixmap.scaled(
                36,
                36,
                Qt.AspectRatioMode.KeepAspectRatio,
                Qt.TransformationMode.SmoothTransformation,
            )
        )
        icon.setFixedHeight(38)

        title = QLabel(feature.title)
        title.setObjectName("CardTitle")

        description = QLabel(feature.description)
        description.setObjectName("MutedText")
        description.setWordWrap(True)

        action = QPushButton("Mở chức năng")
        action.setObjectName("SecondaryButton")
        action.setCursor(Qt.CursorShape.PointingHandCursor)
        action.clicked.connect(lambda: self.requested.emit(self.feature.title))

        layout.addWidget(icon)
        layout.addWidget(title)
        layout.addWidget(description)
        layout.addStretch()
        layout.addWidget(action, alignment=Qt.AlignmentFlag.AlignLeft)


class ClickableBrand(QFrame):
    clicked = Signal()

    def mousePressEvent(self, event: QMouseEvent) -> None:
        if event.button() == Qt.MouseButton.LeftButton:
            self.clicked.emit()
        super().mousePressEvent(event)


class MainWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("AgribankV3")
        self.setWindowIcon(app_icon())
        self.setMinimumSize(1120, 700)
        self.resize(1360, 820)

        self.nav_buttons: list[QPushButton] = []
        self.excel_service = ExcelService()
        self.excel_context: ExcelContext | None = None
        self.sidebar_expanded = True

        root = QWidget()
        root.setObjectName("AppRoot")
        root_layout = QHBoxLayout(root)
        root_layout.setContentsMargins(0, 0, 0, 0)
        root_layout.setSpacing(0)

        root_layout.addWidget(self._build_sidebar())
        root_layout.addWidget(self._build_content(), stretch=1)
        self.setCentralWidget(root)

        status = QStatusBar()
        status.showMessage("AgribankV3 v0.1.0  •  Prototype giao diện")
        self.setStatusBar(status)
        self.select_page(0)
        self.auto_connect_timer = QTimer(self)
        self.auto_connect_timer.setInterval(3000)
        self.auto_connect_timer.timeout.connect(self.auto_detect_excel)
        self.auto_connect_timer.start()
        QTimer.singleShot(500, self.auto_detect_excel)

    def _build_sidebar(self) -> QWidget:
        self.sidebar = QFrame()
        self.sidebar.setObjectName("Sidebar")
        self.sidebar.setFixedWidth(224)
        layout = QVBoxLayout(self.sidebar)
        layout.setContentsMargins(16, 24, 16, 20)
        layout.setSpacing(5)

        self.brand_button = ClickableBrand()
        self.brand_button.setObjectName("BrandButton")
        self.brand_button.setCursor(Qt.CursorShape.PointingHandCursor)
        self.brand_button.setToolTip("Nhấn để thu gọn menu")
        self.brand_button.clicked.connect(self.toggle_sidebar)
        brand_row = QHBoxLayout(self.brand_button)
        brand_row.setContentsMargins(0, 0, 0, 0)
        self.brand_logo = QLabel()
        self.brand_logo.setAttribute(
            Qt.WidgetAttribute.WA_TransparentForMouseEvents,
            True,
        )
        pixmap = QPixmap(icon_path("logoagri.png"))
        self.brand_logo.setPixmap(
            pixmap.scaled(
                42,
                42,
                Qt.AspectRatioMode.KeepAspectRatio,
                Qt.TransformationMode.SmoothTransformation,
            )
        )
        brand_text = QVBoxLayout()
        brand_text.setSpacing(0)
        self.brand_title = QLabel("AgribankV3")
        self.brand_title.setObjectName("BrandTitle")
        self.brand_title.setAttribute(
            Qt.WidgetAttribute.WA_TransparentForMouseEvents,
            True,
        )
        self.brand_subtitle = QLabel("Công cụ Excel nghiệp vụ")
        self.brand_subtitle.setObjectName("BrandSubtitle")
        self.brand_subtitle.setAttribute(
            Qt.WidgetAttribute.WA_TransparentForMouseEvents,
            True,
        )
        brand_text.addWidget(self.brand_title)
        brand_text.addWidget(self.brand_subtitle)
        brand_row.addWidget(self.brand_logo)
        brand_row.addLayout(brand_text)
        brand_row.addStretch()
        layout.addWidget(self.brand_button)
        layout.addSpacing(8)

        for index, name in enumerate(NAVIGATION):
            button = QPushButton(name)
            button.setObjectName("NavButton")
            button.setIcon(QIcon(icon_path(NAVIGATION_ICONS[index])))
            button.setIconSize(QSize(22, 22))
            button.setCheckable(True)
            button.setAutoExclusive(True)
            button.setCursor(Qt.CursorShape.PointingHandCursor)
            button.setToolTip(name)
            button.setProperty("fullText", name)
            button.clicked.connect(lambda checked=False, i=index: self.select_page(i))
            layout.addWidget(button)
            self.nav_buttons.append(button)

        layout.addStretch()
        self.sidebar_excel_button = QPushButton("Mở Excel")
        self.sidebar_excel_button.setObjectName("SidebarExcelButton")
        self.sidebar_excel_button.setIcon(QIcon(icon_path("button_excel.svg")))
        self.sidebar_excel_button.setIconSize(QSize(22, 22))
        self.sidebar_excel_button.setToolTip("Mở, kết nối hoặc hiện cửa sổ Excel")
        self.sidebar_excel_button.setCursor(Qt.CursorShape.PointingHandCursor)
        self.sidebar_excel_button.clicked.connect(self.show_or_connect_excel)
        layout.addWidget(self.sidebar_excel_button)

        self.version_label = QLabel("Phiên bản thử nghiệm 0.1.0")
        self.version_label.setObjectName("BrandSubtitle")
        self.version_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(self.version_label)
        return self.sidebar

    def _build_content(self) -> QWidget:
        content = QWidget()
        layout = QVBoxLayout(content)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        self.pages = QStackedWidget()
        self.pages.addWidget(self._build_dashboard())
        for section, features in SECTIONS.items():
            if section == "Trắc nghiệm":
                try:
                    self.quiz_widget = QuizWidget(self)
                    self.quiz_widget.close_requested.connect(
                        lambda: self.select_page(0)
                    )
                    self.pages.addWidget(self.quiz_widget)
                except QuizDatabaseError as exc:
                    error_page = QWidget()
                    error_layout = QVBoxLayout(error_page)
                    error_layout.setContentsMargins(36, 30, 36, 30)
                    error_title = QLabel("Không thể mở dữ liệu trắc nghiệm")
                    error_title.setObjectName("PageTitle")
                    error_text = QLabel(str(exc))
                    error_text.setWordWrap(True)
                    error_layout.addWidget(error_title)
                    error_layout.addWidget(error_text)
                    error_layout.addStretch()
                    self.pages.addWidget(error_page)
            else:
                self.pages.addWidget(self._build_feature_page(section, features))
        layout.addWidget(self.pages, stretch=1)
        return content

    def _build_dashboard(self) -> QWidget:
        page = self._scroll_page()
        body = page.widget()
        layout = body.layout()

        welcome = QFrame()
        welcome.setObjectName("WelcomeCard")
        welcome_layout = QHBoxLayout(welcome)
        welcome_layout.setContentsMargins(24, 20, 24, 20)
        text_layout = QVBoxLayout()
        heading = QLabel("Trung tâm công cụ AgribankV3")
        heading.setObjectName("PageTitle")
        description = QLabel(
            "Bản thử giao diện mới, tổ chức lại toàn bộ chức năng của add-in "
            "AgribankV2 theo từng nhóm nghiệp vụ."
        )
        description.setObjectName("MutedText")
        description.setWordWrap(True)
        text_layout.addWidget(heading)
        text_layout.addWidget(description)
        welcome_layout.addLayout(text_layout, stretch=1)
        layout.addWidget(welcome)

        metrics = QHBoxLayout()
        for value, label in (
            ("217", "Ribbon callback"),
            ("8", "Nhóm chức năng"),
            ("53", "Biểu mẫu sẽ chuyển đổi"),
            ("13", "Add-in và workbook"),
        ):
            metrics.addWidget(self._metric_card(value, label))
        layout.addLayout(metrics)

        section_title = QLabel("Truy cập nhanh")
        section_title.setObjectName("PageTitle")
        layout.addWidget(section_title)

        grid = QGridLayout()
        grid.setHorizontalSpacing(14)
        grid.setVerticalSpacing(14)
        quick_features = [
            SECTIONS["Chức năng"][0],
            SECTIONS["Dữ liệu"][1],
            SECTIONS["Tín dụng"][0],
            SECTIONS["Quyết toán"][0],
        ]
        for index, feature in enumerate(quick_features):
            card = FeatureCard(feature)
            card.requested.connect(self.open_feature)
            grid.addWidget(card, index // 2, index % 2)
        layout.addLayout(grid)
        layout.addStretch()
        return page

    def _build_feature_page(self, title: str, features: list[Feature]) -> QWidget:
        page = self._scroll_page()
        body = page.widget()
        layout = body.layout()

        intro = QLabel(
            f"{len(features)} chức năng đại diện trong nhóm {title}. "
            "Danh mục sẽ được bổ sung theo ma trận chức năng VBA."
        )
        intro.setObjectName("MutedText")
        intro.setWordWrap(True)
        layout.addWidget(intro)

        grid = QGridLayout()
        grid.setHorizontalSpacing(14)
        grid.setVerticalSpacing(14)
        for index, feature in enumerate(features):
            card = FeatureCard(feature)
            card.requested.connect(self.open_feature)
            grid.addWidget(card, index // 3, index % 3)
        for column in range(3):
            grid.setColumnStretch(column, 1)
        layout.addLayout(grid)
        layout.addStretch()
        return page

    @staticmethod
    def _scroll_page() -> QScrollArea:
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.Shape.NoFrame)
        body = QWidget()
        layout = QVBoxLayout(body)
        layout.setContentsMargins(28, 24, 28, 28)
        layout.setSpacing(18)
        scroll.setWidget(body)
        return scroll

    @staticmethod
    def _metric_card(value: str, label: str) -> QWidget:
        card = QFrame()
        card.setObjectName("MetricCard")
        card.setMinimumHeight(98)
        layout = QVBoxLayout(card)
        layout.setContentsMargins(18, 12, 18, 12)
        metric = QLabel(value)
        metric.setObjectName("MetricValue")
        caption = QLabel(label)
        caption.setObjectName("MutedText")
        layout.addWidget(metric)
        layout.addWidget(caption)
        return card

    def select_page(self, index: int) -> None:
        self.pages.setCurrentIndex(index)
        self.nav_buttons[index].setChecked(True)

    def connect_excel(self, show_error: bool = True) -> bool:
        try:
            self.excel_context = self.excel_service.connect()
        except ExcelConnectionError as exc:
            self._set_excel_disconnected()
            if exc.code == "not_running" and show_error:
                launcher = ExcelLauncherDialog(self.excel_service, self)
                if launcher.exec() == QDialog.DialogCode.Accepted and launcher.context:
                    self._apply_excel_context(launcher.context)
                    return True
                return False
            if exc.code == "no_workbook":
                try:
                    context = self.excel_service.connect(
                        retry_attempts=1,
                        create_workbook_if_missing=True,
                    )
                    self._apply_excel_context(context)
                    return True
                except ExcelConnectionError as create_error:
                    exc = create_error
            if show_error:
                QMessageBox.warning(self, "Không thể kết nối Excel", str(exc))
            return False

        self._apply_excel_context(self.excel_context)
        return True

    def _apply_excel_context(self, context: ExcelContext) -> None:
        self.excel_context = context
        self.sidebar_excel_button.setText(
            "XL" if not self.sidebar_expanded else "Hiện Excel"
        )
        self.sidebar_excel_button.setToolTip(
            f"{context.excel_name} • {context.workbook} • "
            f"{context.worksheet} • {context.selection}"
        )
        self.statusBar().showMessage(
            f"Đã kết nối {context.excel_name}: {context.workbook} / "
            f"{context.worksheet} / {context.selection}"
        )

    def _set_excel_disconnected(self) -> None:
        self.excel_context = None
        self.sidebar_excel_button.setText(
            "XL" if not self.sidebar_expanded else "Mở Excel"
        )
        self.sidebar_excel_button.setToolTip("Excel chưa kết nối")
        self.statusBar().showMessage("Excel chưa kết nối")

    def show_or_connect_excel(self, checked: bool = False) -> None:
        del checked
        if not self.connect_excel(show_error=True):
            return
        application = self.excel_service.application
        context = self.excel_context
        if application is None or context is None:
            return
        try:
            application.Visible = True
            hwnd = ExcelCompatibility.window_handle(application, context.workbook)
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
            win32gui.SetForegroundWindow(hwnd)
        except (AttributeError, win32_error):
            QMessageBox.warning(
                self,
                "Không thể hiện Excel",
                "Không thể đưa cửa sổ Excel ra phía trước. Hãy chọn Excel trên "
                "thanh tác vụ Windows.",
            )

    def auto_detect_excel(self) -> None:
        try:
            if self.excel_service.is_connected:
                context = self.excel_service.get_context()
            else:
                context = self.excel_service.connect(retry_attempts=1)
        except ExcelConnectionError:
            if self.excel_context is not None:
                self._set_excel_disconnected()
            return
        self._apply_excel_context(context)

    def toggle_sidebar(self) -> None:
        self.sidebar_expanded = not self.sidebar_expanded
        if self.sidebar_expanded:
            self.sidebar.setFixedWidth(224)
            self.sidebar.layout().setContentsMargins(16, 24, 16, 20)
            self.brand_title.show()
            self.brand_subtitle.show()
            self.version_label.show()
            self.brand_button.setToolTip("Nhấn để thu gọn menu")
            self.sidebar_excel_button.setText(
                "Hiện Excel" if self.excel_context else "Mở Excel"
            )
            for button in self.nav_buttons:
                button.setText(str(button.property("fullText")))
        else:
            self.sidebar.setFixedWidth(72)
            self.sidebar.layout().setContentsMargins(8, 24, 8, 20)
            self.brand_title.hide()
            self.brand_subtitle.hide()
            self.version_label.hide()
            self.brand_button.setToolTip("Nhấn để mở rộng menu")
            self.sidebar_excel_button.setText("XL")
            for button in self.nav_buttons:
                button.setText("")

    def open_feature(self, title: str) -> None:
        if title == "Kiểm tra nghiệp vụ":
            self.select_page(NAVIGATION.index("Trắc nghiệm"))
            return

        if title == "Chuyển kiểu chữ":
            if not self.connect_excel(show_error=True):
                return
            dialog = CaseConversionDialog(
                self.excel_service,
                self.excel_context,
                self,
            )
            dialog.exec()
            self.connect_excel(show_error=False)
            return

        QMessageBox.information(
            self,
            title,
            "Chức năng này đang nằm trong lộ trình chuyển đổi từ VBA sang Python.",
        )

    def closeEvent(self, event: QCloseEvent) -> None:
        self.excel_service.disconnect()
        super().closeEvent(event)
