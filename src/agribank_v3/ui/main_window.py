from __future__ import annotations

from dataclasses import replace
from pathlib import Path

from PySide6.QtCore import QSize, Qt, QTimer, QUrl, Signal
from PySide6.QtGui import QCloseEvent, QDesktopServices, QIcon, QMouseEvent, QPixmap
from PySide6.QtWidgets import (
    QFrame,
    QDialog,
    QApplication,
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

from agribank_v3.features.catalog import (
    Feature,
    QUYET_TOAN_KE_TOAN_FEATURES,
    QUYET_TOAN_TIN_DUNG_FEATURES,
    SECTIONS,
)
from agribank_v3.settings import AddinMode, SettingsDatabaseError
from agribank_v3.settlement import (
    SETTLEMENT_SPECS,
    SettlementEngine,
    SettlementError,
    SettlementRequest,
)
from agribank_v3.settlement.processors import (
    Mau04Processor,
    Mau05Processor,
    Mau06Processor,
    Mau0708Processor,
    Mau09Processor,
    Mau1314Processor,
    Mau1516Processor,
    Mau18Processor,
    Mau20aProcessor,
    Mau22Processor,
    Mau23Processor,
    Mau24Processor,
    Mau30Processor,
)
from agribank_v3.excel import (
    ExcelCompatibility,
    ExcelConnectionError,
    ExcelContext,
    ExcelService,
)
from agribank_v3.ui.dialogs.author_info import AuthorInfoDialog
from agribank_v3.ui.dialogs.case_conversion import CaseConversionDialog
from agribank_v3.ui.dialogs.excel_launcher import ExcelLauncherDialog
from agribank_v3.ui.dialogs.settlement_mau1516 import Mau1516SettlementDialog
from agribank_v3.ui.dialogs.settlement_mau05 import Mau05SettlementDialog
from agribank_v3.ui.dialogs.settlement_mau06 import Mau06SettlementDialog
from agribank_v3.ui.dialogs.settlement_mau30 import Mau30SettlementDialog
from agribank_v3.ui.dialogs.settlement_multi_source import MultiSourceSettlementDialog
from agribank_v3.ui.dialogs.settlement_simple_source import SimpleSourceSettlementDialog
from agribank_v3.ui.dialogs.quiz import QuizWidget
from agribank_v3.ui.icons import app_icon, icon_path
from agribank_v3.ui.settings import SettingsWidget
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


class FeatureMenuItem(QFrame):
    requested = Signal(str)

    def __init__(self, feature: Feature, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.feature = feature
        self.setObjectName("FeatureMenuItem")
        self.setCursor(Qt.CursorShape.PointingHandCursor)
        self.setMinimumHeight(82)

        row = QHBoxLayout(self)
        row.setContentsMargins(16, 12, 16, 12)
        row.setSpacing(14)

        icon = QLabel()
        pixmap = QPixmap(icon_path(feature.icon))
        icon.setPixmap(
            pixmap.scaled(
                42,
                42,
                Qt.AspectRatioMode.KeepAspectRatio,
                Qt.TransformationMode.SmoothTransformation,
            )
        )
        icon.setFixedSize(46, 46)

        text_layout = QVBoxLayout()
        text_layout.setSpacing(4)
        title = QLabel(feature.title)
        title.setObjectName("CardTitle")
        description = QLabel(feature.description)
        description.setObjectName("MutedText")
        description.setWordWrap(True)
        text_layout.addWidget(title)
        text_layout.addWidget(description)

        action = QPushButton("Mở")
        action.setObjectName("SecondaryButton")
        action.setCursor(Qt.CursorShape.PointingHandCursor)
        action.clicked.connect(lambda: self.requested.emit(self.feature.title))

        row.addWidget(icon)
        row.addLayout(text_layout, stretch=1)
        row.addWidget(action)

    def mousePressEvent(self, event: QMouseEvent) -> None:
        if event.button() == Qt.MouseButton.LeftButton:
            self.requested.emit(self.feature.title)
        super().mousePressEvent(event)


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
        self.settlement_engine = SettlementEngine()
        self.settlement_engine.register("mau04", Mau04Processor())
        self.settlement_engine.register("mau05", Mau05Processor())
        self.settlement_engine.register("mau06", Mau06Processor())
        self.settlement_engine.register("mau07", Mau0708Processor())
        self.settlement_engine.register("mau08", Mau0708Processor())
        self.settlement_engine.register("mau09", Mau09Processor())
        self.settlement_engine.register("mau13_14", Mau1314Processor())
        self.settlement_engine.register("mau15_16", Mau1516Processor())
        self.settlement_engine.register("mau18", Mau18Processor())
        self.settlement_engine.register("mau20a", Mau20aProcessor())
        self.settlement_engine.register("mau22", Mau22Processor())
        self.settlement_engine.register("mau23", Mau23Processor())
        self.settlement_engine.register("mau24", Mau24Processor())
        self.settlement_engine.register("mau30", Mau30Processor())
        self.excel_context: ExcelContext | None = None
        self.sidebar_expanded = True
        self.quyet_toan_tin_dung_page: QWidget | None = None
        self.quyet_toan_ke_toan_page: QWidget | None = None
        self.author_info_dialog: AuthorInfoDialog | None = None

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

        self.author_info_button = QPushButton("Thông tin tác giả")
        self.author_info_button.setObjectName("SidebarAuthorButton")
        self.author_info_button.setIcon(QIcon(icon_path("inforcn.png")))
        self.author_info_button.setIconSize(QSize(18, 18))
        self.author_info_button.setToolTip("Thông tin ứng dụng và tác giả")
        self.author_info_button.setCursor(Qt.CursorShape.PointingHandCursor)
        self.author_info_button.clicked.connect(self.show_author_info)
        layout.addWidget(self.author_info_button)

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
            if section == "Cài đặt":
                self.settings_widget = SettingsWidget(self)
                self.excel_service.set_addin_mode(
                    self.settings_widget.addin_mode
                )
                self.excel_service.configure_addin_states(
                    self.settings_widget.addin_states
                )
                disabled_addins = tuple(
                    name
                    for name, enabled in self.settings_widget.addin_states.items()
                    if not enabled
                )
                if disabled_addins:
                    self.excel_service.cleanup_tool_addins(disabled_addins)
                if self.settings_widget.addin_mode is AddinMode.SESSION:
                    self.excel_service.cleanup_session_addins()
                self.settings_widget.connect_excel_requested.connect(
                    lambda: self.choose_excel_version(show_after_connect=False)
                )
                self.settings_widget.show_excel_requested.connect(
                    self.show_or_connect_excel
                )
                self.settings_widget.addin_mode_changed.connect(
                    self._apply_addin_mode
                )
                self.settings_widget.addin_enabled_changed.connect(
                    self._apply_addin_enabled
                )
                self.pages.addWidget(self.settings_widget)
            elif section == "Trắc nghiệm":
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

    def show_author_info(self) -> None:
        """Open one author-information dialog and keep it above this window."""
        if (
            self.author_info_dialog is not None
            and self.author_info_dialog.isVisible()
        ):
            self.author_info_dialog.raise_()
            self.author_info_dialog.activateWindow()
            return

        dialog = AuthorInfoDialog(self)
        self.author_info_dialog = dialog
        dialog.finished.connect(self._clear_author_info_dialog)
        dialog.move(self.frameGeometry().center() - dialog.rect().center())
        dialog.open()

    def _clear_author_info_dialog(self) -> None:
        if self.author_info_dialog is not None:
            self.author_info_dialog.deleteLater()
            self.author_info_dialog = None

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

    def _build_quyet_toan_tin_dung_page(self) -> QWidget:
        page = self._scroll_page()
        body = page.widget()
        layout = body.layout()

        header = QHBoxLayout()
        title = QLabel("Quyết toán tín dụng")
        title.setObjectName("PageTitle")
        back_button = QPushButton("Quay lại quyết toán")
        back_button.setObjectName("SecondaryButton")
        back_button.setCursor(Qt.CursorShape.PointingHandCursor)
        back_button.clicked.connect(
            lambda: self.select_page(NAVIGATION.index("Quyết toán"))
        )
        header.addWidget(title)
        header.addStretch()
        header.addWidget(back_button)
        layout.addLayout(header)

        for feature in self._quyet_toan_tin_dung_features():
            item = FeatureMenuItem(feature)
            item.requested.connect(self.open_feature)
            layout.addWidget(item)
        layout.addStretch()
        return page

    def _quyet_toan_tin_dung_features(self) -> list[Feature]:
        branch_code = self._current_branch_code()
        return [
            Feature(
                feature.title.replace("{MaCN}", branch_code),
                feature.description,
                feature.icon,
            )
            for feature in QUYET_TOAN_TIN_DUNG_FEATURES
        ]

    def _current_branch_code(self) -> str:
        try:
            branch_code = self.settings_widget.database.load_branch_profile().branch_code
        except (AttributeError, SettingsDatabaseError):
            branch_code = ""
        return branch_code.strip() or "MaCN"

    def _show_quyet_toan_tin_dung_page(self) -> None:
        if self.quyet_toan_tin_dung_page is not None:
            self.pages.removeWidget(self.quyet_toan_tin_dung_page)
            self.quyet_toan_tin_dung_page.deleteLater()
        self.quyet_toan_tin_dung_page = self._build_quyet_toan_tin_dung_page()
        self.pages.addWidget(self.quyet_toan_tin_dung_page)
        self.pages.setCurrentWidget(self.quyet_toan_tin_dung_page)
        self.nav_buttons[NAVIGATION.index("Quyết toán")].setChecked(True)

    def _build_quyet_toan_ke_toan_page(self) -> QWidget:
        page = self._scroll_page()
        body = page.widget()
        layout = body.layout()

        header = QHBoxLayout()
        title = QLabel("Quyết toán kế toán")
        title.setObjectName("PageTitle")
        back_button = QPushButton("Quay lại quyết toán")
        back_button.setObjectName("SecondaryButton")
        back_button.setCursor(Qt.CursorShape.PointingHandCursor)
        back_button.clicked.connect(
            lambda: self.select_page(NAVIGATION.index("Quyết toán"))
        )
        header.addWidget(title)
        header.addStretch()
        header.addWidget(back_button)
        layout.addLayout(header)

        for feature in self._quyet_toan_ke_toan_features():
            item = FeatureMenuItem(feature)
            item.requested.connect(self._open_quyet_toan_ke_toan_feature)
            layout.addWidget(item)
        layout.addStretch()
        return page

    def _quyet_toan_ke_toan_features(self) -> list[Feature]:
        branch_code = self._current_branch_code()
        return [
            Feature(
                feature.title.replace("{MaCN}", branch_code),
                feature.description,
                feature.icon,
            )
            for feature in QUYET_TOAN_KE_TOAN_FEATURES
        ]

    def _show_quyet_toan_ke_toan_page(self) -> None:
        if self.quyet_toan_ke_toan_page is not None:
            self.pages.removeWidget(self.quyet_toan_ke_toan_page)
            self.quyet_toan_ke_toan_page.deleteLater()
        self.quyet_toan_ke_toan_page = self._build_quyet_toan_ke_toan_page()
        self.pages.addWidget(self.quyet_toan_ke_toan_page)
        self.pages.setCurrentWidget(self.quyet_toan_ke_toan_page)
        self.nav_buttons[NAVIGATION.index("Quyết toán")].setChecked(True)

    def _open_quyet_toan_ke_toan_feature(self, title: str) -> None:
        for spec_key in (
            "accounting.04",
            "accounting.07a",
            "accounting.08",
            "accounting.09a",
            "accounting.09b",
            "accounting.09c",
            "accounting.13",
            "accounting.14",
            "accounting.22",
            "accounting.23",
            "accounting.24",
            "accounting.30a",
        ):
            spec = SETTLEMENT_SPECS[spec_key]
            expected_prefix = f"Tạo Mẫu biểu {spec.report_code}/QT"
            if title.casefold().startswith(expected_prefix.casefold()):
                if spec_key == "accounting.30a":
                    self._run_mau30_dialog(spec_key)
                elif spec_key in self._simple_source_settlement_keys():
                    self._run_simple_source_dialog(spec_key)
                elif spec_key in self._multi_source_settlement_keys():
                    self._run_multi_source_dialog(spec_key)
                else:
                    self._run_mau1516_dialog(spec_key)
                return
        QMessageBox.information(
            self,
            title,
            "Chức năng này đang nằm trong lộ trình chuyển đổi từ VBA sang Python.",
        )

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

    def choose_excel_version(self, show_after_connect: bool = False) -> bool:
        launcher = ExcelLauncherDialog(self.excel_service, self)
        if launcher.exec() != QDialog.DialogCode.Accepted or not launcher.context:
            return False
        self._apply_excel_context(launcher.context)
        if show_after_connect:
            self._bring_excel_to_front()
        return True

    def _apply_excel_context(self, context: ExcelContext) -> None:
        self.excel_context = context
        addin_report = self.excel_service.last_addin_report
        if addin_report.failed:
            failed_details = "; ".join(
                f"{name}: {reason}" for name, reason in addin_report.failed
            )
            addin_status = (
                f"Add-in: không nạp được {failed_details}. Kiểm tra Trust Center, "
                f"quyền ghi thư mục AddIns hoặc tệp trong {addin_report.directory}"
            )
        elif addin_report.discovered:
            active_count = (
                len(addin_report.loaded) + len(addin_report.already_loaded)
            )
            addin_status = (
                f"Add-in: {active_count}/{len(addin_report.discovered)} tệp "
                f"đang hoạt động • {addin_report.directory}"
            )
        else:
            addin_status = (
                f"Chưa có add-in trong {addin_report.directory}. "
                "Chép tệp .xla/.xlam vào đây rồi kết nối lại."
            )
        self.settings_widget.set_addin_status(addin_status)
        self.sidebar_excel_button.setText(
            "XL" if not self.sidebar_expanded else "Hiện Excel"
        )
        if self.excel_service.is_system_worksheet(context):
            tooltip = f"{context.excel_name} • workbook hệ thống (đã ẩn)"
            status = f"Đã kết nối {context.excel_name}"
            settings_status = f"Đã kết nối {context.excel_name}"
        else:
            tooltip = (
                f"{context.excel_name} • {context.workbook} • "
                f"{context.worksheet} • {context.selection}"
            )
            status = (
                f"Đã kết nối {context.excel_name}: {context.workbook} / "
                f"{context.worksheet} / {context.selection}"
            )
            settings_status = (
                f"Đã kết nối: {context.excel_name} • {context.workbook} • "
                f"{context.worksheet} • {context.selection}"
            )
        self.sidebar_excel_button.setToolTip(tooltip)
        self.statusBar().showMessage(status)
        self.settings_widget.set_excel_status(settings_status)

    def _set_excel_disconnected(self) -> None:
        self.excel_context = None
        self.sidebar_excel_button.setText(
            "XL" if not self.sidebar_expanded else "Mở Excel"
        )
        self.sidebar_excel_button.setToolTip("Excel chưa kết nối")
        self.statusBar().showMessage("Excel chưa kết nối")
        self.settings_widget.set_excel_status("Excel chưa kết nối")
        self.settings_widget.set_addin_status(
            f"Thư mục add-in: {self.excel_service.tool_addin_directory()} • "
            f"XLSTART: {self.excel_service.excel_xlstart_directory()}"
        )

    def _apply_addin_mode(self, mode_value: str) -> None:
        mode = AddinMode(mode_value)
        self.excel_service.set_addin_mode(mode)
        if mode is AddinMode.SESSION:
            cleanup = self.excel_service.cleanup_session_addins()
            if cleanup.failed:
                details = "; ".join(
                    f"{name}: {reason}" for name, reason in cleanup.failed
                )
                self.settings_widget.set_addin_status(
                    f"Không gỡ sạch add-in thường trực: {details}"
                )
                return
        elif (
            self.excel_service.is_connected
            and self.excel_service.capabilities is not None
            and self.excel_service.capabilities.major_version <= 14
        ):
            install = self.excel_service.install_tool_addins_to_xlstart()
            if install.failed:
                details = "; ".join(
                    f"{name}: {reason}" for name, reason in install.failed
                )
                self.settings_widget.set_addin_status(
                    f"Không cài được add-in thường trực: {details}"
                )
                return

        if self.excel_service.is_connected:
            report = self.excel_service.load_tool_addins()
            if report.failed:
                details = "; ".join(
                    f"{name}: {reason}" for name, reason in report.failed
                )
                self.settings_widget.set_addin_status(
                    f"Không áp dụng được chế độ add-in: {details}"
                )
            else:
                self.settings_widget.set_addin_status(
                    "Đã áp dụng chế độ add-in "
                    + (
                        "thường trực."
                        if mode is AddinMode.PERMANENT
                        else "chỉ dùng trong phiên AgribankV3."
                    )
                )
        else:
            self.settings_widget.set_addin_status(
                "Đã lưu chế độ add-in "
                + (
                    "thường trực. Thiết lập có hiệu lực khi kết nối Excel."
                    if mode is AddinMode.PERMANENT
                    else "theo phiên. Add-in sẽ được nạp khi kết nối Excel."
                )
            )

    def _apply_addin_enabled(self, file_name: str, enabled: bool) -> None:
        self.excel_service.set_addin_enabled(file_name, enabled)
        if not enabled:
            report = self.excel_service.cleanup_tool_addins((file_name,))
            if report.failed:
                details = "; ".join(
                    f"{name}: {reason}" for name, reason in report.failed
                )
                self.settings_widget.set_addin_status(
                    f"Không gỡ được {file_name}: {details}"
                )
                return
            if self.excel_service.is_connected:
                message = f"Đã gỡ add-in {file_name}."
            else:
                message = (
                    f"Đã xóa {file_name} khỏi XLSTART. Nếu add-in đang được "
                    "Excel ghi nhớ, hãy mở Excel và kết nối AgribankV3 một lần "
                    "để hoàn tất gỡ đăng ký."
                )
            self.settings_widget.set_addin_status(message)
            return

        if not self.excel_service.is_connected:
            if self.excel_service.addin_mode is AddinMode.PERMANENT:
                report = self.excel_service.install_tool_addins_to_xlstart()
                if report.failed:
                    details = "; ".join(
                        f"{name}: {reason}" for name, reason in report.failed
                    )
                    self.settings_widget.set_addin_status(
                        f"Không cài được {file_name}: {details}"
                    )
                    return
                self.settings_widget.set_addin_status(
                    f"Đã cài {file_name} vào XLSTART. Add-in sẽ hoạt động khi "
                    "mở Excel."
                )
            else:
                self.settings_widget.set_addin_status(
                    f"Đã bật {file_name}. Add-in sẽ được nạp khi kết nối Excel."
                )
            return

        if (
            self.excel_service.addin_mode is AddinMode.PERMANENT
            and self.excel_service.capabilities is not None
            and self.excel_service.capabilities.major_version <= 14
        ):
            xlstart_report = self.excel_service.install_tool_addins_to_xlstart()
            if xlstart_report.failed:
                details = "; ".join(
                    f"{name}: {reason}" for name, reason in xlstart_report.failed
                )
                self.settings_widget.set_addin_status(
                    f"Không cài được {file_name}: {details}"
                )
                return
        report = self.excel_service.load_tool_addins()
        matching_failures = tuple(
            failure
            for failure in report.failed
            if failure[0].casefold() == file_name.casefold()
        )
        if matching_failures:
            details = "; ".join(reason for _, reason in matching_failures)
            self.settings_widget.set_addin_status(
                f"Không cài được {file_name}: {details}"
            )
        else:
            self.settings_widget.set_addin_status(
                f"Đã cài và nạp add-in {file_name}."
            )

    def show_or_connect_excel(self, checked: bool = False) -> None:
        del checked
        if not self.choose_excel_version(show_after_connect=False):
            return
        self._bring_excel_to_front()

    def _bring_excel_to_front(self) -> None:
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

    def _open_result_file(self, path: Path) -> None:
        timer_was_active = self.auto_connect_timer.isActive()
        if timer_was_active:
            self.auto_connect_timer.stop()
        QDesktopServices.openUrl(QUrl.fromLocalFile(str(path)))
        if timer_was_active:
            QTimer.singleShot(60_000, self.auto_connect_timer.start)

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
            self.author_info_button.setText("Thông tin tác giả")
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
            self.author_info_button.setText("i")
            for button in self.nav_buttons:
                button.setText("")

    def open_feature(self, title: str) -> None:
        if title in {
            "Kết nối Excel",
            "Thông tin chi nhánh",
            "Cơ sở dữ liệu",
            "Sao lưu dữ liệu",
        }:
            self.select_page(NAVIGATION.index("Cài đặt"))
            self.settings_widget.show_tab_for_feature(title)
            return

        if title == "Kiểm tra nghiệp vụ":
            self.select_page(NAVIGATION.index("Trắc nghiệm"))
            return

        if title == "Quyết toán tín dụng":
            self._show_quyet_toan_tin_dung_page()
            return

        if title == "Quyết toán kế toán":
            self._show_quyet_toan_ke_toan_page()
            return

        if title.startswith("Tạo Mẫu biểu 05/QT ("):
            self._run_mau05_dialog()
            return

        if title.startswith("Tạo Mẫu biểu 06/QT ("):
            self._run_mau06_dialog()
            return

        for spec_key in (
            "credit.15a",
            "credit.15b",
            "credit.16",
            "credit.18",
            "credit.20a",
        ):
            spec = SETTLEMENT_SPECS[spec_key]
            expected_prefix = f"Tạo Mẫu biểu {spec.report_code}/QT ("
            if title.casefold().startswith(expected_prefix.casefold()):
                self._run_mau1516_dialog(spec_key)
                return

        if title.casefold().startswith("tạo mẫu biểu 30a/qt"):
            self._run_mau30_dialog("credit.30a")
            return

        if title == "Chuyển kiểu chữ":
            if not self.connect_excel(show_error=True):
                return
            if (
                self.excel_context is not None
                and self.excel_service.is_system_worksheet(self.excel_context)
            ):
                QMessageBox.information(
                    self,
                    title,
                    "Workbook đang hoạt động là add-in hệ thống. Hãy chọn một "
                    "worksheet trong workbook dữ liệu rồi thử lại.",
                )
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

    def _run_python_settlement(self, spec_key: str) -> None:
        spec = SETTLEMENT_SPECS[spec_key]
        try:
            context = self.excel_service.connect(retry_attempts=1)
        except ExcelConnectionError as exc:
            self._set_excel_disconnected()
            QMessageBox.warning(self, f"Không thể tạo {spec.title}", str(exc))
            return
        self._apply_excel_context(context)

        try:
            profile = self.settings_widget.database.load_branch_profile()
            source_path = self.excel_service.active_workbook_path()
            output_path = source_path.with_name(
                f"{profile.branch_code.strip()}QT{spec.report_code.upper()}.xlsx"
            )
            if output_path.exists():
                answer = QMessageBox.question(
                    self,
                    f"Tạo {spec.title}",
                    f"File {output_path.name} đã tồn tại. Ghi đè file này?",
                    QMessageBox.StandardButton.Yes
                    | QMessageBox.StandardButton.No,
                    QMessageBox.StandardButton.No,
                )
                if answer != QMessageBox.StandardButton.Yes:
                    return
            progress = self._show_busy_dialog(
                f"Đang tạo mẫu quyết toán {spec.report_code}/QT...\n"
                "Vui lòng chờ trong giây lát!"
            )
            try:
                result = self.settlement_engine.execute(
                    SettlementRequest(
                        spec=spec,
                        profile=profile,
                        source_paths=(source_path,),
                    )
                )
                if result.output_path is None:
                    raise SettlementError("Processor không trả về file kết quả.")
                context = self.excel_service.open_workbook(result.output_path)
            finally:
                self._close_busy_dialog(progress)
        except (
            ExcelConnectionError,
            SettingsDatabaseError,
            SettlementError,
            OSError,
        ) as exc:
            QMessageBox.warning(self, f"Không thể tạo {spec.title}", str(exc))
            return
        self._apply_excel_context(context)
        self._bring_excel_to_front()
        QMessageBox.information(
            self,
            f"Hoàn thành {spec.title}",
            f"Đã tạo {result.output_path.name} từ {source_path.name}.",
        )

    def _run_mau05_dialog(self) -> None:
        spec = SETTLEMENT_SPECS["credit.05"]
        try:
            profile = self.settings_widget.database.load_branch_profile()
        except SettingsDatabaseError as exc:
            QMessageBox.warning(self, f"Không thể tạo {spec.title}", str(exc))
            return

        dialog = Mau05SettlementDialog(profile, self)
        if dialog.exec() != QDialog.DialogCode.Accepted:
            return
        source_path = dialog.source_path
        output_path = dialog.output_path()
        if source_path is None or output_path is None:
            return
        processing_options = self._ask_mau05_processing_options(dialog.options())
        if processing_options is None:
            return
        if output_path.exists():
            answer = QMessageBox.question(
                self,
                f"Tạo {spec.title}",
                f"File {output_path.name} đã tồn tại. Ghi đè file này?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No,
            )
            if answer != QMessageBox.StandardButton.Yes:
                return
        progress = self._show_busy_dialog(
            "Đang tạo mẫu quyết toán 05/QT...\n"
            "Vui lòng chờ trong giây lát!"
        )
        execution_error: Exception | None = None
        result = None
        try:
            result = self.settlement_engine.execute(
                SettlementRequest(
                    spec=spec,
                    profile=profile,
                    options=processing_options,
                    source_paths=(source_path,),
                )
            )
        except (SettlementError, OSError) as exc:
            execution_error = exc
        finally:
            self._close_busy_dialog(progress)
        if execution_error is not None:
            QMessageBox.warning(
                self,
                f"Không thể tạo {spec.title}",
                str(execution_error),
            )
            return

        if result is None or result.output_path is None:
            QMessageBox.warning(
                self,
                f"Không thể tạo {spec.title}",
                "Processor không trả về file kết quả.",
            )
            return
        self.statusBar().showMessage(
            f"Đã tạo {result.output_path.name} từ {source_path.name}"
        )
        open_answer = QMessageBox.question(
            self,
            f"Hoàn thành {spec.title}",
            (
                f"Đã tạo file:\n{result.output_path}"
                + (f"\n\n{chr(10).join(result.warnings)}" if result.warnings else "")
                + "\n\nMở file kết quả bây giờ?"
            ),
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.Yes,
        )
        if open_answer == QMessageBox.StandardButton.Yes:
            self._open_result_file(result.output_path)

    def _run_mau06_dialog(self) -> None:
        spec = SETTLEMENT_SPECS["credit.06"]
        try:
            profile = self.settings_widget.database.load_branch_profile()
        except SettingsDatabaseError as exc:
            QMessageBox.warning(self, f"Không thể tạo {spec.title}", str(exc))
            return

        dialog = Mau06SettlementDialog(profile, self)
        if dialog.exec() != QDialog.DialogCode.Accepted:
            return
        source_path = dialog.source_path
        output_path = dialog.output_path()
        if source_path is None or output_path is None:
            return
        if output_path.exists():
            answer = QMessageBox.question(
                self,
                f"Tạo {spec.title}",
                f"File {output_path.name} đã tồn tại. Ghi đè file này?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No,
            )
            if answer != QMessageBox.StandardButton.Yes:
                return
        progress = self._show_busy_dialog(
            "Đang tạo mẫu quyết toán 06/QT...\n"
            "Vui lòng chờ trong giây lát!"
        )
        execution_error: Exception | None = None
        result = None
        try:
            result = self.settlement_engine.execute(
                SettlementRequest(
                    spec=spec,
                    profile=profile,
                    options=dialog.options(),
                    source_paths=(source_path,),
                )
            )
        except (SettlementError, OSError) as exc:
            execution_error = exc
        finally:
            self._close_busy_dialog(progress)
        if execution_error is not None:
            QMessageBox.warning(
                self,
                f"Không thể tạo {spec.title}",
                str(execution_error),
            )
            return

        if result is None or result.output_path is None:
            QMessageBox.warning(
                self,
                f"Không thể tạo {spec.title}",
                "Processor không trả về file kết quả.",
            )
            return
        self.statusBar().showMessage(
            f"Đã tạo {result.output_path.name} từ {source_path.name}"
        )
        open_answer = QMessageBox.question(
            self,
            f"Hoàn thành {spec.title}",
            (
                f"Đã tạo file:\n{result.output_path}"
                + (f"\n\n{chr(10).join(result.warnings)}" if result.warnings else "")
                + "\n\nMở file kết quả bây giờ?"
            ),
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.Yes,
        )
        if open_answer == QMessageBox.StandardButton.Yes:
            self._open_result_file(result.output_path)

    def _run_mau1516_dialog(self, spec_key: str) -> None:
        spec = SETTLEMENT_SPECS[spec_key]
        try:
            profile = self.settings_widget.database.load_branch_profile()
        except SettingsDatabaseError as exc:
            QMessageBox.warning(self, f"Không thể tạo {spec.title}", str(exc))
            return

        dialog = Mau1516SettlementDialog(spec, profile, self)
        if dialog.exec() != QDialog.DialogCode.Accepted:
            return
        source_path = dialog.source_path
        output_path = dialog.output_path()
        if source_path is None or output_path is None:
            return
        if output_path.exists():
            answer = QMessageBox.question(
                self,
                f"Tạo {spec.title}",
                f"File {output_path.name} đã tồn tại. Ghi đè file này?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No,
            )
            if answer != QMessageBox.StandardButton.Yes:
                return
        progress = self._show_busy_dialog(
            f"Đang tạo mẫu quyết toán {spec.report_code}/QT...\n"
            "Vui lòng chờ trong giây lát!"
        )
        execution_error: Exception | None = None
        result = None
        try:
            result = self.settlement_engine.execute(
                SettlementRequest(
                    spec=spec,
                    profile=profile,
                    options=dialog.options(),
                    source_paths=(source_path,),
                )
            )
        except (SettlementError, OSError) as exc:
            execution_error = exc
        finally:
            self._close_busy_dialog(progress)
        if execution_error is not None:
            QMessageBox.warning(
                self,
                f"Không thể tạo {spec.title}",
                str(execution_error),
            )
            return

        if result is None or result.output_path is None:
            QMessageBox.warning(
                self,
                f"Không thể tạo {spec.title}",
                "Processor không trả về file kết quả.",
            )
            return
        self.statusBar().showMessage(
            f"Đã tạo {result.output_path.name} từ {source_path.name}"
        )
        open_answer = QMessageBox.question(
            self,
            f"Hoàn thành {spec.title}",
            (
                f"Đã tạo file:\n{result.output_path}"
                + (f"\n\n{chr(10).join(result.warnings)}" if result.warnings else "")
                + "\n\nMở file kết quả bây giờ?"
            ),
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.Yes,
        )
        if open_answer == QMessageBox.StandardButton.Yes:
            self._open_result_file(result.output_path)

    def _run_simple_source_dialog(self, spec_key: str) -> None:
        spec = SETTLEMENT_SPECS[spec_key]
        try:
            profile = self.settings_widget.database.load_branch_profile()
        except SettingsDatabaseError as exc:
            QMessageBox.warning(self, f"Không thể tạo {spec.title}", str(exc))
            return

        dialog = SimpleSourceSettlementDialog(spec, profile, self)
        if dialog.exec() != QDialog.DialogCode.Accepted:
            return
        source_path = dialog.source_path
        output_path = dialog.output_path()
        if source_path is None or output_path is None:
            return
        if output_path.exists():
            answer = QMessageBox.question(
                self,
                f"Tạo {spec.title}",
                f"File {output_path.name} đã tồn tại. Ghi đè file này?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No,
            )
            if answer != QMessageBox.StandardButton.Yes:
                return
        progress = self._show_busy_dialog(
            f"Đang tạo mẫu quyết toán {spec.report_code}/QT...\n"
            "Vui lòng chờ trong giây lát!"
        )
        execution_error: Exception | None = None
        result = None
        try:
            result = self.settlement_engine.execute(
                SettlementRequest(
                    spec=spec,
                    profile=profile,
                    options=dialog.options(),
                    source_paths=(source_path,),
                )
            )
        except (SettlementError, OSError) as exc:
            execution_error = exc
        finally:
            self._close_busy_dialog(progress)
        if execution_error is not None:
            QMessageBox.warning(
                self,
                f"Không thể tạo {spec.title}",
                str(execution_error),
            )
            return

        if result is None or result.output_path is None:
            QMessageBox.warning(
                self,
                f"Không thể tạo {spec.title}",
                "Processor không trả về file kết quả.",
            )
            return
        self.statusBar().showMessage(
            f"Đã tạo {result.output_path.name} từ {source_path.name}"
        )
        open_answer = QMessageBox.question(
            self,
            f"Hoàn thành {spec.title}",
            (
                f"Đã tạo file:\n{result.output_path}"
                + (f"\n\n{chr(10).join(result.warnings)}" if result.warnings else "")
                + "\n\nMở file kết quả bây giờ?"
            ),
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.Yes,
        )
        if open_answer == QMessageBox.StandardButton.Yes:
            self._open_result_file(result.output_path)

    def _run_multi_source_dialog(self, spec_key: str) -> None:
        spec = SETTLEMENT_SPECS[spec_key]
        try:
            profile = self.settings_widget.database.load_branch_profile()
        except SettingsDatabaseError as exc:
            QMessageBox.warning(self, f"Không thể tạo {spec.title}", str(exc))
            return

        dialog = MultiSourceSettlementDialog(spec, profile, self)
        if dialog.exec() != QDialog.DialogCode.Accepted:
            return
        output_path = dialog.output_path()
        if output_path is None:
            return
        if output_path.exists():
            answer = QMessageBox.question(
                self,
                f"Tạo {spec.title}",
                f"File {output_path.name} đã tồn tại. Ghi đè file này?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No,
            )
            if answer != QMessageBox.StandardButton.Yes:
                return
        progress = self._show_busy_dialog(
            f"Đang tạo mẫu quyết toán {spec.report_code}/QT...\n"
            "Vui lòng chờ trong giây lát!"
        )
        execution_error: Exception | None = None
        result = None
        try:
            result = self.settlement_engine.execute(
                SettlementRequest(
                    spec=spec,
                    profile=profile,
                    options=dialog.options(),
                    source_paths=tuple(dialog.source_paths),
                )
            )
        except (SettlementError, OSError) as exc:
            execution_error = exc
        finally:
            self._close_busy_dialog(progress)
        if execution_error is not None:
            QMessageBox.warning(self, f"Không thể tạo {spec.title}", str(execution_error))
            return
        if result is None or result.output_path is None:
            QMessageBox.warning(self, f"Không thể tạo {spec.title}", "Processor không trả về file kết quả.")
            return
        self.statusBar().showMessage(f"Đã tạo {result.output_path.name}")
        open_answer = QMessageBox.question(
            self,
            f"Hoàn thành {spec.title}",
            (
                f"Đã tạo file:\n{result.output_path}"
                + (f"\n\n{chr(10).join(result.warnings)}" if result.warnings else "")
                + "\n\nMở file kết quả bây giờ?"
            ),
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.Yes,
        )
        if open_answer == QMessageBox.StandardButton.Yes:
            self._open_result_file(result.output_path)

    @staticmethod
    def _simple_source_settlement_keys() -> set[str]:
        return {
            "accounting.04",
            "accounting.07a",
            "accounting.08",
            "accounting.09a",
            "accounting.09b",
            "accounting.09c",
        }

    @staticmethod
    def _multi_source_settlement_keys() -> set[str]:
        return {"accounting.22", "accounting.23"}

    def _run_mau30_dialog(self, spec_key: str = "credit.30a") -> None:
        spec = SETTLEMENT_SPECS[spec_key]
        try:
            profile = self.settings_widget.database.load_branch_profile()
        except SettingsDatabaseError as exc:
            QMessageBox.warning(self, f"Không thể tạo {spec.title}", str(exc))
            return

        last_balance_text = self.settings_widget.database.load_preference(
            "mau30_last_balance_path"
        )
        last_balance_path = (
            Path(last_balance_text) if last_balance_text else None
        )
        dialog = Mau30SettlementDialog(profile, spec, last_balance_path, self)
        if dialog.exec() != QDialog.DialogCode.Accepted:
            return
        source_path = dialog.source_path
        output_path = dialog.output_path()
        if source_path is None or output_path is None:
            return
        progress = self._show_busy_dialog(
            f"Đang tạo mẫu quyết toán 30/QT từ Mẫu {dialog.selected_model}/QT...\n"
            "Vui lòng chờ trong giây lát!"
        )
        execution_error: Exception | None = None
        result = None
        try:
            source_paths = (
                (source_path, dialog.balance_path)
                if dialog.balance_path is not None
                else (source_path,)
            )
            result = self.settlement_engine.execute(
                SettlementRequest(
                    spec=spec,
                    profile=profile,
                    options=dialog.options(),
                    source_paths=source_paths,
                )
            )
        except (SettlementError, OSError) as exc:
            execution_error = exc
        finally:
            self._close_busy_dialog(progress)
        if execution_error is not None:
            QMessageBox.warning(
                self,
                f"Không thể tạo {spec.title}",
                str(execution_error),
            )
            return

        if result is None or result.output_path is None:
            QMessageBox.warning(
                self,
                f"Không thể tạo {spec.title}",
                "Processor không trả về file kết quả.",
            )
            return
        if dialog.balance_path is not None:
            try:
                self.settings_widget.database.save_preference(
                    "mau30_last_balance_path",
                    str(dialog.balance_path),
                )
            except SettingsDatabaseError:
                pass
        self.statusBar().showMessage(
            f"Đã thêm {result.worksheet_name} vào {source_path.name}"
        )
        open_answer = QMessageBox.question(
            self,
            f"Hoàn thành {spec.title}",
            f"Đã thêm sheet {result.worksheet_name} vào file:\n{result.output_path}\n\nMở file bây giờ?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.Yes,
        )
        if open_answer == QMessageBox.StandardButton.Yes:
            self._open_result_file(result.output_path)

    def _show_busy_dialog(self, message: str) -> QDialog:
        progress = QDialog(self)
        progress.setWindowTitle("Đang xử lý")
        progress.setWindowModality(Qt.WindowModality.WindowModal)
        progress.setModal(True)
        progress.setFixedSize(420, 92)
        layout = QVBoxLayout(progress)
        layout.setContentsMargins(18, 10, 18, 10)
        layout.setSpacing(0)
        label = QLabel(message, progress)
        label.setWordWrap(True)
        label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        label.setStyleSheet("font-weight: 600;")
        layout.addWidget(label)
        QApplication.setOverrideCursor(Qt.CursorShape.WaitCursor)
        progress.show()
        progress.raise_()
        progress.activateWindow()
        progress.repaint()
        QApplication.processEvents()
        QApplication.processEvents()
        return progress

    @staticmethod
    def _close_busy_dialog(progress: QDialog) -> None:
        progress.close()
        QApplication.restoreOverrideCursor()
        QApplication.processEvents()

    def _ask_mau05_processing_options(self, base_options):
        guarantee_answer = QMessageBox.question(
            self,
            "Yêu cầu",
            "Đối với TSBĐ là bảo lãnh, bạn muốn chạy sao kê theo "
            "họ và tên chính chủ TSBĐ?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.Yes,
        )
        customer_total_answer = QMessageBox.question(
            self,
            "Agribank",
            "Bạn có muốn cộng tổng theo từng khách hàng không?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        bold_customer_answer = QMessageBox.question(
            self,
            "Agribank",
            "Bạn có muốn tô đậm (Bold) dòng tổng cộng theo từng tên khách hàng không?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        return replace(
            base_options,
            use_collateral_owner_for_guarantee=(
                guarantee_answer == QMessageBox.StandardButton.Yes
            ),
            include_customer_totals=(
                customer_total_answer == QMessageBox.StandardButton.Yes
            ),
            remove_customer_total_rows=(
                customer_total_answer != QMessageBox.StandardButton.Yes
            ),
            bold_customer_rows=(
                bold_customer_answer == QMessageBox.StandardButton.Yes
            ),
        )

    def closeEvent(self, event: QCloseEvent) -> None:
        if self.excel_service.addin_mode is AddinMode.SESSION:
            self.excel_service.cleanup_session_addins()
        self.excel_service.disconnect()
        super().closeEvent(event)
