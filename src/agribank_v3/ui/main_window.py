from __future__ import annotations

from dataclasses import asdict, replace
from html import escape
import json
from pathlib import Path
import sys

from PySide6.QtCore import QEvent, QProcess, QRect, QSize, Qt, QTimer, Signal
from PySide6.QtGui import QCloseEvent, QFontMetrics, QIcon, QMouseEvent, QPixmap
from PySide6.QtWidgets import (
    QCheckBox,
    QFrame,
    QDialog,
    QApplication,
    QGridLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
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

from agribank_v3 import __version__
from agribank_v3.features.catalog import (
    Feature,
    FEATURE_GROUPS,
    QUYET_TOAN_KE_TOAN_FEATURES,
    QUICK_ACCESS_DEFAULT_IDS,
    QUICK_ACCESS_FEATURES,
    QuickAccessFeature,
    QUYET_TOAN_TIN_DUNG_FEATURES,
    QUYET_TOAN_TONG_HOP_FEATURES,
    SECTIONS,
)
from agribank_v3.settings import AddinMode, AppSettingsDatabase, SettingsDatabaseError
from agribank_v3.settlement import (
    SETTLEMENT_SPECS,
    SettlementEngine,
    SettlementError,
    SettlementRequest,
    SettlementResult,
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
from agribank_v3.settlement.processors.summary05 import Summary05Processor
from agribank_v3.excel import (
    ExcelCompatibility,
    ExcelConnectionError,
    ExcelContext,
    ExcelService,
)
from agribank_v3.file_merge import (
    FileMergeError,
    merge_same_structure_csv_to_csv,
    merge_same_structure_csv_to_xlsx,
    merge_same_structure_excel_to_xlsx,
)
from agribank_v3.excel_tools import (
    ExcelToolError,
    convert_csv_to_excel,
    split_workbook_sheets_to_files,
)
from agribank_v3.features.credit.auto_interest.placeholder_windows import (
    AUTO_INTEREST_PLACEHOLDER_TITLE,
    AutoInterestPlaceholderDialog,
)
from agribank_v3.features.credit.tovayvon.menu import TOVAYVON_FEATURES
from agribank_v3.features.credit.tovayvon.placeholder_windows import (
    CREDIT_GROUP_MANAGEMENT_ROUTE_TITLES,
    CREDIT_TOVAYVON_PLACEHOLDER_TITLES,
    CreditGroupManagementPlaceholderDialog,
    CreditMigrationPlaceholderDialog,
)
from agribank_v3.word_folder_print import print_word_files
from agribank_v3.ui.dialogs.author_info import AuthorInfoDialog
from agribank_v3.ui.dialogs.case_conversion import CaseConversionDialog
from agribank_v3.ui.dialogs.consolidation_csv import ConsolidationCsvDialog
from agribank_v3.ui.dialogs.excel_file_tools import CsvToExcelDialog, SplitSheetsDialog
from agribank_v3.ui.dialogs.excel_launcher import ExcelLauncherDialog
from agribank_v3.ui.dialogs.settlement_mau1516 import Mau1516SettlementDialog
from agribank_v3.ui.dialogs.settlement_mau05 import Mau05SettlementDialog
from agribank_v3.ui.dialogs.settlement_mau06 import Mau06SettlementDialog
from agribank_v3.ui.dialogs.settlement_mau30 import Mau30SettlementDialog
from agribank_v3.ui.dialogs.settlement_guidance import SettlementGuidanceDialog
from agribank_v3.ui.dialogs.settlement_multi_source import MultiSourceSettlementDialog
from agribank_v3.ui.dialogs.settlement_simple_source import SimpleSourceSettlementDialog
from agribank_v3.ui.dialogs.printer_settings import PrinterSettingsDialog
from agribank_v3.ui.dialogs.same_structure_merge import SameStructureMergeDialog
from agribank_v3.ui.dialogs.word_folder_print import WordFolderPrintDialog
from agribank_v3.ui.dialogs.quiz import QuizWidget
from agribank_v3.ui.icons import app_icon, icon_path
from agribank_v3.ui.settings import SettingsWidget
from agribank_v3.ui.workers import run_in_thread
from agribank_v3.quiz import QuizDatabaseError


APP_FOOTER_TEXT = (
    f"AgribankV3 - Phiên bản v{__version__} - Nguyễn Hoài Nam - 0972.173.064"
)
FEATURE_CARD_GAP = 12
NAVIGATION = [
    "Tổng quan",
    "Chức năng",
    "Dữ liệu",
    "Tín dụng",
    "Kế toán",
    "Quyết toán",
    "Trắc nghiệm",
    "Cài đặt",
]
NAVIGATION_ICONS = (
    "Logo-HNA.png",
    "case.png",
    "access.png",
    "m09a.png",
    "qtkt.png",
    "qt.png",
    "tracnghiem.svg",
    "caidat.png",
)


class AutoFitLabel(QLabel):
    def __init__(
        self,
        text: str,
        *,
        max_pixel_size: int,
        min_pixel_size: int,
        max_lines: int,
        fixed_height: int,
        parent: QWidget | None = None,
    ) -> None:
        super().__init__(parent)
        self._full_text = text
        self._max_pixel_size = max_pixel_size
        self._min_pixel_size = min_pixel_size
        self._max_lines = max_lines
        self._fitting = False
        self.setWordWrap(True)
        self.setFixedHeight(fixed_height)
        self.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        self.setText(text)

    def resizeEvent(self, event) -> None:
        super().resizeEvent(event)
        self.fit_text_to_card()

    def showEvent(self, event) -> None:
        super().showEvent(event)
        self.fit_text_to_card()

    def fit_text_to_card(self) -> None:
        if self._fitting or self.width() <= 0:
            return
        self._fitting = True
        try:
            available_width = max(10, self.width())
            available_height = max(10, self.height())
            text = self._full_text
            chosen_font = self.font()
            for pixel_size in range(self._max_pixel_size, self._min_pixel_size - 1, -1):
                candidate_font = self.font()
                candidate_font.setPixelSize(pixel_size)
                metrics = QFontMetrics(candidate_font)
                max_height = min(available_height, metrics.lineSpacing() * self._max_lines)
                text_rect = metrics.boundingRect(
                    QRect(0, 0, available_width, max_height * 3),
                    Qt.TextFlag.TextWordWrap,
                    text,
                )
                if text_rect.height() <= max_height and text_rect.width() <= available_width:
                    chosen_font = candidate_font
                    self.setFont(chosen_font)
                    QLabel.setText(self, text)
                    return
                chosen_font = candidate_font
            self.setFont(chosen_font)
            QLabel.setText(
                self,
                self._elided_multiline_text(
                    text,
                    available_width,
                    available_height,
                    chosen_font,
                ),
            )
        finally:
            self._fitting = False

    def _elided_multiline_text(
        self,
        text: str,
        available_width: int,
        available_height: int,
        font,
    ) -> str:
        metrics = QFontMetrics(font)
        max_height = min(available_height, metrics.lineSpacing() * self._max_lines)
        if not text:
            return ""
        low = 0
        high = len(text)
        best = "..."
        while low <= high:
            middle = (low + high) // 2
            candidate = text[:middle].rstrip()
            if middle < len(text):
                candidate = f"{candidate}..."
            text_rect = metrics.boundingRect(
                QRect(0, 0, available_width, max_height * 3),
                Qt.TextFlag.TextWordWrap,
                candidate,
            )
            if text_rect.height() <= max_height and text_rect.width() <= available_width:
                best = candidate
                low = middle + 1
            else:
                high = middle - 1
        return best


class FeatureCard(QFrame):
    requested = Signal(str)
    CARD_WIDTH = 260
    CARD_HEIGHT = 126
    ICON_SIZE = 30
    TITLE_HEIGHT = 44
    ACTION_HEIGHT = 30

    def __init__(self, feature: Feature, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.feature = feature
        self.setObjectName("FeatureCard")
        self.setCursor(Qt.CursorShape.PointingHandCursor)
        self.setToolTip(self._tooltip_html(feature.description))
        self.setFixedSize(self.CARD_WIDTH, self.CARD_HEIGHT)
        self.setSizePolicy(QSizePolicy.Policy.Fixed, QSizePolicy.Policy.Fixed)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(14, 11, 14, 10)
        layout.setSpacing(5)

        icon = QLabel()
        pixmap = QPixmap(icon_path(feature.icon))
        icon.setPixmap(
            pixmap.scaled(
                self.ICON_SIZE,
                self.ICON_SIZE,
                Qt.AspectRatioMode.KeepAspectRatio,
                Qt.TransformationMode.SmoothTransformation,
            )
        )
        icon.setFixedSize(self.ICON_SIZE, self.ICON_SIZE)
        icon.setToolTip(self.toolTip())

        title = AutoFitLabel(
            feature.title,
            max_pixel_size=15,
            min_pixel_size=11,
            max_lines=2,
            fixed_height=self.TITLE_HEIGHT,
        )
        title.setObjectName("AutoCardTitle")
        title.setToolTip(self.toolTip())

        layout.addWidget(icon)
        layout.addWidget(title)
        layout.addStretch()
        action_area = QWidget()
        action_area.setFixedHeight(self.ACTION_HEIGHT)
        action_layout = QHBoxLayout(action_area)
        action_layout.setContentsMargins(0, 0, 0, 0)
        action_layout.setSpacing(0)
        action_layout.addStretch()
        arrow = QLabel(">")
        arrow.setObjectName("CardArrow")
        arrow.setToolTip(self.toolTip())
        arrow.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        arrow.setFixedSize(22, self.ACTION_HEIGHT)
        action_layout.addWidget(arrow)
        layout.addWidget(action_area)

    def mousePressEvent(self, event: QMouseEvent) -> None:
        if event.button() == Qt.MouseButton.LeftButton:
            self.requested.emit(self.feature.title)
        super().mousePressEvent(event)

    @staticmethod
    def _tooltip_html(text: str) -> str:
        return (
            "<div style='white-space: normal; width: 320px;'>"
            f"{escape(text)}"
            "</div>"
        )


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

        row.addWidget(icon)
        row.addLayout(text_layout, stretch=1)

    def mousePressEvent(self, event: QMouseEvent) -> None:
        if event.button() == Qt.MouseButton.LeftButton:
            self.requested.emit(self.feature.title)
        super().mousePressEvent(event)


class ResponsiveFeatureGrid(QWidget):
    def __init__(
        self,
        features: tuple[Feature, ...] | list[Feature],
        parent: QWidget | None = None,
    ) -> None:
        super().__init__(parent)
        self.cards = tuple(FeatureCard(feature) for feature in features)
        self._columns = 0
        self._viewport = None
        self.grid = QGridLayout(self)
        self.grid.setContentsMargins(0, 0, 0, 0)
        self.grid.setHorizontalSpacing(FEATURE_CARD_GAP)
        self.grid.setVerticalSpacing(FEATURE_CARD_GAP)
        self.grid.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignTop)
        self.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        self._relayout_cards(force=True)

    def showEvent(self, event) -> None:
        super().showEvent(event)
        self._install_viewport_resize_filter()
        self._relayout_cards(force=True)

    def resizeEvent(self, event) -> None:
        super().resizeEvent(event)
        self._relayout_cards()

    def eventFilter(self, watched, event) -> bool:
        if watched is self._viewport and event.type() == QEvent.Type.Resize:
            self._relayout_cards(force=True)
        return super().eventFilter(watched, event)

    def connect_requested(self, slot) -> None:
        for card in self.cards:
            card.requested.connect(slot)

    def _column_count_for_width(self, available_width: int) -> int:
        card_span = FeatureCard.CARD_WIDTH + FEATURE_CARD_GAP
        return max(1, (max(1, available_width) + FEATURE_CARD_GAP) // card_span)

    def _relayout_cards(self, force: bool = False) -> None:
        columns = self._column_count_for_width(self._available_layout_width())
        if not force and columns == self._columns:
            return
        self._columns = columns
        for card in self.cards:
            self.grid.removeWidget(card)
        for index, card in enumerate(self.cards):
            self.grid.addWidget(
                card,
                index // columns,
                index % columns,
                alignment=Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignTop,
            )
        self.updateGeometry()

    def _available_layout_width(self) -> int:
        available_width = self.width()
        scroll_area = self._ancestor_scroll_area()
        if scroll_area is not None:
            viewport_width = scroll_area.viewport().width()
            body = scroll_area.widget()
            if body is not None and body.layout() is not None:
                margins = body.layout().contentsMargins()
                viewport_width -= margins.left() + margins.right()
            if viewport_width > 0:
                available_width = (
                    min(available_width, viewport_width)
                    if available_width > 0
                    else viewport_width
                )
        return max(FeatureCard.CARD_WIDTH, available_width)

    def _install_viewport_resize_filter(self) -> None:
        scroll_area = self._ancestor_scroll_area()
        viewport = scroll_area.viewport() if scroll_area is not None else None
        if viewport is self._viewport:
            return
        if self._viewport is not None:
            self._viewport.removeEventFilter(self)
        self._viewport = viewport
        if self._viewport is not None:
            self._viewport.installEventFilter(self)

    def _ancestor_scroll_area(self) -> QScrollArea | None:
        parent = self.parentWidget()
        while parent is not None:
            if isinstance(parent, QScrollArea):
                return parent
            parent = parent.parentWidget()
        return None


class QuickAccessSettingsDialog(QDialog):
    MAX_ITEMS = 12
    ROW_HEIGHT = 48
    CHECKBOX_COLUMN_WIDTH = 34
    ICON_COLUMN_WIDTH = 34
    GROUP_COLUMN_WIDTH = 170

    def __init__(
        self,
        candidates: tuple[QuickAccessFeature, ...],
        selected_ids: tuple[str, ...],
        parent: QWidget | None = None,
    ) -> None:
        super().__init__(parent)
        self.candidates = candidates
        self.selected_ids = selected_ids
        self.setWindowTitle("Cài đặt truy cập nhanh")
        self.setModal(True)
        self.setMinimumSize(620, 560)
        self.saved_quick_access_ids = tuple(selected_ids)
        self.temp_quick_access_ids = tuple(selected_ids)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(18, 16, 18, 16)
        layout.setSpacing(10)

        header = QFrame()
        header.setObjectName("QuickAccessDialogHeader")
        header_layout = QVBoxLayout(header)
        header_layout.setContentsMargins(16, 14, 16, 14)
        header_layout.setSpacing(4)
        title = QLabel("Cài đặt truy cập nhanh")
        title.setObjectName("DialogHeaderTitle")
        subtitle = QLabel(
            "Chọn các chức năng thường dùng để hiển thị tại màn hình Tổng quan."
        )
        subtitle.setObjectName("DialogHeaderSubtitle")
        subtitle.setWordWrap(True)
        header_layout.addWidget(title)
        header_layout.addWidget(subtitle)

        self.search_edit = QLineEdit()
        self.search_edit.setPlaceholderText("Tìm chức năng...")
        self.search_edit.textChanged.connect(self._filter_items)

        self.list_widget = QListWidget()
        self.list_widget.setHorizontalScrollBarPolicy(
            Qt.ScrollBarPolicy.ScrollBarAlwaysOff
        )
        self.list_widget.setUniformItemSizes(True)
        self.list_widget.itemClicked.connect(self._toggle_item_checkbox)

        move_row = QHBoxLayout()
        up_button = QPushButton("Lên")
        up_button.clicked.connect(lambda checked=False: self._move_current(-1))
        down_button = QPushButton("Xuống")
        down_button.clicked.connect(lambda checked=False: self._move_current(1))
        default_button = QPushButton("Khôi phục mặc định")
        default_button.clicked.connect(self.restore_defaults)
        move_row.addWidget(up_button)
        move_row.addWidget(down_button)
        move_row.addStretch()
        move_row.addWidget(default_button)

        button_row = QHBoxLayout()
        save_button = QPushButton("Lưu")
        save_button.setObjectName("PrimaryButton")
        save_button.clicked.connect(self.accept)
        cancel_button = QPushButton("Hủy")
        cancel_button.clicked.connect(self.reject)
        button_row.addStretch()
        button_row.addWidget(save_button)
        button_row.addWidget(cancel_button)

        layout.addWidget(header)
        layout.addWidget(self.search_edit)
        layout.addWidget(self.list_widget, stretch=1)
        layout.addLayout(move_row)
        layout.addLayout(button_row)

        self._populate_items()

    def selected_item_ids(self) -> tuple[str, ...]:
        self._sync_temp_from_checkboxes()
        return self.temp_quick_access_ids

    def restore_defaults(self) -> None:
        self.temp_quick_access_ids = tuple(QUICK_ACCESS_DEFAULT_IDS)
        self._rebuild_items(self.temp_quick_access_ids)
        self._set_checked_ids(self.temp_quick_access_ids)
        self.list_widget.setCurrentRow(0 if self.list_widget.count() else -1)

    def accept(self) -> None:
        selected = self.selected_item_ids()
        if len(selected) > self.MAX_ITEMS:
            QMessageBox.warning(
                self,
                "Truy cập nhanh",
                f"Chỉ có thể chọn tối đa {self.MAX_ITEMS} chức năng truy cập nhanh.",
            )
            return
        super().accept()

    def _populate_items(self) -> None:
        self._rebuild_items(self.temp_quick_access_ids)

    def _rebuild_items(self, selected_ids: tuple[str, ...]) -> None:
        self.list_widget.clear()
        selected_order = {item_id: index for index, item_id in enumerate(selected_ids)}
        ordered = sorted(
            self.candidates,
            key=lambda item: (
                0 if item.id in selected_order else 1,
                selected_order.get(item.id, 0),
                item.group.casefold(),
                item.feature.title.casefold(),
            ),
        )
        for candidate in ordered:
            item = QListWidgetItem()
            item.setData(Qt.ItemDataRole.UserRole, candidate.id)
            item.setData(Qt.ItemDataRole.UserRole + 1, candidate.group)
            item.setData(Qt.ItemDataRole.UserRole + 2, candidate.feature.title)
            item.setText("")
            item.setToolTip(candidate.feature.description)
            self.list_widget.addItem(item)
            item.setSizeHint(QSize(0, self.ROW_HEIGHT))
            self._install_item_widget(
                item,
                candidate.feature.title,
                candidate.group,
                candidate.feature.icon,
                candidate.id in selected_ids,
            )
        self._filter_items(self.search_edit.text())

    def _set_checked_ids(self, selected_ids: tuple[str, ...]) -> None:
        selected = set(selected_ids)
        for row in range(self.list_widget.count()):
            item = self.list_widget.item(row)
            checkbox = self._item_checkbox(item)
            if checkbox is not None:
                checkbox.setChecked(
                    str(item.data(Qt.ItemDataRole.UserRole)) in selected
                )
        self.temp_quick_access_ids = self._checked_ids_from_widgets()

    def _checked_ids_from_widgets(self) -> tuple[str, ...]:
        selected: list[str] = []
        for row in range(self.list_widget.count()):
            item = self.list_widget.item(row)
            checkbox = self._item_checkbox(item)
            if checkbox is not None and checkbox.isChecked():
                selected.append(str(item.data(Qt.ItemDataRole.UserRole)))
        return tuple(selected)

    def _sync_temp_from_checkboxes(self) -> None:
        self.temp_quick_access_ids = self._checked_ids_from_widgets()

    def _item_checkbox(self, item: QListWidgetItem) -> QCheckBox | None:
        row_widget = self.list_widget.itemWidget(item)
        if row_widget is None:
            return None
        return row_widget.findChild(QCheckBox)

    def _install_item_widget(
        self,
        item: QListWidgetItem,
        title: str,
        group: str,
        icon_name: str,
        checked: bool,
    ) -> None:
        row_widget = QWidget()
        row_widget.setObjectName("QuickAccessRow")
        row_widget.setFixedHeight(self.ROW_HEIGHT)
        row_layout = QGridLayout(row_widget)
        row_layout.setContentsMargins(8, 5, 10, 5)
        row_layout.setHorizontalSpacing(8)
        row_layout.setVerticalSpacing(0)
        row_layout.setColumnMinimumWidth(0, self.CHECKBOX_COLUMN_WIDTH)
        row_layout.setColumnMinimumWidth(1, self.ICON_COLUMN_WIDTH)
        row_layout.setColumnMinimumWidth(3, self.GROUP_COLUMN_WIDTH)
        row_layout.setColumnStretch(0, 0)
        row_layout.setColumnStretch(1, 0)
        row_layout.setColumnStretch(2, 1)
        row_layout.setColumnStretch(3, 0)

        checkbox = QCheckBox()
        checkbox.setChecked(checked)
        checkbox.setCursor(Qt.CursorShape.PointingHandCursor)
        checkbox.setFixedSize(24, 24)

        icon_label = QLabel()
        pixmap = QPixmap(icon_path(icon_name))
        icon_label.setPixmap(
            pixmap.scaled(
                18,
                18,
                Qt.AspectRatioMode.KeepAspectRatio,
                Qt.TransformationMode.SmoothTransformation,
            )
        )
        icon_label.setFixedSize(24, 24)
        icon_label.setAlignment(Qt.AlignmentFlag.AlignCenter)

        text_label = QLabel(title)
        text_label.setWordWrap(True)
        text_label.setFixedHeight(34)
        text_label.setAlignment(Qt.AlignmentFlag.AlignVCenter)
        text_label.setObjectName("QuickAccessListText")

        group_label = QLabel(group)
        group_label.setObjectName("QuickAccessGroupBadge")
        group_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        group_label.setFixedWidth(self.GROUP_COLUMN_WIDTH)
        group_label.setFixedHeight(26)

        row_layout.addWidget(
            checkbox,
            0,
            0,
            alignment=Qt.AlignmentFlag.AlignCenter,
        )
        row_layout.addWidget(
            icon_label,
            0,
            1,
            alignment=Qt.AlignmentFlag.AlignCenter,
        )
        row_layout.addWidget(text_label, 0, 2)
        row_layout.addWidget(
            group_label,
            0,
            3,
            alignment=Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter,
        )
        self.list_widget.setItemWidget(item, row_widget)

    def _toggle_item_checkbox(self, item: QListWidgetItem) -> None:
        checkbox = self._item_checkbox(item)
        if checkbox is not None:
            checkbox.setChecked(not checkbox.isChecked())
            self._sync_temp_from_checkboxes()

    def _move_current(self, direction: int) -> None:
        row = self.list_widget.currentRow()
        target = row + direction
        if row < 0 or target < 0 or target >= self.list_widget.count():
            return
        current_item = self.list_widget.item(row)
        checkbox = self._item_checkbox(current_item)
        checked = checkbox.isChecked() if checkbox is not None else False
        title = str(current_item.data(Qt.ItemDataRole.UserRole + 2))
        group = str(current_item.data(Qt.ItemDataRole.UserRole + 1))
        item = self.list_widget.takeItem(row)
        self.list_widget.insertItem(target, item)
        item.setSizeHint(QSize(0, self.ROW_HEIGHT))
        candidate = next(
            (
                candidate
                for candidate in self.candidates
                if candidate.id == str(item.data(Qt.ItemDataRole.UserRole))
            ),
            None,
        )
        if candidate is not None:
            self._install_item_widget(
                item,
                candidate.feature.title,
                candidate.group,
                candidate.feature.icon,
                checked,
            )
        else:
            self._install_item_widget(item, title, group, "file.png", checked)
        self.list_widget.setCurrentRow(target)
        self._sync_temp_from_checkboxes()

    def _filter_items(self, text: str) -> None:
        needle = text.strip().casefold()
        for row in range(self.list_widget.count()):
            item = self.list_widget.item(row)
            searchable = (
                str(item.data(Qt.ItemDataRole.UserRole + 2))
                + " "
                + str(item.data(Qt.ItemDataRole.UserRole + 1))
            ).casefold()
            item.setHidden(bool(needle) and needle not in searchable)


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
        self.settings_database = AppSettingsDatabase()
        self.excel_service = ExcelService()
        self.settlement_engine = SettlementEngine()
        self.settlement_engine.register("mau04", Mau04Processor())
        self.settlement_engine.register("mau05", Mau05Processor())
        self.settlement_engine.register("mau06", Mau06Processor())
        self.settlement_engine.register("summary_06", Mau06Processor())
        self.settlement_engine.register("mau07", Mau0708Processor())
        self.settlement_engine.register("mau08", Mau0708Processor())
        self.settlement_engine.register("mau09", Mau09Processor())
        self.settlement_engine.register("mau13_14", Mau1314Processor())
        self.settlement_engine.register("summary_13_14", Mau1314Processor())
        self.settlement_engine.register("mau15_16", Mau1516Processor())
        self.settlement_engine.register("mau18", Mau18Processor())
        self.settlement_engine.register("mau20a", Mau20aProcessor())
        self.settlement_engine.register("mau22", Mau22Processor())
        self.settlement_engine.register("mau23", Mau23Processor())
        self.settlement_engine.register("mau24", Mau24Processor())
        self.settlement_engine.register("mau30", Mau30Processor())
        self.settlement_engine.register("summary_30", Mau30Processor())
        self.excel_context: ExcelContext | None = None
        self.sidebar_expanded = True
        self.quyet_toan_tin_dung_page: QWidget | None = None
        self.quyet_toan_ke_toan_page: QWidget | None = None
        self.quyet_toan_tong_hop_page: QWidget | None = None
        self.tovayvon_page: QWidget | None = None
        self.author_info_dialog: AuthorInfoDialog | None = None
        self.settlement_guidance_dialog: SettlementGuidanceDialog | None = None
        self.printer_settings_dialog: PrinterSettingsDialog | None = None
        self.quick_access_container: QWidget | None = None
        self.quick_access_layout: QVBoxLayout | None = None
        self._background_threads: list[object] = []
        self._background_processes: list[QProcess] = []
        self._nonblocking_messages: list[QMessageBox] = []

        root = QWidget()
        root.setObjectName("AppRoot")
        root_layout = QHBoxLayout(root)
        root_layout.setContentsMargins(0, 0, 0, 0)
        root_layout.setSpacing(0)

        root_layout.addWidget(self._build_sidebar())
        root_layout.addWidget(self._build_content(), stretch=1)
        self.setCentralWidget(root)

        status = QStatusBar()
        status.showMessage(APP_FOOTER_TEXT)
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
        return self.sidebar

    def _build_content(self) -> QWidget:
        content = QWidget()
        layout = QVBoxLayout(content)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        self.pages = QStackedWidget()
        self.pages.addWidget(self._build_dashboard())
        for section in NAVIGATION[1:]:
            features = SECTIONS[section]
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
                self.settings_widget.printer_settings_requested.connect(
                    self._show_printer_settings_dialog
                )
                self.settings_widget.quick_access_settings_requested.connect(
                    self._show_quick_access_settings_dialog
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

    def show_settlement_guidance(
        self,
        active_tab: int = SettlementGuidanceDialog.CREATE_30A_TAB,
    ) -> None:
        if (
            self.settlement_guidance_dialog is not None
            and self.settlement_guidance_dialog.isVisible()
        ):
            self.settlement_guidance_dialog.tabs.setCurrentIndex(active_tab)
            self.settlement_guidance_dialog.raise_()
            self.settlement_guidance_dialog.activateWindow()
            return

        dialog = SettlementGuidanceDialog(self, active_tab=active_tab)
        self.settlement_guidance_dialog = dialog
        dialog.finished.connect(self._clear_settlement_guidance_dialog)
        dialog.move(self.frameGeometry().center() - dialog.rect().center())
        dialog.open()

    def _clear_settlement_guidance_dialog(self) -> None:
        if self.settlement_guidance_dialog is not None:
            self.settlement_guidance_dialog.deleteLater()
            self.settlement_guidance_dialog = None

    def _show_printer_settings_dialog(self) -> None:
        if (
            self.printer_settings_dialog is not None
            and self.printer_settings_dialog.isVisible()
        ):
            self.printer_settings_dialog.raise_()
            self.printer_settings_dialog.activateWindow()
            return

        dialog = PrinterSettingsDialog(self)
        self.printer_settings_dialog = dialog
        dialog.finished.connect(self._clear_printer_settings_dialog)
        dialog.move(self.frameGeometry().center() - dialog.rect().center())
        dialog.open()

    def _clear_printer_settings_dialog(self) -> None:
        if self.printer_settings_dialog is not None:
            self.printer_settings_dialog.deleteLater()
            self.printer_settings_dialog = None

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

        metrics = QGridLayout()
        metrics.setHorizontalSpacing(12)
        metrics.setVerticalSpacing(12)
        metric_items = (
            ("217", "Ribbon callback"),
            ("8", "Nhóm chức năng"),
            ("53", "Biểu mẫu sẽ chuyển đổi"),
            ("13", "Add-in và workbook"),
        )
        for index, (value, label) in enumerate(metric_items):
            metrics.addWidget(self._metric_card(value, label), index // 3, index % 3)
        for column in range(3):
            metrics.setColumnStretch(column, 1)
        layout.addLayout(metrics)

        quick_header = QHBoxLayout()
        section_title = QLabel("Truy cập nhanh")
        section_title.setObjectName("PageTitle")
        settings_button = QPushButton("Cài đặt")
        settings_button.setObjectName("QuickSettingsButton")
        settings_button.setCursor(Qt.CursorShape.PointingHandCursor)
        settings_button.setToolTip(
            "Tùy chỉnh các chức năng hiển thị trong truy cập nhanh"
        )
        settings_button.clicked.connect(self._show_quick_access_settings_dialog)
        quick_header.addWidget(section_title)
        quick_header.addStretch()
        quick_header.addWidget(settings_button)
        layout.addLayout(quick_header)

        self.quick_access_container = QWidget()
        self.quick_access_layout = QVBoxLayout(self.quick_access_container)
        self.quick_access_layout.setContentsMargins(0, 0, 0, 0)
        self.quick_access_layout.setSpacing(12)
        layout.addWidget(self.quick_access_container)
        self._render_quick_access()
        layout.addStretch()
        return page

    def _quick_access_valid_ids(self) -> tuple[str, ...]:
        return tuple(item.id for item in QUICK_ACCESS_FEATURES)

    def _quick_access_by_id(self) -> dict[str, QuickAccessFeature]:
        return {item.id: item for item in QUICK_ACCESS_FEATURES}

    def _load_quick_access_ids(self) -> tuple[str, ...]:
        return self.settings_database.load_quick_access_items(
            QUICK_ACCESS_DEFAULT_IDS,
            self._quick_access_valid_ids(),
        )

    def _quick_access_features(self) -> tuple[Feature, ...]:
        by_id = self._quick_access_by_id()
        return tuple(
            by_id[item_id].feature
            for item_id in self._load_quick_access_ids()
            if item_id in by_id
        )

    def _render_quick_access(self) -> None:
        if self.quick_access_layout is None:
            return
        while self.quick_access_layout.count():
            item = self.quick_access_layout.takeAt(0)
            widget = item.widget()
            if widget is not None:
                widget.deleteLater()

        features = self._quick_access_features()
        if features:
            quick_grid = ResponsiveFeatureGrid(features)
            quick_grid.connect_requested(self.open_feature)
            self.quick_access_layout.addWidget(quick_grid)
            return

        empty_card = QFrame()
        empty_card.setObjectName("SettingsCard")
        empty_layout = QVBoxLayout(empty_card)
        empty_layout.setContentsMargins(18, 16, 18, 16)
        empty_layout.setSpacing(8)
        empty_title = QLabel("Chưa có chức năng truy cập nhanh.")
        empty_title.setObjectName("SectionTitle")
        empty_text = QLabel("Bấm Cài đặt để thêm chức năng thường dùng.")
        empty_text.setObjectName("MutedText")
        empty_button = QPushButton("Cài đặt truy cập nhanh")
        empty_button.setObjectName("PrimaryButton")
        empty_button.clicked.connect(self._show_quick_access_settings_dialog)
        empty_layout.addWidget(empty_title)
        empty_layout.addWidget(empty_text)
        empty_layout.addWidget(empty_button, alignment=Qt.AlignmentFlag.AlignLeft)
        self.quick_access_layout.addWidget(empty_card)

    def _show_quick_access_settings_dialog(self) -> None:
        dialog = QuickAccessSettingsDialog(
            QUICK_ACCESS_FEATURES,
            self._load_quick_access_ids(),
            self,
        )
        if dialog.exec() != QDialog.DialogCode.Accepted:
            return
        selected = self.settings_database.save_quick_access_items(
            dialog.selected_item_ids(),
            self._quick_access_valid_ids(),
            limit=QuickAccessSettingsDialog.MAX_ITEMS,
        )
        self._render_quick_access()
        if not selected:
            self.statusBar().showMessage("Đã bỏ trống Truy cập nhanh.", 4000)
        else:
            self.statusBar().showMessage("Đã lưu cài đặt Truy cập nhanh.", 4000)

    def _build_feature_page(self, title: str, features: list[Feature]) -> QWidget:
        page = self._scroll_page()
        body = page.widget()
        layout = body.layout()

        feature_groups = FEATURE_GROUPS.get(title)
        if feature_groups:
            for group in feature_groups:
                group_title = QLabel(group.title)
                group_title.setObjectName("SectionTitle")
                layout.addWidget(group_title)
                grid = ResponsiveFeatureGrid(group.features)
                grid.connect_requested(self.open_feature)
                layout.addWidget(grid)
            layout.addStretch()
            return page

        grid = ResponsiveFeatureGrid(features)
        grid.connect_requested(self.open_feature)
        layout.addWidget(grid)
        layout.addStretch()
        return page

    def _build_tovayvon_page(self) -> QWidget:
        page = self._scroll_page()
        body = page.widget()
        layout = body.layout()

        header = QHBoxLayout()
        title = QLabel("Tổ vay vốn")
        title.setObjectName("PageTitle")
        back_button = QPushButton("Quay lại Tín dụng")
        back_button.setObjectName("SecondaryButton")
        back_button.setCursor(Qt.CursorShape.PointingHandCursor)
        back_button.clicked.connect(lambda: self.select_page(NAVIGATION.index("Tín dụng")))
        header.addWidget(title)
        header.addStretch()
        header.addWidget(back_button)
        layout.addLayout(header)

        grid = ResponsiveFeatureGrid(TOVAYVON_FEATURES)
        grid.connect_requested(self.open_feature)
        layout.addWidget(grid)
        layout.addStretch()
        return page

    def _show_tovayvon_page(self) -> None:
        if self.tovayvon_page is not None:
            self.pages.removeWidget(self.tovayvon_page)
            self.tovayvon_page.deleteLater()
        self.tovayvon_page = self._build_tovayvon_page()
        self.pages.addWidget(self.tovayvon_page)
        self.pages.setCurrentWidget(self.tovayvon_page)
        self.nav_buttons[NAVIGATION.index("Tín dụng")].setChecked(True)

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

    def _build_quyet_toan_tong_hop_page(self) -> QWidget:
        page = self._scroll_page()
        body = page.widget()
        layout = body.layout()

        header = QHBoxLayout()
        title = QLabel("Quyết toán tổng hợp")
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

        for feature in self._quyet_toan_tong_hop_features():
            item = FeatureMenuItem(feature)
            item.requested.connect(self._open_quyet_toan_tong_hop_feature)
            layout.addWidget(item)
        layout.addStretch()
        return page

    def _quyet_toan_tong_hop_features(self) -> list[Feature]:
        branch_code = self._current_branch_code()
        return [
            Feature(
                feature.title.replace("{MaCN}", branch_code),
                feature.description,
                feature.icon,
            )
            for feature in QUYET_TOAN_TONG_HOP_FEATURES
        ]

    def _show_quyet_toan_tong_hop_page(self) -> None:
        if self.quyet_toan_tong_hop_page is not None:
            self.pages.removeWidget(self.quyet_toan_tong_hop_page)
            self.quyet_toan_tong_hop_page.deleteLater()
        self.quyet_toan_tong_hop_page = self._build_quyet_toan_tong_hop_page()
        self.pages.addWidget(self.quyet_toan_tong_hop_page)
        self.pages.setCurrentWidget(self.quyet_toan_tong_hop_page)
        self.nav_buttons[NAVIGATION.index("Quyết toán")].setChecked(True)

    def _open_quyet_toan_tong_hop_feature(self, title: str) -> None:
        if title == "Hướng dẫn tổng hợp số liệu quyết toán":
            self.show_settlement_guidance(
                SettlementGuidanceDialog.CONSOLIDATION_TAB
            )
            return
        for spec_key in (
            "consolidation.05",
            "consolidation.06",
            "consolidation.13",
            "consolidation.14",
            "consolidation.15a",
            "consolidation.15b",
            "consolidation.16",
            "consolidation.18",
            "consolidation.30a",
        ):
            spec = SETTLEMENT_SPECS[spec_key]
            if title.casefold().startswith(f"Tổng hợp Mẫu biểu {spec.report_code}/QT".casefold()):
                if spec_key == "consolidation.05":
                    self._run_consolidation_05_dialog()
                    return
                if spec_key == "consolidation.06":
                    self._run_mau06_dialog("consolidation.06")
                    return
                if spec_key == "consolidation.13":
                    self._run_consolidation_1314_dialog("consolidation.13")
                    return
                if spec_key == "consolidation.14":
                    self._run_consolidation_1314_dialog("consolidation.14")
                    return
                if spec_key == "consolidation.15a":
                    self._run_consolidation_15ab_dialog("consolidation.15a")
                    return
                if spec_key == "consolidation.15b":
                    self._run_consolidation_15ab_dialog("consolidation.15b")
                    return
                if spec_key == "consolidation.16":
                    self._run_consolidation_15ab_dialog("consolidation.16")
                    return
                if spec_key == "consolidation.18":
                    self._run_consolidation_15ab_dialog("consolidation.18")
                    return
                self._show_consolidation_not_ready(title, spec_key)
                return
            if title.casefold().startswith(f"Tạo Mẫu biểu {spec.report_code}/QT".casefold()):
                if spec_key == "consolidation.06":
                    self._run_mau06_dialog("consolidation.06")
                    return
                if spec_key == "consolidation.30a":
                    self._run_mau30_dialog("consolidation.30a")
                    return
                self._show_consolidation_not_ready(title, spec_key)
                return
        QMessageBox.information(
            self,
            title,
            "Chức năng này đang nằm trong lộ trình chuyển đổi từ VBA sang Python.",
        )

    def _show_consolidation_not_ready(self, title: str, spec_key: str) -> None:
        spec = SETTLEMENT_SPECS[spec_key]
        QMessageBox.information(
            self,
            title,
            f"{spec.title} đã có trong menu tổng hợp. Processor tổng hợp sẽ được chuyển từ VBA sang Python ở bước tiếp theo.",
        )

    def _run_same_structure_merge_dialog(self) -> None:
        dialog = SameStructureMergeDialog(self)
        if dialog.exec() != QDialog.DialogCode.Accepted:
            return
        output_path = dialog.output_path()
        if output_path is None:
            return
        if output_path.exists():
            answer = QMessageBox.question(
                self,
                "Nối file cùng cấu trúc",
                f"File {output_path.name} đã tồn tại. Ghi đè file này?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No,
            )
            if answer != QMessageBox.StandardButton.Yes:
                return

        progress = self._show_busy_dialog(
            "Đang nối các file cùng cấu trúc...\nVui lòng chờ trong giây lát!"
        )
        execution_error: Exception | None = None
        result = None
        try:
            if dialog.source_kind() == "csv":
                result = merge_same_structure_csv_to_xlsx(
                    dialog.source_paths,
                    output_path,
                    include_source_filename=dialog.include_source_filename(),
                )
            else:
                result = merge_same_structure_excel_to_xlsx(
                    dialog.source_paths,
                    output_path,
                    include_source_filename=dialog.include_source_filename(),
                )
        except (FileMergeError, OSError) as exc:
            execution_error = exc
        finally:
            self._close_busy_dialog(progress)

        if execution_error is not None:
            QMessageBox.warning(
                self,
                "Không thể nối file cùng cấu trúc",
                str(execution_error),
            )
            return
        if result is None:
            QMessageBox.warning(
                self,
                "Không thể nối file cùng cấu trúc",
                "Không nhận được kết quả xử lý.",
            )
            return
        self.statusBar().showMessage(
            f"Đã nối {result.source_count} file, {result.row_count} dòng dữ liệu."
        )
        self._show_result_message(
            "Hoàn thành nối file cùng cấu trúc",
            (
                f"Đã tạo file:\n{result.output_path}\n\n"
                f"Số file nguồn: {result.source_count}\n"
                f"Số dòng dữ liệu: {result.row_count}\n"
                f"Số cột: {result.column_count}"
            ),
            result.output_path,
        )

    def _run_csv_to_excel_dialog(self) -> None:
        dialog = CsvToExcelDialog(self)
        if dialog.exec() != QDialog.DialogCode.Accepted:
            return
        conversions = dialog.batch_outputs()
        if not conversions:
            return
        existing_outputs = [output for _, output in conversions if output.exists()]
        if existing_outputs:
            names = ", ".join(path.name for path in existing_outputs[:8])
            if len(existing_outputs) > 8:
                names += f", ... và {len(existing_outputs) - 8} file khác"
            answer = QMessageBox.question(
                self,
                "Chuyển CSV sang Excel",
                f"Có {len(existing_outputs)} file kết quả đã tồn tại:\n{names}\n\nGhi đè các file này?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No,
            )
            if answer != QMessageBox.StandardButton.Yes:
                return
        progress = self._show_busy_dialog(
            "Đang chuyển CSV sang Excel...\nVui lòng chờ trong giây lát!"
        )
        execution_error: Exception | None = None
        results = []
        try:
            for source_path, output_path in conversions:
                results.append(
                    convert_csv_to_excel(
                        source_path,
                        output_path,
                        output_format=dialog.output_format(),
                    )
                )
        except (ExcelToolError, OSError) as exc:
            execution_error = exc
        finally:
            self._close_busy_dialog(progress)
        if execution_error is not None:
            QMessageBox.warning(self, "Không thể chuyển CSV sang Excel", str(execution_error))
            return
        if not results:
            return
        self.statusBar().showMessage(f"Đã chuyển {len(results)} file CSV sang Excel")
        preview = "\n".join(str(result.output_path) for result in results[:12])
        if len(results) > 12:
            preview += f"\n... và {len(results) - 12} file khác."
        self._show_result_message(
            "Hoàn thành chuyển CSV sang Excel",
            (
                f"Đã tạo {len(results)} file:\n{preview}\n\n"
                f"Tổng số dòng: {sum(result.row_count for result in results)}"
            ),
            results[0].output_path,
        )

    def _run_split_sheets_dialog(self) -> None:
        dialog = SplitSheetsDialog(self)
        if dialog.exec() != QDialog.DialogCode.Accepted:
            return
        if dialog.source_path is None or dialog.output_directory is None:
            return
        progress = self._show_busy_dialog(
            "Đang tách các sheet thành file riêng...\nVui lòng chờ trong giây lát!"
        )
        execution_error: Exception | None = None
        result = None
        try:
            result = split_workbook_sheets_to_files(
                dialog.source_path,
                dialog.output_directory,
                sheet_names=dialog.selected_sheet_names(),
            )
        except (ExcelToolError, OSError) as exc:
            execution_error = exc
        finally:
            self._close_busy_dialog(progress)
        if execution_error is not None:
            QMessageBox.warning(self, "Không thể tách sheet", str(execution_error))
            return
        if result is None:
            return
        self.statusBar().showMessage(
            f"Đã tách {len(result.output_paths)} sheet vào {result.output_directory}"
        )
        preview = "\n".join(str(path) for path in result.output_paths[:12])
        if len(result.output_paths) > 12:
            preview += f"\n... và {len(result.output_paths) - 12} file khác."
        self._show_result_message(
            "Hoàn thành tách sheet",
            f"Đã tạo {len(result.output_paths)} file:\n{preview}",
            result.output_directory,
        )

    def _run_word_folder_print_dialog(self) -> None:
        dialog = WordFolderPrintDialog(self)
        if dialog.exec() != QDialog.DialogCode.Accepted:
            return
        source_paths = dialog.word_files
        if not source_paths:
            return
        progress = self._show_busy_dialog(
            "Đang gửi các file Word tới máy in...\nVui lòng chờ trong giây lát!"
        )

        def run(progress_callback):
            return print_word_files(source_paths, progress=progress_callback)

        def on_progress(message: str) -> None:
            self._update_busy_dialog(progress, message)
            self.statusBar().showMessage(message)

        def cleanup_thread(thread) -> None:
            if thread in self._background_threads:
                self._background_threads.remove(thread)

        def on_finished(result) -> None:
            self._close_busy_dialog(progress)
            cleanup_thread(thread)
            detail = (
                f"Đã gửi in {result.printed_count}/{result.file_count} file Word."
            )
            if result.failed:
                failed_lines = "\n".join(
                    f"- {path.name}: {error}" for path, error in result.failed[:20]
                )
                detail += f"\n\nCác file lỗi:\n{failed_lines}"
                if len(result.failed) > 20:
                    detail += f"\n... và {len(result.failed) - 20} file lỗi khác."
            self.statusBar().showMessage(detail.splitlines()[0])
            self._show_result_message(
                "Hoàn thành in file Word",
                detail,
                result.folder_path,
            )

        def on_failed(exc: Exception) -> None:
            self._close_busy_dialog(progress)
            cleanup_thread(thread)
            QMessageBox.warning(
                self,
                "Không thể in file Word",
                str(exc),
            )

        thread = run_in_thread(
            self,
            run,
            on_finished,
            on_failed,
            on_progress,
        )
        self._background_threads.append(thread)

    def _run_consolidation_05_dialog(self) -> None:
        spec = SETTLEMENT_SPECS["consolidation.05"]
        try:
            profile = self.settings_widget.database.load_branch_profile()
        except SettingsDatabaseError as exc:
            QMessageBox.warning(self, f"Không thể tạo {spec.title}", str(exc))
            return
        dialog = ConsolidationCsvDialog(spec, profile, self)
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
            "Đang nối các file CSV và tạo tổng hợp Mẫu 05/QT...\n"
            "Vui lòng chờ trong giây lát!"
        )
        execution_error: Exception | None = None
        merge_result = None
        processed_rows = 0
        merged_csv_path = output_path.with_suffix(".merged.csv")
        try:
            merge_result = merge_same_structure_csv_to_csv(
                dialog.source_paths,
                merged_csv_path,
            )
            processed_rows = Summary05Processor().execute(
                SettlementRequest(
                    spec=spec,
                    profile=profile,
                    options=dialog.options(),
                    source_paths=(merged_csv_path,),
                ),
                merged_csv_path,
                output_path,
            )
        except (FileMergeError, SettlementError, OSError) as exc:
            execution_error = exc
        finally:
            if merged_csv_path.exists():
                try:
                    merged_csv_path.unlink()
                except OSError:
                    pass
            self._close_busy_dialog(progress)

        if execution_error is not None:
            QMessageBox.warning(
                self,
                f"Không thể tạo {spec.title}",
                str(execution_error),
            )
            return
        if merge_result is None:
            QMessageBox.warning(
                self,
                f"Không thể tạo {spec.title}",
                "Không nhận được kết quả nối file.",
            )
            return
        self.statusBar().showMessage(
            f"Đã nối {merge_result.source_count} file CSV và tạo tổng hợp Mẫu 05/QT"
        )
        self._show_result_message(
            f"Hoàn thành {spec.title}",
            "Đã nối file CSV và tạo sheet TongHop_Mau05 trong file:\n"
            f"{output_path}\n\n"
            f"Số dòng dữ liệu đã xử lý: {processed_rows:,}",
            output_path,
        )

    def _run_consolidation_1314_dialog(self, spec_key: str) -> None:
        spec = SETTLEMENT_SPECS[spec_key]
        report_code = spec.report_code
        try:
            profile = self.settings_widget.database.load_branch_profile()
        except SettingsDatabaseError as exc:
            QMessageBox.warning(self, f"Không thể tạo {spec.title}", str(exc))
            return
        dialog = ConsolidationCsvDialog(spec, profile, self)
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

        source_paths = tuple(dialog.source_paths)
        options = dialog.options()
        merged_csv_path = output_path.with_suffix(".merged.csv")
        request_path = output_path.with_suffix(".request.json")
        try:
            request_path.write_text(
                json.dumps(
                    {
                        "source_paths": [str(path) for path in source_paths],
                        "output_path": str(output_path),
                        "merged_csv_path": str(merged_csv_path),
                        "spec_key": spec_key,
                        "profile": asdict(profile),
                        "options": asdict(options),
                    },
                    ensure_ascii=False,
                    indent=2,
                ),
                encoding="utf-8",
            )
        except OSError as exc:
            QMessageBox.warning(
                self,
                f"Không thể tạo {spec.title}",
                f"Không ghi được file request xử lý mẫu {report_code}:\n{exc}",
            )
            return
        progress = self._show_busy_dialog(
            f"Đang nối các file CSV và tạo tổng hợp Mẫu {report_code}/QT...\n"
            "Dữ liệu lớn có thể mất vài phút, vui lòng chờ!"
        )

        process = QProcess(self)
        process.setProgram(sys.executable)
        process.setArguments(
            [
                "-m",
                "agribank_v3.settlement.consolidation1314_worker",
                str(request_path),
            ]
        )
        process.setWorkingDirectory(str(Path(__file__).resolve().parents[3]))
        stdout_chunks: list[str] = []
        stderr_chunks: list[str] = []
        process_done = False

        def read_stdout() -> None:
            text = bytes(process.readAllStandardOutput()).decode("utf-8", errors="replace")
            if not text:
                return
            stdout_chunks.append(text)
            print(text, end="", flush=True)
            last_line = next((line for line in reversed(text.splitlines()) if line.strip()), "")
            if last_line:
                self.statusBar().showMessage(last_line)

        def read_stderr() -> None:
            text = bytes(process.readAllStandardError()).decode("utf-8", errors="replace")
            if not text:
                return
            stderr_chunks.append(text)
            print(text, end="", flush=True)

        def cleanup_request() -> None:
            if request_path.exists():
                try:
                    request_path.unlink()
                except OSError:
                    pass

        def on_finished(exit_code: int, exit_status: QProcess.ExitStatus) -> None:
            nonlocal process_done
            if process_done:
                return
            process_done = True
            read_stdout()
            read_stderr()
            self._close_busy_dialog(progress)
            cleanup_request()
            if process in self._background_processes:
                self._background_processes.remove(process)
            process.deleteLater()

            stdout_text = "".join(stdout_chunks)
            stderr_text = "".join(stderr_chunks).strip()
            if exit_status == QProcess.ExitStatus.NormalExit and exit_code == 0:
                self.statusBar().showMessage(f"Đã tạo {output_path.name}")
                self._show_result_message(
                    f"Hoàn thành {spec.title}",
                    (
                        f"Đã tạo file tổng hợp Mẫu {report_code}/QT:\n"
                        f"{output_path}\n\n"
                        "Do file lớn, ứng dụng không tự mở Excel. Hãy mở file từ thư mục kết quả."
                    ),
                    output_path,
                )
                return

            detail = stderr_text or stdout_text.strip() or f"Tiến trình kết thúc với mã lỗi {exit_code}."
            QMessageBox.warning(
                self,
                f"Không thể tạo {spec.title}",
                detail[-4000:],
            )

        def on_error(error: QProcess.ProcessError) -> None:
            nonlocal process_done
            if process_done:
                return
            process_done = True
            self._close_busy_dialog(progress)
            cleanup_request()
            if process in self._background_processes:
                self._background_processes.remove(process)
            process.deleteLater()
            QMessageBox.warning(
                self,
                f"Không thể tạo {spec.title}",
                f"Không chạy được tiến trình xử lý mẫu {report_code}: {error.name}",
            )

        process.readyReadStandardOutput.connect(read_stdout)
        process.readyReadStandardError.connect(read_stderr)
        process.finished.connect(on_finished)
        process.errorOccurred.connect(on_error)
        self._background_processes.append(process)
        process.start()

    def _run_consolidation_15ab_dialog(self, spec_key: str) -> None:
        spec = SETTLEMENT_SPECS[spec_key]
        report_code = spec.report_code
        report_code_lower = report_code.casefold()
        try:
            profile = self.settings_widget.database.load_branch_profile()
        except SettingsDatabaseError as exc:
            QMessageBox.warning(self, f"Không thể tạo {spec.title}", str(exc))
            return
        dialog = ConsolidationCsvDialog(spec, profile, self)
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

        source_paths = tuple(dialog.source_paths)
        merged_csv_path = output_path.with_suffix(".merged.csv")
        request_path = output_path.with_suffix(".request.json")
        try:
            request_path.write_text(
                json.dumps(
                    {
                        "source_paths": [str(path) for path in source_paths],
                        "output_path": str(output_path),
                        "merged_csv_path": str(merged_csv_path),
                        "profile": asdict(profile),
                        "options": asdict(dialog.options()),
                    },
                    ensure_ascii=False,
                    indent=2,
                ),
                encoding="utf-8",
            )
        except OSError as exc:
            QMessageBox.warning(
                self,
                f"Không thể tạo {spec.title}",
                f"Không ghi được file request xử lý mẫu {report_code_lower}:\n{exc}",
            )
            return
        progress = self._show_busy_dialog(
            f"Đang nối các file CSV và tạo tổng hợp Mẫu {report_code_lower}/QT...\n"
            "Dữ liệu lớn có thể mất vài phút, vui lòng chờ!"
        )
        process = QProcess(self)
        process.setProgram(sys.executable)
        process.setArguments(
            [
                "-m",
                f"agribank_v3.settlement.consolidation{report_code_lower}_worker",
                str(request_path),
            ]
        )
        process.setWorkingDirectory(str(Path(__file__).resolve().parents[3]))
        stdout_chunks: list[str] = []
        stderr_chunks: list[str] = []
        process_done = False

        def read_stdout() -> None:
            text = bytes(process.readAllStandardOutput()).decode("utf-8", errors="replace")
            if not text:
                return
            stdout_chunks.append(text)
            print(text, end="", flush=True)
            last_line = next((line for line in reversed(text.splitlines()) if line.strip()), "")
            if last_line:
                self.statusBar().showMessage(last_line)

        def read_stderr() -> None:
            text = bytes(process.readAllStandardError()).decode("utf-8", errors="replace")
            if not text:
                return
            stderr_chunks.append(text)
            print(text, end="", flush=True)

        def cleanup_request() -> None:
            if request_path.exists():
                try:
                    request_path.unlink()
                except OSError:
                    pass

        def on_finished(exit_code: int, exit_status: QProcess.ExitStatus) -> None:
            nonlocal process_done
            if process_done:
                return
            process_done = True
            read_stdout()
            read_stderr()
            self._close_busy_dialog(progress)
            cleanup_request()
            if process in self._background_processes:
                self._background_processes.remove(process)
            process.deleteLater()
            stdout_text = "".join(stdout_chunks)
            stderr_text = "".join(stderr_chunks).strip()
            if exit_status == QProcess.ExitStatus.NormalExit and exit_code == 0:
                self.statusBar().showMessage(f"Đã tạo {output_path.name}")
                self._show_result_message(
                    f"Hoàn thành {spec.title}",
                    f"Đã tạo file tổng hợp Mẫu {report_code_lower}/QT:\n"
                    f"{output_path}\n\n"
                    "Do file lớn, ứng dụng không tự mở Excel. Hãy mở file từ thư mục kết quả.",
                    output_path,
                )
                return
            detail = stderr_text or stdout_text.strip() or f"Tiến trình kết thúc với mã lỗi {exit_code}."
            QMessageBox.warning(
                self,
                f"Không thể tạo {spec.title}",
                detail[-4000:],
            )

        def on_error(error: QProcess.ProcessError) -> None:
            nonlocal process_done
            if process_done:
                return
            process_done = True
            self._close_busy_dialog(progress)
            cleanup_request()
            if process in self._background_processes:
                self._background_processes.remove(process)
            process.deleteLater()
            QMessageBox.warning(
                self,
                f"Không thể tạo {spec.title}",
                f"Không chạy được tiến trình xử lý mẫu {report_code_lower}: {error.name}",
            )

        process.readyReadStandardOutput.connect(read_stdout)
        process.readyReadStandardError.connect(read_stderr)
        process.finished.connect(on_finished)
        process.errorOccurred.connect(on_error)
        self._background_processes.append(process)
        process.start()

    @staticmethod
    def _scroll_page() -> QScrollArea:
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
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
        target = str(path)

        def open_detached() -> None:
            started = QProcess.startDetached(
                "explorer.exe",
                ["/select,", target],
                str(path.parent),
            )
            message = (
                f"Đã mở thư mục chứa file: {path.name}"
                if started
                else f"Không mở được thư mục chứa file: {path}"
            )
            self.statusBar().showMessage(message)

        QTimer.singleShot(250, open_detached)
        if timer_was_active:
            QTimer.singleShot(300_000, self.auto_connect_timer.start)

    def _show_result_message(
        self,
        title: str,
        text: str,
        result_path: Path | None = None,
    ) -> None:
        message = QMessageBox(
            QMessageBox.Icon.Information,
            title,
            text,
            QMessageBox.StandardButton.Ok,
            self,
        )
        message.setModal(False)
        message.setWindowModality(Qt.WindowModality.NonModal)
        if result_path is not None:
            open_folder_button = message.addButton(
                "Mở thư mục",
                QMessageBox.ButtonRole.ActionRole,
            )

            def on_button_clicked(button) -> None:
                if button is open_folder_button:
                    self._open_result_file(result_path)

            message.buttonClicked.connect(on_button_clicked)
        self._nonblocking_messages.append(message)
        message.finished.connect(
            lambda _=0, box=message: (
                self._nonblocking_messages.remove(box)
                if box in self._nonblocking_messages
                else None
            )
        )
        message.show()
        message.raise_()
        message.activateWindow()

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
            self.brand_button.setToolTip("Nhấn để mở rộng menu")
            self.sidebar_excel_button.setText("XL")
            self.author_info_button.setText("i")
            for button in self.nav_buttons:
                button.setText("")

    def open_feature(self, title: str) -> None:
        if title in NAVIGATION:
            self.select_page(NAVIGATION.index(title))
            return

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

        if title == "Tổ vay vốn":
            self._show_tovayvon_page()
            return

        if title == AUTO_INTEREST_PLACEHOLDER_TITLE:
            AutoInterestPlaceholderDialog(self).exec()
            return

        if title in CREDIT_GROUP_MANAGEMENT_ROUTE_TITLES:
            CreditGroupManagementPlaceholderDialog(self).exec()
            return

        if title in CREDIT_TOVAYVON_PLACEHOLDER_TITLES:
            CreditMigrationPlaceholderDialog(title, self).exec()
            return

        if title == "Quyết toán tín dụng":
            self._show_quyet_toan_tin_dung_page()
            return

        if title == "Quyết toán kế toán":
            self._show_quyet_toan_ke_toan_page()
            return

        if title == "Quyết toán tổng hợp":
            self._show_quyet_toan_tong_hop_page()
            return

        if title.startswith("Hướng dẫn"):
            self.show_settlement_guidance(
                SettlementGuidanceDialog.CREATE_30A_TAB
            )
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

        if title in {"Ghép tệp Excel", "Nối file cùng cấu trúc"}:
            self._run_same_structure_merge_dialog()
            return

        if title == "Tách sheet thành từng file":
            self._run_split_sheets_dialog()
            return

        if title == "Chuyển CSV sang Excel":
            self._run_csv_to_excel_dialog()
            return

        if title == "In tất cả file Word trong 1 folder":
            self._run_word_folder_print_dialog()
            return

        if title == "Cài đặt máy in Word":
            self._show_printer_settings_dialog()
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
        self._show_result_message(
            f"Hoàn thành {spec.title}",
            (
                f"Đã tạo file:\n{result.output_path}"
                + (f"\n\n{chr(10).join(result.warnings)}" if result.warnings else "")
            ),
            result.output_path,
        )

    def _run_mau06_dialog(self, spec_key: str = "credit.06") -> None:
        spec = SETTLEMENT_SPECS[spec_key]
        try:
            profile = self.settings_widget.database.load_branch_profile()
        except SettingsDatabaseError as exc:
            QMessageBox.warning(self, f"Không thể tạo {spec.title}", str(exc))
            return

        dialog_kwargs = {"window_title": f"Tạo {spec.title}"}
        if spec_key == "consolidation.06":
            dialog_kwargs.update(
                {
                    "source_label_text": (
                        "Tên File nguồn tổng hợp Mẫu 05/QT dùng xử lý để tạo ra "
                        "Mẫu biểu 06QT là: file Tổng hợp Mẫu 05/QT đã tạo."
                    ),
                    "output_label_text": (
                        "Tên File tổng hợp Mẫu 06/QT sẽ được tạo ra:"
                    ),
                    "consolidation_output": True,
                }
            )
        dialog = Mau06SettlementDialog(profile, self, **dialog_kwargs)
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
        self._show_result_message(
            f"Hoàn thành {spec.title}",
            (
                f"Đã tạo file:\n{result.output_path}"
                + (f"\n\n{chr(10).join(result.warnings)}" if result.warnings else "")
            ),
            result.output_path,
        )

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
        self._show_result_message(
            f"Hoàn thành {spec.title}",
            (
                f"Đã tạo file:\n{result.output_path}"
                + (f"\n\n{chr(10).join(result.warnings)}" if result.warnings else "")
            ),
            result.output_path,
        )

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
        self._show_result_message(
            f"Hoàn thành {spec.title}",
            (
                f"Đã tạo file:\n{result.output_path}"
                + (f"\n\n{chr(10).join(result.warnings)}" if result.warnings else "")
            ),
            result.output_path,
        )

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
        self._show_result_message(
            f"Hoàn thành {spec.title}",
            (
                f"Đã tạo file:\n{result.output_path}"
                + (f"\n\n{chr(10).join(result.warnings)}" if result.warnings else "")
            ),
            result.output_path,
        )

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
        if spec_key == "consolidation.30a":
            self._run_mau30_background_process(spec, profile, dialog, source_path)
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
        self._show_result_message(
            f"Hoàn thành {spec.title}",
            f"Đã thêm sheet {result.worksheet_name} vào file:\n{result.output_path}",
            result.output_path,
        )

    def _run_mau30_background_process(
        self,
        spec,
        profile,
        dialog: Mau30SettlementDialog,
        source_path: Path,
    ) -> None:
        source_paths = (
            (source_path, dialog.balance_path)
            if dialog.balance_path is not None
            else (source_path,)
        )
        request_path = source_path.with_suffix(".mau30.request.json")
        try:
            request_path.write_text(
                json.dumps(
                    {
                        "spec_key": spec.key,
                        "source_paths": [str(path) for path in source_paths],
                        "profile": asdict(profile),
                        "options": asdict(dialog.options()),
                    },
                    ensure_ascii=False,
                    indent=2,
                ),
                encoding="utf-8",
            )
        except OSError as exc:
            QMessageBox.warning(
                self,
                f"Không thể tạo {spec.title}",
                f"Không ghi được file request xử lý mẫu 30a:\n{exc}",
            )
            return

        progress = self._show_busy_dialog(
            f"Đang tạo mẫu 30a tổng hợp từ Mẫu {dialog.selected_model}/QT...\n"
            "File lớn có thể mất vài phút, vui lòng chờ!"
        )
        process = QProcess(self)
        process.setProgram(sys.executable)
        process.setArguments(
            [
                "-m",
                "agribank_v3.settlement.mau30_worker",
                str(request_path),
            ]
        )
        process.setWorkingDirectory(str(Path(__file__).resolve().parents[3]))
        stdout_chunks: list[str] = []
        stderr_chunks: list[str] = []
        process_done = False

        def read_stdout() -> None:
            text = bytes(process.readAllStandardOutput()).decode("utf-8", errors="replace")
            if not text:
                return
            stdout_chunks.append(text)
            print(text, end="", flush=True)
            last_line = next((line for line in reversed(text.splitlines()) if line.strip()), "")
            if last_line:
                self.statusBar().showMessage(last_line)

        def read_stderr() -> None:
            text = bytes(process.readAllStandardError()).decode("utf-8", errors="replace")
            if not text:
                return
            stderr_chunks.append(text)
            print(text, end="", flush=True)

        def cleanup_request() -> None:
            if request_path.exists():
                try:
                    request_path.unlink()
                except OSError:
                    pass

        def on_finished(exit_code: int, exit_status: QProcess.ExitStatus) -> None:
            nonlocal process_done
            if process_done:
                return
            process_done = True
            read_stdout()
            read_stderr()
            self._close_busy_dialog(progress)
            cleanup_request()
            if process in self._background_processes:
                self._background_processes.remove(process)
            process.deleteLater()

            stdout_text = "".join(stdout_chunks)
            stderr_text = "".join(stderr_chunks).strip()
            if exit_status == QProcess.ExitStatus.NormalExit and exit_code == 0:
                if dialog.balance_path is not None:
                    try:
                        self.settings_widget.database.save_preference(
                            "mau30_last_balance_path",
                            str(dialog.balance_path),
                        )
                    except SettingsDatabaseError:
                        pass
                self.statusBar().showMessage(f"Đã thêm Mau30QT-{dialog.selected_model} vào {source_path.name}")
                self._show_result_message(
                    f"Hoàn thành {spec.title}",
                    f"Đã thêm sheet Mau30QT-{dialog.selected_model} vào file:\n{source_path}",
                    source_path,
                )
                return

            detail = stderr_text or stdout_text.strip() or f"Tiến trình kết thúc với mã lỗi {exit_code}."
            QMessageBox.warning(
                self,
                f"Không thể tạo {spec.title}",
                detail[-4000:],
            )

        def on_error(error: QProcess.ProcessError) -> None:
            nonlocal process_done
            if process_done:
                return
            process_done = True
            self._close_busy_dialog(progress)
            cleanup_request()
            if process in self._background_processes:
                self._background_processes.remove(process)
            process.deleteLater()
            QMessageBox.warning(
                self,
                f"Không thể tạo {spec.title}",
                f"Không chạy được tiến trình xử lý mẫu 30a: {error.name}",
            )

        process.readyReadStandardOutput.connect(read_stdout)
        process.readyReadStandardError.connect(read_stderr)
        process.finished.connect(on_finished)
        process.errorOccurred.connect(on_error)
        self._background_processes.append(process)
        process.start()

    def _show_busy_dialog(self, message: str) -> QDialog:
        progress = QDialog(self)
        progress.setWindowTitle("Đang xử lý")
        progress.setWindowModality(Qt.WindowModality.NonModal)
        progress.setModal(False)
        progress.setFixedSize(420, 92)
        layout = QVBoxLayout(progress)
        layout.setContentsMargins(18, 10, 18, 10)
        layout.setSpacing(0)
        label = QLabel(message, progress)
        label.setWordWrap(True)
        label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        label.setStyleSheet("font-weight: 600;")
        layout.addWidget(label)
        progress.setProperty("message_label", label)
        QApplication.setOverrideCursor(Qt.CursorShape.WaitCursor)
        progress.show()
        progress.raise_()
        progress.activateWindow()
        progress.repaint()
        QApplication.processEvents()
        QApplication.processEvents()
        return progress

    @staticmethod
    def _update_busy_dialog(progress: QDialog, message: str) -> None:
        label = progress.property("message_label")
        if isinstance(label, QLabel):
            label.setText(message)
            label.repaint()
        QApplication.processEvents()

    @staticmethod
    def _close_busy_dialog(progress: QDialog) -> None:
        progress.setModal(False)
        progress.setWindowModality(Qt.WindowModality.NonModal)
        progress.hide()
        progress.reject()
        progress.close()
        progress.deleteLater()
        while QApplication.overrideCursor() is not None:
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
