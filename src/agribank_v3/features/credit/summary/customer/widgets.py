from __future__ import annotations

from dataclasses import dataclass
import math

from PySide6.QtCore import QPoint, QRect, Signal, QSize, Qt, QTimer
from PySide6.QtGui import QColor, QFont, QPainter, QPen, QStandardItem, QStandardItemModel
from PySide6.QtWidgets import (
    QApplication,
    QAbstractItemView,
    QComboBox,
    QCompleter,
    QFrame,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QGridLayout,
    QLineEdit,
    QPushButton,
    QStyledItemDelegate,
    QTableView,
    QToolTip,
    QVBoxLayout,
    QSizePolicy,
    QWidget,
)

from agribank_v3.features.credit.summary.officer_history.widgets import MultiLineHeaderView
from agribank_v3.features.credit.summary.customer.formatters import format_money_vn, format_percent_vn
from agribank_v3.ui.components.controls import (
    apply_agribank_combo_popup_style as _shared_apply_agribank_combo_popup_style,
    combo_box as _shared_combo_box,
    configure_combo_popup_width as _shared_configure_combo_popup_width,
    configure_searchable_combo as _shared_configure_searchable_combo,
    current_data as _shared_current_data,
    danger_button as _shared_danger_button,
    make_button_control as _shared_make_button_control,
    make_compact_control as _shared_make_compact_control,
    populate_combo as _shared_populate_combo,
    populate_officer_combo as _shared_populate_officer_combo,
    primary_button as _shared_primary_button,
    recommended_control_height,
    secondary_button as _shared_secondary_button,
)
from agribank_v3.ui.components.kpi import (
    CompactKpiCard as _SharedCompactKpiCard,
    KpiMetric as _SharedKpiMetric,
    MetricGrid as _SharedMetricGrid,
    ResponsiveKpiGrid as _SharedResponsiveKpiGrid,
    compact_money_vn as _shared_compact_money_vn,
    format_count_vn as _shared_format_count_vn,
    kpi_display_values as _shared_kpi_display_values,
    kpi_tooltip as _shared_kpi_tooltip,
    normalize_kpi_metric as _shared_normalize_kpi_metric,
)


OFFICER_NAME_ROLE = Qt.ItemDataRole.UserRole + 1
OFFICER_FULL_TEXT_ROLE = Qt.ItemDataRole.UserRole + 2
AGRIBANK_TABLE_SELECTION_STYLE = """
QTableView, QTableWidget {
    selection-background-color: rgba(174, 28, 63, 38);
    selection-color: #202020;
}
QTableView::item:hover, QTableWidget::item:hover {
    background-color: rgba(174, 28, 63, 18);
}
QTableView::item:selected, QTableWidget::item:selected {
    background-color: rgba(174, 28, 63, 38);
    color: #202020;
}
QTableView::item:selected:active, QTableWidget::item:selected:active {
    background-color: rgba(174, 28, 63, 48);
    color: #202020;
}
QTableView::item:selected:!active, QTableWidget::item:selected:!active {
    background-color: rgba(174, 28, 63, 28);
    color: #303030;
}
"""
AGRIBANK_COMBO_POPUP_STYLE = """
QAbstractItemView {
    background: #ffffff;
    border: 1px solid #d8dce3;
    border-radius: 6px;
    padding: 4px;
    outline: 0;
    selection-background-color: rgba(174, 28, 63, 42);
    selection-color: #202020;
}
QAbstractItemView::item {
    min-height: 24px;
    padding: 2px 8px;
    color: #202020;
}
QAbstractItemView::item:hover {
    background: rgba(174, 28, 63, 22);
}
QAbstractItemView::item:selected {
    background: rgba(174, 28, 63, 42);
    color: #202020;
}
QAbstractItemView::indicator {
    width: 16px;
    height: 16px;
}
"""
SCOPE_SELECT_ALL_VALUE = "__scope_filter_select_all__"
SCOPE_APPLY_VALUE = "__scope_filter_apply__"
SCOPE_CLEAR_VALUE = "__scope_filter_clear__"


@dataclass(frozen=True, slots=True)
class KpiMetric:
    title: str
    value: object = None
    value_type: str = "text"
    full_value: object = None
    tooltip: str = ""
    group: str = "main"
    signed: bool = False


def primary_button(text: str) -> QPushButton:
    return _shared_primary_button(text)


def secondary_button(text: str) -> QPushButton:
    return _shared_secondary_button(text)


def danger_button(text: str) -> QPushButton:
    return _shared_danger_button(text)


def make_compact_control(widget: QWidget) -> QWidget:
    if isinstance(widget, QPushButton):
        return _shared_make_button_control(widget)
    return _shared_make_compact_control(widget)


def apply_agribank_table_style(table: QWidget) -> QWidget:
    table.setStyleSheet(AGRIBANK_TABLE_SELECTION_STYLE)
    return table


def apply_agribank_combo_popup_style(combo: QComboBox) -> QComboBox:
    return _shared_apply_agribank_combo_popup_style(combo)


class CompactComboItemDelegate(QStyledItemDelegate):
    def sizeHint(self, option, index) -> QSize:
        size = super().sizeHint(option, index)
        flags = index.flags()
        target_height = 30 if flags & Qt.ItemFlag.ItemIsUserCheckable else 28
        size.setHeight(max(24, min(32, target_height)))
        return size


def combo_box(
    first_label: str = "Tất cả",
    *,
    minimum_width: int = 130,
    maximum_width: int | None = 190,
    minimum_contents_length: int = 10,
    searchable: bool = False,
) -> QComboBox:
    return _shared_combo_box(
        first_label,
        minimum_width=minimum_width,
        maximum_width=maximum_width,
        minimum_contents_length=minimum_contents_length,
        searchable=searchable,
    )


def populate_combo(combo: QComboBox, values: list[str] | tuple[tuple[str, str], ...] | list[tuple[str, str]]) -> None:
    _shared_populate_combo(combo, values)


def current_data(combo: QComboBox) -> str:
    return _shared_current_data(combo)


def configure_searchable_combo(combo: QComboBox) -> None:
    _shared_configure_searchable_combo(combo)


def configure_combo_popup_width(
    combo: QComboBox,
    *,
    minimum_popup_width: int = 320,
    maximum_screen_ratio: float = 0.70,
) -> int:
    return _shared_configure_combo_popup_width(
        combo,
        minimum_popup_width=minimum_popup_width,
        maximum_screen_ratio=maximum_screen_ratio,
    )


def populate_officer_combo(
    combo: QComboBox,
    officers: list[dict[str, object]] | tuple[dict[str, object], ...],
    *,
    first_label: str | None = None,
    display_code: bool = True,
    minimum_popup_width: int = 360,
) -> None:
    _shared_populate_officer_combo(
        combo,
        officers,
        first_label=first_label,
        display_code=display_code,
        minimum_popup_width=minimum_popup_width,
    )


def _officer_display_text(code: str, name: str, *, display_code: bool) -> str:
    if display_code and code:
        return f"[{code}] {name or code}"
    return name or code


def _apply_combo_item_tooltips(combo: QComboBox) -> None:
    for index in range(combo.count()):
        text = combo.itemText(index)
        combo.setItemData(index, text, Qt.ItemDataRole.ToolTipRole)
        combo.setItemData(index, text, OFFICER_FULL_TEXT_ROLE)
    _sync_combo_tooltip(combo, combo.currentIndex())
    _refresh_combo_completer(combo)


def _sync_combo_tooltip(combo: QComboBox, index: int) -> None:
    if index < 0:
        combo.setToolTip("")
        return
    tooltip = combo.itemData(index, Qt.ItemDataRole.ToolTipRole) or combo.itemText(index)
    combo.setToolTip(str(tooltip or ""))


def _refresh_combo_completer(combo: QComboBox) -> None:
    completer = combo.completer()
    if completer is not None:
        completer.setModel(combo.model())


class ScopeFilterComboBox(QComboBox):
    applied = Signal()

    def __init__(
        self,
        choices: tuple[tuple[str, str], ...],
        *,
        default_values: tuple[str, ...] = ("cross_branch",),
        parent=None,
    ) -> None:
        super().__init__(parent)
        self.setObjectName("AgribankComboBox")
        self.setModel(QStandardItemModel(self))
        self.setEditable(True)
        line_edit = self.lineEdit()
        if line_edit is not None:
            line_edit.setReadOnly(True)
        self._skip_next_hide = False
        self._updating_all_state = False
        self._choices = tuple((str(label), str(value)) for label, value in choices)
        self.view().pressed.connect(self._handle_pressed)
        apply_agribank_combo_popup_style(self)
        self.setMinimumWidth(230)
        self.setMaximumWidth(340)
        self.setSizePolicy(QSizePolicy.Policy.Preferred, QSizePolicy.Policy.Fixed)
        make_compact_control(self)
        self._rebuild_items()
        self.set_selected_values(default_values, emit=False)

    def selected_values(self) -> tuple[str, ...]:
        values: list[str] = []
        model = self.model()
        for row in range(1, 1 + len(self._choices)):
            item = model.item(row)
            if item is not None and item.checkState() == Qt.CheckState.Checked:
                values.append(str(item.data(Qt.ItemDataRole.UserRole) or ""))
        return tuple(value for value in values if value)

    def set_selected_values(self, values: tuple[str, ...] | list[str], *, emit: bool = False) -> None:
        selected = {str(value or "").strip() for value in values if str(value or "").strip()}
        model = self.model()
        for row in range(1, 1 + len(self._choices)):
            item = model.item(row)
            if item is not None:
                item.setCheckState(
                    Qt.CheckState.Checked
                    if str(item.data(Qt.ItemDataRole.UserRole) or "") in selected
                    else Qt.CheckState.Unchecked
                )
        self._sync_select_all_state()
        self._update_display_text()
        if emit:
            self.applied.emit()

    def select_all_scopes(self, *, emit: bool = False) -> None:
        self.set_selected_values(tuple(value for _label, value in self._choices), emit=emit)

    def hidePopup(self) -> None:
        if self._skip_next_hide:
            self._skip_next_hide = False
            return
        super().hidePopup()

    def _rebuild_items(self) -> None:
        model = self.model()
        model.clear()
        model.appendRow(self._action_item("Chọn tất cả", SCOPE_SELECT_ALL_VALUE, checkable=True))
        for label, value in self._choices:
            model.appendRow(self._action_item(label, value, checkable=True))
        model.appendRow(self._action_item("Áp dụng bộ lọc", SCOPE_APPLY_VALUE, checkable=False))
        model.appendRow(self._action_item("Xóa lọc", SCOPE_CLEAR_VALUE, checkable=False))
        configure_combo_popup_width(self, minimum_popup_width=320)

    def _action_item(self, text: str, value: str, *, checkable: bool) -> QStandardItem:
        item = QStandardItem(text)
        item.setData(value, Qt.ItemDataRole.UserRole)
        flags = Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsSelectable
        if checkable:
            flags |= Qt.ItemFlag.ItemIsUserCheckable
            item.setCheckState(Qt.CheckState.Unchecked)
        item.setFlags(flags)
        item.setData(text, Qt.ItemDataRole.ToolTipRole)
        return item

    def _handle_pressed(self, index) -> None:
        item = self.model().itemFromIndex(index)
        if item is None:
            return
        value = str(item.data(Qt.ItemDataRole.UserRole) or "")
        if value == SCOPE_APPLY_VALUE:
            self.hidePopup()
            self.applied.emit()
            return
        if value == SCOPE_CLEAR_VALUE:
            self.select_all_scopes(emit=False)
            self.hidePopup()
            self.applied.emit()
            return
        self._skip_next_hide = True
        if value == SCOPE_SELECT_ALL_VALUE:
            if item.checkState() == Qt.CheckState.Checked:
                self.set_selected_values((), emit=False)
            else:
                self.select_all_scopes(emit=False)
            return
        item.setCheckState(
            Qt.CheckState.Unchecked
            if item.checkState() == Qt.CheckState.Checked
            else Qt.CheckState.Checked
        )
        self._sync_select_all_state()
        self._update_display_text()

    def _sync_select_all_state(self) -> None:
        if self._updating_all_state:
            return
        self._updating_all_state = True
        try:
            model = self.model()
            all_item = model.item(0)
            if all_item is None:
                return
            total = len(self._choices)
            checked = len(self.selected_values())
            if checked == 0:
                all_item.setCheckState(Qt.CheckState.Unchecked)
            elif checked == total:
                all_item.setCheckState(Qt.CheckState.Checked)
            else:
                all_item.setCheckState(Qt.CheckState.PartiallyChecked)
        finally:
            self._updating_all_state = False

    def _update_display_text(self) -> None:
        selected = set(self.selected_values())
        labels = [label for label, value in self._choices if value in selected]
        if not labels:
            text = "Chưa chọn phạm vi vay"
        elif len(labels) == len(self._choices):
            text = "Tất cả phạm vi vay"
        elif len(labels) == 1:
            text = labels[0]
        else:
            text = f"Đã chọn {len(labels)} phạm vi"
        if self.lineEdit() is not None:
            self.lineEdit().setText(text)
        self.setToolTip(text)


class SearchBox(QLineEdit):
    debouncedTextChanged = Signal(str)

    def __init__(self, placeholder: str = "Tìm mã hoặc tên khách hàng", parent=None) -> None:
        super().__init__(parent)
        self.setObjectName("AgribankSearchBox")
        self.setPlaceholderText(placeholder)
        self.setClearButtonEnabled(True)
        self._timer = QTimer(self)
        self._timer.setSingleShot(True)
        self._timer.setInterval(350)
        self._timer.timeout.connect(lambda: self.debouncedTextChanged.emit(self.text()))
        self.textChanged.connect(lambda: self._timer.start())


class CustomerTableView(QTableView):
    def __init__(self, parent=None) -> None:
        super().__init__(parent)
        self.setObjectName("SummaryDataTable")
        apply_agribank_table_style(self)
        self.setHorizontalHeader(MultiLineHeaderView(Qt.Orientation.Horizontal, self))
        self.setAlternatingRowColors(True)
        self.setSortingEnabled(False)
        self.setSelectionBehavior(QTableView.SelectionBehavior.SelectRows)
        self.setSelectionMode(QTableView.SelectionMode.SingleSelection)
        self.setEditTriggers(QTableView.EditTrigger.NoEditTriggers)
        self.setHorizontalScrollMode(QAbstractItemView.ScrollMode.ScrollPerPixel)
        self.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        self.setTextElideMode(Qt.TextElideMode.ElideRight)
        self.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        self.setMinimumWidth(320)
        header = self.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        header.setStretchLastSection(False)
        header.setSectionsClickable(True)
        self.verticalHeader().setVisible(False)

    def apply_default_widths(self, widths: tuple[int, ...]) -> None:
        for index, width in enumerate(widths):
            self.setColumnWidth(index, width)
        header = self.horizontalHeader()
        if hasattr(header, "refresh_height"):
            header.refresh_height()


class Pager(QWidget):
    pageChanged = Signal(int)
    pageSizeChanged = Signal(int)

    def __init__(self, parent=None) -> None:
        super().__init__(parent)
        self.page = 1
        self.page_size = 100
        self.total_rows = 0
        layout = QHBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        self.first_button = secondary_button("Trang đầu")
        self.prev_button = secondary_button("Trang trước")
        self.next_button = secondary_button("Trang sau")
        self.last_button = secondary_button("Trang cuối")
        self.label = QLabel("Trang 1/1 - 0 dòng")
        self.size_combo = combo_box("100 dòng/trang")
        self.size_combo.clear()
        for size in (50, 100, 200, 500):
            self.size_combo.addItem(f"{size} dòng/trang", size)
        self.size_combo.setCurrentIndex(1)
        _apply_combo_item_tooltips(self.size_combo)
        configure_combo_popup_width(self.size_combo, minimum_popup_width=180)
        for button in (self.first_button, self.prev_button, self.next_button, self.last_button):
            layout.addWidget(button)
        layout.addWidget(self.label)
        layout.addWidget(self.size_combo)
        layout.addStretch()
        self.first_button.clicked.connect(lambda: self.pageChanged.emit(1))
        self.prev_button.clicked.connect(lambda: self.pageChanged.emit(max(1, self.page - 1)))
        self.next_button.clicked.connect(lambda: self.pageChanged.emit(self.page + 1))
        self.last_button.clicked.connect(lambda: self.pageChanged.emit(self.total_pages()))
        self.size_combo.currentIndexChanged.connect(self._emit_page_size)

    def set_state(self, *, page: int, page_size: int, total_rows: int) -> None:
        self.page = max(1, int(page or 1))
        self.page_size = max(1, int(page_size or 100))
        self.total_rows = max(0, int(total_rows or 0))
        total_pages = self.total_pages()
        self.label.setText(f"Trang {self.page}/{total_pages} - {self.total_rows:,} dòng")
        self.first_button.setEnabled(self.page > 1)
        self.prev_button.setEnabled(self.page > 1)
        self.next_button.setEnabled(self.page < total_pages)
        self.last_button.setEnabled(self.page < total_pages)

    def total_pages(self) -> int:
        return max(1, (self.total_rows + self.page_size - 1) // self.page_size)

    def _emit_page_size(self) -> None:
        self.pageSizeChanged.emit(int(self.size_combo.currentData() or 100))


class CompactToolbar(QWidget):
    def __init__(self, parent=None) -> None:
        super().__init__(parent)
        self.setObjectName("CustomerCompactToolbar")
        self.grid = QGridLayout(self)
        self.grid.setContentsMargins(0, 0, 0, 0)
        self.grid.setHorizontalSpacing(6)
        self.grid.setVerticalSpacing(4)
        self._items: list[QWidget] = []
        self.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)

    def addWidget(self, widget: QWidget) -> None:
        make_compact_control(widget)
        self._items.append(widget)
        self._relayout()

    def resizeEvent(self, event) -> None:
        super().resizeEvent(event)
        self._relayout()

    def sizeHint(self) -> QSize:
        rows = 2 if self.width() and self.width() < 760 and len(self._items) > 3 else 1
        return QSize(600, rows * 34 + max(0, rows - 1) * 4)

    def _relayout(self) -> None:
        while self.grid.count():
            self.grid.takeAt(0)
        if not self._items:
            return
        wrap = self.width() and self.width() < 760 and len(self._items) > 3
        split = 2 if wrap else len(self._items)
        for index, widget in enumerate(self._items):
            row = 0 if index < split else 1
            column = index if index < split else index - split
            self.grid.addWidget(widget, row, column)
        self.grid.setColumnStretch(len(self._items), 1)
        self.updateGeometry()


class CompactKpiCard(QFrame):
    CARD_HEIGHT = 66
    SECONDARY_CARD_HEIGHT = 58

    def __init__(self, metric: KpiMetric | None = None, *, secondary: bool = False, parent=None) -> None:
        super().__init__(parent)
        self.secondary = secondary
        self.setObjectName("CompactKpiCardSecondary" if secondary else "CompactKpiCard")
        height = self.SECONDARY_CARD_HEIGHT if secondary else self.CARD_HEIGHT
        self.setMinimumHeight(height)
        self.setMaximumHeight(height + 6)
        self.setMinimumWidth(132 if secondary else 140)
        self.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(8, 6, 8, 6)
        layout.setSpacing(3)
        self.title_label = QLabel()
        self.title_label.setObjectName("CompactKpiLabel")
        self.title_label.setWordWrap(False)
        self.title_label.setSizePolicy(QSizePolicy.Policy.Ignored, QSizePolicy.Policy.Fixed)
        self.value_label = QLabel()
        self.value_label.setObjectName("CompactKpiValueSecondary" if secondary else "CompactKpiValue")
        self.value_label.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        self.value_label.setWordWrap(False)
        self.value_label.setSizePolicy(QSizePolicy.Policy.Ignored, QSizePolicy.Policy.Fixed)
        value_font = QFont(self.value_label.font())
        value_font.setPointSize(14 if secondary else 17)
        value_font.setBold(True)
        self.value_label.setFont(value_font)
        title_font = QFont(self.title_label.font())
        title_font.setPointSize(10)
        self.title_label.setFont(title_font)
        layout.addWidget(self.title_label)
        layout.addWidget(self.value_label)
        if metric is not None:
            self.set_metric(metric)

    def set_metric(self, metric: KpiMetric) -> None:
        display_value, full_value = kpi_display_values(metric)
        tooltip = metric.tooltip or kpi_tooltip(metric.title, full_value, metric.value_type)
        self.title_label.setText(metric.title)
        self.value_label.setText(display_value)
        self.title_label.setToolTip(metric.title)
        self.value_label.setToolTip(tooltip)
        self.setToolTip(tooltip)

    def sizeHint(self) -> QSize:
        return QSize(150, self.SECONDARY_CARD_HEIGHT if self.secondary else self.CARD_HEIGHT)


class ResponsiveKpiGrid(QWidget):
    def __init__(self, parent=None) -> None:
        super().__init__(parent)
        self.setObjectName("ResponsiveKpiGrid")
        self.layout = QVBoxLayout(self)
        self.layout.setContentsMargins(0, 0, 0, 0)
        self.layout.setSpacing(5)
        self.main_container = QWidget()
        self.main_grid = QGridLayout(self.main_container)
        self.main_grid.setContentsMargins(0, 0, 0, 0)
        self.main_grid.setHorizontalSpacing(7)
        self.main_grid.setVerticalSpacing(7)
        self.layout.addWidget(self.main_container)
        self.secondary_toolbar = QWidget()
        self.secondary_toolbar.setObjectName("KpiSecondaryToolbar")
        secondary_toolbar_layout = QHBoxLayout(self.secondary_toolbar)
        secondary_toolbar_layout.setContentsMargins(0, 0, 0, 0)
        secondary_toolbar_layout.addStretch()
        self.toggle_secondary_button = secondary_button("Thu gọn chỉ tiêu")
        self.toggle_secondary_button.setObjectName("KpiToggleButton")
        self.toggle_secondary_button.clicked.connect(self._toggle_secondary)
        secondary_toolbar_layout.addWidget(self.toggle_secondary_button)
        self.layout.addWidget(self.secondary_toolbar)
        self.secondary_container = QWidget()
        self.secondary_grid = QGridLayout(self.secondary_container)
        self.secondary_grid.setContentsMargins(0, 0, 0, 0)
        self.secondary_grid.setHorizontalSpacing(6)
        self.secondary_grid.setVerticalSpacing(6)
        self.layout.addWidget(self.secondary_container)
        self._main_cards: list[CompactKpiCard] = []
        self._secondary_cards: list[CompactKpiCard] = []
        self._main_columns = 0
        self._secondary_columns = 0
        self._user_secondary_expanded: bool | None = None
        self.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)

    def set_metrics(self, metrics: list[KpiMetric | tuple]) -> None:
        self._clear_cards()
        for metric in (normalize_kpi_metric(item) for item in metrics):
            is_secondary = metric.group == "secondary"
            card = CompactKpiCard(metric, secondary=is_secondary)
            if is_secondary:
                self._secondary_cards.append(card)
            else:
                self._main_cards.append(card)
        self._refresh_secondary_visibility()
        self._relayout_cards(force=True)

    def set_loading(self, labels: list[str] | tuple[str, ...]) -> None:
        self.set_metrics([KpiMetric(str(label), "…", "loading") for label in labels])

    def set_empty(self, labels: list[str] | tuple[str, ...]) -> None:
        self.set_metrics([KpiMetric(str(label), None, "text") for label in labels])

    def main_column_count(self) -> int:
        return self._main_columns

    def secondary_column_count(self) -> int:
        return self._secondary_columns

    def sizeHint(self) -> QSize:
        width = self.width() if self.width() > 0 else 1600
        return QSize(width, self.estimated_height_for_width(width))

    def estimated_height_for_width(self, width: int) -> int:
        main_columns = self.column_count_for_width(max(1, width))
        main_rows = _row_count(len(self._main_cards), main_columns)
        height = main_rows * CompactKpiCard.CARD_HEIGHT + max(0, main_rows - 1) * 7
        has_secondary = bool(self._secondary_cards)
        if has_secondary:
            height += self.toggle_secondary_button.minimumHeight() + 5
            expanded = self._user_secondary_expanded if self._user_secondary_expanded is not None else width >= 1600
            if expanded:
                secondary_columns = self.column_count_for_width(max(1, width))
                secondary_rows = _row_count(len(self._secondary_cards), secondary_columns)
                height += secondary_rows * CompactKpiCard.SECONDARY_CARD_HEIGHT + max(0, secondary_rows - 1) * 6
        return height

    def column_count_for_width(self, width: int) -> int:
        if width >= 1600:
            return 8
        if width >= 1250:
            return 6
        if width >= 950:
            return 4
        if width >= 700:
            return 3
        return 2

    def resizeEvent(self, event) -> None:
        super().resizeEvent(event)
        self._refresh_secondary_visibility()
        self._relayout_cards()

    def _clear_cards(self) -> None:
        for grid in (self.main_grid, self.secondary_grid):
            while grid.count():
                item = grid.takeAt(0)
                widget = item.widget()
                if widget is not None:
                    widget.deleteLater()
        self._main_cards = []
        self._secondary_cards = []
        self._main_columns = 0
        self._secondary_columns = 0

    def _toggle_secondary(self) -> None:
        self._user_secondary_expanded = not self.secondary_container.isVisible()
        self._refresh_secondary_visibility()

    def _refresh_secondary_visibility(self) -> None:
        has_secondary = bool(self._secondary_cards)
        if self._user_secondary_expanded is None:
            expanded = self.width() >= 1600
        else:
            expanded = self._user_secondary_expanded
        self.secondary_toolbar.setVisible(has_secondary)
        self.secondary_container.setVisible(has_secondary and expanded)
        self.toggle_secondary_button.setText("Thu gọn chỉ tiêu" if expanded else "Xem thêm chỉ tiêu")

    def _relayout_cards(self, *, force: bool = False) -> None:
        width = max(1, self.width())
        main_columns = self.column_count_for_width(width)
        secondary_columns = self.column_count_for_width(width)
        if force or main_columns != self._main_columns:
            self._main_columns = main_columns
            self._relayout_group(self.main_grid, self._main_cards, main_columns)
        if force or secondary_columns != self._secondary_columns:
            self._secondary_columns = secondary_columns
            self._relayout_group(self.secondary_grid, self._secondary_cards, secondary_columns)
        self.updateGeometry()

    def _relayout_group(self, grid: QGridLayout, cards: list[CompactKpiCard], columns: int) -> None:
        while grid.count():
            grid.takeAt(0)
        for index, card in enumerate(cards):
            grid.addWidget(card, index // columns, index % columns)
        for column in range(max(1, columns)):
            grid.setColumnStretch(column, 1)


class MetricGrid(ResponsiveKpiGrid):
    pass


def normalize_kpi_metric(item: KpiMetric | tuple) -> KpiMetric:
    if isinstance(item, KpiMetric):
        return item
    if isinstance(item, tuple):
        if len(item) == 2:
            return KpiMetric(str(item[0]), item[1], "text")
        if len(item) == 3:
            return KpiMetric(str(item[0]), item[1], str(item[2]))
        if len(item) >= 4:
            return KpiMetric(str(item[0]), item[1], str(item[2]), group=str(item[3]))
    return KpiMetric(str(item), "", "text")


def kpi_display_values(metric: KpiMetric) -> tuple[str, str]:
    value = metric.value
    full_source = metric.full_value if metric.full_value is not None else value
    value_type = (metric.value_type or "text").lower()
    if value_type in {"loading", "busy"}:
        return "…", f"{metric.title}\nĐang tải dữ liệu"
    if _is_no_data(value):
        return "—", f"{metric.title}\nKhông có dữ liệu"
    if value_type == "money":
        display = compact_money_vn(value, signed=metric.signed)
        full = f"{format_money_vn(full_source, signed=metric.signed)} đồng"
        return display, full
    if value_type in {"percentage", "percent"}:
        display = format_percent_vn(value, signed=metric.signed, empty="—")
        return display, display
    if value_type == "count":
        display = format_count_vn(value, signed=metric.signed)
        return display, display
    text = str(value).strip()
    return (text or "—"), (str(full_source).strip() or "—")


def kpi_tooltip(title: str, full_value: str, value_type: str) -> str:
    unit = {
        "money": "",
        "percentage": "",
        "percent": "",
        "count": "",
        "text": "",
    }.get((value_type or "text").lower(), "")
    return f"{title}\n{full_value}{unit}"


def compact_money_vn(value: object, *, signed: bool = False) -> str:
    if _is_no_data(value):
        return "—"
    try:
        number = float(value or 0)
    except (TypeError, ValueError):
        return "—"
    prefix = "+" if signed and number > 0 else "-" if number < 0 else ""
    absolute = abs(number)
    if absolute >= 1_000_000_000:
        return f"{prefix}{_format_decimal_vn(absolute / 1_000_000_000, 2)} tỷ"
    if absolute >= 1_000_000:
        return f"{prefix}{_format_decimal_vn(absolute / 1_000_000, 2)} triệu"
    return format_money_vn(number, signed=signed)


def format_count_vn(value: object, *, signed: bool = False) -> str:
    try:
        number = int(round(float(value or 0)))
    except (TypeError, ValueError):
        return "—"
    prefix = "+" if signed and number > 0 else ""
    return prefix + f"{number:,}".replace(",", ".")


def _format_decimal_vn(value: float, decimals: int) -> str:
    text = f"{value:,.{decimals}f}"
    return text.replace(",", "_").replace(".", ",").replace("_", ".")


def _row_count(items: int, columns: int) -> int:
    if items <= 0:
        return 0
    return max(1, (items + max(1, columns) - 1) // max(1, columns))


def _is_no_data(value: object) -> bool:
    if value is None:
        return True
    if isinstance(value, float) and math.isnan(value):
        return True
    text = str(value).strip().lower()
    return text in {"", "none", "nan", "n/a"}

class QueryStateBanner(QWidget):
    retryRequested = Signal()

    def __init__(self, parent=None) -> None:
        super().__init__(parent)
        self.state = "ready"
        layout = QHBoxLayout(self)
        layout.setContentsMargins(10, 6, 10, 6)
        self.label = QLabel("")
        self.retry_button = secondary_button("Thử lại")
        self.retry_button.clicked.connect(self.retryRequested.emit)
        layout.addWidget(self.label)
        layout.addWidget(self.retry_button)
        layout.addStretch()
        self.hide()

    def clear(self) -> None:
        self.state = "ready"
        self.label.setText("")
        self.retry_button.hide()
        self.hide()

    def set_loading(self, message: str = "Đang tải dữ liệu...") -> None:
        self.state = "loading"
        self.label.setText(message)
        self.retry_button.hide()
        self.show()

    def set_empty(self, message: str = "Không có dữ liệu phù hợp với bộ lọc.") -> None:
        self.state = "empty"
        self.label.setText(message)
        self.retry_button.show()
        self.show()

    def set_error(self, message: str = "Không tải được dữ liệu. Vui lòng thử lại.") -> None:
        self.state = "error"
        self.label.setText(message)
        self.retry_button.show()
        self.show()


class CustomerChartWidget(QWidget):
    def __init__(self, title: str, *, value_kind: str = "money", chart_kind: str = "line", parent=None) -> None:
        super().__init__(parent)
        self.title = title
        self.value_kind = value_kind
        self.chart_kind = chart_kind
        self.series: tuple[tuple[str, tuple[tuple[str, float], ...]], ...] = ()
        self.points: list[tuple[QRect, str, str, float]] = []
        self.setMinimumHeight(220)
        self.setMouseTracking(True)

    def set_series(self, series: tuple[tuple[str, tuple[tuple[str, float], ...]], ...]) -> None:
        self.series = series
        self.update()

    def paintEvent(self, event) -> None:
        super().paintEvent(event)
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        painter.fillRect(self.rect(), QColor("#ffffff"))
        painter.setPen(QColor("#374151"))
        painter.drawText(QRect(8, 4, self.width() - 16, 22), Qt.AlignmentFlag.AlignLeft, self.title)
        rect = self.rect().adjusted(48, 34, -20, -34)
        painter.setPen(QColor("#d8dee8"))
        painter.drawRect(rect)
        self.points = []
        values = [value for _name, rows in self.series for _label, value in rows]
        if not values:
            painter.setPen(QColor("#6b7280"))
            painter.drawText(self.rect(), Qt.AlignmentFlag.AlignCenter, "Chưa có dữ liệu")
            return
        min_value = min(0.0, *values)
        max_value = max(0.0, *values)
        if min_value == max_value:
            max_value += 1
        colors = (QColor("#1f6feb"), QColor("#d97706"), QColor("#059669"), QColor("#7c3aed"))
        for index in range(5):
            ratio = index / 4
            y = rect.bottom() - int(rect.height() * ratio)
            painter.setPen(QColor("#e5e7eb"))
            painter.drawLine(rect.left(), y, rect.right(), y)
            value = min_value + (max_value - min_value) * ratio
            painter.setPen(QColor("#6b7280"))
            painter.drawText(QRect(0, y - 9, 44, 18), Qt.AlignmentFlag.AlignRight, self._format_axis(value))
        for series_index, (name, rows) in enumerate(self.series):
            color = colors[series_index % len(colors)]
            painter.setPen(QPen(color, 2))
            points = self._points(rect, rows, min_value, max_value)
            for first, second in zip(points, points[1:]):
                painter.drawLine(first[0], first[1], second[0], second[1])
            painter.setBrush(color)
            for x, y, label, value in points:
                hit = QRect(x - 4, y - 4, 8, 8)
                painter.drawEllipse(hit)
                self.points.append((hit.adjusted(-4, -4, 4, 4), label, name, value))
            painter.drawText(QRect(56 + series_index * 140, self.height() - 24, 132, 18), Qt.AlignmentFlag.AlignLeft, name)

    def mouseMoveEvent(self, event) -> None:
        position = event.position().toPoint()
        for rect, label, name, value in self.points:
            if rect.contains(position):
                QToolTip.showText(event.globalPosition().toPoint(), f"{label}\n{name}: {self._format_value(value)}", self)
                return
        QToolTip.hideText()
        super().mouseMoveEvent(event)

    def _points(self, rect: QRect, rows: tuple[tuple[str, float], ...], min_value: float, max_value: float) -> list[tuple[int, int, str, float]]:
        if not rows:
            return []
        span = max_value - min_value or 1
        output: list[tuple[int, int, str, float]] = []
        for index, (label, value) in enumerate(rows):
            x = rect.left() + int(rect.width() * index / max(1, len(rows) - 1))
            y = rect.bottom() - int((float(value or 0) - min_value) / span * rect.height())
            output.append((x, y, label, float(value or 0)))
        return output

    def _format_axis(self, value: float) -> str:
        if self.value_kind == "percent":
            return format_percent_vn(value)
        if abs(value) >= 1_000_000_000:
            return f"{value / 1_000_000_000:,.0f} tỷ".replace(",", ".")
        return format_money_vn(value)

    def _format_value(self, value: float) -> str:
        if self.value_kind == "percent":
            return format_percent_vn(value)
        return format_money_vn(value)


def fit_window_to_screen(
    window: QWidget,
    *,
    width_ratio: float = 0.9,
    height_ratio: float = 0.88,
    max_width: int = 1400,
    max_height: int = 900,
    min_width: int = 950,
    min_height: int = 650,
    saved_geometry: QRect | None = None,
) -> QRect:
    available = available_screen_geometry(window)
    minimum_width = min(max(1, int(min_width)), max(1, available.width()))
    minimum_height = min(max(1, int(min_height)), max(1, available.height()))
    window.setMinimumSize(QSize(minimum_width, minimum_height))
    default_width = min(max_width, int(available.width() * width_ratio))
    default_height = min(max_height, int(available.height() * height_ratio))
    default_width = min(available.width(), max(minimum_width, default_width))
    default_height = min(available.height(), max(minimum_height, default_height))
    if saved_geometry is not None and saved_geometry.isValid():
        geometry = ensure_geometry_visible(window, saved_geometry)
        window.setGeometry(geometry)
        return geometry
    window.resize(default_width, default_height)
    center_window_on_screen(window)
    return window.geometry()


def ensure_geometry_visible(window: QWidget, geometry: QRect) -> QRect:
    available = available_screen_geometry(window)
    if not geometry.isValid() or geometry.width() <= 0 or geometry.height() <= 0:
        return _centered_geometry(available, min(available.width(), 1000), min(available.height(), 700))
    width = min(max(1, geometry.width()), available.width())
    height = min(max(1, geometry.height()), available.height())
    candidate = QRect(geometry.topLeft(), QSize(width, height))
    if not available.intersects(candidate) or not available.contains(candidate.center()):
        return _centered_geometry(available, width, height)
    left = min(max(candidate.left(), available.left()), available.right() - width + 1)
    top = min(max(candidate.top(), available.top()), available.bottom() - height + 1)
    return QRect(QPoint(left, top), QSize(width, height))


def center_window_on_screen(window: QWidget) -> None:
    available = available_screen_geometry(window)
    geometry = _centered_geometry(available, window.width(), window.height())
    window.move(geometry.topLeft())


def available_screen_geometry(window: QWidget) -> QRect:
    screen = window.screen() if window is not None else None
    if screen is None:
        screen = QApplication.primaryScreen()
    if screen is None:
        return QRect(0, 0, 1280, 720)
    return screen.availableGeometry()


def _centered_geometry(available: QRect, width: int, height: int) -> QRect:
    width = min(max(1, width), max(1, available.width()))
    height = min(max(1, height), max(1, available.height()))
    left = available.left() + max(0, (available.width() - width) // 2)
    top = available.top() + max(0, (available.height() - height) // 2)
    return QRect(QPoint(left, top), QSize(width, height))


KpiMetric = _SharedKpiMetric
CompactKpiCard = _SharedCompactKpiCard
ResponsiveKpiGrid = _SharedResponsiveKpiGrid
MetricGrid = _SharedMetricGrid
compact_money_vn = _shared_compact_money_vn
format_count_vn = _shared_format_count_vn
kpi_display_values = _shared_kpi_display_values
kpi_tooltip = _shared_kpi_tooltip
normalize_kpi_metric = _shared_normalize_kpi_metric
