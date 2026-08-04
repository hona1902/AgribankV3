from __future__ import annotations

from dataclasses import dataclass
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
import math

from PySide6.QtCore import QSize
from PySide6.QtGui import QFont
from PySide6.QtWidgets import QFrame, QGridLayout, QHBoxLayout, QLabel, QPushButton, QSizePolicy, QVBoxLayout, QWidget
from PySide6.QtCore import Qt

from agribank_v3.ui.components.controls import secondary_button


@dataclass(frozen=True, slots=True)
class KpiMetric:
    title: str
    value: object = None
    value_type: str = "text"
    full_value: object = None
    tooltip: str = ""
    group: str = "main"
    signed: bool = False


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
        self.set_metrics([KpiMetric(str(label), "...", "loading") for label in labels])

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
                    widget.setParent(None)
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
        expanded = self.width() >= 1600 if self._user_secondary_expanded is None else self._user_secondary_expanded
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
    if value_type in {"number", "decimal", "average"}:
        display = format_decimal_vn(value, decimals=2, signed=metric.signed)
        full = format_decimal_vn(full_source, decimals=2, signed=metric.signed)
        return display, full
    text = str(value).strip()
    return (text or "—"), (str(full_source).strip() or "—")


def kpi_tooltip(title: str, full_value: str, value_type: str) -> str:
    _ = value_type
    return f"{title}\n{full_value}"


def format_money_vn(value: object, *, signed: bool = False) -> str:
    try:
        number = int(round(float(value or 0)))
    except (TypeError, ValueError):
        number = 0
    prefix = "+" if signed and number > 0 else ""
    return prefix + f"{number:,}".replace(",", ".")


def format_percent_vn(value: object, *, signed: bool = False, empty: str = "N/A") -> str:
    if value is None:
        return empty
    try:
        number = Decimal(str(value or 0)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    except (InvalidOperation, TypeError, ValueError):
        number = Decimal("0.00")
    prefix = "+" if signed and number > 0 else ""
    text = f"{number:,.2f}%"
    return prefix + text.replace(",", "_").replace(".", ",").replace("_", ".")


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


def format_decimal_vn(value: object, *, decimals: int = 2, signed: bool = False) -> str:
    try:
        number = float(value or 0)
    except (TypeError, ValueError):
        return "—"
    prefix = "+" if signed and number > 0 else ""
    return prefix + _format_decimal_vn(number, decimals)


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
