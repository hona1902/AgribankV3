from __future__ import annotations

from PySide6.QtCore import QPoint, QRect, Qt
from PySide6.QtGui import QColor, QPainter, QPen
from PySide6.QtWidgets import QToolTip, QWidget

from agribank_v3.features.credit.summary.dashboard_service import (
    DashboardBranchRow,
    DashboardPeriodRow,
    latest_branch_rows,
    metric_value,
)
from agribank_v3.features.credit.summary.nim_ui_config import NimUiConfig, NIM_DN_UI_CONFIG
from agribank_v3.features.credit.summary.officer_history.models import (
    METRIC_BALANCE,
    METRIC_BALANCE_GROWTH,
    ChartSeries,
)
from agribank_v3.features.credit.summary.officer_history.widgets import format_money_vn, format_percent_vn


class DashboardBarChart(QWidget):
    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.bars: tuple[tuple[str, str, float | None], ...] = ()
        self.value_kind = "money"
        self.metric_label = ""
        self.empty_message = "Không có dữ liệu."
        self.bar_rects: list[tuple[QRect, str, str, float | None]] = []
        self.setMinimumHeight(280)
        self.setMouseTracking(True)

    def set_bars(
        self,
        bars: tuple[tuple[str, str, float | None], ...],
        *,
        value_kind: str,
        metric_label: str,
        empty_message: str = "Không có dữ liệu.",
    ) -> None:
        self.bars = bars[:16]
        self.value_kind = value_kind
        self.metric_label = metric_label
        self.empty_message = empty_message
        self.setMinimumHeight(max(280, 72 + len(self.bars) * 30))
        self.update()

    def paintEvent(self, event) -> None:
        super().paintEvent(event)
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        painter.fillRect(self.rect(), QColor("#ffffff"))
        labels = [label for _period, label, value in self.bars if value is not None]
        left_margin = 92
        if labels:
            left_margin = max(190, max(self.fontMetrics().horizontalAdvance(label) for label in labels) + 18)
        left_margin = min(left_margin, max(190, self.width() - 220))
        rect = self.rect().adjusted(left_margin, 28, -24, -34)
        painter.setPen(QColor("#d8dee8"))
        painter.drawRect(rect)
        self.bar_rects = []
        visible = tuple((period, label, value) for period, label, value in self.bars if value is not None)
        if not visible:
            painter.setPen(QColor("#6b7280"))
            painter.drawText(self.rect(), Qt.AlignmentFlag.AlignCenter, self.empty_message)
            return
        values = [float(value or 0) for _period, _label, value in visible]
        min_value = min(0.0, *values)
        max_value = max(0.0, *values)
        if min_value == max_value:
            max_value = min_value + 1
        span = max_value - min_value
        zero_x = rect.left() + int((0 - min_value) / span * rect.width())
        painter.setPen(QColor("#6b7280"))
        for index in range(5):
            ratio = index / 4
            x = rect.left() + int(rect.width() * ratio)
            painter.drawLine(x, rect.top(), x, rect.bottom())
            value = min_value + span * ratio
            painter.drawText(QRect(x - 48, rect.bottom() + 6, 96, 18), Qt.AlignmentFlag.AlignCenter, _format_axis(value, self.value_kind))
        unit_label = "Đơn vị: tỷ đồng" if self.value_kind.startswith("money") and max(abs(value) for value in values) >= 1_000_000_000 else "Đơn vị: đồng"
        if self.value_kind.startswith("percent"):
            unit_label = "Đơn vị: %"
        painter.drawText(QRect(rect.left(), 6, rect.width(), 18), Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter, unit_label)

        count = len(visible)
        gap = max(6, min(10, rect.height() // max(1, count * 8)))
        bar_height = max(14, min(24, (rect.height() - gap * (count + 1)) // max(1, count)))
        painter.setPen(QPen(QColor("#2563eb"), 1))
        painter.setBrush(QColor("#2f80ed"))
        for index, (period, label, value) in enumerate(visible):
            number = float(value or 0)
            value_x = rect.left() + int((number - min_value) / span * rect.width())
            y = rect.top() + gap + index * (bar_height + gap)
            x = min(zero_x, value_x)
            width = max(2, abs(value_x - zero_x))
            bar_rect = QRect(x, y, width, bar_height)
            painter.drawRoundedRect(bar_rect, 3, 3)
            self.bar_rects.append((bar_rect.adjusted(-3, -3, 3, 3), period, label, value))
            painter.setPen(QColor("#374151"))
            painter.drawText(
                QRect(6, y - 2, left_margin - 14, bar_height + 4),
                Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter,
                label,
            )
            painter.drawText(
                QRect(max(zero_x, value_x) + 6, y - 2, max(62, self.width() - max(zero_x, value_x) - 10), bar_height + 4),
                Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter,
                _format_value(value, self.value_kind),
            )
            painter.setPen(QPen(QColor("#2563eb"), 1))

    def mouseMoveEvent(self, event) -> None:
        position = event.position().toPoint()
        for rect, period, label, value in self.bar_rects:
            if rect.contains(position):
                QToolTip.showText(
                    event.globalPosition().toPoint(),
                    f"{label}\nKỳ: {period}\n{self.metric_label}: {_format_value(value, self.value_kind)}",
                    self,
                )
                return
        QToolTip.hideText()
        super().mouseMoveEvent(event)


class DashboardBranchComparisonChart(QWidget):
    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.pairs: tuple[tuple[str, str, float | None, float | None], ...] = ()
        self.value_kind = "money"
        self.metric_label = ""
        self.from_period = ""
        self.to_period = ""
        self.empty_message = "Không có dữ liệu chi nhánh."
        self.bar_rects: list[tuple[QRect, str, str, float | None]] = []
        self.orientation = "horizontal_grouped"
        self.setMinimumHeight(280)
        self.setMouseTracking(True)

    def set_pairs(
        self,
        pairs: tuple[tuple[str, str, float | None, float | None], ...],
        *,
        value_kind: str,
        metric_label: str,
        from_period: str,
        to_period: str,
        empty_message: str = "Không có dữ liệu chi nhánh.",
    ) -> None:
        self.pairs = pairs[:16]
        self.value_kind = value_kind
        self.metric_label = metric_label
        self.from_period = from_period
        self.to_period = to_period
        self.empty_message = empty_message
        self.setMinimumHeight(max(300, 100 + len(self.pairs) * 40))
        self.update()

    def paintEvent(self, event) -> None:
        super().paintEvent(event)
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        painter.fillRect(self.rect(), QColor("#ffffff"))
        labels = [label for _code, label, _from_value, _to_value in self.pairs]
        left_margin = 190
        if labels:
            left_margin = max(190, max(self.fontMetrics().horizontalAdvance(label) for label in labels) + 18)
        left_margin = min(left_margin, max(190, self.width() - 240))
        rect = self.rect().adjusted(left_margin, 56, -24, -36)
        painter.setPen(QColor("#d8dee8"))
        painter.drawRect(rect)
        self.bar_rects = []
        visible = tuple(pair for pair in self.pairs if pair[2] is not None or pair[3] is not None)
        if not visible:
            painter.setPen(QColor("#6b7280"))
            painter.drawText(self.rect(), Qt.AlignmentFlag.AlignCenter, self.empty_message)
            return
        values = [float(value or 0) for _code, _label, from_value, to_value in visible for value in (from_value, to_value) if value is not None]
        min_value = min(0.0, *values)
        max_value = max(0.0, *values)
        if min_value == max_value:
            max_value = min_value + 1
        span = max_value - min_value
        zero_x = rect.left() + int((0 - min_value) / span * rect.width())
        painter.setPen(QColor("#6b7280"))
        for index in range(5):
            ratio = index / 4
            x = rect.left() + int(rect.width() * ratio)
            painter.drawLine(x, rect.top(), x, rect.bottom())
            value = min_value + span * ratio
            painter.drawText(QRect(x - 48, rect.bottom() + 6, 96, 18), Qt.AlignmentFlag.AlignCenter, _format_axis(value, self.value_kind))
        self._draw_legend(painter, rect)
        count = len(visible)
        group_gap = max(8, min(14, rect.height() // max(1, count * 6)))
        bar_height = max(8, min(14, (rect.height() - group_gap * (count + 1)) // max(1, count * 2)))
        colors = (QColor("#1f6feb"), QColor("#d97706"))
        for index, (_code, label, from_value, to_value) in enumerate(visible):
            group_top = rect.top() + group_gap + index * (bar_height * 2 + group_gap)
            painter.setPen(QColor("#374151"))
            painter.drawText(QRect(6, group_top - 2, left_margin - 14, bar_height * 2 + 4), Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter, label)
            for bar_index, (period, value) in enumerate(((self.from_period, from_value), (self.to_period, to_value))):
                if value is None:
                    continue
                number = float(value or 0)
                value_x = rect.left() + int((number - min_value) / span * rect.width())
                y = group_top + bar_index * bar_height
                x = min(zero_x, value_x)
                width = max(2, abs(value_x - zero_x))
                bar_rect = QRect(x, y, width, max(2, bar_height - 2))
                painter.setPen(QPen(colors[bar_index], 1))
                painter.setBrush(colors[bar_index])
                painter.drawRoundedRect(bar_rect, 3, 3)
                self.bar_rects.append((bar_rect.adjusted(-3, -3, 3, 3), period, label, value))
                painter.setPen(QColor("#374151"))
                painter.drawText(
                    QRect(max(zero_x, value_x) + 6, y - 2, max(62, self.width() - max(zero_x, value_x) - 10), bar_height + 2),
                    Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter,
                    _format_value(value, self.value_kind),
                )

    def mouseMoveEvent(self, event) -> None:
        position = event.position().toPoint()
        for rect, period, label, value in self.bar_rects:
            if rect.contains(position):
                QToolTip.showText(
                    event.globalPosition().toPoint(),
                    f"{label}\nKỳ: {period}\n{self.metric_label}: {_format_value(value, self.value_kind)}",
                    self,
                )
                return
        QToolTip.hideText()
        super().mouseMoveEvent(event)

    def _draw_legend(self, painter: QPainter, rect: QRect) -> None:
        metrics = self.fontMetrics()
        x = rect.left()
        y = 28
        for label, color in ((self.from_period or "Từ kỳ", QColor("#1f6feb")), (self.to_period or "Đến kỳ", QColor("#d97706"))):
            label_width = metrics.horizontalAdvance(label) + 8
            painter.setPen(QPen(color, 2))
            painter.drawLine(x, y, x + 22, y)
            painter.setBrush(color)
            painter.drawEllipse(QPoint(x + 11, y), 4, 4)
            painter.setPen(QColor("#374151"))
            painter.drawText(QRect(x + 30, y - metrics.height() // 2, label_width, metrics.height() + 2), Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter, label)
            x += label_width + 50


def overview_series(rows: tuple[DashboardPeriodRow, ...], ui_config: NimUiConfig = NIM_DN_UI_CONFIG) -> tuple[ChartSeries, ...]:
    series = [
        ChartSeries("NIM trước ĐC", tuple((row.period, row.nim_before) for row in rows), "percent"),
        ChartSeries("NIM sau ĐC", tuple((row.period, row.nim_after) for row in rows), "percent"),
    ]
    if ui_config.include_average_rate:
        series.append(ChartSeries("Lãi suất bình quân", tuple((row.period, row.average_rate) for row in rows), "percent"))
    return tuple(series)


def growth_series(rows: tuple[DashboardPeriodRow, ...], ui_config: NimUiConfig = NIM_DN_UI_CONFIG) -> tuple[ChartSeries, ...]:
    return (
        ChartSeries(ui_config.growth_percent_label, tuple((row.period, row.balance_growth_percent) for row in rows), "percent_signed"),
        ChartSeries("Biến động NIM trước ĐC", tuple((row.period, row.nim_before_delta) for row in rows), "percent_signed"),
        ChartSeries("Biến động NIM sau ĐC", tuple((row.period, row.nim_after_delta) for row in rows), "percent_signed"),
    )


def branch_bar_values(rows: tuple[DashboardBranchRow, ...], metric: str) -> tuple[tuple[str, str, float | None], ...]:
    current_rows = latest_branch_rows(rows)
    return tuple(
        sorted(
            ((row.period, row.branch, metric_value(row, metric)) for row in current_rows),
            key=lambda item: float("-inf") if item[2] is None else float(item[2]),
            reverse=True,
        )
    )


def branch_period_pair_values(
    rows: tuple[DashboardBranchRow, ...],
    metric: str,
    *,
    from_period: str = "",
    to_period: str = "",
) -> tuple[tuple[str, str, float | None, float | None], str, str]:
    period_from, period_to = _period_bounds(rows, from_period, to_period)
    from_rows = {row.branch_code or row.branch: row for row in rows if row.period == period_from}
    to_rows = {row.branch_code or row.branch: row for row in rows if row.period == period_to}
    keys = sorted(set(from_rows) | set(to_rows), key=lambda key: ((to_rows.get(key) or from_rows.get(key)).branch if (to_rows.get(key) or from_rows.get(key)) else "").casefold())
    pairs = []
    for key in keys:
        from_row = from_rows.get(key)
        to_row = to_rows.get(key)
        display_row = to_row or from_row
        if display_row is None:
            continue
        pairs.append((display_row.branch_code, display_row.branch, metric_value(from_row, metric) if from_row else None, metric_value(to_row, metric) if to_row else None))
    return tuple(pairs), period_from, period_to


def branch_empty_message(metric: str) -> str:
    if metric == METRIC_BALANCE_GROWTH:
        return "Chưa đủ dữ liệu lịch sử để hiển thị tăng trưởng."
    return "Không có dữ liệu chi nhánh."


def _period_bounds(rows: tuple[DashboardBranchRow, ...], from_period: str, to_period: str) -> tuple[str, str]:
    periods = sorted({row.period for row in rows})
    if not periods:
        return "", ""
    clean_from = from_period if from_period in periods else periods[0]
    clean_to = to_period if to_period in periods else periods[-1]
    return clean_from, clean_to


def _format_axis(value: float, value_kind: str) -> str:
    if value_kind.startswith("money"):
        if abs(value) >= 1_000_000_000:
            return f"{value / 1_000_000_000:,.0f} tỷ".replace(",", ".")
        return format_money_vn(value)
    if value_kind.startswith("percent"):
        return format_percent_vn(value, signed=value_kind.endswith("signed"))
    return f"{value:,.0f}"


def _format_value(value: float | None, value_kind: str) -> str:
    if value is None:
        return "N/A"
    if value_kind.startswith("money"):
        return format_money_vn(value, signed=value_kind.endswith("signed"))
    if value_kind.startswith("percent"):
        return format_percent_vn(value, signed=value_kind.endswith("signed"))
    return str(value)
