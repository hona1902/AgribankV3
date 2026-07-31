from __future__ import annotations

from PySide6.QtCore import QRect, Qt
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


def branch_empty_message(metric: str) -> str:
    if metric == METRIC_BALANCE_GROWTH:
        return "Chưa đủ dữ liệu lịch sử để hiển thị tăng trưởng."
    return "Không có dữ liệu chi nhánh."


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
