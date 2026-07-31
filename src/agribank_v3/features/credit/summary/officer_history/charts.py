from __future__ import annotations

from decimal import Decimal, InvalidOperation, ROUND_HALF_UP

from PySide6.QtCore import QPoint, QRect, Qt
from PySide6.QtGui import QColor, QPainter, QPen
from PySide6.QtWidgets import QToolTip, QWidget

from .models import ChartSeries


SERIES_COLORS = (
    QColor("#1f6feb"),
    QColor("#d97706"),
    QColor("#059669"),
    QColor("#7c3aed"),
    QColor("#dc2626"),
    QColor("#2563eb"),
    QColor("#0f766e"),
    QColor("#9333ea"),
    QColor("#be123c"),
    QColor("#0891b2"),
    QColor("#65a30d"),
    QColor("#c2410c"),
    QColor("#4f46e5"),
    QColor("#0d9488"),
    QColor("#b45309"),
    QColor("#db2777"),
    QColor("#16a34a"),
    QColor("#0284c7"),
    QColor("#a21caf"),
    QColor("#ea580c"),
    QColor("#475569"),
    QColor("#15803d"),
    QColor("#7e22ce"),
    QColor("#b91c1c"),
)


class AnalysisLineChart(QWidget):
    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.series: tuple[ChartSeries, ...] = ()
        self.empty_message = "Không có dữ liệu."
        self.single_point_message = ""
        self.point_rects: list[tuple[QRect, str, str, str, float | None, str]] = []
        self.zoom_factor = 1.0
        self.setMinimumHeight(240)
        self.setMouseTracking(True)

    def set_series(
        self,
        series: tuple[ChartSeries, ...],
        *,
        empty_message: str = "Không có dữ liệu.",
        single_point_message: str = "",
    ) -> None:
        self.series = series
        self.empty_message = empty_message
        self.single_point_message = single_point_message
        self.zoom_factor = 1.0
        self.update()

    def paintEvent(self, event) -> None:
        super().paintEvent(event)
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        painter.fillRect(self.rect(), QColor("#ffffff"))
        rect = self.rect().adjusted(58, 16, -18, -42)
        painter.setPen(QColor("#d8dee8"))
        painter.drawRect(rect)
        self.point_rects = []
        periods = self._visible_periods()
        if not periods:
            self._draw_center_text(painter, self.empty_message)
            return
        values = [
            value
            for item in self.series
            for period, value in item.values
            if period in periods and value is not None
        ]
        if not values:
            self._draw_center_text(painter, self.empty_message)
            return
        min_value = min(values)
        max_value = max(values)
        if min_value == max_value:
            min_value -= 1
            max_value += 1
        padding = (max_value - min_value) * 0.08
        min_value -= padding
        max_value += padding
        self._draw_grid(painter, rect, min_value, max_value, self.series[0].value_kind)
        for index, item in enumerate(self.series):
            self._draw_series(painter, rect, periods, item, min_value, max_value, _series_color(item.label, index))
        self._draw_period_labels(painter, rect, periods)
        self._draw_legends(painter)
        if len(periods) == 1 and self.single_point_message:
            painter.setPen(QColor("#6b7280"))
            painter.drawText(QRect(rect.left(), rect.top() + 8, rect.width(), 20), Qt.AlignmentFlag.AlignCenter, self.single_point_message)

    def wheelEvent(self, event) -> None:
        periods = self._all_periods()
        if len(periods) < 3:
            return
        delta = event.angleDelta().y()
        if delta > 0:
            self.zoom_factor = min(6.0, self.zoom_factor * 1.2)
        elif delta < 0:
            self.zoom_factor = max(1.0, self.zoom_factor / 1.2)
        self.update()
        event.accept()

    def mouseMoveEvent(self, event) -> None:
        position = event.position().toPoint()
        for rect, period, label, metric_label, value, value_kind in self.point_rects:
            if rect.contains(position):
                QToolTip.showText(
                    event.globalPosition().toPoint(),
                    _tooltip_text(period, label, metric_label, value, value_kind),
                    self,
                )
                return
        QToolTip.hideText()
        super().mouseMoveEvent(event)

    def _draw_grid(self, painter: QPainter, rect: QRect, min_value: float, max_value: float, value_kind: str) -> None:
        painter.setPen(QColor("#6b7280"))
        for index in range(5):
            ratio = index / 4
            y = rect.bottom() - int(rect.height() * ratio)
            value = min_value + (max_value - min_value) * ratio
            painter.drawLine(rect.left(), y, rect.right(), y)
            painter.drawText(QRect(0, y - 9, 54, 18), Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter, _format_value(value, value_kind, compact=True))

    def _draw_series(
        self,
        painter: QPainter,
        rect: QRect,
        periods: tuple[str, ...],
        series: ChartSeries,
        min_value: float,
        max_value: float,
        color: QColor,
    ) -> None:
        value_by_period = {period: value for period, value in series.values}
        points: list[tuple[str, float, QPoint]] = []
        for index, period in enumerate(periods):
            value = value_by_period.get(period)
            if value is None:
                continue
            x = rect.left() + int(rect.width() * index / max(1, len(periods) - 1))
            y = rect.bottom() - int(((float(value) - min_value) / (max_value - min_value)) * rect.height())
            points.append((period, float(value), QPoint(x, y)))
        painter.setPen(QPen(color, 2))
        for index in range(1, len(points)):
            painter.drawLine(points[index - 1][2], points[index][2])
        painter.setBrush(color)
        painter.setPen(QPen(color, 1))
        for period, value, point in points:
            painter.drawEllipse(point, 4, 4)
            self.point_rects.append((QRect(point.x() - 8, point.y() - 8, 16, 16), period, series.label, series.tooltip_metric, value, series.value_kind))

    def _draw_period_labels(self, painter: QPainter, rect: QRect, periods: tuple[str, ...]) -> None:
        painter.setPen(QColor("#374151"))
        step = max(1, len(periods) // 8)
        for index, period in enumerate(periods):
            if index % step == 0 or index == len(periods) - 1:
                x = rect.left() + int(rect.width() * index / max(1, len(periods) - 1))
                painter.drawText(QRect(x - 38, rect.bottom() + 6, 76, 18), Qt.AlignmentFlag.AlignCenter, period)

    def _draw_legends(self, painter: QPainter) -> None:
        x = 58
        y = self.height() - 20
        for index, item in enumerate(self.series):
            color = _series_color(item.label, index)
            painter.setPen(QPen(color, 2))
            painter.drawLine(x, y, x + 22, y)
            painter.setBrush(color)
            painter.drawEllipse(QPoint(x + 11, y), 4, 4)
            painter.setPen(QColor("#374151"))
            label_width = min(170, max(76, len(item.label) * 7))
            painter.drawText(QRect(x + 28, y - 9, label_width, 18), Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter, item.label)
            x += label_width + 38
            if x > self.width() - 180:
                break

    def _draw_center_text(self, painter: QPainter, text: str) -> None:
        painter.setPen(QColor("#6b7280"))
        painter.drawText(self.rect(), Qt.AlignmentFlag.AlignCenter, text)

    def _visible_periods(self) -> tuple[str, ...]:
        periods = self._all_periods()
        if self.zoom_factor <= 1.01:
            return periods
        count = max(2, int(round(len(periods) / self.zoom_factor)))
        return periods[-count:]

    def _all_periods(self) -> tuple[str, ...]:
        values = sorted({period for item in self.series for period, _value in item.values})
        return tuple(values)


def _format_value(value: float | None, value_kind: str, *, compact: bool = False) -> str:
    if value is None:
        return "N/A"
    if value_kind.startswith("money"):
        text = _format_money_vn(value)
        if compact and abs(float(value)) >= 1_000_000_000:
            return f"{float(value) / 1_000_000_000:,.0f} tỷ".replace(",", ".")
        return text
    if value_kind.startswith("percent"):
        return _format_percent_vn(value, signed=value_kind.endswith("signed"))
    return str(value)


def _series_color(label: str, fallback_index: int) -> QColor:
    if label:
        checksum = sum((index + 1) * ord(char) for index, char in enumerate(label))
        return SERIES_COLORS[checksum % len(SERIES_COLORS)]
    return SERIES_COLORS[fallback_index % len(SERIES_COLORS)]


def _tooltip_text(period: str, label: str, metric_label: str, value: float | None, value_kind: str) -> str:
    if metric_label:
        return f"{period}\nCán bộ: {label}\n{metric_label}: {_format_value(value, value_kind)}"
    return f"{period}\n{label}: {_format_value(value, value_kind)}"


def _format_money_vn(value: object, *, signed: bool = False) -> str:
    try:
        number = int(round(float(value or 0)))
    except (TypeError, ValueError):
        number = 0
    prefix = "+" if signed and number > 0 else ""
    return prefix + f"{number:,}".replace(",", ".")


def _format_percent_vn(value: object, *, signed: bool = False) -> str:
    try:
        number = Decimal(str(value or 0)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    except (InvalidOperation, TypeError, ValueError):
        number = Decimal("0.00")
    prefix = "+" if signed and number > 0 else ""
    text = f"{number:,.2f}%"
    return prefix + text.replace(",", "_").replace(".", ",").replace("_", ".")
