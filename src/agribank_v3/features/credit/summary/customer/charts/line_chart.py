from __future__ import annotations

from PySide6.QtCharts import QCategoryAxis, QLineSeries, QScatterSeries, QValueAxis
from PySide6.QtCore import QPointF, Qt

from agribank_v3.features.credit.summary.customer.charts.base_chart import (
    CHART_COLORS,
    BaseCustomerChart,
)
from agribank_v3.features.credit.summary.customer.charts.chart_formatters import (
    format_chart_value,
    format_period,
    money_axis_scale,
)
from agribank_v3.features.credit.summary.customer.charts.chart_tooltip import ChartTooltip


class CustomerLineChart(BaseCustomerChart):
    def set_series(self, series: tuple[tuple[str, tuple[tuple[str, float], ...]], ...]) -> None:
        self.last_series = series
        self.clear_chart()
        labels = _ordered_labels(series)
        if not series or not labels:
            self.set_empty()
            return
        raw_values = [float(value or 0) for _name, rows in series for _label, value in rows]
        if self.value_kind.startswith("money"):
            divisor, unit = money_axis_scale(raw_values)
        elif self.value_kind.startswith("number"):
            divisor, unit = 1.0, "khách hàng"
        else:
            divisor, unit = 1.0, "%"
        scaled_values = [value / divisor for value in raw_values]
        min_value = min(0.0, *scaled_values)
        max_value = max(0.0, *scaled_values)
        if min_value == max_value:
            max_value = min_value + 1.0
        payloads: dict[tuple[str, int], ChartTooltip] = {}
        for series_index, (name, rows) in enumerate(series):
            values_by_label = {label: float(value or 0) for label, value in rows}
            line = QLineSeries()
            line.setName(name)
            marker = QScatterSeries()
            marker.setName(name)
            marker.setMarkerSize(8.0)
            color = CHART_COLORS[series_index % len(CHART_COLORS)]
            line.setColor(color)
            marker.setColor(color)
            for index, label in enumerate(labels):
                raw_value = values_by_label.get(label)
                if raw_value is None:
                    continue
                point = QPointF(float(index), raw_value / divisor)
                line.append(point)
                marker.append(point)
                payloads[(name, index)] = ChartTooltip(
                    label=label,
                    series_name=name,
                    value=raw_value,
                    value_kind=self.value_kind,
                    detail=f"Đơn vị trục: {unit}",
                )
            self.chart.addSeries(line)
            self.chart.addSeries(marker)
            for legend_marker in self.chart.legend().markers(marker):
                legend_marker.setVisible(False)
            line.hovered.connect(lambda point, state, s=name: self._hover_line(point, state, s, payloads))
            marker.hovered.connect(lambda point, state, s=name: self._hover_line(point, state, s, payloads))
        self.chart.legend().setVisible(len(series) > 1)
        axis_x = QCategoryAxis()
        axis_x.setLabelsPosition(QCategoryAxis.AxisLabelsPosition.AxisLabelsPositionOnValue)
        step = max(1, len(labels) // 8)
        for index, label in enumerate(labels):
            if index % step == 0 or index == len(labels) - 1:
                axis_x.append(format_period(label), float(index))
        axis_x.setRange(0, max(0, len(labels) - 1))
        axis_y = QValueAxis()
        axis_y.setRange(min_value, max_value)
        if self.value_kind.startswith("number"):
            axis_y.setLabelFormat("%d")
        else:
            axis_y.setLabelFormat("%.2f")
        axis_y.setTitleText("%" if self.value_kind.startswith("percent") else unit)
        self.chart.addAxis(axis_x, Qt.AlignmentFlag.AlignBottom)
        self.chart.addAxis(axis_y, Qt.AlignmentFlag.AlignLeft)
        for item in self.chart.series():
            item.attachAxis(axis_x)
            item.attachAxis(axis_y)
        self.show_chart()

    def _hover_line(
        self,
        point: QPointF,
        state: bool,
        series_name: str,
        payloads: dict[tuple[str, int], ChartTooltip],
    ) -> None:
        if not state:
            return
        payload = payloads.get((series_name, int(round(point.x()))))
        if payload:
            self.show_tooltip(payload, self.mapToGlobal(self.rect().center()))


def _ordered_labels(series: tuple[tuple[str, tuple[tuple[str, float], ...]], ...]) -> list[str]:
    labels = {str(label) for _name, rows in series for label, _value in rows}
    return sorted(labels)
