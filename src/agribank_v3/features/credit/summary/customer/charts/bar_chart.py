from __future__ import annotations

from PySide6.QtCharts import QBarCategoryAxis, QBarSeries, QBarSet, QValueAxis
from PySide6.QtCore import Qt

from agribank_v3.features.credit.summary.customer.charts.base_chart import BaseCustomerChart
from agribank_v3.features.credit.summary.customer.charts.chart_formatters import format_period, money_axis_scale
from agribank_v3.features.credit.summary.customer.charts.chart_tooltip import ChartTooltip


class CustomerBarChart(BaseCustomerChart):
    def set_series(self, series: tuple[tuple[str, tuple[tuple[str, float], ...]], ...]) -> None:
        self.last_series = series
        self.clear_chart()
        labels = sorted({str(label) for _name, rows in series for label, _value in rows})
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
        bar_series = QBarSeries()
        payloads: dict[tuple[str, int], ChartTooltip] = {}
        max_value = 0.0
        for name, rows in series:
            values_by_label = {str(label): float(value or 0) for label, value in rows}
            bar_set = QBarSet(name)
            for index, label in enumerate(labels):
                raw_value = values_by_label.get(label, 0.0)
                bar_set.append(raw_value / divisor)
                max_value = max(max_value, abs(raw_value / divisor))
                payloads[(name, index)] = ChartTooltip(label=label, series_name=name, value=raw_value, value_kind=self.value_kind)
            bar_series.append(bar_set)
        bar_series.hovered.connect(lambda state, index, barset: self._hover_bar(state, index, barset, payloads))
        self.chart.addSeries(bar_series)
        axis_x = QBarCategoryAxis()
        axis_x.append([format_period(label) for label in labels])
        axis_y = QValueAxis()
        axis_y.setRange(0, max(1.0, max_value * 1.1))
        axis_y.setTitleText("%" if self.value_kind.startswith("percent") else unit)
        self.chart.addAxis(axis_x, Qt.AlignmentFlag.AlignBottom)
        self.chart.addAxis(axis_y, Qt.AlignmentFlag.AlignLeft)
        bar_series.attachAxis(axis_x)
        bar_series.attachAxis(axis_y)
        self.show_chart()

    def _hover_bar(self, state: bool, index: int, barset: QBarSet, payloads: dict[tuple[str, int], ChartTooltip]) -> None:
        if not state:
            return
        payload = payloads.get((barset.label(), index))
        if payload:
            self.show_tooltip(payload, self.mapToGlobal(self.rect().center()))
