from __future__ import annotations

from PySide6.QtCharts import QBarCategoryAxis, QBarSet, QHorizontalBarSeries, QValueAxis
from PySide6.QtCore import Qt

from agribank_v3.features.credit.summary.customer.charts.base_chart import BaseCustomerChart
from agribank_v3.features.credit.summary.customer.charts.chart_formatters import (
    format_customer_label,
    money_axis_scale,
)
from agribank_v3.features.credit.summary.customer.charts.chart_tooltip import ChartTooltip


class CustomerHorizontalBarChart(BaseCustomerChart):
    def set_rows(
        self,
        rows: tuple[dict[str, object], ...] | list[dict[str, object]],
        *,
        label_field: str = "customer_name",
        value_field: str = "total_balance",
        series_name: str = "Giá trị",
        tooltip_fields: tuple[str, ...] = ("customer_code", "customer_name", "effective_officer_name", "branch_code"),
    ) -> None:
        self.clear_chart()
        clean_rows = [dict(row) for row in rows if row.get(value_field) is not None]
        self.last_series = ((series_name, tuple((str(row.get(label_field) or ""), float(row.get(value_field) or 0)) for row in clean_rows)),)
        if not clean_rows:
            self.set_empty()
            return
        values = [float(row.get(value_field) or 0) for row in clean_rows]
        if self.value_kind.startswith("money"):
            divisor, unit = money_axis_scale(values)
        elif self.value_kind.startswith("number"):
            divisor, unit = 1.0, "khách hàng"
        else:
            divisor, unit = 1.0, "%"
        categories = [
            format_customer_label(row.get("customer_code"), row.get(label_field), max_length=34)
            for row in clean_rows
        ]
        bar_set = QBarSet(series_name)
        payloads: dict[int, ChartTooltip] = {}
        max_value = 0.0
        for index, row in enumerate(clean_rows):
            raw_value = float(row.get(value_field) or 0)
            bar_set.append(abs(raw_value) / divisor)
            max_value = max(max_value, abs(raw_value) / divisor)
            detail = "\n".join(
                str(row.get(field) or "")
                for field in tooltip_fields
                if str(row.get(field) or "").strip()
            )
            payloads[index] = ChartTooltip(
                label=str(row.get(label_field) or ""),
                series_name=series_name,
                value=raw_value,
                value_kind=self.value_kind,
                detail=detail,
            )
        series = QHorizontalBarSeries()
        series.append(bar_set)
        series.hovered.connect(lambda state, index, _barset: self._hover_bar(state, index, payloads))
        self.chart.addSeries(series)
        axis_y = QBarCategoryAxis()
        axis_y.append(categories)
        axis_x = QValueAxis()
        axis_x.setRange(0, max(1.0, max_value * 1.1))
        axis_x.setTitleText("%" if self.value_kind.startswith("percent") else unit)
        self.chart.addAxis(axis_y, Qt.AlignmentFlag.AlignLeft)
        self.chart.addAxis(axis_x, Qt.AlignmentFlag.AlignBottom)
        series.attachAxis(axis_y)
        series.attachAxis(axis_x)
        self.chart_view.setMinimumHeight(max(300, 110 + len(clean_rows) * 24))
        self.show_chart()

    def _hover_bar(self, state: bool, index: int, payloads: dict[int, ChartTooltip]) -> None:
        if state and index in payloads:
            self.show_tooltip(payloads[index], self.mapToGlobal(self.rect().center()))
