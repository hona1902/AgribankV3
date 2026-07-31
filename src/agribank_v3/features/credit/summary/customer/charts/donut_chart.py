from __future__ import annotations

from PySide6.QtCharts import QPieSeries

from agribank_v3.features.credit.summary.customer.charts.base_chart import BaseCustomerChart
from agribank_v3.features.credit.summary.customer.charts.chart_formatters import format_money_full
from agribank_v3.features.credit.summary.customer.charts.chart_tooltip import ChartTooltip


class CustomerDonutChart(BaseCustomerChart):
    def set_slices(self, slices: tuple[tuple[str, float], ...]) -> None:
        self.clear_chart()
        clean_slices = tuple((label, float(value or 0)) for label, value in slices if float(value or 0) > 0)
        self.last_series = (("Cơ cấu", clean_slices),)
        total = sum(value for _label, value in clean_slices)
        if total <= 0:
            self.set_empty("Không có dữ liệu dư nợ trong phạm vi đã chọn.")
            return
        series = QPieSeries()
        series.setHoleSize(0.44)
        for label, value in clean_slices:
            pie_slice = series.append(label, value)
            percent = value / total * 100
            pie_slice.setLabel(f"{label}: {percent:.2f}%")
            pie_slice.setLabelVisible(True)
            tooltip = ChartTooltip(
                label=label,
                series_name="Cơ cấu kỳ hạn",
                value=value,
                value_kind="money",
                detail=f"Tỷ trọng: {percent:.2f}%\nTổng: {format_money_full(total)}",
            )
            pie_slice.hovered.connect(lambda state, payload=tooltip: self._hover_slice(state, payload))
        self.chart.addSeries(series)
        self.chart.setTitle(f"{self.title}\nTổng dư nợ: {format_money_full(total)}")
        self.show_chart()

    def _hover_slice(self, state: bool, payload: ChartTooltip) -> None:
        if state:
            self.show_tooltip(payload, self.mapToGlobal(self.rect().center()))
