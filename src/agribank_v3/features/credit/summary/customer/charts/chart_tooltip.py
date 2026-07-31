from __future__ import annotations

from dataclasses import dataclass

from agribank_v3.features.credit.summary.customer.charts.chart_formatters import (
    format_chart_value,
    format_period,
)


@dataclass(frozen=True, slots=True)
class ChartTooltip:
    label: str
    series_name: str
    value: float
    value_kind: str = "money"
    detail: str = ""

    def text(self) -> str:
        parts = [
            format_period(self.label),
            f"{self.series_name}: {format_chart_value(self.value, self.value_kind, full=True)}",
        ]
        if self.detail:
            parts.append(self.detail)
        return "\n".join(parts)
