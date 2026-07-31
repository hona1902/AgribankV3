from __future__ import annotations

from dataclasses import dataclass

from agribank_v3.features.credit.summary.models import SummaryDataType


METRIC_BALANCE = "balance"
METRIC_AVERAGE_RATE = "average_rate"
METRIC_NIM_BEFORE = "nim_before"
METRIC_NIM_AFTER = "nim_after"
METRIC_BALANCE_GROWTH = "balance_growth"


@dataclass(frozen=True, slots=True)
class OfficerKey:
    code: str
    raw_name: str
    display_name: str
    branch: str = ""
    transaction_office: str = ""


@dataclass(frozen=True, slots=True)
class HistoryFilters:
    period_from: str = ""
    period_to: str = ""
    customer_type: str = ""
    transaction_office: str = ""


@dataclass(frozen=True, slots=True)
class HistoryPoint:
    period: str
    balance: float
    average_rate: float
    nim_before: float
    nim_after: float


@dataclass(frozen=True, slots=True)
class OfficerOverview:
    data_type: SummaryDataType
    officer: OfficerKey
    branch: str
    transaction_office: str
    customer_type: str
    current_period: str
    current_balance: float
    current_average_rate: float
    current_nim_before: float
    current_nim_after: float
    points: tuple[HistoryPoint, ...]


@dataclass(frozen=True, slots=True)
class GrowthPoint:
    period: str
    balance: float
    delta: float | None
    growth_percent: float | None
    nim_before: float
    nim_after: float


@dataclass(frozen=True, slots=True)
class ComparisonRow:
    period: str
    officer: OfficerKey
    branch: str
    transaction_office: str
    customer_type: str
    metric: str
    value: float | None


@dataclass(frozen=True, slots=True)
class ChartSeries:
    label: str
    values: tuple[tuple[str, float | None], ...]
    value_kind: str
    tooltip_metric: str = ""
