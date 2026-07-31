from __future__ import annotations

from dataclasses import dataclass
from typing import Iterable

from agribank_v3.features.credit.summary.dashboard_repository import NimDashboardRepository
from agribank_v3.features.credit.summary.models import SummaryDataType
from agribank_v3.features.credit.summary.nim_ui_config import NimUiConfig, get_nim_ui_config
from agribank_v3.features.credit.summary.officer_history.models import (
    METRIC_AVERAGE_RATE,
    METRIC_BALANCE,
    METRIC_BALANCE_GROWTH,
    METRIC_NIM_AFTER,
    METRIC_NIM_BEFORE,
)
from agribank_v3.features.credit.summary.officer_history.widgets import format_money_vn, format_percent_vn


METRIC_LABELS = {
    METRIC_BALANCE: "Dư nợ",
    METRIC_NIM_BEFORE: "NIM trước ĐC",
    METRIC_NIM_AFTER: "NIM sau ĐC",
    METRIC_AVERAGE_RATE: "Lãi suất bình quân",
    METRIC_BALANCE_GROWTH: "Tăng trưởng dư nợ",
}


@dataclass(frozen=True, slots=True)
class DashboardFilters:
    period_from: str = ""
    period_to: str = ""
    branch: str = ""
    transaction_office: str = ""
    customer_type: str = ""
    metric: str = METRIC_BALANCE

    def as_query_filters(self) -> dict[str, object]:
        return {
            "period_from": self.period_from,
            "period_to": self.period_to,
            "branch": self.branch,
            "transaction_office": self.transaction_office,
            "customer_type": self.customer_type,
        }


@dataclass(frozen=True, slots=True)
class DashboardPeriodRow:
    period: str
    balance: float
    average_rate: float
    nim_before: float
    nim_after: float
    balance_delta: float | None = None
    balance_growth_percent: float | None = None
    nim_before_delta: float | None = None
    nim_after_delta: float | None = None


@dataclass(frozen=True, slots=True)
class DashboardBranchRow:
    period: str
    branch: str
    balance: float
    average_rate: float
    nim_before: float
    nim_after: float
    balance_delta: float | None = None
    balance_growth_percent: float | None = None
    nim_before_delta: float | None = None
    nim_after_delta: float | None = None


@dataclass(frozen=True, slots=True)
class DashboardDetailRow:
    period: str
    branch: str
    transaction_office: str
    customer_type: str
    balance: float
    average_rate: float
    nim_before: float
    nim_after: float
    balance_delta: float | None = None
    balance_growth_percent: float | None = None
    nim_before_delta: float | None = None
    nim_after_delta: float | None = None


@dataclass(frozen=True, slots=True)
class DashboardKpi:
    label: str
    value: str


@dataclass(frozen=True, slots=True)
class DashboardNimData:
    data_type: SummaryDataType
    ui_config: NimUiConfig
    filters: DashboardFilters
    kpis: tuple[DashboardKpi, ...]
    period_rows: tuple[DashboardPeriodRow, ...]
    branch_rows: tuple[DashboardBranchRow, ...]
    detail_rows: tuple[DashboardDetailRow, ...]


def build_nim_dashboard(
    repository: NimDashboardRepository,
    data_type: SummaryDataType,
    filters: DashboardFilters,
) -> DashboardNimData:
    ui_config = get_nim_ui_config(data_type)
    query_filters = filters.as_query_filters()
    period_rows = _period_rows(repository.period_summary(data_type, query_filters))
    branch_rows = _branch_rows(repository.branch_period_summary(data_type, query_filters))
    detail_rows = _detail_rows(repository.detail_summary(data_type, query_filters))
    return DashboardNimData(
        data_type=data_type,
        ui_config=ui_config,
        filters=filters,
        kpis=_kpis(period_rows, branch_rows, ui_config),
        period_rows=period_rows,
        branch_rows=branch_rows,
        detail_rows=detail_rows,
    )


def metric_value(row: DashboardBranchRow | DashboardDetailRow | DashboardPeriodRow, metric: str) -> float | None:
    if metric == METRIC_BALANCE:
        return row.balance
    if metric == METRIC_NIM_BEFORE:
        return row.nim_before
    if metric == METRIC_NIM_AFTER:
        return row.nim_after
    if metric == METRIC_AVERAGE_RATE:
        return row.average_rate
    if metric == METRIC_BALANCE_GROWTH:
        return row.balance_growth_percent
    return row.balance


def metric_value_kind(metric: str) -> str:
    if metric == METRIC_BALANCE:
        return "money"
    if metric == METRIC_BALANCE_GROWTH:
        return "percent_signed"
    return "percent"


def format_metric_value(value: float | None, metric: str) -> str:
    if value is None:
        return "N/A"
    if metric == METRIC_BALANCE:
        return format_money_vn(value)
    if metric == METRIC_BALANCE_GROWTH:
        return format_percent_vn(value, signed=True)
    return format_percent_vn(value)


def metric_labels_for(data_type: SummaryDataType) -> dict[str, str]:
    return get_nim_ui_config(data_type).metric_labels()


def latest_branch_rows(rows: Iterable[DashboardBranchRow]) -> tuple[DashboardBranchRow, ...]:
    all_rows = tuple(rows)
    latest_period = max((row.period for row in all_rows), default="")
    if not latest_period:
        return ()
    return tuple(row for row in all_rows if row.period == latest_period)


def _period_rows(rows: list[dict[str, object]]) -> tuple[DashboardPeriodRow, ...]:
    parsed = [
        DashboardPeriodRow(
            period=str(row.get("period") or ""),
            balance=float(row.get("balance") or 0),
            average_rate=float(row.get("average_rate") or 0),
            nim_before=float(row.get("nim_before") or 0),
            nim_after=float(row.get("nim_after") or 0),
        )
        for row in rows
    ]
    return tuple(_with_period_growth(parsed))


def _branch_rows(rows: list[dict[str, object]]) -> tuple[DashboardBranchRow, ...]:
    parsed = [
        DashboardBranchRow(
            period=str(row.get("period") or ""),
            branch=str(row.get("branch") or ""),
            balance=float(row.get("balance") or 0),
            average_rate=float(row.get("average_rate") or 0),
            nim_before=float(row.get("nim_before") or 0),
            nim_after=float(row.get("nim_after") or 0),
        )
        for row in rows
    ]
    return tuple(_with_branch_growth(parsed))


def _detail_rows(rows: list[dict[str, object]]) -> tuple[DashboardDetailRow, ...]:
    parsed = [
        DashboardDetailRow(
            period=str(row.get("period") or ""),
            branch=str(row.get("branch") or ""),
            transaction_office=str(row.get("transaction_office") or ""),
            customer_type=str(row.get("customer_type") or "Tất cả"),
            balance=float(row.get("balance") or 0),
            average_rate=float(row.get("average_rate") or 0),
            nim_before=float(row.get("nim_before") or 0),
            nim_after=float(row.get("nim_after") or 0),
        )
        for row in rows
    ]
    return tuple(_with_detail_growth(parsed))


def _with_period_growth(rows: list[DashboardPeriodRow]) -> list[DashboardPeriodRow]:
    output: list[DashboardPeriodRow] = []
    previous: DashboardPeriodRow | None = None
    for row in sorted(rows, key=lambda item: item.period):
        delta, growth = _growth(row.balance, previous.balance if previous else None)
        nim_before_delta = None if previous is None else row.nim_before - previous.nim_before
        nim_after_delta = None if previous is None else row.nim_after - previous.nim_after
        output.append(
            DashboardPeriodRow(
                period=row.period,
                balance=row.balance,
                average_rate=row.average_rate,
                nim_before=row.nim_before,
                nim_after=row.nim_after,
                balance_delta=delta,
                balance_growth_percent=growth,
                nim_before_delta=nim_before_delta,
                nim_after_delta=nim_after_delta,
            )
        )
        previous = row
    return output


def _with_branch_growth(rows: list[DashboardBranchRow]) -> list[DashboardBranchRow]:
    grouped: dict[str, list[DashboardBranchRow]] = {}
    for row in rows:
        grouped.setdefault(row.branch, []).append(row)
    output: list[DashboardBranchRow] = []
    for branch, branch_rows in grouped.items():
        previous: DashboardBranchRow | None = None
        for row in sorted(branch_rows, key=lambda item: item.period):
            delta, growth = _growth(row.balance, previous.balance if previous else None)
            nim_before_delta = None if previous is None else row.nim_before - previous.nim_before
            nim_after_delta = None if previous is None else row.nim_after - previous.nim_after
            output.append(
                DashboardBranchRow(
                    period=row.period,
                    branch=branch,
                    balance=row.balance,
                    average_rate=row.average_rate,
                    nim_before=row.nim_before,
                    nim_after=row.nim_after,
                    balance_delta=delta,
                    balance_growth_percent=growth,
                    nim_before_delta=nim_before_delta,
                    nim_after_delta=nim_after_delta,
                )
            )
            previous = row
    return sorted(output, key=lambda item: (item.period, item.branch.casefold()))


def _with_detail_growth(rows: list[DashboardDetailRow]) -> list[DashboardDetailRow]:
    grouped: dict[tuple[str, str, str], list[DashboardDetailRow]] = {}
    for row in rows:
        grouped.setdefault((row.branch, row.transaction_office, row.customer_type), []).append(row)
    output: list[DashboardDetailRow] = []
    for key, detail_rows in grouped.items():
        previous: DashboardDetailRow | None = None
        for row in sorted(detail_rows, key=lambda item: item.period):
            delta, growth = _growth(row.balance, previous.balance if previous else None)
            nim_before_delta = None if previous is None else row.nim_before - previous.nim_before
            nim_after_delta = None if previous is None else row.nim_after - previous.nim_after
            output.append(
                DashboardDetailRow(
                    period=row.period,
                    branch=key[0],
                    transaction_office=key[1],
                    customer_type=key[2],
                    balance=row.balance,
                    average_rate=row.average_rate,
                    nim_before=row.nim_before,
                    nim_after=row.nim_after,
                    balance_delta=delta,
                    balance_growth_percent=growth,
                    nim_before_delta=nim_before_delta,
                    nim_after_delta=nim_after_delta,
                )
            )
            previous = row
    return sorted(output, key=lambda item: (item.period, item.branch.casefold(), item.transaction_office.casefold(), item.customer_type.casefold()))


def _growth(current: float, previous: float | None) -> tuple[float | None, float | None]:
    if previous is None:
        return None, None
    delta = current - previous
    if previous == 0:
        return delta, None
    return delta, (delta / previous) * 100


def _kpis(
    period_rows: tuple[DashboardPeriodRow, ...],
    branch_rows: tuple[DashboardBranchRow, ...],
    ui_config: NimUiConfig,
) -> tuple[DashboardKpi, ...]:
    current = period_rows[-1] if period_rows else None
    latest_period = current.period if current else ""
    branch_count = len({row.branch for row in branch_rows if row.period == latest_period}) if latest_period else 0
    metrics = [
        DashboardKpi("Số chi nhánh", str(branch_count)),
        DashboardKpi(ui_config.total_balance_label, format_money_vn(current.balance if current else 0)),
        DashboardKpi("NIM trước ĐC bình quân", format_percent_vn(current.nim_before if current else 0)),
        DashboardKpi("NIM sau ĐC bình quân", format_percent_vn(current.nim_after if current else 0)),
        DashboardKpi(
            f"{ui_config.growth_label} kỳ gần nhất",
            "N/A" if current is None or current.balance_growth_percent is None else format_percent_vn(current.balance_growth_percent, signed=True),
        ),
    ]
    if ui_config.include_average_rate:
        metrics.insert(4, DashboardKpi("Lãi suất bình quân", format_percent_vn(current.average_rate if current else 0)))
    return tuple(metrics)
