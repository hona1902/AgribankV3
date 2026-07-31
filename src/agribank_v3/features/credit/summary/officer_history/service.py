from __future__ import annotations

from collections import defaultdict

from agribank_v3.features.credit.summary.models import SummaryDataType
from agribank_v3.features.credit.summary.nim_ui_config import NimUiConfig, NIM_DN_UI_CONFIG, get_nim_ui_config
from agribank_v3.features.credit.summary.repository import SummaryRepository

from .models import (
    METRIC_AVERAGE_RATE,
    METRIC_BALANCE,
    METRIC_BALANCE_GROWTH,
    METRIC_NIM_AFTER,
    METRIC_NIM_BEFORE,
    ChartSeries,
    ComparisonRow,
    GrowthPoint,
    HistoryFilters,
    HistoryPoint,
    OfficerKey,
    OfficerOverview,
)
from .repository import OfficerHistoryRepository, officer_key


METRIC_LABELS = {
    METRIC_BALANCE: "Dư nợ",
    METRIC_AVERAGE_RATE: "Lãi suất bình quân",
    METRIC_NIM_BEFORE: "NIM trước ĐC",
    METRIC_NIM_AFTER: "NIM sau ĐC",
    METRIC_BALANCE_GROWTH: "Tăng trưởng dư nợ",
}


def build_officer_overview(
    repository: SummaryRepository,
    data_type: SummaryDataType,
    *,
    officer_code: str = "",
    officer: str = "",
    branch: str = "",
    transaction_office: str = "",
    customer_type: str = "",
    period_from: str = "",
    period_to: str = "",
) -> OfficerOverview:
    filters = HistoryFilters(
        period_from=period_from,
        period_to=period_to,
        customer_type="" if customer_type == "Tất cả" else customer_type,
        transaction_office=transaction_office,
    )
    history_repository = OfficerHistoryRepository(repository)
    rows = history_repository.get_officer_history(
        data_type,
        officer_code=officer_code,
        officer=officer,
        branch=branch,
        filters=filters,
    )
    points = tuple(_history_point(row) for row in rows)
    current = points[-1] if points else None
    key = officer_key(officer)
    if officer_code and not key.code:
        key = OfficerKey(officer_code, officer, key.display_name, branch, transaction_office)
    return OfficerOverview(
        data_type=data_type,
        officer=key,
        branch=branch,
        transaction_office=transaction_office,
        customer_type=filters.customer_type,
        current_period=current.period if current else "",
        current_balance=current.balance if current else 0.0,
        current_average_rate=current.average_rate if current else 0.0,
        current_nim_before=current.nim_before if current else 0.0,
        current_nim_after=current.nim_after if current else 0.0,
        points=points,
    )


def build_officer_growth_history(points: tuple[HistoryPoint, ...]) -> tuple[GrowthPoint, ...]:
    growth: list[GrowthPoint] = []
    previous: HistoryPoint | None = None
    for point in points:
        delta: float | None = None
        growth_percent: float | None = None
        if previous is not None:
            delta = point.balance - previous.balance
            if previous.balance != 0:
                growth_percent = (delta / previous.balance) * 100
        growth.append(
            GrowthPoint(
                period=point.period,
                balance=point.balance,
                delta=delta,
                growth_percent=growth_percent,
                nim_before=point.nim_before,
                nim_after=point.nim_after,
            )
        )
        previous = point
    return tuple(growth)


def build_multiple_officer_comparison(
    repository: SummaryRepository,
    data_type: SummaryDataType,
    *,
    officers: list[OfficerKey],
    metric: str,
    branch: str = "",
    filters: HistoryFilters | None = None,
) -> tuple[ComparisonRow, ...]:
    filters = filters or HistoryFilters()
    rows = OfficerHistoryRepository(repository).get_multiple_officer_history(
        data_type,
        officers=officers,
        branch=branch,
        filters=filters,
    )
    histories: dict[str, list[dict[str, object]]] = defaultdict(list)
    for row in rows:
        histories[str(row.get("officer") or "")].append(row)
    output: list[ComparisonRow] = []
    for raw_officer, officer_rows in histories.items():
        key = officer_key(raw_officer)
        points = [_history_point(row) for row in officer_rows]
        growth_by_period = {point.period: point for point in build_officer_growth_history(tuple(points))}
        for row, point in zip(officer_rows, points, strict=False):
            value = _metric_value(point, metric, growth_by_period.get(point.period))
            output.append(
                ComparisonRow(
                    period=point.period,
                    officer=key,
                    branch=str(row.get("branch") or ""),
                    transaction_office=str(row.get("transaction_office") or ""),
                    customer_type=str(row.get("customer_type") or "Tất cả"),
                    metric=metric,
                    value=value,
                )
            )
    return tuple(sorted(output, key=lambda item: (item.period, item.officer.display_name, item.officer.code)))


def build_officer_branch_comparison(
    repository: SummaryRepository,
    data_type: SummaryDataType,
    *,
    officer_code: str = "",
    officer: str = "",
    branch: str,
    metric: str,
    filters: HistoryFilters | None = None,
) -> tuple[ChartSeries, tuple[ComparisonRow, ...]]:
    filters = filters or HistoryFilters()
    overview = build_officer_overview(
        repository,
        data_type,
        officer_code=officer_code,
        officer=officer,
        branch=branch,
        transaction_office=filters.transaction_office,
        customer_type=filters.customer_type,
        period_from=filters.period_from,
        period_to=filters.period_to,
    )
    branch_rows = OfficerHistoryRepository(repository).get_branch_history(data_type, branch=branch, filters=filters)
    branch_points = tuple(_history_point(row) for row in branch_rows)
    officer_growth = {point.period: point for point in build_officer_growth_history(overview.points)}
    branch_growth = {point.period: point for point in build_officer_growth_history(branch_points)}
    branch_key = OfficerKey("", branch, branch)
    rows: list[ComparisonRow] = []
    for point in overview.points:
        rows.append(
            ComparisonRow(
                period=point.period,
                officer=overview.officer,
                branch=branch,
                transaction_office=filters.transaction_office,
                customer_type=filters.customer_type or "Tất cả",
                metric=metric,
                value=_metric_value(point, metric, officer_growth.get(point.period)),
            )
        )
    for point in branch_points:
        rows.append(
            ComparisonRow(
                period=point.period,
                officer=branch_key,
                branch=branch,
                transaction_office=filters.transaction_office,
                customer_type=filters.customer_type or "Tất cả",
                metric=metric,
                value=_metric_value(point, metric, branch_growth.get(point.period)),
            )
        )
    branch_label = f"Chi nhánh: {branch}" if metric == METRIC_BALANCE else f"Bình quân chi nhánh: {branch}"
    series = (
        ChartSeries(
            label=f"CBTD: {overview.officer.display_name}",
            values=tuple((row.period, row.value) for row in rows if row.officer.raw_name == overview.officer.raw_name),
            value_kind=_metric_value_kind(metric),
        ),
        ChartSeries(
            label=branch_label,
            values=tuple((row.period, row.value) for row in rows if row.officer.raw_name == branch),
            value_kind=_metric_value_kind(metric),
        ),
    )
    return series, tuple(rows)


def overview_series(points: tuple[HistoryPoint, ...], ui_config: NimUiConfig = NIM_DN_UI_CONFIG) -> tuple[ChartSeries, ...]:
    series = [
        ChartSeries("NIM trước ĐC", tuple((point.period, point.nim_before) for point in points), "percent"),
        ChartSeries("NIM sau ĐC", tuple((point.period, point.nim_after) for point in points), "percent"),
    ]
    if ui_config.include_average_rate:
        series.append(ChartSeries("Lãi suất bình quân", tuple((point.period, point.average_rate) for point in points), "percent"))
    return tuple(series)


def balance_series(points: tuple[HistoryPoint, ...], ui_config: NimUiConfig = NIM_DN_UI_CONFIG) -> tuple[ChartSeries, ...]:
    return (ChartSeries(ui_config.balance_label, tuple((point.period, point.balance) for point in points), "money"),)


def growth_series(growth: tuple[GrowthPoint, ...], ui_config: NimUiConfig = NIM_DN_UI_CONFIG) -> tuple[ChartSeries, ...]:
    return (
        ChartSeries(ui_config.balance_delta_label, tuple((point.period, point.delta) for point in growth), "money_signed"),
        ChartSeries(ui_config.growth_percent_label, tuple((point.period, point.growth_percent) for point in growth), "percent_signed"),
    )


def metric_labels(data_type: SummaryDataType) -> dict[str, str]:
    return get_nim_ui_config(data_type).metric_labels()


def comparison_series(rows: tuple[ComparisonRow, ...], metric: str, *, metric_label: str = "") -> tuple[ChartSeries, ...]:
    grouped: dict[str, tuple[str, list[tuple[str, float | None]]]] = {}
    for row in rows:
        key = row.officer.raw_name or row.officer.display_name
        if key not in grouped:
            grouped[key] = (row.officer.display_name, [])
        grouped[key][1].append((row.period, row.value))
    return tuple(
        ChartSeries(label, tuple(values), _metric_value_kind(metric), metric_label)
        for label, values in grouped.values()
    )


def _history_point(row: dict[str, object]) -> HistoryPoint:
    return HistoryPoint(
        period=str(row.get("period") or ""),
        balance=float(row.get("balance") or 0),
        average_rate=float(row.get("average_rate") or 0),
        nim_before=float(row.get("nim_before") or 0),
        nim_after=float(row.get("nim_after") or 0),
    )


def _metric_value(point: HistoryPoint, metric: str, growth: GrowthPoint | None = None) -> float | None:
    if metric == METRIC_BALANCE:
        return point.balance
    if metric == METRIC_AVERAGE_RATE:
        return point.average_rate
    if metric == METRIC_NIM_BEFORE:
        return point.nim_before
    if metric == METRIC_NIM_AFTER:
        return point.nim_after
    if metric == METRIC_BALANCE_GROWTH:
        return None if growth is None else growth.growth_percent
    return point.nim_after


def _metric_value_kind(metric: str) -> str:
    if metric == METRIC_BALANCE:
        return "money"
    if metric == METRIC_BALANCE_GROWTH:
        return "percent_signed"
    return "percent"
