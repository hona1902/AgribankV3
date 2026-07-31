from __future__ import annotations

from agribank_v3.features.credit.summary.customer.filters import (
    MOVEMENT_STATUS_DECREASE,
    MOVEMENT_STATUS_INCREASE,
    MOVEMENT_STATUS_NEW,
    MOVEMENT_STATUS_PAID_OFF,
    MOVEMENT_STATUS_UNCHANGED,
)


SeriesDataset = tuple[tuple[str, tuple[tuple[str, float], ...]], ...]
METRIC_LABELS = {
    "average_rate": "Lãi suất bình quân",
    "nim_before": "NIM trước điều chỉnh",
    "nim_after": "NIM sau điều chỉnh",
}
METRIC_TITLES = {
    "average_rate": "Xu hướng lãi suất bình quân",
    "nim_before": "Xu hướng NIM trước điều chỉnh",
    "nim_after": "Xu hướng NIM sau điều chỉnh",
}
CUSTOMER_TYPE_ORDER = {"CN": 0, "TC": 1, "OTHER": 2}


def dashboard_chart_dataset_balance_trend(rows: list[dict[str, object]] | tuple[dict[str, object], ...]) -> SeriesDataset:
    if rows and "series_key" in dict(rows[0]):
        return _series_from_grouped_rows(rows, value_field="value")
    ordered = _period_ordered(rows)
    return (("Tổng dư nợ", _points(ordered, "total_balance")),)


def dashboard_chart_dataset_term_structure(metrics: dict[str, object]) -> tuple[tuple[str, float], ...]:
    return (
        ("Ngắn hạn", _number(metrics.get("short_term_balance"))),
        ("Trung/dài hạn", _number(metrics.get("medium_long_term_balance"))),
        ("Chưa phân loại", _number(metrics.get("other_balance"))),
    )


def dashboard_chart_dataset_customer_movements(
    kpis: dict[str, object],
    *,
    value_mode: str = "count",
    period: str = "",
) -> SeriesDataset:
    label = period or "Hiện tại"
    if value_mode == "money":
        return (
            ("Vay mới", ((label, _number(kpis.get("new_customer_balance"))),)),
            ("Tất toán", ((label, _number(kpis.get("paid_off_customer_balance"))),)),
            ("Tăng dư nợ", ((label, _number(kpis.get("total_increase"))),)),
            ("Giảm dư nợ", ((label, _number(kpis.get("total_decrease"))),)),
            ("Không thay đổi", ((label, _number(kpis.get("unchanged_customer_balance"))),)),
        )
    return (
        ("Vay mới", ((label, _number(kpis.get("new_customer_count"))),)),
        ("Tất toán", ((label, _number(kpis.get("paid_off_customer_count"))),)),
        ("Tăng dư nợ", ((label, _number(kpis.get("increased_customer_count"))),)),
        ("Giảm dư nợ", ((label, _number(kpis.get("decreased_customer_count"))),)),
        ("Không thay đổi", ((label, _number(kpis.get("unchanged_customer_count"))),)),
    )


def dashboard_chart_dataset_nim_rates(rows: list[dict[str, object]] | tuple[dict[str, object], ...]) -> SeriesDataset:
    ordered = _period_ordered(rows)
    return (
        ("Lãi suất bình quân", _points(ordered, "average_rate")),
        ("NIM trước ĐC", _points(ordered, "nim_before")),
        ("NIM sau ĐC", _points(ordered, "nim_after")),
    )


def dashboard_chart_dataset_metric_trend(
    rows: list[dict[str, object]] | tuple[dict[str, object], ...],
    metric: str,
) -> SeriesDataset:
    metric_key = str(metric or "average_rate")
    label = METRIC_LABELS.get(metric_key, METRIC_LABELS["average_rate"])
    if rows and "value" in dict(rows[0]):
        ordered = _period_ordered(rows)
        return ((label, _points(ordered, "value")),)
    ordered = _period_ordered(rows)
    return ((label, _points(ordered, metric_key if metric_key in METRIC_LABELS else "average_rate")),)


def dashboard_chart_dataset_active_customer_count(
    rows: list[dict[str, object]] | tuple[dict[str, object], ...],
) -> SeriesDataset:
    ordered = _period_ordered(rows)
    return (("Số khách hàng còn dư nợ", _points(ordered, "active_customer_count")),)


def dashboard_top_balance_dataset(rows: list[dict[str, object]] | tuple[dict[str, object], ...]) -> list[dict[str, object]]:
    return [dict(row) for row in rows]


def dashboard_top_increase_dataset(rows: list[dict[str, object]] | tuple[dict[str, object], ...]) -> list[dict[str, object]]:
    return [dict(row) for row in rows if _number(row.get("difference")) > 0]


def dashboard_top_decrease_dataset(rows: list[dict[str, object]] | tuple[dict[str, object], ...]) -> list[dict[str, object]]:
    return [dict(row) for row in rows if _number(row.get("difference")) < 0]


def customer_detail_chart_datasets(rows: list[dict[str, object]] | tuple[dict[str, object], ...]) -> dict[str, SeriesDataset]:
    ordered = _period_ordered(rows)
    compare_money = tuple(
        (str(row.get("period") or ""), _number(row.get("difference")))
        for row in ordered
        if row.get("difference") not in (None, "")
    )
    compare_percent = tuple(
        (str(row.get("period") or ""), _number(row.get("growth_rate")))
        for row in ordered
        if row.get("growth_rate") not in (None, "")
    )
    return {
        "balance": (
            ("Tổng dư nợ", _points(ordered, "total_balance")),
            ("Dư nợ ngắn hạn", _points(ordered, "short_term_balance")),
            ("Dư nợ trung/dài hạn", _points(ordered, "medium_long_term_balance")),
        ),
        "term_money": (
            ("Dư nợ ngắn hạn", _points(ordered, "short_term_balance")),
            ("Dư nợ trung/dài hạn", _points(ordered, "medium_long_term_balance")),
            ("Dư nợ khác", _points(ordered, "other_balance")),
        ),
        "term_ratio": (("Tỷ lệ trung/dài hạn", _points(ordered, "medium_long_ratio")),),
        "nim_rates": dashboard_chart_dataset_nim_rates(ordered),
        "compare_money": (("Tăng/giảm tuyệt đối", compare_money),),
        "compare_percent": (("Tăng trưởng", compare_percent),),
    }


def chart_periods_are_sorted(series: SeriesDataset) -> bool:
    for _name, points in series:
        labels = [label for label, _value in points]
        if labels != sorted(labels):
            return False
    return True


def weighted_metric_from_rows(
    rows: list[dict[str, object]] | tuple[dict[str, object], ...],
    numerator_key: str,
    *,
    balance_key: str = "total_balance",
) -> float:
    denominator = sum(_number(row.get(balance_key)) for row in rows)
    if denominator == 0:
        return 0.0
    return sum(_number(row.get(numerator_key)) for row in rows) / denominator


def movement_status_for_top_mode(mode: str) -> str:
    if mode == "decrease":
        return MOVEMENT_STATUS_DECREASE
    if mode == "new":
        return MOVEMENT_STATUS_NEW
    if mode == "paid_off":
        return MOVEMENT_STATUS_PAID_OFF
    if mode == "unchanged":
        return MOVEMENT_STATUS_UNCHANGED
    return MOVEMENT_STATUS_INCREASE


def _period_ordered(rows: list[dict[str, object]] | tuple[dict[str, object], ...]) -> list[dict[str, object]]:
    return sorted((dict(row) for row in rows), key=lambda row: str(row.get("period") or ""))


def _points(rows: list[dict[str, object]], field: str) -> tuple[tuple[str, float], ...]:
    return tuple((str(row.get("period") or ""), _number(row.get(field))) for row in rows)


def _series_from_grouped_rows(
    rows: list[dict[str, object]] | tuple[dict[str, object], ...],
    *,
    value_field: str,
) -> SeriesDataset:
    grouped: dict[str, list[dict[str, object]]] = {}
    names: dict[str, str] = {}
    for row in rows:
        item = dict(row)
        key = str(item.get("series_key") or "total")
        grouped.setdefault(key, []).append(item)
        names.setdefault(key, str(item.get("series_name") or key))
    series: list[tuple[str, tuple[tuple[str, float], ...]]] = []
    for key in sorted(grouped, key=_series_sort_key):
        ordered = _period_ordered(grouped[key])
        series.append((names[key], _points(ordered, value_field)))
    return tuple(series)


def _series_sort_key(key: str) -> tuple[int, str]:
    if key == "total":
        return (-1, key)
    if key in CUSTOMER_TYPE_ORDER:
        return (CUSTOMER_TYPE_ORDER[key], key)
    return (10, key)


def _number(value: object) -> float:
    try:
        return float(value or 0)
    except (TypeError, ValueError):
        return 0.0
