from __future__ import annotations

from dataclasses import dataclass

from agribank_v3.features.credit.summary.customer.filters import CustomerFilters


@dataclass(frozen=True, slots=True)
class PeriodFilterValidation:
    valid: bool
    period_from: str = ""
    period_to: str = ""
    report_period: str = ""
    reason: str = ""
    no_data: bool = False


def validate_dashboard_period_filters(
    periods: list[str] | tuple[str, ...],
    filters: CustomerFilters,
) -> PeriodFilterValidation:
    clean_periods = [str(period or "").strip() for period in periods if str(period or "").strip()]
    if not clean_periods:
        return PeriodFilterValidation(
            valid=False,
            reason="Chưa có dữ liệu khách hàng.",
            no_data=True,
        )
    period_set = set(clean_periods)
    period_from = str(filters.period_from or "").strip()
    period_to = str(filters.period_to or "").strip()
    report_period = str(filters.current_period or "").strip()
    if period_from not in period_set:
        period_from = clean_periods[0]
    if period_to not in period_set:
        period_to = clean_periods[-1]
    if period_from > period_to:
        period_from = period_to
    if report_period not in period_set or report_period < period_from or report_period > period_to:
        report_period = period_to
    if not period_from or not period_to or not report_period:
        return PeriodFilterValidation(
            valid=False,
            period_from=period_from,
            period_to=period_to,
            report_period=report_period,
            reason="Chưa chọn đủ kỳ dữ liệu.",
        )
    return PeriodFilterValidation(
        valid=True,
        period_from=period_from,
        period_to=period_to,
        report_period=report_period,
    )
