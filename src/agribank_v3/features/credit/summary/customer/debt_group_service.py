from __future__ import annotations

from agribank_v3.features.credit.summary.customer.filters import CustomerFilters
from agribank_v3.features.credit.summary.customer.repository import CustomerRepository


def build_debt_group_payload(
    repository: CustomerRepository,
    report_period: str,
    filters: CustomerFilters,
    *,
    branch_page: int = 1,
    branch_page_size: int = 100,
    branch_sort_by: str = "bad_debt_ratio",
    branch_sort_desc: bool = True,
    officer_page: int = 1,
    officer_page_size: int = 100,
    officer_sort_by: str = "bad_debt_ratio",
    officer_sort_desc: bool = True,
    customer_page: int = 1,
    customer_page_size: int = 100,
    customer_sort_by: str = "bad_debt_ratio",
    customer_sort_desc: bool = True,
) -> dict[str, object]:
    period = str(report_period or filters.current_period or filters.period_to or "").strip()
    return {
        "report_period": period,
        "kpis": repository.get_debt_quality_kpis(period, filters),
        "summary_rows": repository.get_debt_group_summary(period, filters),
        "trend_rows": repository.get_debt_group_trend(filters.period_from, filters.period_to, filters),
        "branch_result": repository.query_debt_group_by_branch(
            period,
            filters,
            page=branch_page,
            page_size=branch_page_size,
            sort_by=branch_sort_by,
            sort_desc=branch_sort_desc,
        ),
        "officer_result": repository.query_debt_group_by_officer(
            period,
            filters,
            page=officer_page,
            page_size=officer_page_size,
            sort_by=officer_sort_by,
            sort_desc=officer_sort_desc,
        ),
        "customer_result": repository.query_debt_group_customers(
            period,
            filters,
            page=customer_page,
            page_size=customer_page_size,
            sort_by=customer_sort_by,
            sort_desc=customer_sort_desc,
        ),
    }
