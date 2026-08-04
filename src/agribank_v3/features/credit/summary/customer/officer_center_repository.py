from __future__ import annotations

from contextlib import closing
from dataclasses import dataclass, replace
import sqlite3
import unicodedata

from agribank_v3.features.credit.summary.customer.filters import (
    CustomerFilters,
    clean_filter_text,
)
from agribank_v3.features.credit.summary.customer.repository import (
    CustomerRepository,
    _append_debt_group_filter,
    _finalize_debt_group_metrics,
    _has_override_sql,
    _override_value_sql,
)
from agribank_v3.features.credit.summary.models import PageResult


OFFICER_MODE_IMPORTED = "imported"
OFFICER_MODE_EFFECTIVE = "effective"

OFFICER_MODE_LABELS: tuple[tuple[str, str], ...] = (
    ("Theo phân bổ dữ liệu import", OFFICER_MODE_IMPORTED),
    ("Theo cán bộ quản lý hiệu lực", OFFICER_MODE_EFFECTIVE),
)

OFFICER_STATUS_FILTERS: tuple[tuple[str, str], ...] = (
    ("Tất cả trạng thái", ""),
    ("Đang sử dụng", "active"),
    ("Ngừng sử dụng", "inactive"),
    ("Chưa có trong danh mục", "missing_directory"),
    ("Chưa có mã cán bộ", "unresolved"),
    ("Tên hệ thống/không xác định", "system"),
)

SYSTEM_OFFICER_NAMES = {
    "the tin dung",
    "khong xac dinh",
    "khong ro",
    "khong co",
    "unknown",
}

TERM_STRUCTURE_TOLERANCE = 0.01
TERM_STRUCTURE_WARNING = (
    "Kỳ này chưa có dữ liệu phân bổ kỳ hạn đầy đủ theo CBTD. "
    "Vui lòng nhập lại kỳ NIM Dư nợ từ file FTP Loan."
)


@dataclass(frozen=True, slots=True)
class OfficerCenterFilters:
    period_from: str = ""
    period_to: str = ""
    report_period: str = ""
    compare_period: str = ""
    branch_code: str = ""
    transaction_office: str = ""
    customer_type: str = ""
    loan_term: str = ""
    debt_group: str = ""
    officer_status: str = ""
    search_text: str = ""
    mode: str = OFFICER_MODE_IMPORTED
    selected_officers: tuple[str, ...] = ()
    benchmark_scope: str = "all"

    def normalized(self) -> "OfficerCenterFilters":
        mode = self.mode if self.mode in {OFFICER_MODE_IMPORTED, OFFICER_MODE_EFFECTIVE} else OFFICER_MODE_IMPORTED
        selected = tuple(str(item or "").strip() for item in self.selected_officers if str(item or "").strip())
        return replace(
            self,
            mode=mode,
            selected_officers=selected,
            period_from=str(self.period_from or "").strip(),
            period_to=str(self.period_to or "").strip(),
            report_period=str(self.report_period or "").strip(),
            compare_period=str(self.compare_period or "").strip(),
            branch_code=str(self.branch_code or "").strip(),
            transaction_office=str(self.transaction_office or "").strip(),
            customer_type=str(self.customer_type or "").strip(),
            loan_term=str(self.loan_term or "").strip(),
            debt_group=str(self.debt_group or "").strip(),
            officer_status=str(self.officer_status or "").strip(),
            search_text=str(self.search_text or "").strip(),
            benchmark_scope=str(self.benchmark_scope or "all").strip() or "all",
        )

    def as_customer_filters(self) -> CustomerFilters:
        return CustomerFilters(
            period_from=self.period_from,
            period_to=self.period_to,
            current_period=self.report_period,
            compare_period=self.compare_period,
            branch_code=self.branch_code,
            customer_type=self.customer_type,
            loan_term=self.loan_term,
            search_text="",
            debt_group=self.debt_group,
        )


class OfficerCenterRepository:
    """Officer-centric analytics built on the existing Customer.db aggregates."""

    def __init__(self, repository: CustomerRepository) -> None:
        self.repository = repository
        self.unit_directory = repository.unit_directory

    def distinct_periods(self) -> list[str]:
        return self.repository.distinct_periods()

    def distinct_branch_codes(self, filters: OfficerCenterFilters | None = None) -> list[str]:
        customer_filters = (filters or OfficerCenterFilters()).normalized().as_customer_filters()
        return self.repository.distinct_branch_codes(customer_filters)

    def distinct_offices(self, period: str, *, branch_code: str = "") -> list[dict[str, object]]:
        return self.repository.distinct_offices(period, branch_code=branch_code)

    def previous_period(self, report_period: str) -> str:
        periods = self.distinct_periods()
        clean = str(report_period or "").strip()
        if not periods:
            return ""
        if clean in periods:
            index = periods.index(clean)
            return periods[index - 1] if index > 0 else ""
        before = [period for period in periods if period < clean]
        return before[-1] if before else ""

    def officer_options(self, filters: OfficerCenterFilters, *, limit: int = 5000) -> list[dict[str, object]]:
        filters = filters.normalized()
        sql, params = self._officer_aggregate_sql(filters, require_report_period=False)
        sql = f"""
            SELECT officer_key, officer_code, officer_name, branch_code, transaction_office, total_balance
            FROM ({sql}) q
            WHERE officer_key <> ''
            ORDER BY officer_name COLLATE NOCASE, officer_code COLLATE NOCASE
            LIMIT ?
        """
        with closing(self.repository.connect()) as database:
            rows = database.execute(sql, (*params, max(1, int(limit or 5000)))).fetchall()
        return [self._finalize_row(dict(row)) for row in rows]

    def dashboard_payload(
        self,
        filters: OfficerCenterFilters,
        *,
        top_metric: str = "total_balance",
        top_limit: int = 10,
    ) -> dict[str, object]:
        filters = filters.normalized()
        limit = _top_limit(top_limit)
        metric = _top_metric(top_metric)
        return {
            "kpis": self.kpis(filters),
            "balance_trend": self.trend(filters, "total_balance"),
            "officer_count_trend": self.trend(filters, "officer_count"),
            "metric_trend": self.trend(filters, "nim_after"),
            "debt_structure": self.debt_structure(filters),
            "top_rows": self.top_officers(filters, metric=metric, limit=limit),
            "top_metric": metric,
            "top_limit": limit,
            "mode_label": _mode_label(filters.mode),
        }

    def top_officers(
        self,
        filters: OfficerCenterFilters,
        *,
        metric: str = "total_balance",
        limit: int = 10,
    ) -> list[dict[str, object]]:
        filters = filters.normalized()
        metric = _top_metric(metric)
        limit = _top_limit(limit)
        if metric == "balance_change":
            rows = self.officer_movement(
                filters,
                page=1,
                page_size=limit,
                sort_by="balance_change",
                sort_desc=True,
            ).rows
            return [
                dict(
                    row,
                    total_balance=row.get("current_balance"),
                    customer_count=row.get("current_customer_count"),
                )
                for row in rows
            ]
        return self.officer_list(filters, page=1, page_size=limit, sort_by=metric, sort_desc=True).rows

    def kpis(self, filters: OfficerCenterFilters) -> dict[str, object]:
        filters = filters.normalized()
        sql, params = self._officer_aggregate_sql(filters)
        base_sql, base_params = self._base_sql(filters, period_mode="report")
        with closing(self.repository.connect()) as database:
            row = database.execute(
                f"""
                WITH officers AS ({sql}),
                base AS ({base_sql})
                SELECT
                    COUNT(CASE WHEN officers.total_balance > 0 THEN 1 END) AS active_officer_count,
                    COALESCE(SUM(officers.total_balance), 0) AS total_balance,
                    COALESCE(SUM(officers.customer_count), 0) AS officer_customer_occurrence_count,
                    (SELECT COUNT(DISTINCT customer_code) FROM base WHERE total_balance > 0) AS unique_customer_count,
                    COALESCE(SUM(officers.interest_rate_numerator), 0) AS interest_rate_numerator,
                    COALESCE(SUM(officers.nim_before_numerator), 0) AS nim_before_numerator,
                    COALESCE(SUM(officers.nim_after_numerator), 0) AS nim_after_numerator,
                    COALESCE(SUM(officers.debt_group_2_balance), 0) AS attention_balance,
                    COALESCE(SUM(officers.bad_debt_balance), 0) AS bad_debt_balance,
                    COALESCE(SUM(CASE WHEN officers.debt_group_2_balance > 0 THEN 1 ELSE 0 END), 0) AS attention_officer_count,
                    COALESCE(SUM(CASE WHEN officers.bad_debt_balance > 0 THEN 1 ELSE 0 END), 0) AS bad_debt_officer_count,
                    COALESCE(SUM(CASE WHEN officers.has_debt_group_data THEN 1 ELSE 0 END), 0) AS has_debt_group_officer_count
                FROM officers
                """,
                (*params, *base_params),
            ).fetchone()
        data = dict(row) if row is not None else {}
        total_balance = _number(data.get("total_balance"))
        active_officers = int(data.get("active_officer_count") or 0)
        occurrences = int(data.get("officer_customer_occurrence_count") or 0)
        data["average_balance_per_officer"] = total_balance / active_officers if active_officers else None
        data["average_customer_per_officer"] = occurrences / active_officers if active_officers else None
        data["average_rate"] = _ratio(_number(data.get("interest_rate_numerator")), total_balance)
        data["nim_before"] = _ratio(_number(data.get("nim_before_numerator")), total_balance)
        data["nim_after"] = _ratio(_number(data.get("nim_after_numerator")), total_balance)
        data["attention_ratio"] = _ratio(_number(data.get("attention_balance")) * 100, total_balance)
        data["bad_debt_ratio"] = _ratio(_number(data.get("bad_debt_balance")) * 100, total_balance)
        data["mode_label"] = _mode_label(filters.mode)
        return data

    def trend(self, filters: OfficerCenterFilters, metric: str = "total_balance") -> list[dict[str, object]]:
        filters = filters.normalized()
        sql, params = self._base_sql(filters, period_mode="range")
        metric_expr = _trend_metric_expr(metric)
        with closing(self.repository.connect()) as database:
            rows = database.execute(
                f"""
                WITH base AS ({sql})
                SELECT
                    period,
                    {metric_expr} AS value
                FROM base
                GROUP BY period
                ORDER BY period
                """,
                params,
            ).fetchall()
        return [dict(row) for row in rows]

    def debt_structure(self, filters: OfficerCenterFilters) -> dict[str, float]:
        filters = filters.normalized()
        sql, params = self._base_sql(filters, period_mode="report")
        with closing(self.repository.connect()) as database:
            row = database.execute(
                f"""
                WITH base AS ({sql})
                SELECT
                    COALESCE(SUM(debt_group_1_balance), 0) AS debt_group_1_balance,
                    COALESCE(SUM(debt_group_2_balance), 0) AS debt_group_2_balance,
                    COALESCE(SUM(debt_group_3_balance), 0) AS debt_group_3_balance,
                    COALESCE(SUM(debt_group_4_balance), 0) AS debt_group_4_balance,
                    COALESCE(SUM(debt_group_5_balance), 0) AS debt_group_5_balance,
                    COALESCE(SUM(debt_group_unknown_balance), 0) AS debt_group_unknown_balance
                FROM base
                """,
                params,
            ).fetchone()
        return dict(row) if row is not None else {}

    def officer_list(
        self,
        filters: OfficerCenterFilters,
        *,
        page: int = 1,
        page_size: int = 100,
        sort_by: str = "total_balance",
        sort_desc: bool = True,
    ) -> PageResult:
        return self._paged_officer_rows(filters, page=page, page_size=page_size, sort_by=sort_by, sort_desc=sort_desc)

    def officer_debt_quality(
        self,
        filters: OfficerCenterFilters,
        *,
        page: int = 1,
        page_size: int = 100,
        sort_by: str = "bad_debt_ratio",
        sort_desc: bool = True,
    ) -> PageResult:
        return self._paged_officer_rows(filters, page=page, page_size=page_size, sort_by=sort_by, sort_desc=sort_desc)

    def compare_officers(
        self,
        filters: OfficerCenterFilters,
        *,
        page: int = 1,
        page_size: int = 100,
        sort_by: str = "total_balance",
        sort_desc: bool = True,
    ) -> PageResult:
        filters = filters.normalized()
        result = self._paged_officer_rows(filters, page=page, page_size=page_size, sort_by=sort_by, sort_desc=sort_desc)
        benchmark = self._benchmark(replace(filters, selected_officers=()))
        rows = [dict(row) for row in result.rows]
        for row in rows:
            row.update(_comparison_values(row, benchmark))
        return PageResult(rows=rows, total_rows=result.total_rows, page=result.page, page_size=result.page_size)

    def officer_customers(
        self,
        filters: OfficerCenterFilters,
        *,
        officer_key: str = "",
        page: int = 1,
        page_size: int = 100,
        sort_by: str = "customer_code",
        sort_desc: bool = False,
    ) -> PageResult:
        filters = filters.normalized()
        base_sql, params = self._base_sql(filters, period_mode="report")
        clauses = ["1 = 1"]
        if officer_key:
            clauses.append("(officer_key = ? OR officer_code = ?)")
            params.extend([officer_key, officer_key])
        where = "WHERE " + " AND ".join(clauses)
        order = _order_sql(sort_by, sort_desc, _CUSTOMER_SORTS, default="customer_code")
        page = max(1, int(page or 1))
        page_size = max(1, min(5000, int(page_size or 100)))
        offset = (page - 1) * page_size
        with closing(self.repository.connect()) as database:
            total = int(
                database.execute(
                    f"WITH base AS ({base_sql}) SELECT COUNT(*) FROM base {where}",
                    params,
                ).fetchone()[0]
                or 0
            )
            rows = database.execute(
                f"""
                WITH base AS ({base_sql})
                SELECT
                    period,
                    customer_code,
                    customer_name,
                    customer_type,
                    branch_code,
                    transaction_office,
                    officer_code AS imported_officer_code,
                    officer_name AS imported_officer_name,
                    officer_key,
                    total_customer_balance,
                    total_balance AS officer_balance,
                    CASE WHEN total_customer_balance <> 0 THEN total_balance / total_customer_balance * 100 ELSE NULL END AS officer_share,
                    short_term_balance,
                    medium_long_term_balance,
                    other_balance,
                    (COALESCE(short_term_balance, 0) + COALESCE(medium_long_term_balance, 0) + COALESCE(other_balance, 0)) AS term_balance_total,
                    CASE
                        WHEN total_balance <= 0 THEN 1
                        WHEN ABS(
                            COALESCE(short_term_balance, 0)
                            + COALESCE(medium_long_term_balance, 0)
                            + COALESCE(other_balance, 0)
                            - total_balance
                        ) <= {TERM_STRUCTURE_TOLERANCE} THEN 1
                        ELSE 0
                    END AS term_structure_available,
                    debt_group_2_balance AS attention_balance,
                    debt_group_3_balance + debt_group_4_balance + debt_group_5_balance AS bad_debt_balance,
                    worst_debt_group,
                    average_rate,
                    nim_before,
                    nim_after,
                    has_multiple_officers,
                    has_override
                FROM base
                {where}
                {order}
                LIMIT ? OFFSET ?
                """,
                (*params, page_size, offset),
            ).fetchall()
        output = [self._finalize_row(dict(row, mode=filters.mode)) for row in rows]
        return PageResult(rows=output, total_rows=total, page=page, page_size=page_size)

    def officer_movement(
        self,
        filters: OfficerCenterFilters,
        *,
        previous_period: str = "",
        current_period: str = "",
        page: int = 1,
        page_size: int = 100,
        sort_by: str = "balance_change",
        sort_desc: bool = True,
    ) -> PageResult:
        filters = filters.normalized()
        current = str(current_period or filters.report_period or "").strip()
        previous = str(previous_period or filters.compare_period or self.previous_period(current)).strip()
        if not previous or not current:
            return PageResult(rows=[], total_rows=0, page=max(1, int(page or 1)), page_size=max(1, int(page_size or 100)))
        prev_filters = replace(filters, report_period=previous)
        curr_filters = replace(filters, report_period=current)
        prev_sql, prev_params = self._base_sql(prev_filters, period_mode="report")
        curr_sql, curr_params = self._base_sql(curr_filters, period_mode="report")
        order = _order_sql(sort_by, sort_desc, _MOVEMENT_SORTS, default="balance_change")
        page = max(1, int(page or 1))
        page_size = max(1, min(5000, int(page_size or 100)))
        offset = (page - 1) * page_size
        with closing(self.repository.connect()) as database:
            aggregate_sql = _movement_sql(prev_sql, curr_sql)
            params = (*prev_params, *curr_params, previous, current)
            total = int(database.execute(f"SELECT COUNT(*) FROM ({aggregate_sql}) q", params).fetchone()[0] or 0)
            rows = database.execute(f"{aggregate_sql} {order} LIMIT ? OFFSET ?", (*params, page_size, offset)).fetchall()
        output = [self._finalize_row(dict(row, mode=filters.mode)) for row in rows]
        return PageResult(rows=output, total_rows=total, page=page, page_size=page_size)

    def officer_period_history(self, filters: OfficerCenterFilters) -> list[dict[str, object]]:
        filters = filters.normalized()
        if not filters.selected_officers:
            return []
        base_sql, params = self._base_sql(filters, period_mode="range")
        with closing(self.repository.connect()) as database:
            rows = database.execute(
                f"""
                WITH base AS ({base_sql})
                SELECT
                    period,
                    '{filters.mode}' AS mode,
                    MAX(officer_key) AS officer_key,
                    MAX(officer_code) AS officer_code,
                    MAX(officer_name) AS officer_name,
                    COALESCE(NULLIF(MAX(directory_branch_code), ''), MIN(NULLIF(branch_code, '')), '') AS branch_code,
                    COALESCE(NULLIF(MAX(directory_transaction_office), ''), MIN(NULLIF(transaction_office, '')), '') AS transaction_office,
                    MAX(COALESCE(directory_is_active, -1)) AS directory_is_active,
                    COUNT(DISTINCT customer_code) AS customer_count,
                    COALESCE(SUM(total_balance), 0) AS total_balance,
                    COALESCE(SUM(short_term_balance), 0) AS short_term_balance,
                    COALESCE(SUM(medium_long_term_balance), 0) AS medium_long_term_balance,
                    COALESCE(SUM(other_balance), 0) AS other_balance,
                    (
                        COALESCE(SUM(short_term_balance), 0)
                        + COALESCE(SUM(medium_long_term_balance), 0)
                        + COALESCE(SUM(other_balance), 0)
                    ) AS term_balance_total,
                    COALESCE(SUM(interest_rate_numerator), 0) AS interest_rate_numerator,
                    COALESCE(SUM(nim_before_numerator), 0) AS nim_before_numerator,
                    COALESCE(SUM(nim_after_numerator), 0) AS nim_after_numerator,
                    MAX(has_debt_group_data) AS has_debt_group_data,
                    COALESCE(SUM(debt_group_1_balance), 0) AS debt_group_1_balance,
                    COALESCE(SUM(debt_group_2_balance), 0) AS debt_group_2_balance,
                    COALESCE(SUM(debt_group_3_balance), 0) AS debt_group_3_balance,
                    COALESCE(SUM(debt_group_4_balance), 0) AS debt_group_4_balance,
                    COALESCE(SUM(debt_group_5_balance), 0) AS debt_group_5_balance,
                    COALESCE(SUM(debt_group_unknown_balance), 0) AS debt_group_unknown_balance,
                    COUNT(DISTINCT CASE WHEN debt_group_2_balance > 0 THEN customer_code END) AS attention_customer_count,
                    COUNT(DISTINCT CASE WHEN (debt_group_3_balance + debt_group_4_balance + debt_group_5_balance) > 0 THEN customer_code END) AS bad_debt_customer_count,
                    COUNT(DISTINCT CASE WHEN has_multiple_officers = 1 THEN customer_code END) AS multiple_officer_customer_count,
                    COUNT(DISTINCT CASE WHEN has_override = 1 THEN customer_code END) AS override_customer_count,
                    CASE
                        WHEN SUM(total_balance) <= 0 THEN 1
                        WHEN ABS(
                            COALESCE(SUM(short_term_balance), 0)
                            + COALESCE(SUM(medium_long_term_balance), 0)
                            + COALESCE(SUM(other_balance), 0)
                            - SUM(total_balance)
                        ) <= {TERM_STRUCTURE_TOLERANCE} THEN 1
                        ELSE 0
                    END AS term_structure_available,
                    CASE WHEN SUM(total_balance) <> 0 THEN SUM(medium_long_term_balance) / SUM(total_balance) * 100 ELSE NULL END AS medium_long_ratio,
                    CASE WHEN SUM(total_balance) <> 0 THEN SUM(interest_rate_numerator) / SUM(total_balance) ELSE NULL END AS average_rate,
                    CASE WHEN SUM(total_balance) <> 0 THEN SUM(nim_before_numerator) / SUM(total_balance) ELSE NULL END AS nim_before,
                    CASE WHEN SUM(total_balance) <> 0 THEN SUM(nim_after_numerator) / SUM(total_balance) ELSE NULL END AS nim_after,
                    CASE WHEN SUM(total_balance) <> 0 THEN SUM(debt_group_2_balance) / SUM(total_balance) * 100 ELSE NULL END AS attention_ratio,
                    CASE WHEN SUM(total_balance) <> 0 THEN SUM(debt_group_3_balance + debt_group_4_balance + debt_group_5_balance) / SUM(total_balance) * 100 ELSE NULL END AS bad_debt_ratio
                FROM base
                WHERE total_balance > 0
                GROUP BY period
                ORDER BY period
                """,
                params,
            ).fetchall()
        return [self._finalize_row(dict(row)) for row in rows]

    def _paged_officer_rows(
        self,
        filters: OfficerCenterFilters,
        *,
        page: int,
        page_size: int,
        sort_by: str,
        sort_desc: bool,
    ) -> PageResult:
        filters = filters.normalized()
        sql, params = self._officer_aggregate_sql(filters)
        order = _order_sql(sort_by, sort_desc, _OFFICER_SORTS, default="total_balance")
        page = max(1, int(page or 1))
        page_size = max(1, min(5000, int(page_size or 100)))
        offset = (page - 1) * page_size
        with closing(self.repository.connect()) as database:
            total = int(database.execute(f"SELECT COUNT(*) FROM ({sql}) q", params).fetchone()[0] or 0)
            rows = database.execute(f"{sql} {order} LIMIT ? OFFSET ?", (*params, page_size, offset)).fetchall()
        output = [self._finalize_row(dict(row)) for row in rows]
        return PageResult(rows=output, total_rows=total, page=page, page_size=page_size)

    def _officer_aggregate_sql(
        self,
        filters: OfficerCenterFilters,
        *,
        require_report_period: bool = True,
    ) -> tuple[str, list[object]]:
        filters = filters.normalized()
        base_sql, params = self._base_sql(filters, period_mode="report" if require_report_period else "range")
        status_clause, status_params = _status_clause(filters.officer_status)
        search_clause = ""
        if filters.search_text:
            search_clause = """
                AND (
                    grouped.officer_code LIKE ?
                    OR grouped.officer_name LIKE ?
                    OR grouped.officer_key LIKE ?
                )
            """
            pattern = f"%{clean_filter_text(filters.search_text)}%"
            status_params.extend([pattern, pattern, pattern])
        sql = f"""
            WITH base AS ({base_sql}),
            grouped AS (
                SELECT
                    officer_key,
                    MAX(CASE WHEN officer_code <> '' THEN officer_code ELSE '' END) AS officer_code,
                    MAX(CASE WHEN officer_name <> '' THEN officer_name ELSE '' END) AS officer_name,
                    COALESCE(NULLIF(MAX(directory_branch_code), ''), MIN(NULLIF(branch_code, '')), '') AS branch_code,
                    COALESCE(NULLIF(MAX(directory_transaction_office), ''), MIN(NULLIF(transaction_office, '')), '') AS transaction_office,
                    MAX(COALESCE(directory_is_active, -1)) AS directory_is_active,
                    COUNT(DISTINCT customer_code) AS customer_count,
                    COALESCE(SUM(total_balance), 0) AS total_balance,
                    SUM(short_term_balance) AS short_term_balance,
                    SUM(medium_long_term_balance) AS medium_long_term_balance,
                    SUM(other_balance) AS other_balance,
                    (
                        COALESCE(SUM(short_term_balance), 0)
                        + COALESCE(SUM(medium_long_term_balance), 0)
                        + COALESCE(SUM(other_balance), 0)
                    ) AS term_balance_total,
                    COALESCE(SUM(interest_rate_numerator), 0) AS interest_rate_numerator,
                    COALESCE(SUM(nim_before_numerator), 0) AS nim_before_numerator,
                    COALESCE(SUM(nim_after_numerator), 0) AS nim_after_numerator,
                    MAX(has_debt_group_data) AS has_debt_group_data,
                    COALESCE(SUM(debt_group_1_balance), 0) AS debt_group_1_balance,
                    COALESCE(SUM(debt_group_2_balance), 0) AS debt_group_2_balance,
                    COALESCE(SUM(debt_group_3_balance), 0) AS debt_group_3_balance,
                    COALESCE(SUM(debt_group_4_balance), 0) AS debt_group_4_balance,
                    COALESCE(SUM(debt_group_5_balance), 0) AS debt_group_5_balance,
                    COALESCE(SUM(debt_group_unknown_balance), 0) AS debt_group_unknown_balance,
                    COUNT(DISTINCT CASE WHEN debt_group_2_balance > 0 THEN customer_code END) AS attention_customer_count,
                    COUNT(DISTINCT CASE WHEN (debt_group_3_balance + debt_group_4_balance + debt_group_5_balance) > 0 THEN customer_code END) AS bad_debt_customer_count,
                    COUNT(DISTINCT CASE WHEN has_multiple_officers = 1 THEN customer_code END) AS multiple_officer_customer_count,
                    COUNT(DISTINCT CASE WHEN has_override = 1 THEN customer_code END) AS override_customer_count,
                    COUNT(DISTINCT period) AS period_count,
                    MAX(period) AS latest_period,
                    MIN(period) AS first_period,
                    COUNT(DISTINCT NULLIF(branch_code, '')) AS branch_count,
                    COUNT(DISTINCT NULLIF(transaction_office, '')) AS office_count
                FROM base
                WHERE total_balance > 0
                GROUP BY officer_key
            )
            SELECT
                grouped.*,
                '{filters.mode}' AS mode,
                (debt_group_3_balance + debt_group_4_balance + debt_group_5_balance) AS bad_debt_balance,
                CASE
                    WHEN total_balance <= 0 THEN 1
                    WHEN ABS(term_balance_total - total_balance) <= {TERM_STRUCTURE_TOLERANCE} THEN 1
                    ELSE 0
                END AS term_structure_available,
                CASE WHEN total_balance <> 0 THEN medium_long_term_balance / total_balance * 100 ELSE NULL END AS medium_long_ratio,
                CASE WHEN total_balance <> 0 THEN interest_rate_numerator / total_balance ELSE NULL END AS average_rate,
                CASE WHEN total_balance <> 0 THEN nim_before_numerator / total_balance ELSE NULL END AS nim_before,
                CASE WHEN total_balance <> 0 THEN nim_after_numerator / total_balance ELSE NULL END AS nim_after,
                CASE WHEN total_balance <> 0 THEN debt_group_2_balance / total_balance * 100 ELSE NULL END AS attention_ratio,
                CASE WHEN total_balance <> 0 THEN (debt_group_3_balance + debt_group_4_balance + debt_group_5_balance) / total_balance * 100 ELSE NULL END AS bad_debt_ratio,
                {_status_expr()} AS officer_status,
                CASE
                    WHEN branch_count > 1 THEN 'Xuất hiện ở nhiều chi nhánh'
                    WHEN office_count > 1 THEN 'Xuất hiện ở nhiều PGD'
                    WHEN officer_code = '' THEN 'Chưa có mã cán bộ'
                    WHEN directory_is_active = 0 AND total_balance > 0 THEN 'Inactive nhưng còn dư nợ'
                    ELSE ''
                END AS data_warning
            FROM grouped
            WHERE 1 = 1
            {status_clause}
            {search_clause}
        """
        return sql, [*params, *status_params]

    def _base_sql(self, filters: OfficerCenterFilters, *, period_mode: str) -> tuple[str, list[object]]:
        filters = filters.normalized()
        if filters.mode == OFFICER_MODE_EFFECTIVE:
            return _effective_base_sql(filters, period_mode)
        return _imported_base_sql(filters, period_mode)

    def _benchmark(self, filters: OfficerCenterFilters) -> dict[str, object]:
        filters = filters.normalized()
        sql, params = self._base_sql(filters, period_mode="report")
        with closing(self.repository.connect()) as database:
            row = database.execute(
                f"""
                WITH base AS ({sql})
                SELECT
                    COALESCE(SUM(total_balance), 0) AS total_balance,
                    COALESCE(SUM(interest_rate_numerator), 0) AS interest_rate_numerator,
                    COALESCE(SUM(nim_before_numerator), 0) AS nim_before_numerator,
                    COALESCE(SUM(nim_after_numerator), 0) AS nim_after_numerator,
                    COALESCE(SUM(debt_group_2_balance), 0) AS debt_group_2_balance,
                    COALESCE(SUM(debt_group_3_balance + debt_group_4_balance + debt_group_5_balance), 0) AS bad_debt_balance,
                    COUNT(DISTINCT officer_key) AS benchmark_officer_count
                FROM base
                WHERE total_balance > 0
                """,
                params,
            ).fetchone()
        output = dict(row) if row is not None else {}
        total = _number(output.get("total_balance"))
        output["benchmark_average_rate"] = _ratio(_number(output.get("interest_rate_numerator")), total)
        output["benchmark_nim_before"] = _ratio(_number(output.get("nim_before_numerator")), total)
        output["benchmark_nim_after"] = _ratio(_number(output.get("nim_after_numerator")), total)
        output["benchmark_attention_ratio"] = _ratio(_number(output.get("debt_group_2_balance")) * 100, total)
        output["benchmark_bad_debt_ratio"] = _ratio(_number(output.get("bad_debt_balance")) * 100, total)
        return output

    def _finalize_row(self, row: dict[str, object]) -> dict[str, object]:
        row = dict(row)
        branch = str(row.get("branch_code") or "")
        office = str(row.get("transaction_office") or "")
        row["branch_name"] = self.unit_directory.get_branch_display_name(branch) if branch else ""
        row["office_name"] = self.unit_directory.get_office_name(branch, office) if branch and office else ""
        row["officer_display"] = str(row.get("officer_name") or row.get("officer_code") or row.get("officer_key") or "")
        row["mode_label"] = _mode_label(str(row.get("mode") or ""))
        if "officer_status" not in row:
            row["officer_status"] = _status_from_row(row)
        if _is_system_officer(row.get("officer_name")):
            row["officer_status"] = "Tên hệ thống/không xác định"
        return _apply_term_structure_display(_finalize_numeric_row(row))


def _imported_base_sql(filters: OfficerCenterFilters, period_mode: str) -> tuple[str, list[object]]:
    clauses, params = _base_clauses(filters, period_mode, alias="op", summary_alias="s", total_column="balance_managed")
    if filters.transaction_office:
        clauses.append("op.transaction_office = ?")
        params.append(filters.transaction_office)
    _append_selected_officer_filter(clauses, params, filters.selected_officers, code_expr="op.officer_code", name_expr="op.officer_name")
    where = "WHERE " + " AND ".join(clauses)
    has_override = _has_override_sql("s")
    return f"""
        SELECT
            op.period,
            op.customer_code,
            s.customer_name,
            s.customer_type,
            op.officer_code,
            op.officer_name,
            {_officer_key_expr("op.officer_code", "op.officer_name")} AS officer_key,
            op.branch_code,
            op.transaction_office,
            d.branch_code AS directory_branch_code,
            d.transaction_office AS directory_transaction_office,
            d.is_active AS directory_is_active,
            op.balance_managed AS total_balance,
            s.total_balance AS total_customer_balance,
            op.short_term_balance,
            op.medium_long_term_balance,
            op.other_balance,
            op.interest_rate_numerator,
            op.nim_before_numerator,
            op.nim_after_numerator,
            op.has_debt_group_data,
            op.worst_debt_group,
            op.debt_group_1_balance,
            op.debt_group_2_balance,
            op.debt_group_3_balance,
            op.debt_group_4_balance,
            op.debt_group_5_balance,
            op.debt_group_unknown_balance,
            op.interest_rate_numerator / NULLIF(op.balance_managed, 0) AS average_rate,
            op.nim_before_numerator / NULLIF(op.balance_managed, 0) AS nim_before,
            op.nim_after_numerator / NULLIF(op.balance_managed, 0) AS nim_after,
            s.officer_count,
            s.has_multiple_officers,
            CASE WHEN {has_override} THEN 1 ELSE 0 END AS has_override
        FROM customer_officer_period op
        JOIN customer_period_summary s
            ON s.period = op.period AND s.customer_code = op.customer_code
        LEFT JOIN customer_officer_directory d
            ON d.officer_code = op.officer_code AND op.officer_code <> ''
        {where}
    """, params


def _effective_base_sql(filters: OfficerCenterFilters, period_mode: str) -> tuple[str, list[object]]:
    clauses, params = _base_clauses(filters, period_mode, alias="s", summary_alias="s", total_column="total_balance")
    if filters.transaction_office:
        clauses.append(
            "EXISTS (SELECT 1 FROM customer_office_period office_filter "
            "WHERE office_filter.period = s.period AND office_filter.customer_code = s.customer_code "
            "AND office_filter.trctcd = ?)"
        )
        params.append(filters.transaction_office)
    code_expr = _override_value_sql("s", "officer_code", fallback="s.primary_officer_code")
    name_expr = _override_value_sql("s", "officer_name", fallback="s.primary_officer_name", null_if_empty=True)
    _append_selected_officer_filter(clauses, params, filters.selected_officers, code_expr=code_expr, name_expr=name_expr)
    where = "WHERE " + " AND ".join(clauses)
    has_override = _has_override_sql("s")
    office_expr = (
        "(SELECT o.trctcd FROM customer_office_period o "
        "WHERE o.period = s.period AND o.customer_code = s.customer_code "
        "ORDER BY o.total_balance DESC, o.trctcd ASC LIMIT 1)"
    )
    return f"""
        SELECT
            s.period,
            s.customer_code,
            s.customer_name,
            s.customer_type,
            {code_expr} AS officer_code,
            {name_expr} AS officer_name,
            {_officer_key_expr(code_expr, name_expr)} AS officer_key,
            s.branch_code,
            COALESCE({office_expr}, '') AS transaction_office,
            d.branch_code AS directory_branch_code,
            d.transaction_office AS directory_transaction_office,
            d.is_active AS directory_is_active,
            s.total_balance,
            s.total_balance AS total_customer_balance,
            s.short_term_balance,
            s.medium_long_term_balance,
            s.other_balance,
            s.interest_rate_numerator,
            s.nim_before_numerator,
            s.nim_after_numerator,
            s.has_debt_group_data,
            s.worst_debt_group,
            s.debt_group_1_balance,
            s.debt_group_2_balance,
            s.debt_group_3_balance,
            s.debt_group_4_balance,
            s.debt_group_5_balance,
            s.debt_group_unknown_balance,
            s.average_rate,
            s.nim_before,
            s.nim_after,
            s.officer_count,
            s.has_multiple_officers,
            CASE WHEN {has_override} THEN 1 ELSE 0 END AS has_override
        FROM customer_period_summary s
        LEFT JOIN customer_officer_directory d
            ON d.officer_code = {code_expr} AND {code_expr} <> ''
        {where}
    """, params


def _base_clauses(
    filters: OfficerCenterFilters,
    period_mode: str,
    *,
    alias: str,
    summary_alias: str,
    total_column: str,
) -> tuple[list[str], list[object]]:
    clauses = ["1 = 1"]
    params: list[object] = []
    prefix = f"{alias}."
    if period_mode == "report":
        if filters.report_period:
            clauses.append(f"{prefix}period = ?")
            params.append(filters.report_period)
        else:
            clauses.append("0 = 1")
    else:
        if filters.period_from:
            clauses.append(f"{prefix}period >= ?")
            params.append(filters.period_from)
        if filters.period_to:
            clauses.append(f"{prefix}period <= ?")
            params.append(filters.period_to)
        if not filters.period_from and not filters.period_to and filters.report_period:
            clauses.append(f"{prefix}period = ?")
            params.append(filters.report_period)
    if filters.branch_code:
        clauses.append(f"{prefix}branch_code = ?")
        params.append(filters.branch_code)
    if filters.customer_type:
        clauses.append(f"{summary_alias}.customer_type = ?")
        params.append(filters.customer_type)
    if filters.loan_term:
        if filters.loan_term == "SHORT_TERM":
            clauses.append(f"{summary_alias}.short_term_balance > 0")
        elif filters.loan_term == "MEDIUM_LONG_TERM":
            clauses.append(f"{summary_alias}.medium_long_term_balance > 0")
        elif filters.loan_term == "OTHER":
            clauses.append(f"{summary_alias}.other_balance > 0")
    if filters.debt_group:
        _append_debt_group_filter(clauses, params, filters.debt_group, alias=alias, total_column=total_column)
    return clauses, params


def _append_selected_officer_filter(
    clauses: list[str],
    params: list[object],
    selected_officers: tuple[str, ...],
    *,
    code_expr: str,
    name_expr: str,
) -> None:
    selected = tuple(_normalize_selected_officer(item) for item in selected_officers if str(item or "").strip())
    if not selected:
        return
    keys = [_officer_key_expr(code_expr, name_expr)]
    predicates = []
    for item in selected:
        if not item:
            continue
        predicates.append(f"{keys[0]} = ?")
        params.append(item)
    if predicates:
        clauses.append("(" + " OR ".join(predicates) + ")")


def _movement_sql(previous_sql: str, current_sql: str) -> str:
    return f"""
        WITH previous AS ({previous_sql}),
        current AS ({current_sql}),
        previous_system AS (
            SELECT customer_code FROM customer_period_summary WHERE period = ? AND total_balance > 0
        ),
        current_system AS (
            SELECT customer_code FROM customer_period_summary WHERE period = ? AND total_balance > 0
        ),
        paired AS (
            SELECT
                COALESCE(c.officer_key, p.officer_key) AS officer_key,
                COALESCE(c.officer_code, p.officer_code) AS officer_code,
                COALESCE(c.officer_name, p.officer_name) AS officer_name,
                COALESCE(c.branch_code, p.branch_code) AS branch_code,
                COALESCE(c.transaction_office, p.transaction_office) AS transaction_office,
                COALESCE(p.customer_code, c.customer_code) AS customer_code,
                COALESCE(p.total_balance, 0) AS previous_balance,
                COALESCE(c.total_balance, 0) AS current_balance,
                COALESCE(p.debt_group_2_balance, 0) AS previous_attention_balance,
                COALESCE(c.debt_group_2_balance, 0) AS current_attention_balance,
                COALESCE(p.debt_group_3_balance + p.debt_group_4_balance + p.debt_group_5_balance, 0) AS previous_bad_debt_balance,
                COALESCE(c.debt_group_3_balance + c.debt_group_4_balance + c.debt_group_5_balance, 0) AS current_bad_debt_balance,
                CASE WHEN p.customer_code IS NULL THEN 0 ELSE 1 END AS had_officer_before,
                CASE WHEN c.customer_code IS NULL THEN 0 ELSE 1 END AS has_officer_now,
                CASE WHEN ps.customer_code IS NULL THEN 0 ELSE 1 END AS had_system_before,
                CASE WHEN cs.customer_code IS NULL THEN 0 ELSE 1 END AS has_system_now,
                COALESCE(p.nim_before_numerator, 0) AS previous_nim_before_numerator,
                COALESCE(c.nim_before_numerator, 0) AS current_nim_before_numerator,
                COALESCE(p.nim_after_numerator, 0) AS previous_nim_after_numerator,
                COALESCE(c.nim_after_numerator, 0) AS current_nim_after_numerator
            FROM current c
            LEFT JOIN previous p ON p.officer_key = c.officer_key AND p.customer_code = c.customer_code
            LEFT JOIN previous_system ps ON ps.customer_code = c.customer_code
            LEFT JOIN current_system cs ON cs.customer_code = c.customer_code
            UNION ALL
            SELECT
                COALESCE(c.officer_key, p.officer_key) AS officer_key,
                COALESCE(c.officer_code, p.officer_code) AS officer_code,
                COALESCE(c.officer_name, p.officer_name) AS officer_name,
                COALESCE(c.branch_code, p.branch_code) AS branch_code,
                COALESCE(c.transaction_office, p.transaction_office) AS transaction_office,
                COALESCE(p.customer_code, c.customer_code) AS customer_code,
                COALESCE(p.total_balance, 0) AS previous_balance,
                COALESCE(c.total_balance, 0) AS current_balance,
                COALESCE(p.debt_group_2_balance, 0) AS previous_attention_balance,
                COALESCE(c.debt_group_2_balance, 0) AS current_attention_balance,
                COALESCE(p.debt_group_3_balance + p.debt_group_4_balance + p.debt_group_5_balance, 0) AS previous_bad_debt_balance,
                COALESCE(c.debt_group_3_balance + c.debt_group_4_balance + c.debt_group_5_balance, 0) AS current_bad_debt_balance,
                CASE WHEN p.customer_code IS NULL THEN 0 ELSE 1 END AS had_officer_before,
                CASE WHEN c.customer_code IS NULL THEN 0 ELSE 1 END AS has_officer_now,
                CASE WHEN ps.customer_code IS NULL THEN 0 ELSE 1 END AS had_system_before,
                CASE WHEN cs.customer_code IS NULL THEN 0 ELSE 1 END AS has_system_now,
                COALESCE(p.nim_before_numerator, 0) AS previous_nim_before_numerator,
                COALESCE(c.nim_before_numerator, 0) AS current_nim_before_numerator,
                COALESCE(p.nim_after_numerator, 0) AS previous_nim_after_numerator,
                COALESCE(c.nim_after_numerator, 0) AS current_nim_after_numerator
            FROM previous p
            LEFT JOIN current c ON c.officer_key = p.officer_key AND c.customer_code = p.customer_code
            LEFT JOIN previous_system ps ON ps.customer_code = p.customer_code
            LEFT JOIN current_system cs ON cs.customer_code = p.customer_code
            WHERE c.customer_code IS NULL
        )
        SELECT
            officer_key,
            MAX(officer_code) AS officer_code,
            MAX(officer_name) AS officer_name,
            MAX(branch_code) AS branch_code,
            MAX(transaction_office) AS transaction_office,
            SUM(previous_balance) AS previous_balance,
            SUM(current_balance) AS current_balance,
            SUM(current_balance - previous_balance) AS balance_change,
            CASE WHEN SUM(previous_balance) <> 0 THEN SUM(current_balance - previous_balance) / SUM(previous_balance) * 100 ELSE NULL END AS growth_rate,
            COUNT(DISTINCT CASE WHEN previous_balance > 0 THEN customer_code END) AS previous_customer_count,
            COUNT(DISTINCT CASE WHEN current_balance > 0 THEN customer_code END) AS current_customer_count,
            COUNT(DISTINCT CASE WHEN had_officer_before = 0 AND has_officer_now = 1 AND had_system_before = 0 THEN customer_code END) AS new_system_customer_count,
            COUNT(DISTINCT CASE WHEN had_officer_before = 1 AND has_officer_now = 0 AND has_system_now = 0 THEN customer_code END) AS paid_off_customer_count,
            COUNT(DISTINCT CASE WHEN had_officer_before = 0 AND has_officer_now = 1 AND had_system_before = 1 THEN customer_code END) AS transfer_in_customer_count,
            COUNT(DISTINCT CASE WHEN had_officer_before = 1 AND has_officer_now = 0 AND has_system_now = 1 THEN customer_code END) AS transfer_out_customer_count,
            SUM(current_attention_balance - previous_attention_balance) AS attention_change,
            SUM(current_bad_debt_balance - previous_bad_debt_balance) AS bad_debt_change,
            CASE
                WHEN SUM(previous_balance) > 0
                    THEN SUM(previous_nim_before_numerator) / SUM(previous_balance)
                ELSE NULL
            END AS previous_nim_before,
            CASE
                WHEN SUM(previous_balance) > 0
                    THEN SUM(previous_nim_after_numerator) / SUM(previous_balance)
                ELSE NULL
            END AS previous_nim_after,
            CASE
                WHEN SUM(current_balance) > 0
                    THEN SUM(current_nim_before_numerator) / SUM(current_balance)
                ELSE NULL
            END AS current_nim_before,
            CASE
                WHEN SUM(current_balance) > 0
                    THEN SUM(current_nim_after_numerator) / SUM(current_balance)
                ELSE NULL
            END AS current_nim_after,
            CASE
                WHEN SUM(previous_balance) > 0 AND SUM(current_balance) > 0
                    THEN SUM(current_nim_before_numerator) / SUM(current_balance)
                        - SUM(previous_nim_before_numerator) / SUM(previous_balance)
                ELSE NULL
            END AS nim_before_change_pp,
            CASE
                WHEN SUM(previous_balance) > 0 AND SUM(current_balance) > 0
                    THEN SUM(current_nim_after_numerator) / SUM(current_balance)
                        - SUM(previous_nim_after_numerator) / SUM(previous_balance)
                ELSE NULL
            END AS nim_after_change_pp
        FROM paired
        GROUP BY officer_key
    """


def _officer_key_expr(code_expr: str, name_expr: str) -> str:
    return (
        "CASE "
        f"WHEN COALESCE({code_expr}, '') <> '' THEN 'CODE:' || {code_expr} "
        f"WHEN COALESCE({name_expr}, '') <> '' THEN 'NAME:' || UPPER(TRIM({name_expr})) "
        "ELSE 'UNRESOLVED' END"
    )


def _status_expr() -> str:
    return (
        "CASE "
        "WHEN officer_key = 'UNRESOLVED' OR officer_code = '' THEN 'Chưa có mã cán bộ' "
        "WHEN directory_is_active = 1 THEN 'Đang sử dụng' "
        "WHEN directory_is_active = 0 THEN 'Ngừng sử dụng' "
        "ELSE 'Chưa có trong danh mục' END"
    )


def _status_clause(status: str) -> tuple[str, list[object]]:
    key = str(status or "").strip()
    if not key:
        return "", []
    if key == "active":
        return "AND grouped.directory_is_active = 1", []
    if key == "inactive":
        return "AND grouped.directory_is_active = 0", []
    if key == "missing_directory":
        return "AND grouped.directory_is_active IS NULL AND grouped.officer_code <> ''", []
    if key == "unresolved":
        return "AND (grouped.officer_key = 'UNRESOLVED' OR grouped.officer_code = '')", []
    if key == "system":
        return (
            "AND (LOWER(grouped.officer_name) LIKE ? OR LOWER(grouped.officer_name) LIKE ? "
            "OR LOWER(grouped.officer_name) LIKE ?)",
            ["%thẻ tín dụng%", "%không xác định%", "%không rõ%"],
        )
    return "", []


def _trend_metric_expr(metric: str) -> str:
    key = str(metric or "total_balance")
    if key == "officer_count":
        return "COUNT(DISTINCT CASE WHEN total_balance > 0 THEN officer_key END)"
    if key == "customer_count":
        return "COUNT(DISTINCT CASE WHEN total_balance > 0 THEN customer_code END)"
    if key == "average_rate":
        return "CASE WHEN SUM(total_balance) <> 0 THEN SUM(interest_rate_numerator) / SUM(total_balance) ELSE NULL END"
    if key == "nim_before":
        return "CASE WHEN SUM(total_balance) <> 0 THEN SUM(nim_before_numerator) / SUM(total_balance) ELSE NULL END"
    if key == "nim_after":
        return "CASE WHEN SUM(total_balance) <> 0 THEN SUM(nim_after_numerator) / SUM(total_balance) ELSE NULL END"
    if key == "attention_ratio":
        return "CASE WHEN SUM(total_balance) <> 0 THEN SUM(debt_group_2_balance) / SUM(total_balance) * 100 ELSE NULL END"
    if key == "bad_debt_ratio":
        return "CASE WHEN SUM(total_balance) <> 0 THEN SUM(debt_group_3_balance + debt_group_4_balance + debt_group_5_balance) / SUM(total_balance) * 100 ELSE NULL END"
    return "COALESCE(SUM(total_balance), 0)"


def _comparison_values(row: dict[str, object], benchmark: dict[str, object]) -> dict[str, object]:
    average = _number(benchmark.get("benchmark_nim_after"))
    value = _number(row.get("nim_after"))
    return {
        "benchmark_nim_after": benchmark.get("benchmark_nim_after"),
        "benchmark_average_rate": benchmark.get("benchmark_average_rate"),
        "benchmark_bad_debt_ratio": benchmark.get("benchmark_bad_debt_ratio"),
        "nim_after_difference": value - average if row.get("nim_after") is not None and benchmark.get("benchmark_nim_after") is not None else None,
        "benchmark_officer_count": int(benchmark.get("benchmark_officer_count") or 0),
    }


def _order_sql(sort_by: str, sort_desc: bool, allowed: dict[str, str], *, default: str) -> str:
    column = allowed.get(str(sort_by or ""), allowed[default])
    direction = "DESC" if sort_desc else "ASC"
    return f"ORDER BY {column} {direction}, officer_name COLLATE NOCASE ASC, officer_code COLLATE NOCASE ASC"


def _finalize_numeric_row(row: dict[str, object]) -> dict[str, object]:
    if "bad_debt_balance" not in row:
        row["bad_debt_balance"] = _number(row.get("debt_group_3_balance")) + _number(row.get("debt_group_4_balance")) + _number(row.get("debt_group_5_balance"))
    if "attention_balance" not in row:
        row["attention_balance"] = _number(row.get("debt_group_2_balance"))
    total = _number(row.get("total_balance"))
    for key in ("average_rate", "nim_before", "nim_after", "attention_ratio", "bad_debt_ratio"):
        if key not in row:
            continue
        row[key] = None if row.get(key) is None else float(row.get(key) or 0)
    if "medium_long_ratio" in row and row.get("medium_long_ratio") is not None:
        row["medium_long_ratio"] = float(row.get("medium_long_ratio") or 0)
    row["has_debt_group_data"] = bool(row.get("has_debt_group_data"))
    if total:
        row.setdefault("attention_ratio", _number(row.get("attention_balance")) / total * 100)
        row.setdefault("bad_debt_ratio", _number(row.get("bad_debt_balance")) / total * 100)
    return row


def _apply_term_structure_display(row: dict[str, object]) -> dict[str, object]:
    term_keys = ("short_term_balance", "medium_long_term_balance", "other_balance")
    if not any(key in row for key in term_keys):
        return row
    total = _number(row.get("total_balance"))
    if not total and "officer_balance" in row:
        total = _number(row.get("officer_balance"))
    term_total = sum(_number(row.get(key)) for key in term_keys)
    if "term_structure_available" in row:
        available = bool(row.get("term_structure_available"))
    else:
        available = total <= 0 or abs(term_total - total) <= TERM_STRUCTURE_TOLERANCE
    row["term_structure_available"] = available
    if available:
        if "medium_long_ratio" in row and row.get("medium_long_ratio") is None and total:
            row["medium_long_ratio"] = _number(row.get("medium_long_term_balance")) / total * 100
        return row
    row["term_structure_warning"] = TERM_STRUCTURE_WARNING
    for key in (*term_keys, "medium_long_ratio"):
        if key in row:
            row[key] = None
    existing_warning = str(row.get("data_warning") or "").strip()
    term_warning = "Dữ liệu kỳ hạn theo CBTD chưa đầy đủ"
    row["data_warning"] = f"{existing_warning}; {term_warning}" if existing_warning else term_warning
    return row


def _status_from_row(row: dict[str, object]) -> str:
    key = str(row.get("officer_key") or "")
    code = str(row.get("officer_code") or "")
    is_active = row.get("directory_is_active")
    if key == "UNRESOLVED" or not code:
        return "Chưa có mã cán bộ"
    if is_active == 1:
        return "Đang sử dụng"
    if is_active == 0:
        return "Ngừng sử dụng"
    return "Chưa có trong danh mục"


def _mode_label(mode: str) -> str:
    return "Theo cán bộ hiệu lực" if mode == OFFICER_MODE_EFFECTIVE else "Theo phân bổ import"


def _normalize_selected_officer(value: object) -> str:
    text = str(value or "").strip()
    if not text:
        return ""
    return text if ":" in text else f"CODE:{text}"


def _row_selected(row: dict[str, object], selected: set[str]) -> bool:
    keys = {
        str(row.get("officer_key") or ""),
        f"CODE:{str(row.get('officer_code') or '').strip()}",
        str(row.get("officer_code") or "").strip(),
    }
    return bool(keys & selected)


def _ratio(numerator: float, denominator: float) -> float | None:
    return numerator / denominator if denominator else None


def _number(value: object) -> float:
    try:
        return float(value or 0)
    except (TypeError, ValueError):
        return 0.0


def _is_system_officer(name: object) -> bool:
    normalized = _strip_accents(str(name or "")).casefold().strip()
    return normalized in SYSTEM_OFFICER_NAMES


def _strip_accents(value: str) -> str:
    decomposed = unicodedata.normalize("NFD", value)
    return "".join(char for char in decomposed if unicodedata.category(char) != "Mn")


_OFFICER_SORTS = {
    "officer_code": "officer_code",
    "officer_name": "officer_name",
    "branch_code": "branch_code",
    "transaction_office": "transaction_office",
    "customer_count": "customer_count",
    "total_balance": "total_balance",
    "short_term_balance": "short_term_balance",
    "medium_long_term_balance": "medium_long_term_balance",
    "other_balance": "other_balance",
    "medium_long_ratio": "medium_long_ratio",
    "average_rate": "average_rate",
    "nim_before": "nim_before",
    "nim_after": "nim_after",
    "attention_balance": "debt_group_2_balance",
    "bad_debt_balance": "bad_debt_balance",
    "attention_ratio": "attention_ratio",
    "bad_debt_ratio": "bad_debt_ratio",
    "multiple_officer_customer_count": "multiple_officer_customer_count",
    "override_customer_count": "override_customer_count",
}

_CUSTOMER_SORTS = {
    "period": "period",
    "customer_code": "customer_code",
    "customer_name": "customer_name",
    "total_customer_balance": "total_customer_balance",
    "officer_balance": "officer_balance",
    "officer_share": "officer_share",
    "attention_balance": "attention_balance",
    "bad_debt_balance": "bad_debt_balance",
    "nim_after": "nim_after",
}

_MOVEMENT_SORTS = {
    "officer_code": "officer_code",
    "officer_name": "officer_name",
    "previous_balance": "previous_balance",
    "current_balance": "current_balance",
    "balance_change": "balance_change",
    "growth_rate": "growth_rate",
    "new_system_customer_count": "new_system_customer_count",
    "paid_off_customer_count": "paid_off_customer_count",
    "transfer_in_customer_count": "transfer_in_customer_count",
    "transfer_out_customer_count": "transfer_out_customer_count",
    "attention_change": "attention_change",
    "bad_debt_change": "bad_debt_change",
    "previous_nim_before": "previous_nim_before",
    "previous_nim_after": "previous_nim_after",
    "current_nim_before": "current_nim_before",
    "current_nim_after": "current_nim_after",
    "nim_before_change_pp": "nim_before_change_pp",
    "nim_after_change_pp": "nim_after_change_pp",
}


_TOP_METRICS = {
    "total_balance",
    "balance_change",
    "nim_after",
    "attention_balance",
    "bad_debt_balance",
    "bad_debt_ratio",
}


def _top_metric(metric: object) -> str:
    key = str(metric or "total_balance").strip()
    return key if key in _TOP_METRICS else "total_balance"


def _top_limit(limit: object) -> int:
    try:
        value = int(limit or 10)
    except (TypeError, ValueError):
        value = 10
    return min(50, max(10, value))
