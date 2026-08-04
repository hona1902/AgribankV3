from __future__ import annotations

from contextlib import closing
from collections.abc import Mapping, Sequence

from agribank_v3.features.credit.summary.models import SummaryDataType, SummaryError
from agribank_v3.features.credit.summary.repository import SummaryRepository
from agribank_v3.features.settings.unit_directory.service import get_unit_directory_service


class NimDashboardRepository:
    """Read-only dashboard queries for NIM dư nợ."""

    def __init__(self, repository: SummaryRepository) -> None:
        self.repository = repository
        self.unit_directory = get_unit_directory_service(repository.main_database_path)

    def distinct_values(
        self,
        data_type: SummaryDataType,
        column_name: str,
        *,
        filters: Mapping[str, object] | None = None,
        exclude: str = "",
    ) -> list[str] | list[tuple[str, str]]:
        allowed = {
            "period": "period",
            "branch": "branch_code",
            "transaction_office": "trctcd",
            "customer_type": "customer_type",
        }
        if column_name not in allowed:
            raise SummaryError("Trường lọc Dashboard NIM không hợp lệ.")
        where, params = self._where(data_type, filters or {}, exclude=exclude)
        database_column = allowed[column_name]
        select_columns = f"{database_column} AS value"
        order_columns = "value COLLATE NOCASE"
        if column_name == "transaction_office":
            select_columns = "branch_code, trctcd AS value"
            order_columns = "branch_code COLLATE NOCASE, value COLLATE NOCASE"
        with closing(self.repository.connect()) as database:
            rows = database.execute(
                f"""
                SELECT DISTINCT {select_columns}
                FROM nim_period_summary
                {where} AND {database_column} <> ''
                ORDER BY {order_columns}
                """,
                params,
            ).fetchall()
        if column_name == "branch":
            return [
                (self.unit_directory.get_branch_display_name(row["value"]), str(row["value"] or ""))
                for row in rows
            ]
        if column_name == "transaction_office":
            return [
                (
                    self.unit_directory.get_office_display_name(row["branch_code"], row["value"])
                    if "branch_code" in row.keys()
                    else str(row["value"] or ""),
                    (
                        f"{str(row['branch_code'] or '')}-{str(row['value'] or '')}"
                        if "branch_code" in row.keys()
                        else str(row["value"] or "")
                    ),
                )
                for row in rows
            ]
        return [str(row["value"] or "") for row in rows]

    def period_summary(self, data_type: SummaryDataType, filters: Mapping[str, object]) -> list[dict[str, object]]:
        where, params = self._where(data_type, filters)
        with closing(self.repository.connect()) as database:
            rows = database.execute(
                f"""
                SELECT
                    period,
                    SUM(balance) AS balance,
                    CASE WHEN SUM(balance) <> 0 THEN SUM(interest_rate_numerator) / SUM(balance) ELSE 0 END AS average_rate,
                    CASE WHEN SUM(balance) <> 0 THEN SUM(numerator_before) / SUM(balance) ELSE 0 END AS nim_before,
                    CASE WHEN SUM(balance) <> 0 THEN SUM(numerator_after) / SUM(balance) ELSE 0 END AS nim_after
                FROM nim_period_summary
                {where}
                GROUP BY period
                ORDER BY period
                """,
                params,
            ).fetchall()
        return [dict(row) for row in rows]

    def branch_period_summary(self, data_type: SummaryDataType, filters: Mapping[str, object]) -> list[dict[str, object]]:
        where, params = self._where(data_type, filters)
        with closing(self.repository.connect()) as database:
            rows = database.execute(
                f"""
                SELECT
                    period,
                    branch_code,
                    MIN(branch_name) AS branch,
                    SUM(balance) AS balance,
                    CASE WHEN SUM(balance) <> 0 THEN SUM(interest_rate_numerator) / SUM(balance) ELSE 0 END AS average_rate,
                    CASE WHEN SUM(balance) <> 0 THEN SUM(numerator_before) / SUM(balance) ELSE 0 END AS nim_before,
                    CASE WHEN SUM(balance) <> 0 THEN SUM(numerator_after) / SUM(balance) ELSE 0 END AS nim_after
                FROM nim_period_summary
                {where}
                GROUP BY period, branch_code
                ORDER BY period, branch_code COLLATE NOCASE
                """,
                params,
            ).fetchall()
        return [self._dynamic_unit_row(dict(row)) for row in rows]

    def detail_summary(
        self,
        data_type: SummaryDataType,
        filters: Mapping[str, object],
        *,
        limit: int | None = None,
        offset: int = 0,
    ) -> list[dict[str, object]]:
        where, params = self._where(data_type, filters)
        has_customer_type = bool(str(filters.get("customer_type") or "").strip())
        customer_select = "customer_type" if has_customer_type else "'Tất cả' AS customer_type"
        group_by = "period, branch_code, trctcd"
        if has_customer_type:
            group_by += ", customer_type"
        query_params = list(params)
        paging_sql = ""
        if limit is not None:
            paging_sql = " LIMIT ? OFFSET ?"
            query_params.extend([max(0, int(limit)), max(0, int(offset))])
        with closing(self.repository.connect()) as database:
            rows = database.execute(
                f"""
                SELECT
                    period,
                    branch_code,
                    trctcd,
                    MIN(branch_name) AS branch,
                    MIN(transaction_office) AS transaction_office,
                    {customer_select},
                    SUM(balance) AS balance,
                    CASE WHEN SUM(balance) <> 0 THEN SUM(interest_rate_numerator) / SUM(balance) ELSE 0 END AS average_rate,
                    CASE WHEN SUM(balance) <> 0 THEN SUM(numerator_before) / SUM(balance) ELSE 0 END AS nim_before,
                    CASE WHEN SUM(balance) <> 0 THEN SUM(numerator_after) / SUM(balance) ELSE 0 END AS nim_after
                FROM nim_period_summary
                {where}
                GROUP BY {group_by}
                ORDER BY period, branch_code COLLATE NOCASE, trctcd COLLATE NOCASE, customer_type COLLATE NOCASE
                {paging_sql}
                """,
                query_params,
            ).fetchall()
        return [self._dynamic_unit_row(dict(row)) for row in rows]

    def detail_summary_count(self, data_type: SummaryDataType, filters: Mapping[str, object]) -> int:
        where, params = self._where(data_type, filters)
        has_customer_type = bool(str(filters.get("customer_type") or "").strip())
        group_by = "period, branch_code, trctcd"
        if has_customer_type:
            group_by += ", customer_type"
        with closing(self.repository.connect()) as database:
            row = database.execute(
                f"""
                SELECT COUNT(*) AS count_rows
                FROM (
                    SELECT 1
                    FROM nim_period_summary
                    {where}
                    GROUP BY {group_by}
                ) AS grouped_detail
                """,
                params,
            ).fetchone()
        return int(row["count_rows"] if row else 0)

    def _where(
        self,
        data_type: SummaryDataType,
        filters: Mapping[str, object],
        *,
        exclude: str = "",
    ) -> tuple[str, list[object]]:
        clauses = ["data_type = ?"]
        params: list[object] = [data_type.value]
        period_from = str(filters.get("period_from") or "").strip()
        period_to = str(filters.get("period_to") or "").strip()
        if exclude != "period_from" and period_from:
            clauses.append("period >= ?")
            params.append(period_from)
        if exclude != "period_to" and period_to:
            clauses.append("period <= ?")
            params.append(period_to)
        self._add_value_filter(
            clauses,
            params,
            "branch_code",
            self._branch_filter_value(filters.get("branch")),
            exclude == "branch",
        )
        if exclude != "transaction_office":
            office_value = str(filters.get("transaction_office") or "").strip()
            if office_value:
                branch, sep, trctcd = office_value.partition("-")
                if sep:
                    clauses.append("branch_code = ?")
                    clauses.append("trctcd = ?")
                    params.extend([branch, trctcd])
                else:
                    clauses.append("transaction_office = ?")
                    params.append(office_value)
        self._add_value_filter(clauses, params, "customer_type", filters.get("customer_type"), exclude == "customer_type")
        return "WHERE " + " AND ".join(clauses), params

    def _branch_filter_value(self, value: object) -> object:
        if isinstance(value, str):
            return _branch_code_from_filter(value)
        if isinstance(value, Sequence):
            return [_branch_code_from_filter(item) for item in value]
        return value

    def _dynamic_unit_row(self, row: dict[str, object]) -> dict[str, object]:
        branch_code = str(row.get("branch_code") or "").strip()
        trctcd = str(row.get("trctcd") or "").strip()
        if branch_code:
            row["branch"] = self.unit_directory.get_branch_display_name(branch_code)
        if branch_code and trctcd:
            row["transaction_office"] = self.unit_directory.get_office_name(branch_code, trctcd)
        return row

    @staticmethod
    def _add_value_filter(
        clauses: list[str],
        params: list[object],
        column_name: str,
        value: object,
        excluded: bool,
    ) -> None:
        if excluded:
            return
        if isinstance(value, str):
            clean = value.strip()
            if not clean or clean == "Tất cả":
                return
            clauses.append(f"{column_name} = ?")
            params.append(clean)
            return
        if isinstance(value, Sequence):
            values = [str(item).strip() for item in value if str(item).strip()]
            if not values:
                return
            placeholders = ", ".join("?" for _item in values)
            clauses.append(f"{column_name} IN ({placeholders})")
            params.extend(values)


def _branch_code_from_filter(value: object) -> str:
    text = "" if value is None else str(value).strip()
    if not text or text == "Tất cả":
        return ""
    prefix, separator, _label = text.partition(" - ")
    if separator and prefix.strip():
        return prefix.strip()
    return text
