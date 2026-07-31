from __future__ import annotations

from contextlib import closing
import re

from agribank_v3.features.credit.summary.models import SummaryDataType, SummaryError
from agribank_v3.features.credit.summary.repository import SummaryRepository

from .models import HistoryFilters, OfficerKey


NIM_OFFICER_DISPLAY_SQL = (
    "CASE WHEN officer_code <> '' "
    "THEN '[' || officer_code || '] ' || officer_name "
    "ELSE officer_name END"
)


class OfficerHistoryRepository:
    def __init__(self, repository: SummaryRepository) -> None:
        self.repository = repository
        self.unit_directory = repository.unit_directory

    def get_officer_history(
        self,
        data_type: SummaryDataType,
        *,
        officer_code: str = "",
        officer: str = "",
        branch: str = "",
        filters: HistoryFilters | None = None,
    ) -> list[dict[str, object]]:
        filters = filters or HistoryFilters()
        where, params = self._history_where(
            data_type,
            officer_code=officer_code,
            officer=officer,
            branch=branch,
            filters=filters,
        )
        return self._period_rows(where, params)

    def get_available_officers(
        self,
        data_type: SummaryDataType,
        *,
        branch: str = "",
        transaction_office: str = "",
    ) -> list[OfficerKey]:
        clauses = ["data_type = ?", "officer_name <> ''"]
        params: list[object] = [data_type.value]
        branch_code = self._add_branch_clause(clauses, params, branch)
        if transaction_office:
            self._add_transaction_office_clause(clauses, params, transaction_office, branch_code=branch_code)
        with closing(self.repository.connect()) as database:
            rows = database.execute(
                f"""
                SELECT
                    {NIM_OFFICER_DISPLAY_SQL} AS officer,
                    branch_code,
                    trctcd,
                    MIN(branch_name) AS branch_name,
                    MIN(transaction_office) AS transaction_office,
                    SUM(balance) AS balance
                FROM nim_period_summary
                WHERE {' AND '.join(clauses)}
                GROUP BY officer_code, officer_name, branch_code, trctcd
                ORDER BY balance DESC, officer COLLATE NOCASE
                """,
                params,
            ).fetchall()
        officers: dict[str, OfficerKey] = {}
        for row in rows:
            key = officer_key(
                str(row["officer"] or ""),
                self._branch_display(row["branch_code"], row["branch_name"]),
                self._office_name(row["branch_code"], row["trctcd"], row["transaction_office"]),
            )
            officers.setdefault(key.raw_name, key)
        return list(officers.values())

    def get_customer_types(
        self,
        data_type: SummaryDataType,
        *,
        officer_code: str = "",
        officer: str = "",
        branch: str = "",
    ) -> list[str]:
        clauses = ["data_type = ?", "customer_type <> ''"]
        params: list[object] = [data_type.value]
        self._add_officer_clause(clauses, params, officer_code=officer_code, officer=officer)
        self._add_branch_clause(clauses, params, branch)
        with closing(self.repository.connect()) as database:
            rows = database.execute(
                f"""
                SELECT DISTINCT customer_type
                FROM nim_period_summary
                WHERE {' AND '.join(clauses)}
                ORDER BY customer_type COLLATE NOCASE
                """,
                params,
            ).fetchall()
        return [str(row["customer_type"] or "") for row in rows]

    def get_periods(
        self,
        data_type: SummaryDataType,
        *,
        officer_code: str = "",
        officer: str = "",
        branch: str = "",
    ) -> list[str]:
        clauses = ["data_type = ?", "period <> ''"]
        params: list[object] = [data_type.value]
        self._add_officer_clause(clauses, params, officer_code=officer_code, officer=officer)
        self._add_branch_clause(clauses, params, branch)
        with closing(self.repository.connect()) as database:
            rows = database.execute(
                f"""
                SELECT DISTINCT period
                FROM nim_period_summary
                WHERE {' AND '.join(clauses)}
                ORDER BY period
                """,
                params,
            ).fetchall()
        return [str(row["period"] or "") for row in rows]

    def get_multiple_officer_history(
        self,
        data_type: SummaryDataType,
        *,
        officers: list[OfficerKey],
        branch: str = "",
        filters: HistoryFilters | None = None,
    ) -> list[dict[str, object]]:
        filters = filters or HistoryFilters()
        if not officers:
            return []
        officer_clauses: list[str] = []
        params: list[object] = [data_type.value]
        for item in officers:
            if item.code:
                officer_clauses.append("officer_code = ?")
                params.append(item.code)
            else:
                officer_clauses.append("officer_name = ?")
                params.append(item.display_name or item.raw_name)
        clauses = ["data_type = ?", f"({' OR '.join(officer_clauses)})"]
        branch_code = self._add_branch_clause(clauses, params, branch)
        self._add_filter_clauses(clauses, params, filters, branch_code=branch_code)
        with closing(self.repository.connect()) as database:
            rows = database.execute(
                f"""
                SELECT
                    period,
                    {NIM_OFFICER_DISPLAY_SQL} AS officer,
                    branch_code,
                    trctcd,
                    MIN(branch_name) AS branch,
                    MIN(transaction_office) AS transaction_office,
                    {'customer_type' if filters.customer_type else "'Tất cả' AS customer_type"},
                    SUM(balance) AS balance,
                    CASE WHEN SUM(balance) <> 0 THEN SUM(interest_rate_numerator) / SUM(balance) ELSE 0 END AS average_rate,
                    CASE WHEN SUM(balance) <> 0 THEN SUM(numerator_before) / SUM(balance) ELSE 0 END AS nim_before,
                    CASE WHEN SUM(balance) <> 0 THEN SUM(numerator_after) / SUM(balance) ELSE 0 END AS nim_after
                FROM nim_period_summary
                WHERE {' AND '.join(clauses)}
                GROUP BY period, officer_code, officer_name, branch_code, trctcd{', customer_type' if filters.customer_type else ''}
                ORDER BY period, officer COLLATE NOCASE
                """,
                params,
            ).fetchall()
        return [self._dynamic_row(dict(row)) for row in rows]

    def get_branch_history(
        self,
        data_type: SummaryDataType,
        *,
        branch: str,
        filters: HistoryFilters | None = None,
    ) -> list[dict[str, object]]:
        filters = filters or HistoryFilters()
        if not branch:
            return []
        clauses = ["data_type = ?"]
        params: list[object] = [data_type.value]
        branch_code = self._add_branch_clause(clauses, params, branch)
        self._add_filter_clauses(clauses, params, filters, branch_code=branch_code)
        return self._period_rows("WHERE " + " AND ".join(clauses), params)

    def _period_rows(self, where: str, params: list[object]) -> list[dict[str, object]]:
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

    def _history_where(
        self,
        data_type: SummaryDataType,
        *,
        officer_code: str = "",
        officer: str = "",
        branch: str = "",
        filters: HistoryFilters,
    ) -> tuple[str, list[object]]:
        if not officer_code and not officer:
            raise SummaryError("Thiếu cán bộ để truy vấn lịch sử NIM.")
        clauses = ["data_type = ?"]
        params: list[object] = [data_type.value]
        self._add_officer_clause(clauses, params, officer_code=officer_code, officer=officer)
        branch_code = self._add_branch_clause(clauses, params, branch)
        self._add_filter_clauses(clauses, params, filters, branch_code=branch_code)
        return "WHERE " + " AND ".join(clauses), params

    def _add_branch_clause(self, clauses: list[str], params: list[object], branch: object) -> str:
        text = str(branch or "").strip()
        if not text:
            return ""
        branch_code = self._branch_code_from_filter(text)
        if branch_code:
            clauses.append("branch_code = ?")
            params.append(branch_code)
            return branch_code
        clauses.append("branch_name = ?")
        params.append(text)
        return ""

    def _branch_code_from_filter(self, value: object) -> str:
        text = str(value or "").strip()
        if not text:
            return ""
        if self.unit_directory.get_branch(text) is not None:
            return text
        prefix, separator, _label = text.partition(" - ")
        if separator and prefix.strip() and self.unit_directory.get_branch(prefix.strip()) is not None:
            return prefix.strip()
        for branch in self.unit_directory.repository.list_branches(active_only=False):
            if self.unit_directory.format_branch_display(branch.branch_code) == text:
                return branch.branch_code
        return ""

    def _add_transaction_office_clause(
        self,
        clauses: list[str],
        params: list[object],
        value: object,
        *,
        branch_code: str = "",
    ) -> None:
        text = str(value or "").strip()
        if not text:
            return
        office_code, separator, _label = text.partition(" - ")
        if separator and "-" in office_code:
            branch, _sep, trctcd = office_code.partition("-")
            if branch and trctcd:
                clauses.append("branch_code = ?")
                clauses.append("trctcd = ?")
                params.extend([branch, trctcd])
                return
        if branch_code:
            for office in self.unit_directory.repository.list_offices(branch_code=branch_code, active_only=False):
                labels = {
                    office.office_name,
                    office.short_name,
                    self.unit_directory.get_office_name(office.branch_code, office.trctcd),
                    self.unit_directory.get_office_display_name(office.branch_code, office.trctcd),
                }
                if text in labels:
                    clauses.append("trctcd = ?")
                    params.append(office.trctcd)
                    return
        clauses.append("transaction_office = ?")
        params.append(text)

    def _branch_display(self, branch_code: object, snapshot: object = "") -> str:
        code = str(branch_code or "").strip()
        if code:
            return self.unit_directory.get_branch_display_name(code)
        return str(snapshot or "")

    def _office_name(self, branch_code: object, trctcd: object, snapshot: object = "") -> str:
        branch = str(branch_code or "").strip()
        code = str(trctcd or "").strip()
        if branch and code:
            return self.unit_directory.get_office_name(branch, code)
        return str(snapshot or "")

    def _dynamic_row(self, row: dict[str, object]) -> dict[str, object]:
        row["branch"] = self._branch_display(row.get("branch_code"), row.get("branch"))
        row["transaction_office"] = self._office_name(row.get("branch_code"), row.get("trctcd"), row.get("transaction_office"))
        row.pop("branch_code", None)
        row.pop("trctcd", None)
        return row

    @staticmethod
    def _add_officer_clause(
        clauses: list[str],
        params: list[object],
        *,
        officer_code: str = "",
        officer: str = "",
    ) -> None:
        if officer_code:
            clauses.append("officer_code = ?")
            params.append(officer_code)
        elif officer:
            match = re.match(r"^\s*\[([^\]]+)\]", str(officer or ""))
            parsed_code = match.group(1).strip() if match else ""
            if parsed_code:
                clauses.append("officer_code = ?")
                params.append(parsed_code)
            else:
                clauses.append("officer_name = ?")
                params.append(officer_display_name(officer))

    def _add_filter_clauses(
        self,
        clauses: list[str],
        params: list[object],
        filters: HistoryFilters,
        *,
        branch_code: str = "",
    ) -> None:
        if filters.period_from:
            clauses.append("period >= ?")
            params.append(filters.period_from)
        if filters.period_to:
            clauses.append("period <= ?")
            params.append(filters.period_to)
        if filters.customer_type:
            clauses.append("customer_type = ?")
            params.append(filters.customer_type)
        if filters.transaction_office:
            self._add_transaction_office_clause(
                clauses,
                params,
                filters.transaction_office,
                branch_code=branch_code,
            )


def officer_key(raw_name: str, branch: str = "", transaction_office: str = "") -> OfficerKey:
    code = officer_code(raw_name)
    return OfficerKey(
        code=code,
        raw_name=raw_name,
        display_name=officer_display_name(raw_name),
        branch=branch,
        transaction_office=transaction_office,
    )


def officer_code(raw_name: str) -> str:
    match = re.match(r"^\s*\[([^\]]+)\]", str(raw_name or ""))
    return match.group(1).strip() if match else ""


def officer_display_name(raw_name: str) -> str:
    text = str(raw_name or "").strip()
    if text.startswith("[") and "]" in text:
        return text.split("]", 1)[1].strip()
    return text
