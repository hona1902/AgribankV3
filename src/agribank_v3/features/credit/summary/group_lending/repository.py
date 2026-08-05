from __future__ import annotations

from contextlib import closing
from dataclasses import replace
from pathlib import Path
import sqlite3

from agribank_v3.features.credit.summary.credit_report import CreditReportRepository, _officer_display
from agribank_v3.features.credit.summary.models import PageResult
from agribank_v3.features.credit.tovayvon.models import (
    ASSOCIATION_FARMERS_UNION,
    ASSOCIATION_OTHER,
    ASSOCIATION_TYPE_LABELS,
    ASSOCIATION_WOMENS_UNION,
    CreditGroup,
)
from agribank_v3.features.credit.tovayvon.repository import CreditGroupRepository

from .models import (
    ASSOCIATION_FILTER_OPTIONS,
    ASSOCIATION_UNKNOWN,
    ASSOCIATION_UNKNOWN_LABEL,
    GROUP_STATUS_DECLARED,
    GROUP_STATUS_INACTIVE,
    GROUP_STATUS_LABELS,
    GROUP_STATUS_NOT_DECLARED,
    GroupAssociationComparisonRow,
    GroupAssociationSummaryRow,
    GroupDirectoryEntry,
    GroupLendingComparisonRow,
    GroupLendingFilters,
    GroupLendingKpi,
    GroupLendingResult,
    GroupLendingRow,
)


class GroupLendingRepository:
    def __init__(self, main_database_path: Path | None = None) -> None:
        self.credit_repository = CreditReportRepository(main_database_path)
        self.group_repository = CreditGroupRepository(self.credit_repository.group_directory_database_path())
        self.group_directory_load_count = 0

    def filter_values(self) -> dict[str, list[tuple[str, str]]]:
        with closing(self.credit_repository.connect()) as connection:
            branches = [
                (str(row[0]), str(row[0]))
                for row in connection.execute(
                    """
                    SELECT DISTINCT branch_code
                    FROM credit_loan_period
                    WHERE branch_code <> ''
                    ORDER BY branch_code COLLATE NOCASE
                    """
                ).fetchall()
            ]
            officers = [
                (display, display)
                for display in sorted(
                    {
                        _officer_display(str(row["officer_code"] or ""), str(row["officer_name"] or ""))
                        for row in connection.execute(
                            """
                            SELECT DISTINCT officer_code, officer_name
                            FROM credit_loan_period
                            WHERE officer_code <> '' OR officer_name <> ''
                            """
                        ).fetchall()
                    }
                )
                if display
            ]
        return {
            "branches": branches,
            "offices": [],
            "officers": officers,
            "association_types": list(ASSOCIATION_FILTER_OPTIONS),
            "statuses": [
                ("Tất cả", ""),
                (GROUP_STATUS_LABELS[GROUP_STATUS_DECLARED], GROUP_STATUS_DECLARED),
                (GROUP_STATUS_LABELS[GROUP_STATUS_INACTIVE], GROUP_STATUS_INACTIVE),
                (GROUP_STATUS_LABELS[GROUP_STATUS_NOT_DECLARED], GROUP_STATUS_NOT_DECLARED),
            ],
        }

    def get_group_lending_snapshot(
        self,
        period: str,
        filters: GroupLendingFilters,
        *,
        page: int = 1,
        page_size: int = 100,
        sort_by: str = "total_balance",
        sort_desc: bool = True,
    ) -> GroupLendingResult:
        rows = self._group_rows(period, filters)
        rows = self._sort_group_rows(rows, sort_by=sort_by, sort_desc=sort_desc)
        total_rows = len(rows)
        page, page_size, page_rows = _page(rows, page, page_size)
        kpis = self._kpis(period, rows, filters)
        notes = self._period_notes(period, rows)
        return GroupLendingResult(
            period=period,
            rows=tuple(page_rows),
            kpis=kpis,
            total_rows=total_rows,
            page=page,
            page_size=page_size,
            notes=notes,
            diagnostics=self._diagnostics(period, filters),
        )

    def get_group_detail_page(
        self,
        period: str,
        filters: GroupLendingFilters,
        *,
        page: int = 1,
        page_size: int = 100,
        sort_by: str = "total_balance",
        sort_desc: bool = True,
    ) -> GroupLendingResult:
        return self.get_group_lending_snapshot(
            period,
            filters,
            page=page,
            page_size=page_size,
            sort_by=sort_by,
            sort_desc=sort_desc,
        )

    def get_group_member_page(
        self,
        period: str,
        group_code: str,
        filters: GroupLendingFilters,
        *,
        page: int = 1,
        page_size: int = 100,
        sort_by: str = "total_balance",
        sort_desc: bool = True,
    ) -> PageResult:
        clean_group = str(group_code or "").strip()
        if not period or not clean_group:
            return PageResult(rows=[], total_rows=0, page=max(1, page), page_size=max(1, page_size))
        where, params = self._loan_where_clause(replace(filters, period=period), require_group_code=False)
        where.append("l.group_code = ?")
        params.append(clean_group)
        order = _member_order_clause(sort_by, sort_desc)
        page = max(1, int(page or 1))
        page_size = max(1, int(page_size or 100))
        offset = (page - 1) * page_size
        where_sql = " AND ".join(where)
        query = f"""
            WITH member AS (
                SELECT
                    l.customer_id,
                    c.customer_code,
                    c.customer_name,
                    l.customer_type_code,
                    l.officer_code,
                    l.officer_name,
                    COUNT(DISTINCT l.loan_key) AS loan_count,
                    SUM(l.outstanding_balance) AS total_balance,
                    MAX(CASE
                        WHEN l.debt_group_code IN ('01', '1') THEN 1
                        WHEN l.debt_group_code IN ('02', '2') THEN 2
                        WHEN l.debt_group_code IN ('03', '3') THEN 3
                        WHEN l.debt_group_code IN ('04', '4') THEN 4
                        WHEN l.debt_group_code IN ('05', '5') THEN 5
                        ELSE 0
                    END) AS worst_debt_group,
                    SUM(CASE WHEN l.term_category = 'SHORT' THEN l.outstanding_balance ELSE 0 END) AS short_term_balance,
                    SUM(CASE WHEN l.term_category = 'MEDIUM' THEN l.outstanding_balance ELSE 0 END) AS medium_term_balance,
                    SUM(CASE WHEN l.term_category = 'LONG' THEN l.outstanding_balance ELSE 0 END) AS long_term_balance
                FROM credit_loan_period l
                JOIN credit_customer_master c ON c.id = l.customer_id
                WHERE {where_sql}
                GROUP BY l.customer_id, c.customer_code, c.customer_name, l.customer_type_code, l.officer_code, l.officer_name
                HAVING SUM(l.outstanding_balance) > 0
            )
            SELECT *
            FROM member
            ORDER BY {order}
            LIMIT ? OFFSET ?
        """
        count_query = f"""
            WITH member AS (
                SELECT l.customer_id, SUM(l.outstanding_balance) AS total_balance
                FROM credit_loan_period l
                JOIN credit_customer_master c ON c.id = l.customer_id
                WHERE {where_sql}
                GROUP BY l.customer_id
                HAVING SUM(l.outstanding_balance) > 0
            )
            SELECT COUNT(*) FROM member
        """
        with closing(self.credit_repository.connect()) as connection:
            total_rows = int(connection.execute(count_query, params).fetchone()[0] or 0)
            rows = connection.execute(query, [*params, page_size, offset]).fetchall()
        return PageResult(
            rows=[_member_row_to_dict(row) for row in rows],
            total_rows=total_rows,
            page=page,
            page_size=page_size,
        )

    def get_association_summary(self, period: str, filters: GroupLendingFilters) -> GroupLendingResult:
        group_rows = self._group_rows(period, filters)
        rows = self._association_rows(group_rows, filters)
        return GroupLendingResult(
            period=period,
            rows=tuple(rows),
            kpis=self._kpis(period, group_rows, filters),
            total_rows=len(rows),
            notes=self._period_notes(period, group_rows),
            diagnostics=self._diagnostics(period, filters),
        )

    def compare_groups(self, from_period: str, to_period: str, filters: GroupLendingFilters) -> GroupLendingResult:
        from_rows = {row.group_code: row for row in self._group_rows(from_period, replace(filters, period=from_period))}
        to_rows = {row.group_code: row for row in self._group_rows(to_period, replace(filters, period=to_period))}
        output: list[GroupLendingComparisonRow] = []
        for code in sorted(set(from_rows) | set(to_rows)):
            before = from_rows.get(code)
            after = to_rows.get(code)
            display = after or before
            if display is None:
                continue
            before_balance = before.total_balance if before else 0.0
            after_balance = after.total_balance if after else 0.0
            diff = after_balance - before_balance
            growth = None if before_balance == 0 else diff / before_balance * 100
            before_members = before.member_count if before else 0
            after_members = after.member_count if after else 0
            output.append(
                GroupLendingComparisonRow(
                    group_code=code,
                    group_name=display.group_name,
                    association_label=display.association_label,
                    member_count_from=before_members,
                    member_count_to=after_members,
                    member_change=after_members - before_members,
                    balance_from=before_balance,
                    balance_to=after_balance,
                    balance_change=diff,
                    balance_growth_rate=growth,
                    movement_category=_movement_category(before_balance, after_balance),
                )
            )
        output.sort(key=lambda row: abs(row.balance_change), reverse=True)
        kpis = self._comparison_kpis(from_period, to_period, tuple(from_rows.values()), tuple(to_rows.values()))
        return GroupLendingResult(
            period=to_period,
            rows=tuple(output),
            kpis=kpis,
            total_rows=len(output),
            notes=_compare_notes(from_period, to_period),
            diagnostics=self._diagnostics(to_period, replace(filters, period=to_period)),
        )

    def compare_associations(self, from_period: str, to_period: str, filters: GroupLendingFilters) -> GroupLendingResult:
        from_filters = replace(filters, period=from_period)
        to_filters = replace(filters, period=to_period)
        from_rows = self._association_rows(self._group_rows(from_period, from_filters), from_filters)
        to_rows = self._association_rows(self._group_rows(to_period, to_filters), to_filters)
        before_by_key = {row.association_type: row for row in from_rows}
        after_by_key = {row.association_type: row for row in to_rows}
        before_total = sum(row.total_balance for row in from_rows)
        after_total = sum(row.total_balance for row in to_rows)
        output: list[GroupAssociationComparisonRow] = []
        for key in _association_order():
            before = before_by_key.get(key)
            after = after_by_key.get(key)
            before_balance = before.total_balance if before else 0.0
            after_balance = after.total_balance if after else 0.0
            diff = after_balance - before_balance
            before_share = _ratio(before_balance, before_total)
            after_share = _ratio(after_balance, after_total)
            output.append(
                GroupAssociationComparisonRow(
                    association_type=key,
                    association_label=_association_label(key),
                    group_count_from=before.group_count if before else 0,
                    group_count_to=after.group_count if after else 0,
                    group_count_change=(after.group_count if after else 0) - (before.group_count if before else 0),
                    member_count_from=before.unique_member_count if before else 0,
                    member_count_to=after.unique_member_count if after else 0,
                    member_count_change=(after.unique_member_count if after else 0) - (before.unique_member_count if before else 0),
                    balance_from=before_balance,
                    balance_to=after_balance,
                    balance_change=diff,
                    growth_rate=None if before_balance == 0 else diff / before_balance * 100,
                    share_from=before_share,
                    share_to=after_share,
                    share_change_pp=None if before_share is None or after_share is None else after_share - before_share,
                )
            )
        return GroupLendingResult(
            period=to_period,
            rows=tuple(output),
            kpis=self._comparison_kpis(from_period, to_period, tuple(before_by_key.values()), tuple(after_by_key.values())),
            total_rows=len(output),
            notes=_compare_notes(from_period, to_period),
            diagnostics=self._diagnostics(to_period, replace(filters, period=to_period)),
        )

    def _group_rows(self, period: str, filters: GroupLendingFilters) -> tuple[GroupLendingRow, ...]:
        if not period:
            return ()
        directory = self._group_directory()
        raw_rows = self._raw_group_rows(period, filters)
        enriched = [self._enrich_group_row(period, row, directory) for row in raw_rows]
        return tuple(row for row in enriched if _passes_group_filters(row, filters))

    def _raw_group_rows(self, period: str, filters: GroupLendingFilters) -> list[sqlite3.Row]:
        where, params = self._loan_where_clause(replace(filters, period=period), require_group_code=True)
        where_sql = " AND ".join(where)
        with closing(self.credit_repository.connect()) as connection:
            return connection.execute(
                f"""
                WITH customer_group AS (
                    SELECT
                        l.period,
                        l.group_code,
                        l.customer_id,
                        MIN(l.branch_code) AS branch_code,
                        COUNT(DISTINCT l.loan_key) AS loan_count,
                        SUM(l.outstanding_balance) AS customer_total_balance
                    FROM credit_loan_period l
                    JOIN credit_customer_master c ON c.id = l.customer_id
                    WHERE {where_sql}
                    GROUP BY l.period, l.group_code, l.customer_id
                    HAVING SUM(l.outstanding_balance) > 0
                )
                SELECT
                    period,
                    group_code,
                    MIN(branch_code) AS branch_code,
                    COUNT(*) AS member_count,
                    SUM(loan_count) AS loan_count,
                    SUM(customer_total_balance) AS total_balance
                FROM customer_group
                GROUP BY period, group_code
                HAVING SUM(customer_total_balance) > 0
                """,
                params,
            ).fetchall()

    def _loan_where_clause(
        self,
        filters: GroupLendingFilters,
        *,
        require_group_code: bool,
    ) -> tuple[list[str], list[object]]:
        where = ["l.period = ?"]
        params: list[object] = [filters.period]
        if require_group_code:
            where.append("l.group_code IS NOT NULL AND TRIM(l.group_code) <> ''")
        if filters.branch_code:
            where.append("l.branch_code = ?")
            params.append(filters.branch_code)
        if filters.officer:
            where.append(
                """
                (
                    l.officer_code = ?
                    OR l.officer_name = ?
                    OR ('[' || l.officer_code || '] ' || l.officer_name) = ?
                )
                """
            )
            params.extend([filters.officer, filters.officer, filters.officer])
        return where, params

    def _group_directory(self) -> dict[str, GroupDirectoryEntry]:
        self.group_directory_load_count += 1
        groups = self.group_repository.list_groups(include_inactive=True)
        return {group.ma_to: _directory_entry_from_group(group) for group in groups}

    def _enrich_group_row(
        self,
        period: str,
        row: sqlite3.Row,
        directory: dict[str, GroupDirectoryEntry],
    ) -> GroupLendingRow:
        group_code = str(row["group_code"] or "").strip()
        entry = directory.get(group_code) or _unknown_entry(group_code)
        member_count = int(row["member_count"] or 0)
        total_balance = float(row["total_balance"] or 0)
        return GroupLendingRow(
            period=period,
            group_code=group_code,
            group_name=entry.group_name,
            association_type=entry.association_type,
            association_label=entry.association_label,
            association_other_name=entry.association_other_name,
            branch_code=str(row["branch_code"] or ""),
            office_name=entry.office_name,
            commune=entry.commune,
            leader_name=entry.leader_name,
            member_count=member_count,
            loan_count=int(row["loan_count"] or 0),
            total_balance=total_balance,
            average_balance_per_member=(total_balance / member_count if member_count else None),
            status=entry.status,
            status_label=entry.status_label,
        )

    def _association_rows(
        self,
        rows: tuple[GroupLendingRow, ...],
        filters: GroupLendingFilters,
    ) -> list[GroupAssociationSummaryRow]:
        total_balance = sum(row.total_balance for row in rows)
        grouped: dict[str, dict[str, object]] = {
            key: {"rows": [], "unique_members": set()}
            for key in _association_order()
        }
        for row in rows:
            grouped.setdefault(row.association_type, {"rows": [], "unique_members": set()})
            grouped[row.association_type]["rows"].append(row)
        output: list[GroupAssociationSummaryRow] = []
        for key in _association_order():
            group_rows = list(grouped.get(key, {}).get("rows", []))
            balance = sum(row.total_balance for row in group_rows)
            member_occurrence = sum(row.member_count for row in group_rows)
            unique_members = self._unique_member_count_for_association(key, tuple(group_rows), filters)
            output.append(
                GroupAssociationSummaryRow(
                    association_type=key,
                    association_label=_association_label(key),
                    group_count=sum(1 for row in group_rows if row.total_balance > 0),
                    unique_member_count=unique_members,
                    member_occurrence_count=member_occurrence,
                    total_balance=balance,
                    share=_ratio(balance, total_balance),
                    average_balance_per_group=(balance / len(group_rows) if group_rows else None),
                    average_balance_per_member=(balance / unique_members if unique_members else None),
                )
            )
        return output

    def _unique_member_count_for_association(
        self,
        association_type: str,
        rows: tuple[GroupLendingRow, ...],
        filters: GroupLendingFilters,
    ) -> int:
        _ = association_type
        if not rows:
            return 0
        group_codes = [row.group_code for row in rows]
        placeholders = ", ".join("?" for _code in group_codes)
        where, params = self._loan_where_clause(replace(filters, period=rows[0].period), require_group_code=True)
        where.append(f"l.group_code IN ({placeholders})")
        params.extend(group_codes)
        with closing(self.credit_repository.connect()) as connection:
            row = connection.execute(
                f"""
                WITH member_group AS (
                    SELECT l.customer_id, l.group_code, SUM(l.outstanding_balance) AS balance
                    FROM credit_loan_period l
                    JOIN credit_customer_master c ON c.id = l.customer_id
                    WHERE {' AND '.join(where)}
                    GROUP BY l.customer_id, l.group_code
                    HAVING SUM(l.outstanding_balance) > 0
                )
                SELECT COUNT(DISTINCT customer_id)
                FROM member_group
                """,
                params,
            ).fetchone()
        return int(row[0] if row else 0)

    def _kpis(
        self,
        period: str,
        rows: tuple[GroupLendingRow, ...],
        filters: GroupLendingFilters,
    ) -> tuple[GroupLendingKpi, ...]:
        total_balance = sum(row.total_balance for row in rows)
        active_group_count = sum(1 for row in rows if row.total_balance > 0)
        unique_member_count = self._unique_member_count_for_group_codes(
            period,
            filters,
            tuple(row.group_code for row in rows),
        )
        member_occurrences = sum(row.member_count for row in rows)
        unknown_rows = [row for row in rows if row.status == GROUP_STATUS_NOT_DECLARED]
        return (
            GroupLendingKpi("Số tổ có dư nợ", active_group_count, "count"),
            GroupLendingKpi("Tổng số tổ viên duy nhất", unique_member_count, "count"),
            GroupLendingKpi("Tổng lượt tổ viên theo tổ", member_occurrences, "count"),
            GroupLendingKpi("Tổng dư nợ cho vay qua tổ", total_balance, "money"),
            GroupLendingKpi("Dư nợ bình quân/tổ", total_balance / active_group_count if active_group_count else None, "money"),
            GroupLendingKpi("Dư nợ bình quân/tổ viên", total_balance / unique_member_count if unique_member_count else None, "money"),
            GroupLendingKpi("Số mã tổ chưa khai báo", len(unknown_rows), "count"),
            GroupLendingKpi("Dư nợ thuộc mã tổ chưa khai báo", sum(row.total_balance for row in unknown_rows), "money"),
        )

    def _comparison_kpis(
        self,
        from_period: str,
        to_period: str,
        from_rows: tuple[GroupLendingRow | GroupAssociationSummaryRow, ...],
        to_rows: tuple[GroupLendingRow | GroupAssociationSummaryRow, ...],
    ) -> tuple[GroupLendingKpi, ...]:
        specs = (
            ("Số tổ có dư nợ", sum(1 for row in from_rows if row.total_balance > 0), sum(1 for row in to_rows if row.total_balance > 0), "count"),
            ("Số lượng tổ viên", _sum_members(from_rows), _sum_members(to_rows), "count"),
            ("Tổng dư nợ cho vay qua tổ", sum(row.total_balance for row in from_rows), sum(row.total_balance for row in to_rows), "money"),
            (
                "Dư nợ bình quân/tổ",
                _average_balance_per_group(from_rows),
                _average_balance_per_group(to_rows),
                "money",
            ),
        )
        output: list[GroupLendingKpi] = []
        for label, before, after, kind in specs:
            diff = float(after or 0) - float(before or 0)
            output.append(
                GroupLendingKpi(
                    label=label,
                    value=after,
                    kind=kind,
                    from_value=before,
                    to_value=after,
                    difference=diff,
                    growth_rate=None if float(before or 0) == 0 else diff / float(before or 0) * 100,
                    tooltip=f"Từ kỳ {from_period}: {before}\nĐến kỳ {to_period}: {after}",
                )
            )
        return tuple(output)

    def _unique_member_count_for_group_codes(
        self,
        period: str,
        filters: GroupLendingFilters,
        group_codes: tuple[str, ...],
    ) -> int:
        clean_codes = tuple(code for code in group_codes if str(code or "").strip())
        if not period or not clean_codes:
            return 0
        placeholders = ", ".join("?" for _code in clean_codes)
        where, params = self._loan_where_clause(replace(filters, period=period), require_group_code=True)
        where.append(f"l.group_code IN ({placeholders})")
        params.extend(clean_codes)
        with closing(self.credit_repository.connect()) as connection:
            rows = connection.execute(
                f"""
                WITH member_group AS (
                    SELECT l.customer_id, l.group_code, SUM(l.outstanding_balance) AS balance
                    FROM credit_loan_period l
                    JOIN credit_customer_master c ON c.id = l.customer_id
                    WHERE {' AND '.join(where)}
                    GROUP BY l.customer_id, l.group_code
                    HAVING SUM(l.outstanding_balance) > 0
                )
                SELECT COUNT(DISTINCT customer_id)
                FROM member_group
                """,
                params,
            ).fetchone()
        return int(rows[0] if rows else 0)

    def _diagnostics(self, period: str, filters: GroupLendingFilters) -> dict[str, object]:
        where, params = self._loan_where_clause(replace(filters, period=period), require_group_code=True)
        where_sql = " AND ".join(where)
        with closing(self.credit_repository.connect()) as connection:
            no_group = connection.execute(
                """
                SELECT COUNT(*) AS row_count, COALESCE(SUM(outstanding_balance), 0) AS balance
                FROM credit_loan_period l
                WHERE l.period = ? AND (l.group_code IS NULL OR TRIM(l.group_code) = '')
                """,
                (period,),
            ).fetchone()
            multi = connection.execute(
                f"""
                WITH customer_groups AS (
                    SELECT l.customer_id, COUNT(DISTINCT l.group_code) AS group_count, SUM(l.outstanding_balance) AS balance
                    FROM credit_loan_period l
                    JOIN credit_customer_master c ON c.id = l.customer_id
                    WHERE {where_sql}
                    GROUP BY l.customer_id
                    HAVING COUNT(DISTINCT l.group_code) >= 2 AND SUM(l.outstanding_balance) > 0
                )
                SELECT COUNT(*) AS customer_count, COALESCE(SUM(balance), 0) AS balance
                FROM customer_groups
                """,
                params,
            ).fetchone()
        return {
            "no_group_row_count": int(no_group["row_count"] or 0),
            "no_group_balance": float(no_group["balance"] or 0),
            "multi_group_customer_count": int(multi["customer_count"] or 0),
            "multi_group_customer_balance": float(multi["balance"] or 0),
            "credit_card_scope": "excluded_without_group_code",
        }

    def _period_notes(self, period: str, rows: tuple[GroupLendingRow, ...]) -> tuple[str, ...]:
        with closing(self.credit_repository.connect()) as connection:
            loan_rows = int(connection.execute("SELECT COUNT(*) FROM credit_loan_period WHERE period = ?", (period,)).fetchone()[0] or 0)
            group_rows = int(
                connection.execute(
                    """
                    SELECT COUNT(*)
                    FROM credit_loan_period
                    WHERE period = ? AND group_code IS NOT NULL AND TRIM(group_code) <> ''
                    """,
                    (period,),
                ).fetchone()[0]
                or 0
            )
        if loan_rows > 0 and group_rows == 0:
            return (f"Kỳ {period} chưa có dữ liệu mã tổ vay vốn. Hãy ghi đè kỳ bằng file LN01 có cột GRPNO.",)
        if not rows and loan_rows == 0:
            return (f"Kỳ {period} chưa có dữ liệu LN01 trong Credit.db.",)
        return ("Dữ liệu thẻ DN15 không được cộng vì nguồn thẻ chưa có mã tổ vay vốn.",)

    def _sort_group_rows(self, rows: tuple[GroupLendingRow, ...], *, sort_by: str, sort_desc: bool) -> tuple[GroupLendingRow, ...]:
        allowed = {
            "group_code": lambda row: row.group_code,
            "member_count": lambda row: row.member_count,
            "loan_count": lambda row: row.loan_count,
            "total_balance": lambda row: row.total_balance,
        }
        key = allowed.get(sort_by, allowed["total_balance"])
        return tuple(sorted(rows, key=key, reverse=sort_desc))


def _directory_entry_from_group(group: CreditGroup) -> GroupDirectoryEntry:
    association_type = group.association_type or ASSOCIATION_OTHER
    return GroupDirectoryEntry(
        group_code=group.ma_to,
        group_name=group.ten_to or group.ten_tvv_day_du or group.ma_to,
        association_type=association_type,
        association_label=_association_label(association_type),
        association_other_name=group.association_other_name,
        branch_name="",
        office_name="",
        commune=group.xa,
        leader_name=group.ten_to_truong,
        active=bool(group.active),
        status=GROUP_STATUS_DECLARED if group.active else GROUP_STATUS_INACTIVE,
        status_label=GROUP_STATUS_LABELS[GROUP_STATUS_DECLARED] if group.active else GROUP_STATUS_LABELS[GROUP_STATUS_INACTIVE],
    )


def _unknown_entry(group_code: str) -> GroupDirectoryEntry:
    return GroupDirectoryEntry(
        group_code=group_code,
        group_name="Chưa khai báo",
        association_type=ASSOCIATION_UNKNOWN,
        association_label=ASSOCIATION_UNKNOWN_LABEL,
        association_other_name="",
        branch_name="",
        office_name="",
        commune="",
        leader_name="",
        active=False,
        status=GROUP_STATUS_NOT_DECLARED,
        status_label=GROUP_STATUS_LABELS[GROUP_STATUS_NOT_DECLARED],
    )


def _passes_group_filters(row: GroupLendingRow, filters: GroupLendingFilters) -> bool:
    if row.status == GROUP_STATUS_NOT_DECLARED and not filters.include_unknown_groups:
        return False
    if filters.association_type and row.association_type != filters.association_type:
        return False
    if filters.group_status and row.status != filters.group_status:
        return False
    if filters.office_code and row.office_name != filters.office_code:
        return False
    needle = str(filters.search or "").strip().casefold()
    if needle:
        haystack = " ".join(
            (
                row.group_code,
                row.group_name,
                row.association_label,
                row.association_other_name,
                row.commune,
                row.leader_name,
            )
        ).casefold()
        if needle not in haystack:
            return False
    return True


def _association_order() -> tuple[str, ...]:
    return (
        ASSOCIATION_FARMERS_UNION,
        ASSOCIATION_WOMENS_UNION,
        ASSOCIATION_OTHER,
        ASSOCIATION_UNKNOWN,
    )


def _association_label(value: str) -> str:
    if value == ASSOCIATION_UNKNOWN:
        return ASSOCIATION_UNKNOWN_LABEL
    return ASSOCIATION_TYPE_LABELS.get(value, ASSOCIATION_UNKNOWN_LABEL)


def _ratio(value: float, total: float) -> float | None:
    return None if not total else float(value or 0) / float(total) * 100


def _page(rows: tuple[GroupLendingRow, ...], page: int, page_size: int) -> tuple[int, int, tuple[GroupLendingRow, ...]]:
    page = max(1, int(page or 1))
    page_size = max(1, int(page_size or 100))
    start = (page - 1) * page_size
    return page, page_size, rows[start : start + page_size]


def _member_order_clause(sort_by: str, sort_desc: bool) -> str:
    allowed = {
        "customer_code": "customer_code",
        "loan_count": "loan_count",
        "total_balance": "total_balance",
        "worst_debt_group": "worst_debt_group",
    }
    column = allowed.get(sort_by, "total_balance")
    direction = "DESC" if sort_desc else "ASC"
    return f"{column} {direction}, customer_code COLLATE NOCASE ASC"


def _member_row_to_dict(row: sqlite3.Row) -> dict[str, object]:
    worst = int(row["worst_debt_group"] or 0)
    return {
        "Mã khách hàng": row["customer_code"],
        "Tên khách hàng": row["customer_name"],
        "Loại khách hàng": row["customer_type_code"],
        "CBTD": _officer_display(str(row["officer_code"] or ""), str(row["officer_name"] or "")),
        "Số món": int(row["loan_count"] or 0),
        "Tổng dư nợ": float(row["total_balance"] or 0),
        "Nhóm nợ cao nhất": f"Nhóm {worst}" if worst else "Chưa xác định",
        "Ngắn hạn": float(row["short_term_balance"] or 0),
        "Trung hạn": float(row["medium_term_balance"] or 0),
        "Dài hạn": float(row["long_term_balance"] or 0),
    }


def _movement_category(before: float, after: float) -> str:
    if before <= 0 < after:
        return "Tổ mới có dư nợ"
    if before > 0 >= after:
        return "Tổ hết dư nợ"
    if after > before:
        return "Tăng dư nợ"
    if after < before:
        return "Giảm dư nợ"
    return "Không thay đổi"


def _compare_notes(from_period: str, to_period: str) -> tuple[str, ...]:
    notes = ["So sánh trực tiếp Từ kỳ với Đến kỳ; không cộng hoặc lấy trung bình kỳ giữa."]
    if from_period == to_period:
        notes.append("Từ kỳ và Đến kỳ đang giống nhau.")
    notes.append("Dữ liệu thẻ DN15 không được cộng vì nguồn thẻ chưa có mã tổ vay vốn.")
    return tuple(notes)


def _sum_members(rows: tuple[GroupLendingRow | GroupAssociationSummaryRow, ...]) -> int:
    total = 0
    for row in rows:
        if isinstance(row, GroupAssociationSummaryRow):
            total += row.unique_member_count
        else:
            total += row.member_count
    return total


def _average_balance_per_group(rows: tuple[GroupLendingRow | GroupAssociationSummaryRow, ...]) -> float | None:
    count = sum(1 for row in rows if row.total_balance > 0)
    if not count:
        return None
    return sum(row.total_balance for row in rows) / count
