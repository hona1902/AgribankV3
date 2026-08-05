from __future__ import annotations

from pathlib import Path

from agribank_v3.features.credit.summary.models import PageResult

from .models import GroupLendingFilters, GroupLendingResult
from .repository import GroupLendingRepository


class GroupLendingService:
    def __init__(self, main_database_path: Path | None = None) -> None:
        self.repository = GroupLendingRepository(main_database_path)

    def filter_values(self) -> dict[str, list[tuple[str, str]]]:
        return self.repository.filter_values()

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
        return self.repository.get_group_lending_snapshot(
            period,
            filters,
            page=page,
            page_size=page_size,
            sort_by=sort_by,
            sort_desc=sort_desc,
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
        return self.repository.get_group_detail_page(
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
        return self.repository.get_group_member_page(
            period,
            group_code,
            filters,
            page=page,
            page_size=page_size,
            sort_by=sort_by,
            sort_desc=sort_desc,
        )

    def get_association_summary(self, period: str, filters: GroupLendingFilters) -> GroupLendingResult:
        return self.repository.get_association_summary(period, filters)

    def compare_groups(self, from_period: str, to_period: str, filters: GroupLendingFilters) -> GroupLendingResult:
        return self.repository.compare_groups(from_period, to_period, filters)

    def compare_associations(self, from_period: str, to_period: str, filters: GroupLendingFilters) -> GroupLendingResult:
        return self.repository.compare_associations(from_period, to_period, filters)
