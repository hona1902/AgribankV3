from __future__ import annotations

from collections import Counter
from pathlib import Path

from agribank_v3.features.credit.tovayvon.models import (
    ASSOCIATION_TYPES,
    association_type_display,
    normalize_association_type,
)
from agribank_v3.features.credit.tovayvon.repository import CreditGroupRepository


class CreditGroupOrganizationService:
    """Read-only facade for report modules that need credit-group organization type."""

    def __init__(self, database_path: Path | str | CreditGroupRepository) -> None:
        self.repository = (
            database_path
            if isinstance(database_path, CreditGroupRepository)
            else CreditGroupRepository(Path(database_path))
        )

    def get_group_organization_type(self, group_code: str) -> str:
        group = self.repository.get_group(str(group_code or "").strip())
        return group.association_type if group is not None else ""

    def get_group_organization_display(self, group_code: str) -> str:
        group = self.repository.get_group(str(group_code or "").strip())
        if group is None:
            return ""
        return association_type_display(group.association_type, group.association_other_name)

    def list_groups_by_organization_type(self, type_code: str) -> list[str]:
        normalized = normalize_association_type(type_code)
        if not normalized:
            return []
        return [
            group.ma_to
            for group in self.repository.list_groups(association_type=normalized)
        ]

    def summarize_group_counts_by_organization_type(self) -> dict[str, int]:
        counter: Counter[str] = Counter()
        for group in self.repository.list_groups():
            type_code = normalize_association_type(group.association_type)
            if type_code:
                counter[type_code] += 1
        return {type_code: int(counter.get(type_code, 0)) for type_code in ASSOCIATION_TYPES}
