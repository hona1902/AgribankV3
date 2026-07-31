from __future__ import annotations

from dataclasses import dataclass


HEAD_OFFICE = "HEAD_OFFICE"
TRANSACTION_OFFICE = "TRANSACTION_OFFICE"
OTHER_OFFICE = "OTHER"
OFFICE_TYPES = (HEAD_OFFICE, TRANSACTION_OFFICE, OTHER_OFFICE)


@dataclass(frozen=True, slots=True)
class BranchDirectoryEntry:
    branch_code: str
    branch_name: str
    short_name: str = ""
    display_name: str = ""
    province_name: str = ""
    is_active: bool = True
    sort_order: int | None = None
    created_at: str = ""
    updated_at: str = ""
    updated_by: str = ""


@dataclass(frozen=True, slots=True)
class OfficeDirectoryEntry:
    id: int | None
    branch_code: str
    trctcd: str
    office_code: str
    office_name: str
    short_name: str = ""
    office_type: str = OTHER_OFFICE
    is_active: bool = True
    sort_order: int | None = None
    created_at: str = ""
    updated_at: str = ""
    updated_by: str = ""


@dataclass(frozen=True, slots=True)
class AppUnitSettings:
    home_branch_code: str = ""
    default_office_code: str = ""
    organization_name: str = ""
    updated_at: str = ""
    updated_by: str = ""

