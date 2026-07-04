from __future__ import annotations

from dataclasses import dataclass, field
from enum import StrEnum
from pathlib import Path
from typing import Any

from agribank_v3.settings import BranchProfile


class SettlementCategory(StrEnum):
    CREDIT = "credit"
    ACCOUNTING = "accounting"
    CONSOLIDATION = "consolidation"
    LEGACY = "legacy"


class SettlementSourceMode(StrEnum):
    ACTIVE_WORKBOOK = "active_workbook"
    MULTIPLE_FILES = "multiple_files"
    GENERATED_WORKBOOK = "generated_workbook"


@dataclass(frozen=True, slots=True)
class SettlementOptions:
    convert_tcvn3_to_unicode: bool = True
    include_branch_in_customer_id: bool = False
    use_collateral_owner_for_guarantee: bool = True
    four_digit_year: bool = True
    create_control_sheet: bool = True
    remove_unused_columns: bool = True
    include_customer_totals: bool = False
    bold_customer_rows: bool = False
    remove_customer_total_rows: bool = True
    include_accrual_accounts: bool = True
    use_default_accrual_accounts: bool = True
    include_loan_deposit_schedule: bool = False
    source_report_code: str = ""


@dataclass(frozen=True, slots=True)
class SettlementSpec:
    key: str
    report_code: str
    title: str
    category: SettlementCategory
    source_hint: str
    source_mode: SettlementSourceMode
    processor_family: str
    legacy_entrypoint: str
    active: bool = True


@dataclass(frozen=True, slots=True)
class SettlementRequest:
    spec: SettlementSpec
    profile: BranchProfile
    options: SettlementOptions = field(default_factory=SettlementOptions)
    source_paths: tuple[Path, ...] = ()
    excel_application: Any | None = None


@dataclass(frozen=True, slots=True)
class SettlementResult:
    spec_key: str
    output_path: Path | None
    workbook_name: str
    worksheet_name: str
    processed_rows: int
    warnings: tuple[str, ...] = ()
