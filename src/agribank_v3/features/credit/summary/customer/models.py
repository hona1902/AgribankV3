from __future__ import annotations

from dataclasses import dataclass
from enum import StrEnum

class CustomerDatabaseError(RuntimeError):
    pass


class CustomerDataType(StrEnum):
    NIM_DN = "DN"


class CustomerTypeCode(StrEnum):
    PERSONAL = "CN"
    ORGANIZATION = "TC"
    OTHER = "OTHER"


class LoanTerm(StrEnum):
    SHORT_TERM = "SHORT_TERM"
    MEDIUM_LONG_TERM = "MEDIUM_LONG_TERM"
    UNKNOWN = "UNKNOWN"


class CustomerOfficeType(StrEnum):
    HEAD_OFFICE = "HEAD_OFFICE"
    TRANSACTION_OFFICE = "TRANSACTION_OFFICE"
    UNKNOWN = "UNKNOWN"


class RepresentativeOfficeReason(StrEnum):
    HAS_HEAD_OFFICE = "HAS_HEAD_OFFICE"
    SINGLE_PGD = "SINGLE_PGD"
    MULTIPLE_PGD_LARGEST_BALANCE = "MULTIPLE_PGD_LARGEST_BALANCE"
    UNKNOWN = "UNKNOWN"


@dataclass(frozen=True, slots=True)
class DebtGroupAggregate:
    balance: float = 0.0
    interest_rate_numerator: float = 0.0
    nim_before_numerator: float = 0.0
    nim_after_numerator: float = 0.0
    source_row_count: int = 0


@dataclass(frozen=True, slots=True)
class OfficerIdentity:
    officer_code: str
    officer_name: str


@dataclass(frozen=True, slots=True)
class CustomerImportFile:
    file_name: str
    file_path: str
    file_hash: str
    branch_code: str
    period: str
    source_row_count: int
    customer_count: int
    status: str = "COMPLETED"
    error_message: str = ""


@dataclass(frozen=True, slots=True)
class CustomerPeriodAggregate:
    period: str
    customer_code: str
    branch_code: str
    customer_sequence: str
    customer_name: str
    customer_type: str
    primary_officer_code: str
    primary_officer_name: str
    officer_count: int
    has_multiple_officers: bool
    total_balance: float
    short_term_balance: float
    medium_long_term_balance: float
    other_balance: float
    medium_long_ratio: float
    interest_rate_numerator: float
    nim_before_numerator: float
    nim_after_numerator: float
    average_rate: float
    nim_before: float
    nim_after: float
    source_loan_count: int
    has_debt_group_data: bool = False
    worst_debt_group: str = ""
    debt_group_unknown_row_count: int = 0
    debt_group_1_balance: float = 0.0
    debt_group_2_balance: float = 0.0
    debt_group_3_balance: float = 0.0
    debt_group_4_balance: float = 0.0
    debt_group_5_balance: float = 0.0
    debt_group_unknown_balance: float = 0.0
    debt_group_1_interest_numerator: float = 0.0
    debt_group_2_interest_numerator: float = 0.0
    debt_group_3_interest_numerator: float = 0.0
    debt_group_4_interest_numerator: float = 0.0
    debt_group_5_interest_numerator: float = 0.0
    debt_group_unknown_interest_numerator: float = 0.0
    debt_group_1_nim_before_numerator: float = 0.0
    debt_group_2_nim_before_numerator: float = 0.0
    debt_group_3_nim_before_numerator: float = 0.0
    debt_group_4_nim_before_numerator: float = 0.0
    debt_group_5_nim_before_numerator: float = 0.0
    debt_group_unknown_nim_before_numerator: float = 0.0
    debt_group_1_nim_after_numerator: float = 0.0
    debt_group_2_nim_after_numerator: float = 0.0
    debt_group_3_nim_after_numerator: float = 0.0
    debt_group_4_nim_after_numerator: float = 0.0
    debt_group_5_nim_after_numerator: float = 0.0
    debt_group_unknown_nim_after_numerator: float = 0.0


@dataclass(frozen=True, slots=True)
class CustomerOfficerAggregate:
    period: str
    customer_code: str
    officer_code: str
    officer_name: str
    balance_managed: float
    short_term_balance: float
    medium_long_term_balance: float
    other_balance: float
    source_loan_count: int
    interest_rate_numerator: float
    nim_before_numerator: float
    nim_after_numerator: float
    is_primary: bool
    first_seen_order: int
    branch_code: str = ""
    transaction_office: str = ""
    has_debt_group_data: bool = False
    worst_debt_group: str = ""
    debt_group_unknown_row_count: int = 0
    debt_group_1_balance: float = 0.0
    debt_group_2_balance: float = 0.0
    debt_group_3_balance: float = 0.0
    debt_group_4_balance: float = 0.0
    debt_group_5_balance: float = 0.0
    debt_group_unknown_balance: float = 0.0
    debt_group_1_interest_numerator: float = 0.0
    debt_group_2_interest_numerator: float = 0.0
    debt_group_3_interest_numerator: float = 0.0
    debt_group_4_interest_numerator: float = 0.0
    debt_group_5_interest_numerator: float = 0.0
    debt_group_unknown_interest_numerator: float = 0.0
    debt_group_1_nim_before_numerator: float = 0.0
    debt_group_2_nim_before_numerator: float = 0.0
    debt_group_3_nim_before_numerator: float = 0.0
    debt_group_4_nim_before_numerator: float = 0.0
    debt_group_5_nim_before_numerator: float = 0.0
    debt_group_unknown_nim_before_numerator: float = 0.0
    debt_group_1_nim_after_numerator: float = 0.0
    debt_group_2_nim_after_numerator: float = 0.0
    debt_group_3_nim_after_numerator: float = 0.0
    debt_group_4_nim_after_numerator: float = 0.0
    debt_group_5_nim_after_numerator: float = 0.0
    debt_group_unknown_nim_after_numerator: float = 0.0


@dataclass(frozen=True, slots=True)
class CustomerOfficeAggregate:
    period: str
    customer_code: str
    customer_sequence: str
    branch_code: str
    trctcd: str
    office_code: str
    office_name: str
    office_type: str
    primary_officer_code: str
    primary_officer_name: str
    officer_count: int
    total_balance: float
    short_term_balance: float
    medium_long_term_balance: float
    other_balance: float
    interest_rate_numerator: float
    nim_before_numerator: float
    nim_after_numerator: float
    source_loan_count: int
    first_seen_order: int
    has_debt_group_data: bool = False
    worst_debt_group: str = ""
    debt_group_unknown_row_count: int = 0
    debt_group_1_balance: float = 0.0
    debt_group_2_balance: float = 0.0
    debt_group_3_balance: float = 0.0
    debt_group_4_balance: float = 0.0
    debt_group_5_balance: float = 0.0
    debt_group_unknown_balance: float = 0.0
    debt_group_1_interest_numerator: float = 0.0
    debt_group_2_interest_numerator: float = 0.0
    debt_group_3_interest_numerator: float = 0.0
    debt_group_4_interest_numerator: float = 0.0
    debt_group_5_interest_numerator: float = 0.0
    debt_group_unknown_interest_numerator: float = 0.0
    debt_group_1_nim_before_numerator: float = 0.0
    debt_group_2_nim_before_numerator: float = 0.0
    debt_group_3_nim_before_numerator: float = 0.0
    debt_group_4_nim_before_numerator: float = 0.0
    debt_group_5_nim_before_numerator: float = 0.0
    debt_group_unknown_nim_before_numerator: float = 0.0
    debt_group_1_nim_after_numerator: float = 0.0
    debt_group_2_nim_after_numerator: float = 0.0
    debt_group_3_nim_after_numerator: float = 0.0
    debt_group_4_nim_after_numerator: float = 0.0
    debt_group_5_nim_after_numerator: float = 0.0
    debt_group_unknown_nim_after_numerator: float = 0.0


@dataclass(frozen=True, slots=True)
class RepresentativeOffice:
    representative_office_code: str
    representative_office_name: str
    representative_office_type: str
    has_head_office: bool
    pgd_count: int
    office_count: int
    reason: str


@dataclass(frozen=True, slots=True)
class CustomerAggregationResult:
    period: str
    source_folder: str
    file_count: int
    source_row_count: int
    customer_count: int
    personal_customer_count: int
    organization_customer_count: int
    total_balance: float
    source_total_balance: float
    short_term_balance: float
    medium_long_term_balance: float
    other_balance: float
    multiple_officer_customer_count: int
    unknown_ftp_codes: tuple[str, ...]
    invalid_row_count: int
    warning_count: int
    warnings: tuple[str, ...]
    files: tuple[CustomerImportFile, ...]
    summaries: tuple[CustomerPeriodAggregate, ...]
    officer_rows: tuple[CustomerOfficerAggregate, ...]
    office_rows: tuple[CustomerOfficeAggregate, ...] = ()
    debt_group_valid_row_count: int = 0
    debt_group_1_row_count: int = 0
    debt_group_2_row_count: int = 0
    debt_group_3_row_count: int = 0
    debt_group_4_row_count: int = 0
    debt_group_5_row_count: int = 0
    debt_group_unknown_row_count: int = 0
    debt_group_invalid_samples: tuple[str, ...] = ()


@dataclass(frozen=True, slots=True)
class CustomerDatabaseStatus:
    database_path: str
    size_bytes: int
    master_count: int
    period_count: int
    period_summary_count: int
    officer_period_count: int
    import_run_count: int
    import_file_count: int
    override_count: int
    action_log_count: int
    officer_directory_count: int
    first_period: str
    last_period: str
    page_count: int
    page_size: int
    freelist_count: int
    reclaimable_bytes: int
    last_optimized_at: str = ""
