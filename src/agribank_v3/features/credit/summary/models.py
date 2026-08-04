from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime
from enum import StrEnum
from pathlib import Path


NIM_DN_TITLE = "NIM dư nợ"
NIM_NV_TITLE = "NIM nguồn vốn"
NIM_TITLE = "NIM DN / NIM NV"
LOAN_COMPARE_TITLE = "So sánh tăng giảm khách hàng"
CREDIT_LIMIT_TITLE = "Hạn mức tín dụng hết hạn"
DEBT_GROUP_NORMAL = "NORMAL"
DEBT_GROUP_ATTENTION = "ATTENTION"
DEBT_GROUP_BAD_DEBT = "BAD_DEBT"
DEBT_GROUP_UNKNOWN = "UNKNOWN"
DEBT_GROUP_VALID_CODES = ("01", "02", "03", "04", "05")
DEBT_GROUP_COLUMN_SUFFIXES = ("1", "2", "3", "4", "5", "unknown")


class SummaryError(RuntimeError):
    pass


class SummaryDataType(StrEnum):
    NIM_DN = "nim_dn"
    NIM_NV = "nim_nv"
    LOAN_COMPARE = "loan_compare"
    CREDIT_LIMIT = "credit_limit"


@dataclass(frozen=True, slots=True)
class NimConfig:
    data_type: SummaryDataType
    title: str
    file_pattern_token: str
    officer_header: str
    has_average_rate: bool


NIM_DN_CONFIG = NimConfig(
    data_type=SummaryDataType.NIM_DN,
    title="NIM Doanh nghiệp",
    file_pattern_token="_ftpln_",
    officer_header="CBTD",
    has_average_rate=True,
)

NIM_NV_CONFIG = NimConfig(
    data_type=SummaryDataType.NIM_NV,
    title="NIM Cá nhân",
    file_pattern_token="_ftpdp_",
    officer_header="CBHD",
    has_average_rate=False,
)


@dataclass(frozen=True, slots=True)
class ImportBatch:
    id: int
    data_type: str
    period: str
    source_path: str
    file_name: str
    imported_by: str
    row_count: int
    status: str
    message: str
    created_at: str
    updated_at: str
    version: int = 1


@dataclass(frozen=True, slots=True)
class NimRow:
    batch_id: int
    data_type: SummaryDataType
    period: str
    branch_code: str
    branch_name: str
    trctcd: str
    transaction_office: str
    customer_type: str
    officer: str
    balance: float
    interest_rate: float
    ftp_rate: float
    adjustment_rate: float
    numerator_before: float
    numerator_after: float
    average_rate_numerator: float
    source_file: str
    source_row_count: int = 1


@dataclass(frozen=True, slots=True)
class NormalizedLoanRow:
    period: str
    source_file: str
    source_row_number: int
    branch_code: str
    trctcd: str
    transaction_office: str
    customer_sequence: str
    customer_code: str
    customer_name: str
    customer_type: str
    ftp_code: str
    balance: float
    ftp: float
    interest_rate: float
    ftp_adjustment: float
    officer_code: str
    officer_name: str
    office_code: str = ""
    office_name: str = ""
    office_type: str = "UNKNOWN"
    debt_group_code: str = DEBT_GROUP_UNKNOWN
    debt_group_number: int | None = None
    debt_group_category: str = DEBT_GROUP_UNKNOWN
    has_valid_debt_group: bool = False


@dataclass(frozen=True, slots=True)
class NimSummaryRow:
    period: str
    branch: str
    transaction_office: str
    customer_type: str
    officer: str
    balance: float
    average_rate: float
    nim_before: float
    nim_after: float


@dataclass(frozen=True, slots=True)
class OfficerHistoryPoint:
    period: str
    balance: float
    average_rate: float
    nim_before: float
    nim_after: float


@dataclass(frozen=True, slots=True)
class OfficerHistory:
    data_type: SummaryDataType
    officer: str
    branch: str
    transaction_office: str
    customer_type: str
    current_period: str
    current_balance: float
    current_average_rate: float
    current_nim_before: float
    current_nim_after: float
    points: tuple[OfficerHistoryPoint, ...]


@dataclass(frozen=True, slots=True)
class LoanSnapshotRow:
    customer_code: str
    customer_name: str
    address: str
    officer: str
    previous_balance: float
    current_balance: float
    category: str


@dataclass(frozen=True, slots=True)
class CreditLimitRow:
    customer_code: str
    customer_name: str
    contract_number: str
    approved_date: date | None
    approved_amount: float
    outstanding_balance: float
    expiry_date: date | None
    address: str
    officer: str
    note: str
    days_to_expiry: int | None
    status: str
    officer_code: str = ""
    branch_code: str = ""
    account_number: str = ""
    credit_line_type: str = "Line of Credit"
    source_row_count: int = 1


@dataclass(frozen=True, slots=True)
class ImportResult:
    batch_id: int | str
    row_count: int
    message: str
    output_path: Path | None = None


@dataclass(frozen=True, slots=True)
class PageResult:
    rows: list[dict[str, object]]
    total_rows: int
    page: int
    page_size: int


@dataclass(frozen=True, slots=True)
class DashboardMetric:
    label: str
    value: str
    detail: str = ""


@dataclass(frozen=True, slots=True)
class DashboardData:
    metrics: tuple[DashboardMetric, ...]
    bars: tuple[tuple[object, ...], ...] = ()
    lines: tuple[tuple[object, ...], ...] = ()
    pies: tuple[tuple[object, ...], ...] = ()


def now_text() -> str:
    return datetime.now().astimezone().isoformat(timespec="seconds")
