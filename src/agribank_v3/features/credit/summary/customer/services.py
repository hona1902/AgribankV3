from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
import re

from agribank_v3.features.credit.summary.models import (
    DEBT_GROUP_ATTENTION,
    DEBT_GROUP_BAD_DEBT,
    DEBT_GROUP_NORMAL,
    DEBT_GROUP_UNKNOWN,
    NormalizedLoanRow,
)
from agribank_v3.features.credit.summary.customer.filters import (
    MOVEMENT_STATUS_DECREASE,
    MOVEMENT_STATUS_INCREASE,
    MOVEMENT_STATUS_NEW,
    MOVEMENT_STATUS_PAID_OFF,
    MOVEMENT_STATUS_UNCHANGED,
)
from agribank_v3.features.credit.summary.customer.models import (
    CustomerAggregationResult,
    CustomerImportFile,
    CustomerOfficeAggregate,
    CustomerOfficeType,
    CustomerOfficerAggregate,
    CustomerPeriodAggregate,
    CustomerTypeCode,
    LoanTerm,
    OfficerIdentity,
    RepresentativeOffice,
    RepresentativeOfficeReason,
)


SHORT_TERM_FTP_CODES = frozenset(
    {
        "DN1",
        "DN2",
        "DN3",
        "DN4",
        "DN5",
        "DN6",
        "DN13",
        "DN14",
        "DN15",
        "DN16",
    }
)
MEDIUM_LONG_TERM_FTP_CODES = frozenset({"DN7", "DN8", "DN9", "DN10", "DN11", "DN12"})
OFFICE_UNKNOWN_SUFFIX = "UNKNOWN"
DEBT_GROUP_SUFFIX_BY_CODE = {
    "01": "1",
    "02": "2",
    "03": "3",
    "04": "4",
    "05": "5",
    DEBT_GROUP_UNKNOWN: "unknown",
}
DEBT_GROUP_CATEGORY_BY_CODE = {
    "01": DEBT_GROUP_NORMAL,
    "02": DEBT_GROUP_ATTENTION,
    "03": DEBT_GROUP_BAD_DEBT,
    "04": DEBT_GROUP_BAD_DEBT,
    "05": DEBT_GROUP_BAD_DEBT,
}


def normalize_customer_sequence(value: object) -> str:
    text = "" if value is None else str(value).strip()
    if text.startswith("'"):
        text = text[1:].strip()
    if re.fullmatch(r"\d+\.0+", text):
        text = text.split(".", 1)[0]
    return text


def build_customer_code(branch_code: object, customer_sequence: object) -> str:
    branch = "" if branch_code is None else str(branch_code).strip()
    sequence = normalize_customer_sequence(customer_sequence)
    return f"{branch}{sequence}"


def normalize_trctcd(value: object) -> str:
    text = "" if value is None else str(value).strip().replace("'", "")
    if re.fullmatch(r"\d+\.0+", text):
        text = text.split(".", 1)[0]
    if text.isdigit() and len(text) < 2:
        text = text.zfill(2)
    return text


def build_office_code(branch_code: object, trctcd: object) -> str:
    branch = "" if branch_code is None else str(branch_code).strip()
    code = normalize_trctcd(trctcd)
    if not branch:
        return ""
    if not code:
        return f"{branch}-{OFFICE_UNKNOWN_SUFFIX}"
    return f"{branch}-{code}"


def classify_office_type(trctcd: object) -> CustomerOfficeType:
    code = normalize_trctcd(trctcd)
    if not code:
        return CustomerOfficeType.UNKNOWN
    if code == "00":
        return CustomerOfficeType.HEAD_OFFICE
    return CustomerOfficeType.TRANSACTION_OFFICE


def map_customer_type_code(value: object) -> CustomerTypeCode:
    text = "" if value is None else str(value).strip().upper()
    if text == CustomerTypeCode.PERSONAL.value:
        return CustomerTypeCode.PERSONAL
    if text == CustomerTypeCode.ORGANIZATION.value:
        return CustomerTypeCode.ORGANIZATION
    return CustomerTypeCode.OTHER


def normalize_debt_group(value: object) -> tuple[str, int | None, str, bool]:
    text = "" if value is None else str(value).strip()
    if text.startswith("'"):
        text = text[1:].strip()
    if not text or text.casefold() in {"null", "none", "n/a", "na"}:
        return DEBT_GROUP_UNKNOWN, None, DEBT_GROUP_UNKNOWN, False
    text = text.replace(",", ".")
    if re.fullmatch(r"\d+\.0+", text):
        text = text.split(".", 1)[0]
    if not re.fullmatch(r"\d+", text):
        return DEBT_GROUP_UNKNOWN, None, DEBT_GROUP_UNKNOWN, False
    try:
        number = int(text)
    except ValueError:
        return DEBT_GROUP_UNKNOWN, None, DEBT_GROUP_UNKNOWN, False
    if number < 1 or number > 5:
        return DEBT_GROUP_UNKNOWN, None, DEBT_GROUP_UNKNOWN, False
    code = f"{number:02d}"
    return code, number, DEBT_GROUP_CATEGORY_BY_CODE[code], True


def customer_type_label(value: object) -> str:
    code = map_customer_type_code(value)
    if code == CustomerTypeCode.PERSONAL:
        return "Ca nhan"
    if code == CustomerTypeCode.ORGANIZATION:
        return "To chuc/Phap nhan"
    return "Khac"


def classify_loan_term(ftp_code: object) -> LoanTerm:
    code = "" if ftp_code is None else str(ftp_code).strip().upper()
    if code in SHORT_TERM_FTP_CODES:
        return LoanTerm.SHORT_TERM
    if code in MEDIUM_LONG_TERM_FTP_CODES:
        return LoanTerm.MEDIUM_LONG_TERM
    return LoanTerm.UNKNOWN


def split_officer(value: object) -> OfficerIdentity:
    text = "" if value is None else str(value).strip()
    match = re.match(r"^\s*\[([^\]]+)\]\s*(.*)$", text)
    if match:
        return OfficerIdentity(match.group(1).strip(), match.group(2).strip())
    return OfficerIdentity("", text)


@dataclass(slots=True)
class _DebtGroupAccumulator:
    balance: float = 0.0
    interest_rate_numerator: float = 0.0
    nim_before_numerator: float = 0.0
    nim_after_numerator: float = 0.0
    source_row_count: int = 0


@dataclass(slots=True)
class _OfficerAccumulator:
    officer_code: str
    officer_name: str
    balance_managed: float = 0.0
    short_term_balance: float = 0.0
    medium_long_term_balance: float = 0.0
    other_balance: float = 0.0
    source_loan_count: int = 0
    interest_rate_numerator: float = 0.0
    nim_before_numerator: float = 0.0
    nim_after_numerator: float = 0.0
    first_seen_order: int = 0
    branch_code: str = ""
    transaction_office: str = ""
    has_debt_group_data: bool = False
    debt_group_unknown_row_count: int = 0
    debt_groups: dict[str, _DebtGroupAccumulator] | None = None

    def __post_init__(self) -> None:
        if self.debt_groups is None:
            self.debt_groups = _empty_debt_group_accumulators()


@dataclass(slots=True)
class _OfficeOfficerAccumulator:
    officer_code: str
    officer_name: str
    balance_managed: float = 0.0
    first_seen_order: int = 0


@dataclass(slots=True)
class _OfficeAccumulator:
    period: str
    customer_code: str
    customer_sequence: str
    branch_code: str
    trctcd: str
    office_code: str
    office_name: str
    office_type: str
    first_seen_order: int
    total_balance: float = 0.0
    short_term_balance: float = 0.0
    medium_long_term_balance: float = 0.0
    other_balance: float = 0.0
    interest_rate_numerator: float = 0.0
    nim_before_numerator: float = 0.0
    nim_after_numerator: float = 0.0
    source_loan_count: int = 0
    officers: dict[str, _OfficeOfficerAccumulator] | None = None
    has_debt_group_data: bool = False
    debt_group_unknown_row_count: int = 0
    debt_groups: dict[str, _DebtGroupAccumulator] | None = None

    def __post_init__(self) -> None:
        if self.officers is None:
            self.officers = {}
        if self.debt_groups is None:
            self.debt_groups = _empty_debt_group_accumulators()


@dataclass(slots=True)
class _CustomerAccumulator:
    period: str
    customer_code: str
    branch_code: str
    customer_sequence: str
    customer_name: str
    customer_type: str
    first_seen_order: int
    total_balance: float = 0.0
    short_term_balance: float = 0.0
    medium_long_term_balance: float = 0.0
    other_balance: float = 0.0
    interest_rate_numerator: float = 0.0
    nim_before_numerator: float = 0.0
    nim_after_numerator: float = 0.0
    source_loan_count: int = 0
    officer_keys: set[str] | None = None
    has_debt_group_data: bool = False
    debt_group_unknown_row_count: int = 0
    debt_groups: dict[str, _DebtGroupAccumulator] | None = None

    def __post_init__(self) -> None:
        if self.officer_keys is None:
            self.officer_keys = set()
        if self.debt_groups is None:
            self.debt_groups = _empty_debt_group_accumulators()


def _empty_debt_group_accumulators() -> dict[str, _DebtGroupAccumulator]:
    return {suffix: _DebtGroupAccumulator() for suffix in ("1", "2", "3", "4", "5", "unknown")}


def _debt_group_suffix(code: object) -> str:
    return DEBT_GROUP_SUFFIX_BY_CODE.get(str(code or "").strip().upper(), "unknown")


def _add_debt_group_amounts(
    target,
    row: NormalizedLoanRow,
    balance: float,
    interest_rate_numerator: float,
    nim_before_numerator: float,
    nim_after_numerator: float,
) -> None:
    suffix = _debt_group_suffix(row.debt_group_code if row.has_valid_debt_group else DEBT_GROUP_UNKNOWN)
    groups = target.debt_groups or _empty_debt_group_accumulators()
    target.debt_groups = groups
    item = groups[suffix]
    item.balance += balance
    item.interest_rate_numerator += interest_rate_numerator
    item.nim_before_numerator += nim_before_numerator
    item.nim_after_numerator += nim_after_numerator
    item.source_row_count += 1
    target.has_debt_group_data = True
    if not row.has_valid_debt_group:
        target.debt_group_unknown_row_count += 1


def _debt_group_kwargs(target) -> dict[str, object]:
    groups = target.debt_groups or _empty_debt_group_accumulators()
    output: dict[str, object] = {
        "has_debt_group_data": bool(target.has_debt_group_data),
        "worst_debt_group": _worst_debt_group(groups, float(getattr(target, "total_balance", 0) or getattr(target, "balance_managed", 0) or 0)),
        "debt_group_unknown_row_count": int(target.debt_group_unknown_row_count or 0),
    }
    for suffix in ("1", "2", "3", "4", "5", "unknown"):
        group = groups[suffix]
        output[f"debt_group_{suffix}_balance"] = float(group.balance or 0)
        output[f"debt_group_{suffix}_interest_numerator"] = float(group.interest_rate_numerator or 0)
        output[f"debt_group_{suffix}_nim_before_numerator"] = float(group.nim_before_numerator or 0)
        output[f"debt_group_{suffix}_nim_after_numerator"] = float(group.nim_after_numerator or 0)
    return output


def _worst_debt_group(groups: dict[str, _DebtGroupAccumulator], total_balance: float) -> str:
    if total_balance <= 0:
        return ""
    for suffix in ("5", "4", "3", "2", "1"):
        if float(groups.get(suffix, _DebtGroupAccumulator()).balance or 0) > 0:
            return f"{int(suffix):02d}"
    if float(groups.get("unknown", _DebtGroupAccumulator()).balance or 0) > 0:
        return DEBT_GROUP_UNKNOWN
    return ""


class CustomerAggregationService:
    """Aggregate FTPLN loan rows into customer-period rows only."""

    def __init__(self, source_folder: Path | str = "") -> None:
        self.source_folder = str(source_folder or "")
        self._customers: dict[tuple[str, str], _CustomerAccumulator] = {}
        self._officers: dict[tuple[str, str, str], _OfficerAccumulator] = {}
        self._offices: dict[tuple[str, str, str], _OfficeAccumulator] = {}
        self._files: list[CustomerImportFile] = []
        self._unknown_ftp_codes: set[str] = set()
        self._warnings: list[str] = []
        self._warning_count = 0
        self._invalid_row_count = 0
        self._source_row_count = 0
        self._source_total_balance = 0.0
        self._order = 0
        self._debt_group_valid_row_count = 0
        self._debt_group_counts = {"01": 0, "02": 0, "03": 0, "04": 0, "05": 0, DEBT_GROUP_UNKNOWN: 0}
        self._debt_group_invalid_samples: list[str] = []

    def add_file(
        self,
        rows: list[NormalizedLoanRow],
        *,
        period: str,
        file_path: Path,
        file_hash: str,
        source_row_count: int,
        source_total_balance: float,
        invalid_row_count: int = 0,
        warning_count: int = 0,
        warnings: tuple[str, ...] = (),
        debt_group_invalid_samples: tuple[str, ...] = (),
        debt_group_header_present: bool = False,
    ) -> None:
        file_path = Path(file_path)
        file_customer_codes: set[str] = set()
        file_branch_code = ""
        self._source_row_count += int(source_row_count)
        self._source_total_balance += float(source_total_balance or 0)
        self._invalid_row_count += int(invalid_row_count)
        self._warning_count += max(0, int(warning_count) - len(warnings))
        for warning in warnings:
            self._add_warning(warning)
        for sample in debt_group_invalid_samples:
            text = str(sample or "").strip()
            if text and text not in self._debt_group_invalid_samples and len(self._debt_group_invalid_samples) < 20:
                self._debt_group_invalid_samples.append(text)
        for row in rows:
            if not row.customer_code:
                self._invalid_row_count += 1
                self._add_warning(
                    f"{row.source_file}:{row.source_row_number} khong co customer_code hop le."
                )
                continue
            if not file_branch_code and row.branch_code:
                file_branch_code = row.branch_code
            file_customer_codes.add(row.customer_code)
            self._add_row(row, debt_group_header_present=debt_group_header_present)
        self._files.append(
            CustomerImportFile(
                file_name=file_path.name,
                file_path=str(file_path),
                file_hash=str(file_hash or ""),
                branch_code=file_branch_code,
                period=period,
                source_row_count=int(source_row_count),
                customer_count=len(file_customer_codes),
            )
        )

    def build_result(self) -> CustomerAggregationResult:
        officer_rows_by_customer: dict[tuple[str, str], list[_OfficerAccumulator]] = {}
        for (period, customer_code, _identity), officer in self._officers.items():
            officer_rows_by_customer.setdefault((period, customer_code), []).append(officer)

        summaries: list[CustomerPeriodAggregate] = []
        officer_rows: list[CustomerOfficerAggregate] = []
        office_rows: list[CustomerOfficeAggregate] = []
        personal_count = 0
        organization_count = 0
        multiple_officer_count = 0
        total_balance = 0.0
        short_term_balance = 0.0
        medium_long_term_balance = 0.0
        other_balance = 0.0

        for key, customer in sorted(
            self._customers.items(),
            key=lambda item: (item[0][0], item[1].first_seen_order),
        ):
            customer_officers = officer_rows_by_customer.get(key, [])
            primary = _primary_officer(customer_officers)
            officer_count = len(customer_officers)
            if officer_count > 1:
                multiple_officer_count += 1
            if customer.customer_type == CustomerTypeCode.PERSONAL.value:
                personal_count += 1
            elif customer.customer_type == CustomerTypeCode.ORGANIZATION.value:
                organization_count += 1
            balance = float(customer.total_balance or 0)
            total_balance += balance
            short_term_balance += float(customer.short_term_balance or 0)
            medium_long_term_balance += float(customer.medium_long_term_balance or 0)
            other_balance += float(customer.other_balance or 0)
            summaries.append(
                CustomerPeriodAggregate(
                    period=customer.period,
                    customer_code=customer.customer_code,
                    branch_code=customer.branch_code,
                    customer_sequence=customer.customer_sequence,
                    customer_name=customer.customer_name,
                    customer_type=customer.customer_type,
                    primary_officer_code=primary.officer_code if primary else "",
                    primary_officer_name=primary.officer_name if primary else "",
                    officer_count=officer_count,
                    has_multiple_officers=officer_count > 1,
                    total_balance=balance,
                    short_term_balance=float(customer.short_term_balance or 0),
                    medium_long_term_balance=float(customer.medium_long_term_balance or 0),
                    other_balance=float(customer.other_balance or 0),
                    medium_long_ratio=(float(customer.medium_long_term_balance or 0) / balance * 100) if balance else 0.0,
                    interest_rate_numerator=float(customer.interest_rate_numerator or 0),
                    nim_before_numerator=float(customer.nim_before_numerator or 0),
                    nim_after_numerator=float(customer.nim_after_numerator or 0),
                    average_rate=(float(customer.interest_rate_numerator or 0) / balance) if balance else 0.0,
                    nim_before=(float(customer.nim_before_numerator or 0) / balance) if balance else 0.0,
                    nim_after=(float(customer.nim_after_numerator or 0) / balance) if balance else 0.0,
                    source_loan_count=int(customer.source_loan_count or 0),
                    **_debt_group_kwargs(customer),
                )
            )
            for officer in sorted(customer_officers, key=lambda item: item.first_seen_order):
                officer_rows.append(
                    CustomerOfficerAggregate(
                        period=customer.period,
                        customer_code=customer.customer_code,
                        officer_code=officer.officer_code,
                        officer_name=officer.officer_name,
                        balance_managed=float(officer.balance_managed or 0),
                        short_term_balance=float(officer.short_term_balance or 0),
                        medium_long_term_balance=float(officer.medium_long_term_balance or 0),
                        other_balance=float(officer.other_balance or 0),
                        source_loan_count=int(officer.source_loan_count or 0),
                        interest_rate_numerator=float(officer.interest_rate_numerator or 0),
                        nim_before_numerator=float(officer.nim_before_numerator or 0),
                        nim_after_numerator=float(officer.nim_after_numerator or 0),
                        is_primary=officer is primary,
                        first_seen_order=officer.first_seen_order,
                        branch_code=officer.branch_code,
                        transaction_office=officer.transaction_office,
                        **_debt_group_kwargs(officer),
                    )
                )

        for (_period, _customer_code, _office_code), office in sorted(
            self._offices.items(),
            key=lambda item: (item[0][0], item[1].first_seen_order),
        ):
            office_officers = list((office.officers or {}).values())
            primary = _primary_office_officer(office_officers)
            office_rows.append(
                CustomerOfficeAggregate(
                    period=office.period,
                    customer_code=office.customer_code,
                    customer_sequence=office.customer_sequence,
                    branch_code=office.branch_code,
                    trctcd=office.trctcd,
                    office_code=office.office_code,
                    office_name=office.office_name,
                    office_type=office.office_type,
                    primary_officer_code=primary.officer_code if primary else "",
                    primary_officer_name=primary.officer_name if primary else "",
                    officer_count=len(office_officers),
                    total_balance=float(office.total_balance or 0),
                    short_term_balance=float(office.short_term_balance or 0),
                    medium_long_term_balance=float(office.medium_long_term_balance or 0),
                    other_balance=float(office.other_balance or 0),
                    interest_rate_numerator=float(office.interest_rate_numerator or 0),
                    nim_before_numerator=float(office.nim_before_numerator or 0),
                    nim_after_numerator=float(office.nim_after_numerator or 0),
                    source_loan_count=int(office.source_loan_count or 0),
                    first_seen_order=office.first_seen_order,
                    **_debt_group_kwargs(office),
                )
            )

        periods = sorted({file.period for file in self._files if file.period})
        if len(periods) == 1:
            period = periods[0]
        elif periods:
            period = "Nhiều kỳ"
        else:
            period = ""
        return CustomerAggregationResult(
            period=period,
            source_folder=self.source_folder,
            file_count=len(self._files),
            source_row_count=self._source_row_count,
            customer_count=len(summaries),
            personal_customer_count=personal_count,
            organization_customer_count=organization_count,
            total_balance=total_balance,
            source_total_balance=self._source_total_balance,
            short_term_balance=short_term_balance,
            medium_long_term_balance=medium_long_term_balance,
            other_balance=other_balance,
            multiple_officer_customer_count=multiple_officer_count,
            unknown_ftp_codes=tuple(sorted(self._unknown_ftp_codes)),
            invalid_row_count=self._invalid_row_count,
            warning_count=self._warning_count,
            warnings=tuple(self._warnings),
            files=tuple(self._files),
            summaries=tuple(summaries),
            officer_rows=tuple(officer_rows),
            office_rows=tuple(office_rows),
            debt_group_valid_row_count=self._debt_group_valid_row_count,
            debt_group_1_row_count=self._debt_group_counts["01"],
            debt_group_2_row_count=self._debt_group_counts["02"],
            debt_group_3_row_count=self._debt_group_counts["03"],
            debt_group_4_row_count=self._debt_group_counts["04"],
            debt_group_5_row_count=self._debt_group_counts["05"],
            debt_group_unknown_row_count=self._debt_group_counts[DEBT_GROUP_UNKNOWN],
            debt_group_invalid_samples=tuple(self._debt_group_invalid_samples),
        )

    def _add_row(self, row: NormalizedLoanRow, *, debt_group_header_present: bool = False) -> None:
        self._order += 1
        key = (row.period, row.customer_code)
        customer_type = map_customer_type_code(row.customer_type).value
        customer = self._customers.get(key)
        if customer is None:
            customer = _CustomerAccumulator(
                period=row.period,
                customer_code=row.customer_code,
                branch_code=row.branch_code,
                customer_sequence=row.customer_sequence,
                customer_name=str(row.customer_name or "").strip(),
                customer_type=customer_type,
                first_seen_order=self._order,
            )
            self._customers[key] = customer
        else:
            self._merge_customer_identity(customer, row, customer_type)

        balance = float(row.balance or 0)
        customer.total_balance += balance
        term = classify_loan_term(row.ftp_code)
        if term == LoanTerm.SHORT_TERM:
            customer.short_term_balance += balance
        elif term == LoanTerm.MEDIUM_LONG_TERM:
            customer.medium_long_term_balance += balance
        else:
            customer.other_balance += balance
            if row.ftp_code:
                self._unknown_ftp_codes.add(row.ftp_code)
        interest_rate_numerator = float(row.interest_rate or 0) * balance
        nim_before_numerator = (float(row.interest_rate or 0) - float(row.ftp or 0)) * balance
        nim_after_numerator = (
            float(row.interest_rate or 0)
            - float(row.ftp or 0)
            - float(row.ftp_adjustment or 0)
        ) * balance
        customer.interest_rate_numerator += interest_rate_numerator
        customer.nim_before_numerator += nim_before_numerator
        customer.nim_after_numerator += nim_after_numerator
        customer.source_loan_count += 1
        if debt_group_header_present:
            if row.has_valid_debt_group:
                self._debt_group_valid_row_count += 1
                self._debt_group_counts[row.debt_group_code] = self._debt_group_counts.get(row.debt_group_code, 0) + 1
            else:
                self._debt_group_counts[DEBT_GROUP_UNKNOWN] = self._debt_group_counts.get(DEBT_GROUP_UNKNOWN, 0) + 1
            _add_debt_group_amounts(
                customer,
                row,
                balance,
                interest_rate_numerator,
                nim_before_numerator,
                nim_after_numerator,
            )
        self._add_office_row(
            row,
            balance,
            term,
            interest_rate_numerator,
            nim_before_numerator,
            nim_after_numerator,
            debt_group_header_present=debt_group_header_present,
        )

        officer_code = str(row.officer_code or "").strip()
        officer_name = str(row.officer_name or "").strip()
        identity = officer_code or f"name:{officer_name.casefold()}"
        officer_key = (row.period, row.customer_code, identity)
        if customer.officer_keys is not None:
            customer.officer_keys.add(identity)
        officer = self._officers.get(officer_key)
        if officer is None:
            officer = _OfficerAccumulator(
                officer_code=officer_code,
                officer_name=officer_name,
                first_seen_order=self._order,
                branch_code=row.branch_code,
                transaction_office=row.transaction_office,
            )
            self._officers[officer_key] = officer
        elif not officer.officer_name and officer_name:
            officer.officer_name = officer_name
        officer.balance_managed += balance
        if term == LoanTerm.SHORT_TERM:
            officer.short_term_balance += balance
        elif term == LoanTerm.MEDIUM_LONG_TERM:
            officer.medium_long_term_balance += balance
        else:
            officer.other_balance += balance
        officer.source_loan_count += 1
        officer.interest_rate_numerator += interest_rate_numerator
        officer.nim_before_numerator += nim_before_numerator
        officer.nim_after_numerator += nim_after_numerator
        if debt_group_header_present:
            _add_debt_group_amounts(
                officer,
                row,
                balance,
                interest_rate_numerator,
                nim_before_numerator,
                nim_after_numerator,
            )

    def _add_office_row(
        self,
        row: NormalizedLoanRow,
        balance: float,
        term: LoanTerm,
        interest_rate_numerator: float,
        nim_before_numerator: float,
        nim_after_numerator: float,
        *,
        debt_group_header_present: bool = False,
    ) -> None:
        trctcd = normalize_trctcd(row.trctcd)
        office_code = str(row.office_code or "").strip() or build_office_code(row.branch_code, trctcd)
        office_type = str(row.office_type or "").strip() or classify_office_type(trctcd).value
        office_name = str(row.office_name or row.transaction_office or "").strip()
        if not office_name and office_type == CustomerOfficeType.UNKNOWN.value:
            office_name = "Không xác định"
        key = (row.period, row.customer_code, office_code)
        office = self._offices.get(key)
        if office is None:
            office = _OfficeAccumulator(
                period=row.period,
                customer_code=row.customer_code,
                customer_sequence=row.customer_sequence,
                branch_code=row.branch_code,
                trctcd=trctcd,
                office_code=office_code,
                office_name=office_name,
                office_type=office_type,
                first_seen_order=self._order,
            )
            self._offices[key] = office
        elif not office.office_name and office_name:
            office.office_name = office_name
        office.total_balance += balance
        if term == LoanTerm.SHORT_TERM:
            office.short_term_balance += balance
        elif term == LoanTerm.MEDIUM_LONG_TERM:
            office.medium_long_term_balance += balance
        else:
            office.other_balance += balance
        office.interest_rate_numerator += interest_rate_numerator
        office.nim_before_numerator += nim_before_numerator
        office.nim_after_numerator += nim_after_numerator
        office.source_loan_count += 1
        if debt_group_header_present:
            _add_debt_group_amounts(
                office,
                row,
                balance,
                interest_rate_numerator,
                nim_before_numerator,
                nim_after_numerator,
            )

        officer_code = str(row.officer_code or "").strip()
        officer_name = str(row.officer_name or "").strip()
        identity = officer_code or f"name:{officer_name.casefold()}"
        if not identity:
            return
        office_officers = office.officers or {}
        office_officer = office_officers.get(identity)
        if office_officer is None:
            office_officer = _OfficeOfficerAccumulator(
                officer_code=officer_code,
                officer_name=officer_name,
                first_seen_order=self._order,
            )
            office_officers[identity] = office_officer
        elif not office_officer.officer_name and officer_name:
            office_officer.officer_name = officer_name
        office_officer.balance_managed += balance

    def _merge_customer_identity(
        self,
        customer: _CustomerAccumulator,
        row: NormalizedLoanRow,
        customer_type: str,
    ) -> None:
        incoming_name = str(row.customer_name or "").strip()
        if not customer.customer_name and incoming_name:
            customer.customer_name = incoming_name
        elif incoming_name and incoming_name.casefold() != customer.customer_name.casefold():
            self._add_warning(
                f"{row.source_file}:{row.source_row_number} ten khach hang khac nhau cho {row.customer_code}; giu gia tri dau tien."
            )
        if customer.customer_type == CustomerTypeCode.OTHER.value and customer_type != CustomerTypeCode.OTHER.value:
            customer.customer_type = customer_type
        elif (
            customer_type != CustomerTypeCode.OTHER.value
            and customer.customer_type != customer_type
        ):
            self._add_warning(
                f"{row.source_file}:{row.source_row_number} loai khach hang xung dot cho {row.customer_code}; giu {customer.customer_type}."
            )

    def _add_warning(self, message: str) -> None:
        self._warning_count += 1
        if len(self._warnings) < 50:
            self._warnings.append(message)


def validate_customer_balance(result: CustomerAggregationResult, *, tolerance: float = 0.0001) -> None:
    difference = float(result.total_balance or 0) - float(result.source_total_balance or 0)
    if abs(difference) > tolerance:
        raise ValueError(
            "Tong du no Customer khong khop nguon FTPLN: "
            f"customer={result.total_balance}, source={result.source_total_balance}, diff={difference}."
        )


def _primary_officer(officers: list[_OfficerAccumulator]) -> _OfficerAccumulator | None:
    if not officers:
        return None
    return sorted(officers, key=lambda item: (-float(item.balance_managed or 0), item.first_seen_order))[0]


def _primary_office_officer(officers: list[_OfficeOfficerAccumulator]) -> _OfficeOfficerAccumulator | None:
    if not officers:
        return None
    return sorted(officers, key=lambda item: (-float(item.balance_managed or 0), item.first_seen_order))[0]


def resolve_representative_office(
    period: object,
    customer_sequence: object,
    branch_code: object,
    office_rows: list[dict[str, object]] | tuple[dict[str, object], ...],
) -> RepresentativeOffice:
    _ = (period, customer_sequence)
    branch = "" if branch_code is None else str(branch_code).strip()
    relevant = [
        dict(row)
        for row in office_rows
        if str(row.get("branch_code") or "").strip() == branch
        and float(row.get("total_balance") or 0) > 0
    ]
    office_count = len({str(row.get("office_code") or "").strip() for row in relevant if str(row.get("office_code") or "").strip()})
    head_offices = [
        row for row in relevant if str(row.get("office_type") or "").strip() == CustomerOfficeType.HEAD_OFFICE.value
    ]
    pgd_rows = [
        row for row in relevant if str(row.get("office_type") or "").strip() == CustomerOfficeType.TRANSACTION_OFFICE.value
    ]
    pgd_count = len({str(row.get("office_code") or "").strip() for row in pgd_rows if str(row.get("office_code") or "").strip()})
    if head_offices:
        selected = sorted(head_offices, key=lambda row: str(row.get("trctcd") or "00"))[0]
        return RepresentativeOffice(
            representative_office_code=str(selected.get("office_code") or build_office_code(branch, "00")),
            representative_office_name=str(selected.get("office_name") or "Hội sở"),
            representative_office_type=CustomerOfficeType.HEAD_OFFICE.value,
            has_head_office=True,
            pgd_count=pgd_count,
            office_count=office_count,
            reason=RepresentativeOfficeReason.HAS_HEAD_OFFICE.value,
        )
    if pgd_count == 1 and pgd_rows:
        selected = pgd_rows[0]
        return RepresentativeOffice(
            representative_office_code=str(selected.get("office_code") or ""),
            representative_office_name=str(selected.get("office_name") or ""),
            representative_office_type=CustomerOfficeType.TRANSACTION_OFFICE.value,
            has_head_office=False,
            pgd_count=pgd_count,
            office_count=office_count,
            reason=RepresentativeOfficeReason.SINGLE_PGD.value,
        )
    if pgd_count > 1:
        selected = sorted(
            pgd_rows,
            key=lambda row: (-float(row.get("total_balance") or 0), normalize_trctcd(row.get("trctcd"))),
        )[0]
        return RepresentativeOffice(
            representative_office_code=str(selected.get("office_code") or ""),
            representative_office_name=str(selected.get("office_name") or ""),
            representative_office_type=CustomerOfficeType.TRANSACTION_OFFICE.value,
            has_head_office=False,
            pgd_count=pgd_count,
            office_count=office_count,
            reason=RepresentativeOfficeReason.MULTIPLE_PGD_LARGEST_BALANCE.value,
        )
    selected_unknown = relevant[0] if relevant else {}
    return RepresentativeOffice(
        representative_office_code=str(selected_unknown.get("office_code") or ""),
        representative_office_name=str(selected_unknown.get("office_name") or ""),
        representative_office_type=str(selected_unknown.get("office_type") or CustomerOfficeType.UNKNOWN.value),
        has_head_office=False,
        pgd_count=0,
        office_count=office_count,
        reason=RepresentativeOfficeReason.UNKNOWN.value,
    )


@dataclass(frozen=True, slots=True)
class EffectiveOfficer:
    officer_code: str
    officer_name: str
    has_override: bool = False
    reason: str = ""
    updated_at: str = ""


@dataclass(frozen=True, slots=True)
class CustomerMovementKpis:
    new_customer_count: int = 0
    paid_off_customer_count: int = 0
    increased_customer_count: int = 0
    decreased_customer_count: int = 0
    total_increase: float = 0.0
    total_decrease: float = 0.0
    net_difference: float = 0.0


def weighted_ratio(numerator: object, denominator: object) -> float:
    total = float(denominator or 0)
    if total == 0:
        return 0.0
    return float(numerator or 0) / total


def classify_customer_movement(previous_balance: object, current_balance: object) -> str:
    previous = float(previous_balance or 0)
    current = float(current_balance or 0)
    if previous <= 0 and current > 0:
        return MOVEMENT_STATUS_NEW
    if previous > 0 and current <= 0:
        return MOVEMENT_STATUS_PAID_OFF
    if current > previous:
        return MOVEMENT_STATUS_INCREASE
    if current < previous:
        return MOVEMENT_STATUS_DECREASE
    return MOVEMENT_STATUS_UNCHANGED


def balance_difference(previous_balance: object, current_balance: object) -> float:
    return float(current_balance or 0) - float(previous_balance or 0)


def growth_rate(previous_balance: object, current_balance: object) -> float | None:
    previous = float(previous_balance or 0)
    if previous == 0:
        return None
    return balance_difference(previous_balance, current_balance) / previous * 100


def resolve_effective_officer(
    imported_officer_code: object,
    imported_officer_name: object,
    override: dict[str, object] | None = None,
) -> EffectiveOfficer:
    if override:
        code = str(override.get("officer_code") or "").strip()
        name = str(override.get("officer_name") or "").strip()
        if code or name:
            return EffectiveOfficer(
                officer_code=code,
                officer_name=name,
                has_override=True,
                reason=str(override.get("reason") or "").strip(),
                updated_at=str(override.get("updated_at") or "").strip(),
            )
    return EffectiveOfficer(
        officer_code=str(imported_officer_code or "").strip(),
        officer_name=str(imported_officer_name or "").strip(),
        has_override=False,
    )
