from __future__ import annotations

import csv
from dataclasses import replace
from datetime import date, datetime
from io import StringIO
import logging
import re
from pathlib import Path
from typing import Iterable

from agribank_v3.features.credit.summary.customer.services import (
    build_customer_code,
    normalize_customer_sequence,
    normalize_debt_group,
)
from agribank_v3.features.credit.summary.models import (
    CreditLimitRow,
    NormalizedLn01Row,
    SummaryError,
)


LOGGER = logging.getLogger(__name__)

LN01_MIN_CREDIT_LIMIT_COLUMNS = 63
LN01_FIELD_ALIASES: dict[str, tuple[str, ...]] = {
    "branch_code": ("BRCD",),
    "customer_sequence": ("CUSTSEQ", "CUSTSED", "CUSTEQ"),
    "customer_name": ("CUSTNM", "CUSTOMER_NAME", "TEN_KHACH_HANG"),
    "account_number": ("TAI_KHOAN", "ACCOUNT_NUMBER", "ACCTNO"),
    "outstanding_balance": ("DU_NO", "OUTSTANDING_BALANCE", "LDRBAL"),
    "approval_sequence": ("APPRSEQ", "APPROVAL_SEQUENCE", "CONTRACT_NUMBER"),
    "approval_date": ("APPRDT", "APPROVED_DATE", "APPROVAL_DATE"),
    "approved_limit": ("APPRAMT", "APPROVED_AMOUNT", "APPROVED_LIMIT"),
    "maturity_date": ("APPRMATDT", "EXPIRY_DATE", "MATURITY_DATE"),
    "officer": ("OFFICER_NAME", "OFFICER", "CBTD"),
    "address": ("ADDR1", "ADDRESS"),
    "credit_line_type": ("CREDIT_LINE_YPE", "CREDIT_LINE_TYPE"),
    "customer_type_code": ("CUSTOMER_TYPE_CODE", "CUSTOMER_TYPE", "CUSTTP"),
    "secured_percent": ("SECURED_PERCENT",),
    "debt_group_code": ("NHOM_NO",),
    "industry_code": ("MA_NGANH_KT", "INDUSTRY_CODE"),
    "group_code": ("GRPNO", "MATOVAYVON", "MATO", "MÃ TỔ VAY VỐN", "MÃ SỐ TỔ"),
}
LN01_FALLBACK_INDEXES: dict[str, int] = {
    "branch_code": 0,
    "customer_sequence": 1,
    "customer_name": 2,
    "account_number": 3,
    "outstanding_balance": 5,
    "approval_sequence": 14,
    "approval_date": 15,
    "approved_limit": 17,
    "maturity_date": 18,
    "officer": 27,
    "address": 35,
    "credit_line_type": 62,
    "customer_type_code": 31,  # AF
    "secured_percent": 43,  # AR
    "debt_group_code": 44,  # AS
    "group_code": 49,  # AX
    "industry_code": 76,  # BY
}


def parse_ln01_bytes(
    file_path: Path,
    raw_data: bytes,
    *,
    period: str | None = None,
) -> tuple[list[NormalizedLn01Row], int]:
    try:
        text = raw_data.decode("utf-8-sig")
    except UnicodeDecodeError as exc:
        raise SummaryError("File LN01 phải dùng mã UTF-8 hoặc UTF-8 BOM.") from exc
    rows = _read_delimited_rows_from_text(text)
    if not rows:
        raise SummaryError("File LN01 không có dữ liệu.")
    headers = [_clean_header(item) for item in rows[0]]
    if len(headers) < LN01_MIN_CREDIT_LIMIT_COLUMNS:
        raise SummaryError("File LN01 không đủ số cột tới CREDIT_LINE_YPE.")
    _validate_ln01_headers(file_path, headers)
    column_map = _build_column_map(headers)
    report_period = _normalize_period(period) or parse_period_from_filename(file_path.name)
    normalized_rows: list[NormalizedLn01Row] = []
    for source_row_number, raw in enumerate(rows[1:], start=2):
        if not raw or not any(str(item).strip() for item in raw):
            continue
        branch_code = _normalize_excel_text(_value(raw, column_map, "branch_code"))
        customer_sequence = normalize_customer_sequence(_value(raw, column_map, "customer_sequence"))
        customer_code = build_customer_code(branch_code, customer_sequence) if branch_code and customer_sequence else ""
        officer_code, officer_name = split_officer(_value(raw, column_map, "officer"))
        debt_group_code, _number, _category, has_valid_debt_group = normalize_debt_group(
            _value(raw, column_map, "debt_group_code")
        )
        normalized_rows.append(
            NormalizedLn01Row(
                period=report_period,
                source_file=file_path.name,
                source_row_number=source_row_number,
                branch_code=branch_code,
                customer_sequence=customer_sequence,
                customer_code=customer_code,
                customer_name=_clean_cell(_value(raw, column_map, "customer_name")),
                account_number=_normalize_excel_text(_value(raw, column_map, "account_number")),
                approval_sequence=_normalize_excel_text(_value(raw, column_map, "approval_sequence")),
                approval_date=parse_date(_value(raw, column_map, "approval_date")),
                approved_limit=safe_amount(_value(raw, column_map, "approved_limit")),
                maturity_date=parse_date(_value(raw, column_map, "maturity_date")),
                outstanding_balance=safe_amount(_value(raw, column_map, "outstanding_balance")),
                customer_type_code=_normalize_excel_text(_value(raw, column_map, "customer_type_code")),
                debt_group_code=debt_group_code if has_valid_debt_group else "UNKNOWN",
                secured_percent=safe_decimal_or_none(_value(raw, column_map, "secured_percent")),
                industry_code=_normalize_excel_text(_value(raw, column_map, "industry_code")),
                officer_code=officer_code,
                officer_name=officer_name,
                address=_clean_cell(_value(raw, column_map, "address")),
                credit_line_type=_clean_cell(_value(raw, column_map, "credit_line_type")),
                group_code=normalize_credit_group_code(_value(raw, column_map, "group_code")),
            )
        )
    return normalized_rows, max(0, len(rows) - 1)


def ln01_has_group_code_header(raw_data: bytes) -> bool:
    try:
        text = raw_data.decode("utf-8-sig")
    except UnicodeDecodeError:
        return False
    rows = _read_delimited_rows_from_text(text)
    if not rows:
        return False
    headers = {_clean_header(item) for item in rows[0]}
    return any(alias in headers for alias in LN01_FIELD_ALIASES["group_code"])


def project_credit_limit_rows(rows: Iterable[NormalizedLn01Row]) -> list[CreditLimitRow]:
    by_contract: dict[tuple[str, str, str], CreditLimitRow] = {}
    for row in rows:
        if row.credit_line_type.upper() != "LINE OF CREDIT":
            continue
        contract_number = row.approval_sequence
        if not contract_number:
            continue
        key = (row.branch_code, row.customer_code, contract_number)
        if key in by_contract:
            existing = by_contract[key]
            by_contract[key] = replace(
                existing,
                outstanding_balance=existing.outstanding_balance + row.outstanding_balance,
                source_row_count=existing.source_row_count + row.source_row_count,
            )
            continue
        by_contract[key] = CreditLimitRow(
            branch_code=row.branch_code,
            customer_code=row.customer_sequence or row.customer_code,
            customer_name=row.customer_name,
            account_number=row.account_number,
            credit_line_type=row.credit_line_type,
            contract_number=contract_number,
            approved_date=row.approval_date,
            approved_amount=row.approved_limit,
            outstanding_balance=row.outstanding_balance,
            expiry_date=row.maturity_date,
            address=row.address,
            officer=row.officer_name,
            officer_code=row.officer_code,
            note="",
            days_to_expiry=None,
            status="",
            source_row_count=row.source_row_count,
            group_code=row.group_code,
        )
    return list(by_contract.values())


def parse_period_from_filename(file_name: str) -> str:
    stem = Path(file_name).stem
    date_part = stem[-8:]
    if len(date_part) == 8 and date_part.isdigit():
        return f"{date_part[:4]}-{date_part[4:6]}"
    return "Không Rõ"


def safe_amount(value: object) -> float:
    text = _clean_cell(value)
    if not text:
        return 0.0
    text = text.replace(" ", "")
    if "," in text and "." in text:
        text = text.replace(",", "")
    elif "," in text:
        parts = text.split(",")
        if len(parts) > 1 and all(len(part) == 3 for part in parts[1:]):
            text = "".join(parts)
        else:
            text = text.replace(",", ".")
    try:
        return float(text)
    except ValueError:
        return 0.0


def safe_decimal_or_none(value: object) -> float | None:
    text = _clean_cell(value)
    if not text:
        return None
    text = text.replace(" ", "").replace(",", ".")
    try:
        return float(text)
    except ValueError:
        return None


def parse_date(value: object) -> date | None:
    text = _clean_cell(value)
    if not text:
        return None
    if " " in text:
        text = text.split(" ", 1)[0]
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%Y%m%d"):
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            continue
    try:
        return datetime.fromisoformat(text).date()
    except ValueError:
        return None


def split_officer(raw_name: object) -> tuple[str, str]:
    text = str(raw_name or "").strip()
    match = re.match(r"^\s*\[([^\]]+)\]\s*(.*)$", text)
    if match:
        return match.group(1).strip(), match.group(2).strip()
    return "", text


def normalize_credit_group_code(value: object) -> str | None:
    text = _clean_cell(value)
    if not text:
        return None
    if text.startswith("'"):
        text = text[1:].strip()
    if not text:
        return None
    if re.fullmatch(r"\d+\.0+", text):
        text = text.split(".", 1)[0]
    return text or None


def _read_delimited_rows_from_text(text: str) -> list[list[str]]:
    first_line = text.splitlines()[0] if text.splitlines() else ""
    delimiter = ","
    if "\t" in first_line:
        delimiter = "\t"
    elif ";" in first_line:
        delimiter = ";"
    return [row for row in csv.reader(StringIO(text), delimiter=delimiter)]


def _validate_ln01_headers(file_path: Path, headers: list[str]) -> None:
    h0 = headers[0] if len(headers) > 0 else ""
    h1 = headers[1] if len(headers) > 1 else ""
    h2 = headers[2] if len(headers) > 2 else ""
    h3 = headers[3] if len(headers) > 3 else ""
    hbk = headers[62] if len(headers) > 62 else ""
    if (
        h0 != "BRCD"
        or h1 not in {"CUSTSED", "CUSTSEQ", "CUSTEQ"}
        or h2 != "CUSTNM"
        or h3 != "TAI_KHOAN"
        or hbk not in {"CREDIT_LINE_YPE", "CREDIT_LINE_TYPE"}
    ):
        raise SummaryError(f"Bạn chọn không đúng file LN01 xuất từ mssr98: {file_path.name}.")


def _build_column_map(headers: list[str]) -> dict[str, int | None]:
    header_map: dict[str, int] = {}
    for index, header in enumerate(headers):
        header_map.setdefault(header, index)
    column_map: dict[str, int | None] = {}
    for field, aliases in LN01_FIELD_ALIASES.items():
        found = next((header_map[alias] for alias in aliases if alias in header_map), None)
        if found is not None:
            column_map[field] = found
            continue
        fallback = LN01_FALLBACK_INDEXES.get(field)
        if fallback is not None and fallback < len(headers):
            column_map[field] = fallback
            LOGGER.info("LN01 thiếu header %s; dùng fallback cột %s.", aliases[0], fallback + 1)
        else:
            column_map[field] = None
            LOGGER.info("LN01 thiếu header %s và không có cột fallback.", aliases[0])
    return column_map


def _value(row: list[str], column_map: dict[str, int | None], field: str) -> str:
    index = column_map.get(field)
    if index is None or index < 0 or index >= len(row):
        return ""
    return _clean_cell(row[index])


def _clean_header(value: object) -> str:
    text = str(value or "").replace("\ufeff", "").strip().replace('"', "")
    upper = text.upper()
    if upper.endswith("CUSTSEQ"):
        return "CUSTSEQ"
    return upper


def _clean_cell(value: object) -> str:
    if value is None:
        return ""
    return str(value).strip().replace('"', "")


def _normalize_excel_text(value: object) -> str:
    text = _clean_cell(value)
    if text.startswith("'"):
        text = text[1:].strip()
    if re.fullmatch(r"\d+\.0+", text):
        text = text.split(".", 1)[0]
    return text


def _normalize_period(value: str | None) -> str:
    text = str(value or "").strip()
    if re.fullmatch(r"\d{4}-\d{2}", text):
        return text
    if re.fullmatch(r"\d{6}", text):
        return f"{text[:4]}-{text[4:]}"
    return ""
