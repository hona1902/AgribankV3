from __future__ import annotations

from datetime import date, datetime
from decimal import Decimal, InvalidOperation
import re
from typing import Any


_TELEX_CODES = (
    "aws", "awf", "awr", "awx", "awj",
    "aas", "aaf", "aar", "aax", "aaj",
    "ees", "eef", "eer", "eex", "eej",
    "oos", "oof", "oor", "oox", "ooj",
    "ows", "owf", "owr", "owx", "owj",
    "uws", "uwf", "uwr", "uwx", "uwj",
    "as", "af", "ar", "ax", "aj", "aw", "aa", "dd",
    "es", "ef", "er", "ex", "ej", "ee",
    "is", "if", "ir", "ix", "ij",
    "os", "of", "or", "ox", "oj", "oo", "ow",
    "us", "uf", "ur", "ux", "uj", "uw",
    "ys", "yf", "yr", "yx", "yj",
)
_TELEX_CHARACTERS = (
    "ắ", "ằ", "ẳ", "ẵ", "ặ",
    "ấ", "ầ", "ẩ", "ẫ", "ậ",
    "ế", "ề", "ể", "ễ", "ệ",
    "ố", "ồ", "ổ", "ỗ", "ộ",
    "ớ", "ờ", "ở", "ỡ", "ợ",
    "ứ", "ừ", "ử", "ữ", "ự",
    "á", "à", "ả", "ã", "ạ", "ă", "â", "đ",
    "é", "è", "ẻ", "ẽ", "ẹ", "ê",
    "í", "ì", "ỉ", "ĩ", "ị",
    "ó", "ò", "ỏ", "õ", "ọ", "ô", "ơ",
    "ú", "ù", "ủ", "ũ", "ụ", "ư",
    "ý", "ỳ", "ỷ", "ỹ", "ỵ",
)
_YYYYMMDD = re.compile(r"^\d{8}$")


def excel_column_name(index: int) -> str:
    if index < 1:
        raise ValueError("Excel column index must be positive.")
    result = ""
    while index:
        index, remainder = divmod(index - 1, 26)
        result = chr(65 + remainder) + result
    return result


def parse_yyyymmdd(value: Any) -> date | None:
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    text = str(value or "").strip()
    if not _YYYYMMDD.fullmatch(text):
        return None
    try:
        return datetime.strptime(text, "%Y%m%d").date()
    except ValueError:
        return None


def vietnamese_report_date(value: Any) -> str:
    parsed = parse_yyyymmdd(value)
    if parsed is None:
        return ""
    return f"Ngày {parsed.day:02d} tháng {parsed.month:02d} năm {parsed.year}"


def normalize_customer_id(
    value: Any,
    branch_code: str,
    *,
    include_branch: bool,
) -> str:
    customer_id = str(value or "").strip()
    if customer_id.startswith("'"):
        customer_id = customer_id[1:]
    return f"{branch_code.strip()}{customer_id}" if include_branch else customer_id


def parse_vietnamese_amount(value: Any) -> Decimal:
    if value is None or isinstance(value, bool):
        return Decimal(0)
    if isinstance(value, Decimal):
        return value
    if isinstance(value, (int, float)):
        return Decimal(str(value))
    text = str(value).strip()
    if not text or text == "-":
        return Decimal(0)
    text = text.replace("\u00a0", "").replace(" ", "")
    text = text.replace(".", "").replace(",", ".")
    try:
        return Decimal(text)
    except InvalidOperation:
        return Decimal(0)


def is_branch_code(value: Any) -> bool:
    text = str(value or "").replace("\u00a0", " ").strip()
    return text.isdigit() and 3 <= len(text) <= 6


def is_code_like(value: Any) -> bool:
    text = str(value or "").replace("\u00a0", " ").strip()
    if not text:
        return False
    prefix = text.split("_", 1)[0].strip()
    return bool(prefix) and prefix.isdigit()


def decode_telex(value: str) -> str:
    """Port of dichchu_telex, preserving its replacement order."""
    result = value
    for code, character in zip(_TELEX_CODES, _TELEX_CHARACTERS, strict=True):
        result = result.replace(code, character)
        result = result.replace(code.upper(), character.upper())
    return result
