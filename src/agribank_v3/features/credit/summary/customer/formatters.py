from __future__ import annotations

from decimal import Decimal, InvalidOperation, ROUND_HALF_UP

from agribank_v3.features.credit.summary.customer.models import CustomerTypeCode


def format_money_vn(value: object, *, signed: bool = False) -> str:
    try:
        number = int(round(float(value or 0)))
    except (TypeError, ValueError):
        number = 0
    prefix = "+" if signed and number > 0 else ""
    return prefix + f"{number:,}".replace(",", ".")


def format_percent_vn(value: object, *, signed: bool = False, empty: str = "N/A") -> str:
    if value is None:
        return empty
    try:
        number = Decimal(str(value or 0)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    except (InvalidOperation, TypeError, ValueError):
        number = Decimal("0.00")
    prefix = "+" if signed and number > 0 else ""
    text = f"{number:,.2f}%"
    return prefix + text.replace(",", "_").replace(".", ",").replace("_", ".")


def format_customer_type(value: object) -> str:
    code = "" if value is None else str(value).strip().upper()
    if code == CustomerTypeCode.PERSONAL.value:
        return "Cá nhân"
    if code == CustomerTypeCode.ORGANIZATION.value:
        return "Tổ chức/Pháp nhân"
    if code == CustomerTypeCode.OTHER.value:
        return "Khác"
    return str(value or "")


def format_officer(code: object, name: object) -> str:
    officer_name = "" if name is None else str(name).strip()
    officer_code = "" if code is None else str(code).strip()
    return officer_name or officer_code


def normalize_officer_code(value: object) -> str:
    return "" if value is None else str(value).strip()


def normalize_officer_name(value: object) -> str:
    return " ".join(str(value or "").strip().split())


def format_bool_yes_no(value: object) -> str:
    return "Có" if bool(value) else "Không"


def format_override_status(value: object) -> str:
    return "Có override" if bool(value) else "Không override"
