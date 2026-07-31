from __future__ import annotations

from decimal import Decimal, InvalidOperation, ROUND_HALF_UP


def format_money_full(value: object) -> str:
    try:
        number = int(round(float(value or 0)))
    except (TypeError, ValueError):
        number = 0
    sign = "-" if number < 0 else ""
    return sign + f"{abs(number):,}".replace(",", ".")


def money_axis_scale(values: list[float] | tuple[float, ...]) -> tuple[float, str]:
    max_value = max((abs(float(value or 0)) for value in values), default=0.0)
    if max_value >= 1_000_000_000:
        return 1_000_000_000.0, "tỷ đồng"
    if max_value >= 1_000_000:
        return 1_000_000.0, "triệu đồng"
    return 1.0, "đồng"


def format_money_axis(value: object, *, divisor: float | None = None, unit: str | None = None) -> str:
    try:
        number = float(value or 0)
    except (TypeError, ValueError):
        number = 0.0
    if divisor is None or unit is None:
        divisor, unit = money_axis_scale([number])
    scaled = number / (divisor or 1.0)
    if unit == "đồng":
        return format_money_full(number)
    text = f"{scaled:,.2f}".replace(",", "_").replace(".", ",").replace("_", ".")
    return f"{text} {unit.split()[0]}"


def format_percentage(value: object) -> str:
    try:
        number = Decimal(str(value or 0)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    except (InvalidOperation, TypeError, ValueError):
        number = Decimal("0.00")
    text = f"{number:,.2f}%"
    return text.replace(",", "_").replace(".", ",").replace("_", ".")


def format_number_full(value: object) -> str:
    try:
        number = int(round(float(value or 0)))
    except (TypeError, ValueError):
        number = 0
    sign = "-" if number < 0 else ""
    return sign + f"{abs(number):,}".replace(",", ".")


def format_period(value: object) -> str:
    text = "" if value is None else str(value).strip()
    if len(text) == 7 and text[4] == "-":
        return text
    return text


def format_customer_label(code: object, name: object, *, max_length: int = 32) -> str:
    customer_name = "" if name is None else str(name).strip()
    customer_code = "" if code is None else str(code).strip()
    text = customer_name or customer_code
    if len(text) <= max_length:
        return text
    return text[: max(8, max_length - 3)] + "..."


def format_chart_value(value: object, value_kind: str, *, full: bool = False, divisor: float | None = None, unit: str | None = None) -> str:
    if value_kind.startswith("percent"):
        return format_percentage(value)
    if value_kind.startswith("number"):
        return format_number_full(value)
    if full:
        return format_money_full(value)
    return format_money_axis(value, divisor=divisor, unit=unit)
