from __future__ import annotations

from dataclasses import dataclass, replace


CUSTOMER_TYPE_FILTERS: tuple[tuple[str, str], ...] = (
    ("Tất cả", ""),
    ("Cá nhân", "CN"),
    ("Tổ chức/Pháp nhân", "TC"),
    ("Khác", "OTHER"),
)

LOAN_TERM_FILTERS: tuple[tuple[str, str], ...] = (
    ("Tất cả", ""),
    ("Có dư nợ ngắn hạn", "SHORT_TERM"),
    ("Có dư nợ trung/dài hạn", "MEDIUM_LONG_TERM"),
    ("Có dư nợ chưa phân loại", "OTHER"),
)

OFFICER_STATUS_FILTERS: tuple[tuple[str, str], ...] = (
    ("Tất cả", ""),
    ("Một cán bộ", "one"),
    ("Nhiều cán bộ", "multiple"),
    ("Có override", "override"),
    ("Không có override", "no_override"),
)

MULTIPLE_OFFICER_FILTERS: tuple[tuple[str, str], ...] = (
    ("Nhiều cán bộ trong cùng kỳ", "same_period"),
    ("Thay đổi cán bộ qua các kỳ", "changed_period"),
    ("Có override", "override"),
    ("Không có override", "no_override"),
)

MOVEMENT_STATUS_NEW = "Vay mới"
MOVEMENT_STATUS_PAID_OFF = "Tất toán"
MOVEMENT_STATUS_INCREASE = "Tăng dư nợ"
MOVEMENT_STATUS_DECREASE = "Giảm dư nợ"
MOVEMENT_STATUS_UNCHANGED = "Không thay đổi"

MOVEMENT_STATUS_FILTERS: tuple[tuple[str, str], ...] = (
    ("Tất cả", ""),
    ("Khách hàng vay mới", MOVEMENT_STATUS_NEW),
    ("Khách hàng tất toán", MOVEMENT_STATUS_PAID_OFF),
    ("Khách hàng tăng dư nợ", MOVEMENT_STATUS_INCREASE),
    ("Khách hàng giảm dư nợ", MOVEMENT_STATUS_DECREASE),
    ("Không thay đổi", MOVEMENT_STATUS_UNCHANGED),
)

OVERRIDE_STATUS_FILTERS: tuple[tuple[str, str], ...] = (
    ("Tất cả", ""),
    ("Có override", "override"),
    ("Không có override", "no_override"),
)


@dataclass(frozen=True, slots=True)
class CustomerFilters:
    period_from: str = ""
    period_to: str = ""
    current_period: str = ""
    compare_period: str = ""
    branch_code: str = ""
    customer_type: str = ""
    officer: str = ""
    loan_term: str = ""
    search_text: str = ""
    movement_status: str = ""
    multi_status: str = ""
    override_status: str = ""

    def without_exact_period(self) -> "CustomerFilters":
        return replace(self, current_period="")

    def with_current_period(self, period: str) -> "CustomerFilters":
        return replace(self, current_period=str(period or "").strip())

    def with_compare_period(self, period: str) -> "CustomerFilters":
        return replace(self, compare_period=str(period or "").strip())


def clean_filter_text(value: object) -> str:
    return "" if value is None else str(value).strip()
