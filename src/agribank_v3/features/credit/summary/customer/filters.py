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

DEBT_GROUP_ALL = "ALL"
DEBT_GROUP_HAS_GROUP_1 = "HAS_GROUP_1"
DEBT_GROUP_HAS_GROUP_2 = "HAS_GROUP_2"
DEBT_GROUP_HAS_GROUP_3 = "HAS_GROUP_3"
DEBT_GROUP_HAS_GROUP_4 = "HAS_GROUP_4"
DEBT_GROUP_HAS_GROUP_5 = "HAS_GROUP_5"
DEBT_GROUP_ATTENTION = "ATTENTION"
DEBT_GROUP_BAD_DEBT = "BAD_DEBT"
DEBT_GROUP_UNKNOWN = "UNKNOWN"
DEBT_GROUP_WORST_1 = "WORST_1"
DEBT_GROUP_WORST_2 = "WORST_2"
DEBT_GROUP_WORST_3 = "WORST_3"
DEBT_GROUP_WORST_4 = "WORST_4"
DEBT_GROUP_WORST_5 = "WORST_5"

DEBT_GROUP_FILTERS: tuple[tuple[str, str], ...] = (
    ("Tất cả nhóm nợ", DEBT_GROUP_ALL),
    ("Có dư nợ nhóm 1", DEBT_GROUP_HAS_GROUP_1),
    ("Có dư nợ nhóm 2", DEBT_GROUP_HAS_GROUP_2),
    ("Có dư nợ nhóm 3", DEBT_GROUP_HAS_GROUP_3),
    ("Có dư nợ nhóm 4", DEBT_GROUP_HAS_GROUP_4),
    ("Có dư nợ nhóm 5", DEBT_GROUP_HAS_GROUP_5),
    ("Nợ cần chú ý", DEBT_GROUP_ATTENTION),
    ("Nợ xấu", DEBT_GROUP_BAD_DEBT),
    ("Chưa xác định nhóm", DEBT_GROUP_UNKNOWN),
    ("Nhóm nợ cao nhất: 01", DEBT_GROUP_WORST_1),
    ("Nhóm nợ cao nhất: 02", DEBT_GROUP_WORST_2),
    ("Nhóm nợ cao nhất: 03", DEBT_GROUP_WORST_3),
    ("Nhóm nợ cao nhất: 04", DEBT_GROUP_WORST_4),
    ("Nhóm nợ cao nhất: 05", DEBT_GROUP_WORST_5),
)

DEBT_GROUP_FILTER_KEYS = frozenset(value for _label, value in DEBT_GROUP_FILTERS)

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
    debt_group: str = ""

    def without_exact_period(self) -> "CustomerFilters":
        return replace(self, current_period="")

    def with_current_period(self, period: str) -> "CustomerFilters":
        return replace(self, current_period=str(period or "").strip())

    def with_compare_period(self, period: str) -> "CustomerFilters":
        return replace(self, compare_period=str(period or "").strip())


def clean_filter_text(value: object) -> str:
    return "" if value is None else str(value).strip()
