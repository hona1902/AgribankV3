from __future__ import annotations

from dataclasses import dataclass

from agribank_v3.features.credit.tovayvon.models import (
    ASSOCIATION_FARMERS_UNION,
    ASSOCIATION_OTHER,
    ASSOCIATION_TYPE_LABELS,
    ASSOCIATION_WOMENS_UNION,
)


DETAIL_BY_GROUP = "DETAIL_BY_GROUP"
SUMMARY_BY_ASSOCIATION = "SUMMARY_BY_ASSOCIATION"
ASSOCIATION_UNKNOWN = "UNKNOWN"
ASSOCIATION_UNKNOWN_LABEL = "Chưa xác định"
GROUP_STATUS_DECLARED = "DECLARED"
GROUP_STATUS_INACTIVE = "INACTIVE"
GROUP_STATUS_NOT_DECLARED = "NOT_DECLARED"
GROUP_STATUS_LABELS = {
    GROUP_STATUS_DECLARED: "Đã khai báo",
    GROUP_STATUS_INACTIVE: "Ngừng sử dụng",
    GROUP_STATUS_NOT_DECLARED: "Chưa khai báo",
}
ASSOCIATION_FILTER_OPTIONS = (
    ("Tất cả", ""),
    (ASSOCIATION_TYPE_LABELS[ASSOCIATION_FARMERS_UNION], ASSOCIATION_FARMERS_UNION),
    (ASSOCIATION_TYPE_LABELS[ASSOCIATION_WOMENS_UNION], ASSOCIATION_WOMENS_UNION),
    (ASSOCIATION_TYPE_LABELS[ASSOCIATION_OTHER], ASSOCIATION_OTHER),
    (ASSOCIATION_UNKNOWN_LABEL, ASSOCIATION_UNKNOWN),
)


@dataclass(frozen=True, slots=True)
class GroupLendingFilters:
    period: str = ""
    from_period: str = ""
    to_period: str = ""
    branch_code: str = ""
    office_code: str = ""
    association_type: str = ""
    group_status: str = ""
    officer: str = ""
    search: str = ""
    include_unknown_groups: bool = True


@dataclass(frozen=True, slots=True)
class GroupDirectoryEntry:
    group_code: str
    group_name: str
    association_type: str
    association_label: str
    association_other_name: str
    branch_name: str
    office_name: str
    commune: str
    leader_name: str
    active: bool
    status: str
    status_label: str


@dataclass(frozen=True, slots=True)
class GroupLendingKpi:
    label: str
    value: object
    kind: str = "number"
    tooltip: str = ""
    from_value: object = None
    to_value: object = None
    difference: object = None
    growth_rate: float | None = None


@dataclass(frozen=True, slots=True)
class GroupLendingRow:
    period: str
    group_code: str
    group_name: str
    association_type: str
    association_label: str
    association_other_name: str
    branch_code: str
    office_name: str
    commune: str
    leader_name: str
    member_count: int
    loan_count: int
    total_balance: float
    average_balance_per_member: float | None
    status: str
    status_label: str

    def to_dict(self, stt: int | None = None) -> dict[str, object]:
        values: dict[str, object] = {}
        if stt is not None:
            values["STT"] = stt
        values.update(
            {
                "Kỳ": self.period,
                "Mã tổ": self.group_code,
                "Tên tổ": self.group_name,
                "Loại tổ chức Hội": self.association_label,
                "Tên tổ chức khác": self.association_other_name,
                "Chi nhánh": self.branch_code,
                "Hội sở/PGD": self.office_name,
                "Xã": self.commune,
                "Tổ trưởng": self.leader_name,
                "Số tổ viên còn dư nợ": self.member_count,
                "Số món": self.loan_count,
                "Tổng dư nợ": self.total_balance,
                "Dư nợ bình quân/tổ viên": self.average_balance_per_member,
                "Trạng thái danh mục": self.status_label,
            }
        )
        return values


@dataclass(frozen=True, slots=True)
class GroupAssociationSummaryRow:
    association_type: str
    association_label: str
    group_count: int
    unique_member_count: int
    member_occurrence_count: int
    total_balance: float
    share: float | None
    average_balance_per_group: float | None
    average_balance_per_member: float | None

    def to_dict(self) -> dict[str, object]:
        return {
            "Loại tổ chức Hội": self.association_label,
            "Số tổ có dư nợ": self.group_count,
            "Số tổ viên duy nhất": self.unique_member_count,
            "Tổng lượt tổ viên theo tổ": self.member_occurrence_count,
            "Tổng dư nợ": self.total_balance,
            "Tỷ trọng": self.share,
            "Dư nợ bình quân/tổ": self.average_balance_per_group,
            "Dư nợ bình quân/tổ viên": self.average_balance_per_member,
        }


@dataclass(frozen=True, slots=True)
class GroupLendingComparisonRow:
    group_code: str
    group_name: str
    association_label: str
    member_count_from: int
    member_count_to: int
    member_change: int
    balance_from: float
    balance_to: float
    balance_change: float
    balance_growth_rate: float | None
    movement_category: str

    def to_dict(self) -> dict[str, object]:
        return {
            "Mã tổ": self.group_code,
            "Tên tổ": self.group_name,
            "Loại Hội": self.association_label,
            "Tổ viên Từ kỳ": self.member_count_from,
            "Tổ viên Đến kỳ": self.member_count_to,
            "Tăng/giảm tổ viên": self.member_change,
            "Dư nợ Từ kỳ": self.balance_from,
            "Dư nợ Đến kỳ": self.balance_to,
            "Tăng/giảm dư nợ": self.balance_change,
            "Tăng trưởng dư nợ (%)": self.balance_growth_rate,
            "Phân loại biến động": self.movement_category,
        }


@dataclass(frozen=True, slots=True)
class GroupAssociationComparisonRow:
    association_type: str
    association_label: str
    group_count_from: int
    group_count_to: int
    group_count_change: int
    member_count_from: int
    member_count_to: int
    member_count_change: int
    balance_from: float
    balance_to: float
    balance_change: float
    growth_rate: float | None
    share_from: float | None
    share_to: float | None
    share_change_pp: float | None

    def to_dict(self) -> dict[str, object]:
        return {
            "Loại Hội": self.association_label,
            "Số tổ Từ kỳ": self.group_count_from,
            "Số tổ Đến kỳ": self.group_count_to,
            "Thay đổi số tổ": self.group_count_change,
            "Tổ viên Từ kỳ": self.member_count_from,
            "Tổ viên Đến kỳ": self.member_count_to,
            "Thay đổi tổ viên": self.member_count_change,
            "Dư nợ Từ kỳ": self.balance_from,
            "Dư nợ Đến kỳ": self.balance_to,
            "Tăng/giảm dư nợ": self.balance_change,
            "Tăng trưởng (%)": self.growth_rate,
            "Tỷ trọng Từ kỳ": self.share_from,
            "Tỷ trọng Đến kỳ": self.share_to,
            "Thay đổi tỷ trọng (điểm %)": self.share_change_pp,
        }


@dataclass(frozen=True, slots=True)
class GroupLendingResult:
    period: str
    rows: tuple[GroupLendingRow | GroupAssociationSummaryRow | GroupLendingComparisonRow | GroupAssociationComparisonRow, ...]
    kpis: tuple[GroupLendingKpi, ...]
    total_rows: int
    page: int = 1
    page_size: int = 100
    notes: tuple[str, ...] = ()
    diagnostics: dict[str, object] | None = None
