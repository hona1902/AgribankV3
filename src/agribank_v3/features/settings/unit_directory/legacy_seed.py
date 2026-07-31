from __future__ import annotations

from agribank_v3.features.settings.unit_directory.models import (
    HEAD_OFFICE,
    TRANSACTION_OFFICE,
)


LEGACY_BRANCHES: tuple[dict[str, object], ...] = (
    {
        "branch_code": "5400",
        "branch_name": "Chi nhánh Lâm Đồng",
        "short_name": "CN Lâm Đồng",
        "province_name": "Lâm Đồng",
        "sort_order": 10,
    },
    {
        "branch_code": "5401",
        "branch_name": "Chi nhánh Lạc Dương",
        "short_name": "CN Lạc Dương",
        "province_name": "Lâm Đồng",
        "sort_order": 20,
    },
    {
        "branch_code": "5404",
        "branch_name": "Chi nhánh Lâm Hà",
        "short_name": "CN Lâm Hà",
        "province_name": "Lâm Đồng",
        "sort_order": 30,
    },
    {
        "branch_code": "5405",
        "branch_name": "Chi nhánh Đơn Dương",
        "short_name": "CN Đơn Dương",
        "province_name": "Lâm Đồng",
        "sort_order": 40,
    },
    {
        "branch_code": "5406",
        "branch_name": "Chi nhánh Đà Lạt",
        "short_name": "CN Đà Lạt",
        "province_name": "Lâm Đồng",
        "sort_order": 50,
    },
    {
        "branch_code": "5412",
        "branch_name": "Chi nhánh Đức Trọng",
        "short_name": "CN Đức Trọng",
        "province_name": "Lâm Đồng",
        "sort_order": 60,
    },
    {
        "branch_code": "5491",
        "branch_name": "Chi nhánh Lộc Phát",
        "short_name": "CN Lộc Phát",
        "province_name": "Lâm Đồng",
        "sort_order": 70,
    },
    {
        "branch_code": "5493",
        "branch_name": "Chi nhánh Đam Rông",
        "short_name": "CN Đam Rông",
        "province_name": "Lâm Đồng",
        "sort_order": 80,
    },
    {
        "branch_code": "5499",
        "branch_name": "Chi nhánh Nam Ban",
        "short_name": "CN Nam Ban",
        "province_name": "Lâm Đồng",
        "sort_order": 90,
    },
)


LEGACY_PGD_OFFICES: tuple[dict[str, object], ...] = (
    {
        "branch_code": "5400",
        "trctcd": "01",
        "office_name": "Phòng giao dịch Hòa Bình",
        "short_name": "PGD Hòa Bình",
        "office_type": TRANSACTION_OFFICE,
        "sort_order": 11,
    },
    {
        "branch_code": "5405",
        "trctcd": "01",
        "office_name": "Phòng giao dịch Ka Đô",
        "short_name": "PGD Ka Đô",
        "office_type": TRANSACTION_OFFICE,
        "sort_order": 41,
    },
    {
        "branch_code": "5405",
        "trctcd": "02",
        "office_name": "Phòng giao dịch Lạc Nghiệp",
        "short_name": "PGD Lạc Nghiệp",
        "office_type": TRANSACTION_OFFICE,
        "sort_order": 42,
    },
    {
        "branch_code": "5406",
        "trctcd": "01",
        "office_name": "Phòng giao dịch Phan Chu Trinh",
        "short_name": "PGD Phan Chu Trinh",
        "office_type": TRANSACTION_OFFICE,
        "sort_order": 51,
    },
    {
        "branch_code": "5406",
        "trctcd": "02",
        "office_name": "Phòng giao dịch Trại Mát",
        "short_name": "PGD Trại Mát",
        "office_type": TRANSACTION_OFFICE,
        "sort_order": 52,
    },
    {
        "branch_code": "5406",
        "trctcd": "03",
        "office_name": "Phòng giao dịch Phù Đổng",
        "short_name": "PGD Phù Đổng",
        "office_type": TRANSACTION_OFFICE,
        "sort_order": 53,
    },
    {
        "branch_code": "5412",
        "trctcd": "01",
        "office_name": "Phòng giao dịch Liên Khương",
        "short_name": "PGD Liên Khương",
        "office_type": TRANSACTION_OFFICE,
        "sort_order": 61,
    },
)


def legacy_head_offices() -> tuple[dict[str, object], ...]:
    rows: list[dict[str, object]] = []
    for branch in LEGACY_BRANCHES:
        short_name = str(branch.get("short_name") or "")
        branch_label = short_name[3:] if short_name.startswith("CN ") else short_name
        rows.append(
            {
                "branch_code": str(branch["branch_code"]),
                "trctcd": "00",
                "office_name": f"Hội sở Chi nhánh {branch_label}",
                "short_name": "Hội sở",
                "office_type": HEAD_OFFICE,
                "sort_order": int(branch.get("sort_order") or 0),
            }
        )
    return tuple(rows)


LEGACY_OFFICES: tuple[dict[str, object], ...] = (
    *legacy_head_offices(),
    *LEGACY_PGD_OFFICES,
)

