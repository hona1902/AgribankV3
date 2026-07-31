from __future__ import annotations

from agribank_v3.features.catalog import Feature
from agribank_v3.features.credit.summary.models import (
    CREDIT_LIMIT_TITLE,
    LOAN_COMPARE_TITLE,
    NIM_DN_TITLE,
    NIM_NV_TITLE,
)


SUMMARY_FEATURES: tuple[Feature, ...] = (
    Feature(
        NIM_DN_TITLE,
        "Import NIM dư nợ và xuất báo cáo đối chiếu theo VBA.",
        "qttd.png",
    ),
    Feature(
        NIM_NV_TITLE,
        "Import NIM nguồn vốn và xuất báo cáo đối chiếu theo VBA.",
        "qttd.png",
    ),
    Feature(
        LOAN_COMPARE_TITLE,
        "Đối chiếu biến động dư nợ khách hàng giữa hai kỳ và lưu lịch sử.",
        "m09b.png",
    ),
    Feature(
        CREDIT_LIMIT_TITLE,
        "Theo dõi hợp đồng hạn mức tín dụng đã hết hạn hoặc sắp hết hạn.",
        "m15A.png",
    ),
)
