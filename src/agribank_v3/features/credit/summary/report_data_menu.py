from __future__ import annotations

from agribank_v3.features.catalog import Feature
from agribank_v3.features.credit.summary.models import (
    REPORT_DATA_TITLE,
    REPORT_SUMMARY_TITLE,
)


REPORT_SUMMARY_ROUTE = "credit.report_summary"


REPORT_DATA_FEATURES: tuple[Feature, ...] = (
    Feature(
        REPORT_SUMMARY_TITLE,
        "Tổng hợp số liệu báo cáo theo kỳ từ LN01 và dư nợ thẻ DN15.",
        "m15A.png",
    ),
)


REPORT_DATA_RESERVED_SLOTS = 3


__all__ = [
    "REPORT_DATA_FEATURES",
    "REPORT_DATA_RESERVED_SLOTS",
    "REPORT_DATA_TITLE",
    "REPORT_SUMMARY_ROUTE",
    "REPORT_SUMMARY_TITLE",
]
