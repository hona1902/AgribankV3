from __future__ import annotations

from dataclasses import dataclass

from agribank_v3.features.credit.summary.models import SummaryDataType
from agribank_v3.features.credit.summary.officer_history.models import (
    METRIC_AVERAGE_RATE,
    METRIC_BALANCE,
    METRIC_BALANCE_GROWTH,
    METRIC_NIM_AFTER,
    METRIC_NIM_BEFORE,
)


@dataclass(frozen=True, slots=True)
class NimUiConfig:
    data_type: SummaryDataType
    main_title: str
    dashboard_title: str
    officer_history_title: str
    officer_label: str
    officer_short_label: str
    officer_selector_placeholder: str
    officer_selector_counter_label: str
    balance_label: str
    total_balance_label: str
    current_balance_label: str
    growth_label: str
    growth_percent_label: str
    balance_delta_label: str
    include_average_rate: bool
    dashboard_sheets: dict[str, str]

    def metric_labels(self) -> dict[str, str]:
        labels = {
            METRIC_BALANCE: self.balance_label,
            METRIC_NIM_BEFORE: "NIM trước ĐC",
            METRIC_NIM_AFTER: "NIM sau ĐC",
            METRIC_BALANCE_GROWTH: self.growth_label,
        }
        if self.include_average_rate:
            labels[METRIC_AVERAGE_RATE] = "Lãi suất bình quân"
        return labels

    def metric_order(self, *, include_growth: bool) -> tuple[str, ...]:
        metrics = [METRIC_BALANCE]
        if self.include_average_rate:
            metrics.append(METRIC_AVERAGE_RATE)
        metrics.extend([METRIC_NIM_BEFORE, METRIC_NIM_AFTER])
        if include_growth:
            metrics.append(METRIC_BALANCE_GROWTH)
        return tuple(metrics)


NIM_DN_UI_CONFIG = NimUiConfig(
    data_type=SummaryDataType.NIM_DN,
    main_title="NIM dư nợ",
    dashboard_title="Dashboard NIM chi nhánh - AgribankV3",
    officer_history_title="Phân tích NIM cán bộ tín dụng - AgribankV3",
    officer_label="Người quản lý KV",
    officer_short_label="CBTD",
    officer_selector_placeholder="Chọn CBTD",
    officer_selector_counter_label="CBTD",
    balance_label="Dư nợ",
    total_balance_label="Tổng dư nợ",
    current_balance_label="Tổng dư nợ kỳ hiện tại",
    growth_label="Tăng trưởng dư nợ",
    growth_percent_label="Tăng trưởng dư nợ (%)",
    balance_delta_label="Tăng/giảm dư nợ tuyệt đối",
    include_average_rate=True,
    dashboard_sheets={
        "overview": "TongQuanTheoKy",
        "branch": "SoSanhChiNhanh",
        "growth": "TangTruong",
        "detail": "BangDuLieuChiTiet",
    },
)

NIM_NV_UI_CONFIG = NimUiConfig(
    data_type=SummaryDataType.NIM_NV,
    main_title="NIM nguồn vốn",
    dashboard_title="Dashboard NIM nguồn vốn - AgribankV3",
    officer_history_title="Phân tích NIM cán bộ nguồn vốn - AgribankV3",
    officer_label="Người quản lý NV",
    officer_short_label="Cán bộ",
    officer_selector_placeholder="Chọn cán bộ",
    officer_selector_counter_label="cán bộ",
    balance_label="Số dư nguồn vốn",
    total_balance_label="Tổng nguồn vốn",
    current_balance_label="Tổng số dư nguồn vốn kỳ hiện tại",
    growth_label="Tăng trưởng nguồn vốn",
    growth_percent_label="Tăng trưởng nguồn vốn (%)",
    balance_delta_label="Tăng/giảm nguồn vốn tuyệt đối",
    include_average_rate=False,
    dashboard_sheets={
        "overview": "TongQuanNIMNV",
        "branch": "SoSanhChiNhanhNIMNV",
        "growth": "TangTruongNIMNV",
        "detail": "BangDuLieuNIMNV",
    },
)


def get_nim_ui_config(data_type: SummaryDataType) -> NimUiConfig:
    if data_type == SummaryDataType.NIM_NV:
        return NIM_NV_UI_CONFIG
    return NIM_DN_UI_CONFIG
