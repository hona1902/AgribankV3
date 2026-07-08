from __future__ import annotations

from agribank_v3.settlement.models import (
    SettlementCategory,
    SettlementSourceMode,
    SettlementSpec,
)


def _spec(
    key: str,
    report_code: str,
    title: str,
    category: SettlementCategory,
    source_hint: str,
    family: str,
    entrypoint: str,
    *,
    source_mode: SettlementSourceMode = SettlementSourceMode.ACTIVE_WORKBOOK,
    active: bool = True,
) -> SettlementSpec:
    return SettlementSpec(
        key=key,
        report_code=report_code,
        title=title,
        category=category,
        source_hint=source_hint,
        source_mode=source_mode,
        processor_family=family,
        legacy_entrypoint=entrypoint,
        active=active,
    )


ACTIVE_SETTLEMENT_SPECS: tuple[SettlementSpec, ...] = (
    _spec("credit.05", "05", "Mẫu 05/QT", SettlementCategory.CREDIT,
          "{MaCN}_rt05.csv", "mau05", "QUYETTOAN_MAU05"),
    _spec("credit.06", "06", "Mẫu 06/QT", SettlementCategory.CREDIT,
          "{MaCN}QT05.xlsx", "mau06", "QUYETTOAN_MAU06"),
    _spec("credit.15a", "15A", "Mẫu 15A/QT", SettlementCategory.CREDIT,
          "{MaCN}_rt15a.csv", "mau15_16", "QUYETTOAN_MAU15A"),
    _spec("credit.15b", "15B", "Mẫu 15B/QT", SettlementCategory.CREDIT,
          "{MaCN}_rt15b.csv", "mau15_16", "QUYETTOAN_MAU15B"),
    _spec("credit.16", "16", "Mẫu 16/QT", SettlementCategory.CREDIT,
          "{MaCN}_rt16.csv", "mau15_16", "QUYETTOAN_MAU16"),
    _spec("credit.18", "18", "Mẫu 18/QT", SettlementCategory.CREDIT,
          "{MaCN}_rt18.csv", "mau18", "QUYETTOAN_MAU18"),
    _spec("credit.20a", "20a", "Mẫu 20a/QT", SettlementCategory.CREDIT,
          "{MaCN}_rt20.xls", "mau20a", "QUYETTOAN_MAU20a"),
    _spec(
        "credit.30a",
        "30A",
        "Mẫu 30A/QT tín dụng",
        SettlementCategory.CREDIT,
        "Workbook quyết toán tín dụng",
        "mau30",
        "QUYETTOAN_MAU30aTD",
        source_mode=SettlementSourceMode.GENERATED_WORKBOOK,
    ),
    _spec("accounting.04", "04", "Mẫu 04/QT", SettlementCategory.ACCOUNTING,
          "IC_100435", "mau04", "QUYETTOAN_MAU04"),
    _spec("accounting.07a", "07A", "Mẫu 07A/QT", SettlementCategory.ACCOUNTING,
          "WT_100642", "mau07", "QUYETTOAN_MAU07_642"),
    _spec("accounting.08", "08", "Mẫu 08/QT", SettlementCategory.ACCOUNTING,
          "FA_100586", "mau08", "QUYETTOAN_MAU08"),
    _spec("accounting.09a", "09a", "Mẫu 09a/QT", SettlementCategory.ACCOUNTING,
          "Màn hình mshr32 - xuất Excel báo cáo TMBCTC_TSCD001", "mau09", "QUYETTOAN_MAU09a"),
    _spec("accounting.09b", "09b", "Mẫu 09b/QT", SettlementCategory.ACCOUNTING,
          "Màn hình mshr32 - xuất Excel báo cáo TMBCTC_TSCD002", "mau09", "QUYETTOAN_MAU09b"),
    _spec("accounting.09c", "09c", "Mẫu 09c/QT", SettlementCategory.ACCOUNTING,
          "Màn hình mshr32 - xuất Excel báo cáo TMBCTC_TSCD003", "mau09", "QUYETTOAN_MAU09c"),
    _spec("accounting.13", "13", "Mẫu 13/QT", SettlementCategory.ACCOUNTING,
          "{MaCN}_rt13.csv", "mau13_14", "QUYETTOAN_MAU13"),
    _spec("accounting.14", "14", "Mẫu 14/QT", SettlementCategory.ACCOUNTING,
          "{MaCN}_rt14.csv", "mau13_14", "QUYETTOAN_MAU14"),
    _spec(
        "accounting.22",
        "22",
        "Mẫu 22/QT",
        SettlementCategory.ACCOUNTING,
        "Một hoặc nhiều file GL_glst34",
        "mau22",
        "QUYETTOAN_MAU22",
        source_mode=SettlementSourceMode.MULTIPLE_FILES,
    ),
    _spec(
        "accounting.23",
        "23",
        "Mẫu 23/QT",
        SettlementCategory.ACCOUNTING,
        "Một hoặc nhiều file GL_glcb06",
        "mau23",
        "QUYETTOAN_MAU23",
        source_mode=SettlementSourceMode.MULTIPLE_FILES,
    ),
    _spec("accounting.24", "24", "Mẫu 24/QT", SettlementCategory.ACCOUNTING,
          "{MaCN}_rt24.csv", "mau24", "QUYETTOAN_MAU24"),
    _spec(
        "accounting.30a",
        "30A",
        "Mẫu 30A/QT kế toán",
        SettlementCategory.ACCOUNTING,
        "Workbook quyết toán kế toán",
        "mau30",
        "QUYETTOAN_MAU30aKT",
        source_mode=SettlementSourceMode.GENERATED_WORKBOOK,
    ),
    *(
        _spec(
            f"consolidation.{code.casefold()}",
            code.upper(),
            f"Tổng hợp Mẫu {code.upper()}/QT",
            SettlementCategory.CONSOLIDATION,
            f"Danh sách file Mẫu {code.upper()}/QT",
            f"summary_{family}",
            entrypoint,
            source_mode=SettlementSourceMode.MULTIPLE_FILES,
        )
        for code, family, entrypoint in (
            ("05", "05", "TH_QT_05"),
            ("06", "06", "TH_QT_06"),
            ("13", "13_14", "TH_QT_13"),
            ("14", "13_14", "TH_QT_14"),
            ("15a", "15_16", "TH_QT_15a"),
            ("15b", "15_16", "TH_QT_15b"),
            ("16", "15_16", "TH_QT_16"),
            ("18", "18", "TH_QT_18"),
            ("30a", "30", "TH_QT_30"),
        )
    ),
)


LEGACY_SETTLEMENT_SPECS: tuple[SettlementSpec, ...] = (
    _spec("legacy.02.loan.rt15", "02", "Dữ liệu 02/QT từ rt15",
          SettlementCategory.LEGACY, "{MaCN}_rt15.csv", "mau02",
          "QUYETTOAN_MAU02_LN_RT15", active=False),
    _spec("legacy.02.deposit.rtdp", "02", "Dữ liệu 02/QT từ rtdp",
          SettlementCategory.LEGACY, "{MaCN}_rtdp.csv", "mau02",
          "QUYETTOAN_MAU02_DP_RTDP", active=False),
    _spec("legacy.02.loan.mis80", "02", "Dữ liệu 02/QT từ MIS80",
          SettlementCategory.LEGACY, "msit80", "mau02",
          "QUYETTOAN_MAU02_LN_MIS80", active=False),
    _spec("legacy.02.deposit.mis81", "02", "Dữ liệu 02/QT từ MIS81",
          SettlementCategory.LEGACY, "msit81", "mau02",
          "QUYETTOAN_MAU02_DP_MIS81", active=False),
    _spec("legacy.04.2023", "04", "Mẫu 04/QT 2023",
          SettlementCategory.LEGACY, "IC_100435", "mau04",
          "QUYETTOAN_MAU04_2023", active=False),
    _spec("legacy.04.old", "04", "Mẫu 04/QT cũ",
          SettlementCategory.LEGACY, "IC_100435", "mau04",
          "QUYETTOAN_MAU04cu", active=False),
    _spec("legacy.07b", "07B", "Mẫu 07B/QT",
          SettlementCategory.LEGACY, "WT_100643", "mau07",
          "QUYETTOAN_MAU07_643", active=False),
)


SETTLEMENT_SPECS: dict[str, SettlementSpec] = {
    spec.key: spec
    for spec in (*ACTIVE_SETTLEMENT_SPECS, *LEGACY_SETTLEMENT_SPECS)
}
