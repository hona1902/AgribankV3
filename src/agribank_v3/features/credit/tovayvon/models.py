from __future__ import annotations

from dataclasses import asdict, dataclass
from datetime import datetime
from typing import Any


DATA_TVV_HEADERS: tuple[str, ...] = (
    "STT",
    "MaTo",
    "TenTo",
    "TenTVV_DayDu",
    "Xa",
    "MaToTruong",
    "Ten_ToTruong",
    "DiaChi",
    "TK_ToTruong",
    "SoDienThoai",
    "ToHoi",
    "TK_ToHoiXa",
    "ToChuc",
    "Ten_Huyen",
    "TK_HUYEN",
    "Ten_Tinh",
    "TK_TINH",
    "Ten_TW",
    "TK_TW",
    "uyquyen",
    "TTLN_TW",
    "TTLN_Tinh",
)


COMMISSION_EXPORT_HEADERS: tuple[str, ...] = (
    "HH_TyLeChung_KhongTSBD",
    "HH_TyLeChung_CoTSBD",
    "HH_KhongBD_ToTruong",
    "HH_KhongBD_CapXa",
    "HH_KhongBD_CapHuyen",
    "HH_KhongBD_CapTinh",
    "HH_KhongBD_CapTW",
    "HH_KhongBD_Tong",
    "HH_CoBDTS_ToTruong",
    "HH_CoBDTS_CapXa",
    "HH_CoBDTS_CapHuyen",
    "HH_CoBDTS_CapTinh",
    "HH_CoBDTS_CapTW",
    "HH_CoBDTS_Tong",
)


COMMISSION_RULE_EXPORT_HEADERS: tuple[str, ...] = (
    "DK_ThuLai_Min_1",
    "DK_ThuLai_Max_1",
    "DK_TyLeChi_1",
    "DK_ThuLai_Min_2",
    "DK_ThuLai_Max_2",
    "DK_TyLeChi_2",
    "DK_ThuLai_Min_3",
    "DK_TyLeChi_3",
    "DK_NoXau_Nguong",
    "DK_NoXau_TyLeChi",
)


DATA_TVV_TEMPLATE_HEADERS: tuple[str, ...] = (
    DATA_TVV_HEADERS + COMMISSION_EXPORT_HEADERS + COMMISSION_RULE_EXPORT_HEADERS
)


DATA_TVV_FIELD_LABELS: dict[str, str] = {
    "STT": "STT",
    "MaTo": "Mã số tổ",
    "TenTo": "Tên tổ vay vốn",
    "TenTVV_DayDu": "Tên tổ VV đầy đủ",
    "Xa": "Xã",
    "MaToTruong": "Mã số KH tổ trưởng",
    "Ten_ToTruong": "Tên tổ trưởng",
    "DiaChi": "Địa chỉ",
    "TK_ToTruong": "Tài khoản tổ trưởng",
    "SoDienThoai": "Số điện thoại",
    "ToHoi": "Tổ Hội cấp xã",
    "TK_ToHoiXa": "Tài khoản tổ hội cấp xã",
    "ToChuc": "Tổ chức quản lý cấp xã",
    "Ten_Huyen": "Tổ Hội cấp huyện",
    "TK_HUYEN": "Tài khoản tổ hội cấp huyện",
    "Ten_Tinh": "Tổ Hội cấp tỉnh",
    "TK_TINH": "Tài khoản tổ hội cấp tỉnh",
    "Ten_TW": "Tổ Hội cấp TW",
    "TK_TW": "Tài khoản tổ hội TW",
    "uyquyen": "Ủy quyền thu lãi",
    "TTLN_TW": "Thỏa thuận liên ngành TW",
    "TTLN_Tinh": "Thỏa thuận liên ngành Tỉnh",
}


DATA_TVV_ATTR_TO_HEADER: dict[str, str] = {
    "stt": "STT",
    "ma_to": "MaTo",
    "ten_to": "TenTo",
    "ten_tvv_day_du": "TenTVV_DayDu",
    "xa": "Xa",
    "ma_to_truong": "MaToTruong",
    "ten_to_truong": "Ten_ToTruong",
    "dia_chi": "DiaChi",
    "tk_to_truong": "TK_ToTruong",
    "so_dien_thoai": "SoDienThoai",
    "to_hoi": "ToHoi",
    "tk_to_hoi_xa": "TK_ToHoiXa",
    "to_chuc": "ToChuc",
    "ten_huyen": "Ten_Huyen",
    "tk_huyen": "TK_HUYEN",
    "ten_tinh": "Ten_Tinh",
    "tk_tinh": "TK_TINH",
    "ten_tw": "Ten_TW",
    "tk_tw": "TK_TW",
    "uy_quyen": "uyquyen",
    "ttln_tw": "TTLN_TW",
    "ttln_tinh": "TTLN_Tinh",
}


def now_text() -> str:
    return datetime.now().isoformat(timespec="seconds")


@dataclass(frozen=True, slots=True)
class CreditGroup:
    stt: int = 0
    ma_to: str = ""
    ten_to: str = ""
    ten_tvv_day_du: str = ""
    xa: str = ""
    ma_to_truong: str = ""
    ten_to_truong: str = ""
    dia_chi: str = ""
    tk_to_truong: str = ""
    so_dien_thoai: str = ""
    to_hoi: str = ""
    tk_to_hoi_xa: str = ""
    to_chuc: str = ""
    ten_huyen: str = ""
    tk_huyen: str = ""
    ten_tinh: str = ""
    tk_tinh: str = ""
    ten_tw: str = ""
    tk_tw: str = ""
    uy_quyen: str = ""
    ttln_tw: str = ""
    ttln_tinh: str = ""
    active: bool = True
    created_at: str = ""
    updated_at: str = ""

    @classmethod
    def from_data_tvv_row(cls, row: list[Any] | tuple[Any, ...]) -> "CreditGroup":
        values = list(row) + [""] * (len(DATA_TVV_HEADERS) - len(row))
        return cls(
            stt=int(values[0] or 0),
            ma_to=str(values[1] or "").strip(),
            ten_to=str(values[2] or "").strip(),
            ten_tvv_day_du=str(values[3] or "").strip(),
            xa=str(values[4] or "").strip(),
            ma_to_truong=str(values[5] or "").strip(),
            ten_to_truong=str(values[6] or "").strip(),
            dia_chi=str(values[7] or "").strip(),
            tk_to_truong=str(values[8] or "").strip(),
            so_dien_thoai=str(values[9] or "").strip(),
            to_hoi=str(values[10] or "").strip(),
            tk_to_hoi_xa=str(values[11] or "").strip(),
            to_chuc=str(values[12] or "").strip(),
            ten_huyen=str(values[13] or "").strip(),
            tk_huyen=str(values[14] or "").strip(),
            ten_tinh=str(values[15] or "").strip(),
            tk_tinh=str(values[16] or "").strip(),
            ten_tw=str(values[17] or "").strip(),
            tk_tw=str(values[18] or "").strip(),
            uy_quyen=str(values[19] or "").strip(),
            ttln_tw=str(values[20] or "").strip(),
            ttln_tinh=str(values[21] or "").strip(),
        )

    def to_data_tvv_row(self) -> tuple[Any, ...]:
        return (
            self.stt,
            self.ma_to,
            self.ten_to,
            self.ten_tvv_day_du,
            self.xa,
            self.ma_to_truong,
            self.ten_to_truong,
            self.dia_chi,
            self.tk_to_truong,
            self.so_dien_thoai,
            self.to_hoi,
            self.tk_to_hoi_xa,
            self.to_chuc,
            self.ten_huyen,
            self.tk_huyen,
            self.ten_tinh,
            self.tk_tinh,
            self.ten_tw,
            self.tk_tw,
            self.uy_quyen,
            self.ttln_tw,
            self.ttln_tinh,
        )


@dataclass(frozen=True, slots=True)
class CreditGroupCommissionRate:
    ma_to: str
    base_no_secured_rate: float = 3.0
    base_secured_rate: float = 2.0
    no_secured_to_truong: float = 80.0
    no_secured_cap_xa: float = 13.0
    no_secured_cap_huyen: float = 3.8
    no_secured_cap_tinh: float = 2.5
    no_secured_cap_tw: float = 0.7
    secured_to_truong: float = 90.0
    secured_cap_xa: float = 10.0
    secured_cap_huyen: float = 0.0
    secured_cap_tinh: float = 0.0
    secured_cap_tw: float = 0.0
    created_at: str = ""
    updated_at: str = ""

    @classmethod
    def default_for_group(cls, ma_to: str) -> "CreditGroupCommissionRate":
        return cls(ma_to=ma_to)

    def total_no_secured(self) -> float:
        return (
            self.no_secured_to_truong
            + self.no_secured_cap_xa
            + self.no_secured_cap_huyen
            + self.no_secured_cap_tinh
            + self.no_secured_cap_tw
        )

    def total_secured(self) -> float:
        return (
            self.secured_to_truong
            + self.secured_cap_xa
            + self.secured_cap_huyen
            + self.secured_cap_tinh
            + self.secured_cap_tw
        )

    def validate(self) -> list[str]:
        errors: list[str] = []
        if not self.ma_to.strip():
            errors.append("Mã tổ không được để trống.")

        labels = {
            "base_no_secured_rate": "Tỷ lệ hoa hồng không TSBĐ",
            "base_secured_rate": "Tỷ lệ hoa hồng có TSBĐ",
            "no_secured_to_truong": "Không BĐ - Tổ trưởng",
            "no_secured_cap_xa": "Không BĐ - Cấp xã",
            "no_secured_cap_huyen": "Không BĐ - Cấp huyện",
            "no_secured_cap_tinh": "Không BĐ - Cấp tỉnh",
            "no_secured_cap_tw": "Không BĐ - Cấp TW",
            "secured_to_truong": "Có BĐTS - Tổ trưởng",
            "secured_cap_xa": "Có BĐTS - Cấp xã",
            "secured_cap_huyen": "Có BĐTS - Cấp huyện",
            "secured_cap_tinh": "Có BĐTS - Cấp tỉnh",
            "secured_cap_tw": "Có BĐTS - Cấp TW",
        }
        for field, label in labels.items():
            value = float(getattr(self, field))
            if value < 0:
                errors.append(f"{label} không được âm.")
            if field.startswith("base_") and value > 100:
                errors.append(f"{label} phải từ 0 đến 100%.")

        if not 99.99 <= self.total_no_secured() <= 100.01:
            errors.append("Tổng Hoa hồng không BĐ phải bằng 100%.")
        if not 99.99 <= self.total_secured() <= 100.01:
            errors.append("Tổng Hoa hồng có BĐTS phải bằng 100%.")
        return errors

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


@dataclass(frozen=True, slots=True)
class CreditCommissionRuleSettings:
    """Global commission eligibility settings for the Tổ vay vốn module."""

    secured_base_rate: float = 2.0
    no_secured_base_rate: float = 3.0
    interest_min_1: float = 85.0
    interest_max_1: float = 90.0
    interest_pay_1: float = 50.0
    interest_min_2: float = 90.0
    interest_max_2: float = 95.0
    interest_pay_2: float = 90.0
    interest_min_3: float = 95.0
    interest_pay_3: float = 100.0
    bad_debt_threshold: float = 2.0
    bad_debt_pay: float = 0.0
    updated_at: str = ""

    def validate(self) -> list[str]:
        errors: list[str] = []
        if not 0 <= self.secured_base_rate <= 100:
            errors.append("Tỷ lệ hoa hồng nền có BĐTS phải từ 0 đến 100%.")
        if not 0 <= self.no_secured_base_rate <= 100:
            errors.append("Tỷ lệ hoa hồng nền không BĐ phải từ 0 đến 100%.")
        ranges = (
            (self.interest_min_1, self.interest_max_1, self.interest_pay_1, "Mức 1"),
            (self.interest_min_2, self.interest_max_2, self.interest_pay_2, "Mức 2"),
        )
        for min_value, max_value, pay_value, label in ranges:
            if min_value < 0 or max_value < 0:
                errors.append(f"{label}: tỷ lệ thu lãi không được âm.")
            if min_value >= max_value:
                errors.append(f"{label}: Min thu lãi phải nhỏ hơn Max thu lãi.")
            if not 0 <= pay_value <= 100:
                errors.append(f"{label}: tỷ lệ chi phải từ 0 đến 100%.")
        if self.interest_min_3 < 0:
            errors.append("Mức 3: Min thu lãi không được âm.")
        if not 0 <= self.interest_pay_3 <= 100:
            errors.append("Mức 3: tỷ lệ chi phải từ 0 đến 100%.")
        if self.bad_debt_threshold < 0:
            errors.append("Ngưỡng nợ xấu không được âm.")
        if not 0 <= self.bad_debt_pay <= 100:
            errors.append("Tỷ lệ chi khi vượt ngưỡng nợ xấu phải từ 0 đến 100%.")
        if (
            not errors
            and not (
                self.interest_min_1
                < self.interest_max_1
                <= self.interest_min_2
                < self.interest_max_2
                <= self.interest_min_3
            )
        ):
            errors.append("Các khoảng thu lãi cần tăng dần và không chồng lấn.")
        return errors


@dataclass(frozen=True, slots=True)
class CreditGroupCommissionRule:
    """Optional per-group commission eligibility override."""

    ma_to: str
    use_custom_rule: bool = False
    interest_min_1: float = 85.0
    interest_max_1: float = 90.0
    interest_pay_1: float = 50.0
    interest_min_2: float = 90.0
    interest_max_2: float = 95.0
    interest_pay_2: float = 90.0
    interest_min_3: float = 95.0
    interest_pay_3: float = 100.0
    bad_debt_threshold: float = 2.0
    bad_debt_pay: float = 0.0
    created_at: str = ""
    updated_at: str = ""

    @classmethod
    def default_for_group(cls, ma_to: str) -> "CreditGroupCommissionRule":
        return cls(ma_to=ma_to)

    def as_settings(self) -> CreditCommissionRuleSettings:
        return CreditCommissionRuleSettings(
            interest_min_1=self.interest_min_1,
            interest_max_1=self.interest_max_1,
            interest_pay_1=self.interest_pay_1,
            interest_min_2=self.interest_min_2,
            interest_max_2=self.interest_max_2,
            interest_pay_2=self.interest_pay_2,
            interest_min_3=self.interest_min_3,
            interest_pay_3=self.interest_pay_3,
            bad_debt_threshold=self.bad_debt_threshold,
            bad_debt_pay=self.bad_debt_pay,
            updated_at=self.updated_at,
        )

    @classmethod
    def from_settings(
        cls,
        ma_to: str,
        settings: CreditCommissionRuleSettings,
        *,
        use_custom_rule: bool = False,
    ) -> "CreditGroupCommissionRule":
        return cls(
            ma_to=ma_to,
            use_custom_rule=use_custom_rule,
            interest_min_1=settings.interest_min_1,
            interest_max_1=settings.interest_max_1,
            interest_pay_1=settings.interest_pay_1,
            interest_min_2=settings.interest_min_2,
            interest_max_2=settings.interest_max_2,
            interest_pay_2=settings.interest_pay_2,
            interest_min_3=settings.interest_min_3,
            interest_pay_3=settings.interest_pay_3,
            bad_debt_threshold=settings.bad_debt_threshold,
            bad_debt_pay=settings.bad_debt_pay,
        )

    def validate(self) -> list[str]:
        if not self.ma_to.strip():
            return ["Mã tổ không được để trống."]
        return self.as_settings().validate()
