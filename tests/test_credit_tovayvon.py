from __future__ import annotations

from pathlib import Path
from contextlib import closing
import os
import sqlite3
from tempfile import TemporaryDirectory
import unittest
from unittest.mock import patch

from openpyxl import Workbook, load_workbook

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtWidgets import QApplication, QComboBox, QLineEdit

from agribank_v3.features.credit.tovayvon.models import (
    COMMISSION_EXPORT_HEADERS,
    COMMISSION_RULE_EXPORT_HEADERS,
    DATA_TVV_FIELD_LABELS,
    DATA_TVV_HEADERS,
    DATA_TVV_TEMPLATE_HEADERS,
    CreditGroup,
    CreditGroupCommissionRate,
    CreditCommissionRuleSettings,
)
from agribank_v3.features.credit.tovayvon.excel_templates import (
    DATA_TVV_TEXT_COLUMNS,
    create_data_tvv_template,
)
from agribank_v3.features.credit.tovayvon.repository import (
    CreditGroupRepository,
    CreditGroupRepositoryError,
)
from agribank_v3.features.credit.tovayvon.placeholder_windows import (
    CreditGroupEditDialog,
    FIELD_PLACEHOLDERS,
    get_default_credit_group_info,
    get_suggested_credit_group_default_info,
    load_default_credit_group_info,
    normalize_uy_quyen,
    save_default_credit_group_info,
)


class CreditGroupRepositoryTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temporary_directory = TemporaryDirectory()
        self.database_path = Path(self.temporary_directory.name) / "DuLieuV3.db"
        self.repository = CreditGroupRepository(self.database_path)

    def tearDown(self) -> None:
        self.temporary_directory.cleanup()

    def test_new_group_gets_vba_default_commission_rates(self) -> None:
        self.repository.save_group(
            CreditGroup(stt=1, ma_to="T001", ten_to="Tổ 001")
        )

        rate = self.repository.get_or_create_commission_rate("T001")

        self.assertEqual(rate.no_secured_to_truong, 80.0)
        self.assertEqual(rate.no_secured_cap_xa, 13.0)
        self.assertEqual(rate.no_secured_cap_huyen, 3.8)
        self.assertEqual(rate.no_secured_cap_tinh, 2.5)
        self.assertEqual(rate.no_secured_cap_tw, 0.7)
        self.assertEqual(rate.secured_to_truong, 90.0)
        self.assertEqual(rate.secured_cap_xa, 10.0)
        self.assertEqual(rate.total_no_secured(), 100.0)
        self.assertEqual(rate.total_secured(), 100.0)

    def test_data_tvv_field_labels_complete(self) -> None:
        self.assertEqual(set(DATA_TVV_FIELD_LABELS), set(DATA_TVV_HEADERS))
        self.assertEqual(DATA_TVV_FIELD_LABELS["MaTo"], "Mã số tổ")
        self.assertEqual(DATA_TVV_FIELD_LABELS["TenTo"], "Tên tổ vay vốn")
        self.assertEqual(
            DATA_TVV_FIELD_LABELS["TK_HUYEN"],
            "Tài khoản tổ hội cấp huyện",
        )

    def test_default_info_empty_does_not_crash(self) -> None:
        self.assertEqual(get_default_credit_group_info(), {})

    def test_default_info_settings_save_load(self) -> None:
        saved = save_default_credit_group_info(
            {
                "xa": "Đan Phượng",
                "ten_huyen": "Hội nông dân huyện Lâm Hà",
                "uy_quyen": "1",
                "ma_to": "SHOULD_NOT_SAVE",
            },
            self.database_path,
        )

        loaded = load_default_credit_group_info(self.database_path)

        self.assertEqual(saved["xa"], "Đan Phượng")
        self.assertEqual(loaded["ten_huyen"], "Hội nông dân huyện Lâm Hà")
        self.assertEqual(loaded["uy_quyen"], "Có")
        self.assertNotIn("ma_to", loaded)

    def test_apply_default_info_fills_empty_only(self) -> None:
        app = QApplication.instance() or QApplication([])
        _ = app
        save_default_credit_group_info(
            {
                "xa": "Đan Phượng",
                "dia_chi": "Địa chỉ mặc định",
                "ten_tinh": "Hội nông dân tỉnh Lâm Đồng",
                "uy_quyen": "Có",
            },
            self.database_path,
        )
        dialog = CreditGroupEditDialog(self.repository, mode="add")
        try:
            assert isinstance(dialog.inputs["xa"], QLineEdit)
            dialog.inputs["xa"].setText("Xã đã nhập")
            with patch(
                "agribank_v3.features.credit.tovayvon.placeholder_windows.QMessageBox.information"
            ):
                dialog.apply_default_info()

            self.assertEqual(dialog.inputs["xa"].text(), "Xã đã nhập")
            self.assertEqual(dialog.inputs["dia_chi"].text(), "Địa chỉ mặc định")
            self.assertEqual(dialog.inputs["ten_tinh"].text(), "Hội nông dân tỉnh Lâm Đồng")
            uy_quyen = dialog.inputs["uy_quyen"]
            self.assertIsInstance(uy_quyen, QComboBox)
            assert isinstance(uy_quyen, QComboBox)
            self.assertEqual(uy_quyen.currentText(), "Có")
        finally:
            dialog.close()

    def test_uyquyen_option_normalization(self) -> None:
        for value in ("Có", "Co", "1", "True", "Yes"):
            self.assertEqual(normalize_uy_quyen(value), "Có")
        for value in ("Không", "Khong", "0", "False", "No", ""):
            self.assertEqual(normalize_uy_quyen(value), "Không")

    def test_form_placeholders_mapping(self) -> None:
        self.assertEqual(FIELD_PLACEHOLDERS["ma_to"], "5491LLG202100003")
        self.assertEqual(FIELD_PLACEHOLDERS["ten_to"], "Tổ dịch vụ TD HND Đan Phượng 01")
        self.assertEqual(
            FIELD_PLACEHOLDERS["ttln_tw"],
            "012025/TTLN-HNDVN-AGRIBANK ngày 26/12/2025",
        )
        self.assertEqual(
            get_suggested_credit_group_default_info()["uy_quyen"],
            "Không",
        )

    def test_custom_commission_rate_is_saved_per_ma_to(self) -> None:
        self.repository.save_group(
            CreditGroup(stt=1, ma_to="T001", ten_to="Tổ 001")
        )
        self.repository.save_group(
            CreditGroup(stt=2, ma_to="T002", ten_to="Tổ 002")
        )

        self.repository.save_commission_rate(
            CreditGroupCommissionRate(
                ma_to="T001",
                no_secured_to_truong=70,
                no_secured_cap_xa=15,
                no_secured_cap_huyen=5,
                no_secured_cap_tinh=5,
                no_secured_cap_tw=5,
                secured_to_truong=85,
                secured_cap_xa=10,
                secured_cap_huyen=2,
                secured_cap_tinh=2,
                secured_cap_tw=1,
            )
        )

        custom_rate = self.repository.get_or_create_commission_rate("T001")
        default_rate = self.repository.get_or_create_commission_rate("T002")
        self.assertEqual(custom_rate.no_secured_to_truong, 70.0)
        self.assertEqual(custom_rate.secured_to_truong, 85.0)
        self.assertEqual(default_rate.no_secured_to_truong, 80.0)
        self.assertEqual(default_rate.secured_to_truong, 90.0)

    def test_create_and_update_group_keeps_text_fields(self) -> None:
        self.repository.save_group(
            CreditGroup(
                stt=1,
                ma_to="T001",
                ten_to="Tổ 001",
                tk_to_truong="001234567890",
                so_dien_thoai="0912.345.678",
            )
        )

        self.repository.save_group(
            CreditGroup(
                stt=1,
                ma_to="T001",
                ten_to="Tổ 001 cập nhật",
                tk_to_truong="000000123456",
                so_dien_thoai="0987.000.111",
            )
        )

        group = self.repository.get_group("T001")
        self.assertIsNotNone(group)
        assert group is not None
        self.assertEqual(group.ten_to, "Tổ 001 cập nhật")
        self.assertEqual(group.tk_to_truong, "000000123456")
        self.assertEqual(group.so_dien_thoai, "0987.000.111")
        self.assertEqual(
            self.repository.get_or_create_commission_rate("T001").total_secured(),
            100.0,
        )

    def test_create_group_without_stt_assigns_next(self) -> None:
        self.repository.save_group(CreditGroup(stt=1, ma_to="T001", ten_to="Tổ 001"))
        self.repository.save_group(CreditGroup(stt=2, ma_to="T002", ten_to="Tổ 002"))

        self.repository.save_group(CreditGroup(stt=0, ma_to="T003", ten_to="Tổ 003"))

        groups = self.repository.list_groups()
        self.assertEqual(
            [(group.ma_to, group.stt) for group in groups],
            [("T001", 1), ("T002", 2), ("T003", 3)],
        )

    def test_insert_group_with_duplicate_stt_reorders_sequence(self) -> None:
        self.repository.save_group(CreditGroup(stt=1, ma_to="T001", ten_to="Tổ 001"))
        self.repository.save_group(CreditGroup(stt=2, ma_to="T002", ten_to="Tổ 002"))
        self.repository.save_group(CreditGroup(stt=3, ma_to="T003", ten_to="Tổ 003"))

        self.repository.save_group(CreditGroup(stt=2, ma_to="T004", ten_to="Tổ 004"))

        groups = self.repository.list_groups()
        self.assertEqual(
            [(group.ma_to, group.stt) for group in groups],
            [("T001", 1), ("T004", 2), ("T002", 3), ("T003", 4)],
        )

    def test_update_group_with_duplicate_stt_reorders_sequence(self) -> None:
        self.repository.save_group(CreditGroup(stt=1, ma_to="T001", ten_to="Tổ 001"))
        self.repository.save_group(CreditGroup(stt=2, ma_to="T002", ten_to="Tổ 002"))
        self.repository.save_group(CreditGroup(stt=3, ma_to="T003", ten_to="Tổ 003"))

        self.repository.save_group(
            CreditGroup(stt=1, ma_to="T003", ten_to="Tổ 003 cập nhật")
        )

        groups = self.repository.list_groups()
        self.assertEqual(
            [(group.ma_to, group.stt) for group in groups],
            [("T003", 1), ("T001", 2), ("T002", 3)],
        )
        self.assertEqual(groups[0].ten_to, "Tổ 003 cập nhật")

    def test_resequence_existing_duplicate_stt(self) -> None:
        now = "2026-01-01T00:00:00"
        with closing(sqlite3.connect(self.database_path)) as connection:
            connection.executescript(
                """
                INSERT INTO credit_groups(
                    ma_to, stt, ten_to, created_at, updated_at
                )
                VALUES
                    ('T001', 1, 'Tổ 001', '2026-01-01T00:00:01', '2026-01-01T00:00:01'),
                    ('T002', 1, 'Tổ 002', '2026-01-01T00:00:02', '2026-01-01T00:00:02'),
                    ('T003', 0, 'Tổ 003', '2026-01-01T00:00:03', '2026-01-01T00:00:03')
                """
            )
            connection.commit()

        self.repository.resequence_group_stt()

        groups = self.repository.list_groups()
        self.assertEqual(
            [(group.ma_to, group.stt) for group in groups],
            [("T001", 1), ("T002", 2), ("T003", 3)],
        )

    def test_invalid_commission_total_is_rejected(self) -> None:
        self.repository.save_group(
            CreditGroup(stt=1, ma_to="T001", ten_to="Tổ 001")
        )

        with self.assertRaises(CreditGroupRepositoryError):
            self.repository.save_commission_rate(
                CreditGroupCommissionRate(
                    ma_to="T001",
                    no_secured_to_truong=50,
                    no_secured_cap_xa=10,
                    no_secured_cap_huyen=10,
                    no_secured_cap_tinh=10,
                    no_secured_cap_tw=10,
                )
            )

        rate = self.repository.get_or_create_commission_rate("T001")
        self.assertEqual(rate.no_secured_to_truong, 80.0)

    def test_import_data_tvv_creates_default_commissions(self) -> None:
        source_path = Path(self.temporary_directory.name) / "DaTa_ToVayVon.xlsx"
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "Data_TVV"
        sheet.append(DATA_TVV_HEADERS)
        sheet.append(self._data_tvv_row(1, "T001", "Tổ 001"))
        sheet.append(self._data_tvv_row(2, "T002", "Tổ 002"))
        workbook.save(source_path)
        workbook.close()

        imported_count = self.repository.import_data_tvv(source_path)

        self.assertEqual(imported_count, 2)
        self.assertEqual(
            self.repository.get_or_create_commission_rate("T001").total_secured(),
            100.0,
        )
        self.assertEqual(
            self.repository.get_or_create_commission_rate("T002").total_no_secured(),
            100.0,
        )

    def test_export_data_tvv_preserves_original_22_columns_by_default(self) -> None:
        self.repository.save_group(
            CreditGroup(stt=1, ma_to="T001", ten_to="Tổ 001")
        )
        self.repository.save_commission_rate(
            CreditGroupCommissionRate(
                ma_to="T001",
                no_secured_to_truong=70,
                no_secured_cap_xa=15,
                no_secured_cap_huyen=5,
                no_secured_cap_tinh=5,
                no_secured_cap_tw=5,
                secured_to_truong=85,
                secured_cap_xa=10,
                secured_cap_huyen=2,
                secured_cap_tinh=2,
                secured_cap_tw=1,
            )
        )
        output_path = Path(self.temporary_directory.name) / "Data_TVV.xlsx"

        self.repository.export_data_tvv(output_path)

        workbook = load_workbook(output_path, read_only=True)
        try:
            headers = [cell.value for cell in next(workbook["Data_TVV"].iter_rows(max_row=1))]
        finally:
            workbook.close()
        self.assertEqual(headers, list(DATA_TVV_HEADERS))
        self.assertEqual(len(headers), 22)

    def test_create_data_tvv_template_has_44_columns(self) -> None:
        output_path = Path(self.temporary_directory.name) / "Mau_Data_TVV.xlsx"

        create_data_tvv_template(output_path)

        self.assertTrue(output_path.is_file())
        workbook = load_workbook(output_path)
        try:
            self.assertIn("Data_TVV", workbook.sheetnames)
            self.assertIn("HuongDan", workbook.sheetnames)
            headers = [
                cell.value
                for cell in next(workbook["Data_TVV"].iter_rows(max_row=1))
            ]
        finally:
            workbook.close()
        self.assertEqual(headers, list(DATA_TVV_TEMPLATE_HEADERS))
        self.assertEqual(headers[:22], list(DATA_TVV_HEADERS))
        self.assertEqual(headers[22:34], list(COMMISSION_EXPORT_HEADERS))
        self.assertEqual(headers[34:44], list(COMMISSION_RULE_EXPORT_HEADERS))
        self.assertEqual(len(headers), 44)

    def test_template_text_columns_are_formatted_as_text(self) -> None:
        output_path = Path(self.temporary_directory.name) / "Mau_Data_TVV.xlsx"

        create_data_tvv_template(output_path)

        workbook = load_workbook(output_path)
        try:
            sheet = workbook["Data_TVV"]
            header_to_column = {
                sheet.cell(row=1, column=column_index).value: column_index
                for column_index in range(1, len(DATA_TVV_TEMPLATE_HEADERS) + 1)
            }
            for header in DATA_TVV_TEXT_COLUMNS:
                column_index = header_to_column[header]
                self.assertEqual(
                    sheet.cell(row=2, column=column_index).number_format,
                    "@",
                )
                self.assertEqual(
                    sheet.cell(row=20, column=column_index).number_format,
                    "@",
                )
        finally:
            workbook.close()

    def test_import_created_template_no_crash(self) -> None:
        output_path = Path(self.temporary_directory.name) / "Mau_Data_TVV.xlsx"
        create_data_tvv_template(output_path)

        imported_count = self.repository.import_data_tvv(output_path)

        self.assertEqual(imported_count, 1)
        self.assertIsNotNone(self.repository.get_group("TVV001"))
        self.assertEqual(
            self.repository.get_or_create_commission_rate("TVV001").total_secured(),
            100.0,
        )

    def test_import_old_22_column_template_still_works(self) -> None:
        source_path = Path(self.temporary_directory.name) / "Old_Data_TVV.xlsx"
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "Data_TVV"
        sheet.append(DATA_TVV_HEADERS)
        sheet.append(self._data_tvv_row(1, "T001", "Tổ 001"))
        workbook.save(source_path)
        workbook.close()

        imported_count = self.repository.import_data_tvv(source_path)

        self.assertEqual(imported_count, 1)
        self.assertEqual(
            self.repository.get_or_create_commission_rate("T001").no_secured_to_truong,
            80.0,
        )

    def test_import_new_44_column_template_with_commission(self) -> None:
        source_path = Path(self.temporary_directory.name) / "New_Data_TVV.xlsx"
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "Data_TVV"
        sheet.append(DATA_TVV_TEMPLATE_HEADERS)
        sheet.append(
            self._data_tvv_row(1, "T001", "Tổ 001")
            + [70, 15, 5, 5, 5, 100, 85, 10, 2, 2, 1, 100]
            + [80, 88, 40, 88, 96, 85, 96, 100, 3, 0]
        )
        workbook.save(source_path)
        workbook.close()

        imported_count = self.repository.import_data_tvv(
            source_path,
            update_commission_rules=True,
        )

        self.assertEqual(imported_count, 1)
        rate = self.repository.get_or_create_commission_rate("T001")
        self.assertEqual(rate.no_secured_to_truong, 70.0)
        self.assertEqual(rate.secured_to_truong, 85.0)
        settings = self.repository.get_commission_rule_settings()
        self.assertEqual(settings.interest_min_1, 80.0)
        self.assertEqual(settings.bad_debt_threshold, 3.0)

    def test_invalid_commission_total_in_template_is_reported(self) -> None:
        source_path = Path(self.temporary_directory.name) / "Invalid_Data_TVV.xlsx"
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "Data_TVV"
        sheet.append(DATA_TVV_TEMPLATE_HEADERS)
        sheet.append(
            self._data_tvv_row(1, "T001", "Tổ 001")
            + [50, 10, 10, 10, 10, 90, 90, 10, 0, 0, 0, 100]
            + [85, 90, 50, 90, 95, 90, 95, 100, 2, 0]
        )
        workbook.save(source_path)
        workbook.close()

        with self.assertRaises(CreditGroupRepositoryError):
            self.repository.import_data_tvv(source_path)

    def test_export_with_commission_has_44_columns(self) -> None:
        self.repository.save_group(
            CreditGroup(stt=1, ma_to="T001", ten_to="Tổ 001")
        )
        output_path = Path(self.temporary_directory.name) / "Data_TVV_44.xlsx"

        self.repository.export_data_tvv(output_path, include_commission=True)

        workbook = load_workbook(output_path, read_only=True)
        try:
            headers = [
                cell.value
                for cell in next(workbook["Data_TVV"].iter_rows(max_row=1))
            ]
        finally:
            workbook.close()
        self.assertEqual(headers, list(DATA_TVV_TEMPLATE_HEADERS))
        self.assertEqual(len(headers), 44)

    def test_commission_rule_settings_have_vba_like_defaults(self) -> None:
        settings = self.repository.get_commission_rule_settings()

        self.assertEqual(settings.interest_min_1, 85.0)
        self.assertEqual(settings.interest_max_1, 90.0)
        self.assertEqual(settings.interest_pay_1, 50.0)
        self.assertEqual(settings.interest_min_2, 90.0)
        self.assertEqual(settings.interest_max_2, 95.0)
        self.assertEqual(settings.interest_pay_2, 90.0)
        self.assertEqual(settings.interest_min_3, 95.0)
        self.assertEqual(settings.interest_pay_3, 100.0)
        self.assertEqual(settings.bad_debt_threshold, 2.0)
        self.assertEqual(settings.bad_debt_pay, 0.0)

    def test_commission_rule_settings_are_saved_and_validated(self) -> None:
        self.repository.save_commission_rule_settings(
            CreditCommissionRuleSettings(
                interest_min_1=80,
                interest_max_1=88,
                interest_pay_1=40,
                interest_min_2=88,
                interest_max_2=96,
                interest_pay_2=85,
                interest_min_3=96,
                interest_pay_3=100,
                bad_debt_threshold=3,
                bad_debt_pay=0,
            )
        )

        settings = self.repository.get_commission_rule_settings()
        self.assertEqual(settings.interest_min_1, 80.0)
        self.assertEqual(settings.interest_pay_2, 85.0)
        self.assertEqual(settings.bad_debt_threshold, 3.0)

        with self.assertRaises(CreditGroupRepositoryError):
            self.repository.save_commission_rule_settings(
                CreditCommissionRuleSettings(
                    interest_min_1=90,
                    interest_max_1=85,
                    interest_pay_1=50,
                )
            )

    @staticmethod
    def _data_tvv_row(stt: int, ma_to: str, ten_to: str) -> list[str | int]:
        return [
            stt,
            ma_to,
            ten_to,
            f"{ten_to} đầy đủ",
            "Xã A",
            f"TT{stt:03d}",
            "Nguyễn Văn A",
            "Ấp 1",
            "123456789",
            "0900000000",
            "Hội nông dân",
            "987654321",
            "Tổ chức A",
            "Huyện A",
            "111",
            "Tỉnh A",
            "222",
            "Trung ương",
            "333",
            "Có",
            "TW",
            "Tỉnh",
        ]


if __name__ == "__main__":
    unittest.main()
