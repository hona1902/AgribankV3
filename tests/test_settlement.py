from __future__ import annotations

import unittest
from datetime import datetime
from decimal import Decimal
from io import BytesIO
from pathlib import Path

from openpyxl import Workbook, load_workbook

from agribank_v3.settings import BranchProfile
from agribank_v3.settlement import (
    ACTIVE_SETTLEMENT_SPECS,
    LEGACY_SETTLEMENT_SPECS,
    SETTLEMENT_SPECS,
    SettlementEngine,
    SettlementError,
    SettlementOptions,
    SettlementRequest,
)
from agribank_v3.settlement.transforms import (
    decode_telex,
    excel_column_name,
    is_branch_code,
    is_code_like,
    normalize_customer_id,
    parse_vietnamese_amount,
    parse_yyyymmdd,
    vietnamese_report_date,
)
from agribank_v3.settlement.processors import (
    Mau05Processor,
    Mau06Processor,
    Mau1516Processor,
    Mau18Processor,
    Mau20aProcessor,
    Mau30Processor,
)
from agribank_v3.settlement.processors.mau05 import CollateralRecord


class SettlementRegistryTests(unittest.TestCase):
    def test_registry_contains_every_active_ribbon_entry(self) -> None:
        self.assertEqual(len(ACTIVE_SETTLEMENT_SPECS), 29)
        self.assertEqual(len(LEGACY_SETTLEMENT_SPECS), 7)
        self.assertEqual(len(SETTLEMENT_SPECS), 36)
        self.assertEqual(
            len(SETTLEMENT_SPECS),
            len(set(SETTLEMENT_SPECS)),
        )

    def test_shared_report_families_are_explicit(self) -> None:
        self.assertEqual(
            SETTLEMENT_SPECS["accounting.13"].processor_family,
            SETTLEMENT_SPECS["accounting.14"].processor_family,
        )
        self.assertEqual(
            SETTLEMENT_SPECS["credit.15a"].processor_family,
            SETTLEMENT_SPECS["credit.16"].processor_family,
        )


class SettlementEngineTests(unittest.TestCase):
    def test_requires_database_branch_code(self) -> None:
        request = SettlementRequest(
            spec=SETTLEMENT_SPECS["credit.05"],
            profile=BranchProfile(),
        )
        with self.assertRaisesRegex(SettlementError, "mã chi nhánh"):
            SettlementEngine().execute(request)

    def test_reports_unmigrated_processor_without_using_vba(self) -> None:
        request = SettlementRequest(
            spec=SETTLEMENT_SPECS["credit.05"],
            profile=BranchProfile(branch_code="5491"),
        )
        with self.assertRaises(SettlementError) as raised:
            SettlementEngine().execute(request)
        self.assertEqual(raised.exception.code, "processor_not_migrated")


class SettlementTransformTests(unittest.TestCase):
    def test_excel_column_name_supports_full_xlsx_range(self) -> None:
        self.assertEqual(excel_column_name(1), "A")
        self.assertEqual(excel_column_name(26), "Z")
        self.assertEqual(excel_column_name(27), "AA")
        self.assertEqual(excel_column_name(16_384), "XFD")

    def test_parses_ipcas_date_without_locale_dependency(self) -> None:
        self.assertEqual(str(parse_yyyymmdd("20251231")), "2025-12-31")
        self.assertEqual(
            vietnamese_report_date("20251231"),
            "Ngày 31 tháng 12 năm 2025",
        )
        self.assertIsNone(parse_yyyymmdd("20250230"))

    def test_normalizes_customer_id_from_vba_text_values(self) -> None:
        self.assertEqual(
            normalize_customer_id("'000123", "5491", include_branch=True),
            "5491000123",
        )
        self.assertEqual(
            normalize_customer_id("'000123", "5491", include_branch=False),
            "000123",
        )

    def test_parses_vietnamese_amount(self) -> None:
        self.assertEqual(
            parse_vietnamese_amount("1.234.567,89"),
            parse_vietnamese_amount(1234567.89),
        )
        self.assertEqual(parse_vietnamese_amount("-"), 0)

    def test_recognizes_balance_codes(self) -> None:
        self.assertTrue(is_branch_code(" 5491 "))
        self.assertFalse(is_branch_code("Tổng cộng"))
        self.assertTrue(is_code_like("211101_VND"))
        self.assertFalse(is_code_like("Cộng chi nhánh"))

    def test_decodes_legacy_telex_labels(self) -> None:
        self.assertEqual(
            decode_telex("Toorng howjp quyeest toasn"),
            "Tổng hợp quyết toán",
        )


class SettlementFixtureParityTests(unittest.TestCase):
    FIXTURE_DIR = (
        Path(__file__).resolve().parents[1] / "DuLieuTEST" / "SoSanh"
    )
    PROFILE = BranchProfile(
        branch_code="5491",
        branch_name="Chi nhánh Lộc Phát Lâm Đồng",
        report_location="Tân Hà",
    )

    @staticmethod
    def _normalized(value: object) -> object:
        if value is None:
            return ""
        if isinstance(value, Decimal):
            return int(value) if value == value.to_integral_value() else float(value)
        return value.date() if isinstance(value, datetime) else value

    def test_mau05_business_region_matches_vba_fixture(self) -> None:
        source = self.FIXTURE_DIR / "5491_rt05.csv"
        expected = load_workbook(
            self.FIXTURE_DIR / "5491QT05.xlsx",
            data_only=False,
        )["05"]
        fixture_uses_branch_code = str(expected["A16"].value or "").startswith(
            self.PROFILE.branch_code
        )
        fixture_uses_customer_totals = any(
            str(expected.cell(row, 1).value or "").startswith(
                "   Cộng khách hàng:"
            )
            for row in range(12, expected.max_row + 1)
        )
        processor = Mau05Processor()
        generated = None
        records = []
        for use_owner in (False, True):
            request = SettlementRequest(
                SETTLEMENT_SPECS["credit.05"],
                self.PROFILE,
                options=SettlementOptions(
                    include_branch_in_customer_id=fixture_uses_branch_code,
                    include_customer_totals=fixture_uses_customer_totals,
                    use_collateral_owner_for_guarantee=use_owner,
                ),
                source_paths=(source,),
            )
            records, report_date = processor.read_source(request, source)
            generated_workbook = processor.build_workbook(
                request, records, report_date
            )
            buffer = BytesIO()
            generated_workbook.save(buffer)
            buffer.seek(0)
            candidate = load_workbook(buffer, data_only=False)["05"]
            if self._normalized(candidate["A1533"].value) == self._normalized(
                expected["A1533"].value
            ):
                generated = candidate
                break
        self.assertIsNotNone(generated)

        self.assertEqual(len(records), 1837)
        for row in range(12, 1861):
            for column in range(1, 19):
                self.assertEqual(
                    self._normalized(generated.cell(row, column).value),
                    self._normalized(expected.cell(row, column).value),
                    f"Mismatch at {generated.cell(row, column).coordinate}",
                )

    def test_mau06_business_region_matches_vba_fixture(self) -> None:
        csv_source = self.FIXTURE_DIR / "5491_rt05.csv"
        request05 = SettlementRequest(
            SETTLEMENT_SPECS["credit.05"],
            self.PROFILE,
            source_paths=(csv_source,),
        )
        processor05 = Mau05Processor()
        records, report_date = processor05.read_source(request05, csv_source)
        source_sheet = processor05.build_workbook(
            request05, records, report_date
        )["05"]
        request06 = SettlementRequest(
            SETTLEMENT_SPECS["credit.06"],
            self.PROFILE,
            source_paths=(self.FIXTURE_DIR / "5491QT05.xlsx",),
        )
        generated = Mau06Processor().build_workbook(
            request06, source_sheet
        )["06"]
        expected = load_workbook(
            self.FIXTURE_DIR / "5491QT06.xlsx",
            data_only=False,
        )["06"]

        for row in range(12, 25):
            for column in range(1, 12):
                self.assertEqual(
                    generated.cell(row, column).value,
                    expected.cell(row, column).value,
                    f"Mismatch at {generated.cell(row, column).coordinate}",
                )
        self.assertEqual(generated["A6"].value, "Ngày 31 tháng 12 năm 2025")

    def test_mau05_header_and_print_layout_are_ready_to_print(self) -> None:
        source = self.FIXTURE_DIR / "5491_rt05.csv"
        request = SettlementRequest(
            SETTLEMENT_SPECS["credit.05"],
            self.PROFILE,
            source_paths=(source,),
        )
        processor = Mau05Processor()
        records, report_date = processor.read_source(request, source)
        sheet = processor.build_workbook(request, records, report_date)["05"]

        merged_ranges = {str(merged) for merged in sheet.merged_cells.ranges}
        self.assertTrue({"A1:D1", "A2:D2", "A3:D3", "A4:D4"} <= merged_ranges)
        for address in ("A1", "A2", "A3", "A4"):
            self.assertEqual(sheet[address].alignment.horizontal, "center")
            self.assertTrue(sheet[address].font.bold)
        self.assertIsNone(sheet["A2"].border.bottom.style)
        self.assertEqual(sheet["R7"].alignment.horizontal, "right")
        self.assertTrue(sheet["R7"].font.italic)
        total_row = next(
            row
            for row in range(12, sheet.max_row + 1)
            if sheet.cell(row, 1).value == "Cộng II:"
        )
        self.assertEqual(sheet.cell(total_row, 10).number_format, "0")
        signature_row = next(
            row
            for row in range(1, sheet.max_row + 1)
            if sheet.cell(row, 1).value == "LẬP BIỂU"
        )
        for column in (1, 4, 7, 12):
            self.assertTrue(sheet.cell(signature_row, column).font.bold)
            self.assertTrue(sheet.cell(signature_row + 1, column).font.italic)
        self.assertTrue(sheet.cell(signature_row - 1, 12).font.italic)
        self.assertGreaterEqual(sheet.column_dimensions["A"].width, 26)
        self.assertGreaterEqual(sheet.column_dimensions["E"].width, 17)
        self.assertGreaterEqual(sheet.column_dimensions["H"].width, 17)
        self.assertEqual(sheet.page_setup.paperSize, 9)
        self.assertEqual(sheet.page_setup.orientation, "landscape")
        self.assertEqual(sheet.page_setup.fitToWidth, 1)
        self.assertEqual(sheet.page_setup.fitToHeight, 0)
        self.assertTrue(sheet.sheet_properties.pageSetUpPr.fitToPage)
        self.assertEqual(sheet.print_title_rows, "$8:$10")
        self.assertEqual(str(sheet.print_area), f"'05'!$A$1:$R${sheet.max_row}")
        self.assertTrue(sheet.HeaderFooter.differentFirst)
        self.assertEqual(sheet.firstHeader.right.text, "")
        self.assertEqual(sheet.oddHeader.center.text, "&P/&N")

    def test_mau06_print_layout_is_ready_to_print(self) -> None:
        csv_source = self.FIXTURE_DIR / "5491_rt05.csv"
        request05 = SettlementRequest(
            SETTLEMENT_SPECS["credit.05"],
            self.PROFILE,
            source_paths=(csv_source,),
        )
        processor05 = Mau05Processor()
        records, report_date = processor05.read_source(request05, csv_source)
        source_sheet = processor05.build_workbook(
            request05, records, report_date
        )["05"]
        request06 = SettlementRequest(
            SETTLEMENT_SPECS["credit.06"],
            self.PROFILE,
            source_paths=(self.FIXTURE_DIR / "5491QT05.xlsx",),
        )
        sheet = Mau06Processor().build_workbook(request06, source_sheet)["06"]

        self.assertEqual(sheet["K7"].alignment.horizontal, "right")
        self.assertTrue(sheet["K7"].font.italic)
        for address in ("A26", "B26", "E26", "H26"):
            self.assertTrue(sheet[address].font.bold)
            self.assertEqual(sheet[address].alignment.horizontal, "center")
        for address in ("A27", "B27", "E27", "H27"):
            self.assertTrue(sheet[address].font.italic)
            self.assertEqual(sheet[address].alignment.horizontal, "center")
        self.assertEqual(sheet.page_setup.paperSize, 9)
        self.assertEqual(sheet.page_setup.orientation, "landscape")
        self.assertEqual(sheet.page_setup.fitToWidth, 1)
        self.assertEqual(sheet.page_setup.fitToHeight, 0)
        self.assertTrue(sheet.sheet_properties.pageSetUpPr.fitToPage)
        self.assertEqual(str(sheet.print_area), "'06'!$A$1:$K$40")
        self.assertTrue(sheet.HeaderFooter.differentFirst)
        self.assertEqual(sheet.firstHeader.right.text, "")
        self.assertEqual(sheet.oddHeader.center.text, "&P/&N")

    def test_mau1516_processors_build_expected_report_shapes(self) -> None:
        fixture_dir = Path(__file__).resolve().parents[1] / "DuLieuTEST" / "5491"
        cases = (
            ("credit.15a", "5491_rt15a_20251231.csv", "15a", 20, True),
            ("credit.15b", "5491_rt15b_20251231.csv", "15b", 19, True),
            ("credit.16", "5491_rt16_20251231.csv", "16", 17, False),
        )
        for spec_key, file_name, sheet_name, last_column, has_control in cases:
            with self.subTest(spec_key=spec_key):
                source = fixture_dir / file_name
                request = SettlementRequest(
                    SETTLEMENT_SPECS[spec_key],
                    self.PROFILE,
                    options=SettlementOptions(
                        include_branch_in_customer_id=True,
                        create_control_sheet=True,
                    ),
                    source_paths=(source,),
                )
                processor = Mau1516Processor()
                records, report_date = processor.read_source(request, source)
                workbook = processor.build_workbook(request, records, report_date)
                sheet = workbook[sheet_name]

                self.assertEqual(sheet.max_column, last_column)
                self.assertEqual(str(sheet.print_area), f"'{sheet_name}'!$A$1:${excel_column_name(last_column)}${sheet.max_row}")
                self.assertEqual(sheet.print_title_rows, "$9:$10")
                self.assertEqual(sheet.oddHeader.center.text, "&P/&N")
                self.assertEqual(sheet.cell(10, last_column).value, last_column)
                self.assertTrue(sheet["A1"].font.bold)
                if has_control:
                    self.assertIn("SoLieuTongHop", workbook.sheetnames)
                else:
                    self.assertNotIn("SoLieuTongHop", workbook.sheetnames)
                if records:
                    self.assertTrue(str(sheet["B11"].value).startswith("5491"))
                if sheet_name == "16":
                    self.assertTrue(
                        any(
                            str(sheet.cell(row, 2).value or "").startswith("Cộng KH:")
                            for row in range(11, sheet.max_row + 1)
                        )
                    )

    def test_mau1516_can_keep_unused_source_columns_off_print_area(self) -> None:
        source = (
            Path(__file__).resolve().parents[1]
            / "DuLieuTEST"
            / "5491"
            / "5491_rt15a_20251231.csv"
        )
        request = SettlementRequest(
            SETTLEMENT_SPECS["credit.15a"],
            self.PROFILE,
            options=SettlementOptions(remove_unused_columns=False),
            source_paths=(source,),
        )
        processor = Mau1516Processor()
        records, report_date = processor.read_source(request, source)
        sheet = processor.build_workbook(request, records, report_date)["15a"]

        self.assertGreater(sheet.max_column, 20)
        self.assertEqual(sheet.cell(10, 21).value, "NGAY")
        self.assertEqual(str(sheet.print_area), f"'15a'!$A$1:$T${sheet.max_row}")

    def test_mau18_matches_vba_fixture_values(self) -> None:
        source = self.FIXTURE_DIR / "5400_rt18.csv"
        expected = load_workbook(
            self.FIXTURE_DIR / "5491QT18.xlsx",
            data_only=False,
        )
        request = SettlementRequest(
            SETTLEMENT_SPECS["credit.18"],
            BranchProfile(
                branch_code="5491",
                branch_name="Chi nhánh Lộc Phát Lâm Đồng",
                report_location="Tân Hà",
                report_preparer="Nam",
            ),
            options=SettlementOptions(
                include_branch_in_customer_id=False,
                create_control_sheet=True,
            ),
            source_paths=(source,),
        )
        processor = Mau18Processor()
        records, report_date = processor.read_source(request, source)
        generated = processor.build_workbook(request, records, report_date)

        self.assertEqual(len(records), 288)
        for sheet_name in ("18", "SoLieuTongHop"):
            with self.subTest(sheet=sheet_name):
                actual_sheet = generated[sheet_name]
                expected_sheet = expected[sheet_name]
                self.assertEqual(actual_sheet.max_row, expected_sheet.max_row)
                self.assertEqual(actual_sheet.max_column, expected_sheet.max_column)
                for row in range(1, expected_sheet.max_row + 1):
                    for column in range(1, expected_sheet.max_column + 1):
                        if sheet_name == "18" and row == 21 and column == 13:
                            self.assertEqual(
                                self._normalized(actual_sheet.cell(row, column).value),
                                719240.85,
                            )
                            self.assertEqual(
                                actual_sheet.cell(row, column).number_format,
                                "#,##0",
                            )
                            continue
                        self.assertEqual(
                            self._normalized(actual_sheet.cell(row, column).value),
                            self._normalized(expected_sheet.cell(row, column).value),
                            f"Mismatch at {sheet_name}!{actual_sheet.cell(row, column).coordinate}",
                        )

    def test_mau20a_matches_vba_xls_transform_shape(self) -> None:
        source = self.FIXTURE_DIR / "5491_rt20a.XLS"
        if not source.exists():
            self.skipTest("Mẫu 20a fixture is not available.")
        request = SettlementRequest(
            SETTLEMENT_SPECS["credit.20a"],
            BranchProfile(
                branch_code="5491",
                branch_name="Chi nhánh Lộc Phát Lâm Đồng",
                reporting_branch_name="Chi nhánh Lộc Phát Lâm Đồng",
                report_location="Tân Hà",
            ),
            options=SettlementOptions(),
            source_paths=(source,),
        )
        processor = Mau20aProcessor()
        records = processor.read_source(request, source)
        sheet = processor.build_workbook(request, records, source)["20a"]

        self.assertEqual(len(records), 3)
        self.assertEqual(sheet.max_column, 16)
        self.assertEqual(sheet["A9"].value, "STT")
        self.assertEqual(sheet["A11"].value, 1)
        self.assertEqual(sheet["A12"].value, 1)
        self.assertEqual(sheet["D12"].value, "5491147744864")
        self.assertEqual(sheet["F12"].value, "5491LAV201703312")
        self.assertEqual(sheet["G12"].value, "5491LDS201705592")
        self.assertEqual(sheet["I12"].value, "")
        self.assertEqual(sheet["J12"].value, 1600000)
        self.assertEqual(processor._principal_account("Vay ngắn hạn (TK 211)"), "971101")
        self.assertEqual(processor._principal_account("Vay trung hạn (TK 212)"), "791102")
        self.assertEqual(processor._principal_account("Nguồn vốn (TK 252101)"), "791101")
        self.assertEqual(processor._principal_account("Nguồn vốn (TK 252102)"), "791102")
        self.assertEqual(processor._principal_account("Nguồn vốn (TK 252103)"), "791103")
        self.assertEqual(processor._principal_account("Nguồn vốn (TK 271101)"), "791101")
        self.assertEqual(processor._principal_account("Nguồn vốn (TK 271102)"), "791102")
        self.assertEqual(processor._principal_account("Nguồn vốn (TK 271103)"), "791103")
        self.assertEqual(sheet["A15"].value, "TỔNG CỘNG:")
        self.assertEqual(sheet["J15"].value, "=SUM(J12:J14)")
        self.assertEqual(sheet["N18"].value, "Tân Hà, Ngày 31 tháng 12 năm 2025")
        self.assertTrue(sheet["A19"].font.bold)
        self.assertTrue(sheet["A20"].font.italic)
        self.assertEqual(sheet["A20"].alignment.horizontal, "center")
        self.assertFalse(sheet["B29"].alignment.wrap_text)
        self.assertEqual(sheet["B29"].alignment.horizontal, "left")
        self.assertEqual(sheet.print_area, "'20a'!$A$1:$P$32")
        self.assertEqual(sheet.print_title_rows, "$11:$11")

    def test_mau05_control_sheet_has_table_borders(self) -> None:
        workbook = Workbook()
        workbook.remove(workbook.active)
        records = [
            CollateralRecord(
                sort_customer_id="1",
                customer_id="1",
                customer_name="A",
                contract_number="HD1",
                contract_date=None,
                total=100,
                real_estate=100,
                movable_property=0,
                valuable_papers=0,
                other_assets=0,
                legal_valid="X",
                legal_other="",
                dossier_complete="X",
                dossier_in_progress="",
                dossier_impossible="",
                dossier_other="",
                marketable="X",
                less_marketable="",
                not_marketable="",
                account="994001",
                collateral_type="real_estate",
            )
        ]
        Mau05Processor._write_control_sheet(workbook, records)
        sheet = workbook["SoLieuTongHop"]

        self.assertEqual(sheet["A2"].border.left.style, "thin")
        self.assertEqual(sheet["G4"].border.right.style, "thin")
        self.assertTrue(sheet["A2"].font.bold)
        self.assertTrue(sheet["A4"].font.bold)

    def test_mau30_builds_from_selected_qt_workbook(self) -> None:
        source = self.FIXTURE_DIR / "DaTest" / "5491QT15a.xlsx"
        if not source.exists():
            self.skipTest("Mẫu 30 source fixture is not available.")
        request = SettlementRequest(
            SETTLEMENT_SPECS["credit.30a"],
            BranchProfile(
                branch_code="5491",
                branch_name="Chi nhánh Lộc Phát Lâm Đồng",
                reporting_branch_name="Chi nhánh Lộc Phát Lâm Đồng",
                parent_branch_code="5400",
                report_location="Tân Hà",
            ),
            options=SettlementOptions(
                source_report_code="15a",
                include_accrual_accounts=True,
            ),
            source_paths=(source,),
        )
        workbook = load_workbook(source, data_only=False)
        values = load_workbook(source, data_only=True)
        processor = Mau30Processor()
        balance = processor._read_balance(self.FIXTURE_DIR / "CanDoiNam-30062026.XLS")
        sheet = processor.build_sheet(request, workbook, values, "15a", balance)

        self.assertEqual(sheet.title, "Mau30QT-15a")
        self.assertEqual(sheet["A5"].value, "TỔNG HỢP SỐ LIỆU TOÀN CHI NHÁNH MẪU BIỂU QUYẾT TOÁN SỐ 15a")
        self.assertEqual(sheet["B12"].value, "5400")
        self.assertEqual(sheet["C12"].value, "5491")
        self.assertEqual(sheet["D12"].value, "211101")
        self.assertEqual(sheet["E12"].value, 1959747671000)
        self.assertEqual(sheet["F12"].value, 2169363622401)
        self.assertEqual(sheet["G12"].value, "=F12-E12")
        self.assertTrue(sheet["A1"].font.bold)
        self.assertEqual(sheet["A1"].alignment.horizontal, "center")
        self.assertTrue(sheet["A5"].font.bold)
        self.assertEqual(sheet["A5"].alignment.horizontal, "center")
        self.assertEqual(sheet["G27"].value, "Tân Hà, Ngày 31 tháng 12 năm 2025")
        self.assertEqual(sheet["G27"].alignment.horizontal, "center")
        self.assertTrue(sheet["G27"].alignment.shrink_to_fit)
        self.assertTrue(sheet["G27"].font.italic)
        self.assertEqual(sheet.row_dimensions[27].height, 21)
        self.assertEqual(sheet.print_title_rows, "$11:$11")
        self.assertTrue(str(sheet.print_area).startswith("'Mau30QT-15a'!$A$1:$H$"))

    def test_mau30_builds_20a_without_control_sheet(self) -> None:
        workbook = Workbook()
        source = workbook.active
        source.title = "20a"
        source["A6"] = "BÁO CÁO NỢ ĐƯỢC XỬ LÝ BẰNG NGUỒN DỰ PHÒNG"
        source["A7"] = "Ngày 31 tháng 12 năm 2025"
        source["I12"] = "971101"
        source["J12"] = 1_000_000
        source["K12"] = "971201"
        source["L12"] = 200_000
        source["I13"] = "971101"
        source["J13"] = 500_000
        source["K13"] = "971202"
        source["L13"] = 300_000
        source["A14"] = "TỔNG CỘNG:"
        source["J14"] = 1_500_000
        source["L14"] = 500_000
        source["I17"] = "TRƯỞNG PHÒNG KẾ TOÁN"
        source["K18"] = "(Ký, ghi rõ họ tên)"
        request = SettlementRequest(
            SETTLEMENT_SPECS["credit.30a"],
            BranchProfile(
                branch_code="5491",
                branch_name="Chi nhánh Lộc Phát Lâm Đồng",
                reporting_branch_name="Chi nhánh Lộc Phát Lâm Đồng",
                parent_branch_code="5400",
                report_location="Tân Hà",
            ),
            options=SettlementOptions(source_report_code="20a"),
            source_paths=(Path("5491QT20a.xlsx"),),
        )

        sheet = Mau30Processor().build_sheet(
            request,
            workbook,
            workbook,
            "20a",
        )

        self.assertNotIn("SoLieuTongHop", workbook.sheetnames)
        self.assertEqual(sheet["D12"].value, "971101")
        self.assertEqual(sheet["F12"].value, 1_500_000)
        self.assertEqual(sheet["D13"].value, "971201")
        self.assertEqual(sheet["F13"].value, 200_000)
        self.assertEqual(sheet["D14"].value, "971202")
        self.assertEqual(sheet["F14"].value, 300_000)

    def test_mau05_can_keep_guarantee_borrower_customer(self) -> None:
        source = self.FIXTURE_DIR / "5491_rt05.csv"
        request = SettlementRequest(
            SETTLEMENT_SPECS["credit.05"],
            self.PROFILE,
            options=SettlementOptions(
                use_collateral_owner_for_guarantee=False,
            ),
            source_paths=(source,),
        )
        records, _ = Mau05Processor().read_source(request, source)

        self.assertTrue(
            any(
                record.collateral_type.casefold() == "bao lanh"
                and record.customer_name == "Cty TNHH Phú Cường"
                for record in records
            )
        )

    def test_mau05_can_insert_customer_total_rows(self) -> None:
        source = self.FIXTURE_DIR / "5491_rt05.csv"
        request = SettlementRequest(
            SETTLEMENT_SPECS["credit.05"],
            self.PROFILE,
            options=SettlementOptions(include_customer_totals=True),
            source_paths=(source,),
        )
        processor = Mau05Processor()
        records, report_date = processor.read_source(request, source)
        sheet = processor.build_workbook(request, records, report_date)["05"]

        self.assertTrue(
            any(
                str(sheet.cell(row, 1).value or "").startswith(
                    "   Cộng khách hàng:"
                )
                for row in range(12, sheet.max_row + 1)
            )
        )

    def test_mau05_can_bold_customer_total_rows(self) -> None:
        source = self.FIXTURE_DIR / "5491_rt05.csv"
        request = SettlementRequest(
            SETTLEMENT_SPECS["credit.05"],
            self.PROFILE,
            options=SettlementOptions(
                include_customer_totals=True,
                bold_customer_rows=True,
            ),
            source_paths=(source,),
        )
        processor = Mau05Processor()
        records, report_date = processor.read_source(request, source)
        sheet = processor.build_workbook(request, records, report_date)["05"]
        total_row = next(
            row
            for row in range(12, sheet.max_row + 1)
            if str(sheet.cell(row, 1).value or "").startswith(
                "   Cộng khách hàng:"
            )
        )

        self.assertTrue(sheet.cell(total_row, 1).font.bold)
        self.assertTrue(sheet.cell(total_row, 5).font.bold)

    def test_mau05_can_keep_unused_source_columns_off_print_area(self) -> None:
        source = self.FIXTURE_DIR / "5491_rt05.csv"
        request = SettlementRequest(
            SETTLEMENT_SPECS["credit.05"],
            self.PROFILE,
            options=SettlementOptions(
                include_customer_totals=True,
                remove_unused_columns=False,
            ),
            source_paths=(source,),
        )
        processor = Mau05Processor()
        records, report_date = processor.read_source(request, source)
        sheet = processor.build_workbook(request, records, report_date)["05"]

        self.assertGreater(sheet.max_column, 18)
        self.assertEqual(sheet.cell(11, 19).value, "NGAY")
        self.assertEqual(str(sheet.print_area), f"'05'!$A$1:$R${sheet.max_row}")


if __name__ == "__main__":
    unittest.main()
