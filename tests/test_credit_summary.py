from __future__ import annotations

from contextlib import closing
from datetime import date
import os
from pathlib import Path
import sqlite3
from tempfile import TemporaryDirectory
import unittest
from unittest.mock import patch
import zipfile

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from openpyxl import load_workbook
from PySide6.QtCore import Qt
from PySide6.QtWidgets import QApplication, QComboBox, QPushButton, QSizePolicy, QStyleOptionViewItem

from agribank_v3.features.credit.summary.dashboard_repository import NimDashboardRepository
from agribank_v3.features.credit.summary.dashboard_charts import branch_bar_values
from agribank_v3.features.credit.summary.dashboard_export import (
    SHEET_BRANCH,
    SHEET_DETAIL,
    SHEET_GROWTH,
    SHEET_OVERVIEW,
    DashboardNimExportService,
)
from agribank_v3.features.credit.summary.dashboard_service import DashboardFilters, build_nim_dashboard
from agribank_v3.features.credit.summary.dashboard_window import NimDashboardWindow
from agribank_v3.features.credit.summary.database import CREDIT_SUMMARY_DATABASE_NAME, credit_summary_database_path
from agribank_v3.features.credit.summary.menu import SUMMARY_FEATURES
from agribank_v3.features.credit.summary.models import (
    LOAN_COMPARE_TITLE,
    DashboardData,
    DashboardMetric,
    SummaryDataType,
    SummaryError,
)
from agribank_v3.features.credit.summary.nim_ui_config import NIM_DN_UI_CONFIG, NIM_NV_UI_CONFIG
from agribank_v3.features.credit.summary.officer_history.export import SHEET_NAMES, export_analysis_rows
from agribank_v3.features.credit.summary.officer_history.models import (
    METRIC_AVERAGE_RATE,
    METRIC_BALANCE,
    METRIC_BALANCE_GROWTH,
    METRIC_NIM_AFTER,
    METRIC_NIM_BEFORE,
    HistoryFilters,
    HistoryPoint,
)
from agribank_v3.features.credit.summary.officer_history.repository import OfficerHistoryRepository, officer_key
from agribank_v3.features.credit.summary.officer_history.service import (
    build_multiple_officer_comparison,
    build_officer_branch_comparison,
    build_officer_growth_history,
    build_officer_overview,
)
from agribank_v3.features.credit.summary.officer_history.window import OfficerHistoryDialog
from agribank_v3.features.credit.summary.officer_history.widgets import OfficerMultiSelectCombo
from agribank_v3.features.credit.summary.reports import (
    CREDIT_LIMIT_VBA_HEADERS,
    LOAN_COMPARE_VBA_HEADERS,
    NIM_DN_DETAIL_HEADERS,
    NIM_NV_DETAIL_HEADERS,
    export_rows,
)
from agribank_v3.features.credit.summary.repository import SummaryRepository
from agribank_v3.features.credit.summary.regression import compare_workbooks
from agribank_v3.features.credit.summary.services import (
    build_officer_history,
    compare_loan_balances,
    export_officer_history_excel,
    import_credit_limit_file,
    import_nim_dn,
    import_nim_nv,
)
from agribank_v3.features.credit.summary.windows import (
    CreditLimitTab,
    DeleteNimPeriodDialog,
    LoanCompareTab,
    NimTab,
    SummaryMaintenanceDialog,
    _display_officer_name,
    _format_money_vn,
    _format_percent_vn,
    _populate_combo,
)
from agribank_v3.features.settings.unit_directory.models import (
    AppUnitSettings,
    BranchDirectoryEntry,
    OfficeDirectoryEntry,
    TRANSACTION_OFFICE,
)
from agribank_v3.ui.components.controls import (
    configure_combo_popup_width,
    danger_button,
    primary_button,
    recommended_control_height,
    secondary_button,
)
from agribank_v3.ui.components.kpi import CompactKpiCard, MetricGrid
from agribank_v3.update.db_migrations import MigrationSpec, apply_migrations


class CreditSummaryTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temporary_directory = TemporaryDirectory()
        self.root = Path(self.temporary_directory.name)
        self.database_path = self.root / "DuLieuV3.db"
        self.repository = SummaryRepository(self.database_path)

    def tearDown(self) -> None:
        self.temporary_directory.cleanup()

    def _set_sample_nim_kpis(self, tab: NimTab) -> list[CompactKpiCard]:
        tab.metrics.set_data(
            DashboardData(
                metrics=(
                    DashboardMetric(tab.ui_config.total_balance_label, "28.056.021.646.654"),
                    DashboardMetric("NIM trước ĐC", "2,31%"),
                    DashboardMetric("NIM sau ĐC", "2,40%"),
                    DashboardMetric("Lãi suất bình quân", "8,01%"),
                )
            )
        )
        return tab.metrics.findChildren(CompactKpiCard)

    def _assert_button_text_fits(self, button: QPushButton, text: str) -> None:
        button.setText(text)
        required = recommended_control_height(button)
        self.assertGreaterEqual(button.minimumHeight(), required)
        self.assertGreaterEqual(max(button.sizeHint().height(), button.minimumHeight()), required)
        self.assertEqual(button.maximumHeight(), 16777215)

    def _nim_combo_with_long_officer(self) -> QComboBox:
        tab = NimTab(self.repository, SummaryDataType.NIM_DN)
        self.addCleanup(tab.deleteLater)
        combo = tab.officer_filter
        _populate_combo(
            combo,
            [("Nguyễn Văn A Phòng GD Trung tâm", "540000321")],
        )
        return combo

    def test_nim_dn_and_nv_import_use_vba_formulas(self) -> None:
        dn = self.root / "5491_FTPLN_20260331.csv"
        dn.write_text(
            "\n".join(
                [
                    "BRCD,FTP,INTRT,MUCFTPDC,CBTD,TRCTCD,LDRBAL,TRREF,CUSTTP",
                    "5491,2,10,1,CB1,00,1000,,CN",
                ]
            ),
            encoding="utf-8",
        )
        nv = self.root / "5491_FTPDP_20260331.csv"
        nv.write_text(
            "\n".join(
                [
                    "BRCD,FTP,INTRT,MUCFTPDC,CBHD,TRCTCD,LDRBAL,TRREF,CUSTTP",
                    "5491,2,10,1,CB2,00,1000,,TC",
                ]
            ),
            encoding="utf-8",
        )

        dn_output = self.root / "BaoCaoNIM_CSDL.xlsx"
        nv_output = self.root / "BaoCaoNIM_NV_CSDL.xlsx"
        result_dn = import_nim_dn(self.repository, self.root, export_path=dn_output)
        result_nv = import_nim_nv(self.repository, self.root, export_path=nv_output)

        self.assertEqual(result_dn.row_count, 1)
        self.assertEqual(result_nv.row_count, 1)
        self.assertEqual(result_dn.output_path, dn_output)
        self.assertEqual(result_nv.output_path, nv_output)
        page_dn = self.repository.query_nim(SummaryDataType.NIM_DN)
        self.assertEqual(page_dn.total_rows, 1)
        row_dn = page_dn.rows[0]
        self.assertEqual(row_dn["period"], "2026-03")
        self.assertAlmostEqual(float(row_dn["average_rate"]), 10.0)
        self.assertAlmostEqual(float(row_dn["nim_before"]), 8.0)
        self.assertAlmostEqual(float(row_dn["nim_after"]), 7.0)

        page_nv = self.repository.query_nim(SummaryDataType.NIM_NV)
        row_nv = page_nv.rows[0]
        self.assertAlmostEqual(float(row_nv["nim_before"]), -8.0)
        self.assertAlmostEqual(float(row_nv["nim_after"]), -7.0)
        workbook_dn = load_workbook(dn_output, data_only=True)
        try:
            self.assertEqual(workbook_dn.sheetnames, ["Cache_Nim", "Báo Cáo NIM DN", "Sheet1", "Sheet2", "Sheet3"])
            self.assertEqual(workbook_dn["Cache_Nim"].sheet_state, "hidden")
            self.assertEqual(
                [workbook_dn["Cache_Nim"].cell(1, col).value for col in range(1, 9)],
                list(NIM_DN_DETAIL_HEADERS),
            )
            self.assertEqual(
                [workbook_dn["Báo Cáo NIM DN"].cell(15, col).value for col in range(1, 9)],
                list(NIM_DN_DETAIL_HEADERS),
            )
        finally:
            workbook_dn.close()
        workbook_nv = load_workbook(nv_output, data_only=True)
        try:
            self.assertEqual(workbook_nv.sheetnames, ["Cache_Nim_NV", "Báo Cáo NIM NV", "Sheet1", "Sheet2", "Sheet3"])
            self.assertEqual(workbook_nv["Cache_Nim_NV"].sheet_state, "hidden")
            self.assertEqual(
                [workbook_nv["Cache_Nim_NV"].cell(1, col).value for col in range(1, 8)],
                list(NIM_NV_DETAIL_HEADERS),
            )
            self.assertEqual(
                [workbook_nv["Báo Cáo NIM NV"].cell(15, col).value for col in range(1, 8)],
                list(NIM_NV_DETAIL_HEADERS),
            )
        finally:
            workbook_nv.close()

    def test_nim_branch_combo_uses_unit_directory(self) -> None:
        self.repository.unit_directory.save_branch(
            BranchDirectoryEntry(branch_code="6501", branch_name="Chi nhánh Động", short_name="CN Động")
        )
        self._write_nim_dn_file("6501_FTPLN_20260331.csv", ["6501,2,10,1,CB1,00,1000,,CN"])
        import_nim_dn(self.repository, self.root)
        app = QApplication.instance() or QApplication([])
        _ = app
        tab = NimTab(self.repository, SummaryDataType.NIM_DN)
        self.addCleanup(tab.deleteLater)

        labels = [tab.branch_filter.itemText(index) for index in range(tab.branch_filter.count())]

        self.assertIn("6501 - CN Động", labels)
        self.assertEqual(tab.branch_filter.itemData(labels.index("6501 - CN Động")), "6501")

    def test_nim_office_combo_uses_unit_directory(self) -> None:
        self.repository.unit_directory.save_branch(
            BranchDirectoryEntry(branch_code="6501", branch_name="Chi nhánh Động", short_name="CN Động")
        )
        self.repository.unit_directory.save_office(
            OfficeDirectoryEntry(
                id=None,
                branch_code="6501",
                trctcd="07",
                office_code="",
                office_name="Phòng giao dịch Kiểm thử",
                short_name="PGD Kiểm thử",
                office_type=TRANSACTION_OFFICE,
            )
        )
        self._write_nim_dn_file("6501_FTPLN_20260331.csv", ["6501,2,10,1,CB1,07,1000,,CN"])
        import_nim_dn(self.repository, self.root)
        app = QApplication.instance() or QApplication([])
        _ = app
        tab = NimTab(self.repository, SummaryDataType.NIM_DN)
        self.addCleanup(tab.deleteLater)

        labels = [tab.transaction_office_filter.itemText(index) for index in range(tab.transaction_office_filter.count())]

        self.assertIn("6501-07 - PGD Kiểm thử", labels)
        self.assertEqual(tab.transaction_office_filter.itemData(labels.index("6501-07 - PGD Kiểm thử")), "6501-07")

    def test_table_branch_name_changes_after_directory_update(self) -> None:
        self.repository.unit_directory.save_branch(
            BranchDirectoryEntry(branch_code="6501", branch_name="Chi nhánh Cũ", short_name="CN Cũ")
        )
        self._write_nim_dn_file("6501_FTPLN_20260331.csv", ["6501,2,10,1,CB1,00,1000,,CN"])
        import_nim_dn(self.repository, self.root)
        self.repository.unit_directory.save_branch(
            BranchDirectoryEntry(branch_code="6501", branch_name="Chi nhánh Mới", short_name="CN Mới")
        )

        row = self.repository.query_nim(SummaryDataType.NIM_DN).rows[0]

        self.assertEqual(row["branch"], "6501 - CN Mới")

    def test_officer_history_uses_current_branch_name_after_directory_update(self) -> None:
        self.repository.unit_directory.save_branch(
            BranchDirectoryEntry(branch_code="6501", branch_name="Chi nhánh Cũ", short_name="CN Cũ")
        )
        self._write_nim_dn_file("6501_FTPLN_20260331.csv", ["6501,2,10,1,[650100001] Nguyễn Văn A,00,1000,,CN"])
        import_nim_dn(self.repository, self.root)
        self.repository.unit_directory.save_branch(
            BranchDirectoryEntry(branch_code="6501", branch_name="Chi nhánh Mới", short_name="CN Mới")
        )

        history = build_officer_history(
            self.repository,
            SummaryDataType.NIM_DN,
            officer="[650100001] Nguyễn Văn A",
            branch="6501 - CN Mới",
            transaction_office="Hội sở",
        )

        self.assertEqual(history.current_period, "2026-03")
        self.assertAlmostEqual(history.current_balance, 1000.0)

    def test_export_uses_current_branch_name(self) -> None:
        self.repository.unit_directory.save_branch(
            BranchDirectoryEntry(branch_code="6501", branch_name="Chi nhánh Xuất", short_name="CN Xuất")
        )
        self._write_nim_dn_file("6501_FTPLN_20260331.csv", ["6501,2,10,1,CB1,00,1000,,CN"])
        import_nim_dn(self.repository, self.root)
        self.repository.unit_directory.save_branch(
            BranchDirectoryEntry(branch_code="6501", branch_name="Chi nhánh Xuất đổi", short_name="CN Xuất đổi")
        )
        data = build_nim_dashboard(
            NimDashboardRepository(self.repository),
            SummaryDataType.NIM_DN,
            DashboardFilters(branch="6501"),
        )
        output = self.root / "nim-dashboard.xlsx"

        DashboardNimExportService(data).export_detail(output)
        workbook = load_workbook(output, data_only=True)
        try:
            sheet = workbook[NIM_DN_UI_CONFIG.dashboard_sheets["detail"]]
            names = [sheet.cell(row, 2).value for row in range(2, sheet.max_row + 1)]
        finally:
            workbook.close()

        self.assertIn("6501 - CN Xuất đổi", names)

    def test_export_reflects_branch_rename_without_reimport(self) -> None:
        self.repository.unit_directory.save_branch(
            BranchDirectoryEntry(branch_code="6501", branch_name="Chi nhánh Export cũ", short_name="CN Export cũ")
        )
        self._write_nim_dn_file("6501_FTPLN_20260331.csv", ["6501,2,10,1,CB1,00,1000,,CN"])
        import_nim_dn(self.repository, self.root)
        self.repository.unit_directory.save_branch(
            BranchDirectoryEntry(branch_code="6501", branch_name="Chi nhánh Export mới", short_name="CN Export mới")
        )

        data = build_nim_dashboard(
            NimDashboardRepository(self.repository),
            SummaryDataType.NIM_DN,
            DashboardFilters(branch="6501"),
        )
        output = self.root / "nim-export-rename.xlsx"
        DashboardNimExportService(data).export_detail(output)
        workbook = load_workbook(output, data_only=True)
        try:
            sheet = workbook[NIM_DN_UI_CONFIG.dashboard_sheets["detail"]]
            exported_names = [sheet.cell(row, 2).value for row in range(2, sheet.max_row + 1)]
        finally:
            workbook.close()

        self.assertIn("6501 - CN Export mới", exported_names)

    def test_unit_directory_changed_refreshes_open_windows(self) -> None:
        self.repository.unit_directory.save_branch(
            BranchDirectoryEntry(branch_code="6501", branch_name="Chi nhánh đang mở", short_name="CN đang mở")
        )
        self._write_nim_dn_file("6501_FTPLN_20260331.csv", ["6501,2,10,1,CB1,00,1000,,CN"])
        import_nim_dn(self.repository, self.root)
        app = QApplication.instance() or QApplication([])
        window = NimDashboardWindow(self.repository, SummaryDataType.NIM_DN)
        self.addCleanup(window.close)

        self.repository.unit_directory.save_branch(
            BranchDirectoryEntry(branch_code="6501", branch_name="Chi nhánh đã đổi", short_name="CN đã đổi")
        )
        app.processEvents()

        labels = [window.branch_combo.itemText(index) for index in range(window.branch_combo.count())]
        self.assertIn("6501 - CN đã đổi", labels)

    def test_unit_directory_changed_preserves_selected_code(self) -> None:
        self.repository.unit_directory.save_branch(
            BranchDirectoryEntry(branch_code="6501", branch_name="Chi nhánh chọn", short_name="CN chọn")
        )
        self._write_nim_dn_file("6501_FTPLN_20260331.csv", ["6501,2,10,1,CB1,00,1000,,CN"])
        import_nim_dn(self.repository, self.root)
        app = QApplication.instance() or QApplication([])
        window = NimDashboardWindow(self.repository, SummaryDataType.NIM_DN)
        self.addCleanup(window.close)
        index = window.branch_combo.findData("6501")
        self.assertGreaterEqual(index, 0)
        window.branch_combo.setCurrentIndex(index)

        self.repository.unit_directory.save_branch(
            BranchDirectoryEntry(branch_code="6501", branch_name="Chi nhánh chọn mới", short_name="CN chọn mới")
        )
        app.processEvents()

        self.assertEqual(window.branch_combo.currentData(), "6501")

    def test_no_runtime_legacy_mapping_lookup(self) -> None:
        self.repository.unit_directory.save_branch(
            BranchDirectoryEntry(branch_code="6501", branch_name="Chi nhánh Triển khai", short_name="CN Triển khai")
        )
        self._write_nim_dn_file("6501_FTPLN_20260331.csv", ["6501,2,10,1,CB1,00,1000,,CN"])
        import_nim_dn(self.repository, self.root)

        row = self.repository.query_nim(SummaryDataType.NIM_DN).rows[0]

        self.assertEqual(row["branch"], "6501 - CN Triển khai")

    def test_import_unknown_branch_does_not_fail(self) -> None:
        self._write_nim_dn_file("6501_FTPLN_20260331.csv", ["6501,2,10,1,CB1,00,1000,,CN"])

        result = import_nim_dn(self.repository, self.root)

        self.assertEqual(result.row_count, 1)
        self.assertIsNotNone(self.repository.unit_directory.get_branch("6501"))

    def test_import_unknown_office_does_not_fail(self) -> None:
        self._write_nim_dn_file("6501_FTPLN_20260331.csv", ["6501,2,10,1,CB1,03,1000,,CN"])

        result = import_nim_dn(self.repository, self.root)

        self.assertEqual(result.row_count, 1)
        self.assertIsNotNone(self.repository.unit_directory.get_office("6501", "03"))

    def test_unknown_branch_placeholder_or_warning(self) -> None:
        self._write_nim_dn_file("6501_FTPLN_20260331.csv", ["6501,2,10,1,CB1,00,1000,,CN"])

        result = import_nim_dn(self.repository, self.root)

        self.assertIn("6501", result.message)
        self.assertEqual(self.repository.unit_directory.get_branch_display_name("6501"), "6501 - CN chưa khai báo")

    def test_unknown_office_not_assigned_wrong_name(self) -> None:
        self._write_nim_dn_file("6501_FTPLN_20260331.csv", ["6501,2,10,1,CB1,03,1000,,CN"])
        import_nim_dn(self.repository, self.root)

        self.assertEqual(self.repository.unit_directory.get_office_display_name("6501", "03"), "6501-03 - PGD 03")

    def test_nim_dn_result_unchanged(self) -> None:
        self._write_nim_dn_file("5491_FTPLN_20260331.csv", ["5491,2,10,1,CB1,00,1000,,CN"])

        import_nim_dn(self.repository, self.root)
        row = self.repository.query_nim(SummaryDataType.NIM_DN).rows[0]

        self.assertAlmostEqual(float(row["average_rate"]), 10.0)
        self.assertAlmostEqual(float(row["nim_before"]), 8.0)
        self.assertAlmostEqual(float(row["nim_after"]), 7.0)

    def test_nim_nv_result_unchanged(self) -> None:
        self._write_nim_nv_file("5491_FTPDP_20260331.csv", ["5491,2,10,1,CB1,00,1000,,CN"])

        import_nim_nv(self.repository, self.root)
        row = self.repository.query_nim(SummaryDataType.NIM_NV).rows[0]

        self.assertAlmostEqual(float(row["nim_before"]), -8.0)
        self.assertAlmostEqual(float(row["nim_after"]), -7.0)

    def test_nim_ui_display_format_helpers(self) -> None:
        self.assertEqual(_display_officer_name("[540000321] Nguyễn Văn A"), "Nguyễn Văn A")
        self.assertEqual(_display_officer_name("Nguyễn Văn B"), "Nguyễn Văn B")
        self.assertEqual(_format_money_vn(1234567890), "1.234.567.890")
        self.assertEqual(_format_percent_vn(7.15), "7,15%")
        self.assertEqual(_format_percent_vn(2.34), "2,34%")
        self.assertEqual(_format_percent_vn(0.58), "0,58%")

    def test_summary_buttons_use_shared_style(self) -> None:
        app = QApplication.instance() or QApplication([])
        _ = app
        tabs = (
            NimTab(self.repository, SummaryDataType.NIM_DN),
            LoanCompareTab(self.repository),
            CreditLimitTab(self.repository),
        )
        try:
            for tab in tabs:
                buttons = [button for button in tab.findChildren(QPushButton) if button.text()]
                self.assertTrue(buttons)
                self.assertTrue(
                    all(button.objectName() in {"PrimaryButton", "SecondaryButton", "DangerButton", "KpiToggleButton"} for button in buttons),
                    [f"{button.text()}={button.objectName()}" for button in buttons],
                )
        finally:
            for tab in tabs:
                tab.deleteLater()

    def test_summary_comboboxes_use_shared_style(self) -> None:
        app = QApplication.instance() or QApplication([])
        _ = app
        tabs = (
            NimTab(self.repository, SummaryDataType.NIM_DN),
            LoanCompareTab(self.repository),
            CreditLimitTab(self.repository),
        )
        try:
            combos = [combo for tab in tabs for combo in tab.findChildren(QComboBox)]
            self.assertTrue(combos)
            self.assertTrue(all(combo.objectName() == "AgribankComboBox" for combo in combos))
        finally:
            for tab in tabs:
                tab.deleteLater()

    def test_credit_limit_uses_unit_directory(self) -> None:
        self.repository.unit_directory.save_branch(
            BranchDirectoryEntry(branch_code="6501", branch_name="Chi nhánh Hạn mức", short_name="CN Hạn mức")
        )
        self.repository.unit_directory.save_office(
            OfficeDirectoryEntry(
                id=None,
                branch_code="6501",
                trctcd="00",
                office_code="6501-00",
                office_name="Hội sở Hạn mức",
                short_name="Hội sở Hạn mức",
                office_type="HEAD_OFFICE",
            )
        )
        self.repository.unit_directory.save_settings(
            AppUnitSettings(home_branch_code="6501", default_office_code="6501-00")
        )
        app = QApplication.instance() or QApplication([])
        _ = app
        tab = CreditLimitTab(self.repository)
        try:
            tab.reload()
            self.assertIs(tab.repository.unit_directory, self.repository.unit_directory)
            self.assertEqual(tab.repository.unit_directory.get_home_branch().branch_code, "6501")
        finally:
            tab.deleteLater()

    def test_nim_kpi_cards_use_shared_compact_component(self) -> None:
        app = QApplication.instance() or QApplication([])
        _ = app
        tabs = (
            NimTab(self.repository, SummaryDataType.NIM_DN),
            NimTab(self.repository, SummaryDataType.NIM_NV),
        )
        try:
            for tab in tabs:
                tab.metrics.set_data(
                    DashboardData(
                        metrics=(
                            DashboardMetric(tab.ui_config.total_balance_label, "28.056.021.646.654"),
                            DashboardMetric("NIM trước ĐC", "2,31%"),
                            DashboardMetric("NIM sau ĐC", "2,11%"),
                        )
                    )
                )
                self.assertIsInstance(tab.metrics, MetricGrid)
                cards = tab.metrics.findChildren(CompactKpiCard)
                self.assertEqual(len(cards), 3)
                self.assertTrue(all(card.objectName().startswith("CompactKpiCard") for card in cards))
                self.assertTrue(all(card.maximumHeight() <= 80 for card in cards))
        finally:
            for tab in tabs:
                tab.deleteLater()

    def test_nim_kpi_money_is_compact_with_full_tooltip(self) -> None:
        app = QApplication.instance() or QApplication([])
        _ = app
        tab = NimTab(self.repository, SummaryDataType.NIM_DN)
        try:
            tab.metrics.set_data(
                DashboardData(
                    metrics=(DashboardMetric("Tổng dư nợ", "28.056.021.646.654"),)
                )
            )
            card = tab.metrics.findChildren(CompactKpiCard)[0]
            self.assertIn("tỷ", card.value_label.text())
            self.assertNotEqual(card.value_label.text(), "28.056.021.646.654")
            self.assertIn("28.056.021.646.654 đồng", card.toolTip())
        finally:
            tab.deleteLater()

    def test_nim_kpi_percentage_display_preserves_scale(self) -> None:
        app = QApplication.instance() or QApplication([])
        _ = app
        tab = NimTab(self.repository, SummaryDataType.NIM_DN)
        try:
            tab.metrics.set_data(
                DashboardData(
                    metrics=(
                        DashboardMetric("NIM trước ĐC", "2,31%"),
                        DashboardMetric("Lãi suất bình quân", "7,15%"),
                    )
                )
            )
            values = [card.value_label.text() for card in tab.metrics.findChildren(CompactKpiCard)]
            self.assertEqual(values, ["2,31%", "7,15%"])
            self.assertNotIn("231,00%", values)
            self.assertNotIn("715,00%", values)
        finally:
            tab.deleteLater()

    def test_nim_kpi_empty_value_shows_no_data(self) -> None:
        app = QApplication.instance() or QApplication([])
        _ = app
        tab = NimTab(self.repository, SummaryDataType.NIM_NV)
        try:
            tab.metrics.set_data(DashboardData(metrics=(DashboardMetric("Tổng nguồn vốn", ""),)))
            card = tab.metrics.findChildren(CompactKpiCard)[0]
            self.assertEqual(card.value_label.text(), "—")
            self.assertIn("Không có dữ liệu", card.toolTip())
        finally:
            tab.deleteLater()

    def test_officer_history_info_uses_shared_kpi_grid(self) -> None:
        app = QApplication.instance() or QApplication([])
        _ = app
        self._write_nim_dn_file(
            "5491_FTPLN_20260131.csv",
            ["5491,2,10,1,[540000321] Nguyễn Văn A,00,1000,,CN"],
        )
        import_nim_dn(self.repository, self.root)
        dialog = self._open_officer_history_dialog(SummaryDataType.NIM_DN, "[540000321] Nguyễn Văn A")
        try:
            self.assertIsInstance(dialog.info_metrics, MetricGrid)
            titles = [card.title_label.text() for card in dialog.info_metrics.findChildren(CompactKpiCard)]
            self.assertIn("CBTD", titles)
            self.assertIn("Tổng dư nợ kỳ hiện tại", titles)
            self.assertIn("NIM sau ĐC hiện tại", titles)
        finally:
            dialog.close()
            dialog.deleteLater()

    def test_nim_combos_use_shared_popup_and_preserve_data(self) -> None:
        app = QApplication.instance() or QApplication([])
        _ = app
        tab = NimTab(self.repository, SummaryDataType.NIM_DN)
        try:
            combo = tab.officer_filter
            _populate_combo(combo, [("Nguyễn Văn A", "[540000321] Nguyễn Văn A")])
            self.assertEqual(combo.objectName(), "AgribankComboBox")
            self.assertEqual(combo.view().objectName(), "AgribankComboPopup")
            self.assertEqual(combo.view().textElideMode(), Qt.TextElideMode.ElideNone)
            self.assertEqual(combo.itemData(1), "[540000321] Nguyễn Văn A")
            changes = {"count": 0}
            combo.currentIndexChanged.connect(lambda _index: changes.__setitem__("count", changes["count"] + 1))
            combo.setCurrentIndex(1)
            self.assertEqual(changes["count"], 1)
        finally:
            tab.deleteLater()

    def test_officer_history_combos_use_shared_popup(self) -> None:
        app = QApplication.instance() or QApplication([])
        _ = app
        self._write_nim_dn_file(
            "5491_FTPLN_20260131.csv",
            ["5491,2,10,1,[540000321] Nguyễn Văn A,00,1000,,CN"],
        )
        import_nim_dn(self.repository, self.root)
        dialog = self._open_officer_history_dialog(SummaryDataType.NIM_DN, "[540000321] Nguyễn Văn A")
        try:
            combos = dialog.findChildren(QComboBox)
            self.assertTrue(combos)
            self.assertTrue(all(combo.objectName() == "AgribankComboBox" for combo in combos))
            self.assertTrue(all(combo.view().objectName() == "AgribankComboPopup" for combo in combos))
        finally:
            dialog.close()
            dialog.deleteLater()

    def test_shared_buttons_have_font_metric_safe_height(self) -> None:
        app = QApplication.instance() or QApplication([])
        _ = app
        buttons = [
            primary_button("Áp dụng cập nhật"),
            secondary_button("Nhập dữ liệu"),
            danger_button("Xóa dữ liệu"),
        ]
        try:
            for button in buttons:
                self.assertGreaterEqual(button.minimumHeight(), recommended_control_height(button))
                self.assertEqual(button.maximumHeight(), 16777215)
        finally:
            for button in buttons:
                button.deleteLater()

    def test_shared_button_height_handles_scaled_font(self) -> None:
        app = QApplication.instance() or QApplication([])
        _ = app
        button = primary_button("Đồng ý")
        try:
            font = button.font()
            font.setPointSize(max(font.pointSize() + 5, 17))
            button.setFont(font)
            button.setMinimumHeight(recommended_control_height(button))
            self.assertGreaterEqual(button.minimumHeight(), recommended_control_height(button))
        finally:
            button.deleteLater()

    def test_button_size_hint_stable_when_pressed_and_focused(self) -> None:
        app = QApplication.instance() or QApplication([])
        _ = app
        button = secondary_button("Áp dụng")
        try:
            before = button.sizeHint()
            button.setFocus(Qt.FocusReason.OtherFocusReason)
            button.setDown(True)
            app.processEvents()
            self.assertEqual(button.sizeHint(), before)
            button.setDown(False)
        finally:
            button.deleteLater()

    def test_nim_debt_uses_shared_compact_kpi_card(self) -> None:
        app = QApplication.instance() or QApplication([])
        _ = app
        tab = NimTab(self.repository, SummaryDataType.NIM_DN)
        try:
            cards = self._set_sample_nim_kpis(tab)
            self.assertIsInstance(tab.metrics, MetricGrid)
            self.assertTrue(cards)
            self.assertTrue(all(isinstance(card, CompactKpiCard) for card in cards))
        finally:
            tab.deleteLater()

    def test_nim_funding_uses_shared_compact_kpi_card(self) -> None:
        app = QApplication.instance() or QApplication([])
        _ = app
        tab = NimTab(self.repository, SummaryDataType.NIM_NV)
        try:
            cards = self._set_sample_nim_kpis(tab)
            self.assertIsInstance(tab.metrics, MetricGrid)
            self.assertTrue(cards)
            self.assertTrue(all(isinstance(card, CompactKpiCard) for card in cards))
        finally:
            tab.deleteLater()

    def test_nim_kpi_money_is_compact(self) -> None:
        app = QApplication.instance() or QApplication([])
        _ = app
        tab = NimTab(self.repository, SummaryDataType.NIM_DN)
        try:
            cards = self._set_sample_nim_kpis(tab)
            balance_card = next(card for card in cards if card.title_label.text() == "Tổng dư nợ")
            self.assertEqual(balance_card.value_label.text(), "28.056,02 tỷ")
        finally:
            tab.deleteLater()

    def test_nim_kpi_money_tooltip_has_full_value(self) -> None:
        app = QApplication.instance() or QApplication([])
        _ = app
        tab = NimTab(self.repository, SummaryDataType.NIM_DN)
        try:
            cards = self._set_sample_nim_kpis(tab)
            balance_card = next(card for card in cards if card.title_label.text() == "Tổng dư nợ")
            self.assertIn("Tổng dư nợ", balance_card.toolTip())
            self.assertIn("28.056.021.646.654 đồng", balance_card.toolTip())
        finally:
            tab.deleteLater()

    def test_nim_kpi_percentage_format(self) -> None:
        app = QApplication.instance() or QApplication([])
        _ = app
        tab = NimTab(self.repository, SummaryDataType.NIM_DN)
        try:
            cards = self._set_sample_nim_kpis(tab)
            values = {card.title_label.text(): card.value_label.text() for card in cards}
            self.assertEqual(values["NIM trước ĐC"], "2,31%")
            self.assertEqual(values["NIM sau ĐC"], "2,40%")
            self.assertEqual(values["Lãi suất bình quân"], "8,01%")
        finally:
            tab.deleteLater()

    def test_nim_kpi_no_data_state(self) -> None:
        app = QApplication.instance() or QApplication([])
        _ = app
        tab = NimTab(self.repository, SummaryDataType.NIM_DN)
        try:
            tab.metrics.set_data(DashboardData(metrics=(DashboardMetric("Tổng dư nợ", ""),)))
            card = tab.metrics.findChildren(CompactKpiCard)[0]
            self.assertEqual(card.value_label.text(), "—")
            self.assertNotIn("None", card.toolTip())
            self.assertNotIn("nan", card.toolTip().casefold())
        finally:
            tab.deleteLater()

    def test_nim_kpi_grid_responsive(self) -> None:
        app = QApplication.instance() or QApplication([])
        _ = app
        tab = NimTab(self.repository, SummaryDataType.NIM_DN)
        try:
            tab.metrics.set_metrics([("KPI " + str(index), index, "count") for index in range(8)])
            tab.metrics.resize(1600, 120)
            tab.metrics.resizeEvent(None)
            self.assertEqual(tab.metrics.main_column_count(), 8)
            tab.metrics.resize(700, 220)
            tab.metrics.resizeEvent(None)
            self.assertEqual(tab.metrics.main_column_count(), 3)
        finally:
            tab.deleteLater()

    def test_nim_kpi_values_unchanged_after_style_refactor(self) -> None:
        app = QApplication.instance() or QApplication([])
        _ = app
        self._write_nim_dn_file(
            "5491_FTPLN_20260131.csv",
            ["5491,2,10,1,[540000321] Nguyễn Văn A,00,1000,,CN"],
        )
        import_nim_dn(self.repository, self.root)
        tab = NimTab(self.repository, SummaryDataType.NIM_DN)
        try:
            dashboard = tab.query_dashboard()
            before = tuple((metric.label, metric.value) for metric in dashboard.metrics)
            tab.metrics.set_data(dashboard)
            after = tuple((metric.label, metric.value) for metric in dashboard.metrics)
            self.assertEqual(after, before)
            self.assertIn(("Tổng dư nợ", "1.000"), before)
            self.assertIn(("NIM sau ĐC", "7,00%"), before)
        finally:
            tab.deleteLater()

    def test_nim_combos_use_shared_agribank_style(self) -> None:
        app = QApplication.instance() or QApplication([])
        _ = app
        tab = NimTab(self.repository, SummaryDataType.NIM_DN)
        dashboard = NimDashboardWindow(self.repository, SummaryDataType.NIM_DN, parent=tab)
        try:
            combos = tab.findChildren(QComboBox) + dashboard.findChildren(QComboBox)
            self.assertTrue(combos)
            self.assertTrue(all(combo.objectName() == "AgribankComboBox" for combo in combos))
            self.assertTrue(all(combo.view().objectName() == "AgribankComboPopup" for combo in combos))
        finally:
            dashboard.close()
            dashboard.deleteLater()
            tab.deleteLater()

    def test_nim_combo_closed_height_compact(self) -> None:
        app = QApplication.instance() or QApplication([])
        _ = app
        combo = self._nim_combo_with_long_officer()
        self.assertGreaterEqual(combo.minimumHeight(), 30)
        self.assertLessEqual(combo.maximumHeight(), 34)

    def test_nim_combo_popup_item_height(self) -> None:
        app = QApplication.instance() or QApplication([])
        _ = app
        combo = self._nim_combo_with_long_officer()
        option = QStyleOptionViewItem()
        height = combo.view().itemDelegate().sizeHint(option, combo.model().index(0, 0)).height()
        self.assertGreaterEqual(height, 24)
        self.assertLessEqual(height, 32)

    def test_nim_combo_selected_item_contrast(self) -> None:
        app = QApplication.instance() or QApplication([])
        _ = app
        combo = self._nim_combo_with_long_officer()
        style = combo.view().styleSheet()
        self.assertIn("selection-background-color", style)
        self.assertIn("selection-color: #202020", style)
        self.assertIn("174, 28, 63", style)

    def test_nim_combo_popup_full_officer_name(self) -> None:
        app = QApplication.instance() or QApplication([])
        _ = app
        combo = self._nim_combo_with_long_officer()
        long_text = combo.itemText(1)
        width = configure_combo_popup_width(combo, minimum_popup_width=260)
        self.assertGreaterEqual(width, combo.fontMetrics().horizontalAdvance(long_text) + 40)

    def test_nim_combo_item_data_unchanged(self) -> None:
        app = QApplication.instance() or QApplication([])
        _ = app
        combo = self._nim_combo_with_long_officer()
        self.assertEqual(combo.itemData(1), "540000321")

    def test_nim_combo_signals_unchanged(self) -> None:
        app = QApplication.instance() or QApplication([])
        _ = app
        combo = self._nim_combo_with_long_officer()
        changes = {"count": 0}
        combo.currentIndexChanged.connect(lambda _index: changes.__setitem__("count", changes["count"] + 1))
        combo.setCurrentIndex(1)
        self.assertEqual(changes["count"], 1)

    def test_button_height_fits_font_metrics(self) -> None:
        app = QApplication.instance() or QApplication([])
        _ = app
        button = primary_button("Áp dụng cập nhật Đồng ý")
        try:
            self._assert_button_text_fits(button, button.text())
        finally:
            button.deleteLater()

    def test_button_apply_text_not_clipped(self) -> None:
        app = QApplication.instance() or QApplication([])
        _ = app
        button = primary_button("Áp dụng")
        try:
            self._assert_button_text_fits(button, "Áp dụng")
        finally:
            button.deleteLater()

    def test_button_descender_g_not_clipped(self) -> None:
        app = QApplication.instance() or QApplication([])
        _ = app
        button = secondary_button("Tăng trưởng")
        try:
            self._assert_button_text_fits(button, "Tăng trưởng g")
        finally:
            button.deleteLater()

    def test_button_descender_p_not_clipped(self) -> None:
        app = QApplication.instance() or QApplication([])
        _ = app
        button = secondary_button("Nhập dữ liệu")
        try:
            self._assert_button_text_fits(button, "Nhập p q y")
        finally:
            button.deleteLater()

    def test_button_vietnamese_dot_below_not_clipped(self) -> None:
        app = QApplication.instance() or QApplication([])
        _ = app
        button = danger_button("Khôi phục")
        try:
            self._assert_button_text_fits(button, "dụng động nhập ạ ặ ậ ệ ị ọ ộ ợ ự")
        finally:
            button.deleteLater()

    def test_button_primary_secondary_same_height(self) -> None:
        app = QApplication.instance() or QApplication([])
        _ = app
        primary = primary_button("Áp dụng")
        secondary = secondary_button("Xóa lọc")
        try:
            self.assertEqual(primary.minimumHeight(), secondary.minimumHeight())
        finally:
            primary.deleteLater()
            secondary.deleteLater()

    def test_button_focus_does_not_change_geometry(self) -> None:
        app = QApplication.instance() or QApplication([])
        _ = app
        button = primary_button("Đồng ý")
        try:
            before = button.sizeHint()
            button.setFocus(Qt.FocusReason.OtherFocusReason)
            app.processEvents()
            self.assertEqual(button.sizeHint(), before)
        finally:
            button.deleteLater()

    def test_button_pressed_text_does_not_shift(self) -> None:
        app = QApplication.instance() or QApplication([])
        _ = app
        button = secondary_button("Cập nhật")
        try:
            before = button.sizeHint()
            button.setDown(True)
            app.processEvents()
            self.assertEqual(button.sizeHint(), before)
            button.setDown(False)
        finally:
            button.deleteLater()

    def test_button_no_conflicting_fixed_max_height(self) -> None:
        app = QApplication.instance() or QApplication([])
        _ = app
        button = primary_button("Áp dụng")
        try:
            self.assertEqual(button.maximumHeight(), 16777215)
            stylesheet = Path("src/agribank_v3/ui/styles.py").read_text(encoding="utf-8")
            self.assertNotRegex(stylesheet, r"(?s)QPushButton[^{]*\{[^}]*max-height")
        finally:
            button.deleteLater()

    def test_toolbar_button_text_visible_at_125_percent(self) -> None:
        app = QApplication.instance() or QApplication([])
        _ = app
        button = secondary_button("Phòng giao dịch")
        try:
            font = button.font()
            font.setPointSize(max(font.pointSize() + 2, 13))
            button.setFont(font)
            button.setMinimumHeight(recommended_control_height(button))
            self._assert_button_text_fits(button, "Phòng giao dịch")
        finally:
            button.deleteLater()

    def test_toolbar_button_text_visible_at_150_percent(self) -> None:
        app = QApplication.instance() or QApplication([])
        _ = app
        button = secondary_button("Đề nghị thanh toán")
        try:
            font = button.font()
            font.setPointSize(max(font.pointSize() + 5, 16))
            button.setFont(font)
            button.setMinimumHeight(recommended_control_height(button))
            self._assert_button_text_fits(button, "Đề nghị thanh toán")
        finally:
            button.deleteLater()

    def test_maintenance_delete_batch_button_not_expanding(self) -> None:
        app = QApplication.instance() or QApplication([])
        _ = app
        dialog = SummaryMaintenanceDialog(self.repository)
        try:
            self.assertNotEqual(dialog.delete_batch_button.sizePolicy().horizontalPolicy(), QSizePolicy.Policy.Expanding)
            self.assertLessEqual(dialog.delete_batch_button.sizeHint().width(), 240)
        finally:
            dialog.close()
            dialog.deleteLater()

    def test_maintenance_batch_row_compact(self) -> None:
        app = QApplication.instance() or QApplication([])
        _ = app
        dialog = SummaryMaintenanceDialog(self.repository)
        try:
            self.assertEqual(dialog.batch_row.spacing(), 8)
            self.assertGreaterEqual(dialog.batch_combo.minimumWidth(), 360)
            self.assertEqual(dialog.batch_combo.sizePolicy().horizontalPolicy(), QSizePolicy.Policy.Expanding)
            self.assertEqual(dialog.batch_row.indexOf(dialog.delete_batch_button), 1)
            self.assertIsNotNone(dialog.batch_row.itemAt(2).spacerItem())
        finally:
            dialog.close()
            dialog.deleteLater()

    def test_maintenance_actions_use_two_rows(self) -> None:
        app = QApplication.instance() or QApplication([])
        _ = app
        dialog = SummaryMaintenanceDialog(self.repository)
        try:
            self.assertEqual(dialog.data_actions_row.count(), 4)
            self.assertEqual(dialog.system_actions_row.count(), 3)
            self.assertEqual(dialog.system_actions_flow.count(), 3)
            self.assertIs(dialog.data_actions_row.itemAt(0).widget(), dialog.refresh_button)
            self.assertIs(dialog.system_actions_flow.itemAt(0).widget(), dialog.optimize_button)
        finally:
            dialog.close()
            dialog.deleteLater()

    def test_maintenance_close_button_aligned_right(self) -> None:
        app = QApplication.instance() or QApplication([])
        _ = app
        dialog = SummaryMaintenanceDialog(self.repository)
        try:
            self.assertEqual(dialog.system_actions_row.indexOf(dialog.close_button), dialog.system_actions_row.count() - 1)
            self.assertIsNotNone(dialog.system_actions_row.itemAt(dialog.system_actions_row.count() - 2).spacerItem())
        finally:
            dialog.close()
            dialog.deleteLater()

    def test_maintenance_buttons_do_not_clip_text(self) -> None:
        app = QApplication.instance() or QApplication([])
        _ = app
        dialog = SummaryMaintenanceDialog(self.repository)
        try:
            buttons = dialog.findChildren(QPushButton)
            self.assertTrue(buttons)
            for button in buttons:
                self.assertGreaterEqual(button.minimumHeight(), recommended_control_height(button))
                self.assertEqual(button.maximumHeight(), 16777215)
        finally:
            dialog.close()
            dialog.deleteLater()

    def test_maintenance_layout_responsive(self) -> None:
        app = QApplication.instance() or QApplication([])
        _ = app
        dialog = SummaryMaintenanceDialog(self.repository)
        try:
            dialog.resize(760, 320)
            app.processEvents()
            self.assertLessEqual(dialog.minimumWidth(), 760)
            self.assertTrue(dialog.data_actions_row.hasHeightForWidth())
            self.assertGreater(
                dialog.data_actions_row.heightForWidth(420),
                dialog.data_actions_row.heightForWidth(900),
            )
            self.assertTrue(dialog.system_actions_flow.hasHeightForWidth())
            self.assertLessEqual(dialog.system_actions_row.sizeHint().width(), dialog.width())
        finally:
            dialog.close()
            dialog.deleteLater()

    def test_maintenance_no_horizontal_overflow(self) -> None:
        app = QApplication.instance() or QApplication([])
        _ = app
        dialog = SummaryMaintenanceDialog(self.repository)
        try:
            dialog.resize(max(760, dialog.minimumWidth()), 320)
            app.processEvents()
            available = dialog.width() - dialog.layout().contentsMargins().left() - dialog.layout().contentsMargins().right()
            self.assertLessEqual(dialog.batch_row.sizeHint().width(), available)
            for index in range(dialog.data_actions_row.count()):
                self.assertLessEqual(dialog.data_actions_row.itemAt(index).sizeHint().width(), available)
            self.assertLessEqual(dialog.system_actions_row.sizeHint().width(), available)
        finally:
            dialog.close()
            dialog.deleteLater()

    def test_menu_label_is_customer_increase_decrease_comparison(self) -> None:
        self.assertEqual(LOAN_COMPARE_TITLE, "So sánh tăng giảm khách hàng")
        titles = [feature.title for feature in SUMMARY_FEATURES]
        self.assertIn("So sánh tăng giảm khách hàng", titles)
        self.assertNotIn("So sánh dư nợ", titles)

    def test_loan_compare_headers_are_vietnamese(self) -> None:
        app = QApplication.instance() or QApplication([])
        _ = app
        self._seed_loan_compare_rows()
        tab = LoanCompareTab(self.repository)
        try:
            tab.reload()
            headers = [tab.table.horizontalHeaderItem(index).text() for index in range(tab.table.columnCount())]
            self.assertEqual(
                headers,
                [
                    "Mã KH",
                    "Tên khách hàng",
                    "Dư nợ kỳ trước",
                    "Dư nợ kỳ này",
                    "Tăng/giảm",
                    "Loại KH",
                    "Cán bộ QL",
                    "Địa chỉ",
                ],
            )
        finally:
            tab.deleteLater()

    def test_loan_compare_money_display_uses_dot_separator(self) -> None:
        app = QApplication.instance() or QApplication([])
        _ = app
        self._seed_loan_compare_rows()
        tab = LoanCompareTab(self.repository)
        try:
            tab.reload()
            values = [tab.table.item(0, column).text() for column in (2, 3, 4)]
            self.assertEqual(values, ["13.500.000.000", "0", "-13.500.000.000"])
            for column in (2, 3, 4):
                self.assertEqual(tab.table.item(0, column).textAlignment(), Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
            ui_rows = tab._export_rows()
            output = export_rows(ui_rows, self.root / "loan-ui.xlsx", title=tab.title, sheet_name=tab.title)
            workbook = load_workbook(output, data_only=True)
            try:
                worksheet = workbook[tab.title[:31]]
                self.assertEqual(worksheet["A3"].value, "000123")
                self.assertEqual(worksheet["A3"].number_format, "@")
                self.assertEqual(worksheet["C3"].value, 13500000000)
                self.assertEqual(worksheet["C3"].number_format, "#,##0")
            finally:
                workbook.close()
        finally:
            tab.deleteLater()

    def test_delete_nim_dn_period(self) -> None:
        self._write_nim_dn_file("5491_FTPLN_20260131.csv", ["5491,2,10,1,CB1,00,1000,,CN"])
        self._write_nim_dn_file("5491_FTPLN_20260228.csv", ["5491,2,10,1,CB1,00,2000,,CN"])
        import_nim_dn(self.repository, self.root)

        info = self.repository.delete_nim_period(SummaryDataType.NIM_DN, "2026-01")

        self.assertEqual(info["row_count"], 1)
        self.assertEqual(self._count_nim_rows(SummaryDataType.NIM_DN, "2026-01"), 0)
        self.assertEqual(self._count_nim_rows(SummaryDataType.NIM_DN, "2026-02"), 1)

    def test_delete_nim_nv_period(self) -> None:
        self._write_nim_nv_file("5491_FTPDP_20260131.csv", ["5491,10,2,1,CB1,00,1000,,CN"])
        self._write_nim_nv_file("5491_FTPDP_20260228.csv", ["5491,10,2,1,CB1,00,2000,,CN"])
        import_nim_nv(self.repository, self.root)

        info = self.repository.delete_nim_period(SummaryDataType.NIM_NV, "2026-01")

        self.assertEqual(info["row_count"], 1)
        self.assertEqual(self._count_nim_rows(SummaryDataType.NIM_NV, "2026-01"), 0)
        self.assertEqual(self._count_nim_rows(SummaryDataType.NIM_NV, "2026-02"), 1)

    def test_delete_nim_period_does_not_delete_other_type(self) -> None:
        self._write_nim_dn_file("5491_FTPLN_20260131.csv", ["5491,2,10,1,CB1,00,1000,,CN"])
        self._write_nim_nv_file("5491_FTPDP_20260131.csv", ["5491,10,2,1,CB2,00,2000,,CN"])
        import_nim_dn(self.repository, self.root)
        import_nim_nv(self.repository, self.root)

        self.repository.delete_nim_period(SummaryDataType.NIM_DN, "2026-01")

        self.assertEqual(self._count_nim_rows(SummaryDataType.NIM_DN, "2026-01"), 0)
        self.assertEqual(self._count_nim_rows(SummaryDataType.NIM_NV, "2026-01"), 1)

    def test_delete_nim_period_transaction_rollback(self) -> None:
        self._write_nim_dn_file("5491_FTPLN_20260131.csv", ["5491,2,10,1,CB1,00,1000,,CN"])
        import_nim_dn(self.repository, self.root)

        with patch.object(self.repository, "log_action", side_effect=RuntimeError("boom")):
            with self.assertRaises(RuntimeError):
                self.repository.delete_nim_period(SummaryDataType.NIM_DN, "2026-01")

        self.assertEqual(self._count_nim_rows(SummaryDataType.NIM_DN, "2026-01"), 1)
        self.assertEqual(len(self.repository.list_batches(SummaryDataType.NIM_DN)), 1)

    def test_delete_nim_period_refreshes_available_periods(self) -> None:
        app = QApplication.instance() or QApplication([])
        _ = app
        self._write_nim_dn_file("5491_FTPLN_20260131.csv", ["5491,2,10,1,CB1,00,1000,,CN"])
        self._write_nim_dn_file("5491_FTPLN_20260228.csv", ["5491,2,10,1,CB1,00,2000,,CN"])
        import_nim_dn(self.repository, self.root)
        tab = NimTab(self.repository, SummaryDataType.NIM_DN)
        try:
            self.repository.delete_nim_period(SummaryDataType.NIM_DN, "2026-01")
            tab.reload()
            values = [tab.period_filter.itemData(index) for index in range(tab.period_filter.count())]
            self.assertNotIn("2026-01", values)
            self.assertIn("2026-02", values)
        finally:
            tab.deleteLater()

    def test_credit_summary_database_created(self) -> None:
        self.assertEqual(self.repository.database_path.name, CREDIT_SUMMARY_DATABASE_NAME)
        self.assertTrue(self.repository.database_path.is_file())
        with closing(sqlite3.connect(self.repository.database_path)) as connection:
            table = connection.execute(
                """
                SELECT 1
                FROM sqlite_master
                WHERE type = 'table' AND name = 'credit_summary_schema_migrations'
                """
            ).fetchone()
        self.assertIsNotNone(table)

    def test_summary_repository_uses_separate_database(self) -> None:
        expected = credit_summary_database_path(self.database_path)
        self.assertEqual(self.repository.database_path, expected)
        self.assertNotEqual(self.repository.database_path, self.database_path)

    def test_migrate_existing_summary_data_preserves_rows(self) -> None:
        legacy_path = self.root / "legacy" / "DuLieuV3.db"
        self._create_legacy_summary_database(legacy_path)

        migrated = SummaryRepository(legacy_path)

        self.assertEqual(migrated.database_path, legacy_path.parent / CREDIT_SUMMARY_DATABASE_NAME)
        self.assertEqual(self._count_nim_rows(SummaryDataType.NIM_DN, "2026-03", repository=migrated), 1)
        self.assertEqual(len(migrated.list_batches(SummaryDataType.NIM_DN)), 1)

    def test_migrate_summary_data_idempotent(self) -> None:
        legacy_path = self.root / "legacy-idempotent" / "DuLieuV3.db"
        self._create_legacy_summary_database(legacy_path)

        first = SummaryRepository(legacy_path)
        second = SummaryRepository(legacy_path)

        self.assertEqual(self._count_nim_rows(SummaryDataType.NIM_DN, "2026-03", repository=first), 1)
        self.assertEqual(self._count_nim_rows(SummaryDataType.NIM_DN, "2026-03", repository=second), 1)

    def test_duplicate_import_hash_is_detected(self) -> None:
        self._write_nim_dn_file("5491_FTPLN_20260131.csv", ["5491,2,10,1,CB1,00,1000,,CN"])
        import_nim_dn(self.repository, self.root)

        with self.assertRaises(SummaryError):
            import_nim_dn(self.repository, self.root)

    def test_nim_import_stores_aggregated_rows_only(self) -> None:
        rows = [
            f"5491,2,10,1,[540000{index % 4:03d}] CBTD {index % 4},00,100,,{'CN' if index % 2 == 0 else 'TC'}"
            for index in range(1000)
        ]
        self._write_nim_dn_file("5491_FTPLN_20260331.csv", rows)

        import_nim_dn(self.repository, self.root)

        self.assertEqual(self._count_raw_nim_rows(SummaryDataType.NIM_DN, "2026-03"), 0)
        self.assertEqual(self._count_nim_rows(SummaryDataType.NIM_DN, "2026-03"), 4)
        with closing(self.repository.connect()) as connection:
            source_rows = int(
                connection.execute(
                    """
                    SELECT SUM(source_row_count)
                    FROM nim_period_summary
                    WHERE data_type = ? AND period = ?
                    """,
                    (SummaryDataType.NIM_DN.value, "2026-03"),
                ).fetchone()[0]
                or 0
            )
        self.assertEqual(source_rows, 1000)

    def test_nim_aggregate_matches_raw_totals(self) -> None:
        self._write_nim_dn_file(
            "5491_FTPLN_20260331.csv",
            [
                "5491,2,10,1,[540000321] Nguyễn Văn A,00,1000,,CN",
                "5491,4,8,1,[540000321] Nguyễn Văn A,00,3000,,TC",
            ],
        )

        import_nim_dn(self.repository, self.root)

        with closing(self.repository.connect()) as connection:
            row = connection.execute(
                """
                SELECT
                    SUM(balance) AS balance,
                    SUM(interest_rate_numerator) AS interest_rate_numerator,
                    SUM(numerator_before) AS numerator_before,
                    SUM(numerator_after) AS numerator_after,
                    SUM(source_row_count) AS source_row_count
                FROM nim_period_summary
                WHERE data_type = ? AND period = ?
                """,
                (SummaryDataType.NIM_DN.value, "2026-03"),
            ).fetchone()
        self.assertAlmostEqual(float(row["balance"]), 4000.0)
        self.assertAlmostEqual(float(row["interest_rate_numerator"]), 34000.0)
        self.assertAlmostEqual(float(row["numerator_before"]), 20000.0)
        self.assertAlmostEqual(float(row["numerator_after"]), 16000.0)
        self.assertEqual(int(row["source_row_count"]), 2)

    def test_all_customer_types_calculated_on_query(self) -> None:
        self._write_nim_dn_file(
            "5491_FTPLN_20260331.csv",
            [
                "5491,2,10,1,[540000321] Nguyễn Văn A,00,1000,,CN",
                "5491,4,8,1,[540000321] Nguyễn Văn A,00,3000,,TC",
            ],
        )
        import_nim_dn(self.repository, self.root)

        with closing(self.repository.connect()) as connection:
            saved_total_rows = int(
                connection.execute(
                    """
                    SELECT COUNT(*)
                    FROM nim_period_summary
                    WHERE customer_type = ?
                    """,
                    ("[Tất cả KH]",),
                ).fetchone()[0]
                or 0
            )
        page = self.repository.query_nim(SummaryDataType.NIM_DN)

        self.assertEqual(saved_total_rows, 0)
        self.assertEqual(page.total_rows, 1)
        self.assertEqual(page.rows[0]["customer_type"], "Tất cả")
        self.assertAlmostEqual(float(page.rows[0]["balance"]), 4000.0)
        self.assertAlmostEqual(float(page.rows[0]["average_rate"]), 8.5)

    def test_weighted_nim_not_average_percent(self) -> None:
        self._write_nim_dn_file(
            "5491_FTPLN_20260331.csv",
            [
                "5491,2,10,0,[540000321] Nguyễn Văn A,00,100,,CN",
                "5491,1,2,0,[540000321] Nguyễn Văn A,00,300,,TC",
            ],
        )
        import_nim_dn(self.repository, self.root)

        row = self.repository.query_nim(SummaryDataType.NIM_DN).rows[0]

        self.assertAlmostEqual(float(row["nim_before"]), 2.75)
        self.assertNotAlmostEqual(float(row["nim_before"]), 4.5)

    def test_existing_raw_migration_preserves_totals(self) -> None:
        legacy_path = self.root / "legacy-raw" / "DuLieuV3.db"
        self._create_legacy_summary_database(legacy_path)
        with closing(sqlite3.connect(legacy_path)) as connection:
            batch_id = int(connection.execute("SELECT id FROM summary_import_history LIMIT 1").fetchone()[0])
            connection.execute(
                """
                INSERT INTO nim_details(
                    batch_id, data_type, period, branch_code, branch_name, trctcd,
                    transaction_office, customer_type, officer, balance, interest_rate,
                    ftp_rate, adjustment_rate, numerator_before, numerator_after,
                    average_rate_numerator, source_file, created_at
                )
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    batch_id,
                    SummaryDataType.NIM_DN.value,
                    "2026-03",
                    "5491",
                    "5491 - CN Lộc Phát",
                    "00",
                    "Hội sở",
                    "Cá nhân (CN)",
                    "CB1",
                    3000,
                    8,
                    4,
                    1,
                    12000,
                    9000,
                    24000,
                    "legacy.csv",
                    "2026-03-31T08:00:00+07:00",
                ),
            )
            connection.commit()

        migrated = SummaryRepository(legacy_path)
        verification = migrated.verify_nim_summary_totals()

        self.assertTrue(verification["matches"])
        self.assertEqual(verification["raw_rows"], 2)
        self.assertEqual(verification["summary_source_rows_for_raw"], 2)
        self.assertEqual(self._count_nim_rows(SummaryDataType.NIM_DN, "2026-03", repository=migrated), 1)

    def test_same_officer_has_one_row_per_customer_type_period(self) -> None:
        self._write_nim_dn_file(
            "5491_FTPLN_20260331.csv",
            [
                "5491,2,10,1,[540000321] Nguyễn Văn A,00,1000,,CN",
                "5491,2,11,1,[540000321] Nguyễn Văn A,00,2000,,CN",
                "5491,4,8,1,[540000321] Nguyễn Văn A,00,3000,,TC",
                "5491,4,9,1,[540000321] Nguyễn Văn A,00,4000,,TC",
            ],
        )
        import_nim_dn(self.repository, self.root)

        with closing(self.repository.connect()) as connection:
            rows = connection.execute(
                """
                SELECT customer_type, COUNT(*) AS count_rows, SUM(source_row_count) AS source_rows
                FROM nim_period_summary
                WHERE data_type = ? AND period = ? AND officer_code = ?
                GROUP BY customer_type
                ORDER BY customer_type
                """,
                (SummaryDataType.NIM_DN.value, "2026-03", "540000321"),
            ).fetchall()

        self.assertEqual([(row["customer_type"], row["count_rows"]) for row in rows], [("Cá nhân (CN)", 1), ("Tổ chức (TC)", 1)])
        self.assertEqual(sum(int(row["source_rows"]) for row in rows), 4)

    def test_dashboard_uses_summary_table(self) -> None:
        self._write_nim_dn_file("5491_FTPLN_20260331.csv", ["5491,2,10,1,[540000321] Nguyễn Văn A,00,1000,,CN"])
        import_nim_dn(self.repository, self.root)

        data = build_nim_dashboard(
            NimDashboardRepository(self.repository),
            SummaryDataType.NIM_DN,
            DashboardFilters(customer_type="Cá nhân (CN)"),
        )

        self.assertEqual(self._count_raw_nim_rows(SummaryDataType.NIM_DN, "2026-03"), 0)
        self.assertAlmostEqual(data.period_rows[0].balance, 1000.0)

    def test_officer_history_uses_summary_table(self) -> None:
        self._write_nim_dn_file("5491_FTPLN_20260331.csv", ["5491,2,10,1,[540000321] Nguyễn Văn A,00,1000,,CN"])
        import_nim_dn(self.repository, self.root)

        history = build_officer_history(
            self.repository,
            SummaryDataType.NIM_DN,
            officer="[540000321] Nguyễn Văn A",
        )

        self.assertEqual(self._count_raw_nim_rows(SummaryDataType.NIM_DN, "2026-03"), 0)
        self.assertEqual(len(history.points), 1)
        self.assertAlmostEqual(history.current_balance, 1000.0)

    def test_export_vba_result_unchanged(self) -> None:
        self._write_nim_dn_file(
            "5491_FTPLN_20260331.csv",
            [
                "5491,2,10,1,[540000321] Nguyễn Văn A,00,1000,,CN",
                "5491,4,8,1,[540000321] Nguyễn Văn A,00,3000,,TC",
            ],
        )
        output = self.root / "BaoCaoNIM_CSDL.xlsx"

        import_nim_dn(self.repository, self.root, export_path=output)

        workbook = load_workbook(output, data_only=True)
        try:
            cache = workbook["Cache_Nim"]
            customer_types = [cache.cell(row_index, 4).value for row_index in range(2, 5)]
            self.assertEqual(customer_types, ["Cá nhân (CN)", "[Tất cả KH]", "Tổ chức (TC)"])
            self.assertAlmostEqual(cache.cell(3, 6).value, 8.5)
            self.assertAlmostEqual(cache.cell(3, 7).value, 5.0)
            self.assertAlmostEqual(cache.cell(3, 8).value, 4.0)
        finally:
            workbook.close()

    def test_delete_period_summary_rows(self) -> None:
        self._write_nim_dn_file("5491_FTPLN_20260331.csv", ["5491,2,10,1,[540000321] Nguyễn Văn A,00,1000,,CN"])
        import_nim_dn(self.repository, self.root)

        info = self.repository.delete_nim_period(SummaryDataType.NIM_DN, "2026-03")

        self.assertEqual(info["row_count"], 1)
        self.assertEqual(self._count_nim_rows(SummaryDataType.NIM_DN, "2026-03"), 0)
        self.assertEqual(self._count_raw_nim_rows(SummaryDataType.NIM_DN, "2026-03"), 0)

    def test_duplicate_file_hash_not_imported_twice(self) -> None:
        self._write_nim_dn_file("5491_FTPLN_20260331.csv", ["5491,2,10,1,[540000321] Nguyễn Văn A,00,1000,,CN"])
        import_nim_dn(self.repository, self.root)

        with self.assertRaises(SummaryError):
            import_nim_dn(self.repository, self.root)

        self.assertEqual(self._count_nim_rows(SummaryDataType.NIM_DN, "2026-03"), 1)
        self.assertEqual(self._count_raw_nim_rows(SummaryDataType.NIM_DN, "2026-03"), 0)

    def test_same_file_hash_different_period_is_allowed(self) -> None:
        self._write_nim_dn_file("5414_FTPLN_20260331.csv", [])
        self._write_nim_dn_file("5414_FTPLN_20260420.csv", [])

        result = import_nim_dn(self.repository, self.root)

        self.assertEqual(result.row_count, 0)
        self.assertEqual(len(self.repository.list_batches(SummaryDataType.NIM_DN)), 2)

    def test_database_optimize_does_not_change_data(self) -> None:
        self._write_nim_dn_file("5491_FTPLN_20260131.csv", ["5491,2,10,1,CB1,00,1000,,CN"])
        import_nim_dn(self.repository, self.root)
        before = self._count_nim_rows(SummaryDataType.NIM_DN, "2026-01")

        result = self.repository.optimize_database(vacuum=True)

        self.assertTrue(result["vacuum"])
        self.assertEqual(self._count_nim_rows(SummaryDataType.NIM_DN, "2026-01"), before)

    def test_summary_backup_restore(self) -> None:
        self._write_nim_dn_file("5491_FTPLN_20260131.csv", ["5491,2,10,1,CB1,00,1000,,CN"])
        import_nim_dn(self.repository, self.root)
        backup = self.repository.backup_database(self.root / "summary-backup.zip")
        with zipfile.ZipFile(backup) as archive:
            self.assertIn(CREDIT_SUMMARY_DATABASE_NAME, archive.namelist())
        self.repository.delete_nim_period(SummaryDataType.NIM_DN, "2026-01")
        self.assertEqual(self._count_nim_rows(SummaryDataType.NIM_DN, "2026-01"), 0)

        self.repository.restore_database(backup)

        self.assertEqual(self._count_nim_rows(SummaryDataType.NIM_DN, "2026-01"), 1)

    def test_loan_compare_customer_code_preserves_leading_zero(self) -> None:
        app = QApplication.instance() or QApplication([])
        _ = app
        self._seed_loan_compare_rows()
        tab = LoanCompareTab(self.repository)
        try:
            tab.reload()
            self.assertEqual(tab.table.item(0, 0).text(), "000123")
        finally:
            tab.deleteLater()

    def test_loan_compare_category_display_is_vietnamese(self) -> None:
        app = QApplication.instance() or QApplication([])
        _ = app
        self._seed_loan_compare_rows()
        tab = LoanCompareTab(self.repository)
        try:
            tab.reload()
            categories = {
                tab.table.item(row, 5).text()
                for row in range(tab.table.rowCount())
            }
            self.assertIn("Khách hàng tất toán", categories)
            self.assertIn("Khách hàng vay tăng", categories)
            labels = [tab.category_filter.itemText(index) for index in range(tab.category_filter.count())]
            self.assertEqual(labels[0], "Tất cả loại")
            self.assertIn("Khách hàng vay tăng", labels)
        finally:
            tab.deleteLater()

    def test_nim_tab_groups_all_customer_types_by_officer_only(self) -> None:
        app = QApplication.instance() or QApplication([])
        _ = app
        dn = self.root / "5491_FTPLN_20260331.csv"
        dn.write_text(
            "\n".join(
                [
                    "BRCD,FTP,INTRT,MUCFTPDC,CBTD,TRCTCD,LDRBAL,TRREF,CUSTTP",
                    "5491,2,10,1,[540000321] Nguyễn Văn A,00,1000,,CN",
                    "5491,4,8,1,[540000321] Nguyễn Văn A,00,3000,,TC",
                ]
            ),
            encoding="utf-8",
        )
        import_nim_dn(self.repository, self.root)

        tab = NimTab(self.repository, SummaryDataType.NIM_DN)
        try:
            all_types = tab.query_page()
            self.assertEqual(all_types.total_rows, 1)
            row = all_types.rows[0]
            self.assertEqual(row["officer"], "[540000321] Nguyễn Văn A")
            self.assertEqual(row["customer_type"], "Tất cả")
            self.assertAlmostEqual(float(row["balance"]), 4000.0)
            self.assertAlmostEqual(float(row["average_rate"]), 8.5)
            self.assertAlmostEqual(float(row["nim_before"]), 5.0)
            self.assertAlmostEqual(float(row["nim_after"]), 4.0)
            self.assertFalse(hasattr(tab, "rate_input"))
            self.assertNotIn("Số dòng", [metric.label for metric in tab.query_dashboard().metrics])
            tab._render_table(all_types.rows)
            tab._restore_column_widths()
            total_width = sum(tab.table.columnWidth(index) for index in range(tab.table.columnCount()))
            self.assertLessEqual(total_width, tab.table.viewport().width() + 2)
            for column in (5, 6, 7, 8):
                self.assertEqual(tab.table.item(0, column).textAlignment(), Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)

            tab.customer_type_filter.setCurrentIndex(tab.customer_type_filter.findData("Cá nhân (CN)"))
            personal = tab.query_page()
            self.assertEqual(personal.total_rows, 1)
            self.assertEqual(personal.rows[0]["customer_type"], "Cá nhân (CN)")
            self.assertAlmostEqual(float(personal.rows[0]["balance"]), 1000.0)
            self.assertAlmostEqual(float(personal.rows[0]["average_rate"]), 10.0)
        finally:
            tab.deleteLater()

    def test_nim_dashboard_weighted_branch_metrics(self) -> None:
        first = self.root / "5491_FTPLN_20260131.csv"
        first.write_text(
            "\n".join(
                [
                    "BRCD,FTP,INTRT,MUCFTPDC,CBTD,TRCTCD,LDRBAL,TRREF,CUSTTP",
                    "5491,2,10,1,[540000321] Nguyễn Văn A,00,100,,CN",
                    "5491,1,2,0.5,[540000322] Nguyễn Văn B,00,300,,CN",
                    "5400,1,5,0,[540000323] Nguyễn Văn C,00,200,,CN",
                ]
            ),
            encoding="utf-8",
        )
        import_nim_dn(self.repository, self.root)

        data = build_nim_dashboard(
            NimDashboardRepository(self.repository),
            SummaryDataType.NIM_DN,
            DashboardFilters(customer_type="Cá nhân (CN)", metric=METRIC_NIM_BEFORE),
        )

        branch_rows = {row.branch: row for row in data.branch_rows}
        self.assertAlmostEqual(branch_rows["5491 - CN Lộc Phát"].nim_before, 2.75)
        self.assertNotAlmostEqual(branch_rows["5491 - CN Lộc Phát"].nim_before, 4.5)
        self.assertEqual(data.kpis[0].label, "Số chi nhánh")
        self.assertEqual(data.kpis[0].value, "2")

    def test_nim_dashboard_all_customer_types_grouped_for_detail(self) -> None:
        path = self.root / "5491_FTPLN_20260131.csv"
        path.write_text(
            "\n".join(
                [
                    "BRCD,FTP,INTRT,MUCFTPDC,CBTD,TRCTCD,LDRBAL,TRREF,CUSTTP",
                    "5491,2,10,1,[540000321] Nguyễn Văn A,00,1000,,CN",
                    "5491,4,8,1,[540000322] Nguyễn Văn B,00,3000,,TC",
                ]
            ),
            encoding="utf-8",
        )
        import_nim_dn(self.repository, self.root)

        data = build_nim_dashboard(
            NimDashboardRepository(self.repository),
            SummaryDataType.NIM_DN,
            DashboardFilters(),
        )

        self.assertEqual(len(data.detail_rows), 1)
        self.assertEqual(data.detail_rows[0].customer_type, "Tất cả")
        self.assertAlmostEqual(data.detail_rows[0].balance, 4000.0)
        self.assertAlmostEqual(data.detail_rows[0].average_rate, 8.5)

    def test_nim_dashboard_window_is_modeless(self) -> None:
        app = QApplication.instance() or QApplication([])
        _ = app
        window = NimDashboardWindow(self.repository, SummaryDataType.NIM_DN)
        try:
            self.assertFalse(window.isModal())
            self.assertEqual(window.windowModality(), Qt.WindowModality.NonModal)
            self.assertEqual(window.tabs.count(), 4)
        finally:
            window.close()
            window.deleteLater()

    def test_export_overview_tab(self) -> None:
        data = self._seed_dashboard_data(DashboardFilters(customer_type="Cá nhân (CN)"))
        output = DashboardNimExportService(data).export_overview_by_period(self.root / "overview.xlsx")

        workbook = load_workbook(output, data_only=True)
        try:
            worksheet = workbook[SHEET_OVERVIEW]
            self.assertEqual([worksheet.cell(1, col).value for col in range(1, 8)], [
                "Kỳ",
                "Tổng dư nợ",
                "Lãi suất bình quân",
                "NIM trước ĐC",
                "NIM sau ĐC",
                "Tăng/giảm dư nợ tuyệt đối",
                "Tăng trưởng dư nợ (%)",
            ])
            self.assertEqual(worksheet.freeze_panes, "A2")
            self.assertEqual(worksheet.auto_filter.ref, worksheet.dimensions)
        finally:
            workbook.close()

    def test_export_branch_comparison_tab(self) -> None:
        data = self._seed_dashboard_data(DashboardFilters(period_from="2026-02", period_to="2026-02", customer_type="Cá nhân (CN)"))
        output = DashboardNimExportService(data, metric=METRIC_NIM_AFTER).export_branch_comparison(self.root / "branch.xlsx")

        workbook = load_workbook(output, data_only=True)
        try:
            worksheet = workbook[SHEET_BRANCH]
            headers = [worksheet.cell(1, col).value for col in range(1, 9)]
            self.assertEqual(headers[-2:], ["Chỉ tiêu đang chọn", "Giá trị chỉ tiêu"])
            branches = [worksheet.cell(row, 2).value for row in range(2, worksheet.max_row + 1)]
            self.assertEqual(set(branches), {"5400 - CN Lâm Đồng", "5491 - CN Lộc Phát"})
            self.assertEqual(worksheet.cell(2, 8).number_format, '0.00"%"')
        finally:
            workbook.close()

    def test_export_growth_tab(self) -> None:
        data = self._seed_dashboard_data(DashboardFilters(customer_type="Cá nhân (CN)"))
        output = DashboardNimExportService(data).export_growth(self.root / "growth.xlsx")

        workbook = load_workbook(output, data_only=True)
        try:
            worksheet = workbook[SHEET_GROWTH]
            self.assertEqual([worksheet.cell(1, col).value for col in range(1, 10)], [
                "Kỳ",
                "Tên chi nhánh",
                "Phòng GD",
                "Loại KH",
                "Dư nợ",
                "Tăng/giảm dư nợ tuyệt đối",
                "Tăng trưởng dư nợ (%)",
                "Biến động NIM trước ĐC",
                "Biến động NIM sau ĐC",
            ])
            self.assertIn("5491 - CN Lộc Phát", [worksheet.cell(row, 2).value for row in range(2, worksheet.max_row + 1)])
        finally:
            workbook.close()

    def test_export_detail_tab(self) -> None:
        data = self._seed_dashboard_data(DashboardFilters(customer_type="Cá nhân (CN)"))
        output = DashboardNimExportService(data).export_detail(self.root / "detail.xlsx")

        workbook = load_workbook(output, data_only=True)
        try:
            worksheet = workbook[SHEET_DETAIL]
            self.assertEqual(worksheet.max_row - 1, len(data.detail_rows))
            self.assertEqual(worksheet.cell(1, 10).value, "Tăng/giảm dư nợ tuyệt đối")
        finally:
            workbook.close()

    def test_export_all_tabs(self) -> None:
        data = self._seed_dashboard_data(DashboardFilters(customer_type="Cá nhân (CN)"))
        output = DashboardNimExportService(data, metric=METRIC_NIM_AFTER).export_all_tabs(self.root / "all.xlsx")

        workbook = load_workbook(output, data_only=True)
        try:
            self.assertEqual(workbook.sheetnames, [SHEET_OVERVIEW, SHEET_BRANCH, SHEET_GROWTH, SHEET_DETAIL])
        finally:
            workbook.close()

    def test_export_uses_current_filters(self) -> None:
        data = self._seed_dashboard_data(
            DashboardFilters(
                period_from="2026-02",
                period_to="2026-02",
                branch="5491 - CN Lộc Phát",
                customer_type="Cá nhân (CN)",
            )
        )
        output = DashboardNimExportService(data).export_all_tabs(self.root / "filtered.xlsx")

        workbook = load_workbook(output, data_only=True)
        try:
            detail = workbook[SHEET_DETAIL]
            self.assertEqual(detail.max_row, 2)
            self.assertEqual(detail["A2"].value, "2026-02")
            self.assertEqual(detail["B2"].value, "5491 - CN Lộc Phát")
            self.assertEqual(detail["D2"].value, "Cá nhân (CN)")
        finally:
            workbook.close()

    def test_branch_chart_uses_full_branch_names(self) -> None:
        data = self._seed_dashboard_data(DashboardFilters(customer_type="Cá nhân (CN)"))
        bars = branch_bar_values(data.branch_rows, METRIC_BALANCE)

        names = [name for _period, name, _value in bars]
        self.assertIn("5400 - CN Lâm Đồng", names)
        self.assertIn("5491 - CN Lộc Phát", names)
        self.assertFalse(any("..." in name for name in names))

    def test_multiline_header_size(self) -> None:
        app = QApplication.instance() or QApplication([])
        _ = app
        window = NimDashboardWindow(self.repository, SummaryDataType.NIM_DN)
        try:
            window.detail_table.setColumnWidth(8, 70)
            header = window.detail_table.horizontalHeader()
            header.refresh_height()
            self.assertGreater(header.height(), 38)
        finally:
            window.close()
            window.deleteLater()

    def test_percentage_export_format(self) -> None:
        data = self._seed_dashboard_data(DashboardFilters(period_from="2026-02", period_to="2026-02", branch="5491 - CN Lộc Phát", customer_type="Cá nhân (CN)"))
        output = DashboardNimExportService(data).export_overview_by_period(self.root / "percent.xlsx")

        workbook = load_workbook(output, data_only=True)
        try:
            worksheet = workbook[SHEET_OVERVIEW]
            self.assertAlmostEqual(worksheet["C2"].value, 8.0)
            self.assertEqual(worksheet["C2"].number_format, '0.00"%"')
            self.assertNotAlmostEqual(worksheet["C2"].value, 0.08)
        finally:
            workbook.close()

    def test_officer_multi_select_combo_suppresses_popup_close_on_row_toggle(self) -> None:
        app = QApplication.instance() or QApplication([])
        _ = app
        combo = OfficerMultiSelectCombo()
        try:
            combo.set_officers([officer_key("[540000321] Nguyễn Văn A"), officer_key("[540000322] Nguyễn Văn B")])
            combo._toggle_item(combo.model().index(1, 0))
            self.assertEqual([item.code for item in combo.selected_officers()], ["540000321"])
            combo._suppress_next_hide = True
            combo.hidePopup()
            self.assertFalse(combo._suppress_next_hide)
        finally:
            combo.deleteLater()

    def test_select_more_than_8_officers_is_allowed(self) -> None:
        officers = self._seed_many_nim_dn_officers(12)
        dialog = self._open_officer_history_dialog(SummaryDataType.NIM_DN, officers[0])
        try:
            dialog.officer_selector.select_all()
            with patch("agribank_v3.features.credit.summary.officer_history.window.QMessageBox.warning") as warning:
                dialog.apply_officer_comparison()
            self.assertFalse(warning.called)
            self.assertEqual(len({row.officer.code for row in dialog.compare_rows}), 12)
            self.assertEqual(dialog.compare_status_label.text(), "Đã chọn 12 cán bộ. Biểu đồ đang hiển thị 8 cán bộ theo từng trang.")
        finally:
            dialog.close()
            dialog.deleteLater()

    def test_officer_chart_pagination(self) -> None:
        officers = self._seed_many_nim_dn_officers(18)
        dialog = self._open_officer_history_dialog(SummaryDataType.NIM_DN, officers[0])
        try:
            dialog.officer_selector.select_all()
            dialog.apply_officer_comparison()
            self.assertEqual(dialog._compare_total_pages(), 3)
            self.assertEqual(len(dialog.officer_compare_chart.series), 8)
            self.assertEqual(dialog.compare_page_label.text(), "Trang 1/3")
            dialog.next_compare_chart_page()
            self.assertEqual(len(dialog.officer_compare_chart.series), 8)
            self.assertEqual(dialog.compare_page_label.text(), "Trang 2/3")
            dialog.next_compare_chart_page()
            self.assertEqual(len(dialog.officer_compare_chart.series), 2)
            self.assertEqual(dialog.compare_page_label.text(), "Trang 3/3")
        finally:
            dialog.close()
            dialog.deleteLater()

    def test_chart_page_does_not_limit_table_data(self) -> None:
        officers = self._seed_many_nim_dn_officers(18)
        dialog = self._open_officer_history_dialog(SummaryDataType.NIM_DN, officers[0])
        try:
            dialog.officer_selector.select_all()
            dialog.apply_officer_comparison()
            self.assertEqual(len(dialog.officer_compare_chart.series), 8)
            self.assertEqual(dialog.officer_compare_table.rowCount(), 18)
            dialog.next_compare_chart_page()
            self.assertEqual(len(dialog.officer_compare_chart.series), 8)
            self.assertEqual(dialog.officer_compare_table.rowCount(), 18)
        finally:
            dialog.close()
            dialog.deleteLater()

    def test_export_contains_all_selected_officers(self) -> None:
        officers = self._seed_many_nim_dn_officers(18)
        dialog = self._open_officer_history_dialog(SummaryDataType.NIM_DN, officers[0])
        try:
            dialog.officer_selector.select_all()
            dialog.apply_officer_comparison()
            dialog.tabs.setCurrentIndex(2)
            tab_key, rows = dialog._current_export_rows()
            output = export_analysis_rows(rows, self.root / "compare_all.xlsx", tab_key=tab_key, metadata=dialog._export_metadata(tab_key))

            workbook = load_workbook(output, data_only=True)
            try:
                worksheet = workbook["SoSanhCBTD"]
                exported = {worksheet.cell(row, 2).value for row in range(2, worksheet.max_row + 1)}
                self.assertEqual(len(exported), 18)
                self.assertIn("ThongTin", workbook.sheetnames)
                self.assertEqual(workbook["ThongTin"]["B3"].value, 18)
            finally:
                workbook.close()
        finally:
            dialog.close()
            dialog.deleteLater()

    def test_top_n_officers(self) -> None:
        officers = self._seed_many_nim_dn_officers(18)
        dialog = self._open_officer_history_dialog(SummaryDataType.NIM_DN, officers[0])
        try:
            dialog.officer_selector.select_all()
            dialog.apply_officer_comparison()
            dialog.compare_mode_combo.setCurrentIndex(dialog.compare_mode_combo.findText("Top N cán bộ"))
            dialog.compare_top_combo.setCurrentIndex(dialog.compare_top_combo.findData(10))
            dialog._render_officer_compare_chart()

            labels = [series.label for series in dialog.officer_compare_chart.series]
            self.assertEqual(len(labels), 10)
            self.assertEqual(labels[0], "Cán bộ 18")
            self.assertEqual(labels[-1], "Cán bộ 09")
        finally:
            dialog.close()
            dialog.deleteLater()

    def test_select_all_officers(self) -> None:
        officers = self._seed_many_nim_dn_officers(18)
        dialog = self._open_officer_history_dialog(SummaryDataType.NIM_DN, officers[0])
        try:
            dialog.officer_selector.select_all()
            self.assertEqual(dialog.officer_selector.selected_count(), 18)
            dialog.apply_officer_comparison()
            self.assertEqual(len({row.officer.code for row in dialog.compare_rows}), 18)
        finally:
            dialog.close()
            dialog.deleteLater()

    def test_officer_popup_stays_open(self) -> None:
        app = QApplication.instance() or QApplication([])
        _ = app
        combo = OfficerMultiSelectCombo()
        try:
            combo.set_officers([officer_key("[540000321] Nguyễn Văn A"), officer_key("[540000322] Nguyễn Văn B")])
            combo._toggle_item(combo.model().index(1, 0))
            combo._suppress_next_hide = True
            combo.hidePopup()
            self.assertFalse(combo._suppress_next_hide)
            self.assertEqual([item.code for item in combo.selected_officers()], ["540000321"])
        finally:
            combo.deleteLater()

    def test_page_navigation_does_not_query_database_again(self) -> None:
        officers = self._seed_many_nim_dn_officers(18)
        dialog = self._open_officer_history_dialog(SummaryDataType.NIM_DN, officers[0])
        try:
            dialog.officer_selector.select_all()
            dialog.apply_officer_comparison()
            with patch("agribank_v3.features.credit.summary.officer_history.window.build_multiple_officer_comparison") as query:
                dialog.next_compare_chart_page()
                dialog.previous_compare_chart_page()
            self.assertFalse(query.called)
        finally:
            dialog.close()
            dialog.deleteLater()

    def test_nim_nv_more_than_8_officers(self) -> None:
        officers = self._seed_many_nim_nv_officers(12)
        dialog = self._open_officer_history_dialog(SummaryDataType.NIM_NV, officers[0])
        try:
            dialog.officer_selector.select_all()
            dialog.apply_officer_comparison()
            self.assertEqual(len({row.officer.code for row in dialog.compare_rows}), 12)
            self.assertEqual(len(dialog.officer_compare_chart.series), 8)
            self.assertEqual(dialog.officer_compare_table.rowCount(), 12)
            self.assertEqual(dialog.compare_page_label.text(), "Trang 1/2")
        finally:
            dialog.close()
            dialog.deleteLater()

    def test_nim_nv_main_table_responsive_columns(self) -> None:
        app = QApplication.instance() or QApplication([])
        _ = app
        self._write_nim_nv_file(
            "5491_FTPDP_20260131.csv",
            ["5491,10,2,1,[540500321] Nguyễn Văn A,00,1000,,CN"],
        )
        import_nim_nv(self.repository, self.root)

        tab = NimTab(self.repository, SummaryDataType.NIM_NV)
        try:
            rows = tab.query_page().rows
            tab._render_table(rows)
            tab._restore_column_widths()
            headers = [tab.table.horizontalHeaderItem(index).text() for index in range(tab.table.columnCount())]
            self.assertEqual(
                headers,
                ["Kỳ", "Tên chi nhánh", "Phòng GD", "Loại KH", "Người quản lý NV", "Số dư nguồn vốn", "NIM trước ĐC", "NIM sau ĐC"],
            )
            self.assertNotIn("Lãi suất bình quân", headers)
            self.assertIn("Mở Dashboard", [button.text() for button in tab.findChildren(QPushButton)])
            self.assertEqual([metric.label for metric in tab.query_dashboard().metrics], ["Tổng nguồn vốn", "NIM trước ĐC", "NIM sau ĐC"])
            self.assertLessEqual(sum(tab.table.columnWidth(index) for index in range(tab.table.columnCount())), tab.table.viewport().width() + 2)
            self.assertEqual(tab.table.item(0, 0).textAlignment(), Qt.AlignmentFlag.AlignCenter)
            for column in (5, 6, 7):
                self.assertEqual(tab.table.item(0, column).textAlignment(), Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        finally:
            tab.deleteLater()

    def test_nim_nv_all_customer_types_grouped_per_officer(self) -> None:
        app = QApplication.instance() or QApplication([])
        _ = app
        self._write_nim_nv_file(
            "5491_FTPDP_20260131.csv",
            [
                "5491,2,10,1,[540500321] Nguyễn Văn A,00,1000,,CN",
                "5491,4,8,1,[540500321] Nguyễn Văn A,00,3000,,TC",
            ],
        )
        import_nim_nv(self.repository, self.root)

        tab = NimTab(self.repository, SummaryDataType.NIM_NV)
        try:
            all_types = tab.query_page()
            self.assertEqual(all_types.total_rows, 1)
            row = all_types.rows[0]
            self.assertEqual(row["officer"], "[540500321] Nguyễn Văn A")
            self.assertEqual(row["customer_type"], "Tất cả")
            self.assertAlmostEqual(float(row["balance"]), 4000.0)
            self.assertAlmostEqual(float(row["nim_before"]), -5.0)
            self.assertAlmostEqual(float(row["nim_after"]), -4.0)
            self.assertNotIn("average_rate", row)

            tab.customer_type_filter.setCurrentIndex(tab.customer_type_filter.findData("Cá nhân (CN)"))
            personal = tab.query_page()
            self.assertEqual(personal.total_rows, 1)
            self.assertEqual(personal.rows[0]["customer_type"], "Cá nhân (CN)")
            self.assertAlmostEqual(float(personal.rows[0]["balance"]), 1000.0)
            self.assertAlmostEqual(float(personal.rows[0]["nim_before"]), -8.0)
            self.assertAlmostEqual(float(personal.rows[0]["nim_after"]), -7.0)
        finally:
            tab.deleteLater()

    def test_nim_nv_officer_history(self) -> None:
        self._write_nim_nv_file(
            "5491_FTPDP_20260131.csv",
            [
                "5491,2,10,1,[540500321] Nguyễn Văn A,00,1000,,CN",
                "5491,4,8,1,[540500321] Nguyễn Văn A,00,3000,,TC",
            ],
        )
        self._write_nim_nv_file(
            "5491_FTPDP_20260228.csv",
            ["5491,4,8,1,[540500321] Nguyễn Văn A,00,2000,,CN"],
        )
        import_nim_nv(self.repository, self.root)

        overview = build_officer_overview(
            self.repository,
            SummaryDataType.NIM_NV,
            officer_code="540500321",
            officer="[540500321] Nguyễn Văn A",
            branch="5491 - CN Lộc Phát",
        )

        self.assertEqual([point.period for point in overview.points], ["2026-01", "2026-02"])
        self.assertEqual(overview.officer.display_name, "Nguyễn Văn A")
        self.assertAlmostEqual(overview.points[0].balance, 4000.0)
        self.assertAlmostEqual(overview.points[0].nim_before, -5.0)
        self.assertAlmostEqual(overview.points[0].nim_after, -4.0)
        self.assertAlmostEqual(overview.current_balance, 2000.0)
        self.assertAlmostEqual(overview.current_nim_before, -4.0)
        self.assertAlmostEqual(overview.current_nim_after, -3.0)

    def test_nim_nv_growth_history(self) -> None:
        history_points = (
            HistoryPoint("2026-01", 1000.0, 0.0, -8.0, -7.0),
            HistoryPoint("2026-02", 1500.0, 0.0, -5.0, -4.0),
            HistoryPoint("2026-03", 1200.0, 0.0, -4.0, -3.0),
        )

        growth = build_officer_growth_history(history_points)

        self.assertIsNone(growth[0].delta)
        self.assertIsNone(growth[0].growth_percent)
        self.assertAlmostEqual(growth[1].delta, 500.0)
        self.assertAlmostEqual(growth[1].growth_percent, 50.0)
        self.assertAlmostEqual(growth[2].delta, -300.0)
        self.assertAlmostEqual(growth[2].growth_percent, -20.0)

    def test_nim_nv_zero_previous_balance(self) -> None:
        history_points = (
            HistoryPoint("2026-01", 0.0, 0.0, 0.0, 0.0),
            HistoryPoint("2026-02", 1000.0, 0.0, -8.0, -7.0),
        )

        growth = build_officer_growth_history(history_points)

        self.assertAlmostEqual(growth[1].delta, 1000.0)
        self.assertIsNone(growth[1].growth_percent)

    def test_nim_nv_multiple_officer_comparison(self) -> None:
        self._write_nim_nv_file(
            "5491_FTPDP_20260131.csv",
            [
                "5491,10,2,1,[540500321] Nguyễn Văn A,00,1000,,CN",
                "5491,8,4,1,[540500322] Nguyễn Văn B,00,3000,,CN",
            ],
        )
        import_nim_nv(self.repository, self.root)

        rows = build_multiple_officer_comparison(
            self.repository,
            SummaryDataType.NIM_NV,
            officers=[
                officer_key("[540500321] Nguyễn Văn A"),
                officer_key("[540500322] Nguyễn Văn B"),
            ],
            metric=METRIC_BALANCE,
            branch="5491 - CN Lộc Phát",
        )

        values = {(row.officer.code, row.period): row.value for row in rows}
        self.assertEqual(values[("540500321", "2026-01")], 1000)
        self.assertEqual(values[("540500322", "2026-01")], 3000)

    def test_nim_nv_branch_weighted_nim(self) -> None:
        self._write_nim_nv_file(
            "5491_FTPDP_20260131.csv",
            [
                "5491,2,10,1,[540500321] Nguyễn Văn A,00,100,,CN",
                "5491,1,2,0.5,[540500322] Nguyễn Văn B,00,300,,CN",
            ],
        )
        import_nim_nv(self.repository, self.root)

        _series, rows = build_officer_branch_comparison(
            self.repository,
            SummaryDataType.NIM_NV,
            officer_code="540500321",
            officer="[540500321] Nguyễn Văn A",
            branch="5491 - CN Lộc Phát",
            metric=METRIC_NIM_BEFORE,
            filters=HistoryFilters(customer_type="Cá nhân (CN)"),
        )

        branch_row = [row for row in rows if row.officer.display_name == "5491 - CN Lộc Phát"][0]
        self.assertAlmostEqual(branch_row.value or 0, -2.75)
        self.assertNotAlmostEqual(branch_row.value or 0, -4.5)

    def test_nim_nv_dashboard_overview(self) -> None:
        data = self._seed_nim_nv_dashboard_data(DashboardFilters(customer_type="Cá nhân (CN)"))

        rows = DashboardNimExportService(data).overview_by_period_rows()

        self.assertEqual(data.ui_config, NIM_NV_UI_CONFIG)
        self.assertEqual([metric.label for metric in data.kpis], ["Số chi nhánh", "Tổng nguồn vốn", "NIM trước ĐC bình quân", "NIM sau ĐC bình quân", "Tăng trưởng nguồn vốn kỳ gần nhất"])
        self.assertEqual(list(rows[0].keys()), ["Kỳ", "Tổng nguồn vốn", "NIM trước ĐC", "NIM sau ĐC", "Tăng/giảm nguồn vốn tuyệt đối", "Tăng trưởng nguồn vốn (%)"])
        self.assertNotIn("Lãi suất bình quân", rows[0])

    def test_nim_nv_dashboard_branch_comparison(self) -> None:
        data = self._seed_nim_nv_dashboard_data(DashboardFilters(period_from="2026-02", period_to="2026-02", customer_type="Cá nhân (CN)", metric=METRIC_NIM_AFTER))
        rows = DashboardNimExportService(data, metric=METRIC_NIM_AFTER).branch_comparison_rows()

        self.assertEqual(list(rows[0].keys()), ["Kỳ", "Tên chi nhánh", "Số dư nguồn vốn", "NIM trước ĐC", "NIM sau ĐC", "Chỉ tiêu đang chọn", "Giá trị chỉ tiêu"])
        self.assertEqual({row["Tên chi nhánh"] for row in rows}, {"5400 - CN Lâm Đồng", "5491 - CN Lộc Phát"})
        self.assertEqual({row["Chỉ tiêu đang chọn"] for row in rows}, {"NIM sau ĐC"})
        self.assertFalse(any("..." in name for _period, name, _value in branch_bar_values(data.branch_rows, METRIC_BALANCE)))

    def test_nim_nv_dashboard_growth(self) -> None:
        data = self._seed_nim_nv_dashboard_data(DashboardFilters(customer_type="Cá nhân (CN)"))

        rows = DashboardNimExportService(data).growth_rows()

        self.assertIn("Số dư nguồn vốn", rows[0])
        self.assertIn("Tăng/giảm nguồn vốn tuyệt đối", rows[0])
        self.assertIn("Tăng trưởng nguồn vốn (%)", rows[0])
        self.assertNotIn("Dư nợ", rows[0])
        growth_values = sorted(row["Tăng trưởng nguồn vốn (%)"] for row in rows if row["Tăng trưởng nguồn vốn (%)"] is not None)
        self.assertEqual(growth_values, [25.0, 50.0])

    def test_nim_nv_dashboard_detail(self) -> None:
        data = self._seed_nim_nv_dashboard_data(DashboardFilters(customer_type="Cá nhân (CN)"))

        rows = DashboardNimExportService(data).detail_rows()

        self.assertEqual(list(rows[0].keys()), ["Kỳ", "Tên chi nhánh", "Phòng GD", "Loại KH", "Số dư nguồn vốn", "NIM trước ĐC", "NIM sau ĐC", "Tăng trưởng nguồn vốn (%)", "Tăng/giảm nguồn vốn tuyệt đối"])
        self.assertNotIn("Lãi suất bình quân", rows[0])
        self.assertEqual({row["Loại KH"] for row in rows}, {"Cá nhân (CN)"})

    def test_nim_nv_export_each_tab(self) -> None:
        data = self._seed_nim_nv_dashboard_data(DashboardFilters(customer_type="Cá nhân (CN)"))
        service = DashboardNimExportService(data, metric=METRIC_NIM_AFTER)

        exports = {
            "overview": service.export_overview_by_period(self.root / "nv_overview.xlsx"),
            "branch": service.export_branch_comparison(self.root / "nv_branch.xlsx"),
            "growth": service.export_growth(self.root / "nv_growth.xlsx"),
            "detail": service.export_detail(self.root / "nv_detail.xlsx"),
        }

        for tab_key, output in exports.items():
            workbook = load_workbook(output, data_only=True)
            try:
                worksheet = workbook[NIM_NV_UI_CONFIG.dashboard_sheets[tab_key]]
                headers = [worksheet.cell(1, column).value for column in range(1, worksheet.max_column + 1)]
                self.assertNotIn("Dư nợ", headers)
                self.assertNotIn("Lãi suất bình quân", headers)
                self.assertEqual(worksheet.freeze_panes, "A2")
                self.assertEqual(worksheet.auto_filter.ref, worksheet.dimensions)
            finally:
                workbook.close()

    def test_nim_nv_export_all_tabs(self) -> None:
        data = self._seed_nim_nv_dashboard_data(DashboardFilters(customer_type="Cá nhân (CN)"))
        output = DashboardNimExportService(data, metric=METRIC_NIM_AFTER).export_all_tabs(self.root / "nim_nv_all.xlsx")

        workbook = load_workbook(output, data_only=True)
        try:
            self.assertEqual(workbook.sheetnames, ["TongQuanNIMNV", "SoSanhChiNhanhNIMNV", "TangTruongNIMNV", "BangDuLieuNIMNV"])
            self.assertEqual(workbook["BangDuLieuNIMNV"].cell(1, 5).value, "Số dư nguồn vốn")
        finally:
            workbook.close()

    def test_nim_nv_multiselect_popup_stays_open(self) -> None:
        app = QApplication.instance() or QApplication([])
        _ = app
        combo = OfficerMultiSelectCombo(placeholder="Chọn cán bộ", counter_label="cán bộ")
        try:
            combo.set_officers([officer_key("[540500321] Nguyễn Văn A"), officer_key("[540500322] Nguyễn Văn B")])
            self.assertEqual(combo.lineEdit().text(), "Chọn cán bộ")
            combo._toggle_item(combo.model().index(1, 0))
            self.assertEqual([item.code for item in combo.selected_officers()], ["540500321"])
            combo._toggle_item(combo.model().index(2, 0))
            self.assertEqual(combo.lineEdit().text(), "2 cán bộ đã chọn")
            combo._suppress_next_hide = True
            combo.hidePopup()
            self.assertFalse(combo._suppress_next_hide)
            self.assertEqual([item.code for item in combo.selected_officers()], ["540500321", "540500322"])
        finally:
            combo.deleteLater()

    def test_nim_nv_multiline_header(self) -> None:
        app = QApplication.instance() or QApplication([])
        _ = app
        window = NimDashboardWindow(self.repository, SummaryDataType.NIM_NV)
        try:
            self.assertEqual(window.windowTitle(), "Dashboard NIM nguồn vốn - AgribankV3")
            headers = [window.detail_table.horizontalHeaderItem(index).text() for index in range(window.detail_table.columnCount())]
            self.assertIn("Số dư nguồn vốn", headers)
            self.assertNotIn("Dư nợ", headers)
            window.detail_table.setColumnWidth(8, 70)
            header = window.detail_table.horizontalHeader()
            header.refresh_height()
            self.assertGreater(header.height(), 38)
        finally:
            window.close()
            window.deleteLater()

    def test_nim_nv_non_modal_windows(self) -> None:
        app = QApplication.instance() or QApplication([])
        _ = app
        self._write_nim_nv_file(
            "5491_FTPDP_20260131.csv",
            ["5491,10,2,1,[540500321] Nguyễn Văn A,00,1000,,CN"],
        )
        import_nim_nv(self.repository, self.root)

        tab = NimTab(self.repository, SummaryDataType.NIM_NV)
        dashboard = NimDashboardWindow(self.repository, SummaryDataType.NIM_NV, parent=tab)
        history = OfficerHistoryDialog(
            self.repository,
            SummaryDataType.NIM_NV,
            officer="[540500321] Nguyễn Văn A",
            branch="5491 - CN Lộc Phát",
            transaction_office="Hội sở",
            customer_type="Cá nhân (CN)",
            parent=tab,
        )
        try:
            tab.open_dashboard()
            self.assertFalse(dashboard.isModal())
            self.assertEqual(dashboard.windowModality(), Qt.WindowModality.NonModal)
            self.assertFalse(history.isModal())
            self.assertEqual(history.windowModality(), Qt.WindowModality.NonModal)
            self.assertEqual(len(tab.dashboard_windows), 1)
            self.assertFalse(tab.dashboard_windows[0].isModal())
        finally:
            for dialog in list(tab.dashboard_windows):
                dialog.close()
                dialog.deleteLater()
            history.close()
            history.deleteLater()
            dashboard.close()
            dashboard.deleteLater()
            tab.deleteLater()

    def test_nim_nv_percentage_not_multiplied_twice(self) -> None:
        self._write_nim_nv_file(
            "5491_FTPDP_20260131.csv",
            ["5491,10,7.684322,0,[540500321] Nguyễn Văn A,00,1000,,CN"],
        )
        import_nim_nv(self.repository, self.root)
        data = build_nim_dashboard(
            NimDashboardRepository(self.repository),
            SummaryDataType.NIM_NV,
            DashboardFilters(customer_type="Cá nhân (CN)"),
        )
        output = DashboardNimExportService(data).export_overview_by_period(self.root / "nv_percent.xlsx")

        workbook = load_workbook(output, data_only=True)
        try:
            worksheet = workbook["TongQuanNIMNV"]
            self.assertAlmostEqual(worksheet["C2"].value, 2.315678)
            self.assertEqual(worksheet["C2"].number_format, '0.00"%"')
            self.assertNotAlmostEqual(worksheet["C2"].value, 0.02315678)
        finally:
            workbook.close()

    def _seed_dashboard_data(self, filters: DashboardFilters) -> object:
        january = self.root / "5491_FTPLN_20260131.csv"
        february = self.root / "5491_FTPLN_20260228.csv"
        january.write_text(
            "\n".join(
                [
                    "BRCD,FTP,INTRT,MUCFTPDC,CBTD,TRCTCD,LDRBAL,TRREF,CUSTTP",
                    "5491,2,10,1,[540000321] Nguyễn Văn A,00,1000,,CN",
                    "5400,1,6,0,[540000322] Nguyễn Văn B,00,2000,,CN",
                    "5491,3,12,1,[540000323] Nguyễn Văn C,00,7000,,TC",
                ]
            ),
            encoding="utf-8",
        )
        february.write_text(
            "\n".join(
                [
                    "BRCD,FTP,INTRT,MUCFTPDC,CBTD,TRCTCD,LDRBAL,TRREF,CUSTTP",
                    "5491,3,8,1,[540000321] Nguyễn Văn A,00,1500,,CN",
                    "5400,1,6,0,[540000322] Nguyễn Văn B,00,2500,,CN",
                    "5491,3,12,1,[540000323] Nguyễn Văn C,00,9000,,TC",
                ]
            ),
            encoding="utf-8",
        )
        import_nim_dn(self.repository, self.root)
        return build_nim_dashboard(NimDashboardRepository(self.repository), SummaryDataType.NIM_DN, filters)

    def _write_nim_dn_file(self, filename: str, rows: list[str]) -> Path:
        path = self.root / filename
        path.write_text(
            "\n".join(["BRCD,FTP,INTRT,MUCFTPDC,CBTD,TRCTCD,LDRBAL,TRREF,CUSTTP", *rows]),
            encoding="utf-8",
        )
        return path

    def _write_nim_nv_file(self, filename: str, rows: list[str]) -> Path:
        path = self.root / filename
        path.write_text(
            "\n".join(["BRCD,FTP,INTRT,MUCFTPDC,CBHD,TRCTCD,LDRBAL,TRREF,CUSTTP", *rows]),
            encoding="utf-8",
        )
        return path

    def _seed_many_nim_dn_officers(self, count: int) -> list[str]:
        officers = [f"[540{i:06d}] Cán bộ {i:02d}" for i in range(1, count + 1)]
        self._write_nim_dn_file(
            "5491_FTPLN_20260131.csv",
            [
                f"5491,2,10,1,{officer},00,{index * 1000},,CN"
                for index, officer in enumerate(officers, start=1)
            ],
        )
        import_nim_dn(self.repository, self.root)
        return officers

    def _seed_many_nim_nv_officers(self, count: int) -> list[str]:
        officers = [f"[550{i:06d}] Cán bộ NV {i:02d}" for i in range(1, count + 1)]
        self._write_nim_nv_file(
            "5491_FTPDP_20260131.csv",
            [
                f"5491,10,2,1,{officer},00,{index * 1000},,CN"
                for index, officer in enumerate(officers, start=1)
            ],
        )
        import_nim_nv(self.repository, self.root)
        return officers

    def _open_officer_history_dialog(self, data_type: SummaryDataType, officer: str) -> OfficerHistoryDialog:
        app = QApplication.instance() or QApplication([])
        _ = app
        return OfficerHistoryDialog(
            self.repository,
            data_type,
            officer=officer,
            branch="5491 - CN Lộc Phát",
            transaction_office="Hội sở",
            customer_type="Cá nhân (CN)",
        )

    def _seed_loan_compare_rows(self) -> None:
        previous = self.root / "loan_previous.csv"
        previous.write_text(
            "\n".join(
                [
                    "CUSTSEQ,CUSTNM,DU_NO,ADDR1,OFFICER_NAME",
                    "000123,Khach tat toan,13500000000,Dia chi rat dai khong lam cot qua rong,CB1",
                    "000124,Khach tang,1000000000,Dia chi 2,CB2",
                ]
            ),
            encoding="utf-8",
        )
        current = self.root / "loan_current.csv"
        current.write_text(
            "\n".join(
                [
                    "CUSTSEQ,CUSTNM,DU_NO,ADDR1,OFFICER_NAME",
                    "000124,Khach tang,1500000000,Dia chi 2,CB2",
                ]
            ),
            encoding="utf-8",
        )
        compare_loan_balances(self.repository, previous, current)

    def _count_nim_rows(
        self,
        data_type: SummaryDataType,
        period: str,
        *,
        repository: SummaryRepository | None = None,
    ) -> int:
        repo = repository or self.repository
        with closing(repo.connect()) as connection:
            return int(
                connection.execute(
                    """
                    SELECT COUNT(*)
                    FROM nim_period_summary
                    WHERE data_type = ? AND period = ?
                    """,
                    (data_type.value, period),
                ).fetchone()[0]
                or 0
            )

    def _count_raw_nim_rows(
        self,
        data_type: SummaryDataType,
        period: str,
        *,
        repository: SummaryRepository | None = None,
    ) -> int:
        repo = repository or self.repository
        with closing(repo.connect()) as connection:
            return int(
                connection.execute(
                    """
                    SELECT COUNT(*)
                    FROM nim_details
                    WHERE data_type = ? AND period = ?
                    """,
                    (data_type.value, period),
                ).fetchone()[0]
                or 0
            )

    def _create_legacy_summary_database(self, database_path: Path) -> None:
        database_path.parent.mkdir(parents=True, exist_ok=True)
        with closing(sqlite3.connect(database_path)) as connection:
            connection.row_factory = sqlite3.Row
            apply_migrations(
                connection,
                (MigrationSpec(version="0.1.6", description="legacy summary schema"),),
                update_root=self.root,
            )
            now = "2026-03-31T08:00:00+07:00"
            cursor = connection.execute(
                """
                INSERT INTO summary_import_history(
                    data_type, period, source_path, file_name, imported_by,
                    row_count, duration_ms, version, status, message, source_hash,
                    created_at, updated_at
                )
                VALUES (?, ?, ?, ?, ?, ?, ?, 1, 'success', ?, ?, ?, ?)
                """,
                (
                    SummaryDataType.NIM_DN.value,
                    "2026-03",
                    "legacy.csv",
                    "legacy.csv",
                    "tester",
                    1,
                    10,
                    "legacy",
                    "legacy-hash",
                    now,
                    now,
                ),
            )
            batch_id = int(cursor.lastrowid)
            connection.execute(
                """
                INSERT INTO nim_details(
                    batch_id, data_type, period, branch_code, branch_name, trctcd,
                    transaction_office, customer_type, officer, balance, interest_rate,
                    ftp_rate, adjustment_rate, numerator_before, numerator_after,
                    average_rate_numerator, source_file, created_at
                )
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    batch_id,
                    SummaryDataType.NIM_DN.value,
                    "2026-03",
                    "5491",
                    "5491 - CN Lộc Phát",
                    "00",
                    "Hội sở",
                    "Cá nhân (CN)",
                    "CB1",
                    1000,
                    10,
                    2,
                    1,
                    8000,
                    7000,
                    10000,
                    "legacy.csv",
                    now,
                ),
            )
            connection.commit()

    def _seed_nim_nv_dashboard_data(self, filters: DashboardFilters) -> object:
        self._write_nim_nv_file(
            "5491_FTPDP_20260131.csv",
            [
                "5491,10,2,1,[540500321] Nguyễn Văn A,00,1000,,CN",
                "5400,6,1,0,[540500322] Nguyễn Văn B,00,2000,,CN",
                "5491,12,3,1,[540500323] Nguyễn Văn C,00,7000,,TC",
            ],
        )
        self._write_nim_nv_file(
            "5491_FTPDP_20260228.csv",
            [
                "5491,8,3,1,[540500321] Nguyễn Văn A,00,1500,,CN",
                "5400,6,1,0,[540500322] Nguyễn Văn B,00,2500,,CN",
                "5491,12,3,1,[540500323] Nguyễn Văn C,00,9000,,TC",
            ],
        )
        import_nim_nv(self.repository, self.root)
        return build_nim_dashboard(NimDashboardRepository(self.repository), SummaryDataType.NIM_NV, filters)

    def test_officer_history_no_data(self) -> None:
        history = build_officer_history(
            self.repository,
            SummaryDataType.NIM_DN,
            officer="[540000999] Không Có",
        )

        self.assertEqual(history.points, ())
        self.assertEqual(history.current_period, "")
        self.assertEqual(history.current_balance, 0)

    def test_officer_history_single_period(self) -> None:
        path = self.root / "5491_FTPLN_20260131.csv"
        path.write_text(
            "\n".join(
                [
                    "BRCD,FTP,INTRT,MUCFTPDC,CBTD,TRCTCD,LDRBAL,TRREF,CUSTTP",
                    "5491,2,10,1,[540000321] Nguyễn Văn A,00,1000,,CN",
                ]
            ),
            encoding="utf-8",
        )
        import_nim_dn(self.repository, self.root)

        rows = self.repository.get_officer_history(SummaryDataType.NIM_DN, officer="[540000321] Nguyễn Văn A")
        history = build_officer_history(
            self.repository,
            SummaryDataType.NIM_DN,
            officer="[540000321] Nguyễn Văn A",
        )

        self.assertEqual(len(rows), 1)
        self.assertEqual(len(history.points), 1)
        self.assertEqual(history.current_period, "2026-01")
        self.assertAlmostEqual(history.current_balance, 1000.0)
        self.assertAlmostEqual(history.current_average_rate, 10.0)
        self.assertAlmostEqual(history.current_nim_before, 8.0)
        self.assertAlmostEqual(history.current_nim_after, 7.0)

    def test_officer_history_multiple_periods_sorted_and_weighted(self) -> None:
        january = self.root / "5491_FTPLN_20260131.csv"
        february = self.root / "5491_FTPLN_20260228.csv"
        january.write_text(
            "\n".join(
                [
                    "BRCD,FTP,INTRT,MUCFTPDC,CBTD,TRCTCD,LDRBAL,TRREF,CUSTTP",
                    "5491,2,10,1,[540000321] Nguyễn Văn A,00,1000,,CN",
                    "5491,4,8,1,[540000321] Nguyễn Văn A,00,3000,,TC",
                ]
            ),
            encoding="utf-8",
        )
        february.write_text(
            "\n".join(
                [
                    "BRCD,FTP,INTRT,MUCFTPDC,CBTD,TRCTCD,LDRBAL,TRREF,CUSTTP",
                    "5491,4,8,1,[540000321] Nguyễn Văn A,00,2000,,CN",
                ]
            ),
            encoding="utf-8",
        )
        import_nim_dn(self.repository, self.root)

        history = build_officer_history(
            self.repository,
            SummaryDataType.NIM_DN,
            officer="[540000321] Nguyễn Văn A",
            branch="5491 - CN Lộc Phát",
            transaction_office="Hội sở",
        )

        self.assertEqual([point.period for point in history.points], ["2026-01", "2026-02"])
        self.assertAlmostEqual(history.points[0].balance, 4000.0)
        self.assertAlmostEqual(history.points[0].average_rate, 8.5)
        self.assertAlmostEqual(history.points[0].nim_before, 5.0)
        self.assertAlmostEqual(history.points[0].nim_after, 4.0)
        self.assertAlmostEqual(history.points[1].balance, 2000.0)
        self.assertAlmostEqual(history.current_nim_after, 3.0)

    def test_officer_history_export_excel(self) -> None:
        path = self.root / "5491_FTPLN_20260131.csv"
        path.write_text(
            "\n".join(
                [
                    "BRCD,FTP,INTRT,MUCFTPDC,CBTD,TRCTCD,LDRBAL,TRREF,CUSTTP",
                    "5491,2,10,1,[540000321] Nguyễn Văn A,00,1000,,CN",
                ]
            ),
            encoding="utf-8",
        )
        import_nim_dn(self.repository, self.root)
        history = build_officer_history(
            self.repository,
            SummaryDataType.NIM_DN,
            officer="[540000321] Nguyễn Văn A",
            branch="5491 - CN Lộc Phát",
            transaction_office="Hội sở",
            customer_type="Cá nhân (CN)",
        )

        output_path = export_officer_history_excel(history, self.root / "history.xlsx")

        workbook = load_workbook(output_path, data_only=True)
        try:
            worksheet = workbook["LichSu_NIM_CBTD"]
            self.assertEqual(worksheet["A1"].value, "Lịch sử NIM CBTD")
            self.assertEqual(worksheet["B3"].value, "[540000321] Nguyễn Văn A")
            self.assertEqual([worksheet.cell(8, col).value for col in range(1, 6)], ["Kỳ", "Dư nợ", "Lãi suất bình quân", "NIM trước ĐC", "NIM sau ĐC"])
            self.assertEqual(worksheet["A9"].value, "2026-01")
            self.assertEqual(worksheet["B9"].value, 1000)
            self.assertEqual(worksheet["C9"].value, 10)
            self.assertEqual(worksheet["C9"].number_format, '0.00"%"')
        finally:
            workbook.close()

    def test_officer_overview_has_nim_before_and_after(self) -> None:
        january = self.root / "5491_FTPLN_20260131.csv"
        february = self.root / "5491_FTPLN_20260228.csv"
        january.write_text(
            "\n".join(
                [
                    "BRCD,FTP,INTRT,MUCFTPDC,CBTD,TRCTCD,LDRBAL,TRREF,CUSTTP",
                    "5491,2,10,1,[540000321] Nguyễn Văn A,00,1000,,CN",
                ]
            ),
            encoding="utf-8",
        )
        february.write_text(
            "\n".join(
                [
                    "BRCD,FTP,INTRT,MUCFTPDC,CBTD,TRCTCD,LDRBAL,TRREF,CUSTTP",
                    "5491,4,8,1,[540000321] Nguyễn Văn A,00,2000,,CN",
                ]
            ),
            encoding="utf-8",
        )
        import_nim_dn(self.repository, self.root)

        overview = build_officer_overview(
            self.repository,
            SummaryDataType.NIM_DN,
            officer_code="540000321",
            officer="[540000321] Nguyễn Văn A",
            branch="5491 - CN Lộc Phát",
        )

        self.assertEqual(overview.current_period, "2026-02")
        self.assertAlmostEqual(overview.current_nim_before, 4.0)
        self.assertAlmostEqual(overview.current_nim_after, 3.0)
        self.assertAlmostEqual(overview.current_average_rate, 8.0)

    def test_officer_growth_history(self) -> None:
        points = (
            build_officer_overview(
                self.repository,
                SummaryDataType.NIM_DN,
                officer="[none]",
            ).points
        )
        self.assertEqual(points, ())
        from agribank_v3.features.credit.summary.officer_history.models import HistoryPoint

        history_points = (
            HistoryPoint("2026-01", 1000.0, 10.0, 8.0, 7.0),
            HistoryPoint("2026-02", 1500.0, 9.0, 5.0, 4.0),
            HistoryPoint("2026-03", 1200.0, 8.0, 4.0, 3.0),
        )

        growth = build_officer_growth_history(history_points)

        self.assertIsNone(growth[0].delta)
        self.assertIsNone(growth[0].growth_percent)
        self.assertAlmostEqual(growth[1].delta, 500.0)
        self.assertAlmostEqual(growth[1].growth_percent, 50.0)
        self.assertAlmostEqual(growth[2].delta, -300.0)
        self.assertAlmostEqual(growth[2].growth_percent, -20.0)

    def test_officer_growth_zero_previous_balance(self) -> None:
        from agribank_v3.features.credit.summary.officer_history.models import HistoryPoint

        history_points = (
            HistoryPoint("2026-01", 0.0, 0.0, 0.0, 0.0),
            HistoryPoint("2026-02", 1000.0, 10.0, 8.0, 7.0),
        )

        growth = build_officer_growth_history(history_points)

        self.assertAlmostEqual(growth[1].delta, 1000.0)
        self.assertIsNone(growth[1].growth_percent)

    def test_multiple_officer_comparison(self) -> None:
        path = self.root / "5491_FTPLN_20260131.csv"
        path.write_text(
            "\n".join(
                [
                    "BRCD,FTP,INTRT,MUCFTPDC,CBTD,TRCTCD,LDRBAL,TRREF,CUSTTP",
                    "5491,2,10,1,[540000321] Nguyễn Văn A,00,1000,,CN",
                    "5491,4,8,1,[540000322] Nguyễn Văn B,00,3000,,CN",
                ]
            ),
            encoding="utf-8",
        )
        import_nim_dn(self.repository, self.root)

        rows = build_multiple_officer_comparison(
            self.repository,
            SummaryDataType.NIM_DN,
            officers=[
                officer_key("[540000321] Nguyễn Văn A"),
                officer_key("[540000322] Nguyễn Văn B"),
            ],
            metric=METRIC_BALANCE,
            branch="5491 - CN Lộc Phát",
        )

        values = {(row.officer.code, row.period): row.value for row in rows}
        self.assertEqual(values[("540000321", "2026-01")], 1000)
        self.assertEqual(values[("540000322", "2026-01")], 3000)

    def test_duplicate_officer_names_use_officer_code(self) -> None:
        path = self.root / "5491_FTPLN_20260131.csv"
        path.write_text(
            "\n".join(
                [
                    "BRCD,FTP,INTRT,MUCFTPDC,CBTD,TRCTCD,LDRBAL,TRREF,CUSTTP",
                    "5491,2,10,1,[540000321] Nguyễn Văn A,00,1000,,CN",
                    "5491,4,8,1,[540000999] Nguyễn Văn A,00,3000,,CN",
                ]
            ),
            encoding="utf-8",
        )
        import_nim_dn(self.repository, self.root)

        first = build_officer_overview(
            self.repository,
            SummaryDataType.NIM_DN,
            officer_code="540000321",
            officer="[540000321] Nguyễn Văn A",
            branch="5491 - CN Lộc Phát",
        )
        second = build_officer_overview(
            self.repository,
            SummaryDataType.NIM_DN,
            officer_code="540000999",
            officer="[540000999] Nguyễn Văn A",
            branch="5491 - CN Lộc Phát",
        )

        self.assertEqual(first.officer.display_name, "Nguyễn Văn A")
        self.assertEqual(second.officer.display_name, "Nguyễn Văn A")
        self.assertAlmostEqual(first.current_balance, 1000.0)
        self.assertAlmostEqual(second.current_balance, 3000.0)
        comparison = build_multiple_officer_comparison(
            self.repository,
            SummaryDataType.NIM_DN,
            officers=[
                officer_key("[540000321] Nguyễn Văn A"),
                officer_key("[540000999] Nguyễn Văn A"),
            ],
            metric=METRIC_BALANCE,
            branch="5491 - CN Lộc Phát",
        )
        self.assertEqual(sorted(row.officer.code for row in comparison), ["540000321", "540000999"])

    def test_branch_nim_uses_weighted_formula(self) -> None:
        path = self.root / "5491_FTPLN_20260131.csv"
        path.write_text(
            "\n".join(
                [
                    "BRCD,FTP,INTRT,MUCFTPDC,CBTD,TRCTCD,LDRBAL,TRREF,CUSTTP",
                    "5491,2,10,1,[540000321] Nguyễn Văn A,00,100,,CN",
                    "5491,1,2,0.5,[540000322] Nguyễn Văn B,00,300,,CN",
                ]
            ),
            encoding="utf-8",
        )
        import_nim_dn(self.repository, self.root)

        _series, rows = build_officer_branch_comparison(
            self.repository,
            SummaryDataType.NIM_DN,
            officer_code="540000321",
            officer="[540000321] Nguyễn Văn A",
            branch="5491 - CN Lộc Phát",
            metric=METRIC_NIM_BEFORE,
            filters=HistoryFilters(customer_type="Cá nhân (CN)"),
        )

        branch_row = [row for row in rows if row.officer.display_name == "5491 - CN Lộc Phát"][0]
        self.assertAlmostEqual(branch_row.value or 0, 2.75)
        self.assertNotAlmostEqual(branch_row.value or 0, 4.5)

    def test_branch_average_rate_uses_weighted_formula(self) -> None:
        path = self.root / "5491_FTPLN_20260131.csv"
        path.write_text(
            "\n".join(
                [
                    "BRCD,FTP,INTRT,MUCFTPDC,CBTD,TRCTCD,LDRBAL,TRREF,CUSTTP",
                    "5491,2,10,1,[540000321] Nguyễn Văn A,00,100,,CN",
                    "5491,1,2,0.5,[540000322] Nguyễn Văn B,00,300,,CN",
                ]
            ),
            encoding="utf-8",
        )
        import_nim_dn(self.repository, self.root)

        _series, rows = build_officer_branch_comparison(
            self.repository,
            SummaryDataType.NIM_DN,
            officer_code="540000321",
            officer="[540000321] Nguyễn Văn A",
            branch="5491 - CN Lộc Phát",
            metric=METRIC_AVERAGE_RATE,
            filters=HistoryFilters(customer_type="Cá nhân (CN)"),
        )

        branch_row = [row for row in rows if row.officer.display_name == "5491 - CN Lộc Phát"][0]
        self.assertAlmostEqual(branch_row.value or 0, 4.0)
        self.assertNotAlmostEqual(branch_row.value or 0, 6.0)

    def test_all_customer_types_grouped_per_officer_period(self) -> None:
        path = self.root / "5491_FTPLN_20260131.csv"
        path.write_text(
            "\n".join(
                [
                    "BRCD,FTP,INTRT,MUCFTPDC,CBTD,TRCTCD,LDRBAL,TRREF,CUSTTP",
                    "5491,2,10,1,[540000321] Nguyễn Văn A,00,1000,,CN",
                    "5491,4,8,1,[540000321] Nguyễn Văn A,00,3000,,TC",
                ]
            ),
            encoding="utf-8",
        )
        import_nim_dn(self.repository, self.root)

        rows = OfficerHistoryRepository(self.repository).get_multiple_officer_history(
            SummaryDataType.NIM_DN,
            officers=[officer_key("[540000321] Nguyễn Văn A")],
            branch="5491 - CN Lộc Phát",
            filters=HistoryFilters(),
        )

        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0]["customer_type"], "Tất cả")
        self.assertAlmostEqual(float(rows[0]["balance"]), 4000.0)

    def test_export_each_analysis_tab(self) -> None:
        sample_rows = [{"Kỳ": "2026-01", "Giá trị": 1.23}]

        for tab_key, sheet_name in SHEET_NAMES.items():
            output = export_analysis_rows(sample_rows, self.root / f"{tab_key}.xlsx", tab_key=tab_key)
            workbook = load_workbook(output, data_only=True)
            try:
                self.assertEqual(workbook.sheetnames, [sheet_name])
                self.assertEqual(workbook[sheet_name]["A2"].value, "2026-01")
            finally:
                workbook.close()

    def test_loan_compare_preserves_vba_classification(self) -> None:
        previous = self.root / "2003.csv"
        previous.write_text(
            "\n".join(
                [
                    "CUSTSEQ,CUSTNM,DU_NO,ADDR1,OFFICER_NAME",
                    "B,Khach B,100,,CB1",
                    "A,Khach A,100,,CB1",
                    "C,Khach C,100,,CB2",
                ]
            ),
            encoding="utf-8",
        )
        current = self.root / "2010.csv"
        current.write_text(
            "\n".join(
                [
                    "CUSTSEQ,CUSTNM,DU_NO,ADDR1,OFFICER_NAME",
                    "A,Khach A,150,,CB1",
                    "B,Khach B,0,,CB1",
                    "D,Khach D,70,,CB3",
                ]
            ),
            encoding="utf-8",
        )

        output_path = self.root / "BaoCaoTangGiamKH.xlsx"
        result = compare_loan_balances(self.repository, previous, current, export_path=output_path)

        self.assertEqual(result.row_count, 4)
        self.assertEqual(result.output_path, output_path)
        rows = {row["customer_code"]: row for row in self.repository.query_loan_compare(page_size=20).rows}
        self.assertEqual(rows["A"]["category"], "Khach hang vay tang")
        self.assertEqual(rows["B"]["category"], "Khach hang tat toan")
        self.assertEqual(rows["C"]["category"], "Khach hang tat toan")
        self.assertEqual(rows["D"]["category"], "Khach hang vay moi")
        workbook = load_workbook(output_path, data_only=True)
        try:
            self.assertEqual(workbook.sheetnames, ["SoSanh_DuNo", "Sheet2", "Sheet3"])
            worksheet = workbook["SoSanh_DuNo"]
            self.assertEqual([worksheet.cell(1, col).value for col in range(1, 7)], list(LOAN_COMPARE_VBA_HEADERS))
            self.assertEqual([worksheet.cell(row, 1).value for row in range(2, 6)], ["B", "A", "C", "D"])
            self.assertEqual(worksheet["C2"].number_format, "#,##0")
            self.assertEqual(worksheet["D2"].number_format, "#,##0")
        finally:
            workbook.close()

    def test_credit_limit_import_filters_ln01_like_vba(self) -> None:
        path = self.root / "5491_ln01_20260427.csv"
        headers = [f"H{i}" for i in range(63)]
        headers[0] = "BRCD"
        headers[1] = "CUSTSEQ"
        headers[2] = "CUSTNM"
        headers[3] = "TAI_KHOAN"
        headers[62] = "CREDIT_LINE_YPE"
        expired = [""] * 63
        expired[1] = "KH01"
        expired[2] = "Khach 01"
        expired[5] = "100"
        expired[14] = "HD01"
        expired[15] = "20240101"
        expired[17] = "1000"
        expired[18] = "20240115"
        expired[27] = "CB1"
        expired[35] = "Dia chi"
        expired[62] = "Line of Credit"
        soon = expired.copy()
        soon[1] = "KH02"
        soon[14] = "HD02"
        soon[18] = "20240210"
        path.write_text(
            ",".join(headers) + "\n" + ",".join(expired) + "\n" + ",".join(soon),
            encoding="utf-8",
        )

        output_path = self.root / "BaoCaoHanMucHetHan.xlsx"
        result = import_credit_limit_file(
            self.repository,
            path,
            warn_days=30,
            reference_date=date(2024, 1, 20),
            export_path=output_path,
        )

        self.assertEqual(result.row_count, 2)
        self.assertEqual(result.output_path, output_path)
        rows = {row["contract_number"]: row for row in self.repository.query_credit_limits(page_size=20).rows}
        self.assertEqual(rows["HD01"]["status"], "Đã hết hạn")
        self.assertEqual(rows["HD02"]["status"], "Sắp hết hạn")
        workbook = load_workbook(output_path, data_only=True)
        try:
            self.assertEqual(workbook.sheetnames, ["HanMuc_HetHan"])
            worksheet = workbook["HanMuc_HetHan"]
            self.assertEqual([worksheet.cell(1, col).value for col in range(1, 11)], list(CREDIT_LIMIT_VBA_HEADERS))
            self.assertEqual(worksheet["D2"].number_format, "dd/mm/yyyy")
            self.assertEqual(worksheet["E2"].number_format, "#,##0")
            self.assertEqual(worksheet["G2"].number_format, "dd/mm/yyyy")
            self.assertEqual(worksheet["J2"].value, "Hợp đồng hạn mức tín dụng ã quá hạn đến thời đểm hiện tại")
            self.assertEqual(worksheet["J3"].value, "Hợp đồng tín dụng dến hạn trong vòng 30 ngày tới theo thời đểm hiện tại")
        finally:
            workbook.close()

    def test_export_rows_supports_excel_csv_pdf(self) -> None:
        rows = [{"A": "x", "B": 1}]
        xlsx = export_rows(rows, self.root / "out.xlsx", title="Test")
        csv_path = export_rows(rows, self.root / "out.csv", title="Test")
        pdf = export_rows(rows, self.root / "out.pdf", title="Test")

        self.assertTrue(xlsx.is_file())
        self.assertTrue(csv_path.read_text(encoding="utf-8-sig").startswith("A,B"))
        self.assertTrue(pdf.read_bytes().startswith(b"%PDF"))

    def test_backup_and_restore_roundtrip(self) -> None:
        backup = self.repository.backup_database(self.root / "backup.zip")

        self.assertTrue(backup.is_file())
        safety = self.repository.restore_database(backup)
        self.assertTrue(safety.is_file())

    def test_workbook_compare_reports_cell_differences(self) -> None:
        from openpyxl import Workbook

        vba = self.root / "vba.xlsx"
        python = self.root / "python.xlsx"
        for path, value in ((vba, "A"), (python, "B")):
            workbook = Workbook()
            worksheet = workbook.active
            worksheet.title = "SoSanh_DuNo"
            worksheet["A1"] = value
            workbook.save(path)

        result = compare_workbooks(vba, python, self.root / "diff.xlsx")

        self.assertFalse(result.same)
        self.assertGreaterEqual(result.diff_count, 1)
        self.assertTrue(result.report_path.is_file())


if __name__ == "__main__":
    unittest.main()
