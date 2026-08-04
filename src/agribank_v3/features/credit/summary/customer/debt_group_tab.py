from __future__ import annotations

from dataclasses import replace
from pathlib import Path

from PySide6.QtCore import Qt
from PySide6.QtWidgets import QFileDialog, QGridLayout, QMessageBox, QTabWidget, QVBoxLayout, QWidget

from agribank_v3.features.credit.summary.customer.async_query import AsyncQueryController
from agribank_v3.features.credit.summary.customer.charts import CustomerDonutChart, CustomerLineChart
from agribank_v3.features.credit.summary.customer.customer_detail_window import CustomerDetailWindow
from agribank_v3.features.credit.summary.customer.debt_group_service import build_debt_group_payload
from agribank_v3.features.credit.summary.customer.export_service import (
    DEBT_GROUP_BRANCH_COLUMNS,
    DEBT_GROUP_CUSTOMER_COLUMNS,
    DEBT_GROUP_OFFICER_COLUMNS,
    DEBT_GROUP_SUMMARY_COLUMNS,
    export_debt_group_analysis,
    suggested_customer_export_name,
)
from agribank_v3.features.credit.summary.customer.filters import DEBT_GROUP_FILTERS, CustomerFilters
from agribank_v3.features.credit.summary.customer.formatters import format_money_vn, format_percent_vn
from agribank_v3.features.credit.summary.customer.repository import CustomerRepository
from agribank_v3.features.credit.summary.customer.table_models import CustomerTableModel
from agribank_v3.features.credit.summary.customer.widgets import (
    CompactToolbar,
    CustomerTableView,
    KpiMetric,
    MetricGrid,
    Pager,
    QueryStateBanner,
    combo_box,
    current_data,
    populate_combo,
    secondary_button,
)


DEBT_GROUP_KPI_LABELS = (
    "Tổng dư nợ",
    "Dư nợ nhóm 1",
    "Nợ cần chú ý",
    "Nợ xấu",
    "Dư nợ chưa xác định nhóm",
    "Tỷ lệ nợ cần chú ý",
    "Tỷ lệ nợ xấu",
    "KH có nợ cần chú ý",
    "KH có nợ xấu",
    "KH có nhóm nợ UNKNOWN",
)


class DebtGroupAnalysisTab(QWidget):
    def __init__(self, repository: CustomerRepository, filters_provider, parent=None) -> None:
        super().__init__(parent)
        self.repository = repository
        self.filters_provider = filters_provider
        self.branch_page = 1
        self.branch_page_size = 100
        self.branch_sort_by = "bad_debt_ratio"
        self.branch_sort_desc = True
        self.officer_page = 1
        self.officer_page_size = 100
        self.officer_sort_by = "bad_debt_ratio"
        self.officer_sort_desc = True
        self.customer_page = 1
        self.customer_page_size = 100
        self.customer_sort_by = "bad_debt_ratio"
        self.customer_sort_desc = True
        self.detail_windows: list[CustomerDetailWindow] = []
        self.query_controller = AsyncQueryController(self, max_cache_entries=24)
        self.export_controller = AsyncQueryController(self, max_cache_entries=1)

        layout = QVBoxLayout(self)
        layout.setSpacing(6)
        toolbar = CompactToolbar()
        self.debt_group_combo = combo_box("Nhóm nợ", minimum_width=180, maximum_width=280)
        populate_combo(self.debt_group_combo, DEBT_GROUP_FILTERS[1:])
        self.debt_group_combo.currentIndexChanged.connect(self._filter_changed)
        refresh_button = secondary_button("Làm mới")
        refresh_button.clicked.connect(lambda: self.refresh(use_cache=False))
        self.export_button = secondary_button("Xuất Excel")
        self.export_button.clicked.connect(self.export_excel)
        for widget in (self.debt_group_combo, refresh_button, self.export_button):
            toolbar.addWidget(widget)
        layout.addWidget(toolbar)

        self.state_banner = QueryStateBanner()
        self.state_banner.retryRequested.connect(lambda: self.refresh(use_cache=False))
        layout.addWidget(self.state_banner)

        self.tabs = QTabWidget()
        layout.addWidget(self.tabs, stretch=1)
        self.metrics = MetricGrid()
        self.structure_chart = CustomerDonutChart("Cơ cấu dư nợ theo nhóm nợ", value_kind="money")
        self.trend_chart = CustomerLineChart("Xu hướng nợ cần chú ý và nợ xấu", value_kind="percent")
        self.summary_model = CustomerTableModel(DEBT_GROUP_SUMMARY_COLUMNS, self)
        self.branch_model = CustomerTableModel(DEBT_GROUP_BRANCH_COLUMNS, self)
        self.officer_model = CustomerTableModel(DEBT_GROUP_OFFICER_COLUMNS, self)
        self.customer_model = CustomerTableModel(DEBT_GROUP_CUSTOMER_COLUMNS, self)
        self._build_overview_tab()
        self.branch_table, self.branch_pager = self._build_table_tab("Theo chi nhánh", self.branch_model)
        self.officer_table, self.officer_pager = self._build_table_tab("Theo cán bộ", self.officer_model)
        self.customer_table, self.customer_pager = self._build_table_tab("Theo khách hàng", self.customer_model)
        self.branch_table.horizontalHeader().sectionClicked.connect(lambda section: self._sort_changed("branch", section))
        self.officer_table.horizontalHeader().sectionClicked.connect(lambda section: self._sort_changed("officer", section))
        self.customer_table.horizontalHeader().sectionClicked.connect(lambda section: self._sort_changed("customer", section))
        self.customer_table.doubleClicked.connect(self._open_customer_detail)
        self.branch_pager.pageChanged.connect(lambda page: self._page_changed("branch", page))
        self.branch_pager.pageSizeChanged.connect(lambda size: self._page_size_changed("branch", size))
        self.officer_pager.pageChanged.connect(lambda page: self._page_changed("officer", page))
        self.officer_pager.pageSizeChanged.connect(lambda size: self._page_size_changed("officer", size))
        self.customer_pager.pageChanged.connect(lambda page: self._page_changed("customer", page))
        self.customer_pager.pageSizeChanged.connect(lambda size: self._page_size_changed("customer", size))
        self.branch_table.apply_default_widths((54, 100, 180, 130, 110, 110, 110, 110, 110, 120, 100, 100, 110, 110, 105, 105, 105))
        self.officer_table.apply_default_widths((54, 110, 180, 160, 130, 130, 105, 105, 105, 105, 105, 120, 100, 100, 105, 105, 105, 105, 105))
        self.customer_table.apply_default_widths((54, 80, 120, 220, 95, 160, 180, 130, 120, 110, 110, 110, 110, 110, 120, 100, 100, 105, 105, 105))

    def refresh(self, *args, use_cache: bool = True) -> None:
        if not self.repository.has_period_data():
            self.set_empty_state()
            return
        filters = self._filters()
        report_period = filters.current_period or filters.period_to or ""
        cache_key = (
            "debt_group",
            filters,
            report_period,
            self.branch_page,
            self.branch_page_size,
            self.branch_sort_by,
            self.branch_sort_desc,
            self.officer_page,
            self.officer_page_size,
            self.officer_sort_by,
            self.officer_sort_desc,
            self.customer_page,
            self.customer_page_size,
            self.customer_sort_by,
            self.customer_sort_desc,
        )
        self.query_controller.run(
            "customer_debt_group",
            lambda: build_debt_group_payload(
                self.repository,
                report_period,
                filters,
                branch_page=self.branch_page,
                branch_page_size=self.branch_page_size,
                branch_sort_by=self.branch_sort_by,
                branch_sort_desc=self.branch_sort_desc,
                officer_page=self.officer_page,
                officer_page_size=self.officer_page_size,
                officer_sort_by=self.officer_sort_by,
                officer_sort_desc=self.officer_sort_desc,
                customer_page=self.customer_page,
                customer_page_size=self.customer_page_size,
                customer_sort_by=self.customer_sort_by,
                customer_sort_desc=self.customer_sort_desc,
            ),
            self._apply_payload,
            self._query_failed,
            cache_key=cache_key,
            use_cache=use_cache,
            state_callback=self._state_changed,
        )

    def export_excel(self) -> None:
        if not self.repository.has_period_data():
            self.state_banner.set_empty("Chưa có dữ liệu để xuất.")
            return
        path, _ = QFileDialog.getSaveFileName(
            self,
            "Xuất phân tích nhóm nợ",
            suggested_customer_export_name("PhanTichNhomNo"),
            "Excel (*.xlsx)",
        )
        if not path:
            return
        filters = self._filters()
        report_period = filters.current_period or filters.period_to or ""
        self.export_button.setEnabled(False)
        self.export_controller.run(
            "customer_debt_group_export",
            lambda: export_debt_group_analysis(
                self.repository,
                filters,
                Path(path),
                report_period=report_period,
                branch_sort_by=self.branch_sort_by,
                branch_sort_desc=self.branch_sort_desc,
                officer_sort_by=self.officer_sort_by,
                officer_sort_desc=self.officer_sort_desc,
                customer_sort_by=self.customer_sort_by,
                customer_sort_desc=self.customer_sort_desc,
            ),
            self._export_finished,
            self._export_failed,
            cache_key=("debt_group_export", path, filters, report_period),
            use_cache=False,
            state_callback=self._export_state_changed,
        )

    def invalidate_cache(self) -> None:
        self.query_controller.invalidate_cache()
        self.export_controller.invalidate_cache()

    def cancel_queries(self) -> None:
        self.query_controller.cancel_pending()
        self.export_controller.cancel_pending()

    def wait_for_queries(self) -> None:
        self.query_controller.wait_for_idle()
        self.export_controller.wait_for_idle()

    def set_empty_state(self, message: str = "Chưa có dữ liệu khách hàng.") -> None:
        self.cancel_queries()
        self.metrics.set_empty(DEBT_GROUP_KPI_LABELS)
        for model in (self.summary_model, self.branch_model, self.officer_model, self.customer_model):
            model.set_rows([])
        for pager, page_size in (
            (self.branch_pager, self.branch_page_size),
            (self.officer_pager, self.officer_page_size),
            (self.customer_pager, self.customer_page_size),
        ):
            pager.set_state(page=1, page_size=page_size, total_rows=0)
        self.structure_chart.set_empty(message)
        self.trend_chart.set_empty(message)
        self.state_banner.set_empty(message)
        self.export_button.setEnabled(False)

    def _build_overview_tab(self) -> None:
        tab = QWidget()
        layout = QVBoxLayout(tab)
        layout.addWidget(self.metrics)
        chart_area = QWidget()
        chart_grid = QGridLayout(chart_area)
        chart_grid.setContentsMargins(0, 0, 0, 0)
        chart_grid.setSpacing(10)
        chart_grid.addWidget(self.structure_chart, 0, 0)
        chart_grid.addWidget(self.trend_chart, 0, 1)
        layout.addWidget(chart_area)
        table = CustomerTableView()
        table.setModel(self.summary_model)
        table.apply_default_widths((140, 130, 100, 150, 120, 110, 110))
        layout.addWidget(table, stretch=1)
        self.tabs.addTab(tab, "Tổng quan")

    def _build_table_tab(self, title: str, model: CustomerTableModel) -> tuple[CustomerTableView, Pager]:
        tab = QWidget()
        layout = QVBoxLayout(tab)
        table = CustomerTableView()
        table.setModel(model)
        layout.addWidget(table, stretch=1)
        pager = Pager()
        layout.addWidget(pager)
        self.tabs.addTab(tab, title)
        return table, pager

    def _filters(self) -> CustomerFilters:
        return replace(self.filters_provider(), debt_group=current_data(self.debt_group_combo))

    def _apply_payload(self, payload: dict[str, object]) -> None:
        self.export_button.setEnabled(True)
        kpis = dict(payload.get("kpis") or {})
        if not kpis.get("has_debt_group_data"):
            self.state_banner.set_empty(
                "Kỳ này chưa có dữ liệu nhóm nợ. Vui lòng nhập lại kỳ NIM Dư nợ từ file FTP Loan có cột AQCCDFIN."
            )
        else:
            self.state_banner.clear()
        self.metrics.set_metrics(
            [
                KpiMetric("Tổng dư nợ", kpis.get("total_balance", 0), "money"),
                KpiMetric("Dư nợ nhóm 1", kpis.get("debt_group_1_balance", 0), "money"),
                KpiMetric("Nợ cần chú ý", kpis.get("attention_balance", 0), "money"),
                KpiMetric("Nợ xấu", kpis.get("bad_debt_balance", 0), "money"),
                KpiMetric("Dư nợ chưa xác định nhóm", kpis.get("debt_group_unknown_balance", 0), "money"),
                KpiMetric("Tỷ lệ nợ cần chú ý", kpis.get("attention_ratio"), "percentage"),
                KpiMetric("Tỷ lệ nợ xấu", kpis.get("bad_debt_ratio"), "percentage"),
                KpiMetric("KH có nợ cần chú ý", kpis.get("attention_customer_count", 0), "count"),
                KpiMetric("KH có nợ xấu", kpis.get("bad_debt_customer_count", 0), "count"),
                KpiMetric("KH có nhóm nợ UNKNOWN", kpis.get("unknown_customer_count", 0), "count"),
            ]
        )
        summary_rows = list(payload.get("summary_rows") or [])
        self.summary_model.set_rows(summary_rows)
        self.structure_chart.set_slices(
            tuple((str(row.get("debt_group") or ""), float(row.get("balance") or 0)) for row in summary_rows)
        )
        trend_rows = list(payload.get("trend_rows") or [])
        self.trend_chart.set_series(
            (
                ("Tỷ lệ nhóm 2", tuple((str(row.get("period") or ""), float(row.get("attention_ratio") or 0)) for row in trend_rows)),
                ("Tỷ lệ nợ xấu", tuple((str(row.get("period") or ""), float(row.get("bad_debt_ratio") or 0)) for row in trend_rows)),
            )
        )
        branch_result = payload["branch_result"]
        officer_result = payload["officer_result"]
        customer_result = payload["customer_result"]
        self.branch_model.set_rows(_rank_page_rows(branch_result.rows, branch_result.page, branch_result.page_size))
        self.officer_model.set_rows(_rank_page_rows(officer_result.rows, officer_result.page, officer_result.page_size))
        self.customer_model.set_rows(_rank_page_rows(customer_result.rows, customer_result.page, customer_result.page_size))
        self.branch_pager.set_state(page=branch_result.page, page_size=branch_result.page_size, total_rows=branch_result.total_rows)
        self.officer_pager.set_state(page=officer_result.page, page_size=officer_result.page_size, total_rows=officer_result.total_rows)
        self.customer_pager.set_state(page=customer_result.page, page_size=customer_result.page_size, total_rows=customer_result.total_rows)

    def _query_failed(self, exc: Exception) -> None:
        self.set_empty_state("Không tải được dữ liệu phân tích nhóm nợ.")

    def _state_changed(self, state: str, message: str) -> None:
        if state == "loading":
            self.state_banner.set_loading("Đang tải phân tích nhóm nợ...")
            self.metrics.set_loading(DEBT_GROUP_KPI_LABELS)
        elif state == "error":
            self.state_banner.set_error(message or "Không tải được phân tích nhóm nợ.")

    def _export_state_changed(self, state: str, message: str) -> None:
        if state == "loading":
            self.state_banner.set_loading("Đang xuất Excel phân tích nhóm nợ...")
        elif state == "error":
            self.state_banner.set_error(message or "Không xuất được Excel phân tích nhóm nợ.")

    def _export_finished(self, output: Path) -> None:
        self.export_button.setEnabled(True)
        if self.state_banner.state == "loading":
            self.state_banner.clear()
        QMessageBox.information(self, "Xuất phân tích nhóm nợ", f"Đã xuất: {output}")

    def _export_failed(self, exc: Exception) -> None:
        self.export_button.setEnabled(True)
        QMessageBox.warning(self, "Xuất phân tích nhóm nợ", str(exc))

    def _filter_changed(self) -> None:
        self.branch_page = 1
        self.officer_page = 1
        self.customer_page = 1
        self.refresh()

    def _page_changed(self, target: str, page: int) -> None:
        setattr(self, f"{target}_page", max(1, int(page or 1)))
        self.refresh()

    def _page_size_changed(self, target: str, page_size: int) -> None:
        setattr(self, f"{target}_page", 1)
        setattr(self, f"{target}_page_size", int(page_size or 100))
        self.refresh()

    def _sort_changed(self, target: str, section: int) -> None:
        columns = {
            "branch": DEBT_GROUP_BRANCH_COLUMNS,
            "officer": DEBT_GROUP_OFFICER_COLUMNS,
            "customer": DEBT_GROUP_CUSTOMER_COLUMNS,
        }[target]
        if not (0 <= section < len(columns)):
            return
        field = columns[section][0]
        sort_by_name = f"{target}_sort_by"
        sort_desc_name = f"{target}_sort_desc"
        if getattr(self, sort_by_name) == field:
            setattr(self, sort_desc_name, not getattr(self, sort_desc_name))
        else:
            setattr(self, sort_by_name, field)
            setattr(self, sort_desc_name, field not in {"period", "customer_code", "customer_name", "branch_code", "officer_name"})
        setattr(self, f"{target}_page", 1)
        self.refresh()

    def _open_customer_detail(self, index) -> None:
        row = self.customer_model.raw_row(index.row())
        if not row:
            return
        dialog = CustomerDetailWindow(
            self.repository,
            str(row.get("customer_code") or ""),
            period=str(row.get("period") or ""),
            initial_tab=5,
            parent=self,
        )
        dialog.setAttribute(Qt.WidgetAttribute.WA_DeleteOnClose, True)
        self.detail_windows.append(dialog)
        dialog.finished.connect(lambda _result, item=dialog: self._forget_detail(item))
        dialog.show()
        dialog.raise_()

    def _forget_detail(self, dialog: CustomerDetailWindow) -> None:
        if dialog in self.detail_windows:
            self.detail_windows.remove(dialog)


def _rank_page_rows(rows: list[dict[str, object]], page: int, page_size: int) -> list[dict[str, object]]:
    start = (max(1, int(page or 1)) - 1) * max(1, int(page_size or 100))
    return [dict(row, rank=start + index) for index, row in enumerate(rows, start=1)]
