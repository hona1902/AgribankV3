from __future__ import annotations

from dataclasses import replace
from pathlib import Path

from PySide6.QtCore import Qt, QTimer, Signal
from PySide6.QtWidgets import (
    QFileDialog,
    QGridLayout,
    QHBoxLayout,
    QLabel,
    QLayout,
    QMessageBox,
    QScrollArea,
    QSizePolicy,
    QTabWidget,
    QVBoxLayout,
    QDialog,
    QWidget,
)

from agribank_v3.features.credit.summary.customer.async_query import AsyncQueryController
from agribank_v3.features.credit.summary.customer.charts import CustomerDonutChart, CustomerLineChart
from agribank_v3.features.credit.summary.customer.customer_detail_window import CustomerDetailWindow
from agribank_v3.features.credit.summary.customer.filters import (
    CUSTOMER_TYPE_FILTERS,
    DEBT_GROUP_FILTERS,
    LOAN_TERM_FILTERS,
)
from agribank_v3.features.credit.summary.customer.officer_center_export import (
    OFFICER_COMPARE_COLUMNS,
    OFFICER_CUSTOMER_COLUMNS,
    OFFICER_LIST_COLUMNS,
    OFFICER_MOVEMENT_COLUMNS,
    OFFICER_TOP_COLUMNS,
    export_officer_center_workbook,
    export_officer_top_workbook,
    suggested_officer_export_name,
)
from agribank_v3.features.credit.summary.customer.officer_center_repository import (
    OFFICER_MODE_IMPORTED,
    OFFICER_MODE_LABELS,
    OFFICER_STATUS_FILTERS,
    OfficerCenterFilters,
    OfficerCenterRepository,
    TERM_STRUCTURE_WARNING,
)
from agribank_v3.features.credit.summary.customer.officer_detail_registry import open_shared_officer_detail
from agribank_v3.features.credit.summary.customer.officer_management_tab import OfficerManagementTab
from agribank_v3.features.credit.summary.customer.repository import CustomerRepository
from agribank_v3.features.credit.summary.customer.table_models import ColumnSpec, CustomerTableModel
from agribank_v3.features.credit.summary.customer.widgets import (
    CompactToolbar,
    CustomerTableView,
    KpiMetric,
    MetricGrid,
    Pager,
    QueryStateBanner,
    SearchBox,
    combo_box,
    configure_combo_popup_width,
    current_data,
    fit_window_to_screen,
    populate_combo,
    primary_button,
    secondary_button,
)
from agribank_v3.features.credit.summary.officer_history.models import OfficerKey
from agribank_v3.features.credit.summary.officer_history.widgets import OfficerMultiSelectCombo
from agribank_v3.ui.workers import run_in_thread


OFFICER_CENTER_TITLE = "Quản lý CBTD"


class OfficerCenterWindow(QDialog):
    openNimDnRequested = Signal()

    def __init__(self, main_database_path: Path, parent=None) -> None:
        super().__init__(parent)
        self.main_database_path = Path(main_database_path)
        self.customer_repository = CustomerRepository(self.main_database_path)
        self.repository = OfficerCenterRepository(self.customer_repository)
        self._customer_detail_windows: list[CustomerDetailWindow] = []
        self._export_thread = None
        self._updating_filters = False
        self._available_periods: list[str] = []
        self._filter_timer = QTimer(self)
        self._filter_timer.setSingleShot(True)
        self._filter_timer.setInterval(350)
        self._filter_timer.timeout.connect(self._refresh_current_tab)
        self.setWindowTitle("Quản lý cán bộ tín dụng - AgribankV3")
        self.setModal(False)
        self.setWindowFlag(Qt.WindowType.WindowMinMaxButtonsHint, True)
        fit_window_to_screen(
            self,
            width_ratio=0.92,
            height_ratio=0.90,
            max_width=1480,
            max_height=920,
            min_width=1080,
            min_height=680,
        )
        layout = QVBoxLayout(self)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(8)
        layout.addWidget(self._build_filter_panel())
        self.tabs = QTabWidget()
        self.dashboard_tab = OfficerDashboardTab(self.repository, self.current_filters, self)
        self.list_tab = OfficerPagedTableTab(
            self.repository,
            self.current_filters,
            "officer_list",
            OFFICER_LIST_COLUMNS,
            "total_balance",
            detail_callback=self.open_officer_detail,
            open_nim_callback=self.openNimDnRequested.emit,
        )
        self.movement_tab = OfficerPagedTableTab(
            self.repository,
            self.current_filters,
            "officer_movement",
            OFFICER_MOVEMENT_COLUMNS,
            "balance_change",
            detail_callback=self.open_officer_detail,
        )
        self.compare_tab = OfficerCompareTab(
            self.repository,
            self.current_filters,
            detail_callback=self.open_officer_detail,
            open_nim_callback=self.openNimDnRequested.emit,
        )
        self.debt_quality_tab = OfficerPagedTableTab(
            self.repository,
            self.current_filters,
            "officer_debt_quality",
            OFFICER_LIST_COLUMNS,
            "bad_debt_ratio",
            detail_callback=lambda row: self.open_officer_detail(row, initial_tab=4),
            open_nim_callback=self.openNimDnRequested.emit,
        )
        self.customer_tab = OfficerPagedTableTab(
            self.repository,
            self.current_filters,
            "officer_customers",
            OFFICER_CUSTOMER_COLUMNS,
            "customer_code",
            detail_callback=self.open_customer_detail,
            open_nim_callback=self.openNimDnRequested.emit,
        )
        self.directory_tab = OfficerManagementTab(self.customer_repository, self)
        self.tabs.addTab(self.dashboard_tab, "Dashboard")
        self.tabs.addTab(self.list_tab, "Danh sách CBTD")
        self.tabs.addTab(self.movement_tab, "Biến động danh mục")
        self.tabs.addTab(self.compare_tab, "So sánh CBTD")
        self.tabs.addTab(self.debt_quality_tab, "Chất lượng tín dụng")
        self.tabs.addTab(self.customer_tab, "Khách hàng quản lý")
        self.tabs.addTab(self.directory_tab, "Danh mục CBTD")
        self.tabs.currentChanged.connect(lambda _index: self._refresh_current_tab())
        layout.addWidget(self.tabs, stretch=1)
        self.refresh_filters()
        self._refresh_current_tab()

    def _build_filter_panel(self) -> QWidget:
        panel = QWidget()
        panel.setObjectName("OfficerCenterFilterPanel")
        panel_layout = QVBoxLayout(panel)
        panel_layout.setContentsMargins(10, 8, 10, 8)
        panel_layout.setSpacing(7)
        self.period_from_combo = combo_box("Từ kỳ", minimum_width=135, maximum_width=155)
        self.period_to_combo = combo_box("Đến kỳ", minimum_width=135, maximum_width=155)
        self.report_period_combo = combo_box("Kỳ báo cáo", minimum_width=140, maximum_width=165)
        self.branch_combo = combo_box("Tất cả chi nhánh", minimum_width=190, maximum_width=280)
        self.office_combo = combo_box("Tất cả Phòng GD", minimum_width=170, maximum_width=250)
        self.customer_type_combo = combo_box("Loại khách hàng", minimum_width=155, maximum_width=210)
        self.loan_term_combo = combo_box("Loại thời hạn", minimum_width=165, maximum_width=220)
        self.debt_group_combo = combo_box("Nhóm nợ", minimum_width=155, maximum_width=220)
        self.status_combo = combo_box("Trạng thái CBTD", minimum_width=190, maximum_width=280)
        self.mode_combo = combo_box("Cách xác định dữ liệu", minimum_width=260, maximum_width=420)
        self.search_box = SearchBox("Tìm mã hoặc tên CBTD")
        self.search_box.setMinimumWidth(320)
        self.search_box.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        self.refresh_button = secondary_button("Làm mới")
        self.clear_button = secondary_button("Xóa lọc")
        self.export_button = primary_button("Xuất toàn bộ")
        populate_combo(self.customer_type_combo, CUSTOMER_TYPE_FILTERS[1:])
        populate_combo(self.loan_term_combo, LOAN_TERM_FILTERS[1:])
        populate_combo(self.debt_group_combo, DEBT_GROUP_FILTERS[1:])
        populate_combo(self.status_combo, OFFICER_STATUS_FILTERS[1:])
        populate_combo(self.mode_combo, OFFICER_MODE_LABELS)
        mode_index = self.mode_combo.findData(OFFICER_MODE_IMPORTED)
        if mode_index >= 0:
            self.mode_combo.setCurrentIndex(mode_index)
            self.mode_combo.setToolTip(self.mode_combo.currentText())
        configure_combo_popup_width(self.mode_combo, minimum_popup_width=340, maximum_screen_ratio=0.90)
        configure_combo_popup_width(self.status_combo, minimum_popup_width=260, maximum_screen_ratio=0.90)
        for combo in (
            self.period_from_combo,
            self.period_to_combo,
            self.report_period_combo,
            self.branch_combo,
            self.office_combo,
            self.customer_type_combo,
            self.loan_term_combo,
            self.debt_group_combo,
            self.status_combo,
            self.mode_combo,
        ):
            combo.currentIndexChanged.connect(self._filter_changed)
        self.mode_combo.currentIndexChanged.connect(lambda index: self.mode_combo.setToolTip(self.mode_combo.itemText(index)))
        self.branch_combo.currentIndexChanged.connect(lambda _index: self._refresh_office_filter())
        self.search_box.debouncedTextChanged.connect(lambda _text: self._filter_changed())
        self.refresh_button.clicked.connect(lambda: self.refresh_all(use_cache=False))
        self.clear_button.clicked.connect(self.clear_filters)
        self.export_button.clicked.connect(self.export_all)
        period_toolbar = CompactToolbar()
        period_toolbar.setObjectName("OfficerCenterPeriodToolbar")
        for widget in (
            self.period_from_combo,
            self.period_to_combo,
            self.report_period_combo,
            self.branch_combo,
            self.office_combo,
        ):
            period_toolbar.addWidget(widget)
        analysis_toolbar = CompactToolbar()
        analysis_toolbar.setObjectName("OfficerCenterAnalysisToolbar")
        for widget in (
            self.customer_type_combo,
            self.loan_term_combo,
            self.debt_group_combo,
            self.status_combo,
            self.mode_combo,
        ):
            analysis_toolbar.addWidget(widget)
        action_toolbar = CompactToolbar()
        action_toolbar.setObjectName("OfficerCenterActionToolbar")
        for widget in (self.search_box, self.refresh_button, self.clear_button, self.export_button):
            action_toolbar.addWidget(widget)
        self.filter_toolbars = (period_toolbar, analysis_toolbar, action_toolbar)
        panel_layout.addWidget(period_toolbar)
        panel_layout.addWidget(analysis_toolbar)
        panel_layout.addWidget(action_toolbar)
        return panel

    def refresh_filters(self) -> None:
        self._updating_filters = True
        try:
            periods = self.repository.distinct_periods()
            self._available_periods = periods
            for combo in (self.period_from_combo, self.period_to_combo, self.report_period_combo):
                current = current_data(combo)
                combo.clear()
                combo.addItem(combo.toolTip() or "Tất cả", "")
                for period in periods:
                    combo.addItem(period, period)
                if current:
                    index = combo.findData(current)
                    if index >= 0:
                        combo.setCurrentIndex(index)
            if periods:
                self._set_combo_data_if_empty(self.period_from_combo, periods[0])
                self._set_combo_data_if_empty(self.period_to_combo, periods[-1])
                self._set_combo_data_if_empty(self.report_period_combo, periods[-1])
            current_branch = current_data(self.branch_combo)
            self.branch_combo.clear()
            self.branch_combo.addItem("Tất cả chi nhánh", "")
            for code in self.repository.distinct_branch_codes(self.current_filters()):
                self.branch_combo.addItem(self.repository.unit_directory.get_branch_display_name(code), code)
            if current_branch:
                index = self.branch_combo.findData(current_branch)
                if index >= 0:
                    self.branch_combo.setCurrentIndex(index)
            self._refresh_office_filter()
            self.directory_tab.refresh_filters()
        finally:
            self._updating_filters = False

    def current_filters(self) -> OfficerCenterFilters:
        return OfficerCenterFilters(
            period_from=current_data(self.period_from_combo),
            period_to=current_data(self.period_to_combo),
            report_period=current_data(self.report_period_combo),
            compare_period=self.repository.previous_period(current_data(self.report_period_combo)),
            branch_code=current_data(self.branch_combo),
            transaction_office=current_data(self.office_combo),
            customer_type=current_data(self.customer_type_combo),
            loan_term=current_data(self.loan_term_combo),
            debt_group=current_data(self.debt_group_combo),
            officer_status=current_data(self.status_combo),
            mode=current_data(self.mode_combo) or OFFICER_MODE_IMPORTED,
            search_text=self.search_box.text(),
        ).normalized()

    def clear_filters(self) -> None:
        self._updating_filters = True
        try:
            for combo in (
                self.branch_combo,
                self.office_combo,
                self.customer_type_combo,
                self.loan_term_combo,
                self.debt_group_combo,
                self.status_combo,
            ):
                combo.setCurrentIndex(0)
            self.mode_combo.setCurrentIndex(0)
            self.search_box.clear()
            if self._available_periods:
                self._set_combo_data(self.period_from_combo, self._available_periods[0])
                self._set_combo_data(self.period_to_combo, self._available_periods[-1])
                self._set_combo_data(self.report_period_combo, self._available_periods[-1])
        finally:
            self._updating_filters = False
        self.refresh_all(use_cache=False)

    def refresh_all(self, *, use_cache: bool = True) -> None:
        self.refresh_filters()
        for tab in (
            self.dashboard_tab,
            self.list_tab,
            self.movement_tab,
            self.compare_tab,
            self.debt_quality_tab,
            self.customer_tab,
            self.directory_tab,
        ):
            if hasattr(tab, "refresh"):
                tab.refresh(use_cache=use_cache)

    def invalidate_cache(self) -> None:
        for tab in (
            self.dashboard_tab,
            self.list_tab,
            self.movement_tab,
            self.compare_tab,
            self.debt_quality_tab,
            self.customer_tab,
            self.directory_tab,
        ):
            if hasattr(tab, "invalidate_cache"):
                tab.invalidate_cache()

    def export_all(self) -> None:
        path, _ = QFileDialog.getSaveFileName(
            self,
            "Xuất toàn bộ Quản lý CBTD",
            suggested_officer_export_name("QuanLyCBTD"),
            "Excel (*.xlsx)",
        )
        if not path:
            return
        filters = self.current_filters()
        self.export_button.setEnabled(False)
        self._export_thread = run_in_thread(
            lambda: export_officer_center_workbook(self.repository, filters, Path(path)),
            lambda output: self._export_finished(output),
            lambda exc: self._export_failed(exc),
            self,
        )

    def open_officer_detail(self, row: dict[str, object], *, initial_tab: int | None = None) -> None:
        open_shared_officer_detail(
            self,
            self.main_database_path,
            row,
            filters=self.current_filters(),
            initial_tab=initial_tab,
        )

    def open_customer_detail(self, row: dict[str, object]) -> None:
        customer_code = str(row.get("customer_code") or "").strip()
        period = str(row.get("period") or self.current_filters().report_period)
        if not customer_code:
            return
        dialog = CustomerDetailWindow(self.customer_repository, customer_code=customer_code, period=period, parent=self)
        self._customer_detail_windows.append(dialog)
        dialog.finished.connect(lambda _result, item=dialog: self._discard_customer_detail(item))
        dialog.show()

    def select_tab(self, tab: str | int) -> None:
        if isinstance(tab, int):
            self.tabs.setCurrentIndex(max(0, min(self.tabs.count() - 1, tab)))
            return
        mapping = {
            "dashboard": 0,
            "list": 1,
            "officers": 1,
            "movement": 2,
            "compare": 3,
            "debt_quality": 4,
            "customers": 5,
            "directory": 6,
        }
        index = mapping.get(str(tab or "").strip().casefold())
        if index is not None:
            self.tabs.setCurrentIndex(index)

    def _refresh_current_tab(self) -> None:
        if self._updating_filters:
            return
        tab = self.tabs.currentWidget()
        if hasattr(tab, "refresh"):
            tab.refresh()

    def _filter_changed(self) -> None:
        if self._updating_filters:
            return
        self._filter_timer.start()

    def _refresh_office_filter(self) -> None:
        if self._updating_filters:
            return
        current = current_data(self.office_combo)
        period = current_data(self.report_period_combo)
        branch = current_data(self.branch_combo)
        self.office_combo.blockSignals(True)
        try:
            self.office_combo.clear()
            self.office_combo.addItem("Tất cả PGD", "")
            for row in self.repository.distinct_offices(period, branch_code=branch):
                self.office_combo.addItem(str(row.get("office_display") or row.get("trctcd") or ""), str(row.get("trctcd") or ""))
            if current:
                index = self.office_combo.findData(current)
                if index >= 0:
                    self.office_combo.setCurrentIndex(index)
        finally:
            self.office_combo.blockSignals(False)

    def _export_finished(self, output: Path) -> None:
        self.export_button.setEnabled(True)
        QMessageBox.information(self, "Xuất toàn bộ Quản lý CBTD", f"Đã xuất: {output}")

    def _export_failed(self, exc: Exception) -> None:
        self.export_button.setEnabled(True)
        QMessageBox.warning(self, "Xuất toàn bộ Quản lý CBTD", str(exc))

    def _discard_customer_detail(self, dialog: CustomerDetailWindow) -> None:
        if dialog in self._customer_detail_windows:
            self._customer_detail_windows.remove(dialog)

    @staticmethod
    def _set_combo_data_if_empty(combo, value: str) -> None:
        if not current_data(combo):
            OfficerCenterWindow._set_combo_data(combo, value)

    @staticmethod
    def _set_combo_data(combo, value: str) -> None:
        index = combo.findData(value)
        if index >= 0:
            combo.setCurrentIndex(index)

    def closeEvent(self, event) -> None:
        for tab in (
            self.dashboard_tab,
            self.list_tab,
            self.movement_tab,
            self.compare_tab,
            self.debt_quality_tab,
            self.customer_tab,
            self.directory_tab,
        ):
            if hasattr(tab, "cancel_queries"):
                tab.cancel_queries()
            if hasattr(tab, "wait_for_queries"):
                tab.wait_for_queries()
        super().closeEvent(event)


class OfficerDashboardTab(QWidget):
    def __init__(self, repository: OfficerCenterRepository, filters_provider, parent=None) -> None:
        super().__init__(parent)
        self.repository = repository
        self.filters_provider = filters_provider
        self.query_controller = AsyncQueryController(self, max_cache_entries=24)
        self._last_mode_label = ""
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(8)
        self.state_banner = QueryStateBanner()
        self.state_banner.retryRequested.connect(lambda: self.refresh(use_cache=False))
        layout.addWidget(self.state_banner)
        self.dashboard_scroll = QScrollArea()
        self.dashboard_scroll.setWidgetResizable(True)
        self.dashboard_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.dashboard_content = QWidget()
        self.dashboard_content.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.MinimumExpanding)
        self.dashboard_content_layout = QVBoxLayout(self.dashboard_content)
        self.dashboard_content_layout.setContentsMargins(0, 0, 0, 0)
        self.dashboard_content_layout.setSpacing(10)
        self.dashboard_content_layout.setSizeConstraint(QLayout.SizeConstraint.SetMinimumSize)
        self.metrics = MetricGrid()
        self.dashboard_content_layout.addWidget(self.metrics)
        chart_panel = QWidget()
        chart_grid = QGridLayout(chart_panel)
        chart_grid.setContentsMargins(0, 0, 0, 0)
        chart_grid.setHorizontalSpacing(10)
        chart_grid.setVerticalSpacing(10)
        self.balance_chart = CustomerLineChart("Xu hướng tổng dư nợ theo kỳ", value_kind="money")
        self.count_chart = CustomerLineChart("Xu hướng số CBTD có dư nợ", value_kind="number")
        self.metric_chart = CustomerLineChart("Xu hướng NIM sau ĐC", value_kind="percent")
        self.debt_chart = CustomerDonutChart("Cơ cấu chất lượng dư nợ", value_kind="money")
        chart_grid.addWidget(self.balance_chart, 0, 0)
        chart_grid.addWidget(self.count_chart, 0, 1)
        chart_grid.addWidget(self.metric_chart, 1, 0)
        chart_grid.addWidget(self.debt_chart, 1, 1)
        chart_grid.setColumnStretch(0, 1)
        chart_grid.setColumnStretch(1, 1)
        self.dashboard_content_layout.addWidget(chart_panel)
        top_toolbar = QHBoxLayout()
        top_toolbar.setContentsMargins(0, 0, 0, 0)
        top_toolbar.setSpacing(8)
        self.top_title = QLabel("Top CBTD theo dư nợ")
        self.top_title.setObjectName("SectionTitle")
        self.top_metric_combo = combo_box("Chỉ tiêu", minimum_width=150, maximum_width=220)
        populate_combo(
            self.top_metric_combo,
            (
                ("Dư nợ", "total_balance"),
                ("Tăng trưởng", "balance_change"),
                ("NIM sau ĐC", "nim_after"),
                ("Nợ nhóm 2", "attention_balance"),
                ("Nợ xấu", "bad_debt_balance"),
                ("Tỷ lệ nợ xấu", "bad_debt_ratio"),
            ),
        )
        metric_index = self.top_metric_combo.findData("total_balance")
        if metric_index >= 0:
            self.top_metric_combo.setCurrentIndex(metric_index)
        self.top_limit_combo = combo_box("Top 10", minimum_width=105, maximum_width=130)
        self.top_limit_combo.clear()
        for limit in (10, 20, 50):
            self.top_limit_combo.addItem(f"Top {limit}", limit)
        self.top_export_button = secondary_button("Xuất Excel")
        self.top_metric_combo.currentIndexChanged.connect(lambda _index: self.refresh(use_cache=False))
        self.top_limit_combo.currentIndexChanged.connect(lambda _index: self.refresh(use_cache=False))
        self.top_export_button.clicked.connect(self.export_top_excel)
        top_toolbar.addWidget(self.top_title)
        top_toolbar.addStretch()
        top_toolbar.addWidget(self.top_metric_combo)
        top_toolbar.addWidget(self.top_limit_combo)
        top_toolbar.addWidget(self.top_export_button)
        self.dashboard_content_layout.addLayout(top_toolbar)
        self.top_model = CustomerTableModel(OFFICER_TOP_COLUMNS, self)
        self.top_table = CustomerTableView()
        self.top_table.setModel(self.top_model)
        self.top_table.setMinimumHeight(320)
        self.top_table.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.MinimumExpanding)
        self.top_table.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        self.top_table.apply_default_widths(_default_widths(OFFICER_TOP_COLUMNS))
        self.dashboard_content_layout.addWidget(self.top_table)
        self.dashboard_scroll.setWidget(self.dashboard_content)
        layout.addWidget(self.dashboard_scroll, stretch=1)

    def refresh(self, *args, use_cache: bool = True) -> None:
        filters = self.filters_provider()
        top_metric = current_data(self.top_metric_combo) or "total_balance"
        top_limit = int(current_data(self.top_limit_combo) or 10)
        self.query_controller.run(
            "officer_dashboard",
            lambda: self.repository.dashboard_payload(filters, top_metric=top_metric, top_limit=top_limit),
            self._apply_payload,
            self._query_failed,
            cache_key=("officer_dashboard", filters, top_metric, top_limit),
            use_cache=use_cache,
            state_callback=self._state_changed,
        )

    def cancel_queries(self) -> None:
        self.query_controller.cancel_pending()

    def invalidate_cache(self) -> None:
        self.query_controller.invalidate_cache()

    def wait_for_queries(self) -> None:
        self.query_controller.wait_for_idle()

    def _apply_payload(self, payload: dict[str, object]) -> None:
        kpis = dict(payload.get("kpis") or {})
        self.metrics.set_metrics(_dashboard_metrics(kpis))
        self._last_mode_label = str(payload.get("mode_label") or kpis.get("mode_label") or "")
        self.balance_chart.set_series((("Tổng dư nợ", _trend_points(payload.get("balance_trend"))),))
        self.count_chart.set_series((("Số CBTD", _trend_points(payload.get("officer_count_trend"))),))
        self.metric_chart.set_series((("NIM sau ĐC", _trend_points(payload.get("metric_trend"))),))
        debt = dict(payload.get("debt_structure") or {})
        self.debt_chart.set_slices(
            (
                ("Nhóm 1", float(debt.get("debt_group_1_balance") or 0)),
                ("Nhóm 2", float(debt.get("debt_group_2_balance") or 0)),
                (
                    "Nợ xấu 3-5",
                    float(debt.get("debt_group_3_balance") or 0)
                    + float(debt.get("debt_group_4_balance") or 0)
                    + float(debt.get("debt_group_5_balance") or 0),
                ),
                ("UNKNOWN", float(debt.get("debt_group_unknown_balance") or 0)),
            )
        )
        limit = int(payload.get("top_limit") or current_data(self.top_limit_combo) or 10)
        metric_label = self.top_metric_combo.currentText() or "Dư nợ"
        self.top_title.setText(f"Top {limit} CBTD theo {metric_label}")
        self.top_model.set_rows([dict(row, rank=index) for index, row in enumerate(payload.get("top_rows") or (), start=1)])
        self.top_table.updateGeometry()
        self.dashboard_content.updateGeometry()
        self.dashboard_content.adjustSize()
        self.state_banner.clear()

    def _query_failed(self, exc: Exception) -> None:
        self.metrics.set_metrics(_dashboard_placeholder_metrics())
        self.top_model.set_rows([])
        self.balance_chart.set_error()
        self.count_chart.set_error()
        self.metric_chart.set_error()
        self.debt_chart.set_error()
        self.state_banner.set_error(str(exc))

    def _state_changed(self, state: str, message: str) -> None:
        if state == "loading":
            self.state_banner.set_loading("Đang tải Dashboard CBTD...")
        elif state == "error":
            self.state_banner.set_error(message or "Không tải được Dashboard CBTD.")

    def export_top_excel(self) -> None:
        path, _ = QFileDialog.getSaveFileName(
            self,
            "Xuất Top CBTD",
            suggested_officer_export_name("TopCBTD"),
            "Excel (*.xlsx)",
        )
        if not path:
            return
        rows = [dict(row) for row in self.top_model.rows]
        try:
            output = export_officer_top_workbook(
                rows,
                Path(path),
                metric_label=self.top_metric_combo.currentText() or "Dư nợ",
                limit=int(current_data(self.top_limit_combo) or len(rows) or 10),
                mode_label=self._last_mode_label,
            )
        except Exception as exc:
            QMessageBox.warning(self, "Xuất Top CBTD", str(exc))
            return
        QMessageBox.information(self, "Xuất Top CBTD", f"Đã xuất: {output}")


class OfficerPagedTableTab(QWidget):
    def __init__(
        self,
        repository: OfficerCenterRepository,
        filters_provider,
        query_name: str,
        columns: tuple[ColumnSpec, ...],
        default_sort: str,
        *,
        detail_callback=None,
        open_nim_callback=None,
        parent=None,
    ) -> None:
        super().__init__(parent)
        self.repository = repository
        self.filters_provider = filters_provider
        self.query_name = query_name
        self.columns = columns
        self.sort_by = default_sort
        self.sort_desc = True
        self.page = 1
        self.page_size = 100
        self.detail_callback = detail_callback
        self.open_nim_callback = open_nim_callback
        self.query_controller = AsyncQueryController(self, max_cache_entries=32)
        layout = QVBoxLayout(self)
        self.state_banner = QueryStateBanner()
        self.state_banner.retryRequested.connect(lambda: self.refresh(use_cache=False))
        layout.addWidget(self.state_banner)
        self.term_warning_bar = QWidget()
        warning_layout = QHBoxLayout(self.term_warning_bar)
        warning_layout.setContentsMargins(10, 6, 10, 6)
        warning_layout.setSpacing(8)
        self.term_warning_label = QLabel("Dữ liệu kỳ hạn theo CBTD của kỳ này chưa đầy đủ. Hãy nhập lại kỳ NIM Dư nợ để cập nhật.")
        self.term_warning_label.setWordWrap(True)
        self.term_warning_label.setToolTip(TERM_STRUCTURE_WARNING)
        self.open_nim_button = secondary_button("Mở NIM Dư nợ")
        self.open_nim_button.setVisible(callable(self.open_nim_callback))
        self.open_nim_button.clicked.connect(self._open_nim_dn)
        warning_layout.addWidget(self.term_warning_label, stretch=1)
        warning_layout.addWidget(self.open_nim_button)
        self.term_warning_bar.setVisible(False)
        layout.addWidget(self.term_warning_bar)
        self.model = CustomerTableModel(columns, self)
        self.table = CustomerTableView()
        self.table.setModel(self.model)
        self.table.horizontalHeader().sectionClicked.connect(self._sort_requested)
        self.table.doubleClicked.connect(self._open_detail)
        self.table.apply_default_widths(_default_widths(columns))
        layout.addWidget(self.table, stretch=1)
        self.pager = Pager()
        self.pager.pageChanged.connect(self._page_changed)
        self.pager.pageSizeChanged.connect(self._page_size_changed)
        layout.addWidget(self.pager)

    def refresh(self, *args, use_cache: bool = True) -> None:
        filters = self.filters_provider()
        page = self.page
        page_size = self.page_size
        sort_by = self.sort_by
        sort_desc = self.sort_desc
        method = getattr(self.repository, self.query_name)
        self.query_controller.run(
            self.query_name,
            lambda: method(filters, page=page, page_size=page_size, sort_by=sort_by, sort_desc=sort_desc),
            self._apply_result,
            self._query_failed,
            cache_key=(self.query_name, filters, page, page_size, sort_by, sort_desc),
            use_cache=use_cache,
            state_callback=self._state_changed,
        )

    def cancel_queries(self) -> None:
        self.query_controller.cancel_pending()

    def invalidate_cache(self) -> None:
        self.query_controller.invalidate_cache()

    def wait_for_queries(self) -> None:
        self.query_controller.wait_for_idle()

    def _apply_result(self, result) -> None:
        self.model.set_rows([dict(row, rank=(result.page - 1) * result.page_size + index) for index, row in enumerate(result.rows, start=1)])
        self.pager.set_state(page=result.page, page_size=result.page_size, total_rows=result.total_rows)
        self._set_term_warning(_has_incomplete_term_rows(result.rows) and _columns_include_term_structure(self.columns))
        if result.total_rows:
            self.state_banner.clear()
        else:
            self.state_banner.set_empty("Không có dữ liệu CBTD phù hợp với bộ lọc.")

    def _query_failed(self, exc: Exception) -> None:
        self.model.set_rows([])
        self._set_term_warning(False)
        self.pager.set_state(page=1, page_size=self.page_size, total_rows=0)
        self.state_banner.set_error(str(exc))

    def _state_changed(self, state: str, message: str) -> None:
        if state == "loading":
            self.state_banner.set_loading("Đang tải dữ liệu CBTD...")
        elif state == "error":
            self.state_banner.set_error(message or "Không tải được dữ liệu CBTD.")

    def _sort_requested(self, section: int) -> None:
        if not (0 <= section < len(self.columns)):
            return
        field = self.columns[section][0]
        if field == "rank":
            return
        if self.sort_by == field:
            self.sort_desc = not self.sort_desc
        else:
            self.sort_by = field
            self.sort_desc = True
        self.page = 1
        self.refresh(use_cache=False)

    def _open_detail(self, index) -> None:
        if self.detail_callback is None or not index.isValid():
            return
        row = self.model.raw_row(index.row())
        self.detail_callback(row)

    def _page_changed(self, page: int) -> None:
        self.page = max(1, int(page or 1))
        self.refresh()

    def _page_size_changed(self, page_size: int) -> None:
        self.page = 1
        self.page_size = int(page_size or 100)
        self.refresh()

    def _set_term_warning(self, visible: bool) -> None:
        self.term_warning_bar.setVisible(bool(visible))

    def _open_nim_dn(self) -> None:
        if callable(self.open_nim_callback):
            self.open_nim_callback()


class OfficerCompareTab(OfficerPagedTableTab):
    def __init__(
        self,
        repository: OfficerCenterRepository,
        filters_provider,
        *,
        detail_callback=None,
        open_nim_callback=None,
        parent=None,
    ) -> None:
        self._base_filters_provider = filters_provider
        self._selector_cache_key = None
        self.selector = OfficerMultiSelectCombo(placeholder="Chọn CBTD so sánh", counter_label="CBTD")
        self.search_box = SearchBox("Tìm trong danh sách CBTD")
        self.apply_button = secondary_button("Áp dụng so sánh")
        self.series_status_label = QLabel("Biểu đồ đang hiển thị 0/0 CBTD")
        self.series_status_label.setObjectName("MutedText")
        super().__init__(
            repository,
            filters_provider,
            "compare_officers",
            OFFICER_COMPARE_COLUMNS,
            "total_balance",
            detail_callback=detail_callback,
            open_nim_callback=open_nim_callback,
            parent=parent,
        )
        toolbar = QHBoxLayout()
        toolbar.setContentsMargins(0, 0, 0, 0)
        toolbar.addWidget(self.search_box, stretch=1)
        toolbar.addWidget(self.selector)
        toolbar.addWidget(self.apply_button)
        toolbar.addWidget(self.series_status_label)
        self.layout().insertLayout(0, toolbar)
        self.search_box.debouncedTextChanged.connect(self.selector.set_filter_text)
        self.apply_button.clicked.connect(lambda: self.refresh(use_cache=False))
        self.filters_provider = self.refresh_query_filters

    def refresh(self, *args, use_cache: bool = True) -> None:
        self._refresh_selector()
        super().refresh(use_cache=use_cache)

    def invalidate_cache(self) -> None:
        super().invalidate_cache()
        self._selector_cache_key = None

    def _apply_result(self, result) -> None:
        super()._apply_result(result)
        selected_count = len(self.selector.selected_officers()) or result.total_rows
        visible_count = min(20, int(result.total_rows or 0))
        self.series_status_label.setText(f"Biểu đồ đang hiển thị {visible_count}/{selected_count} CBTD; bảng giữ đầy đủ dữ liệu.")

    def _refresh_selector(self) -> None:
        filters = self._base_filters()
        cache_key = replace(filters, selected_officers=())
        if cache_key == self._selector_cache_key:
            return
        officers = [
            OfficerKey(
                code=str(row.get("officer_code") or ""),
                raw_name=str(row.get("officer_key") or row.get("officer_name") or ""),
                display_name=str(row.get("officer_display") or row.get("officer_name") or row.get("officer_code") or ""),
                branch=str(row.get("branch_name") or ""),
                transaction_office=str(row.get("office_name") or ""),
            )
            for row in self.repository.officer_options(filters)
        ]
        selected_codes = {officer.code for officer in self.selector.selected_officers() if officer.code}
        self.selector.set_officers(officers, selected_codes=selected_codes)
        self._selector_cache_key = cache_key

    def _base_filters(self) -> OfficerCenterFilters:
        return self._base_filters_provider()

    def refresh_query_filters(self) -> OfficerCenterFilters:
        filters = self._base_filters()
        selected = tuple(officer.code or officer.raw_name for officer in self.selector.selected_officers())
        return replace(filters, selected_officers=selected)


def _dashboard_metrics(kpis: dict[str, object]) -> list[KpiMetric]:
    occurrences = int(float(kpis.get("officer_customer_occurrence_count") or 0))
    active_officers = int(float(kpis.get("active_officer_count") or 0))
    average_customer = kpis.get("average_customer_per_officer")
    average_customer_tooltip = (
        "Tổng lượt khách hàng theo CBTD: "
        f"{_format_integer_vn(occurrences)}\n"
        f"Số CBTD có dư nợ: {_format_integer_vn(active_officers)}\n"
        f"Bình quân: {_format_decimal_vn(average_customer)} khách hàng/CBTD"
    )
    return [
        KpiMetric("Số CBTD có dư nợ", kpis.get("active_officer_count"), "count"),
        KpiMetric("Tổng dư nợ", kpis.get("total_balance"), "money"),
        KpiMetric("Tổng lượt KH theo CBTD", kpis.get("officer_customer_occurrence_count"), "count"),
        KpiMetric("Số KH duy nhất", kpis.get("unique_customer_count"), "count"),
        KpiMetric("Dư nợ bình quân/CBTD", kpis.get("average_balance_per_officer"), "money"),
        KpiMetric("KH bình quân/CBTD", average_customer, "number", tooltip=average_customer_tooltip),
        KpiMetric("Lãi suất bình quân", kpis.get("average_rate"), "percent"),
        KpiMetric("NIM trước ĐC", kpis.get("nim_before"), "percent"),
        KpiMetric("NIM sau ĐC", kpis.get("nim_after"), "percent"),
        KpiMetric("Nợ cần chú ý", kpis.get("attention_balance"), "money"),
        KpiMetric("Nợ xấu", kpis.get("bad_debt_balance"), "money"),
        KpiMetric("Tỷ lệ nợ cần chú ý", kpis.get("attention_ratio"), "percent"),
        KpiMetric("Tỷ lệ nợ xấu", kpis.get("bad_debt_ratio"), "percent"),
        KpiMetric("CBTD có nợ nhóm 2", kpis.get("attention_officer_count"), "count"),
        KpiMetric("CBTD có nợ xấu", kpis.get("bad_debt_officer_count"), "count"),
    ]


def _dashboard_placeholder_metrics() -> list[KpiMetric]:
    return [KpiMetric(label, None, "text") for label in (
        "Số CBTD có dư nợ",
        "Tổng dư nợ",
        "Tổng lượt KH theo CBTD",
        "Số KH duy nhất",
        "Dư nợ bình quân/CBTD",
        "KH bình quân/CBTD",
        "Lãi suất bình quân",
        "NIM trước ĐC",
        "NIM sau ĐC",
        "Nợ cần chú ý",
        "Nợ xấu",
        "Tỷ lệ nợ cần chú ý",
        "Tỷ lệ nợ xấu",
        "CBTD có nợ nhóm 2",
        "CBTD có nợ xấu",
    )]


def _format_integer_vn(value: object) -> str:
    try:
        return f"{int(float(value or 0)):,}".replace(",", ".")
    except (TypeError, ValueError):
        return "0"


def _format_decimal_vn(value: object) -> str:
    if value in (None, ""):
        return "—"
    try:
        number = float(value or 0)
    except (TypeError, ValueError):
        return "—"
    return f"{number:,.2f}".replace(",", "_").replace(".", ",").replace("_", ".")


def _trend_points(rows: object) -> tuple[tuple[str, float], ...]:
    return tuple((str(row.get("period") or ""), float(row.get("value") or 0)) for row in (rows or ()))


def _columns_include_term_structure(columns: tuple[ColumnSpec, ...]) -> bool:
    fields = {field for field, _label, _kind in columns}
    return bool(fields & {"short_term_balance", "medium_long_term_balance", "other_balance", "medium_long_ratio"})


def _has_incomplete_term_rows(rows: object) -> bool:
    return any(row.get("term_structure_available") is False for row in (rows or ()))


def _default_widths(columns: tuple[ColumnSpec, ...]) -> tuple[int, ...]:
    output: list[int] = []
    for field, _label, kind in columns:
        if field == "rank":
            output.append(55)
        elif kind in {"money", "money_signed", "term_money_or_dash"}:
            output.append(125)
        elif kind in {"percent_point_signed"}:
            output.append(130)
        elif kind.startswith("percent") or kind == "term_percent_or_dash":
            output.append(95)
        elif "name" in field:
            output.append(170)
        elif "code" in field:
            output.append(105)
        elif kind == "integer":
            output.append(85)
        else:
            output.append(120)
    return tuple(output)
