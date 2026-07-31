from __future__ import annotations

from dataclasses import replace
from pathlib import Path

from PySide6.QtCore import Qt
from PySide6.QtWidgets import QFileDialog, QComboBox, QGridLayout, QHBoxLayout, QLabel, QMessageBox, QScrollArea, QSizePolicy, QVBoxLayout, QWidget

from agribank_v3.features.credit.summary.customer import chart_service
from agribank_v3.features.credit.summary.customer.async_query import AsyncQueryController
from agribank_v3.features.credit.summary.customer.charts import (
    CustomerDonutChart,
    CustomerLineChart,
)
from agribank_v3.features.credit.summary.customer.customer_detail_window import CustomerDetailWindow
from agribank_v3.features.credit.summary.customer.export_service import (
    TOP_BALANCE_COLUMNS,
    TOP_MOVEMENT_COLUMNS,
    export_customer_dashboard,
    export_top_customer_balance_rows,
    export_top_customer_movement_rows,
    suggested_customer_export_name,
)
from agribank_v3.features.credit.summary.customer.filters import CustomerFilters
from agribank_v3.features.credit.summary.customer.formatters import format_money_vn, format_percent_vn
from agribank_v3.features.credit.summary.customer.period_validation import validate_dashboard_period_filters
from agribank_v3.features.credit.summary.customer.repository import CustomerRepository
from agribank_v3.features.credit.summary.customer.table_models import CustomerTableModel
from agribank_v3.features.credit.summary.customer.widgets import (
    CustomerTableView,
    KpiMetric,
    MetricGrid,
    QueryStateBanner,
    combo_box,
    configure_combo_popup_width,
    current_data,
    secondary_button,
)


DASHBOARD_MAIN_KPI_LABELS = (
    "Số khách hàng còn dư nợ",
    "Tổng dư nợ",
    "Dư nợ ngắn hạn",
    "Dư nợ trung/dài hạn",
    "Tỷ lệ trung/dài hạn",
    "Lãi suất bình quân",
    "NIM trước ĐC",
    "NIM sau ĐC",
)

DASHBOARD_SECONDARY_KPI_LABELS = (
    "Dư nợ chưa phân loại",
    "Khách hàng vay mới",
    "Khách hàng tất toán",
    "Khách hàng tăng dư nợ",
    "Khách hàng giảm dư nợ",
    "Khách hàng nhiều cán bộ quản lý",
    "Khách hàng có override cán bộ",
)

DASHBOARD_METRIC_LABELS = DASHBOARD_MAIN_KPI_LABELS + DASHBOARD_SECONDARY_KPI_LABELS


class CustomerDashboardTab(QWidget):
    def __init__(self, repository: CustomerRepository, filters_provider, parent=None) -> None:
        super().__init__(parent)
        self.repository = repository
        self.filters_provider = filters_provider
        self.query_controller = AsyncQueryController(self, max_cache_entries=32)
        self.balance_chart_controller = AsyncQueryController(self, max_cache_entries=32)
        self.metric_chart_controller = AsyncQueryController(self, max_cache_entries=32)
        self.customer_count_chart_controller = AsyncQueryController(self, max_cache_entries=32)
        self.detail_windows: list[CustomerDetailWindow] = []
        layout = QVBoxLayout(self)
        layout.setSpacing(6)
        self.dashboard_toolbar = QWidget()
        self.dashboard_toolbar.setObjectName("CustomerDashboardToolbar")
        actions = QHBoxLayout(self.dashboard_toolbar)
        actions.setContentsMargins(0, 0, 0, 0)
        actions.setSpacing(7)
        self.refresh_button = secondary_button("Làm mới")
        self.refresh_button.clicked.connect(lambda: self.refresh(use_cache=False))
        self.export_button = secondary_button("Xuất Excel")
        self.export_button.clicked.connect(self.export_excel)
        for button in (self.refresh_button, self.export_button):
            button.setSizePolicy(QSizePolicy.Policy.Preferred, QSizePolicy.Policy.Fixed)
        actions.addWidget(self.refresh_button)
        actions.addWidget(self.export_button)
        actions.addStretch(1)
        layout.addWidget(self.dashboard_toolbar)
        self.state_banner = QueryStateBanner()
        self.state_banner.retryRequested.connect(lambda: self.refresh(use_cache=False))
        layout.addWidget(self.state_banner)
        self.metrics = MetricGrid()
        layout.addWidget(self.metrics)
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.chart_container = QWidget()
        self.chart_grid = QGridLayout(self.chart_container)
        self.chart_grid.setContentsMargins(0, 0, 0, 0)
        self.chart_grid.setSpacing(12)
        scroll.setWidget(self.chart_container)
        layout.addWidget(scroll, stretch=1)
        self.balance_chart = CustomerLineChart("Xu hướng dư nợ theo kỳ", value_kind="money")
        self.balance_group_combo = _chart_combo(
            (
                ("Phân nhóm: Tổng hợp", "total"),
                ("Phân nhóm: Theo chi nhánh", "branch"),
                ("Phân nhóm: Theo loại khách hàng", "customer_type"),
            )
        )
        self.balance_group_combo.currentIndexChanged.connect(self.refresh_balance_chart)
        self.balance_chart.add_header_widget(self.balance_group_combo)
        self.balance_chart.set_save_name("Customer_TotalBalance_Total")
        self.term_chart = CustomerDonutChart("Cơ cấu kỳ hạn", value_kind="money")
        self.customer_count_chart = CustomerLineChart("Số lượng khách hàng còn dư nợ theo kỳ", value_kind="number")
        self.customer_count_chart.set_save_name("Customer_ActiveCustomerCount")
        self.movement_chart = self.customer_count_chart
        self.nim_chart = CustomerLineChart("Xu hướng lãi suất bình quân", value_kind="percent")
        self.metric_combo = _chart_combo(
            (
                ("Chỉ tiêu: Lãi suất bình quân", "average_rate"),
                ("Chỉ tiêu: NIM trước điều chỉnh", "nim_before"),
                ("Chỉ tiêu: NIM sau điều chỉnh", "nim_after"),
            )
        )
        self.metric_combo.currentIndexChanged.connect(self.refresh_metric_chart)
        self.nim_chart.add_header_widget(self.metric_combo)
        self.nim_chart.set_save_name("Customer_AverageRate")
        self._chart_widgets = (
            self.balance_chart,
            self.term_chart,
            self.nim_chart,
            self.customer_count_chart,
        )
        for chart in self._chart_widgets:
            chart.setMinimumWidth(480)
        self.top_balance_model = CustomerTableModel(TOP_BALANCE_COLUMNS, self)
        self.top_movement_model = CustomerTableModel(TOP_MOVEMENT_COLUMNS, self)
        self.top_balance_section, self.top_balance_limit_combo, self.top_balance_table = self._build_top_balance_section()
        self.top_movement_section, self.top_movement_mode_combo, self.top_movement_limit_combo, self.top_movement_table = self._build_top_movement_section()
        self.top_n_combo = self.top_balance_limit_combo
        self._data_widgets = self._chart_widgets + (self.top_balance_section, self.top_movement_section)
        self._arrange_charts()

    def refresh(self, *args, use_cache: bool = True) -> None:
        filters = self.filters_provider()
        periods = self.repository.distinct_periods()
        validation = validate_dashboard_period_filters(periods, filters)
        if not validation.valid:
            self.set_empty_state(validation.reason)
            return
        filters = replace(
            filters,
            period_from=validation.period_from,
            period_to=validation.period_to,
            current_period=validation.report_period,
        )
        previous_period, report_period = self._compare_periods(filters)
        balance_top_n = int(self.top_balance_limit_combo.currentData() or 10)
        movement_top_n = int(self.top_movement_limit_combo.currentData() or 10)
        top_mode = current_data(self.top_movement_mode_combo) or "increase"
        cache_key = ("dashboard", filters, previous_period, "report_period", report_period, balance_top_n, movement_top_n, top_mode)
        self.query_controller.run(
            "customer_dashboard",
            lambda: self._load_payload(filters, previous_period, report_period, balance_top_n, movement_top_n, top_mode),
            self._apply_payload,
            self._query_failed,
            cache_key=cache_key,
            use_cache=use_cache,
            state_callback=self._state_changed,
        )
        self.refresh_balance_chart(use_cache=use_cache)
        self.refresh_metric_chart(use_cache=use_cache)
        self.refresh_customer_count_chart(use_cache=use_cache)

    def refresh_balance_chart(self, *args, use_cache: bool = True) -> None:
        filters = self.filters_provider()
        validation = validate_dashboard_period_filters(self.repository.distinct_periods(), filters)
        if not validation.valid:
            self.balance_chart_controller.cancel_pending()
            self.balance_chart.set_empty("Chưa có dữ liệu để hiển thị.")
            return
        filters = replace(filters, period_from=validation.period_from, period_to=validation.period_to, current_period=validation.report_period)
        chart_filters = filters.without_exact_period()
        group_by = current_data(self.balance_group_combo) or "total"
        save_suffix = {"branch": "ByBranch", "customer_type": "ByCustomerType"}.get(group_by, "Total")
        self.balance_chart.set_title("Xu hướng dư nợ theo kỳ")
        self.balance_chart.set_save_name(f"Customer_TotalBalance_{save_suffix}")
        cache_key = ("dashboard_balance_chart", chart_filters, chart_filters.period_from, chart_filters.period_to, group_by)
        self.balance_chart_controller.run(
            "customer_dashboard_balance_chart",
            lambda: self.repository.get_total_balance_trend(filters, filters.period_from, filters.period_to, group_by=group_by),
            lambda rows, selected_group=group_by: self._apply_balance_chart(rows, selected_group),
            lambda exc: self.balance_chart.set_error("Không tải được xu hướng dư nợ."),
            cache_key=cache_key,
            use_cache=use_cache,
            state_callback=lambda state, _message: self._chart_state_changed(self.balance_chart, state),
        )

    def refresh_metric_chart(self, *args, use_cache: bool = True) -> None:
        filters = self.filters_provider()
        validation = validate_dashboard_period_filters(self.repository.distinct_periods(), filters)
        if not validation.valid:
            self.metric_chart_controller.cancel_pending()
            self.nim_chart.set_empty("Chưa có dữ liệu để hiển thị.")
            return
        filters = replace(filters, period_from=validation.period_from, period_to=validation.period_to, current_period=validation.report_period)
        chart_filters = filters.without_exact_period()
        metric = current_data(self.metric_combo) or "average_rate"
        title = chart_service.METRIC_TITLES.get(metric, chart_service.METRIC_TITLES["average_rate"])
        save_name = {
            "average_rate": "Customer_AverageRate",
            "nim_before": "Customer_NimBefore",
            "nim_after": "Customer_NimAfter",
        }.get(metric, "Customer_AverageRate")
        self.nim_chart.set_title(title)
        self.nim_chart.set_save_name(save_name)
        cache_key = ("dashboard_metric_chart", chart_filters, chart_filters.period_from, chart_filters.period_to, metric)
        self.metric_chart_controller.run(
            "customer_dashboard_metric_chart",
            lambda: self.repository.get_customer_metric_trend(filters, filters.period_from, filters.period_to, metric=metric),
            lambda rows, selected_metric=metric: self._apply_metric_chart(rows, selected_metric),
            lambda exc: self.nim_chart.set_error("Không tải được xu hướng NIM/lãi suất."),
            cache_key=cache_key,
            use_cache=use_cache,
            state_callback=lambda state, _message: self._chart_state_changed(self.nim_chart, state),
        )

    def refresh_customer_count_chart(self, *args, use_cache: bool = True) -> None:
        filters = self.filters_provider()
        validation = validate_dashboard_period_filters(self.repository.distinct_periods(), filters)
        if not validation.valid:
            self.customer_count_chart_controller.cancel_pending()
            self.customer_count_chart.set_empty("Chưa có dữ liệu để hiển thị.")
            return
        filters = replace(filters, period_from=validation.period_from, period_to=validation.period_to, current_period=validation.report_period)
        chart_filters = filters.without_exact_period()
        self.customer_count_chart.set_title("Số lượng khách hàng còn dư nợ theo kỳ")
        self.customer_count_chart.set_save_name("Customer_ActiveCustomerCount")
        cache_key = ("dashboard_active_customer_count_chart", chart_filters, chart_filters.period_from, chart_filters.period_to)
        self.customer_count_chart_controller.run(
            "customer_dashboard_active_customer_count_chart",
            lambda: self.repository.get_active_customer_count_trend(filters, filters.period_from, filters.period_to),
            self._apply_customer_count_chart,
            lambda exc: self.customer_count_chart.set_error("Không tải được số lượng khách hàng còn dư nợ."),
            cache_key=cache_key,
            use_cache=use_cache,
            state_callback=lambda state, _message: self._chart_state_changed(self.customer_count_chart, state),
        )

    def export_excel(self) -> None:
        if not self.repository.has_period_data():
            self.state_banner.set_empty("Chưa có dữ liệu để xuất.")
            return
        path, _ = QFileDialog.getSaveFileName(
            self,
            "Xuất Dashboard khách hàng",
            suggested_customer_export_name("TongQuanKhachHang"),
            "Excel (*.xlsx)",
        )
        if not path:
            return
        try:
            output = export_customer_dashboard(self.repository, self.filters_provider(), Path(path))
        except Exception as exc:
            QMessageBox.warning(self, "Xuất Dashboard khách hàng", str(exc))
            return
        QMessageBox.information(self, "Xuất Dashboard khách hàng", f"Đã xuất: {output}")

    def invalidate_cache(self) -> None:
        self.query_controller.invalidate_cache()
        self.balance_chart_controller.invalidate_cache()
        self.metric_chart_controller.invalidate_cache()
        self.customer_count_chart_controller.invalidate_cache()

    def cancel_queries(self) -> None:
        for controller in (
            self.query_controller,
            self.balance_chart_controller,
            self.metric_chart_controller,
            self.customer_count_chart_controller,
        ):
            controller.cancel_pending()

    def wait_for_queries(self, timeout_ms: int = 5000) -> None:
        for controller in (
            self.query_controller,
            self.balance_chart_controller,
            self.metric_chart_controller,
            self.customer_count_chart_controller,
        ):
            controller.wait_for_idle(timeout_ms)

    def closeEvent(self, event) -> None:
        self.cancel_queries()
        super().closeEvent(event)

    def set_empty_state(self, message: str = "Chưa có dữ liệu khách hàng.") -> None:
        self.cancel_queries()
        self.state_banner.set_empty(
            message
            or "Chưa có dữ liệu khách hàng. Hãy nhập thư mục FTP Loan tại chức năng NIM Dư nợ để tạo dữ liệu khách hàng."
        )
        self.metrics.set_empty(DASHBOARD_METRIC_LABELS)
        self.term_chart.set_empty("Chưa có dữ liệu để hiển thị.")
        self.balance_chart.set_empty("Chưa có dữ liệu để hiển thị.")
        self.nim_chart.set_empty("Chưa có dữ liệu để hiển thị.")
        self.customer_count_chart.set_empty("Chưa có dữ liệu để hiển thị.")
        self.top_balance_model.set_rows([])
        self.top_movement_model.set_rows([])
        self.top_movement_period_label.setText("Chưa có dữ liệu để so sánh")
        for button_name in ("export_button", "top_balance_export_button", "top_movement_export_button"):
            button = getattr(self, button_name, None)
            if button is not None:
                button.setEnabled(False)
                button.setToolTip("Chưa có dữ liệu để xuất")

    def resizeEvent(self, event) -> None:
        super().resizeEvent(event)
        self._arrange_charts()

    def _load_payload(
        self,
        filters: CustomerFilters,
        previous_period: str,
        report_period: str,
        balance_top_n: int,
        movement_top_n: int,
        top_mode: str,
    ) -> dict[str, object]:
        metrics = self.repository.get_dashboard_kpis(filters, report_period)
        movement = {}
        top_movement_rows: list[dict[str, object]] = []
        if previous_period and report_period:
            movement_filters = replace(filters, movement_status="")
            movement = self.repository.movement_kpis(previous_period, report_period, movement_filters)
            top_movement_rows = self.repository.get_top_customer_movements(
                movement_filters,
                previous_period,
                report_period,
                direction=top_mode,
                limit=movement_top_n,
            )
        return {
            "metrics": metrics,
            "movement": movement,
            "previous_period": previous_period,
            "current_period": report_period,
            "top_balance_rows": self.repository.get_top_customers_by_balance(filters, report_period, balance_top_n),
            "top_movement_rows": top_movement_rows,
            "top_mode": top_mode,
            "balance_top_n": balance_top_n,
            "movement_top_n": movement_top_n,
        }

    def _apply_payload(self, payload: dict[str, object]) -> None:
        for button_name in ("export_button", "top_balance_export_button", "top_movement_export_button"):
            button = getattr(self, button_name, None)
            if button is not None:
                button.setEnabled(True)
                button.setToolTip("")
        metrics = dict(payload.get("metrics") or {})
        movement = dict(payload.get("movement") or {})
        customer_count = int(metrics.get("customer_count") or 0)
        total_balance = metrics.get("total_balance", 0)
        medium_long_balance = metrics.get("medium_long_term_balance", 0)
        medium_long_ratio = metrics.get("medium_long_ratio", 0)
        if customer_count <= 0:
            self.state_banner.set_empty()
        else:
            self.state_banner.clear()
        self.metrics.set_metrics(
            [
                KpiMetric("Số khách hàng còn dư nợ", customer_count, "count"),
                KpiMetric("Tổng dư nợ", total_balance, "money"),
                KpiMetric("Dư nợ ngắn hạn", metrics.get("short_term_balance", 0), "money"),
                KpiMetric("Dư nợ trung/dài hạn", medium_long_balance, "money"),
                KpiMetric(
                    "Tỷ lệ trung/dài hạn",
                    medium_long_ratio,
                    "percentage",
                    tooltip=(
                        "Tỷ lệ trung/dài hạn\n"
                        f"{format_money_vn(medium_long_balance)} / {format_money_vn(total_balance)} = "
                        f"{format_percent_vn(medium_long_ratio, empty='—')}"
                    ),
                ),
                KpiMetric("Lãi suất bình quân", metrics.get("average_rate", 0), "percentage"),
                KpiMetric("NIM trước ĐC", metrics.get("nim_before", 0), "percentage"),
                KpiMetric("NIM sau ĐC", metrics.get("nim_after", 0), "percentage"),
                KpiMetric("Dư nợ chưa phân loại", metrics.get("other_balance", 0), "money", group="secondary"),
                KpiMetric("Khách hàng vay mới", movement.get("new_customer_count", 0), "count", group="secondary"),
                KpiMetric("Khách hàng tất toán", movement.get("paid_off_customer_count", 0), "count", group="secondary"),
                KpiMetric("Khách hàng tăng dư nợ", movement.get("increased_customer_count", 0), "count", group="secondary"),
                KpiMetric("Khách hàng giảm dư nợ", movement.get("decreased_customer_count", 0), "count", group="secondary"),
                KpiMetric("Khách hàng nhiều cán bộ quản lý", metrics.get("multiple_officer_customer_count", 0), "count", group="secondary"),
                KpiMetric("Khách hàng có override cán bộ", metrics.get("override_customer_count", 0), "count", group="secondary"),
            ]
        )
        self.term_chart.set_slices(chart_service.dashboard_chart_dataset_term_structure(metrics))
        self.top_balance_model.set_rows(_rank_rows(list(payload.get("top_balance_rows") or [])))
        top_rows = _rank_rows(list(payload.get("top_movement_rows") or []))
        self.top_movement_model.set_rows(top_rows)
        top_mode = str(payload.get("top_mode") or "increase")
        self.top_movement_title.setText("Top khách hàng giảm dư nợ" if top_mode == "decrease" else "Top khách hàng tăng dư nợ")
        previous = str(payload.get("previous_period") or "")
        current = str(payload.get("current_period") or "")
        self.top_movement_period_label.setText(f"So sánh {previous} → {current}" if previous and current else "Chưa đủ hai kỳ để so sánh")

    def _query_failed(self, exc: Exception) -> None:
        self.metrics.set_metrics(_dashboard_placeholder_metrics(None))
        self.term_chart.set_error("Không tải được dữ liệu biểu đồ.")
        self.top_balance_model.set_rows([])
        self.top_movement_model.set_rows([])

    def _state_changed(self, state: str, message: str) -> None:
        if state == "loading":
            self.state_banner.set_loading()
            self.metrics.set_metrics(_dashboard_placeholder_metrics("…", value_type="loading"))
            self.term_chart.set_loading()
            self.top_balance_model.set_rows([])
            self.top_movement_model.set_rows([])
        elif state == "error":
            self.state_banner.set_error(message or "Không tải được dữ liệu Dashboard.")
        elif state == "ready":
            self.state_banner.clear()

    def _apply_balance_chart(self, rows: list[dict[str, object]], group_by: str) -> None:
        dataset = chart_service.dashboard_chart_dataset_balance_trend(rows)
        if not dataset:
            self.balance_chart.set_empty("Không có dữ liệu dư nợ phù hợp với bộ lọc.")
            return
        self.balance_chart.value_kind = "money"
        self.balance_chart.set_series(dataset)

    def _apply_metric_chart(self, rows: list[dict[str, object]], metric: str) -> None:
        title = chart_service.METRIC_TITLES.get(metric, chart_service.METRIC_TITLES["average_rate"])
        self.nim_chart.set_title(title)
        dataset = chart_service.dashboard_chart_dataset_metric_trend(rows, metric)
        if not dataset:
            self.nim_chart.set_empty("Không có dữ liệu NIM/lãi suất phù hợp với bộ lọc.")
            return
        self.nim_chart.value_kind = "percent"
        self.nim_chart.set_series(dataset)

    def _apply_customer_count_chart(self, rows: list[dict[str, object]]) -> None:
        dataset = chart_service.dashboard_chart_dataset_active_customer_count(rows)
        if not dataset or not dataset[0][1]:
            self.customer_count_chart.set_empty("Không có khách hàng còn dư nợ phù hợp với bộ lọc.")
            return
        self.customer_count_chart.value_kind = "number"
        self.customer_count_chart.set_series(dataset)

    def _chart_state_changed(self, chart: CustomerLineChart, state: str) -> None:
        if state == "loading":
            chart.set_loading()

    def _compare_periods(self, filters: CustomerFilters) -> tuple[str, str]:
        periods = self.repository.distinct_periods()
        current = filters.current_period or filters.period_to or (periods[-1] if periods else "")
        previous = filters.compare_period
        if current and not previous and current in periods:
            index = periods.index(current)
            previous = periods[index - 1] if index > 0 else ""
        return previous, current

    def _build_top_balance_section(self) -> tuple[QWidget, QComboBox, CustomerTableView]:
        section = QWidget()
        section.setObjectName("DashboardTopTableSection")
        layout = QVBoxLayout(section)
        layout.setContentsMargins(0, 0, 0, 0)
        header = QHBoxLayout()
        title = QLabel("Top khách hàng dư nợ lớn")
        title.setObjectName("SectionTitle")
        combo = _top_n_combo()
        combo.currentIndexChanged.connect(self.refresh)
        self.top_balance_export_button = secondary_button("Xuất Excel")
        self.top_balance_export_button.clicked.connect(self.export_top_balance_excel)
        self.top_balance_export_button.setSizePolicy(QSizePolicy.Policy.Preferred, QSizePolicy.Policy.Fixed)
        header.setSpacing(7)
        header.addWidget(title)
        header.addStretch()
        header.addWidget(combo)
        header.addWidget(self.top_balance_export_button)
        layout.addLayout(header)
        table = CustomerTableView()
        table.setModel(self.top_balance_model)
        table.doubleClicked.connect(lambda index: self._open_customer_detail_from_model(self.top_balance_model, index))
        table.setMinimumHeight(260)
        table.apply_default_widths((50, 110, 210, 90, 90, 180, 130, 130, 140, 110, 100))
        layout.addWidget(table, stretch=1)
        return section, combo, table

    def _build_top_movement_section(self) -> tuple[QWidget, QComboBox, QComboBox, CustomerTableView]:
        section = QWidget()
        section.setObjectName("DashboardTopTableSection")
        layout = QVBoxLayout(section)
        layout.setContentsMargins(0, 0, 0, 0)
        header = QHBoxLayout()
        self.top_movement_title = QLabel("Top khách hàng tăng dư nợ")
        self.top_movement_title.setObjectName("SectionTitle")
        mode_combo = combo_box("Top tăng dư nợ", minimum_width=142, maximum_width=190)
        mode_combo.clear()
        mode_combo.addItem("Top tăng dư nợ", "increase")
        mode_combo.addItem("Top giảm dư nợ", "decrease")
        configure_combo_popup_width(mode_combo, minimum_popup_width=190)
        mode_combo.currentIndexChanged.connect(self.refresh)
        limit_combo = _top_n_combo()
        limit_combo.currentIndexChanged.connect(self.refresh)
        self.top_movement_export_button = secondary_button("Xuất Excel")
        self.top_movement_export_button.clicked.connect(self.export_top_movement_excel)
        self.top_movement_export_button.setSizePolicy(QSizePolicy.Policy.Preferred, QSizePolicy.Policy.Fixed)
        header.setSpacing(7)
        header.addWidget(self.top_movement_title)
        header.addStretch()
        header.addWidget(mode_combo)
        header.addWidget(limit_combo)
        header.addWidget(self.top_movement_export_button)
        layout.addLayout(header)
        self.top_movement_period_label = QLabel("Chưa đủ hai kỳ để so sánh")
        self.top_movement_period_label.setObjectName("MutedText")
        layout.addWidget(self.top_movement_period_label)
        table = CustomerTableView()
        table.setModel(self.top_movement_model)
        table.doubleClicked.connect(lambda index: self._open_customer_detail_from_model(self.top_movement_model, index))
        table.setMinimumHeight(260)
        table.apply_default_widths((50, 110, 210, 90, 90, 180, 130, 130, 130, 105, 120))
        layout.addWidget(table, stretch=1)
        return section, mode_combo, limit_combo, table

    def export_top_balance_excel(self) -> None:
        _previous, current = self._compare_periods(self.filters_provider())
        path, _ = QFileDialog.getSaveFileName(
            self,
            "Xuất Top khách hàng dư nợ lớn",
            suggested_customer_export_name(f"TopKhachHangDuNo_{current or 'TatCa'}"),
            "Excel (*.xlsx)",
        )
        if not path:
            return
        try:
            output = export_top_customer_balance_rows(self.top_balance_model.rows, Path(path))
        except Exception as exc:
            QMessageBox.warning(self, "Xuất Top khách hàng dư nợ lớn", str(exc))
            return
        QMessageBox.information(self, "Xuất Top khách hàng dư nợ lớn", f"Đã xuất: {output}")

    def export_top_movement_excel(self) -> None:
        filters = self.filters_provider()
        previous, current = self._compare_periods(filters)
        direction = current_data(self.top_movement_mode_combo) or "increase"
        title = "Top giảm dư nợ" if direction == "decrease" else "Top tăng dư nợ"
        path, _ = QFileDialog.getSaveFileName(
            self,
            f"Xuất {title}",
            suggested_customer_export_name(f"{'TopGiamDuNo' if direction == 'decrease' else 'TopTangDuNo'}_{previous}_{current}"),
            "Excel (*.xlsx)",
        )
        if not path:
            return
        try:
            output = export_top_customer_movement_rows(self.top_movement_model.rows, direction, Path(path))
        except Exception as exc:
            QMessageBox.warning(self, f"Xuất {title}", str(exc))
            return
        QMessageBox.information(self, f"Xuất {title}", f"Đã xuất: {output}")

    def _open_customer_detail_from_model(self, model: CustomerTableModel, index) -> None:
        row = model.raw_row(index.row())
        if not row:
            return
        period = str(row.get("current_period") or row.get("period") or self._compare_periods(self.filters_provider())[1])
        dialog = CustomerDetailWindow(
            self.repository,
            str(row.get("customer_code") or ""),
            period=period,
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

    def _arrange_charts(self) -> None:
        columns = 2 if self.width() >= 1080 else 1
        for index, chart in enumerate(self._data_widgets):
            row = index // columns
            column = index % columns
            self.chart_grid.addWidget(chart, row, column)
        for column in range(2):
            self.chart_grid.setColumnStretch(column, 1 if column < columns else 0)


def _dashboard_placeholder_metrics(value: object, *, value_type: str = "text") -> list[KpiMetric]:
    return [
        KpiMetric(label, value, value_type, group="main")
        for label in DASHBOARD_MAIN_KPI_LABELS
    ] + [
        KpiMetric(label, value, value_type, group="secondary")
        for label in DASHBOARD_SECONDARY_KPI_LABELS
    ]


def _top_n_combo() -> QComboBox:
    combo = combo_box("Top 10", minimum_width=84, maximum_width=96, minimum_contents_length=6)
    combo.clear()
    for value in (10, 20, 50):
        combo.addItem(f"Top {value}", value)
    configure_combo_popup_width(combo, minimum_popup_width=110)
    return combo


def _chart_combo(items: tuple[tuple[str, str], ...]) -> QComboBox:
    combo = combo_box("Chỉ tiêu", minimum_width=150, maximum_width=220)
    combo.clear()
    for label, value in items:
        combo.addItem(label, value)
    configure_combo_popup_width(combo, minimum_popup_width=240)
    return combo


def _rank_rows(rows: list[dict[str, object]]) -> list[dict[str, object]]:
    return [dict(row, rank=index) for index, row in enumerate(rows, start=1)]
