from __future__ import annotations

import getpass
import os
from pathlib import Path

from PySide6.QtCore import Qt
from PySide6.QtWidgets import QFileDialog, QGridLayout, QMessageBox, QTabWidget, QVBoxLayout, QDialog, QWidget

from agribank_v3.features.credit.summary.customer import chart_service
from agribank_v3.features.credit.summary.customer.async_query import AsyncQueryController
from agribank_v3.features.credit.summary.customer.charts import CustomerBarChart, CustomerLineChart
from agribank_v3.features.credit.summary.customer.export_service import (
    DETAIL_BALANCE_COLUMNS,
    DETAIL_NIM_COLUMNS,
    DETAIL_OFFICER_COLUMNS,
    export_customer_detail,
    suggested_customer_export_name,
)
from agribank_v3.features.credit.summary.customer.formatters import (
    format_customer_type,
    format_money_vn,
    format_percent_vn,
)
from agribank_v3.features.credit.summary.customer.officer_override_dialog import OfficerOverrideDialog
from agribank_v3.features.credit.summary.customer.repository import CustomerRepository
from agribank_v3.features.credit.summary.customer.table_models import CustomerTableModel
from agribank_v3.features.credit.summary.customer.widgets import (
    CompactToolbar,
    CustomerTableView,
    KpiMetric,
    MetricGrid,
    QueryStateBanner,
    fit_window_to_screen,
    primary_button,
    secondary_button,
)


DETAIL_MAIN_KPI_LABELS = (
    "Mã khách hàng",
    "Tên khách hàng",
    "Kỳ đang chọn",
    "Tổng dư nợ",
    "Lãi suất bình quân",
    "NIM trước ĐC",
    "NIM sau ĐC",
)

DETAIL_SECONDARY_KPI_LABELS = (
    "Loại khách hàng",
    "Chi nhánh",
    "Cán bộ quản lý hiệu lực",
    "Cán bộ chính từ import",
    "Trạng thái override",
    "Dư nợ ngắn hạn",
    "Dư nợ trung/dài hạn",
    "Dư nợ chưa phân loại",
    "Tỷ lệ trung/dài hạn",
)

DETAIL_METRIC_LABELS = DETAIL_MAIN_KPI_LABELS + DETAIL_SECONDARY_KPI_LABELS


class CustomerDetailWindow(QDialog):
    def __init__(
        self,
        repository: CustomerRepository,
        customer_code: str,
        *,
        period: str = "",
        initial_tab: int = 0,
        parent=None,
    ) -> None:
        super().__init__(parent)
        self.repository = repository
        self.customer_code = str(customer_code or "").strip()
        self.period = str(period or "").strip()
        self.query_controller = AsyncQueryController(self, max_cache_entries=16)
        self.setWindowTitle("Phân tích khách hàng - AgribankV3")
        self.setModal(False)
        self.setWindowFlag(Qt.WindowType.WindowMinMaxButtonsHint, True)
        fit_window_to_screen(
            self,
            width_ratio=0.82,
            height_ratio=0.84,
            max_width=1250,
            max_height=820,
            min_width=850,
            min_height=600,
        )
        layout = QVBoxLayout(self)
        layout.setSpacing(6)
        actions = CompactToolbar()
        override_button = primary_button("Cập nhật cán bộ quản lý")
        restore_button = secondary_button("Khôi phục theo dữ liệu import")
        export_button = secondary_button("Xuất Excel")
        refresh_button = secondary_button("Làm mới")
        override_button.clicked.connect(self.update_officer_override)
        restore_button.clicked.connect(self.restore_imported_officer)
        export_button.clicked.connect(self.export_excel)
        refresh_button.clicked.connect(lambda: self.refresh(use_cache=False))
        for button in (override_button, restore_button, export_button, refresh_button):
            actions.addWidget(button)
        layout.addWidget(actions)
        self.state_banner = QueryStateBanner()
        self.state_banner.retryRequested.connect(lambda: self.refresh(use_cache=False))
        layout.addWidget(self.state_banner)
        self.metrics = MetricGrid()
        layout.addWidget(self.metrics)
        self.tabs = QTabWidget()
        layout.addWidget(self.tabs, stretch=1)
        self.balance_model = CustomerTableModel(DETAIL_BALANCE_COLUMNS, self)
        self.term_model = CustomerTableModel(DETAIL_BALANCE_COLUMNS[:5] + (("medium_long_ratio", "Tỷ lệ trung/dài hạn", "percent"),), self)
        self.nim_model = CustomerTableModel(DETAIL_NIM_COLUMNS, self)
        self.officer_model = CustomerTableModel(DETAIL_OFFICER_COLUMNS, self)
        self.compare_model = CustomerTableModel(DETAIL_BALANCE_COLUMNS, self)
        self.balance_chart = CustomerLineChart("Xu hướng dư nợ", value_kind="money")
        self.term_money_chart = CustomerBarChart("Cơ cấu kỳ hạn theo kỳ", value_kind="money")
        self.term_ratio_chart = CustomerLineChart("Tỷ lệ trung/dài hạn", value_kind="percent")
        self.nim_chart = CustomerLineChart("NIM và lãi suất", value_kind="percent")
        self.compare_money_chart = CustomerBarChart("So sánh tăng/giảm dư nợ", value_kind="money")
        self.compare_percent_chart = CustomerLineChart("Tăng trưởng dư nợ", value_kind="percent")
        self._dual_chart_grids: list[tuple[QGridLayout, QWidget, QWidget]] = []
        self._add_chart_table_tab("Xu hướng dư nợ", self.balance_chart, self.balance_model)
        self._add_dual_chart_table_tab("Cơ cấu kỳ hạn", self.term_money_chart, self.term_ratio_chart, self.term_model)
        self._add_chart_table_tab("NIM và lãi suất", self.nim_chart, self.nim_model)
        self.officer_table = self._add_table_tab("Cán bộ quản lý", self.officer_model)
        self.officer_table.doubleClicked.connect(self._officer_row_double_clicked)
        self._add_dual_chart_table_tab("So sánh các kỳ", self.compare_money_chart, self.compare_percent_chart, self.compare_model)
        self.tabs.setCurrentIndex(max(0, min(initial_tab, self.tabs.count() - 1)))
        self.refresh()

    def refresh(self, *args, use_cache: bool = True) -> None:
        customer_code = self.customer_code
        period = self.period
        cache_key = ("customer_detail", customer_code, period)
        self.query_controller.run(
            "customer_detail",
            lambda: self._load_payload(customer_code, period),
            self._apply_payload,
            self._query_failed,
            cache_key=cache_key,
            use_cache=use_cache,
            state_callback=self._state_changed,
        )

    def update_officer_override(self) -> None:
        self.open_officer_override_dialog(self.period)

    def open_officer_override_dialog(self, period: str = "") -> None:
        selected_period = str(period or self.period).strip()
        detail = self.repository.customer_detail(self.customer_code, selected_period) or {}
        if not detail:
            detail = self.repository.customer_detail(self.customer_code, self.period) or {}
        dialog = OfficerOverrideDialog(
            self.repository,
            customer_code=self.customer_code,
            customer_name=str(detail.get("customer_name") or ""),
            period=str(detail.get("period") or selected_period or self.period),
            imported_officer_code=str(detail.get("imported_officer_code") or ""),
            imported_officer_name=str(detail.get("imported_officer_name") or ""),
            parent=self,
        )
        if dialog.exec() == QDialog.DialogCode.Accepted:
            self.invalidate_cache()
            self._notify_parent_cache_invalidated()
            self.refresh(use_cache=False)

    def restore_imported_officer(self) -> None:
        answer = QMessageBox.question(
            self,
            "Khôi phục theo dữ liệu import",
            "Bạn có chắc muốn ngừng áp dụng thông tin cán bộ đã điều chỉnh cho khách hàng này không?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        if answer != QMessageBox.StandardButton.Yes:
            return
        self.repository.restore_imported_officer(
            customer_code=self.customer_code,
            period=self.period,
            user_name=_current_user(),
            computer_name=os.environ.get("COMPUTERNAME", ""),
        )
        self.invalidate_cache()
        self._notify_parent_cache_invalidated()
        self.refresh(use_cache=False)

    def export_excel(self) -> None:
        path, _ = QFileDialog.getSaveFileName(
            self,
            "Xuất phân tích khách hàng",
            suggested_customer_export_name("ChiTietKhachHang"),
            "Excel (*.xlsx)",
        )
        if not path:
            return
        try:
            output = export_customer_detail(self.repository, self.customer_code, Path(path))
        except Exception as exc:
            QMessageBox.warning(self, "Xuất phân tích khách hàng", str(exc))
            return
        QMessageBox.information(self, "Xuất phân tích khách hàng", f"Đã xuất: {output}")

    def invalidate_cache(self) -> None:
        self.query_controller.invalidate_cache()

    def wait_for_queries(self) -> None:
        self.query_controller.wait_for_idle()

    def closeEvent(self, event) -> None:
        self.query_controller.cancel_pending()
        super().closeEvent(event)

    def resizeEvent(self, event) -> None:
        super().resizeEvent(event)
        self._arrange_dual_charts()

    def _load_payload(self, customer_code: str, period: str) -> dict[str, object]:
        detail = self.repository.customer_detail(customer_code, period) or {}
        history = self.repository.customer_history(customer_code)
        if history and not detail:
            detail = history[-1]
        officer_history = self.repository.customer_officer_history(customer_code)
        compare_rows = [row for row in history if row.get("difference") != ""]
        return {
            "detail": detail,
            "history": history,
            "officer_history": officer_history,
            "compare_rows": compare_rows,
            "charts": chart_service.customer_detail_chart_datasets(history),
        }

    def _apply_payload(self, payload: dict[str, object]) -> None:
        detail = dict(payload.get("detail") or {})
        history = list(payload.get("history") or [])
        if detail:
            self.period = str(detail.get("period") or self.period)
        total_balance = detail.get("total_balance", 0)
        medium_long_balance = detail.get("medium_long_term_balance", 0)
        medium_long_ratio = detail.get("medium_long_ratio", 0)
        self.metrics.set_metrics(
            [
                KpiMetric("Mã khách hàng", str(detail.get("customer_code") or self.customer_code), "text"),
                KpiMetric("Tên khách hàng", str(detail.get("customer_name") or ""), "text"),
                KpiMetric("Kỳ đang chọn", str(detail.get("period") or self.period), "text"),
                KpiMetric("Tổng dư nợ", total_balance, "money"),
                KpiMetric("Lãi suất bình quân", detail.get("average_rate", 0), "percentage"),
                KpiMetric("NIM trước ĐC", detail.get("nim_before", 0), "percentage"),
                KpiMetric("NIM sau ĐC", detail.get("nim_after", 0), "percentage"),
                KpiMetric("Loại khách hàng", format_customer_type(detail.get("customer_type")), "text", group="secondary"),
                KpiMetric("Chi nhánh", str(detail.get("branch_code") or ""), "text", group="secondary"),
                KpiMetric("Cán bộ quản lý hiệu lực", str(detail.get("effective_officer_name") or ""), "text", group="secondary"),
                KpiMetric("Cán bộ chính từ import", str(detail.get("imported_officer_name") or ""), "text", group="secondary"),
                KpiMetric("Trạng thái override", str(detail.get("override_status") or ""), "text", group="secondary"),
                KpiMetric("Dư nợ ngắn hạn", detail.get("short_term_balance", 0), "money", group="secondary"),
                KpiMetric("Dư nợ trung/dài hạn", medium_long_balance, "money", group="secondary"),
                KpiMetric("Dư nợ chưa phân loại", detail.get("other_balance", 0), "money", group="secondary"),
                KpiMetric(
                    "Tỷ lệ trung/dài hạn",
                    medium_long_ratio,
                    "percentage",
                    group="secondary",
                    tooltip=(
                        "Tỷ lệ trung/dài hạn\n"
                        f"{format_money_vn(medium_long_balance)} / {format_money_vn(total_balance)} = "
                        f"{format_percent_vn(medium_long_ratio, empty='—')}"
                    ),
                ),
            ]
        )
        officer_history = list(payload.get("officer_history") or [])
        compare_rows = list(payload.get("compare_rows") or [])
        self.balance_model.set_rows(history)
        self.term_model.set_rows(history)
        self.nim_model.set_rows(history)
        self.officer_model.set_rows(officer_history)
        self.compare_model.set_rows(compare_rows)
        charts = dict(payload.get("charts") or {})
        self.balance_chart.set_series(charts.get("balance", ()))
        self.term_money_chart.set_series(charts.get("term_money", ()))
        self.term_ratio_chart.set_series(charts.get("term_ratio", ()))
        self.nim_chart.set_series(charts.get("nim_rates", ()))
        self.compare_money_chart.set_series(charts.get("compare_money", ()))
        self.compare_percent_chart.set_series(charts.get("compare_percent", ()))
        if history:
            self.state_banner.clear()
        else:
            self.state_banner.set_empty("Không có dữ liệu lịch sử cho khách hàng này.")

    def _query_failed(self, exc: Exception) -> None:
        self.metrics.set_metrics(_detail_placeholder_metrics(None))
        self.balance_model.set_rows([])
        self.term_model.set_rows([])
        self.nim_model.set_rows([])
        self.officer_model.set_rows([])
        self.compare_model.set_rows([])
        for chart in (
            self.balance_chart,
            self.term_money_chart,
            self.term_ratio_chart,
            self.nim_chart,
            self.compare_money_chart,
            self.compare_percent_chart,
        ):
            chart.set_error("Không tải được dữ liệu biểu đồ.")

    def _state_changed(self, state: str, message: str) -> None:
        if state == "loading":
            self.state_banner.set_loading()
            self.metrics.set_metrics(_detail_placeholder_metrics("…", value_type="loading"))
            for chart in (
                self.balance_chart,
                self.term_money_chart,
                self.term_ratio_chart,
                self.nim_chart,
                self.compare_money_chart,
                self.compare_percent_chart,
            ):
                chart.set_loading()
        elif state == "error":
            self.state_banner.set_error(message or "Không tải được dữ liệu khách hàng.")

    def _add_chart_table_tab(self, label: str, chart: QWidget, model: CustomerTableModel) -> None:
        tab = QWidget()
        layout = QVBoxLayout(tab)
        layout.addWidget(chart)
        table = CustomerTableView()
        table.setModel(model)
        layout.addWidget(table, stretch=1)
        self.tabs.addTab(tab, label)

    def _add_dual_chart_table_tab(self, label: str, first_chart: QWidget, second_chart: QWidget, model: CustomerTableModel) -> None:
        tab = QWidget()
        layout = QVBoxLayout(tab)
        chart_area = QWidget()
        chart_grid = QGridLayout(chart_area)
        chart_grid.setContentsMargins(0, 0, 0, 0)
        chart_grid.setSpacing(10)
        layout.addWidget(chart_area)
        self._dual_chart_grids.append((chart_grid, first_chart, second_chart))
        self._arrange_dual_chart_grid(chart_grid, first_chart, second_chart)
        table = CustomerTableView()
        table.setModel(model)
        layout.addWidget(table, stretch=1)
        self.tabs.addTab(tab, label)

    def _add_table_tab(self, label: str, model: CustomerTableModel) -> CustomerTableView:
        tab = QWidget()
        layout = QVBoxLayout(tab)
        table = CustomerTableView()
        table.setModel(model)
        layout.addWidget(table, stretch=1)
        self.tabs.addTab(tab, label)
        return table

    def _officer_row_double_clicked(self, index) -> None:
        row = self.officer_model.raw_row(index.row())
        period = str(row.get("period") or self.period).strip()
        self.open_officer_override_dialog(period)

    def _notify_parent_cache_invalidated(self) -> None:
        parent = self.parent()
        while parent is not None:
            if hasattr(parent, "invalidate_customer_caches"):
                parent.invalidate_customer_caches()
                return
            if hasattr(parent, "invalidate_cache"):
                parent.invalidate_cache()
            parent = parent.parent()

    def _arrange_dual_charts(self) -> None:
        for grid, first_chart, second_chart in self._dual_chart_grids:
            self._arrange_dual_chart_grid(grid, first_chart, second_chart)

    def _arrange_dual_chart_grid(self, grid: QGridLayout, first_chart: QWidget, second_chart: QWidget) -> None:
        columns = 2 if self.width() >= 1050 else 1
        grid.addWidget(first_chart, 0, 0)
        grid.addWidget(second_chart, 0 if columns == 2 else 1, 1 if columns == 2 else 0)
        grid.setColumnStretch(0, 1)
        grid.setColumnStretch(1, 1 if columns == 2 else 0)


def _detail_placeholder_metrics(value: object, *, value_type: str = "text") -> list[KpiMetric]:
    return [
        KpiMetric(label, value, value_type, group="main")
        for label in DETAIL_MAIN_KPI_LABELS
    ] + [
        KpiMetric(label, value, value_type, group="secondary")
        for label in DETAIL_SECONDARY_KPI_LABELS
    ]


def _current_user() -> str:
    try:
        return getpass.getuser()
    except Exception:
        return ""
