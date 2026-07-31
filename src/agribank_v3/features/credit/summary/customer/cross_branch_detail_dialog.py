from __future__ import annotations

from PySide6.QtCore import Qt
from pathlib import Path

from PySide6.QtWidgets import QFileDialog, QDialog, QHBoxLayout, QTabWidget, QVBoxLayout, QWidget

from agribank_v3.features.credit.summary.customer.async_query import AsyncQueryController
from agribank_v3.features.credit.summary.customer.customer_detail_window import CustomerDetailWindow
from agribank_v3.features.credit.summary.customer.export_service import (
    CROSS_BRANCH_DETAIL_COLUMNS,
    CROSS_BRANCH_HISTORY_COLUMNS,
    export_cross_branch_customer_detail,
    suggested_customer_export_name,
)
from agribank_v3.features.credit.summary.customer.formatters import format_customer_type
from agribank_v3.features.credit.summary.customer.repository import CustomerRepository
from agribank_v3.features.credit.summary.customer.table_models import CustomerTableModel
from agribank_v3.features.credit.summary.customer.widgets import (
    CustomerTableView,
    KpiMetric,
    MetricGrid,
    QueryStateBanner,
    combo_box,
    current_data,
    fit_window_to_screen,
    populate_combo,
    secondary_button,
)


DETAIL_SCOPE_FILTERS = (
    ("Tất cả", "all"),
    ("Vay tại nhiều chi nhánh", "cross_branch"),
    ("Hội sở và PGD cùng chi nhánh", "head_and_pgd"),
    ("Nhiều PGD cùng chi nhánh", "multi_pgd"),
    ("Chỉ Hội sở", "only_head_office"),
    ("Chỉ một PGD", "only_one_pgd"),
)


class CrossBranchCustomerDetailDialog(QDialog):
    def __init__(
        self,
        repository: CustomerRepository,
        period: str,
        customer_sequence: str,
        parent=None,
    ) -> None:
        super().__init__(parent)
        self.repository = repository
        self.period = str(period or "").strip()
        self.customer_sequence = str(customer_sequence or "").strip()
        self.query_controller = AsyncQueryController(self, max_cache_entries=12)
        self.detail_windows: list[CustomerDetailWindow] = []
        self._updating_filters = False
        self.setWindowTitle("Phân tích khách hàng vay liên chi nhánh - AgribankV3")
        self.setModal(False)
        self.setWindowFlag(Qt.WindowType.WindowMinMaxButtonsHint, True)
        fit_window_to_screen(
            self,
            width_ratio=0.84,
            height_ratio=0.82,
            max_width=1280,
            max_height=820,
            min_width=900,
            min_height=600,
        )
        layout = QVBoxLayout(self)
        toolbar = QWidget()
        toolbar_layout = QHBoxLayout(toolbar)
        toolbar_layout.setContentsMargins(0, 0, 0, 0)
        self.period_from_combo = combo_box("Từ kỳ", minimum_width=110, maximum_width=150)
        self.period_to_combo = combo_box("Đến kỳ", minimum_width=110, maximum_width=150)
        self.report_period_combo = combo_box("Kỳ báo cáo", minimum_width=118, maximum_width=160)
        self.branch_combo = combo_box("Tất cả chi nhánh", minimum_width=170, maximum_width=260)
        self.office_combo = combo_box("Tất cả đơn vị", minimum_width=180, maximum_width=300)
        self.scope_combo = combo_box("Loại phạm vi vay", minimum_width=210, maximum_width=310)
        populate_combo(self.scope_combo, DETAIL_SCOPE_FILTERS)
        apply_button = secondary_button("Áp dụng")
        clear_button = secondary_button("Xóa lọc")
        refresh_button = secondary_button("Làm mới")
        export_button = secondary_button("Xuất Excel")
        apply_button.clicked.connect(lambda: self.refresh(use_cache=False))
        clear_button.clicked.connect(self.clear_filters)
        refresh_button.clicked.connect(lambda: self.refresh(use_cache=False))
        export_button.clicked.connect(self.export_excel)
        for widget in (
            self.period_from_combo,
            self.period_to_combo,
            self.report_period_combo,
            self.branch_combo,
            self.office_combo,
            self.scope_combo,
            apply_button,
            clear_button,
            refresh_button,
            export_button,
        ):
            toolbar_layout.addWidget(widget)
        toolbar_layout.addStretch()
        layout.addWidget(toolbar)
        self.state_banner = QueryStateBanner()
        self.state_banner.retryRequested.connect(lambda: self.refresh(use_cache=False))
        layout.addWidget(self.state_banner)
        self.metrics = MetricGrid()
        layout.addWidget(self.metrics)
        self.tabs = QTabWidget()
        layout.addWidget(self.tabs, stretch=1)
        self.detail_model = CustomerTableModel(CROSS_BRANCH_DETAIL_COLUMNS, self)
        self.history_model = CustomerTableModel(CROSS_BRANCH_HISTORY_COLUMNS, self)
        self.detail_table = self._add_table_tab("Chi tiết theo đơn vị", self.detail_model)
        self.detail_table.doubleClicked.connect(self._branch_row_double_clicked)
        self.history_table = self._add_table_tab("Lịch sử theo kỳ", self.history_model)
        self.detail_table.apply_default_widths((80, 150, 90, 180, 120, 120, 160, 220, 120, 150, 150, 130, 130, 140, 110, 110, 100, 100, 90, 120))
        self.history_table.apply_default_widths((90, 110, 110, 90, 90, 140, 130, 220, 260, 130, 130, 130, 130, 110, 100, 100))
        for combo in (
            self.period_from_combo,
            self.period_to_combo,
            self.report_period_combo,
            self.branch_combo,
            self.office_combo,
            self.scope_combo,
        ):
            combo.currentIndexChanged.connect(self._filter_changed)
        self._populate_filter_options()
        self.refresh()

    def refresh(self, *args, use_cache: bool = True) -> None:
        report_period = current_data(self.report_period_combo) or self.period
        period_from = current_data(self.period_from_combo)
        period_to = current_data(self.period_to_combo)
        branch_code = current_data(self.branch_combo)
        office_code = current_data(self.office_combo)
        scope_type = current_data(self.scope_combo) or "all"
        cache_key = (
            "cross_branch_detail",
            self.customer_sequence,
            period_from,
            period_to,
            report_period,
            branch_code,
            office_code,
            scope_type,
        )
        self.query_controller.run(
            "cross_branch_detail",
            lambda: {
                "detail": self.repository.get_cross_branch_customer_offices(
                    self.customer_sequence,
                    report_period,
                    branch_code=branch_code,
                    office_code=office_code,
                    scope_type=scope_type,
                ),
                "history": self.repository.get_cross_branch_customer_unit_history(
                    self.customer_sequence,
                    period_from,
                    period_to,
                    branch_code=branch_code,
                    office_code=office_code,
                    scope_type=scope_type,
                ),
                "kpis": self.repository.get_cross_branch_customer_filtered_kpis(
                    self.customer_sequence,
                    report_period,
                    branch_code=branch_code,
                    office_code=office_code,
                    scope_type=scope_type,
                ),
            },
            self._apply_payload,
            self._query_failed,
            cache_key=cache_key,
            use_cache=use_cache,
            state_callback=self._state_changed,
        )

    def invalidate_cache(self) -> None:
        self.query_controller.invalidate_cache()

    def wait_for_queries(self) -> None:
        self.query_controller.wait_for_idle()

    def clear_filters(self) -> None:
        self._populate_filter_options(reset=True)
        self.refresh(use_cache=False)

    def export_excel(self) -> None:
        report_period = current_data(self.report_period_combo) or self.period
        path, _ = QFileDialog.getSaveFileName(
            self,
            "Xuất phân tích khách hàng vay liên chi nhánh",
            suggested_customer_export_name(f"PhanTichLienChiNhanh_{self.customer_sequence}_{report_period or 'TatCa'}"),
            "Excel (*.xlsx)",
        )
        if not path:
            return
        output = export_cross_branch_customer_detail(
            self.repository,
            self.customer_sequence,
            Path(path),
            period_from=current_data(self.period_from_combo),
            period_to=current_data(self.period_to_combo),
            report_period=report_period,
            branch_code=current_data(self.branch_combo),
            office_code=current_data(self.office_combo),
            scope_type=current_data(self.scope_combo) or "all",
        )
        self.state_banner.set_empty(f"Đã xuất: {output}")

    def closeEvent(self, event) -> None:
        self.query_controller.cancel_pending()
        super().closeEvent(event)

    def _apply_payload(self, payload: dict[str, object]) -> None:
        detail = list(payload.get("detail") or [])
        history = list(payload.get("history") or [])
        self.detail_model.set_rows(detail)
        self.history_model.set_rows(history)
        self.metrics.set_metrics(_detail_metrics(dict(payload.get("kpis") or {}), self.customer_sequence))
        kpis = dict(payload.get("kpis") or {})
        if detail:
            self.state_banner.clear()
        elif int(kpis.get("office_detail_missing") or 0):
            self.state_banner.set_empty(
                "Chưa có dữ liệu chi tiết Hội sở/Phòng giao dịch cho kỳ này. "
                "Vui lòng nhập lại kỳ NIM Dư nợ từ file FTP Loan."
            )
        else:
            self.state_banner.set_empty("Không có dữ liệu chi tiết liên chi nhánh phù hợp.")

    def _query_failed(self, exc: Exception) -> None:
        self.detail_model.set_rows([])
        self.history_model.set_rows([])
        self.metrics.set_metrics(_placeholder_metrics(self.period, self.customer_sequence))

    def _state_changed(self, state: str, message: str) -> None:
        if state == "loading":
            self.state_banner.set_loading()
            self.metrics.set_metrics(_placeholder_metrics("…", "…"))
        elif state == "error":
            self.state_banner.set_error(message or "Không tải được chi tiết khách hàng liên chi nhánh.")

    def _add_table_tab(self, label: str, model: CustomerTableModel) -> CustomerTableView:
        tab = QWidget()
        layout = QVBoxLayout(tab)
        table = CustomerTableView()
        table.setModel(model)
        table.setSortingEnabled(True)
        layout.addWidget(table, stretch=1)
        self.tabs.addTab(tab, label)
        return table

    def _branch_row_double_clicked(self, index) -> None:
        row = self.detail_model.raw_row(index.row())
        customer_code = str(row.get("customer_code") or "").strip()
        period = str(row.get("period") or self.period).strip()
        if not customer_code:
            return
        dialog = CustomerDetailWindow(self.repository, customer_code, period=period, parent=self)
        dialog.setAttribute(Qt.WidgetAttribute.WA_DeleteOnClose, True)
        self.detail_windows.append(dialog)
        dialog.finished.connect(lambda _result, item=dialog: self._forget_detail(item))
        dialog.show()
        dialog.raise_()

    def _forget_detail(self, dialog: CustomerDetailWindow) -> None:
        if dialog in self.detail_windows:
            self.detail_windows.remove(dialog)

    def _populate_filter_options(self, *, reset: bool = False) -> None:
        self._updating_filters = True
        try:
            periods = self.repository.customer_sequence_periods(self.customer_sequence)
            populate_combo(self.period_from_combo, periods)
            populate_combo(self.period_to_combo, periods)
            populate_combo(self.report_period_combo, periods)
            if periods:
                from_period = periods[0] if reset or current_data(self.period_from_combo) not in periods else current_data(self.period_from_combo)
                to_period = periods[-1] if reset or current_data(self.period_to_combo) not in periods else current_data(self.period_to_combo)
                report_period = self.period if self.period in periods and not reset else periods[-1]
                if not reset and current_data(self.report_period_combo) in periods:
                    report_period = current_data(self.report_period_combo)
                self._set_combo_current_data(self.period_from_combo, from_period)
                self._set_combo_current_data(self.period_to_combo, to_period)
                self._set_combo_current_data(self.report_period_combo, report_period)
            branches = self.repository.get_customer_available_branches(self.customer_sequence, current_data(self.report_period_combo) or self.period)
            populate_combo(
                self.branch_combo,
                [(str(row.get("branch_name") or row.get("branch_code") or ""), str(row.get("branch_code") or "")) for row in branches],
            )
            if reset:
                self.branch_combo.setCurrentIndex(0)
            self._refresh_office_options()
            if reset:
                self.office_combo.setCurrentIndex(0)
                self.scope_combo.setCurrentIndex(0)
        finally:
            self._updating_filters = False

    def _refresh_office_options(self) -> None:
        current = current_data(self.office_combo)
        offices = self.repository.get_customer_available_offices(
            self.customer_sequence,
            current_data(self.report_period_combo) or self.period,
            branch_code=current_data(self.branch_combo),
        )
        populate_combo(
            self.office_combo,
            [(str(row.get("office_display") or row.get("office_code") or ""), str(row.get("office_code") or "")) for row in offices],
        )
        if current and self.office_combo.findData(current) >= 0:
            self.office_combo.setCurrentIndex(self.office_combo.findData(current))

    def _filter_changed(self) -> None:
        if self._updating_filters:
            return
        sender = self.sender()
        if sender is self.report_period_combo:
            self._populate_filter_options()
        elif sender is self.branch_combo:
            self._refresh_office_options()
        self.refresh()

    @staticmethod
    def _set_combo_current_data(combo, value: str) -> None:
        index = combo.findData(value)
        if index >= 0:
            combo.setCurrentIndex(index)


def _detail_metrics(kpis: dict[str, object], customer_sequence: str) -> list[KpiMetric]:
    return [
        KpiMetric("Kỳ báo cáo", kpis.get("period"), "text"),
        KpiMetric("Mã khách hàng gốc", customer_sequence, "text"),
        KpiMetric("Tên khách hàng", kpis.get("customer_name", ""), "text"),
        KpiMetric("Loại khách hàng", kpis.get("customer_type_display", ""), "text"),
        KpiMetric("Số chi nhánh vay", kpis.get("branch_count", 0), "count"),
        KpiMetric("Số đơn vị vay", kpis.get("office_count", 0), "count"),
        KpiMetric("Số Hội sở có dư nợ", kpis.get("head_office_count", 0), "count"),
        KpiMetric("Số PGD có dư nợ", kpis.get("pgd_count", 0), "count"),
        KpiMetric("Tổng dư nợ", kpis.get("total_balance", 0), "money"),
        KpiMetric("Dư nợ tại Hội sở", kpis.get("head_office_balance", 0), "money"),
        KpiMetric("Dư nợ tại PGD", kpis.get("pgd_balance", 0), "money"),
        KpiMetric("Lãi suất bình quân", kpis.get("average_rate", 0), "percentage"),
        KpiMetric("NIM trước ĐC", kpis.get("nim_before", 0), "percentage"),
        KpiMetric("NIM sau ĐC", kpis.get("nim_after", 0), "percentage"),
    ]


def _placeholder_metrics(period: object, customer_sequence: object) -> list[KpiMetric]:
    return [
        KpiMetric("Kỳ", period, "text"),
        KpiMetric("Mã khách hàng gốc", customer_sequence, "text"),
        KpiMetric("Tên khách hàng", None, "text"),
        KpiMetric("Loại khách hàng", None, "text"),
        KpiMetric("Tổng số chi nhánh", None, "text"),
        KpiMetric("Tổng dư nợ", None, "money"),
        KpiMetric("Lãi suất bình quân", None, "percentage"),
        KpiMetric("NIM trước ĐC", None, "percentage"),
        KpiMetric("NIM sau ĐC", None, "percentage"),
    ]


def _weighted(rows: list[dict[str, object]], field: str, denominator: float) -> float:
    if denominator == 0:
        return 0.0
    return sum(_number(row.get(field)) * _number(row.get("total_balance")) for row in rows) / denominator


def _number(value: object) -> float:
    try:
        return float(value or 0)
    except (TypeError, ValueError):
        return 0.0
