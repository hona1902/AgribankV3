from __future__ import annotations

import logging
from pathlib import Path
from time import perf_counter

from PySide6.QtCore import Qt, QTimer
from PySide6.QtWidgets import QFileDialog, QGridLayout, QMessageBox, QSizePolicy, QVBoxLayout, QWidget

from agribank_v3.features.credit.summary.customer.async_query import AsyncQueryController
from agribank_v3.features.credit.summary.customer.cross_branch_detail_dialog import CrossBranchCustomerDetailDialog
from agribank_v3.features.credit.summary.customer.export_service import (
    CROSS_BRANCH_COLUMNS,
    export_cross_branch_customers,
    suggested_customer_export_name,
)
from agribank_v3.features.credit.summary.customer.filters import CUSTOMER_TYPE_FILTERS, CustomerFilters
from agribank_v3.features.credit.summary.customer.formatters import format_customer_type
from agribank_v3.features.credit.summary.customer.repository import CustomerRepository
from agribank_v3.features.credit.summary.customer.table_models import CustomerTableModel
from agribank_v3.features.credit.summary.customer.widgets import (
    CustomerTableView,
    KpiMetric,
    MetricGrid,
    Pager,
    QueryStateBanner,
    SearchBox,
    ScopeFilterComboBox,
    combo_box,
    configure_combo_popup_width,
    current_data,
    populate_combo,
    populate_officer_combo,
    secondary_button,
)


LOGGER = logging.getLogger(__name__)

CROSS_BRANCH_SCOPE_FILTERS = (
    ("Liên chi nhánh", "cross_branch"),
    ("Hội sở và PGD cùng chi nhánh", "head_and_pgd"),
    ("Nhiều PGD cùng chi nhánh", "multi_pgd"),
)

OFFICE_FILTER_MODE_FILTERS = (
    ("Theo đơn vị đại diện", "representative"),
    ("Theo nơi có dư nợ thực tế", "actual"),
)


class CrossBranchCustomersTab(QWidget):
    def __init__(self, repository: CustomerRepository, parent=None) -> None:
        super().__init__(parent)
        self.repository = repository
        self.page = 1
        self.page_size = 100
        self.sort_by = "branch_count"
        self.sort_desc = True
        self._updating_filters = False
        self._filter_timer = QTimer(self)
        self._filter_timer.setSingleShot(True)
        self._filter_timer.setInterval(350)
        self._filter_timer.timeout.connect(self._apply_filter_changed)
        self.query_controller = AsyncQueryController(self, max_cache_entries=32)
        self.last_benchmark: dict[str, object] = {}
        self.detail_windows: list[CrossBranchCustomerDetailDialog] = []
        layout = QVBoxLayout(self)
        toolbar = QWidget()
        toolbar.setObjectName("CustomerCrossBranchFilterPanel")
        toolbar.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        toolbar_layout = QGridLayout(toolbar)
        toolbar_layout.setContentsMargins(0, 0, 0, 0)
        toolbar_layout.setHorizontalSpacing(8)
        toolbar_layout.setVerticalSpacing(7)
        self.period_combo = combo_box("Kỳ báo cáo", minimum_width=118, maximum_width=160)
        self.branch_combo = combo_box("Tất cả chi nhánh", minimum_width=160, maximum_width=240)
        self.office_combo = combo_box("Tất cả đơn vị", minimum_width=170, maximum_width=260)
        self.scope_combo = ScopeFilterComboBox(CROSS_BRANCH_SCOPE_FILTERS)
        self.office_filter_mode_combo = combo_box("Cách xác định đơn vị", minimum_width=190, maximum_width=280)
        populate_combo(self.office_filter_mode_combo, OFFICE_FILTER_MODE_FILTERS)
        representative_index = self.office_filter_mode_combo.findData("representative")
        if representative_index >= 0:
            self.office_filter_mode_combo.setCurrentIndex(representative_index)
        self.customer_type_combo = combo_box("Tất cả loại khách hàng", minimum_width=170, maximum_width=240)
        self.minimum_branch_combo = combo_box("Từ 2 chi nhánh", minimum_width=150, maximum_width=210)
        self.minimum_branch_combo.clear()
        self.minimum_branch_combo.addItem("Từ 2 chi nhánh", 2)
        self.minimum_branch_combo.addItem("Từ 3 chi nhánh", 3)
        self.minimum_branch_combo.addItem("Từ 4 chi nhánh", 4)
        self.minimum_branch_combo.addItem("Tất cả trường hợp liên chi nhánh", 2)
        for index in range(self.minimum_branch_combo.count()):
            self.minimum_branch_combo.setItemData(index, self.minimum_branch_combo.itemText(index), Qt.ItemDataRole.ToolTipRole)
        configure_combo_popup_width(self.minimum_branch_combo, minimum_popup_width=300)
        self.officer_combo = combo_box(
            "Tất cả cán bộ",
            minimum_width=220,
            maximum_width=320,
            minimum_contents_length=18,
            searchable=True,
        )
        self.search_box = SearchBox("Tìm mã hoặc tên khách hàng")
        self.search_box.setMinimumWidth(220)
        self.search_box.setMaximumWidth(340)
        apply_button = secondary_button("Áp dụng")
        refresh_button = secondary_button("Làm mới")
        clear_button = secondary_button("Xóa lọc")
        self.export_button = secondary_button("Xuất Excel")
        apply_button.clicked.connect(lambda: self.refresh(use_cache=True))
        refresh_button.clicked.connect(lambda: self.refresh(use_cache=False))
        clear_button.clicked.connect(self.clear_filters)
        self.export_button.clicked.connect(self.export_excel)
        toolbar_layout.addWidget(self.period_combo, 0, 0)
        toolbar_layout.addWidget(self.branch_combo, 0, 1)
        toolbar_layout.addWidget(self.office_combo, 0, 2)
        toolbar_layout.addWidget(self.scope_combo, 0, 3)
        toolbar_layout.addWidget(self.office_filter_mode_combo, 1, 0)
        toolbar_layout.addWidget(self.customer_type_combo, 1, 1)
        toolbar_layout.addWidget(self.minimum_branch_combo, 1, 2)
        toolbar_layout.addWidget(self.officer_combo, 1, 3)
        toolbar_layout.addWidget(self.search_box, 2, 0, 1, 2)
        toolbar_layout.addWidget(apply_button, 2, 2)
        toolbar_layout.addWidget(refresh_button, 2, 3)
        toolbar_layout.addWidget(clear_button, 2, 4)
        toolbar_layout.addWidget(self.export_button, 2, 5)
        toolbar_layout.setColumnStretch(6, 1)
        layout.addWidget(toolbar)
        self.state_banner = QueryStateBanner()
        self.state_banner.retryRequested.connect(lambda: self.refresh(use_cache=False))
        layout.addWidget(self.state_banner)
        self.metrics = MetricGrid()
        layout.addWidget(self.metrics)
        self.model = CustomerTableModel(CROSS_BRANCH_COLUMNS, self)
        self.table = CustomerTableView()
        self.table.setModel(self.model)
        self.table.doubleClicked.connect(self._double_clicked)
        self.table.horizontalHeader().sectionClicked.connect(self._sort_changed)
        layout.addWidget(self.table, stretch=1)
        self.pager = Pager()
        self.pager.pageChanged.connect(self._page_changed)
        self.pager.pageSizeChanged.connect(self._page_size_changed)
        layout.addWidget(self.pager)
        self.table.apply_default_widths((
            50, 82, 130, 220, 130, 105, 100, 90, 90, 120,
            110, 180, 230, 130, 170, 240, 260, 130, 130, 130,
            130, 140, 130, 110, 110, 100, 100, 95, 220, 95,
            170,
        ))
        self.search_box.debouncedTextChanged.connect(lambda _text: self._filter_changed())
        self.scope_combo.applied.connect(self._filter_changed)
        for combo in (
            self.period_combo,
            self.branch_combo,
            self.office_combo,
            self.office_filter_mode_combo,
            self.customer_type_combo,
            self.minimum_branch_combo,
            self.officer_combo,
        ):
            combo.currentIndexChanged.connect(self._filter_changed)
        self.refresh_filter_options()

    def current_filters(self) -> CustomerFilters:
        return CustomerFilters(
            current_period=current_data(self.period_combo),
            branch_code=current_data(self.branch_combo),
            customer_type=current_data(self.customer_type_combo),
            officer=current_data(self.officer_combo),
            search_text=self.search_box.text().strip(),
        )

    def current_period(self) -> str:
        period = current_data(self.period_combo)
        if period:
            return period
        periods = self.repository.distinct_periods()
        return periods[-1] if periods else ""

    def minimum_branch_count(self) -> int:
        return max(2, int(current_data(self.minimum_branch_combo) or 2))

    def office_code(self) -> str:
        return current_data(self.office_combo)

    def scope_type(self) -> object:
        selected = self.scope_combo.selected_values()
        return selected if selected else ("__none__",)

    def office_filter_mode(self) -> str:
        return current_data(self.office_filter_mode_combo) or "representative"

    def refresh_filter_options(self) -> None:
        self._updating_filters = True
        current_period = current_data(self.period_combo)
        current_branch = current_data(self.branch_combo)
        periods = self.repository.distinct_periods()
        try:
            self.period_combo.blockSignals(True)
            if periods:
                populate_combo(self.period_combo, periods)
                target_period = current_period if current_period in periods else periods[-1]
                self.period_combo.setCurrentIndex(max(0, self.period_combo.findData(target_period)))
            else:
                self.period_combo.clear()
                self.period_combo.addItem("Chưa có dữ liệu", "")
                self.period_combo.setCurrentIndex(0)
            self.period_combo.setEnabled(bool(periods))
            self.period_combo.blockSignals(False)
            if periods:
                period_filter = CustomerFilters(current_period=self.current_period())
                branch_values = [
                    (self.repository.unit_directory.get_branch_display_name(code), code)
                    for code in self.repository.distinct_branch_codes(period_filter)
                ]
                populate_combo(self.branch_combo, branch_values)
                if current_branch and self.branch_combo.findData(current_branch) >= 0:
                    self.branch_combo.setCurrentIndex(self.branch_combo.findData(current_branch))
                populate_combo(
                    self.customer_type_combo,
                    [(format_customer_type(value), value) for value in self.repository.distinct_customer_types(period_filter)],
                )
                populate_officer_combo(self.officer_combo, self.repository.distinct_officers(period_filter))
                self._refresh_office_filter_options()
            else:
                populate_combo(self.branch_combo, [])
                populate_combo(self.office_combo, [])
                populate_combo(self.customer_type_combo, [])
                populate_officer_combo(self.officer_combo, [], first_label="Tất cả cán bộ")
        finally:
            self._updating_filters = False

    def _refresh_office_filter_options(self) -> None:
        current = current_data(self.office_combo)
        offices = self.repository.distinct_offices(self.current_period(), branch_code=current_data(self.branch_combo))
        populate_combo(
            self.office_combo,
            [(str(row.get("office_display") or row.get("office_code") or ""), str(row.get("office_code") or "")) for row in offices],
        )
        if current and self.office_combo.findData(current) >= 0:
            self.office_combo.setCurrentIndex(self.office_combo.findData(current))

    def refresh(self, *args, use_cache: bool = True) -> None:
        self._filter_timer.stop()
        period = self.current_period()
        if not period or not self.repository.has_period_data():
            self.set_empty_state()
            return
        if not self.repository.has_period(period):
            self.set_empty_state(f"Không có dữ liệu khách hàng cho kỳ {period}.")
            return
        if not self.repository.has_office_detail_for_period(period):
            self.set_empty_state(f"Kỳ {period} chưa có dữ liệu chi tiết Hội sở/Phòng giao dịch để phân tích liên chi nhánh.")
            return
        filters = self.current_filters()
        minimum = self.minimum_branch_count()
        office_code = self.office_code()
        scope_type = self.scope_type()
        office_filter_mode = self.office_filter_mode()
        page = self.page
        page_size = self.page_size
        sort_by = self.sort_by
        sort_desc = self.sort_desc
        cache_key = (
            "cross_branch",
            period,
            filters,
            minimum,
            scope_type,
            office_code,
            office_filter_mode,
            page,
            page_size,
            sort_by,
            sort_desc,
        )
        self.cancel_queries()
        self.query_controller.run(
            "cross_branch_customers",
            lambda: self._load_payload(period, filters, minimum, scope_type, office_code, office_filter_mode, page, page_size, sort_by, sort_desc),
            self._apply_payload,
            self._query_failed,
            cache_key=cache_key,
            use_cache=use_cache,
            state_callback=self._state_changed,
            log_context={
                "period": period,
                "page": page,
                "page_size": page_size,
                "scope": scope_type,
                "branch": filters.branch_code,
                "office": office_code,
                "search": bool(filters.search_text),
            },
        )

    def export_excel(self) -> None:
        period = self.current_period()
        if not period or not self.repository.has_period_data():
            self.state_banner.set_empty("Chưa có dữ liệu để xuất.")
            return
        path, _ = QFileDialog.getSaveFileName(
            self,
            "Xuất khách hàng vay liên chi nhánh",
            suggested_customer_export_name(f"KhachHangLienChiNhanh_{period or 'TatCa'}"),
            "Excel (*.xlsx)",
        )
        if not path:
            return
        try:
            output = export_cross_branch_customers(
                self.repository,
                period,
                self.current_filters(),
                self.minimum_branch_count(),
                Path(path),
                scope_type=self.scope_type(),
                office_code=self.office_code(),
                office_filter_mode=self.office_filter_mode(),
                sort_by=self.sort_by,
                sort_desc=self.sort_desc,
            )
        except Exception as exc:
            QMessageBox.warning(self, "Xuất khách hàng vay liên chi nhánh", str(exc))
            return
        QMessageBox.information(self, "Xuất khách hàng vay liên chi nhánh", f"Đã xuất: {output}")

    def clear_filters(self) -> None:
        self._filter_timer.stop()
        self._updating_filters = True
        try:
            self.branch_combo.setCurrentIndex(0)
            self.office_combo.setCurrentIndex(0)
            self.scope_combo.select_all_scopes(emit=False)
            representative_index = self.office_filter_mode_combo.findData("representative")
            self.office_filter_mode_combo.setCurrentIndex(representative_index if representative_index >= 0 else 0)
            self.customer_type_combo.setCurrentIndex(0)
            self.minimum_branch_combo.setCurrentIndex(0)
            self.officer_combo.setCurrentIndex(0)
            self.search_box.clear()
        finally:
            self._updating_filters = False
        self.page = 1
        self.refresh()

    def invalidate_cache(self) -> None:
        self.query_controller.invalidate_cache()

    def cancel_queries(self) -> None:
        self.query_controller.cancel_pending()

    def wait_for_queries(self) -> None:
        self.query_controller.wait_for_idle()

    def set_empty_state(self, message: str = "Chưa có dữ liệu khách hàng.") -> None:
        self.cancel_queries()
        self._filter_timer.stop()
        self.model.set_rows([])
        self.pager.set_state(page=1, page_size=self.page_size, total_rows=0)
        self.metrics.set_empty(
            (
                "Số khách hàng vay liên chi nhánh",
                "Tổng dư nợ liên chi nhánh",
                "Tổng số lượt chi nhánh tham gia",
                "Tổng số lượt đơn vị tham gia",
                "Khách hàng Hội sở + PGD",
                "Khách hàng nhiều PGD",
                "Dư nợ Hội sở",
                "Dư nợ PGD",
                "Số khách hàng vay tại từ 3 chi nhánh",
                "Khách hàng có nhiều cán bộ quản lý",
                "Khách hàng có override cán bộ",
                "Chi nhánh có nhiều khách hàng liên chi nhánh nhất",
            )
        )
        self.state_banner.set_empty(message)
        self.export_button.setEnabled(False)
        self.export_button.setToolTip("Chưa có dữ liệu để xuất")

    def _load_payload(
        self,
        period: str,
        filters: CustomerFilters,
        minimum: int,
        scope_type: object,
        office_code: str,
        office_filter_mode: str,
        page: int,
        page_size: int,
        sort_by: str,
        sort_desc: bool,
    ) -> dict[str, object]:
        return self.repository.query_cross_branch_tab_payload(
            period,
            filters,
            minimum_branch_count=minimum,
            scope_type=scope_type,
            office_code=office_code,
            office_filter_mode=office_filter_mode,
            page=page,
            page_size=page_size,
            sort_by=sort_by,
            sort_desc=sort_desc,
        )

    def _apply_payload(self, payload: dict[str, object]) -> None:
        self.export_button.setEnabled(True)
        self.export_button.setToolTip("")
        result = payload.get("result")
        kpis = dict(payload.get("kpis") or {})
        benchmark = dict(payload.get("benchmark") or {})
        if result is None:
            self._query_failed(RuntimeError("empty result"))
            return
        ui_started = perf_counter()
        self.model.set_rows(result.rows)
        self.pager.set_state(page=result.page, page_size=result.page_size, total_rows=result.total_rows)
        self.metrics.set_metrics(
            [
                KpiMetric("Số khách hàng vay liên chi nhánh", kpis.get("cross_customer_count", 0), "count"),
                KpiMetric("Tổng dư nợ liên chi nhánh", kpis.get("total_balance", 0), "money"),
                KpiMetric("Tổng số lượt chi nhánh tham gia", kpis.get("branch_occurrence_count", 0), "count"),
                KpiMetric("Tổng số lượt đơn vị tham gia", kpis.get("office_occurrence_count", 0), "count"),
                KpiMetric("Khách hàng Hội sở + PGD", kpis.get("head_and_pgd_customer_count", 0), "count"),
                KpiMetric("Khách hàng nhiều PGD", kpis.get("multi_pgd_customer_count", 0), "count"),
                KpiMetric("Dư nợ Hội sở", kpis.get("head_office_balance", 0), "money"),
                KpiMetric("Dư nợ PGD", kpis.get("pgd_balance", 0), "money"),
                KpiMetric("Số khách hàng vay tại từ 3 chi nhánh", kpis.get("three_branch_customer_count", 0), "count"),
                KpiMetric("Khách hàng có nhiều cán bộ quản lý", kpis.get("multiple_officer_customer_count", 0), "count"),
                KpiMetric("Khách hàng có override cán bộ", kpis.get("override_customer_count", 0), "count"),
                KpiMetric("Chi nhánh có nhiều khách hàng liên chi nhánh nhất", kpis.get("top_branch_name", ""), "text"),
            ]
        )
        benchmark["ui_update"] = {
            "elapsed_ms": (perf_counter() - ui_started) * 1000,
            "sql_count": 0,
        }
        self.last_benchmark = {
            "stages": benchmark,
            "sql_statement_count": int(payload.get("sql_statement_count") or 0),
            "row_count": int(result.total_rows or 0),
        }
        LOGGER.info(
            "Cross-branch tab loaded: rows=%s sql=%s benchmark=%s",
            result.total_rows,
            self.last_benchmark["sql_statement_count"],
            {
                key: round(float(value.get("elapsed_ms", 0)), 1)
                for key, value in benchmark.items()
                if isinstance(value, dict)
            },
        )
        if result.total_rows:
            self.state_banner.clear()
        else:
            self.state_banner.set_empty(
                str(payload.get("empty_message") or "Không có khách hàng vay liên chi nhánh phù hợp với bộ lọc.")
            )

    def _query_failed(self, exc: Exception) -> None:
        self.model.set_rows([])
        self.pager.set_state(page=1, page_size=self.page_size, total_rows=0)
        self.metrics.set_empty(
            (
                "Số khách hàng vay liên chi nhánh",
                "Tổng dư nợ liên chi nhánh",
                "Tổng số lượt chi nhánh tham gia",
                "Tổng số lượt đơn vị tham gia",
                "Khách hàng Hội sở + PGD",
                "Khách hàng nhiều PGD",
                "Dư nợ Hội sở",
                "Dư nợ PGD",
                "Số khách hàng vay tại từ 3 chi nhánh",
                "Khách hàng có nhiều cán bộ quản lý",
                "Khách hàng có override cán bộ",
                "Chi nhánh có nhiều khách hàng liên chi nhánh nhất",
            )
        )

    def _state_changed(self, state: str, message: str) -> None:
        if state == "loading":
            self.state_banner.set_loading()
        elif state == "error":
            self.state_banner.set_error(message or "Không tải được khách hàng vay liên chi nhánh.")

    def _filter_changed(self) -> None:
        if self._updating_filters:
            return
        sender = self.sender()
        if sender is self.period_combo:
            self.refresh_filter_options()
        elif sender is self.branch_combo:
            self._refresh_office_filter_options()
        self.page = 1
        self._filter_timer.start()

    def _apply_filter_changed(self) -> None:
        self.refresh()

    def _page_changed(self, page: int) -> None:
        self.page = max(1, int(page or 1))
        self.refresh()

    def _page_size_changed(self, page_size: int) -> None:
        self.page = 1
        self.page_size = int(page_size or 100)
        self.refresh()

    def _sort_changed(self, section: int) -> None:
        if 0 <= section < len(CROSS_BRANCH_COLUMNS):
            field = CROSS_BRANCH_COLUMNS[section][0]
            if field not in {"branch_count", "office_count", "pgd_count", "total_balance", "nim_after", "medium_long_ratio", "customer_name"}:
                return
            if self.sort_by == field:
                self.sort_desc = not self.sort_desc
            else:
                self.sort_by = field
                self.sort_desc = field != "customer_name"
            self.page = 1
            self.refresh()

    def _double_clicked(self, index) -> None:
        row = self.model.raw_row(index.row())
        if not row:
            return
        dialog = CrossBranchCustomerDetailDialog(
            self.repository,
            str(row.get("period") or self.current_period()),
            str(row.get("customer_sequence") or ""),
            parent=self,
        )
        dialog.setAttribute(Qt.WidgetAttribute.WA_DeleteOnClose, True)
        self.detail_windows.append(dialog)
        dialog.finished.connect(lambda _result, item=dialog: self._forget_detail(item))
        dialog.show()
        dialog.raise_()

    def _forget_detail(self, dialog: CrossBranchCustomerDetailDialog) -> None:
        if dialog in self.detail_windows:
            self.detail_windows.remove(dialog)
