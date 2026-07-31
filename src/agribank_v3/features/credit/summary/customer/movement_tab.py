from __future__ import annotations

from dataclasses import replace
import logging
from pathlib import Path

from PySide6.QtWidgets import QFileDialog, QMessageBox, QVBoxLayout, QWidget

from agribank_v3.features.credit.summary.customer.async_query import AsyncQueryController
from agribank_v3.features.credit.summary.customer.export_service import (
    MOVEMENT_COLUMNS,
    export_customer_growth,
    suggested_customer_export_name,
)
from agribank_v3.features.credit.summary.customer.filters import MOVEMENT_STATUS_FILTERS, CustomerFilters
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


LOGGER = logging.getLogger(__name__)
MOVEMENT_KPI_LABELS = (
    "Số khách hàng vay mới",
    "Số khách hàng tất toán",
    "Số khách hàng tăng",
    "Số khách hàng giảm",
    "Số khách hàng không thay đổi",
    "Tổng dư nợ tăng",
    "Tổng dư nợ giảm",
    "Chênh lệch ròng",
)


class CustomerMovementTab(QWidget):
    def __init__(self, repository: CustomerRepository, filters_provider, parent=None) -> None:
        super().__init__(parent)
        self.repository = repository
        self.filters_provider = filters_provider
        self.page = 1
        self.page_size = 100
        self.sort_by = "difference"
        self.sort_desc = True
        self.query_controller = AsyncQueryController(self, max_cache_entries=32)
        self.export_controller = AsyncQueryController(self, max_cache_entries=1)
        self._updating_filters = False
        self._closing = False
        layout = QVBoxLayout(self)
        layout.setSpacing(6)
        filters = CompactToolbar()
        self.previous_combo = combo_box("Kỳ trước", minimum_width=110, maximum_width=150)
        self.current_combo = combo_box("Kỳ hiện tại", minimum_width=110, maximum_width=150)
        self.status_combo = combo_box("Tất cả phân loại", minimum_width=150, maximum_width=220)
        populate_combo(self.status_combo, MOVEMENT_STATUS_FILTERS[1:])
        for combo in (self.previous_combo, self.current_combo, self.status_combo):
            combo.currentIndexChanged.connect(self._filter_changed)
        refresh_button = secondary_button("Làm mới")
        refresh_button.clicked.connect(lambda: self.refresh(use_cache=False))
        self.export_button = secondary_button("Xuất Excel")
        self.export_button.clicked.connect(self.export_excel)
        self.open_nim_button = secondary_button("Mở NIM dư nợ")
        self.open_nim_button.clicked.connect(self._open_nim_dn)
        self.open_nim_button.hide()
        for combo in (self.previous_combo, self.current_combo, self.status_combo):
            filters.addWidget(combo)
        filters.addWidget(refresh_button)
        filters.addWidget(self.export_button)
        filters.addWidget(self.open_nim_button)
        layout.addWidget(filters)
        self.state_banner = QueryStateBanner()
        self.state_banner.retryRequested.connect(lambda: self.refresh(use_cache=False))
        layout.addWidget(self.state_banner)
        self.metrics = MetricGrid()
        layout.addWidget(self.metrics)
        self.model = CustomerTableModel(MOVEMENT_COLUMNS, self)
        self.table = CustomerTableView()
        self.table.setModel(self.model)
        self.table.horizontalHeader().sectionClicked.connect(self._sort_changed)
        layout.addWidget(self.table, stretch=1)
        self.pager = Pager()
        self.pager.pageChanged.connect(self._page_changed)
        self.pager.pageSizeChanged.connect(self._page_size_changed)
        layout.addWidget(self.pager)
        self.table.apply_default_widths((120, 220, 100, 90, 180, 130, 130, 130, 100, 140))

    def refresh_periods(self) -> None:
        self._updating_filters = True
        try:
            periods = self.repository.distinct_periods()
            if periods:
                self.previous_combo.setEnabled(True)
                self.current_combo.setEnabled(True)
                populate_combo(self.previous_combo, periods)
                populate_combo(self.current_combo, periods)
                if not current_data(self.current_combo):
                    self.current_combo.setCurrentIndex(self.current_combo.findData(periods[-1]))
                if len(periods) > 1 and not current_data(self.previous_combo):
                    self.previous_combo.setCurrentIndex(self.previous_combo.findData(periods[-2]))
            else:
                for combo in (self.previous_combo, self.current_combo):
                    combo.blockSignals(True)
                    try:
                        combo.clear()
                        combo.addItem("Chưa có dữ liệu", "")
                        combo.setCurrentIndex(0)
                        combo.setEnabled(False)
                    finally:
                        combo.blockSignals(False)
            self.open_nim_button.setVisible(len(periods) < 2)
        finally:
            self._updating_filters = False

    def refresh(self, *args, use_cache: bool = True) -> None:
        if not self.repository.has_period_data():
            self.set_empty_state()
            return
        previous_period = current_data(self.previous_combo)
        current_period = current_data(self.current_combo)
        filters = replace(self.filters_provider(), movement_status=current_data(self.status_combo))
        if not previous_period or not current_period:
            periods = self.repository.distinct_periods()
            message = (
                "Cần tối thiểu hai kỳ dữ liệu khách hàng để thực hiện so sánh."
                if len(periods) < 2
                else "Chưa chọn đủ kỳ so sánh."
            )
            self.model.set_rows([])
            self.metrics.set_empty(MOVEMENT_KPI_LABELS)
            self.state_banner.set_empty(message)
            self.pager.set_state(page=1, page_size=self.page_size, total_rows=0)
            self.export_button.setEnabled(False)
            self.export_button.setToolTip("Chưa đủ dữ liệu để xuất")
            return
        page = self.page
        page_size = self.page_size
        sort_by = self.sort_by
        sort_desc = self.sort_desc
        cache_key = ("movement", previous_period, current_period, filters, page, page_size, sort_by, sort_desc)
        generation = self.query_controller.generation + 1
        LOGGER.info(
            "movement_refresh_requested source=refresh generation=%s previous_period=%s current_period=%s page=%s page_size=%s",
            generation,
            previous_period,
            current_period,
            page,
            page_size,
        )
        self.query_controller.run(
            "customer_movement",
            lambda: self._load_payload(
                previous_period,
                current_period,
                filters,
                page,
                page_size,
                sort_by,
                sort_desc,
                generation=generation,
            ),
            self._apply_payload,
            self._query_failed,
            cache_key=cache_key,
            use_cache=use_cache,
            state_callback=self._state_changed,
            log_context={
                "previous_period": previous_period,
                "current_period": current_period,
                "page": page,
                "page_size": page_size,
            },
        )

    def _load_payload(
        self,
        previous_period: str,
        current_period: str,
        filters,
        page: int,
        page_size: int,
        sort_by: str,
        sort_desc: bool,
        generation: int | None = None,
    ) -> dict[str, object]:
        return self.repository.movement_payload(
            previous_period,
            current_period,
            filters,
            page=page,
            page_size=page_size,
            sort_by=sort_by,
            sort_desc=sort_desc,
            generation=generation,
        )

    def _apply_payload(self, payload: dict[str, object]) -> None:
        if self._closing:
            return
        self.export_button.setEnabled(True)
        self.export_button.setToolTip("")
        kpis = dict(payload.get("kpis") or {})
        self.metrics.set_metrics(
            [
                KpiMetric("Số khách hàng vay mới", kpis.get("new_customer_count", 0), "count"),
                KpiMetric("Số khách hàng tất toán", kpis.get("paid_off_customer_count", 0), "count"),
                KpiMetric("Số khách hàng tăng", kpis.get("increased_customer_count", 0), "count"),
                KpiMetric("Số khách hàng giảm", kpis.get("decreased_customer_count", 0), "count"),
                KpiMetric("Số khách hàng không thay đổi", kpis.get("unchanged_customer_count", 0), "count"),
                KpiMetric("Tổng dư nợ tăng", kpis.get("total_increase", 0), "money"),
                KpiMetric("Tổng dư nợ giảm", kpis.get("total_decrease", 0), "money"),
                KpiMetric("Chênh lệch ròng", kpis.get("net_difference", 0), "money", signed=True),
            ]
        )
        result = payload["result"]
        self.model.set_rows(result.rows)
        self.pager.set_state(page=result.page, page_size=result.page_size, total_rows=result.total_rows)
        if result.total_rows:
            self.state_banner.clear()
        else:
            self.state_banner.set_empty()

    def export_excel(self) -> None:
        previous_period = current_data(self.previous_combo)
        current_period = current_data(self.current_combo)
        if not previous_period or not current_period:
            QMessageBox.warning(self, "Xuất biến động dư nợ", "Chưa chọn đủ kỳ so sánh.")
            return
        path, _ = QFileDialog.getSaveFileName(
            self,
            "Xuất biến động dư nợ",
            suggested_customer_export_name("BienDongDuNo"),
            "Excel (*.xlsx)",
        )
        if not path:
            return
        filters = replace(self.filters_provider(), movement_status=current_data(self.status_combo))
        self.export_button.setEnabled(False)
        self.export_button.setToolTip("Đang xuất Excel")
        export_generation = self.export_controller.generation + 1
        self.export_controller.run(
            "customer_movement_export",
            lambda: export_customer_growth(
                self.repository,
                previous_period,
                current_period,
                filters,
                Path(path),
                sort_by=self.sort_by,
                sort_desc=self.sort_desc,
            ),
            self._export_finished,
            self._export_failed,
            cache_key=("movement_export", export_generation, previous_period, current_period, filters, path, self.sort_by, self.sort_desc),
            use_cache=False,
            state_callback=self._export_state_changed,
            log_context={"previous_period": previous_period, "current_period": current_period},
        )

    def _export_finished(self, output: Path) -> None:
        if self._closing:
            return
        self.export_button.setEnabled(True)
        self.export_button.setToolTip("")
        if self.state_banner.state == "loading":
            self.state_banner.clear()
        QMessageBox.information(self, "Xuất biến động dư nợ", f"Đã xuất: {output}")

    def _export_failed(self, exc: Exception) -> None:
        if self._closing:
            return
        self.export_button.setEnabled(True)
        self.export_button.setToolTip("")
        QMessageBox.warning(self, "Xuất biến động dư nợ", str(exc))

    def _filter_changed(self) -> None:
        if self._updating_filters:
            return
        self.page = 1
        self.refresh()

    def _open_nim_dn(self) -> None:
        parent = self.parent()
        while parent is not None:
            if hasattr(parent, "openNimDnRequested"):
                parent.openNimDnRequested.emit()
                return
            parent = parent.parent()

    def _page_changed(self, page: int) -> None:
        self.page = max(1, int(page or 1))
        self.refresh()

    def _page_size_changed(self, page_size: int) -> None:
        self.page = 1
        self.page_size = int(page_size or 100)
        self.refresh()

    def _sort_changed(self, section: int) -> None:
        if 0 <= section < len(MOVEMENT_COLUMNS):
            field = MOVEMENT_COLUMNS[section][0]
            if self.sort_by == field:
                self.sort_desc = not self.sort_desc
            else:
                self.sort_by = field
                self.sort_desc = field in {"difference", "current_balance"}
            self.page = 1
            self.refresh()

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
        self.model.set_rows([])
        self.metrics.set_empty(MOVEMENT_KPI_LABELS)
        self.pager.set_state(page=1, page_size=self.page_size, total_rows=0)
        self.state_banner.set_empty(message)
        self.export_button.setEnabled(False)
        self.export_button.setToolTip("Chưa có dữ liệu để xuất")
        self.open_nim_button.show()

    def _query_failed(self, exc: Exception) -> None:
        if self._closing:
            return
        self.model.set_rows([])
        self.pager.set_state(page=1, page_size=self.page_size, total_rows=0)
        self.metrics.set_empty(MOVEMENT_KPI_LABELS)

    def _state_changed(self, state: str, message: str) -> None:
        if self._closing:
            return
        if state == "loading":
            self.state_banner.set_loading("Đang tải dữ liệu biến động khách hàng...")
            self.metrics.set_loading(MOVEMENT_KPI_LABELS)
        elif state == "error":
            self.state_banner.set_error(message or "Không thể tải dữ liệu biến động khách hàng.")

    def _export_state_changed(self, state: str, message: str) -> None:
        if self._closing:
            return
        if state == "loading":
            self.state_banner.set_loading("Đang xuất Excel biến động dư nợ...")
        elif state == "error":
            self.state_banner.set_error(message or "Không thể xuất Excel biến động dư nợ.")
        elif state == "ready" and self.state_banner.state == "loading":
            self.state_banner.clear()

    def closeEvent(self, event) -> None:
        self._closing = True
        self.cancel_queries()
        super().closeEvent(event)
