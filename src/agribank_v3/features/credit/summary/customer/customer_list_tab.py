from __future__ import annotations

from pathlib import Path

from PySide6.QtCore import Qt
from PySide6.QtWidgets import QFileDialog, QHBoxLayout, QMessageBox, QVBoxLayout, QWidget

from agribank_v3.features.credit.summary.customer.async_query import AsyncQueryController
from agribank_v3.features.credit.summary.customer.customer_detail_window import CustomerDetailWindow
from agribank_v3.features.credit.summary.customer.export_service import (
    CUSTOMER_LIST_COLUMNS,
    export_customer_list,
    suggested_customer_export_name,
)
from agribank_v3.features.credit.summary.customer.repository import CustomerRepository
from agribank_v3.features.credit.summary.customer.table_models import CustomerTableModel
from agribank_v3.features.credit.summary.customer.widgets import CustomerTableView, Pager, QueryStateBanner, secondary_button


class CustomerListTab(QWidget):
    def __init__(self, repository: CustomerRepository, filters_provider, parent=None) -> None:
        super().__init__(parent)
        self.repository = repository
        self.filters_provider = filters_provider
        self.page = 1
        self.page_size = 100
        self.sort_by = "period"
        self.sort_desc = True
        self.detail_windows: list[CustomerDetailWindow] = []
        self.query_controller = AsyncQueryController(self, max_cache_entries=40)
        layout = QVBoxLayout(self)
        actions = QHBoxLayout()
        refresh_button = secondary_button("Làm mới")
        refresh_button.clicked.connect(lambda: self.refresh(use_cache=False))
        self.export_button = secondary_button("Xuất Excel")
        self.export_button.clicked.connect(self.export_excel)
        actions.addWidget(refresh_button)
        actions.addWidget(self.export_button)
        actions.addStretch()
        layout.addLayout(actions)
        self.state_banner = QueryStateBanner()
        self.state_banner.retryRequested.connect(lambda: self.refresh(use_cache=False))
        layout.addWidget(self.state_banner)
        self.model = CustomerTableModel(CUSTOMER_LIST_COLUMNS, self)
        self.table = CustomerTableView()
        self.table.setModel(self.model)
        self.table.doubleClicked.connect(self._double_clicked)
        self.table.horizontalHeader().sectionClicked.connect(self._sort_changed)
        layout.addWidget(self.table, stretch=1)
        self.pager = Pager()
        self.pager.pageChanged.connect(self._page_changed)
        self.pager.pageSizeChanged.connect(self._page_size_changed)
        layout.addWidget(self.pager)
        self.table.apply_default_widths((80, 120, 220, 100, 90, 180, 180, 86, 130, 130, 140, 130, 110, 110, 100, 100, 90, 120))

    def refresh(self, *args, use_cache: bool = True) -> None:
        if not self.repository.has_period_data():
            self.set_empty_state()
            return
        filters = self.filters_provider()
        page = self.page
        page_size = self.page_size
        sort_by = self.sort_by
        sort_desc = self.sort_desc
        cache_key = ("customer_list", filters, page, page_size, sort_by, sort_desc)
        self.query_controller.run(
            "customer_list",
            lambda: self.repository.query_customer_list(
                filters,
                page=page,
                page_size=page_size,
                sort_by=sort_by,
                sort_desc=sort_desc,
            ),
            self._apply_result,
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
            "Xuất danh sách khách hàng",
            suggested_customer_export_name("DanhSachKhachHang"),
            "Excel (*.xlsx)",
        )
        if not path:
            return
        try:
            output = export_customer_list(
                self.repository,
                self.filters_provider(),
                Path(path),
                sort_by=self.sort_by,
                sort_desc=self.sort_desc,
            )
        except Exception as exc:
            QMessageBox.warning(self, "Xuất danh sách khách hàng", str(exc))
            return
        QMessageBox.information(self, "Xuất danh sách khách hàng", f"Đã xuất: {output}")

    def _page_changed(self, page: int) -> None:
        self.page = max(1, int(page or 1))
        self.refresh()

    def _page_size_changed(self, page_size: int) -> None:
        self.page = 1
        self.page_size = int(page_size or 100)
        self.refresh()

    def _sort_changed(self, section: int) -> None:
        if 0 <= section < len(CUSTOMER_LIST_COLUMNS):
            field = CUSTOMER_LIST_COLUMNS[section][0]
            if self.sort_by == field:
                self.sort_desc = not self.sort_desc
            else:
                self.sort_by = field
                self.sort_desc = field in {"period", "total_balance"}
            self.page = 1
            self.refresh()

    def _double_clicked(self, index) -> None:
        row = self.model.raw_row(index.row())
        if not row:
            return
        dialog = CustomerDetailWindow(
            self.repository,
            str(row.get("customer_code") or ""),
            period=str(row.get("period") or ""),
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

    def invalidate_cache(self) -> None:
        self.query_controller.invalidate_cache()

    def cancel_queries(self) -> None:
        self.query_controller.cancel_pending()

    def wait_for_queries(self) -> None:
        self.query_controller.wait_for_idle()

    def set_empty_state(self, message: str = "Chưa có dữ liệu khách hàng.") -> None:
        self.cancel_queries()
        self.model.set_rows([])
        self.pager.set_state(page=1, page_size=self.page_size, total_rows=0)
        self.state_banner.set_empty(message)
        self.export_button.setEnabled(False)
        self.export_button.setToolTip("Chưa có dữ liệu để xuất")

    def _apply_result(self, result) -> None:
        self.export_button.setEnabled(True)
        self.export_button.setToolTip("")
        self.model.set_rows(result.rows)
        self.pager.set_state(page=result.page, page_size=result.page_size, total_rows=result.total_rows)
        if result.total_rows:
            self.state_banner.clear()
        else:
            self.state_banner.set_empty()

    def _query_failed(self, exc: Exception) -> None:
        self.model.set_rows([])
        self.pager.set_state(page=1, page_size=self.page_size, total_rows=0)

    def _state_changed(self, state: str, message: str) -> None:
        if state == "loading":
            self.state_banner.set_loading()
        elif state == "error":
            self.state_banner.set_error(message or "Không tải được danh sách khách hàng.")
