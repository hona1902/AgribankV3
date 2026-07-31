from __future__ import annotations

from pathlib import Path

from PySide6.QtGui import QDesktopServices
from PySide6.QtCore import QUrl
from PySide6.QtWidgets import QFileDialog, QHBoxLayout, QMessageBox, QVBoxLayout, QWidget

from agribank_v3.features.credit.summary.customer.async_query import AsyncQueryController
from agribank_v3.features.credit.summary.customer.export_service import (
    IMPORT_FILE_COLUMNS,
    IMPORT_RUN_COLUMNS,
    export_import_history,
    suggested_customer_export_name,
)
from agribank_v3.features.credit.summary.customer.repository import CustomerRepository
from agribank_v3.features.credit.summary.customer.table_models import CustomerTableModel
from agribank_v3.features.credit.summary.customer.widgets import CustomerTableView, Pager, QueryStateBanner, secondary_button


class ImportHistoryTab(QWidget):
    def __init__(self, repository: CustomerRepository, parent=None) -> None:
        super().__init__(parent)
        self.repository = repository
        self.page = 1
        self.page_size = 100
        self.query_controller = AsyncQueryController(self, max_cache_entries=16)
        layout = QVBoxLayout(self)
        actions = QHBoxLayout()
        refresh_button = secondary_button("Làm mới")
        refresh_button.clicked.connect(lambda: self.refresh(use_cache=False))
        export_button = secondary_button("Xuất Excel")
        export_button.clicked.connect(self.export_excel)
        folder_button = secondary_button("Mở thư mục nguồn")
        folder_button.clicked.connect(self.open_source_folder)
        actions.addWidget(refresh_button)
        actions.addWidget(export_button)
        actions.addWidget(folder_button)
        actions.addStretch()
        layout.addLayout(actions)
        self.state_banner = QueryStateBanner()
        self.state_banner.retryRequested.connect(lambda: self.refresh(use_cache=False))
        layout.addWidget(self.state_banner)
        self.run_model = CustomerTableModel(IMPORT_RUN_COLUMNS, self)
        self.run_table = CustomerTableView()
        self.run_table.setModel(self.run_model)
        self.run_table.doubleClicked.connect(self._run_selected)
        layout.addWidget(self.run_table, stretch=2)
        self.file_model = CustomerTableModel(IMPORT_FILE_COLUMNS, self)
        self.file_table = CustomerTableView()
        self.file_table.setModel(self.file_model)
        layout.addWidget(self.file_table, stretch=1)
        self.pager = Pager()
        self.pager.pageChanged.connect(self._page_changed)
        self.pager.pageSizeChanged.connect(self._page_size_changed)
        layout.addWidget(self.pager)
        self.run_table.apply_default_widths((70, 90, 260, 70, 110, 110, 90, 160, 160, 120, 120, 220))
        self.file_table.apply_default_widths((180, 320, 260, 90, 90, 110, 110, 90, 220))

    def refresh(self, *args, use_cache: bool = True) -> None:
        page = self.page
        page_size = self.page_size
        cache_key = ("import_history", page, page_size)
        self.query_controller.run(
            "customer_import_history",
            lambda: self._load_payload(page, page_size),
            self._apply_payload,
            self._query_failed,
            cache_key=cache_key,
            use_cache=use_cache,
            state_callback=self._state_changed,
        )

    def export_excel(self) -> None:
        path, _ = QFileDialog.getSaveFileName(
            self,
            "Xuất lịch sử import",
            suggested_customer_export_name("LichSuImport"),
            "Excel (*.xlsx)",
        )
        if not path:
            return
        try:
            output = export_import_history(self.repository, Path(path))
        except Exception as exc:
            QMessageBox.warning(self, "Xuất lịch sử import", str(exc))
            return
        QMessageBox.information(self, "Xuất lịch sử import", f"Đã xuất: {output}")

    def open_source_folder(self) -> None:
        row = self._selected_run()
        if not row:
            return
        folder = Path(str(row.get("source_folder") or ""))
        if folder.exists():
            QDesktopServices.openUrl(QUrl.fromLocalFile(str(folder)))

    def _run_selected(self, index) -> None:
        row = self.run_model.raw_row(index.row())
        if row:
            self.file_model.set_rows(self.repository.import_files(int(row.get("id") or 0)))

    def _selected_run(self) -> dict[str, object]:
        indexes = self.run_table.selectionModel().selectedRows() if self.run_table.selectionModel() else []
        if indexes:
            return self.run_model.raw_row(indexes[0].row())
        return self.run_model.raw_row(0)

    def _page_changed(self, page: int) -> None:
        self.page = max(1, int(page or 1))
        self.refresh()

    def _page_size_changed(self, page_size: int) -> None:
        self.page = 1
        self.page_size = int(page_size or 100)
        self.refresh()

    def invalidate_cache(self) -> None:
        self.query_controller.invalidate_cache()

    def cancel_queries(self) -> None:
        self.query_controller.cancel_pending()

    def wait_for_queries(self) -> None:
        self.query_controller.wait_for_idle()

    def _load_payload(self, page: int, page_size: int) -> dict[str, object]:
        result = self.repository.import_runs(page=page, page_size=page_size)
        files = self.repository.import_files(int(result.rows[0].get("id") or 0)) if result.rows else []
        return {"result": result, "files": files}

    def _apply_payload(self, payload: dict[str, object]) -> None:
        result = payload["result"]
        self.run_model.set_rows(result.rows)
        self.pager.set_state(page=result.page, page_size=result.page_size, total_rows=result.total_rows)
        self.file_model.set_rows(list(payload.get("files") or []))
        if result.total_rows:
            self.state_banner.clear()
        else:
            self.state_banner.set_empty("Chưa có lịch sử import.")

    def _query_failed(self, exc: Exception) -> None:
        self.run_model.set_rows([])
        self.file_model.set_rows([])
        self.pager.set_state(page=1, page_size=self.page_size, total_rows=0)

    def _state_changed(self, state: str, message: str) -> None:
        if state == "loading":
            self.state_banner.set_loading()
        elif state == "error":
            self.state_banner.set_error(message or "Không tải được lịch sử import.")
