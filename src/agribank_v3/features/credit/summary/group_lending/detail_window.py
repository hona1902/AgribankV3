from __future__ import annotations

from PySide6.QtWidgets import QDialog, QHBoxLayout, QLabel, QPushButton, QVBoxLayout, QWidget

from agribank_v3.features.credit.summary.customer.table_models import CustomerTableModel
from agribank_v3.features.credit.summary.customer.widgets import CustomerTableView, Pager
from agribank_v3.ui.components.kpi import KpiMetric, MetricGrid
from agribank_v3.ui.components.controls import secondary_button

from .models import GroupLendingFilters, GroupLendingRow
from .service import GroupLendingService


MEMBER_COLUMNS = (
    ("Mã khách hàng", "Mã khách hàng", "text"),
    ("Tên khách hàng", "Tên khách hàng", "text"),
    ("Loại khách hàng", "Loại khách hàng", "text"),
    ("CBTD", "CBTD", "text"),
    ("Số món", "Số món", "integer"),
    ("Tổng dư nợ", "Tổng dư nợ", "money"),
    ("Nhóm nợ cao nhất", "Nhóm nợ cao nhất", "text"),
    ("Ngắn hạn", "Ngắn hạn", "money"),
    ("Trung hạn", "Trung hạn", "money"),
    ("Dài hạn", "Dài hạn", "money"),
)


class GroupLendingDetailWindow(QDialog):
    def __init__(
        self,
        service: GroupLendingService,
        period: str,
        group_row: GroupLendingRow,
        filters: GroupLendingFilters,
        parent: QWidget | None = None,
    ) -> None:
        super().__init__(parent)
        self.service = service
        self.period = period
        self.group_row = group_row
        self.filters = filters
        self.page = 1
        self.page_size = 100
        self.setWindowTitle("Chi tiết cho vay qua tổ - AgribankV3")
        self.resize(980, 620)
        self._build_ui()
        self.refresh()

    def _build_ui(self) -> None:
        layout = QVBoxLayout(self)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(8)
        title = QLabel(f"{self.group_row.group_code} - {self.group_row.group_name}")
        title.setObjectName("PageTitle")
        layout.addWidget(title)
        self.metrics = MetricGrid()
        self.metrics.set_metrics(
            [
                KpiMetric("Kỳ", self.period, "text"),
                KpiMetric("Loại Hội", self.group_row.association_label, "text"),
                KpiMetric("Tổ trưởng", self.group_row.leader_name, "text"),
                KpiMetric("Số tổ viên", self.group_row.member_count, "count"),
                KpiMetric("Tổng dư nợ", self.group_row.total_balance, "money"),
            ]
        )
        layout.addWidget(self.metrics)
        self.model = CustomerTableModel(MEMBER_COLUMNS, self)
        self.table = CustomerTableView(self)
        self.table.setModel(self.model)
        self.table.apply_default_widths((110, 220, 110, 180, 80, 140, 120, 120, 120, 120))
        layout.addWidget(self.table, stretch=1)
        self.pager = Pager(self)
        self.pager.pageChanged.connect(self._page_changed)
        self.pager.pageSizeChanged.connect(self._page_size_changed)
        layout.addWidget(self.pager)
        row = QHBoxLayout()
        close_button = secondary_button("Đóng")
        close_button.clicked.connect(self.accept)
        row.addStretch()
        row.addWidget(close_button)
        layout.addLayout(row)

    def refresh(self) -> None:
        result = self.service.get_group_member_page(
            self.period,
            self.group_row.group_code,
            self.filters,
            page=self.page,
            page_size=self.page_size,
        )
        self.model.set_rows(result.rows)
        self.pager.set_state(page=result.page, page_size=result.page_size, total_rows=result.total_rows)

    def _page_changed(self, page: int) -> None:
        self.page = int(page or 1)
        self.refresh()

    def _page_size_changed(self, page_size: int) -> None:
        self.page = 1
        self.page_size = int(page_size or 100)
        self.refresh()
