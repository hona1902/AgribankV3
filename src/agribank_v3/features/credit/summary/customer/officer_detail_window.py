from __future__ import annotations

from dataclasses import replace
from pathlib import Path

from PySide6.QtCore import Qt
from PySide6.QtWidgets import QDialog, QGridLayout, QLabel, QTabWidget, QVBoxLayout, QWidget

from agribank_v3.features.credit.summary.customer.charts import CustomerLineChart
from agribank_v3.features.credit.summary.customer.officer_center_export import (
    OFFICER_COMPARE_COLUMNS,
    OFFICER_CUSTOMER_COLUMNS,
    OFFICER_LIST_COLUMNS,
    OFFICER_MOVEMENT_COLUMNS,
)
from agribank_v3.features.credit.summary.customer.officer_center_repository import (
    OFFICER_MODE_IMPORTED,
    OfficerCenterFilters,
    OfficerCenterRepository,
)
from agribank_v3.features.credit.summary.customer.repository import CustomerRepository
from agribank_v3.features.credit.summary.customer.table_models import ColumnSpec, CustomerTableModel
from agribank_v3.features.credit.summary.customer.widgets import (
    CustomerTableView,
    KpiMetric,
    MetricGrid,
    QueryStateBanner,
    fit_window_to_screen,
    primary_button,
    secondary_button,
)


OFFICER_DETAIL_TITLE = "Phân tích cán bộ tín dụng - AgribankV3"


class OfficerDetailWindow(QDialog):
    """Customer.db officer detail window shared by officer-centric screens."""

    REQUIRED_TAB_LABELS = (
        "Tổng quan",
        "Dư nợ & tăng trưởng",
        "Khách hàng quản lý",
        "NIM và lãi suất",
        "Chất lượng tín dụng",
        "Biến động danh mục",
        "So sánh đơn vị",
        "Lịch sử cán bộ",
    )

    def __init__(
        self,
        main_database_path: Path,
        *,
        officer_code: str = "",
        officer_name: str = "",
        officer_key: str = "",
        branch_code: str = "",
        transaction_office: str = "",
        filters: OfficerCenterFilters | None = None,
        parent=None,
    ) -> None:
        super().__init__(parent)
        self.main_database_path = Path(main_database_path)
        self.customer_repository = CustomerRepository(self.main_database_path)
        self.repository = OfficerCenterRepository(self.customer_repository)
        self.officer_code = str(officer_code or "").strip()
        self.officer_name = str(officer_name or "").strip()
        self.officer_key = str(officer_key or "").strip() or _officer_identity(self.officer_code, self.officer_name)
        self.branch_code = str(branch_code or "").strip()
        self.transaction_office = str(transaction_office or "").strip()
        base_filters = filters.normalized() if filters is not None else OfficerCenterFilters()
        selected = (self.officer_key if self.officer_key.startswith(("CODE:", "NAME:", "UNRESOLVED")) else self.officer_code,)
        self.filters = replace(
            base_filters,
            selected_officers=tuple(item for item in selected if item),
            mode=base_filters.mode or OFFICER_MODE_IMPORTED,
            branch_code=base_filters.branch_code or self.branch_code,
            transaction_office=base_filters.transaction_office or self.transaction_office,
        ).normalized()
        self.history_rows: list[dict[str, object]] = []
        self.customer_rows: list[dict[str, object]] = []
        self.movement_rows: list[dict[str, object]] = []
        self.compare_rows: list[dict[str, object]] = []
        self.setWindowTitle(OFFICER_DETAIL_TITLE)
        self.setModal(False)
        self.setWindowFlag(Qt.WindowType.WindowMinMaxButtonsHint, True)
        fit_window_to_screen(self, width_ratio=0.86, height_ratio=0.86, max_width=1360, max_height=860, min_width=980, min_height=650)
        self._build_ui()
        self.reload()

    def _build_ui(self) -> None:
        layout = QVBoxLayout(self)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(8)
        self.state_banner = QueryStateBanner()
        layout.addWidget(self.state_banner)
        self.metrics = MetricGrid()
        layout.addWidget(self.metrics)
        self.tabs = QTabWidget()
        self.overview_chart = CustomerLineChart("Xu hướng tổng dư nợ", value_kind="money")
        self.growth_chart = CustomerLineChart("Tăng trưởng dư nợ", value_kind="money")
        self.nim_chart = CustomerLineChart("NIM và lãi suất bình quân", value_kind="percent")
        self.debt_chart = CustomerLineChart("Chất lượng dư nợ", value_kind="percent")
        self.overview_model = CustomerTableModel(_HISTORY_COLUMNS, self)
        self.growth_model = CustomerTableModel(_GROWTH_COLUMNS, self)
        self.customer_model = CustomerTableModel(OFFICER_CUSTOMER_COLUMNS, self)
        self.nim_model = CustomerTableModel(_NIM_COLUMNS, self)
        self.debt_model = CustomerTableModel(_DEBT_COLUMNS, self)
        self.movement_model = CustomerTableModel(OFFICER_MOVEMENT_COLUMNS, self)
        self.compare_model = CustomerTableModel(OFFICER_COMPARE_COLUMNS, self)
        self.history_model = CustomerTableModel(OFFICER_LIST_COLUMNS, self)
        self._add_chart_table_tab("Tổng quan", self.overview_chart, self.overview_model, (95, 130, 100, 100, 100, 95, 120))
        self._add_chart_table_tab("Dư nợ & tăng trưởng", self.growth_chart, self.growth_model, (95, 135, 135, 105, 120, 120))
        self._add_table_tab("Khách hàng quản lý", self.customer_model, OFFICER_CUSTOMER_COLUMNS)
        self._add_chart_table_tab("NIM và lãi suất", self.nim_chart, self.nim_model, (95, 130, 130, 120, 120))
        self._add_chart_table_tab("Chất lượng dư nợ", self.debt_chart, self.debt_model, (95, 135, 135, 105, 105, 95, 95, 90, 90))
        self._add_table_tab("Biến động danh mục", self.movement_model, OFFICER_MOVEMENT_COLUMNS)
        self._add_table_tab("So sánh đơn vị", self.compare_model, OFFICER_COMPARE_COLUMNS)
        self._add_table_tab("Lịch sử cán bộ", self.history_model, OFFICER_LIST_COLUMNS)
        layout.addWidget(self.tabs, stretch=1)
        actions = QGridLayout()
        refresh_button = secondary_button("Làm mới")
        refresh_button.clicked.connect(self.reload)
        close_button = primary_button("Đóng")
        close_button.clicked.connect(self.accept)
        actions.addWidget(refresh_button, 0, 0)
        actions.addWidget(QLabel(""), 0, 1)
        actions.addWidget(close_button, 0, 2)
        actions.setColumnStretch(1, 1)
        layout.addLayout(actions)

    def _add_chart_table_tab(self, label: str, chart: CustomerLineChart, model: CustomerTableModel, widths: tuple[int, ...]) -> None:
        tab = QWidget()
        layout = QVBoxLayout(tab)
        layout.addWidget(chart, stretch=2)
        table = CustomerTableView()
        table.setModel(model)
        table.apply_default_widths(widths)
        layout.addWidget(table, stretch=1)
        self.tabs.addTab(tab, label)

    def _add_table_tab(self, label: str, model: CustomerTableModel, columns: tuple[ColumnSpec, ...]) -> None:
        tab = QWidget()
        layout = QVBoxLayout(tab)
        table = CustomerTableView()
        table.setModel(model)
        table.apply_default_widths(_default_widths(columns))
        layout.addWidget(table, stretch=1)
        self.tabs.addTab(tab, label)

    def reload(self) -> None:
        try:
            self._load_data()
        except Exception as exc:
            self.state_banner.set_error(str(exc))
            return
        self.state_banner.clear()
        self._render()

    def _load_data(self) -> None:
        periods = self.repository.distinct_periods()
        report_period = self.filters.report_period or (periods[-1] if periods else "")
        period_from = self.filters.period_from or (periods[0] if periods else "")
        period_to = self.filters.period_to or report_period
        self.filters = replace(self.filters, report_period=report_period, period_from=period_from, period_to=period_to).normalized()
        self.history_rows = self.repository.officer_period_history(self.filters)
        self.customer_rows = self.repository.officer_customers(self.filters, page=1, page_size=500, sort_by="total_customer_balance", sort_desc=True).rows
        self.movement_rows = self.repository.officer_movement(self.filters, page=1, page_size=50).rows
        self.compare_rows = self.repository.compare_officers(self.filters, page=1, page_size=50).rows

    def _render(self) -> None:
        current = self.history_rows[-1] if self.history_rows else {}
        self.metrics.set_metrics(
            [
                KpiMetric("CBTD", str(current.get("officer_name") or self.officer_name or self.officer_code), "text"),
                KpiMetric("Chi nhánh", str(current.get("branch_name") or self.branch_code), "text"),
                KpiMetric("Phòng GD", str(current.get("office_name") or self.transaction_office or "Tất cả"), "text"),
                KpiMetric("Chế độ phân tích", str(current.get("mode_label") or self.filters.mode), "text"),
                KpiMetric("Tổng dư nợ kỳ hiện tại", current.get("total_balance"), "money"),
                KpiMetric("Lãi suất bình quân hiện tại", current.get("average_rate"), "percent"),
                KpiMetric("NIM trước ĐC hiện tại", current.get("nim_before"), "percent"),
                KpiMetric("NIM sau ĐC hiện tại", current.get("nim_after"), "percent"),
            ]
        )
        self.overview_chart.set_series((("Tổng dư nợ", _points(self.history_rows, "total_balance")),))
        self.growth_chart.set_series((("Tăng/giảm", _points(_growth_rows(self.history_rows), "balance_change")),))
        self.nim_chart.set_series(
            (
                ("Lãi suất bình quân", _points(self.history_rows, "average_rate")),
                ("NIM trước ĐC", _points(self.history_rows, "nim_before")),
                ("NIM sau ĐC", _points(self.history_rows, "nim_after")),
            )
        )
        self.debt_chart.set_series(
            (
                ("Tỷ lệ nhóm 2", _points(self.history_rows, "attention_ratio")),
                ("Tỷ lệ nợ xấu", _points(self.history_rows, "bad_debt_ratio")),
            )
        )
        self.overview_model.set_rows(self.history_rows)
        self.growth_model.set_rows(_growth_rows(self.history_rows))
        self.customer_model.set_rows(self.customer_rows)
        self.nim_model.set_rows(self.history_rows)
        self.debt_model.set_rows(self.history_rows)
        self.movement_model.set_rows([dict(row, rank=index) for index, row in enumerate(self.movement_rows, start=1)])
        self.compare_model.set_rows([dict(row, rank=index) for index, row in enumerate(self.compare_rows, start=1)])
        self.history_model.set_rows([dict(row, rank=index) for index, row in enumerate(self.history_rows, start=1)])


def _officer_identity(code: str, name: str) -> str:
    code = str(code or "").strip()
    name = str(name or "").strip()
    if code:
        return f"CODE:{code}"
    if name:
        return f"NAME:{name.upper()}"
    return "UNRESOLVED"


def _points(rows: list[dict[str, object]], key: str) -> tuple[tuple[str, float], ...]:
    return tuple((str(row.get("period") or ""), float(row.get(key) or 0)) for row in rows)


def _growth_rows(rows: list[dict[str, object]]) -> list[dict[str, object]]:
    output: list[dict[str, object]] = []
    previous_balance: float | None = None
    for row in rows:
        balance = float(row.get("total_balance") or 0)
        change = None if previous_balance is None else balance - previous_balance
        growth = None if previous_balance in (None, 0) else change / previous_balance * 100
        output.append(
            {
                "period": row.get("period"),
                "total_balance": balance,
                "balance_change": change,
                "growth_rate": growth,
                "nim_before": row.get("nim_before"),
                "nim_after": row.get("nim_after"),
            }
        )
        previous_balance = balance
    return output


def _default_widths(columns: tuple[ColumnSpec, ...]) -> tuple[int, ...]:
    output: list[int] = []
    for field, _label, kind in columns:
        if field == "rank":
            output.append(55)
        elif kind in {"money", "money_signed", "money_or_blank"}:
            output.append(130)
        elif kind.startswith("percent"):
            output.append(100)
        elif "name" in field:
            output.append(170)
        elif "code" in field:
            output.append(105)
        elif kind == "integer":
            output.append(86)
        else:
            output.append(120)
    return tuple(output)


_HISTORY_COLUMNS: tuple[ColumnSpec, ...] = (
    ("period", "Kỳ", "text"),
    ("total_balance", "Tổng dư nợ", "money"),
    ("average_rate", "Lãi suất bình quân", "percent_or_blank"),
    ("nim_before", "NIM trước ĐC", "percent_or_blank"),
    ("nim_after", "NIM sau ĐC", "percent_or_blank"),
    ("customer_count", "Số KH", "integer"),
    ("officer_status", "Trạng thái", "text"),
)

_GROWTH_COLUMNS: tuple[ColumnSpec, ...] = (
    ("period", "Kỳ", "text"),
    ("total_balance", "Tổng dư nợ", "money"),
    ("balance_change", "Tăng/giảm", "money_signed"),
    ("growth_rate", "Tăng trưởng", "percent_or_blank"),
    ("nim_before", "NIM trước ĐC", "percent_or_blank"),
    ("nim_after", "NIM sau ĐC", "percent_or_blank"),
)

_NIM_COLUMNS: tuple[ColumnSpec, ...] = (
    ("period", "Kỳ", "text"),
    ("total_balance", "Tổng dư nợ", "money"),
    ("average_rate", "Lãi suất bình quân", "percent_or_blank"),
    ("nim_before", "NIM trước ĐC", "percent_or_blank"),
    ("nim_after", "NIM sau ĐC", "percent_or_blank"),
)

_DEBT_COLUMNS: tuple[ColumnSpec, ...] = (
    ("period", "Kỳ", "text"),
    ("total_balance", "Tổng dư nợ", "money"),
    ("attention_balance", "Nợ nhóm 2", "money"),
    ("bad_debt_balance", "Nợ xấu", "money"),
    ("attention_ratio", "Tỷ lệ nhóm 2", "percent_or_blank"),
    ("bad_debt_ratio", "Tỷ lệ nợ xấu", "percent_or_blank"),
    ("debt_group_1_balance", "Nhóm 1", "money"),
    ("debt_group_2_balance", "Nhóm 2", "money"),
    ("debt_group_3_balance", "Nhóm 3", "money"),
    ("debt_group_4_balance", "Nhóm 4", "money"),
    ("debt_group_5_balance", "Nhóm 5", "money"),
    ("debt_group_unknown_balance", "UNKNOWN", "money"),
    ("attention_customer_count", "KH nhóm 2", "integer"),
    ("bad_debt_customer_count", "KH nợ xấu", "integer"),
)
