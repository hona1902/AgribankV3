from __future__ import annotations

from datetime import datetime
from pathlib import Path

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QAbstractItemView,
    QComboBox,
    QDialog,
    QFileDialog,
    QHBoxLayout,
    QMessageBox,
    QTabWidget,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from agribank_v3.features.credit.summary.dashboard_charts import (
    DashboardBarChart,
    branch_bar_values,
    branch_empty_message,
    growth_series,
    overview_series,
)
from agribank_v3.features.credit.summary.dashboard_repository import NimDashboardRepository
from agribank_v3.features.credit.summary.dashboard_export import (
    MONEY_HEADERS,
    PERCENT_HEADERS,
    SHEET_BY_TAB,
    DashboardNimExportService,
    export_dashboard_rows,
)
from agribank_v3.features.credit.summary.dashboard_service import (
    DashboardFilters,
    DashboardNimData,
    build_nim_dashboard,
    metric_value_kind,
)
from agribank_v3.features.credit.summary.models import SummaryDataType
from agribank_v3.features.credit.summary.nim_ui_config import NimUiConfig, get_nim_ui_config
from agribank_v3.features.credit.summary.officer_history.charts import AnalysisLineChart
from agribank_v3.features.credit.summary.officer_history.models import (
    METRIC_AVERAGE_RATE,
    METRIC_BALANCE,
    METRIC_BALANCE_GROWTH,
    METRIC_NIM_AFTER,
    METRIC_NIM_BEFORE,
)
from agribank_v3.features.credit.summary.officer_history.widgets import (
    FitTableWidget,
    MultiLineHeaderView,
    NumericTableWidgetItem,
    format_money_vn,
    format_percent_vn,
)
from agribank_v3.features.credit.summary.repository import SummaryRepository
from agribank_v3.ui.components.controls import (
    combo_box as shared_combo_box,
    configure_combo_popup_width,
    populate_combo as shared_populate_combo,
    primary_button,
    secondary_button,
)
from agribank_v3.ui.components.kpi import KpiMetric, MetricGrid


class NimDashboardWindow(QDialog):
    def __init__(
        self,
        repository: SummaryRepository,
        data_type: SummaryDataType = SummaryDataType.NIM_DN,
        parent: QWidget | None = None,
    ) -> None:
        super().__init__(parent)
        self.repository = NimDashboardRepository(repository)
        self.data_type = data_type
        self.ui_config = get_nim_ui_config(data_type)
        self.dashboard_data: DashboardNimData | None = None
        self.visible_export_rows: dict[str, list[dict[str, object]]] = {}
        self.repository.unit_directory.add_listener(self._unit_directory_changed)
        self.setWindowTitle(self.ui_config.dashboard_title)
        self.setWindowFlags(
            self.windowFlags()
            | Qt.WindowType.WindowMinimizeButtonHint
            | Qt.WindowType.WindowMaximizeButtonHint
            | Qt.WindowType.WindowCloseButtonHint
        )
        self.setWindowModality(Qt.WindowModality.NonModal)
        self.setSizeGripEnabled(True)
        self.resize(1220, 780)
        self.setMinimumSize(980, 640)
        self._build_ui()
        self._reload_filter_options()
        self.reload()

    def closeEvent(self, event) -> None:
        self.repository.unit_directory.remove_listener(self._unit_directory_changed)
        super().closeEvent(event)

    def _unit_directory_changed(self) -> None:
        self.reload()

    def _build_ui(self) -> None:
        layout = QVBoxLayout(self)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(10)

        filter_row = QHBoxLayout()
        self.period_from_combo = _combo("Từ kỳ", "Từ kỳ")
        self.period_to_combo = _combo("Đến kỳ", "Đến kỳ")
        self.branch_combo = _combo("Chi nhánh", "Tất cả chi nhánh")
        self.transaction_office_combo = _combo("Phòng GD", "Tất cả Phòng GD")
        self.customer_type_combo = _combo("Loại KH", "Tất cả loại KH")
        self.metric_combo = _metric_combo(self.ui_config)
        self.metric_combo.currentIndexChanged.connect(lambda _index: self._render_branch_tab())
        apply_button = primary_button("Áp dụng")
        apply_button.clicked.connect(self.reload)
        clear_button = secondary_button("Xóa lọc")
        clear_button.clicked.connect(self.clear_filters)
        export_button = secondary_button("Xuất toàn bộ")
        export_button.clicked.connect(self.export_all_tabs)
        close_button = secondary_button("Đóng")
        close_button.clicked.connect(self.close)
        for widget in (
            self.period_from_combo,
            self.period_to_combo,
            self.branch_combo,
            self.transaction_office_combo,
            self.customer_type_combo,
            self.metric_combo,
            apply_button,
            clear_button,
            export_button,
            close_button,
        ):
            filter_row.addWidget(widget)
        filter_row.addStretch()
        layout.addLayout(filter_row)

        self.kpi_grid = MetricGrid()
        layout.addWidget(self.kpi_grid)

        self.tabs = QTabWidget()
        self._build_overview_tab()
        self._build_branch_tab()
        self._build_growth_tab()
        self._build_detail_tab()
        layout.addWidget(self.tabs, stretch=1)

    def _build_overview_tab(self) -> None:
        tab = QWidget()
        layout = QVBoxLayout(tab)
        layout.addLayout(self._tab_actions("overview"))
        self.overview_chart = AnalysisLineChart()
        self.overview_table = _table()
        layout.addWidget(self.overview_chart, stretch=2)
        layout.addWidget(self.overview_table, stretch=1)
        self.tabs.addTab(tab, "Tổng quan theo kỳ")

    def _build_branch_tab(self) -> None:
        tab = QWidget()
        layout = QVBoxLayout(tab)
        layout.addLayout(self._tab_actions("branch"))
        self.branch_chart = DashboardBarChart()
        self.branch_table = _table()
        layout.addWidget(self.branch_chart, stretch=2)
        layout.addWidget(self.branch_table, stretch=1)
        self.tabs.addTab(tab, "So sánh chi nhánh")

    def _build_growth_tab(self) -> None:
        tab = QWidget()
        layout = QVBoxLayout(tab)
        layout.addLayout(self._tab_actions("growth"))
        self.growth_chart = AnalysisLineChart()
        self.growth_table = _table()
        layout.addWidget(self.growth_chart, stretch=2)
        layout.addWidget(self.growth_table, stretch=1)
        self.tabs.addTab(tab, "Tăng trưởng")

    def _build_detail_tab(self) -> None:
        tab = QWidget()
        layout = QVBoxLayout(tab)
        layout.addLayout(self._tab_actions("detail"))
        self.detail_table = _table(wrap_header=True)
        layout.addWidget(self.detail_table)
        self.tabs.addTab(tab, "Bảng dữ liệu chi tiết")

    def _tab_actions(self, tab_key: str) -> QHBoxLayout:
        row = QHBoxLayout()
        row.addStretch()
        button = secondary_button("Xuất Excel")
        button.clicked.connect(lambda _checked=False, key=tab_key: self.export_tab(key))
        row.addWidget(button)
        return row

    def reload(self) -> None:
        self._reload_filter_options()
        try:
            self.dashboard_data = build_nim_dashboard(self.repository, self.data_type, self._filters())
        except Exception as exc:
            QMessageBox.warning(self, self.windowTitle(), str(exc))
            return
        self._render_kpis()
        self._render_overview_tab()
        self._render_branch_tab()
        self._render_growth_tab()
        self._render_detail_tab()

    def clear_filters(self) -> None:
        for combo in (
            self.period_from_combo,
            self.period_to_combo,
            self.branch_combo,
            self.transaction_office_combo,
            self.customer_type_combo,
        ):
            combo.setCurrentIndex(0)
        metric_index = self.metric_combo.findData(METRIC_BALANCE)
        self.metric_combo.setCurrentIndex(metric_index if metric_index >= 0 else 0)
        self.reload()

    def export_tab(self, tab_key: str) -> None:
        service = self._export_service()
        if service is None:
            return
        rows = self.visible_export_rows.get(tab_key) or service.rows_for_tab(tab_key)
        path, _ = QFileDialog.getSaveFileName(
            self,
            "Xuất Dashboard NIM",
            _default_export_file(tab_key, self.ui_config),
            "Excel (*.xlsx)",
        )
        if not path:
            return
        try:
            output = export_dashboard_rows(rows, Path(path), sheet_name=self.ui_config.dashboard_sheets.get(tab_key, "DashboardNIM"))
        except Exception as exc:
            QMessageBox.warning(self, "Xuất Excel", str(exc))
            return
        QMessageBox.information(self, "Xuất Excel", f"Đã xuất: {output}")

    def export_all_tabs(self) -> None:
        service = self._export_service()
        if service is None:
            return
        path, _ = QFileDialog.getSaveFileName(
            self,
            "Xuất toàn bộ Dashboard NIM",
            _default_export_file("TatCa", self.ui_config),
            "Excel (*.xlsx)",
        )
        if not path:
            return
        try:
            output = service.export_all_tabs(Path(path))
        except Exception as exc:
            QMessageBox.warning(self, "Xuất toàn bộ", str(exc))
            return
        QMessageBox.information(self, "Xuất toàn bộ", f"Đã xuất: {output}")

    def _reload_filter_options(self) -> None:
        filters = self._filters().as_query_filters() if hasattr(self, "period_from_combo") else {}
        _populate_combo_preserve(
            self.period_from_combo,
            self.repository.distinct_values(self.data_type, "period", filters=filters, exclude="period_from"),
        )
        _populate_combo_preserve(
            self.period_to_combo,
            self.repository.distinct_values(self.data_type, "period", filters=filters, exclude="period_to"),
        )
        _populate_combo_preserve(
            self.branch_combo,
            self.repository.distinct_values(self.data_type, "branch", filters=filters, exclude="branch"),
        )
        _populate_combo_preserve(
            self.transaction_office_combo,
            self.repository.distinct_values(self.data_type, "transaction_office", filters=filters, exclude="transaction_office"),
        )
        _populate_combo_preserve(
            self.customer_type_combo,
            _customer_type_items(self.repository.distinct_values(self.data_type, "customer_type", filters=filters, exclude="customer_type")),
        )

    def _filters(self) -> DashboardFilters:
        return DashboardFilters(
            period_from=str(self.period_from_combo.currentData() or ""),
            period_to=str(self.period_to_combo.currentData() or ""),
            branch=str(self.branch_combo.currentData() or ""),
            transaction_office=str(self.transaction_office_combo.currentData() or ""),
            customer_type=str(self.customer_type_combo.currentData() or ""),
            metric=str(self.metric_combo.currentData() or METRIC_BALANCE),
        )

    def _render_kpis(self) -> None:
        data = self.dashboard_data
        if data is None:
            return
        self.kpi_grid.set_metrics([_dashboard_kpi_metric(metric) for metric in data.kpis])

    def _render_overview_tab(self) -> None:
        data = self.dashboard_data
        if data is None:
            return
        rows = self._export_service().overview_by_period_rows()
        self.visible_export_rows["overview"] = rows
        self.overview_chart.set_series(overview_series(data.period_rows, self.ui_config), empty_message="Không có dữ liệu NIM theo kỳ.")
        headers = [
            "Kỳ",
            self.ui_config.total_balance_label,
        ]
        widths = [76, 138]
        if self.ui_config.include_average_rate:
            headers.append("Lãi suất bình quân")
            widths.append(108)
        headers.extend(["NIM trước ĐC", "NIM sau ĐC", self.ui_config.balance_delta_label, self.ui_config.growth_percent_label])
        widths.extend([96, 96, 142, 130])
        _render_dict_table(
            self.overview_table,
            tuple(headers),
            rows,
            tuple(widths),
            metric=str(self.metric_combo.currentData() or METRIC_BALANCE),
        )

    def _render_branch_tab(self) -> None:
        data = self.dashboard_data
        if data is None:
            return
        metric = str(self.metric_combo.currentData() or METRIC_BALANCE)
        metric_label = self.ui_config.metric_labels().get(metric, metric)
        rows = self._export_service().branch_comparison_rows()
        self.visible_export_rows["branch"] = rows
        bars = branch_bar_values(data.branch_rows, metric)
        self.branch_chart.set_bars(
            bars,
            value_kind=metric_value_kind(metric),
            metric_label=metric_label,
            empty_message=branch_empty_message(metric),
        )
        headers = ["Kỳ", "Tên chi nhánh", self.ui_config.balance_label]
        widths = [78, 180, 130]
        if self.ui_config.include_average_rate:
            headers.append("Lãi suất bình quân")
            widths.append(110)
        headers.extend(["NIM trước ĐC", "NIM sau ĐC", "Chỉ tiêu đang chọn", "Giá trị chỉ tiêu"])
        widths.extend([100, 100, 140, 112])
        _render_dict_table(
            self.branch_table,
            tuple(headers),
            rows,
            tuple(widths),
            metric=metric,
        )

    def _render_growth_tab(self) -> None:
        data = self.dashboard_data
        if data is None:
            return
        rows = self._export_service().growth_rows()
        self.visible_export_rows["growth"] = rows
        self.growth_chart.set_series(
            growth_series(data.period_rows, self.ui_config),
            empty_message="Không có dữ liệu tăng trưởng.",
            single_point_message="Chưa đủ dữ liệu lịch sử để hiển thị xu hướng tăng trưởng.",
        )
        _render_dict_table(
            self.growth_table,
            (
                "Kỳ",
                "Tên chi nhánh",
                "Phòng GD",
                "Loại KH",
                self.ui_config.balance_label,
                self.ui_config.balance_delta_label,
                self.ui_config.growth_percent_label,
                "Biến động NIM trước ĐC",
                "Biến động NIM sau ĐC",
            ),
            rows,
            (74, 150, 105, 84, 124, 138, 118, 118, 118),
            metric=str(self.metric_combo.currentData() or METRIC_BALANCE),
        )

    def _render_detail_tab(self) -> None:
        data = self.dashboard_data
        if data is None:
            return
        rows = self._export_service().detail_rows()
        self.visible_export_rows["detail"] = rows
        headers = ["Kỳ", "Tên chi nhánh", "Phòng GD", "Loại KH", self.ui_config.balance_label, "NIM trước ĐC", "NIM sau ĐC", self.ui_config.growth_percent_label, self.ui_config.balance_delta_label]
        widths = [74, 150, 112, 82, 124, 94, 94, 124, 136]
        if self.ui_config.include_average_rate:
            headers.insert(5, "Lãi suất bình quân")
            widths.insert(5, 102)
        _render_dict_table(
            self.detail_table,
            tuple(headers),
            rows,
            tuple(widths),
            metric=str(self.metric_combo.currentData() or METRIC_BALANCE),
        )

    def _export_service(self) -> DashboardNimExportService | None:
        data = self.dashboard_data
        if data is None:
            return None
        return DashboardNimExportService(data, metric=str(self.metric_combo.currentData() or METRIC_BALANCE))


def _table(*, wrap_header: bool = False) -> FitTableWidget:
    table = FitTableWidget()
    if wrap_header:
        header = MultiLineHeaderView(Qt.Orientation.Horizontal, table)
        table.setHorizontalHeader(header)
        header.sectionResized.connect(lambda *_args: header.refresh_height())
    table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
    table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
    table.setSortingEnabled(True)
    return table


def _render_dict_table(
    table: QTableWidget,
    headers: tuple[str, ...],
    rows: list[dict[str, object]],
    widths: tuple[int, ...],
    *,
    metric: str,
) -> None:
    table.setSortingEnabled(False)
    table.clear()
    table.setColumnCount(len(headers))
    table.setHorizontalHeaderLabels(headers)
    table.setRowCount(len(rows))
    for row_index, row in enumerate(rows):
        for column_index, header in enumerate(headers):
            value = row.get(header)
            if _is_numeric_header(header, metric):
                item = NumericTableWidgetItem(_format_display_value(header, value, metric), None if value in (None, "") else float(value))
            else:
                item = QTableWidgetItem(str(value or ""))
            item.setTextAlignment(_alignment_for_header(header, metric))
            table.setItem(row_index, column_index, item)
    if isinstance(table, FitTableWidget):
        table.set_default_widths(widths)
    else:
        for index, width in enumerate(widths):
            table.setColumnWidth(index, width)
    header = table.horizontalHeader()
    if hasattr(header, "refresh_height"):
        header.refresh_height()
    table.setSortingEnabled(True)


def _is_numeric_header(header: str, metric: str) -> bool:
    if header in MONEY_HEADERS or header in PERCENT_HEADERS:
        return True
    if header == "Giá trị chỉ tiêu":
        return metric in {METRIC_BALANCE, METRIC_NIM_BEFORE, METRIC_NIM_AFTER, METRIC_AVERAGE_RATE, METRIC_BALANCE_GROWTH}
    return False


def _format_display_value(header: str, value: object, metric: str) -> str:
    if value is None or value == "":
        return "N/A"
    if header in MONEY_HEADERS:
        return format_money_vn(value, signed=header.startswith("Tăng/giảm"))
    if header in PERCENT_HEADERS:
        return format_percent_vn(value, signed=header.startswith("Tăng trưởng") or header.startswith("Biến động"))
    if header == "Giá trị chỉ tiêu":
        if metric == METRIC_BALANCE:
            return format_money_vn(value)
        return format_percent_vn(value, signed=metric == METRIC_BALANCE_GROWTH)
    return str(value)


def _alignment_for_header(header: str, metric: str) -> Qt.AlignmentFlag:
    if header == "Kỳ":
        return Qt.AlignmentFlag.AlignCenter
    if _is_numeric_header(header, metric):
        return Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter
    return Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter


def _combo(tooltip: str, first_label: str) -> QComboBox:
    combo = shared_combo_box(first_label, minimum_width=110, maximum_width=180)
    combo.setToolTip(tooltip)
    return combo


def _metric_combo(ui_config: NimUiConfig) -> QComboBox:
    combo = shared_combo_box("Chỉ tiêu chính", minimum_width=150, maximum_width=220)
    combo.clear()
    combo.setToolTip("Chỉ tiêu chính")
    labels = ui_config.metric_labels()
    for metric in ui_config.metric_order(include_growth=True):
        combo.addItem(labels[metric], metric)
    configure_combo_popup_width(combo, minimum_popup_width=240)
    return combo


def _populate_combo_preserve(combo: QComboBox, values: list[str] | list[tuple[str, str]]) -> None:
    shared_populate_combo(combo, values)


def _dashboard_kpi_metric(metric) -> KpiMetric:
    value = str(getattr(metric, "value", "") or "").strip()
    label = str(getattr(metric, "label", "") or "")
    label_key = label.casefold()
    if "dư nợ" in label_key or "nguồn vốn" in label_key or "số dư" in label_key:
        number = _parse_vn_number(value)
        return KpiMetric(label, number, "money", full_value=number)
    if "nim" in label_key or "lãi suất" in label_key:
        number = _parse_vn_number(value)
        return KpiMetric(label, number, "percent", full_value=number)
    return KpiMetric(label, value, "text")


def _parse_vn_number(value: str) -> float | None:
    text = str(value or "").strip()
    if not text:
        return None
    text = text.replace("đồng", "").replace("%", "").strip()
    text = text.replace(".", "").replace(",", ".")
    try:
        return float(text)
    except ValueError:
        return None


def _customer_type_items(values: list[str]) -> list[tuple[str, str]]:
    preferred = {
        "Cá nhân (CN)": "Cá nhân",
        "Pháp nhân": "Pháp nhân",
        "Tổ chức (TC)": "Tổ chức",
    }
    ordered: list[tuple[str, str]] = []
    for value, label in preferred.items():
        if value in values:
            ordered.append((label, value))
    for value in values:
        if value not in preferred:
            ordered.append((value, value))
    return ordered


def _default_export_file(tab_key: str, ui_config: NimUiConfig) -> str:
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    suffix = ui_config.dashboard_sheets.get(tab_key, str(tab_key))
    return f"DashboardNIM_{suffix}_{stamp}.xlsx"
