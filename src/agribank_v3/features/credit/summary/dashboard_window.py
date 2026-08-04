from __future__ import annotations

from datetime import datetime
from pathlib import Path

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QAbstractItemView,
    QButtonGroup,
    QComboBox,
    QDialog,
    QFileDialog,
    QHBoxLayout,
    QLabel,
    QMessageBox,
    QRadioButton,
    QTabWidget,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from agribank_v3.features.credit.summary.dashboard_charts import (
    DashboardBarChart,
    DashboardBranchComparisonChart,
    branch_bar_values,
    branch_empty_message,
    branch_period_pair_values,
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
    export_dashboard_workbook,
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


OVERVIEW_MODE_ALL = "all"
OVERVIEW_MODE_ENDPOINTS = "endpoints"
BRANCH_MODE_PERIOD_COMPARE = "period_compare"
BRANCH_MODE_CURRENT = "current"
DETAIL_MODE_BRANCH = "branch"
DETAIL_MODE_OFFICE = "office"


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
        self.overview_table_mode = OVERVIEW_MODE_ALL
        self.branch_compare_mode = BRANCH_MODE_PERIOD_COMPARE
        self.detail_group_mode = DETAIL_MODE_BRANCH
        self._radio_groups: list[QButtonGroup] = []
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
        layout.addLayout(
            self._radio_row(
                "Chế độ bảng:",
                (
                    ("Hiện toàn bộ các kỳ", OVERVIEW_MODE_ALL),
                    ("Chỉ so sánh Từ kỳ và Đến kỳ", OVERVIEW_MODE_ENDPOINTS),
                ),
                selected=OVERVIEW_MODE_ALL,
                callback=self._overview_mode_changed,
            )
        )
        layout.addWidget(self.overview_chart, stretch=2)
        layout.addWidget(self.overview_table, stretch=1)
        self.tabs.addTab(tab, "Tổng quan theo kỳ")

    def _build_branch_tab(self) -> None:
        tab = QWidget()
        layout = QVBoxLayout(tab)
        layout.addLayout(self._tab_actions("branch"))
        layout.addLayout(
            self._radio_row(
                "Chế độ so sánh:",
                (
                    ("So sánh Từ kỳ đến kỳ", BRANCH_MODE_PERIOD_COMPARE),
                    ("Kỳ hiện tại", BRANCH_MODE_CURRENT),
                ),
                selected=BRANCH_MODE_PERIOD_COMPARE,
                callback=self._branch_mode_changed,
            )
        )
        self.branch_compare_chart = DashboardBranchComparisonChart()
        self.branch_current_chart = DashboardBarChart()
        self.branch_current_chart.hide()
        self.branch_chart = self.branch_compare_chart
        self.branch_table = _table()
        layout.addWidget(self.branch_compare_chart, stretch=2)
        layout.addWidget(self.branch_current_chart, stretch=2)
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
        layout.addLayout(
            self._radio_row(
                "Chế độ nhóm:",
                (
                    ("Tổng hợp theo chi nhánh", DETAIL_MODE_BRANCH),
                    ("Chi tiết theo Hội sở/PGD", DETAIL_MODE_OFFICE),
                ),
                selected=DETAIL_MODE_BRANCH,
                callback=self._detail_mode_changed,
            )
        )
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

    def _radio_row(
        self,
        label: str,
        options: tuple[tuple[str, str], ...],
        *,
        selected: str,
        callback,
    ) -> QHBoxLayout:
        row = QHBoxLayout()
        row.setSpacing(12)
        row.addWidget(QLabel(label))
        group = QButtonGroup(self)
        group.setExclusive(True)
        self._radio_groups.append(group)
        for text, value in options:
            radio = QRadioButton(text)
            radio.setProperty("modeValue", value)
            radio.setChecked(value == selected)
            group.addButton(radio)
            row.addWidget(radio)
            radio.toggled.connect(lambda checked, mode=value: checked and callback(mode))
        row.addStretch()
        return row

    def _overview_mode_changed(self, mode: str) -> None:
        self.overview_table_mode = mode
        self._render_overview_tab()

    def _branch_mode_changed(self, mode: str) -> None:
        self.branch_compare_mode = mode
        self._render_branch_tab()

    def _detail_mode_changed(self, mode: str) -> None:
        self.detail_group_mode = mode
        self._render_detail_tab()

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
            output = export_dashboard_rows(
                rows,
                Path(path),
                sheet_name=self.ui_config.dashboard_sheets.get(tab_key, "DashboardNIM"),
                metadata=self._export_metadata(tab_key),
            )
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
            output = export_dashboard_workbook(
                [
                    (self.ui_config.dashboard_sheets["overview"], self.visible_export_rows.get("overview") or service.overview_by_period_rows()),
                    (self.ui_config.dashboard_sheets["branch"], self.visible_export_rows.get("branch") or service.branch_comparison_rows()),
                    (self.ui_config.dashboard_sheets["growth"], self.visible_export_rows.get("growth") or service.growth_rows()),
                    (self.ui_config.dashboard_sheets["detail"], self.visible_export_rows.get("detail") or service.detail_rows()),
                ],
                Path(path),
                metadata=self._export_metadata("all"),
            )
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
        service = self._export_service()
        if service is None:
            return
        rows = service.overview_endpoint_rows() if self.overview_table_mode == OVERVIEW_MODE_ENDPOINTS else service.overview_by_period_rows()
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
        service = self._export_service()
        if service is None:
            return
        metric = str(self.metric_combo.currentData() or METRIC_BALANCE)
        metric_label = self.ui_config.metric_labels().get(metric, metric)
        if self.branch_compare_mode == BRANCH_MODE_CURRENT:
            rows = service.branch_current_rows()
            self.visible_export_rows["branch"] = rows
            bars = branch_bar_values(data.branch_rows, metric)
            self.branch_compare_chart.hide()
            self.branch_current_chart.show()
            self.branch_chart = self.branch_current_chart
            self.branch_current_chart.set_bars(
                bars,
                value_kind=metric_value_kind(metric),
                metric_label=metric_label,
                empty_message=branch_empty_message(metric),
            )
            headers = ["Kỳ", "Mã chi nhánh", "Tên chi nhánh", self.ui_config.balance_label]
            widths = [72, 96, 178, 126]
            if self.ui_config.include_average_rate:
                headers.append("Lãi suất bình quân")
                widths.append(108)
            headers.extend(["NIM trước ĐC", "NIM sau ĐC", "Chỉ tiêu đang chọn", "Giá trị chỉ tiêu"])
            widths.extend([96, 96, 136, 112])
        else:
            rows = service.branch_period_comparison_rows()
            self.visible_export_rows["branch"] = rows
            pairs, from_period, to_period = branch_period_pair_values(
                data.branch_rows,
                metric,
                from_period=data.filters.period_from,
                to_period=data.filters.period_to,
            )
            self.branch_current_chart.hide()
            self.branch_compare_chart.show()
            self.branch_chart = self.branch_compare_chart
            self.branch_compare_chart.set_pairs(
                pairs,
                value_kind=metric_value_kind(metric),
                metric_label=metric_label,
                from_period=from_period,
                to_period=to_period,
                empty_message=branch_empty_message(metric),
            )
            headers = tuple(rows[0].keys()) if rows else _branch_compare_headers(self.ui_config, metric)
            widths = _branch_compare_widths(metric)
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
        service = self._export_service()
        if service is None:
            return
        if self.detail_group_mode == DETAIL_MODE_OFFICE:
            rows = service.detail_office_rows()
            headers = ["Kỳ", "Mã chi nhánh", "Tên chi nhánh", "Mã đơn vị", "Hội sở/Phòng GD", "Loại đơn vị", "Loại KH", self.ui_config.balance_label, "NIM trước ĐC", "NIM sau ĐC", self.ui_config.growth_percent_label, self.ui_config.balance_delta_label]
            widths = [72, 94, 150, 96, 120, 106, 82, 122, 92, 92, 116, 130]
            if self.ui_config.include_average_rate:
                headers.insert(8, "Lãi suất bình quân")
                widths.insert(8, 104)
        else:
            rows = service.detail_branch_rows()
            headers = ["Kỳ", "Mã chi nhánh", "Tên chi nhánh", "Loại KH", self.ui_config.balance_label, "NIM trước ĐC", "NIM sau ĐC", self.ui_config.growth_percent_label, self.ui_config.balance_delta_label]
            widths = [72, 94, 170, 82, 126, 94, 94, 122, 136]
            if self.ui_config.include_average_rate:
                headers.insert(5, "Lãi suất bình quân")
                widths.insert(5, 106)
        self.visible_export_rows["detail"] = rows
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

    def _export_metadata(self, tab_key: str) -> list[tuple[str, object]]:
        filters = self._filters()
        metadata: list[tuple[str, object]] = [
            ("Chức năng", self.ui_config.dashboard_title),
            ("Tab", self.ui_config.dashboard_sheets.get(tab_key, tab_key)),
            ("Từ kỳ", filters.period_from or "Tất cả"),
            ("Đến kỳ", filters.period_to or "Tất cả"),
            ("Chi nhánh", self.branch_combo.currentText()),
            ("Phòng GD", self.transaction_office_combo.currentText()),
            ("Loại KH", self.customer_type_combo.currentText()),
            ("Chỉ tiêu", self.metric_combo.currentText()),
        ]
        if tab_key == "overview":
            metadata.append(("Chế độ bảng", "Chỉ so sánh Từ kỳ và Đến kỳ" if self.overview_table_mode == OVERVIEW_MODE_ENDPOINTS else "Hiện toàn bộ các kỳ"))
        elif tab_key == "branch":
            metadata.append(("Chế độ bảng", "Kỳ hiện tại" if self.branch_compare_mode == BRANCH_MODE_CURRENT else "So sánh Từ kỳ đến kỳ"))
        elif tab_key == "detail":
            metadata.append(("Chế độ bảng", "Chi tiết theo Hội sở/PGD" if self.detail_group_mode == DETAIL_MODE_OFFICE else "Tổng hợp theo chi nhánh"))
        elif tab_key == "all":
            metadata.extend(
                [
                    ("Chế độ bảng Tổng quan", "Chỉ so sánh Từ kỳ và Đến kỳ" if self.overview_table_mode == OVERVIEW_MODE_ENDPOINTS else "Hiện toàn bộ các kỳ"),
                    ("Chế độ bảng So sánh chi nhánh", "Kỳ hiện tại" if self.branch_compare_mode == BRANCH_MODE_CURRENT else "So sánh Từ kỳ đến kỳ"),
                    ("Chế độ bảng Chi tiết", "Chi tiết theo Hội sở/PGD" if self.detail_group_mode == DETAIL_MODE_OFFICE else "Tổng hợp theo chi nhánh"),
                ]
            )
        return metadata


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


def _branch_compare_headers(ui_config: NimUiConfig, metric: str) -> tuple[str, ...]:
    if metric == METRIC_BALANCE:
        return (
            "Mã chi nhánh",
            "Tên chi nhánh",
            f"{ui_config.balance_label} Từ kỳ",
            f"{ui_config.balance_label} Đến kỳ",
            "Tăng/giảm tuyệt đối",
            "Tăng trưởng (%)",
        )
    metric_label = ui_config.metric_labels().get(metric, metric)
    return (
        "Mã chi nhánh",
        "Tên chi nhánh",
        f"{metric_label} Từ kỳ",
        f"{metric_label} Đến kỳ",
        "Thay đổi (điểm %)",
    )


def _branch_compare_widths(metric: str) -> tuple[int, ...]:
    if metric == METRIC_BALANCE:
        return (96, 190, 136, 136, 136, 116)
    return (96, 190, 130, 130, 132)


def _is_numeric_header(header: str, metric: str) -> bool:
    if _is_money_header(header) or _is_percent_header(header):
        return True
    if header == "Giá trị chỉ tiêu":
        return metric in {METRIC_BALANCE, METRIC_NIM_BEFORE, METRIC_NIM_AFTER, METRIC_AVERAGE_RATE, METRIC_BALANCE_GROWTH}
    return False


def _format_display_value(header: str, value: object, metric: str) -> str:
    if value is None or value == "":
        return "N/A"
    if _is_money_header(header):
        return format_money_vn(value, signed=header.startswith("Tăng/giảm"))
    if _is_percent_header(header):
        return format_percent_vn(value, signed=header.startswith("Tăng trưởng") or header.startswith("Biến động") or header.startswith("Thay đổi"))
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


def _is_money_header(header: str) -> bool:
    if header in MONEY_HEADERS:
        return True
    lowered = header.casefold()
    return (
        "dư nợ" in lowered
        or "nguồn vốn" in lowered
        or "số dư" in lowered
        or "tăng/giảm tuyệt đối" in lowered
    ) and "tăng trưởng" not in lowered


def _is_percent_header(header: str) -> bool:
    if header in PERCENT_HEADERS:
        return True
    lowered = header.casefold()
    return "nim" in lowered or "lãi suất" in lowered or "tăng trưởng" in lowered or "điểm %" in lowered


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
