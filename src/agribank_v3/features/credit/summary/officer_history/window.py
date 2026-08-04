from __future__ import annotations

from pathlib import Path

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QAbstractItemView,
    QComboBox,
    QDialog,
    QFileDialog,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QTabWidget,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from agribank_v3.features.credit.summary.models import SummaryDataType
from agribank_v3.features.credit.summary.customer.filters import CustomerFilters
from agribank_v3.features.credit.summary.customer.repository import CustomerRepository
from agribank_v3.features.credit.summary.nim_ui_config import NimUiConfig, get_nim_ui_config
from agribank_v3.features.credit.summary.repository import SummaryRepository
from agribank_v3.ui.components.controls import (
    combo_box as shared_combo_box,
    configure_combo_popup_width,
    primary_button,
    secondary_button,
)
from agribank_v3.ui.components.kpi import KpiMetric, MetricGrid

from .charts import AnalysisLineChart
from .export import export_analysis_rows
from .models import (
    METRIC_AVERAGE_RATE,
    METRIC_BALANCE,
    METRIC_BALANCE_GROWTH,
    METRIC_NIM_AFTER,
    METRIC_NIM_BEFORE,
    ChartSeries,
    ComparisonRow,
    GrowthPoint,
    HistoryFilters,
    HistoryPoint,
    OfficerOverview,
)
from .repository import OfficerHistoryRepository, officer_code, officer_display_name, officer_key
from .service import (
    balance_series,
    build_multiple_officer_comparison,
    build_officer_branch_comparison,
    build_officer_growth_history,
    build_officer_overview,
    comparison_series,
    growth_series,
    metric_labels,
    overview_series,
)
from .widgets import NumericTableWidgetItem, OfficerMultiSelectCombo, format_money_vn, format_percent_vn
from .widgets import FitTableWidget


COMPARE_DEFAULT_PAGE_SIZE = 8
COMPARE_PAGE_SIZE_OPTIONS = (5, 8, 10)
COMPARE_TOP_OPTIONS = (5, 8, 10, 15, 20)
COMPARE_MODE_PAGED = "paged"
COMPARE_MODE_TOP = "top"
COMPARE_MODE_TABLE = "table"
COMPARE_MODE_ALL = "all"
COMPARE_EXPORT_ALL = "all"
COMPARE_EXPORT_VISIBLE = "visible"


class OfficerHistoryDialog(QDialog):
    def __init__(
        self,
        repository: SummaryRepository,
        data_type: SummaryDataType,
        *,
        officer: str,
        branch: str,
        transaction_office: str,
        customer_type: str,
        parent: QWidget | None = None,
    ) -> None:
        super().__init__(parent)
        self.repository = repository
        self.history_repository = OfficerHistoryRepository(repository)
        self.customer_repository = CustomerRepository(repository.main_database_path)
        self.data_type = data_type
        self.ui_config = get_nim_ui_config(data_type)
        self.officer = officer
        self.officer_code = officer_code(officer)
        self.branch = branch
        self.transaction_office = transaction_office
        self.customer_type = "" if customer_type == "Tất cả" else customer_type
        self.overview: OfficerOverview | None = None
        self.growth_rows: tuple[GrowthPoint, ...] = ()
        self.compare_rows: tuple[ComparisonRow, ...] = ()
        self.compare_selected_officers: tuple = ()
        self.compare_chart_page = 0
        self.branch_compare_rows: tuple[ComparisonRow, ...] = ()
        self.branch_compare_series = ()
        self.debt_quality_rows: tuple[dict[str, object], ...] = ()
        self.setWindowTitle(self.ui_config.officer_history_title)
        self.setWindowFlags(
            self.windowFlags()
            | Qt.WindowType.WindowMinimizeButtonHint
            | Qt.WindowType.WindowMaximizeButtonHint
            | Qt.WindowType.WindowCloseButtonHint
        )
        self.resize(1120, 760)
        self._build_ui()
        self._reload_filter_options()
        self.reload()

    def _build_ui(self) -> None:
        layout = QVBoxLayout(self)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(10)
        self._build_info(layout)
        self._build_filters(layout)

        self.tabs = QTabWidget()
        self.tabs.currentChanged.connect(lambda _index: self._render_current_tab())
        self._build_overview_tab()
        self._build_growth_tab()
        self._build_officer_compare_tab()
        self._build_branch_compare_tab()
        self._build_debt_quality_tab()
        layout.addWidget(self.tabs, stretch=1)

        actions = QHBoxLayout()
        self.refresh_button = secondary_button("Làm mới")
        self.refresh_button.clicked.connect(self.reload)
        self.export_button = secondary_button("Xuất Excel")
        self.export_button.clicked.connect(self.export_excel)
        close_button = primary_button("Đóng")
        close_button.clicked.connect(self.accept)
        actions.addWidget(self.refresh_button)
        actions.addWidget(self.export_button)
        actions.addStretch()
        actions.addWidget(close_button)
        layout.addLayout(actions)

    def _build_info(self, layout: QVBoxLayout) -> None:
        self.info_metrics = MetricGrid()
        self.info_metrics.set_empty(
            [
                self.ui_config.officer_short_label,
                "Chi nhánh",
                "Phòng GD",
                "Loại khách hàng",
                self.ui_config.current_balance_label,
                "NIM trước ĐC hiện tại",
                "NIM sau ĐC hiện tại",
            ]
        )
        layout.addWidget(self.info_metrics)

    def _build_filters(self, layout: QVBoxLayout) -> None:
        row = QHBoxLayout()
        self.period_from_combo = _combo("Từ kỳ")
        self.period_to_combo = _combo("Đến kỳ")
        self.customer_type_combo = _combo("Loại KH")
        self.transaction_office_combo = _combo("Phòng GD")
        self.apply_filter_button = primary_button("Áp dụng")
        self.apply_filter_button.clicked.connect(self.apply_filters)
        self.clear_filter_button = secondary_button("Xóa lọc")
        self.clear_filter_button.clicked.connect(self.clear_filters)
        for widget in (
            self.period_from_combo,
            self.period_to_combo,
            self.customer_type_combo,
            self.transaction_office_combo,
            self.apply_filter_button,
            self.clear_filter_button,
        ):
            row.addWidget(widget)
        row.addStretch()
        layout.addLayout(row)

    def _build_overview_tab(self) -> None:
        tab = QWidget()
        layout = QVBoxLayout(tab)
        self.overview_chart = AnalysisLineChart()
        self.overview_table = _table()
        layout.addWidget(self.overview_chart, stretch=2)
        layout.addWidget(self.overview_table, stretch=1)
        self.tabs.addTab(tab, "Tổng quan NIM")

    def _build_growth_tab(self) -> None:
        tab = QWidget()
        layout = QVBoxLayout(tab)
        self.balance_chart = AnalysisLineChart()
        self.growth_chart = AnalysisLineChart()
        self.growth_table = _table()
        layout.addWidget(self.balance_chart, stretch=1)
        layout.addWidget(self.growth_chart, stretch=1)
        layout.addWidget(self.growth_table, stretch=1)
        self.tabs.addTab(tab, f"{self.ui_config.balance_label.replace('Số dư ', '').capitalize()} & tăng trưởng")

    def _build_officer_compare_tab(self) -> None:
        tab = QWidget()
        layout = QVBoxLayout(tab)
        controls = QHBoxLayout()
        self.officer_search_input = QLineEdit()
        self.officer_search_input.setObjectName("AgribankSearchBox")
        self.officer_search_input.setPlaceholderText(f"Tìm {self.ui_config.officer_short_label.lower()}")
        self.officer_search_input.setClearButtonEnabled(True)
        self.officer_selector = OfficerMultiSelectCombo(
            placeholder=self.ui_config.officer_selector_placeholder,
            counter_label=self.ui_config.officer_selector_counter_label,
        )
        self.officer_search_input.textChanged.connect(self.officer_selector.set_filter_text)
        select_visible_button = secondary_button("Chọn đang hiển thị")
        select_visible_button.clicked.connect(self.officer_selector.select_visible)
        select_all_button = secondary_button("Chọn tất cả")
        select_all_button.clicked.connect(self.officer_selector.select_all)
        self.compare_metric_combo = _metric_combo(self.ui_config, include_growth=True)
        self.compare_metric_combo.currentIndexChanged.connect(lambda _index: self.apply_officer_comparison(show_limit_warning=False))
        apply_button = primary_button("Áp dụng")
        apply_button.clicked.connect(self.apply_officer_comparison)
        clear_button = secondary_button("Xóa lựa chọn")
        clear_button.clicked.connect(self.clear_officer_comparison)
        for widget in (
            self.officer_search_input,
            self.officer_selector,
            select_visible_button,
            select_all_button,
            self.compare_metric_combo,
            apply_button,
            clear_button,
        ):
            controls.addWidget(widget)
        controls.addStretch()
        layout.addLayout(controls)

        display_controls = QHBoxLayout()
        self.compare_mode_combo = shared_combo_box("Chế độ hiển thị", minimum_width=145, maximum_width=210)
        self.compare_mode_combo.clear()
        self.compare_mode_combo.addItem("Biểu đồ theo trang", COMPARE_MODE_PAGED)
        self.compare_mode_combo.addItem("Top N cán bộ", COMPARE_MODE_TOP)
        self.compare_mode_combo.addItem("Bảng dữ liệu", COMPARE_MODE_TABLE)
        self.compare_mode_combo.addItem("Tất cả trên một biểu đồ", COMPARE_MODE_ALL)
        configure_combo_popup_width(self.compare_mode_combo, minimum_popup_width=240)
        self.compare_mode_combo.currentIndexChanged.connect(lambda _index: self._compare_display_changed())
        self.compare_page_size_combo = shared_combo_box("Số cán bộ/trang", minimum_width=126, maximum_width=170)
        self.compare_page_size_combo.clear()
        for size in COMPARE_PAGE_SIZE_OPTIONS:
            self.compare_page_size_combo.addItem(f"{size} cán bộ/trang", size)
        self.compare_page_size_combo.setCurrentIndex(self.compare_page_size_combo.findData(COMPARE_DEFAULT_PAGE_SIZE))
        configure_combo_popup_width(self.compare_page_size_combo, minimum_popup_width=190)
        self.compare_page_size_combo.currentIndexChanged.connect(lambda _index: self._compare_page_size_changed())
        self.compare_top_combo = shared_combo_box("Top", minimum_width=82, maximum_width=110)
        self.compare_top_combo.clear()
        for size in COMPARE_TOP_OPTIONS:
            self.compare_top_combo.addItem(f"Top {size}", size)
        self.compare_top_combo.setCurrentIndex(self.compare_top_combo.findData(10))
        configure_combo_popup_width(self.compare_top_combo, minimum_popup_width=120)
        self.compare_top_combo.currentIndexChanged.connect(lambda _index: self._render_officer_compare_chart())
        self.compare_export_scope_combo = shared_combo_box("Phạm vi xuất", minimum_width=148, maximum_width=205)
        self.compare_export_scope_combo.clear()
        self.compare_export_scope_combo.addItem("Xuất toàn bộ", COMPARE_EXPORT_ALL)
        self.compare_export_scope_combo.addItem("Chỉ xuất đang hiển thị", COMPARE_EXPORT_VISIBLE)
        configure_combo_popup_width(self.compare_export_scope_combo, minimum_popup_width=220)
        self.compare_prev_button = secondary_button("Trang trước")
        self.compare_prev_button.clicked.connect(self.previous_compare_chart_page)
        self.compare_next_button = secondary_button("Trang sau")
        self.compare_next_button.clicked.connect(self.next_compare_chart_page)
        self.compare_page_label = QLabel("Trang 1/1")
        self.compare_status_label = QLabel("")
        self.compare_status_label.setObjectName("MutedText")
        for widget in (
            QLabel("Chế độ hiển thị"),
            self.compare_mode_combo,
            self.compare_page_size_combo,
            self.compare_top_combo,
            self.compare_prev_button,
            self.compare_page_label,
            self.compare_next_button,
            self.compare_export_scope_combo,
        ):
            display_controls.addWidget(widget)
        display_controls.addStretch()
        display_controls.addWidget(self.compare_status_label)
        layout.addLayout(display_controls)

        self.officer_compare_chart = AnalysisLineChart()
        self.officer_compare_table = _table()
        layout.addWidget(self.officer_compare_chart, stretch=2)
        table_filters = QHBoxLayout()
        self.compare_table_search = QLineEdit()
        self.compare_table_search.setObjectName("AgribankSearchBox")
        self.compare_table_search.setPlaceholderText(f"Tìm {self.ui_config.officer_short_label.lower()}")
        self.compare_table_search.setClearButtonEnabled(True)
        self.compare_table_search.textChanged.connect(lambda _text: self._render_officer_compare_table())
        self.compare_table_period_combo = _combo("Kỳ")
        self.compare_table_period_combo.currentTextChanged.connect(lambda _text: self._render_officer_compare_table())
        self.compare_table_branch_combo = _combo("Chi nhánh")
        self.compare_table_branch_combo.currentTextChanged.connect(lambda _text: self._render_officer_compare_table())
        for widget in (self.compare_table_search, self.compare_table_period_combo, self.compare_table_branch_combo):
            table_filters.addWidget(widget)
        table_filters.addStretch()
        layout.addLayout(table_filters)
        layout.addWidget(self.officer_compare_table, stretch=1)
        self.tabs.addTab(tab, "So sánh cán bộ")

    def _build_branch_compare_tab(self) -> None:
        tab = QWidget()
        layout = QVBoxLayout(tab)
        controls = QHBoxLayout()
        self.branch_metric_combo = _metric_combo(self.ui_config, include_growth=True, default_metric=METRIC_NIM_AFTER)
        self.branch_metric_combo.currentIndexChanged.connect(lambda _index: self.reload_branch_comparison())
        controls.addWidget(self.branch_metric_combo)
        controls.addStretch()
        layout.addLayout(controls)
        self.branch_compare_chart = AnalysisLineChart()
        self.branch_compare_table = _table()
        layout.addWidget(self.branch_compare_chart, stretch=2)
        layout.addWidget(self.branch_compare_table, stretch=1)
        self.tabs.addTab(tab, "So sánh chi nhánh")

    def _build_debt_quality_tab(self) -> None:
        tab = QWidget()
        layout = QVBoxLayout(tab)
        self.debt_quality_metrics = MetricGrid()
        self.debt_quality_metrics.set_empty(
            (
                "Tổng dư nợ cán bộ",
                "Nợ cần chú ý",
                "Nợ xấu",
                "Tỷ lệ nhóm 2",
                "Tỷ lệ nợ xấu",
                "Số KH có nhóm 2",
                "Số KH có nợ xấu",
            )
        )
        self.debt_quality_chart = AnalysisLineChart()
        self.debt_quality_table = _table()
        layout.addWidget(self.debt_quality_metrics)
        layout.addWidget(self.debt_quality_chart, stretch=2)
        layout.addWidget(self.debt_quality_table, stretch=1)
        self.tabs.addTab(tab, "Chất lượng tín dụng")

    def reload(self) -> None:
        self._reload_filter_options()
        filters = self._filters()
        self.overview = build_officer_overview(
            self.repository,
            self.data_type,
            officer_code=self.officer_code,
            officer=self.officer,
            branch=self.branch,
            transaction_office=filters.transaction_office,
            customer_type=filters.customer_type,
            period_from=filters.period_from,
            period_to=filters.period_to,
        )
        self.growth_rows = build_officer_growth_history(self.overview.points)
        self.apply_officer_comparison(show_limit_warning=False)
        self.reload_branch_comparison()
        self.reload_debt_quality()
        self._render_info()
        self._render_current_tab()

    def apply_filters(self) -> None:
        self.reload()

    def clear_filters(self) -> None:
        for combo in (self.period_from_combo, self.period_to_combo, self.customer_type_combo, self.transaction_office_combo):
            combo.setCurrentIndex(0)
        self.reload()

    def apply_officer_comparison(self, *, show_limit_warning: bool = True) -> None:
        _ = show_limit_warning
        self.officer_selector.force_hide_popup()
        selected = self.officer_selector.selected_officers()
        if not selected:
            selected = [officer_key(self.officer, self.branch, self.transaction_office)]
        self.compare_selected_officers = tuple(selected)
        self.compare_chart_page = 0
        metric = str(self.compare_metric_combo.currentData() or METRIC_NIM_AFTER)
        self.compare_rows = build_multiple_officer_comparison(
            self.repository,
            self.data_type,
            officers=selected,
            metric=metric,
            branch=self.branch,
            filters=self._filters(),
        )
        self._refresh_compare_table_filters()
        self._render_officer_compare()

    def clear_officer_comparison(self) -> None:
        self.officer_selector.clear_selection()
        self.apply_officer_comparison(show_limit_warning=False)

    def previous_compare_chart_page(self) -> None:
        self.compare_chart_page = max(0, self.compare_chart_page - 1)
        self._render_officer_compare_chart()

    def next_compare_chart_page(self) -> None:
        self.compare_chart_page = min(self._compare_total_pages() - 1, self.compare_chart_page + 1)
        self._render_officer_compare_chart()

    def _compare_display_changed(self) -> None:
        self.compare_chart_page = 0
        self._render_officer_compare()

    def _compare_page_size_changed(self) -> None:
        self.compare_chart_page = 0
        self._render_officer_compare_chart()

    def reload_branch_comparison(self) -> None:
        metric = str(self.branch_metric_combo.currentData() or METRIC_NIM_AFTER)
        self.branch_compare_series, self.branch_compare_rows = build_officer_branch_comparison(
            self.repository,
            self.data_type,
            officer_code=self.officer_code,
            officer=self.officer,
            branch=self.branch,
            metric=metric,
            filters=self._filters(),
        )
        self._render_branch_compare()

    def reload_debt_quality(self) -> None:
        history_filters = self._filters()
        customer_filters = CustomerFilters(
            period_from=history_filters.period_from,
            period_to=history_filters.period_to,
            customer_type=history_filters.customer_type,
            officer=self.officer_code or self.officer,
        )
        self.debt_quality_rows = tuple(
            self.customer_repository.get_officer_debt_group_history(
                officer_code=self.officer_code,
                officer_name=self.officer,
                filters=customer_filters,
            )
        )
        self._render_debt_quality()

    def export_excel(self) -> None:
        tab_key, rows = self._current_export_rows()
        path, _ = QFileDialog.getSaveFileName(
            self,
            "Xuất dữ liệu phân tích",
            f"{tab_key}.xlsx",
            "Excel (*.xlsx)",
        )
        if not path:
            return
        try:
            output = export_analysis_rows(rows, Path(path), tab_key=tab_key, metadata=self._export_metadata(tab_key))
        except Exception as exc:
            QMessageBox.warning(self, "Xuất Excel", str(exc))
            return
        QMessageBox.information(self, "Xuất Excel", f"Đã xuất: {output}")

    def _reload_filter_options(self) -> None:
        periods = self.history_repository.get_periods(
            self.data_type,
            officer_code=self.officer_code,
            officer=self.officer,
            branch=self.branch,
        )
        customer_types = self.history_repository.get_customer_types(
            self.data_type,
            officer_code=self.officer_code,
            officer=self.officer,
            branch=self.branch,
        )
        officers = self.history_repository.get_available_officers(self.data_type, branch=self.branch)
        self._populate_combo_preserve(self.period_from_combo, periods)
        self._populate_combo_preserve(self.period_to_combo, periods)
        self._populate_combo_preserve(self.customer_type_combo, customer_types)
        transaction_offices = sorted({officer.transaction_office for officer in officers if officer.transaction_office})
        if self.transaction_office and self.transaction_office not in transaction_offices:
            transaction_offices.insert(0, self.transaction_office)
        self._populate_combo_preserve(self.transaction_office_combo, transaction_offices)
        selected_codes = {self.officer_code} if self.officer_code else set()
        self.officer_selector.set_officers(officers, selected_codes=selected_codes, selected_raw=self.officer)

    def _populate_combo_preserve(self, combo: QComboBox, values: list[str]) -> None:
        current = combo.currentData()
        label = combo.itemText(0) if combo.count() else "Tất cả"
        combo.blockSignals(True)
        combo.clear()
        combo.addItem(label, "")
        for value in values:
            combo.addItem(value, value)
        index = combo.findData(current)
        combo.setCurrentIndex(index if index >= 0 else 0)
        configure_combo_popup_width(combo, minimum_popup_width=max(220, combo.minimumWidth()))
        combo.blockSignals(False)

    def _filters(self) -> HistoryFilters:
        return HistoryFilters(
            period_from=str(self.period_from_combo.currentData() or ""),
            period_to=str(self.period_to_combo.currentData() or ""),
            customer_type=str(self.customer_type_combo.currentData() or ""),
            transaction_office=str(self.transaction_office_combo.currentData() or ""),
        )

    def _render_info(self) -> None:
        history = self.overview
        if history is None:
            return
        metrics = [
            KpiMetric(self.ui_config.officer_short_label, history.officer.display_name, "text"),
            KpiMetric("Chi nhánh", history.branch or "", "text"),
            KpiMetric("Phòng GD", history.transaction_office or "Tất cả", "text"),
            KpiMetric("Loại khách hàng", history.customer_type or "Tất cả", "text"),
            KpiMetric(self.ui_config.current_balance_label, history.current_balance, "money"),
        ]
        if self.ui_config.include_average_rate:
            metrics.append(KpiMetric("Lãi suất bình quân hiện tại", history.current_average_rate, "percent"))
        metrics.extend(
            [
                KpiMetric("NIM trước ĐC hiện tại", history.current_nim_before, "percent"),
                KpiMetric("NIM sau ĐC hiện tại", history.current_nim_after, "percent"),
            ]
        )
        self.info_metrics.set_metrics(metrics)

    def _render_current_tab(self) -> None:
        index = self.tabs.currentIndex()
        if index == 0:
            self._render_overview()
        elif index == 1:
            self._render_growth()
        elif index == 2:
            self._render_officer_compare()
        elif index == 3:
            self._render_branch_compare()
        elif index == 4:
            self._render_debt_quality()

    def _render_overview(self) -> None:
        points = self.overview.points if self.overview else ()
        self.overview_chart.set_series(
            overview_series(points, self.ui_config),
            empty_message="Không có dữ liệu lịch sử.",
            single_point_message="Chưa đủ dữ liệu lịch sử để đánh giá xu hướng.",
        )
        headers = ["Kỳ", self.ui_config.balance_label]
        widths = [90, 150]
        if self.ui_config.include_average_rate:
            headers.append("Lãi suất bình quân")
            widths.append(130)
        headers.extend(["NIM trước ĐC", "NIM sau ĐC"])
        widths.extend([120, 120])
        rows = []
        for point in points:
            payload = [
                point.period,
                (point.balance, format_money_vn(point.balance)),
            ]
            if self.ui_config.include_average_rate:
                payload.append((point.average_rate, format_percent_vn(point.average_rate)))
            payload.extend(
                [
                    (point.nim_before, format_percent_vn(point.nim_before)),
                    (point.nim_after, format_percent_vn(point.nim_after)),
                ]
            )
            rows.append(tuple(payload))
        _render_table(
            self.overview_table,
            tuple(headers),
            rows,
            tuple(widths),
        )

    def _render_growth(self) -> None:
        points = self.overview.points if self.overview else ()
        self.balance_chart.set_series(balance_series(points, self.ui_config), empty_message=f"Không có dữ liệu {self.ui_config.balance_label.lower()}.")
        self.growth_chart.set_series(growth_series(self.growth_rows, self.ui_config), empty_message="Không có dữ liệu tăng trưởng.")
        _render_table(
            self.growth_table,
            ("Kỳ", self.ui_config.balance_label, self.ui_config.balance_delta_label, self.ui_config.growth_percent_label, "NIM trước ĐC", "NIM sau ĐC"),
            [
                (
                    row.period,
                    (row.balance, format_money_vn(row.balance)),
                    (row.delta, "" if row.delta is None else format_money_vn(row.delta, signed=True)),
                    (row.growth_percent, format_percent_vn(row.growth_percent, signed=True)),
                    (row.nim_before, format_percent_vn(row.nim_before)),
                    (row.nim_after, format_percent_vn(row.nim_after)),
                )
                for row in self.growth_rows
            ],
            (90, 150, 150, 130, 120, 120),
        )

    def _render_officer_compare(self) -> None:
        self._render_officer_compare_chart()
        self._render_officer_compare_table()

    def _render_officer_compare_chart(self) -> None:
        metric = str(self.compare_metric_combo.currentData() or METRIC_NIM_AFTER)
        labels = metric_labels(self.data_type)
        mode = self._compare_mode()
        self.officer_compare_chart.setVisible(mode != COMPARE_MODE_TABLE)
        chart_rows = self._compare_chart_rows()
        self.officer_compare_chart.set_series(
            comparison_series(chart_rows, metric, metric_label=labels.get(metric, metric)),
            empty_message="Chưa có dữ liệu so sánh cán bộ.",
        )
        self._update_compare_chart_controls()

    def _render_officer_compare_table(self) -> None:
        metric = str(self.compare_metric_combo.currentData() or METRIC_NIM_AFTER)
        labels = metric_labels(self.data_type)
        rows = self._filtered_compare_rows()
        _render_table(
            self.officer_compare_table,
            ("Kỳ", self.ui_config.officer_short_label, "Chi nhánh", "Phòng GD", "Loại KH", "Chỉ tiêu", "Giá trị"),
            [
                (
                    row.period,
                    row.officer.display_name,
                    row.branch,
                    row.transaction_office,
                    row.customer_type,
                    labels.get(row.metric, row.metric),
                    (row.value, _format_metric_value(row.value, row.metric)),
                )
                for row in rows
            ],
            (90, 160, 170, 120, 100, 150, 120),
        )

    def _refresh_compare_table_filters(self) -> None:
        if not hasattr(self, "compare_table_period_combo"):
            return
        self._populate_combo_preserve(self.compare_table_period_combo, sorted({row.period for row in self.compare_rows if row.period}))
        self._populate_combo_preserve(self.compare_table_branch_combo, sorted({row.branch for row in self.compare_rows if row.branch}))

    def _filtered_compare_rows(self) -> tuple[ComparisonRow, ...]:
        rows = self.compare_rows
        if hasattr(self, "compare_table_search"):
            needle = self.compare_table_search.text().strip().casefold()
            if needle:
                rows = tuple(
                    row
                    for row in rows
                    if needle in row.officer.display_name.casefold()
                    or needle in row.branch.casefold()
                    or needle in row.transaction_office.casefold()
                    or needle in row.period.casefold()
                )
            period = str(self.compare_table_period_combo.currentData() or "")
            if period:
                rows = tuple(row for row in rows if row.period == period)
            branch = str(self.compare_table_branch_combo.currentData() or "")
            if branch:
                rows = tuple(row for row in rows if row.branch == branch)
        return rows

    def _compare_chart_rows(self) -> tuple[ComparisonRow, ...]:
        if self._compare_mode() == COMPARE_MODE_TABLE:
            return ()
        keys = self._compare_chart_officer_keys()
        rows_by_key: dict[str, list[ComparisonRow]] = {key: [] for key in keys}
        for row in self.compare_rows:
            key = _comparison_key(row)
            if key in rows_by_key:
                rows_by_key[key].append(row)
        return tuple(row for key in keys for row in rows_by_key[key])

    def _compare_chart_officer_keys(self) -> tuple[str, ...]:
        keys = self._compare_officer_keys()
        if self._compare_mode() == COMPARE_MODE_TOP:
            return self._top_compare_officer_keys(keys)
        if self._compare_mode() == COMPARE_MODE_ALL:
            return keys
        page_size = self._compare_page_size()
        total_pages = max(1, (len(keys) + page_size - 1) // page_size)
        self.compare_chart_page = min(max(0, self.compare_chart_page), total_pages - 1)
        start = self.compare_chart_page * page_size
        return keys[start:start + page_size]

    def _top_compare_officer_keys(self, keys: tuple[str, ...]) -> tuple[str, ...]:
        latest: dict[str, tuple[str, float | None]] = {}
        for row in self.compare_rows:
            key = _comparison_key(row)
            current = latest.get(key)
            if current is None or row.period >= current[0]:
                latest[key] = (row.period, row.value)
        ranked = sorted(
            keys,
            key=lambda key: float("-inf") if latest.get(key, ("", None))[1] is None else float(latest[key][1] or 0),
            reverse=True,
        )
        return tuple(ranked[:self._compare_top_n()])

    def _compare_officer_keys(self) -> tuple[str, ...]:
        if self.compare_selected_officers:
            selected = tuple(_officer_key_identity(officer) for officer in self.compare_selected_officers)
            available = {_comparison_key(row) for row in self.compare_rows}
            return tuple(key for key in selected if key in available)
        keys: list[str] = []
        seen: set[str] = set()
        for row in self.compare_rows:
            key = _comparison_key(row)
            if key not in seen:
                seen.add(key)
                keys.append(key)
        return tuple(keys)

    def _compare_total_pages(self) -> int:
        if self._compare_mode() != COMPARE_MODE_PAGED:
            return 1
        count = len(self._compare_officer_keys())
        return max(1, (count + self._compare_page_size() - 1) // self._compare_page_size())

    def _update_compare_chart_controls(self) -> None:
        mode = self._compare_mode()
        total_pages = self._compare_total_pages()
        self.compare_chart_page = min(max(0, self.compare_chart_page), total_pages - 1)
        is_paged = mode == COMPARE_MODE_PAGED
        self.compare_page_size_combo.setEnabled(is_paged)
        self.compare_top_combo.setEnabled(mode == COMPARE_MODE_TOP)
        self.compare_prev_button.setEnabled(is_paged and self.compare_chart_page > 0)
        self.compare_next_button.setEnabled(is_paged and self.compare_chart_page < total_pages - 1)
        self.compare_page_label.setText(f"Trang {self.compare_chart_page + 1}/{total_pages}")
        selected_count = len(self._compare_officer_keys())
        visible_count = len(self._compare_chart_officer_keys()) if mode != COMPARE_MODE_TABLE else 0
        if mode == COMPARE_MODE_TABLE:
            message = f"Đã chọn {selected_count} cán bộ. Đang ưu tiên bảng dữ liệu."
        elif mode == COMPARE_MODE_ALL and selected_count > self._compare_page_size():
            message = f"Đã chọn {selected_count} cán bộ. Biểu đồ có nhiều series và có thể khó đọc."
        elif is_paged and selected_count > self._compare_page_size():
            message = f"Đã chọn {selected_count} cán bộ. Biểu đồ đang hiển thị {visible_count} cán bộ theo từng trang."
        elif mode == COMPARE_MODE_TOP:
            message = f"Đã chọn {selected_count} cán bộ. Biểu đồ đang hiển thị Top {visible_count}."
        else:
            message = f"Đã chọn {selected_count} cán bộ."
        self.compare_status_label.setText(message)

    def _compare_mode(self) -> str:
        return str(self.compare_mode_combo.currentData() or COMPARE_MODE_PAGED)

    def _compare_page_size(self) -> int:
        return int(self.compare_page_size_combo.currentData() or COMPARE_DEFAULT_PAGE_SIZE)

    def _compare_top_n(self) -> int:
        return int(self.compare_top_combo.currentData() or 10)

    def _compare_export_scope(self) -> str:
        return str(self.compare_export_scope_combo.currentData() or COMPARE_EXPORT_ALL)

    def _render_branch_compare(self) -> None:
        metric = str(self.branch_metric_combo.currentData() or METRIC_NIM_AFTER)
        labels = metric_labels(self.data_type)
        self.branch_compare_chart.set_series(self.branch_compare_series, empty_message="Chưa có dữ liệu so sánh chi nhánh.")
        _render_table(
            self.branch_compare_table,
            ("Kỳ", "Đối tượng", "Chi nhánh", "Phòng GD", "Loại KH", "Chỉ tiêu", "Giá trị"),
            [
                (
                    row.period,
                    row.officer.display_name,
                    row.branch,
                    row.transaction_office,
                    row.customer_type,
                    labels.get(row.metric, row.metric),
                    (row.value, _format_metric_value(row.value, metric)),
                )
                for row in self.branch_compare_rows
            ],
            (90, 190, 180, 120, 100, 150, 120),
        )

    def _render_debt_quality(self) -> None:
        rows = list(self.debt_quality_rows)
        current = rows[-1] if rows else {}
        self.debt_quality_metrics.set_metrics(
            [
                KpiMetric("Tổng dư nợ cán bộ", current.get("total_balance", 0), "money"),
                KpiMetric("Nợ cần chú ý", current.get("attention_balance", 0), "money"),
                KpiMetric("Nợ xấu", current.get("bad_debt_balance", 0), "money"),
                KpiMetric("Tỷ lệ nhóm 2", current.get("attention_ratio"), "percent"),
                KpiMetric("Tỷ lệ nợ xấu", current.get("bad_debt_ratio"), "percent"),
                KpiMetric("Số KH có nhóm 2", current.get("attention_customer_count", 0), "count"),
                KpiMetric("Số KH có nợ xấu", current.get("bad_debt_customer_count", 0), "count"),
            ]
        )
        self.debt_quality_chart.set_series(
            (
                ChartSeries(
                    "Tỷ lệ nhóm 2",
                    tuple((str(row.get("period") or ""), row.get("attention_ratio")) for row in rows),
                    "percent",
                ),
                ChartSeries(
                    "Tỷ lệ nợ xấu",
                    tuple((str(row.get("period") or ""), row.get("bad_debt_ratio")) for row in rows),
                    "percent",
                ),
            ),
            empty_message="Chưa có dữ liệu nhóm nợ cho cán bộ này.",
            single_point_message="Chưa đủ dữ liệu lịch sử để hiển thị xu hướng.",
        )
        _render_table(
            self.debt_quality_table,
            ("Kỳ", "Tổng dư nợ", "Nợ cần chú ý", "Nợ xấu", "Tỷ lệ nhóm 2", "Tỷ lệ nợ xấu", "Lãi suất bình quân", "NIM trước ĐC", "NIM sau ĐC"),
            [
                (
                    row.get("period", ""),
                    (row.get("total_balance"), format_money_vn(row.get("total_balance"))),
                    (row.get("attention_balance"), format_money_vn(row.get("attention_balance"))),
                    (row.get("bad_debt_balance"), format_money_vn(row.get("bad_debt_balance"))),
                    (row.get("attention_ratio"), format_percent_vn(row.get("attention_ratio"))),
                    (row.get("bad_debt_ratio"), format_percent_vn(row.get("bad_debt_ratio"))),
                    (row.get("average_rate"), format_percent_vn(row.get("average_rate"))),
                    (row.get("nim_before"), format_percent_vn(row.get("nim_before"))),
                    (row.get("nim_after"), format_percent_vn(row.get("nim_after"))),
                )
                for row in rows
            ],
            (90, 140, 140, 140, 115, 115, 130, 120, 120),
        )

    def _current_export_rows(self) -> tuple[str, list[dict[str, object]]]:
        index = self.tabs.currentIndex()
        if index == 1:
            return "growth", [
                {
                    "Kỳ": row.period,
                    self.ui_config.balance_label: row.balance,
                    self.ui_config.balance_delta_label: row.delta,
                    self.ui_config.growth_percent_label: row.growth_percent,
                    "NIM trước ĐC": row.nim_before,
                    "NIM sau ĐC": row.nim_after,
                }
                for row in self.growth_rows
            ]
        if index == 2:
            rows = self._compare_chart_rows() if self._compare_export_scope() == COMPARE_EXPORT_VISIBLE else self.compare_rows
            return "officer_compare", _comparison_export_rows(rows, self.ui_config)
        if index == 3:
            return "branch_compare", _comparison_export_rows(self.branch_compare_rows, self.ui_config)
        if index == 4:
            return "debt_quality", [
                {
                    "Kỳ": row.get("period", ""),
                    "Tổng dư nợ": row.get("total_balance", 0),
                    "Nợ cần chú ý": row.get("attention_balance", 0),
                    "Nợ xấu": row.get("bad_debt_balance", 0),
                    "Tỷ lệ nhóm 2": row.get("attention_ratio"),
                    "Tỷ lệ nợ xấu": row.get("bad_debt_ratio"),
                    "Lãi suất bình quân": row.get("average_rate"),
                    "NIM trước ĐC": row.get("nim_before"),
                    "NIM sau ĐC": row.get("nim_after"),
                }
                for row in self.debt_quality_rows
            ]
        points = self.overview.points if self.overview else ()
        output = []
        for point in points:
            payload = {
                "Kỳ": point.period,
                self.ui_config.balance_label: point.balance,
                "NIM trước ĐC": point.nim_before,
                "NIM sau ĐC": point.nim_after,
            }
            if self.ui_config.include_average_rate:
                payload["Lãi suất bình quân"] = point.average_rate
            output.append(payload)
        return "overview", [
            row for row in output
        ]

    def _export_metadata(self, tab_key: str) -> dict[str, object] | None:
        if tab_key != "officer_compare":
            return None
        metric = str(self.compare_metric_combo.currentData() or METRIC_NIM_AFTER)
        labels = metric_labels(self.data_type)
        periods = sorted({row.period for row in self.compare_rows if row.period})
        filters = self._filters()
        return {
            "Chỉ tiêu": labels.get(metric, metric),
            "Số cán bộ đã chọn": len(self._compare_officer_keys()),
            "Các kỳ": ", ".join(periods),
            "Từ kỳ": filters.period_from or "Tất cả",
            "Đến kỳ": filters.period_to or "Tất cả",
            "Loại KH": filters.customer_type or "Tất cả",
            "Phòng GD": filters.transaction_office or "Tất cả",
            "Phạm vi xuất": "Đang hiển thị" if self._compare_export_scope() == COMPARE_EXPORT_VISIBLE else "Toàn bộ cán bộ đã chọn",
        }


def _table() -> QTableWidget:
    table = FitTableWidget()
    table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
    table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
    table.setSortingEnabled(True)
    return table


def _combo(placeholder: str) -> QComboBox:
    combo = shared_combo_box("Tất cả", minimum_width=120, maximum_width=185)
    combo.setToolTip(placeholder)
    return combo


def _metric_combo(ui_config: NimUiConfig, *, include_growth: bool, default_metric: str = METRIC_BALANCE) -> QComboBox:
    combo = shared_combo_box("Chỉ tiêu", minimum_width=145, maximum_width=220)
    combo.clear()
    labels = ui_config.metric_labels()
    for metric in ui_config.metric_order(include_growth=include_growth):
        combo.addItem(labels[metric], metric)
    index = combo.findData(default_metric)
    combo.setCurrentIndex(index if index >= 0 else 0)
    configure_combo_popup_width(combo, minimum_popup_width=240)
    return combo


def _render_table(table: QTableWidget, headers: tuple[str, ...], rows: list[tuple[object, ...]], widths: tuple[int, ...]) -> None:
    table.setSortingEnabled(False)
    table.clear()
    table.setColumnCount(len(headers))
    table.setHorizontalHeaderLabels(headers)
    table.setRowCount(len(rows))
    for row_index, row in enumerate(rows):
        for col_index, value in enumerate(row):
            if isinstance(value, tuple):
                number, text = value
                item = NumericTableWidgetItem(text, number)
                item.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
            else:
                item = QTableWidgetItem(str(value or ""))
                item.setTextAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
            table.setItem(row_index, col_index, item)
    if isinstance(table, FitTableWidget):
        table.set_default_widths(widths)
    else:
        for index, width in enumerate(widths):
            table.setColumnWidth(index, width)
    table.setSortingEnabled(True)


def _format_metric_value(value: float | None, metric: str) -> str:
    if value is None:
        return "N/A"
    if metric == METRIC_BALANCE:
        return format_money_vn(value)
    if metric == METRIC_BALANCE_GROWTH:
        return format_percent_vn(value, signed=True)
    return format_percent_vn(value)


def _comparison_key(row: ComparisonRow) -> str:
    return row.officer.code or row.officer.raw_name or row.officer.display_name


def _officer_key_identity(officer) -> str:
    return getattr(officer, "code", "") or getattr(officer, "raw_name", "") or getattr(officer, "display_name", "")


def _comparison_export_rows(rows: tuple[ComparisonRow, ...], ui_config: NimUiConfig) -> list[dict[str, object]]:
    labels = ui_config.metric_labels()
    return [
        {
            "Kỳ": row.period,
            ui_config.officer_short_label: row.officer.display_name,
            "Chi nhánh": row.branch,
            "Phòng GD": row.transaction_office,
            "Loại KH": row.customer_type,
            "Chỉ tiêu": labels.get(row.metric, row.metric),
            "Giá trị": row.value,
        }
        for row in rows
    ]
