from __future__ import annotations

from pathlib import Path
from typing import Callable

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QButtonGroup,
    QCheckBox,
    QComboBox,
    QFileDialog,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QRadioButton,
    QSizePolicy,
    QWidget,
    QVBoxLayout,
)

from agribank_v3.features.credit.summary.credit_report import VIEW_COMPARE_PERIODS, VIEW_CURRENT_PERIOD
from agribank_v3.features.credit.summary.customer.table_models import CustomerTableModel
from agribank_v3.features.credit.summary.customer.widgets import CustomerTableView, Pager
from agribank_v3.features.credit.summary.group_lending.detail_window import GroupLendingDetailWindow
from agribank_v3.features.credit.summary.group_lending.export_service import GroupLendingExportService
from agribank_v3.features.credit.summary.group_lending.models import (
    DETAIL_BY_GROUP,
    SUMMARY_BY_ASSOCIATION,
    GroupLendingFilters,
    GroupLendingKpi,
    GroupLendingResult,
    GroupLendingRow,
)
from agribank_v3.features.credit.summary.group_lending.service import GroupLendingService
from agribank_v3.features.credit.tovayvon.placeholder_windows import CreditGroupManagementPlaceholderDialog
from agribank_v3.ui.components.controls import combo_box, current_data, populate_combo, primary_button, secondary_button
from agribank_v3.ui.components.flow_layout import FlowLayout
from agribank_v3.ui.components.kpi import KpiMetric, MetricGrid
from agribank_v3.ui.workers import run_in_thread


GROUP_DETAIL_COLUMNS = (
    ("STT", "STT", "integer"),
    ("Kỳ", "Kỳ", "text"),
    ("Mã tổ", "Mã tổ", "text"),
    ("Tên tổ", "Tên tổ", "text"),
    ("Loại tổ chức Hội", "Loại tổ chức Hội", "text"),
    ("Tên tổ chức khác", "Tên tổ chức khác", "text"),
    ("Chi nhánh", "Chi nhánh", "text"),
    ("Hội sở/PGD", "Hội sở/PGD", "text"),
    ("Xã", "Xã", "text"),
    ("Tổ trưởng", "Tổ trưởng", "text"),
    ("Số tổ viên còn dư nợ", "Số tổ viên còn dư nợ", "integer"),
    ("Số món", "Số món", "integer"),
    ("Tổng dư nợ", "Tổng dư nợ", "money"),
    ("Dư nợ bình quân/tổ viên", "Dư nợ bình quân/tổ viên", "money_or_blank"),
    ("Trạng thái danh mục", "Trạng thái danh mục", "text"),
)
ASSOCIATION_COLUMNS = (
    ("Loại tổ chức Hội", "Loại tổ chức Hội", "text"),
    ("Số tổ có dư nợ", "Số tổ có dư nợ", "integer"),
    ("Số tổ viên duy nhất", "Số tổ viên duy nhất", "integer"),
    ("Tổng lượt tổ viên theo tổ", "Tổng lượt tổ viên theo tổ", "integer"),
    ("Tổng dư nợ", "Tổng dư nợ", "money"),
    ("Tỷ trọng", "Tỷ trọng", "percent_or_blank"),
    ("Dư nợ bình quân/tổ", "Dư nợ bình quân/tổ", "money_or_blank"),
    ("Dư nợ bình quân/tổ viên", "Dư nợ bình quân/tổ viên", "money_or_blank"),
)
GROUP_COMPARE_COLUMNS = (
    ("Mã tổ", "Mã tổ", "text"),
    ("Tên tổ", "Tên tổ", "text"),
    ("Loại Hội", "Loại Hội", "text"),
    ("Tổ viên Từ kỳ", "Tổ viên Từ kỳ", "integer"),
    ("Tổ viên Đến kỳ", "Tổ viên Đến kỳ", "integer"),
    ("Tăng/giảm tổ viên", "Tăng/giảm tổ viên", "integer"),
    ("Dư nợ Từ kỳ", "Dư nợ Từ kỳ", "money"),
    ("Dư nợ Đến kỳ", "Dư nợ Đến kỳ", "money"),
    ("Tăng/giảm dư nợ", "Tăng/giảm dư nợ", "money_signed"),
    ("Tăng trưởng dư nợ (%)", "Tăng trưởng dư nợ (%)", "percent_signed"),
    ("Phân loại biến động", "Phân loại biến động", "text"),
)
ASSOCIATION_COMPARE_COLUMNS = (
    ("Loại Hội", "Loại Hội", "text"),
    ("Số tổ Từ kỳ", "Số tổ Từ kỳ", "integer"),
    ("Số tổ Đến kỳ", "Số tổ Đến kỳ", "integer"),
    ("Thay đổi số tổ", "Thay đổi số tổ", "integer"),
    ("Tổ viên Từ kỳ", "Tổ viên Từ kỳ", "integer"),
    ("Tổ viên Đến kỳ", "Tổ viên Đến kỳ", "integer"),
    ("Thay đổi tổ viên", "Thay đổi tổ viên", "integer"),
    ("Dư nợ Từ kỳ", "Dư nợ Từ kỳ", "money"),
    ("Dư nợ Đến kỳ", "Dư nợ Đến kỳ", "money"),
    ("Tăng/giảm dư nợ", "Tăng/giảm dư nợ", "money_signed"),
    ("Tăng trưởng (%)", "Tăng trưởng (%)", "percent_signed"),
    ("Tỷ trọng Từ kỳ", "Tỷ trọng Từ kỳ", "percent_or_blank"),
    ("Tỷ trọng Đến kỳ", "Tỷ trọng Đến kỳ", "percent_or_blank"),
    ("Thay đổi tỷ trọng (điểm %)", "Thay đổi tỷ trọng (điểm %)", "percent_point_signed"),
)


class GroupLendingTab(QWidget):
    def __init__(
        self,
        main_database_path: Path | None,
        *,
        period_provider: Callable[[], tuple[str, str, str]],
        view_mode_provider: Callable[[], str],
        parent: QWidget | None = None,
    ) -> None:
        super().__init__(parent)
        self.service = GroupLendingService(main_database_path)
        self.period_provider = period_provider
        self.view_mode_provider = view_mode_provider
        self.page = 1
        self.page_size = 100
        self._generation = 0
        self._current_result: GroupLendingResult | None = None
        self._current_group_rows: dict[str, GroupLendingRow] = {}
        self._build_ui()
        self.reload_filter_values()
        self.refresh()

    def _build_ui(self) -> None:
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(8)

        option_area = QWidget()
        option_flow = FlowLayout(option_area, spacing=8)
        self.option_button_group = QButtonGroup(self)
        self.option_button_group.setExclusive(True)
        self.detail_radio = QRadioButton("Chi tiết theo tổ")
        self.association_radio = QRadioButton("Tổng hợp theo Hội")
        self.detail_radio.setProperty("group_lending_mode", DETAIL_BY_GROUP)
        self.association_radio.setProperty("group_lending_mode", SUMMARY_BY_ASSOCIATION)
        self.detail_radio.setChecked(True)
        for button in (self.detail_radio, self.association_radio):
            self.option_button_group.addButton(button)
            option_flow.addWidget(button)
        self.mode_label = QLabel("")
        self.mode_label.setObjectName("MutedText")
        option_flow.addWidget(self.mode_label)
        layout.addWidget(option_area)

        filter_area = QWidget()
        filter_area.setObjectName("GroupLendingFilterArea")
        filter_flow = FlowLayout(filter_area, spacing=7)
        self.branch_combo = _group_combo("Chi nhánh")
        self.office_combo = _group_combo("Phòng giao dịch")
        self.association_combo = _group_combo("Loại tổ chức Hội")
        self.status_combo = _group_combo("Trạng thái tổ")
        self.officer_combo = _group_combo("CBTD", maximum_width=260)
        self.search_edit = QLineEdit()
        self.search_edit.setObjectName("AgribankSearchBox")
        self.search_edit.setPlaceholderText("Tìm mã hoặc tên tổ")
        self.search_edit.setClearButtonEnabled(True)
        self.search_edit.setMinimumWidth(280)
        self.search_edit.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        self.include_unknown_check = QCheckBox("Hiển thị tổ chưa khai báo")
        self.include_unknown_check.setChecked(True)
        self.refresh_button = secondary_button("Làm mới")
        self.clear_button = secondary_button("Xóa lọc")
        self.export_button = secondary_button("Xuất Excel")
        self.open_group_manager_button = primary_button("Mở Quản lý tổ vay vốn")
        for label, widget in (
            ("Chi nhánh", self.branch_combo),
            ("Phòng giao dịch", self.office_combo),
            ("Loại tổ chức Hội", self.association_combo),
            ("Trạng thái tổ", self.status_combo),
            ("CBTD", self.officer_combo),
            ("Tìm kiếm", self.search_edit),
        ):
            filter_flow.addWidget(_labeled(label, widget))
        for widget in (
            self.include_unknown_check,
            self.refresh_button,
            self.clear_button,
            self.export_button,
            self.open_group_manager_button,
        ):
            filter_flow.addWidget(widget)
        layout.addWidget(filter_area)

        self.metrics = MetricGrid()
        layout.addWidget(self.metrics)
        self.note_label = QLabel("")
        self.note_label.setObjectName("MutedText")
        self.note_label.setWordWrap(True)
        layout.addWidget(self.note_label)

        self.model = CustomerTableModel(GROUP_DETAIL_COLUMNS, self)
        self.table = CustomerTableView(self)
        self.table.setModel(self.model)
        self.table.apply_default_widths((56, 78, 128, 180, 150, 150, 90, 120, 110, 130, 120, 90, 140, 150, 120))
        self.table.doubleClicked.connect(self._open_selected_group_detail)
        layout.addWidget(self.table, stretch=1)
        self.pager = Pager(self)
        self.pager.pageChanged.connect(self._page_changed)
        self.pager.pageSizeChanged.connect(self._page_size_changed)
        layout.addWidget(self.pager)

        self.option_button_group.buttonToggled.connect(self._option_toggled)
        for combo in (self.branch_combo, self.office_combo, self.association_combo, self.status_combo, self.officer_combo):
            combo.currentIndexChanged.connect(lambda _index: self._filters_changed())
        self.search_edit.textChanged.connect(lambda _text: self._filters_changed())
        self.include_unknown_check.toggled.connect(lambda _checked: self._filters_changed())
        self.refresh_button.clicked.connect(self.refresh)
        self.clear_button.clicked.connect(self.clear_filters)
        self.export_button.clicked.connect(self.export_excel)
        self.open_group_manager_button.clicked.connect(self.open_group_manager)

    def reload_filter_values(self) -> None:
        values = self.service.filter_values()
        for combo, key in (
            (self.branch_combo, "branches"),
            (self.office_combo, "offices"),
            (self.association_combo, "association_types"),
            (self.status_combo, "statuses"),
            (self.officer_combo, "officers"),
        ):
            current = current_data(combo)
            populate_combo(combo, values.get(key, []))
            if current:
                _select_data(combo, current)

    def refresh(self) -> None:
        self._generation += 1
        generation = self._generation
        filters = self._filters()
        view_mode = self.view_mode_provider()
        display_mode = self._display_mode()
        self._update_mode_label(view_mode, filters)
        if self.isVisible():
            self.metrics.set_loading(("Số tổ có dư nợ", "Tổng số tổ viên duy nhất", "Tổng dư nợ cho vay qua tổ"))

            def done(result: GroupLendingResult) -> None:
                if generation != self._generation:
                    return
                self._render(result, view_mode, display_mode)

            def failed(exc: Exception) -> None:
                if generation != self._generation:
                    return
                self.metrics.set_empty(("Số tổ có dư nợ", "Tổng dư nợ cho vay qua tổ"))
                QMessageBox.warning(self, "Cho vay qua tổ", str(exc))

            run_in_thread(self, lambda _progress: self._load_result(filters, view_mode, display_mode), done, failed)
            return
        self._render(self._load_result(filters, view_mode, display_mode), view_mode, display_mode)

    def clear_filters(self) -> None:
        for combo in (self.branch_combo, self.office_combo, self.association_combo, self.status_combo, self.officer_combo):
            combo.setCurrentIndex(0)
        self.search_edit.clear()
        self.include_unknown_check.setChecked(True)
        self.page = 1
        self.refresh()

    def export_excel(self) -> None:
        path, _ = QFileDialog.getSaveFileName(self, "Xuất Cho vay qua tổ", "ChoVayQuaTo.xlsx", "Excel (*.xlsx)")
        if not path:
            return
        filters = self._filters()
        view_mode = self.view_mode_provider()
        display_mode = self._display_mode()
        try:
            output = GroupLendingExportService(self.service).export(
                Path(path),
                filters=filters,
                view_mode=view_mode,
                display_mode=display_mode,
            )
        except Exception as exc:
            QMessageBox.warning(self, "Xuất Cho vay qua tổ", str(exc))
            return
        QMessageBox.information(self, "Xuất Cho vay qua tổ", f"Đã xuất: {output}")

    def open_group_manager(self) -> None:
        CreditGroupManagementPlaceholderDialog(self).exec()
        self.reload_filter_values()
        self.refresh()

    def _load_result(
        self,
        filters: GroupLendingFilters,
        view_mode: str,
        display_mode: str,
    ) -> GroupLendingResult:
        if view_mode == VIEW_COMPARE_PERIODS:
            if display_mode == SUMMARY_BY_ASSOCIATION:
                return self.service.compare_associations(filters.from_period, filters.to_period, filters)
            return self.service.compare_groups(filters.from_period, filters.to_period, filters)
        if display_mode == SUMMARY_BY_ASSOCIATION:
            return self.service.get_association_summary(filters.period, filters)
        return self.service.get_group_lending_snapshot(filters.period, filters, page=self.page, page_size=self.page_size)

    def _render(self, result: GroupLendingResult, view_mode: str, display_mode: str) -> None:
        self._current_result = result
        self._current_group_rows = {
            row.group_code: row for row in result.rows if isinstance(row, GroupLendingRow)
        }
        self.metrics.set_metrics([_kpi_to_metric(item, compare=view_mode == VIEW_COMPARE_PERIODS) for item in result.kpis])
        rows = [_result_row_to_dict(row, index) for index, row in enumerate(result.rows, start=(result.page - 1) * result.page_size + 1)]
        self.model = CustomerTableModel(_columns_for(view_mode, display_mode), self)
        self.model.set_rows(rows)
        self.table.setModel(self.model)
        self.table.apply_default_widths(_widths_for(view_mode, display_mode))
        self.pager.setVisible(view_mode != VIEW_COMPARE_PERIODS and display_mode == DETAIL_BY_GROUP)
        self.pager.set_state(page=result.page, page_size=result.page_size, total_rows=result.total_rows)
        notes = list(result.notes)
        diagnostics = result.diagnostics or {}
        if diagnostics.get("multi_group_customer_count"):
            notes.append(
                "Có {count} khách hàng có dư nợ ở từ hai tổ trở lên trong kỳ; tổng dư nợ {balance:,.0f}.".format(
                    count=diagnostics.get("multi_group_customer_count"),
                    balance=float(diagnostics.get("multi_group_customer_balance") or 0),
                )
            )
        self.note_label.setText(" ".join(notes))

    def _filters(self) -> GroupLendingFilters:
        period, from_period, to_period = self.period_provider()
        return GroupLendingFilters(
            period=period,
            from_period=from_period,
            to_period=to_period,
            branch_code=current_data(self.branch_combo),
            office_code=current_data(self.office_combo),
            association_type=current_data(self.association_combo),
            group_status=current_data(self.status_combo),
            officer=current_data(self.officer_combo),
            search=self.search_edit.text(),
            include_unknown_groups=self.include_unknown_check.isChecked(),
        )

    def _display_mode(self) -> str:
        button = self.option_button_group.checkedButton()
        return str(button.property("group_lending_mode") if button is not None else DETAIL_BY_GROUP)

    def _option_toggled(self, button: QRadioButton, checked: bool) -> None:
        if checked:
            self.page = 1
            self.refresh()

    def _filters_changed(self) -> None:
        self.page = 1
        self.refresh()

    def _page_changed(self, page: int) -> None:
        self.page = int(page or 1)
        self.refresh()

    def _page_size_changed(self, page_size: int) -> None:
        self.page = 1
        self.page_size = int(page_size or 100)
        self.refresh()

    def _open_selected_group_detail(self, _index=None) -> None:
        if self.view_mode_provider() == VIEW_COMPARE_PERIODS or self._display_mode() != DETAIL_BY_GROUP:
            return
        index = self.table.currentIndex()
        if not index.isValid():
            return
        row = self.model.raw_row(index.row())
        group_code = str(row.get("Mã tổ") or "").strip()
        group_row = self._current_group_rows.get(group_code)
        if group_row is None:
            return
        dialog = GroupLendingDetailWindow(self.service, self._filters().period, group_row, self._filters(), self)
        dialog.exec()

    def _update_mode_label(self, view_mode: str, filters: GroupLendingFilters) -> None:
        if view_mode == VIEW_COMPARE_PERIODS:
            self.mode_label.setText(f"So sánh trực tiếp {filters.from_period or '...'} -> {filters.to_period or '...'}")
        else:
            self.mode_label.setText(f"Kỳ hiện tại: {filters.period or '...'}")


def _group_combo(label: str, *, maximum_width: int | None = None) -> QComboBox:
    widths = {
        "Chi nhánh": (168, 220),
        "Phòng giao dịch": (184, 250),
        "Loại tổ chức Hội": (200, 270),
        "Trạng thái tổ": (166, 210),
        "CBTD": (170, 280),
    }
    minimum_width, default_maximum = widths.get(label, (160, 220))
    return combo_box(
        label,
        minimum_width=minimum_width,
        maximum_width=maximum_width if maximum_width is not None else default_maximum,
        minimum_contents_length=max(10, len(label) + 3),
        searchable=True,
    )


def _labeled(label: str, widget: QWidget) -> QWidget:
    container = QWidget()
    container.setSizePolicy(widget.sizePolicy().horizontalPolicy(), QSizePolicy.Policy.Fixed)
    layout = QVBoxLayout(container)
    layout.setContentsMargins(0, 0, 0, 0)
    layout.setSpacing(2)
    caption = QLabel(label)
    caption.setObjectName("MutedText")
    layout.addWidget(caption)
    layout.addWidget(widget)
    return container


def _select_data(combo: QComboBox, value: str) -> None:
    index = combo.findData(value)
    if index >= 0:
        combo.setCurrentIndex(index)


def _kpi_to_metric(kpi: GroupLendingKpi, *, compare: bool) -> KpiMetric:
    if compare:
        value_type = "count" if kpi.kind == "count" else "money"
        tooltip = (
            f"{kpi.label}\n"
            f"Từ kỳ: {kpi.from_value}\n"
            f"Đến kỳ: {kpi.to_value}\n"
            f"Tăng/giảm: {kpi.difference}\n"
            f"Tăng trưởng: {'N/A' if kpi.growth_rate is None else kpi.growth_rate}"
        )
        return KpiMetric(kpi.label, kpi.difference, value_type, tooltip=tooltip, signed=True)
    return KpiMetric(kpi.label, kpi.value, "count" if kpi.kind == "count" else "money", tooltip=kpi.tooltip)


def _result_row_to_dict(row: object, index: int) -> dict[str, object]:
    if isinstance(row, GroupLendingRow):
        return row.to_dict(index)
    if hasattr(row, "to_dict"):
        return row.to_dict()
    return dict(row) if isinstance(row, dict) else {"Thông tin": row}


def _columns_for(view_mode: str, display_mode: str):
    if view_mode == VIEW_COMPARE_PERIODS:
        return ASSOCIATION_COMPARE_COLUMNS if display_mode == SUMMARY_BY_ASSOCIATION else GROUP_COMPARE_COLUMNS
    return ASSOCIATION_COLUMNS if display_mode == SUMMARY_BY_ASSOCIATION else GROUP_DETAIL_COLUMNS


def _widths_for(view_mode: str, display_mode: str) -> tuple[int, ...]:
    if view_mode == VIEW_COMPARE_PERIODS:
        if display_mode == SUMMARY_BY_ASSOCIATION:
            return (150, 100, 100, 105, 110, 110, 110, 130, 130, 130, 120, 110, 110, 150)
        return (120, 180, 130, 110, 110, 120, 130, 130, 130, 130, 150)
    if display_mode == SUMMARY_BY_ASSOCIATION:
        return (160, 120, 130, 145, 140, 100, 140, 150)
    return (56, 78, 128, 180, 150, 150, 90, 120, 110, 130, 120, 90, 140, 150, 120)
