from __future__ import annotations

from datetime import date
from pathlib import Path
import re
from typing import Callable

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QAbstractItemView,
    QButtonGroup,
    QCheckBox,
    QComboBox,
    QDialog,
    QFileDialog,
    QHBoxLayout,
    QHeaderView,
    QInputDialog,
    QLabel,
    QMessageBox,
    QProgressBar,
    QPushButton,
    QRadioButton,
    QSizePolicy,
    QTabWidget,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from agribank_v3.features.credit.summary.credit_report import (
    CREDIT_DATABASE_NAME,
    CREDIT_QUALITY_DISPLAY_NAME,
    GROUP_CREDIT_QUALITY,
    GROUP_CUSTOMER_TYPE,
    GROUP_DECREE55,
    GROUP_INDUSTRY,
    GROUP_SUMMARY,
    GROUP_TERM_STRUCTURE,
    REPORT_GROUP_LABELS,
    VIEW_COMPARE_PERIODS,
    VIEW_CURRENT_PERIOD,
    CreditReportFilters,
    CreditReportRepository,
    ReportComparisonResult,
    ReportMetric,
    ReportSnapshot,
    ReportTableRow,
    compare_report_snapshots,
    get_report_snapshot,
)
from agribank_v3.features.credit.summary.models import (
    REPORT_SUMMARY_TITLE,
    ImportResult,
)
from agribank_v3.features.credit.summary.repository import SummaryRepository
from agribank_v3.features.credit.summary.ln01_import_dialog import ask_ln01_duplicate_decision
from agribank_v3.features.credit.summary.services import (
    Ln01DuplicateDecision,
    Ln01ImportCoordinator,
    parse_period_from_filename,
)
from agribank_v3.features.credit.summary.group_lending.tab import GroupLendingTab
from agribank_v3.settings import AppSettingsDatabase
from agribank_v3.ui.components.controls import (
    center_window_on_screen,
    combo_box,
    current_data,
    danger_button,
    fit_window_to_screen,
    populate_combo,
    primary_button,
    secondary_button,
)
from agribank_v3.ui.components.flow_layout import FlowLayout
from agribank_v3.ui.components.kpi import KpiMetric, MetricGrid
from agribank_v3.ui.workers import run_in_thread


GROUP_ORDER = (
    GROUP_SUMMARY,
    GROUP_TERM_STRUCTURE,
    GROUP_CUSTOMER_TYPE,
    GROUP_CREDIT_QUALITY,
    GROUP_DECREE55,
    GROUP_INDUSTRY,
)

MONEY_HEADERS = {
    "Giá trị",
    "Dư nợ LN01",
    "Dư nợ thẻ",
    "Dư nợ thẻ tín dụng",
    "Tổng dư nợ",
    "Dư nợ Từ kỳ",
    "Dư nợ Đến kỳ",
    "Giá trị Từ kỳ",
    "Giá trị Đến kỳ",
    "Tăng/giảm",
    "Tăng/giảm tuyệt đối",
}
PERCENT_HEADERS = {
    "Tỷ trọng",
    "Tỷ lệ",
    "Tỷ trọng LN01",
    "Tỷ trọng trên dư nợ LN01",
    "Tăng trưởng (%)",
    "Tỷ trọng Từ kỳ",
    "Tỷ trọng Đến kỳ",
}
PERCENT_POINT_HEADERS = {"Thay đổi tỷ trọng (điểm %)"}
COUNT_HEADERS = {
    "Số khách hàng",
    "Số món",
    "Số lượng khách hàng",
    "Số khách hàng cá nhân",
    "Số khách hàng pháp nhân",
    "Số dòng nguồn",
    "Số món sau chuẩn hóa",
    "Số cảnh báo",
    "Số KH Từ kỳ",
    "Số KH Đến kỳ",
    "Số món Từ kỳ",
    "Số món Đến kỳ",
    "KH cá nhân Từ kỳ",
    "KH cá nhân Đến kỳ",
    "KH pháp nhân Từ kỳ",
    "KH pháp nhân Đến kỳ",
    "Thay đổi số KH",
}


class ReportSummaryWindow(QDialog):
    def __init__(self, main_database_path: Path | None = None, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setWindowTitle(f"{REPORT_SUMMARY_TITLE} - AgribankV3")
        self.setModal(False)
        self.setWindowModality(Qt.WindowModality.NonModal)
        self.setWindowFlags(
            self.windowFlags()
            | Qt.WindowType.WindowMinimizeButtonHint
            | Qt.WindowType.WindowMaximizeButtonHint
        )
        settings_path = main_database_path or AppSettingsDatabase().database_path
        self.main_database_path = settings_path
        self.summary_repository = SummaryRepository(settings_path)
        self.repository = CreditReportRepository(settings_path)
        self._loading_filters = False
        self._overview_generation = 0
        self._current_snapshot: ReportSnapshot | None = None
        self._from_snapshot: ReportSnapshot | None = None
        self._to_snapshot: ReportSnapshot | None = None
        self._build_ui()
        fit_window_to_screen(self, width_ratio=0.94, height_ratio=0.90)
        center_window_on_screen(self)
        self.reload()

    def _build_ui(self) -> None:
        layout = QVBoxLayout(self)
        layout.setContentsMargins(12, 10, 12, 10)
        layout.setSpacing(8)

        title_row = QHBoxLayout()
        title = QLabel(REPORT_SUMMARY_TITLE)
        title.setObjectName("PageTitle")
        self.status_label = QLabel("")
        self.status_label.setObjectName("MutedText")
        title_row.addWidget(title)
        title_row.addStretch()
        title_row.addWidget(self.status_label)
        layout.addLayout(title_row)

        filter_area = QWidget()
        filter_area.setObjectName("ReportSummaryFilterArea")
        filter_flow = FlowLayout(filter_area, spacing=7)
        self.from_period_combo = _report_combo("Từ kỳ")
        self.to_period_combo = _report_combo("Đến kỳ")
        self.period_combo = _report_combo("Kỳ báo cáo")
        self.branch_combo = _report_combo("Chi nhánh")
        self.office_combo = _report_combo("Phòng giao dịch")
        self.customer_type_combo = _report_combo("Loại khách hàng")
        self.debt_group_combo = _report_combo("Nhóm nợ")
        self.term_combo = _report_combo("Loại thời hạn")
        self.officer_combo = _report_combo("CBTD", maximum_width=260)
        self.view_mode_widget = create_report_view_mode_widget(self)

        for label, widget in (
            ("Từ kỳ", self.from_period_combo),
            ("Đến kỳ", self.to_period_combo),
            ("Kỳ báo cáo", self.period_combo),
            ("Chi nhánh", self.branch_combo),
            ("Phòng giao dịch", self.office_combo),
            ("Loại khách hàng", self.customer_type_combo),
            ("Nhóm nợ", self.debt_group_combo),
            ("Loại thời hạn", self.term_combo),
            ("CBTD", self.officer_combo),
        ):
            filter_flow.addWidget(_labeled_control(label, widget))
        filter_flow.addWidget(self.view_mode_widget)
        layout.addWidget(filter_area)

        action_area = QWidget()
        action_area.setObjectName("ReportSummaryActionArea")
        action_flow = FlowLayout(action_area, spacing=7)
        self.import_button = primary_button("Import LN01")
        self.refresh_button = secondary_button("Làm mới")
        self.clear_button = secondary_button("Xóa lọc")
        self.delete_period_button = danger_button("Xóa dữ liệu kỳ")
        self.export_button = secondary_button("Xuất Excel")
        self.backup_button = secondary_button("Sao lưu")
        self.restore_button = secondary_button("Khôi phục")
        self.maintenance_button = secondary_button("Bảo trì dữ liệu")
        self.progress = QProgressBar()
        self.progress.setVisible(False)
        self.progress.setMinimumWidth(260)
        for button in (
            self.import_button,
            self.refresh_button,
            self.clear_button,
            self.delete_period_button,
            self.export_button,
            self.backup_button,
            self.restore_button,
            self.maintenance_button,
        ):
            action_flow.addWidget(button)
        action_flow.addWidget(self.progress)
        layout.addWidget(action_area)

        self.tabs = QTabWidget()
        self.overview_tab = self._build_overview_tab()
        self.group_lending_tab = GroupLendingTab(
            self.main_database_path,
            period_provider=self._group_lending_period_context,
            view_mode_provider=self._view_mode,
            parent=self,
        )
        self.history_table = _table()
        self.rules_table = _table()
        self.rules_tab = self._build_rules_tab()
        self.tabs.addTab(self.overview_tab, "Tổng quan")
        self.tabs.addTab(self.group_lending_tab, "Cho vay qua tổ")
        self.tabs.addTab(_table_tab(self.history_table), "Lịch sử import")
        self.tabs.addTab(self.rules_tab, "Cài đặt phân loại")
        layout.addWidget(self.tabs, stretch=1)

        for combo in (
            self.from_period_combo,
            self.to_period_combo,
            self.period_combo,
            self.branch_combo,
            self.office_combo,
            self.customer_type_combo,
            self.debt_group_combo,
            self.term_combo,
            self.officer_combo,
        ):
            combo.currentIndexChanged.connect(lambda _index: self._filter_changed())
        self.mode_button_group.buttonToggled.connect(self._mode_toggled)
        self.import_button.clicked.connect(self.import_ln01)
        self.refresh_button.clicked.connect(self.reload)
        self.clear_button.clicked.connect(self.clear_filters)
        self.delete_period_button.clicked.connect(self.delete_period)
        self.export_button.clicked.connect(self.export_excel)
        self.backup_button.clicked.connect(self.backup_data)
        self.restore_button.clicked.connect(self.restore_data)
        self.maintenance_button.clicked.connect(self.maintenance)
        self.tabs.currentChanged.connect(lambda _index: self._active_tab_changed())

    def _build_overview_tab(self) -> QWidget:
        tab = QWidget()
        layout = QVBoxLayout(tab)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(8)

        group_area = QWidget()
        group_area.setObjectName("ReportOverviewGroupArea")
        group_flow = FlowLayout(group_area, spacing=7)
        group_label = QLabel("Nhóm số liệu")
        group_label.setObjectName("MutedText")
        group_flow.addWidget(group_label)
        self.analysis_group = QButtonGroup(self)
        self.analysis_group.setExclusive(True)
        self.group_buttons: dict[str, QRadioButton] = {}
        for key in GROUP_ORDER:
            button = QRadioButton(REPORT_GROUP_LABELS[key])
            button.setProperty("group_key", key)
            if key == GROUP_SUMMARY:
                button.setChecked(True)
            self.analysis_group.addButton(button)
            self.group_buttons[key] = button
            group_flow.addWidget(button)
        layout.addWidget(group_area)

        self.metrics = MetricGrid()
        layout.addWidget(self.metrics)
        self.overview_table = _table()
        layout.addWidget(self.overview_table, stretch=1)
        self.note_label = QLabel("")
        self.note_label.setObjectName("MutedText")
        self.note_label.setWordWrap(True)
        layout.addWidget(self.note_label)

        self.analysis_group.buttonToggled.connect(self._group_toggled)
        return tab

    def _build_rules_tab(self) -> QWidget:
        tab = QWidget()
        layout = QVBoxLayout(tab)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(8)
        toolbar = QHBoxLayout()
        add_button = primary_button("Thêm mã")
        edit_button = secondary_button("Sửa mã")
        delete_button = danger_button("Xóa mã")
        restore_button = secondary_button("Khôi phục mặc định")
        self.rule_enabled_check = QCheckBox("Bật mã đang chọn")
        self.rule_enabled_check.setChecked(True)
        toolbar.addWidget(add_button)
        toolbar.addWidget(edit_button)
        toolbar.addWidget(delete_button)
        toolbar.addWidget(restore_button)
        toolbar.addWidget(self.rule_enabled_check)
        toolbar.addStretch()
        layout.addLayout(toolbar)
        layout.addWidget(self.rules_table, stretch=1)
        add_button.clicked.connect(self.add_rule)
        edit_button.clicked.connect(self.edit_rule)
        delete_button.clicked.connect(self.delete_rule)
        restore_button.clicked.connect(self.restore_rules)
        return tab

    def reload(self) -> None:
        try:
            self._reload_filter_values()
            self.reload_overview()
            self.group_lending_tab.reload_filter_values()
            self.group_lending_tab.refresh()
            _render_table(self.history_table, self.repository.import_history_rows(self._filters()))
            _render_table(self.rules_table, self.repository.customer_type_rule_rows())
            self._update_status()
        except Exception as exc:
            QMessageBox.warning(self, REPORT_SUMMARY_TITLE, str(exc))

    def reload_overview(self) -> None:
        if self._loading_filters:
            return
        self._overview_generation += 1
        generation = self._overview_generation
        filters = self._filters()
        view_mode = self._view_mode()
        if self.isVisible():
            self.progress.setVisible(True)
            self.progress.setRange(0, 0)
            self.progress.setFormat("Đang tải số liệu báo cáo")

            def done(payload: tuple[ReportSnapshot | None, ReportSnapshot | None, ReportSnapshot | None]) -> None:
                if generation != self._overview_generation:
                    return
                self.progress.setVisible(False)
                self.progress.setRange(0, 100)
                self._current_snapshot, self._from_snapshot, self._to_snapshot = payload
                self._render_overview_from_cache()

            def failed(exc: Exception) -> None:
                if generation != self._overview_generation:
                    return
                self.progress.setVisible(False)
                self.progress.setRange(0, 100)
                QMessageBox.warning(self, REPORT_SUMMARY_TITLE, str(exc))

            run_in_thread(
                self,
                lambda _progress: self._build_overview_payload(filters, view_mode),
                done,
                failed,
                lambda message: self.progress.setFormat(message),
            )
            return
        try:
            self._current_snapshot, self._from_snapshot, self._to_snapshot = self._build_overview_payload(filters, view_mode)
            if generation == self._overview_generation:
                self._render_overview_from_cache()
        except Exception as exc:
            QMessageBox.warning(self, REPORT_SUMMARY_TITLE, str(exc))

    def _build_overview_payload(
        self,
        filters: CreditReportFilters,
        view_mode: str,
    ) -> tuple[ReportSnapshot | None, ReportSnapshot | None, ReportSnapshot | None]:
        if view_mode == VIEW_COMPARE_PERIODS:
            from_period = filters.from_period
            to_period = filters.to_period
            from_snapshot = get_report_snapshot(self.repository, from_period, filters)
            to_snapshot = from_snapshot if to_period == from_period else get_report_snapshot(self.repository, to_period, filters)
            return None, from_snapshot, to_snapshot
        return get_report_snapshot(self.repository, filters.period, filters), None, None

    def _render_overview_from_cache(self) -> None:
        group = self._group_key()
        if self._view_mode() == VIEW_COMPARE_PERIODS:
            if self._from_snapshot is None or self._to_snapshot is None:
                self.metrics.set_empty(("Tổng dư nợ", "Tổng khách hàng", "Nợ xấu"))
                _render_table(self.overview_table, [])
                self.note_label.setText("Chưa có dữ liệu so sánh.")
                return
            result = compare_report_snapshots(self._from_snapshot, self._to_snapshot, group)
            self._render_compare_metrics(result)
            _render_table(self.overview_table, result.rows)
            self.note_label.setText(" ".join(result.notes))
            return
        if self._current_snapshot is None or not self._current_snapshot.exists:
            self.metrics.set_empty(_empty_metric_labels(group))
            _render_table(self.overview_table, [])
            self.note_label.setText(self._current_snapshot.note if self._current_snapshot else "Chưa có dữ liệu báo cáo.")
            return
        self._render_current_metrics(self._current_snapshot, group)
        _render_table(self.overview_table, self._current_snapshot.groups.get(group, ()))
        self.note_label.setText(self._current_snapshot.card_message if group in {GROUP_SUMMARY, GROUP_TERM_STRUCTURE, GROUP_CUSTOMER_TYPE, GROUP_CREDIT_QUALITY} else "")

    def _render_current_metrics(self, snapshot: ReportSnapshot, group: str) -> None:
        if group == GROUP_SUMMARY:
            self._render_summary_metrics(snapshot.summary)
            return
        rows = snapshot.groups.get(group, ())
        if group in {GROUP_TERM_STRUCTURE, GROUP_CUSTOMER_TYPE, GROUP_CREDIT_QUALITY, GROUP_INDUSTRY}:
            label_header, value_header = {
                GROUP_TERM_STRUCTURE: ("Loại thời hạn", "Tổng dư nợ"),
                GROUP_CUSTOMER_TYPE: ("Loại khách hàng", "Tổng dư nợ"),
                GROUP_CREDIT_QUALITY: ("Nhóm nợ", "Tổng dư nợ"),
                GROUP_INDUSTRY: ("Nhóm ngành", "Tổng dư nợ"),
            }[group]
            self.metrics.set_metrics(
                [
                    KpiMetric(str(row.get(label_header)), row.get(value_header), "money")
                    for row in rows
                ]
            )
            return
        if group == GROUP_DECREE55 and rows:
            row = rows[0]
            self.metrics.set_metrics(
                [
                    KpiMetric("Dư nợ Nghị định 55", row.get("Tổng dư nợ"), "money"),
                    KpiMetric("Tỷ trọng LN01", row.get("Tỷ trọng trên dư nợ LN01"), "percent"),
                    KpiMetric("Số món", row.get("Số món"), "count"),
                    KpiMetric("Số khách hàng", row.get("Số lượng khách hàng"), "count"),
                    KpiMetric("Khách hàng cá nhân", row.get("Số khách hàng cá nhân"), "count", group="secondary"),
                    KpiMetric("Khách hàng pháp nhân", row.get("Số khách hàng pháp nhân"), "count", group="secondary"),
                ]
            )
            return
        self.metrics.set_empty(_empty_metric_labels(group))

    def _render_summary_metrics(self, summary: dict[str, object]) -> None:
        tooltip = (
            f"LN01: {_format_money(summary['ln01_total_balance'])}\n"
            f"Thẻ tín dụng: {_format_money(summary['credit_card_balance'])}\n"
            f"Tổng cộng: {_format_money(summary['total_balance'])}"
        )
        self.metrics.set_metrics(
            [
                KpiMetric("Tổng dư nợ", summary["total_balance"], "money", tooltip=tooltip),
                KpiMetric("Dư nợ LN01", summary["ln01_total_balance"], "money"),
                KpiMetric("Dư nợ thẻ tín dụng", summary["credit_card_balance"], "money", tooltip=str(summary["card_data_message"])),
                KpiMetric("Tổng khách hàng còn dư nợ", summary["customer_count"], "count"),
                KpiMetric("Dư nợ ngắn hạn", summary["short_term_balance"], "money"),
                KpiMetric("Dư nợ trung hạn", summary["medium_term_balance"], "money"),
                KpiMetric("Dư nợ dài hạn", summary["long_term_balance"], "money"),
                KpiMetric("Dư nợ cá nhân", summary["personal_balance"], "money", group="secondary"),
                KpiMetric("Dư nợ pháp nhân", summary["legal_balance"], "money", group="secondary"),
                KpiMetric("Nợ cần chú ý", summary["debt_group_2_balance"], "money", group="secondary"),
                KpiMetric("Nợ xấu", summary["bad_debt_balance"], "money", group="secondary"),
                KpiMetric("Dư nợ Nghị định 55", summary["decree55_balance"], "money", group="secondary"),
            ]
        )

    def _render_compare_metrics(self, result: ReportComparisonResult) -> None:
        metrics: list[KpiMetric] = []
        for metric in result.kpis:
            metrics.append(
                KpiMetric(
                    metric.label,
                    metric.difference,
                    "count" if metric.value_kind == "count" else "money",
                    signed=True,
                    tooltip=_compare_metric_tooltip(metric, result.from_snapshot.period, result.to_snapshot.period),
                )
            )
        self.metrics.set_metrics(metrics)

    def _reload_filter_values(self) -> None:
        values = self.repository.filter_values()
        self._loading_filters = True
        try:
            current_period = current_data(self.period_combo)
            current_from = current_data(self.from_period_combo)
            current_to = current_data(self.to_period_combo)
            populate_combo(self.from_period_combo, values["periods"])
            populate_combo(self.to_period_combo, values["periods"])
            populate_combo(self.period_combo, values["periods"])
            populate_combo(self.branch_combo, values["branches"])
            populate_combo(self.customer_type_combo, values["customer_types"])
            populate_combo(self.debt_group_combo, values["debt_groups"])
            populate_combo(self.term_combo, values["terms"])
            populate_combo(self.officer_combo, values["officers"])
            populate_combo(self.office_combo, [])
            if current_period:
                _select_combo_data(self.period_combo, current_period)
            elif self.period_combo.count() > 1:
                self.period_combo.setCurrentIndex(self.period_combo.count() - 1)
            if current_from:
                _select_combo_data(self.from_period_combo, current_from)
            elif self.from_period_combo.count() > 1:
                self.from_period_combo.setCurrentIndex(max(1, self.from_period_combo.count() - 2))
            if current_to:
                _select_combo_data(self.to_period_combo, current_to)
            elif self.to_period_combo.count() > 1:
                self.to_period_combo.setCurrentIndex(self.to_period_combo.count() - 1)
        finally:
            self._loading_filters = False
        self._update_period_combo_state()

    def _filters(self) -> CreditReportFilters:
        return CreditReportFilters(
            period=current_data(self.period_combo),
            from_period=current_data(self.from_period_combo),
            to_period=current_data(self.to_period_combo),
            branch_code=current_data(self.branch_combo),
            transaction_office=current_data(self.office_combo),
            customer_type=current_data(self.customer_type_combo),
            debt_group=current_data(self.debt_group_combo),
            term_category=current_data(self.term_combo),
            officer=current_data(self.officer_combo),
        )

    def _view_mode(self) -> str:
        button = self.mode_button_group.checkedButton()
        return str(button.property("view_mode") if button is not None else VIEW_CURRENT_PERIOD)

    def _group_key(self) -> str:
        button = self.analysis_group.checkedButton()
        return str(button.property("group_key") if button is not None else GROUP_SUMMARY)

    def _mode_toggled(self, button: QRadioButton, checked: bool) -> None:
        if checked:
            self._update_period_combo_state()
            current = self.tabs.currentWidget()
            if current is self.group_lending_tab:
                self.group_lending_tab.refresh()
            elif current is self.overview_tab:
                self.reload_overview()

    def _group_toggled(self, button: QRadioButton, checked: bool) -> None:
        if checked:
            self._render_overview_from_cache()

    def _filter_changed(self) -> None:
        if not self._loading_filters:
            self.reload_overview()
            self.group_lending_tab.refresh()

    def _update_period_combo_state(self) -> None:
        compare = self._view_mode() == VIEW_COMPARE_PERIODS
        self.from_period_combo.setEnabled(compare)
        self.to_period_combo.setEnabled(compare)
        self.period_combo.setEnabled(not compare)

    def import_ln01(self) -> None:
        path, _ = QFileDialog.getOpenFileName(self, "Chọn file LN01", "", "CSV (*.csv)")
        if not path:
            return
        default_period = parse_period_from_filename(Path(path).name)
        if not re.fullmatch(r"\d{4}-\d{2}", default_period):
            default_period = ""
        period, ok = QInputDialog.getText(
            self,
            "Kỳ dữ liệu",
            f"File nguồn: {Path(path).name}\nKỳ dữ liệu báo cáo (YYYY-MM):",
            text=default_period,
        )
        if not ok:
            return
        period = str(period or "").strip()
        if not re.fullmatch(r"\d{4}-\d{2}", period):
            QMessageBox.warning(self, REPORT_SUMMARY_TITLE, "Kỳ dữ liệu phải có dạng YYYY-MM.")
            return
        try:
            coordinator = Ln01ImportCoordinator(self.summary_repository, self.repository)
            prepared = coordinator.prepare_import(
                Path(path),
                period=period,
                reference_date=date.today(),
            )
        except Exception as exc:
            QMessageBox.warning(self, REPORT_SUMMARY_TITLE, str(exc))
            return
        decision = ask_ln01_duplicate_decision(self, prepared.context)
        if decision == Ln01DuplicateDecision.CANCEL:
            return
        self._run_background(
            "Import LN01",
            lambda progress: coordinator.execute_prepared_import(
                prepared,
                duplicate_decision=decision,
                progress=progress,
            ),
            self._import_done,
        )

    def _import_done(self, result: object) -> None:
        message = result.message if isinstance(result, ImportResult) else "Import xong."
        QMessageBox.information(self, REPORT_SUMMARY_TITLE, message)
        self.reload()

    def clear_filters(self) -> None:
        self._loading_filters = True
        try:
            for combo in (
                self.from_period_combo,
                self.to_period_combo,
                self.period_combo,
                self.branch_combo,
                self.office_combo,
                self.customer_type_combo,
                self.debt_group_combo,
                self.term_combo,
                self.officer_combo,
            ):
                combo.setCurrentIndex(0)
            self.current_mode_radio.setChecked(True)
        finally:
            self._loading_filters = False
        self.reload()

    def delete_period(self) -> None:
        period = current_data(self.period_combo)
        if not period:
            QMessageBox.warning(self, REPORT_SUMMARY_TITLE, "Chưa chọn kỳ cần xóa.")
            return
        answer = QMessageBox.question(
            self,
            REPORT_SUMMARY_TITLE,
            (
                f"Thao tác này chỉ xóa dữ liệu Tổng hợp số liệu báo cáo kỳ {period} trong {CREDIT_DATABASE_NAME}. "
                "File Hạn mức HMHETHAN và dữ liệu NIM Dư nợ không bị xóa."
            ),
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        if answer != QMessageBox.StandardButton.Yes:
            return
        try:
            result = self.repository.delete_report_period(period)
        except Exception as exc:
            QMessageBox.warning(self, REPORT_SUMMARY_TITLE, str(exc))
            return
        QMessageBox.information(
            self,
            REPORT_SUMMARY_TITLE,
            f"Đã xóa kỳ {period}: {result['loan_rows']} dòng LN01, {result['card_rows']} dòng thẻ.",
        )
        self.reload()

    def export_excel(self) -> None:
        if self.tabs.currentWidget() is self.group_lending_tab:
            self.group_lending_tab.export_excel()
            return
        path, _ = QFileDialog.getSaveFileName(
            self,
            "Xuất Excel",
            f"{REPORT_SUMMARY_TITLE}.xlsx",
            "Excel (*.xlsx)",
        )
        if not path:
            return
        filters = self._filters()
        view_mode = self._view_mode()
        self._run_background(
            "Xuất Excel",
            lambda _progress: self.repository.export_workbook(Path(path), filters=filters, view_mode=view_mode),
            lambda output: QMessageBox.information(self, REPORT_SUMMARY_TITLE, f"Đã xuất: {output}"),
        )

    def backup_data(self) -> None:
        path, _ = QFileDialog.getSaveFileName(
            self,
            "Sao lưu Credit.db",
            "Credit-backup.zip",
            "Backup (*.zip);;SQLite (*.db)",
        )
        if not path:
            return
        try:
            output = self.repository.backup_database(Path(path))
        except Exception as exc:
            QMessageBox.warning(self, REPORT_SUMMARY_TITLE, str(exc))
            return
        QMessageBox.information(self, REPORT_SUMMARY_TITLE, f"Đã sao lưu: {output}")

    def restore_data(self) -> None:
        path, _ = QFileDialog.getOpenFileName(
            self,
            "Khôi phục Credit.db",
            "",
            "Backup (*.zip);;SQLite (*.db)",
        )
        if not path:
            return
        try:
            output = self.repository.restore_database(Path(path))
        except Exception as exc:
            QMessageBox.warning(self, REPORT_SUMMARY_TITLE, str(exc))
            return
        QMessageBox.information(self, REPORT_SUMMARY_TITLE, f"Đã khôi phục: {output}")
        self.reload()

    def maintenance(self) -> None:
        dialog = CreditMaintenanceDialog(self.repository, self)
        dialog.exec()
        self._update_status()

    def add_rule(self) -> None:
        code, ok = QInputDialog.getText(self, "Thêm mã cá nhân", "Mã CUSTOMER_TYPE_CODE:")
        if not ok:
            return
        try:
            self.repository.save_personal_type_rule(code, enabled=self.rule_enabled_check.isChecked())
        except Exception as exc:
            QMessageBox.warning(self, REPORT_SUMMARY_TITLE, str(exc))
            return
        self.reload()

    def edit_rule(self) -> None:
        code = self._selected_rule_code()
        if not code:
            QMessageBox.warning(self, REPORT_SUMMARY_TITLE, "Chưa chọn mã cần sửa.")
            return
        new_code, ok = QInputDialog.getText(self, "Sửa mã cá nhân", "Mã CUSTOMER_TYPE_CODE:", text=code)
        if not ok:
            return
        try:
            if new_code != code:
                self.repository.delete_personal_type_rule(code)
            self.repository.save_personal_type_rule(new_code, enabled=self.rule_enabled_check.isChecked())
        except Exception as exc:
            QMessageBox.warning(self, REPORT_SUMMARY_TITLE, str(exc))
            return
        self.reload()

    def delete_rule(self) -> None:
        code = self._selected_rule_code()
        if not code:
            QMessageBox.warning(self, REPORT_SUMMARY_TITLE, "Chưa chọn mã cần xóa.")
            return
        try:
            self.repository.delete_personal_type_rule(code)
        except Exception as exc:
            QMessageBox.warning(self, REPORT_SUMMARY_TITLE, str(exc))
            return
        self.reload()

    def restore_rules(self) -> None:
        self.repository.restore_default_type_rules()
        self.reload()

    def _selected_rule_code(self) -> str:
        row = self.rules_table.currentRow()
        if row < 0:
            return ""
        item = self.rules_table.item(row, 0)
        return item.text().strip() if item is not None else ""

    def _update_status(self) -> None:
        status = self.repository.status()
        self.status_label.setText(
            f"{status['period_count']} kỳ | {status['loan_rows']:,} dòng LN01 | {status['card_rows']:,} dòng thẻ | {status['size_bytes']:,} bytes"
        )

    def _run_background(
        self,
        title: str,
        function: Callable[[Callable[[str], None]], object],
        on_done: Callable[[object], None],
    ) -> None:
        self.progress.setVisible(True)
        self.progress.setRange(0, 0)
        self.progress.setFormat(title)

        def done(payload: object) -> None:
            self.progress.setVisible(False)
            self.progress.setRange(0, 100)
            on_done(payload)

        def failed(exc: Exception) -> None:
            self.progress.setVisible(False)
            self.progress.setRange(0, 100)
            QMessageBox.warning(self, title, str(exc))

        run_in_thread(self, function, done, failed, lambda message: self.progress.setFormat(message))

    def _group_lending_period_context(self) -> tuple[str, str, str]:
        return (
            current_data(self.period_combo),
            current_data(self.from_period_combo),
            current_data(self.to_period_combo),
        )

    def _active_tab_changed(self) -> None:
        if self.tabs.currentWidget() is self.group_lending_tab:
            self.group_lending_tab.refresh()


class CreditMaintenanceDialog(QDialog):
    def __init__(self, repository: CreditReportRepository, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.repository = repository
        self.setWindowTitle("Bảo trì Credit.db")
        self.resize(980, 640)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(8)
        self.summary_label = QLabel("")
        self.summary_label.setObjectName("MutedText")
        self.summary_label.setWordWrap(True)
        layout.addWidget(self.summary_label)
        self.table = _table()
        layout.addWidget(self.table, stretch=1)
        action_row = QHBoxLayout()
        self.refresh_button = secondary_button("Làm mới")
        self.cleanup_button = secondary_button("Xóa dữ liệu khách hàng")
        self.compact_button = primary_button("Thu gọn Credit.db")
        self.close_button = secondary_button("Đóng")
        self.progress = QProgressBar()
        self.progress.setVisible(False)
        self.progress.setMinimumWidth(220)
        for button in (self.refresh_button, self.cleanup_button, self.compact_button, self.close_button):
            action_row.addWidget(button)
        action_row.addWidget(self.progress)
        action_row.addStretch()
        layout.addLayout(action_row)
        self.refresh_button.clicked.connect(self.reload)
        self.cleanup_button.clicked.connect(self.cleanup_orphans)
        self.compact_button.clicked.connect(self.compact_database)
        self.close_button.clicked.connect(self.accept)
        self.reload()

    def reload(self) -> None:
        diagnostics = self.repository.diagnostics()
        _render_table(self.table, _diagnostic_rows(diagnostics))
        self.summary_label.setText(
            "SQLite DELETE không tự làm nhỏ file vật lý. "
            f"Page trống: {_format_count(diagnostics['freelist_count'])}; "
            f"ước tính có thể thu hồi: {_format_money(diagnostics['reclaimable_bytes'])} bytes."
        )

    def cleanup_orphans(self) -> None:
        self._run_background(
            "Làm sạch dữ liệu khách hàng",
            lambda _progress: self.repository.cleanup_orphan_customers(),
            lambda result: QMessageBox.information(
                self,
                self.windowTitle(),
                f"Đã xóa {_format_count(result['deleted_customers'])} khách hàng.",
            ),
        )

    def compact_database(self) -> None:
        answer = QMessageBox.question(
            self,
            self.windowTitle(),
            (
                "Thu gọn Credit.db sẽ tạo backup, checkpoint WAL và chạy VACUUM.\n\n"
                "Tiếp tục?"
            ),
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        if answer != QMessageBox.StandardButton.Yes:
            return
        self._run_background(
            "Thu gọn Credit.db",
            lambda _progress: self.repository.compact_database(),
            self._compact_done,
        )

    def _compact_done(self, result: object) -> None:
        payload = dict(result) if isinstance(result, dict) else {}
        QMessageBox.information(
            self,
            self.windowTitle(),
            (
                f"Backup: {payload.get('backup_path', '')}\n"
                f"Dung lượng trước: {_format_money(payload.get('before_size_bytes'))} bytes\n"
                f"Dung lượng sau: {_format_money(payload.get('after_size_bytes'))} bytes\n"
                f"Đã giảm: {_format_money(payload.get('reduced_bytes'))} bytes\n"
                f"Thời gian: {_format_count(payload.get('duration_ms'))} ms\n"
                f"Integrity: {payload.get('integrity_check', '')}"
            ),
        )

    def _run_background(
        self,
        title: str,
        function: Callable[[Callable[[str], None]], object],
        on_done: Callable[[object], None],
    ) -> None:
        self.progress.setVisible(True)
        self.progress.setRange(0, 0)
        self.progress.setFormat(title)
        for button in (self.refresh_button, self.cleanup_button, self.compact_button):
            button.setEnabled(False)

        def done(payload: object) -> None:
            self.progress.setVisible(False)
            self.progress.setRange(0, 100)
            for button in (self.refresh_button, self.cleanup_button, self.compact_button):
                button.setEnabled(True)
            on_done(payload)
            self.reload()

        def failed(exc: Exception) -> None:
            self.progress.setVisible(False)
            self.progress.setRange(0, 100)
            for button in (self.refresh_button, self.cleanup_button, self.compact_button):
                button.setEnabled(True)
            QMessageBox.warning(self, title, str(exc))

        run_in_thread(self, function, done, failed, lambda message: self.progress.setFormat(message))


def _report_combo(label: str, *, maximum_width: int | None = None) -> QComboBox:
    widths = {
        "Từ kỳ": (132, 150),
        "Đến kỳ": (132, 150),
        "Kỳ báo cáo": (160, 190),
        "Chi nhánh": (168, 220),
        "Phòng giao dịch": (184, 250),
        "Loại khách hàng": (176, 230),
        "Nhóm nợ": (152, 200),
        "Loại thời hạn": (166, 220),
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


def create_report_view_mode_widget(window: ReportSummaryWindow) -> QWidget:
    container = QWidget()
    container.setObjectName("ReportViewModeWidget")
    container.setSizePolicy(QSizePolicy.Policy.Fixed, QSizePolicy.Policy.Fixed)
    layout = QVBoxLayout(container)
    layout.setContentsMargins(0, 0, 0, 0)
    layout.setSpacing(2)
    caption = QLabel("Chế độ xem")
    caption.setObjectName("MutedText")
    layout.addWidget(caption)

    radio_row = QWidget()
    radio_row.setObjectName("ReportViewModeRadioRow")
    radio_layout = QHBoxLayout(radio_row)
    radio_layout.setContentsMargins(0, 0, 0, 0)
    radio_layout.setSpacing(12)
    window.mode_button_group = QButtonGroup(window)
    window.mode_button_group.setExclusive(True)
    window.current_mode_radio = QRadioButton("Kỳ hiện tại")
    window.compare_mode_radio = QRadioButton("So sánh các kỳ")
    window.current_mode_radio.setObjectName("ReportViewModeCurrentRadio")
    window.compare_mode_radio.setObjectName("ReportViewModeCompareRadio")
    window.current_mode_radio.setProperty("view_mode", VIEW_CURRENT_PERIOD)
    window.compare_mode_radio.setProperty("view_mode", VIEW_COMPARE_PERIODS)
    window.current_mode_radio.setChecked(True)
    font_metrics = window.current_mode_radio.fontMetrics()
    for button in (window.current_mode_radio, window.compare_mode_radio):
        button.setMinimumHeight(30)
        button.setMinimumWidth(font_metrics.horizontalAdvance(button.text()) + 32)
        button.setToolTip(button.text())
        window.mode_button_group.addButton(button)
        radio_layout.addWidget(button)
    layout.addWidget(radio_row)
    minimum_width = (
        window.current_mode_radio.minimumWidth()
        + window.compare_mode_radio.minimumWidth()
        + radio_layout.spacing()
        + 4
    )
    container.setMinimumWidth(minimum_width)
    container.setMaximumWidth(max(minimum_width, 270))
    return container


def _diagnostic_rows(diagnostics: dict[str, object]) -> list[dict[str, object]]:
    rows: list[dict[str, object]] = [
        {"Nhóm": "File", "Chỉ tiêu": "Đường dẫn database", "Giá trị": diagnostics.get("database_path", ""), "Ghi chú": ""},
        {"Nhóm": "File", "Chỉ tiêu": "Dung lượng .db", "Giá trị": diagnostics.get("db_size_bytes", 0), "Ghi chú": "bytes"},
        {"Nhóm": "File", "Chỉ tiêu": "Dung lượng -wal", "Giá trị": diagnostics.get("wal_size_bytes", 0), "Ghi chú": "bytes"},
        {"Nhóm": "File", "Chỉ tiêu": "Dung lượng -shm", "Giá trị": diagnostics.get("shm_size_bytes", 0), "Ghi chú": "bytes"},
        {"Nhóm": "PRAGMA", "Chỉ tiêu": "page_size", "Giá trị": diagnostics.get("page_size", 0), "Ghi chú": ""},
        {"Nhóm": "PRAGMA", "Chỉ tiêu": "page_count", "Giá trị": diagnostics.get("page_count", 0), "Ghi chú": ""},
        {"Nhóm": "PRAGMA", "Chỉ tiêu": "freelist_count", "Giá trị": diagnostics.get("freelist_count", 0), "Ghi chú": ""},
        {"Nhóm": "PRAGMA", "Chỉ tiêu": "Dung lượng có thể thu hồi", "Giá trị": diagnostics.get("reclaimable_bytes", 0), "Ghi chú": "freelist_count x page_size"},
        {"Nhóm": "PRAGMA", "Chỉ tiêu": "auto_vacuum", "Giá trị": diagnostics.get("auto_vacuum", 0), "Ghi chú": "0 = NONE"},
        {"Nhóm": "PRAGMA", "Chỉ tiêu": "journal_mode", "Giá trị": diagnostics.get("journal_mode", ""), "Ghi chú": ""},
        {"Nhóm": "PRAGMA", "Chỉ tiêu": "integrity_check", "Giá trị": diagnostics.get("integrity_check", ""), "Ghi chú": ""},
        {"Nhóm": "PRAGMA", "Chỉ tiêu": "schema_version", "Giá trị": diagnostics.get("schema_version", 0), "Ghi chú": ""},
        {"Nhóm": "Logic", "Chỉ tiêu": "Số kỳ", "Giá trị": diagnostics.get("period_count", 0), "Ghi chú": ", ".join(diagnostics.get("periods", []))},
    ]
    for table, count in dict(diagnostics.get("tables", {})).items():
        note = ""
        if table == "credit_customer_master" and int(count or 0) > 0 and int(diagnostics.get("period_count", 0) or 0) == 0:
            note = "Có thể là khách hàng đã xóa sau khi xóa kỳ cuối."
        rows.append({"Nhóm": "Bảng", "Chỉ tiêu": table, "Giá trị": count, "Ghi chú": note})
    table_bytes = dict(diagnostics.get("table_bytes", {}))
    index_bytes = dict(diagnostics.get("index_bytes", {}))
    if table_bytes or index_bytes:
        for name, size in sorted(table_bytes.items(), key=lambda item: int(item[1] or 0), reverse=True):
            rows.append({"Nhóm": "Dung lượng bảng", "Chỉ tiêu": name, "Giá trị": size, "Ghi chú": "bytes"})
        for name, size in sorted(index_bytes.items(), key=lambda item: int(item[1] or 0), reverse=True):
            rows.append({"Nhóm": "Dung lượng index", "Chỉ tiêu": name, "Giá trị": size, "Ghi chú": "bytes"})
    else:
        rows.append({"Nhóm": "dbstat", "Chỉ tiêu": "Trạng thái", "Giá trị": "", "Ghi chú": diagnostics.get("dbstat_error", "SQLite build không hỗ trợ dbstat.")})
    return rows


def _labeled_control(label: str, widget: QWidget) -> QWidget:
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


def _table() -> QTableWidget:
    table = QTableWidget()
    table.setObjectName("SummaryDataTable")
    table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
    table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
    table.setAlternatingRowColors(True)
    table.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
    table.setWordWrap(True)
    header = table.horizontalHeader()
    header.setStretchLastSection(True)
    header.setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
    header.setDefaultAlignment(Qt.AlignmentFlag.AlignCenter | Qt.AlignmentFlag.AlignVCenter)
    return table


def _table_tab(table: QTableWidget) -> QWidget:
    tab = QWidget()
    layout = QVBoxLayout(tab)
    layout.setContentsMargins(0, 0, 0, 0)
    layout.addWidget(table)
    return tab


def _render_table(table: QTableWidget, rows: tuple[ReportTableRow, ...] | list[dict[str, object]]) -> None:
    normalized_rows = [_row_to_dict(row) for row in rows]
    headers = list(normalized_rows[0].keys()) if normalized_rows else ["Thông tin"]
    table.clear()
    table.setColumnCount(len(headers))
    table.setHorizontalHeaderLabels(headers)
    table.setRowCount(len(normalized_rows))
    for row_index, row in enumerate(normalized_rows):
        for column_index, header in enumerate(headers):
            value = row.get(header, "")
            item = QTableWidgetItem(_display_value(header, value))
            if _is_numeric_header(header):
                item.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
            else:
                item.setTextAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
            table.setItem(row_index, column_index, item)
    for index, header in enumerate(headers):
        table.setColumnWidth(index, _column_width(header))


def _row_to_dict(row: ReportTableRow | dict[str, object]) -> dict[str, object]:
    if isinstance(row, ReportTableRow):
        return row.to_dict()
    return dict(row)


def _display_value(header: str, value: object) -> str:
    if value is None:
        return "N/A" if header in PERCENT_HEADERS or header in PERCENT_POINT_HEADERS else "—"
    if header in PERCENT_POINT_HEADERS:
        return _format_percent(value, signed=True, suffix=" điểm %")
    if header == "Tăng trưởng (%)":
        return _format_percent(value, signed=True)
    if header in PERCENT_HEADERS:
        return _format_percent(value)
    if header in MONEY_HEADERS:
        return _format_money(value, signed=header.startswith("Tăng/giảm"))
    if header in COUNT_HEADERS:
        return _format_count(value, signed=header.startswith("Thay đổi"))
    return str(value or "")


def _format_money(value: object, *, signed: bool = False) -> str:
    try:
        number = int(round(float(value or 0)))
    except (TypeError, ValueError):
        return "0"
    prefix = "+" if signed and number > 0 else ""
    return prefix + f"{number:,}".replace(",", ".")


def _format_count(value: object, *, signed: bool = False) -> str:
    try:
        number = int(round(float(value or 0)))
    except (TypeError, ValueError):
        return "0"
    prefix = "+" if signed and number > 0 else ""
    return prefix + f"{number:,}".replace(",", ".")


def _format_percent(value: object, *, signed: bool = False, suffix: str = "%") -> str:
    try:
        number = float(value)
    except (TypeError, ValueError):
        return "N/A"
    prefix = "+" if signed and number > 0 else ""
    text = f"{number:,.2f}".replace(",", "_").replace(".", ",").replace("_", ".")
    return f"{prefix}{text}{suffix}"


def _is_numeric_header(header: str) -> bool:
    return header in MONEY_HEADERS or header in PERCENT_HEADERS or header in PERCENT_POINT_HEADERS or header in COUNT_HEADERS


def _column_width(header: str) -> int:
    if header in {"Ghi chú", "SHA-256", "Tên file nguồn"}:
        return 280
    if header in PERCENT_POINT_HEADERS:
        return 170
    return min(max(110, len(header) * 8 + 24), 230)


def _empty_metric_labels(group: str) -> tuple[str, ...]:
    if group == GROUP_SUMMARY:
        return (
            "Tổng dư nợ",
            "Dư nợ LN01",
            "Dư nợ thẻ tín dụng",
            "Tổng khách hàng còn dư nợ",
            "Dư nợ ngắn hạn",
            "Dư nợ trung hạn",
        )
    return (REPORT_GROUP_LABELS.get(group, "Chỉ tiêu"),)


def _compare_metric_tooltip(metric: ReportMetric, from_period: str, to_period: str) -> str:
    growth = "N/A" if metric.growth_rate is None else _format_percent(metric.growth_rate, signed=True)
    formatter = _format_count if metric.value_kind == "count" else _format_money
    return (
        f"{metric.label}\n"
        f"Từ kỳ {from_period}: {formatter(metric.from_value)}\n"
        f"Đến kỳ {to_period}: {formatter(metric.to_value)}\n"
        f"Chênh lệch: {formatter(metric.difference, signed=True)}\n"
        f"Tăng trưởng: {growth}"
    )


def _select_combo_data(combo: QComboBox, value: str) -> None:
    index = combo.findData(value)
    if index >= 0:
        combo.setCurrentIndex(index)
