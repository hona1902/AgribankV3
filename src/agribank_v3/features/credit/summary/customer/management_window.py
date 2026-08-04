from __future__ import annotations

import logging
from pathlib import Path

from PySide6.QtCore import Signal, Qt, QTimer
from PySide6.QtWidgets import QFileDialog, QComboBox, QHBoxLayout, QLabel, QMessageBox, QProgressDialog, QTabWidget, QVBoxLayout, QDialog, QWidget

from agribank_v3.features.credit.summary.customer.dashboard_tab import CustomerDashboardTab
from agribank_v3.features.credit.summary.customer.customer_list_tab import CustomerListTab
from agribank_v3.features.credit.summary.customer.cross_branch_tab import CrossBranchCustomersTab
from agribank_v3.features.credit.summary.customer.debt_group_tab import DebtGroupAnalysisTab
from agribank_v3.features.credit.summary.customer.delete_period_dialog import DeleteCustomerPeriodDialog
from agribank_v3.features.credit.summary.customer.export_service import (
    export_all_customer_sheets,
    suggested_customer_export_name,
)
from agribank_v3.features.credit.summary.customer.filters import (
    CUSTOMER_TYPE_FILTERS,
    LOAN_TERM_FILTERS,
    CustomerFilters,
)
from agribank_v3.features.credit.summary.customer.formatters import format_customer_type
from agribank_v3.features.credit.summary.customer.import_history_tab import ImportHistoryTab
from agribank_v3.features.credit.summary.customer.maintenance_dialog import CustomerMaintenanceDialog
from agribank_v3.features.credit.summary.customer.movement_tab import CustomerMovementTab
from agribank_v3.features.credit.summary.customer.multiple_officers_tab import MultipleOfficersTab
from agribank_v3.features.credit.summary.customer.officer_management_tab import OfficerManagementTab
from agribank_v3.features.credit.summary.customer.period_validation import validate_dashboard_period_filters
from agribank_v3.features.credit.summary.customer.repository import CustomerRepository
from agribank_v3.features.credit.summary.customer.widgets import (
    SearchBox,
    combo_box,
    current_data,
    danger_button,
    fit_window_to_screen,
    populate_combo,
    populate_officer_combo,
    secondary_button,
)
from agribank_v3.ui.workers import run_in_thread


LOGGER = logging.getLogger(__name__)


class CustomerManagementWindow(QDialog):
    openNimDnRequested = Signal()

    def __init__(self, main_database_path: Path, parent=None) -> None:
        super().__init__(parent)
        self.repository = CustomerRepository(Path(main_database_path))
        self.repository.unit_directory.add_listener(self._unit_directory_changed)
        self.setWindowTitle("Quản lý dữ liệu khách hàng - AgribankV3")
        self.setModal(False)
        self.setWindowFlag(Qt.WindowType.WindowMinMaxButtonsHint, True)
        fit_window_to_screen(
            self,
            width_ratio=0.90,
            height_ratio=0.88,
            max_width=1400,
            max_height=900,
            min_width=980,
            min_height=650,
        )
        layout = QVBoxLayout(self)
        self._export_thread = None
        self._vacuum_thread = None
        self._updating_filters = False
        self._has_period_data = False
        self._available_periods: list[str] = []
        self._filter_timer = QTimer(self)
        self._filter_timer.setSingleShot(True)
        self._filter_timer.setInterval(350)
        self._filter_timer.timeout.connect(self._apply_filter_changed)
        filter_panel = QWidget()
        filter_panel.setObjectName("CustomerFilterPanel")
        filter_layout = QVBoxLayout(filter_panel)
        filter_layout.setContentsMargins(0, 0, 0, 0)
        filter_layout.setSpacing(6)
        filter_row_1 = QHBoxLayout()
        filter_row_2 = QHBoxLayout()
        filter_row_3 = QHBoxLayout()
        self.search_box = SearchBox()
        self.search_box.setMinimumWidth(220)
        self.search_box.setMaximumWidth(340)
        self.period_from_combo = combo_box("Từ kỳ", minimum_width=110, maximum_width=150)
        self.period_to_combo = combo_box("Đến kỳ", minimum_width=110, maximum_width=150)
        self.current_period_combo = combo_box("Kỳ báo cáo", minimum_width=120, maximum_width=160)
        self.branch_combo = combo_box("Tất cả chi nhánh", minimum_width=150, maximum_width=220)
        self.customer_type_combo = combo_box("Tất cả loại khách hàng", minimum_width=160, maximum_width=230)
        self.officer_combo = combo_box(
            "Tất cả cán bộ",
            minimum_width=220,
            maximum_width=320,
            minimum_contents_length=18,
            searchable=True,
        )
        self.loan_term_combo = combo_box("Tất cả thời hạn", minimum_width=150, maximum_width=220)
        populate_combo(self.customer_type_combo, CUSTOMER_TYPE_FILTERS[1:])
        populate_combo(self.loan_term_combo, LOAN_TERM_FILTERS[1:])
        for widget in (
            self.search_box,
            self.period_from_combo,
            self.period_to_combo,
            self.current_period_combo,
        ):
            filter_row_1.addWidget(widget)
        filter_row_1.addStretch()
        for widget in (
            self.branch_combo,
            self.customer_type_combo,
            self.officer_combo,
            self.loan_term_combo,
        ):
            filter_row_2.addWidget(widget)
        filter_row_2.addStretch()
        self.refresh_button = secondary_button("Làm mới")
        self.export_all_button = secondary_button("Xuất toàn bộ")
        self.delete_period_button = danger_button("Xóa dữ liệu kỳ")
        self.maintenance_button = secondary_button("Bảo trì dữ liệu")
        self.refresh_button.clicked.connect(self.refresh_all)
        self.export_all_button.clicked.connect(self.export_all)
        self.delete_period_button.clicked.connect(self.delete_period)
        self.maintenance_button.clicked.connect(self.open_maintenance)
        filter_row_3.addWidget(self.refresh_button)
        filter_row_3.addWidget(self.export_all_button)
        filter_row_3.addWidget(self.delete_period_button)
        filter_row_3.addWidget(self.maintenance_button)
        filter_row_3.addStretch()
        filter_layout.addLayout(filter_row_1)
        filter_layout.addLayout(filter_row_2)
        filter_layout.addLayout(filter_row_3)
        layout.addWidget(filter_panel)
        self.empty_data_banner = self._build_empty_data_banner()
        layout.addWidget(self.empty_data_banner)
        self.tabs = QTabWidget()
        layout.addWidget(self.tabs, stretch=1)
        self.dashboard_tab = CustomerDashboardTab(self.repository, self.current_filters, self)
        self.debt_group_tab = DebtGroupAnalysisTab(self.repository, self.current_filters, self)
        self.list_tab = CustomerListTab(self.repository, self.current_filters, self)
        self.movement_tab = CustomerMovementTab(self.repository, self.current_filters, self)
        self.multiple_tab = MultipleOfficersTab(self.repository, self.current_filters, self)
        self.cross_branch_tab = CrossBranchCustomersTab(self.repository, self)
        self.import_tab = ImportHistoryTab(self.repository, self)
        self.officer_tab = OfficerManagementTab(self.repository, self)
        self.tabs.addTab(self.dashboard_tab, "Dashboard")
        self.tabs.addTab(self.debt_group_tab, "Phân tích nhóm nợ")
        self.tabs.addTab(self.list_tab, "Danh sách khách hàng")
        self.tabs.addTab(self.movement_tab, "Biến động dư nợ")
        self.tabs.addTab(self.multiple_tab, "Nhiều cán bộ quản lý")
        self.tabs.addTab(self.cross_branch_tab, "Khách hàng vay liên chi nhánh")
        self.tabs.addTab(self.import_tab, "Lịch sử import")
        self.tabs.addTab(self.officer_tab, "Quản lý cán bộ")
        self.search_box.debouncedTextChanged.connect(self._filter_changed)
        for combo in (
            self.period_from_combo,
            self.period_to_combo,
            self.current_period_combo,
            self.branch_combo,
            self.customer_type_combo,
            self.officer_combo,
            self.loan_term_combo,
        ):
            combo.currentIndexChanged.connect(self._filter_changed)
        self.tabs.currentChanged.connect(lambda _index: self._refresh_active_tab())
        LOGGER.debug("Customer management window open")
        self.refresh_filters()
        self.refresh_all()

    def _unit_directory_changed(self) -> None:
        self.refresh_all()

    def current_filters(self) -> CustomerFilters:
        return CustomerFilters(
            period_from=current_data(self.period_from_combo),
            period_to=current_data(self.period_to_combo),
            current_period=current_data(self.current_period_combo),
            branch_code=current_data(self.branch_combo),
            customer_type=current_data(self.customer_type_combo),
            officer=current_data(self.officer_combo),
            loan_term=current_data(self.loan_term_combo),
            search_text=self.search_box.text().strip(),
        )

    def refresh_filters(self) -> list[str]:
        self._updating_filters = True
        filters = self.current_filters()
        try:
            periods = self.repository.distinct_periods()
            self._available_periods = periods
            self._has_period_data = bool(periods)
            LOGGER.debug("Customer available period count: %s", len(periods))
            if periods:
                self._populate_period_combo(self.period_from_combo, "Từ kỳ", periods)
                self._populate_period_combo(self.period_to_combo, "Đến kỳ", periods)
                self._populate_period_combo(self.current_period_combo, "Kỳ báo cáo", periods)
                self._normalize_period_combos(periods)
            else:
                self._reset_period_combos_empty()
            self._set_period_combo_tooltips()
            filters = self.current_filters()
            if periods:
                populate_combo(
                    self.branch_combo,
                    [
                        (self.repository.unit_directory.get_branch_display_name(code), code)
                        for code in self.repository.distinct_branch_codes(filters)
                    ],
                )
                populate_combo(
                    self.customer_type_combo,
                    [(format_customer_type(value), value) for value in self.repository.distinct_customer_types(filters)],
                )
                populate_officer_combo(self.officer_combo, self.repository.distinct_officers(filters))
            else:
                populate_combo(self.branch_combo, [])
                populate_combo(self.customer_type_combo, [])
                populate_officer_combo(self.officer_combo, [], first_label="Tất cả cán bộ")
            self.movement_tab.refresh_periods()
            self.cross_branch_tab.refresh_filter_options()
            self.officer_tab.refresh_filters()
        finally:
            self._updating_filters = False
        return list(self._available_periods)

    def refresh_all(self) -> None:
        LOGGER.debug("Customer refresh_all requested")
        self.invalidate_customer_caches()
        periods = self.refresh_filters()
        self._update_empty_data_banner()
        if not periods:
            self.handle_customer_data_became_empty()
            self.import_tab.refresh()
            self.officer_tab.refresh()
            return
        for tab in (
            self.dashboard_tab,
            self.debt_group_tab,
            self.list_tab,
            self.movement_tab,
            self.multiple_tab,
            self.cross_branch_tab,
            self.import_tab,
            self.officer_tab,
        ):
            if hasattr(tab, "refresh"):
                tab.refresh()

    def export_all(self) -> None:
        path, _ = QFileDialog.getSaveFileName(
            self,
            "Xuất toàn bộ dữ liệu khách hàng",
            suggested_customer_export_name("TatCa"),
            "Excel (*.xlsx)",
        )
        if not path:
            return
        filters = self.current_filters()
        previous_period = current_data(self.movement_tab.previous_combo)
        current_period = current_data(self.movement_tab.current_combo)
        progress = QProgressDialog("Đang xuất toàn bộ dữ liệu khách hàng...", "Hủy", 0, 0, self)
        progress.setWindowTitle("Xuất toàn bộ")
        progress.setWindowModality(Qt.WindowModality.WindowModal)
        progress.setMinimumDuration(0)
        progress.show()

        def task():
            return export_all_customer_sheets(
                self.repository,
                filters,
                Path(path),
                previous_period=previous_period,
                current_period=current_period,
            )

        def done(output):
            progress.close()
            QMessageBox.information(self, "Xuất toàn bộ dữ liệu khách hàng", f"Đã xuất: {output}")

        def failed(exc: Exception):
            progress.close()
            QMessageBox.warning(self, "Xuất toàn bộ dữ liệu khách hàng", str(exc))

        self._export_thread = run_in_thread(self, task, done, failed)

    def delete_period(self) -> None:
        self.cancel_period_data_queries()
        dialog = DeleteCustomerPeriodDialog(self.repository, self)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            LOGGER.debug("Customer delete completed: %s", dialog.deleted_info)
            self.cancel_period_data_queries()
            self.invalidate_customer_caches()
            if not self.repository.has_period_data():
                self.refresh_filters()
                self.handle_customer_data_became_empty()
                self.import_tab.refresh()
                self.officer_tab.refresh()
                self._refresh_open_related_windows()
                self._offer_vacuum_after_last_period_delete()
            else:
                self.refresh_all()
                self._refresh_open_related_windows()

    def open_maintenance(self) -> None:
        self.cancel_period_data_queries()
        dialog = CustomerMaintenanceDialog(self.repository, self)
        dialog.exec()
        self.invalidate_customer_caches()
        self.refresh_all()

    def _offer_vacuum_after_last_period_delete(self) -> None:
        status = self.repository.maintenance_status()
        message = QMessageBox(self)
        message.setWindowTitle("Thu hồi dung lượng Customer.db")
        message.setIcon(QMessageBox.Icon.Question)
        message.setText("Đã xóa kỳ dữ liệu cuối cùng khỏi Customer.db.")
        message.setInformativeText(
            "SQLite chỉ đánh dấu trang trống sau DELETE; dung lượng file chỉ giảm khi chạy VACUUM. "
            f"Dung lượng hiện tại: {status.size_bytes:,} bytes; "
            f"có thể thu hồi khoảng {status.reclaimable_bytes:,} bytes."
        )
        reclaim_button = message.addButton("Thu hồi dung lượng", QMessageBox.ButtonRole.AcceptRole)
        message.addButton("Để sau", QMessageBox.ButtonRole.RejectRole)
        message.exec()
        if message.clickedButton() is reclaim_button:
            self._run_vacuum_after_last_period_delete()

    def _run_vacuum_after_last_period_delete(self) -> None:
        progress = QProgressDialog("Đang thu hồi dung lượng Customer.db...", "", 0, 0, self)
        progress.setWindowTitle("Thu hồi dung lượng Customer.db")
        progress.setCancelButton(None)
        progress.setMinimumDuration(0)
        progress.setWindowModality(Qt.WindowModality.WindowModal)
        progress.show()

        def task():
            return self.repository.optimize_database(vacuum=True)

        def done(result):
            progress.close()
            self.invalidate_customer_caches()
            self.refresh_all()
            QMessageBox.information(
                self,
                "Thu hồi dung lượng Customer.db",
                "Đã thu hồi dung lượng Customer.db.\n"
                f"Trước: {int(result.get('before_size_bytes', 0)):,} bytes\n"
                f"Sau: {int(result.get('after_size_bytes', 0)):,} bytes",
            )

        def failed(exc: Exception):
            progress.close()
            QMessageBox.warning(self, "Thu hồi dung lượng Customer.db", str(exc))

        self._vacuum_thread = run_in_thread(self, task, done, failed)

    def select_tab(self, tab: str | int) -> None:
        if isinstance(tab, int):
            if 0 <= tab < self.tabs.count():
                self.tabs.setCurrentIndex(tab)
            return
        key = str(tab or "").strip().casefold()
        mapping = {
            "dashboard": self.dashboard_tab,
            "debt_group": self.debt_group_tab,
            "debt-group": self.debt_group_tab,
            "nhom_no": self.debt_group_tab,
            "list": self.list_tab,
            "customers": self.list_tab,
            "movement": self.movement_tab,
            "comparison": self.movement_tab,
            "multiple": self.multiple_tab,
            "cross_branch": self.cross_branch_tab,
            "cross-branch": self.cross_branch_tab,
            "lien_chi_nhanh": self.cross_branch_tab,
            "import": self.import_tab,
            "officer": self.officer_tab,
        }
        widget = mapping.get(key)
        if widget is not None:
            self.tabs.setCurrentWidget(widget)
            self._refresh_active_tab()

    def _filter_changed(self) -> None:
        if self._updating_filters:
            return
        self._filter_timer.start()

    def _apply_filter_changed(self) -> None:
        self.list_tab.page = 1
        self.debt_group_tab.branch_page = 1
        self.debt_group_tab.officer_page = 1
        self.debt_group_tab.customer_page = 1
        self.movement_tab.page = 1
        self.multiple_tab.page = 1
        self.cross_branch_tab.page = 1
        periods = self.refresh_filters()
        if not periods:
            self.handle_customer_data_became_empty()
            return
        self._refresh_active_tab()

    def _refresh_active_tab(self) -> None:
        if not self._has_period_data and self.tabs.currentWidget() not in {self.import_tab, self.officer_tab}:
            self.handle_customer_data_became_empty()
            return
        tab = self.tabs.currentWidget()
        if hasattr(tab, "refresh"):
            tab.refresh()

    def invalidate_customer_caches(self) -> None:
        for tab in (
            self.dashboard_tab,
            self.debt_group_tab,
            self.list_tab,
            self.movement_tab,
            self.multiple_tab,
            self.cross_branch_tab,
            self.import_tab,
            self.officer_tab,
        ):
            if hasattr(tab, "invalidate_cache"):
                tab.invalidate_cache()

    def cancel_period_data_queries(self) -> None:
        self._filter_timer.stop()
        for tab in (
            self.dashboard_tab,
            self.debt_group_tab,
            self.list_tab,
            self.movement_tab,
            self.multiple_tab,
            self.cross_branch_tab,
        ):
            if hasattr(tab, "cancel_queries"):
                tab.cancel_queries()

    def wait_for_all_queries(self) -> None:
        for tab in (
            self.dashboard_tab,
            self.debt_group_tab,
            self.list_tab,
            self.movement_tab,
            self.multiple_tab,
            self.cross_branch_tab,
            self.import_tab,
            self.officer_tab,
        ):
            if hasattr(tab, "wait_for_queries"):
                tab.wait_for_queries()

    def closeEvent(self, event) -> None:
        self.cancel_period_data_queries()
        self.import_tab.cancel_queries()
        self.officer_tab.cancel_queries()
        self.repository.unit_directory.remove_listener(self._unit_directory_changed)
        super().closeEvent(event)

    def _build_empty_data_banner(self) -> QWidget:
        banner = QWidget()
        banner.setObjectName("CustomerEmptyStateBanner")
        layout = QHBoxLayout(banner)
        layout.setContentsMargins(10, 8, 10, 8)
        self.empty_data_message = QLabel("Chưa có dữ liệu khách hàng. Hãy vào NIM dư nợ và import thư mục FTP Loan để tạo dữ liệu.")
        self.empty_data_message.setObjectName("MutedText")
        self.empty_data_message.setWordWrap(True)
        open_nim_button = secondary_button("Mở NIM dư nợ")
        open_nim_button.clicked.connect(self.openNimDnRequested.emit)
        layout.addWidget(self.empty_data_message, stretch=1)
        layout.addWidget(open_nim_button)
        banner.hide()
        return banner

    def _update_empty_data_banner(self) -> None:
        has_data = bool(self._available_periods) if self._available_periods is not None else self.repository.has_period_data()
        if not has_data:
            if self.repository.officer_directory_count():
                self.empty_data_message.setText("Chưa có dữ liệu khách hàng theo kỳ. Danh mục CBTD vẫn được giữ lại.")
            else:
                self.empty_data_message.setText(
                    "Chưa có dữ liệu khách hàng. Hãy vào NIM dư nợ và import thư mục FTP Loan để tạo dữ liệu."
                )
        self.empty_data_banner.setVisible(not has_data)
        self.export_all_button.setEnabled(has_data)
        self.export_all_button.setToolTip("" if has_data else "Chưa có dữ liệu để xuất")
        self.delete_period_button.setEnabled(has_data)
        self.delete_period_button.setToolTip("" if has_data else "Chưa có dữ liệu kỳ để xóa")
        self.maintenance_button.setEnabled(True)
        if not has_data:
            LOGGER.debug("Customer empty state entered")

    def _normalize_period_combos(self, periods: list[str]) -> None:
        if not periods:
            return
        for combo in (self.period_from_combo, self.period_to_combo, self.current_period_combo):
            combo.setEnabled(True)
        period_set = set(periods)
        period_from = current_data(self.period_from_combo)
        period_to = current_data(self.period_to_combo)
        report_period = current_data(self.current_period_combo)
        validation = validate_dashboard_period_filters(
            periods,
            CustomerFilters(period_from=period_from, period_to=period_to, current_period=report_period),
        )
        period_from = validation.period_from
        period_to = validation.period_to
        report_period = validation.report_period
        LOGGER.debug(
            "Customer selected period range: from=%s to=%s report=%s valid=%s",
            period_from,
            period_to,
            report_period,
            validation.valid,
        )
        self._set_combo_current_data(self.period_from_combo, period_from)
        self._set_combo_current_data(self.period_to_combo, period_to)
        self._set_combo_current_data(self.current_period_combo, report_period)

    def handle_customer_data_became_empty(self) -> None:
        self._has_period_data = False
        self._available_periods = []
        self.cancel_period_data_queries()
        self._update_empty_data_banner()
        for tab in (
            self.dashboard_tab,
            self.debt_group_tab,
            self.list_tab,
            self.movement_tab,
            self.multiple_tab,
            self.cross_branch_tab,
        ):
            if hasattr(tab, "set_empty_state"):
                tab.set_empty_state()

    @staticmethod
    def _populate_period_combo(combo: QComboBox, first_label: str, periods: list[str]) -> None:
        combo.blockSignals(True)
        try:
            current = combo.currentData()
            combo.clear()
            combo.addItem(first_label, "")
            for period in periods:
                combo.addItem(period, period)
            index = combo.findData(current)
            combo.setCurrentIndex(index if index >= 0 else 0)
            combo.setEnabled(True)
        finally:
            combo.blockSignals(False)

    def _reset_period_combos_empty(self) -> None:
        for combo in (self.period_from_combo, self.period_to_combo, self.current_period_combo):
            combo.blockSignals(True)
            try:
                combo.clear()
                combo.addItem("Chưa có dữ liệu", "")
                combo.setCurrentIndex(0)
                combo.setEnabled(False)
            finally:
                combo.blockSignals(False)

    @staticmethod
    def _set_combo_current_data(combo: QComboBox, value: str) -> None:
        index = combo.findData(value)
        if index >= 0 and combo.currentIndex() != index:
            combo.setCurrentIndex(index)

    def _set_period_combo_tooltips(self) -> None:
        self.period_from_combo.setToolTip("Từ kỳ dùng cho các biểu đồ xu hướng nhiều kỳ.")
        self.period_to_combo.setToolTip("Đến kỳ dùng cho các biểu đồ xu hướng nhiều kỳ.")
        self.current_period_combo.setToolTip("Kỳ báo cáo dùng cho KPI, cơ cấu kỳ hạn và các bảng Top.")

    def _refresh_open_related_windows(self) -> None:
        seen: set[int] = set()
        for owner in self._owner_chain():
            center = getattr(owner, "_officer_center_window", None)
            self._refresh_related_window(center, seen)
            registry = getattr(owner, "_shared_officer_detail_windows", None)
            if isinstance(registry, dict):
                for window in list(registry.values()):
                    self._refresh_related_window(window, seen)

    def _owner_chain(self):
        current = self
        while current is not None:
            yield current
            try:
                current = current.parent()
            except RuntimeError:
                return

    @staticmethod
    def _refresh_related_window(window, seen: set[int]) -> None:
        if window is None:
            return
        marker = id(window)
        if marker in seen:
            return
        seen.add(marker)
        try:
            if hasattr(window, "invalidate_cache"):
                window.invalidate_cache()
            if hasattr(window, "refresh_all"):
                window.refresh_all(use_cache=False)
            elif hasattr(window, "refresh"):
                window.refresh(use_cache=False)
            elif hasattr(window, "reload"):
                window.reload()
        except RuntimeError:
            return
