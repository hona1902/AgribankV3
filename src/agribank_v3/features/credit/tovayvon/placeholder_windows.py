from __future__ import annotations

import json
from pathlib import Path

from PySide6.QtCore import Qt, QUrl
from PySide6.QtGui import QDesktopServices
from PySide6.QtWidgets import (
    QApplication,
    QFileDialog,
    QComboBox,
    QDialog,
    QDialogButtonBox,
    QFrame,
    QGridLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QScrollArea,
    QTabWidget,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from agribank_v3.features.credit.tovayvon.models import (
    CreditGroup,
    CreditGroupCommissionRate,
    CreditCommissionRuleSettings,
    DATA_TVV_ATTR_TO_HEADER,
    DATA_TVV_FIELD_LABELS,
)
from agribank_v3.features.credit.tovayvon.excel_templates import (
    create_data_tvv_template,
)
from agribank_v3.features.credit.tovayvon.repository import (
    CreditGroupRepository,
    CreditGroupRepositoryError,
)
from agribank_v3.runtime_paths import application_root
from agribank_v3.settings import AppSettingsDatabase, SettingsDatabaseError


CREDIT_GROUP_MANAGEMENT_TITLE = "Quản lý tổ vay vốn"

CREDIT_GROUP_MANAGEMENT_ROUTE_TITLES: frozenset[str] = frozenset(
    {
        CREDIT_GROUP_MANAGEMENT_TITLE,
        "Danh sách tổ vay vốn",
        "Import/Export dữ liệu tổ vay vốn",
        "Cấu hình hoa hồng tổ vay vốn",
    }
)


CREDIT_TOVAYVON_PLACEHOLDER_TITLES: frozenset[str] = frozenset(
    {
        "Bảng kê thu lãi tổ vay vốn",
        "Đề nghị thanh toán hoa hồng tổ vay vốn",
        "Đối chiếu dư nợ theo tổ vay vốn",
        "Hướng dẫn tổ vay vốn",
    }
)


class CreditGroupManagementPlaceholderDialog(QDialog):
    """Temporary central screen for the Tổ vay vốn data management migration."""

    def __init__(
        self,
        parent: QWidget | None = None,
        database_path: Path | None = None,
    ) -> None:
        super().__init__(parent)
        self.setWindowTitle(CREDIT_GROUP_MANAGEMENT_TITLE)
        self.setModal(True)
        self.setMinimumSize(950, 720)
        self._apply_initial_window_size()
        self.groups: list[CreditGroup] = []
        self.filtered_groups: list[CreditGroup] = []
        self.current_rate: CreditGroupCommissionRate | None = None
        self.commission_inputs: dict[str, QLineEdit] = {}
        self.rule_inputs: dict[str, QLineEdit] = {}
        self.repository_error: str | None = None
        self.repository: CreditGroupRepository | None = None
        self._loading_rate = False
        self._loading_rules = False
        self._stt_normalized = False

        resolved_database_path = self._resolve_database_path(parent, database_path)
        try:
            self.repository = CreditGroupRepository(resolved_database_path)
        except CreditGroupRepositoryError as exc:
            self.repository_error = str(exc)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(18, 16, 18, 14)
        layout.setSpacing(12)

        title = QLabel(CREDIT_GROUP_MANAGEMENT_TITLE)
        title.setObjectName("PageTitle")
        layout.addWidget(title)

        action_bar = QFrame()
        action_layout = QHBoxLayout(action_bar)
        action_layout.setContentsMargins(0, 0, 0, 0)
        action_layout.setSpacing(8)

        add_button = QPushButton("Thêm mới")
        add_button.setObjectName("SecondaryButton")
        add_button.clicked.connect(self._add_group)
        action_layout.addWidget(add_button)

        edit_button = QPushButton("Sửa")
        edit_button.setObjectName("SecondaryButton")
        edit_button.clicked.connect(self._edit_selected_group)
        action_layout.addWidget(edit_button)

        delete_button = QPushButton("Xóa / Ngừng sử dụng")
        delete_button.setObjectName("SecondaryButton")
        delete_button.clicked.connect(self._show_not_migrated_notice)
        action_layout.addWidget(delete_button)

        refresh_button = QPushButton("Làm mới")
        refresh_button.setObjectName("SecondaryButton")
        refresh_button.clicked.connect(self._load_groups)
        action_layout.addWidget(refresh_button)
        action_layout.addStretch()
        layout.addWidget(action_bar)

        self.tabs = QTabWidget()
        self.groups_table = QTableWidget(0, 6)
        self.tabs.addTab(self._groups_tab(), "Danh sách tổ vay vốn")
        self.tabs.addTab(self._commission_tab(), "Tỷ lệ hoa hồng")
        layout.addWidget(self.tabs, stretch=1)

        utility_bar = QFrame()
        utility_layout = QHBoxLayout(utility_bar)
        utility_layout.setContentsMargins(0, 0, 0, 0)
        utility_layout.setSpacing(8)
        utility_layout.addStretch()

        template_button = QPushButton("Tải file Excel mẫu")
        template_button.setObjectName("SecondaryButton")
        template_button.setToolTip(
            "Tải file Excel mẫu đúng cấu trúc Data_TVV để nhập danh sách tổ vay vốn."
        )
        template_button.clicked.connect(self._create_data_tvv_template)
        utility_layout.addWidget(template_button)

        import_button = QPushButton("Import Excel")
        import_button.setObjectName("SecondaryButton")
        import_button.clicked.connect(self._import_data_tvv)
        utility_layout.addWidget(import_button)

        export_button = QPushButton("Export Excel")
        export_button.setObjectName("SecondaryButton")
        export_button.clicked.connect(self._export_data_tvv)
        utility_layout.addWidget(export_button)

        close_button = QPushButton("Đóng")
        close_button.setObjectName("PrimaryButton")
        close_button.clicked.connect(self.reject)
        utility_layout.addWidget(close_button)
        layout.addWidget(utility_bar)
        self._load_groups()
        self._load_rule_settings()

    def _apply_initial_window_size(self) -> None:
        screen = QApplication.primaryScreen()
        if screen is None:
            self.resize(1100, 820)
            return
        available = screen.availableGeometry()
        width = min(1100, max(950, available.width() - 80))
        height = min(820, max(720, available.height() - 80))
        self.resize(width, height)
        frame = self.frameGeometry()
        frame.moveCenter(available.center())
        self.move(frame.topLeft())

    @staticmethod
    def _resolve_database_path(
        parent: QWidget | None,
        database_path: Path | None,
    ) -> Path:
        if database_path is not None:
            return Path(database_path)
        parent_database = getattr(parent, "settings_database", None)
        parent_path = getattr(parent_database, "database_path", None)
        if parent_path is not None:
            return Path(parent_path)
        return application_root() / "data" / "DuLieuV3.db"

    def _groups_tab(self) -> QWidget:
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(8)

        search_layout = QHBoxLayout()
        search_layout.setContentsMargins(0, 0, 0, 0)
        search_layout.setSpacing(8)
        self.search_edit = QLineEdit()
        self.search_edit.setPlaceholderText(
            "Tìm theo mã tổ, tên tổ, tổ trưởng, xã, tổ hội..."
        )
        self.search_edit.returnPressed.connect(self._apply_group_filter)
        self.search_edit.textChanged.connect(self._apply_group_filter)
        search_layout.addWidget(self.search_edit, stretch=1)

        search_button = QPushButton("Tìm")
        search_button.setObjectName("SecondaryButton")
        search_button.clicked.connect(self._apply_group_filter)
        search_layout.addWidget(search_button)

        clear_search_button = QPushButton("Xóa tìm kiếm")
        clear_search_button.setObjectName("SecondaryButton")
        clear_search_button.clicked.connect(self._clear_group_filter)
        search_layout.addWidget(clear_search_button)
        layout.addLayout(search_layout)

        self.groups_table.setHorizontalHeaderLabels(
            ["STT", "MaTo", "Tên tổ", "Tổ trưởng", "Số điện thoại", "Xã"]
        )
        self.groups_table.horizontalHeaderItem(0).setTextAlignment(
            Qt.AlignmentFlag.AlignCenter
        )
        self.groups_table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.groups_table.setSelectionBehavior(
            QTableWidget.SelectionBehavior.SelectRows
        )
        self.groups_table.setHorizontalScrollBarPolicy(
            Qt.ScrollBarPolicy.ScrollBarAlwaysOff
        )
        self.groups_table.verticalHeader().setVisible(False)
        self.groups_table.verticalHeader().setDefaultSectionSize(30)
        self.groups_table.horizontalHeader().setStretchLastSection(True)
        self.groups_table.cellDoubleClicked.connect(self._open_group_detail_from_row)
        layout.addWidget(self.groups_table, stretch=1)
        return page

    def _commission_tab(self) -> QWidget:
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(10)

        commission_tabs = QTabWidget()
        commission_tabs.addTab(self._commission_rate_tab(), "Tỷ lệ hoa hồng")
        commission_tabs.addTab(
            self._commission_rule_settings_tab(),
            "Điều kiện chi hoa hồng",
        )
        layout.addWidget(commission_tabs, stretch=1)
        return page

    def _commission_rate_tab(self) -> QWidget:
        page = QScrollArea()
        page.setWidgetResizable(True)
        page.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        content = QWidget()
        page.setWidget(content)
        layout = QVBoxLayout(content)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(10)

        section_title = QLabel("Tỷ lệ hoa hồng theo tổ")
        section_title.setObjectName("SectionTitle")
        layout.addWidget(section_title)

        search_layout = QHBoxLayout()
        search_layout.setContentsMargins(0, 0, 0, 0)
        search_layout.setSpacing(8)
        search_layout.addWidget(QLabel("Tìm tổ vay vốn"))
        self.commission_group_filter_edit = QLineEdit()
        self.commission_group_filter_edit.setPlaceholderText(
            "Nhập mã tổ, tên tổ, tổ trưởng, xã..."
        )
        self.commission_group_filter_edit.returnPressed.connect(
            self._filter_commission_group_options
        )
        search_layout.addWidget(self.commission_group_filter_edit, stretch=1)

        filter_button = QPushButton("Tìm")
        filter_button.setObjectName("SecondaryButton")
        filter_button.clicked.connect(self._filter_commission_group_options)
        search_layout.addWidget(filter_button)

        clear_filter_button = QPushButton("Xóa tìm kiếm")
        clear_filter_button.setObjectName("SecondaryButton")
        clear_filter_button.clicked.connect(self._clear_commission_group_filter)
        search_layout.addWidget(clear_filter_button)
        layout.addLayout(search_layout)

        selector_layout = QHBoxLayout()
        selector_layout.setContentsMargins(0, 0, 0, 0)
        selector_layout.setSpacing(8)
        selector_layout.addWidget(QLabel("Tổ vay vốn"))
        self.group_combo = QComboBox()
        self.group_combo.currentIndexChanged.connect(self._load_selected_rate)
        selector_layout.addWidget(self.group_combo, stretch=1)
        layout.addLayout(selector_layout)

        self.group_summary = QLabel()
        self.group_summary.setObjectName("MutedText")
        self.group_summary.setWordWrap(True)
        layout.addWidget(self.group_summary)

        self.commission_filter_status = QLabel()
        self.commission_filter_status.setObjectName("MutedText")
        self.commission_filter_status.setWordWrap(True)
        layout.addWidget(self.commission_filter_status)

        rates_layout = QGridLayout()
        rates_layout.setContentsMargins(0, 0, 0, 0)
        rates_layout.setHorizontalSpacing(12)
        rates_layout.setVerticalSpacing(8)
        rates_layout.addWidget(
            self._rate_group_box(
                "Hoa hồng không BĐ",
                (
                    ("Tổ trưởng (%)", "no_secured_to_truong"),
                    ("Cấp xã (%)", "no_secured_cap_xa"),
                    ("Cấp huyện (%)", "no_secured_cap_huyen"),
                    ("Cấp tỉnh (%)", "no_secured_cap_tinh"),
                    ("Cấp TW (%)", "no_secured_cap_tw"),
                ),
                "no_secured_total_label",
            ),
            0,
            0,
        )
        rates_layout.addWidget(
            self._rate_group_box(
                "Hoa hồng có BĐTS",
                (
                    ("Tổ trưởng (%)", "secured_to_truong"),
                    ("Cấp xã (%)", "secured_cap_xa"),
                    ("Cấp huyện (%)", "secured_cap_huyen"),
                    ("Cấp tỉnh (%)", "secured_cap_tinh"),
                    ("Cấp TW (%)", "secured_cap_tw"),
                ),
                "secured_total_label",
            ),
            0,
            1,
        )
        rates_layout.setColumnStretch(0, 1)
        rates_layout.setColumnStretch(1, 1)
        layout.addLayout(rates_layout)

        actions = QHBoxLayout()
        actions.addStretch()
        reset_button = QPushButton("Khôi phục mặc định")
        reset_button.setObjectName("SecondaryButton")
        reset_button.clicked.connect(self._reset_selected_rate)
        actions.addWidget(reset_button)
        save_button = QPushButton("Lưu tỷ lệ")
        save_button.setObjectName("PrimaryButton")
        save_button.clicked.connect(self._save_selected_rate)
        actions.addWidget(save_button)
        layout.addLayout(actions)

        self.rate_status = QLabel()
        self.rate_status.setObjectName("MutedText")
        self.rate_status.setWordWrap(True)
        layout.addWidget(self.rate_status)

        layout.addStretch()
        return page

    def _commission_rule_settings_tab(self) -> QWidget:
        page = QScrollArea()
        page.setWidgetResizable(True)
        page.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        content = QWidget()
        page.setWidget(content)
        layout = QVBoxLayout(content)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(10)

        settings_title = QLabel("Cài đặt tỷ lệ")
        settings_title.setObjectName("SectionTitle")
        layout.addWidget(settings_title)
        layout.addWidget(self._commission_rule_settings_box())
        layout.addStretch()
        return page

    def _commission_rule_settings_box(self) -> QGroupBox:
        box = QGroupBox("Cài đặt điều kiện chi hoa hồng")
        box.setMinimumHeight(300)
        layout = QVBoxLayout(box)
        layout.setSpacing(12)

        interest_box = QGroupBox("Chi theo tỷ lệ thu lãi")
        interest_grid = QGridLayout(interest_box)
        interest_grid.setContentsMargins(12, 14, 12, 12)
        interest_grid.setHorizontalSpacing(8)
        interest_grid.setVerticalSpacing(10)
        for column, header in enumerate(
            ("Khoảng", "Min thu lãi (%)", "Max thu lãi (%)", "Tỷ lệ chi (%)")
        ):
            label = QLabel(header)
            label.setStyleSheet("font-weight: 600;")
            interest_grid.addWidget(label, 0, column)

        rows = (
            ("Mức 1", "interest_min_1", "interest_max_1", "interest_pay_1"),
            ("Mức 2", "interest_min_2", "interest_max_2", "interest_pay_2"),
            ("Mức 3", "interest_min_3", None, "interest_pay_3"),
        )
        for row_index, (label, min_field, max_field, pay_field) in enumerate(rows, start=1):
            interest_grid.setRowMinimumHeight(row_index, 34)
            interest_grid.addWidget(QLabel(label), row_index, 0)
            interest_grid.addWidget(self._rule_input(min_field), row_index, 1)
            if max_field is None:
                max_label = QLabel("Trở lên")
                max_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
                interest_grid.addWidget(max_label, row_index, 2)
            else:
                interest_grid.addWidget(self._rule_input(max_field), row_index, 2)
            interest_grid.addWidget(self._rule_input(pay_field), row_index, 3)
        layout.addWidget(interest_box)

        bad_debt_box = QGroupBox("Điều kiện nợ xấu")
        bad_debt_layout = QHBoxLayout(bad_debt_box)
        bad_debt_layout.setContentsMargins(12, 14, 12, 12)
        bad_debt_layout.setSpacing(8)
        bad_debt_layout.addWidget(QLabel("Ngưỡng nợ xấu tối đa (%)"))
        bad_debt_layout.addWidget(self._rule_input("bad_debt_threshold"))
        bad_debt_layout.addWidget(QLabel("Tỷ lệ chi khi vượt ngưỡng (%)"))
        bad_debt_layout.addWidget(self._rule_input("bad_debt_pay"))
        layout.addWidget(bad_debt_box)

        button_layout = QHBoxLayout()
        button_layout.addStretch()
        reload_button = QPushButton("Làm mới cấu hình")
        reload_button.setObjectName("SecondaryButton")
        reload_button.clicked.connect(self._load_rule_settings)
        button_layout.addWidget(reload_button)
        reset_button = QPushButton("Khôi phục mặc định")
        reset_button.setObjectName("SecondaryButton")
        reset_button.clicked.connect(self._reset_rule_settings)
        button_layout.addWidget(reset_button)
        save_button = QPushButton("Lưu điều kiện")
        save_button.setObjectName("PrimaryButton")
        save_button.clicked.connect(self._save_rule_settings)
        button_layout.addWidget(save_button)
        layout.addLayout(button_layout)

        self.rule_status = QLabel()
        self.rule_status.setObjectName("MutedText")
        self.rule_status.setWordWrap(True)
        layout.addWidget(self.rule_status)
        return box

    def _rule_input(self, field_name: str) -> QLineEdit:
        edit = QLineEdit()
        edit.setAlignment(Qt.AlignmentFlag.AlignRight)
        edit.setPlaceholderText("0")
        edit.setFixedWidth(96)
        edit.setMinimumHeight(30)
        edit.textChanged.connect(self._update_rule_status)
        self.rule_inputs[field_name] = edit
        return edit

    def _rate_group_box(
        self,
        title: str,
        fields: tuple[tuple[str, str], ...],
        total_attr: str,
    ) -> QGroupBox:
        box = QGroupBox(title)
        box.setMinimumHeight(300)
        grid = QGridLayout(box)
        grid.setContentsMargins(14, 18, 14, 16)
        grid.setHorizontalSpacing(8)
        grid.setVerticalSpacing(10)
        grid.setColumnMinimumWidth(0, 150)
        grid.setColumnMinimumWidth(1, 96)
        grid.setColumnMinimumWidth(2, 18)
        for row_index, (label, field_name) in enumerate(fields):
            grid.setRowMinimumHeight(row_index, 34)
            label_widget = QLabel(label)
            label_widget.setMinimumWidth(150)
            grid.addWidget(label_widget, row_index, 0, alignment=Qt.AlignmentFlag.AlignVCenter)
            edit = QLineEdit()
            edit.setAlignment(Qt.AlignmentFlag.AlignRight)
            edit.setPlaceholderText("0")
            edit.setFixedWidth(96)
            edit.setMinimumHeight(30)
            edit.textChanged.connect(self._update_rate_totals)
            self.commission_inputs[field_name] = edit
            grid.addWidget(edit, row_index, 1, alignment=Qt.AlignmentFlag.AlignVCenter)
            percent_label = QLabel("%")
            percent_label.setFixedWidth(18)
            percent_label.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
            grid.addWidget(percent_label, row_index, 2)
        total_row = len(fields)
        grid.setRowMinimumHeight(total_row, 34)
        total_title = QLabel("Tổng")
        total_title.setMinimumWidth(150)
        grid.addWidget(total_title, total_row, 0, alignment=Qt.AlignmentFlag.AlignVCenter)
        total_label = QLineEdit("0")
        total_label.setReadOnly(True)
        total_label.setMinimumHeight(30)
        total_label.setFixedWidth(96)
        total_label.setAlignment(Qt.AlignmentFlag.AlignRight)
        setattr(self, total_attr, total_label)
        grid.addWidget(total_label, total_row, 1, alignment=Qt.AlignmentFlag.AlignVCenter)
        total_percent_label = QLabel("%")
        total_percent_label.setFixedWidth(18)
        total_percent_label.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
        grid.addWidget(total_percent_label, total_row, 2)
        return box

    @staticmethod
    def _placeholder_tab(message: str) -> QWidget:
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(8)

        label = QLabel(message)
        label.setObjectName("MutedText")
        label.setWordWrap(True)
        label.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignTop)
        layout.addWidget(label)
        layout.addStretch()
        return page

    def _show_not_migrated_notice(self) -> None:
        QMessageBox.information(
            self,
            CREDIT_GROUP_MANAGEMENT_TITLE,
            "Chức năng xóa/ngừng sử dụng tổ vay vốn sẽ được chuyển ở bước tiếp theo.",
        )

    def _load_groups(self) -> None:
        if self.repository is None:
            self._set_repository_error()
            return
        try:
            if not self._stt_normalized:
                self.repository.resequence_group_stt()
                self._stt_normalized = True
            self.groups = self.repository.list_groups()
        except CreditGroupRepositoryError as exc:
            QMessageBox.warning(self, CREDIT_GROUP_MANAGEMENT_TITLE, str(exc))
            self.groups = []
        self._apply_group_filter()
        self._populate_group_combo()

    def _set_repository_error(self) -> None:
        message = self.repository_error or "Không thể mở dữ liệu tổ vay vốn."
        self.group_summary.setText(message)
        self.rate_status.setText(message)
        if hasattr(self, "rule_status"):
            self.rule_status.setText(message)

    def _populate_groups_table(self) -> None:
        self.groups_table.setRowCount(len(self.filtered_groups))
        for row_index, group in enumerate(self.filtered_groups):
            values = (
                group.stt,
                group.ma_to,
                group.ten_to,
                group.ten_to_truong,
                group.so_dien_thoai,
                group.xa,
            )
            for column, value in enumerate(values):
                item = QTableWidgetItem(str(value))
                if column == 0:
                    item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                else:
                    item.setTextAlignment(
                        Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter
                    )
                self.groups_table.setItem(row_index, column, item)
        self.groups_table.resizeColumnsToContents()

    def _apply_group_filter(self) -> None:
        keyword = self.search_edit.text().strip().casefold()
        if not keyword:
            self.filtered_groups = list(self.groups)
        else:
            self.filtered_groups = [
                group
                for group in self.groups
                if self._group_matches_keyword(group, keyword)
            ]
        self._populate_groups_table()

    def _clear_group_filter(self) -> None:
        self.search_edit.clear()
        self._apply_group_filter()

    @staticmethod
    def _group_matches_keyword(group: CreditGroup, keyword: str) -> bool:
        fields = (
            group.ma_to,
            group.ten_to,
            group.ten_tvv_day_du,
            group.ten_to_truong,
            group.xa,
            group.dia_chi,
            group.to_hoi,
            group.to_chuc,
        )
        return any(keyword in str(value).casefold() for value in fields)

    def _open_group_detail_from_row(self, row: int, column: int) -> None:
        del column
        if row < 0 or row >= len(self.filtered_groups):
            return
        self._edit_group(self.filtered_groups[row])

    def _selected_table_group(self) -> CreditGroup | None:
        selected_rows = self.groups_table.selectionModel().selectedRows()
        if not selected_rows:
            return None
        row = selected_rows[0].row()
        if row < 0 or row >= len(self.filtered_groups):
            return None
        return self.filtered_groups[row]

    def _add_group(self) -> None:
        if self.repository is None:
            self._set_repository_error()
            return
        dialog = CreditGroupEditDialog(
            repository=self.repository,
            mode="add",
            parent=self,
        )
        if dialog.exec() == QDialog.DialogCode.Accepted:
            self._load_groups()
            if dialog.saved_ma_to:
                self._select_commission_group(dialog.saved_ma_to)
            QMessageBox.information(
                self,
                CREDIT_GROUP_MANAGEMENT_TITLE,
                "Đã thêm mới tổ vay vốn.",
            )

    def _edit_selected_group(self) -> None:
        group = self._selected_table_group()
        if group is None:
            QMessageBox.information(
                self,
                CREDIT_GROUP_MANAGEMENT_TITLE,
                "Vui lòng chọn một tổ vay vốn để sửa.",
            )
            return
        self._edit_group(group)

    def _edit_group(self, group: CreditGroup) -> None:
        if self.repository is None:
            self._set_repository_error()
            return
        dialog = CreditGroupEditDialog(
            repository=self.repository,
            mode="edit",
            group=group,
            parent=self,
        )
        if dialog.exec() == QDialog.DialogCode.Accepted:
            self._load_groups()
            self._select_commission_group(group.ma_to)
            QMessageBox.information(
                self,
                CREDIT_GROUP_MANAGEMENT_TITLE,
                "Đã cập nhật thông tin tổ vay vốn.",
            )

    def _populate_group_combo(self) -> None:
        keyword = ""
        if hasattr(self, "commission_group_filter_edit"):
            keyword = self.commission_group_filter_edit.text().strip()
        self.refresh_commission_group_dropdown(
            self.filter_commission_group_options(keyword)
        )

    def filter_commission_group_options(self, keyword: str) -> list[CreditGroup]:
        normalized = keyword.strip().casefold()
        if not normalized:
            return list(self.groups)
        return [
            group
            for group in self.groups
            if self._group_matches_keyword(group, normalized)
        ]

    def refresh_commission_group_dropdown(self, groups: list[CreditGroup]) -> None:
        current_ma_to = self._selected_ma_to() if self.group_combo.count() else ""
        self.group_combo.blockSignals(True)
        self.group_combo.clear()
        for group in groups:
            display = f"{group.ma_to} - {group.ten_to or group.ten_tvv_day_du}"
            self.group_combo.addItem(display, group.ma_to)
        if current_ma_to:
            index = self.group_combo.findData(current_ma_to)
            if index >= 0:
                self.group_combo.setCurrentIndex(index)
        elif len(groups) == 1:
            self.group_combo.setCurrentIndex(0)
        self.group_combo.blockSignals(False)
        if hasattr(self, "commission_filter_status"):
            self.commission_filter_status.setText(
                "Không tìm thấy tổ vay vốn phù hợp."
                if not groups and self.commission_group_filter_edit.text().strip()
                else ""
            )
        new_ma_to = self._selected_ma_to()
        if not current_ma_to or new_ma_to != current_ma_to:
            self._load_selected_rate()

    def _filter_commission_group_options(self) -> None:
        self.refresh_commission_group_dropdown(
            self.filter_commission_group_options(
                self.commission_group_filter_edit.text()
            )
        )

    def _clear_commission_group_filter(self) -> None:
        self.commission_group_filter_edit.clear()
        self.refresh_commission_group_dropdown(list(self.groups))

    def _select_commission_group(self, ma_to: str) -> None:
        index = self.group_combo.findData(ma_to)
        if index >= 0:
            self.group_combo.setCurrentIndex(index)

    def _selected_ma_to(self) -> str:
        return str(self.group_combo.currentData() or "").strip()

    def _selected_group(self) -> CreditGroup | None:
        selected_ma_to = self._selected_ma_to()
        return next((group for group in self.groups if group.ma_to == selected_ma_to), None)

    def _load_selected_rate(self) -> None:
        if self.repository is None:
            self._set_repository_error()
            return
        selected_group = self._selected_group()
        if selected_group is None:
            self.current_rate = None
            self.group_summary.setText(
                "Chưa có tổ vay vốn. Hãy import file Data_TVV trước khi cấu hình hoa hồng."
            )
            self._set_rate_inputs_enabled(False)
            self._clear_rate_inputs()
            self._update_rate_totals()
            return
        try:
            self.current_rate = self.repository.get_or_create_commission_rate(
                selected_group.ma_to
            )
        except CreditGroupRepositoryError as exc:
            QMessageBox.warning(self, CREDIT_GROUP_MANAGEMENT_TITLE, str(exc))
            return
        self.group_summary.setText(
            f"MaTo: {selected_group.ma_to} | Tên tổ: {selected_group.ten_to} | "
            f"Tổ trưởng: {selected_group.ten_to_truong} | Xã: {selected_group.xa}"
        )
        self._set_rate_inputs_enabled(True)
        self._set_rate_inputs(self.current_rate)

    def _set_rate_inputs_enabled(self, enabled: bool) -> None:
        for edit in self.commission_inputs.values():
            edit.setEnabled(enabled)

    def _clear_rate_inputs(self) -> None:
        self._loading_rate = True
        try:
            for edit in self.commission_inputs.values():
                edit.clear()
        finally:
            self._loading_rate = False

    def _set_rate_inputs(self, rate: CreditGroupCommissionRate) -> None:
        self._loading_rate = True
        try:
            for field_name, edit in self.commission_inputs.items():
                edit.setText(self._format_percent_value(getattr(rate, field_name)))
        finally:
            self._loading_rate = False
        self._update_rate_totals()

    @staticmethod
    def _format_percent_value(value: float) -> str:
        formatted = f"{float(value):.4f}".rstrip("0").rstrip(".")
        return formatted or "0"

    def _rate_from_inputs(self) -> CreditGroupCommissionRate:
        ma_to = self._selected_ma_to()
        values = {
            field_name: self._parse_percent(edit.text(), field_name)
            for field_name, edit in self.commission_inputs.items()
        }
        return CreditGroupCommissionRate(ma_to=ma_to, **values)

    @staticmethod
    def _parse_percent(text: str, field_name: str) -> float:
        cleaned = text.strip().replace(",", ".")
        if not cleaned:
            return 0.0
        try:
            return float(cleaned)
        except ValueError as exc:
            raise CreditGroupRepositoryError(
                f"Giá trị tỷ lệ không hợp lệ ở trường {field_name}."
            ) from exc

    def _update_rate_totals(self) -> None:
        if self._loading_rate:
            return
        try:
            rate = self._rate_from_inputs()
            no_secured_total = rate.total_no_secured()
            secured_total = rate.total_secured()
            errors = rate.validate()
        except CreditGroupRepositoryError as exc:
            self.no_secured_total_label.setText("--")
            self.secured_total_label.setText("--")
            self.rate_status.setText(str(exc))
            return
        self.no_secured_total_label.setText(self._format_percent_value(no_secured_total))
        self.secured_total_label.setText(self._format_percent_value(secured_total))
        self._style_total_label(self.no_secured_total_label, no_secured_total)
        self._style_total_label(self.secured_total_label, secured_total)
        self.rate_status.setText(
            "Tỷ lệ hợp lệ, có thể lưu."
            if not errors and self._selected_ma_to()
            else "\n".join(errors)
        )

    @staticmethod
    def _style_total_label(label: QLabel, total: float) -> None:
        if 99.99 <= total <= 100.01:
            label.setStyleSheet("color: #16794c; font-weight: 600;")
        else:
            label.setStyleSheet("color: #b42318; font-weight: 600;")

    def _save_selected_rate(self) -> None:
        if self.repository is None:
            self._set_repository_error()
            return
        try:
            rate = self._rate_from_inputs()
            self.repository.save_commission_rate(rate)
        except CreditGroupRepositoryError as exc:
            QMessageBox.warning(self, "Không thể lưu tỷ lệ hoa hồng", str(exc))
            return
        self.current_rate = self.repository.get_or_create_commission_rate(rate.ma_to)
        self._set_rate_inputs(self.current_rate)
        QMessageBox.information(
            self,
            "Tỷ lệ hoa hồng",
            f"Đã lưu tỷ lệ hoa hồng cho tổ {rate.ma_to}.",
        )

    def _reset_selected_rate(self) -> None:
        if self.repository is None:
            self._set_repository_error()
            return
        ma_to = self._selected_ma_to()
        if not ma_to:
            return
        try:
            self.current_rate = self.repository.reset_commission_rate_to_default(ma_to)
        except CreditGroupRepositoryError as exc:
            QMessageBox.warning(self, "Không thể khôi phục mặc định", str(exc))
            return
        self._set_rate_inputs(self.current_rate)

    def _load_rule_settings(self) -> None:
        if self.repository is None:
            if hasattr(self, "rule_status"):
                self.rule_status.setText(
                    self.repository_error or "Không thể mở dữ liệu tổ vay vốn."
                )
            return
        try:
            settings = self.repository.get_commission_rule_settings()
        except CreditGroupRepositoryError as exc:
            QMessageBox.warning(self, "Không thể tải cấu hình tỷ lệ", str(exc))
            return
        self._set_rule_inputs(settings)

    def _set_rule_inputs(self, settings: CreditCommissionRuleSettings) -> None:
        self._loading_rules = True
        try:
            for field_name, edit in self.rule_inputs.items():
                edit.setText(self._format_percent_value(getattr(settings, field_name)))
        finally:
            self._loading_rules = False
        self._update_rule_status()

    def _rule_settings_from_inputs(self) -> CreditCommissionRuleSettings:
        values = {
            field_name: self._parse_percent(edit.text(), field_name)
            for field_name, edit in self.rule_inputs.items()
        }
        return CreditCommissionRuleSettings(**values)

    def _update_rule_status(self) -> None:
        if self._loading_rules or not hasattr(self, "rule_status"):
            return
        try:
            settings = self._rule_settings_from_inputs()
        except CreditGroupRepositoryError as exc:
            self.rule_status.setText(str(exc))
            return
        errors = settings.validate()
        self.rule_status.setText(
            "Cấu hình điều kiện chi hợp lệ, có thể lưu."
            if not errors
            else "\n".join(errors)
        )

    def _save_rule_settings(self) -> None:
        if self.repository is None:
            self._set_repository_error()
            return
        try:
            settings = self._rule_settings_from_inputs()
            self.repository.save_commission_rule_settings(settings)
        except CreditGroupRepositoryError as exc:
            QMessageBox.warning(self, "Không thể lưu cấu hình tỷ lệ", str(exc))
            return
        self._load_rule_settings()
        QMessageBox.information(
            self,
            "Cài đặt tỷ lệ",
            "Đã lưu cấu hình điều kiện chi hoa hồng.",
        )

    def _reset_rule_settings(self) -> None:
        self._set_rule_inputs(CreditCommissionRuleSettings())

    def _import_data_tvv(self) -> None:
        if self.repository is None:
            self._set_repository_error()
            return
        source, _ = QFileDialog.getOpenFileName(
            self,
            "Chọn file Data_TVV",
            "",
            "Excel files (*.xlsx *.xlsm);;All files (*.*)",
        )
        if not source:
            return
        update_commission_rules = False
        try:
            has_rule_columns = self.repository.has_commission_rule_columns(Path(source))
        except OSError:
            has_rule_columns = False
        if has_rule_columns:
            answer = QMessageBox.question(
                self,
                "Import Data_TVV",
                "File có chứa cấu hình điều kiện chi hoa hồng. "
                "Bạn có muốn cập nhật cấu hình điều kiện chi hiện tại không?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No,
            )
            update_commission_rules = answer == QMessageBox.StandardButton.Yes
        try:
            imported_count = self.repository.import_data_tvv(
                Path(source),
                update_commission_rules=update_commission_rules,
            )
        except (CreditGroupRepositoryError, OSError) as exc:
            QMessageBox.warning(self, "Không thể import Data_TVV", str(exc))
            return
        self._load_groups()
        if update_commission_rules:
            self._load_rule_settings()
        QMessageBox.information(
            self,
            "Import Data_TVV",
            f"Đã import {imported_count} tổ vay vốn. "
            "Tỷ lệ hoa hồng mặc định đã được tạo cho các MaTo chưa có cấu hình riêng.",
        )

    def _create_data_tvv_template(self) -> None:
        output, _ = QFileDialog.getSaveFileName(
            self,
            "Tải file Excel mẫu Data_TVV",
            "Mau_Data_TVV.xlsx",
            "Excel files (*.xlsx);;All files (*.*)",
        )
        if not output:
            return
        output_path = Path(output)
        if output_path.suffix.lower() != ".xlsx":
            output_path = output_path.with_suffix(".xlsx")
        try:
            created_path = create_data_tvv_template(output_path)
        except OSError as exc:
            QMessageBox.warning(
                self,
                "Không thể tạo file Excel mẫu",
                f"Không tạo được file mẫu Data_TVV:\n{exc}",
            )
            return
        answer = QMessageBox.question(
            self,
            "Tải file Excel mẫu",
            "Đã tạo file Excel mẫu thành công.\nBạn có muốn mở thư mục chứa file không?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        if answer == QMessageBox.StandardButton.Yes:
            QDesktopServices.openUrl(QUrl.fromLocalFile(str(created_path.parent)))

    def _export_data_tvv(self) -> None:
        if self.repository is None:
            self._set_repository_error()
            return
        answer = QMessageBox.question(
            self,
            "Export Data_TVV",
            "Bạn có muốn xuất kèm tỷ lệ hoa hồng và điều kiện chi không?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        include_commission = answer == QMessageBox.StandardButton.Yes
        output, _ = QFileDialog.getSaveFileName(
            self,
            "Xuất Data_TVV",
            "Data_TVV_Kem_TyLe.xlsx" if include_commission else "Data_TVV.xlsx",
            "Excel files (*.xlsx);;All files (*.*)",
        )
        if not output:
            return
        output_path = Path(output)
        if output_path.suffix.lower() != ".xlsx":
            output_path = output_path.with_suffix(".xlsx")
        try:
            self.repository.export_data_tvv(
                output_path,
                include_commission=include_commission,
            )
        except (CreditGroupRepositoryError, OSError) as exc:
            QMessageBox.warning(self, "Không thể export Data_TVV", str(exc))
            return
        QMessageBox.information(
            self,
            "Export Data_TVV",
            (
                "Đã xuất file Data_TVV kèm tỷ lệ hoa hồng và điều kiện chi:\n"
                if include_commission
                else "Đã xuất file Data_TVV 22 cột gốc:\n"
            )
            + str(output_path),
        )


DEFAULT_CREDIT_GROUP_INFO_KEY = "credit_group_default_info"

DEFAULT_CREDIT_GROUP_INFO_FIELDS: tuple[str, ...] = (
    "xa",
    "dia_chi",
    "to_hoi",
    "tk_to_hoi_xa",
    "to_chuc",
    "ten_huyen",
    "tk_huyen",
    "ten_tinh",
    "tk_tinh",
    "ten_tw",
    "tk_tw",
    "uy_quyen",
    "ttln_tw",
    "ttln_tinh",
)

FIELD_PLACEHOLDERS: dict[str, str] = {
    "ma_to": "5491LLG202100003",
    "ten_to": "Tổ dịch vụ TD HND Đan Phượng 01",
    "ten_tvv_day_du": (
        "Tổ dịch vụ tín dụng Hội nông dân Đan Phượng 01 (Thôn Đoàn Kết)"
    ),
    "xa": "Đan Phượng",
    "dia_chi": "thôn Đoàn Kết, xã Tân Hà Lâm Hà, tỉnh Lâm Đồng",
    "ma_to_truong": "5491XXXXXXXXX",
    "ten_to_truong": "Nguyễn Văn A",
    "tk_to_truong": "5491205XXXXXX",
    "so_dien_thoai": "0912345678",
    "to_hoi": "Hội nông dân/Hội phụ nữ",
    "tk_to_hoi_xa": "5491201XXXXXX",
    "to_chuc": "Hội nông dân xã Đan Phượng",
    "ten_huyen": "Hội nông dân huyện Lâm Hà (có thể bỏ trống)",
    "tk_huyen": "5404201XXXXXX (có thể bỏ trống)",
    "ten_tinh": "Hội nông dân tỉnh Lâm Đồng",
    "tk_tinh": "5400XXXXXXXXX",
    "ten_tw": "Quỹ hỗ trợ nông dân TW/Hội liên hiệp phụ nữ Việt Nam",
    "tk_tw": "5491XXXXXXXXX",
    "ttln_tinh": "012026/TTLN-HNDVN-AGRIBANK ngày 27/01/2026",
    "ttln_tw": "012025/TTLN-HNDVN-AGRIBANK ngày 26/12/2025",
}


def normalize_uy_quyen(value: object) -> str:
    text = str(value or "").strip().casefold()
    if text in {"có", "co", "1", "true", "yes", "y"}:
        return "Có"
    return "Không"


def get_suggested_credit_group_default_info() -> dict[str, str]:
    return {
        "xa": "Đan Phượng",
        "dia_chi": "thôn Đoàn Kết, xã Tân Hà Lâm Hà, tỉnh Lâm Đồng",
        "to_hoi": "Hội nông dân/Hội phụ nữ",
        "tk_to_hoi_xa": "5491201XXXXXX",
        "to_chuc": "Hội nông dân xã Đan Phượng",
        "ten_huyen": "Hội nông dân huyện Lâm Hà",
        "tk_huyen": "5404201XXXXXX",
        "ten_tinh": "Hội nông dân tỉnh Lâm Đồng",
        "tk_tinh": "5400XXXXXXXXX",
        "ten_tw": "Quỹ hỗ trợ nông dân TW/Hội liên hiệp phụ nữ Việt Nam",
        "tk_tw": "5491XXXXXXXXX",
        "uy_quyen": "Không",
        "ttln_tw": "012025/TTLN-HNDVN-AGRIBANK ngày 26/12/2025",
        "ttln_tinh": "012026/TTLN-HNDVN-AGRIBANK ngày 27/01/2026",
    }


def _sanitize_default_credit_group_info(data: dict[str, object]) -> dict[str, str]:
    sanitized: dict[str, str] = {}
    for attr in DEFAULT_CREDIT_GROUP_INFO_FIELDS:
        value = data.get(attr, "")
        text = normalize_uy_quyen(value) if attr == "uy_quyen" else str(value or "").strip()
        if text:
            sanitized[attr] = text
    return sanitized


def get_default_credit_group_info() -> dict[str, str]:
    """Return safe default credit group values when branch settings exist later."""

    # Backward-compatible no-arg call for tests and old code paths.
    return {}


def load_default_credit_group_info(database_path: Path) -> dict[str, str]:
    try:
        raw_value = AppSettingsDatabase(database_path).load_preference(
            DEFAULT_CREDIT_GROUP_INFO_KEY,
            "",
        )
    except SettingsDatabaseError:
        return {}
    if not raw_value:
        return {}
    try:
        parsed = json.loads(raw_value)
    except json.JSONDecodeError:
        return {}
    if not isinstance(parsed, dict):
        return {}
    return _sanitize_default_credit_group_info(parsed)


def save_default_credit_group_info(
    data: dict[str, object],
    database_path: Path,
) -> dict[str, str]:
    sanitized = _sanitize_default_credit_group_info(data)
    payload = json.dumps(sanitized, ensure_ascii=False, sort_keys=True)
    try:
        AppSettingsDatabase(database_path).save_preference(
            DEFAULT_CREDIT_GROUP_INFO_KEY,
            payload,
        )
    except SettingsDatabaseError as exc:
        raise CreditGroupRepositoryError(
            f"Không thể lưu thông tin mặc định tổ vay vốn: {exc}"
        ) from exc
    return sanitized


class CreditGroupDefaultInfoDialog(QDialog):
    """Dialog for durable default values used by the Data_TVV add/edit form."""

    FIELD_GROUPS: tuple[tuple[str, tuple[str, ...]], ...] = (
        ("Địa bàn", ("xa", "dia_chi")),
        ("Tổ hội / tổ chức", ("to_hoi", "tk_to_hoi_xa", "to_chuc")),
        ("Cấp huyện", ("ten_huyen", "tk_huyen")),
        ("Cấp tỉnh", ("ten_tinh", "tk_tinh")),
        ("Cấp TW", ("ten_tw", "tk_tw")),
        ("Ủy quyền / thỏa thuận", ("uy_quyen", "ttln_tw", "ttln_tinh")),
    )

    def __init__(
        self,
        database_path: Path,
        parent: QWidget | None = None,
    ) -> None:
        super().__init__(parent)
        self.database_path = Path(database_path)
        self.inputs: dict[str, QLineEdit | QComboBox] = {}
        self.setWindowTitle("Cài đặt thông tin mặc định")
        self.setModal(True)
        self.setMinimumSize(780, 620)
        self.resize(820, 680)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(18, 16, 18, 14)
        layout.setSpacing(12)

        title = QLabel("Cài đặt thông tin mặc định")
        title.setObjectName("SectionTitle")
        layout.addWidget(title)

        description = QLabel(
            "Thiết lập các thông tin thường dùng để tự điền khi thêm mới tổ vay vốn."
        )
        description.setObjectName("MutedText")
        description.setWordWrap(True)
        layout.addWidget(description)

        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        body = QWidget()
        body_layout = QVBoxLayout(body)
        body_layout.setContentsMargins(0, 0, 0, 0)
        body_layout.setSpacing(10)
        for group_title, attrs in self.FIELD_GROUPS:
            box = QGroupBox(group_title)
            grid = QGridLayout(box)
            grid.setContentsMargins(14, 14, 14, 12)
            grid.setHorizontalSpacing(8)
            grid.setVerticalSpacing(8)
            grid.setColumnMinimumWidth(0, 200)
            grid.setColumnStretch(1, 1)
            for row_index, attr in enumerate(attrs):
                label = QLabel(CreditGroupEditDialog._display_label(attr, attr))
                label.setMinimumWidth(200)
                label.setWordWrap(True)
                label.setAlignment(
                    Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter
                )
                editor = self._create_editor(attr)
                self.inputs[attr] = editor
                grid.addWidget(label, row_index, 0)
                grid.addWidget(editor, row_index, 1)
            body_layout.addWidget(box)
        body_layout.addStretch()
        scroll_area.setWidget(body)
        layout.addWidget(scroll_area, stretch=1)

        button_layout = QHBoxLayout()
        restore_button = QPushButton("Khôi phục mặc định gợi ý")
        restore_button.setObjectName("SecondaryButton")
        restore_button.clicked.connect(self.restore_suggested_defaults)
        button_layout.addWidget(restore_button)
        button_layout.addStretch()

        close_button = QPushButton("Đóng")
        close_button.setObjectName("SecondaryButton")
        close_button.clicked.connect(self.reject)
        button_layout.addWidget(close_button)

        save_button = QPushButton("Lưu thông tin mặc định")
        save_button.setObjectName("PrimaryButton")
        save_button.clicked.connect(self.save_defaults)
        button_layout.addWidget(save_button)
        layout.addLayout(button_layout)

        self.load_saved_defaults()

    def _create_editor(self, attr: str) -> QLineEdit | QComboBox:
        if attr == "uy_quyen":
            combo = QComboBox()
            combo.addItems(("Có", "Không"))
            combo.setCurrentText("Không")
            combo.setMinimumWidth(500)
            combo.setToolTip("Chọn Có hoặc Không")
            combo.setProperty("user_changed", False)
            combo.currentIndexChanged.connect(
                lambda _index, field=combo: field.setProperty("user_changed", True)
            )
            return combo
        edit = QLineEdit()
        edit.setObjectName("TextInput")
        edit.setMinimumWidth(500)
        placeholder = FIELD_PLACEHOLDERS.get(attr)
        if placeholder:
            edit.setPlaceholderText(placeholder)
            edit.setToolTip(placeholder)
        return edit

    def load_saved_defaults(self) -> None:
        self._set_values(load_default_credit_group_info(self.database_path))

    def restore_suggested_defaults(self) -> None:
        self._set_values(get_suggested_credit_group_default_info())

    def save_defaults(self) -> None:
        values = {attr: self._field_text(editor) for attr, editor in self.inputs.items()}
        try:
            save_default_credit_group_info(values, self.database_path)
        except CreditGroupRepositoryError as exc:
            QMessageBox.warning(self, self.windowTitle(), str(exc))
            return
        QMessageBox.information(
            self,
            self.windowTitle(),
            "Đã lưu thông tin mặc định cho tổ vay vốn.",
        )
        self.accept()

    def _set_values(self, values: dict[str, str]) -> None:
        for attr, value in values.items():
            editor = self.inputs.get(attr)
            if editor is not None:
                self._set_field_text(editor, value)

    @staticmethod
    def _field_text(editor: QLineEdit | QComboBox) -> str:
        if isinstance(editor, QComboBox):
            return normalize_uy_quyen(editor.currentText())
        return editor.text().strip()

    @staticmethod
    def _set_field_text(editor: QLineEdit | QComboBox, value: object) -> None:
        if isinstance(editor, QComboBox):
            editor.setCurrentText(normalize_uy_quyen(value))
            return
        editor.setText(str(value or "").strip())


class CreditGroupEditDialog(QDialog):
    """Shared add/edit dialog for Data_TVV credit group records."""

    FIELD_GROUPS: tuple[tuple[str, tuple[tuple[str, str], ...]], ...] = (
        (
            "Thông tin tổ",
            (
                ("STT", "stt"),
                ("MaTo", "ma_to"),
                ("TenTo", "ten_to"),
                ("TenTVV_DayDu", "ten_tvv_day_du"),
                ("Xa", "xa"),
                ("DiaChi", "dia_chi"),
            ),
        ),
        (
            "Tổ trưởng",
            (
                ("MaToTruong", "ma_to_truong"),
                ("Ten_ToTruong", "ten_to_truong"),
                ("TK_ToTruong", "tk_to_truong"),
                ("SoDienThoai", "so_dien_thoai"),
            ),
        ),
        (
            "Tổ hội / tổ chức",
            (
                ("ToHoi", "to_hoi"),
                ("TK_ToHoiXa", "tk_to_hoi_xa"),
                ("ToChuc", "to_chuc"),
            ),
        ),
        (
            "Cấp huyện / tỉnh / trung ương",
            (
                ("Ten_Huyen", "ten_huyen"),
                ("TK_HUYEN", "tk_huyen"),
                ("Ten_Tinh", "ten_tinh"),
                ("TK_TINH", "tk_tinh"),
                ("Ten_TW", "ten_tw"),
                ("TK_TW", "tk_tw"),
            ),
        ),
        (
            "Ủy quyền / thỏa thuận",
            (
                ("uyquyen", "uy_quyen"),
                ("TTLN_TW", "ttln_tw"),
                ("TTLN_Tinh", "ttln_tinh"),
            ),
        ),
    )
    FIELD_PLACEHOLDERS = FIELD_PLACEHOLDERS

    def __init__(
        self,
        repository: CreditGroupRepository,
        mode: str,
        group: CreditGroup | None = None,
        parent: QWidget | None = None,
    ) -> None:
        super().__init__(parent)
        self.repository = repository
        self.mode = mode
        self.group = group
        self.inputs: dict[str, QLineEdit | QComboBox] = {}
        self.saved_ma_to = ""
        title_text = (
            "Thêm mới tổ vay vốn" if mode == "add" else "Sửa thông tin tổ vay vốn"
        )
        action_text = "Thêm mới" if mode == "add" else "Sửa"
        self.setWindowTitle(title_text)
        self.setModal(True)
        self.setMinimumSize(820, 700)
        self.resize(860, 740)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(18, 16, 18, 14)
        layout.setSpacing(12)

        title = QLabel(title_text)
        title.setObjectName("SectionTitle")
        layout.addWidget(title)

        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        body = QWidget()
        body_layout = QVBoxLayout(body)
        body_layout.setContentsMargins(0, 0, 0, 0)
        body_layout.setSpacing(10)
        for group_title, fields in self.FIELD_GROUPS:
            box = QGroupBox(group_title)
            grid = QGridLayout(box)
            grid.setContentsMargins(14, 16, 14, 14)
            grid.setHorizontalSpacing(8)
            grid.setVerticalSpacing(8)
            grid.setColumnMinimumWidth(0, 200)
            grid.setColumnStretch(1, 1)
            for row_index, (technical_label, attr) in enumerate(fields):
                label = QLabel(self._display_label(attr, technical_label))
                label.setMinimumWidth(200)
                label.setWordWrap(True)
                label.setAlignment(
                    Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter
                )
                editor = self._create_editor(attr)
                if attr == "stt" and isinstance(editor, QLineEdit):
                    editor.setToolTip(
                        "Nếu STT bị trùng, hệ thống sẽ tự sắp xếp lại thứ tự."
                    )
                self.inputs[attr] = editor
                grid.addWidget(label, row_index, 0)
                grid.addWidget(editor, row_index, 1)
            body_layout.addWidget(box)
        body_layout.addStretch()
        scroll_area.setWidget(body)
        layout.addWidget(scroll_area, stretch=1)

        button_layout = QHBoxLayout()
        default_button = QPushButton("Lấy thông tin mặc định")
        default_button.setObjectName("SecondaryButton")
        default_button.clicked.connect(self.apply_default_info)
        button_layout.addWidget(default_button)

        settings_button = QPushButton("Cài đặt thông tin mặc định")
        settings_button.setObjectName("SecondaryButton")
        settings_button.clicked.connect(self.open_default_info_settings)
        button_layout.addWidget(settings_button)
        button_layout.addStretch()

        cancel_button = QPushButton("Hủy")
        cancel_button.setObjectName("SecondaryButton")
        cancel_button.clicked.connect(self.reject)
        button_layout.addWidget(cancel_button)

        save_button = QPushButton(action_text)
        save_button.setObjectName("PrimaryButton")
        save_button.clicked.connect(self.save_group)
        button_layout.addWidget(save_button)
        layout.addLayout(button_layout)

        if group is not None:
            self.load_group_to_form(group)
        if mode == "edit":
            self.inputs["ma_to"].setReadOnly(True)

    def _create_editor(self, attr: str) -> QLineEdit | QComboBox:
        if attr == "uy_quyen":
            combo = QComboBox()
            combo.addItems(("Có", "Không"))
            combo.setCurrentText("Không")
            combo.setMinimumWidth(500)
            combo.setToolTip("Chọn Có hoặc Không")
            return combo
        edit = QLineEdit()
        edit.setObjectName("TextInput")
        edit.setMinimumWidth(500)
        placeholder = self.FIELD_PLACEHOLDERS.get(attr)
        if placeholder:
            edit.setPlaceholderText(placeholder)
            edit.setToolTip(placeholder)
        return edit

    def load_group_to_form(self, group: CreditGroup) -> None:
        for attr, editor in self.inputs.items():
            self._set_field_text(editor, getattr(group, attr))

    def collect_form_data(self) -> CreditGroup:
        values = {
            attr: self._field_text(editor)
            for attr, editor in self.inputs.items()
        }
        stt_text = values.pop("stt", "")
        stt = int(stt_text) if stt_text else 0
        return CreditGroup(stt=stt, **values)

    def validate_form_data(self, group: CreditGroup) -> list[str]:
        errors: list[str] = []
        if not group.ma_to:
            errors.append("Mã tổ không được để trống.")
        if not group.ten_to:
            errors.append("Tên tổ không được để trống.")
        stt_text = self._field_text(self.inputs["stt"])
        if stt_text:
            try:
                stt_value = int(stt_text)
            except ValueError:
                errors.append("STT phải là số hoặc để trống.")
            else:
                if stt_value <= 0:
                    errors.append("STT phải lớn hơn 0.")
        if self.mode == "add" and group.ma_to and self.repository.get_group(group.ma_to):
            errors.append("Mã tổ đã tồn tại.")
        return errors

    def apply_default_info(self) -> None:
        defaults = load_default_credit_group_info(self.repository.database_path)
        if not defaults:
            QMessageBox.information(
                self,
                "Lấy thông tin mặc định",
                "Chưa có thông tin mặc định. "
                "Vui lòng bấm ‘Cài đặt thông tin mặc định’ để thiết lập.",
            )
            return
        updated_count = self._apply_default_values(defaults)
        message = (
            "Đã điền thông tin mặc định vào các trường còn trống."
            if updated_count
            else "Không có trường trống phù hợp để điền thông tin mặc định."
        )
        QMessageBox.information(self, "Lấy thông tin mặc định", message)

    def _apply_default_values(self, defaults: dict[str, str]) -> int:
        updated_count = 0
        for attr, value in defaults.items():
            editor = self.inputs.get(attr)
            if editor is None:
                continue
            current_value = self._field_text(editor)
            if attr == "uy_quyen":
                user_changed = bool(editor.property("user_changed"))
                if self.mode == "add" and current_value == "Không" and not user_changed:
                    normalized = normalize_uy_quyen(value)
                    if normalized != current_value:
                        self._set_field_text(editor, normalized)
                        updated_count += 1
                continue
            if not current_value:
                self._set_field_text(editor, value)
                updated_count += 1
        return updated_count

    def open_default_info_settings(self) -> None:
        dialog = CreditGroupDefaultInfoDialog(
            self.repository.database_path,
            self,
        )
        dialog.exec()

    def save_group(self) -> None:
        try:
            group = self.collect_form_data()
        except ValueError:
            QMessageBox.warning(
                self,
                self.windowTitle(),
                "STT phải là số hoặc để trống.",
            )
            return
        errors = self.validate_form_data(group)
        if errors:
            QMessageBox.warning(
                self,
                self.windowTitle(),
                "Vui lòng kiểm tra lại thông tin nhập:\n" + "\n".join(errors),
            )
            return
        try:
            self.repository.save_group(group)
        except CreditGroupRepositoryError as exc:
            QMessageBox.warning(self, self.windowTitle(), str(exc))
            return
        self.saved_ma_to = group.ma_to
        self.accept()

    @staticmethod
    def _display_label(attr: str, fallback: str) -> str:
        header = DATA_TVV_ATTR_TO_HEADER.get(attr, fallback)
        return DATA_TVV_FIELD_LABELS.get(header, fallback)

    @staticmethod
    def _field_text(editor: QLineEdit | QComboBox) -> str:
        if isinstance(editor, QComboBox):
            return normalize_uy_quyen(editor.currentText())
        return editor.text().strip()

    @staticmethod
    def _set_field_text(editor: QLineEdit | QComboBox, value: object) -> None:
        if isinstance(editor, QComboBox):
            was_blocked = editor.blockSignals(True)
            editor.setCurrentText(normalize_uy_quyen(value))
            editor.blockSignals(was_blocked)
            return
        editor.setText(str(value or "").strip())


class CreditMigrationPlaceholderDialog(QDialog):
    """Temporary window for credit features not migrated yet."""

    def __init__(self, title: str, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setWindowTitle(title)
        self.setModal(True)
        self.setMinimumWidth(560)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(18, 16, 18, 14)
        layout.setSpacing(12)

        title_label = QLabel(title)
        title_label.setObjectName("SectionTitle")
        layout.addWidget(title_label)

        message = QLabel(
            "Chức năng đang được chuyển đổi từ agribank-tool.xlam. "
            "Vui lòng xem docs/migration_5491_tovayvon.md để biết trạng thái."
        )
        message.setObjectName("MutedText")
        message.setWordWrap(True)
        message.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignTop)
        layout.addWidget(message)

        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Close)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)
