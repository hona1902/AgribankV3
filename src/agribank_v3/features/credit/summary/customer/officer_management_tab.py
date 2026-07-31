from __future__ import annotations

from pathlib import Path

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QCheckBox,
    QDialog,
    QFileDialog,
    QFormLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QSizePolicy,
    QVBoxLayout,
    QWidget,
)

from agribank_v3.features.credit.summary.customer.export_service import (
    OFFICER_DIRECTORY_COLUMNS,
    export_officer_directory,
    suggested_customer_export_name,
)
from agribank_v3.features.credit.summary.customer.async_query import AsyncQueryController
from agribank_v3.features.credit.summary.customer.formatters import normalize_officer_code, normalize_officer_name
from agribank_v3.features.credit.summary.customer.repository import CustomerRepository
from agribank_v3.features.credit.summary.customer.table_models import CustomerTableModel
from agribank_v3.features.credit.summary.customer.widgets import (
    CustomerTableView,
    Pager,
    QueryStateBanner,
    SearchBox,
    combo_box,
    current_data,
    make_compact_control,
    populate_combo,
    fit_window_to_screen,
    primary_button,
    secondary_button,
)


class OfficerManagementTab(QWidget):
    def __init__(self, repository: CustomerRepository, parent=None) -> None:
        super().__init__(parent)
        self.repository = repository
        self.page = 1
        self.page_size = 100
        self.query_controller = AsyncQueryController(self, max_cache_entries=16)
        layout = QVBoxLayout(self)
        filters = QHBoxLayout()
        self.search_box = SearchBox("Tìm mã hoặc tên cán bộ")
        self.branch_combo = combo_box("Tất cả chi nhánh", minimum_width=180, maximum_width=260)
        self.status_combo = combo_box("Tất cả trạng thái", minimum_width=160, maximum_width=220)
        populate_combo(self.status_combo, [("Đang sử dụng", "active"), ("Ngừng sử dụng", "inactive")])
        self.search_box.debouncedTextChanged.connect(self._filter_changed)
        self.branch_combo.currentIndexChanged.connect(self._filter_changed)
        self.status_combo.currentIndexChanged.connect(self._filter_changed)
        filters.addWidget(self.search_box, stretch=1)
        filters.addWidget(self.branch_combo)
        filters.addWidget(self.status_combo)
        layout.addLayout(filters)
        actions = QHBoxLayout()
        add_button = primary_button("Thêm cán bộ")
        edit_button = secondary_button("Sửa thông tin")
        disable_button = secondary_button("Ngừng sử dụng")
        export_button = secondary_button("Xuất Excel")
        add_button.clicked.connect(self.add_officer)
        edit_button.clicked.connect(self.edit_officer)
        disable_button.clicked.connect(self.disable_officer)
        export_button.clicked.connect(self.export_excel)
        for button in (add_button, edit_button, disable_button, export_button):
            actions.addWidget(button)
        actions.addStretch()
        layout.addLayout(actions)
        self.state_banner = QueryStateBanner()
        self.state_banner.retryRequested.connect(lambda: self.refresh(use_cache=False))
        layout.addWidget(self.state_banner)
        self.model = CustomerTableModel(OFFICER_DIRECTORY_COLUMNS, self)
        self.table = CustomerTableView()
        self.table.setModel(self.model)
        layout.addWidget(self.table, stretch=1)
        self.pager = Pager()
        self.pager.pageChanged.connect(self._page_changed)
        self.pager.pageSizeChanged.connect(self._page_size_changed)
        layout.addWidget(self.pager)
        self.table.apply_default_widths((120, 220, 100, 120, 110, 160))

    def refresh_filters(self) -> None:
        populate_combo(
            self.branch_combo,
            [
                (self.repository.unit_directory.get_branch_display_name(code), code)
                for code in self.repository.distinct_branch_codes()
            ],
        )

    def refresh(self, *args, use_cache: bool = True) -> None:
        search_text = self.search_box.text()
        branch_code = current_data(self.branch_combo)
        status = current_data(self.status_combo)
        page = self.page
        page_size = self.page_size
        cache_key = ("officer_directory", search_text, branch_code, status, page, page_size)
        self.query_controller.run(
            "customer_officer_directory",
            lambda: self.repository.officer_directory(
                search_text=search_text,
                branch_code=branch_code,
                status=status,
                page=page,
                page_size=page_size,
            ),
            self._apply_result,
            self._query_failed,
            cache_key=cache_key,
            use_cache=use_cache,
            state_callback=self._state_changed,
        )

    def add_officer(self) -> None:
        dialog = OfficerDirectoryDialog(self.repository, parent=self)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            self.invalidate_cache()
            self._notify_parent_cache_invalidated()
            self.refresh_filters()
            self.refresh(use_cache=False)

    def edit_officer(self) -> None:
        row = self._selected_row()
        if not row:
            return
        dialog = OfficerDirectoryDialog(self.repository, row=row, parent=self)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            self.invalidate_cache()
            self._notify_parent_cache_invalidated()
            self.refresh_filters()
            self.refresh(use_cache=False)

    def disable_officer(self) -> None:
        row = self._selected_row()
        if not row:
            return
        self.repository.disable_officer(str(row.get("officer_code") or ""))
        self.invalidate_cache()
        self._notify_parent_cache_invalidated()
        self.refresh(use_cache=False)

    def export_excel(self) -> None:
        path, _ = QFileDialog.getSaveFileName(
            self,
            "Xuất danh mục cán bộ",
            suggested_customer_export_name("DanhMucCanBo"),
            "Excel (*.xlsx)",
        )
        if not path:
            return
        try:
            output = export_officer_directory(self.repository, Path(path))
        except Exception as exc:
            QMessageBox.warning(self, "Xuất danh mục cán bộ", str(exc))
            return
        QMessageBox.information(self, "Xuất danh mục cán bộ", f"Đã xuất: {output}")

    def _selected_row(self) -> dict[str, object]:
        indexes = self.table.selectionModel().selectedRows() if self.table.selectionModel() else []
        if not indexes:
            return {}
        return self.model.raw_row(indexes[0].row())

    def _filter_changed(self) -> None:
        self.page = 1
        self.refresh()

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

    def _apply_result(self, result) -> None:
        self.model.set_rows(result.rows)
        self.pager.set_state(page=result.page, page_size=result.page_size, total_rows=result.total_rows)
        if result.total_rows:
            self.state_banner.clear()
        else:
            self.state_banner.set_empty("Không có cán bộ phù hợp với bộ lọc.")

    def _query_failed(self, exc: Exception) -> None:
        self.model.set_rows([])
        self.pager.set_state(page=1, page_size=self.page_size, total_rows=0)

    def _state_changed(self, state: str, message: str) -> None:
        if state == "loading":
            self.state_banner.set_loading()
        elif state == "error":
            self.state_banner.set_error(message or "Không tải được danh mục cán bộ.")

    def _notify_parent_cache_invalidated(self) -> None:
        parent = self.parent()
        while parent is not None:
            if hasattr(parent, "invalidate_customer_caches"):
                parent.invalidate_customer_caches()
                return
            parent = parent.parent()


class OfficerDirectoryDialog(QDialog):
    def __init__(self, repository: CustomerRepository, row: dict[str, object] | None = None, parent=None) -> None:
        super().__init__(parent)
        self.repository = repository
        self.original_officer_code = normalize_officer_code((row or {}).get("officer_code"))
        self.setWindowTitle("Cập nhật danh mục cán bộ")
        fit_window_to_screen(
            self,
            width_ratio=0.42,
            height_ratio=0.34,
            max_width=650,
            max_height=350,
            min_width=520,
            min_height=260,
        )
        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 14, 16, 14)
        layout.setSpacing(8)
        layout.setSizeConstraint(QVBoxLayout.SizeConstraint.SetMinAndMaxSize)
        form = QFormLayout()
        self.form_layout = form
        form.setLabelAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        form.setFormAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignTop)
        form.setHorizontalSpacing(12)
        form.setVerticalSpacing(8)
        self.code_input = QLineEdit(str((row or {}).get("officer_code") or ""))
        self.name_input = QLineEdit(str((row or {}).get("officer_name") or ""))
        self.branch_input = QLineEdit(str((row or {}).get("branch_code") or ""))
        self.office_input = QLineEdit(str((row or {}).get("transaction_office") or ""))
        for field in (self.code_input, self.name_input, self.branch_input, self.office_input):
            field.setMinimumWidth(340)
            field.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
            make_compact_control(field)
            field.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
            field.textEdited.connect(lambda _text: self._set_error(""))
        self.active_check = QCheckBox("Đang sử dụng")
        self.active_check.setChecked(int((row or {}).get("is_active", 1) or 0) == 1)
        for label, widget in (
            ("Mã cán bộ", self.code_input),
            ("Tên cán bộ", self.name_input),
            ("Chi nhánh", self.branch_input),
            ("Phòng GD", self.office_input),
            ("Trạng thái", self.active_check),
        ):
            form.addRow(label, widget)
        layout.addLayout(form)
        self.error_label = QLabel("")
        self.error_label.setObjectName("ValidationErrorLabel")
        self.error_label.setWordWrap(True)
        self.error_label.setStyleSheet("color: #b91c1c;")
        layout.addWidget(self.error_label)
        actions = QHBoxLayout()
        self.actions_layout = actions
        actions.setSpacing(7)
        self.save_button = primary_button("Lưu")
        self.cancel_button = secondary_button("Hủy")
        self.save_button.setDefault(True)
        self.save_button.clicked.connect(self._save)
        self.cancel_button.clicked.connect(self.reject)
        actions.addStretch()
        actions.addWidget(self.save_button)
        actions.addWidget(self.cancel_button)
        layout.addLayout(actions)
        self.adjustSize()
        self.setMinimumWidth(520)
        compact_width = min(650, max(560, self.sizeHint().width()))
        compact_height = min(360, max(300, self.sizeHint().height()))
        self.resize(compact_width, compact_height)
        if self.original_officer_code:
            self.name_input.setFocus()
        else:
            self.code_input.setFocus()

    def _save(self) -> None:
        code = normalize_officer_code(self.code_input.text())
        name = normalize_officer_name(self.name_input.text())
        self.code_input.setText(code)
        self.name_input.setText(name)
        self.branch_input.setText(self.branch_input.text().strip())
        self.office_input.setText(self.office_input.text().strip())
        if not code:
            self._set_error("Mã cán bộ không được để trống.")
            self.code_input.setFocus()
            return
        if not name:
            self._set_error("Tên cán bộ không được để trống.")
            self.name_input.setFocus()
            return
        if self._is_duplicate_code(code):
            self._set_error("Mã cán bộ đã tồn tại trong danh mục.")
            self.code_input.setFocus()
            return
        try:
            self.repository.upsert_officer_directory(
                officer_code=code,
                officer_name=name,
                branch_code=self.branch_input.text(),
                transaction_office=self.office_input.text(),
                is_active=self.active_check.isChecked(),
            )
        except Exception as exc:
            self._set_error(str(exc))
            return
        self.accept()

    def _is_duplicate_code(self, code: str) -> bool:
        current = normalize_officer_code(code)
        if not current:
            return False
        original = self.original_officer_code.casefold()
        rows = self.repository.officer_directory(search_text=current, page=1, page_size=50).rows
        for row in rows:
            existing = normalize_officer_code(row.get("officer_code"))
            if existing.casefold() == current.casefold() and existing.casefold() != original:
                return True
        return False

    def _set_error(self, message: str) -> None:
        self.error_label.setText(message)
