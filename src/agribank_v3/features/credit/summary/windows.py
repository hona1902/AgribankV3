from __future__ import annotations

from contextlib import closing
from datetime import date
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
import getpass
from pathlib import Path
from typing import Callable

from PySide6.QtCore import QDate, QEvent, QPoint, QRect, QSettings, Qt, QTimer, QUrl
from PySide6.QtGui import QColor, QDesktopServices, QPainter, QPen
from PySide6.QtWidgets import (
    QAbstractItemView,
    QComboBox,
    QDateEdit,
    QDialog,
    QDoubleSpinBox,
    QFileDialog,
    QFormLayout,
    QFrame,
    QGroupBox,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QMessageBox,
    QProgressBar,
    QPushButton,
    QSizePolicy,
    QSpinBox,
    QTabWidget,
    QTableWidget,
    QTableWidgetItem,
    QToolTip,
    QVBoxLayout,
    QWidget,
)

from agribank_v3.features.credit.summary.models import (
    CREDIT_LIMIT_TITLE,
    LOAN_COMPARE_TITLE,
    NIM_DN_TITLE,
    NIM_NV_TITLE,
    NIM_TITLE,
    DashboardData,
    DashboardMetric,
    NIM_DN_CONFIG,
    NIM_NV_CONFIG,
    PageResult,
    SummaryDataType,
    SummaryError,
)
from agribank_v3.features.credit.summary.history_dialog import OfficerHistoryDialog
from agribank_v3.features.credit.summary.nim_ui_config import NimUiConfig, get_nim_ui_config
from agribank_v3.features.credit.summary.officer_history.widgets import MultiLineHeaderView, NumericTableWidgetItem
from agribank_v3.features.credit.summary.reports import export_rows
from agribank_v3.features.credit.summary.repository import NIM_OFFICER_DISPLAY_SQL, SummaryRepository
from agribank_v3.features.credit.summary.services import (
    compare_loan_balances,
    import_credit_limit_file,
    import_nim_dn,
    import_nim_nv,
    parse_period_from_filename,
)
from agribank_v3.features.settings.unit_directory.service import UnitDirectoryService, get_unit_directory_service
from agribank_v3.runtime_paths import application_root
from agribank_v3.settings import AppSettingsDatabase
from agribank_v3.ui.components.controls import (
    combo_box as shared_combo_box,
    current_data as shared_current_data,
    danger_button as shared_danger_button,
    populate_combo as shared_populate_combo,
    primary_button as shared_primary_button,
    secondary_button as shared_secondary_button,
)
from agribank_v3.ui.components.kpi import KpiMetric, MetricGrid
from agribank_v3.ui.components.flow_layout import FlowLayout
from agribank_v3.ui.workers import run_in_thread


LOAN_COMPARE_FIELD_ORDER = (
    "customer_code",
    "customer_name",
    "previous_balance",
    "current_balance",
    "difference",
    "category",
    "officer",
    "address",
)
LOAN_COMPARE_HEADERS = {
    "customer_code": "Mã KH",
    "customer_name": "Tên khách hàng",
    "previous_balance": "Dư nợ kỳ trước",
    "previous_blance": "Dư nợ kỳ trước",
    "current_balance": "Dư nợ kỳ này",
    "current_blance": "Dư nợ kỳ này",
    "difference": "Tăng/giảm",
    "category": "Loại KH",
    "officer": "Cán bộ QL",
    "address": "Địa chỉ",
}
LOAN_COMPARE_MONEY_FIELDS = {"previous_balance", "previous_blance", "current_balance", "current_blance", "difference"}
LOAN_COMPARE_FIELD_ALIASES = {
    "previous_balance": "previous_blance",
    "current_balance": "current_blance",
}
LOAN_COMPARE_CATEGORY_LABELS = {
    "Khach hang tat toan": "Khách hàng tất toán",
    "Khach hang vay giam": "Khách hàng vay giảm",
    "Khach hang vay moi": "Khách hàng vay mới",
    "Khach hang vay tang": "Khách hàng vay tăng",
    "Khong thay doi": "Không thay đổi",
}
LOAN_COMPARE_COLUMN_WIDTHS = (86, 188, 132, 132, 118, 140, 150, 210)


class MiniChart(QWidget):
    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.values: tuple[tuple[str, float], ...] = ()
        self.setMinimumHeight(150)

    def set_values(self, values: tuple[tuple[str, float], ...]) -> None:
        self.values = values[:10]
        self.update()

    def paintEvent(self, event) -> None:
        super().paintEvent(event)
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        rect = self.rect().adjusted(8, 8, -8, -20)
        painter.fillRect(self.rect(), QColor("#ffffff"))
        painter.setPen(QColor("#d8dee8"))
        painter.drawRect(rect)
        if not self.values:
            painter.setPen(QColor("#6b7280"))
            painter.drawText(self.rect(), Qt.AlignmentFlag.AlignCenter, "Chưa có dữ liệu biểu đồ")
            return
        max_value = max(abs(value) for _, value in self.values) or 1
        bar_width = max(12, rect.width() // max(1, len(self.values) * 2))
        gap = max(8, (rect.width() - bar_width * len(self.values)) // max(1, len(self.values) + 1))
        base_y = rect.bottom() if all(value >= 0 for _, value in self.values) else rect.center().y()
        painter.setPen(QColor("#2563eb"))
        painter.setBrush(QColor("#2f80ed"))
        for index, (label, value) in enumerate(self.values):
            height = int((abs(value) / max_value) * (rect.height() - 28))
            x = rect.left() + gap + index * (bar_width + gap)
            y = base_y - height if value >= 0 else base_y
            painter.drawRoundedRect(QRect(x, y, bar_width, max(2, height)), 3, 3)
            painter.setPen(QColor("#374151"))
            painter.drawText(QRect(x - gap // 2, rect.bottom() + 2, bar_width + gap, 16), Qt.AlignmentFlag.AlignCenter, _short(label))
            painter.setPen(QColor("#2563eb"))


class NimTrendChart(QWidget):
    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.points_data: tuple[tuple[str, float, float], ...] = ()
        self.point_rects: list[tuple[QRect, tuple[str, float, float]]] = []
        self.setMinimumHeight(190)
        self.setMouseTracking(True)

    def set_values(self, values: tuple[tuple[str, float], ...]) -> None:
        _ = values

    def set_trend(self, values: tuple[tuple[str, float, float], ...]) -> None:
        self.points_data = values
        self.update()

    def paintEvent(self, event) -> None:
        super().paintEvent(event)
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        painter.fillRect(self.rect(), QColor("#ffffff"))
        rect = self.rect().adjusted(46, 16, -18, -38)
        painter.setPen(QColor("#d8dee8"))
        painter.drawRect(rect)
        self.point_rects = []
        if not self.points_data:
            painter.setPen(QColor("#6b7280"))
            painter.drawText(self.rect(), Qt.AlignmentFlag.AlignCenter, "Chưa có dữ liệu xu hướng NIM")
            return

        values = [before for _, before, _ in self.points_data] + [after for _, _, after in self.points_data]
        min_value = min(values)
        max_value = max(values)
        if min_value == max_value:
            min_value -= 1
            max_value += 1
        padding = (max_value - min_value) * 0.08
        min_value -= padding
        max_value += padding

        painter.setPen(QColor("#6b7280"))
        for index in range(5):
            ratio = index / 4
            y = rect.bottom() - int(rect.height() * ratio)
            value = min_value + (max_value - min_value) * ratio
            painter.drawLine(rect.left(), y, rect.right(), y)
            painter.drawText(QRect(0, y - 9, 42, 18), Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter, _format_percent_axis(value))

        before_points = self._line_points(rect, min_value, max_value, series_index=1)
        after_points = self._line_points(rect, min_value, max_value, series_index=2)
        self._draw_series(painter, before_points, QColor("#1f6feb"), "NIM trước ĐC")
        self._draw_series(painter, after_points, QColor("#d97706"), "NIM sau ĐC")

        painter.setPen(QColor("#374151"))
        step = max(1, len(self.points_data) // 8)
        for index, (period, _, _) in enumerate(self.points_data):
            if index % step == 0 or index == len(self.points_data) - 1:
                x = before_points[index].x()
                painter.drawText(QRect(x - 36, rect.bottom() + 6, 72, 18), Qt.AlignmentFlag.AlignCenter, period)

        legend_y = self.height() - 20
        self._draw_legend(painter, 54, legend_y, QColor("#1f6feb"), "NIM trước ĐC")
        self._draw_legend(painter, 170, legend_y, QColor("#d97706"), "NIM sau ĐC")

    def mouseMoveEvent(self, event) -> None:
        position = event.position().toPoint()
        for rect, payload in self.point_rects:
            if rect.contains(position):
                period, before, after = payload
                QToolTip.showText(
                    event.globalPosition().toPoint(),
                    f"{period}\nNIM trước ĐC: {_format_percent_vn(before)}\nNIM sau ĐC: {_format_percent_vn(after)}",
                    self,
                )
                return
        QToolTip.hideText()
        super().mouseMoveEvent(event)

    def _line_points(self, rect: QRect, min_value: float, max_value: float, *, series_index: int) -> list[QPoint]:
        points: list[QPoint] = []
        count = len(self.points_data)
        for index, payload in enumerate(self.points_data):
            value = payload[series_index]
            x = rect.left() + int(rect.width() * index / max(1, count - 1))
            y = rect.bottom() - int(((value - min_value) / (max_value - min_value)) * rect.height())
            points.append(QPoint(x, y))
        return points

    def _draw_series(self, painter: QPainter, points: list[QPoint], color: QColor, label: str) -> None:
        _ = label
        painter.setPen(QPen(color, 2))
        for index in range(1, len(points)):
            painter.drawLine(points[index - 1], points[index])
        painter.setBrush(color)
        painter.setPen(QPen(color, 1))
        for index, point in enumerate(points):
            painter.drawEllipse(point, 4, 4)
            self.point_rects.append((QRect(point.x() - 8, point.y() - 8, 16, 16), self.points_data[index]))

    @staticmethod
    def _draw_legend(painter: QPainter, x: int, y: int, color: QColor, label: str) -> None:
        painter.setPen(QPen(color, 2))
        painter.drawLine(x, y, x + 22, y)
        painter.setBrush(color)
        painter.drawEllipse(QPoint(x + 11, y), 4, 4)
        painter.setPen(QColor("#374151"))
        painter.drawText(QRect(x + 28, y - 9, 120, 18), Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter, label)


class MetricStrip(MetricGrid):
    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)

    def set_data(self, dashboard: DashboardData) -> None:
        self.set_metrics([_dashboard_metric_to_kpi(metric) for metric in dashboard.metrics])


def _dashboard_metric_to_kpi(metric: DashboardMetric) -> KpiMetric:
    value = str(metric.value or "").strip()
    label_key = metric.label.casefold()
    if "dư nợ" in label_key or "nguồn vốn" in label_key or "số dư" in label_key:
        number = _parse_vn_number(value)
        return KpiMetric(metric.label, number, "money", full_value=number, tooltip=metric.detail)
    if "nim" in label_key or "lãi suất" in label_key:
        number = _parse_vn_percent(value)
        return KpiMetric(metric.label, number, "percent", full_value=number, tooltip=metric.detail)
    return KpiMetric(metric.label, value, "text", tooltip=metric.detail)


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


def _parse_vn_percent(value: str) -> float | None:
    return _parse_vn_number(value)


class SummaryDataTab(QWidget):
    title = "Báo cáo"

    def __init__(self, repository: SummaryRepository, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.repository = repository
        self.page = 1
        self.page_size = 200
        self.current_rows: list[dict[str, object]] = []
        self.settings = QSettings("AgribankV3", "AgribankV3")
        self._restoring_columns = False
        self.search_timer = QTimer(self)
        self.search_timer.setInterval(250)
        self.search_timer.setSingleShot(True)
        self.search_timer.timeout.connect(self.reload)
        self._build_ui()

    def _build_ui(self) -> None:
        layout = QVBoxLayout(self)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(10)

        self.metrics = MetricStrip()
        self.chart = self._create_chart()
        layout.addWidget(self.metrics)
        layout.addWidget(self.chart)

        filter_row = QHBoxLayout()
        self.search_input = QLineEdit()
        self.search_input.setObjectName("AgribankSearchBox")
        self.search_input.setPlaceholderText("Tìm kiếm")
        self.search_input.setClearButtonEnabled(True)
        self.search_input.textChanged.connect(lambda: self.search_timer.start())
        filter_row.addWidget(self.search_input, stretch=1)
        self._add_filters(filter_row)
        layout.addLayout(filter_row)

        action_row = QHBoxLayout()
        self._add_action_buttons(action_row)
        self.progress = QProgressBar()
        self.progress.setVisible(False)
        action_row.addWidget(self.progress, stretch=1)
        layout.addLayout(action_row)

        self.table = QTableWidget()
        self.table.setObjectName("SummaryDataTable")
        self.table.setHorizontalHeader(MultiLineHeaderView(Qt.Orientation.Horizontal, self.table))
        self.table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.table.itemDoubleClicked.connect(self._table_item_double_clicked)
        header = self.table.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        header.setStretchLastSection(True)
        header.setSectionsMovable(True)
        header.setSectionsClickable(True)
        header.sectionResized.connect(self._save_column_width)
        header.sectionDoubleClicked.connect(self._autofit_column)
        if hasattr(header, "sectionHandleDoubleClicked"):
            header.sectionHandleDoubleClicked.connect(self._autofit_column)
        layout.addWidget(self.table, stretch=1)

        pager = QHBoxLayout()
        self.prev_button = _secondary_button("Trang trước")
        self.prev_button.clicked.connect(self.previous_page)
        self.next_button = _secondary_button("Trang sau")
        self.next_button.clicked.connect(self.next_page)
        self.page_label = QLabel("Trang 1")
        pager.addStretch()
        pager.addWidget(self.prev_button)
        pager.addWidget(self.page_label)
        pager.addWidget(self.next_button)
        layout.addLayout(pager)

    def _create_chart(self) -> QWidget:
        return MiniChart()

    def _add_filters(self, row: QHBoxLayout) -> None:
        _ = row

    def _add_action_buttons(self, row: QHBoxLayout) -> None:
        export_button = _secondary_button("Xuất")
        export_button.clicked.connect(self.export_current_rows)
        row.addWidget(export_button)
        refresh_button = _secondary_button("Làm mới")
        refresh_button.clicked.connect(self.reload)
        row.addWidget(refresh_button)
        backup_button = _secondary_button("Sao lưu")
        backup_button.clicked.connect(self.backup_data)
        row.addWidget(backup_button)
        restore_button = _secondary_button("Khôi phục")
        restore_button.clicked.connect(self.restore_data)
        row.addWidget(restore_button)
        maintenance_button = _secondary_button("Bảo trì dữ liệu")
        maintenance_button.clicked.connect(self.open_maintenance)
        row.addWidget(maintenance_button)

    def reload_filters(self) -> None:
        pass

    def query_page(self) -> PageResult:
        raise NotImplementedError

    def query_dashboard(self) -> DashboardData:
        raise NotImplementedError

    def reload(self) -> None:
        try:
            self.reload_filters()
            result = self.query_page()
            dashboard = self.query_dashboard()
        except Exception as exc:
            QMessageBox.warning(self, self.title, str(exc))
            return
        self.current_rows = result.rows
        self._render_table(result.rows)
        self.metrics.set_data(dashboard)
        self._update_chart(dashboard)
        total_pages = max(1, (result.total_rows + result.page_size - 1) // result.page_size)
        self.page_label.setText(f"Trang {result.page}/{total_pages} - {result.total_rows:,} dòng")
        self.prev_button.setEnabled(result.page > 1)
        self.next_button.setEnabled(result.page < total_pages)

    def _update_chart(self, dashboard: DashboardData) -> None:
        if hasattr(self.chart, "set_values"):
            self.chart.set_values(dashboard.bars or dashboard.lines or dashboard.pies)

    def previous_page(self) -> None:
        self.page = max(1, self.page - 1)
        self.reload()

    def next_page(self) -> None:
        self.page += 1
        self.reload()

    def export_current_rows(self) -> None:
        path, _ = QFileDialog.getSaveFileName(
            self,
            "Xuất báo cáo",
            f"{self.title}.xlsx",
            "Excel (*.xlsx);;PDF (*.pdf);;CSV (*.csv)",
        )
        if not path:
            return
        try:
            output = export_rows(
                self._export_rows(),
                Path(path),
                title=self.title,
                sheet_name=self.title,
            )
        except Exception as exc:
            QMessageBox.warning(self, "Xuất báo cáo", str(exc))
        return

    def _export_rows(self) -> list[dict[str, object]]:
        return self.current_rows
        QMessageBox.information(self, "Xuất báo cáo", f"Đã xuất: {output}")

    def backup_data(self) -> None:
        path, _ = QFileDialog.getSaveFileName(
            self,
            "Sao lưu dữ liệu",
            "summary-backup.zip",
            "Backup (*.zip)",
        )
        if not path:
            return
        try:
            output = self.repository.backup_database(Path(path))
        except Exception as exc:
            QMessageBox.warning(self, "Sao lưu dữ liệu", str(exc))
            return
        QMessageBox.information(self, "Sao lưu dữ liệu", f"Đã sao lưu: {output}")

    def restore_data(self) -> None:
        path, _ = QFileDialog.getOpenFileName(
            self,
            "Khôi phục dữ liệu",
            "",
            "Backup (*.zip)",
        )
        if not path:
            return
        if QMessageBox.question(
            self,
            "Khôi phục dữ liệu",
            "Khôi phục sẽ thay thế database hiện tại sau khi tạo một bản sao lưu an toàn. Tiếp tục?",
        ) != QMessageBox.StandardButton.Yes:
            return
        try:
            safety_backup = self.repository.restore_database(Path(path))
        except Exception as exc:
            QMessageBox.warning(self, "Khôi phục dữ liệu", str(exc))
            return
        QMessageBox.information(
            self,
            "Khôi phục dữ liệu",
            f"Đã khôi phục. Bản sao lưu an toàn: {safety_backup}",
        )
        self.reload()

    def open_maintenance(self) -> None:
        dialog = SummaryMaintenanceDialog(self.repository, parent=self)
        dialog.exec()
        self.reload()

    def run_background(self, title: str, function: Callable[[Callable[[str], None]], object], on_done: Callable[[object], None]) -> None:
        self.progress.setVisible(True)
        self.progress.setRange(0, 0)

        def wrapped(progress):
            return function(progress)

        def done(payload):
            self.progress.setVisible(False)
            self.progress.setRange(0, 100)
            on_done(payload)
            self.page = 1
            self.reload()

        def failed(exc: Exception):
            self.progress.setVisible(False)
            QMessageBox.warning(self, title, str(exc))

        run_in_thread(self, wrapped, done, failed, lambda message: self.progress.setFormat(message))

    def _render_table(self, rows: list[dict[str, object]]) -> None:
        headers = list(rows[0].keys()) if rows else []
        self.table.clear()
        self.table.setColumnCount(len(headers))
        self.table.setHorizontalHeaderLabels([self._display_header(header) for header in headers])
        self.table.setRowCount(len(rows))
        for r_index, row in enumerate(rows):
            for c_index, header in enumerate(headers):
                item = QTableWidgetItem(self._display_value(header, row.get(header, "")))
                item.setTextAlignment(self._display_alignment(header))
                self.table.setItem(r_index, c_index, item)
        QTimer.singleShot(0, self._restore_column_widths)

    def _display_header(self, header: str) -> str:
        return header

    def _display_value(self, header: str, value: object) -> str:
        _ = header
        return "" if value is None else str(value)

    def _display_alignment(self, header: str) -> Qt.AlignmentFlag:
        _ = header
        return Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter

    def _table_item_double_clicked(self, item: QTableWidgetItem) -> None:
        row_index = item.row()
        if 0 <= row_index < len(self.current_rows):
            self._row_double_clicked(self.current_rows[row_index])

    def _row_double_clicked(self, row: dict[str, object]) -> None:
        _ = row

    def _autofit_column(self, logical_index: int) -> None:
        self.table.resizeColumnToContents(logical_index)

    def _save_column_width(self, logical_index: int, old_size: int, new_size: int) -> None:
        _ = old_size
        if self._restoring_columns or logical_index < 0:
            return
        header = self.table.horizontalHeaderItem(logical_index)
        if header is None:
            return
        self.settings.setValue(self._column_width_key(header.text()), int(new_size))

    def _restore_column_widths(self) -> None:
        self._restoring_columns = True
        try:
            for index in range(self.table.columnCount()):
                header = self.table.horizontalHeaderItem(index)
                if header is None:
                    continue
                width = self.settings.value(self._column_width_key(header.text()), type=int)
                if width:
                    self.table.setColumnWidth(index, max(40, int(width)))
                else:
                    self.table.resizeColumnToContents(index)
            header = self.table.horizontalHeader()
            if hasattr(header, "refresh_height"):
                header.refresh_height()
        finally:
            self._restoring_columns = False

    def _column_width_key(self, header: str) -> str:
        return f"summary/{self.title}/columns/{header}/width"


class DeleteNimPeriodDialog(QDialog):
    def __init__(
        self,
        repository: SummaryRepository,
        data_type: SummaryDataType,
        parent: QWidget | None = None,
    ) -> None:
        super().__init__(parent)
        self.repository = repository
        self.data_type = data_type
        self.deleted_info: dict[str, object] | None = None
        self.setWindowTitle("Xóa dữ liệu NIM theo kỳ")
        self.setMinimumWidth(460)
        self._build_ui()
        self._reload_periods()

    def _build_ui(self) -> None:
        layout = QVBoxLayout(self)
        layout.setContentsMargins(14, 14, 14, 14)
        layout.setSpacing(10)

        form = QFormLayout()
        self.type_label = QLabel(_nim_delete_title(self.data_type))
        self.type_label.setObjectName("SectionTitle")
        self.period_combo = _combo("Kỳ", "Chọn kỳ")
        self.period_combo.currentIndexChanged.connect(lambda _index: self._reload_info())
        form.addRow("Loại dữ liệu", self.type_label)
        form.addRow("Kỳ cần xóa", self.period_combo)
        layout.addLayout(form)

        self.info_label = QLabel("")
        self.info_label.setWordWrap(True)
        self.info_label.setObjectName("MutedText")
        layout.addWidget(self.info_label)

        buttons = QHBoxLayout()
        buttons.addStretch()
        cancel_button = _secondary_button("Hủy")
        cancel_button.clicked.connect(self.reject)
        self.delete_button = _danger_button("Xóa dữ liệu")
        self.delete_button.clicked.connect(self._delete_current_period)
        buttons.addWidget(cancel_button)
        buttons.addWidget(self.delete_button)
        layout.addLayout(buttons)

    def _reload_periods(self) -> None:
        self.period_combo.blockSignals(True)
        self.period_combo.clear()
        periods = self.repository.nim_periods(self.data_type)
        if not periods:
            self.period_combo.addItem("Không có kỳ dữ liệu", "")
            self.delete_button.setEnabled(False)
        else:
            for period in periods:
                self.period_combo.addItem(period, period)
            self.period_combo.setCurrentIndex(len(periods) - 1)
            self.delete_button.setEnabled(True)
        self.period_combo.blockSignals(False)
        self._reload_info()

    def _reload_info(self) -> None:
        period = _current_filter(self.period_combo)
        if not period:
            self.info_label.setText("Không có dữ liệu NIM để xóa.")
            return
        try:
            info = self.repository.nim_period_info(self.data_type, period)
        except Exception as exc:
            self.info_label.setText(str(exc))
            return
        files = ", ".join(str(item) for item in info.get("source_files", ()) if item) or "Không có thông tin"
        self.info_label.setText(
            "Số batch: {batch_count}\n"
            "Số dòng dữ liệu: {row_count}\n"
            "Ngày import gần nhất: {latest_import_at}\n"
            "File nguồn: {files}".format(
                batch_count=_format_integer_vn(info.get("batch_count", 0)),
                row_count=_format_integer_vn(info.get("row_count", 0)),
                latest_import_at=info.get("latest_import_at") or "Không có thông tin",
                files=files,
            )
        )

    def _delete_current_period(self) -> None:
        period = _current_filter(self.period_combo)
        if not period:
            return
        info = self.repository.nim_period_info(self.data_type, period)
        if QMessageBox.question(
            self,
            "Xác nhận xóa kỳ dữ liệu",
            (
                f"Bạn có chắc muốn xóa toàn bộ dữ liệu {_nim_delete_title(self.data_type)} kỳ {period} không?\n\n"
                "Thao tác này sẽ xóa dữ liệu chi tiết, dữ liệu tổng hợp và các batch import thuộc kỳ đã chọn.\n\n"
                f"Loại NIM: {_nim_delete_title(self.data_type)}\n"
                f"Kỳ: {period}\n"
                f"Số dòng: {_format_integer_vn(info.get('row_count', 0))}\n"
                f"Số batch: {_format_integer_vn(info.get('batch_count', 0))}"
            ),
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        ) != QMessageBox.StandardButton.Yes:
            return
        try:
            self.deleted_info = self.repository.delete_nim_period(
                self.data_type,
                period,
                created_by=_current_user(),
            )
        except Exception as exc:
            QMessageBox.warning(self, "Xóa dữ liệu NIM theo kỳ", str(exc))
            return
        QMessageBox.information(self, "Xóa dữ liệu NIM theo kỳ", f"Đã xóa kỳ {period}.")
        self.accept()


class SummaryMaintenanceDialog(QDialog):
    def __init__(self, repository: SummaryRepository, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.repository = repository
        self.setWindowTitle("Bảo trì dữ liệu Tổng hợp số liệu")
        self.setMinimumWidth(620)
        self._build_ui()
        self.reload()

    def _build_ui(self) -> None:
        layout = QVBoxLayout(self)
        layout.setContentsMargins(14, 14, 14, 14)
        layout.setSpacing(10)

        self.status_label = QLabel("")
        self.status_label.setWordWrap(True)
        layout.addWidget(self.status_label)

        batch_row = QHBoxLayout()
        self.batch_row = batch_row
        batch_row.setSpacing(8)
        self.batch_combo = _combo("Batch", "Chọn batch để xóa")
        self.batch_combo.setMinimumWidth(360)
        self.batch_combo.setMaximumWidth(16777215)
        self.batch_combo.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        batch_row.addWidget(self.batch_combo, stretch=1)
        self.delete_batch_button = _danger_button("Xóa batch đã chọn")
        self.delete_batch_button.setSizePolicy(QSizePolicy.Policy.Preferred, QSizePolicy.Policy.Fixed)
        self.delete_batch_button.clicked.connect(self.delete_selected_batch)
        batch_row.addWidget(self.delete_batch_button)
        batch_row.addStretch(1)
        layout.addLayout(batch_row)

        self.progress = QProgressBar()
        self.progress.setVisible(False)
        layout.addWidget(self.progress)

        self.data_actions_widget = QWidget()
        self.data_actions_row = FlowLayout(self.data_actions_widget, spacing=8)
        self.refresh_button = _secondary_button("Làm mới")
        self.refresh_button.clicked.connect(self.reload)
        self.delete_dn_button = _danger_button("Xóa kỳ NIM Dư nợ")
        self.delete_dn_button.clicked.connect(lambda: self.open_delete_period(SummaryDataType.NIM_DN))
        self.delete_nv_button = _danger_button("Xóa kỳ NIM Nguồn vốn")
        self.delete_nv_button.clicked.connect(lambda: self.open_delete_period(SummaryDataType.NIM_NV))
        self.compact_nim_button = _danger_button("Xóa raw NIM đã tổng hợp")
        self.compact_nim_button.clicked.connect(self.compact_legacy_nim)
        for button in (
            self.refresh_button,
            self.delete_dn_button,
            self.delete_nv_button,
            self.compact_nim_button,
        ):
            button.setSizePolicy(QSizePolicy.Policy.Preferred, QSizePolicy.Policy.Fixed)
            self.data_actions_row.addWidget(button)
        layout.addWidget(self.data_actions_widget)

        self.system_actions_row = QHBoxLayout()
        self.system_actions_row.setSpacing(8)
        self.system_actions_flow_widget = QWidget()
        self.system_actions_flow = FlowLayout(self.system_actions_flow_widget, spacing=8)
        self.optimize_button = _secondary_button("Tối ưu cơ sở dữ liệu")
        self.optimize_button.clicked.connect(self.optimize_database)
        self.folder_button = _secondary_button("Mở thư mục database")
        self.folder_button.clicked.connect(self.open_database_folder)
        self.backup_button = _secondary_button("Sao lưu database")
        self.backup_button.clicked.connect(self.backup_database)
        self.close_button = _primary_button("Đóng")
        self.close_button.clicked.connect(self.accept)
        for button in (self.optimize_button, self.folder_button, self.backup_button):
            button.setSizePolicy(QSizePolicy.Policy.Preferred, QSizePolicy.Policy.Fixed)
            self.system_actions_flow.addWidget(button)
        self.system_actions_row.addWidget(self.system_actions_flow_widget, stretch=1)
        self.system_actions_row.addStretch(1)
        self.close_button.setSizePolicy(QSizePolicy.Policy.Preferred, QSizePolicy.Policy.Fixed)
        self.system_actions_row.addWidget(self.close_button)
        layout.addLayout(self.system_actions_row)

    def reload(self) -> None:
        status = self.repository.maintenance_status()
        self.status_label.setText(
            "Database: {path}\n"
            "Dung lượng: {size}\n"
            "Số kỳ NIM DN: {nim_dn}\n"
            "Số kỳ NIM NV: {nim_nv}\n"
            "Số dòng raw NIM legacy: {raw_nim_rows}\n"
            "Số batch so sánh tăng giảm: {loan_batches}\n"
            "Số batch hạn mức: {limit_batches}".format(
                path=status["database_path"],
                size=_format_size(int(status["size_bytes"] or 0)),
                nim_dn=_format_integer_vn(status["nim_dn_periods"]),
                nim_nv=_format_integer_vn(status["nim_nv_periods"]),
                raw_nim_rows=_format_integer_vn(status.get("raw_nim_rows", 0)),
                loan_batches=_format_integer_vn(status["loan_compare_batches"]),
                limit_batches=_format_integer_vn(status["credit_limit_batches"]),
            )
        )
        self.batch_combo.blockSignals(True)
        self.batch_combo.clear()
        self.batch_combo.addItem("Chọn batch để xóa", "")
        for batch in self.repository.list_batches(limit=200):
            self.batch_combo.addItem(
                f"{batch.id} - {batch.data_type} - {batch.period} - {batch.file_name}",
                batch.id,
            )
        self.batch_combo.blockSignals(False)

    def open_delete_period(self, data_type: SummaryDataType) -> None:
        dialog = DeleteNimPeriodDialog(self.repository, data_type, parent=self)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            self.reload()

    def delete_selected_batch(self) -> None:
        batch_id = self.batch_combo.currentData()
        if not batch_id:
            QMessageBox.information(self, self.windowTitle(), "Chưa chọn batch cần xóa.")
            return
        if QMessageBox.question(
            self,
            "Xác nhận xóa batch",
            f"Bạn có chắc muốn xóa batch {batch_id} không?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        ) != QMessageBox.StandardButton.Yes:
            return
        try:
            result = self.repository.delete_batch(int(batch_id), created_by=_current_user())
        except Exception as exc:
            QMessageBox.warning(self, self.windowTitle(), str(exc))
            return
        QMessageBox.information(
            self,
            self.windowTitle(),
            f"Đã xóa batch {result['batch_id']} ({_format_integer_vn(result['row_count'])} dòng).",
        )
        self.reload()

    def compact_legacy_nim(self) -> None:
        if QMessageBox.question(
            self,
            "Xóa raw NIM đã tổng hợp",
            "Chỉ xóa các dòng raw NIM legacy sau khi bảng tổng hợp đã được đối soát. Tiếp tục và chạy VACUUM?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        ) != QMessageBox.StandardButton.Yes:
            return
        self.progress.setVisible(True)
        self.progress.setRange(0, 0)

        def done(result: object) -> None:
            self.progress.setVisible(False)
            data = result if isinstance(result, dict) else {}
            QMessageBox.information(
                self,
                "Xóa raw NIM đã tổng hợp",
                "Đã xóa {rows} dòng raw NIM. Dung lượng: {before} -> {after}".format(
                    rows=_format_integer_vn(data.get("deleted_rows", 0)),
                    before=_format_size(int(data.get("before_size_bytes", 0))),
                    after=_format_size(int(data.get("after_size_bytes", 0))),
                ),
            )
            self.reload()

        def failed(exc: Exception) -> None:
            self.progress.setVisible(False)
            QMessageBox.warning(self, "Xóa raw NIM đã tổng hợp", str(exc))

        run_in_thread(
            self,
            lambda progress: self.repository.compact_legacy_nim_details(vacuum=True, created_by=_current_user()),
            done,
            failed,
            lambda message: self.progress.setFormat(message),
        )

    def optimize_database(self) -> None:
        vacuum = (
            QMessageBox.question(
                self,
                "Tối ưu cơ sở dữ liệu",
                "Chạy VACUUM để thu hồi dung lượng sau khi xóa dữ liệu?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No,
            )
            == QMessageBox.StandardButton.Yes
        )
        self.progress.setVisible(True)
        self.progress.setRange(0, 0)

        def done(result: object) -> None:
            self.progress.setVisible(False)
            data = result if isinstance(result, dict) else {}
            QMessageBox.information(
                self,
                "Tối ưu cơ sở dữ liệu",
                "Đã tối ưu. Dung lượng: {before} -> {after}".format(
                    before=_format_size(int(data.get("before_size_bytes", 0))),
                    after=_format_size(int(data.get("after_size_bytes", 0))),
                ),
            )
            self.reload()

        def failed(exc: Exception) -> None:
            self.progress.setVisible(False)
            QMessageBox.warning(self, "Tối ưu cơ sở dữ liệu", str(exc))

        run_in_thread(
            self,
            lambda progress: self.repository.optimize_database(vacuum=vacuum),
            done,
            failed,
            lambda message: self.progress.setFormat(message),
        )

    def open_database_folder(self) -> None:
        QDesktopServices.openUrl(QUrl.fromLocalFile(str(self.repository.database_path.parent)))

    def backup_database(self) -> None:
        path, _ = QFileDialog.getSaveFileName(
            self,
            "Sao lưu dữ liệu Tổng hợp số liệu",
            "CreditSummary-backup.zip",
            "Backup (*.zip)",
        )
        if not path:
            return
        try:
            output = self.repository.backup_database(Path(path))
        except Exception as exc:
            QMessageBox.warning(self, "Sao lưu dữ liệu", str(exc))
            return
        QMessageBox.information(self, "Sao lưu dữ liệu", f"Đã sao lưu: {output}")
        self.reload()


class NimTab(SummaryDataTab):
    BASE_COLUMN_WIDTHS = {
        "Kỳ": 72,
        "Tên chi nhánh": 145,
        "Phòng GD": 105,
        "Loại KH": 78,
        "Người quản lý KV": 155,
        "Người quản lý NV": 155,
        "Dư nợ": 118,
        "Số dư nguồn vốn": 124,
        "Lãi suất bình quân": 96,
        "NIM trước ĐC": 92,
        "NIM sau ĐC": 92,
    }

    def __init__(self, repository: SummaryRepository, data_type: SummaryDataType, parent: QWidget | None = None) -> None:
        self.data_type = data_type
        self.ui_config = get_nim_ui_config(data_type)
        self.title = self.ui_config.main_title
        self.history_dialogs: list[OfficerHistoryDialog] = []
        self.dashboard_windows: list[QDialog] = []
        self.customer_management_window: QDialog | None = None
        self.unit_directory: UnitDirectoryService
        super().__init__(repository, parent)
        self.unit_directory = get_unit_directory_service(repository.main_database_path)
        self._unit_directory_listener = self.reload
        self.unit_directory.add_listener(self._unit_directory_listener)
        self.destroyed.connect(lambda *_args: self.unit_directory.remove_listener(self._unit_directory_listener))
        self.table.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.table.horizontalHeader().setStretchLastSection(False)
        self.table.viewport().installEventFilter(self)
        self.reload()

    def _create_chart(self) -> QWidget:
        return NimTrendChart()

    def _add_filters(self, row: QHBoxLayout) -> None:
        self.period_filter = _combo("Kỳ")
        self.branch_filter = _combo("Chi nhánh")
        self.transaction_office_filter = _combo("Phòng GD")
        self.customer_type_filter = _combo("Loại KH")
        self.officer_filter = _combo(self.ui_config.officer_short_label)
        for combo in (
            self.period_filter,
            self.branch_filter,
            self.transaction_office_filter,
            self.customer_type_filter,
            self.officer_filter,
        ):
            combo.currentTextChanged.connect(self._filter_changed)
            row.addWidget(combo)

    def _add_action_buttons(self, row: QHBoxLayout) -> None:
        import_button = _primary_button("Import thư mục")
        import_button.clicked.connect(self.import_folder)
        row.addWidget(import_button)
        super()._add_action_buttons(row)
        delete_button = _danger_button("Xóa kỳ dữ liệu")
        delete_button.clicked.connect(self.delete_period)
        row.addWidget(delete_button)
        dashboard_button = _secondary_button("Mở Dashboard")
        dashboard_button.clicked.connect(self.open_dashboard)
        row.addWidget(dashboard_button)
        if self.data_type == SummaryDataType.NIM_DN:
            customer_button = _secondary_button("Dữ liệu khách hàng")
            customer_button.clicked.connect(self.open_customer_management)
            row.addWidget(customer_button)

    def eventFilter(self, watched, event) -> bool:
        if watched == self.table.viewport() and event.type() == QEvent.Type.Resize:
            QTimer.singleShot(0, self._restore_column_widths)
        return super().eventFilter(watched, event)

    def reload_filters(self) -> None:
        filters = self._filters()
        _populate_combo(
            self.period_filter,
            self._distinct_values("period", filters, exclude="period"),
        )
        _populate_combo(
            self.branch_filter,
            self._branch_filter_items(filters, exclude="branch"),
        )
        _populate_combo(
            self.transaction_office_filter,
            self._office_filter_items(filters, exclude="transaction_office"),
        )
        _populate_combo(
            self.customer_type_filter,
            self._customer_type_items(self._distinct_values("customer_type", filters, exclude="customer_type")),
        )
        _populate_combo(
            self.officer_filter,
            [
                (_display_officer_name(value), value)
                for value in self._distinct_values("officer", filters, exclude="officer")
            ],
        )

    def query_page(self) -> PageResult:
        filters = self._filters()
        where, params = self._nim_where(filters, search=self.search_input.text())
        page = max(1, int(self.page))
        page_size = max(10, min(2000, int(self.page_size)))
        offset = (page - 1) * page_size
        has_customer_type_filter = bool(filters.get("customer_type"))
        customer_type_select = "customer_type" if has_customer_type_filter else "'Tất cả' AS customer_type"
        group_columns = "period, branch_code, trctcd, officer_code, officer_name"
        if has_customer_type_filter:
            group_columns += ", customer_type"
        group_sql = f"""
            SELECT
                period,
                branch_code,
                trctcd,
                MIN(branch_name) AS branch,
                MIN(transaction_office) AS transaction_office,
                {customer_type_select},
                {NIM_OFFICER_DISPLAY_SQL} AS officer,
                SUM(balance) AS balance,
                CASE WHEN SUM(balance) <> 0 THEN SUM(interest_rate_numerator) / SUM(balance) ELSE 0 END AS average_rate,
                CASE WHEN SUM(balance) <> 0 THEN SUM(numerator_before) / SUM(balance) ELSE 0 END AS nim_before,
                CASE WHEN SUM(balance) <> 0 THEN SUM(numerator_after) / SUM(balance) ELSE 0 END AS nim_after
            FROM nim_period_summary
            {where}
            GROUP BY {group_columns}
        """
        order_sql = """
            ORDER BY period DESC,
                     branch_code COLLATE NOCASE,
                     transaction_office COLLATE NOCASE,
                     customer_type COLLATE NOCASE,
                     officer COLLATE NOCASE
        """
        with closing(self.repository.connect()) as database:
            total = int(database.execute(f"SELECT COUNT(*) FROM ({group_sql}) AS q", params).fetchone()[0])
            rows = database.execute(
                f"{group_sql} {order_sql} LIMIT ? OFFSET ?",
                (*params, page_size, offset),
            ).fetchall()
        output_rows = [self._visible_row(self._dynamic_unit_row(dict(row))) for row in rows]
        return PageResult(
            rows=output_rows,
            total_rows=total,
            page=page,
            page_size=page_size,
        )

    def query_dashboard(self) -> DashboardData:
        row = self._summary_row(self._filters())
        metrics = [
            DashboardMetric(self.ui_config.total_balance_label, _format_money_vn(row.get("balance", 0))),
            DashboardMetric("NIM trước ĐC", _format_percent_vn(row.get("nim_before", 0))),
            DashboardMetric("NIM sau ĐC", _format_percent_vn(row.get("nim_after", 0))),
        ]
        if self.ui_config.include_average_rate:
            metrics.append(DashboardMetric("Lãi suất bình quân", _format_percent_vn(row.get("average_rate", 0))))
        return DashboardData(metrics=tuple(metrics))

    def _update_chart(self, dashboard: DashboardData) -> None:
        _ = dashboard
        if isinstance(self.chart, NimTrendChart):
            self.chart.set_trend(self._trend_rows(self._filters()))

    def import_folder(self) -> None:
        folder = QFileDialog.getExistingDirectory(self, "Chọn thư mục chứa CSV NIM")
        if not folder:
            return
        service = import_nim_dn if self.data_type == SummaryDataType.NIM_DN else import_nim_nv
        replace_existing_periods = False
        if self.data_type == SummaryDataType.NIM_DN:
            from agribank_v3.features.credit.summary.customer.repository import CustomerRepository

            folder_path = Path(folder)
            periods = sorted(
                {
                    parse_period_from_filename(path.name)
                    for path in folder_path.iterdir()
                    if path.is_file()
                    and path.suffix.casefold() == ".csv"
                    and "_ftpln_" in path.name.casefold()
                }
            )
            existing_periods = CustomerRepository(self.repository.main_database_path).periods_with_data(periods)
            if existing_periods:
                answer = QMessageBox.question(
                    self,
                    self.title,
                    "Dữ liệu khách hàng kỳ "
                    f"{', '.join(existing_periods)} đã tồn tại. "
                    "Bạn có muốn ghi đè toàn bộ dữ liệu kỳ này không?",
                    QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                    QMessageBox.StandardButton.No,
                )
                if answer != QMessageBox.StandardButton.Yes:
                    return
                replace_existing_periods = True
        self.run_background(
            self.title,
            lambda progress: service(
                self.repository,
                Path(folder),
                export_path=_nim_vba_export_path(self.data_type),
                progress=progress,
                replace_existing_periods=replace_existing_periods,
            ),
            lambda result: QMessageBox.information(
                self,
                self.title,
                _result_message(result, "Import xong."),
            ),
        )

    def open_dashboard(self) -> None:
        from agribank_v3.features.credit.summary.dashboard_window import NimDashboardWindow

        dialog = NimDashboardWindow(self.repository, self.data_type, parent=self)
        dialog.setModal(False)
        self.dashboard_windows.append(dialog)
        dialog.finished.connect(lambda _result, item=dialog: self._forget_dashboard_window(item))
        dialog.show()
        dialog.raise_()

    def open_customer_management(self) -> None:
        if self.data_type != SummaryDataType.NIM_DN:
            return
        from agribank_v3.features.credit.summary.customer.window_controller import open_customer_management_window

        self.customer_management_window = open_customer_management_window(
            self,
            self.repository.main_database_path,
            open_nim_dn_callback=lambda: self.window().raise_(),
        )

    def delete_period(self) -> None:
        dialog = DeleteNimPeriodDialog(self.repository, self.data_type, parent=self)
        if dialog.exec() != QDialog.DialogCode.Accepted:
            return
        remaining = self.repository.nim_periods(self.data_type)
        target_period = max(remaining) if remaining else ""
        self.page = 1
        self.reload()
        if target_period:
            index = self.period_filter.findData(target_period)
            if index >= 0:
                self.period_filter.setCurrentIndex(index)
        for dashboard in list(self.dashboard_windows):
            if hasattr(dashboard, "reload"):
                dashboard.reload()

    def _filter_changed(self) -> None:
        self.page = 1
        self.reload()

    def _filters(self) -> dict[str, object]:
        return {
            "period": _current_filter(self.period_filter),
            "branch": _current_filter(self.branch_filter),
            "transaction_office": _current_filter(self.transaction_office_filter),
            "customer_type": _current_filter(self.customer_type_filter),
            "officer": _current_filter(self.officer_filter),
        }

    def _display_header(self, header: str) -> str:
        return self._header_labels().get(header, header)

    def _render_table(self, rows: list[dict[str, object]]) -> None:
        display_rows = [
            {key: value for key, value in row.items() if key not in {"branch_code", "trctcd"}}
            for row in rows
        ]
        super()._render_table(display_rows)

    def _export_rows(self) -> list[dict[str, object]]:
        output: list[dict[str, object]] = []
        for row in self.current_rows:
            item = {
                self._display_header(key): self._export_value(key, value)
                for key, value in row.items()
                if key not in {"branch_code", "trctcd"}
            }
            output.append(item)
        return output

    def _export_value(self, key: str, value: object) -> object:
        if key == "officer":
            return _display_officer_name(value)
        return value

    def _display_value(self, header: str, value: object) -> str:
        if header == "officer":
            return _display_officer_name(value)
        if header in {"branch_code", "trctcd"}:
            return ""
        if header == "balance":
            return _format_money_vn(value)
        if header in {"average_rate", "nim_before", "nim_after"}:
            return _format_percent_vn(value)
        return str(value or "")

    def _display_alignment(self, header: str) -> Qt.AlignmentFlag:
        if header == "period":
            return Qt.AlignmentFlag.AlignCenter
        if header in {"balance", "average_rate", "nim_before", "nim_after"}:
            return Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter
        return super()._display_alignment(header)

    def _visible_row(self, row: dict[str, object]) -> dict[str, object]:
        if self.ui_config.include_average_rate:
            return row
        row.pop("average_rate", None)
        return row

    def _header_labels(self) -> dict[str, str]:
        labels = {
            "period": "Kỳ",
            "branch_code": "",
            "trctcd": "",
            "branch": "Tên chi nhánh",
            "transaction_office": "Phòng GD",
            "customer_type": "Loại KH",
            "officer": self.ui_config.officer_label,
            "balance": self.ui_config.balance_label,
            "nim_before": "NIM trước ĐC",
            "nim_after": "NIM sau ĐC",
        }
        if self.ui_config.include_average_rate:
            labels["average_rate"] = "Lãi suất bình quân"
        return labels

    def _row_double_clicked(self, row: dict[str, object]) -> None:
        officer = str(row.get("officer") or "").strip()
        if not officer:
            return
        branch_code = str(row.get("branch_code") or "").strip()
        trctcd = str(row.get("trctcd") or "").strip()
        dialog = OfficerHistoryDialog(
            self.repository,
            self.data_type,
            officer=officer,
            branch=self.unit_directory.get_branch_display_name(branch_code) if branch_code else str(row.get("branch") or ""),
            transaction_office=self.unit_directory.get_office_name(branch_code, trctcd) if branch_code else str(row.get("transaction_office") or ""),
            customer_type=str(row.get("customer_type") or ""),
            parent=self,
        )
        self.history_dialogs.append(dialog)
        dialog.finished.connect(lambda _result, item=dialog: self._forget_history_dialog(item))
        dialog.show()
        dialog.raise_()

    def _forget_history_dialog(self, dialog: OfficerHistoryDialog) -> None:
        if dialog in self.history_dialogs:
            self.history_dialogs.remove(dialog)

    def _forget_dashboard_window(self, dialog: QDialog) -> None:
        if dialog in self.dashboard_windows:
            self.dashboard_windows.remove(dialog)

    def _forget_customer_management_window(self) -> None:
        self.customer_management_window = None

    def _restore_column_widths(self) -> None:
        self._restoring_columns = True
        try:
            column_count = self.table.columnCount()
            if column_count <= 0:
                return
            base_widths: list[int] = []
            for index in range(column_count):
                header = self.table.horizontalHeaderItem(index)
                base_widths.append(self.BASE_COLUMN_WIDTHS.get(header.text() if header else "", 96))
            available = max(240, self.table.viewport().width() - 2)
            widths = _fit_column_widths(base_widths, available)
            for index, width in enumerate(widths):
                self.table.setColumnWidth(index, width)
            header = self.table.horizontalHeader()
            if hasattr(header, "refresh_height"):
                header.refresh_height()
        finally:
            self._restoring_columns = False

    def _column_width_key(self, header: str) -> str:
        return f"summary/{self.title}/columns/v2/{header}/width"

    def _distinct_values(
        self,
        column_name: str,
        filters: dict[str, object],
        *,
        exclude: str,
    ) -> list[str]:
        allowed = {
            "period",
            "branch_code",
            "trctcd",
            "branch_name",
            "transaction_office",
            "customer_type",
            "officer",
        }
        if column_name not in allowed:
            raise SummaryError("Trường lọc NIM không hợp lệ.")
        where, params = self._nim_where(filters, exclude=exclude)
        select_column = f"{column_name} AS value"
        not_empty_column = column_name
        if column_name == "officer":
            select_column = f"{NIM_OFFICER_DISPLAY_SQL} AS value"
            not_empty_column = "officer_name"
        with closing(self.repository.connect()) as database:
            rows = database.execute(
                f"""
                SELECT DISTINCT {select_column}
                FROM nim_period_summary
                {where} AND {not_empty_column} <> ''
                ORDER BY value COLLATE NOCASE
                """,
                params,
            ).fetchall()
        return [str(row["value"]) for row in rows]

    def _branch_filter_items(self, filters: dict[str, object], *, exclude: str) -> list[tuple[str, str]]:
        where, params = self._nim_where(filters, exclude=exclude)
        with closing(self.repository.connect()) as database:
            rows = database.execute(
                f"""
                SELECT DISTINCT branch_code
                FROM nim_period_summary
                {where} AND branch_code <> ''
                ORDER BY branch_code COLLATE NOCASE
                """,
                params,
            ).fetchall()
        return [
            (self.unit_directory.get_branch_display_name(row["branch_code"]), str(row["branch_code"] or ""))
            for row in rows
        ]

    def _office_filter_items(self, filters: dict[str, object], *, exclude: str) -> list[tuple[str, str]]:
        where, params = self._nim_where(filters, exclude=exclude)
        with closing(self.repository.connect()) as database:
            rows = database.execute(
                f"""
                SELECT DISTINCT branch_code, trctcd
                FROM nim_period_summary
                {where} AND branch_code <> '' AND trctcd <> ''
                ORDER BY branch_code COLLATE NOCASE, trctcd COLLATE NOCASE
                """,
                params,
            ).fetchall()
        return [
            (
                self.unit_directory.get_office_display_name(row["branch_code"], row["trctcd"]),
                f"{str(row['branch_code'] or '')}-{str(row['trctcd'] or '')}",
            )
            for row in rows
        ]

    def _summary_row(self, filters: dict[str, object]) -> dict[str, object]:
        where, params = self._nim_where(filters)
        with closing(self.repository.connect()) as database:
            row = database.execute(
                f"""
                SELECT
                    COUNT(*) AS rows_count,
                    SUM(balance) AS balance,
                    CASE WHEN SUM(balance) <> 0 THEN SUM(numerator_before) / SUM(balance) ELSE 0 END AS nim_before,
                    CASE WHEN SUM(balance) <> 0 THEN SUM(numerator_after) / SUM(balance) ELSE 0 END AS nim_after,
                    CASE WHEN SUM(balance) <> 0 THEN SUM(interest_rate_numerator) / SUM(balance) ELSE 0 END AS average_rate
                FROM nim_period_summary
                {where}
                """,
                params,
            ).fetchone()
        return dict(row or {})

    def _trend_rows(self, filters: dict[str, object]) -> tuple[tuple[str, float, float], ...]:
        chart_filters = dict(filters)
        chart_filters["period"] = ""
        where, params = self._nim_where(chart_filters)
        with closing(self.repository.connect()) as database:
            rows = database.execute(
                f"""
                SELECT
                    period,
                    CASE WHEN SUM(balance) <> 0 THEN SUM(numerator_before) / SUM(balance) ELSE 0 END AS nim_before,
                    CASE WHEN SUM(balance) <> 0 THEN SUM(numerator_after) / SUM(balance) ELSE 0 END AS nim_after
                FROM nim_period_summary
                {where}
                GROUP BY period
                ORDER BY period
                """,
                params,
            ).fetchall()
        return tuple((str(row["period"]), float(row["nim_before"] or 0), float(row["nim_after"] or 0)) for row in rows)

    def _nim_where(
        self,
        filters: dict[str, object],
        *,
        exclude: str = "",
        search: str = "",
    ) -> tuple[str, list[object]]:
        clauses = ["data_type = ?"]
        params: list[object] = [self.data_type.value]
        for key, column in (
            ("period", "period"),
            ("customer_type", "customer_type"),
        ):
            if key == exclude:
                continue
            value = filters.get(key)
            if value in (None, "", 0):
                continue
            clauses.append(f"{column} = ?")
            params.append(value)
        if exclude != "branch":
            branch_value = filters.get("branch")
            if branch_value not in (None, "", 0):
                clauses.append("branch_code = ?")
                params.append(branch_value)
        if exclude != "transaction_office":
            office_value = str(filters.get("transaction_office") or "").strip()
            if office_value:
                branch_code, sep, trctcd = office_value.partition("-")
                if sep:
                    clauses.append("branch_code = ?")
                    clauses.append("trctcd = ?")
                    params.extend([branch_code, trctcd])
                else:
                    clauses.append("transaction_office = ?")
                    params.append(office_value)
        if exclude != "officer":
            officer_filter = filters.get("officer")
            if officer_filter not in (None, "", 0):
                _add_nim_officer_clause(clauses, params, officer_filter)
        if search.strip():
            needle = f"%{search.strip()}%"
            clauses.append(
                "(branch_code LIKE ? OR branch_name LIKE ? OR trctcd LIKE ? OR transaction_office LIKE ? OR customer_type LIKE ? "
                "OR officer_code LIKE ? OR officer_name LIKE ? OR period LIKE ?)"
            )
            params.extend([needle, needle, needle, needle, needle, needle, needle, needle])
        return "WHERE " + " AND ".join(clauses), params

    def _dynamic_unit_row(self, row: dict[str, object]) -> dict[str, object]:
        branch_code = str(row.get("branch_code") or "").strip()
        trctcd = str(row.get("trctcd") or "").strip()
        if branch_code:
            row["branch"] = self.unit_directory.get_branch_display_name(branch_code)
        if branch_code and trctcd:
            row["transaction_office"] = self.unit_directory.get_office_name(branch_code, trctcd)
        return row

    @staticmethod
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


class NimWindow(QDialog):
    def __init__(
        self,
        parent: QWidget | None = None,
        data_type: SummaryDataType | None = None,
    ) -> None:
        super().__init__(parent)
        self.setWindowTitle(_nim_window_title(data_type))
        self.setWindowFlags(
            self.windowFlags()
            | Qt.WindowType.WindowMinimizeButtonHint
            | Qt.WindowType.WindowMaximizeButtonHint
            | Qt.WindowType.WindowCloseButtonHint
        )
        self.setSizeGripEnabled(True)
        self.setMinimumSize(920, 620)
        self.resize(1180, 760)
        repository = SummaryRepository(_database_path(parent))
        layout = QVBoxLayout(self)
        if data_type is None:
            tabs = QTabWidget()
            tabs.addTab(NimTab(repository, SummaryDataType.NIM_DN), NIM_DN_CONFIG.title)
            tabs.addTab(NimTab(repository, SummaryDataType.NIM_NV), NIM_NV_CONFIG.title)
            layout.addWidget(tabs)
        else:
            layout.addWidget(NimTab(repository, data_type))


class LoanCompareWindow(QDialog):
    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setWindowTitle(f"{LOAN_COMPARE_TITLE} - AgribankV3")
        self.resize(1180, 760)
        self.repository = SummaryRepository(_database_path(parent))
        layout = QVBoxLayout(self)
        self.tab = LoanCompareTab(self.repository)
        layout.addWidget(self.tab)
        self.tab.reload()


class LoanCompareTab(SummaryDataTab):
    title = LOAN_COMPARE_TITLE

    def __init__(self, repository: SummaryRepository, parent: QWidget | None = None) -> None:
        super().__init__(repository, parent)
        self.table.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.table.horizontalHeader().setStretchLastSection(False)
        self.table.viewport().installEventFilter(self)

    def _add_filters(self, row: QHBoxLayout) -> None:
        self.batch_filter = _combo("Batch")
        self.category_filter = _combo("Loại", "Tất cả loại")
        self.officer_filter = _combo("CBTD")
        for combo in (self.batch_filter, self.category_filter, self.officer_filter):
            combo.currentTextChanged.connect(self._filter_changed)
            row.addWidget(combo)

    def _add_action_buttons(self, row: QHBoxLayout) -> None:
        import_button = _primary_button("Đối chiếu 2 file")
        import_button.clicked.connect(self.import_files)
        row.addWidget(import_button)
        super()._add_action_buttons(row)

    def reload_filters(self) -> None:
        batches = self.repository.list_batches(SummaryDataType.LOAN_COMPARE)
        values = [f"{batch.id} - {batch.period}" for batch in batches]
        _populate_combo(self.batch_filter, values)
        _populate_combo(self.category_filter, _loan_compare_category_items(self.repository.distinct_values("loan_compare_details", "category")))
        _populate_combo(self.officer_filter, self.repository.distinct_values("loan_compare_details", "officer"))

    def query_page(self) -> PageResult:
        return self.repository.query_loan_compare(
            batch_id=_batch_id(self.batch_filter),
            search=self.search_input.text(),
            category=_current_filter(self.category_filter),
            officer=_current_filter(self.officer_filter),
            page=self.page,
            page_size=self.page_size,
        )

    def query_dashboard(self) -> DashboardData:
        return self.repository.dashboard_loan_compare(_batch_id(self.batch_filter))

    def _update_chart(self, dashboard: DashboardData) -> None:
        if hasattr(self.chart, "set_values"):
            values = dashboard.bars or dashboard.lines or dashboard.pies
            self.chart.set_values(
                tuple((_display_loan_compare_category(label), value) for label, value in values)
            )

    def import_files(self) -> None:
        previous, _ = QFileDialog.getOpenFileName(self, "Chọn file kỳ trước", "", "Data (*.csv *.txt *.xlsx)")
        if not previous:
            return
        current, _ = QFileDialog.getOpenFileName(self, "Chọn file kỳ này", "", "Data (*.csv *.txt *.xlsx)")
        if not current:
            return
        self.run_background(
            self.title,
            lambda progress: compare_loan_balances(
                self.repository,
                Path(previous),
                Path(current),
                export_path=_vba_output_path("DuLieu", "BaoCaoTangGiamKH.xlsx"),
                progress=progress,
            ),
            lambda result: QMessageBox.information(
                self,
                self.title,
                _result_message(result, "Đối chiếu xong."),
            ),
        )

    def _filter_changed(self) -> None:
        self.page = 1
        self.reload()

    def eventFilter(self, watched, event) -> bool:
        if watched == self.table.viewport() and event.type() == QEvent.Type.Resize:
            QTimer.singleShot(0, self._restore_column_widths)
        return super().eventFilter(watched, event)

    def _render_table(self, rows: list[dict[str, object]]) -> None:
        self.table.clear()
        self.table.setColumnCount(len(LOAN_COMPARE_FIELD_ORDER))
        self.table.setHorizontalHeaderLabels([LOAN_COMPARE_HEADERS[field] for field in LOAN_COMPARE_FIELD_ORDER])
        self.table.setRowCount(len(rows))
        for row_index, row in enumerate(rows):
            for column_index, field in enumerate(LOAN_COMPARE_FIELD_ORDER):
                raw_value = _loan_compare_row_value(row, field)
                text = self._display_value(field, raw_value)
                if field in LOAN_COMPARE_MONEY_FIELDS:
                    item = NumericTableWidgetItem(text, float(raw_value or 0))
                else:
                    item = QTableWidgetItem(text)
                item.setTextAlignment(self._display_alignment(field))
                self.table.setItem(row_index, column_index, item)
        QTimer.singleShot(0, self._restore_column_widths)

    def _display_header(self, header: str) -> str:
        return LOAN_COMPARE_HEADERS.get(header, header)

    def _display_value(self, header: str, value: object) -> str:
        if header in {"customer_code"}:
            return _display_customer_code(value)
        if header in LOAN_COMPARE_MONEY_FIELDS:
            return _format_money_vn(value)
        if header == "category":
            return _display_loan_compare_category(value)
        return "" if value is None else str(value)

    def _display_alignment(self, header: str) -> Qt.AlignmentFlag:
        if header in LOAN_COMPARE_MONEY_FIELDS:
            return Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter
        return Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter

    def _export_rows(self) -> list[dict[str, object]]:
        output: list[dict[str, object]] = []
        for row in self.current_rows:
            output.append(
                {
                    LOAN_COMPARE_HEADERS[field]: self._export_value(field, _loan_compare_row_value(row, field))
                    for field in LOAN_COMPARE_FIELD_ORDER
                }
            )
        return output

    def _export_value(self, field: str, value: object) -> object:
        if field == "customer_code":
            return _display_customer_code(value)
        if field == "category":
            return _display_loan_compare_category(value)
        if field in LOAN_COMPARE_MONEY_FIELDS:
            try:
                return int(round(float(value or 0)))
            except (TypeError, ValueError):
                return 0
        return value

    def _restore_column_widths(self) -> None:
        self._restoring_columns = True
        try:
            available = max(320, self.table.viewport().width() - 2)
            fitted = _fit_column_widths(list(LOAN_COMPARE_COLUMN_WIDTHS), available)
            for index, default_width in enumerate(fitted):
                header = self.table.horizontalHeaderItem(index)
                saved = self.settings.value(self._column_width_key(header.text()), type=int) if header is not None else None
                self.table.setColumnWidth(index, max(54, int(saved or default_width)))
            header = self.table.horizontalHeader()
            if hasattr(header, "refresh_height"):
                header.refresh_height()
        finally:
            self._restoring_columns = False


class CreditLimitWindow(QDialog):
    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setWindowTitle(CREDIT_LIMIT_TITLE)
        self.resize(1180, 760)
        self.repository = SummaryRepository(_database_path(parent))
        layout = QVBoxLayout(self)
        self.tab = CreditLimitTab(self.repository)
        layout.addWidget(self.tab)
        self.tab.reload()


class CreditLimitTab(SummaryDataTab):
    title = CREDIT_LIMIT_TITLE

    def _add_filters(self, row: QHBoxLayout) -> None:
        self.batch_filter = _combo("Batch")
        self.status_filter = _combo("Trạng thái")
        self.officer_filter = _combo("CBTD")
        for combo in (self.batch_filter, self.status_filter, self.officer_filter):
            combo.currentTextChanged.connect(self._filter_changed)
            row.addWidget(combo)

    def _add_action_buttons(self, row: QHBoxLayout) -> None:
        group = QGroupBox("Tham số")
        form = QFormLayout(group)
        self.min_limit_input = QDoubleSpinBox()
        self.min_limit_input.setRange(0, 10_000_000_000_000)
        self.min_limit_input.setDecimals(0)
        self.min_limit_input.setSingleStep(100_000_000)
        self.warn_days_input = QSpinBox()
        self.warn_days_input.setRange(0, 3660)
        self.warn_days_input.setValue(30)
        self.reference_date_input = QDateEdit(QDate.currentDate())
        self.reference_date_input.setCalendarPopup(True)
        form.addRow("Hạn mức tối thiểu", self.min_limit_input)
        form.addRow("Ngày cảnh báo", self.warn_days_input)
        form.addRow("Ngày tham chiếu", self.reference_date_input)
        row.addWidget(group)
        import_button = _primary_button("Import LN01")
        import_button.clicked.connect(self.import_file)
        row.addWidget(import_button)
        super()._add_action_buttons(row)

    def reload_filters(self) -> None:
        batches = self.repository.list_batches(SummaryDataType.CREDIT_LIMIT)
        values = [f"{batch.id} - {batch.period}" for batch in batches]
        _populate_combo(self.batch_filter, values)
        _populate_combo(self.status_filter, self.repository.distinct_values("credit_limit_details", "status"))
        _populate_combo(self.officer_filter, self.repository.distinct_values("credit_limit_details", "officer"))

    def query_page(self) -> PageResult:
        return self.repository.query_credit_limits(
            batch_id=_batch_id(self.batch_filter),
            search=self.search_input.text(),
            status=_current_filter(self.status_filter),
            officer=_current_filter(self.officer_filter),
            page=self.page,
            page_size=self.page_size,
        )

    def query_dashboard(self) -> DashboardData:
        return self.repository.dashboard_credit_limits(_batch_id(self.batch_filter))

    def import_file(self) -> None:
        path, _ = QFileDialog.getOpenFileName(self, "Chọn file LN01", "", "CSV (*.csv)")
        if not path:
            return
        qdate = self.reference_date_input.date()
        reference = date(qdate.year(), qdate.month(), qdate.day())
        self.run_background(
            self.title,
            lambda progress: import_credit_limit_file(
                self.repository,
                Path(path),
                min_limit=self.min_limit_input.value(),
                warn_days=self.warn_days_input.value(),
                reference_date=reference,
                export_path=_vba_output_path("DuLieu", "BaoCaoHanMucHetHan.xlsx"),
                progress=progress,
            ),
            lambda result: QMessageBox.information(
                self,
                self.title,
                _result_message(result, "Import xong."),
            ),
        )

    def _filter_changed(self) -> None:
        self.page = 1
        self.reload()


def _primary_button(text: str) -> QPushButton:
    return shared_primary_button(text)


def _secondary_button(text: str) -> QPushButton:
    return shared_secondary_button(text)


def _danger_button(text: str) -> QPushButton:
    return shared_danger_button(text)


def _combo(placeholder: str, first_label: str | None = None) -> QComboBox:
    combo = shared_combo_box(
        first_label or f"Tất cả {placeholder}",
        minimum_width=150,
        maximum_width=220,
        minimum_contents_length=10,
    )
    combo.setToolTip(placeholder)
    return combo


def _populate_combo(combo: QComboBox, values: list[str] | list[tuple[str, str]]) -> None:
    shared_populate_combo(combo, values)


def _fit_column_widths(base_widths: list[int], available: int) -> list[int]:
    if not base_widths:
        return []
    available = max(len(base_widths) * 44, int(available))
    base_total = sum(base_widths) or available
    minimums = [max(46, min(width, 82)) for width in base_widths]
    minimum_total = sum(minimums)
    if available >= base_total:
        extra = available - base_total
        widths = [width + int(extra * width / base_total) for width in base_widths]
    elif available >= minimum_total:
        scale = available / base_total
        widths = [max(minimums[index], int(width * scale)) for index, width in enumerate(base_widths)]
    else:
        scale = available / minimum_total
        widths = [max(44, int(width * scale)) for width in minimums]
    drift = available - sum(widths)
    widths[-1] = max(44, widths[-1] + drift)
    return widths


def _current_filter(combo: QComboBox) -> str:
    return shared_current_data(combo)


def _batch_id(combo: QComboBox) -> int | None:
    value = _current_filter(combo)
    if not value:
        return None
    try:
        return int(value.split(" - ", 1)[0])
    except ValueError:
        return None


def _loan_compare_category_items(values: list[str]) -> list[tuple[str, str]]:
    preferred = [
        "Khach hang tat toan",
        "Khach hang vay giam",
        "Khach hang vay moi",
        "Khach hang vay tang",
        "Khong thay doi",
    ]
    ordered: list[tuple[str, str]] = []
    seen: set[str] = set()
    for value in preferred:
        if value in values:
            ordered.append((_display_loan_compare_category(value), value))
            seen.add(value)
    for value in values:
        if value not in seen:
            ordered.append((_display_loan_compare_category(value), value))
    return ordered


def _display_loan_compare_category(value: object) -> str:
    text = "" if value is None else str(value).strip()
    return LOAN_COMPARE_CATEGORY_LABELS.get(text, text)


def _display_customer_code(value: object) -> str:
    text = "" if value is None else str(value).strip()
    return text[1:] if text.startswith("'") else text


def _loan_compare_row_value(row: dict[str, object], field: str) -> object:
    value = row.get(field)
    if value is None and field in LOAN_COMPARE_FIELD_ALIASES:
        value = row.get(LOAN_COMPARE_FIELD_ALIASES[field])
    return "" if value is None else value


def _database_path(parent: QWidget | None) -> Path:
    settings = getattr(parent, "settings_database", None)
    if isinstance(settings, AppSettingsDatabase):
        return settings.database_path
    return AppSettingsDatabase().database_path


def _short(text: object) -> str:
    value = str(text or "")
    return value[:12] + "..." if len(value) > 15 else value


def _display_officer_name(value: object) -> str:
    text = "" if value is None else str(value).strip()
    if text.startswith("[") and "]" in text:
        return text.split("]", 1)[1].strip()
    return text


def _add_nim_officer_clause(clauses: list[str], params: list[object], value: object) -> None:
    text = "" if value is None else str(value).strip()
    if not text:
        return
    if text.startswith("[") and "]" in text:
        code = text[1:].split("]", 1)[0].strip()
        if code:
            clauses.append("officer_code = ?")
            params.append(code)
            return
    clauses.append("officer_name = ?")
    params.append(_display_officer_name(text))


def _format_integer_vn(value: object) -> str:
    try:
        number = int(round(float(value or 0)))
    except (TypeError, ValueError):
        number = 0
    return f"{number:,}".replace(",", ".")


def _format_money_vn(value: object) -> str:
    return _format_integer_vn(value)


def _format_percent_vn(value: object) -> str:
    try:
        number = Decimal(str(value or 0))
        number = number.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    except (InvalidOperation, TypeError, ValueError):
        number = Decimal("0.00")
    text = f"{number:,.2f}%"
    return text.replace(",", "_").replace(".", ",").replace("_", ".")


def _format_size(size_bytes: int) -> str:
    value = float(size_bytes)
    for unit in ("B", "KB", "MB", "GB"):
        if value < 1024 or unit == "GB":
            if unit == "B":
                return f"{int(value)} B"
            return f"{value:.2f} {unit}"
        value /= 1024
    return f"{value:.2f} GB"


def _nim_delete_title(data_type: SummaryDataType) -> str:
    return "NIM Dư nợ" if data_type == SummaryDataType.NIM_DN else "NIM Nguồn vốn"


def _current_user() -> str:
    try:
        return getpass.getuser()
    except Exception:
        return ""


def _format_percent_axis(value: float) -> str:
    return _format_percent_vn(value)


def _nim_window_title(data_type: SummaryDataType | None) -> str:
    if data_type == SummaryDataType.NIM_DN:
        return NIM_DN_TITLE
    if data_type == SummaryDataType.NIM_NV:
        return NIM_NV_TITLE
    return NIM_TITLE


def _vba_output_path(*parts: str) -> Path:
    return application_root().joinpath(*parts)


def _nim_vba_export_path(data_type: SummaryDataType) -> Path:
    if data_type == SummaryDataType.NIM_DN:
        return _vba_output_path("DuLieu", "BaoCaoNIM_CSDL.xlsx")
    return _vba_output_path("DuLieu", "DP", "BaoCaoNIM_NV_CSDL.xlsx")


def _result_message(result: object, default: str) -> str:
    message = str(getattr(result, "message", default) or default)
    output_path = getattr(result, "output_path", None)
    if output_path:
        message += f"\nĐã xuất: {output_path}"
    return message
