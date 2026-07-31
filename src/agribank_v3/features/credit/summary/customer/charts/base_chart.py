from __future__ import annotations

from datetime import datetime
from pathlib import Path

from PySide6.QtCharts import QChart, QChartView
from PySide6.QtCore import QEvent, QPoint, Qt
from PySide6.QtGui import QColor, QPainter
from PySide6.QtWidgets import QFileDialog, QHBoxLayout, QLabel, QPushButton, QVBoxLayout, QWidget

from agribank_v3.features.credit.summary.customer.charts.chart_formatters import format_chart_value
from agribank_v3.features.credit.summary.customer.charts.chart_tooltip import ChartTooltip


CHART_COLORS = (
    QColor("#1f6feb"),
    QColor("#d97706"),
    QColor("#059669"),
    QColor("#7c3aed"),
    QColor("#dc2626"),
    QColor("#2563eb"),
)


class ChartLoadingState(QLabel):
    def __init__(self, parent=None) -> None:
        super().__init__("Đang tải dữ liệu...", parent)
        self.setAlignment(Qt.AlignmentFlag.AlignCenter)


class ChartEmptyState(QLabel):
    def __init__(self, parent=None) -> None:
        super().__init__("Không có dữ liệu phù hợp với bộ lọc.", parent)
        self.setAlignment(Qt.AlignmentFlag.AlignCenter)


class BaseCustomerChart(QWidget):
    def __init__(self, title: str, *, value_kind: str = "money", parent=None) -> None:
        super().__init__(parent)
        self.title = title
        self.value_kind = value_kind
        self.state = "empty"
        self.save_name = ""
        self.last_tooltip_text = ""
        self.last_series: tuple[tuple[str, tuple[tuple[str, float], ...]], ...] = ()
        self.chart = QChart()
        self.chart.setTitle(title)
        self.chart.legend().setVisible(True)
        self.chart.legend().setAlignment(Qt.AlignmentFlag.AlignBottom)
        self.chart.setAnimationOptions(QChart.AnimationOption.SeriesAnimations)
        self.chart_view = QChartView(self.chart)
        self.chart_view.setRenderHint(QPainter.RenderHint.Antialiasing)
        self.chart_view.setMinimumHeight(260)
        self.chart_view.setMouseTracking(True)
        self.chart_view.installEventFilter(self)
        self.state_label = QLabel("Không có dữ liệu phù hợp với bộ lọc.")
        self.state_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.tooltip_label = QLabel(self)
        self.tooltip_label.setObjectName("CustomerChartTooltip")
        self.tooltip_label.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
        self.tooltip_label.setWordWrap(True)
        self.tooltip_label.setStyleSheet(
            """
            QLabel#CustomerChartTooltip {
                background-color: #ffffff;
                border: 1px solid #b8c2cc;
                border-radius: 4px;
                color: #202020;
                padding: 6px 8px;
            }
            """
        )
        self.tooltip_label.hide()
        self.save_button = QPushButton("Lưu biểu đồ")
        self.save_button.setObjectName("SecondaryButton")
        self.save_button.clicked.connect(self.save_chart_dialog)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        self.toolbar = QHBoxLayout()
        self.title_label = QLabel(title)
        self.title_label.setObjectName("SectionTitle")
        self.toolbar.addWidget(self.title_label)
        self.toolbar.addStretch()
        self.toolbar.addWidget(self.save_button)
        layout.addLayout(self.toolbar)
        layout.addWidget(self.chart_view)
        layout.addWidget(self.state_label)
        self.setMinimumHeight(320)
        self.set_empty()

    def set_loading(self) -> None:
        self.clear_chart()
        self.save_button.setEnabled(False)
        self.last_series = ()
        self.state = "loading"
        self.state_label.setText("Đang tải dữ liệu...")
        self.state_label.show()
        self.chart_view.hide()

    def set_empty(self, message: str = "Không có dữ liệu phù hợp với bộ lọc.") -> None:
        self.clear_chart()
        self.save_button.setEnabled(False)
        self.last_series = ()
        self.state = "empty"
        self.state_label.setText(message)
        self.state_label.show()
        self.chart_view.hide()

    def set_error(self, message: str = "Không tải được dữ liệu biểu đồ.") -> None:
        self.clear_chart()
        self.save_button.setEnabled(False)
        self.last_series = ()
        self.state = "error"
        self.state_label.setText(message)
        self.state_label.show()
        self.chart_view.hide()

    def show_chart(self) -> None:
        self.state = "ready"
        self.save_button.setEnabled(True)
        self.state_label.hide()
        self.chart_view.show()

    def clear_chart(self) -> None:
        self.hide_tooltip()
        self.chart.removeAllSeries()
        for axis in list(self.chart.axes()):
            self.chart.removeAxis(axis)
        self.chart.setTitle(self.title)

    def set_title(self, title: str) -> None:
        self.title = str(title or "").strip() or self.title
        self.title_label.setText(self.title)
        self.chart.setTitle(self.title)

    def set_save_name(self, name: str) -> None:
        self.save_name = str(name or "").strip()

    def add_header_widget(self, widget: QWidget) -> None:
        self.toolbar.insertWidget(self.toolbar.indexOf(self.save_button), widget)

    def show_tooltip(self, tooltip: ChartTooltip, global_position) -> None:
        self.last_tooltip_text = tooltip.text()
        self.tooltip_label.setText(self.last_tooltip_text)
        self.tooltip_label.adjustSize()
        local_position = self.mapFromGlobal(global_position) + QPoint(14, 18)
        local_position.setX(min(max(0, local_position.x()), max(0, self.width() - self.tooltip_label.width() - 4)))
        local_position.setY(min(max(0, local_position.y()), max(0, self.height() - self.tooltip_label.height() - 4)))
        self.tooltip_label.move(local_position)
        self.tooltip_label.raise_()
        self.tooltip_label.show()

    def hide_tooltip(self) -> None:
        self.tooltip_label.hide()

    def eventFilter(self, watched, event) -> bool:
        if watched is self.chart_view and event.type() in {
            QEvent.Type.Leave,
            QEvent.Type.Hide,
            QEvent.Type.Close,
        }:
            self.hide_tooltip()
        return super().eventFilter(watched, event)

    def hideEvent(self, event) -> None:
        self.hide_tooltip()
        super().hideEvent(event)

    def closeEvent(self, event) -> None:
        self.hide_tooltip()
        super().closeEvent(event)

    def save_chart_dialog(self) -> None:
        base_name = _safe_name(self.save_name or f"CustomerChart_{self.title}")
        path, _ = QFileDialog.getSaveFileName(
            self,
            "Lưu biểu đồ",
            f"{base_name}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png",
            "PNG (*.png)",
        )
        if path:
            self.save_png(Path(path))

    def save_png(self, destination: Path) -> Path:
        destination = Path(destination)
        destination.parent.mkdir(parents=True, exist_ok=True)
        self.grab().save(str(destination), "PNG")
        return destination

    def _format_axis_label(self, value: float, *, divisor: float = 1.0, unit: str = "đồng") -> str:
        if self.value_kind.startswith("percent"):
            return format_chart_value(value, "percent", full=False)
        if self.value_kind.startswith("number"):
            return format_chart_value(value, "number", full=False)
        return format_chart_value(value * divisor, "money", full=False, divisor=divisor, unit=unit)


def _safe_name(value: str) -> str:
    return "".join(char for char in value if char.isalnum() or char in {"_", "-"}).strip() or "Chart"
