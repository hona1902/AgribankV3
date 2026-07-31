from __future__ import annotations

from decimal import Decimal, InvalidOperation, ROUND_HALF_UP

from PySide6.QtCore import QEvent, QSize, Qt
from PySide6.QtGui import QStandardItem, QStandardItemModel
from PySide6.QtWidgets import QComboBox, QHeaderView, QStyle, QStyleOptionHeader, QTableWidget, QTableWidgetItem

from agribank_v3.ui.components.controls import apply_agribank_combo_popup_style, configure_combo_popup_width, make_compact_control

from .models import OfficerKey


class OfficerMultiSelectCombo(QComboBox):
    def __init__(self, parent=None, *, placeholder: str = "Chọn CBTD", counter_label: str = "CBTD") -> None:
        super().__init__(parent)
        self.setObjectName("AgribankComboBox")
        self._suppress_next_hide = False
        self._placeholder = placeholder
        self._counter_label = counter_label
        self._officers: list[OfficerKey] = []
        self._selected_keys: set[str] = set()
        self._filter_text = ""
        self.setEditable(True)
        if self.lineEdit() is not None:
            self.lineEdit().setReadOnly(True)
            self.lineEdit().setPlaceholderText(self._placeholder)
        self.setModel(QStandardItemModel(self))
        self.view().viewport().installEventFilter(self)
        self.setMinimumWidth(240)
        self.setMaximumWidth(340)
        make_compact_control(self)
        apply_agribank_combo_popup_style(self)
        self._refresh_label()

    def set_officers(self, officers: list[OfficerKey], *, selected_codes: set[str] | None = None, selected_raw: str = "") -> None:
        selected_codes = selected_codes or set()
        previous = set(self._selected_keys)
        self._officers = list(officers)
        available_keys = {_officer_identity(officer) for officer in self._officers}
        self._selected_keys = previous.intersection(available_keys)
        for officer in self._officers:
            if (officer.code and officer.code in selected_codes) or (selected_raw and officer.raw_name == selected_raw):
                self._selected_keys.add(_officer_identity(officer))
        self._rebuild_model()
        self._refresh_label()

    def selected_officers(self) -> list[OfficerKey]:
        return [officer for officer in self._officers if _officer_identity(officer) in self._selected_keys]

    def select_all(self) -> None:
        self._selected_keys = {_officer_identity(officer) for officer in self._officers}
        self._rebuild_model()
        self._refresh_label()

    def select_visible(self) -> None:
        self._selected_keys.update(_officer_identity(officer) for officer in self._visible_officers())
        self._rebuild_model()
        self._refresh_label()

    def set_filter_text(self, text: str) -> None:
        self._filter_text = str(text or "").strip()
        self._rebuild_model()
        self._refresh_label()

    def selected_count(self) -> int:
        return len(self._selected_keys)

    def clear_selection(self) -> None:
        self._selected_keys.clear()
        self._rebuild_model()
        self._refresh_label()

    def force_hide_popup(self) -> None:
        self._suppress_next_hide = False
        super().hidePopup()

    def hidePopup(self) -> None:
        if self._suppress_next_hide:
            self._suppress_next_hide = False
            return
        super().hidePopup()

    def eventFilter(self, watched, event) -> bool:
        if watched == self.view().viewport() and event.type() == QEvent.Type.MouseButtonRelease:
            index = self.view().indexAt(event.pos())
            if index.isValid():
                self._toggle_item(index)
                self._suppress_next_hide = True
                return True
        return super().eventFilter(watched, event)

    def _toggle_item(self, index) -> None:
        item = self.model().itemFromIndex(index)
        if item is None:
            return
        if index.row() == 0:
            visible_keys = {_officer_identity(officer) for officer in self._visible_officers()}
            if visible_keys and visible_keys.issubset(self._selected_keys):
                self._selected_keys.difference_update(visible_keys)
            else:
                self._selected_keys.update(visible_keys)
            self._rebuild_model()
            self._refresh_label()
            return
        value = item.data(Qt.ItemDataRole.UserRole)
        if isinstance(value, OfficerKey):
            key = _officer_identity(value)
            if key in self._selected_keys:
                self._selected_keys.remove(key)
            else:
                self._selected_keys.add(key)
        self._rebuild_model()
        self._refresh_label()

    def _refresh_label(self) -> None:
        selected = self.selected_officers()
        line_edit = self.lineEdit()
        if not selected:
            if line_edit is not None:
                line_edit.setText(self._placeholder)
        elif len(selected) == 1:
            if line_edit is not None:
                line_edit.setText(selected[0].display_name)
        else:
            if line_edit is not None:
                line_edit.setText(f"{len(selected)} {self._counter_label} đã chọn")

    def _rebuild_model(self) -> None:
        model = QStandardItemModel(self)
        visible = self._visible_officers()
        all_item = QStandardItem("Chọn tất cả đang hiển thị")
        all_item.setCheckable(True)
        all_item.setData(None, Qt.ItemDataRole.UserRole)
        visible_keys = {_officer_identity(officer) for officer in visible}
        if visible_keys and visible_keys.issubset(self._selected_keys):
            all_item.setCheckState(Qt.CheckState.Checked)
        model.appendRow(all_item)
        for officer in visible:
            item = QStandardItem(officer.display_name)
            item.setCheckable(True)
            item.setData(officer, Qt.ItemDataRole.UserRole)
            if _officer_identity(officer) in self._selected_keys:
                item.setCheckState(Qt.CheckState.Checked)
            model.appendRow(item)
        self.setModel(model)
        configure_combo_popup_width(self, minimum_popup_width=360)

    def _visible_officers(self) -> list[OfficerKey]:
        if not self._filter_text:
            return list(self._officers)
        needle = self._filter_text.casefold()
        return [
            officer
            for officer in self._officers
            if needle in officer.display_name.casefold()
            or needle in officer.raw_name.casefold()
            or needle in officer.code.casefold()
            or needle in officer.branch.casefold()
            or needle in officer.transaction_office.casefold()
        ]


def _officer_identity(officer: OfficerKey) -> str:
    return officer.code or officer.raw_name or officer.display_name


class FitTableWidget(QTableWidget):
    def __init__(self, parent=None) -> None:
        super().__init__(parent)
        self.setObjectName("SummaryDataTable")
        self.default_widths: tuple[int, ...] = ()
        self.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        header = self.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        header.setStretchLastSection(False)
        if hasattr(header, "refresh_height"):
            header.sectionResized.connect(lambda *_args: header.refresh_height())

    def set_default_widths(self, widths: tuple[int, ...]) -> None:
        self.default_widths = widths
        self.apply_fit_widths()
        header = self.horizontalHeader()
        if hasattr(header, "refresh_height"):
            header.refresh_height()

    def resizeEvent(self, event) -> None:
        super().resizeEvent(event)
        self.apply_fit_widths()

    def apply_fit_widths(self) -> None:
        column_count = self.columnCount()
        if column_count <= 0:
            return
        base_widths = list(self.default_widths[:column_count])
        if len(base_widths) < column_count:
            base_widths.extend([100] * (column_count - len(base_widths)))
        available = max(120, self.viewport().width() - 2)
        base_total = sum(base_widths) or available
        minimums = [max(56, min(width, 90)) for width in base_widths]
        minimum_total = sum(minimums)
        if available >= base_total:
            extra = available - base_total
            widths = list(base_widths)
            for index, width in enumerate(base_widths):
                widths[index] += int(extra * width / base_total)
        elif available >= minimum_total:
            scale = available / base_total
            widths = [max(minimums[index], int(width * scale)) for index, width in enumerate(base_widths)]
        else:
            scale = available / minimum_total
            widths = [max(44, int(width * scale)) for width in minimums]
        drift = available - sum(widths)
        if widths:
            widths[-1] = max(44, widths[-1] + drift)
        for index, width in enumerate(widths):
            self.setColumnWidth(index, width)
        header = self.horizontalHeader()
        if hasattr(header, "refresh_height"):
            header.refresh_height()


class MultiLineHeaderView(QHeaderView):
    def __init__(self, orientation: Qt.Orientation, parent=None) -> None:
        super().__init__(orientation, parent)
        self.setDefaultAlignment(Qt.AlignmentFlag.AlignCenter)
        self.setSectionsClickable(True)
        self.setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        self.setMinimumSectionSize(44)
        self.setFixedHeight(38)

    def paintSection(self, painter, rect, logical_index: int) -> None:
        if not rect.isValid():
            return
        painter.save()
        option = QStyleOptionHeader()
        self.initStyleOption(option)
        option.rect = rect
        option.text = ""
        option.section = logical_index
        self.style().drawControl(QStyle.ControlElement.CE_Header, option, painter, self)
        text = self._header_text(logical_index)
        painter.drawText(
            rect.adjusted(5, 3, -5, -3),
            Qt.AlignmentFlag.AlignCenter | Qt.TextFlag.TextWordWrap,
            text,
        )
        painter.restore()

    def sectionSizeFromContents(self, logical_index: int) -> QSize:
        base = super().sectionSizeFromContents(logical_index)
        width = max(44, self.sectionSize(logical_index) or base.width())
        height = self._text_height(logical_index, width)
        return QSize(base.width(), max(base.height(), height))

    def refresh_height(self) -> None:
        heights = [
            self._text_height(index, max(44, self.sectionSize(index)))
            for index in range(self.count())
            if not self.isSectionHidden(index)
        ]
        self.setFixedHeight(max(32, *(heights or [32])))
        self.viewport().update()

    def _text_height(self, logical_index: int, width: int) -> int:
        text = self._header_text(logical_index)
        if not text:
            return 32
        bounds = self.fontMetrics().boundingRect(
            0,
            0,
            max(24, width - 10),
            1000,
            int(Qt.AlignmentFlag.AlignCenter | Qt.TextFlag.TextWordWrap),
            text,
        )
        return max(32, bounds.height() + 10)

    def _header_text(self, logical_index: int) -> str:
        model = self.model()
        if model is None:
            return ""
        value = model.headerData(logical_index, self.orientation(), Qt.ItemDataRole.DisplayRole)
        return str(value or "")


class NumericTableWidgetItem(QTableWidgetItem):
    def __init__(self, text: str, number: float | None) -> None:
        super().__init__(text)
        self.number = float("-inf") if number is None else float(number)

    def __lt__(self, other: QTableWidgetItem) -> bool:
        if isinstance(other, NumericTableWidgetItem):
            return self.number < other.number
        return super().__lt__(other)


def format_money_vn(value: object, *, signed: bool = False) -> str:
    try:
        number = int(round(float(value or 0)))
    except (TypeError, ValueError):
        number = 0
    prefix = "+" if signed and number > 0 else ""
    return prefix + f"{number:,}".replace(",", ".")


def format_percent_vn(value: object, *, signed: bool = False) -> str:
    if value is None:
        return "N/A"
    try:
        number = Decimal(str(value or 0)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    except (InvalidOperation, TypeError, ValueError):
        number = Decimal("0.00")
    prefix = "+" if signed and number > 0 else ""
    text = f"{number:,.2f}%"
    return prefix + text.replace(",", "_").replace(".", ",").replace("_", ".")
