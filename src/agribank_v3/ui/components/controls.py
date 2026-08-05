from __future__ import annotations

from PySide6.QtCore import QEvent, QPoint, QRect, QSize, Qt
from PySide6.QtGui import QFontMetrics
from PySide6.QtWidgets import (
    QApplication,
    QComboBox,
    QCompleter,
    QPushButton,
    QSizePolicy,
    QStyle,
    QStyleOptionComboBox,
    QStyledItemDelegate,
    QWidget,
)


OFFICER_NAME_ROLE = Qt.ItemDataRole.UserRole + 1
OFFICER_FULL_TEXT_ROLE = Qt.ItemDataRole.UserRole + 2
AGRIBANK_COMBO_POPUP_STYLE = """
QAbstractItemView {
    background: #ffffff;
    border: 1px solid #d8dce3;
    border-radius: 6px;
    padding: 4px;
    outline: 0;
    selection-background-color: rgba(174, 28, 63, 42);
    selection-color: #202020;
}
QAbstractItemView::item {
    min-height: 24px;
    padding: 2px 8px;
    color: #202020;
}
QAbstractItemView::item:hover {
    background: rgba(174, 28, 63, 22);
}
QAbstractItemView::item:selected {
    background: rgba(174, 28, 63, 42);
    color: #202020;
}
QAbstractItemView::indicator {
    width: 16px;
    height: 16px;
}
"""

BUTTON_SAMPLE_TEXT = "Áp dụng cập nhật Đồng ý Nhập dữ liệu g p q y ụ ạ ặ ậ ệ ị ọ ộ ợ ự"
BUTTON_HORIZONTAL_PADDING = 42
COMBO_TEXT_PADDING = 18
COMBO_CLEAR_BUTTON_WIDTH = 24
COMBO_POPUP_PADDING = 92


def recommended_control_height(
    widget: QWidget,
    *,
    minimum: int = 34,
    vertical_padding: int = 10,
    border_width: int = 2,
) -> int:
    metrics = QFontMetrics(widget.font())
    text_height = max(metrics.lineSpacing(), metrics.boundingRect(BUTTON_SAMPLE_TEXT).height())
    return max(int(minimum), int(text_height + vertical_padding + border_width))


def make_button_control(button: QPushButton) -> QPushButton:
    height = recommended_control_height(button)
    button.setMinimumHeight(height)
    button.setMaximumHeight(16777215)
    if button.text():
        button.setMinimumWidth(max(button.minimumWidth(), recommended_button_width(button)))
    button.setSizePolicy(QSizePolicy.Policy.Preferred, QSizePolicy.Policy.Fixed)
    return button


def recommended_button_width(
    button: QPushButton,
    *,
    horizontal_padding: int = BUTTON_HORIZONTAL_PADDING,
    icon_spacing: int = 6,
) -> int:
    metrics = QFontMetrics(button.font())
    icon_width = 0
    if not button.icon().isNull():
        icon_width = button.iconSize().width() + int(icon_spacing)
    return max(1, metrics.horizontalAdvance(button.text()) + icon_width + int(horizontal_padding))


class AgribankActionButton(QPushButton):
    def __init__(self, text: str = "", parent: QWidget | None = None) -> None:
        self._metrics_enabled = False
        super().__init__(text, parent)
        self._metrics_enabled = True
        make_button_control(self)

    def setText(self, text: str) -> None:
        super().setText(text)
        if getattr(self, "_metrics_enabled", False):
            make_button_control(self)

    def changeEvent(self, event: QEvent) -> None:
        super().changeEvent(event)
        if event.type() in {
            QEvent.Type.FontChange,
            QEvent.Type.ApplicationFontChange,
            QEvent.Type.StyleChange,
        } and getattr(self, "_metrics_enabled", False):
            make_button_control(self)


def make_compact_control(widget: QWidget) -> QWidget:
    widget.setMinimumHeight(30)
    widget.setMaximumHeight(34)
    widget.setSizePolicy(QSizePolicy.Policy.Preferred, QSizePolicy.Policy.Fixed)
    return widget


def primary_button(text: str) -> QPushButton:
    button = AgribankActionButton(text)
    button.setObjectName("PrimaryButton")
    return make_button_control(button)


def secondary_button(text: str) -> QPushButton:
    button = AgribankActionButton(text)
    button.setObjectName("SecondaryButton")
    return make_button_control(button)


def danger_button(text: str) -> QPushButton:
    button = AgribankActionButton(text)
    button.setObjectName("DangerButton")
    return make_button_control(button)


def apply_agribank_combo_popup_style(combo: QComboBox) -> QComboBox:
    combo.view().setObjectName("AgribankComboPopup")
    combo.view().setStyleSheet(AGRIBANK_COMBO_POPUP_STYLE)
    combo.view().setAlternatingRowColors(False)
    combo.view().setUniformItemSizes(True)
    combo.view().setTextElideMode(Qt.TextElideMode.ElideNone)
    combo.view().setItemDelegate(CompactComboItemDelegate(combo.view()))
    return combo


class CompactComboItemDelegate(QStyledItemDelegate):
    def sizeHint(self, option, index) -> QSize:
        size = super().sizeHint(option, index)
        flags = index.flags()
        target_height = 30 if flags & Qt.ItemFlag.ItemIsUserCheckable else 28
        size.setHeight(max(24, min(32, target_height)))
        return size


def combo_box(
    first_label: str = "Tất cả",
    *,
    minimum_width: int = 130,
    maximum_width: int | None = 190,
    minimum_contents_length: int = 10,
    searchable: bool = False,
) -> QComboBox:
    combo = QComboBox()
    combo.setObjectName("AgribankComboBox")
    combo.addItem(first_label, "")
    combo.setMinimumWidth(max(1, int(minimum_width)))
    if maximum_width is None:
        combo.setMaximumWidth(16777215)
        combo.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
    else:
        combo.setMaximumWidth(max(int(minimum_width), int(maximum_width)))
        combo.setSizePolicy(QSizePolicy.Policy.Preferred, QSizePolicy.Policy.Fixed)
    combo.setMinimumContentsLength(max(1, int(minimum_contents_length)))
    combo.setSizeAdjustPolicy(QComboBox.SizeAdjustPolicy.AdjustToMinimumContentsLengthWithIcon)
    combo.setToolTip(first_label)
    if searchable:
        configure_searchable_combo(combo)
    make_compact_control(combo)
    apply_agribank_combo_popup_style(combo)
    if maximum_width is None:
        combo.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
    configure_filter_combo_width(combo, minimum_width=minimum_width, maximum_width=maximum_width)
    return combo


def populate_combo(combo: QComboBox, values: list[str] | tuple[tuple[str, str], ...] | list[tuple[str, str]]) -> None:
    current = combo.currentData()
    first = combo.itemText(0) if combo.count() else "Tất cả"
    combo.blockSignals(True)
    combo.clear()
    combo.addItem(first, "")
    for value in values:
        if isinstance(value, tuple):
            combo.addItem(value[0], value[1])
        else:
            combo.addItem(str(value), str(value))
    _apply_combo_item_tooltips(combo)
    index = combo.findData(current)
    combo.setCurrentIndex(index if index >= 0 else 0)
    configure_filter_combo_width(
        combo,
        minimum_width=int(combo.property("filter_minimum_width") or combo.minimumWidth()),
        maximum_width=_combo_configured_maximum(combo),
        popup_extra_padding=COMBO_POPUP_PADDING,
    )
    combo.blockSignals(False)


def current_data(combo: QComboBox) -> str:
    if combo.isEditable() and combo.currentIndex() >= 0:
        current_text = str(combo.currentText() or "")
        if current_text and current_text != combo.itemText(combo.currentIndex()):
            return ""
    return str(combo.currentData() or "")


def configure_searchable_combo(combo: QComboBox) -> None:
    combo.setEditable(True)
    combo.setInsertPolicy(QComboBox.InsertPolicy.NoInsert)
    combo.setCompleter(QCompleter(combo.model(), combo))
    completer = combo.completer()
    if completer is not None:
        completer.setCaseSensitivity(Qt.CaseSensitivity.CaseInsensitive)
        completer.setFilterMode(Qt.MatchFlag.MatchContains)
        completer.setCompletionMode(QCompleter.CompletionMode.PopupCompletion)
    if combo.lineEdit() is not None:
        combo.lineEdit().setClearButtonEnabled(True)
    combo.currentIndexChanged.connect(lambda index, item=combo: (_sync_combo_tooltip(item, index), _show_combo_text_from_start(item)))


def configure_filter_combo_width(
    combo: QComboBox,
    *,
    minimum_width: int,
    maximum_width: int | None = None,
    popup_extra_padding: int = COMBO_POPUP_PADDING,
    maximum_screen_ratio: float = 0.90,
) -> int:
    combo.setProperty("filter_minimum_width", int(minimum_width))
    combo.setProperty("filter_maximum_width", "" if maximum_width is None else int(maximum_width))
    content_width = _combo_content_width(combo)
    target_width = max(int(minimum_width), content_width + _combo_chrome_width(combo))
    if maximum_width is None:
        combo.setMaximumWidth(16777215)
        combo.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
    else:
        target_width = min(target_width, max(int(minimum_width), int(maximum_width)))
        combo.setMaximumWidth(max(int(minimum_width), int(maximum_width)))
        combo.setSizePolicy(QSizePolicy.Policy.Preferred, QSizePolicy.Policy.Fixed)
    combo.setMinimumWidth(target_width)
    combo.setMinimumContentsLength(max(combo.minimumContentsLength(), max(1, min(28, content_width // max(1, combo.fontMetrics().averageCharWidth())))))
    popup_width = configure_combo_popup_width(
        combo,
        minimum_popup_width=max(target_width, content_width + int(popup_extra_padding)),
        maximum_screen_ratio=maximum_screen_ratio,
    )
    _sync_combo_tooltip(combo, combo.currentIndex())
    _show_combo_text_from_start(combo)
    return popup_width


def configure_combo_popup_width(
    combo: QComboBox,
    *,
    minimum_popup_width: int = 320,
    maximum_screen_ratio: float = 0.70,
) -> int:
    available = available_screen_geometry(combo)
    maximum_width = max(combo.minimumWidth(), int(available.width() * max(0.1, min(0.9, maximum_screen_ratio))))
    metrics = combo.fontMetrics()
    longest = 0
    for index in range(combo.count()):
        longest = max(longest, metrics.horizontalAdvance(combo.itemText(index)))
        tooltip = combo.itemData(index, Qt.ItemDataRole.ToolTipRole)
        if tooltip:
            longest = max(longest, metrics.horizontalAdvance(str(tooltip).splitlines()[0]))
    width = min(max(int(minimum_popup_width), combo.minimumWidth(), longest + COMBO_POPUP_PADDING), maximum_width)
    combo.view().setMinimumWidth(width)
    combo.view().setTextElideMode(Qt.TextElideMode.ElideNone)
    combo.view().setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
    return width


def populate_officer_combo(
    combo: QComboBox,
    officers: list[dict[str, object]] | tuple[dict[str, object], ...],
    *,
    first_label: str | None = None,
    display_code: bool = True,
    minimum_popup_width: int = 360,
) -> None:
    current = combo.currentData()
    first = first_label if first_label is not None else combo.itemText(0) if combo.count() else "Tất cả cán bộ"
    combo.blockSignals(True)
    combo.clear()
    combo.addItem(first, "")
    combo.setItemData(0, first, Qt.ItemDataRole.ToolTipRole)
    combo.setItemData(0, "", OFFICER_NAME_ROLE)
    combo.setItemData(0, first, OFFICER_FULL_TEXT_ROLE)
    for row in officers:
        code = str(row.get("officer_code") or "").strip()
        name = str(row.get("officer_name") or "").strip()
        if not code and not name:
            continue
        display = _officer_display_text(code, name, display_code=display_code)
        full_text = f"{display}\nMã cán bộ: {code or 'N/A'}"
        combo.addItem(display, code)
        index = combo.count() - 1
        combo.setItemData(index, name, OFFICER_NAME_ROLE)
        combo.setItemData(index, full_text, OFFICER_FULL_TEXT_ROLE)
        combo.setItemData(index, full_text, Qt.ItemDataRole.ToolTipRole)
    _refresh_combo_completer(combo)
    index = combo.findData(current)
    combo.setCurrentIndex(index if index >= 0 else 0)
    configure_filter_combo_width(
        combo,
        minimum_width=int(combo.property("filter_minimum_width") or combo.minimumWidth()),
        maximum_width=_combo_configured_maximum(combo),
        popup_extra_padding=max(COMBO_POPUP_PADDING, minimum_popup_width - combo.minimumWidth()),
    )
    _sync_combo_tooltip(combo, combo.currentIndex())
    combo.blockSignals(False)


def _officer_display_text(code: str, name: str, *, display_code: bool) -> str:
    if display_code and code:
        return f"[{code}] {name or code}"
    return name or code


def _apply_combo_item_tooltips(combo: QComboBox) -> None:
    for index in range(combo.count()):
        text = combo.itemText(index)
        combo.setItemData(index, text, Qt.ItemDataRole.ToolTipRole)
        combo.setItemData(index, text, OFFICER_FULL_TEXT_ROLE)
    _sync_combo_tooltip(combo, combo.currentIndex())
    _refresh_combo_completer(combo)


def _sync_combo_tooltip(combo: QComboBox, index: int) -> None:
    if index < 0:
        combo.setToolTip("")
        return
    tooltip = combo.itemData(index, Qt.ItemDataRole.ToolTipRole) or combo.itemText(index)
    combo.setToolTip(str(tooltip or ""))


def _refresh_combo_completer(combo: QComboBox) -> None:
    completer = combo.completer()
    if completer is not None:
        completer.setModel(combo.model())


def _combo_configured_maximum(combo: QComboBox) -> int | None:
    value = combo.property("filter_maximum_width")
    if value in (None, ""):
        return None
    try:
        return int(value)
    except (TypeError, ValueError):
        return combo.maximumWidth()


def _combo_content_width(combo: QComboBox) -> int:
    metrics = QFontMetrics(combo.font())
    values = [combo.currentText()]
    if combo.count():
        values.append(combo.itemText(0))
    if combo.lineEdit() is not None:
        values.append(combo.lineEdit().placeholderText())
    values.extend(combo.itemText(index) for index in range(combo.count()))
    return max((metrics.horizontalAdvance(str(value or "")) for value in values), default=0)


def _combo_chrome_width(combo: QComboBox) -> int:
    option = QStyleOptionComboBox()
    combo.initStyleOption(option)
    arrow_width = combo.style().subControlRect(
        QStyle.ComplexControl.CC_ComboBox,
        option,
        QStyle.SubControl.SC_ComboBoxArrow,
        combo,
    ).width()
    line_edit = combo.lineEdit() if combo.isEditable() else None
    clear_enabled = bool(getattr(line_edit, "isClearButtonEnabled", lambda: True)()) if line_edit is not None else False
    clear_width = COMBO_CLEAR_BUTTON_WIDTH if clear_enabled else 0
    scrollbar_width = combo.style().pixelMetric(QStyle.PixelMetric.PM_ScrollBarExtent, None, combo)
    return max(54, arrow_width + clear_width + scrollbar_width + COMBO_TEXT_PADDING)


def _show_combo_text_from_start(combo: QComboBox) -> None:
    line_edit = combo.lineEdit() if combo.isEditable() else None
    if line_edit is None:
        return
    line_edit.setCursorPosition(0)
    line_edit.deselect()


def available_screen_geometry(window: QWidget) -> QRect:
    screen = window.screen() if window is not None else None
    if screen is None:
        screen = QApplication.primaryScreen()
    if screen is None:
        return QRect(0, 0, 1280, 720)
    return screen.availableGeometry()


def centered_geometry(available: QRect, width: int, height: int) -> QRect:
    width = min(max(1, width), max(1, available.width()))
    height = min(max(1, height), max(1, available.height()))
    left = available.left() + max(0, (available.width() - width) // 2)
    top = available.top() + max(0, (available.height() - height) // 2)
    return QRect(QPoint(left, top), QSize(width, height))


def ensure_geometry_visible(candidate: QRect, available: QRect) -> QRect:
    width = min(max(1, candidate.width()), max(1, available.width()))
    height = min(max(1, candidate.height()), max(1, available.height()))
    if width >= available.width() and height >= available.height():
        return centered_geometry(available, width, height)
    left = min(max(candidate.left(), available.left()), available.right() - width + 1)
    top = min(max(candidate.top(), available.top()), available.bottom() - height + 1)
    return QRect(QPoint(left, top), QSize(width, height))


def fit_window_to_screen(window: QWidget, *, width_ratio: float = 0.94, height_ratio: float = 0.92) -> QRect:
    available = available_screen_geometry(window)
    width = int(available.width() * max(0.3, min(1.0, width_ratio)))
    height = int(available.height() * max(0.3, min(1.0, height_ratio)))
    geometry = centered_geometry(available, width, height)
    window.setGeometry(geometry)
    return geometry


def center_window_on_screen(window: QWidget) -> None:
    available = available_screen_geometry(window)
    window.move(centered_geometry(available, window.width(), window.height()).topLeft())
