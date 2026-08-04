from __future__ import annotations

from PySide6.QtCore import QAbstractTableModel, QModelIndex, Qt

from agribank_v3.features.credit.summary.customer.formatters import (
    format_customer_type,
    format_money_vn,
    format_override_status,
    format_percent_vn,
)


ColumnSpec = tuple[str, str, str]
TERM_STRUCTURE_WARNING = (
    "Kỳ này chưa có dữ liệu phân bổ kỳ hạn đầy đủ theo CBTD. "
    "Vui lòng nhập lại kỳ NIM Dư nợ từ file FTP Loan."
)
TERM_STRUCTURE_FIELDS = {
    "short_term_balance",
    "medium_long_term_balance",
    "other_balance",
    "medium_long_ratio",
}


class CustomerTableModel(QAbstractTableModel):
    def __init__(self, columns: tuple[ColumnSpec, ...], parent=None) -> None:
        super().__init__(parent)
        self.columns = columns
        self.rows: list[dict[str, object]] = []

    def set_rows(self, rows: list[dict[str, object]]) -> None:
        self.beginResetModel()
        self.rows = list(rows)
        self.endResetModel()

    def rowCount(self, parent: QModelIndex = QModelIndex()) -> int:
        return 0 if parent.isValid() else len(self.rows)

    def columnCount(self, parent: QModelIndex = QModelIndex()) -> int:
        return 0 if parent.isValid() else len(self.columns)

    def data(self, index: QModelIndex, role: int = Qt.ItemDataRole.DisplayRole):
        if not index.isValid():
            return None
        row = self.rows[index.row()]
        field, _label, kind = self.columns[index.column()]
        value = row.get(field)
        if role == Qt.ItemDataRole.UserRole:
            return row
        if role == Qt.ItemDataRole.DisplayRole:
            return _display_value(value, kind, row=row)
        if role == Qt.ItemDataRole.ToolTipRole:
            if field in TERM_STRUCTURE_FIELDS and row.get("term_structure_available") is False:
                return str(row.get("term_structure_warning") or TERM_STRUCTURE_WARNING)
            text = _display_value(value, kind, row=row)
            return text if text else None
        if role == Qt.ItemDataRole.TextAlignmentRole:
            return _alignment(kind)
        return None

    def headerData(
        self,
        section: int,
        orientation: Qt.Orientation,
        role: int = Qt.ItemDataRole.DisplayRole,
    ):
        if role != Qt.ItemDataRole.DisplayRole or orientation != Qt.Orientation.Horizontal:
            return None
        if 0 <= section < len(self.columns):
            return self.columns[section][1]
        return None

    def raw_row(self, row_index: int) -> dict[str, object]:
        if 0 <= row_index < len(self.rows):
            return dict(self.rows[row_index])
        return {}

    def sort(self, column: int, order: Qt.SortOrder = Qt.SortOrder.AscendingOrder) -> None:
        if not (0 <= column < len(self.columns)):
            return
        field, _label, kind = self.columns[column]
        reverse = order == Qt.SortOrder.DescendingOrder
        self.layoutAboutToBeChanged.emit()
        self.rows.sort(key=lambda row: _sort_key(row.get(field), kind), reverse=reverse)
        self.layoutChanged.emit()


def _display_value(value: object, kind: str, *, row: dict[str, object] | None = None) -> str:
    if kind in {"money", "money_signed"}:
        return format_money_vn(value, signed=kind.endswith("signed"))
    if kind == "money_or_blank":
        return "" if value in (None, "") else format_money_vn(value)
    if kind == "term_money_or_dash":
        return "—" if value in (None, "") else format_money_vn(value)
    if kind in {"percent", "percent_signed"}:
        return format_percent_vn(value, signed=kind.endswith("signed"))
    if kind == "percent_or_blank":
        return "" if value in (None, "") else format_percent_vn(value)
    if kind == "term_percent_or_dash":
        if value in (None, ""):
            return "—" if row and row.get("term_structure_available") is False else "N/A"
        return format_percent_vn(value)
    if kind == "percent_point_signed":
        if value in (None, ""):
            return "N/A"
        return format_percent_vn(value, signed=True).replace("%", " đ.%")
    if kind == "customer_type":
        return format_customer_type(value)
    if kind == "yes_no":
        return "Có" if int(value or 0) == 1 else "Không"
    if kind == "active_status":
        return "Đang sử dụng" if int(value or 0) == 1 else "Ngừng sử dụng"
    if kind == "override_status_bool":
        return format_override_status(value)
    if kind == "integer":
        try:
            return f"{int(value or 0):,}".replace(",", ".")
        except (TypeError, ValueError):
            return "0"
    return "" if value is None else str(value)


def _alignment(kind: str) -> Qt.AlignmentFlag:
    if kind in {
        "money",
        "money_signed",
        "money_or_blank",
        "term_money_or_dash",
        "percent",
        "percent_signed",
        "percent_or_blank",
        "term_percent_or_dash",
        "percent_point_signed",
        "integer",
    }:
        return Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter
    if kind == "center":
        return Qt.AlignmentFlag.AlignCenter
    return Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter


def _sort_key(value: object, kind: str):
    if kind in {
        "money",
        "money_signed",
        "money_or_blank",
        "term_money_or_dash",
        "percent",
        "percent_signed",
        "percent_or_blank",
        "term_percent_or_dash",
        "percent_point_signed",
        "integer",
    }:
        try:
            return (0, float(value or 0))
        except (TypeError, ValueError):
            return (1, 0.0)
    return (0, "" if value is None else str(value).casefold())
