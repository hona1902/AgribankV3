from __future__ import annotations

from PySide6.QtWidgets import QComboBox

from agribank_v3.features.settings.unit_directory.service import UnitDirectoryService
from agribank_v3.ui.components.controls import populate_combo


def populate_branch_combo(
    combo: QComboBox,
    service: UnitDirectoryService,
    *,
    include_all: bool = True,
    default_to_home: bool = False,
) -> None:
    first = combo.itemText(0) if combo.count() else "Tất cả chi nhánh"
    if not include_all:
        combo.clear()
        combo.addItem("Chọn chi nhánh", "")
    values = [
        (service.format_branch_display(item.branch_code), item.branch_code)
        for item in service.get_active_branches()
    ]
    populate_combo(combo, values)
    if not include_all and combo.count() > 1 and combo.currentData() == "":
        combo.setCurrentIndex(1)
    if default_to_home:
        home = service.get_settings().home_branch_code
        index = combo.findData(home)
        if index >= 0:
            combo.setCurrentIndex(index)
    elif include_all and combo.count() and combo.itemText(0) != first:
        combo.setItemText(0, first)


def populate_office_combo(
    combo: QComboBox,
    service: UnitDirectoryService,
    *,
    branch_code: str = "",
    include_all: bool = True,
    default_to_home: bool = False,
) -> None:
    first = combo.itemText(0) if combo.count() else "Tất cả đơn vị"
    values = [
        (service.format_office_display(item.branch_code, item.trctcd), item.office_code)
        for item in service.get_active_offices(branch_code)
    ]
    populate_combo(combo, values)
    if not include_all and combo.count() > 1 and combo.currentData() == "":
        combo.setCurrentIndex(1)
    if default_to_home:
        default_office = service.get_settings().default_office_code
        index = combo.findData(default_office)
        if index >= 0:
            combo.setCurrentIndex(index)
    elif include_all and combo.count() and combo.itemText(0) != first:
        combo.setItemText(0, first)

