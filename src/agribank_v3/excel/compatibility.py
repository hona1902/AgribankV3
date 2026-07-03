from __future__ import annotations

from dataclasses import dataclass
from typing import Any

from pywintypes import com_error
import win32gui


@dataclass(frozen=True, slots=True)
class ExcelCapabilities:
    version: str
    major_version: int
    supports_ribbon14: bool
    supports_count_large: bool
    supports_formula2: bool
    supports_dynamic_arrays: bool

    @property
    def display_name(self) -> str:
        names = {
            12: "Excel 2007",
            14: "Excel 2010",
            15: "Excel 2013",
            16: "Excel 2016/2019/2021/2024/Microsoft 365",
        }
        return names.get(self.major_version, f"Excel {self.version}")


class ExcelCompatibility:
    """Normalizes COM API differences between Excel type libraries."""

    def __init__(self, application: Any) -> None:
        self.application = application
        version = str(application.Version)
        try:
            major_version = int(float(version))
        except ValueError:
            major_version = 0
        self.capabilities = ExcelCapabilities(
            version=version,
            major_version=major_version,
            supports_ribbon14=major_version >= 14,
            supports_count_large=major_version >= 14,
            supports_formula2=major_version >= 16,
            supports_dynamic_arrays=major_version >= 16,
        )

    @staticmethod
    def range_address(range_object: Any) -> str:
        # Dynamic dispatch exposes Address as a method. Generated wrappers for
        # older Office type libraries may expose it directly as a string.
        address_member = range_object.Address
        if callable(address_member):
            try:
                value = address_member(False, False)
            except TypeError:
                value = address_member()
        else:
            value = address_member
        return str(value).replace("$", "")

    @staticmethod
    def cell_count(range_object: Any) -> int:
        try:
            return int(range_object.CountLarge)
        except (AttributeError, com_error, TypeError, ValueError):
            return int(range_object.Count)

    @staticmethod
    def worksheet(workbook: Any, name: str) -> Any:
        worksheets = workbook.Worksheets
        try:
            return worksheets.Item(name)
        except (AttributeError, TypeError):
            return worksheets(name)

    @staticmethod
    def window_handle(application: Any, workbook_name: str = "") -> int:
        getters = (
            lambda: application.Hwnd,
            lambda: application.ActiveWindow.Hwnd,
            lambda: application.Windows.Item(1).Hwnd,
        )
        for getter in getters:
            try:
                hwnd = int(getter())
                if hwnd and win32gui.IsWindow(hwnd):
                    return hwnd
            except (AttributeError, com_error, TypeError, ValueError):
                continue

        active_caption = ""
        try:
            active_caption = str(application.ActiveWindow.Caption).casefold()
        except (AttributeError, com_error):
            pass
        if not active_caption and workbook_name:
            active_caption = workbook_name.casefold()
        candidates: list[int] = []

        def collect(hwnd: int, _: object) -> None:
            try:
                if win32gui.GetClassName(hwnd) != "XLMAIN":
                    return
                candidates.append(hwnd)
            except win32gui.error:
                return

        win32gui.EnumWindows(collect, None)
        if active_caption:
            matching = [
                hwnd
                for hwnd in candidates
                if active_caption in win32gui.GetWindowText(hwnd).casefold()
            ]
            if len(matching) == 1:
                return matching[0]
        if len(candidates) == 1:
            return candidates[0]
        raise AttributeError("Excel không cung cấp HWND có thể xác định duy nhất.")
