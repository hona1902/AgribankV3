"""Excel integration services."""

from agribank_v3.excel.compatibility import ExcelCapabilities, ExcelCompatibility
from agribank_v3.excel.launcher import (
    ExcelInstallation,
    ExcelLaunchHandle,
    discover_excel_installations,
    launch_excel,
)
from agribank_v3.excel.service import (
    CaseMode,
    ConversionResult,
    ExcelConnectionError,
    ExcelContext,
    ExcelService,
)

__all__ = [
    "CaseMode",
    "ConversionResult",
    "ExcelCapabilities",
    "ExcelCompatibility",
    "ExcelInstallation",
    "ExcelLaunchHandle",
    "ExcelConnectionError",
    "ExcelContext",
    "ExcelService",
    "discover_excel_installations",
    "launch_excel",
]
