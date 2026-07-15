from __future__ import annotations

import unittest
from unittest.mock import patch

from agribank_v3 import printer_manager
from agribank_v3.printer_manager import (
    get_installed_printers,
    powershell_single_quote,
)


class PrinterManagerValidationTests(unittest.TestCase):
    def test_escapes_powershell_single_quoted_arguments(self) -> None:
        self.assertEqual(powershell_single_quote("HP O'Brien"), "'HP O''Brien'")

    def test_empty_printer_list_from_powershell_is_supported(self) -> None:
        completed = printer_manager.subprocess.CompletedProcess(
            args=["powershell.exe"],
            returncode=0,
            stdout="",
            stderr="",
        )
        with (
            patch.object(printer_manager, "is_windows", return_value=True),
            patch.object(printer_manager, "win32print", None),
            patch.object(printer_manager, "get_default_printer", return_value=None),
            patch.object(printer_manager, "_run_powershell", return_value=completed),
        ):
            self.assertEqual(get_installed_printers(), ())


if __name__ == "__main__":
    unittest.main()
