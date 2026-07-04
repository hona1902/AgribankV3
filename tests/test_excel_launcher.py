from __future__ import annotations

import unittest

from agribank_v3.excel.launcher import _bootstrap_workbook


class ExcelLauncherTests(unittest.TestCase):
    def test_bootstrap_workbook_is_unique_per_launch(self) -> None:
        first = _bootstrap_workbook()
        second = _bootstrap_workbook()

        self.assertNotEqual(first, second)
        self.assertTrue(first.exists())
        self.assertTrue(second.exists())
        self.assertEqual(first.suffix, ".xlsx")


if __name__ == "__main__":
    unittest.main()
