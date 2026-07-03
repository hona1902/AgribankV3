from __future__ import annotations

import unittest

from agribank_v3.excel.compatibility import ExcelCompatibility
from agribank_v3.excel.service import CaseMode, ExcelService, transform_text


class TransformTextTests(unittest.TestCase):
    def test_vietnamese_upper(self) -> None:
        self.assertEqual(transform_text("Nguyễn Văn An", CaseMode.UPPER), "NGUYỄN VĂN AN")

    def test_vietnamese_lower(self) -> None:
        self.assertEqual(transform_text("NGUYỄN VĂN AN", CaseMode.LOWER), "nguyễn văn an")

    def test_vietnamese_title(self) -> None:
        self.assertEqual(transform_text("nGUYỄN vĂN aN", CaseMode.TITLE), "Nguyễn Văn An")

    def test_formula_matrix_preserves_formulas_and_numbers(self) -> None:
        source = (
            ("nguyễn văn an", "=A1", 123),
            ("AGRIBANK", None, ""),
        )
        transformed, stats = ExcelService._transform_formula_matrix(
            source, CaseMode.UPPER
        )
        self.assertEqual(
            transformed,
            (
                ("NGUYỄN VĂN AN", "=A1", 123),
                ("AGRIBANK", None, ""),
            ),
        )
        self.assertEqual(stats, (1, 1, 2))


class CompatibilityTests(unittest.TestCase):
    def test_address_as_property_from_generated_excel_2013_wrapper(self) -> None:
        range_object = type("PropertyRange", (), {"Address": "$A$1:$B$2"})()
        self.assertEqual(
            ExcelCompatibility.range_address(range_object),
            "A1:B2",
        )

    def test_count_falls_back_for_older_excel(self) -> None:
        range_object = type("LegacyRange", (), {"Count": 12})()
        self.assertEqual(ExcelCompatibility.cell_count(range_object), 12)


class _FakeRange:
    def __init__(self, address: str, formulas: object) -> None:
        self._address = address
        self.Formula = formulas

    def Address(self, *_: object) -> str:
        return self._address


class _FakeAreas:
    def __init__(self, areas: list[_FakeRange]) -> None:
        self._areas = areas
        self.Count = len(areas)

    def Item(self, index: int) -> _FakeRange:
        return self._areas[index - 1]


class _FakeSelection(_FakeRange):
    def __init__(self, area: _FakeRange) -> None:
        super().__init__(area._address, area.Formula)
        self.Areas = _FakeAreas([area])
        self.Count = 4
        self.CountLarge = 4


class _FakeWorksheet:
    def __init__(self, area: _FakeRange) -> None:
        self.Name = "Sheet1"
        self._area = area

    def Range(self, address: str) -> _FakeRange:
        if address != self._area._address:
            raise KeyError(address)
        return self._area


class _FakeWorksheets:
    def __init__(self, worksheet: _FakeWorksheet) -> None:
        self._worksheet = worksheet

    def __call__(self, name: str) -> _FakeWorksheet:
        if name != self._worksheet.Name:
            raise KeyError(name)
        return self._worksheet


class _FakeWorkbook:
    def __init__(self, worksheet: _FakeWorksheet) -> None:
        self.Name = "Book1.xlsx"
        self.Worksheets = _FakeWorksheets(worksheet)


class _FakeWorkbooks:
    def __init__(self, workbook: _FakeWorkbook) -> None:
        self._workbook = workbook
        self.Count = 1

    def Item(self, index: int) -> _FakeWorkbook:
        if index != 1:
            raise IndexError(index)
        return self._workbook


class _FakeApplication:
    def __init__(self) -> None:
        area = _FakeRange("A1:B2", (("an", "=A1"), (1, "BÌNH")))
        worksheet = _FakeWorksheet(area)
        workbook = _FakeWorkbook(worksheet)
        self.Version = "16.0"
        self.ActiveWorkbook = workbook
        self.ActiveSheet = worksheet
        self.Selection = _FakeSelection(area)
        self.Workbooks = _FakeWorkbooks(workbook)
        self.ScreenUpdating = True
        self.EnableEvents = True
        self.area = area


class ExcelServiceFlowTests(unittest.TestCase):
    def test_convert_and_undo(self) -> None:
        application = _FakeApplication()
        service = ExcelService(application)

        result = service.convert_selection_case(CaseMode.UPPER)

        self.assertEqual(application.area.Formula, (("AN", "=A1"), (1, "BÌNH")))
        self.assertEqual(result.changed_cells, 1)
        self.assertEqual(result.skipped_formulas, 1)
        self.assertTrue(service.can_undo)

        service.undo_last_change()

        self.assertEqual(application.area.Formula, (("an", "=A1"), (1, "BÌNH")))
        self.assertFalse(service.can_undo)


if __name__ == "__main__":
    unittest.main()
