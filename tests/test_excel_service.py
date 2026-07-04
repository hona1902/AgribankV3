from __future__ import annotations

import unittest
from pathlib import Path
from tempfile import TemporaryDirectory

from agribank_v3.excel.compatibility import ExcelCompatibility
from agribank_v3.excel.service import (
    CaseMode,
    ExcelContext,
    ExcelService,
    transform_text,
)
from agribank_v3.settings import AddinMode


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

    def test_legacy_system_sheets_only_hidden_for_addins(self) -> None:
        addin_context = ExcelContext(
            "16.0", "Excel 2016+", "AgribankV2.xlam", "SYS", "A1", 1
        )
        user_context = ExcelContext(
            "16.0", "Excel 2016+", "BaoCao.xlsx", "DATA", "A1", 1
        )

        self.assertTrue(ExcelService.is_system_worksheet(addin_context))
        self.assertFalse(ExcelService.is_system_worksheet(user_context))

    def test_loads_each_tool_addin_once(self) -> None:
        class Addin:
            def __init__(self, full_name: str) -> None:
                self.FullName = full_name
                self.Name = Path(full_name).name
                self.Installed = False

        class Addins:
            def __init__(self) -> None:
                self.items: list[Addin] = []

            @property
            def Count(self) -> int:
                return len(self.items)

            def Item(self, index: int) -> Addin:
                return self.items[index - 1]

            def Add(self, filename: str, copy_file: bool) -> Addin:
                del copy_file
                addin = Addin(filename)
                self.items.append(addin)
                return addin

        class Workbooks:
            Count = 0

        application = type(
            "AddinApplication",
            (),
            {"Version": "16.0", "AddIns": Addins(), "Workbooks": Workbooks()},
        )()
        service = ExcelService(application)

        with TemporaryDirectory() as temporary:
            directory = Path(temporary)
            (directory / "CustomFunctions.xlam").touch()
            (directory / "Ignore.txt").touch()

            first = service.load_tool_addins(directory)
            second = service.load_tool_addins(directory)

        self.assertEqual(first.loaded, ("CustomFunctions.xlam",))
        self.assertEqual(first.failed, ())
        self.assertTrue(application.AddIns.Item(1).Installed)
        self.assertEqual(second.loaded, ())
        self.assertEqual(second.already_loaded, ("CustomFunctions.xlam",))

    def test_reopens_registered_addin_after_excel_restart(self) -> None:
        class Addin:
            def __init__(self, full_name: str) -> None:
                self.FullName = full_name
                self.Name = Path(full_name).name
                self.Installed = True

        class Addins:
            def __init__(self, addin: Addin) -> None:
                self.items = [addin]

            @property
            def Count(self) -> int:
                return len(self.items)

            def Item(self, index: int) -> Addin:
                return self.items[index - 1]

        class Workbook:
            def __init__(self, full_name: str) -> None:
                self.FullName = full_name
                self.Name = Path(full_name).name
                self.IsAddin = False

        class Workbooks:
            def __init__(self) -> None:
                self.items: list[Workbook] = []

            @property
            def Count(self) -> int:
                return len(self.items)

            def Item(self, index: int) -> Workbook:
                return self.items[index - 1]

            def Open(self, **kwargs: object) -> Workbook:
                workbook = Workbook(str(kwargs["Filename"]))
                self.items.append(workbook)
                return workbook

        with TemporaryDirectory() as temporary:
            directory = Path(temporary)
            path = directory / "RestartFunctions.xlam"
            path.touch()
            workbooks = Workbooks()
            application = type(
                "AddinApplication",
                (),
                {
                    "Version": "16.0",
                    "AddIns": Addins(Addin(str(path))),
                    "Workbooks": workbooks,
                },
            )()
            service = ExcelService(application)

            report = service.load_tool_addins(directory)

        self.assertEqual(report.loaded, ("RestartFunctions.xlam",))
        self.assertEqual(report.already_loaded, ())
        self.assertEqual(report.failed, ())
        self.assertEqual(workbooks.Count, 1)
        self.assertTrue(workbooks.Item(1).IsAddin)

    def test_enables_existing_uninstalled_addin(self) -> None:
        class Addin:
            def __init__(self, full_name: str) -> None:
                self.FullName = full_name
                self.Name = Path(full_name).name
                self.Installed = False

        class Addins:
            def __init__(self, addin: Addin) -> None:
                self.items = [addin]

            @property
            def Count(self) -> int:
                return len(self.items)

            def Item(self, index: int) -> Addin:
                return self.items[index - 1]

            def Add(self, filename: str, copy_file: bool) -> Addin:
                raise AssertionError("Existing add-in should be reused")

        with TemporaryDirectory() as temporary:
            directory = Path(temporary)
            path = directory / "ExistingFunctions.xlam"
            path.touch()
            addin = Addin(str(path))
            application = type(
                "AddinApplication",
                (),
                {"Version": "14.0", "AddIns": Addins(addin)},
            )()
            service = ExcelService(application)

            report = service.load_tool_addins(directory)

        self.assertEqual(report.loaded, ("ExistingFunctions.xlam",))
        self.assertEqual(report.failed, ())
        self.assertTrue(addin.Installed)

    def test_retries_addin_add_with_copy_file_for_legacy_excel(self) -> None:
        class Addin:
            def __init__(self, full_name: str) -> None:
                self.FullName = full_name
                self.Name = Path(full_name).name
                self.Installed = False

        class Addins:
            def __init__(self) -> None:
                self.items: list[Addin] = []
                self.copy_modes: list[bool] = []

            @property
            def Count(self) -> int:
                return len(self.items)

            def Item(self, index: int) -> Addin:
                return self.items[index - 1]

            def Add(self, filename: str, copy_file: bool) -> Addin:
                self.copy_modes.append(copy_file)
                if not copy_file:
                    raise OSError("Excel 2010 rejected external add-in path")
                addin = Addin(filename)
                self.items.append(addin)
                return addin

        addins = Addins()
        application = type(
            "AddinApplication",
            (),
            {"Version": "14.0", "AddIns": addins},
        )()
        service = ExcelService(application)

        with TemporaryDirectory() as temporary:
            directory = Path(temporary)
            (directory / "LegacyFunctions.xlam").touch()

            report = service.load_tool_addins(directory)

        self.assertEqual(addins.copy_modes, [False, True])
        self.assertEqual(report.loaded, ("LegacyFunctions.xlam",))
        self.assertEqual(report.failed, ())
        self.assertTrue(addins.Item(1).Installed)

    def test_opens_addin_workbook_when_addins_collection_fails(self) -> None:
        class Addins:
            Count = 0

            def Add(self, filename: str, copy_file: bool) -> object:
                raise OSError("AddIns collection unavailable")

        class Workbook:
            def __init__(self, full_name: str) -> None:
                self.FullName = full_name
                self.Name = Path(full_name).name
                self.IsAddin = False

        class Workbooks:
            def __init__(self) -> None:
                self.items: list[Workbook] = []

            @property
            def Count(self) -> int:
                return len(self.items)

            def Item(self, index: int) -> Workbook:
                return self.items[index - 1]

            def Open(self, **kwargs: object) -> Workbook:
                workbook = Workbook(str(kwargs["Filename"]))
                self.items.append(workbook)
                return workbook

        workbooks = Workbooks()
        application = type(
            "AddinApplication",
            (),
            {"Version": "14.0", "AddIns": Addins(), "Workbooks": workbooks},
        )()
        service = ExcelService(application)

        with TemporaryDirectory() as temporary:
            directory = Path(temporary)
            (directory / "FallbackFunctions.xlam").touch()

            report = service.load_tool_addins(directory)

        self.assertEqual(report.loaded, ("FallbackFunctions.xlam",))
        self.assertEqual(report.failed, ())
        self.assertEqual(workbooks.Count, 1)
        self.assertTrue(workbooks.Item(1).IsAddin)

    def test_installs_tool_addins_to_xlstart(self) -> None:
        service = ExcelService()
        with TemporaryDirectory() as source_temporary:
            with TemporaryDirectory() as xlstart_temporary:
                source = Path(source_temporary)
                xlstart = Path(xlstart_temporary)
                addin = source / "StartupFunctions.xlam"
                addin.write_bytes(b"addin")
                original_xlstart = ExcelService.excel_xlstart_directory
                ExcelService.excel_xlstart_directory = staticmethod(lambda: xlstart)
                try:
                    first = service.install_tool_addins_to_xlstart(source)
                    second = service.install_tool_addins_to_xlstart(source)
                finally:
                    ExcelService.excel_xlstart_directory = staticmethod(
                        original_xlstart
                    )

        self.assertEqual(first.loaded, ("StartupFunctions.xlam",))
        self.assertEqual(first.failed, ())
        self.assertEqual(second.loaded, ())
        self.assertEqual(second.already_loaded, ("StartupFunctions.xlam",))

    def test_session_mode_opens_addin_without_persistent_registration(self) -> None:
        class Addin:
            def __init__(self, full_name: str) -> None:
                self.FullName = full_name
                self.Name = Path(full_name).name
                self.Installed = True

        class Addins:
            def __init__(self, addin: Addin) -> None:
                self.items = [addin]

            @property
            def Count(self) -> int:
                return len(self.items)

            def Item(self, index: int) -> Addin:
                return self.items[index - 1]

        class Workbook:
            def __init__(self, full_name: str) -> None:
                self.FullName = full_name
                self.Name = Path(full_name).name
                self.IsAddin = False

        class Workbooks:
            def __init__(self) -> None:
                self.items: list[Workbook] = []

            @property
            def Count(self) -> int:
                return len(self.items)

            def Item(self, index: int) -> Workbook:
                return self.items[index - 1]

            def Open(self, **kwargs: object) -> Workbook:
                workbook = Workbook(str(kwargs["Filename"]))
                self.items.append(workbook)
                return workbook

        with TemporaryDirectory() as temporary:
            directory = Path(temporary)
            path = directory / "SessionFunctions.xlam"
            path.touch()
            addin = Addin(str(path))
            workbooks = Workbooks()
            application = type(
                "AddinApplication",
                (),
                {
                    "Version": "14.0",
                    "AddIns": Addins(addin),
                    "Workbooks": workbooks,
                },
            )()
            service = ExcelService(application, addin_mode=AddinMode.SESSION)

            report = service.load_tool_addins(directory)

        self.assertFalse(addin.Installed)
        self.assertEqual(report.loaded, ("SessionFunctions.xlam",))
        self.assertEqual(report.failed, ())
        self.assertEqual(workbooks.Count, 1)
        self.assertTrue(workbooks.Item(1).IsAddin)

    def test_session_cleanup_closes_addin_and_removes_xlstart_copy(self) -> None:
        class Addin:
            def __init__(self, full_name: str) -> None:
                self.FullName = full_name
                self.Name = Path(full_name).name
                self.Installed = True

        class Addins:
            def __init__(self, addin: Addin) -> None:
                self.items = [addin]

            @property
            def Count(self) -> int:
                return len(self.items)

            def Item(self, index: int) -> Addin:
                return self.items[index - 1]

        class Workbook:
            def __init__(self, full_name: str, owner: object) -> None:
                self.FullName = full_name
                self.Name = Path(full_name).name
                self.owner = owner

            def Close(self, SaveChanges: bool = False) -> None:
                del SaveChanges
                self.owner.items.remove(self)

        class Workbooks:
            def __init__(self, full_name: str) -> None:
                self.items = [Workbook(full_name, self)]

            @property
            def Count(self) -> int:
                return len(self.items)

            def Item(self, index: int) -> Workbook:
                return self.items[index - 1]

        with TemporaryDirectory() as source_temporary:
            with TemporaryDirectory() as xlstart_temporary:
                source = Path(source_temporary)
                xlstart = Path(xlstart_temporary)
                path = source / "SessionFunctions.xlam"
                path.write_bytes(b"addin")
                (xlstart / path.name).write_bytes(b"addin")
                addin = Addin(str(path))
                workbooks = Workbooks(str(path))
                application = type(
                    "AddinApplication",
                    (),
                    {
                        "Version": "14.0",
                        "AddIns": Addins(addin),
                        "Workbooks": workbooks,
                    },
                )()
                service = ExcelService(
                    application,
                    addin_mode=AddinMode.SESSION,
                )
                original_xlstart = ExcelService.excel_xlstart_directory
                ExcelService.excel_xlstart_directory = staticmethod(
                    lambda: xlstart
                )
                try:
                    report = service.cleanup_session_addins(source)
                finally:
                    ExcelService.excel_xlstart_directory = staticmethod(
                        original_xlstart
                    )

                self.assertFalse((xlstart / path.name).exists())

        self.assertFalse(addin.Installed)
        self.assertEqual(workbooks.Count, 0)
        self.assertEqual(report.loaded, ("SessionFunctions.xlam",))
        self.assertEqual(report.failed, ())

    def test_disabled_addin_is_not_copied_back_to_xlstart(self) -> None:
        service = ExcelService()
        service.configure_addin_states(
            {
                "DisabledFunctions.xlam": False,
                "EnabledFunctions.xlam": True,
            }
        )
        with TemporaryDirectory() as source_temporary:
            with TemporaryDirectory() as xlstart_temporary:
                source = Path(source_temporary)
                xlstart = Path(xlstart_temporary)
                disabled = source / "DisabledFunctions.xlam"
                enabled = source / "EnabledFunctions.xlam"
                disabled.write_bytes(b"disabled")
                enabled.write_bytes(b"enabled")
                (xlstart / disabled.name).write_bytes(b"old")
                original_xlstart = ExcelService.excel_xlstart_directory
                ExcelService.excel_xlstart_directory = staticmethod(
                    lambda: xlstart
                )
                try:
                    cleanup = service.cleanup_tool_addins(
                        (disabled.name,),
                        source,
                    )
                    install = service.install_tool_addins_to_xlstart(source)
                finally:
                    ExcelService.excel_xlstart_directory = staticmethod(
                        original_xlstart
                    )

                self.assertFalse((xlstart / disabled.name).exists())
                self.assertTrue((xlstart / enabled.name).exists())

        self.assertEqual(cleanup.failed, ())
        self.assertEqual(install.discovered, ("EnabledFunctions.xlam",))


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
        self.FullName = "C:\\Temp\\Book1.xlsx"
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

    def test_application_has_preferred_workbook_path(self) -> None:
        application = _FakeApplication()
        self.assertTrue(
            ExcelService._application_has_workbook(
                application,
                Path("C:/Temp/Book1.xlsx"),
            )
        )
        self.assertFalse(
            ExcelService._application_has_workbook(
                application,
                Path("C:/Temp/Other.xlsx"),
            )
        )

if __name__ == "__main__":
    unittest.main()
