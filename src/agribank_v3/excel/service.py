from __future__ import annotations

from dataclasses import dataclass
from enum import StrEnum
import time
from typing import Any

import pythoncom
import win32com.client
from pywintypes import com_error

from agribank_v3.excel.compatibility import ExcelCapabilities, ExcelCompatibility


class ExcelConnectionError(RuntimeError):
    """Raised when AgribankV3 cannot access a usable Excel instance."""

    def __init__(self, message: str, code: str = "general") -> None:
        super().__init__(message)
        self.code = code


class CaseMode(StrEnum):
    UPPER = "upper"
    LOWER = "lower"
    TITLE = "title"


@dataclass(frozen=True, slots=True)
class ExcelContext:
    excel_version: str
    excel_name: str
    workbook: str
    worksheet: str
    selection: str
    cell_count: int


@dataclass(frozen=True, slots=True)
class ConversionResult:
    changed_cells: int
    skipped_formulas: int
    skipped_non_text: int
    context: ExcelContext


@dataclass(slots=True)
class _AreaSnapshot:
    address: str
    formulas: Any


@dataclass(slots=True)
class _UndoSnapshot:
    workbook: str
    worksheet: str
    areas: list[_AreaSnapshot]


def transform_text(value: str, mode: CaseMode) -> str:
    if mode is CaseMode.UPPER:
        return value.upper()
    if mode is CaseMode.LOWER:
        return value.lower()
    if mode is CaseMode.TITLE:
        return value.lower().title()
    raise ValueError(f"Unsupported case mode: {mode}")


class ExcelService:
    MAX_SELECTION_CELLS = 50_000
    EXCEL_EXTENSIONS = (
        ".xls",
        ".xlsx",
        ".xlsm",
        ".xlsb",
        ".xla",
        ".xlam",
        ".xlt",
        ".xltx",
        ".xltm",
    )

    def __init__(self, application: Any | None = None) -> None:
        pythoncom.CoInitialize()
        self._application: Any | None = application
        self._undo_snapshot: _UndoSnapshot | None = None
        self._compatibility: ExcelCompatibility | None = (
            ExcelCompatibility(application) if application is not None else None
        )

    @property
    def capabilities(self) -> ExcelCapabilities | None:
        return self._compatibility.capabilities if self._compatibility else None

    @property
    def application(self) -> Any | None:
        return self._application if self.is_connected else None

    @property
    def is_connected(self) -> bool:
        if self._application is None:
            return False
        try:
            # Version is available across old and new Excel releases. Hwnd is not
            # exposed consistently by older type libraries (notably Excel 2013).
            _ = self._application.Version
            return True
        except (com_error, AttributeError):
            self._application = None
            self._compatibility = None
            return False

    @property
    def can_undo(self) -> bool:
        return self._undo_snapshot is not None

    def disconnect(self) -> None:
        # AgribankV3 never quits Excel. It only releases its COM references.
        self._compatibility = None
        self._application = None

    def connect(
        self,
        *,
        retry_attempts: int = 5,
        create_workbook_if_missing: bool = False,
        required_major_version: int | None = None,
    ) -> ExcelContext:
        if self.is_connected:
            compatibility = self._require_compatibility()
            version_matches = (
                required_major_version is None
                or compatibility.capabilities.major_version
                == required_major_version
            )
            if version_matches:
                try:
                    if (
                        create_workbook_if_missing
                        and int(self._application.Workbooks.Count) == 0
                    ):
                        self._application.Workbooks.Add()
                    return self.get_context()
                except (com_error, AttributeError, ExcelConnectionError):
                    pass

        last_error: Exception | None = None
        context_error: ExcelConnectionError | None = None
        attempts = max(1, retry_attempts)
        for attempt in range(attempts):
            applications, discovery_error = self._discover_applications()
            last_error = discovery_error or last_error
            for application in applications:
                try:
                    self._application = application
                    self._compatibility = ExcelCompatibility(application)
                    if (
                        required_major_version is not None
                        and self._compatibility.capabilities.major_version
                        != required_major_version
                    ):
                        self._application = None
                        self._compatibility = None
                        continue
                    if (
                        create_workbook_if_missing
                        and int(application.Workbooks.Count) == 0
                    ):
                        application.Workbooks.Add()
                    return self.get_context()
                except ExcelConnectionError as exc:
                    context_error = exc
                except (com_error, AttributeError) as exc:
                    last_error = exc
                self._application = None
                self._compatibility = None
            if attempt < attempts - 1:
                time.sleep(0.2)
        if context_error is not None:
            raise context_error
        raise ExcelConnectionError(
            "Không tìm thấy phiên Excel đang mở. Hãy mở Excel và chọn một vùng ô, "
            "sau đó thử kết nối lại.",
            code="not_running",
        ) from last_error

    def _discover_applications(self) -> tuple[list[Any], Exception | None]:
        applications: list[Any] = []
        last_error: Exception | None = None
        try:
            applications.append(
                win32com.client.GetActiveObject("Excel.Application")
            )
        except (com_error, AttributeError) as exc:
            last_error = exc

        # GetActiveObject only returns one Excel instance. Workbook/add-in
        # monikers in the Running Object Table allow us to find other instances,
        # including side-by-side Office installations.
        try:
            running_objects = pythoncom.GetRunningObjectTable()
            bind_context = pythoncom.CreateBindCtx(0)
            monikers = running_objects.EnumRunning()
            while True:
                batch = monikers.Next(1)
                if not batch:
                    break
                moniker = batch[0]
                try:
                    display_name = str(
                        moniker.GetDisplayName(bind_context, None)
                    ).casefold()
                    if not (
                        display_name.startswith("!{000245")
                        or display_name.endswith(self.EXCEL_EXTENSIONS)
                    ):
                        continue
                    running_object = running_objects.GetObject(moniker)
                    dispatch = win32com.client.Dispatch(running_object)
                    application = getattr(dispatch, "Application", dispatch)
                    if "excel" in str(application.Name).casefold():
                        applications.append(application)
                except (com_error, AttributeError, TypeError):
                    continue
        except (com_error, AttributeError) as exc:
            last_error = exc

        unique: list[Any] = []
        fingerprints: set[tuple[str, str]] = set()
        for application in applications:
            try:
                version = str(application.Version)
                try:
                    window_id = str(application.Hwnd)
                except (com_error, AttributeError):
                    window_id = str(id(application._oleobj_))
                fingerprint = (version, window_id)
                if fingerprint not in fingerprints:
                    fingerprints.add(fingerprint)
                    unique.append(application)
            except (com_error, AttributeError) as exc:
                last_error = exc
        return unique, last_error

    def get_context(self) -> ExcelContext:
        application = self._require_application()
        try:
            workbook = application.ActiveWorkbook
            worksheet = application.ActiveSheet
            selection = application.Selection
            if workbook is None or worksheet is None or selection is None:
                raise ExcelConnectionError(
                    "Excel chưa có workbook, worksheet hoặc vùng ô đang hoạt động.",
                    code="no_workbook",
                )
            compatibility = self._require_compatibility()
            address = compatibility.range_address(selection)
            count = compatibility.cell_count(selection)
            return ExcelContext(
                excel_version=compatibility.capabilities.version,
                excel_name=compatibility.capabilities.display_name,
                workbook=str(workbook.Name),
                worksheet=str(worksheet.Name),
                selection=str(address),
                cell_count=count,
            )
        except ExcelConnectionError:
            raise
        except (com_error, AttributeError, TypeError) as exc:
            raise ExcelConnectionError(
                "Đối tượng đang chọn trong Excel không phải là một vùng ô hợp lệ.",
                code="invalid_selection",
            ) from exc

    def convert_selection_case(self, mode: CaseMode) -> ConversionResult:
        application = self._require_application()
        context = self.get_context()
        if context.cell_count > self.MAX_SELECTION_CELLS:
            raise ExcelConnectionError(
                f"Vùng chọn có {context.cell_count:,} ô, vượt giới hạn "
                f"{self.MAX_SELECTION_CELLS:,} ô cho một lần xử lý."
            )

        selection = application.Selection
        snapshots: list[_AreaSnapshot] = []
        changed = 0
        skipped_formulas = 0
        skipped_non_text = 0
        previous_screen_updating = bool(application.ScreenUpdating)
        previous_enable_events = bool(application.EnableEvents)

        try:
            application.ScreenUpdating = False
            application.EnableEvents = False
            areas = selection.Areas
            for index in range(1, int(areas.Count) + 1):
                area = areas.Item(index)
                original = area.Formula
                transformed, stats = self._transform_formula_matrix(original, mode)
                snapshots.append(
                    _AreaSnapshot(
                        address=self._require_compatibility().range_address(area),
                        formulas=original,
                    )
                )
                changed += stats[0]
                skipped_formulas += stats[1]
                skipped_non_text += stats[2]
                if stats[0]:
                    # Formula accepts both formulas and constants. Reassigning through this
                    # property keeps untouched formulas as formulas.
                    area.Formula = transformed

            self._undo_snapshot = _UndoSnapshot(
                workbook=context.workbook,
                worksheet=context.worksheet,
                areas=snapshots,
            )
        except (com_error, AttributeError, TypeError) as exc:
            raise ExcelConnectionError(
                "Excel không thể cập nhật vùng đang chọn. Hãy kiểm tra sheet có bị "
                "bảo vệ hoặc vùng chọn có chứa ô gộp hay không."
            ) from exc
        finally:
            try:
                application.EnableEvents = previous_enable_events
                application.ScreenUpdating = previous_screen_updating
            except com_error:
                pass

        return ConversionResult(
            changed_cells=changed,
            skipped_formulas=skipped_formulas,
            skipped_non_text=skipped_non_text,
            context=self.get_context(),
        )

    def undo_last_change(self) -> int:
        application = self._require_application()
        snapshot = self._undo_snapshot
        if snapshot is None:
            return 0

        try:
            workbook = self._find_open_workbook(snapshot.workbook)
            worksheet = self._require_compatibility().worksheet(
                workbook, snapshot.worksheet
            )
            previous_screen_updating = bool(application.ScreenUpdating)
            previous_enable_events = bool(application.EnableEvents)
            try:
                application.ScreenUpdating = False
                application.EnableEvents = False
                for area in snapshot.areas:
                    worksheet.Range(area.address).Formula = area.formulas
            finally:
                application.EnableEvents = previous_enable_events
                application.ScreenUpdating = previous_screen_updating
        except (com_error, AttributeError) as exc:
            raise ExcelConnectionError(
                "Không thể hoàn tác vì workbook hoặc worksheet ban đầu không còn mở."
            ) from exc

        restored = len(snapshot.areas)
        self._undo_snapshot = None
        return restored

    def _require_application(self) -> Any:
        if not self.is_connected:
            raise ExcelConnectionError(
                "AgribankV3 chưa kết nối Excel. Hãy nhấn “Kết nối Excel” trước."
            )
        return self._application

    def _require_compatibility(self) -> ExcelCompatibility:
        application = self._require_application()
        if self._compatibility is None:
            self._compatibility = ExcelCompatibility(application)
        return self._compatibility

    def _find_open_workbook(self, name: str) -> Any:
        application = self._require_application()
        for index in range(1, int(application.Workbooks.Count) + 1):
            workbook = application.Workbooks.Item(index)
            if str(workbook.Name).casefold() == name.casefold():
                return workbook
        raise ExcelConnectionError(f"Workbook “{name}” không còn mở.")

    @classmethod
    def _transform_formula_matrix(
        cls, value: Any, mode: CaseMode
    ) -> tuple[Any, tuple[int, int, int]]:
        changed = 0
        skipped_formulas = 0
        skipped_non_text = 0

        def visit(item: Any) -> Any:
            nonlocal changed, skipped_formulas, skipped_non_text
            if isinstance(item, tuple):
                return tuple(visit(child) for child in item)
            if isinstance(item, str):
                if item.startswith("="):
                    skipped_formulas += 1
                    return item
                converted = transform_text(item, mode)
                if converted != item:
                    changed += 1
                return converted
            skipped_non_text += 1
            return item

        return visit(value), (changed, skipped_formulas, skipped_non_text)
