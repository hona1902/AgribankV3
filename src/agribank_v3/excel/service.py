from __future__ import annotations

from dataclasses import dataclass
from enum import StrEnum
from pathlib import Path
import os
import shutil
import time
from typing import Any

import pythoncom
import win32com.client
from pywintypes import com_error

from agribank_v3.excel.compatibility import ExcelCapabilities, ExcelCompatibility
from agribank_v3.runtime_paths import application_root
from agribank_v3.settings import AddinMode


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


@dataclass(frozen=True, slots=True)
class AddinLoadReport:
    directory: Path
    discovered: tuple[str, ...] = ()
    loaded: tuple[str, ...] = ()
    already_loaded: tuple[str, ...] = ()
    failed: tuple[tuple[str, str], ...] = ()


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
    LEGACY_SYSTEM_SHEETS = frozenset(
        {"SYS", "UNICODE", "NHANVIEN", "DATA", "TOADO", "QUYETTOAN"}
    )
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

    def __init__(
        self,
        application: Any | None = None,
        addin_mode: AddinMode = AddinMode.PERMANENT,
    ) -> None:
        pythoncom.CoInitialize()
        self._application: Any | None = application
        self._addin_mode = AddinMode(addin_mode)
        self._disabled_addins: set[str] = set()
        self._undo_snapshot: _UndoSnapshot | None = None
        self._compatibility: ExcelCompatibility | None = (
            ExcelCompatibility(application) if application is not None else None
        )
        self._last_addin_report = AddinLoadReport(
            directory=self.tool_addin_directory()
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

    @property
    def last_addin_report(self) -> AddinLoadReport:
        return self._last_addin_report

    @property
    def addin_mode(self) -> AddinMode:
        return self._addin_mode

    def set_addin_mode(self, mode: AddinMode | str) -> None:
        self._addin_mode = AddinMode(mode)

    def configure_addin_states(self, states: dict[str, bool]) -> None:
        self._disabled_addins = {
            name.casefold() for name, enabled in states.items() if not enabled
        }

    def set_addin_enabled(self, file_name: str, enabled: bool) -> None:
        normalized = Path(file_name).name.casefold()
        if enabled:
            self._disabled_addins.discard(normalized)
        else:
            self._disabled_addins.add(normalized)

    @staticmethod
    def tool_addin_directory() -> Path:
        return application_root() / "tools" / "addins"

    @staticmethod
    def excel_xlstart_directory() -> Path:
        appdata = os.environ.get("APPDATA")
        if appdata:
            return Path(appdata) / "Microsoft" / "Excel" / "XLSTART"
        return Path.home() / "AppData" / "Roaming" / "Microsoft" / "Excel" / "XLSTART"

    def disconnect(self) -> None:
        # AgribankV3 never quits Excel. It only releases its COM references.
        self._compatibility = None
        self._application = None
        self._last_addin_report = AddinLoadReport(
            directory=self.tool_addin_directory()
        )

    def connect(
        self,
        *,
        retry_attempts: int = 5,
        create_workbook_if_missing: bool = False,
        required_major_version: int | None = None,
        preferred_workbook_path: Path | None = None,
    ) -> ExcelContext:
        if self.is_connected:
            compatibility = self._require_compatibility()
            version_matches = (
                required_major_version is None
                or compatibility.capabilities.major_version
                == required_major_version
            )
            workbook_matches = (
                preferred_workbook_path is None
                or self._application_has_workbook(
                    self._application, preferred_workbook_path
                )
            )
            if version_matches and workbook_matches:
                try:
                    if (
                        create_workbook_if_missing
                        and int(self._application.Workbooks.Count) == 0
                    ):
                        self._application.Workbooks.Add()
                    context = self.get_context()
                    self._last_addin_report = self.load_tool_addins()
                    return self._restore_context(context)
                except (com_error, AttributeError, ExcelConnectionError):
                    pass

        last_error: Exception | None = None
        context_error: ExcelConnectionError | None = None
        attempts = max(1, retry_attempts)
        for attempt in range(attempts):
            applications, discovery_error = self._discover_applications(
                preferred_workbook_path=preferred_workbook_path
            )
            last_error = discovery_error or last_error
            applications = self._rank_applications(
                applications,
                required_major_version=required_major_version,
                preferred_workbook_path=preferred_workbook_path,
            )
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
                        preferred_workbook_path is not None
                        and not self._application_has_workbook(
                            application, preferred_workbook_path
                        )
                    ):
                        self._application = None
                        self._compatibility = None
                        continue
                    if (
                        create_workbook_if_missing
                        and int(application.Workbooks.Count) == 0
                    ):
                        application.Workbooks.Add()
                    context = self.get_context()
                    self._last_addin_report = self.load_tool_addins()
                    return self._restore_context(context)
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

    def load_tool_addins(
        self, directory: Path | None = None
    ) -> AddinLoadReport:
        """Install tool add-ins at Excel application scope for all workbooks."""
        application = self._require_application()
        addin_directory = Path(directory or self.tool_addin_directory())
        try:
            addin_directory.mkdir(parents=True, exist_ok=True)
            all_addin_paths = tuple(
                sorted(
                    (
                        path.resolve()
                        for path in addin_directory.iterdir()
                        if path.is_file()
                        and not path.name.startswith("~$")
                        and path.suffix.casefold() in {".xla", ".xlam"}
                    ),
                    key=lambda path: path.name.casefold(),
                )
            )
        except OSError as exc:
            report = AddinLoadReport(
                directory=addin_directory,
                failed=(("Thư mục add-in", self._short_com_error(exc)),),
            )
            self._last_addin_report = report
            return report

        disabled_paths = tuple(
            path
            for path in all_addin_paths
            if path.name.casefold() in self._disabled_addins
        )
        disabled_cleanup = self.cleanup_tool_addins(
            tuple(path.name for path in disabled_paths),
            addin_directory,
        )
        addin_paths = tuple(
            path
            for path in all_addin_paths
            if path.name.casefold() not in self._disabled_addins
        )
        loaded: list[str] = []
        already_loaded: list[str] = []
        failed: list[tuple[str, str]] = list(disabled_cleanup.failed)

        if self._addin_mode is AddinMode.SESSION:
            report = self._load_session_addins(
                application,
                addin_directory,
                addin_paths,
            )
            self._last_addin_report = report
            return report

        for path in addin_paths:
            existing_addin = self._find_matching_addin(application, path)
            if existing_addin is not None:
                was_installed = self._is_addin_installed(existing_addin)
                try:
                    existing_addin.Installed = True
                except (com_error, AttributeError, TypeError, OSError) as exc:
                    failed.append((path.name, self._short_com_error(exc)))
                    continue
                if self._is_addin_installed(existing_addin):
                    if self._ensure_addin_workbook_open(application, path):
                        loaded.append(path.name)
                    elif was_installed:
                        already_loaded.append(path.name)
                    else:
                        loaded.append(path.name)
                else:
                    failed.append(
                        (path.name, "Excel không bật được add-in đã có sẵn.")
                    )
                continue

            try:
                addin = self._add_and_install_addin(application, path)
                if not self._is_addin_installed(addin):
                    raise ExcelConnectionError("Excel không xác nhận Installed=True")
                self._ensure_addin_workbook_open(application, path)
                loaded.append(path.name)
            except (
                ExcelConnectionError,
                com_error,
                AttributeError,
                TypeError,
                OSError,
            ) as exc:
                addin_error = self._short_com_error(exc)
                try:
                    workbook = self._open_addin_workbook(application, path)
                    if workbook is not None:
                        loaded.append(path.name)
                    else:
                        failed.append(
                            (
                                path.name,
                                f"AddIns thất bại: {addin_error}; "
                                "Workbooks.Open không trả workbook.",
                            )
                        )
                except (
                    ExcelConnectionError,
                    com_error,
                    AttributeError,
                    TypeError,
                    OSError,
                ) as open_exc:
                    failed.append(
                        (
                            path.name,
                            f"AddIns thất bại: {addin_error}; "
                            f"Workbooks.Open thất bại: {self._short_com_error(open_exc)}",
                        )
                    )

        report = AddinLoadReport(
            directory=addin_directory.resolve(),
            discovered=tuple(path.name for path in addin_paths),
            loaded=tuple(loaded),
            already_loaded=tuple(already_loaded),
            failed=tuple(failed),
        )
        self._last_addin_report = report
        return report

    def _load_session_addins(
        self,
        application: Any,
        directory: Path,
        addin_paths: tuple[Path, ...],
    ) -> AddinLoadReport:
        loaded: list[str] = []
        already_loaded: list[str] = []
        failed: list[tuple[str, str]] = []
        for path in addin_paths:
            try:
                existing_addin = self._find_matching_addin(application, path)
                if existing_addin is not None and self._is_addin_installed(
                    existing_addin
                ):
                    existing_addin.Installed = False
                was_open = (
                    self._find_open_workbook_by_path(application, path) is not None
                )
                workbook = self._open_addin_workbook(application, path)
                if workbook is None:
                    raise ExcelConnectionError(
                        "Workbooks.Open không trả về workbook add-in."
                    )
                if was_open:
                    already_loaded.append(path.name)
                else:
                    loaded.append(path.name)
            except (
                ExcelConnectionError,
                com_error,
                AttributeError,
                TypeError,
                OSError,
            ) as exc:
                failed.append((path.name, self._short_com_error(exc)))
        return AddinLoadReport(
            directory=directory.resolve(),
            discovered=tuple(path.name for path in addin_paths),
            loaded=tuple(loaded),
            already_loaded=tuple(already_loaded),
            failed=tuple(failed),
        )

    def install_tool_addins_to_xlstart(
        self, directory: Path | None = None
    ) -> AddinLoadReport:
        source_directory = Path(directory or self.tool_addin_directory())
        xlstart_directory = self.excel_xlstart_directory()
        try:
            xlstart_directory.mkdir(parents=True, exist_ok=True)
            source_directory.mkdir(parents=True, exist_ok=True)
            addin_paths = tuple(
                sorted(
                    (
                        path.resolve()
                        for path in source_directory.iterdir()
                        if path.is_file()
                        and not path.name.startswith("~$")
                        and path.suffix.casefold() in {".xla", ".xlam"}
                        and path.name.casefold() not in self._disabled_addins
                    ),
                    key=lambda path: path.name.casefold(),
                )
            )
        except OSError as exc:
            return AddinLoadReport(
                directory=xlstart_directory,
                failed=(("XLSTART", self._short_com_error(exc)),),
            )

        loaded: list[str] = []
        already_loaded: list[str] = []
        failed: list[tuple[str, str]] = []
        for path in addin_paths:
            destination = xlstart_directory / path.name
            try:
                if destination.exists() and self._same_file_content(path, destination):
                    already_loaded.append(path.name)
                    continue
                shutil.copy2(path, destination)
                loaded.append(path.name)
            except OSError as exc:
                failed.append((path.name, self._short_com_error(exc)))

        return AddinLoadReport(
            directory=xlstart_directory,
            discovered=tuple(path.name for path in addin_paths),
            loaded=tuple(loaded),
            already_loaded=tuple(already_loaded),
            failed=tuple(failed),
        )

    def cleanup_tool_addins(
        self,
        file_names: tuple[str, ...] | list[str] | None = None,
        directory: Path | None = None,
    ) -> AddinLoadReport:
        source_directory = Path(directory or self.tool_addin_directory())
        xlstart_directory = self.excel_xlstart_directory()
        requested = (
            {Path(name).name.casefold() for name in file_names}
            if file_names is not None
            else None
        )
        try:
            addin_paths = tuple(
                sorted(
                    (
                        path.resolve()
                        for path in source_directory.iterdir()
                        if path.is_file()
                        and not path.name.startswith("~$")
                        and path.suffix.casefold() in {".xla", ".xlam"}
                        and (
                            requested is None
                            or path.name.casefold() in requested
                        )
                    ),
                    key=lambda path: path.name.casefold(),
                )
            )
        except OSError as exc:
            return AddinLoadReport(
                directory=xlstart_directory,
                failed=(("Thư mục add-in", self._short_com_error(exc)),),
            )

        cleaned: list[str] = []
        failed: list[tuple[str, str]] = []
        application = self._application if self.is_connected else None
        for path in addin_paths:
            errors: list[str] = []
            if application is not None:
                existing_addin = self._find_matching_addin(application, path)
                if existing_addin is not None:
                    try:
                        existing_addin.Installed = False
                    except (com_error, AttributeError, TypeError, OSError) as exc:
                        errors.append(
                            f"gỡ đăng ký: {self._short_com_error(exc)}"
                        )
                try:
                    self._close_open_addin_workbooks(application, path)
                except (com_error, AttributeError, TypeError, OSError) as exc:
                    errors.append(f"đóng add-in: {self._short_com_error(exc)}")

            destination = xlstart_directory / path.name
            try:
                destination.unlink(missing_ok=True)
            except OSError as exc:
                errors.append(f"xóa XLSTART: {self._short_com_error(exc)}")

            if errors:
                failed.append((path.name, "; ".join(errors)))
            else:
                cleaned.append(path.name)

        return AddinLoadReport(
            directory=xlstart_directory,
            discovered=tuple(path.name for path in addin_paths),
            loaded=tuple(cleaned),
            failed=tuple(failed),
        )

    def cleanup_session_addins(
        self, directory: Path | None = None
    ) -> AddinLoadReport:
        return self.cleanup_tool_addins(directory=directory)

    @classmethod
    def _close_open_addin_workbooks(
        cls, application: Any, path: Path
    ) -> None:
        workbooks = application.Workbooks
        for index in range(int(workbooks.Count), 0, -1):
            workbook = workbooks.Item(index)
            if not cls._workbook_matches(workbook, path):
                continue
            try:
                workbook.Close(SaveChanges=False)
            except TypeError:
                workbook.Close(False)

    @staticmethod
    def _same_file_content(left: Path, right: Path) -> bool:
        try:
            if left.stat().st_size != right.stat().st_size:
                return False
            return left.read_bytes() == right.read_bytes()
        except OSError:
            return False

    def _add_and_install_addin(self, application: Any, path: Path) -> Any:
        errors: list[str] = []
        for copy_file in (False, True):
            try:
                addin = application.AddIns.Add(str(path), copy_file)
                addin.Installed = True
                return addin
            except (com_error, AttributeError, TypeError, OSError) as exc:
                mode = "copy vào thư mục AddIns" if copy_file else "giữ nguyên vị trí"
                errors.append(f"{mode}: {self._short_com_error(exc)}")
        raise ExcelConnectionError("; ".join(errors))

    def _ensure_addin_workbook_open(self, application: Any, path: Path) -> bool:
        if self._find_open_workbook_by_path(application, path) is not None:
            return False
        try:
            workbook = self._open_addin_workbook(application, path)
        except (ExcelConnectionError, com_error, AttributeError, TypeError, OSError):
            return False
        return workbook is not None

    def _open_addin_workbook(self, application: Any, path: Path) -> Any:
        existing_workbook = self._find_open_workbook_by_path(application, path)
        if existing_workbook is not None:
            return existing_workbook
        try:
            workbook = application.Workbooks.Open(
                Filename=str(path),
                ReadOnly=True,
                AddToMru=False,
            )
        except TypeError:
            workbook = application.Workbooks.Open(str(path))
        try:
            workbook.IsAddin = True
        except (com_error, AttributeError, TypeError):
            pass
        return workbook

    @staticmethod
    def _find_open_workbook_by_path(application: Any, path: Path) -> Any | None:
        try:
            workbooks = application.Workbooks
            for index in range(1, int(workbooks.Count) + 1):
                workbook = workbooks.Item(index)
                if ExcelService._workbook_matches(workbook, path):
                    return workbook
        except (com_error, AttributeError, TypeError, OSError):
            return None
        return None

    @staticmethod
    def _workbook_matches(workbook: Any, path: Path) -> bool:
        expected_name = path.name.casefold()
        try:
            expected_path = str(path.resolve()).casefold()
        except OSError:
            expected_path = str(path).casefold()
        full_name = str(getattr(workbook, "FullName", "") or "")
        if full_name:
            try:
                if str(Path(full_name).resolve()).casefold() == expected_path:
                    return True
            except OSError:
                if full_name.casefold() == expected_path:
                    return True
        return (
            str(getattr(workbook, "Name", "") or "").casefold()
            == expected_name
        )

    @staticmethod
    def _installed_addin_identifiers(application: Any) -> set[str]:
        identifiers: set[str] = set()
        try:
            addins = application.AddIns
            for index in range(1, int(addins.Count) + 1):
                addin = addins.Item(index)
                try:
                    if not bool(getattr(addin, "Installed", False)):
                        continue
                except (com_error, AttributeError, TypeError):
                    continue
                full_name = str(getattr(addin, "FullName", "") or "")
                if full_name:
                    try:
                        identifiers.add(
                            f"path:{str(Path(full_name).resolve()).casefold()}"
                        )
                    except OSError:
                        identifiers.add(f"path:{full_name.casefold()}")
                name = str(getattr(addin, "Name", "") or "")
                if not name and full_name:
                    name = Path(full_name).name
                if name:
                    identifiers.add(f"name:{name.casefold()}")
        except (com_error, AttributeError, TypeError, OSError):
            pass
        return identifiers

    @classmethod
    def _find_matching_addin(cls, application: Any, path: Path) -> Any | None:
        try:
            addins = application.AddIns
            for index in range(1, int(addins.Count) + 1):
                addin = addins.Item(index)
                if cls._addin_matches(addin, path):
                    return addin
        except (com_error, AttributeError, TypeError, OSError):
            return None
        return None

    @staticmethod
    def _addin_matches(addin: Any, path: Path) -> bool:
        expected_name = path.name.casefold()
        try:
            expected_path = str(path.resolve()).casefold()
        except OSError:
            expected_path = str(path).casefold()
        full_name = str(getattr(addin, "FullName", "") or "")
        if full_name:
            try:
                if str(Path(full_name).resolve()).casefold() == expected_path:
                    return True
            except OSError:
                if full_name.casefold() == expected_path:
                    return True
        name = str(getattr(addin, "Name", "") or "")
        if not name and full_name:
            name = Path(full_name).name
        return name.casefold() == expected_name

    @staticmethod
    def _is_addin_installed(addin: Any) -> bool:
        try:
            return bool(getattr(addin, "Installed", False))
        except (com_error, AttributeError, TypeError):
            return False

    @staticmethod
    def _open_workbook_identifiers(application: Any) -> set[str]:
        identifiers: set[str] = set()
        try:
            workbooks = application.Workbooks
            for index in range(1, int(workbooks.Count) + 1):
                workbook = workbooks.Item(index)
                full_name = str(getattr(workbook, "FullName", "") or "")
                if full_name:
                    identifiers.add(
                        f"path:{str(Path(full_name).resolve()).casefold()}"
                    )
                name = str(getattr(workbook, "Name", "") or "")
                if not name and full_name:
                    name = Path(full_name).name
                if name:
                    identifiers.add(f"name:{name.casefold()}")
        except (com_error, AttributeError, TypeError, OSError):
            pass
        return identifiers

    def _restore_context(self, context: ExcelContext) -> ExcelContext:
        """Return focus to the user's workbook after hidden add-ins are opened."""
        try:
            workbook = self._find_open_workbook(context.workbook)
            workbook.Activate()
            worksheet = self._require_compatibility().worksheet(
                workbook, context.worksheet
            )
            worksheet.Activate()
            return self.get_context()
        except (com_error, AttributeError, ExcelConnectionError):
            return context

    @staticmethod
    def _short_com_error(error: Exception) -> str:
        text = str(error).strip().replace("\r", " ").replace("\n", " ")
        return text[:240] or error.__class__.__name__

    def _discover_applications(
        self, preferred_workbook_path: Path | None = None
    ) -> tuple[list[Any], Exception | None]:
        del preferred_workbook_path
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

    def _rank_applications(
        self,
        applications: list[Any],
        *,
        required_major_version: int | None,
        preferred_workbook_path: Path | None,
    ) -> list[Any]:
        def score(application: Any) -> tuple[int, int]:
            version_score = 0
            workbook_score = 0
            try:
                capabilities = ExcelCompatibility(application).capabilities
                if (
                    required_major_version is not None
                    and capabilities.major_version == required_major_version
                ):
                    version_score = 1
            except (com_error, AttributeError, TypeError):
                pass
            if (
                preferred_workbook_path is not None
                and self._application_has_workbook(
                    application, preferred_workbook_path
                )
            ):
                workbook_score = 1
            return (workbook_score, version_score)

        return sorted(applications, key=score, reverse=True)

    @staticmethod
    def _application_has_workbook(application: Any, workbook_path: Path) -> bool:
        expected_path = str(workbook_path.resolve()).casefold()
        expected_name = workbook_path.name.casefold()
        try:
            workbooks = application.Workbooks
            for index in range(1, int(workbooks.Count) + 1):
                workbook = workbooks.Item(index)
                full_name = str(getattr(workbook, "FullName", "") or "")
                if full_name:
                    try:
                        if str(Path(full_name).resolve()).casefold() == expected_path:
                            return True
                    except OSError:
                        if full_name.casefold() == expected_path:
                            return True
                name = str(getattr(workbook, "Name", "") or "")
                if name.casefold() == expected_name:
                    return True
        except (com_error, AttributeError, TypeError, OSError):
            return False
        return False

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

    def active_workbook_path(self) -> Path:
        application = self._require_application()
        try:
            full_name = str(application.ActiveWorkbook.FullName)
        except (com_error, AttributeError, TypeError) as exc:
            raise ExcelConnectionError(
                "Workbook đang hoạt động chưa được lưu thành file.",
                code="workbook_has_no_path",
            ) from exc
        path = Path(full_name)
        if not path.is_file():
            raise ExcelConnectionError(
                f"Không tìm thấy file workbook đang hoạt động: {path}",
                code="workbook_has_no_path",
            )
        return path

    def open_workbook(self, path: Path) -> ExcelContext:
        application = self._require_application()
        resolved = Path(path).resolve()
        workbook = self._find_open_workbook_by_path(application, resolved)
        try:
            if workbook is None:
                workbook = application.Workbooks.Open(str(resolved))
            workbook.Activate()
            return self.get_context()
        except (com_error, AttributeError, TypeError, OSError) as exc:
            raise ExcelConnectionError(
                f"Không thể mở workbook kết quả: {resolved}",
                code="open_output_failed",
            ) from exc

    @classmethod
    def is_system_worksheet(cls, context: ExcelContext) -> bool:
        """Return True for internal worksheets shipped inside legacy add-ins."""
        workbook_suffix = context.workbook.casefold().rsplit(".", 1)
        is_addin = len(workbook_suffix) == 2 and workbook_suffix[-1] in {
            "xla",
            "xlam",
        }
        return is_addin and context.worksheet.upper() in cls.LEGACY_SYSTEM_SHEETS

    def convert_selection_case(self, mode: CaseMode) -> ConversionResult:
        application = self._require_application()
        context = self.get_context()
        if self.is_system_worksheet(context):
            raise ExcelConnectionError(
                "Không thể xử lý worksheet hệ thống của add-in."
            )
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
