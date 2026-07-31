from __future__ import annotations

from pathlib import Path
from threading import RLock
from typing import Callable

from agribank_v3.features.settings.unit_directory.models import (
    AppUnitSettings,
    BranchDirectoryEntry,
    OfficeDirectoryEntry,
)
from agribank_v3.features.settings.unit_directory.repository import (
    UnitDirectoryRepository,
    build_office_code,
    default_office_type,
    normalize_branch_code,
    normalize_trctcd,
)


DirectoryChangedCallback = Callable[[], None]


class UnitDirectoryService:
    def __init__(self, database_path: Path | str) -> None:
        self.database_path = Path(database_path)
        self.repository = UnitDirectoryRepository(self.database_path)
        self._lock = RLock()
        self._loaded = False
        self._branches_by_code: dict[str, BranchDirectoryEntry] = {}
        self._offices_by_code: dict[str, OfficeDirectoryEntry] = {}
        self._offices_by_branch_trctcd: dict[tuple[str, str], OfficeDirectoryEntry] = {}
        self._offices_by_branch: dict[str, tuple[OfficeDirectoryEntry, ...]] = {}
        self._settings = AppUnitSettings()
        self._listeners: list[DirectoryChangedCallback] = []

    def add_listener(self, callback: DirectoryChangedCallback) -> None:
        with self._lock:
            if callback not in self._listeners:
                self._listeners.append(callback)

    def remove_listener(self, callback: DirectoryChangedCallback) -> None:
        with self._lock:
            if callback in self._listeners:
                self._listeners.remove(callback)

    def invalidate(self, *, notify: bool = True) -> None:
        listeners: list[DirectoryChangedCallback] = []
        with self._lock:
            self._loaded = False
            self._branches_by_code = {}
            self._offices_by_code = {}
            self._offices_by_branch_trctcd = {}
            self._offices_by_branch = {}
            if notify:
                listeners = list(self._listeners)
        for callback in listeners:
            callback()

    def refresh(self) -> None:
        with self._lock:
            branches = self.repository.list_branches()
            offices = self.repository.list_offices()
            self._branches_by_code = {item.branch_code: item for item in branches}
            self._offices_by_code = {item.office_code: item for item in offices}
            self._offices_by_branch_trctcd = {
                (item.branch_code, item.trctcd): item
                for item in offices
            }
            by_branch: dict[str, list[OfficeDirectoryEntry]] = {}
            for item in offices:
                by_branch.setdefault(item.branch_code, []).append(item)
            self._offices_by_branch = {
                branch: tuple(items)
                for branch, items in by_branch.items()
            }
            self._settings = self.repository.load_settings()
            self._loaded = True

    def get_active_branches(self) -> tuple[BranchDirectoryEntry, ...]:
        self._ensure_loaded()
        with self._lock:
            return tuple(
                item for item in self._branches_by_code.values()
                if item.is_active
            )

    def get_branch(self, branch_code: object) -> BranchDirectoryEntry | None:
        code = normalize_branch_code(branch_code)
        if not code:
            return None
        self._ensure_loaded()
        with self._lock:
            return self._branches_by_code.get(code)

    def get_branch_name(self, branch_code: object) -> str:
        return self.format_branch_display(branch_code)

    def get_branch_display_name(self, branch_code: object) -> str:
        return self.format_branch_display(branch_code)

    def get_active_offices(self, branch_code: object = "") -> tuple[OfficeDirectoryEntry, ...]:
        branch = normalize_branch_code(branch_code)
        self._ensure_loaded()
        with self._lock:
            if branch:
                offices = self._offices_by_branch.get(branch, ())
            else:
                offices = tuple(self._offices_by_code.values())
            return tuple(item for item in offices if item.is_active)

    def get_office(self, branch_code: object, trctcd: object) -> OfficeDirectoryEntry | None:
        branch = normalize_branch_code(branch_code)
        code = normalize_trctcd(trctcd)
        if not branch or not code:
            return None
        self._ensure_loaded()
        with self._lock:
            return self._offices_by_branch_trctcd.get((branch, code))

    def get_office_by_code(self, office_code: object) -> OfficeDirectoryEntry | None:
        code = "" if office_code is None else str(office_code).strip()
        if not code:
            return None
        self._ensure_loaded()
        with self._lock:
            return self._offices_by_code.get(code)

    def get_office_name(self, branch_code: object, trctcd: object) -> str:
        office = self.get_office(branch_code, trctcd)
        if office is not None:
            return _office_label(office)
        return _fallback_office_name(trctcd)

    def get_office_display_name(self, branch_code: object, trctcd: object) -> str:
        office = self.get_office(branch_code, trctcd)
        if office is not None:
            return f"{office.office_code} - {_office_label(office)}"
        branch = normalize_branch_code(branch_code)
        code = normalize_trctcd(trctcd)
        office_code = build_office_code(branch, code)
        label = _fallback_office_name(code)
        return f"{office_code} - {label}" if office_code else label

    def get_home_branch(self) -> BranchDirectoryEntry | None:
        settings = self.get_settings()
        return self.get_branch(settings.home_branch_code)

    def get_default_office(self) -> OfficeDirectoryEntry | None:
        settings = self.get_settings()
        return self.get_office_by_code(settings.default_office_code)

    def get_settings(self) -> AppUnitSettings:
        self._ensure_loaded()
        with self._lock:
            return self._settings

    def save_settings(self, settings: AppUnitSettings) -> AppUnitSettings:
        saved = self.repository.save_settings(settings)
        self.invalidate()
        return saved

    def save_branch(self, entry: BranchDirectoryEntry) -> BranchDirectoryEntry:
        saved = self.repository.save_branch(entry)
        self.invalidate()
        return saved

    def create_branch(self, entry: BranchDirectoryEntry) -> BranchDirectoryEntry:
        saved = self.repository.create_branch(entry)
        self.invalidate()
        return saved

    def set_branch_active(self, branch_code: object, active: bool, *, updated_by: str = "") -> None:
        self.repository.set_branch_active(branch_code, active, updated_by=updated_by)
        self.invalidate()

    def save_office(self, entry: OfficeDirectoryEntry) -> OfficeDirectoryEntry:
        saved = self.repository.save_office(entry)
        self.invalidate()
        return saved

    def create_office(self, entry: OfficeDirectoryEntry) -> OfficeDirectoryEntry:
        saved = self.repository.create_office(entry)
        self.invalidate()
        return saved

    def set_office_active(self, office_code: object, active: bool, *, updated_by: str = "") -> None:
        self.repository.set_office_active(office_code, active, updated_by=updated_by)
        self.invalidate()

    def format_branch_display(self, branch_code: object) -> str:
        code = normalize_branch_code(branch_code)
        if not code:
            return ""
        branch = self.get_branch(code)
        if branch is None:
            return f"{code} - CN chưa khai báo"
        if branch.display_name:
            return branch.display_name
        label = branch.short_name or branch.branch_name
        return f"{branch.branch_code} - {label}" if label else branch.branch_code

    def format_office_display(self, branch_code: object, trctcd: object) -> str:
        return self.get_office_display_name(branch_code, trctcd)

    def ensure_known_unit(
        self,
        branch_code: object,
        trctcd: object = "",
        *,
        updated_by: str = "system",
    ) -> tuple[str, ...]:
        branch = normalize_branch_code(branch_code)
        code = normalize_trctcd(trctcd)
        if not branch:
            return ()
        warnings: list[str] = []
        branch_created = self.repository.ensure_branch_placeholder(branch, updated_by=updated_by)
        if branch_created:
            warnings.append(
                f"Phát hiện mã chi nhánh {branch} chưa được khai báo. Đã tạo placeholder trong Cài đặt thông tin đơn vị."
            )
        if code:
            office_created = self.repository.ensure_office_placeholder(branch, code, updated_by=updated_by)
            if office_created:
                office_label = "Hội sở" if code == "00" else f"PGD {code}"
                warnings.append(
                    f"Phát hiện mã đơn vị {branch}-{code} chưa được khai báo ({office_label}). Đã tạo placeholder."
                )
        if branch_created or (code and len(warnings) > (1 if branch_created else 0)):
            self.invalidate()
        return tuple(warnings)

    def _ensure_loaded(self) -> None:
        with self._lock:
            loaded = self._loaded
        if not loaded:
            self.refresh()


def get_unit_directory_service(database_path: Path | str | None = None) -> UnitDirectoryService:
    if database_path is None:
        from agribank_v3.settings import AppSettingsDatabase

        database_path = AppSettingsDatabase().database_path
    path = Path(database_path)
    key = str(path.resolve())
    with _SERVICES_LOCK:
        service = _SERVICES.get(key)
        if service is None:
            service = UnitDirectoryService(path)
            _SERVICES[key] = service
        return service


def invalidate_unit_directory_cache(database_path: Path | str | None = None) -> None:
    if database_path is None:
        with _SERVICES_LOCK:
            services = list(_SERVICES.values())
        for service in services:
            service.invalidate()
        return
    key = str(Path(database_path).resolve())
    with _SERVICES_LOCK:
        service = _SERVICES.get(key)
    if service is not None:
        service.invalidate()


def _office_label(office: OfficeDirectoryEntry) -> str:
    return office.short_name or office.office_name or _fallback_office_name(office.trctcd)


def _fallback_office_name(trctcd: object) -> str:
    code = normalize_trctcd(trctcd)
    if not code:
        return "Không xác định"
    if code == "00":
        return "Hội sở"
    if code.isdigit():
        return f"PGD {code}"
    office_type = default_office_type(code)
    return "Hội sở" if office_type == "HEAD_OFFICE" else str(code)


_SERVICES_LOCK = RLock()
_SERVICES: dict[str, UnitDirectoryService] = {}
