from __future__ import annotations

from contextlib import closing
from dataclasses import dataclass, field
from datetime import date, datetime
import fnmatch
import hashlib
import json
import os
import re
import shutil
import sys
from pathlib import Path
import sqlite3
import tempfile
from typing import Callable, Iterable
from uuid import uuid4
import zipfile

from update_manifest import ManifestMigration, manifest_dict
from db_schema_diff import (
    SchemaDiff,
    SQLiteSchema,
    compare_sqlite_schema,
    generate_python_migration,
    migration_function_name,
    read_sqlite_schema,
    validate_generated_migration_on_copy,
    write_generated_python_migration,
)


DEFAULT_UPDATE_PATH = (
    r"X:\public\5491\06- PHONG KHKD\NAM\AgribankV3\Update"
)

PACKAGE_MODE_RUNTIME = "runtime"
PACKAGE_MODE_SOURCE = "source"
PACKAGE_MODE_DELTA = "delta"
PACKAGE_MODE_APP = "app"
PACKAGE_MODES = {
    PACKAGE_MODE_RUNTIME,
    PACKAGE_MODE_SOURCE,
    PACKAGE_MODE_DELTA,
    PACKAGE_MODE_APP,
}
PAYLOAD_LAYOUT_SOURCE = "source"
PAYLOAD_LAYOUT_APP_ROOT = "app_root"
APP_BUILD_INFO_FILE_NAME = "agribank_v3_build_info.json"

PACKAGE_SIZE_WARNING_BYTES = 100 * 1024 * 1024

DEFAULT_EXCLUDES = [
    ".git",
    ".venv",
    "venv",
    "__pycache__",
    ".pytest_cache",
    ".mypy_cache",
    "dist",
    "build",
    "tests",
    "docs",
    "DuLieuTEST",
    "Test",
    "test_data",
    "sample_data",
    "KetQua",
    "outputs",
    "logs",
    "backups",
    "backup",
    "temp",
    "Update/*",
    "_backup",
    "tools/update_builder",
    "tools/update_builder/logs",
    "tools/update_builder/generated_migrations",
    "tools/update_builder/schema_snapshots",
    "tools/update_builder/release_snapshots",
    "data/*.db",
    "data/*.sqlite",
    "data/*.sqlite3",
    "*.db",
    "*.sqlite",
    "*.sqlite3",
    "*.mdb",
    "*.accdb",
    "*.zip",
    "*.rar",
    "*.7z",
    "*.exe",
    "*.msi",
    "*.spec",
    "*.log",
    "*.tmp",
    "data/backups",
]

APP_PACKAGE_EXCLUDES = tuple(
    item
    for item in DEFAULT_EXCLUDES
    if item.casefold() not in {"dist", "*.exe"}
)
APP_BUILD_RELATIVE_ROOT = Path("dist") / "AgribankV3"

RUNTIME_INCLUDE_DIRS = (
    "src/agribank_v3",
    "assets",
    "resources",
    "templates",
    "MauBieu",
    "mau_bieu",
    "static",
)

RUNTIME_INCLUDE_FILES = (
    "pyproject.toml",
    "requirements.txt",
    "run.ps1",
    "run.bat",
    "README.md",
)

ALREADY_COMPRESSED_SUFFIXES = {
    ".zip",
    ".rar",
    ".7z",
    ".jpg",
    ".jpeg",
    ".png",
    ".gif",
    ".webp",
    ".pdf",
    ".xlsx",
    ".xlsm",
    ".docx",
    ".pptx",
}

DANGEROUS_SQL_PATTERNS = (
    r"\bDROP\s+TABLE\b",
    r"\bDELETE\s+FROM\b",
    r"\bTRUNCATE\b",
)


class UpdateBuilderError(RuntimeError):
    pass


@dataclass(frozen=True, slots=True)
class MigrationItem:
    version: str
    description: str
    source_file: Path | None = None
    use_python_migration: bool = False

    @property
    def manifest_file(self) -> str:
        if self.use_python_migration or self.source_file is None:
            return ""
        return f"migrations/{self.source_file.name}"


@dataclass(frozen=True, slots=True)
class BuildUpdateConfig:
    source_path: Path
    update_path: Path
    new_version: str
    release_date: str = field(default_factory=lambda: date.today().isoformat())
    required_app_restart: bool = True
    notes: tuple[str, ...] = ()
    migrations: tuple[MigrationItem, ...] = ()
    auto_update_source_version: bool = False
    database_changed: bool = False
    previous_release_version: str | None = None
    allow_rebuild_same_version: bool = False
    package_mode: str = PACKAGE_MODE_RUNTIME
    excludes: tuple[str, ...] = tuple(DEFAULT_EXCLUDES)
    verbose: bool = False


@dataclass(frozen=True, slots=True)
class CollectedFile:
    path: Path
    relative_path: Path
    include_reason: str = "included"


@dataclass(frozen=True, slots=True)
class ExcludedPath:
    path: Path
    relative_path: Path
    reason: str


@dataclass(frozen=True, slots=True)
class FileSizeInfo:
    relative_path: Path
    size: int


@dataclass(frozen=True, slots=True)
class DirectorySizeInfo:
    relative_path: Path
    size: int


@dataclass(frozen=True, slots=True)
class PackageFileReport:
    included_count: int
    excluded_count: int
    total_size: int
    estimated_zip_size: int
    top_files: tuple[FileSizeInfo, ...]
    top_directories: tuple[DirectorySizeInfo, ...]


@dataclass(frozen=True, slots=True)
class PackagePlan:
    package_mode: str
    package_type: str
    payload_layout: str
    base_version: str
    included: tuple[CollectedFile, ...]
    snapshot_files: tuple[CollectedFile, ...]
    excluded: tuple[ExcludedPath, ...]
    delete_files: tuple[Path, ...]
    report: PackageFileReport


@dataclass(frozen=True, slots=True)
class ValidationResult:
    current_version: str
    previous_release_version: str | None
    warnings: tuple[str, ...]
    package_plan: PackagePlan | None = None


@dataclass(frozen=True, slots=True)
class BuildResult:
    package_path: Path
    manifest_path: Path
    backup_path: Path | None
    migration_paths: tuple[Path, ...]
    included_count: int
    excluded_count: int
    warnings: tuple[str, ...]
    log_path: Path
    package_mode: str = PACKAGE_MODE_RUNTIME
    package_type: str = "full"
    payload_layout: str = PAYLOAD_LAYOUT_SOURCE
    base_version: str = ""
    delete_files: tuple[Path, ...] = ()
    package_report: PackageFileReport | None = None
    release_snapshot_path: Path | None = None


@dataclass(frozen=True, slots=True)
class AppBuildInfo:
    app: str
    version: str


@dataclass(frozen=True, slots=True)
class AutoBuildConfig:
    source_path: Path
    update_path: Path
    new_version: str
    release_date: str = field(default_factory=lambda: date.today().isoformat())
    required_app_restart: bool = True
    notes: tuple[str, ...] = ()
    auto_update_source_version: bool = False
    dev_db_path: Path | None = None
    baseline_db_path: Path | None = None
    snapshot_dir: Path | None = None
    generated_migration_dir: Path | None = None
    code_only_if_missing_baseline: bool = False
    previous_release_version: str | None = None
    allow_rebuild_same_version: bool = False
    package_mode: str = PACKAGE_MODE_RUNTIME
    verbose: bool = False


@dataclass(frozen=True, slots=True)
class SchemaSnapshotInfo:
    version: str
    path: Path


@dataclass(frozen=True, slots=True)
class AutoDatabaseDetectionResult:
    status: str
    current_version: str
    new_version: str
    dev_db_path: Path | None
    snapshot_path: Path | None
    baseline_db_path: Path | None
    diff: SchemaDiff | None
    message: str

    @property
    def has_database_changes(self) -> bool:
        return self.status in {"safe_changes", "dangerous_changes"}


@dataclass(frozen=True, slots=True)
class AutoBuildResult:
    build_result: BuildResult
    detection: AutoDatabaseDetectionResult
    generated_migration_path: Path | None
    inserted_migration_backup_path: Path | None
    migration_test_database_path: Path | None
    saved_snapshot_path: Path | None


class UpdateBuilder:
    def __init__(
        self,
        logger: Callable[[str], None] | None = None,
        progress_callback: Callable[[int], None] | None = None,
    ) -> None:
        self.logger = logger or (lambda message: None)
        self.progress_callback = progress_callback or (lambda value: None)

    def validate(self, config: BuildUpdateConfig) -> ValidationResult:
        return self.validate_build_deep(config)

    def validate_build_deep(self, config: BuildUpdateConfig) -> ValidationResult:
        source_path = Path(config.source_path)
        update_path = Path(config.update_path)
        warnings: list[str] = []
        if not source_path.is_dir():
            raise UpdateBuilderError(f"Không tìm thấy thư mục source: {source_path}")
        if not (source_path / "src" / "agribank_v3").is_dir():
            raise UpdateBuilderError(
                "Thư mục source không hợp lệ. Vui lòng chọn thư mục gốc "
                "AgribankV3 có src/agribank_v3."
            )
        if not config.notes:
            raise UpdateBuilderError("Release notes không được để trống.")
        _parse_semver(config.new_version)
        normalize_package_mode(config.package_mode)
        _parse_date(config.release_date)
        update_path.mkdir(parents=True, exist_ok=True)
        if _extra_excluded_paths_for_update(source_path, update_path):
            warnings.append(
                "Thư mục Update nằm trong source, Builder sẽ loại trừ khỏi gói zip."
            )
        current_version = read_current_version(source_path)
        previous_release_version = config.previous_release_version
        if previous_release_version is None:
            previous_release_version = detect_previous_release_version(
                update_path,
                source_path,
                config.new_version,
            )
        if previous_release_version is None:
            warnings.append(
                "Không xác định được version phát hành trước. "
                f"Builder sẽ tạo gói version {config.new_version}."
            )
        else:
            comparison = compare_versions(config.new_version, previous_release_version)
            if comparison < 0:
                raise UpdateBuilderError(
                    "Version mới phải lớn hơn version phát hành trước "
                    f"({previous_release_version})."
                )
            if comparison == 0:
                if not config.allow_rebuild_same_version:
                    raise UpdateBuilderError(
                        "Bản cập nhật "
                        f"{config.new_version} đã là version phát hành trước. "
                        "Cần xác nhận cho phép tạo lại cùng version để test."
                    )
                warnings.append(
                    f"Đang tạo lại bản cập nhật cùng version {config.new_version}; "
                    "file cũ sẽ được backup nếu tồn tại."
                )
        source_comparison = compare_versions(config.new_version, current_version)
        if source_comparison < 0:
            raise UpdateBuilderError(
                "Phiên bản mới thấp hơn version trong source "
                f"({current_version}). Vui lòng kiểm tra lại."
            )
        if normalize_package_mode(config.package_mode) == PACKAGE_MODE_APP:
            app_build_info = read_app_build_info(source_path / APP_BUILD_RELATIVE_ROOT)
            if compare_versions(app_build_info.version, config.new_version) != 0:
                raise UpdateBuilderError(
                    "Version thật của bản build app "
                    f"({app_build_info.version}) không khớp version gói "
                    f"({config.new_version}). Hãy chạy build_portable.ps1 sau khi "
                    "cập nhật version source rồi tạo lại gói app."
                )
        if source_comparison > 0 and not config.auto_update_source_version:
            warnings.append(
                "Phiên bản mới lớn hơn version trong source. Nên tích tự cập nhật "
                "version hoặc sửa source trước khi đóng gói."
            )
        package_path = update_path / package_name(config.new_version)
        if package_path.exists():
            warnings.append(f"Package đã tồn tại và sẽ được backup: {package_path.name}")
        if (update_path / "manifest.json").exists():
            warnings.append("manifest.json cũ sẽ được backup trước khi ghi mới.")
        for migration in config.migrations:
            _parse_semver(migration.version)
            if not migration.description.strip():
                raise UpdateBuilderError("Migration phải có description.")
            if migration.use_python_migration:
                if not source_has_python_migration(source_path, migration.version):
                    warnings.append(
                        "Manifest khai báo Python migration version "
                        f"{migration.version} nhưng chưa tìm thấy "
                        f"{migration_function_name(migration.version)} trong db_migrations.py."
                    )
            else:
                if migration.source_file is None:
                    raise UpdateBuilderError("Migration SQL chưa chọn file.")
                if not migration.source_file.is_file():
                    raise UpdateBuilderError(
                        f"Không tìm thấy file migration: {migration.source_file}"
                    )
                warnings.extend(warn_dangerous_sql(migration.source_file))
        if config.database_changed and not config.migrations:
            warnings.append(
                "Bạn đã chọn có thay đổi database nhưng chưa thêm migration vào manifest."
            )
        package_plan = create_package_plan(config, previous_release_version)
        database_files = [
            item.relative_path
            for item in package_plan.included
            if _is_database_like(item.relative_path)
        ]
        if database_files:
            raise UpdateBuilderError(
                "Phát hiện database trong danh sách đóng gói: "
                + ", ".join(str(path) for path in database_files[:10])
            )
        if package_plan.report.total_size > PACKAGE_SIZE_WARNING_BYTES:
            warnings.append(
                "Gói cập nhật có dung lượng lớn. Vui lòng kiểm tra danh sách file "
                "trước khi phát hành."
            )
        return ValidationResult(
            current_version=current_version,
            previous_release_version=previous_release_version,
            warnings=tuple(warnings),
            package_plan=package_plan,
        )

    def build(self, config: BuildUpdateConfig) -> BuildResult:
        validation = self.validate(config)
        self._log(f"Source: {config.source_path}")
        self._log(f"Phiên bản trong source: {validation.current_version}")
        self._log(
            "Phiên bản phát hành trước: "
            + (validation.previous_release_version or "không xác định")
        )
        self._log(f"Phiên bản gói cập nhật sẽ tạo: {config.new_version}")
        if validation.previous_release_version:
            comparison = compare_versions(config.new_version, validation.previous_release_version)
            if comparison > 0:
                self._log(
                    f"Version hợp lệ: {config.new_version} > "
                    f"{validation.previous_release_version}."
                )
            elif comparison == 0 and config.allow_rebuild_same_version:
                self._log(
                    f"Đang tạo lại cùng version {config.new_version} theo xác nhận."
                )
        if compare_versions(config.new_version, validation.current_version) == 0:
            self._log("Version trong source đã khớp với version gói cập nhật.")
        if (
            config.auto_update_source_version
            and compare_versions(config.new_version, validation.current_version) != 0
        ):
            updated_files = update_source_version(config.source_path, config.new_version)
            for path in updated_files:
                self._log(f"Đã cập nhật version: {path}")
        elif compare_versions(config.new_version, validation.current_version) != 0:
            self._log(
                "Cảnh báo: version trong source chưa khớp version mới; "
                "package vẫn được tạo theo version đã nhập."
            )

        update_path = Path(config.update_path)
        package_path = update_path / package_name(config.new_version)
        backup_path = backup_existing_update_files(
            update_path,
            package_path.name,
        )
        if backup_path:
            self._log(f"Đã backup file update cũ: {backup_path}")

        extra_excluded_paths = _extra_excluded_paths_for_update(
            config.source_path,
            config.update_path,
        )
        self._log("Đang quét file source...")
        self._log("Đang phân tích gói cập nhật...")
        package_plan = create_package_plan(config, validation.previous_release_version)
        included = list(package_plan.included)
        excluded = list(package_plan.excluded)
        self._log(f"Kiểu gói cập nhật: {package_plan.package_mode}")
        self._log(f"Layout payload: {package_plan.payload_layout}")
        if package_plan.package_type == "delta":
            self._log(f"Phiên bản nền delta: {package_plan.base_version}")
            self._log(f"File sẽ xóa khi cập nhật delta: {len(package_plan.delete_files)}")
        self._log(format_package_report(package_plan.report))
        self._log(
            f"Đã quét source: {len(included)} file sẽ đóng gói, "
            f"{len(excluded)} mục bị loại trừ."
        )
        if package_plan.report.total_size > PACKAGE_SIZE_WARNING_BYTES:
            self._log(
                "Cảnh báo: gói cập nhật có dung lượng lớn. "
                "Vui lòng kiểm tra danh sách file trước khi phát hành."
            )
        if config.verbose:
            for item in excluded:
                self._log(f"Bỏ qua: {item.relative_path} ({item.reason})")
        self._log("Đang tạo zip...")
        create_update_zip(
            config.source_path,
            package_path,
            config.excludes,
            logger=self._log,
            progress_callback=self.progress_callback,
            precollected_files=included,
            extra_excluded_paths=extra_excluded_paths,
        )
        self._log(f"Đã tạo zip: {package_path}")

        migration_paths = copy_migrations(config)
        for path in migration_paths:
            self._log(f"Đã copy migration: {path}")

        release_snapshot_path = save_release_file_snapshot(
            config.source_path,
            config.new_version,
            package_plan.snapshot_files,
            package_mode=package_plan.package_mode,
        )
        self._log(f"Đã lưu file snapshot: {release_snapshot_path}")

        manifest = build_manifest(
            config,
            package_path.name,
            package_type=package_plan.package_type,
            payload_layout=package_plan.payload_layout,
            base_version=package_plan.base_version,
            delete_files=package_plan.delete_files,
        )
        manifest_path = write_manifest(update_path, manifest)
        self._log(f"Đã tạo manifest: {manifest_path}")
        log_path = write_builder_log(
            {
                "created_at": _now(),
                "source_path": str(config.source_path),
                "current_version": validation.current_version,
                "new_version": config.new_version,
                "update_path": str(config.update_path),
                "zip_output": str(package_path),
                "manifest_output": str(manifest_path),
                "migrations": [str(path) for path in migration_paths],
                "included_count": len(included),
                "excluded_count": len(excluded),
                "package_mode": package_plan.package_mode,
                "package_type": package_plan.package_type,
                "payload_layout": package_plan.payload_layout,
                "base_version": package_plan.base_version,
                "delete_files": [path.as_posix() for path in package_plan.delete_files],
                "package_size_bytes": package_plan.report.total_size,
                "estimated_zip_size_bytes": package_plan.report.estimated_zip_size,
                "release_snapshot": str(release_snapshot_path),
                "warnings": list(validation.warnings),
            }
        )
        return BuildResult(
            package_path=package_path,
            manifest_path=manifest_path,
            backup_path=backup_path,
            migration_paths=tuple(migration_paths),
            included_count=len(included),
            excluded_count=len(excluded),
            warnings=validation.warnings,
            log_path=log_path,
            package_mode=package_plan.package_mode,
            package_type=package_plan.package_type,
            payload_layout=package_plan.payload_layout,
            base_version=package_plan.base_version,
            delete_files=package_plan.delete_files,
            package_report=package_plan.report,
            release_snapshot_path=release_snapshot_path,
        )

    def _log(self, message: str) -> None:
        self.logger(message)


def read_current_version(source_path: Path | str) -> str:
    source = Path(source_path)
    candidates = (
        source / "src" / "agribank_v3" / "__init__.py",
        source / "src" / "agribank_v3" / "version.py",
    )
    for path in candidates:
        if not path.is_file():
            continue
        match = re.search(
            r"__version__\s*=\s*['\"]([^'\"]+)['\"]",
            path.read_text(encoding="utf-8"),
        )
        if match:
            return match.group(1)
    pyproject = source / "pyproject.toml"
    if pyproject.is_file():
        match = re.search(
            r"(?m)^version\s*=\s*['\"]([^'\"]+)['\"]",
            pyproject.read_text(encoding="utf-8"),
        )
        if match:
            return match.group(1)
    raise UpdateBuilderError("Không đọc được version hiện tại từ source.")


def read_app_build_info(app_root: Path | str) -> AppBuildInfo:
    root = Path(app_root)
    build_info_path = root / APP_BUILD_INFO_FILE_NAME
    if not build_info_path.is_file():
        raise UpdateBuilderError(
            f"Thư mục build thiếu {APP_BUILD_INFO_FILE_NAME}: {root}. "
            "Hãy build lại AgribankV3 bằng build_portable.ps1 trước khi tạo gói app."
        )
    try:
        raw = json.loads(build_info_path.read_text(encoding="utf-8-sig"))
    except Exception as exc:
        raise UpdateBuilderError(
            f"Không đọc được {APP_BUILD_INFO_FILE_NAME}: {exc}"
        ) from exc
    if not isinstance(raw, dict):
        raise UpdateBuilderError(f"{APP_BUILD_INFO_FILE_NAME} không đúng cấu trúc.")
    app_name = str(raw.get("app", "")).strip()
    version = str(raw.get("version", "")).strip()
    if app_name and app_name != "AgribankV3":
        raise UpdateBuilderError(f"Build app không phải AgribankV3: {app_name}")
    if not version:
        raise UpdateBuilderError(f"{APP_BUILD_INFO_FILE_NAME} thiếu trường version.")
    _parse_semver(version)
    return AppBuildInfo(app=app_name or "AgribankV3", version=version)


def update_source_version(source_path: Path | str, new_version: str) -> list[Path]:
    _parse_semver(new_version)
    source = Path(source_path)
    updated: list[Path] = []
    for path in (
        source / "src" / "agribank_v3" / "__init__.py",
        source / "src" / "agribank_v3" / "version.py",
    ):
        if not path.is_file():
            continue
        text = path.read_text(encoding="utf-8")
        new_text, count = re.subn(
            r"(__version__\s*=\s*['\"])([^'\"]+)(['\"])",
            rf"\g<1>{new_version}\3",
            text,
            count=1,
        )
        if count:
            _backup_source_file(path)
            path.write_text(new_text, encoding="utf-8")
            updated.append(path)
    pyproject = source / "pyproject.toml"
    if pyproject.is_file():
        text = pyproject.read_text(encoding="utf-8")
        new_text, count = re.subn(
            r"(?m)^(version\s*=\s*['\"])([^'\"]+)(['\"])",
            rf"\g<1>{new_version}\3",
            text,
            count=1,
        )
        if count:
            _backup_source_file(pyproject)
            pyproject.write_text(new_text, encoding="utf-8")
            updated.append(pyproject)
    if not updated:
        raise UpdateBuilderError("Không tìm thấy file version để cập nhật.")
    return updated


def normalize_package_mode(value: str | None) -> str:
    text = str(value or PACKAGE_MODE_RUNTIME).strip().casefold()
    aliases = {
        "full": PACKAGE_MODE_SOURCE,
        "source_full": PACKAGE_MODE_SOURCE,
        "runtime_minimal": PACKAGE_MODE_RUNTIME,
        "minimum": PACKAGE_MODE_RUNTIME,
        "exe": PACKAGE_MODE_APP,
        "frozen": PACKAGE_MODE_APP,
        "pyinstaller": PACKAGE_MODE_APP,
        "app_root": PACKAGE_MODE_APP,
    }
    mode = aliases.get(text, text)
    if mode not in PACKAGE_MODES:
        raise UpdateBuilderError(f"Kiểu gói cập nhật không hợp lệ: {value}")
    return mode


def _payload_layout_for_package_mode(package_mode: str) -> str:
    return (
        PAYLOAD_LAYOUT_APP_ROOT
        if normalize_package_mode(package_mode) == PACKAGE_MODE_APP
        else PAYLOAD_LAYOUT_SOURCE
    )


def collect_package_files(
    source_path: Path | str,
    package_mode: str = PACKAGE_MODE_RUNTIME,
    excludes: Iterable[str] = DEFAULT_EXCLUDES,
    *,
    extra_excluded_paths: Iterable[Path | str] = (),
) -> tuple[list[CollectedFile], list[ExcludedPath]]:
    mode = normalize_package_mode(package_mode)
    if mode == PACKAGE_MODE_SOURCE:
        return collect_files(
            source_path,
            excludes,
            extra_excluded_paths=extra_excluded_paths,
        )
    if mode == PACKAGE_MODE_APP:
        app_excludes = tuple(
            item
            for item in excludes
            if str(item).casefold() not in {"dist", "*.exe"}
        )
        return collect_app_root_files(
            source_path,
            app_excludes,
            extra_excluded_paths=extra_excluded_paths,
        )
    return collect_runtime_files(
        source_path,
        excludes,
        extra_excluded_paths=extra_excluded_paths,
    )


def collect_runtime_files(
    source_path: Path | str,
    excludes: Iterable[str] = DEFAULT_EXCLUDES,
    *,
    extra_excluded_paths: Iterable[Path | str] = (),
) -> tuple[list[CollectedFile], list[ExcludedPath]]:
    source = Path(source_path).resolve()
    included: list[CollectedFile] = []
    excluded: list[ExcludedPath] = []
    exclude_tuple = tuple(excludes)
    extra_roots = tuple(Path(path).resolve() for path in extra_excluded_paths)

    for relative_text in RUNTIME_INCLUDE_DIRS:
        root = source / relative_text
        relative_root = Path(relative_text)
        if not root.exists():
            continue
        if not root.is_dir():
            reason = _exclude_reason(relative_root, False, exclude_tuple)
            if reason:
                excluded.append(ExcludedPath(root, relative_root, reason))
            else:
                included.append(
                    CollectedFile(root, relative_root, f"runtime file: {relative_text}")
                )
            continue
        for root_text, dirs, files in os.walk(root):
            walk_root = Path(root_text)
            kept_dirs: list[str] = []
            for dirname in sorted(dirs, key=str.casefold):
                directory = walk_root / dirname
                relative = directory.relative_to(source)
                reason = _exclude_reason(relative, True, exclude_tuple)
                if not reason and _is_within_any(directory, extra_roots):
                    reason = "path exclude: update/output folder"
                if reason:
                    excluded.append(ExcludedPath(directory, relative, reason))
                    continue
                kept_dirs.append(dirname)
            dirs[:] = kept_dirs

            for filename in sorted(files, key=str.casefold):
                path = walk_root / filename
                relative = path.relative_to(source)
                reason = _exclude_reason(relative, False, exclude_tuple)
                if not reason and _is_within_any(path, extra_roots):
                    reason = "path exclude: update/output file"
                if reason:
                    excluded.append(ExcludedPath(path, relative, reason))
                    continue
                included.append(
                    CollectedFile(path, relative, f"runtime dir: {relative_root.as_posix()}")
                )

    for relative_text in RUNTIME_INCLUDE_FILES:
        path = source / relative_text
        if not path.is_file():
            continue
        relative = Path(relative_text)
        reason = _exclude_reason(relative, False, exclude_tuple)
        if not reason and _is_within_any(path, extra_roots):
            reason = "path exclude: update/output file"
        if reason:
            excluded.append(ExcludedPath(path, relative, reason))
            continue
        included.append(CollectedFile(path, relative, "runtime file"))

    return _dedupe_collected_files(included), excluded


def collect_app_root_files(
    source_path: Path | str,
    excludes: Iterable[str] = APP_PACKAGE_EXCLUDES,
    *,
    extra_excluded_paths: Iterable[Path | str] = (),
) -> tuple[list[CollectedFile], list[ExcludedPath]]:
    source = Path(source_path).resolve()
    app_root = source / APP_BUILD_RELATIVE_ROOT
    if not app_root.is_dir():
        raise UpdateBuilderError(
            f"Không tìm thấy bản build app tại: {app_root}. "
            "Hãy build AgribankV3 trước khi tạo gói cập nhật cho bản exe."
        )
    if not (app_root / "AgribankV3.exe").is_file():
        raise UpdateBuilderError(
            f"Thư mục build không có AgribankV3.exe: {app_root}"
        )
    read_app_build_info(app_root)

    included: list[CollectedFile] = []
    excluded: list[ExcludedPath] = []
    exclude_tuple = tuple(excludes)
    extra_roots = tuple(Path(path).resolve() for path in extra_excluded_paths)
    for root_text, dirs, files in os.walk(app_root):
        root = Path(root_text)
        kept_dirs: list[str] = []
        for dirname in sorted(dirs, key=str.casefold):
            directory = root / dirname
            relative = directory.relative_to(app_root)
            reason = _exclude_reason(relative, True, exclude_tuple)
            if not reason and _is_within_any(directory, extra_roots):
                reason = "path exclude: update/output folder"
            if reason:
                excluded.append(ExcludedPath(directory, relative, reason))
                continue
            kept_dirs.append(dirname)
        dirs[:] = kept_dirs

        for filename in sorted(files, key=str.casefold):
            path = root / filename
            relative = path.relative_to(app_root)
            reason = _exclude_reason(relative, False, exclude_tuple)
            if not reason and _is_within_any(path, extra_roots):
                reason = "path exclude: update/output file"
            if reason:
                excluded.append(ExcludedPath(path, relative, reason))
                continue
            included.append(CollectedFile(path, relative, "app build file"))

    return _dedupe_collected_files(included), excluded


def create_package_plan(
    config: BuildUpdateConfig,
    previous_release_version: str | None = None,
) -> PackagePlan:
    mode = normalize_package_mode(config.package_mode)
    source = Path(config.source_path)
    base_version = previous_release_version or config.previous_release_version or ""
    extra_excluded_paths = _extra_excluded_paths_for_update(
        config.source_path,
        config.update_path,
    )
    collect_mode = PACKAGE_MODE_RUNTIME if mode == PACKAGE_MODE_DELTA else mode
    current_files, excluded = collect_package_files(
        source,
        collect_mode,
        config.excludes,
        extra_excluded_paths=extra_excluded_paths,
    )
    snapshot_files = tuple(current_files)
    delete_files: tuple[Path, ...] = ()
    included = current_files
    package_type = "full"
    if mode == PACKAGE_MODE_DELTA:
        if not base_version:
            raise UpdateBuilderError(
                "Gói delta cần xác định phiên bản nền trước đó."
            )
        previous_snapshot = load_release_file_snapshot(source, base_version)
        if previous_snapshot is None:
            raise UpdateBuilderError(
                f"Không tìm thấy snapshot file của phiên bản nền {base_version}. "
                "Hãy tạo một gói runtime tối thiểu trước khi tạo delta."
            )
        current_snapshot = file_snapshot_for_files(snapshot_files)
        included = [
            item
            for item in current_files
            if _snapshot_entry_changed(
                current_snapshot[item.relative_path.as_posix()],
                previous_snapshot.get(item.relative_path.as_posix()),
            )
        ]
        current_paths = {item.relative_path.as_posix() for item in current_files}
        delete_files = tuple(
            Path(path)
            for path in sorted(previous_snapshot)
            if path not in current_paths and _is_safe_delta_delete_path(Path(path))
        )
        package_type = "delta"
    report = package_file_report(included, excluded)
    payload_layout = _payload_layout_for_package_mode(mode)
    return PackagePlan(
        package_mode=mode,
        package_type=package_type,
        payload_layout=payload_layout,
        base_version=base_version if package_type == "delta" else "",
        included=tuple(included),
        snapshot_files=snapshot_files,
        excluded=tuple(excluded),
        delete_files=delete_files,
        report=report,
    )


def collect_files(
    source_path: Path | str,
    excludes: Iterable[str] = DEFAULT_EXCLUDES,
    *,
    extra_excluded_paths: Iterable[Path | str] = (),
) -> tuple[list[CollectedFile], list[ExcludedPath]]:
    source = Path(source_path).resolve()
    included: list[CollectedFile] = []
    excluded: list[ExcludedPath] = []
    exclude_tuple = tuple(excludes)
    extra_roots = tuple(Path(path).resolve() for path in extra_excluded_paths)
    for root_text, dirs, files in os.walk(source):
        root = Path(root_text)
        kept_dirs: list[str] = []
        for dirname in sorted(dirs, key=str.casefold):
            directory = root / dirname
            relative = directory.relative_to(source)
            reason = _exclude_reason(relative, True, exclude_tuple)
            if not reason and _is_within_any(directory, extra_roots):
                reason = "path exclude: update/output folder"
            if reason:
                excluded.append(ExcludedPath(directory, relative, reason))
                continue
            kept_dirs.append(dirname)
        dirs[:] = kept_dirs

        for filename in sorted(files, key=str.casefold):
            path = root / filename
            relative = path.relative_to(source)
            reason = _exclude_reason(relative, False, exclude_tuple)
            if not reason and _is_within_any(path, extra_roots):
                reason = "path exclude: update/output file"
            if reason:
                excluded.append(ExcludedPath(path, relative, reason))
                continue
            included.append(CollectedFile(path, relative, "source file"))
    return included, excluded


def create_update_zip(
    source_path: Path | str,
    output_zip: Path | str,
    excludes: Iterable[str] = DEFAULT_EXCLUDES,
    *,
    logger: Callable[[str], None] | None = None,
    progress_callback: Callable[[int], None] | None = None,
    precollected_files: Iterable[CollectedFile] | None = None,
    extra_excluded_paths: Iterable[Path | str] = (),
) -> Path:
    source = Path(source_path).resolve()
    output = Path(output_zip).resolve()
    log = logger or (lambda message: None)
    progress = progress_callback or (lambda value: None)
    output.parent.mkdir(parents=True, exist_ok=True)
    temporary_zip = output.with_name(f".{output.name}.{uuid4().hex}.tmp")
    extra_roots = [Path(path).resolve() for path in extra_excluded_paths]
    extra_roots.extend([output, temporary_zip])
    if precollected_files is None:
        included, _ = collect_files(
            source,
            excludes,
            extra_excluded_paths=extra_roots,
        )
    else:
        included = [
            item
            for item in precollected_files
            if not _is_within_any(item.path, tuple(extra_roots))
        ]
    try:
        with zipfile.ZipFile(
            temporary_zip,
            "w",
            compression=zipfile.ZIP_DEFLATED,
            compresslevel=9,
        ) as archive:
            total = len(included)
            last_percent = -1
            for index, item in enumerate(included, start=1):
                archive.write(
                    item.path,
                    item.relative_path.as_posix(),
                    compress_type=(
                        zipfile.ZIP_STORED
                        if _should_store_without_compression(item.relative_path)
                        else zipfile.ZIP_DEFLATED
                    ),
                )
                if total:
                    percent = int(index * 100 / total)
                    if percent != last_percent and (
                        percent % 10 == 0 or index == total
                    ):
                        progress(percent)
                        log(f"Đang đóng gói file {index}/{total}...")
                        last_percent = percent
        temporary_zip.replace(output)
    finally:
        temporary_zip.unlink(missing_ok=True)
    return output


def package_file_report(
    included: Iterable[CollectedFile],
    excluded: Iterable[ExcludedPath] = (),
    *,
    top_file_count: int = 20,
    top_directory_count: int = 10,
) -> PackageFileReport:
    files = tuple(included)
    excluded_tuple = tuple(excluded)
    sizes: list[FileSizeInfo] = []
    directory_sizes: dict[Path, int] = {}
    total_size = 0
    for item in files:
        try:
            size = item.path.stat().st_size
        except OSError:
            size = 0
        total_size += size
        sizes.append(FileSizeInfo(item.relative_path, size))
        parts = item.relative_path.parts
        if len(parts) > 1:
            for depth in range(1, len(parts)):
                directory = Path(*parts[:depth])
                directory_sizes[directory] = directory_sizes.get(directory, 0) + size
        else:
            directory_sizes[Path(".")] = directory_sizes.get(Path("."), 0) + size
    top_files = tuple(
        sorted(sizes, key=lambda item: item.size, reverse=True)[:top_file_count]
    )
    top_directories = tuple(
        DirectorySizeInfo(path, size)
        for path, size in sorted(
            directory_sizes.items(),
            key=lambda item: item[1],
            reverse=True,
        )[:top_directory_count]
    )
    return PackageFileReport(
        included_count=len(files),
        excluded_count=len(excluded_tuple),
        total_size=total_size,
        estimated_zip_size=max(1, int(total_size * 0.65)) if total_size else 0,
        top_files=top_files,
        top_directories=top_directories,
    )


def format_package_report(report: PackageFileReport) -> str:
    lines = [
        "Đang phân tích gói cập nhật...",
        f"Số file sẽ đóng gói: {report.included_count}",
        f"Dung lượng trước nén: {_format_bytes(report.total_size)}",
        f"Dự kiến dung lượng zip: {_format_bytes(report.estimated_zip_size)}",
        "Top file lớn:",
    ]
    if report.top_files:
        lines.extend(
            f"{index}. {item.relative_path.as_posix()} - {_format_bytes(item.size)}"
            for index, item in enumerate(report.top_files, start=1)
        )
    else:
        lines.append("(không có file)")
    lines.append("Top thư mục chiếm dung lượng:")
    if report.top_directories:
        lines.extend(
            f"{index}. {item.relative_path.as_posix()} - {_format_bytes(item.size)}"
            for index, item in enumerate(report.top_directories, start=1)
        )
    else:
        lines.append("(không có thư mục)")
    return "\n".join(lines)


def write_package_file_list(
    output_path: Path | str,
    files: Iterable[CollectedFile],
) -> Path:
    path = Path(output_path)
    path.parent.mkdir(parents=True, exist_ok=True)
    lines = ["relative_path\tsize_bytes\tinclude_reason"]
    for item in sorted(files, key=lambda file: file.relative_path.as_posix().casefold()):
        try:
            size = item.path.stat().st_size
        except OSError:
            size = 0
        lines.append(
            f"{item.relative_path.as_posix()}\t{size}\t{item.include_reason}"
        )
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")
    return path


def file_snapshot_for_files(
    files: Iterable[CollectedFile],
) -> dict[str, dict[str, object]]:
    snapshot: dict[str, dict[str, object]] = {}
    for item in files:
        try:
            stat = item.path.stat()
        except OSError:
            continue
        snapshot[item.relative_path.as_posix()] = {
            "relative_path": item.relative_path.as_posix(),
            "size": stat.st_size,
            "sha256": _sha256_file(item.path),
            "modified_time": stat.st_mtime,
        }
    return snapshot


def save_release_file_snapshot(
    source_path: Path | str,
    version: str,
    files: Iterable[CollectedFile],
    *,
    package_mode: str = PACKAGE_MODE_RUNTIME,
) -> Path:
    _parse_semver(version)
    snapshot_dir = _release_snapshot_dir(source_path)
    snapshot_dir.mkdir(parents=True, exist_ok=True)
    path = snapshot_dir / _release_file_snapshot_name(version, package_mode)
    payload = {
        "version": version,
        "package_mode": normalize_package_mode(package_mode),
        "created_at": _now(),
        "files": file_snapshot_for_files(files),
    }
    path.write_text(
        json.dumps(payload, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )
    return path


def load_release_file_snapshot(
    source_path: Path | str,
    version: str,
    *,
    package_mode: str = PACKAGE_MODE_RUNTIME,
) -> dict[str, dict[str, object]] | None:
    _parse_semver(version)
    path = _release_snapshot_dir(source_path) / _release_file_snapshot_name(
        version,
        package_mode,
    )
    if not path.is_file():
        return None
    try:
        raw = json.loads(path.read_text(encoding="utf-8-sig"))
    except Exception as exc:
        raise UpdateBuilderError(f"Không đọc được file snapshot {path}: {exc}") from exc
    files = raw.get("files") if isinstance(raw, dict) else None
    if not isinstance(files, dict):
        raise UpdateBuilderError(f"File snapshot không đúng cấu trúc: {path}")
    return {
        str(key): value
        for key, value in files.items()
        if isinstance(value, dict)
    }


def build_manifest(
    config: BuildUpdateConfig,
    package: str,
    *,
    package_type: str | None = None,
    payload_layout: str | None = None,
    base_version: str = "",
    delete_files: Iterable[Path | str] = (),
) -> dict[str, object]:
    migrations = [
        ManifestMigration(
            version=item.version,
            file=item.manifest_file,
            description=item.description,
        )
        for item in config.migrations
    ]
    return manifest_dict(
        latest_version=config.new_version,
        package=package,
        release_date=config.release_date,
        required_app_restart=config.required_app_restart,
        notes=list(config.notes),
        migrations=migrations,
        package_type=package_type
        or ("delta" if normalize_package_mode(config.package_mode) == PACKAGE_MODE_DELTA else "full"),
        payload_layout=payload_layout
        or _payload_layout_for_package_mode(config.package_mode),
        base_version=base_version,
        delete_files=[Path(path).as_posix() for path in delete_files],
    )


def write_manifest(update_path: Path | str, manifest: dict[str, object]) -> Path:
    path = Path(update_path) / "manifest.json"
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(
        json.dumps(manifest, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )
    return path


def copy_migrations(config: BuildUpdateConfig) -> list[Path]:
    copied: list[Path] = []
    if not config.migrations:
        return copied
    migrations_dir = Path(config.update_path) / "migrations"
    migrations_dir.mkdir(parents=True, exist_ok=True)
    for migration in config.migrations:
        if migration.use_python_migration or migration.source_file is None:
            continue
        target = migrations_dir / migration.source_file.name
        if migration.source_file.resolve() != target.resolve():
            shutil.copy2(migration.source_file, target)
        copied.append(target)
    return copied


def compare_databases_for_migration(
    old_db_path: Path | str,
    new_db_path: Path | str,
) -> SchemaDiff:
    return compare_sqlite_schema(old_db_path, new_db_path)


def create_suggested_python_migration(
    *,
    old_db_path: Path | str,
    new_db_path: Path | str,
    version: str,
    output_dir: Path | str | None = None,
) -> tuple[Path, SchemaDiff]:
    _parse_semver(version)
    diff = compare_sqlite_schema(old_db_path, new_db_path)
    path = write_generated_python_migration(diff, version, output_dir)
    return path, diff


def test_generated_python_migration(
    *,
    old_db_path: Path | str,
    new_db_path: Path | str,
    migration_file: Path | str,
    version: str,
):
    return validate_generated_migration_on_copy(
        old_db_path=old_db_path,
        new_db_path=new_db_path,
        migration_file=migration_file,
        version=version,
    )


def find_dev_database(source_path: Path | str) -> Path | None:
    source = Path(source_path)
    candidates = (
        source / "data" / "DuLieuV3.db",
        source / "src" / "agribank_v3" / "data" / "DuLieuV3.db",
        source / "DuLieuV3.db",
    )
    for path in candidates:
        if path.is_file():
            return path
    matches = sorted(
        (
            path
            for path in source.rglob("DuLieuV3.db")
            if path.is_file()
            and not any(part in {".git", ".venv", "build", "dist"} for part in path.parts)
        ),
        key=lambda item: len(item.relative_to(source).parts),
    )
    return matches[0] if matches else None


def save_schema_snapshot(
    db_path: Path | str,
    version: str,
    snapshot_dir: Path | str | None = None,
) -> Path:
    _parse_semver(version)
    schema = read_sqlite_schema(db_path)
    directory = Path(snapshot_dir) if snapshot_dir is not None else _schema_snapshot_dir()
    directory.mkdir(parents=True, exist_ok=True)
    payload = _schema_snapshot_payload(schema, version, Path(db_path))
    payload["checksum"] = _schema_snapshot_checksum(payload)
    path = directory / f"schema_{version}.json"
    path.write_text(
        json.dumps(payload, ensure_ascii=False, indent=2, sort_keys=True) + "\n",
        encoding="utf-8",
    )
    return path


def find_latest_schema_snapshot(
    before_version: str,
    snapshot_dir: Path | str | None = None,
) -> SchemaSnapshotInfo | None:
    _parse_semver(before_version)
    directory = Path(snapshot_dir) if snapshot_dir is not None else _schema_snapshot_dir()
    if not directory.is_dir():
        return None
    candidates: list[SchemaSnapshotInfo] = []
    for path in directory.glob("schema_*.json"):
        version = path.stem.removeprefix("schema_")
        try:
            if compare_versions(version, before_version) < 0:
                candidates.append(SchemaSnapshotInfo(version=version, path=path))
        except UpdateBuilderError:
            continue
    if not candidates:
        return None
    return max(candidates, key=lambda item: _parse_semver(item.version))


def detect_previous_release_version(
    update_path: Path | str,
    source_path: Path | str | None = None,
    new_version: str | None = None,
    snapshot_dir: Path | str | None = None,
) -> str | None:
    manifest_version = _read_manifest_version(Path(update_path) / "manifest.json")
    if manifest_version:
        return manifest_version
    snapshot_before = new_version or "9999.9999.9999"
    snapshot = find_latest_schema_snapshot(snapshot_before, snapshot_dir)
    if snapshot is not None:
        return snapshot.version
    config_version = _read_build_config_last_release_version(source_path)
    if config_version:
        return config_version
    return None


def auto_detect_database_changes(
    source_path: Path | str,
    update_path: Path | str,
    current_version: str,
    new_version: str,
    *,
    dev_db_path: Path | str | None = None,
    baseline_db_path: Path | str | None = None,
    snapshot_dir: Path | str | None = None,
) -> AutoDatabaseDetectionResult:
    del update_path
    _parse_semver(current_version)
    _parse_semver(new_version)
    dev_db = Path(dev_db_path) if dev_db_path else find_dev_database(source_path)
    if dev_db is None or not dev_db.is_file():
        return AutoDatabaseDetectionResult(
            status="action_required",
            current_version=current_version,
            new_version=new_version,
            dev_db_path=None,
            snapshot_path=None,
            baseline_db_path=Path(baseline_db_path) if baseline_db_path else None,
            diff=None,
            message="Không tìm thấy database dev để so sánh.",
        )
    if baseline_db_path:
        baseline_db = Path(baseline_db_path)
        if not baseline_db.is_file():
            return AutoDatabaseDetectionResult(
                status="action_required",
                current_version=current_version,
                new_version=new_version,
                dev_db_path=dev_db,
                snapshot_path=None,
                baseline_db_path=baseline_db,
                diff=None,
                message=f"Không tìm thấy database baseline: {baseline_db}",
            )
        diff = compare_sqlite_schema(baseline_db, dev_db)
        return _database_detection_from_diff(
            current_version=current_version,
            new_version=new_version,
            dev_db_path=dev_db,
            snapshot_path=None,
            baseline_db_path=baseline_db,
            diff=diff,
        )
    snapshot = find_latest_schema_snapshot(new_version, snapshot_dir)
    if snapshot is None:
        return AutoDatabaseDetectionResult(
            status="action_required",
            current_version=current_version,
            new_version=new_version,
            dev_db_path=dev_db,
            snapshot_path=None,
            baseline_db_path=None,
            diff=None,
            message="Chưa có schema snapshot bản trước.",
        )
    with tempfile.TemporaryDirectory(prefix="agribankv3-schema-snapshot-") as temp_dir:
        baseline_db = _materialize_schema_snapshot(snapshot.path, Path(temp_dir))
        diff = compare_sqlite_schema(baseline_db, dev_db)
    return _database_detection_from_diff(
        current_version=current_version,
        new_version=new_version,
        dev_db_path=dev_db,
        snapshot_path=snapshot.path,
        baseline_db_path=None,
        diff=diff,
    )


def auto_build_update(
    config: AutoBuildConfig,
    logger: Callable[[str], None] | None = None,
    progress_callback: Callable[[int], None] | None = None,
) -> AutoBuildResult:
    log = logger or (lambda message: None)
    progress = progress_callback or (lambda value: None)
    source_path = Path(config.source_path)
    update_path = Path(config.update_path)
    log("Đang đọc source...")
    current_version = read_current_version(source_path)
    previous_release_version = config.previous_release_version
    if previous_release_version is None:
        previous_release_version = detect_previous_release_version(
            update_path,
            source_path,
            config.new_version,
            config.snapshot_dir,
        )
    log(f"Phiên bản trong source: {current_version}")
    log(
        "Phiên bản phát hành trước: "
        + (previous_release_version or "không xác định")
    )
    log(f"Phiên bản gói cập nhật sẽ tạo: {config.new_version}")
    log("Đang tìm database dev...")
    dev_db = Path(config.dev_db_path) if config.dev_db_path else find_dev_database(source_path)
    if dev_db is not None and dev_db.is_file():
        log(f"Database dev: {dev_db}")
    else:
        log(
            "Không tìm thấy database dev để so sánh. Bản cập nhật sẽ được coi là "
            "chỉ đổi code nếu người dùng xác nhận."
        )

    generated_migration_path: Path | None = None
    inserted_backup_path: Path | None = None
    migration_test_database_path: Path | None = None
    saved_snapshot_path: Path | None = None

    detection = auto_detect_database_changes(
        source_path,
        update_path,
        current_version,
        config.new_version,
        dev_db_path=dev_db,
        baseline_db_path=config.baseline_db_path,
        snapshot_dir=config.snapshot_dir,
    )
    if detection.snapshot_path:
        log(f"Snapshot bản trước: {detection.snapshot_path}")
    elif config.baseline_db_path:
        log(f"Baseline database: {config.baseline_db_path}")
    else:
        log("Snapshot bản trước: chưa có")

    migrations: tuple[MigrationItem, ...] = ()
    database_changed = False

    if detection.status == "action_required":
        if not config.code_only_if_missing_baseline:
            raise UpdateBuilderError(
                "Không đủ dữ liệu để tự kiểm tra thay đổi database. "
                "Vui lòng chọn database baseline hoặc tạo update chỉ đổi code."
            )
        log("Người dùng chọn tạo update chỉ đổi code, không kèm migration database.")
    elif detection.status == "dangerous_changes":
        if detection.diff is not None:
            for change in detection.diff.dangerous_changes:
                log(f"Thay đổi nguy hiểm: {change.kind} {change.name} - {change.detail}")
        raise UpdateBuilderError(
            "Phát hiện thay đổi database cần xử lý thủ công. Không thể tạo update tự động."
        )
    elif detection.status == "no_changes":
        log("Không phát hiện thay đổi database.")
    elif detection.status == "safe_changes":
        log("Có thay đổi database an toàn.")
        if detection.diff is None or dev_db is None:
            raise UpdateBuilderError("Không có dữ liệu diff để tạo migration tự động.")
        database_changed = True
        log("Đang tạo migration...")
        generated_migration_path = write_generated_python_migration(
            detection.diff,
            config.new_version,
            config.generated_migration_dir
            or _generated_migration_dir_for_source(source_path),
        )
        log(f"Đã tạo migration: {generated_migration_path}")
        log("Đang chèn migration vào db_migrations.py...")
        inserted_backup_path = insert_python_migration_into_source(
            source_path=source_path,
            generated_migration_file=generated_migration_path,
            version=config.new_version,
        )
        log(f"Đã backup db_migrations.py: {inserted_backup_path}")
        if not source_has_python_migration(source_path, config.new_version):
            raise UpdateBuilderError(
                f"Không xác nhận được {migration_function_name(config.new_version)} "
                "trong db_migrations.py sau khi chèn."
            )
        log("Đang thử migration trên bản copy...")
        with tempfile.TemporaryDirectory(prefix="agribankv3-auto-build-") as temp_dir:
            if config.baseline_db_path:
                baseline_db = Path(config.baseline_db_path)
            elif detection.snapshot_path:
                baseline_db = _materialize_schema_snapshot(detection.snapshot_path, Path(temp_dir))
            else:
                raise UpdateBuilderError("Không có baseline để test migration.")
            validation = validate_generated_migration_on_copy(
                old_db_path=baseline_db,
                new_db_path=dev_db,
                migration_file=generated_migration_path,
                version=config.new_version,
            )
        migration_test_database_path = validation.test_database_path
        if not validation.success:
            detail = validation.error or ", ".join(validation.missing_items)
            raise UpdateBuilderError(f"Test migration tự động thất bại: {detail}")
        log("Migration OK.")
        migrations = (
            MigrationItem(
                version=config.new_version,
                description=f"Migration database tự động cho phiên bản {config.new_version}",
                use_python_migration=True,
            ),
        )
    else:
        raise UpdateBuilderError(f"Trạng thái auto detect không hợp lệ: {detection.status}")

    if config.baseline_db_path:
        save_schema_snapshot(
            config.baseline_db_path,
            current_version,
            config.snapshot_dir,
        )

    build_config = BuildUpdateConfig(
        source_path=source_path,
        update_path=update_path,
        new_version=config.new_version,
        release_date=config.release_date,
        required_app_restart=config.required_app_restart,
        notes=config.notes,
        migrations=migrations,
        auto_update_source_version=config.auto_update_source_version,
        database_changed=database_changed,
        previous_release_version=previous_release_version,
        allow_rebuild_same_version=config.allow_rebuild_same_version,
        package_mode=config.package_mode,
        verbose=config.verbose,
    )
    log("Đang tạo zip...")
    log("Đang tạo manifest...")
    build_result = UpdateBuilder(log, progress).build(build_config)
    log("Đang lưu schema snapshot mới...")
    if dev_db is not None and dev_db.is_file():
        saved_snapshot_path = save_schema_snapshot(
            dev_db,
            config.new_version,
            config.snapshot_dir,
        )
        log(f"Đã lưu schema snapshot mới: {saved_snapshot_path}")
    else:
        log("Không lưu schema snapshot mới vì không có database dev.")
    log("Hoàn thành.")
    return AutoBuildResult(
        build_result=build_result,
        detection=detection,
        generated_migration_path=generated_migration_path,
        inserted_migration_backup_path=inserted_backup_path,
        migration_test_database_path=migration_test_database_path,
        saved_snapshot_path=saved_snapshot_path,
    )


def source_has_python_migration(source_path: Path | str, version: str) -> bool:
    migration_path = (
        Path(source_path)
        / "src"
        / "agribank_v3"
        / "update"
        / "db_migrations.py"
    )
    if not migration_path.is_file():
        return False
    text = migration_path.read_text(encoding="utf-8")
    function_name = migration_function_name(version)
    return (
        re.search(rf"^def\s+{re.escape(function_name)}\s*\(", text, re.MULTILINE)
        is not None
        and f'"{version}"' in text
    )


def insert_python_migration_into_source(
    *,
    source_path: Path | str,
    generated_migration_file: Path | str,
    version: str,
) -> Path:
    migration_path = (
        Path(source_path)
        / "src"
        / "agribank_v3"
        / "update"
        / "db_migrations.py"
    )
    if not migration_path.is_file():
        raise UpdateBuilderError(f"Không tìm thấy db_migrations.py: {migration_path}")
    text = migration_path.read_text(encoding="utf-8")
    function_name = migration_function_name(version)
    if re.search(rf"^def\s+{re.escape(function_name)}\s*\(", text, re.MULTILINE):
        raise UpdateBuilderError(f"{function_name} đã tồn tại trong db_migrations.py.")
    generated_text = Path(generated_migration_file).read_text(encoding="utf-8")
    match = re.search(
        rf"(^def\s+{re.escape(function_name)}\s*\(conn\):\n(?:^[ \t].*\n|^\s*$)+)",
        generated_text,
        flags=re.MULTILINE,
    )
    if not match:
        raise UpdateBuilderError(
            f"Không tìm thấy function {function_name} trong file migration gợi ý."
        )
    if "def default_python_migrations()" not in text or "return {" not in text:
        raise UpdateBuilderError(
            "Không tìm được cấu trúc default_python_migrations() để đăng ký migration."
        )
    backup = migration_path.with_name(
        f"db_migrations.py.bak_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
    )
    shutil.copy2(migration_path, backup)
    function_block = "\n\n" + match.group(1).rstrip() + "\n"
    updated = text.rstrip() + function_block + "\n"
    updated = re.sub(
        r"(def\s+default_python_migrations\(\)\s*->\s*dict\[str,\s*PythonMigration\]:\n\s*return\s*\{\n)",
        rf'\1        "{version}": {function_name},\n',
        updated,
        count=1,
    )
    if updated == text.rstrip() + function_block + "\n":
        raise UpdateBuilderError("Không đăng ký được migration vào default_python_migrations().")
    migration_path.write_text(updated, encoding="utf-8")
    return backup


def backup_existing_update_files(
    update_path: Path | str,
    package_name_value: str,
) -> Path | None:
    update_root = Path(update_path)
    files = [
        path
        for path in (
            update_root / "manifest.json",
            update_root / package_name_value,
        )
        if path.exists()
    ]
    if not files:
        return None
    backup_root = update_root / "_backup" / datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_root.mkdir(parents=True, exist_ok=True)
    for file_path in files:
        shutil.copy2(file_path, backup_root / file_path.name)
    return backup_root


def warn_dangerous_sql(sql_path: Path | str) -> list[str]:
    path = Path(sql_path)
    text = path.read_text(encoding="utf-8-sig", errors="replace")
    warnings = []
    for pattern in DANGEROUS_SQL_PATTERNS:
        if re.search(pattern, text, flags=re.IGNORECASE):
            warnings.append(
                f"Cảnh báo SQL nguy hiểm trong {path.name}: {pattern}"
            )
    return warnings


def package_name(version: str) -> str:
    _parse_semver(version)
    return f"AgribankV3_{version}.zip"


def compare_versions(current: str, latest: str) -> int:
    current_parts = _parse_semver(current)
    latest_parts = _parse_semver(latest)
    if current_parts < latest_parts:
        return -1
    if current_parts > latest_parts:
        return 1
    return 0


def write_builder_log(entries: dict[str, object]) -> Path:
    log_dir = _builder_output_root() / "logs"
    log_dir.mkdir(parents=True, exist_ok=True)
    log_path = log_dir / "update_builder.log"
    with log_path.open("a", encoding="utf-8") as stream:
        stream.write(json.dumps(entries, ensure_ascii=False, default=str))
        stream.write("\n")
    return log_path


def config_from_json(path: Path | str) -> BuildUpdateConfig:
    raw = json.loads(Path(path).read_text(encoding="utf-8-sig"))
    migrations = []
    for item in raw.get("migrations", []) or []:
        migrations.append(
            MigrationItem(
                version=str(item.get("version", "")),
                description=str(item.get("description", "")),
                source_file=(
                    Path(item["source_file"])
                    if item.get("source_file")
                    else None
                ),
                use_python_migration=bool(item.get("use_python_migration", False)),
            )
        )
    return BuildUpdateConfig(
        source_path=Path(raw["source_path"]),
        update_path=Path(raw.get("update_path") or DEFAULT_UPDATE_PATH),
        new_version=str(raw["new_version"]),
        release_date=str(raw.get("release_date") or date.today().isoformat()),
        required_app_restart=bool(raw.get("required_app_restart", True)),
        notes=tuple(str(note) for note in raw.get("notes", [])),
        migrations=tuple(migrations),
        auto_update_source_version=bool(raw.get("auto_update_source_version", False)),
        previous_release_version=(
            str(raw["previous_release_version"])
            if raw.get("previous_release_version")
            else str(raw["last_release_version"])
            if raw.get("last_release_version")
            else None
        ),
        allow_rebuild_same_version=bool(raw.get("allow_rebuild_same_version", False)),
        package_mode=normalize_package_mode(raw.get("package_mode")),
        verbose=bool(raw.get("verbose", False)),
    )


def _database_detection_from_diff(
    *,
    current_version: str,
    new_version: str,
    dev_db_path: Path,
    snapshot_path: Path | None,
    baseline_db_path: Path | None,
    diff: SchemaDiff,
) -> AutoDatabaseDetectionResult:
    if diff.dangerous_changes:
        status = "dangerous_changes"
        message = "Phát hiện thay đổi database nguy hiểm."
    elif diff.has_safe_changes:
        status = "safe_changes"
        message = "Phát hiện thay đổi database an toàn."
    else:
        status = "no_changes"
        message = "Không phát hiện thay đổi database."
    return AutoDatabaseDetectionResult(
        status=status,
        current_version=current_version,
        new_version=new_version,
        dev_db_path=dev_db_path,
        snapshot_path=snapshot_path,
        baseline_db_path=baseline_db_path,
        diff=diff,
        message=message,
    )


def _generated_migration_dir_for_source(source_path: Path | str) -> Path:
    return Path(source_path) / "tools" / "update_builder" / "generated_migrations"


def _release_snapshot_dir(source_path: Path | str) -> Path:
    return Path(source_path) / "tools" / "update_builder" / "release_snapshots"


def _release_file_snapshot_name(version: str, package_mode: str) -> str:
    mode = normalize_package_mode(package_mode)
    suffix = "_app" if mode == PACKAGE_MODE_APP else ""
    return f"files_{version}{suffix}.json"


def _extra_excluded_paths_for_update(
    source_path: Path | str,
    update_path: Path | str,
) -> tuple[Path, ...]:
    source = Path(source_path).resolve()
    update = Path(update_path).resolve()
    if _is_relative_to(update, source):
        return (update,)
    return ()


def _is_within_any(path: Path, roots: Iterable[Path]) -> bool:
    resolved = Path(path).resolve()
    return any(_is_relative_to(resolved, root) for root in roots)


def _is_relative_to(path: Path, root: Path) -> bool:
    try:
        Path(path).resolve().relative_to(Path(root).resolve())
    except ValueError:
        return False
    return True


def _dedupe_collected_files(files: Iterable[CollectedFile]) -> list[CollectedFile]:
    result: list[CollectedFile] = []
    seen: set[str] = set()
    for item in files:
        key = item.relative_path.as_posix().casefold()
        if key in seen:
            continue
        seen.add(key)
        result.append(item)
    return result


def _snapshot_entry_changed(
    current: dict[str, object],
    previous: dict[str, object] | None,
) -> bool:
    if previous is None:
        return True
    return str(current.get("sha256") or "") != str(previous.get("sha256") or "")


def _sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _format_bytes(value: int) -> str:
    size = float(value)
    for unit in ("B", "KB", "MB", "GB"):
        if size < 1024 or unit == "GB":
            if unit == "B":
                return f"{int(size)} {unit}"
            return f"{size:.2f} {unit}"
        size /= 1024
    return f"{size:.2f} GB"


def _should_store_without_compression(relative: Path) -> bool:
    return relative.suffix.casefold() in ALREADY_COMPRESSED_SUFFIXES


def _is_safe_delta_delete_path(relative: Path) -> bool:
    if relative.is_absolute():
        return False
    parts = [part.casefold() for part in relative.parts]
    if not parts or any(part in {"", ".", ".."} for part in parts):
        return False
    blocked_dirs = {
        "data",
        "logs",
        "backups",
        "backup",
        "temp",
        "ketqua",
        "outputs",
        "update",
    }
    if any(part in blocked_dirs for part in parts[:-1]):
        return False
    if _is_database_like(relative):
        return False
    return True


def _read_manifest_version(manifest_path: Path) -> str | None:
    if not manifest_path.is_file():
        return None
    try:
        raw = json.loads(manifest_path.read_text(encoding="utf-8-sig"))
    except Exception:
        return None
    version = raw.get("latest_version") if isinstance(raw, dict) else None
    if not version:
        return None
    text = str(version).strip()
    try:
        _parse_semver(text)
    except UpdateBuilderError:
        return None
    return text


def _read_build_config_last_release_version(source_path: Path | str | None) -> str | None:
    candidates = [_builder_output_root() / "build_config.json"]
    if source_path is not None:
        candidates.append(Path(source_path) / "tools" / "update_builder" / "build_config.json")
        candidates.append(Path(source_path) / "build_config.json")
    for path in candidates:
        if not path.is_file():
            continue
        try:
            raw = json.loads(path.read_text(encoding="utf-8-sig"))
        except Exception:
            continue
        if not isinstance(raw, dict):
            continue
        version = raw.get("last_release_version") or raw.get("previous_release_version")
        if not version:
            continue
        text = str(version).strip()
        try:
            _parse_semver(text)
        except UpdateBuilderError:
            continue
        return text
    return None


def _schema_snapshot_dir() -> Path:
    return _builder_output_root() / "schema_snapshots"


def _schema_snapshot_payload(
    schema: SQLiteSchema,
    version: str,
    source_database: Path,
) -> dict[str, object]:
    tables = {
        name: {
            "create_sql": table.create_sql,
            "columns": {
                column_name: {
                    "data_type": column.data_type,
                    "not_null": column.not_null,
                    "default_value": column.default_value,
                    "primary_key": column.primary_key,
                }
                for column_name, column in sorted(table.columns.items())
            },
        }
        for name, table in sorted(schema.tables.items())
    }
    indexes = {
        name: {
            "table": index.table,
            "unique": index.unique,
            "columns": list(index.columns),
            "create_sql": index.create_sql,
        }
        for name, index in sorted(schema.indexes.items())
    }
    app_preferences = {
        key: {
            "value": value,
            "updated_at": updated_at,
        }
        for key, (value, updated_at) in sorted(schema.app_preferences.items())
    }
    return {
        "version": version,
        "created_at": _now(),
        "source_database": str(source_database),
        "tables": tables,
        "indexes": indexes,
        "app_preferences": app_preferences,
    }


def _schema_snapshot_checksum(payload: dict[str, object]) -> str:
    clean_payload = dict(payload)
    clean_payload.pop("checksum", None)
    clean_payload.pop("created_at", None)
    clean_payload.pop("source_database", None)
    raw = json.dumps(clean_payload, ensure_ascii=False, sort_keys=True).encode("utf-8")
    return hashlib.sha256(raw).hexdigest()


def _materialize_schema_snapshot(snapshot_path: Path | str, output_dir: Path) -> Path:
    snapshot = json.loads(Path(snapshot_path).read_text(encoding="utf-8-sig"))
    database_path = output_dir / f"{Path(snapshot_path).stem}.db"
    database_path.parent.mkdir(parents=True, exist_ok=True)
    if database_path.exists():
        database_path.unlink()
    with closing(sqlite3.connect(database_path)) as connection:
        with connection:
            for _name, table in sorted((snapshot.get("tables") or {}).items()):
                create_sql = str((table or {}).get("create_sql") or "").strip()
                if create_sql:
                    connection.execute(create_sql)
            for _name, index in sorted((snapshot.get("indexes") or {}).items()):
                create_sql = str((index or {}).get("create_sql") or "").strip()
                if create_sql:
                    connection.execute(create_sql)
            _insert_snapshot_app_preferences(
                connection,
                snapshot.get("app_preferences") or {},
            )
    return database_path


def _insert_snapshot_app_preferences(
    connection: sqlite3.Connection,
    preferences: object,
) -> None:
    table_exists = connection.execute(
        "SELECT 1 FROM sqlite_master WHERE type = 'table' AND name = 'app_preferences'"
    ).fetchone()
    if table_exists is None or not isinstance(preferences, dict):
        return
    columns = {
        str(row[1])
        for row in connection.execute("PRAGMA table_info(app_preferences)").fetchall()
    }
    if not {"key", "value"}.issubset(columns):
        return
    has_updated_at = "updated_at" in columns
    for key, raw_value in sorted(preferences.items()):
        if not isinstance(raw_value, dict):
            continue
        if has_updated_at:
            connection.execute(
                """
                INSERT OR IGNORE INTO app_preferences(key, value, updated_at)
                VALUES (?, ?, ?)
                """,
                (
                    str(key),
                    str(raw_value.get("value", "")),
                    str(raw_value.get("updated_at", "1970-01-01T00:00:00")),
                ),
            )
        else:
            connection.execute(
                "INSERT OR IGNORE INTO app_preferences(key, value) VALUES (?, ?)",
                (str(key), str(raw_value.get("value", ""))),
            )


def _exclude_reason(relative: Path, is_dir: bool, excludes: tuple[str, ...]) -> str:
    normalized = relative.as_posix()
    name = relative.name
    parts = relative.parts
    for pattern in excludes:
        normalized_pattern = pattern.replace("\\", "/")
        if "/" not in normalized_pattern and "*" not in normalized_pattern:
            if name.casefold() == normalized_pattern.casefold() or any(
                part.casefold() == normalized_pattern.casefold()
                for part in parts
            ):
                return f"folder/name exclude: {pattern}"
        if fnmatch.fnmatch(normalized.casefold(), normalized_pattern.casefold()):
            return f"pattern exclude: {pattern}"
        if fnmatch.fnmatch(name.casefold(), normalized_pattern.casefold()):
            return f"name pattern exclude: {pattern}"
    if not is_dir and _is_database_like(relative):
        return "database file exclude"
    return ""


def _is_database_like(relative: Path) -> bool:
    return relative.name.casefold() in {"dulieuv3.db", "quiz.db"} or (
        relative.suffix.casefold()
        in {".db", ".sqlite", ".sqlite3", ".mdb", ".accdb"}
    )


def _parse_semver(version: str) -> tuple[int, int, int, tuple[str, ...]]:
    text = str(version).strip()
    match = re.fullmatch(
        r"v?(\d+)\.(\d+)\.(\d+)(?:[-+]([0-9A-Za-z.-]+))?",
        text,
    )
    if not match:
        raise UpdateBuilderError(f"Version không đúng semantic version: {version}")
    major, minor, patch = (int(match.group(index)) for index in (1, 2, 3))
    suffix = tuple((match.group(4) or "").split(".")) if match.group(4) else ()
    return major, minor, patch, suffix


def _parse_date(value: str) -> None:
    try:
        date.fromisoformat(value)
    except ValueError as exc:
        raise UpdateBuilderError("Release date phải có dạng yyyy-mm-dd.") from exc


def _backup_source_file(path: Path) -> None:
    backup = path.with_name(f"{path.name}.bak")
    shutil.copy2(path, backup)


def _now() -> str:
    return datetime.now().astimezone().isoformat(timespec="seconds")


def _builder_output_root() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent
