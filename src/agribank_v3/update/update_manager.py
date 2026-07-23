from __future__ import annotations

from contextlib import closing
from dataclasses import dataclass
from datetime import datetime
from importlib import util
import json
import os
from pathlib import Path
import re
import shutil
import sqlite3
import sys
from uuid import uuid4
import zipfile

from agribank_v3 import __version__
from agribank_v3.runtime_paths import application_root
from agribank_v3.settings import AppSettingsDatabase, SettingsDatabaseError
from agribank_v3.update.db_migrations import (
    AppliedMigration,
    DatabaseMigrationError,
    apply_migrations,
    ensure_schema_migrations_table,
    latest_schema_version,
)
from agribank_v3.update.update_manifest import (
    UpdateManifest,
    UpdateManifestError,
    read_update_manifest,
)


DEFAULT_UPDATE_PATH = (
    r"X:\public\5491\06- PHONG KHKD\NAM\AgribankV3\Update"
)
UPDATE_PATH_PREFERENCE_KEY = "update_path"
APP_VERSION_PREFERENCE_KEY = "last_successful_update_version"
DATABASE_FILE_NAMES = {
    "dulieuv3.db",
    "quiz.db",
    "agribank_v3.sqlite3",
}
DATABASE_SUFFIXES = {".db", ".sqlite", ".sqlite3", ".mdb", ".accdb"}
SKIPPED_DIRECTORIES = {
    ".git",
    ".venv",
    "__pycache__",
    "backups",
    "backup",
    "logs",
    "temp",
    "ketqua",
}
PAYLOAD_LAYOUT_AUTO = "auto"
PAYLOAD_LAYOUT_SOURCE = "source"
PAYLOAD_LAYOUT_APP_ROOT = "app_root"
APP_BUILD_INFO_FILE_NAME = "agribank_v3_build_info.json"


class UpdateError(RuntimeError):
    pass


@dataclass(frozen=True, slots=True)
class UpdateSettings:
    update_path: Path


@dataclass(frozen=True, slots=True)
class UpdateCheckResult:
    status: str
    current_version: str
    latest_version: str = ""
    manifest: UpdateManifest | None = None
    message: str = ""

    @property
    def update_available(self) -> bool:
        return self.status == "update_available"


@dataclass(frozen=True, slots=True)
class PreparedUpdate:
    manifest: UpdateManifest
    package_path: Path
    copied_package_path: Path
    staging_root: Path
    payload_root: Path
    payload_layout: str
    skipped_payload_files: tuple[Path, ...]


@dataclass(frozen=True, slots=True)
class UpdateApplyResult:
    old_version: str
    new_version: str
    backup_path: Path
    log_path: Path
    copied_files: tuple[Path, ...]
    skipped_files: tuple[Path, ...]
    applied_migrations: tuple[AppliedMigration, ...]
    restart_required: bool
    deleted_files: tuple[Path, ...] = ()
    updater_script: Path | None = None


def get_current_version() -> str:
    return __version__


def load_update_settings(
    settings_database: AppSettingsDatabase | None = None,
) -> UpdateSettings:
    database = settings_database or AppSettingsDatabase()
    value = database.load_preference(
        UPDATE_PATH_PREFERENCE_KEY,
        DEFAULT_UPDATE_PATH,
    )
    return UpdateSettings(update_path=Path(value or DEFAULT_UPDATE_PATH))


def save_update_settings(
    update_path: Path | str,
    settings_database: AppSettingsDatabase | None = None,
) -> UpdateSettings:
    database = settings_database or AppSettingsDatabase()
    saved = database.save_preference(UPDATE_PATH_PREFERENCE_KEY, str(update_path))
    return UpdateSettings(update_path=Path(saved))


def compare_versions(current: str, latest: str) -> int:
    current_parts = _parse_semver(current)
    latest_parts = _parse_semver(latest)
    if current_parts < latest_parts:
        return -1
    if current_parts > latest_parts:
        return 1
    return 0


def check_for_update(
    *,
    update_path: Path | str | None = None,
    current_version: str | None = None,
    settings_database: AppSettingsDatabase | None = None,
) -> UpdateCheckResult:
    version = current_version or get_current_version()
    root = Path(update_path) if update_path is not None else load_update_settings(
        settings_database
    ).update_path
    if not root.is_dir():
        return UpdateCheckResult(
            status="missing_update_directory",
            current_version=version,
            message="Không tìm thấy thư mục cập nhật",
        )
    try:
        manifest = read_update_manifest(root)
        comparison = compare_versions(version, manifest.latest_version)
    except (UpdateManifestError, ValueError) as exc:
        return UpdateCheckResult(
            status="manifest_error",
            current_version=version,
            message=f"Lỗi đọc thông tin cập nhật: {exc}",
        )
    if comparison >= 0:
        return UpdateCheckResult(
            status="up_to_date",
            current_version=version,
            latest_version=manifest.latest_version,
            manifest=manifest,
            message="Đang dùng phiên bản mới nhất",
        )
    return UpdateCheckResult(
        status="update_available",
        current_version=version,
        latest_version=manifest.latest_version,
        manifest=manifest,
        message="Có bản cập nhật mới",
    )


def prepare_update(
    *,
    update_path: Path | str | None = None,
    manifest: UpdateManifest | None = None,
    settings_database: AppSettingsDatabase | None = None,
    work_root: Path | str | None = None,
) -> PreparedUpdate:
    root = Path(update_path) if update_path is not None else load_update_settings(
        settings_database
    ).update_path
    loaded_manifest = manifest or read_update_manifest(root)
    package_path = (root / loaded_manifest.package).resolve()
    if not package_path.is_file():
        raise UpdateError(f"Không tìm thấy gói cập nhật: {package_path}")
    if package_path.suffix.casefold() != ".zip":
        raise UpdateError("Gói cập nhật phải là file .zip.")
    root_for_temp = Path(work_root) if work_root is not None else application_root()
    temporary_root = root_for_temp / "temp" / "update" / _stamp()
    staging_root = temporary_root / "staging"
    staging_root.mkdir(parents=True, exist_ok=True)
    copied_package = temporary_root / package_path.name
    shutil.copy2(package_path, copied_package)
    try:
        with zipfile.ZipFile(copied_package) as archive:
            archive.extractall(staging_root)
    except zipfile.BadZipFile as exc:
        raise UpdateError(f"Gói cập nhật không phải zip hợp lệ: {exc}") from exc
    payload_root = (
        staging_root
        if loaded_manifest.package_type == "delta"
        else _payload_root(staging_root)
    )
    payload_layout = _resolve_payload_layout(loaded_manifest, payload_root)
    _validate_payload(payload_root, payload_layout, loaded_manifest)
    skipped = tuple(
        path
        for path in payload_root.rglob("*")
        if path.is_file() and _should_skip_payload_file(path, payload_root)
    )
    return PreparedUpdate(
        manifest=loaded_manifest,
        package_path=package_path,
        copied_package_path=copied_package,
        staging_root=staging_root,
        payload_root=payload_root,
        payload_layout=payload_layout,
        skipped_payload_files=skipped,
    )


def apply_update(
    *,
    update_path: Path | str | None = None,
    settings_database: AppSettingsDatabase | None = None,
    current_version: str | None = None,
    app_root: Path | str | None = None,
) -> UpdateApplyResult:
    database = settings_database or AppSettingsDatabase()
    check = check_for_update(
        update_path=update_path,
        current_version=current_version,
        settings_database=database,
    )
    if not check.update_available or check.manifest is None:
        raise UpdateError(check.message or "Không có bản cập nhật mới.")
    if check.manifest.package_type == "delta":
        if not check.manifest.base_version:
            raise UpdateError("Bản cập nhật delta thiếu base_version.")
        if check.current_version != check.manifest.base_version:
            raise UpdateError(
                "Bản cập nhật delta yêu cầu phiên bản nền "
                f"{check.manifest.base_version}. Máy hiện tại là {check.current_version}."
            )
    prepared = prepare_update(
        update_path=update_path,
        manifest=check.manifest,
        settings_database=database,
        work_root=app_root,
    )
    validate_payload_layout_for_current_runtime(prepared)
    backup_path = backup_user_databases(database)
    log_entries: dict[str, object] = {
        "started_at": _now(),
        "old_version": check.current_version,
        "new_version": prepared.manifest.latest_version,
        "update_path": str(update_path or load_update_settings(database).update_path),
        "package": str(prepared.package_path),
        "backup_path": str(backup_path),
        "payload_layout": prepared.payload_layout,
        "migrations": [],
        "errors": [],
    }
    copied_files: tuple[Path, ...] = ()
    skipped_files: tuple[Path, ...] = ()
    deleted_files: tuple[Path, ...] = ()
    applied: tuple[AppliedMigration, ...] = ()
    updater_script: Path | None = None
    try:
        staged_python_migrations = load_python_migrations_from_payload(
            prepared.payload_root
        )
        applied = tuple(
            apply_database_migrations(
                database.database_path,
                prepared.manifest,
                update_root=(
                    Path(update_path)
                    if update_path is not None
                    else load_update_settings(database).update_path
                ),
                python_migrations=staged_python_migrations,
            )
        )
        if getattr(sys, "frozen", False):
            updater_script = create_updater_script(
                prepared,
                target_root=Path(app_root) if app_root is not None else None,
            )
        else:
            target_root = Path(app_root) if app_root is not None else application_root()
            copied_files, skipped_files = install_staged_files(
                prepared.payload_root,
                target_root,
            )
            deleted_files = delete_manifest_files(
                prepared.manifest.delete_files,
                target_root,
            )
        database.save_preference(
            APP_VERSION_PREFERENCE_KEY,
            prepared.manifest.latest_version,
        )
        log_entries["migrations"] = [
            {
                "version": migration.version,
                "migration_name": migration.migration_name,
                "checksum": migration.checksum,
            }
            for migration in applied
        ]
        log_entries["copied_files"] = [str(path) for path in copied_files]
        log_entries["skipped_files"] = [str(path) for path in skipped_files]
        log_entries["deleted_files"] = [str(path) for path in deleted_files]
        log_entries["completed_at"] = _now()
    except Exception as exc:
        log_entries["errors"] = [str(exc)]
        log_path = write_update_log(
            log_entries,
            root=Path(app_root) if app_root is not None else None,
        )
        raise UpdateError(
            f"Cập nhật không thành công. Dữ liệu đã được sao lưu tại: {backup_path}. "
            f"Chi tiết lỗi: {exc}"
        ) from exc
    log_path = write_update_log(
        log_entries,
        root=Path(app_root) if app_root is not None else None,
    )
    return UpdateApplyResult(
        old_version=check.current_version,
        new_version=prepared.manifest.latest_version,
        backup_path=backup_path,
        log_path=log_path,
        copied_files=copied_files,
        skipped_files=skipped_files + prepared.skipped_payload_files,
        applied_migrations=applied,
        restart_required=prepared.manifest.required_app_restart,
        deleted_files=deleted_files,
        updater_script=updater_script,
    )


def apply_database_migrations(
    database_path: Path,
    manifest: UpdateManifest,
    *,
    update_root: Path,
    python_migrations: dict[str, object] | None = None,
) -> list[AppliedMigration]:
    if not manifest.database_migrations:
        return []
    try:
        with closing(sqlite3.connect(database_path, timeout=15)) as connection:
            connection.row_factory = sqlite3.Row
            connection.execute("PRAGMA foreign_keys = ON")
            connection.execute("PRAGMA busy_timeout = 15000")
            return apply_migrations(
                connection,
                manifest.database_migrations,
                update_root=update_root,
                python_migrations=python_migrations,
            )
    except (sqlite3.Error, DatabaseMigrationError) as exc:
        raise UpdateError(f"Không thể cập nhật database: {exc}") from exc


def backup_user_databases(
    settings_database: AppSettingsDatabase,
    *,
    backup_root: Path | None = None,
) -> Path:
    destination = (
        Path(backup_root)
        if backup_root is not None
        else settings_database.database_path.parent / "backups" / "update" / _stamp()
    )
    destination.mkdir(parents=True, exist_ok=True)
    _backup_sqlite_database(
        settings_database.database_path,
        destination / settings_database.database_path.name,
        required=False,
    )
    if settings_database.quiz_database_path.is_file():
        _backup_sqlite_database(
            settings_database.quiz_database_path,
            destination / settings_database.quiz_database_path.name,
            required=False,
        )
    return destination


def install_staged_files(
    payload_root: Path,
    target_root: Path,
) -> tuple[tuple[Path, ...], tuple[Path, ...]]:
    copied: list[Path] = []
    skipped: list[Path] = []
    for source in payload_root.rglob("*"):
        if not source.is_file():
            continue
        relative = source.relative_to(payload_root)
        if _should_skip_relative(relative):
            skipped.append(relative)
            continue
        target = target_root / relative
        target.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(source, target)
        copied.append(relative)
    return tuple(copied), tuple(skipped)


def delete_manifest_files(
    delete_files: tuple[Path, ...],
    target_root: Path,
) -> tuple[Path, ...]:
    deleted: list[Path] = []
    root = target_root.resolve()
    for relative in delete_files:
        if not _is_safe_delete_relative(relative):
            continue
        target = (root / relative).resolve()
        if not _is_relative_to(target, root):
            continue
        if not target.is_file():
            continue
        target.unlink()
        deleted.append(relative)
    return tuple(deleted)


def create_updater_script(
    prepared: PreparedUpdate,
    *,
    target_root: Path | None = None,
) -> Path:
    script_path = prepared.staging_root.parent / "apply-update.bat"
    target_root = target_root or application_root()
    source_root = prepared.payload_root
    excluded_files = " ".join(sorted(DATABASE_FILE_NAMES))
    current_pid = os.getpid()
    delete_commands = [
        f'if exist "{target_root / relative}" del /f /q "{target_root / relative}"'
        for relative in prepared.manifest.delete_files
        if _is_safe_delete_relative(relative)
    ]
    log_path = target_root / "logs" / "apply-update.log"
    script_path.write_text(
        "\r\n".join(
            [
                "@echo off",
                "setlocal",
                f'set "SOURCE={source_root}"',
                f'set "TARGET={target_root}"',
                f'set "LOGFILE={log_path}"',
                f'set "APP_PID={current_pid}"',
                'if not exist "%TARGET%\\logs" mkdir "%TARGET%\\logs"',
                'echo [%DATE% %TIME%] Start AgribankV3 update >> "%LOGFILE%"',
                'echo [%DATE% %TIME%] Waiting for old process %APP_PID% >> "%LOGFILE%"',
                "set /a WAIT_COUNT=0",
                ":WAIT_FOR_APP_EXIT",
                'tasklist /FI "PID eq %APP_PID%" /NH | findstr /R /C:"^ *%APP_PID% " > nul',
                "if errorlevel 1 goto APP_PROCESS_EXITED",
                "if %WAIT_COUNT% GEQ 60 goto APP_WAIT_TIMEOUT",
                "set /a WAIT_COUNT+=1",
                "timeout /t 1 /nobreak > nul",
                "goto WAIT_FOR_APP_EXIT",
                ":APP_WAIT_TIMEOUT",
                'echo [%DATE% %TIME%] Wait timeout; continuing update >> "%LOGFILE%"',
                ":APP_PROCESS_EXITED",
                (
                    'robocopy "%SOURCE%" "%TARGET%" /E '
                    f"/XD .git .venv __pycache__ backups backup logs temp KetQua "
                    f"/XF {excluded_files} *.db *.sqlite *.sqlite3 *.mdb *.accdb "
                    '>> "%LOGFILE%" 2>&1'
                ),
                "set ROBOLVL=%ERRORLEVEL%",
                "if %ROBOLVL% GEQ 8 exit /b %ROBOLVL%",
                *delete_commands,
                'echo [%DATE% %TIME%] Done AgribankV3 update >> "%LOGFILE%"',
                'echo [%DATE% %TIME%] Relaunch AgribankV3 >> "%LOGFILE%"',
                'start "" /d "%TARGET%" "%TARGET%\\AgribankV3.exe"',
                "set STARTLVL=%ERRORLEVEL%",
                'echo [%DATE% %TIME%] Relaunch exit code %STARTLVL% >> "%LOGFILE%"',
                "if %STARTLVL% NEQ 0 exit /b %STARTLVL%",
                "exit /b 0",
                "",
            ]
        ),
        encoding="utf-8",
    )
    return script_path


def get_database_schema_version(
    settings_database: AppSettingsDatabase | None = None,
) -> str:
    database = settings_database or AppSettingsDatabase()
    with closing(sqlite3.connect(database.database_path, timeout=15)) as connection:
        ensure_schema_migrations_table(connection)
        connection.commit()
        return latest_schema_version(connection)


def write_update_log(entries: dict[str, object], *, root: Path | None = None) -> Path:
    log_directory = (root or application_root()) / "logs"
    log_directory.mkdir(parents=True, exist_ok=True)
    log_path = log_directory / "update.log"
    with log_path.open("a", encoding="utf-8") as stream:
        stream.write(json.dumps(entries, ensure_ascii=False, default=str))
        stream.write("\n")
    return log_path


def load_python_migrations_from_payload(payload_root: Path) -> dict[str, object]:
    candidates = (
        payload_root / "src" / "agribank_v3" / "update" / "db_migrations.py",
        payload_root / "agribank_v3" / "update" / "db_migrations.py",
    )
    migration_file = next((path for path in candidates if path.is_file()), None)
    if migration_file is None:
        return {}
    module_name = f"agribank_v3_staged_migrations_{uuid4().hex}"
    spec = util.spec_from_file_location(module_name, migration_file)
    if spec is None or spec.loader is None:
        raise UpdateError(f"Không nạp được Python migration từ {migration_file}")
    module = util.module_from_spec(spec)
    sys.modules[module_name] = module
    try:
        spec.loader.exec_module(module)
        loader = getattr(module, "default_python_migrations", None)
        if not callable(loader):
            sys.modules.pop(module_name, None)
            return {}
        migrations = loader()
    except Exception:
        sys.modules.pop(module_name, None)
        raise
    if not isinstance(migrations, dict):
        sys.modules.pop(module_name, None)
        raise UpdateError("default_python_migrations() trong gói update không trả về dict.")
    return migrations


def _parse_semver(version: str) -> tuple[int, int, int, tuple[str, ...]]:
    text = str(version).strip()
    match = re.fullmatch(
        r"v?(\d+)\.(\d+)\.(\d+)(?:[-+]([0-9A-Za-z.-]+))?",
        text,
    )
    if not match:
        raise ValueError(f"Phiên bản không đúng semantic version: {version}")
    major, minor, patch = (int(match.group(index)) for index in (1, 2, 3))
    suffix = tuple((match.group(4) or "").split(".")) if match.group(4) else ()
    return major, minor, patch, suffix


def _payload_root(staging_root: Path) -> Path:
    children = [path for path in staging_root.iterdir() if not path.name.startswith("__")]
    directories = [path for path in children if path.is_dir()]
    files = [path for path in children if path.is_file()]
    if len(directories) == 1 and not files:
        return directories[0]
    return staging_root


def validate_payload_layout_for_current_runtime(
    prepared: PreparedUpdate,
    *,
    frozen: bool | None = None,
) -> None:
    is_frozen = getattr(sys, "frozen", False) if frozen is None else frozen
    if is_frozen and prepared.payload_layout != PAYLOAD_LAYOUT_APP_ROOT:
        raise UpdateError(
            "Gói cập nhật này là dạng source, không thể thay code của bản "
            "AgribankV3.exe đã build. Hãy build AgribankV3 trước rồi tạo gói "
            "'Gói app đã build EXE' bằng Update Builder."
        )
    if not is_frozen and prepared.payload_layout == PAYLOAD_LAYOUT_APP_ROOT:
        raise UpdateError(
            "Gói cập nhật này dành cho bản AgribankV3.exe đã build. App hiện "
            "đang chạy từ source, hãy tạo gói runtime/source hoặc chạy bản exe "
            "để cập nhật."
        )


def _resolve_payload_layout(manifest: UpdateManifest, payload_root: Path) -> str:
    layout = str(manifest.payload_layout or PAYLOAD_LAYOUT_AUTO).strip().casefold()
    if layout == PAYLOAD_LAYOUT_AUTO:
        return _infer_payload_layout(payload_root)
    return layout


def _infer_payload_layout(payload_root: Path) -> str:
    if (payload_root / "AgribankV3.exe").is_file():
        return PAYLOAD_LAYOUT_APP_ROOT
    source_markers = (
        payload_root / "pyproject.toml",
        payload_root / "src" / "agribank_v3",
        payload_root / "agribank_v3",
    )
    if any(path.exists() for path in source_markers):
        return PAYLOAD_LAYOUT_SOURCE
    raise UpdateError(
        "Gói cập nhật không hợp lệ: không nhận diện được layout payload."
    )


def _validate_payload(
    payload_root: Path,
    payload_layout: str,
    manifest: UpdateManifest,
) -> None:
    if payload_layout == PAYLOAD_LAYOUT_APP_ROOT:
        if not (payload_root / "AgribankV3.exe").is_file():
            raise UpdateError(
                "Gói cập nhật dạng app_root không có AgribankV3.exe."
            )
        _validate_app_build_info(payload_root, manifest.latest_version)
        return
    source_candidates = (
        payload_root / "pyproject.toml",
        payload_root / "src" / "agribank_v3",
        payload_root / "agribank_v3",
    )
    if not any(path.exists() for path in source_candidates):
        raise UpdateError(
            "Gói cập nhật không hợp lệ: không tìm thấy pyproject.toml, "
            "src/agribank_v3 hoặc agribank_v3."
        )


def _validate_app_build_info(payload_root: Path, expected_version: str) -> None:
    build_info_path = payload_root / APP_BUILD_INFO_FILE_NAME
    if not build_info_path.is_file():
        raise UpdateError(
            "Gói cập nhật dạng app_root thiếu agribank_v3_build_info.json. "
            "Hãy build lại AgribankV3 bằng build_portable.ps1 rồi tạo lại "
            "gói 'Gói app đã build EXE'."
        )
    try:
        raw = json.loads(build_info_path.read_text(encoding="utf-8-sig"))
    except Exception as exc:
        raise UpdateError(f"Không đọc được {APP_BUILD_INFO_FILE_NAME}: {exc}") from exc
    if not isinstance(raw, dict):
        raise UpdateError(f"{APP_BUILD_INFO_FILE_NAME} không đúng cấu trúc.")
    app_name = str(raw.get("app", "")).strip()
    build_version = str(raw.get("version", "")).strip()
    if app_name and app_name != "AgribankV3":
        raise UpdateError(f"Gói cập nhật không phải của AgribankV3: {app_name}")
    if not build_version:
        raise UpdateError(f"{APP_BUILD_INFO_FILE_NAME} thiếu trường version.")
    try:
        versions_match = compare_versions(build_version, expected_version) == 0
    except ValueError as exc:
        raise UpdateError(
            f"Version trong {APP_BUILD_INFO_FILE_NAME} không hợp lệ: {exc}"
        ) from exc
    if not versions_match:
        raise UpdateError(
            "Version thật của bản app trong gói cập nhật "
            f"({build_version}) không khớp manifest ({expected_version}). "
            "Hãy cập nhật version trong source, build lại AgribankV3.exe, "
            "rồi tạo lại gói app."
        )


def _should_skip_payload_file(path: Path, payload_root: Path) -> bool:
    return _should_skip_relative(path.relative_to(payload_root))


def _should_skip_relative(relative: Path) -> bool:
    parts = [part.casefold() for part in relative.parts]
    if any(part in SKIPPED_DIRECTORIES for part in parts[:-1]):
        return True
    name = relative.name.casefold()
    if name in DATABASE_FILE_NAMES or name.endswith("-wal") or name.endswith("-shm"):
        return True
    if relative.suffix.casefold() in DATABASE_SUFFIXES:
        return True
    return False


def _is_safe_delete_relative(relative: Path) -> bool:
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
    if _should_skip_relative(relative):
        return False
    return True


def _is_relative_to(path: Path, root: Path) -> bool:
    try:
        path.resolve().relative_to(root.resolve())
    except ValueError:
        return False
    return True


def _backup_sqlite_database(source: Path, destination: Path, *, required: bool) -> None:
    if not source.is_file():
        if required:
            raise UpdateError(f"Không tìm thấy database cần sao lưu: {source}")
        return
    temporary_destination = destination.with_name(
        f".{destination.name}.{uuid4().hex}.tmp"
    )
    try:
        with closing(sqlite3.connect(source, timeout=15)) as source_connection:
            source_connection.execute("PRAGMA busy_timeout = 15000")
            with closing(sqlite3.connect(temporary_destination)) as target:
                source_connection.backup(target)
        os.replace(temporary_destination, destination)
    except sqlite3.Error as exc:
        raise UpdateError(f"Không thể sao lưu database {source.name}: {exc}") from exc
    finally:
        temporary_destination.unlink(missing_ok=True)


def _stamp() -> str:
    return datetime.now().strftime("%Y%m%d_%H%M%S_%f")


def _now() -> str:
    return datetime.now().astimezone().isoformat(timespec="seconds")
