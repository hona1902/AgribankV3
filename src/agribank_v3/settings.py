from __future__ import annotations

from contextlib import closing, contextmanager
from dataclasses import dataclass
from datetime import datetime
from enum import StrEnum
from pathlib import Path
import json
import os
import shutil
import sqlite3
import tempfile
from typing import Iterator
from uuid import uuid4
import zipfile

from agribank_v3.runtime_paths import application_root


class SettingsDatabaseError(RuntimeError):
    pass


class AddinMode(StrEnum):
    PERMANENT = "permanent"
    SESSION = "session"


@dataclass(frozen=True, slots=True)
class BranchProfile:
    branch_code: str = ""
    transaction_office_code: str = ""
    branch_name: str = ""
    reporting_branch_name: str = ""
    department_name: str = ""
    address: str = ""
    tax_code: str = ""
    phone: str = ""
    fax: str = ""
    parent_branch_name: str = ""
    report_location: str = ""
    parent_branch_code: str = ""
    report_preparer: str = ""
    revision: int = 0
    updated_at: str = ""


@dataclass(frozen=True, slots=True)
class DatabaseStatus:
    path: Path
    size_bytes: int
    integrity: str
    branch_revision: int
    last_updated_at: str


@dataclass(frozen=True, slots=True)
class ManagedDatabaseStatus:
    settings: DatabaseStatus
    quiz_path: Path
    quiz_size_bytes: int
    quiz_integrity: str

    @property
    def total_size_bytes(self) -> int:
        return self.settings.size_bytes + self.quiz_size_bytes


class AppSettingsDatabase:
    """Owns durable application settings stored outside Excel workbooks."""

    COMPONENT = "app_settings"
    SCHEMA_VERSION = 3
    DEFAULT_DATABASE_NAME = "DuLieuV3.db"
    QUIZ_DATABASE_NAME = "quiz.db"
    BACKUP_FORMAT = "agribank-v3-database-bundle"
    BACKUP_FORMAT_VERSION = 1
    LEGACY_DATABASE_NAME = "agribank_v3.sqlite3"
    PROFILE_FIELDS = (
        "branch_code",
        "transaction_office_code",
        "branch_name",
        "reporting_branch_name",
        "department_name",
        "address",
        "tax_code",
        "phone",
        "fax",
        "parent_branch_name",
        "report_location",
        "parent_branch_code",
        "report_preparer",
    )

    def __init__(self, database_path: Path | None = None) -> None:
        app_root = application_root()
        self.legacy_database_path = app_root / "data" / self.LEGACY_DATABASE_NAME
        self.database_path = (
            Path(database_path)
            if database_path is not None
            else app_root / "data" / self.DEFAULT_DATABASE_NAME
        )
        self.backup_directory = self.database_path.parent / "backups"
        self.quiz_database_path = (
            self.database_path.parent / self.QUIZ_DATABASE_NAME
        )
        self.database_path.parent.mkdir(parents=True, exist_ok=True)
        self.backup_directory.mkdir(parents=True, exist_ok=True)
        self.initialize_schema()
        if database_path is None:
            self._migrate_from_legacy_database()

    def connect(self) -> sqlite3.Connection:
        connection = sqlite3.connect(self.database_path, timeout=15)
        connection.row_factory = sqlite3.Row
        connection.execute("PRAGMA foreign_keys = ON")
        connection.execute("PRAGMA journal_mode = WAL")
        connection.execute("PRAGMA busy_timeout = 15000")
        return connection

    @contextmanager
    def _database(self) -> Iterator[sqlite3.Connection]:
        database = self.connect()
        try:
            with database:
                yield database
        finally:
            database.close()

    def initialize_schema(self) -> None:
        try:
            with self._database() as database:
                database.executescript(
                    """
                    CREATE TABLE IF NOT EXISTS app_schema_versions (
                        component TEXT PRIMARY KEY,
                        version INTEGER NOT NULL,
                        migrated_at TEXT NOT NULL
                    );

                    CREATE TABLE IF NOT EXISTS branch_profiles (
                        id INTEGER PRIMARY KEY CHECK (id = 1),
                        branch_code TEXT NOT NULL DEFAULT '',
                        transaction_office_code TEXT NOT NULL DEFAULT '',
                        branch_name TEXT NOT NULL DEFAULT '',
                        reporting_branch_name TEXT NOT NULL DEFAULT '',
                        department_name TEXT NOT NULL DEFAULT '',
                        address TEXT NOT NULL DEFAULT '',
                        tax_code TEXT NOT NULL DEFAULT '',
                        phone TEXT NOT NULL DEFAULT '',
                        fax TEXT NOT NULL DEFAULT '',
                        parent_branch_name TEXT NOT NULL DEFAULT '',
                        report_location TEXT NOT NULL DEFAULT '',
                        parent_branch_code TEXT NOT NULL DEFAULT '',
                        report_preparer TEXT NOT NULL DEFAULT '',
                        revision INTEGER NOT NULL DEFAULT 1,
                        created_at TEXT NOT NULL,
                        updated_at TEXT NOT NULL
                    );

                    CREATE TABLE IF NOT EXISTS branch_profile_history (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        profile_id INTEGER NOT NULL,
                        branch_code TEXT NOT NULL,
                        transaction_office_code TEXT NOT NULL,
                        branch_name TEXT NOT NULL,
                        reporting_branch_name TEXT NOT NULL,
                        department_name TEXT NOT NULL,
                        address TEXT NOT NULL,
                        tax_code TEXT NOT NULL,
                        phone TEXT NOT NULL,
                        fax TEXT NOT NULL,
                        parent_branch_name TEXT NOT NULL,
                        report_location TEXT NOT NULL,
                        parent_branch_code TEXT NOT NULL,
                        report_preparer TEXT NOT NULL,
                        revision INTEGER NOT NULL,
                        saved_at TEXT NOT NULL
                    );
                    CREATE INDEX IF NOT EXISTS idx_branch_history_revision
                        ON branch_profile_history(profile_id, revision DESC);

                    CREATE TABLE IF NOT EXISTS app_preferences (
                        key TEXT PRIMARY KEY,
                        value TEXT NOT NULL,
                        updated_at TEXT NOT NULL
                    );

                    CREATE TABLE IF NOT EXISTS addin_preferences (
                        file_name TEXT PRIMARY KEY COLLATE NOCASE,
                        enabled INTEGER NOT NULL DEFAULT 1
                            CHECK (enabled IN (0, 1)),
                        updated_at TEXT NOT NULL
                    );
                    """
                )
                database.execute(
                    """
                    INSERT INTO app_schema_versions(component, version, migrated_at)
                    VALUES (?, ?, ?)
                    ON CONFLICT(component) DO UPDATE SET
                        version = excluded.version,
                        migrated_at = excluded.migrated_at
                    """,
                    (self.COMPONENT, self.SCHEMA_VERSION, self._now()),
                )
        except sqlite3.Error as exc:
            raise SettingsDatabaseError(
                f"Không thể khởi tạo cơ sở dữ liệu cài đặt: {exc}"
            ) from exc

    def load_branch_profile(self) -> BranchProfile:
        try:
            with self._database() as database:
                row = database.execute(
                    "SELECT * FROM branch_profiles WHERE id = 1"
                ).fetchone()
        except sqlite3.Error as exc:
            raise SettingsDatabaseError(
                f"Không thể đọc thông tin chi nhánh: {exc}"
            ) from exc
        if row is None:
            return BranchProfile()
        values = {field: str(row[field] or "") for field in self.PROFILE_FIELDS}
        return BranchProfile(
            **values,
            revision=int(row["revision"]),
            updated_at=str(row["updated_at"]),
        )

    def save_branch_profile(self, profile: BranchProfile) -> BranchProfile:
        values = {
            field: self._clean(getattr(profile, field))
            for field in self.PROFILE_FIELDS
        }
        now = self._now()
        try:
            with self._database() as database:
                existing = database.execute(
                    "SELECT revision, created_at FROM branch_profiles WHERE id = 1"
                ).fetchone()
                revision = int(existing["revision"]) + 1 if existing else 1
                created_at = str(existing["created_at"]) if existing else now
                columns = ", ".join(self.PROFILE_FIELDS)
                placeholders = ", ".join("?" for _ in self.PROFILE_FIELDS)
                updates = ", ".join(
                    f"{field} = excluded.{field}" for field in self.PROFILE_FIELDS
                )
                parameters = [values[field] for field in self.PROFILE_FIELDS]
                database.execute(
                    f"""
                    INSERT INTO branch_profiles(
                        id, {columns}, revision, created_at, updated_at
                    )
                    VALUES (1, {placeholders}, ?, ?, ?)
                    ON CONFLICT(id) DO UPDATE SET
                        {updates},
                        revision = excluded.revision,
                        updated_at = excluded.updated_at
                    """,
                    (*parameters, revision, created_at, now),
                )
                database.execute(
                    f"""
                    INSERT INTO branch_profile_history(
                        profile_id, {columns}, revision, saved_at
                    )
                    VALUES (1, {placeholders}, ?, ?)
                    """,
                    (*parameters, revision, now),
                )
        except sqlite3.Error as exc:
            raise SettingsDatabaseError(
                f"Không thể lưu thông tin chi nhánh: {exc}"
            ) from exc
        return BranchProfile(
            **values,
            revision=revision,
            updated_at=now,
        )

    def load_addin_mode(self) -> AddinMode:
        try:
            with self._database() as database:
                row = database.execute(
                    "SELECT value FROM app_preferences WHERE key = 'addin_mode'"
                ).fetchone()
        except sqlite3.Error as exc:
            raise SettingsDatabaseError(
                f"Không thể đọc chế độ add-in: {exc}"
            ) from exc
        if row is None:
            return AddinMode.PERMANENT
        try:
            return AddinMode(str(row["value"]))
        except ValueError:
            return AddinMode.PERMANENT

    def save_addin_mode(self, mode: AddinMode | str) -> AddinMode:
        try:
            normalized = AddinMode(mode)
        except ValueError as exc:
            raise SettingsDatabaseError("Chế độ add-in không hợp lệ.") from exc
        try:
            with self._database() as database:
                database.execute(
                    """
                    INSERT INTO app_preferences(key, value, updated_at)
                    VALUES ('addin_mode', ?, ?)
                    ON CONFLICT(key) DO UPDATE SET
                        value = excluded.value,
                        updated_at = excluded.updated_at
                    """,
                    (normalized.value, self._now()),
                )
        except sqlite3.Error as exc:
            raise SettingsDatabaseError(
                f"Không thể lưu chế độ add-in: {exc}"
            ) from exc
        return normalized

    def load_preference(self, key: str, default: str = "") -> str:
        normalized_key = self._clean(key)
        if not normalized_key:
            return default
        try:
            with self._database() as database:
                row = database.execute(
                    "SELECT value FROM app_preferences WHERE key = ?",
                    (normalized_key,),
                ).fetchone()
        except sqlite3.Error as exc:
            raise SettingsDatabaseError(
                f"Không thể đọc tùy chọn ứng dụng: {exc}"
            ) from exc
        return str(row["value"]) if row is not None else default

    def save_preference(self, key: str, value: str) -> str:
        normalized_key = self._clean(key)
        if not normalized_key:
            raise SettingsDatabaseError("Tên tùy chọn ứng dụng không hợp lệ.")
        normalized_value = self._clean(value)
        try:
            with self._database() as database:
                database.execute(
                    """
                    INSERT INTO app_preferences(key, value, updated_at)
                    VALUES (?, ?, ?)
                    ON CONFLICT(key) DO UPDATE SET
                        value = excluded.value,
                        updated_at = excluded.updated_at
                    """,
                    (normalized_key, normalized_value, self._now()),
                )
        except sqlite3.Error as exc:
            raise SettingsDatabaseError(
                f"Không thể lưu tùy chọn ứng dụng: {exc}"
            ) from exc
        return normalized_value

    def load_addin_states(
        self, file_names: tuple[str, ...] | list[str]
    ) -> dict[str, bool]:
        names = tuple(self._validate_addin_name(name) for name in file_names)
        if not names:
            return {}
        try:
            with self._database() as database:
                rows = database.execute(
                    "SELECT file_name, enabled FROM addin_preferences"
                ).fetchall()
        except sqlite3.Error as exc:
            raise SettingsDatabaseError(
                f"Không thể đọc trạng thái add-in: {exc}"
            ) from exc
        saved = {
            str(row["file_name"]).casefold(): bool(row["enabled"])
            for row in rows
        }
        return {name: saved.get(name.casefold(), True) for name in names}

    def save_addin_enabled(self, file_name: str, enabled: bool) -> bool:
        name = self._validate_addin_name(file_name)
        try:
            with self._database() as database:
                database.execute(
                    """
                    INSERT INTO addin_preferences(file_name, enabled, updated_at)
                    VALUES (?, ?, ?)
                    ON CONFLICT(file_name) DO UPDATE SET
                        enabled = excluded.enabled,
                        updated_at = excluded.updated_at
                    """,
                    (name, int(bool(enabled)), self._now()),
                )
        except sqlite3.Error as exc:
            raise SettingsDatabaseError(
                f"Không thể lưu trạng thái add-in: {exc}"
            ) from exc
        return bool(enabled)

    def status(self) -> DatabaseStatus:
        try:
            with self._database() as database:
                integrity = str(
                    database.execute("PRAGMA integrity_check").fetchone()[0]
                )
                row = database.execute(
                    "SELECT revision, updated_at FROM branch_profiles WHERE id = 1"
                ).fetchone()
        except sqlite3.Error as exc:
            raise SettingsDatabaseError(
                f"Không thể kiểm tra cơ sở dữ liệu: {exc}"
            ) from exc
        return DatabaseStatus(
            path=self.database_path,
            size_bytes=self.database_path.stat().st_size,
            integrity=integrity,
            branch_revision=int(row["revision"]) if row else 0,
            last_updated_at=str(row["updated_at"]) if row else "",
        )

    def managed_status(self) -> ManagedDatabaseStatus:
        settings_status = self.status()
        if not self.quiz_database_path.is_file():
            return ManagedDatabaseStatus(
                settings=settings_status,
                quiz_path=self.quiz_database_path,
                quiz_size_bytes=0,
                quiz_integrity="Chưa tạo",
            )
        try:
            with closing(sqlite3.connect(self.quiz_database_path)) as database:
                integrity = str(
                    database.execute("PRAGMA integrity_check").fetchone()[0]
                )
        except sqlite3.Error as exc:
            raise SettingsDatabaseError(
                f"Không thể kiểm tra cơ sở dữ liệu trắc nghiệm: {exc}"
            ) from exc
        return ManagedDatabaseStatus(
            settings=settings_status,
            quiz_path=self.quiz_database_path,
            quiz_size_bytes=self.quiz_database_path.stat().st_size,
            quiz_integrity=integrity,
        )

    def create_backup(self, destination: Path | None = None) -> Path:
        if destination is None:
            stamp = datetime.now().strftime("%Y%m%d-%H%M%S-%f")
            unique = uuid4().hex[:8]
            destination = (
                self.backup_directory
                / f"AgribankV3-{stamp}-{unique}.zip"
            )
        destination = Path(destination)
        destination.parent.mkdir(parents=True, exist_ok=True)
        temporary_destination = destination.with_name(
            f".{destination.name}.{uuid4().hex}.tmp"
        )
        try:
            with tempfile.TemporaryDirectory(
                prefix="agribank-v3-backup-"
            ) as temporary_directory:
                temporary_root = Path(temporary_directory)
                settings_snapshot = temporary_root / self.DEFAULT_DATABASE_NAME
                self._snapshot_database(
                    self.database_path,
                    settings_snapshot,
                    required_table="branch_profiles",
                )
                database_entries = [
                    {
                        "role": "settings",
                        "filename": self.DEFAULT_DATABASE_NAME,
                    }
                ]
                if not self.quiz_database_path.is_file():
                    raise SettingsDatabaseError(
                        "Chưa tìm thấy quiz.db. Hãy mở mục Trắc nghiệm một lần "
                        "để khởi tạo database trước khi sao lưu."
                    )
                quiz_snapshot = temporary_root / self.QUIZ_DATABASE_NAME
                self._snapshot_database(
                    self.quiz_database_path,
                    quiz_snapshot,
                    required_table="questions",
                )
                database_entries.append(
                    {
                        "role": "quiz",
                        "filename": self.QUIZ_DATABASE_NAME,
                    }
                )
                manifest = {
                    "format": self.BACKUP_FORMAT,
                    "version": self.BACKUP_FORMAT_VERSION,
                    "created_at": self._now(),
                    "databases": database_entries,
                }
                with zipfile.ZipFile(
                    temporary_destination,
                    "w",
                    compression=zipfile.ZIP_DEFLATED,
                ) as archive:
                    archive.writestr(
                        "manifest.json",
                        json.dumps(manifest, ensure_ascii=False, indent=2),
                    )
                    for entry in database_entries:
                        filename = str(entry["filename"])
                        archive.write(temporary_root / filename, filename)
                os.replace(temporary_destination, destination)
        except SettingsDatabaseError:
            raise
        except (sqlite3.Error, OSError, ValueError, zipfile.BadZipFile) as exc:
            raise SettingsDatabaseError(
                f"Không thể sao lưu cơ sở dữ liệu: {exc}"
            ) from exc
        finally:
            temporary_destination.unlink(missing_ok=True)
        return destination

    def restore_backup(self, source_path: Path) -> Path:
        source_path = Path(source_path)
        if not source_path.is_file():
            raise SettingsDatabaseError("Không tìm thấy tệp sao lưu đã chọn.")
        if source_path.resolve() == self.database_path.resolve():
            raise SettingsDatabaseError(
                "Tệp phục hồi không được trùng với database đang sử dụng."
            )
        if zipfile.is_zipfile(source_path):
            return self._restore_bundle(source_path)
        return self._restore_legacy_settings_backup(source_path)

    def _restore_legacy_settings_backup(self, source_path: Path) -> Path:
        safety_backup = self.create_backup()
        restore_temp = (
            self.database_path.parent / f".restore-{uuid4().hex}.db"
        )
        try:
            uri = source_path.resolve().as_uri() + "?mode=ro"
            with closing(sqlite3.connect(uri, uri=True)) as source:
                integrity = str(
                    source.execute("PRAGMA integrity_check").fetchone()[0]
                )
                if integrity.casefold() != "ok":
                    raise SettingsDatabaseError(
                        f"Tệp sao lưu bị lỗi (integrity_check: {integrity})."
                    )
                tables = {
                    str(row[0])
                    for row in source.execute(
                        "SELECT name FROM sqlite_master WHERE type = 'table'"
                    )
                }
                if "branch_profiles" not in tables:
                    raise SettingsDatabaseError(
                        "Tệp đã chọn không phải database của AgribankV3."
                    )
                with closing(sqlite3.connect(restore_temp)) as target:
                    source.backup(target)
                    restored_integrity = str(
                        target.execute("PRAGMA integrity_check").fetchone()[0]
                    )
                    if restored_integrity.casefold() != "ok":
                        raise SettingsDatabaseError(
                            "Snapshot phục hồi không vượt qua kiểm tra toàn vẹn."
                        )

            # Normalize and close the current WAL before atomically replacing
            # the database file. This prevents old WAL frames from being
            # replayed over the restored snapshot.
            with closing(
                sqlite3.connect(self.database_path, timeout=15)
            ) as current:
                current.execute("PRAGMA busy_timeout = 15000")
                current.execute("PRAGMA wal_checkpoint(TRUNCATE)")
                current.execute("PRAGMA journal_mode = DELETE")
            for suffix in ("-wal", "-shm"):
                sidecar = Path(f"{self.database_path}{suffix}")
                sidecar.unlink(missing_ok=True)
            os.replace(restore_temp, self.database_path)
            self.initialize_schema()
        except SettingsDatabaseError:
            raise
        except (sqlite3.Error, OSError) as exc:
            raise SettingsDatabaseError(
                f"Không thể phục hồi cơ sở dữ liệu: {exc}"
            ) from exc
        finally:
            restore_temp.unlink(missing_ok=True)
        return safety_backup

    def _restore_bundle(self, source_path: Path) -> Path:
        with tempfile.TemporaryDirectory(
            prefix="agribank-v3-restore-"
        ) as temporary_directory:
            temporary_root = Path(temporary_directory)
            try:
                with zipfile.ZipFile(source_path, "r") as archive:
                    manifest = json.loads(archive.read("manifest.json"))
                    if (
                        manifest.get("format") != self.BACKUP_FORMAT
                        or manifest.get("version") != self.BACKUP_FORMAT_VERSION
                    ):
                        raise SettingsDatabaseError(
                            "Gói sao lưu không đúng định dạng AgribankV3."
                        )
                    entries = {
                        str(entry["role"]): str(entry["filename"])
                        for entry in manifest.get("databases", ())
                    }
                    if entries.get("settings") != self.DEFAULT_DATABASE_NAME:
                        raise SettingsDatabaseError(
                            "Gói sao lưu không chứa database cài đặt."
                        )
                    if entries.get("quiz") != self.QUIZ_DATABASE_NAME:
                        raise SettingsDatabaseError(
                            "Gói sao lưu không chứa database trắc nghiệm."
                        )
                    allowed_files = {
                        self.DEFAULT_DATABASE_NAME,
                        self.QUIZ_DATABASE_NAME,
                    }
                    if any(name not in allowed_files for name in entries.values()):
                        raise SettingsDatabaseError(
                            "Gói sao lưu chứa tên database không hợp lệ."
                        )
                    staged: dict[str, Path] = {}
                    for role, filename in entries.items():
                        target = temporary_root / f"restore-{filename}"
                        target.write_bytes(archive.read(filename))
                        required_table = (
                            "branch_profiles" if role == "settings" else "questions"
                        )
                        self._validate_snapshot(target, required_table)
                        staged[role] = target
            except SettingsDatabaseError:
                raise
            except (
                KeyError,
                TypeError,
                ValueError,
                json.JSONDecodeError,
                OSError,
                zipfile.BadZipFile,
            ) as exc:
                raise SettingsDatabaseError(
                    f"Không thể đọc gói sao lưu: {exc}"
                ) from exc

            safety_backup = self.create_backup()
            rollback_settings = temporary_root / "rollback-settings.db"
            self._snapshot_database(
                self.database_path,
                rollback_settings,
                required_table="branch_profiles",
            )
            rollback_quiz: Path | None = None
            quiz_existed = self.quiz_database_path.is_file()
            if quiz_existed:
                rollback_quiz = temporary_root / "rollback-quiz.db"
                self._snapshot_database(
                    self.quiz_database_path,
                    rollback_quiz,
                    required_table="questions",
                )

            try:
                self._replace_database(staged["settings"], self.database_path)
                if "quiz" in staged:
                    self._replace_database(
                        staged["quiz"],
                        self.quiz_database_path,
                    )
                self.initialize_schema()
            except (sqlite3.Error, OSError, SettingsDatabaseError) as exc:
                try:
                    self._replace_database(
                        rollback_settings,
                        self.database_path,
                        consume_source=False,
                    )
                    if rollback_quiz is not None:
                        self._replace_database(
                            rollback_quiz,
                            self.quiz_database_path,
                            consume_source=False,
                        )
                    elif not quiz_existed:
                        self.quiz_database_path.unlink(missing_ok=True)
                    self.initialize_schema()
                except (sqlite3.Error, OSError, SettingsDatabaseError):
                    pass
                raise SettingsDatabaseError(
                    f"Không thể phục hồi gói cơ sở dữ liệu: {exc}"
                ) from exc
            return safety_backup

    @staticmethod
    def _validate_snapshot(path: Path, required_table: str) -> None:
        uri = path.resolve().as_uri() + "?mode=ro"
        with closing(sqlite3.connect(uri, uri=True)) as database:
            integrity = str(
                database.execute("PRAGMA integrity_check").fetchone()[0]
            )
            if integrity.casefold() != "ok":
                raise SettingsDatabaseError(
                    f"Snapshot {path.name} bị lỗi (integrity_check: {integrity})."
                )
            tables = {
                str(row[0])
                for row in database.execute(
                    "SELECT name FROM sqlite_master WHERE type = 'table'"
                )
            }
            if required_table not in tables:
                raise SettingsDatabaseError(
                    f"Snapshot {path.name} thiếu bảng {required_table}."
                )

    def _snapshot_database(
        self,
        source_path: Path,
        destination: Path,
        required_table: str,
    ) -> None:
        with closing(sqlite3.connect(source_path, timeout=15)) as source:
            source.execute("PRAGMA busy_timeout = 15000")
            with closing(sqlite3.connect(destination)) as target:
                source.backup(target)
        self._validate_snapshot(destination, required_table)

    @staticmethod
    def _replace_database(
        source: Path,
        destination: Path,
        *,
        consume_source: bool = True,
    ) -> None:
        destination.parent.mkdir(parents=True, exist_ok=True)
        if destination.is_file():
            with closing(sqlite3.connect(destination, timeout=15)) as current:
                current.execute("PRAGMA busy_timeout = 15000")
                current.execute("PRAGMA wal_checkpoint(TRUNCATE)")
                current.execute("PRAGMA journal_mode = DELETE")
        for suffix in ("-wal", "-shm"):
            Path(f"{destination}{suffix}").unlink(missing_ok=True)
        replacement = source
        if not consume_source:
            replacement = destination.with_name(
                f".rollback-{uuid4().hex}.db"
            )
            shutil.copy2(source, replacement)
        os.replace(replacement, destination)

    def _migrate_from_legacy_database(self) -> None:
        if not self.legacy_database_path.is_file():
            return
        if self.legacy_database_path.resolve() == self.database_path.resolve():
            return
        try:
            with self._database() as database:
                if self._table_has_rows(database, "main", "branch_profiles"):
                    return
                database.execute(
                    "ATTACH DATABASE ? AS legacy",
                    (str(self.legacy_database_path),),
                )
                try:
                    if not self._table_exists(database, "legacy", "branch_profiles"):
                        return
                    database.execute(
                        """
                        INSERT OR IGNORE INTO branch_profiles
                        SELECT * FROM legacy.branch_profiles
                        """
                    )
                    if self._table_exists(
                        database, "legacy", "branch_profile_history"
                    ):
                        database.execute(
                            """
                            INSERT OR IGNORE INTO branch_profile_history
                            SELECT * FROM legacy.branch_profile_history
                            """
                        )
                    database.commit()
                finally:
                    database.execute("DETACH DATABASE legacy")
        except sqlite3.Error as exc:
            raise SettingsDatabaseError(
                f"Không thể chuyển dữ liệu app từ database cũ: {exc}"
            ) from exc

    @staticmethod
    def _table_exists(
        database: sqlite3.Connection,
        schema_name: str,
        table_name: str,
    ) -> bool:
        return (
            database.execute(
                f"""
                SELECT 1
                FROM {schema_name}.sqlite_master
                WHERE type = 'table' AND name = ?
                """,
                (table_name,),
            ).fetchone()
            is not None
        )

    @classmethod
    def _table_has_rows(
        cls,
        database: sqlite3.Connection,
        schema_name: str,
        table_name: str,
    ) -> bool:
        if not cls._table_exists(database, schema_name, table_name):
            return False
        return int(
            database.execute(
                f"SELECT COUNT(*) FROM {schema_name}.{table_name}"
            ).fetchone()[0]
        ) > 0

    @staticmethod
    def _clean(value: object) -> str:
        return " ".join(str(value or "").strip().split())

    @staticmethod
    def _validate_addin_name(file_name: object) -> str:
        name = str(file_name or "").strip()
        if (
            not name
            or Path(name).name != name
            or Path(name).suffix.casefold() not in {".xla", ".xlam"}
        ):
            raise SettingsDatabaseError("Tên tệp add-in không hợp lệ.")
        return name

    @staticmethod
    def _now() -> str:
        return datetime.now().astimezone().isoformat(timespec="seconds")
