from __future__ import annotations

from contextlib import closing
import json
from pathlib import Path
import sqlite3
import sys
from tempfile import TemporaryDirectory
import unittest
from unittest.mock import patch

from agribank_v3.settings import AppSettingsDatabase
from agribank_v3.update.db_migrations import MigrationSpec
from agribank_v3.update.update_manager import (
    DEFAULT_UPDATE_PATH,
    UpdateError,
    apply_update,
    backup_user_databases,
    check_for_update,
    compare_versions,
    install_staged_files,
    load_python_migrations_from_payload,
    load_update_settings,
    save_update_settings,
    apply_database_migrations,
)
from agribank_v3.update.update_manifest import UpdateManifest, read_update_manifest


class UpdateManagerTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temporary_directory = TemporaryDirectory()
        self.root = Path(self.temporary_directory.name)
        self.database_path = self.root / "data" / "DuLieuV3.db"
        self.database = AppSettingsDatabase(self.database_path)

    def tearDown(self) -> None:
        self.temporary_directory.cleanup()

    def test_default_update_path(self) -> None:
        settings = load_update_settings(self.database)

        self.assertEqual(str(settings.update_path), DEFAULT_UPDATE_PATH)

    def test_save_update_path(self) -> None:
        custom_path = self.root / "Update"

        save_update_settings(custom_path, self.database)
        reopened = AppSettingsDatabase(self.database_path)

        self.assertEqual(load_update_settings(reopened).update_path, custom_path)

    def test_read_manifest(self) -> None:
        update_root = self._write_manifest(
            latest_version="0.1.2",
            package="AgribankV3_0.1.2.zip",
            notes=["Bổ sung cập nhật phiên bản"],
        )

        manifest = read_update_manifest(update_root)

        self.assertEqual(manifest.latest_version, "0.1.2")
        self.assertEqual(manifest.package, "AgribankV3_0.1.2.zip")
        self.assertEqual(manifest.notes, ("Bổ sung cập nhật phiên bản",))
        self.assertEqual(manifest.payload_layout, "auto")

    def test_compare_versions(self) -> None:
        self.assertLess(compare_versions("0.1.1", "0.1.2"), 0)
        self.assertLess(compare_versions("0.1.2", "0.1.10"), 0)
        self.assertLess(compare_versions("0.1.9", "0.2.0"), 0)

    def test_no_update_when_latest_equal_current(self) -> None:
        update_root = self._write_manifest(latest_version="0.1.0")

        result = check_for_update(
            update_path=update_root,
            current_version="0.1.0",
            settings_database=self.database,
        )

        self.assertEqual(result.status, "up_to_date")
        self.assertFalse(result.update_available)

    def test_update_available(self) -> None:
        update_root = self._write_manifest(latest_version="0.1.2")

        result = check_for_update(
            update_path=update_root,
            current_version="0.1.0",
            settings_database=self.database,
        )

        self.assertEqual(result.status, "update_available")
        self.assertTrue(result.update_available)

    def test_database_backup_before_migration(self) -> None:
        with closing(sqlite3.connect(self.database_path)) as connection:
            connection.execute("CREATE TABLE user_data(id INTEGER PRIMARY KEY, name TEXT)")
            connection.execute("INSERT INTO user_data(name) VALUES ('Du lieu goc')")
            connection.commit()

        backup_path = backup_user_databases(
            self.database,
            backup_root=self.root / "backups" / "update-test",
        )

        self.assertTrue((backup_path / "DuLieuV3.db").is_file())
        with closing(sqlite3.connect(backup_path / "DuLieuV3.db")) as connection:
            value = connection.execute("SELECT name FROM user_data").fetchone()[0]
        self.assertEqual(value, "Du lieu goc")

    def test_migration_add_column_preserves_data(self) -> None:
        with closing(sqlite3.connect(self.database_path)) as connection:
            connection.execute("CREATE TABLE credit_groups(ma_to TEXT PRIMARY KEY)")
            connection.execute("INSERT INTO credit_groups(ma_to) VALUES ('T001')")
            connection.commit()
        manifest = UpdateManifest(
            latest_version="0.1.1",
            package="AgribankV3_0.1.1.zip",
            database_migrations=(MigrationSpec(version="0.1.1"),),
        )

        applied = apply_database_migrations(
            self.database_path,
            manifest,
            update_root=self.root,
        )

        self.assertEqual([migration.version for migration in applied], ["0.1.1"])
        with closing(sqlite3.connect(self.database_path)) as connection:
            row = connection.execute(
                "SELECT ma_to, is_active FROM credit_groups WHERE ma_to = 'T001'"
            ).fetchone()
        self.assertEqual(row, ("T001", 1))

    def test_migration_add_table_preserves_data(self) -> None:
        with closing(sqlite3.connect(self.database_path)) as connection:
            connection.execute("CREATE TABLE user_data(id INTEGER PRIMARY KEY, name TEXT)")
            connection.execute("INSERT INTO user_data(name) VALUES ('Du lieu cu')")
            connection.commit()
        manifest = UpdateManifest(
            latest_version="0.1.2",
            package="AgribankV3_0.1.2.zip",
            database_migrations=(MigrationSpec(version="0.1.2"),),
        )

        apply_database_migrations(self.database_path, manifest, update_root=self.root)

        with closing(sqlite3.connect(self.database_path)) as connection:
            table_exists = connection.execute(
                """
                SELECT 1 FROM sqlite_master
                WHERE type = 'table' AND name = 'app_update_settings'
                """
            ).fetchone()
            user_value = connection.execute("SELECT name FROM user_data").fetchone()[0]
        self.assertIsNotNone(table_exists)
        self.assertEqual(user_value, "Du lieu cu")

    def test_migration_idempotent(self) -> None:
        manifest = UpdateManifest(
            latest_version="0.1.2",
            package="AgribankV3_0.1.2.zip",
            database_migrations=(MigrationSpec(version="0.1.2"),),
        )

        first = apply_database_migrations(self.database_path, manifest, update_root=self.root)
        second = apply_database_migrations(self.database_path, manifest, update_root=self.root)

        self.assertEqual(len(first), 1)
        self.assertEqual(second, [])

    def test_insert_default_setting_does_not_overwrite_user_value(self) -> None:
        self.database.save_preference("commission_rate", "user-value")
        migration_file = self.root / "0.1.9.sql"
        migration_file.write_text(
            """
            INSERT OR IGNORE INTO app_preferences(key, value, updated_at)
            VALUES ('commission_rate', 'default-value', '2026-07-20T00:00:00+07:00');
            """,
            encoding="utf-8",
        )
        manifest = UpdateManifest(
            latest_version="0.1.9",
            package="AgribankV3_0.1.9.zip",
            database_migrations=(
                MigrationSpec(version="0.1.9", file=migration_file.name),
            ),
        )

        apply_database_migrations(self.database_path, manifest, update_root=self.root)

        self.assertEqual(
            AppSettingsDatabase(self.database_path).load_preference("commission_rate"),
            "user-value",
        )

    def test_update_package_does_not_overwrite_user_database(self) -> None:
        payload = self.root / "payload"
        package_db = payload / "data" / "DuLieuV3.db"
        package_db.parent.mkdir(parents=True)
        (payload / "pyproject.toml").write_text("[project]\nname='demo'\n", encoding="utf-8")
        package_db.write_text("empty package database", encoding="utf-8")
        target = self.root / "target"
        user_db = target / "data" / "DuLieuV3.db"
        user_db.parent.mkdir(parents=True)
        user_db.write_text("real user database", encoding="utf-8")

        copied, skipped = install_staged_files(payload, target)

        self.assertNotIn(Path("data") / "DuLieuV3.db", copied)
        self.assertIn(Path("data") / "DuLieuV3.db", skipped)
        self.assertEqual(user_db.read_text(encoding="utf-8"), "real user database")

    def test_delta_update_requires_base_version(self) -> None:
        update_root = self._write_delta_update(
            latest_version="0.1.2",
            base_version="0.1.1",
            files={"src/agribank_v3/new.py": "new"},
        )

        with self.assertRaisesRegex(UpdateError, "yêu cầu phiên bản nền 0.1.1"):
            apply_update(
                update_path=update_root,
                settings_database=self.database,
                current_version="0.1.0",
                app_root=self.root / "target",
            )

    def test_delta_update_copies_changed_and_deletes_safe_files(self) -> None:
        update_root = self._write_delta_update(
            latest_version="0.1.2",
            base_version="0.1.1",
            files={"src/agribank_v3/new.py": "new"},
            delete_files=["old.py", "data/DuLieuV3.db"],
        )
        target = self.root / "target"
        target.mkdir(parents=True)
        (target / "old.py").write_text("old", encoding="utf-8")
        user_db = target / "data" / "DuLieuV3.db"
        user_db.parent.mkdir(parents=True)
        user_db.write_text("real user database", encoding="utf-8")

        result = apply_update(
            update_path=update_root,
            settings_database=self.database,
            current_version="0.1.1",
            app_root=target,
        )

        self.assertEqual((target / "src" / "agribank_v3" / "new.py").read_text(encoding="utf-8"), "new")
        self.assertFalse((target / "old.py").exists())
        self.assertTrue(user_db.exists())
        self.assertIn(Path("old.py"), result.deleted_files)
        self.assertNotIn(Path("data") / "DuLieuV3.db", result.deleted_files)

    def test_frozen_update_rejects_source_payload(self) -> None:
        update_root = self._write_update_package(
            latest_version="0.1.2",
            payload_layout="source",
            files={
                "pyproject.toml": "[project]\nname='demo'\n",
                "src/agribank_v3/__init__.py": '__version__ = "0.1.2"\n',
            },
        )

        with patch.object(sys, "frozen", True, create=True):
            with self.assertRaisesRegex(UpdateError, "dạng source"):
                apply_update(
                    update_path=update_root,
                    settings_database=self.database,
                    current_version="0.1.1",
                    app_root=self.root / "target",
                )

    def test_frozen_update_creates_app_root_updater_script(self) -> None:
        update_root = self._write_update_package(
            latest_version="0.1.2",
            payload_layout="app_root",
            files={
                "AgribankV3.exe": "new exe",
                "agribank_v3_build_info.json": (
                    '{"app":"AgribankV3","version":"0.1.2"}'
                ),
                "_internal/library.dat": "new lib",
                "data/DuLieuV3.db": "package db",
            },
        )
        target = self.root / "target"
        target.mkdir()

        with patch.object(sys, "frozen", True, create=True):
            result = apply_update(
                update_path=update_root,
                settings_database=self.database,
                current_version="0.1.1",
                app_root=target,
            )

        self.assertIsNotNone(result.updater_script)
        script_text = result.updater_script.read_text(encoding="utf-8")
        self.assertIn('robocopy "%SOURCE%" "%TARGET%"', script_text)
        self.assertIn("*.db *.sqlite *.sqlite3 *.mdb *.accdb", script_text)
        self.assertIn("apply-update.log", script_text)
        self.assertIn('start "" /d "%TARGET%" "%TARGET%\\AgribankV3.exe"', script_text)

    def test_app_root_update_rejects_build_version_mismatch(self) -> None:
        update_root = self._write_update_package(
            latest_version="0.1.2",
            payload_layout="app_root",
            files={
                "AgribankV3.exe": "old exe",
                "agribank_v3_build_info.json": (
                    '{"app":"AgribankV3","version":"0.1.1"}'
                ),
            },
        )

        with patch.object(sys, "frozen", True, create=True):
            with self.assertRaisesRegex(UpdateError, "không khớp manifest"):
                apply_update(
                    update_path=update_root,
                    settings_database=self.database,
                    current_version="0.1.1",
                    app_root=self.root / "target",
                )

    def test_load_python_migrations_from_payload_supports_dataclasses(self) -> None:
        migration_file = (
            self.root
            / "payload"
            / "src"
            / "agribank_v3"
            / "update"
            / "db_migrations.py"
        )
        migration_file.parent.mkdir(parents=True)
        migration_file.write_text(
            """
from __future__ import annotations

from dataclasses import dataclass
import sqlite3
from typing import Callable


@dataclass(frozen=True, slots=True)
class LocalMigrationSpec:
    version: str


PythonMigration = Callable[[sqlite3.Connection], None]


def migrate_9_9_9(connection: sqlite3.Connection) -> None:
    connection.execute("CREATE TABLE IF NOT EXISTS staged_ok(id INTEGER)")


def default_python_migrations() -> dict[str, PythonMigration]:
    LocalMigrationSpec("9.9.9")
    return {"9.9.9": migrate_9_9_9}
""",
            encoding="utf-8",
        )

        migrations = load_python_migrations_from_payload(self.root / "payload")

        self.assertIn("9.9.9", migrations)

    def _write_manifest(
        self,
        *,
        latest_version: str,
        package: str = "AgribankV3_0.1.1.zip",
        notes: list[str] | None = None,
    ) -> Path:
        update_root = self.root / "Update"
        update_root.mkdir(exist_ok=True)
        (update_root / "manifest.json").write_text(
            json.dumps(
                {
                    "latest_version": latest_version,
                    "package": package,
                    "release_date": "2026-07-20",
                    "required_app_restart": True,
                    "notes": notes or [],
                    "database_migrations": [],
                },
                ensure_ascii=False,
            ),
            encoding="utf-8",
        )
        return update_root

    def _write_update_package(
        self,
        *,
        latest_version: str,
        payload_layout: str,
        files: dict[str, str],
    ) -> Path:
        update_root = self.root / f"Update_{latest_version}_{payload_layout}"
        payload = self.root / f"payload_{latest_version}_{payload_layout}"
        package = update_root / f"AgribankV3_{latest_version}.zip"
        update_root.mkdir(exist_ok=True)
        for relative, text in files.items():
            path = payload / relative
            path.parent.mkdir(parents=True, exist_ok=True)
            path.write_text(text, encoding="utf-8")
        import zipfile

        with zipfile.ZipFile(package, "w", compression=zipfile.ZIP_DEFLATED) as archive:
            for file_path in payload.rglob("*"):
                if file_path.is_file():
                    archive.write(file_path, file_path.relative_to(payload).as_posix())
        (update_root / "manifest.json").write_text(
            json.dumps(
                {
                    "latest_version": latest_version,
                    "package": package.name,
                    "payload_layout": payload_layout,
                    "release_date": "2026-07-20",
                    "required_app_restart": True,
                    "notes": [],
                    "database_migrations": [],
                },
                ensure_ascii=False,
            ),
            encoding="utf-8",
        )
        return update_root

    def _write_delta_update(
        self,
        *,
        latest_version: str,
        base_version: str,
        files: dict[str, str],
        delete_files: list[str] | None = None,
    ) -> Path:
        update_root = self.root / "Update"
        payload = self.root / "delta_payload"
        package = update_root / f"AgribankV3_{latest_version}.zip"
        update_root.mkdir(exist_ok=True)
        for relative, text in files.items():
            path = payload / relative
            path.parent.mkdir(parents=True, exist_ok=True)
            path.write_text(text, encoding="utf-8")
        import zipfile

        with zipfile.ZipFile(package, "w", compression=zipfile.ZIP_DEFLATED) as archive:
            for file_path in payload.rglob("*"):
                if file_path.is_file():
                    archive.write(file_path, file_path.relative_to(payload).as_posix())
        (update_root / "manifest.json").write_text(
            json.dumps(
                {
                    "latest_version": latest_version,
                    "package": package.name,
                    "package_type": "delta",
                    "base_version": base_version,
                    "release_date": "2026-07-20",
                    "required_app_restart": True,
                    "notes": ["Delta"],
                    "delete_files": delete_files or [],
                    "database_migrations": [],
                },
                ensure_ascii=False,
            ),
            encoding="utf-8",
        )
        return update_root


if __name__ == "__main__":
    unittest.main()
