from __future__ import annotations

from contextlib import closing
import json
from pathlib import Path
import sqlite3
import sys
from tempfile import TemporaryDirectory
import unittest
import zipfile


ROOT = Path(__file__).resolve().parents[1]
BUILDER_DIR = ROOT / "tools" / "update_builder"
for path in (ROOT / "src", BUILDER_DIR):
    if str(path) not in sys.path:
        sys.path.insert(0, str(path))

from agribank_v3.settings import AppSettingsDatabase  # noqa: E402
from agribank_v3.update.update_manager import (  # noqa: E402
    UpdateError,
    apply_update,
    check_for_update,
)
from db_schema_diff import compare_sqlite_schema, write_generated_python_migration  # noqa: E402
from run_update_flow_demo import (  # noqa: E402
    assert_zip_has_no_database,
    create_demo_app,
    create_demo_layout,
    create_new_dev_database,
    create_old_user_database,
    verify_updated_database,
)
from update_builder_core import (  # noqa: E402
    BuildUpdateConfig,
    MigrationItem,
    UpdateBuilder,
    insert_python_migration_into_source,
)


class UpdateFlowEndToEndTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temporary_directory = TemporaryDirectory()
        self.root = Path(self.temporary_directory.name)

    def tearDown(self) -> None:
        self.temporary_directory.cleanup()

    def test_update_flow_preserves_user_data(self) -> None:
        paths, _applied = self._build_and_apply_demo_update()

        with closing(sqlite3.connect(paths["user_db"])) as connection:
            row = connection.execute(
                "SELECT MaTo, TenTo FROM credit_groups WHERE MaTo = '5491TEST001'"
            ).fetchone()
        self.assertEqual(row, ("5491TEST001", "Tổ test cũ"))

    def test_update_flow_adds_new_column(self) -> None:
        paths, _applied = self._build_and_apply_demo_update()

        with closing(sqlite3.connect(paths["user_db"])) as connection:
            row = connection.execute(
                "SELECT is_active FROM credit_groups WHERE MaTo = '5491TEST001'"
            ).fetchone()
        self.assertEqual(row[0], 1)

    def test_update_flow_adds_new_table(self) -> None:
        paths, _applied = self._build_and_apply_demo_update()

        with closing(sqlite3.connect(paths["user_db"])) as connection:
            exists = connection.execute(
                """
                SELECT 1 FROM sqlite_master
                WHERE type = 'table' AND name = 'credit_group_commission_rules'
                """
            ).fetchone()
        self.assertIsNotNone(exists)

    def test_update_flow_adds_new_default_setting(self) -> None:
        paths, _applied = self._build_and_apply_demo_update()

        with closing(sqlite3.connect(paths["user_db"])) as connection:
            value = connection.execute(
                "SELECT value FROM app_preferences WHERE key = 'new_default_setting'"
            ).fetchone()[0]
        self.assertEqual(value, "default_value")

    def test_update_flow_does_not_overwrite_user_setting(self) -> None:
        paths, _applied = self._build_and_apply_demo_update()

        with closing(sqlite3.connect(paths["user_db"])) as connection:
            value = connection.execute(
                "SELECT value FROM app_preferences WHERE key = 'user_custom_setting'"
            ).fetchone()[0]
        self.assertEqual(value, "gia_tri_nguoi_dung_da_sua")

    def test_update_flow_creates_backup(self) -> None:
        _paths, applied = self._build_and_apply_demo_update()

        self.assertTrue((applied.backup_path / "DuLieuV3.db").is_file())
        with closing(sqlite3.connect(applied.backup_path / "DuLieuV3.db")) as connection:
            row = connection.execute(
                "SELECT TenTo FROM credit_groups WHERE MaTo = '5491TEST001'"
            ).fetchone()
        self.assertEqual(row[0], "Tổ test cũ")

    def test_update_flow_records_migration(self) -> None:
        paths, applied = self._build_and_apply_demo_update()

        self.assertEqual([item.version for item in applied.applied_migrations], ["0.1.1"])
        with closing(sqlite3.connect(paths["user_db"])) as connection:
            row = connection.execute(
                "SELECT version, success FROM app_schema_migrations WHERE version = '0.1.1'"
            ).fetchone()
        self.assertEqual(row, ("0.1.1", 1))

    def test_update_flow_skips_database_inside_package(self) -> None:
        paths, applied = self._build_and_apply_demo_update(include_database_in_package=True)

        self.assertIn(Path("data") / "DuLieuV3.db", applied.skipped_files)
        verify_updated_database(paths["user_db"])

    def test_update_flow_rollback_on_migration_error(self) -> None:
        paths = self._create_base_demo()
        self._write_broken_update_package(paths)
        settings_db = AppSettingsDatabase(paths["user_db"])

        with self.assertRaises(UpdateError):
            apply_update(
                update_path=paths["update_server"],
                settings_database=settings_db,
                current_version="0.1.0",
                app_root=paths["old_app"],
            )

        backup_root = paths["user_db"].parent / "backups" / "update"
        self.assertTrue(any(backup_root.iterdir()))
        with closing(sqlite3.connect(paths["user_db"])) as connection:
            user_row = connection.execute(
                "SELECT TenTo FROM credit_groups WHERE MaTo = '5491TEST001'"
            ).fetchone()
            bad_table = connection.execute(
                """
                SELECT 1 FROM sqlite_master
                WHERE type = 'table' AND name = 'bad_partial_table'
                """
            ).fetchone()
            migration = connection.execute(
                "SELECT 1 FROM app_schema_migrations WHERE version = '0.1.9'"
            ).fetchone()
        self.assertEqual(user_row[0], "Tổ test cũ")
        self.assertIsNone(bad_table)
        self.assertIsNone(migration)

    def test_update_flow_idempotent_migration(self) -> None:
        paths, first = self._build_and_apply_demo_update()
        second = apply_update(
            update_path=paths["update_server"],
            settings_database=AppSettingsDatabase(paths["user_db"]),
            current_version="0.1.0",
            app_root=paths["old_app"],
        )

        self.assertEqual([item.version for item in first.applied_migrations], ["0.1.1"])
        self.assertEqual(second.applied_migrations, ())
        verify_updated_database(paths["user_db"])

    def test_manifest_missing_package_reports_clear_error(self) -> None:
        paths = self._create_base_demo()
        (paths["update_server"] / "manifest.json").write_text(
            json.dumps({"latest_version": "0.1.1"}),
            encoding="utf-8",
        )

        result = check_for_update(
            update_path=paths["update_server"],
            current_version="0.1.0",
            settings_database=AppSettingsDatabase(paths["user_db"]),
        )

        self.assertEqual(result.status, "manifest_error")
        self.assertIn("package", result.message)

    def test_missing_package_reports_clear_error(self) -> None:
        paths = self._create_base_demo()
        (paths["update_server"] / "manifest.json").write_text(
            json.dumps(
                {
                    "latest_version": "0.1.1",
                    "package": "missing.zip",
                    "database_migrations": [],
                }
            ),
            encoding="utf-8",
        )

        with self.assertRaisesRegex(UpdateError, "Không tìm thấy gói cập nhật"):
            apply_update(
                update_path=paths["update_server"],
                settings_database=AppSettingsDatabase(paths["user_db"]),
                current_version="0.1.0",
                app_root=paths["old_app"],
            )

    def _build_and_apply_demo_update(self, *, include_database_in_package: bool = False):
        paths = self._create_base_demo()
        diff = compare_sqlite_schema(paths["user_db"], paths["new_dev_db"])
        self.assertEqual(
            [(column.table, column.name) for column in diff.new_columns],
            [("credit_groups", "is_active")],
        )
        self.assertEqual(
            [table.name for table in diff.new_tables],
            ["credit_group_commission_rules"],
        )
        self.assertEqual(
            [item.key for item in diff.new_app_preferences],
            ["new_default_setting"],
        )
        generated = write_generated_python_migration(diff, "0.1.1", self.root)
        insert_python_migration_into_source(
            source_path=paths["new_app"],
            generated_migration_file=generated,
            version="0.1.1",
        )
        config = BuildUpdateConfig(
            source_path=paths["new_app"],
            update_path=paths["update_server"],
            new_version="0.1.1",
            release_date="2026-07-20",
            notes=("End-to-end update",),
            migrations=(
                MigrationItem(
                    version="0.1.1",
                    description="Migration database cho phiên bản 0.1.1",
                    use_python_migration=True,
                ),
            ),
            auto_update_source_version=True,
        )
        result = UpdateBuilder().build(config)
        assert_zip_has_no_database(result.package_path)
        if include_database_in_package:
            with zipfile.ZipFile(result.package_path, "a") as archive:
                archive.writestr("data/DuLieuV3.db", "malicious empty database")
        check = check_for_update(
            update_path=paths["update_server"],
            current_version="0.1.0",
            settings_database=AppSettingsDatabase(paths["user_db"]),
        )
        self.assertTrue(check.update_available)
        applied = apply_update(
            update_path=paths["update_server"],
            settings_database=AppSettingsDatabase(paths["user_db"]),
            current_version="0.1.0",
            app_root=paths["old_app"],
        )
        self.assertTrue(
            (paths["old_app"] / "src" / "agribank_v3" / "new_feature.py").is_file()
        )
        return paths, applied

    def _create_base_demo(self) -> dict[str, Path]:
        paths = create_demo_layout(self.root)
        create_old_user_database(paths["user_db"], paths["update_server"])
        create_new_dev_database(paths["new_dev_db"])
        create_demo_app(paths["old_app"], version="0.1.0", marker=False)
        create_demo_app(paths["new_app"], version="0.1.0", marker=True)
        return paths

    def _write_broken_update_package(self, paths: dict[str, Path]) -> None:
        package_path = paths["update_server"] / "AgribankV3_0.1.9.zip"
        broken_source = paths["new_app"]
        migration_file = (
            broken_source
            / "src"
            / "agribank_v3"
            / "update"
            / "db_migrations.py"
        )
        migration_file.write_text(
            """
from __future__ import annotations


def default_python_migrations():
    return {"0.1.9": migrate_0_1_9}


def migrate_0_1_9(conn):
    conn.execute("CREATE TABLE bad_partial_table(id INTEGER)")
    raise RuntimeError("forced migration failure")
""",
            encoding="utf-8",
        )
        with zipfile.ZipFile(package_path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
            for path in broken_source.rglob("*"):
                if path.is_file() and path.suffix.casefold() != ".db":
                    archive.write(path, path.relative_to(broken_source).as_posix())
        (paths["update_server"] / "manifest.json").write_text(
            json.dumps(
                {
                    "latest_version": "0.1.9",
                    "package": package_path.name,
                    "release_date": "2026-07-20",
                    "required_app_restart": True,
                    "notes": ["Broken migration"],
                    "database_migrations": [
                        {
                            "version": "0.1.9",
                            "file": "",
                            "description": "Broken migration",
                        }
                    ],
                },
                ensure_ascii=False,
            ),
            encoding="utf-8",
        )


if __name__ == "__main__":
    unittest.main()
