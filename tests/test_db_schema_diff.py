from __future__ import annotations

from contextlib import closing
import json
from pathlib import Path
import sqlite3
import sys
from tempfile import TemporaryDirectory
import unittest


BUILDER_DIR = Path(__file__).resolve().parents[1] / "tools" / "update_builder"
if str(BUILDER_DIR) not in sys.path:
    sys.path.insert(0, str(BUILDER_DIR))

from db_schema_diff import (  # noqa: E402
    compare_sqlite_schema,
    validate_generated_migration_on_copy,
    write_generated_python_migration,
)
from update_builder_core import BuildUpdateConfig, MigrationItem, UpdateBuilder, build_manifest  # noqa: E402


class DbSchemaDiffTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temporary_directory = TemporaryDirectory()
        self.root = Path(self.temporary_directory.name)
        self.old_db = self.root / "old.db"
        self.new_db = self.root / "new.db"

    def tearDown(self) -> None:
        self.temporary_directory.cleanup()

    def test_schema_diff_detect_new_table(self) -> None:
        self._exec(self.old_db, "CREATE TABLE app_preferences(key TEXT PRIMARY KEY, value TEXT, updated_at TEXT)")
        self._exec(
            self.new_db,
            """
            CREATE TABLE app_preferences(key TEXT PRIMARY KEY, value TEXT, updated_at TEXT);
            CREATE TABLE demo(id INTEGER PRIMARY KEY, name TEXT);
            """,
        )

        diff = compare_sqlite_schema(self.old_db, self.new_db)

        self.assertEqual([table.name for table in diff.new_tables], ["demo"])

    def test_schema_diff_detect_new_column(self) -> None:
        self._exec(self.old_db, "CREATE TABLE credit_groups(ma_to TEXT PRIMARY KEY)")
        self._exec(
            self.new_db,
            "CREATE TABLE credit_groups(ma_to TEXT PRIMARY KEY, is_active INTEGER DEFAULT 1)",
        )

        diff = compare_sqlite_schema(self.old_db, self.new_db)

        self.assertEqual(
            [(column.table, column.name) for column in diff.new_columns],
            [("credit_groups", "is_active")],
        )

    def test_schema_diff_detect_new_index(self) -> None:
        self._exec(self.old_db, "CREATE TABLE credit_groups(ma_to TEXT PRIMARY KEY, ten_to TEXT)")
        self._exec(
            self.new_db,
            """
            CREATE TABLE credit_groups(ma_to TEXT PRIMARY KEY, ten_to TEXT);
            CREATE INDEX idx_credit_groups_ten_to ON credit_groups(ten_to);
            """,
        )

        diff = compare_sqlite_schema(self.old_db, self.new_db)

        self.assertEqual([index.name for index in diff.new_indexes], ["idx_credit_groups_ten_to"])

    def test_schema_diff_detect_removed_column_as_dangerous(self) -> None:
        self._exec(self.old_db, "CREATE TABLE credit_groups(ma_to TEXT PRIMARY KEY, old_column TEXT)")
        self._exec(self.new_db, "CREATE TABLE credit_groups(ma_to TEXT PRIMARY KEY)")

        diff = compare_sqlite_schema(self.old_db, self.new_db)

        self.assertTrue(
            any(change.kind == "removed_column" for change in diff.dangerous_changes)
        )

    def test_generate_python_migration_add_table_idempotent(self) -> None:
        self._exec(self.old_db, "CREATE TABLE app_preferences(key TEXT PRIMARY KEY, value TEXT, updated_at TEXT)")
        self._exec(
            self.new_db,
            """
            CREATE TABLE app_preferences(key TEXT PRIMARY KEY, value TEXT, updated_at TEXT);
            CREATE TABLE demo(id INTEGER PRIMARY KEY, name TEXT);
            """,
        )
        diff = compare_sqlite_schema(self.old_db, self.new_db)
        migration = write_generated_python_migration(diff, "0.1.5", self.root)

        first = validate_generated_migration_on_copy(
            old_db_path=self.old_db,
            new_db_path=self.new_db,
            migration_file=migration,
            version="0.1.5",
        )
        second = validate_generated_migration_on_copy(
            old_db_path=first.test_database_path,
            new_db_path=self.new_db,
            migration_file=migration,
            version="0.1.5",
        )

        self.assertTrue(first.success, first.error or str(first.missing_items))
        self.assertTrue(second.success, second.error or str(second.missing_items))

    def test_generate_python_migration_add_column_idempotent(self) -> None:
        self._exec(self.old_db, "CREATE TABLE credit_groups(ma_to TEXT PRIMARY KEY)")
        self._exec(
            self.new_db,
            "CREATE TABLE credit_groups(ma_to TEXT PRIMARY KEY, is_active INTEGER DEFAULT 1)",
        )
        diff = compare_sqlite_schema(self.old_db, self.new_db)
        migration = write_generated_python_migration(diff, "0.1.5", self.root)

        first = validate_generated_migration_on_copy(
            old_db_path=self.old_db,
            new_db_path=self.new_db,
            migration_file=migration,
            version="0.1.5",
        )
        second = validate_generated_migration_on_copy(
            old_db_path=first.test_database_path,
            new_db_path=self.new_db,
            migration_file=migration,
            version="0.1.5",
        )

        self.assertTrue(first.success, first.error or str(first.missing_items))
        self.assertTrue(second.success, second.error or str(second.missing_items))

    def test_generated_migration_preserves_existing_data(self) -> None:
        self._exec(
            self.old_db,
            "CREATE TABLE credit_groups(ma_to TEXT PRIMARY KEY);"
            "INSERT INTO credit_groups(ma_to) VALUES ('T001');",
        )
        self._exec(
            self.new_db,
            "CREATE TABLE credit_groups(ma_to TEXT PRIMARY KEY, is_active INTEGER DEFAULT 1)",
        )
        diff = compare_sqlite_schema(self.old_db, self.new_db)
        migration = write_generated_python_migration(diff, "0.1.5", self.root)

        result = validate_generated_migration_on_copy(
            old_db_path=self.old_db,
            new_db_path=self.new_db,
            migration_file=migration,
            version="0.1.5",
        )

        with closing(sqlite3.connect(result.test_database_path)) as connection:
            row = connection.execute(
                "SELECT ma_to, is_active FROM credit_groups WHERE ma_to = 'T001'"
            ).fetchone()
        self.assertEqual(row, ("T001", 1))

    def test_detect_new_app_preferences_key(self) -> None:
        self._exec(
            self.old_db,
            "CREATE TABLE app_preferences(key TEXT PRIMARY KEY, value TEXT, updated_at TEXT)",
        )
        self._exec(
            self.new_db,
            """
            CREATE TABLE app_preferences(key TEXT PRIMARY KEY, value TEXT, updated_at TEXT);
            INSERT INTO app_preferences(key, value, updated_at)
            VALUES ('update_path', 'X:/Update', '2026-07-20');
            """,
        )

        diff = compare_sqlite_schema(self.old_db, self.new_db)
        migration_text = write_generated_python_migration(diff, "0.1.5", self.root).read_text(encoding="utf-8")

        self.assertEqual([item.key for item in diff.new_app_preferences], ["update_path"])
        self.assertIn("INSERT OR IGNORE INTO app_preferences", migration_text)

    def test_existing_app_preferences_value_not_overwritten(self) -> None:
        self._exec(
            self.old_db,
            """
            CREATE TABLE app_preferences(key TEXT PRIMARY KEY, value TEXT, updated_at TEXT);
            INSERT INTO app_preferences(key, value, updated_at)
            VALUES ('update_path', 'USER_VALUE', '2026-07-19');
            """,
        )
        self._exec(
            self.new_db,
            """
            CREATE TABLE app_preferences(key TEXT PRIMARY KEY, value TEXT, updated_at TEXT);
            INSERT INTO app_preferences(key, value, updated_at)
            VALUES ('update_path', 'DEFAULT_VALUE', '2026-07-20');
            INSERT INTO app_preferences(key, value, updated_at)
            VALUES ('new_key', 'DEFAULT_NEW', '2026-07-20');
            """,
        )
        diff = compare_sqlite_schema(self.old_db, self.new_db)
        migration = write_generated_python_migration(diff, "0.1.5", self.root)

        result = validate_generated_migration_on_copy(
            old_db_path=self.old_db,
            new_db_path=self.new_db,
            migration_file=migration,
            version="0.1.5",
        )

        with closing(sqlite3.connect(result.test_database_path)) as connection:
            user_value = connection.execute(
                "SELECT value FROM app_preferences WHERE key = 'update_path'"
            ).fetchone()[0]
            new_value = connection.execute(
                "SELECT value FROM app_preferences WHERE key = 'new_key'"
            ).fetchone()[0]
        self.assertEqual(user_value, "USER_VALUE")
        self.assertEqual(new_value, "DEFAULT_NEW")
        self.assertTrue(
            any(change.kind == "changed_app_preference" for change in diff.dangerous_changes)
        )

    def test_manifest_adds_python_migration_entry(self) -> None:
        config = BuildUpdateConfig(
            source_path=self.root,
            update_path=self.root / "Update",
            new_version="0.1.5",
            release_date="2026-07-20",
            notes=("Migration",),
            migrations=(
                MigrationItem(
                    version="0.1.5",
                    description="Migration database cho phiên bản 0.1.5",
                    use_python_migration=True,
                ),
            ),
        )

        manifest = build_manifest(config, "AgribankV3_0.1.5.zip")

        self.assertEqual(
            manifest["database_migrations"],
            [
                {
                    "version": "0.1.5",
                    "file": "",
                    "description": "Migration database cho phiên bản 0.1.5",
                }
            ],
        )

    def test_warning_manifest_python_migration_missing_function(self) -> None:
        source = self.root / "source"
        (source / "src" / "agribank_v3" / "update").mkdir(parents=True)
        (source / "src" / "agribank_v3" / "__init__.py").write_text(
            '__version__ = "0.1.4"\n',
            encoding="utf-8",
        )
        (source / "src" / "agribank_v3" / "update" / "db_migrations.py").write_text(
            "def default_python_migrations():\n    return {}\n",
            encoding="utf-8",
        )
        config = BuildUpdateConfig(
            source_path=source,
            update_path=self.root / "Update",
            new_version="0.1.5",
            release_date="2026-07-20",
            notes=("Migration",),
            migrations=(
                MigrationItem(
                    version="0.1.5",
                    description="Migration database cho phiên bản 0.1.5",
                    use_python_migration=True,
                ),
            ),
        )

        result = UpdateBuilder().validate(config)

        self.assertTrue(
            any("chưa tìm thấy migrate_0_1_5" in warning for warning in result.warnings)
        )

    @staticmethod
    def _exec(path: Path, script: str) -> None:
        with closing(sqlite3.connect(path)) as connection:
            connection.executescript(script)
            connection.commit()


if __name__ == "__main__":
    unittest.main()

