from __future__ import annotations

from contextlib import closing
import json
from pathlib import Path
import shutil
import sqlite3
import sys
import tempfile
import zipfile

ROOT = Path(__file__).resolve().parents[2]
if str(ROOT / "src") not in sys.path:
    sys.path.insert(0, str(ROOT / "src"))
if str(Path(__file__).resolve().parent) not in sys.path:
    sys.path.insert(0, str(Path(__file__).resolve().parent))

from agribank_v3.settings import AppSettingsDatabase
from agribank_v3.update.update_manager import apply_update, check_for_update
from db_schema_diff import compare_sqlite_schema, write_generated_python_migration
from update_builder_core import (
    BuildUpdateConfig,
    MigrationItem,
    UpdateBuilder,
    insert_python_migration_into_source,
)


def main() -> int:
    with tempfile.TemporaryDirectory(prefix="agribankv3-update-flow-") as temp:
        demo = Path(temp)
        paths = create_demo_layout(demo)
        create_old_user_database(paths["user_db"], paths["update_server"])
        print("PASS: created old user database")
        create_new_dev_database(paths["new_dev_db"])
        create_demo_app(paths["old_app"], version="0.1.0", marker=False)
        create_demo_app(paths["new_app"], version="0.1.0", marker=True)

        diff = compare_sqlite_schema(paths["user_db"], paths["new_dev_db"])
        assert diff.new_columns and diff.new_tables and diff.new_app_preferences
        print("PASS: detected schema changes")

        generated = write_generated_python_migration(diff, "0.1.1")
        insert_python_migration_into_source(
            source_path=paths["new_app"],
            generated_migration_file=generated,
            version="0.1.1",
        )
        print("PASS: generated migration")

        config = BuildUpdateConfig(
            source_path=paths["new_app"],
            update_path=paths["update_server"],
            new_version="0.1.1",
            release_date="2026-07-20",
            notes=("Demo update flow",),
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
        assert result.package_path.is_file()
        assert result.manifest_path.is_file()
        assert_zip_has_no_database(result.package_path)
        print("PASS: created update package")

        settings_db = AppSettingsDatabase(paths["user_db"])
        check = check_for_update(
            update_path=paths["update_server"],
            current_version="0.1.0",
            settings_database=settings_db,
        )
        assert check.update_available
        applied = apply_update(
            update_path=paths["update_server"],
            settings_database=settings_db,
            current_version="0.1.0",
            app_root=paths["old_app"],
        )
        assert (applied.backup_path / "DuLieuV3.db").is_file()
        print("PASS: database backup created")

        assert applied.applied_migrations
        assert applied.applied_migrations[0].version == "0.1.1"
        print("PASS: migration applied")

        verify_updated_database(paths["user_db"])
        print("PASS: user data preserved")
        print("PASS: new schema exists")
        print("PASS: user setting not overwritten")
        assert (paths["old_app"] / "src" / "agribank_v3" / "new_feature.py").is_file()
        print("PASS: code updated")
        print("PASS: demo used temporary data only")
    return 0


def create_demo_layout(root: Path) -> dict[str, Path]:
    paths = {
        "old_app": root / "old_app",
        "new_app": root / "new_app",
        "update_server": root / "update_server",
        "user_data": root / "user_data",
        "backups": root / "backups",
    }
    for path in paths.values():
        path.mkdir(parents=True, exist_ok=True)
    paths["user_db"] = paths["user_data"] / "DuLieuV3.db"
    paths["new_dev_db"] = root / "new_dev.db"
    return paths


def create_old_user_database(path: Path, update_server: Path) -> None:
    with closing(sqlite3.connect(path)) as connection:
        connection.executescript(
            """
            CREATE TABLE credit_groups (
                MaTo TEXT PRIMARY KEY,
                TenTo TEXT NOT NULL
            );
            INSERT INTO credit_groups(MaTo, TenTo)
            VALUES ('5491TEST001', 'Tổ test cũ');

            CREATE TABLE app_preferences (
                key TEXT PRIMARY KEY,
                value TEXT NOT NULL,
                updated_at TEXT NOT NULL
            );
            """
        )
        connection.execute(
            """
            INSERT INTO app_preferences(key, value, updated_at)
            VALUES ('update_path', ?, '2026-07-20')
            """,
            (str(update_server),),
        )
        connection.execute(
            """
            INSERT INTO app_preferences(key, value, updated_at)
            VALUES ('user_custom_setting', 'gia_tri_nguoi_dung_da_sua', '2026-07-20')
            """
        )
        connection.commit()


def create_new_dev_database(path: Path) -> None:
    with closing(sqlite3.connect(path)) as connection:
        connection.executescript(
            """
            CREATE TABLE credit_groups (
                MaTo TEXT PRIMARY KEY,
                TenTo TEXT NOT NULL,
                is_active INTEGER DEFAULT 1
            );
            CREATE TABLE credit_group_commission_rules (
                MaTo TEXT PRIMARY KEY,
                use_custom_rule INTEGER DEFAULT 0,
                updated_at TEXT
            );
            CREATE TABLE app_preferences (
                key TEXT PRIMARY KEY,
                value TEXT NOT NULL,
                updated_at TEXT NOT NULL
            );
            INSERT INTO app_preferences(key, value, updated_at)
            VALUES ('new_default_setting', 'default_value', '2026-07-20');
            INSERT INTO app_preferences(key, value, updated_at)
            VALUES ('user_custom_setting', 'dev_default_should_not_overwrite', '2026-07-20');
            """
        )
        connection.commit()


def create_demo_app(root: Path, *, version: str, marker: bool) -> None:
    package = root / "src" / "agribank_v3"
    update_pkg = package / "update"
    update_pkg.mkdir(parents=True, exist_ok=True)
    (package / "__init__.py").write_text(f'__version__ = "{version}"\n', encoding="utf-8")
    (package / "__main__.py").write_text("print('AgribankV3 demo')\n", encoding="utf-8")
    if marker:
        (package / "new_feature.py").write_text("UPDATED = True\n", encoding="utf-8")
    (root / "pyproject.toml").write_text(
        f'[project]\nname = "agribank-v3-demo"\nversion = "{version}"\n',
        encoding="utf-8",
    )
    (update_pkg / "__init__.py").write_text("", encoding="utf-8")
    (update_pkg / "db_migrations.py").write_text(_base_db_migrations(), encoding="utf-8")
    data = root / "data"
    data.mkdir(exist_ok=True)
    (data / "DuLieuV3.db").write_text("package database must be skipped", encoding="utf-8")


def verify_updated_database(path: Path) -> None:
    with closing(sqlite3.connect(path)) as connection:
        row = connection.execute(
            "SELECT MaTo, TenTo, is_active FROM credit_groups WHERE MaTo = '5491TEST001'"
        ).fetchone()
        assert row == ("5491TEST001", "Tổ test cũ", 1), row
        table = connection.execute(
            """
            SELECT 1 FROM sqlite_master
            WHERE type = 'table' AND name = 'credit_group_commission_rules'
            """
        ).fetchone()
        assert table is not None
        new_setting = connection.execute(
            "SELECT value FROM app_preferences WHERE key = 'new_default_setting'"
        ).fetchone()[0]
        assert new_setting == "default_value"
        user_setting = connection.execute(
            "SELECT value FROM app_preferences WHERE key = 'user_custom_setting'"
        ).fetchone()[0]
        assert user_setting == "gia_tri_nguoi_dung_da_sua"
        migration = connection.execute(
            "SELECT version FROM app_schema_migrations WHERE version = '0.1.1'"
        ).fetchone()
        assert migration is not None


def assert_zip_has_no_database(path: Path) -> None:
    with zipfile.ZipFile(path) as archive:
        names = set(archive.namelist())
    assert not any(name.casefold().endswith(".db") for name in names), names
    assert not any(name.startswith(("logs/", "backups/", "temp/", "KetQua/")) for name in names)


def _base_db_migrations() -> str:
    return """
from __future__ import annotations

import sqlite3
from typing import Callable

PythonMigration = Callable[[sqlite3.Connection], None]


def table_exists(conn, table_name: str) -> bool:
    return conn.execute(
        "SELECT 1 FROM sqlite_master WHERE type = 'table' AND name = ?",
        (table_name,),
    ).fetchone() is not None


def column_exists(conn, table_name: str, column_name: str) -> bool:
    if not table_exists(conn, table_name):
        return False
    return column_name in {
        str(row[1]) for row in conn.execute(f"PRAGMA table_info({table_name})")
    }


def index_exists(conn, index_name: str) -> bool:
    return conn.execute(
        "SELECT 1 FROM sqlite_master WHERE type = 'index' AND name = ?",
        (index_name,),
    ).fetchone() is not None


def default_python_migrations() -> dict[str, PythonMigration]:
    return {
    }
"""


if __name__ == "__main__":
    raise SystemExit(main())
