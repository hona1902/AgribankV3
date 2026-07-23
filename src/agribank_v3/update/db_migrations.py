from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
import hashlib
from pathlib import Path
import sqlite3
from typing import Callable, Iterable


class DatabaseMigrationError(RuntimeError):
    pass


@dataclass(frozen=True, slots=True)
class MigrationSpec:
    version: str
    file: str = ""
    description: str = ""


@dataclass(frozen=True, slots=True)
class AppliedMigration:
    version: str
    migration_name: str
    checksum: str


PythonMigration = Callable[[sqlite3.Connection], None]


def table_exists(connection: sqlite3.Connection, table_name: str) -> bool:
    return (
        connection.execute(
            "SELECT 1 FROM sqlite_master WHERE type = 'table' AND name = ?",
            (table_name,),
        ).fetchone()
        is not None
    )


def column_exists(
    connection: sqlite3.Connection,
    table_name: str,
    column_name: str,
) -> bool:
    if not table_exists(connection, table_name):
        return False
    columns = {
        str(row[1])
        for row in connection.execute(f"PRAGMA table_info({table_name})").fetchall()
    }
    return column_name in columns


def index_exists(connection: sqlite3.Connection, index_name: str) -> bool:
    return (
        connection.execute(
            "SELECT 1 FROM sqlite_master WHERE type = 'index' AND name = ?",
            (index_name,),
        ).fetchone()
        is not None
    )


def ensure_schema_migrations_table(connection: sqlite3.Connection) -> None:
    connection.execute(
        """
        CREATE TABLE IF NOT EXISTS app_schema_migrations (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            version TEXT NOT NULL,
            migration_name TEXT,
            applied_at TEXT NOT NULL,
            checksum TEXT,
            success INTEGER DEFAULT 1 CHECK(success IN (0, 1))
        )
        """
    )
    connection.execute(
        """
        CREATE UNIQUE INDEX IF NOT EXISTS idx_app_schema_migrations_version
        ON app_schema_migrations(version)
        WHERE success = 1
        """
    )


def default_python_migrations() -> dict[str, PythonMigration]:
    return {
        "0.1.1": migrate_0_1_1,
        "0.1.2": migrate_0_1_2,
    }


def migrate_0_1_1(connection: sqlite3.Connection) -> None:
    if table_exists(connection, "credit_groups") and not column_exists(
        connection,
        "credit_groups",
        "is_active",
    ):
        connection.execute(
            "ALTER TABLE credit_groups ADD COLUMN is_active INTEGER DEFAULT 1"
        )
    if column_exists(connection, "credit_groups", "is_active"):
        connection.execute(
            "UPDATE credit_groups SET is_active = 1 WHERE is_active IS NULL"
        )


def migrate_0_1_2(connection: sqlite3.Connection) -> None:
    ensure_schema_migrations_table(connection)
    connection.execute(
        """
        CREATE TABLE IF NOT EXISTS app_update_settings (
            key TEXT PRIMARY KEY,
            value TEXT NOT NULL,
            updated_at TEXT NOT NULL
        )
        """
    )


def applied_versions(connection: sqlite3.Connection) -> set[str]:
    ensure_schema_migrations_table(connection)
    return {
        str(row[0])
        for row in connection.execute(
            """
            SELECT version
            FROM app_schema_migrations
            WHERE success = 1
            """
        ).fetchall()
    }


def latest_schema_version(connection: sqlite3.Connection) -> str:
    ensure_schema_migrations_table(connection)
    row = connection.execute(
        """
        SELECT version
        FROM app_schema_migrations
        WHERE success = 1
        ORDER BY id DESC
        LIMIT 1
        """
    ).fetchone()
    return str(row[0]) if row is not None else "Chưa có migration"


def apply_migrations(
    connection: sqlite3.Connection,
    migrations: Iterable[MigrationSpec],
    *,
    update_root: Path,
    python_migrations: dict[str, PythonMigration] | None = None,
) -> list[AppliedMigration]:
    ensure_schema_migrations_table(connection)
    migration_functions = default_python_migrations()
    if python_migrations:
        migration_functions.update(python_migrations)
    applied = applied_versions(connection)
    results: list[AppliedMigration] = []
    for migration in migrations:
        version = str(migration.version).strip()
        if not version or version in applied:
            continue
        migration_name = migration.file or migration.description or version
        sql_text = ""
        checksum = ""
        migration_file = None
        if migration.file:
            migration_file = (update_root / migration.file).resolve()
            if not migration_file.is_file():
                raise DatabaseMigrationError(
                    f"Không tìm thấy file migration: {migration.file}"
                )
            sql_text = migration_file.read_text(encoding="utf-8-sig")
            checksum = _sha256(sql_text.encode("utf-8"))
        function = migration_functions.get(version)
        if function is not None:
            checksum_source = sql_text.encode("utf-8") or function.__name__.encode("utf-8")
            checksum = checksum or _sha256(checksum_source)
        elif sql_text:
            function = lambda conn, script=sql_text: conn.executescript(script)
        else:
            raise DatabaseMigrationError(
                f"Migration {version} không có SQL hoặc Python migration."
            )

        try:
            connection.execute("BEGIN")
            function(connection)
            connection.execute(
                """
                INSERT INTO app_schema_migrations(
                    version, migration_name, applied_at, checksum, success
                )
                VALUES (?, ?, ?, ?, 1)
                """,
                (version, migration_name, _now(), checksum),
            )
            connection.execute("COMMIT")
        except Exception as exc:
            try:
                connection.execute("ROLLBACK")
            except sqlite3.Error:
                pass
            raise DatabaseMigrationError(
                f"Migration {version} không thành công: {exc}"
            ) from exc
        applied.add(version)
        results.append(
            AppliedMigration(
                version=version,
                migration_name=migration_name,
                checksum=checksum,
            )
        )
    return results


def _sha256(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def _now() -> str:
    return datetime.now().astimezone().isoformat(timespec="seconds")
