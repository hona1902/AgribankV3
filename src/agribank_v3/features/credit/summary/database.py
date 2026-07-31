from __future__ import annotations

from contextlib import closing
from datetime import datetime
from pathlib import Path
import hashlib
import sqlite3

from agribank_v3.runtime_paths import application_root


CREDIT_SUMMARY_DATABASE_NAME = "CreditSummary.db"
CREDIT_SUMMARY_SCHEMA_VERSION = "0.1.6"
CREDIT_SUMMARY_SCHEMA_MIGRATION_NAME = "credit-summary-schema"
CREDIT_SUMMARY_DATA_MIGRATION_VERSION = "0.1.6-data"
CREDIT_SUMMARY_DATA_MIGRATION_NAME = "copy-summary-data-from-DuLieuV3"
CREDIT_SUMMARY_SCHEMA_CHECKSUM = hashlib.sha256(
    CREDIT_SUMMARY_SCHEMA_MIGRATION_NAME.encode("utf-8")
).hexdigest()
CREDIT_SUMMARY_DATA_CHECKSUM = hashlib.sha256(
    CREDIT_SUMMARY_DATA_MIGRATION_NAME.encode("utf-8")
).hexdigest()

SUMMARY_TABLES: tuple[str, ...] = (
    "summary_import_history",
    "nim_details",
    "nim_period_summary",
    "loan_compare_details",
    "credit_limit_details",
    "summary_action_log",
    "summary_query_cache",
    "summary_sync_state",
    "summary_sync_outbox",
)


def credit_summary_database_path(main_database_path: Path | None = None) -> Path:
    if main_database_path is None:
        return application_root() / "data" / CREDIT_SUMMARY_DATABASE_NAME
    path = Path(main_database_path)
    if path.name.casefold() == CREDIT_SUMMARY_DATABASE_NAME.casefold():
        return path
    return path.parent / CREDIT_SUMMARY_DATABASE_NAME


def get_credit_summary_connection(database_path: Path | None = None) -> sqlite3.Connection:
    path = credit_summary_database_path(database_path)
    path.parent.mkdir(parents=True, exist_ok=True)
    connection = sqlite3.connect(path, timeout=30)
    connection.row_factory = sqlite3.Row
    connection.execute("PRAGMA foreign_keys = ON")
    connection.execute("PRAGMA journal_mode = WAL")
    connection.execute("PRAGMA busy_timeout = 30000")
    return connection


def ensure_credit_summary_migration_table(connection: sqlite3.Connection) -> None:
    connection.execute(
        """
        CREATE TABLE IF NOT EXISTS credit_summary_schema_migrations (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            version TEXT NOT NULL,
            migration_name TEXT NOT NULL DEFAULT '',
            applied_at TEXT NOT NULL,
            checksum TEXT NOT NULL DEFAULT '',
            success INTEGER NOT NULL DEFAULT 1 CHECK(success IN (0, 1))
        )
        """
    )
    connection.execute(
        """
        CREATE UNIQUE INDEX IF NOT EXISTS idx_credit_summary_schema_migrations_version
        ON credit_summary_schema_migrations(version, migration_name)
        WHERE success = 1
        """
    )


def mark_credit_summary_migration(
    connection: sqlite3.Connection,
    *,
    version: str,
    migration_name: str,
    checksum: str,
    success: bool = True,
) -> None:
    ensure_credit_summary_migration_table(connection)
    existing = connection.execute(
        """
        SELECT id
        FROM credit_summary_schema_migrations
        WHERE version = ? AND migration_name = ? AND success = 1
        LIMIT 1
        """,
        (version, migration_name),
    ).fetchone()
    params = (
        version,
        migration_name,
        datetime.now().astimezone().isoformat(timespec="seconds"),
        checksum,
        1 if success else 0,
    )
    if existing is None:
        connection.execute(
            """
            INSERT INTO credit_summary_schema_migrations(
                version, migration_name, applied_at, checksum, success
            )
            VALUES (?, ?, ?, ?, ?)
            """,
            params,
        )
    else:
        connection.execute(
            """
            UPDATE credit_summary_schema_migrations
            SET applied_at = ?, checksum = ?, success = ?
            WHERE id = ?
            """,
            (params[2], params[3], params[4], int(existing["id"])),
        )


def credit_summary_migration_applied(
    connection: sqlite3.Connection,
    *,
    version: str,
    migration_name: str,
) -> bool:
    ensure_credit_summary_migration_table(connection)
    return (
        connection.execute(
            """
            SELECT 1
            FROM credit_summary_schema_migrations
            WHERE version = ? AND migration_name = ? AND success = 1
            LIMIT 1
            """,
            (version, migration_name),
        ).fetchone()
        is not None
    )


def migrate_existing_summary_data(
    source_database_path: Path,
    target_database_path: Path,
) -> dict[str, int]:
    source = Path(source_database_path)
    target = Path(target_database_path)
    if not source.is_file() or not target.is_file():
        return {}
    if source.resolve() == target.resolve():
        return {}

    with closing(sqlite3.connect(target, timeout=30)) as connection:
        connection.row_factory = sqlite3.Row
        connection.execute("PRAGMA foreign_keys = ON")
        connection.execute("PRAGMA busy_timeout = 30000")
        ensure_credit_summary_migration_table(connection)
        if credit_summary_migration_applied(
            connection,
            version=CREDIT_SUMMARY_DATA_MIGRATION_VERSION,
            migration_name=CREDIT_SUMMARY_DATA_MIGRATION_NAME,
        ):
            return {}
        copied: dict[str, int] = {}
        connection.execute("ATTACH DATABASE ? AS legacy", (str(source),))
        try:
            with connection:
                for table_name in SUMMARY_TABLES:
                    if not _table_exists(connection, "legacy", table_name):
                        continue
                    if not _table_exists(connection, "main", table_name):
                        continue
                    source_count = _table_row_count(connection, "legacy", table_name)
                    if source_count <= 0:
                        copied[table_name] = 0
                        continue
                    target_count = _table_row_count(connection, "main", table_name)
                    if target_count > 0:
                        copied[table_name] = 0
                        continue
                    columns = _common_columns(connection, table_name)
                    if not columns:
                        copied[table_name] = 0
                        continue
                    column_sql = ", ".join(_quote_identifier(column) for column in columns)
                    connection.execute(
                        f"""
                        INSERT INTO main.{_quote_identifier(table_name)} ({column_sql})
                        SELECT {column_sql}
                        FROM legacy.{_quote_identifier(table_name)}
                        """
                    )
                    copied[table_name] = source_count
                    if _table_row_count(connection, "main", table_name) < source_count:
                        raise sqlite3.DatabaseError(f"Không copy đủ dữ liệu bảng {table_name}.")
                mark_credit_summary_migration(
                    connection,
                    version=CREDIT_SUMMARY_DATA_MIGRATION_VERSION,
                    migration_name=CREDIT_SUMMARY_DATA_MIGRATION_NAME,
                    checksum=CREDIT_SUMMARY_DATA_CHECKSUM,
                )
        finally:
            connection.execute("DETACH DATABASE legacy")
    return copied


def _table_exists(connection: sqlite3.Connection, schema_name: str, table_name: str) -> bool:
    return (
        connection.execute(
            f"""
            SELECT 1
            FROM {_quote_identifier(schema_name)}.sqlite_master
            WHERE type = 'table' AND name = ?
            LIMIT 1
            """,
            (table_name,),
        ).fetchone()
        is not None
    )


def _table_row_count(connection: sqlite3.Connection, schema_name: str, table_name: str) -> int:
    row = connection.execute(
        f"SELECT COUNT(*) FROM {_quote_identifier(schema_name)}.{_quote_identifier(table_name)}"
    ).fetchone()
    return int(row[0] or 0)


def _common_columns(connection: sqlite3.Connection, table_name: str) -> list[str]:
    source_columns = _columns(connection, "legacy", table_name)
    target_columns = set(_columns(connection, "main", table_name))
    return [column for column in source_columns if column in target_columns]


def _columns(connection: sqlite3.Connection, schema_name: str, table_name: str) -> list[str]:
    rows = connection.execute(
        f"PRAGMA {_quote_identifier(schema_name)}.table_info({_quote_identifier(table_name)})"
    ).fetchall()
    return [str(row[1]) for row in rows]


def _quote_identifier(value: str) -> str:
    return '"' + str(value).replace('"', '""') + '"'
