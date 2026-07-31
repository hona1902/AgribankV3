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
        "0.1.6": migrate_0_1_6,
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


def migrate_0_1_6(connection: sqlite3.Connection) -> None:
    ensure_schema_migrations_table(connection)
    connection.executescript(
        """
        CREATE TABLE IF NOT EXISTS summary_import_history (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            data_type TEXT NOT NULL,
            period TEXT NOT NULL DEFAULT '',
            source_path TEXT NOT NULL DEFAULT '',
            file_name TEXT NOT NULL DEFAULT '',
            imported_by TEXT NOT NULL DEFAULT '',
            row_count INTEGER NOT NULL DEFAULT 0,
            duration_ms INTEGER NOT NULL DEFAULT 0,
            version INTEGER NOT NULL DEFAULT 1,
            status TEXT NOT NULL DEFAULT 'success',
            message TEXT NOT NULL DEFAULT '',
            source_hash TEXT NOT NULL DEFAULT '',
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL
        );
        CREATE INDEX IF NOT EXISTS idx_summary_import_history_type_period
            ON summary_import_history(data_type, period, created_at DESC);
        CREATE UNIQUE INDEX IF NOT EXISTS idx_summary_import_history_source_hash
            ON summary_import_history(data_type, period, source_hash)
            WHERE source_hash <> '' AND status = 'success';

        CREATE TABLE IF NOT EXISTS nim_details (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            batch_id INTEGER NOT NULL,
            data_type TEXT NOT NULL,
            period TEXT NOT NULL,
            branch_code TEXT NOT NULL DEFAULT '',
            branch_name TEXT NOT NULL DEFAULT '',
            trctcd TEXT NOT NULL DEFAULT '',
            transaction_office TEXT NOT NULL DEFAULT '',
            customer_type TEXT NOT NULL DEFAULT '',
            officer TEXT NOT NULL DEFAULT '',
            balance REAL NOT NULL DEFAULT 0,
            interest_rate REAL NOT NULL DEFAULT 0,
            ftp_rate REAL NOT NULL DEFAULT 0,
            adjustment_rate REAL NOT NULL DEFAULT 0,
            numerator_before REAL NOT NULL DEFAULT 0,
            numerator_after REAL NOT NULL DEFAULT 0,
            average_rate_numerator REAL NOT NULL DEFAULT 0,
            source_file TEXT NOT NULL DEFAULT '',
            created_at TEXT NOT NULL,
            FOREIGN KEY(batch_id) REFERENCES summary_import_history(id)
                ON DELETE CASCADE
        );
        CREATE INDEX IF NOT EXISTS idx_nim_details_lookup
            ON nim_details(data_type, period, branch_name, transaction_office, customer_type, officer);
        CREATE INDEX IF NOT EXISTS idx_nim_details_batch
            ON nim_details(batch_id);

        CREATE TABLE IF NOT EXISTS nim_period_summary (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            batch_id INTEGER NOT NULL,
            data_type TEXT NOT NULL,
            period TEXT NOT NULL,
            branch_code TEXT NOT NULL DEFAULT '',
            branch_name TEXT NOT NULL DEFAULT '',
            trctcd TEXT NOT NULL DEFAULT '',
            transaction_office TEXT NOT NULL DEFAULT '',
            customer_type TEXT NOT NULL DEFAULT '',
            officer_code TEXT NOT NULL DEFAULT '',
            officer_name TEXT NOT NULL DEFAULT '',
            balance REAL NOT NULL DEFAULT 0,
            interest_rate_numerator REAL NOT NULL DEFAULT 0,
            numerator_before REAL NOT NULL DEFAULT 0,
            numerator_after REAL NOT NULL DEFAULT 0,
            source_row_count INTEGER NOT NULL DEFAULT 0,
            created_at TEXT NOT NULL,
            FOREIGN KEY(batch_id) REFERENCES summary_import_history(id)
                ON DELETE CASCADE,
            UNIQUE(
                batch_id,
                data_type,
                period,
                branch_code,
                trctcd,
                customer_type,
                officer_code,
                officer_name
            )
        );
        CREATE INDEX IF NOT EXISTS idx_nim_period_summary_type_period
            ON nim_period_summary(data_type, period);
        CREATE INDEX IF NOT EXISTS idx_nim_period_summary_type_period_branch
            ON nim_period_summary(data_type, period, branch_code);
        CREATE INDEX IF NOT EXISTS idx_nim_period_summary_type_period_officer
            ON nim_period_summary(data_type, period, officer_code);
        CREATE INDEX IF NOT EXISTS idx_nim_period_summary_type_period_pgd_type
            ON nim_period_summary(data_type, period, branch_code, trctcd, customer_type);

        CREATE TABLE IF NOT EXISTS loan_compare_details (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            batch_id INTEGER NOT NULL,
            customer_code TEXT NOT NULL,
            customer_name TEXT NOT NULL DEFAULT '',
            address TEXT NOT NULL DEFAULT '',
            officer TEXT NOT NULL DEFAULT '',
            previous_balance REAL NOT NULL DEFAULT 0,
            current_balance REAL NOT NULL DEFAULT 0,
            difference REAL NOT NULL DEFAULT 0,
            category TEXT NOT NULL DEFAULT '',
            created_at TEXT NOT NULL,
            FOREIGN KEY(batch_id) REFERENCES summary_import_history(id)
                ON DELETE CASCADE
        );
        CREATE INDEX IF NOT EXISTS idx_loan_compare_details_batch
            ON loan_compare_details(batch_id, category, officer, customer_code);

        CREATE TABLE IF NOT EXISTS credit_limit_details (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            batch_id INTEGER NOT NULL,
            customer_code TEXT NOT NULL DEFAULT '',
            customer_name TEXT NOT NULL DEFAULT '',
            contract_number TEXT NOT NULL DEFAULT '',
            approved_date TEXT,
            approved_amount REAL NOT NULL DEFAULT 0,
            outstanding_balance REAL NOT NULL DEFAULT 0,
            expiry_date TEXT,
            address TEXT NOT NULL DEFAULT '',
            officer TEXT NOT NULL DEFAULT '',
            note TEXT NOT NULL DEFAULT '',
            days_to_expiry INTEGER,
            status TEXT NOT NULL DEFAULT '',
            reference_date TEXT NOT NULL DEFAULT '',
            warn_days INTEGER NOT NULL DEFAULT 30,
            min_limit REAL NOT NULL DEFAULT 0,
            created_at TEXT NOT NULL,
            FOREIGN KEY(batch_id) REFERENCES summary_import_history(id)
                ON DELETE CASCADE
        );
        CREATE INDEX IF NOT EXISTS idx_credit_limit_details_batch
            ON credit_limit_details(batch_id, status, officer, expiry_date);
        CREATE INDEX IF NOT EXISTS idx_credit_limit_details_expiry
            ON credit_limit_details(expiry_date, officer);

        CREATE TABLE IF NOT EXISTS summary_action_log (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            data_type TEXT NOT NULL DEFAULT '',
            action TEXT NOT NULL,
            target_id TEXT NOT NULL DEFAULT '',
            detail TEXT NOT NULL DEFAULT '',
            created_by TEXT NOT NULL DEFAULT '',
            created_at TEXT NOT NULL
        );
        CREATE INDEX IF NOT EXISTS idx_summary_action_log_created
            ON summary_action_log(data_type, created_at DESC);

        CREATE TABLE IF NOT EXISTS summary_query_cache (
            cache_key TEXT PRIMARY KEY,
            data_type TEXT NOT NULL DEFAULT '',
            payload TEXT NOT NULL,
            created_at TEXT NOT NULL,
            expires_at TEXT NOT NULL
        );
        CREATE INDEX IF NOT EXISTS idx_summary_query_cache_type
            ON summary_query_cache(data_type, expires_at);

        CREATE TABLE IF NOT EXISTS summary_sync_state (
            id INTEGER PRIMARY KEY CHECK(id = 1),
            device_id TEXT NOT NULL DEFAULT '',
            last_sync_at TEXT NOT NULL DEFAULT '',
            sync_enabled INTEGER NOT NULL DEFAULT 0 CHECK(sync_enabled IN (0, 1)),
            updated_at TEXT NOT NULL
        );

        CREATE TABLE IF NOT EXISTS summary_sync_outbox (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            data_type TEXT NOT NULL DEFAULT '',
            entity_table TEXT NOT NULL,
            entity_id TEXT NOT NULL,
            operation TEXT NOT NULL,
            payload TEXT NOT NULL DEFAULT '',
            created_at TEXT NOT NULL,
            synced_at TEXT NOT NULL DEFAULT ''
        );
        CREATE INDEX IF NOT EXISTS idx_summary_sync_outbox_pending
            ON summary_sync_outbox(synced_at, created_at);
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
