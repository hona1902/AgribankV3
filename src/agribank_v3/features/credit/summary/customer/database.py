from __future__ import annotations

from datetime import datetime
from pathlib import Path
import hashlib
import sqlite3
import threading

from agribank_v3.runtime_paths import application_root


CUSTOMER_DATABASE_NAME = "Customer.db"
CUSTOMER_SCHEMA_VERSION = "0.1.5"
CUSTOMER_SCHEMA_MIGRATION_NAME = "customer-database-schema"
CUSTOMER_SCHEMA_CHECKSUM = hashlib.sha256(
    CUSTOMER_SCHEMA_MIGRATION_NAME.encode("utf-8")
).hexdigest()


def customer_database_path(main_database_path: Path | None = None) -> Path:
    if main_database_path is None:
        return application_root() / "data" / CUSTOMER_DATABASE_NAME
    path = Path(main_database_path)
    if path.name.casefold() == CUSTOMER_DATABASE_NAME.casefold():
        return path
    return path.parent / CUSTOMER_DATABASE_NAME


def get_customer_database_connection(database_path: Path | None = None) -> sqlite3.Connection:
    path = customer_database_path(database_path)
    path.parent.mkdir(parents=True, exist_ok=True)
    connection = sqlite3.connect(path, timeout=10)
    connection.row_factory = sqlite3.Row
    connection.execute("PRAGMA foreign_keys = ON")
    connection.execute("PRAGMA journal_mode = WAL")
    connection.execute("PRAGMA busy_timeout = 10000")
    return connection


class CustomerDatabaseOperationLock:
    _guard = threading.Lock()
    _active_operation = ""

    def __init__(self, operation: str) -> None:
        self.operation = str(operation or "bảo trì Customer.db").strip() or "bảo trì Customer.db"
        self.acquired = False

    def __enter__(self) -> "CustomerDatabaseOperationLock":
        with self._guard:
            if self.__class__._active_operation:
                raise RuntimeError(
                    "Customer.db đang chạy thao tác "
                    f"{self.__class__._active_operation}; vui lòng chờ thao tác hiện tại hoàn tất."
                )
            self.__class__._active_operation = self.operation
            self.acquired = True
        return self

    def __exit__(self, exc_type, exc, traceback) -> None:
        if not self.acquired:
            return
        with self._guard:
            if self.__class__._active_operation == self.operation:
                self.__class__._active_operation = ""
            self.acquired = False

    @classmethod
    def active_operation(cls) -> str:
        with cls._guard:
            return cls._active_operation

    @classmethod
    def assert_writable(cls) -> None:
        operation = cls.active_operation()
        if operation:
            raise RuntimeError(f"Customer.db đang {operation}; vui lòng chờ thao tác hiện tại hoàn tất.")


def ensure_customer_schema(connection: sqlite3.Connection) -> None:
    ensure_customer_migration_table(connection)
    connection.executescript(
        """
        CREATE TABLE IF NOT EXISTS customer_master (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            customer_code TEXT NOT NULL UNIQUE,
            branch_code TEXT NOT NULL DEFAULT '',
            customer_sequence TEXT NOT NULL DEFAULT '',
            customer_name TEXT NOT NULL DEFAULT '',
            customer_type TEXT NOT NULL DEFAULT '',
            latest_officer_code TEXT NOT NULL DEFAULT '',
            latest_officer_name TEXT NOT NULL DEFAULT '',
            first_seen_period TEXT NOT NULL DEFAULT '',
            last_seen_period TEXT NOT NULL DEFAULT '',
            is_active INTEGER NOT NULL DEFAULT 1 CHECK(is_active IN (0, 1)),
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL
        );

        CREATE TABLE IF NOT EXISTS customer_import_runs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            period TEXT NOT NULL DEFAULT '',
            source_folder TEXT NOT NULL DEFAULT '',
            data_type TEXT NOT NULL DEFAULT 'DN',
            file_count INTEGER NOT NULL DEFAULT 0,
            source_row_count INTEGER NOT NULL DEFAULT 0,
            customer_count INTEGER NOT NULL DEFAULT 0,
            started_at TEXT NOT NULL DEFAULT '',
            completed_at TEXT NOT NULL DEFAULT '',
            status TEXT NOT NULL DEFAULT 'PENDING',
            error_message TEXT NOT NULL DEFAULT '',
            created_by TEXT NOT NULL DEFAULT '',
            computer_name TEXT NOT NULL DEFAULT ''
        );

        CREATE TABLE IF NOT EXISTS customer_import_files (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            run_id INTEGER NOT NULL,
            file_name TEXT NOT NULL DEFAULT '',
            file_path TEXT NOT NULL DEFAULT '',
            file_hash TEXT NOT NULL DEFAULT '',
            branch_code TEXT NOT NULL DEFAULT '',
            period TEXT NOT NULL DEFAULT '',
            source_row_count INTEGER NOT NULL DEFAULT 0,
            customer_count INTEGER NOT NULL DEFAULT 0,
            status TEXT NOT NULL DEFAULT 'PENDING',
            error_message TEXT NOT NULL DEFAULT '',
            FOREIGN KEY(run_id) REFERENCES customer_import_runs(id)
                ON DELETE CASCADE
        );
        CREATE UNIQUE INDEX IF NOT EXISTS idx_customer_import_files_period_hash
            ON customer_import_files(period, file_hash)
            WHERE file_hash <> '' AND status = 'COMPLETED';

        CREATE TABLE IF NOT EXISTS customer_period_summary (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            run_id INTEGER,
            period TEXT NOT NULL,
            customer_code TEXT NOT NULL,
            branch_code TEXT NOT NULL DEFAULT '',
            customer_sequence TEXT NOT NULL DEFAULT '',
            customer_name TEXT NOT NULL DEFAULT '',
            customer_type TEXT NOT NULL DEFAULT '',
            primary_officer_code TEXT NOT NULL DEFAULT '',
            primary_officer_name TEXT NOT NULL DEFAULT '',
            officer_count INTEGER NOT NULL DEFAULT 0,
            has_multiple_officers INTEGER NOT NULL DEFAULT 0 CHECK(has_multiple_officers IN (0, 1)),
            total_balance REAL NOT NULL DEFAULT 0,
            short_term_balance REAL NOT NULL DEFAULT 0,
            medium_long_term_balance REAL NOT NULL DEFAULT 0,
            other_balance REAL NOT NULL DEFAULT 0,
            medium_long_ratio REAL,
            interest_rate_numerator REAL NOT NULL DEFAULT 0,
            nim_before_numerator REAL NOT NULL DEFAULT 0,
            nim_after_numerator REAL NOT NULL DEFAULT 0,
            average_rate REAL,
            nim_before REAL,
            nim_after REAL,
            source_loan_count INTEGER NOT NULL DEFAULT 0,
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL,
            FOREIGN KEY(run_id) REFERENCES customer_import_runs(id)
                ON DELETE SET NULL,
            UNIQUE(period, customer_code)
        );
        CREATE INDEX IF NOT EXISTS idx_customer_period_summary_period
            ON customer_period_summary(period);
        CREATE INDEX IF NOT EXISTS idx_customer_period_summary_branch_period
            ON customer_period_summary(branch_code, period);
        CREATE INDEX IF NOT EXISTS idx_customer_period_summary_officer_period
            ON customer_period_summary(primary_officer_code, period);
        CREATE INDEX IF NOT EXISTS idx_customer_period_summary_type_period
            ON customer_period_summary(customer_type, period);
        CREATE INDEX IF NOT EXISTS idx_customer_period_summary_code_period
            ON customer_period_summary(customer_code, period);
        CREATE INDEX IF NOT EXISTS idx_customer_period_summary_period_sequence_branch
            ON customer_period_summary(period, customer_sequence, branch_code);
        CREATE INDEX IF NOT EXISTS idx_customer_period_summary_period_officer_sequence
            ON customer_period_summary(period, primary_officer_code, customer_sequence);

        CREATE TABLE IF NOT EXISTS customer_officer_period (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            period TEXT NOT NULL,
            customer_code TEXT NOT NULL,
            officer_code TEXT NOT NULL DEFAULT '',
            officer_name TEXT NOT NULL DEFAULT '',
            branch_code TEXT NOT NULL DEFAULT '',
            transaction_office TEXT NOT NULL DEFAULT '',
            balance_managed REAL NOT NULL DEFAULT 0,
            short_term_balance REAL NOT NULL DEFAULT 0,
            medium_long_term_balance REAL NOT NULL DEFAULT 0,
            other_balance REAL NOT NULL DEFAULT 0,
            source_loan_count INTEGER NOT NULL DEFAULT 0,
            interest_rate_numerator REAL NOT NULL DEFAULT 0,
            nim_before_numerator REAL NOT NULL DEFAULT 0,
            nim_after_numerator REAL NOT NULL DEFAULT 0,
            is_primary INTEGER NOT NULL DEFAULT 0 CHECK(is_primary IN (0, 1)),
            created_at TEXT NOT NULL,
            UNIQUE(period, customer_code, officer_code, officer_name)
        );
        CREATE INDEX IF NOT EXISTS idx_customer_officer_period_customer_period
            ON customer_officer_period(customer_code, period);
        CREATE INDEX IF NOT EXISTS idx_customer_officer_period_officer_period
            ON customer_officer_period(officer_code, period);
        CREATE INDEX IF NOT EXISTS idx_customer_officer_period_period_primary
            ON customer_officer_period(period, is_primary);

        CREATE TABLE IF NOT EXISTS customer_office_period (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            run_id INTEGER,
            period TEXT NOT NULL,
            customer_code TEXT NOT NULL,
            customer_sequence TEXT NOT NULL DEFAULT '',
            branch_code TEXT NOT NULL DEFAULT '',
            trctcd TEXT NOT NULL DEFAULT '',
            office_code TEXT NOT NULL,
            office_name TEXT NOT NULL DEFAULT '',
            office_type TEXT NOT NULL DEFAULT 'UNKNOWN',
            primary_officer_code TEXT NOT NULL DEFAULT '',
            primary_officer_name TEXT NOT NULL DEFAULT '',
            officer_count INTEGER NOT NULL DEFAULT 0,
            total_balance REAL NOT NULL DEFAULT 0,
            short_term_balance REAL NOT NULL DEFAULT 0,
            medium_long_term_balance REAL NOT NULL DEFAULT 0,
            other_balance REAL NOT NULL DEFAULT 0,
            interest_rate_numerator REAL NOT NULL DEFAULT 0,
            nim_before_numerator REAL NOT NULL DEFAULT 0,
            nim_after_numerator REAL NOT NULL DEFAULT 0,
            source_loan_count INTEGER NOT NULL DEFAULT 0,
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL,
            FOREIGN KEY(run_id) REFERENCES customer_import_runs(id)
                ON DELETE SET NULL,
            UNIQUE(period, customer_code, office_code)
        );
        CREATE INDEX IF NOT EXISTS idx_customer_office_period_customer_period
            ON customer_office_period(customer_sequence, period);
        CREATE INDEX IF NOT EXISTS idx_customer_office_period_period_customer_branch_trctcd
            ON customer_office_period(period, customer_sequence, branch_code, trctcd);
        CREATE INDEX IF NOT EXISTS idx_customer_office_period_period_branch_office
            ON customer_office_period(period, branch_code, office_code);
        CREATE INDEX IF NOT EXISTS idx_customer_office_period_period_customer_branch_office
            ON customer_office_period(period, customer_sequence, branch_code, office_code);
        CREATE INDEX IF NOT EXISTS idx_customer_office_period_period_branch_type_sequence
            ON customer_office_period(period, branch_code, office_type, customer_sequence);

        CREATE TABLE IF NOT EXISTS customer_officer_override (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            customer_code TEXT NOT NULL,
            effective_from_period TEXT NOT NULL DEFAULT '',
            effective_to_period TEXT NOT NULL DEFAULT '',
            officer_code TEXT NOT NULL DEFAULT '',
            officer_name TEXT NOT NULL DEFAULT '',
            reason TEXT NOT NULL DEFAULT '',
            is_active INTEGER NOT NULL DEFAULT 1 CHECK(is_active IN (0, 1)),
            created_by TEXT NOT NULL DEFAULT '',
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL
        );
        CREATE INDEX IF NOT EXISTS idx_customer_officer_override_customer_effective
            ON customer_officer_override(customer_code, effective_from_period, is_active);

        CREATE TABLE IF NOT EXISTS customer_action_log (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            action_type TEXT NOT NULL DEFAULT '',
            customer_code TEXT NOT NULL DEFAULT '',
            period TEXT NOT NULL DEFAULT '',
            old_value TEXT NOT NULL DEFAULT '',
            new_value TEXT NOT NULL DEFAULT '',
            reason TEXT NOT NULL DEFAULT '',
            user_name TEXT NOT NULL DEFAULT '',
            computer_name TEXT NOT NULL DEFAULT '',
            created_at TEXT NOT NULL
        );
        CREATE INDEX IF NOT EXISTS idx_customer_action_log_customer_period
            ON customer_action_log(customer_code, period, created_at DESC);

        CREATE TABLE IF NOT EXISTS customer_officer_directory (
            officer_code TEXT PRIMARY KEY,
            officer_name TEXT NOT NULL DEFAULT '',
            branch_code TEXT NOT NULL DEFAULT '',
            transaction_office TEXT NOT NULL DEFAULT '',
            is_active INTEGER NOT NULL DEFAULT 1 CHECK(is_active IN (0, 1)),
            updated_at TEXT NOT NULL
        );
        """
    )
    _ensure_import_run_columns(connection)
    _ensure_officer_period_context_columns(connection)
    _ensure_debt_group_columns(connection)
    _ensure_debt_group_indexes(connection)
    mark_customer_migration(
        connection,
        version=CUSTOMER_SCHEMA_VERSION,
        migration_name=CUSTOMER_SCHEMA_MIGRATION_NAME,
        checksum=CUSTOMER_SCHEMA_CHECKSUM,
    )


def _ensure_import_run_columns(connection: sqlite3.Connection) -> None:
    columns = {
        "personal_customer_count": "INTEGER NOT NULL DEFAULT 0",
        "organization_customer_count": "INTEGER NOT NULL DEFAULT 0",
        "total_balance": "REAL NOT NULL DEFAULT 0",
        "short_term_balance": "REAL NOT NULL DEFAULT 0",
        "medium_long_term_balance": "REAL NOT NULL DEFAULT 0",
        "other_balance": "REAL NOT NULL DEFAULT 0",
        "multiple_officer_customer_count": "INTEGER NOT NULL DEFAULT 0",
        "unknown_ftp_code_count": "INTEGER NOT NULL DEFAULT 0",
        "invalid_row_count": "INTEGER NOT NULL DEFAULT 0",
        "warning_count": "INTEGER NOT NULL DEFAULT 0",
        "duration_ms": "INTEGER NOT NULL DEFAULT 0",
        "debt_group_valid_row_count": "INTEGER NOT NULL DEFAULT 0",
        "debt_group_1_row_count": "INTEGER NOT NULL DEFAULT 0",
        "debt_group_2_row_count": "INTEGER NOT NULL DEFAULT 0",
        "debt_group_3_row_count": "INTEGER NOT NULL DEFAULT 0",
        "debt_group_4_row_count": "INTEGER NOT NULL DEFAULT 0",
        "debt_group_5_row_count": "INTEGER NOT NULL DEFAULT 0",
        "debt_group_unknown_row_count": "INTEGER NOT NULL DEFAULT 0",
        "debt_group_invalid_samples": "TEXT NOT NULL DEFAULT ''",
    }
    for column_name, definition in columns.items():
        if not _column_exists(connection, "customer_import_runs", column_name):
            connection.execute(
                f"ALTER TABLE customer_import_runs ADD COLUMN {column_name} {definition}"
            )


def _ensure_officer_period_context_columns(connection: sqlite3.Connection) -> None:
    columns = {
        "branch_code": "TEXT NOT NULL DEFAULT ''",
        "transaction_office": "TEXT NOT NULL DEFAULT ''",
        "short_term_balance": "REAL NOT NULL DEFAULT 0",
        "medium_long_term_balance": "REAL NOT NULL DEFAULT 0",
        "other_balance": "REAL NOT NULL DEFAULT 0",
    }
    for column_name, definition in columns.items():
        if not _column_exists(connection, "customer_officer_period", column_name):
            connection.execute(
                f"ALTER TABLE customer_officer_period ADD COLUMN {column_name} {definition}"
            )


def _ensure_debt_group_columns(connection: sqlite3.Connection) -> None:
    for table_name in ("customer_period_summary", "customer_officer_period", "customer_office_period"):
        for column_name, definition in _debt_group_column_definitions().items():
            if not _column_exists(connection, table_name, column_name):
                connection.execute(f"ALTER TABLE {table_name} ADD COLUMN {column_name} {definition}")


def _ensure_debt_group_indexes(connection: sqlite3.Connection) -> None:
    connection.executescript(
        """
        CREATE INDEX IF NOT EXISTS idx_customer_period_summary_period_branch_worst_debt
            ON customer_period_summary(period, branch_code, worst_debt_group);
        CREATE INDEX IF NOT EXISTS idx_customer_period_summary_period_officer_worst_debt
            ON customer_period_summary(period, primary_officer_code, worst_debt_group);
        CREATE INDEX IF NOT EXISTS idx_customer_officer_period_period_officer_branch
            ON customer_officer_period(period, officer_code, branch_code);
        """
    )


def _debt_group_column_definitions() -> dict[str, str]:
    columns: dict[str, str] = {
        "has_debt_group_data": "INTEGER NOT NULL DEFAULT 0",
        "worst_debt_group": "TEXT NOT NULL DEFAULT ''",
        "debt_group_unknown_row_count": "INTEGER NOT NULL DEFAULT 0",
    }
    for suffix in ("1", "2", "3", "4", "5", "unknown"):
        columns[f"debt_group_{suffix}_balance"] = "REAL NOT NULL DEFAULT 0"
    for suffix in ("1", "2", "3", "4", "5", "unknown"):
        columns[f"debt_group_{suffix}_interest_numerator"] = "REAL NOT NULL DEFAULT 0"
    for suffix in ("1", "2", "3", "4", "5", "unknown"):
        columns[f"debt_group_{suffix}_nim_before_numerator"] = "REAL NOT NULL DEFAULT 0"
    for suffix in ("1", "2", "3", "4", "5", "unknown"):
        columns[f"debt_group_{suffix}_nim_after_numerator"] = "REAL NOT NULL DEFAULT 0"
    return columns


def _column_exists(connection: sqlite3.Connection, table_name: str, column_name: str) -> bool:
    return any(
        str(row[1]) == column_name
        for row in connection.execute(f"PRAGMA table_info({table_name})").fetchall()
    )


def ensure_customer_migration_table(connection: sqlite3.Connection) -> None:
    connection.execute(
        """
        CREATE TABLE IF NOT EXISTS customer_schema_migrations (
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
        CREATE UNIQUE INDEX IF NOT EXISTS idx_customer_schema_migrations_version
            ON customer_schema_migrations(version, migration_name)
            WHERE success = 1
        """
    )


def mark_customer_migration(
    connection: sqlite3.Connection,
    *,
    version: str,
    migration_name: str,
    checksum: str,
    success: bool = True,
) -> None:
    ensure_customer_migration_table(connection)
    existing = connection.execute(
        """
        SELECT id
        FROM customer_schema_migrations
        WHERE version = ? AND migration_name = ? AND success = 1
        LIMIT 1
        """,
        (version, migration_name),
    ).fetchone()
    applied_at = datetime.now().astimezone().isoformat(timespec="seconds")
    if existing is None:
        connection.execute(
            """
            INSERT INTO customer_schema_migrations(
                version, migration_name, applied_at, checksum, success
            )
            VALUES (?, ?, ?, ?, ?)
            """,
            (version, migration_name, applied_at, checksum, 1 if success else 0),
        )
        return
    connection.execute(
        """
        UPDATE customer_schema_migrations
        SET applied_at = ?, checksum = ?, success = ?
        WHERE id = ?
        """,
        (applied_at, checksum, 1 if success else 0, int(existing["id"])),
    )
