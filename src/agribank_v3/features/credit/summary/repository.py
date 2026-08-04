from __future__ import annotations

from contextlib import closing, contextmanager
from dataclasses import asdict
from datetime import date, datetime, timedelta
from pathlib import Path
import hashlib
import json
import os
import shutil
import sqlite3
import tempfile
from typing import Iterator, Sequence
from uuid import uuid4
import zipfile

from agribank_v3.features.credit.summary.models import (
    CreditLimitRow,
    DashboardData,
    DashboardMetric,
    ImportBatch,
    LoanSnapshotRow,
    NimRow,
    PageResult,
    SummaryDataType,
    SummaryError,
    now_text,
)
from agribank_v3.features.credit.summary.credit_limit import (
    CreditLimitBatchMetadata,
    CreditLimitExcelBatchStore,
    CreditLimitStorageStatus,
)
from agribank_v3.features.credit.summary.database import (
    CREDIT_SUMMARY_DATABASE_NAME,
    CREDIT_SUMMARY_SCHEMA_CHECKSUM,
    CREDIT_SUMMARY_SCHEMA_MIGRATION_NAME,
    CREDIT_SUMMARY_SCHEMA_VERSION,
    credit_summary_database_path,
    ensure_credit_summary_migration_table,
    get_credit_summary_connection,
    mark_credit_summary_migration,
    migrate_existing_summary_data,
)
from agribank_v3.features.settings.unit_directory.service import get_unit_directory_service
from agribank_v3.runtime_paths import application_root
from agribank_v3.update.db_migrations import MigrationSpec, apply_migrations, column_exists


NIM_PERIOD_SUMMARY_TABLE = "nim_period_summary"
NIM_SUMMARY_MIGRATION_VERSION = "0.1.7"
NIM_SUMMARY_MIGRATION_NAME = "nim-period-summary-from-raw"
NIM_SUMMARY_MIGRATION_CHECKSUM = "nim-period-summary-v1"
NIM_OFFICER_DISPLAY_SQL = (
    "CASE WHEN officer_code <> '' "
    "THEN '[' || officer_code || '] ' || officer_name "
    "ELSE officer_name END"
)
NIM_PARSED_OFFICER_CODE_SQL = (
    "CASE WHEN officer LIKE '[%]%' AND instr(officer, ']') > 1 "
    "THEN substr(officer, 2, instr(officer, ']') - 2) "
    "ELSE '' END"
)
NIM_PARSED_OFFICER_NAME_SQL = (
    "CASE WHEN officer LIKE '[%]%' AND instr(officer, ']') > 1 "
    "THEN trim(substr(officer, instr(officer, ']') + 1)) "
    "ELSE officer END"
)


class SummaryRepository:
    """SQLite gateway for the 5491-THSoLieu migration."""

    CACHE_SECONDS = 20

    def __init__(self, database_path: Path) -> None:
        self.main_database_path = Path(database_path)
        self.database_path = credit_summary_database_path(self.main_database_path)
        self.database_path.parent.mkdir(parents=True, exist_ok=True)
        self._memory_cache: dict[str, object] = {}
        self.unit_directory = get_unit_directory_service(self.main_database_path)
        self.credit_limit_store = CreditLimitExcelBatchStore(self.main_database_path)
        self.ensure_schema()
        if self.main_database_path.resolve() != self.database_path.resolve():
            try:
                copied = migrate_existing_summary_data(self.main_database_path, self.database_path)
                if copied:
                    self.ensure_schema()
            except Exception as exc:
                raise SummaryError(f"Không thể chuyển dữ liệu Tổng hợp số liệu sang {CREDIT_SUMMARY_DATABASE_NAME}: {exc}") from exc

    def connect(self) -> sqlite3.Connection:
        return get_credit_summary_connection(self.database_path)

    @contextmanager
    def _database(self) -> Iterator[sqlite3.Connection]:
        database = self.connect()
        try:
            with database:
                yield database
        finally:
            database.close()

    def ensure_schema(self) -> None:
        try:
            with closing(self.connect()) as database:
                apply_migrations(
                    database,
                    (MigrationSpec(version="0.1.6", description="5491-THSoLieu SQLite schema"),),
                    update_root=application_root(),
                )
                self._ensure_summary_extensions(database)
                ensure_credit_summary_migration_table(database)
                mark_credit_summary_migration(
                    database,
                    version=CREDIT_SUMMARY_SCHEMA_VERSION,
                    migration_name=CREDIT_SUMMARY_SCHEMA_MIGRATION_NAME,
                    checksum=CREDIT_SUMMARY_SCHEMA_CHECKSUM,
                )
                database.commit()
        except Exception as exc:
            raise SummaryError(f"Không thể khởi tạo SQLite cho Tổng hợp số liệu: {exc}") from exc

    def _ensure_summary_extensions(self, database: sqlite3.Connection) -> None:
        if not column_exists(database, "summary_import_history", "source_hash"):
            database.execute(
                "ALTER TABLE summary_import_history ADD COLUMN source_hash TEXT NOT NULL DEFAULT ''"
            )
        self._ensure_nim_period_summary_schema(database)
        duplicates = database.execute(
            """
            SELECT 1
            FROM summary_import_history
            WHERE source_hash <> '' AND status = 'success'
            GROUP BY data_type, period, source_hash
            HAVING COUNT(*) > 1
            LIMIT 1
            """
        ).fetchone()
        if duplicates is None:
            database.execute("DROP INDEX IF EXISTS idx_summary_import_history_source_hash")
            database.execute(
                """
                CREATE UNIQUE INDEX IF NOT EXISTS idx_summary_import_history_period_source_hash
                    ON summary_import_history(data_type, period, source_hash)
                    WHERE source_hash <> '' AND status = 'success'
                """
            )
        self._migrate_raw_nim_to_period_summary(database)

    def _ensure_nim_period_summary_schema(self, database: sqlite3.Connection) -> None:
        database.executescript(
            """
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
            """
        )

    def _migrate_raw_nim_to_period_summary(self, database: sqlite3.Connection) -> None:
        ensure_credit_summary_migration_table(database)
        migration_done = (
            database.execute(
                """
                SELECT 1
                FROM credit_summary_schema_migrations
                WHERE version = ? AND migration_name = ? AND success = 1
                LIMIT 1
                """,
                (NIM_SUMMARY_MIGRATION_VERSION, NIM_SUMMARY_MIGRATION_NAME),
            ).fetchone()
            is not None
        )
        raw_count = self._table_row_count(database, "nim_details")
        if raw_count <= 0:
            mark_credit_summary_migration(
                database,
                version=NIM_SUMMARY_MIGRATION_VERSION,
                migration_name=NIM_SUMMARY_MIGRATION_NAME,
                checksum=NIM_SUMMARY_MIGRATION_CHECKSUM,
            )
            return
        if migration_done and self._table_row_count(database, NIM_PERIOD_SUMMARY_TABLE) > 0:
            verification = self.verify_nim_summary_totals(database=database)
            if verification["matches"]:
                return
        self._backup_before_nim_summary_migration()
        database.execute(
            f"""
            INSERT OR IGNORE INTO nim_period_summary(
                batch_id,
                data_type,
                period,
                branch_code,
                branch_name,
                trctcd,
                transaction_office,
                customer_type,
                officer_code,
                officer_name,
                balance,
                interest_rate_numerator,
                numerator_before,
                numerator_after,
                source_row_count,
                created_at
            )
            SELECT
                batch_id,
                data_type,
                period,
                branch_code,
                branch_name,
                trctcd,
                transaction_office,
                customer_type,
                {NIM_PARSED_OFFICER_CODE_SQL} AS officer_code,
                {NIM_PARSED_OFFICER_NAME_SQL} AS officer_name,
                SUM(balance) AS balance,
                SUM(average_rate_numerator) AS interest_rate_numerator,
                SUM(numerator_before) AS numerator_before,
                SUM(numerator_after) AS numerator_after,
                COUNT(*) AS source_row_count,
                MIN(created_at) AS created_at
            FROM nim_details
            GROUP BY
                batch_id,
                data_type,
                period,
                branch_code,
                branch_name,
                trctcd,
                transaction_office,
                customer_type,
                officer_code,
                officer_name
            """
        )
        verification = self.verify_nim_summary_totals(database=database)
        if not verification["matches"]:
            raise SummaryError(f"Migration NIM sang bảng tổng hợp chưa khớp tổng: {verification}")
        mark_credit_summary_migration(
            database,
            version=NIM_SUMMARY_MIGRATION_VERSION,
            migration_name=NIM_SUMMARY_MIGRATION_NAME,
            checksum=NIM_SUMMARY_MIGRATION_CHECKSUM,
        )

    def create_batch(
        self,
        data_type: SummaryDataType,
        *,
        period: str,
        source_path: Path,
        imported_by: str,
        row_count: int,
        duration_ms: int,
        status: str = "success",
        message: str = "",
        source_hash: str = "",
    ) -> int:
        now = now_text()
        source_hash = str(source_hash or "").strip()
        with self._database() as database:
            if source_hash and status == "success":
                duplicate = database.execute(
                    """
                    SELECT id, period, file_name, created_at
                    FROM summary_import_history
                    WHERE data_type = ? AND period = ? AND source_hash = ? AND status = 'success'
                    LIMIT 1
                    """,
                    (data_type.value, period, source_hash),
                ).fetchone()
                if duplicate is not None:
                    raise SummaryError(
                        "File đã được import trước đó: "
                        f"{duplicate['file_name']} kỳ {duplicate['period']} "
                        f"(batch {duplicate['id']})."
                    )
            cursor = database.execute(
                """
                INSERT INTO summary_import_history(
                    data_type, period, source_path, file_name, imported_by,
                    row_count, duration_ms, version, status, message, source_hash,
                    created_at, updated_at
                )
                VALUES (?, ?, ?, ?, ?, ?, ?, 1, ?, ?, ?, ?, ?)
                """,
                (
                    data_type.value,
                    period,
                    str(source_path),
                    source_path.name,
                    imported_by,
                    int(row_count),
                    int(duration_ms),
                    status,
                    message,
                    source_hash,
                    now,
                    now,
                ),
            )
            batch_id = int(cursor.lastrowid)
            self.log_action(
                database,
                data_type=data_type,
                action="import",
                target_id=str(batch_id),
                detail=message or source_path.name,
                created_by=imported_by,
            )
            self._enqueue_sync(
                database,
                data_type=data_type,
                entity_table="summary_import_history",
                entity_id=str(batch_id),
                operation="insert",
                payload={"period": period, "file": source_path.name, "rows": row_count},
            )
        self.clear_cache()
        return batch_id

    def save_nim_rows(self, rows: Sequence[NimRow]) -> None:
        if not rows:
            return
        summary_rows = self._aggregate_nim_rows_for_storage(rows)
        now = now_text()
        with self._database() as database:
            database.executemany(
                """
                INSERT INTO nim_period_summary(
                    batch_id, data_type, period, branch_code, branch_name, trctcd,
                    transaction_office, customer_type, officer_code, officer_name,
                    balance, interest_rate_numerator, numerator_before, numerator_after,
                    source_row_count, created_at
                )
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                ON CONFLICT(
                    batch_id,
                    data_type,
                    period,
                    branch_code,
                    trctcd,
                    customer_type,
                    officer_code,
                    officer_name
                ) DO UPDATE SET
                    branch_name = excluded.branch_name,
                    transaction_office = excluded.transaction_office,
                    balance = excluded.balance,
                    interest_rate_numerator = excluded.interest_rate_numerator,
                    numerator_before = excluded.numerator_before,
                    numerator_after = excluded.numerator_after,
                    source_row_count = excluded.source_row_count
                """,
                [
                    (
                        row.batch_id,
                        row.data_type.value,
                        row.period,
                        row.branch_code,
                        row.branch_name,
                        row.trctcd,
                        row.transaction_office,
                        row.customer_type,
                        _split_nim_officer(row.officer)[0],
                        _split_nim_officer(row.officer)[1],
                        row.balance,
                        row.average_rate_numerator,
                        row.numerator_before,
                        row.numerator_after,
                        int(row.source_row_count or 0),
                        now,
                    )
                    for row in summary_rows
                ],
            )
        self.clear_cache()

    def save_loan_compare_rows(self, batch_id: int, rows: Sequence[LoanSnapshotRow]) -> None:
        now = now_text()
        with self._database() as database:
            database.executemany(
                """
                INSERT INTO loan_compare_details(
                    batch_id, customer_code, customer_name, address, officer,
                    previous_balance, current_balance, difference, category, created_at
                )
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                [
                    (
                        batch_id,
                        row.customer_code,
                        row.customer_name,
                        row.address,
                        row.officer,
                        row.previous_balance,
                        row.current_balance,
                        row.current_balance - row.previous_balance,
                        row.category,
                        now,
                    )
                    for row in rows
                ],
            )
        self.clear_cache()

    def save_credit_limit_rows(
        self,
        batch_id: int,
        rows: Sequence[CreditLimitRow],
        *,
        reference_date: str,
        warn_days: int,
        min_limit: float,
    ) -> None:
        now = now_text()
        with self._database() as database:
            database.executemany(
                """
                INSERT INTO credit_limit_details(
                    batch_id, customer_code, customer_name, contract_number, approved_date,
                    approved_amount, outstanding_balance, expiry_date, address, officer,
                    note, days_to_expiry, status, reference_date, warn_days, min_limit, created_at
                )
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                [
                    (
                        batch_id,
                        row.customer_code,
                        row.customer_name,
                        row.contract_number,
                        _date_text(row.approved_date),
                        row.approved_amount,
                        row.outstanding_balance,
                        _date_text(row.expiry_date),
                        row.address,
                        row.officer,
                        row.note,
                        row.days_to_expiry,
                        row.status,
                        reference_date,
                        warn_days,
                        min_limit,
                        now,
                    )
                    for row in rows
                ],
            )
        self.clear_cache()

    def save_credit_limit_excel_batch(
        self,
        *,
        source_path: Path,
        rows: Sequence[CreditLimitRow],
        accepted_rows: Sequence[CreditLimitRow],
        reference_date: date | str,
        warn_days: int,
        min_limit: float,
        source_file_sha256: str,
        source_file_size: int,
        imported_by: str = "",
        source_row_count: int = 0,
        duplicate_policy: str = "error",
        progress=None,
    ) -> CreditLimitBatchMetadata:
        metadata = self.credit_limit_store.create_batch(
            source_path=Path(source_path),
            rows=rows,
            accepted_rows=accepted_rows,
            reference_date=_coerce_date(reference_date) or date.today(),
            min_limit=float(min_limit or 0),
            warn_days=int(warn_days),
            source_file_sha256=source_file_sha256,
            source_file_size=int(source_file_size or 0),
            imported_by=imported_by,
            source_row_count=int(source_row_count or len(rows)),
            duplicate_policy=duplicate_policy,
            progress=progress,
        )
        self._memory_cache.clear()
        return metadata

    def list_credit_limit_batches(self) -> list[CreditLimitBatchMetadata]:
        return self.credit_limit_store.list_batches()

    def credit_limit_duplicate_batches(self, source_path: Path) -> list[CreditLimitBatchMetadata]:
        digest = _file_hash(Path(source_path))
        return self.credit_limit_store.find_duplicate_by_sha(digest)

    def delete_credit_limit_batch(self, batch_id: str | int) -> Path | dict[str, object]:
        if _looks_like_int(batch_id):
            return self.delete_batch(int(batch_id))
        deleted_path = self.credit_limit_store.delete_batch(batch_id)
        self._memory_cache.clear()
        return deleted_path

    def backup_credit_limit_storage(self, destination: Path | None = None, *, progress=None) -> Path:
        return self.credit_limit_store.backup_storage(destination, progress=progress)

    def restore_credit_limit_storage(
        self,
        source_path: Path,
        *,
        conflict_policy: str = "skip",
        progress=None,
    ) -> dict[str, int]:
        result = self.credit_limit_store.restore_storage(
            source_path,
            conflict_policy=conflict_policy,
            progress=progress,
        )
        self._memory_cache.clear()
        return result

    def credit_limit_storage_status(self) -> CreditLimitStorageStatus:
        return self.credit_limit_store.maintenance_status()

    def migrate_legacy_credit_limit_batches(self, *, progress=None) -> dict[str, int]:
        migrated = skipped = 0
        with self._database() as database:
            batches = database.execute(
                """
                SELECT *
                FROM summary_import_history
                WHERE data_type = ?
                ORDER BY id
                """,
                (SummaryDataType.CREDIT_LIMIT.value,),
            ).fetchall()
            for batch in batches:
                source_hash = str(batch["source_hash"] or "").strip()
                if not source_hash:
                    source_hash = _legacy_credit_limit_hash(int(batch["id"]))
                if self.credit_limit_store.find_duplicate_by_sha(source_hash):
                    skipped += 1
                    continue
                detail_rows = database.execute(
                    """
                    SELECT *
                    FROM credit_limit_details
                    WHERE batch_id = ?
                    ORDER BY expiry_date, approved_amount DESC, customer_code COLLATE NOCASE
                    """,
                    (int(batch["id"]),),
                ).fetchall()
                if not detail_rows:
                    skipped += 1
                    continue
                first = detail_rows[0]
                rows = [
                    CreditLimitRow(
                        customer_code=str(item["customer_code"] or ""),
                        customer_name=str(item["customer_name"] or ""),
                        contract_number=str(item["contract_number"] or ""),
                        approved_date=_coerce_date(item["approved_date"]),
                        approved_amount=float(item["approved_amount"] or 0),
                        outstanding_balance=float(item["outstanding_balance"] or 0),
                        expiry_date=_coerce_date(item["expiry_date"]),
                        address=str(item["address"] or ""),
                        officer=str(item["officer"] or ""),
                        note="",
                        days_to_expiry=None,
                        status="",
                    )
                    for item in detail_rows
                ]
                accepted_rows = [
                    CreditLimitRow(
                        customer_code=str(item["customer_code"] or ""),
                        customer_name=str(item["customer_name"] or ""),
                        contract_number=str(item["contract_number"] or ""),
                        approved_date=_coerce_date(item["approved_date"]),
                        approved_amount=float(item["approved_amount"] or 0),
                        outstanding_balance=float(item["outstanding_balance"] or 0),
                        expiry_date=_coerce_date(item["expiry_date"]),
                        address=str(item["address"] or ""),
                        officer=str(item["officer"] or ""),
                        note=str(item["note"] or ""),
                        days_to_expiry=int(item["days_to_expiry"]) if item["days_to_expiry"] is not None else None,
                        status=str(item["status"] or ""),
                    )
                    for item in detail_rows
                ]
                if progress:
                    progress(f"Đang chuyển batch legacy {batch['id']} sang file Excel...")
                self.credit_limit_store.create_batch(
                    source_path=Path(str(batch["source_path"] or batch["file_name"] or f"legacy_{batch['id']}.csv")),
                    rows=rows,
                    accepted_rows=accepted_rows,
                    reference_date=_coerce_date(first["reference_date"]) or _coerce_date(batch["period"]) or date.today(),
                    min_limit=float(first["min_limit"] or 0),
                    warn_days=int(first["warn_days"] or 30),
                    source_file_sha256=source_hash,
                    source_file_size=0,
                    imported_by=str(batch["imported_by"] or ""),
                    source_row_count=int(batch["row_count"] or len(rows)),
                    duplicate_policy="new",
                    progress=progress,
                )
                migrated += 1
        self._memory_cache.clear()
        return {"migrated": migrated, "skipped": skipped}

    def list_batches(self, data_type: SummaryDataType | None = None, limit: int = 80) -> list[ImportBatch]:
        sql = "SELECT * FROM summary_import_history"
        params: list[object] = []
        if data_type is not None:
            sql += " WHERE data_type = ?"
            params.append(data_type.value)
        sql += " ORDER BY created_at DESC, id DESC LIMIT ?"
        params.append(int(limit))
        with self._database() as database:
            rows = database.execute(sql, params).fetchall()
        return [self._batch_from_row(row) for row in rows]

    def nim_periods(self, data_type: SummaryDataType) -> list[str]:
        if data_type not in {SummaryDataType.NIM_DN, SummaryDataType.NIM_NV}:
            raise SummaryError("Loại dữ liệu NIM không hợp lệ.")
        return self.distinct_values(NIM_PERIOD_SUMMARY_TABLE, "period", data_type=data_type)

    def nim_period_info(self, data_type: SummaryDataType, period: str) -> dict[str, object]:
        if data_type not in {SummaryDataType.NIM_DN, SummaryDataType.NIM_NV}:
            raise SummaryError("Loại dữ liệu NIM không hợp lệ.")
        clean_period = str(period or "").strip()
        if not clean_period:
            raise SummaryError("Chưa chọn kỳ dữ liệu NIM.")
        with self._database() as database:
            return self._nim_period_info(database, data_type, clean_period)

    def delete_nim_period(
        self,
        data_type: SummaryDataType,
        period: str,
        *,
        created_by: str = "",
    ) -> dict[str, object]:
        if data_type not in {SummaryDataType.NIM_DN, SummaryDataType.NIM_NV}:
            raise SummaryError("Chỉ hỗ trợ xóa kỳ dữ liệu NIM Dư nợ/Nguồn vốn.")
        clean_period = str(period or "").strip()
        if not clean_period:
            raise SummaryError("Chưa chọn kỳ dữ liệu cần xóa.")
        with self._database() as database:
            info = self._nim_period_info(database, data_type, clean_period)
            detail_rows = int(info["row_count"])
            batch_rows = int(info["batch_count"])
            if detail_rows <= 0 and batch_rows <= 0:
                raise SummaryError(f"Không có dữ liệu {data_type.value} kỳ {clean_period} để xóa.")
            batch_ids = [
                str(row["id"])
                for row in database.execute(
                    """
                    SELECT id
                    FROM summary_import_history
                    WHERE data_type = ? AND period = ?
                    """,
                    (data_type.value, clean_period),
                ).fetchall()
            ]
            database.execute(
                """
                DELETE FROM nim_period_summary
                WHERE data_type = ? AND period = ?
                """,
                (data_type.value, clean_period),
            )
            database.execute(
                """
                DELETE FROM summary_import_history
                WHERE data_type = ? AND period = ?
                """,
                (data_type.value, clean_period),
            )
            if batch_ids:
                placeholders = ", ".join("?" for _item in batch_ids)
                database.execute(
                    f"""
                    DELETE FROM summary_sync_outbox
                    WHERE data_type = ?
                      AND entity_table = 'summary_import_history'
                      AND entity_id IN ({placeholders})
                    """,
                    (data_type.value, *batch_ids),
                )
            database.execute("DELETE FROM summary_query_cache")
            detail = json.dumps(
                {
                    "data_type": data_type.value,
                    "period": clean_period,
                    "row_count": detail_rows,
                    "batch_count": batch_rows,
                    "latest_import_at": info["latest_import_at"],
                    "source_files": info["source_files"],
                },
                ensure_ascii=False,
            )
            self.log_action(
                database,
                data_type=data_type,
                action="delete_nim_period",
                target_id=clean_period,
                detail=detail,
                created_by=created_by,
            )
        self._memory_cache.clear()
        return info

    def delete_batch(self, batch_id: int, *, created_by: str = "") -> dict[str, object]:
        clean_batch_id = int(batch_id)
        with self._database() as database:
            batch = database.execute(
                """
                SELECT id, data_type, period, file_name, row_count
                FROM summary_import_history
                WHERE id = ?
                """,
                (clean_batch_id,),
            ).fetchone()
            if batch is None:
                raise SummaryError(f"Không tìm thấy batch {clean_batch_id}.")
            data_type_text = str(batch["data_type"] or "")
            data_type = SummaryDataType(data_type_text)
            detail_tables = {
                SummaryDataType.NIM_DN: NIM_PERIOD_SUMMARY_TABLE,
                SummaryDataType.NIM_NV: NIM_PERIOD_SUMMARY_TABLE,
                SummaryDataType.LOAN_COMPARE: "loan_compare_details",
                SummaryDataType.CREDIT_LIMIT: "credit_limit_details",
            }
            detail_table = detail_tables[data_type]
            detail_count = int(
                database.execute(
                    f"SELECT COUNT(*) FROM {detail_table} WHERE batch_id = ?",
                    (clean_batch_id,),
                ).fetchone()[0]
                or 0
            )
            database.execute(
                """
                DELETE FROM summary_import_history
                WHERE id = ?
                """,
                (clean_batch_id,),
            )
            database.execute(
                """
                DELETE FROM summary_sync_outbox
                WHERE entity_table = 'summary_import_history' AND entity_id = ?
                """,
                (str(clean_batch_id),),
            )
            database.execute("DELETE FROM summary_query_cache")
            result = {
                "batch_id": clean_batch_id,
                "data_type": data_type.value,
                "period": str(batch["period"] or ""),
                "file_name": str(batch["file_name"] or ""),
                "row_count": detail_count,
            }
            self.log_action(
                database,
                data_type=data_type,
                action="delete_batch",
                target_id=str(clean_batch_id),
                detail=json.dumps(result, ensure_ascii=False),
                created_by=created_by,
            )
        self._memory_cache.clear()
        return result

    def query_nim(
        self,
        data_type: SummaryDataType,
        *,
        search: str = "",
        filters: dict[str, object] | None = None,
        page: int = 1,
        page_size: int = 200,
    ) -> PageResult:
        filters = filters or {}
        where, params = self._nim_where(data_type, search, filters)
        has_customer_type_filter = bool(str(filters.get("customer_type") or "").strip())
        customer_type_select = "customer_type" if has_customer_type_filter else "'Tất cả' AS customer_type"
        group_columns = "period, branch_code, trctcd, officer_code, officer_name"
        if has_customer_type_filter:
            group_columns += ", customer_type"
        group_sql = f"""
            SELECT
                period,
                branch_code,
                trctcd,
                MIN(branch_name) AS branch,
                MIN(transaction_office) AS transaction_office,
                {customer_type_select},
                {NIM_OFFICER_DISPLAY_SQL} AS officer,
                SUM(balance) AS balance,
                CASE WHEN SUM(balance) <> 0 THEN SUM(interest_rate_numerator) / SUM(balance) ELSE 0 END AS average_rate,
                CASE WHEN SUM(balance) <> 0 THEN SUM(numerator_before) / SUM(balance) ELSE 0 END AS nim_before,
                CASE WHEN SUM(balance) <> 0 THEN SUM(numerator_after) / SUM(balance) ELSE 0 END AS nim_after
            FROM nim_period_summary
            {where}
            GROUP BY {group_columns}
        """
        result = self._cached_page(
            key_parts=("nim", data_type.value, search, filters, page, page_size),
            sql=group_sql,
            params=params,
            order_sql="ORDER BY period DESC, branch_code COLLATE NOCASE, trctcd COLLATE NOCASE, customer_type COLLATE NOCASE, officer COLLATE NOCASE",
            page=page,
            page_size=page_size,
        )
        return PageResult(
            rows=[self._dynamic_nim_unit_row(row) for row in result.rows],
            total_rows=result.total_rows,
            page=result.page,
            page_size=result.page_size,
        )

    def dashboard_nim(
        self,
        data_type: SummaryDataType,
        *,
        filters: dict[str, object] | None = None,
    ) -> DashboardData:
        filters = filters or {}
        where, params = self._nim_where(data_type, "", filters)
        cache_key = self._cache_key(("dashboard_nim", data_type.value, filters))
        cached = self._load_cache(cache_key)
        if isinstance(cached, dict):
            return DashboardData(
                metrics=tuple(DashboardMetric(**item) for item in cached.get("metrics", ())),
                bars=tuple(tuple(item) for item in cached.get("bars", ())),
                lines=tuple(tuple(item) for item in cached.get("lines", ())),
                pies=tuple(tuple(item) for item in cached.get("pies", ())),
            )
        with self._database() as database:
            row = database.execute(
                f"""
                SELECT
                    COUNT(*) AS rows_count,
                    SUM(balance) AS balance,
                    CASE WHEN SUM(balance) <> 0 THEN SUM(numerator_before) / SUM(balance) ELSE 0 END AS nim_before,
                    CASE WHEN SUM(balance) <> 0 THEN SUM(numerator_after) / SUM(balance) ELSE 0 END AS nim_after,
                    CASE WHEN SUM(balance) <> 0 THEN SUM(interest_rate_numerator) / SUM(balance) ELSE 0 END AS average_rate
                FROM nim_period_summary
                {where}
                """,
                params,
            ).fetchone()
            bars = database.execute(
                f"""
                SELECT {NIM_OFFICER_DISPLAY_SQL} AS officer, SUM(balance) AS value
                FROM nim_period_summary
                {where}
                GROUP BY officer_code, officer_name
                ORDER BY value DESC
                LIMIT 10
                """,
                params,
            ).fetchall()
            lines = database.execute(
                f"""
                SELECT period, CASE WHEN SUM(balance) <> 0 THEN SUM(numerator_after) / SUM(balance) ELSE 0 END AS value
                FROM nim_period_summary
                {where}
                GROUP BY period
                ORDER BY period
                """,
                params,
            ).fetchall()
            pies = database.execute(
                f"""
                SELECT customer_type, SUM(balance) AS value
                FROM nim_period_summary
                {where}
                GROUP BY customer_type
                ORDER BY value DESC
                """,
                params,
            ).fetchall()
        metrics = (
            DashboardMetric("Tổng dư nợ", f"{float(row['balance'] or 0):,.0f}"),
            DashboardMetric("NIM trước ĐC", f"{float(row['nim_before'] or 0):,.2f}%"),
            DashboardMetric("NIM sau ĐC", f"{float(row['nim_after'] or 0):,.2f}%"),
            DashboardMetric("Lãi suất bình quân", f"{float(row['average_rate'] or 0):,.2f}%"),
        )
        result = DashboardData(
            metrics=metrics,
            bars=tuple((str(item["officer"]), float(item["value"] or 0)) for item in bars),
            lines=tuple((str(item["period"]), float(item["value"] or 0)) for item in lines),
            pies=tuple((str(item["customer_type"]), float(item["value"] or 0)) for item in pies),
        )
        self._save_cache(cache_key, asdict(result), data_type.value)
        return result

    def get_officer_history(
        self,
        data_type: SummaryDataType,
        *,
        officer_code: str = "",
        officer: str = "",
        branch: str = "",
        transaction_office: str = "",
        customer_type: str = "",
        period_from: str = "",
        period_to: str = "",
    ) -> list[dict[str, object]]:
        officer = str(officer or "").strip()
        officer_code = str(officer_code or "").strip()
        if not officer and not officer_code:
            return []
        clauses = ["data_type = ?"]
        params: list[object] = [data_type.value]
        if officer_code:
            clauses.append("officer_code = ?")
            params.append(officer_code)
        else:
            parsed_code, parsed_name = _split_nim_officer(officer)
            if parsed_code:
                clauses.append("officer_code = ?")
                params.append(parsed_code)
            else:
                clauses.append("officer_name = ?")
                params.append(parsed_name)
        clean_branch = str(branch or "").strip()
        branch_code = _branch_code_from_filter(clean_branch)
        if clean_branch and clean_branch != "Tất cả":
            if branch_code:
                clauses.append("branch_code = ?")
                params.append(branch_code)
            else:
                clauses.append("branch_name = ?")
                params.append(clean_branch)
        clean_office = str(transaction_office or "").strip()
        if clean_office and clean_office != "Tất cả":
            office_branch, separator, office_trctcd = clean_office.partition("-")
            if separator and office_branch and office_trctcd:
                trctcd = office_trctcd.split(" ", 1)[0].strip()
                clauses.append("branch_code = ?")
                clauses.append("trctcd = ?")
                params.extend([office_branch.strip(), trctcd])
            else:
                matched_trctcd = self._trctcd_from_office_filter(branch_code, clean_office)
                if matched_trctcd:
                    clauses.append("trctcd = ?")
                    params.append(matched_trctcd)
                else:
                    clauses.append("transaction_office = ?")
                    params.append(clean_office)
        clean_customer_type = str(customer_type or "").strip()
        if clean_customer_type and clean_customer_type != "Tất cả":
            clauses.append("customer_type = ?")
            params.append(clean_customer_type)
        if period_from:
            clauses.append("period >= ?")
            params.append(str(period_from))
        if period_to:
            clauses.append("period <= ?")
            params.append(str(period_to))
        where = "WHERE " + " AND ".join(clauses)
        with self._database() as database:
            rows = database.execute(
                f"""
                SELECT
                    period,
                    SUM(balance) AS balance,
                    CASE WHEN SUM(balance) <> 0 THEN SUM(interest_rate_numerator) / SUM(balance) ELSE 0 END AS average_rate,
                    CASE WHEN SUM(balance) <> 0 THEN SUM(numerator_before) / SUM(balance) ELSE 0 END AS nim_before,
                    CASE WHEN SUM(balance) <> 0 THEN SUM(numerator_after) / SUM(balance) ELSE 0 END AS nim_after
                FROM nim_period_summary
                {where}
                GROUP BY period
                ORDER BY period
                """,
                params,
            ).fetchall()
        return [dict(row) for row in rows]

    def _trctcd_from_office_filter(self, branch_code: str, value: str) -> str:
        if not branch_code or not value:
            return ""
        for office in self.unit_directory.repository.list_offices(branch_code=branch_code, active_only=False):
            labels = {
                office.office_name,
                office.short_name,
                self.unit_directory.get_office_name(office.branch_code, office.trctcd),
                self.unit_directory.get_office_display_name(office.branch_code, office.trctcd),
            }
            if value in labels:
                return office.trctcd
        return ""

    def query_loan_compare(
        self,
        *,
        batch_id: int | None = None,
        search: str = "",
        category: str = "",
        officer: str = "",
        page: int = 1,
        page_size: int = 200,
    ) -> PageResult:
        clauses = ["1 = 1"]
        params: list[object] = []
        if batch_id:
            clauses.append("batch_id = ?")
            params.append(int(batch_id))
        if category:
            clauses.append("category = ?")
            params.append(category)
        if officer:
            clauses.append("officer = ?")
            params.append(officer)
        if search.strip():
            needle = f"%{search.strip()}%"
            clauses.append("(customer_code LIKE ? OR customer_name LIKE ? OR officer LIKE ?)")
            params.extend([needle, needle, needle])
        sql = f"""
            SELECT customer_code, customer_name, previous_balance, current_balance,
                   difference, category, officer, address
            FROM loan_compare_details
            WHERE {' AND '.join(clauses)}
        """
        return self._cached_page(
            key_parts=("loan", batch_id, search, category, officer, page, page_size),
            sql=sql,
            params=params,
            order_sql="ORDER BY ABS(difference) DESC, customer_code COLLATE NOCASE",
            page=page,
            page_size=page_size,
        )

    def dashboard_loan_compare(self, batch_id: int | None = None) -> DashboardData:
        clauses = ["1 = 1"]
        params: list[object] = []
        if batch_id:
            clauses.append("batch_id = ?")
            params.append(int(batch_id))
        where = "WHERE " + " AND ".join(clauses)
        with self._database() as database:
            by_category = database.execute(
                f"""
                SELECT category, COUNT(*) AS count_rows, SUM(difference) AS amount
                FROM loan_compare_details
                {where}
                GROUP BY category
                """,
                params,
            ).fetchall()
            top = database.execute(
                f"""
                SELECT officer, SUM(difference) AS amount
                FROM loan_compare_details
                {where}
                GROUP BY officer
                ORDER BY ABS(amount) DESC
                LIMIT 10
                """,
                params,
            ).fetchall()
        category_map = {
            str(row["category"]): {
                "count_rows": int(row["count_rows"] or 0),
                "amount": float(row["amount"] or 0),
            }
            for row in by_category
        }
        metrics = (
            DashboardMetric("KH tăng", str(category_map.get("Khach hang vay tang", {}).get("count_rows", 0))),
            DashboardMetric("KH giảm", str(category_map.get("Khach hang vay giam", {}).get("count_rows", 0))),
            DashboardMetric("Tổng tăng", f"{sum(float(row['amount'] or 0) for row in by_category if float(row['amount'] or 0) > 0):,.0f}"),
            DashboardMetric("Tổng giảm", f"{sum(float(row['amount'] or 0) for row in by_category if float(row['amount'] or 0) < 0):,.0f}"),
        )
        return DashboardData(
            metrics=metrics,
            bars=tuple((str(row["officer"] or "Không rõ"), float(row["amount"] or 0)) for row in top),
            pies=tuple((str(row["category"]), float(row["count_rows"] or 0)) for row in by_category),
        )

    def query_credit_limits(
        self,
        *,
        batch_id: str | int | None = None,
        search: str = "",
        status: str = "",
        officer: str = "",
        page: int = 1,
        page_size: int = 200,
        min_limit: float | None = None,
        warn_days: int | None = None,
        reference_date: date | None = None,
    ) -> PageResult:
        if self.list_credit_limit_batches() or (batch_id and not _looks_like_int(batch_id)):
            return self.credit_limit_store.query_credit_limits(
                batch_id=batch_id,
                search=search,
                status=status,
                officer=officer,
                page=page,
                page_size=page_size,
                min_limit=min_limit,
                warn_days=warn_days,
                reference_date=reference_date,
            )
        clauses = ["1 = 1"]
        params: list[object] = []
        if batch_id:
            clauses.append("batch_id = ?")
            params.append(int(batch_id))
        if status:
            clauses.append("status = ?")
            params.append(status)
        if officer:
            clauses.append("officer = ?")
            params.append(officer)
        if search.strip():
            needle = f"%{search.strip()}%"
            clauses.append("(customer_code LIKE ? OR customer_name LIKE ? OR contract_number LIKE ? OR officer LIKE ?)")
            params.extend([needle, needle, needle, needle])
        sql = f"""
            SELECT customer_code, customer_name, contract_number, approved_date,
                   approved_amount, outstanding_balance, expiry_date, address,
                   officer, days_to_expiry, status, note
            FROM credit_limit_details
            WHERE {' AND '.join(clauses)}
        """
        return self._cached_page(
            key_parts=("credit_limit", batch_id, search, status, officer, page, page_size),
            sql=sql,
            params=params,
            order_sql="ORDER BY expiry_date, approved_amount DESC, customer_code COLLATE NOCASE",
            page=page,
            page_size=page_size,
        )

    def dashboard_credit_limits(
        self,
        batch_id: str | int | None = None,
        *,
        search: str = "",
        status: str = "",
        officer: str = "",
        min_limit: float | None = None,
        warn_days: int | None = None,
        reference_date: date | None = None,
    ) -> DashboardData:
        if self.list_credit_limit_batches() or (batch_id and not _looks_like_int(batch_id)):
            return self.credit_limit_store.dashboard_credit_limits(
                batch_id=batch_id,
                search=search,
                status=status,
                officer=officer,
                min_limit=min_limit,
                warn_days=warn_days,
                reference_date=reference_date,
            )
        clauses = ["1 = 1"]
        params: list[object] = []
        if batch_id:
            clauses.append("batch_id = ?")
            params.append(int(batch_id))
        if status:
            clauses.append("status = ?")
            params.append(status)
        if officer:
            clauses.append("officer = ?")
            params.append(officer)
        if search.strip():
            needle = f"%{search.strip()}%"
            clauses.append("(customer_code LIKE ? OR customer_name LIKE ? OR contract_number LIKE ? OR officer LIKE ?)")
            params.extend([needle, needle, needle, needle])
        where = "WHERE " + " AND ".join(clauses)
        with self._database() as database:
            status_rows = database.execute(
                f"""
                SELECT status, COUNT(*) AS count_rows, SUM(approved_amount) AS amount
                FROM credit_limit_details
                {where}
                GROUP BY status
                """,
                params,
            ).fetchall()
            by_officer = database.execute(
                f"""
                SELECT officer, COUNT(*) AS count_rows
                FROM credit_limit_details
                {where}
                GROUP BY officer
                ORDER BY count_rows DESC
                LIMIT 10
                """,
                params,
            ).fetchall()
            by_month = database.execute(
                f"""
                SELECT substr(expiry_date, 1, 7) AS month_key, COUNT(*) AS count_rows
                FROM credit_limit_details
                {where}
                GROUP BY month_key
                ORDER BY month_key
                """,
                params,
            ).fetchall()
        status_map = {str(row["status"]): int(row["count_rows"] or 0) for row in status_rows}
        metrics = (
            DashboardMetric("Đã hết hạn", str(status_map.get("Đã hết hạn", 0))),
            DashboardMetric("Sắp hết hạn", str(status_map.get("Sắp hết hạn", 0))),
            DashboardMetric("Tổng cảnh báo", str(sum(status_map.values()))),
        )
        return DashboardData(
            metrics=metrics,
            bars=tuple((str(row["officer"] or "Không rõ"), float(row["count_rows"] or 0)) for row in by_officer),
            lines=tuple((str(row["month_key"] or "Không rõ"), float(row["count_rows"] or 0)) for row in by_month),
            pies=tuple((str(row["status"]), float(row["count_rows"] or 0)) for row in status_rows),
        )

    def distinct_credit_limit_values(
        self,
        column_name: str,
        *,
        batch_id: str | int | None = None,
        min_limit: float | None = None,
        warn_days: int | None = None,
        reference_date: date | None = None,
    ) -> list[str]:
        if self.list_credit_limit_batches() or (batch_id and not _looks_like_int(batch_id)):
            return self.credit_limit_store.distinct_values(
                column_name,
                batch_id=batch_id,
                min_limit=min_limit,
                warn_days=warn_days,
                reference_date=reference_date,
            )
        return self.distinct_values("credit_limit_details", column_name)

    def distinct_values(self, table_name: str, column_name: str, *, data_type: SummaryDataType | None = None) -> list[str]:
        allowed = {
            (NIM_PERIOD_SUMMARY_TABLE, "period"),
            (NIM_PERIOD_SUMMARY_TABLE, "branch_name"),
            (NIM_PERIOD_SUMMARY_TABLE, "transaction_office"),
            (NIM_PERIOD_SUMMARY_TABLE, "customer_type"),
            (NIM_PERIOD_SUMMARY_TABLE, "officer"),
            ("loan_compare_details", "category"),
            ("loan_compare_details", "officer"),
            ("credit_limit_details", "status"),
            ("credit_limit_details", "officer"),
        }
        if (table_name, column_name) not in allowed:
            raise SummaryError("Trường lọc không hợp lệ.")
        params: list[object] = []
        select_column = f"{column_name} AS value"
        where = f"{column_name} <> ''"
        if table_name == NIM_PERIOD_SUMMARY_TABLE and column_name == "officer":
            select_column = f"{NIM_OFFICER_DISPLAY_SQL} AS value"
            where = "officer_name <> ''"
        if data_type is not None and table_name == NIM_PERIOD_SUMMARY_TABLE:
            where += " AND data_type = ?"
            params.append(data_type.value)
        with self._database() as database:
            rows = database.execute(
                f"""
                SELECT DISTINCT {select_column}
                FROM {table_name}
                WHERE {where}
                ORDER BY value COLLATE NOCASE
                """,
                params,
            ).fetchall()
        return [str(row["value"]) for row in rows]

    def clear_cache(self) -> None:
        self._memory_cache.clear()
        try:
            with self._database() as database:
                database.execute("DELETE FROM summary_query_cache")
        except sqlite3.Error:
            pass

    def log_action(
        self,
        database: sqlite3.Connection | None = None,
        *,
        data_type: SummaryDataType | None,
        action: str,
        target_id: str = "",
        detail: str = "",
        created_by: str = "",
    ) -> None:
        now = now_text()
        params = (
            data_type.value if data_type else "",
            action,
            target_id,
            detail,
            created_by,
            now,
        )
        sql = """
            INSERT INTO summary_action_log(
                data_type, action, target_id, detail, created_by, created_at
            )
            VALUES (?, ?, ?, ?, ?, ?)
        """
        if database is not None:
            database.execute(sql, params)
            return
        with self._database() as own_database:
            own_database.execute(sql, params)

    def backup_database(self, destination: Path | None = None) -> Path:
        if destination is None:
            stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
            destination = self.database_path.parent / "backups" / f"summary-{stamp}.zip"
        destination = Path(destination)
        destination.parent.mkdir(parents=True, exist_ok=True)
        with tempfile.TemporaryDirectory(prefix="summary-backup-") as tmp_dir:
            snapshot = Path(tmp_dir) / self.database_path.name
            with closing(sqlite3.connect(self.database_path, timeout=30)) as source:
                with closing(sqlite3.connect(snapshot)) as target:
                    source.backup(target)
            manifest = {
                "format": "agribank-v3-summary-backup",
                "created_at": now_text(),
                "database": self.database_path.name,
            }
            with zipfile.ZipFile(destination, "w", compression=zipfile.ZIP_DEFLATED) as archive:
                archive.writestr("manifest.json", json.dumps(manifest, ensure_ascii=False, indent=2))
                archive.write(snapshot, self.database_path.name)
        self.log_action(data_type=None, action="backup", target_id=str(destination), detail="Sao lưu dữ liệu tổng hợp")
        return destination

    def restore_database(self, source_path: Path) -> Path:
        source_path = Path(source_path)
        if not zipfile.is_zipfile(source_path):
            raise SummaryError("File phục hồi không phải gói sao lưu hợp lệ.")
        safety_backup = self.backup_database()
        with tempfile.TemporaryDirectory(prefix="summary-restore-") as tmp_dir:
            tmp_root = Path(tmp_dir)
            with zipfile.ZipFile(source_path, "r") as archive:
                manifest = json.loads(archive.read("manifest.json"))
                if manifest.get("format") != "agribank-v3-summary-backup":
                    raise SummaryError("Gói sao lưu không đúng định dạng Tổng hợp số liệu.")
                database_name = str(manifest.get("database") or self.database_path.name)
                if Path(database_name).name != self.database_path.name:
                    raise SummaryError("Gói sao lưu không trùng database hiện tại.")
                candidate = tmp_root / database_name
                candidate.write_bytes(archive.read(database_name))
            with closing(sqlite3.connect(candidate)) as database:
                integrity = str(database.execute("PRAGMA integrity_check").fetchone()[0])
                if integrity.casefold() != "ok":
                    raise SummaryError(f"Snapshot phục hồi bị lỗi: {integrity}")
            with closing(sqlite3.connect(self.database_path, timeout=30)) as current:
                current.execute("PRAGMA wal_checkpoint(TRUNCATE)")
                current.execute("PRAGMA journal_mode = DELETE")
            for suffix in ("-wal", "-shm"):
                Path(f"{self.database_path}{suffix}").unlink(missing_ok=True)
            replacement = self.database_path.with_name(f".restore-{uuid4().hex}.db")
            shutil.copy2(candidate, replacement)
            os.replace(replacement, self.database_path)
        self.clear_cache()
        self.log_action(data_type=None, action="restore", target_id=str(source_path), detail="Khôi phục dữ liệu tổng hợp")
        return safety_backup

    def maintenance_status(self) -> dict[str, object]:
        with self._database() as database:
            nim_dn_periods = int(
                database.execute(
                    """
                    SELECT COUNT(DISTINCT period)
                    FROM nim_period_summary
                    WHERE data_type = ?
                    """,
                    (SummaryDataType.NIM_DN.value,),
                ).fetchone()[0]
                or 0
            )
            nim_nv_periods = int(
                database.execute(
                    """
                    SELECT COUNT(DISTINCT period)
                    FROM nim_period_summary
                    WHERE data_type = ?
                    """,
                    (SummaryDataType.NIM_NV.value,),
                ).fetchone()[0]
                or 0
            )
            raw_nim_rows = int(
                database.execute("SELECT COUNT(*) FROM nim_details").fetchone()[0]
                or 0
            )
            loan_batches = int(
                database.execute(
                    """
                    SELECT COUNT(*)
                    FROM summary_import_history
                    WHERE data_type = ?
                    """,
                    (SummaryDataType.LOAN_COMPARE.value,),
                ).fetchone()[0]
                or 0
            )
            limit_batches = int(
                database.execute(
                    """
                    SELECT COUNT(*)
                    FROM summary_import_history
                    WHERE data_type = ?
                    """,
                    (SummaryDataType.CREDIT_LIMIT.value,),
                ).fetchone()[0]
                or 0
            )
        credit_limit_files = self.credit_limit_store.maintenance_status()
        return {
            "database_path": self.database_path,
            "size_bytes": self.database_path.stat().st_size if self.database_path.is_file() else 0,
            "nim_dn_periods": nim_dn_periods,
            "nim_nv_periods": nim_nv_periods,
            "raw_nim_rows": raw_nim_rows,
            "loan_compare_batches": loan_batches,
            "credit_limit_batches": limit_batches,
            "credit_limit_file_path": credit_limit_files.storage_path,
            "credit_limit_file_batches": credit_limit_files.valid_files,
            "credit_limit_invalid_files": credit_limit_files.invalid_files,
            "credit_limit_file_size_bytes": credit_limit_files.total_size_bytes,
        }

    def optimize_database(self, *, vacuum: bool = False) -> dict[str, object]:
        before = self.database_path.stat().st_size if self.database_path.is_file() else 0
        with closing(self.connect()) as database:
            database.execute("PRAGMA optimize")
            database.commit()
            if vacuum:
                database.execute("PRAGMA wal_checkpoint(TRUNCATE)")
                database.execute("VACUUM")
                database.commit()
        self.clear_cache()
        after = self.database_path.stat().st_size if self.database_path.is_file() else 0
        self.log_action(
            data_type=None,
            action="optimize",
            target_id=self.database_path.name,
            detail=json.dumps({"vacuum": bool(vacuum), "before": before, "after": after}, ensure_ascii=False),
        )
        return {"before_size_bytes": before, "after_size_bytes": after, "vacuum": bool(vacuum)}

    def _nim_period_info(
        self,
        database: sqlite3.Connection,
        data_type: SummaryDataType,
        period: str,
    ) -> dict[str, object]:
        row_count = int(
            database.execute(
                """
                SELECT COUNT(*)
                FROM nim_period_summary
                WHERE data_type = ? AND period = ?
                """,
                (data_type.value, period),
            ).fetchone()[0]
            or 0
        )
        batch_summary = database.execute(
            """
            SELECT
                COUNT(*) AS batch_count,
                MAX(created_at) AS latest_import_at
            FROM summary_import_history
            WHERE data_type = ? AND period = ?
            """,
            (data_type.value, period),
        ).fetchone()
        file_rows = database.execute(
            """
            SELECT file_name
            FROM summary_import_history
            WHERE data_type = ? AND period = ? AND file_name <> ''
            ORDER BY created_at DESC, id DESC
            LIMIT 8
            """,
            (data_type.value, period),
        ).fetchall()
        return {
            "data_type": data_type.value,
            "period": period,
            "row_count": row_count,
            "batch_count": int(batch_summary["batch_count"] or 0) if batch_summary else 0,
            "latest_import_at": str(batch_summary["latest_import_at"] or "") if batch_summary else "",
            "source_files": [str(row["file_name"] or "") for row in file_rows],
        }

    def _cached_page(
        self,
        *,
        key_parts: tuple[object, ...],
        sql: str,
        params: Sequence[object],
        order_sql: str,
        page: int,
        page_size: int,
    ) -> PageResult:
        page = max(1, int(page))
        page_size = max(10, min(2000, int(page_size)))
        cache_key = self._cache_key(key_parts)
        cached = self._load_cache(cache_key)
        if isinstance(cached, dict):
            return PageResult(
                rows=list(cached["rows"]),
                total_rows=int(cached["total_rows"]),
                page=page,
                page_size=page_size,
            )
        offset = (page - 1) * page_size
        with self._database() as database:
            total = int(database.execute(f"SELECT COUNT(*) FROM ({sql}) AS q", params).fetchone()[0])
            rows = database.execute(
                f"{sql} {order_sql} LIMIT ? OFFSET ?",
                (*params, page_size, offset),
            ).fetchall()
        payload = {
            "rows": [dict(row) for row in rows],
            "total_rows": total,
        }
        self._save_cache(cache_key, payload, str(key_parts[0]))
        return PageResult(
            rows=payload["rows"],
            total_rows=total,
            page=page,
            page_size=page_size,
        )

    def _nim_where(
        self,
        data_type: SummaryDataType,
        search: str,
        filters: dict[str, object],
    ) -> tuple[str, list[object]]:
        clauses = ["data_type = ?"]
        params: list[object] = [data_type.value]
        for key, column in (
            ("period", "period"),
            ("customer_type", "customer_type"),
            ("batch_id", "batch_id"),
        ):
            value = filters.get(key)
            if value in (None, "", 0):
                continue
            clauses.append(f"{column} = ?")
            params.append(value)
        branch_value = str(filters.get("branch") or "").strip()
        if branch_value:
            branch_code = _branch_code_from_filter(branch_value)
            if branch_code:
                clauses.append("branch_code = ?")
                params.append(branch_code)
            else:
                clauses.append("branch_name = ?")
                params.append(branch_value)
        office_value = str(filters.get("transaction_office") or "").strip()
        if office_value:
            branch_code, sep, trctcd = office_value.partition("-")
            if sep and branch_code and trctcd:
                clauses.append("branch_code = ?")
                clauses.append("trctcd = ?")
                params.extend([branch_code, trctcd])
            else:
                clauses.append("transaction_office = ?")
                params.append(office_value)
        officer_filter = filters.get("officer")
        if officer_filter not in (None, "", 0):
            self._add_summary_officer_clause(clauses, params, officer=officer_filter)
        if search.strip():
            needle = f"%{search.strip()}%"
            clauses.append(
                "(branch_name LIKE ? OR transaction_office LIKE ? OR customer_type LIKE ? "
                "OR officer_code LIKE ? OR officer_name LIKE ? OR period LIKE ?)"
            )
            params.extend([needle, needle, needle, needle, needle, needle])
        return "WHERE " + " AND ".join(clauses), params

    def _dynamic_nim_unit_row(self, row: dict[str, object]) -> dict[str, object]:
        branch_code = str(row.get("branch_code") or "").strip()
        trctcd = str(row.get("trctcd") or "").strip()
        if branch_code:
            row["branch"] = self.unit_directory.get_branch_display_name(branch_code)
        if branch_code and trctcd:
            row["transaction_office"] = self.unit_directory.get_office_name(branch_code, trctcd)
        row.pop("branch_code", None)
        row.pop("trctcd", None)
        return row

    def verify_nim_summary_totals(
        self,
        *,
        database: sqlite3.Connection | None = None,
    ) -> dict[str, object]:
        def _verify(connection: sqlite3.Connection) -> dict[str, object]:
            raw_rows = connection.execute(
                """
                SELECT
                    batch_id,
                    data_type,
                    period,
                    SUM(balance) AS balance,
                    SUM(average_rate_numerator) AS interest_rate_numerator,
                    SUM(numerator_before) AS numerator_before,
                    SUM(numerator_after) AS numerator_after,
                    COUNT(*) AS source_row_count
                FROM nim_details
                GROUP BY batch_id, data_type, period
                ORDER BY batch_id, data_type, period
                """
            ).fetchall()
            mismatches: list[dict[str, object]] = []
            summary_source_rows = 0
            for raw in raw_rows:
                summary = connection.execute(
                    """
                    SELECT
                        SUM(balance) AS balance,
                        SUM(interest_rate_numerator) AS interest_rate_numerator,
                        SUM(numerator_before) AS numerator_before,
                        SUM(numerator_after) AS numerator_after,
                        SUM(source_row_count) AS source_row_count,
                        COUNT(*) AS summary_row_count
                    FROM nim_period_summary
                    WHERE batch_id = ? AND data_type = ? AND period = ?
                    """,
                    (raw["batch_id"], raw["data_type"], raw["period"]),
                ).fetchone()
                summary_source_rows += int((summary["source_row_count"] if summary else 0) or 0)
                for field in ("balance", "interest_rate_numerator", "numerator_before", "numerator_after", "source_row_count"):
                    raw_value = float(raw[field] or 0)
                    summary_value = float(summary[field] or 0) if summary else 0.0
                    tolerance = 0.0001 if field != "source_row_count" else 0.0
                    if abs(raw_value - summary_value) > tolerance:
                        mismatches.append(
                            {
                                "batch_id": int(raw["batch_id"]),
                                "data_type": str(raw["data_type"]),
                                "period": str(raw["period"]),
                                "field": field,
                                "raw": raw_value,
                                "summary": summary_value,
                            }
                        )
            raw_total_rows = sum(int(row["source_row_count"] or 0) for row in raw_rows)
            summary_rows = int(connection.execute("SELECT COUNT(*) FROM nim_period_summary").fetchone()[0] or 0)
            return {
                "matches": not mismatches,
                "raw_rows": raw_total_rows,
                "summary_rows": summary_rows,
                "summary_source_rows_for_raw": summary_source_rows,
                "mismatches": mismatches[:20],
            }

        if database is not None:
            return _verify(database)
        with self._database() as own_database:
            return _verify(own_database)

    def compact_legacy_nim_details(
        self,
        *,
        vacuum: bool = False,
        created_by: str = "",
    ) -> dict[str, object]:
        before_size = self.database_path.stat().st_size if self.database_path.is_file() else 0
        with self._database() as database:
            verification = self.verify_nim_summary_totals(database=database)
            if not verification["matches"]:
                raise SummaryError(f"Chưa thể xóa raw NIM vì tổng chưa khớp: {verification}")
            raw_rows = int(database.execute("SELECT COUNT(*) FROM nim_details").fetchone()[0] or 0)
            if raw_rows <= 0:
                return {
                    "deleted_rows": 0,
                    "before_size_bytes": before_size,
                    "after_size_bytes": before_size,
                    "vacuum": False,
                    "verification": verification,
                }
            backup_path = self._backup_before_nim_summary_migration()
            database.execute("DELETE FROM nim_details")
            database.execute("DELETE FROM summary_query_cache")
            self.log_action(
                database,
                data_type=None,
                action="compact_legacy_nim_details",
                target_id=NIM_PERIOD_SUMMARY_TABLE,
                detail=json.dumps(
                    {
                        "deleted_rows": raw_rows,
                        "backup": str(backup_path),
                        "verification": verification,
                    },
                    ensure_ascii=False,
                ),
                created_by=created_by,
            )
        optimization = self.optimize_database(vacuum=vacuum) if vacuum else {}
        after_size = self.database_path.stat().st_size if self.database_path.is_file() else before_size
        self.clear_cache()
        return {
            "deleted_rows": raw_rows,
            "before_size_bytes": before_size,
            "after_size_bytes": after_size,
            "vacuum": bool(vacuum),
            "verification": verification,
            "optimization": optimization,
        }

    @staticmethod
    def _aggregate_nim_rows_for_storage(rows: Sequence[NimRow]) -> list[NimRow]:
        grouped: dict[tuple[object, ...], dict[str, object]] = {}
        for row in rows:
            officer_code, officer_name = _split_nim_officer(row.officer)
            key = (
                row.batch_id,
                row.data_type,
                row.period,
                row.branch_code,
                row.branch_name,
                row.trctcd,
                row.transaction_office,
                row.customer_type,
                officer_code,
                officer_name,
            )
            item = grouped.get(key)
            if item is None:
                item = {
                    "row": row,
                    "balance": 0.0,
                    "interest_rate_numerator": 0.0,
                    "numerator_before": 0.0,
                    "numerator_after": 0.0,
                    "source_row_count": 0,
                }
                grouped[key] = item
            item["balance"] = float(item["balance"]) + float(row.balance or 0)
            item["interest_rate_numerator"] = float(item["interest_rate_numerator"]) + float(row.average_rate_numerator or 0)
            item["numerator_before"] = float(item["numerator_before"]) + float(row.numerator_before or 0)
            item["numerator_after"] = float(item["numerator_after"]) + float(row.numerator_after or 0)
            item["source_row_count"] = int(item["source_row_count"]) + int(row.source_row_count or 1)
        return [
            NimRow(
                batch_id=int(item["row"].batch_id),  # type: ignore[union-attr]
                data_type=item["row"].data_type,  # type: ignore[union-attr]
                period=str(item["row"].period),  # type: ignore[union-attr]
                branch_code=str(item["row"].branch_code),  # type: ignore[union-attr]
                branch_name=str(item["row"].branch_name),  # type: ignore[union-attr]
                trctcd=str(item["row"].trctcd),  # type: ignore[union-attr]
                transaction_office=str(item["row"].transaction_office),  # type: ignore[union-attr]
                customer_type=str(item["row"].customer_type),  # type: ignore[union-attr]
                officer=str(item["row"].officer),  # type: ignore[union-attr]
                balance=float(item["balance"] or 0),
                interest_rate=0.0,
                ftp_rate=0.0,
                adjustment_rate=0.0,
                numerator_before=float(item["numerator_before"] or 0),
                numerator_after=float(item["numerator_after"] or 0),
                average_rate_numerator=float(item["interest_rate_numerator"] or 0),
                source_file=str(item["row"].source_file),  # type: ignore[union-attr]
                source_row_count=int(item["source_row_count"] or 0),
            )
            for item in grouped.values()
            if isinstance(item.get("row"), NimRow)
        ]

    @staticmethod
    def _add_summary_officer_clause(
        clauses: list[str],
        params: list[object],
        *,
        officer: object = "",
        officer_code: str = "",
    ) -> None:
        clean_code = str(officer_code or "").strip()
        clean_officer = str(officer or "").strip()
        if clean_code:
            clauses.append("officer_code = ?")
            params.append(clean_code)
            return
        parsed_code, parsed_name = _split_nim_officer(clean_officer)
        if parsed_code:
            clauses.append("officer_code = ?")
            params.append(parsed_code)
        elif parsed_name:
            clauses.append("officer_name = ?")
            params.append(parsed_name)

    def _backup_before_nim_summary_migration(self) -> Path | None:
        if not self.database_path.is_file():
            return None
        destination = self.database_path.parent / "backups" / (
            f"nim-summary-migration-{datetime.now().strftime('%Y%m%d-%H%M%S')}-{uuid4().hex[:8]}.db"
        )
        destination.parent.mkdir(parents=True, exist_ok=True)
        with closing(sqlite3.connect(self.database_path, timeout=30)) as source:
            with closing(sqlite3.connect(destination)) as target:
                source.backup(target)
        return destination

    @staticmethod
    def _table_row_count(database: sqlite3.Connection, table_name: str) -> int:
        row = database.execute(
            """
            SELECT 1
            FROM sqlite_master
            WHERE type = 'table' AND name = ?
            LIMIT 1
            """,
            (table_name,),
        ).fetchone()
        if row is None:
            return 0
        return int(database.execute(f"SELECT COUNT(*) FROM {table_name}").fetchone()[0] or 0)

    def _load_cache(self, cache_key: str) -> object | None:
        item = self._memory_cache.get(cache_key)
        if item is not None:
            return item
        now = now_text()
        try:
            with self._database() as database:
                row = database.execute(
                    """
                    SELECT payload FROM summary_query_cache
                    WHERE cache_key = ? AND expires_at > ?
                    """,
                    (cache_key, now),
                ).fetchone()
        except sqlite3.Error:
            return None
        if row is None:
            return None
        try:
            payload = json.loads(str(row["payload"]))
        except json.JSONDecodeError:
            return None
        self._memory_cache[cache_key] = payload
        return payload

    def _save_cache(self, cache_key: str, payload: object, data_type: str) -> None:
        self._memory_cache[cache_key] = payload
        now_dt = datetime.now().astimezone()
        now = now_dt.isoformat(timespec="seconds")
        expires = (now_dt + timedelta(seconds=self.CACHE_SECONDS)).isoformat(timespec="seconds")
        try:
            with self._database() as database:
                database.execute(
                    """
                    INSERT INTO summary_query_cache(cache_key, data_type, payload, created_at, expires_at)
                    VALUES (?, ?, ?, ?, ?)
                    ON CONFLICT(cache_key) DO UPDATE SET
                        data_type = excluded.data_type,
                        payload = excluded.payload,
                        created_at = excluded.created_at,
                        expires_at = excluded.expires_at
                    """,
                    (cache_key, data_type, json.dumps(payload, ensure_ascii=False), now, expires),
                )
        except sqlite3.Error:
            pass

    @staticmethod
    def _cache_key(parts: tuple[object, ...]) -> str:
        return json.dumps(parts, ensure_ascii=False, sort_keys=True, default=str)

    @staticmethod
    def _batch_from_row(row: sqlite3.Row) -> ImportBatch:
        return ImportBatch(
            id=int(row["id"]),
            data_type=str(row["data_type"]),
            period=str(row["period"]),
            source_path=str(row["source_path"]),
            file_name=str(row["file_name"]),
            imported_by=str(row["imported_by"]),
            row_count=int(row["row_count"]),
            status=str(row["status"]),
            message=str(row["message"]),
            created_at=str(row["created_at"]),
            updated_at=str(row["updated_at"]),
            version=int(row["version"]),
        )

    @staticmethod
    def _enqueue_sync(
        database: sqlite3.Connection,
        *,
        data_type: SummaryDataType,
        entity_table: str,
        entity_id: str,
        operation: str,
        payload: dict[str, object],
    ) -> None:
        database.execute(
            """
            INSERT INTO summary_sync_outbox(
                data_type, entity_table, entity_id, operation, payload, created_at, synced_at
            )
            VALUES (?, ?, ?, ?, ?, ?, '')
            """,
            (
                data_type.value,
                entity_table,
                entity_id,
                operation,
                json.dumps(payload, ensure_ascii=False, separators=(",", ":")),
                now_text(),
            ),
        )


def _split_nim_officer(raw_name: object) -> tuple[str, str]:
    text = str(raw_name or "").strip()
    if text.startswith("[") and "]" in text:
        code, name = text[1:].split("]", 1)
        return code.strip(), name.strip()
    return "", text


def _branch_code_from_filter(value: object) -> str:
    text = str(value or "").strip()
    if not text:
        return ""
    first = text.split(" - ", 1)[0].strip()
    if first and all(not char.isspace() for char in first):
        return first
    return ""


def _date_text(value: object) -> str | None:
    if value is None:
        return None
    if hasattr(value, "isoformat"):
        return value.isoformat()
    text = str(value).strip()
    return text or None


def _coerce_date(value: object) -> date | None:
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    text = str(value or "").strip()
    if not text:
        return None
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%Y/%m/%d"):
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            continue
    return None


def _looks_like_int(value: object) -> bool:
    text = str(value or "").strip()
    return text.isdigit()


def _file_hash(path: Path) -> str:
    digest = hashlib.sha256()
    with Path(path).open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _legacy_credit_limit_hash(batch_id: int) -> str:
    return hashlib.sha256(f"legacy-credit-limit:{batch_id}".encode("utf-8")).hexdigest()
