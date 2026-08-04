from __future__ import annotations

from contextlib import closing, contextmanager, nullcontext
from dataclasses import replace
from pathlib import Path
import json
import logging
import os
import shutil
import sqlite3
import tempfile
from time import perf_counter
from typing import Iterator
import unicodedata
from uuid import uuid4
import zipfile

from agribank_v3.features.credit.summary.customer.database import (
    CUSTOMER_DATABASE_NAME,
    CustomerDatabaseOperationLock,
    customer_database_path,
    ensure_customer_schema,
    get_customer_database_connection,
)
from agribank_v3.features.credit.summary.customer.models import (
    CustomerAggregationResult,
    CustomerDatabaseError,
    CustomerDatabaseStatus,
    CustomerDataType,
    CustomerOfficeType,
)
from agribank_v3.features.credit.summary.customer.services import (
    build_office_code,
    normalize_trctcd,
    resolve_representative_office,
)
from agribank_v3.features.credit.summary.customer.filters import (
    DEBT_GROUP_ALL,
    DEBT_GROUP_ATTENTION,
    DEBT_GROUP_BAD_DEBT,
    DEBT_GROUP_HAS_GROUP_1,
    DEBT_GROUP_HAS_GROUP_2,
    DEBT_GROUP_HAS_GROUP_3,
    DEBT_GROUP_HAS_GROUP_4,
    DEBT_GROUP_HAS_GROUP_5,
    DEBT_GROUP_UNKNOWN,
    DEBT_GROUP_WORST_1,
    DEBT_GROUP_WORST_2,
    DEBT_GROUP_WORST_3,
    DEBT_GROUP_WORST_4,
    DEBT_GROUP_WORST_5,
    MOVEMENT_STATUS_DECREASE,
    MOVEMENT_STATUS_INCREASE,
    MOVEMENT_STATUS_NEW,
    MOVEMENT_STATUS_PAID_OFF,
    MOVEMENT_STATUS_UNCHANGED,
    CustomerFilters,
    clean_filter_text,
)
from agribank_v3.features.credit.summary.customer.formatters import (
    format_money_vn,
    normalize_officer_code,
    normalize_officer_name,
)
from agribank_v3.features.credit.summary.models import PageResult, now_text
from agribank_v3.features.settings.unit_directory.service import (
    UnitDirectoryService,
    get_unit_directory_service,
)


LOGGER = logging.getLogger(__name__)
CUSTOMER_TYPE_LABELS = {
    "CN": "Cá nhân",
    "TC": "Tổ chức/Pháp nhân",
    "OTHER": "Khác",
}
CUSTOMER_TYPE_ORDER = {"CN": 0, "TC": 1, "OTHER": 2}
CUSTOMER_METRIC_SQL = {
    "average_rate": ("interest_rate_numerator", "Lãi suất bình quân"),
    "nim_before": ("nim_before_numerator", "NIM trước điều chỉnh"),
    "nim_after": ("nim_after_numerator", "NIM sau điều chỉnh"),
}
OFFICE_TYPE_LABELS = {
    CustomerOfficeType.HEAD_OFFICE.value: "Hội sở",
    CustomerOfficeType.TRANSACTION_OFFICE.value: "Phòng giao dịch",
    CustomerOfficeType.UNKNOWN.value: "Không xác định",
}
REPRESENTATIVE_REASON_LABELS = {
    "HAS_HEAD_OFFICE": "Có dư nợ tại Hội sở",
    "SINGLE_PGD": "Chỉ có một PGD",
    "MULTIPLE_PGD_LARGEST_BALANCE": "Nhiều PGD, chọn PGD dư nợ lớn nhất",
    "UNKNOWN": "Không xác định",
}
CUSTOMER_PERIOD_TABLES = (
    "customer_period_summary",
    "customer_officer_period",
    "customer_office_period",
    "customer_import_files",
    "customer_import_runs",
)
CUSTOMER_RETAINED_TABLES = (
    "customer_officer_override",
    "customer_action_log",
    "customer_officer_directory",
    "customer_schema_migrations",
)
CUSTOMER_DIAGNOSTIC_TABLES = (
    "customer_master",
    *CUSTOMER_PERIOD_TABLES,
    *CUSTOMER_RETAINED_TABLES,
)
CUSTOMER_MOVEMENT_TEMP_TABLE = "temp_customer_movements"
DEBT_GROUP_AGGREGATE_COLUMNS = (
    "has_debt_group_data",
    "worst_debt_group",
    "debt_group_unknown_row_count",
    "debt_group_1_balance",
    "debt_group_2_balance",
    "debt_group_3_balance",
    "debt_group_4_balance",
    "debt_group_5_balance",
    "debt_group_unknown_balance",
    "debt_group_1_interest_numerator",
    "debt_group_2_interest_numerator",
    "debt_group_3_interest_numerator",
    "debt_group_4_interest_numerator",
    "debt_group_5_interest_numerator",
    "debt_group_unknown_interest_numerator",
    "debt_group_1_nim_before_numerator",
    "debt_group_2_nim_before_numerator",
    "debt_group_3_nim_before_numerator",
    "debt_group_4_nim_before_numerator",
    "debt_group_5_nim_before_numerator",
    "debt_group_unknown_nim_before_numerator",
    "debt_group_1_nim_after_numerator",
    "debt_group_2_nim_after_numerator",
    "debt_group_3_nim_after_numerator",
    "debt_group_4_nim_after_numerator",
    "debt_group_5_nim_after_numerator",
    "debt_group_unknown_nim_after_numerator",
)
DEBT_GROUP_SUFFIXES = ("1", "2", "3", "4", "5", "unknown")
DEBT_GROUP_LABELS = {
    "1": "Nợ nhóm 1",
    "2": "Nợ nhóm 2",
    "3": "Nợ nhóm 3",
    "4": "Nợ nhóm 4",
    "5": "Nợ nhóm 5",
    "unknown": "Chưa xác định",
}
DEBT_GROUP_CODE_LABELS = {
    "01": "Nợ nhóm 1",
    "02": "Nợ nhóm 2",
    "03": "Nợ nhóm 3",
    "04": "Nợ nhóm 4",
    "05": "Nợ nhóm 5",
    "UNKNOWN": "Chưa xác định",
    "": "",
}


class CustomerRepository:
    def __init__(self, database_path: Path) -> None:
        self.main_database_path = Path(database_path)
        self.database_path = customer_database_path(self.main_database_path)
        self.database_path.parent.mkdir(parents=True, exist_ok=True)
        self.unit_directory = get_unit_directory_service(self.main_database_path)
        self.ensure_schema()

    def connect(self) -> sqlite3.Connection:
        return get_customer_database_connection(self.database_path)

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
            if self.database_path.is_file() and self._needs_debt_group_migration_backup():
                self.backup_database()
            with closing(self.connect()) as database:
                ensure_customer_schema(database)
                database.commit()
        except Exception as exc:
            raise CustomerDatabaseError(f"Khong the khoi tao {CUSTOMER_DATABASE_NAME}: {exc}") from exc

    def _needs_debt_group_migration_backup(self) -> bool:
        try:
            with closing(sqlite3.connect(self.database_path, timeout=10)) as database:
                tables = {
                    str(row[0] or "")
                    for row in database.execute(
                        "SELECT name FROM sqlite_master WHERE type = 'table'"
                    ).fetchall()
                }
                if "customer_period_summary" not in tables:
                    return False
                required = {"has_debt_group_data", "worst_debt_group", "debt_group_1_balance"}
                period_columns = _table_columns(database, "customer_period_summary")
                officer_columns = _table_columns(database, "customer_officer_period")
                return not required.issubset(period_columns) or not {"branch_code", "transaction_office"}.issubset(officer_columns)
        except Exception:
            LOGGER.exception("Could not inspect Customer.db before schema migration")
            return False

    def maintenance_status(self) -> CustomerDatabaseStatus:
        with self._database() as database:
            period_range = database.execute(
                """
                SELECT MIN(period) AS first_period, MAX(period) AS last_period
                FROM customer_period_summary
                """
            ).fetchone()
            page_count = self._pragma_int(database, "page_count")
            page_size = self._pragma_int(database, "page_size")
            freelist_count = self._pragma_int(database, "freelist_count")
            return CustomerDatabaseStatus(
                database_path=str(self.database_path),
                size_bytes=self.database_path.stat().st_size if self.database_path.is_file() else 0,
                master_count=self._count(database, "customer_master"),
                period_count=self._count_distinct(database, "customer_period_summary", "period"),
                period_summary_count=self._count(database, "customer_period_summary"),
                officer_period_count=self._count(database, "customer_officer_period"),
                import_run_count=self._count(database, "customer_import_runs"),
                import_file_count=self._count(database, "customer_import_files"),
                override_count=self._count(database, "customer_officer_override"),
                action_log_count=self._count(database, "customer_action_log"),
                officer_directory_count=self._count(database, "customer_officer_directory"),
                first_period=str(period_range["first_period"] or "") if period_range else "",
                last_period=str(period_range["last_period"] or "") if period_range else "",
                page_count=page_count,
                page_size=page_size,
                freelist_count=freelist_count,
                reclaimable_bytes=page_size * freelist_count,
            )

    def database_diagnostics(self) -> dict[str, object]:
        with self._database() as database:
            table_counts = {
                table_name: self._count(database, table_name)
                for table_name in CUSTOMER_DIAGNOSTIC_TABLES
                if self._table_exists(database, table_name)
            }
            page_size = self._pragma_int(database, "page_size")
            page_count = self._pragma_int(database, "page_count")
            freelist_count = self._pragma_int(database, "freelist_count")
            retained_counts = {
                table_name: table_counts.get(table_name, 0)
                for table_name in CUSTOMER_RETAINED_TABLES
                if table_name in table_counts
            }
            period_counts = {
                table_name: table_counts.get(table_name, 0)
                for table_name in CUSTOMER_PERIOD_TABLES
                if table_name in table_counts
            }
            return {
                "database_path": str(self.database_path),
                "size_bytes": self.database_path.stat().st_size if self.database_path.is_file() else 0,
                "page_size": page_size,
                "page_count": page_count,
                "freelist_count": freelist_count,
                "database_size_bytes": page_size * page_count,
                "reclaimable_bytes": page_size * freelist_count,
                "auto_vacuum": self._pragma_int(database, "auto_vacuum"),
                "journal_mode": self._pragma_text(database, "journal_mode"),
                "table_counts": table_counts,
                "period_table_counts": period_counts,
                "retained_table_counts": retained_counts,
                "has_period_data": bool(table_counts.get("customer_period_summary", 0)),
                "vacuum_recommended": bool(
                    not table_counts.get("customer_period_summary", 0)
                    and self.database_path.is_file()
                    and self.database_path.stat().st_size > 0
                ),
            }

    def periods_with_data(self, periods: list[str] | tuple[str, ...] | set[str]) -> list[str]:
        clean_periods = sorted({str(period or "").strip() for period in periods if str(period or "").strip()})
        if not clean_periods:
            return []
        placeholders = ", ".join("?" for _period in clean_periods)
        with self._database() as database:
            rows = database.execute(
                f"""
                SELECT DISTINCT period
                FROM customer_period_summary
                WHERE period IN ({placeholders})
                ORDER BY period
                """,
                clean_periods,
            ).fetchall()
        return [str(row["period"] or "") for row in rows]

    def duplicate_import_files(self, files: list[tuple[str, str]]) -> list[dict[str, object]]:
        clean_files = [
            (str(period or "").strip(), str(file_hash or "").strip())
            for period, file_hash in files
            if str(period or "").strip() and str(file_hash or "").strip()
        ]
        if not clean_files:
            return []
        duplicates: list[dict[str, object]] = []
        with self._database() as database:
            for period, file_hash in clean_files:
                row = database.execute(
                    """
                    SELECT id, run_id, file_name, period, file_hash, status
                    FROM customer_import_files
                    WHERE period = ? AND file_hash = ? AND status = 'COMPLETED'
                    LIMIT 1
                    """,
                    (period, file_hash),
                ).fetchone()
                if row is not None:
                    duplicates.append(dict(row))
        return duplicates

    def save_aggregation(
        self,
        result: CustomerAggregationResult,
        *,
        replace_periods: bool = False,
        created_by: str = "",
        duration_ms: int = 0,
    ) -> int:
        self._validate_aggregation_balance(result)
        self._validate_office_aggregation_balance(result)
        self._validate_debt_group_aggregation(result)
        periods = sorted({row.period for row in result.summaries if row.period})
        if not periods:
            raise CustomerDatabaseError("Khong co du lieu khach hang hop le de luu vao Customer.db.")
        self._assert_write_allowed()
        now = now_text()
        computer_name = os.environ.get("COMPUTERNAME", "")
        with self._database() as database:
            if not replace_periods:
                existing_periods = self._periods_with_data(database, periods)
                if existing_periods:
                    raise CustomerDatabaseError(
                        "Du lieu khach hang ky "
                        f"{', '.join(existing_periods)} da ton tai. Can xac nhan ghi de truoc khi import lai."
                    )
                duplicates = self._duplicate_import_files(
                    database,
                    [(file.period, file.file_hash) for file in result.files],
                )
                if duplicates:
                    duplicate = duplicates[0]
                    raise CustomerDatabaseError(
                        "File FTPLN da duoc import vao Customer.db: "
                        f"{duplicate.get('file_name')} ky {duplicate.get('period')}."
                    )
            else:
                self._delete_periods(database, periods)

            self._ensure_result_units(result, updated_by=created_by or "customer_import")

            cursor = database.execute(
                """
                INSERT INTO customer_import_runs(
                    period, source_folder, data_type, file_count, source_row_count,
                    customer_count, started_at, completed_at, status, error_message,
                    created_by, computer_name, personal_customer_count,
                    organization_customer_count, total_balance, short_term_balance,
                    medium_long_term_balance, other_balance,
                    multiple_officer_customer_count, unknown_ftp_code_count,
                    invalid_row_count, warning_count, duration_ms,
                    debt_group_valid_row_count, debt_group_1_row_count,
                    debt_group_2_row_count, debt_group_3_row_count,
                    debt_group_4_row_count, debt_group_5_row_count,
                    debt_group_unknown_row_count, debt_group_invalid_samples
                )
                VALUES (?, ?, ?, ?, ?, ?, ?, '', 'PENDING', '', ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    result.period,
                    result.source_folder,
                    CustomerDataType.NIM_DN.value,
                    result.file_count,
                    result.source_row_count,
                    result.customer_count,
                    now,
                    created_by,
                    computer_name,
                    result.personal_customer_count,
                    result.organization_customer_count,
                    result.total_balance,
                    result.short_term_balance,
                    result.medium_long_term_balance,
                    result.other_balance,
                    result.multiple_officer_customer_count,
                    len(result.unknown_ftp_codes),
                    result.invalid_row_count,
                    result.warning_count,
                    int(duration_ms),
                    result.debt_group_valid_row_count,
                    result.debt_group_1_row_count,
                    result.debt_group_2_row_count,
                    result.debt_group_3_row_count,
                    result.debt_group_4_row_count,
                    result.debt_group_5_row_count,
                    result.debt_group_unknown_row_count,
                    ", ".join(result.debt_group_invalid_samples),
                ),
            )
            run_id = int(cursor.lastrowid)
            self._insert_import_files(database, run_id, result, now)
            self._insert_period_summaries(database, run_id, result, now)
            self._insert_officer_rows(database, result, now)
            self._insert_office_rows(database, run_id, result, now)
            self._upsert_customer_master(database, result, now)
            self._upsert_officer_directory(database, result, now)
            database.execute(
                """
                UPDATE customer_import_runs
                SET completed_at = ?, status = 'COMPLETED', error_message = ''
                WHERE id = ?
                """,
                (now_text(), run_id),
            )
        return run_id

    def delete_import_run(self, run_id: int) -> None:
        self._assert_write_allowed()
        clean_run_id = int(run_id)
        with self._database() as database:
            rows = database.execute(
                """
                SELECT period, customer_code
                FROM customer_period_summary
                WHERE run_id = ?
                """,
                (clean_run_id,),
            ).fetchall()
            for row in rows:
                database.execute(
                    """
                    DELETE FROM customer_officer_period
                    WHERE period = ? AND customer_code = ?
                    """,
                    (row["period"], row["customer_code"]),
                )
                database.execute(
                    """
                    DELETE FROM customer_office_period
                    WHERE period = ? AND customer_code = ?
                    """,
                    (row["period"], row["customer_code"]),
                )
            database.execute("DELETE FROM customer_period_summary WHERE run_id = ?", (clean_run_id,))
            database.execute("DELETE FROM customer_import_runs WHERE id = ?", (clean_run_id,))

    def backup_database(self, destination: Path | None = None) -> Path:
        if destination is None:
            destination = self.database_path.parent / "backups" / f"customer-{now_text().replace(':', '')}.zip"
        destination = Path(destination)
        destination.parent.mkdir(parents=True, exist_ok=True)
        with tempfile.TemporaryDirectory(prefix="customer-backup-") as tmp_dir:
            snapshot = Path(tmp_dir) / self.database_path.name
            with closing(sqlite3.connect(self.database_path, timeout=10)) as source:
                with closing(sqlite3.connect(snapshot)) as target:
                    source.backup(target)
            manifest = {
                "format": "agribank-v3-customer-backup",
                "created_at": now_text(),
                "database": self.database_path.name,
            }
            with zipfile.ZipFile(destination, "w", compression=zipfile.ZIP_DEFLATED) as archive:
                archive.writestr("manifest.json", json.dumps(manifest, ensure_ascii=False, indent=2))
                archive.write(snapshot, self.database_path.name)
        return destination

    def restore_database(self, source_path: Path) -> Path:
        source_path = Path(source_path)
        if not zipfile.is_zipfile(source_path):
            raise CustomerDatabaseError("File phuc hoi Customer.db khong hop le.")
        with CustomerDatabaseOperationLock("khôi phục Customer.db"):
            safety_backup = self.backup_database()
            with tempfile.TemporaryDirectory(prefix="customer-restore-") as tmp_dir:
                tmp_root = Path(tmp_dir)
                with zipfile.ZipFile(source_path, "r") as archive:
                    manifest = json.loads(archive.read("manifest.json"))
                    if manifest.get("format") != "agribank-v3-customer-backup":
                        raise CustomerDatabaseError("Goi sao luu Customer.db khong dung dinh dang.")
                    database_name = str(manifest.get("database") or self.database_path.name)
                    if Path(database_name).name != self.database_path.name:
                        raise CustomerDatabaseError("Goi sao luu khong trung Customer.db hien tai.")
                    candidate = tmp_root / database_name
                    candidate.write_bytes(archive.read(database_name))
                with closing(sqlite3.connect(candidate)) as database:
                    integrity = str(database.execute("PRAGMA integrity_check").fetchone()[0])
                    if integrity.casefold() != "ok":
                        raise CustomerDatabaseError(f"Snapshot Customer.db bi loi: {integrity}")
                with closing(sqlite3.connect(self.database_path, timeout=10)) as current:
                    current.execute("PRAGMA wal_checkpoint(TRUNCATE)")
                    current.execute("PRAGMA journal_mode = DELETE")
                for suffix in ("-wal", "-shm"):
                    Path(f"{self.database_path}{suffix}").unlink(missing_ok=True)
                replacement = self.database_path.with_name(f".customer-restore-{uuid4().hex}.db")
                shutil.copy2(candidate, replacement)
                os.replace(replacement, self.database_path)
            self.ensure_schema()
            return safety_backup

    def optimize_database(self, *, vacuum: bool = False) -> dict[str, object]:
        operation = "thu hồi dung lượng Customer.db" if vacuum else "tối ưu Customer.db"
        if not vacuum:
            self._assert_write_allowed()
        started = perf_counter()
        before = self.database_path.stat().st_size if self.database_path.is_file() else 0
        backup_path = ""
        try:
            context = CustomerDatabaseOperationLock(operation) if vacuum else nullcontext()
            with context:
                if vacuum:
                    backup_path = str(self.backup_database())
                with closing(self.connect()) as database:
                    database.execute("PRAGMA optimize")
                    database.commit()
                    if vacuum:
                        database.execute("PRAGMA wal_checkpoint(TRUNCATE)")
                        database.execute("VACUUM")
                        database.commit()
        except Exception:
            LOGGER.exception("Customer.db maintenance failed: %s", operation)
            raise
        after = self.database_path.stat().st_size if self.database_path.is_file() else 0
        duration_ms = int((perf_counter() - started) * 1000)
        LOGGER.info(
            "Customer.db maintenance finished: operation=%s before=%s after=%s backup=%s duration_ms=%s",
            operation,
            before,
            after,
            backup_path,
            duration_ms,
        )
        return {
            "before_size_bytes": before,
            "after_size_bytes": after,
            "recovered_bytes": max(0, before - after),
            "duration_ms": duration_ms,
            "vacuum": bool(vacuum),
            "backup_path": backup_path,
        }

    def check_database(self, *, full: bool = False) -> dict[str, object]:
        mode = "integrity_check" if full else "quick_check"
        started = perf_counter()
        try:
            with closing(self.connect()) as database:
                rows = database.execute(f"PRAGMA {mode}").fetchall()
        except Exception:
            LOGGER.exception("Customer.db %s failed", mode)
            raise
        messages = [str(row[0]) for row in rows]
        ok = len(messages) == 1 and messages[0].casefold() == "ok"
        return {
            "mode": mode,
            "ok": ok,
            "messages": messages,
            "duration_ms": int((perf_counter() - started) * 1000),
        }

    def distinct_periods(self) -> list[str]:
        with self._database() as database:
            rows = database.execute(
                """
                SELECT DISTINCT period
                FROM customer_period_summary
                WHERE period <> ''
                ORDER BY period
                """
            ).fetchall()
        return [str(row["period"] or "") for row in rows]

    def has_period_data(self) -> bool:
        with self._database() as database:
            row = database.execute(
                """
                SELECT 1
                FROM customer_period_summary
                WHERE period <> ''
                LIMIT 1
                """
            ).fetchone()
        return row is not None

    def officer_directory_count(self) -> int:
        with self._database() as database:
            return self._count(database, "customer_officer_directory")

    def officer_override_count(self) -> int:
        with self._database() as database:
            return self._count(database, "customer_officer_override")

    def action_log_count(self) -> int:
        with self._database() as database:
            return self._count(database, "customer_action_log")

    def has_period(self, period: str) -> bool:
        clean_period = str(period or "").strip()
        if not clean_period:
            return False
        with self._database() as database:
            row = database.execute(
                """
                SELECT 1
                FROM customer_period_summary
                WHERE period = ?
                LIMIT 1
                """,
                (clean_period,),
            ).fetchone()
        return row is not None

    def has_office_detail_for_period(self, period: str) -> bool:
        clean_period = str(period or "").strip()
        if not clean_period:
            return False
        with self._database() as database:
            row = database.execute(
                """
                SELECT 1
                FROM customer_office_period
                WHERE period = ?
                  AND total_balance > 0
                LIMIT 1
                """,
                (clean_period,),
            ).fetchone()
        return row is not None

    def distinct_branch_codes(self, filters: CustomerFilters | None = None) -> list[str]:
        filters = filters or CustomerFilters()
        where, params = _summary_where(filters, exclude={"branch_code"})
        with self._database() as database:
            rows = database.execute(
                f"""
                SELECT DISTINCT s.branch_code AS value
                FROM customer_period_summary s
                {where} AND s.branch_code <> ''
                ORDER BY s.branch_code
                """,
                params,
            ).fetchall()
        return [str(row["value"] or "") for row in rows]

    def _branch_display(self, branch_code: object) -> str:
        return self.unit_directory.get_branch_display_name(branch_code)

    def distinct_offices(
        self,
        period: str,
        *,
        branch_code: str = "",
    ) -> list[dict[str, object]]:
        clean_period = str(period or "").strip()
        if not clean_period:
            return []
        clauses = ["period = ?", "total_balance > 0"]
        params: list[object] = [clean_period]
        branch = str(branch_code or "").strip()
        if branch:
            clauses.append("branch_code = ?")
            params.append(branch)
        where = "WHERE " + " AND ".join(clauses)
        with self._database() as database:
            rows = database.execute(
                f"""
                SELECT DISTINCT branch_code, trctcd, office_code, office_name, office_type
                FROM customer_office_period
                {where}
                ORDER BY branch_code,
                    CASE
                        WHEN office_type = 'HEAD_OFFICE' THEN 0
                        WHEN office_type = 'TRANSACTION_OFFICE' THEN 1
                        ELSE 2
                    END,
                    trctcd COLLATE NOCASE,
                    office_code COLLATE NOCASE
                """,
                params,
            ).fetchall()
        return [_finalize_office_option(dict(row), self.unit_directory) for row in rows]

    def distinct_customer_types(self, filters: CustomerFilters | None = None) -> list[str]:
        filters = filters or CustomerFilters()
        where, params = _summary_where(filters, exclude={"customer_type"})
        with self._database() as database:
            rows = database.execute(
                f"""
                SELECT DISTINCT s.customer_type AS value
                FROM customer_period_summary s
                {where} AND s.customer_type <> ''
                ORDER BY s.customer_type
                """,
                params,
            ).fetchall()
        return [str(row["value"] or "") for row in rows]

    def distinct_officers(self, filters: CustomerFilters | None = None) -> list[dict[str, object]]:
        filters = filters or CustomerFilters()
        where, params = _summary_where(filters, exclude={"officer"})
        code_expr = _override_value_sql("s", "officer_code", fallback="s.primary_officer_code")
        name_expr = _override_value_sql("s", "officer_name", fallback="s.primary_officer_name", null_if_empty=True)
        with self._database() as database:
            rows = database.execute(
                f"""
                SELECT DISTINCT {code_expr} AS officer_code, {name_expr} AS officer_name
                FROM customer_period_summary s
                {where}
                ORDER BY officer_name COLLATE NOCASE, officer_code COLLATE NOCASE
                """,
                params,
            ).fetchall()
        return [dict(row) for row in rows if row["officer_code"] or row["officer_name"]]

    def query_customer_list(
        self,
        filters: CustomerFilters | None = None,
        *,
        page: int = 1,
        page_size: int = 100,
        sort_by: str = "period",
        sort_desc: bool = True,
    ) -> PageResult:
        filters = filters or CustomerFilters()
        page = max(1, int(page or 1))
        page_size = max(1, min(5000, int(page_size or 100)))
        offset = (page - 1) * page_size
        where, params = _summary_where(filters)
        base_sql = _customer_list_base_sql(where)
        order_sql = _order_sql(
            sort_by,
            sort_desc,
            {
                "period": "period",
                "customer_code": "customer_code",
                "customer_name": "customer_name",
                "customer_type": "customer_type",
                "branch_code": "branch_code",
                "effective_officer_name": "effective_officer_name",
                "imported_officer_name": "imported_officer_name",
                "officer_count": "officer_count",
                "total_balance": "total_balance",
                "short_term_balance": "short_term_balance",
                "medium_long_term_balance": "medium_long_term_balance",
                "other_balance": "other_balance",
                "medium_long_ratio": "medium_long_ratio",
                "average_rate": "average_rate",
                "nim_before": "nim_before",
                "nim_after": "nim_after",
                "source_loan_count": "source_loan_count",
                "override_status": "override_status",
            },
            default="period",
        )
        with self._database() as database:
            total_rows = int(
                database.execute(f"SELECT COUNT(*) FROM ({base_sql}) q", params).fetchone()[0] or 0
            )
            rows = database.execute(
                f"SELECT * FROM ({base_sql}) q {order_sql} LIMIT ? OFFSET ?",
                (*params, page_size, offset),
            ).fetchall()
        return PageResult(rows=[dict(row) for row in rows], total_rows=total_rows, page=page, page_size=page_size)

    def all_customer_rows(
        self,
        filters: CustomerFilters | None = None,
        *,
        sort_by: str = "period",
        sort_desc: bool = True,
    ) -> list[dict[str, object]]:
        rows: list[dict[str, object]] = []
        page = 1
        while True:
            result = self.query_customer_list(
                filters,
                page=page,
                page_size=1000,
                sort_by=sort_by,
                sort_desc=sort_desc,
            )
            rows.extend(result.rows)
            if page * result.page_size >= result.total_rows:
                break
            page += 1
        return rows

    def has_debt_group_data(self, period: str, filters: CustomerFilters | None = None) -> bool:
        clean_period = str(period or "").strip()
        if not clean_period:
            return False
        filters = replace(filters or CustomerFilters(), current_period=clean_period, debt_group="")
        where, params = _summary_where(filters, exclude={"debt_group"})
        with self._database() as database:
            row = database.execute(
                f"""
                SELECT MAX(COALESCE(s.has_debt_group_data, 0)) AS has_data
                FROM customer_period_summary s
                {where}
                """,
                params,
            ).fetchone()
        return bool(row and int(row["has_data"] or 0))

    def get_debt_quality_kpis(
        self,
        report_period: str,
        filters: CustomerFilters | None = None,
    ) -> dict[str, object]:
        filters = _debt_report_filters(filters, report_period)
        where, params = _summary_where(filters)
        with self._database() as database:
            row = database.execute(
                f"""
                SELECT
                    MAX(COALESCE(s.has_debt_group_data, 0)) AS has_debt_group_data,
                    COUNT(DISTINCT s.customer_code) AS customer_count,
                    COALESCE(SUM(s.total_balance), 0) AS total_balance,
                    COALESCE(SUM(s.debt_group_1_balance), 0) AS debt_group_1_balance,
                    COALESCE(SUM(s.debt_group_2_balance), 0) AS debt_group_2_balance,
                    COALESCE(SUM(s.debt_group_3_balance), 0) AS debt_group_3_balance,
                    COALESCE(SUM(s.debt_group_4_balance), 0) AS debt_group_4_balance,
                    COALESCE(SUM(s.debt_group_5_balance), 0) AS debt_group_5_balance,
                    COALESCE(SUM(s.debt_group_unknown_balance), 0) AS debt_group_unknown_balance,
                    COALESCE(SUM(s.interest_rate_numerator), 0) AS interest_rate_numerator,
                    COALESCE(SUM(s.nim_before_numerator), 0) AS nim_before_numerator,
                    COALESCE(SUM(s.nim_after_numerator), 0) AS nim_after_numerator,
                    SUM(CASE WHEN s.debt_group_2_balance > 0 THEN 1 ELSE 0 END) AS attention_customer_count,
                    SUM(CASE WHEN (
                        s.debt_group_3_balance + s.debt_group_4_balance + s.debt_group_5_balance
                    ) > 0 THEN 1 ELSE 0 END) AS bad_debt_customer_count,
                    SUM(CASE WHEN s.debt_group_unknown_balance > 0 THEN 1 ELSE 0 END) AS unknown_customer_count,
                    SUM(CASE WHEN s.worst_debt_group = '01' THEN 1 ELSE 0 END) AS worst_group_1_customer_count,
                    SUM(CASE WHEN s.worst_debt_group = '02' THEN 1 ELSE 0 END) AS worst_group_2_customer_count,
                    SUM(CASE WHEN s.worst_debt_group = '03' THEN 1 ELSE 0 END) AS worst_group_3_customer_count,
                    SUM(CASE WHEN s.worst_debt_group = '04' THEN 1 ELSE 0 END) AS worst_group_4_customer_count,
                    SUM(CASE WHEN s.worst_debt_group = '05' THEN 1 ELSE 0 END) AS worst_group_5_customer_count
                FROM customer_period_summary s
                {where}
                """,
                params,
            ).fetchone()
        output = dict(row) if row is not None else {}
        return _finalize_debt_group_metrics(output)

    def get_debt_group_summary(
        self,
        report_period: str,
        filters: CustomerFilters | None = None,
    ) -> list[dict[str, object]]:
        filters = _debt_report_filters(filters, report_period)
        where, params = _summary_where(filters)
        select_parts = ["COALESCE(SUM(s.total_balance), 0) AS total_balance", "MAX(COALESCE(s.has_debt_group_data, 0)) AS has_debt_group_data"]
        for suffix in DEBT_GROUP_SUFFIXES:
            select_parts.extend(
                [
                    f"COALESCE(SUM(s.debt_group_{suffix}_balance), 0) AS debt_group_{suffix}_balance",
                    f"COALESCE(SUM(s.debt_group_{suffix}_interest_numerator), 0) AS debt_group_{suffix}_interest_numerator",
                    f"COALESCE(SUM(s.debt_group_{suffix}_nim_before_numerator), 0) AS debt_group_{suffix}_nim_before_numerator",
                    f"COALESCE(SUM(s.debt_group_{suffix}_nim_after_numerator), 0) AS debt_group_{suffix}_nim_after_numerator",
                    f"SUM(CASE WHEN s.debt_group_{suffix}_balance > 0 THEN 1 ELSE 0 END) AS debt_group_{suffix}_customer_count",
                ]
            )
        with self._database() as database:
            row = database.execute(
                f"""
                SELECT {", ".join(select_parts)}
                FROM customer_period_summary s
                {where}
                """,
                params,
            ).fetchone()
        data = dict(row) if row is not None else {}
        total_balance = _number(data.get("total_balance"))
        rows: list[dict[str, object]] = []
        for suffix in DEBT_GROUP_SUFFIXES:
            balance = _number(data.get(f"debt_group_{suffix}_balance"))
            rows.append(
                {
                    "debt_group": DEBT_GROUP_LABELS[suffix],
                    "debt_group_key": suffix,
                    "balance": balance,
                    "share_ratio": _ratio(balance * 100, total_balance) if total_balance else None,
                    "customer_count": int(data.get(f"debt_group_{suffix}_customer_count") or 0),
                    "average_rate": _ratio(
                        _number(data.get(f"debt_group_{suffix}_interest_numerator")),
                        balance,
                    ) if balance else None,
                    "nim_before": _ratio(
                        _number(data.get(f"debt_group_{suffix}_nim_before_numerator")),
                        balance,
                    ) if balance else None,
                    "nim_after": _ratio(
                        _number(data.get(f"debt_group_{suffix}_nim_after_numerator")),
                        balance,
                    ) if balance else None,
                    "has_debt_group_data": bool(data.get("has_debt_group_data")),
                }
            )
        return rows

    def get_debt_group_trend(
        self,
        period_from: str,
        period_to: str,
        filters: CustomerFilters | None = None,
    ) -> list[dict[str, object]]:
        filters = _trend_filters(replace(filters or CustomerFilters(), current_period=""), period_from, period_to)
        where, params = _summary_where(filters)
        with self._database() as database:
            rows = database.execute(
                f"""
                SELECT
                    s.period,
                    MAX(COALESCE(s.has_debt_group_data, 0)) AS has_debt_group_data,
                    COALESCE(SUM(s.total_balance), 0) AS total_balance,
                    COALESCE(SUM(s.debt_group_2_balance), 0) AS attention_balance,
                    COALESCE(SUM(s.debt_group_3_balance + s.debt_group_4_balance + s.debt_group_5_balance), 0) AS bad_debt_balance
                FROM customer_period_summary s
                {where}
                GROUP BY s.period
                ORDER BY s.period ASC
                """,
                params,
            ).fetchall()
        output: list[dict[str, object]] = []
        for row in rows:
            item = dict(row)
            total_balance = _number(item.get("total_balance"))
            attention_balance = _number(item.get("attention_balance"))
            bad_debt_balance = _number(item.get("bad_debt_balance"))
            item["attention_ratio"] = _ratio(attention_balance * 100, total_balance) if total_balance else None
            item["bad_debt_ratio"] = _ratio(bad_debt_balance * 100, total_balance) if total_balance else None
            output.append(item)
        return output

    def get_debt_group_by_branch(
        self,
        report_period: str,
        filters: CustomerFilters | None = None,
        *,
        limit: int = 100,
        offset: int = 0,
        sort_by: str = "bad_debt_ratio",
        sort_desc: bool = True,
    ) -> list[dict[str, object]]:
        filters = _debt_report_filters(filters, report_period)
        rows = self._debt_group_by_branch_page(filters, limit=limit, offset=offset, sort_by=sort_by, sort_desc=sort_desc)
        _enrich_debt_group_branch_rows(rows, self.unit_directory)
        return rows

    def _debt_group_by_branch_page(
        self,
        filters: CustomerFilters,
        *,
        limit: int,
        offset: int,
        sort_by: str,
        sort_desc: bool,
    ) -> list[dict[str, object]]:
        base_sql, params = _debt_group_by_branch_base_sql(filters)
        order_sql = _debt_group_order_sql(sort_by, sort_desc, default="bad_debt_ratio")
        with self._database() as database:
            rows = [
                dict(row)
                for row in database.execute(
                    f"SELECT * FROM ({base_sql}) q {order_sql} LIMIT ? OFFSET ?",
                    (*params, max(1, int(limit or 100)), max(0, int(offset or 0))),
                ).fetchall()
            ]
        _finalize_debt_group_rows(rows)
        return rows

    def count_debt_group_by_branch(
        self,
        report_period: str,
        filters: CustomerFilters | None = None,
    ) -> int:
        filters = _debt_report_filters(filters, report_period)
        base_sql, params = _debt_group_by_branch_base_sql(filters)
        with self._database() as database:
            return int(database.execute(f"SELECT COUNT(*) FROM ({base_sql}) q", params).fetchone()[0] or 0)

    def query_debt_group_by_branch(
        self,
        report_period: str,
        filters: CustomerFilters | None = None,
        *,
        page: int = 1,
        page_size: int = 100,
        sort_by: str = "bad_debt_ratio",
        sort_desc: bool = True,
    ) -> PageResult:
        page = max(1, int(page or 1))
        page_size = max(1, min(5000, int(page_size or 100)))
        total_rows = self.count_debt_group_by_branch(report_period, filters)
        rows = self.get_debt_group_by_branch(
            report_period,
            filters,
            limit=page_size,
            offset=(page - 1) * page_size,
            sort_by=sort_by,
            sort_desc=sort_desc,
        )
        return PageResult(rows=rows, total_rows=total_rows, page=page, page_size=page_size)

    def get_debt_group_by_officer(
        self,
        report_period: str,
        filters: CustomerFilters | None = None,
        *,
        limit: int = 100,
        offset: int = 0,
        sort_by: str = "bad_debt_ratio",
        sort_desc: bool = True,
    ) -> list[dict[str, object]]:
        filters = _debt_report_filters(filters, report_period)
        base_sql, params = _debt_group_by_officer_base_sql(filters)
        order_sql = _debt_group_order_sql(sort_by, sort_desc, default="bad_debt_ratio")
        with self._database() as database:
            rows = [
                dict(row)
                for row in database.execute(
                    f"SELECT * FROM ({base_sql}) q {order_sql} LIMIT ? OFFSET ?",
                    (*params, max(1, int(limit or 100)), max(0, int(offset or 0))),
                ).fetchall()
            ]
        _finalize_debt_group_rows(rows)
        _enrich_debt_group_branch_rows(rows, self.unit_directory)
        return rows

    def query_debt_group_by_officer(
        self,
        report_period: str,
        filters: CustomerFilters | None = None,
        *,
        page: int = 1,
        page_size: int = 100,
        sort_by: str = "bad_debt_ratio",
        sort_desc: bool = True,
    ) -> PageResult:
        page = max(1, int(page or 1))
        page_size = max(1, min(5000, int(page_size or 100)))
        filters = _debt_report_filters(filters, report_period)
        base_sql, params = _debt_group_by_officer_base_sql(filters)
        with self._database() as database:
            total_rows = int(database.execute(f"SELECT COUNT(*) FROM ({base_sql}) q", params).fetchone()[0] or 0)
        rows = self.get_debt_group_by_officer(
            report_period,
            filters,
            limit=page_size,
            offset=(page - 1) * page_size,
            sort_by=sort_by,
            sort_desc=sort_desc,
        )
        return PageResult(rows=rows, total_rows=total_rows, page=page, page_size=page_size)

    def get_officer_debt_group_history(
        self,
        *,
        officer_code: str = "",
        officer_name: str = "",
        filters: CustomerFilters | None = None,
    ) -> list[dict[str, object]]:
        filters = filters or CustomerFilters()
        clauses = ["1 = 1"]
        params: list[object] = []
        if filters.current_period:
            clauses.append("op.period = ?")
            params.append(filters.current_period)
        else:
            if filters.period_from:
                clauses.append("op.period >= ?")
                params.append(filters.period_from)
            if filters.period_to:
                clauses.append("op.period <= ?")
                params.append(filters.period_to)
        code = clean_filter_text(officer_code)
        name = clean_filter_text(officer_name)
        if code:
            clauses.append("op.officer_code = ?")
            params.append(code)
        elif name:
            clauses.append("(op.officer_name = ? OR op.officer_name LIKE ?)")
            params.extend([name, f"%{name}%"])
        if filters.branch_code:
            clauses.append("(op.branch_code = ? OR s.branch_code = ?)")
            params.extend([filters.branch_code, filters.branch_code])
        if filters.customer_type:
            clauses.append("s.customer_type = ?")
            params.append(filters.customer_type)
        if filters.loan_term:
            if filters.loan_term == "SHORT_TERM":
                clauses.append("s.short_term_balance > 0")
            elif filters.loan_term == "MEDIUM_LONG_TERM":
                clauses.append("s.medium_long_term_balance > 0")
            elif filters.loan_term == "OTHER":
                clauses.append("s.other_balance > 0")
        if filters.debt_group:
            _append_debt_group_filter(clauses, params, filters.debt_group, alias="op", total_column="balance_managed")
        where = "WHERE " + " AND ".join(clauses)
        with self._database() as database:
            rows = [
                dict(row)
                for row in database.execute(
                    f"""
                    SELECT
                        op.period,
                        {_debt_group_aggregate_select_sql(alias="op", total_column="balance_managed", customer_count_expr="COUNT(DISTINCT {alias}.customer_code)")}
                    FROM customer_officer_period op
                    LEFT JOIN customer_period_summary s
                        ON s.period = op.period AND s.customer_code = op.customer_code
                    {where}
                    GROUP BY op.period
                    ORDER BY op.period ASC
                    """,
                    params,
                ).fetchall()
            ]
        _finalize_debt_group_rows(rows)
        return rows

    def get_debt_group_by_office(
        self,
        report_period: str,
        filters: CustomerFilters | None = None,
        *,
        limit: int = 100,
        offset: int = 0,
        sort_by: str = "bad_debt_ratio",
        sort_desc: bool = True,
    ) -> list[dict[str, object]]:
        filters = _debt_report_filters(filters, report_period)
        base_sql, params = _debt_group_by_office_base_sql(filters)
        order_sql = _debt_group_order_sql(sort_by, sort_desc, default="bad_debt_ratio")
        with self._database() as database:
            rows = [
                dict(row)
                for row in database.execute(
                    f"SELECT * FROM ({base_sql}) q {order_sql} LIMIT ? OFFSET ?",
                    (*params, max(1, int(limit or 100)), max(0, int(offset or 0))),
                ).fetchall()
            ]
        _finalize_debt_group_rows(rows)
        _enrich_debt_group_branch_rows(rows, self.unit_directory)
        return rows

    def get_debt_group_customers(
        self,
        report_period: str,
        filters: CustomerFilters | None = None,
        *,
        limit: int = 100,
        offset: int = 0,
        sort: str = "bad_debt_ratio",
        sort_by: str | None = None,
        sort_desc: bool = True,
    ) -> list[dict[str, object]]:
        filters = _debt_report_filters(filters, report_period)
        base_sql, params = _debt_group_customer_base_sql(filters)
        order_sql = _debt_group_order_sql(sort_by or sort, sort_desc, default="bad_debt_ratio")
        with self._database() as database:
            rows = [
                dict(row)
                for row in database.execute(
                    f"SELECT * FROM ({base_sql}) q {order_sql} LIMIT ? OFFSET ?",
                    (*params, max(1, int(limit or 100)), max(0, int(offset or 0))),
                ).fetchall()
            ]
        _finalize_debt_group_rows(rows)
        _enrich_debt_group_branch_rows(rows, self.unit_directory)
        return rows

    def count_debt_group_customers(
        self,
        report_period: str,
        filters: CustomerFilters | None = None,
    ) -> int:
        filters = _debt_report_filters(filters, report_period)
        base_sql, params = _debt_group_customer_base_sql(filters)
        with self._database() as database:
            return int(database.execute(f"SELECT COUNT(*) FROM ({base_sql}) q", params).fetchone()[0] or 0)

    def query_debt_group_customers(
        self,
        report_period: str,
        filters: CustomerFilters | None = None,
        *,
        page: int = 1,
        page_size: int = 100,
        sort_by: str = "bad_debt_ratio",
        sort_desc: bool = True,
    ) -> PageResult:
        page = max(1, int(page or 1))
        page_size = max(1, min(5000, int(page_size or 100)))
        total_rows = self.count_debt_group_customers(report_period, filters)
        rows = self.get_debt_group_customers(
            report_period,
            filters,
            limit=page_size,
            offset=(page - 1) * page_size,
            sort_by=sort_by,
            sort_desc=sort_desc,
        )
        return PageResult(rows=rows, total_rows=total_rows, page=page, page_size=page_size)

    def get_customer_debt_group_history(
        self,
        customer_code: str,
        period_from: str | None = None,
        period_to: str | None = None,
    ) -> list[dict[str, object]]:
        code = str(customer_code or "").strip()
        if not code:
            return []
        clauses = ["s.customer_code = ?"]
        params: list[object] = [code]
        if period_from:
            clauses.append("s.period >= ?")
            params.append(str(period_from))
        if period_to:
            clauses.append("s.period <= ?")
            params.append(str(period_to))
        where = "WHERE " + " AND ".join(clauses)
        base_sql = _debt_group_customer_select_sql(where)
        with self._database() as database:
            rows = [dict(row) for row in database.execute(f"{base_sql} ORDER BY period ASC", params).fetchall()]
        _finalize_debt_group_rows(rows)
        _enrich_debt_group_branch_rows(rows, self.unit_directory)
        return rows

    def get_customer_debt_group_detail(
        self,
        customer_code: str,
        report_period: str,
    ) -> dict[str, object] | None:
        rows = self.get_customer_debt_group_history(customer_code, report_period, report_period)
        return rows[0] if rows else None

    def query_cross_branch_customers(
        self,
        period: str,
        filters: CustomerFilters | None = None,
        *,
        minimum_branch_count: int = 2,
        scope_type: object = "cross_branch",
        office_code: str = "",
        office_filter_mode: str = "actual",
        page: int = 1,
        page_size: int = 100,
        sort_by: str = "branch_count",
        sort_desc: bool = True,
    ) -> PageResult:
        page = max(1, int(page or 1))
        page_size = max(1, min(5000, int(page_size or 100)))
        offset = (page - 1) * page_size
        total_rows = self.count_cross_branch_customers(
            period,
            filters,
            minimum_branch_count=minimum_branch_count,
            scope_type=scope_type,
            office_code=office_code,
            office_filter_mode=office_filter_mode,
        )
        rows = self.get_cross_branch_customers(
            period,
            filters,
            minimum_branch_count=minimum_branch_count,
            scope_type=scope_type,
            office_code=office_code,
            office_filter_mode=office_filter_mode,
            limit=page_size,
            offset=offset,
            sort_by=sort_by,
            sort_order="desc" if sort_desc else "asc",
        )
        return PageResult(rows=rows, total_rows=total_rows, page=page, page_size=page_size)

    def query_cross_branch_tab_payload(
        self,
        period: str,
        filters: CustomerFilters | None = None,
        *,
        minimum_branch_count: int = 2,
        scope_type: object = "cross_branch",
        office_code: str = "",
        office_filter_mode: str = "actual",
        page: int = 1,
        page_size: int = 100,
        sort_by: str = "branch_count",
        sort_desc: bool = True,
    ) -> dict[str, object]:
        period = str(period or "").strip()
        if not period:
            periods = self.distinct_periods()
            period = periods[-1] if periods else ""
        page = max(1, int(page or 1))
        page_size = max(1, min(5000, int(page_size or 100)))
        offset = (page - 1) * page_size
        filters = filters or CustomerFilters()
        minimum = max(2, int(minimum_branch_count or 2))
        benchmark: dict[str, dict[str, object]] = {}
        statement_counter = {"count": 0}

        def trace(statement: str) -> None:
            text = str(statement or "").strip()
            if text:
                statement_counter["count"] += 1

        def empty_payload(message: str = "") -> dict[str, object]:
            return {
                "result": PageResult(rows=[], total_rows=0, page=page, page_size=page_size),
                "kpis": _empty_cross_branch_kpis(),
                "benchmark": benchmark,
                "sql_statement_count": statement_counter["count"],
                "empty_message": message,
            }

        with closing(self.connect()) as database:
            database.set_trace_callback(trace)

            def measured(label: str, function):
                before_sql = statement_counter["count"]
                started = perf_counter()
                value = function()
                benchmark[label] = {
                    "elapsed_ms": (perf_counter() - started) * 1000,
                    "sql_count": statement_counter["count"] - before_sql,
                }
                return value

            if not period:
                return empty_payload("Chưa có dữ liệu khách hàng.")
            has_period = measured(
                "validate_period",
                lambda: database.execute(
                    """
                    SELECT 1
                    FROM customer_period_summary
                    WHERE period = ?
                    LIMIT 1
                    """,
                    (period,),
                ).fetchone()
                is not None,
            )
            if not has_period:
                return empty_payload(f"Không có dữ liệu khách hàng cho kỳ {period}.")
            has_office_detail = measured(
                "validate_office_detail",
                lambda: database.execute(
                    """
                    SELECT 1
                    FROM customer_office_period
                    WHERE period = ?
                      AND total_balance > 0
                    LIMIT 1
                    """,
                    (period,),
                ).fetchone()
                is not None,
            )
            if not has_office_detail:
                return empty_payload(
                    f"Kỳ {period} chưa có dữ liệu chi tiết Hội sở/Phòng giao dịch để phân tích liên chi nhánh."
                )
            scopes = _normalize_cross_branch_scopes(scope_type)
            if scopes == ("cross_branch",):
                branch_count = measured(
                    "fast_scope_branch_count",
                    lambda: int(
                        database.execute(
                            """
                            SELECT COUNT(DISTINCT branch_code)
                            FROM customer_office_period
                            WHERE period = ?
                              AND total_balance > 0
                            """,
                            (period,),
                        ).fetchone()[0]
                        or 0
                    ),
                )
                if branch_count < minimum:
                    return empty_payload("Không có khách hàng vay liên chi nhánh phù hợp với bộ lọc.")
            sql, params = _cross_branch_candidate_sql(
                period,
                filters,
                minimum_branch_count=minimum,
                scope_type=scope_type,
                office_code=office_code,
                office_filter_mode=office_filter_mode,
            )

            def materialize_candidates() -> None:
                database.execute("DROP TABLE IF EXISTS temp_cross_branch_candidates")
                database.execute(f"CREATE TEMP TABLE temp_cross_branch_candidates AS {sql}", params)
                database.execute(
                    """
                    CREATE INDEX IF NOT EXISTS temp_idx_cross_branch_candidates_sequence
                    ON temp_cross_branch_candidates(period, customer_sequence)
                    """
                )

            measured("candidate_materialize", materialize_candidates)
            total_rows = measured(
                "count",
                lambda: int(database.execute("SELECT COUNT(*) FROM temp_cross_branch_candidates").fetchone()[0] or 0),
            )
            kpis = measured("kpi", lambda: _cross_branch_kpis_from_temp(database, period, self.unit_directory))
            order_sql = _cross_branch_order_sql(sort_by, "desc" if sort_desc else "asc")
            page_rows = measured(
                "page",
                lambda: [
                    dict(row)
                    for row in database.execute(
                        f"""
                        SELECT *
                        FROM temp_cross_branch_candidates
                        {order_sql}
                        LIMIT ? OFFSET ?
                        """,
                        (page_size, offset),
                    ).fetchall()
                ],
            )
            detail_map = measured(
                "enrichment",
                lambda: CustomerRepository._cross_branch_detail_rows_for_keys_from_database(
                    database,
                    [
                        (str(row.get("period") or ""), str(row.get("customer_sequence") or ""))
                        for row in page_rows
                    ],
                ),
            )

            def format_rows() -> list[dict[str, object]]:
                output: list[dict[str, object]] = []
                for index, row in enumerate(page_rows, start=offset + 1):
                    key = (str(row.get("period") or ""), str(row.get("customer_sequence") or ""))
                    item = self._finalize_cross_branch_summary(dict(row), detail_rows=detail_map.get(key, []))
                    item["rank"] = index
                    output.append(item)
                return output

            rows = measured("format_rows", format_rows)
            return {
                "result": PageResult(rows=rows, total_rows=total_rows, page=page, page_size=page_size),
                "kpis": kpis,
                "benchmark": benchmark,
                "sql_statement_count": statement_counter["count"],
                "empty_message": "" if total_rows else "Không có khách hàng vay liên chi nhánh phù hợp với bộ lọc.",
            }

    def get_cross_branch_customers(
        self,
        period: str,
        filters: CustomerFilters | None = None,
        *,
        minimum_branch_count: int = 2,
        scope_type: object = "cross_branch",
        office_code: str = "",
        office_filter_mode: str = "actual",
        limit: int = 100,
        offset: int = 0,
        sort_by: str = "branch_count",
        sort_order: str = "desc",
    ) -> list[dict[str, object]]:
        period = str(period or "").strip()
        if not period:
            periods = self.distinct_periods()
            period = periods[-1] if periods else ""
        if not period:
            return []
        if not self.has_office_detail_for_period(period):
            return []
        sql, params = _cross_branch_candidate_sql(
            period,
            filters or CustomerFilters(),
            minimum_branch_count=minimum_branch_count,
            scope_type=scope_type,
            office_code=office_code,
            office_filter_mode=office_filter_mode,
        )
        order_sql = _cross_branch_order_sql(sort_by, sort_order)
        limit = max(1, min(5000, int(limit or 100)))
        offset = max(0, int(offset or 0))
        with self._database() as database:
            rows = database.execute(
                f"{sql} {order_sql} LIMIT ? OFFSET ?",
                (*params, limit, offset),
            ).fetchall()
        detail_map = self._cross_branch_detail_rows_for_keys(
            [
                (str(row["period"] or ""), str(row["customer_sequence"] or ""))
                for row in rows
            ]
        )
        output: list[dict[str, object]] = []
        for index, row in enumerate(rows, start=offset + 1):
            key = (str(row["period"] or ""), str(row["customer_sequence"] or ""))
            item = self._finalize_cross_branch_summary(dict(row), detail_rows=detail_map.get(key, []))
            item["rank"] = index
            output.append(item)
        return output

    def all_cross_branch_customers(
        self,
        period: str,
        filters: CustomerFilters | None = None,
        *,
        minimum_branch_count: int = 2,
        scope_type: object = "cross_branch",
        office_code: str = "",
        office_filter_mode: str = "actual",
        sort_by: str = "branch_count",
        sort_desc: bool = True,
    ) -> list[dict[str, object]]:
        rows: list[dict[str, object]] = []
        offset = 0
        page_size = 1000
        while True:
            page_rows = self.get_cross_branch_customers(
                period,
                filters,
                minimum_branch_count=minimum_branch_count,
                scope_type=scope_type,
                office_code=office_code,
                office_filter_mode=office_filter_mode,
                limit=page_size,
                offset=offset,
                sort_by=sort_by,
                sort_order="desc" if sort_desc else "asc",
            )
            rows.extend(page_rows)
            if len(page_rows) < page_size:
                break
            offset += page_size
        return rows

    def count_cross_branch_customers(
        self,
        period: str,
        filters: CustomerFilters | None = None,
        *,
        minimum_branch_count: int = 2,
        scope_type: object = "cross_branch",
        office_code: str = "",
        office_filter_mode: str = "actual",
    ) -> int:
        period = str(period or "").strip()
        if not period:
            periods = self.distinct_periods()
            period = periods[-1] if periods else ""
        if not period:
            return 0
        if not self.has_office_detail_for_period(period):
            return 0
        sql, params = _cross_branch_candidate_sql(
            period,
            filters or CustomerFilters(),
            minimum_branch_count=minimum_branch_count,
            scope_type=scope_type,
            office_code=office_code,
            office_filter_mode=office_filter_mode,
        )
        with self._database() as database:
            return int(database.execute(f"SELECT COUNT(*) FROM ({sql}) q", params).fetchone()[0] or 0)

    def get_cross_branch_kpis(
        self,
        period: str,
        filters: CustomerFilters | None = None,
        *,
        minimum_branch_count: int = 2,
        scope_type: object = "cross_branch",
        office_code: str = "",
        office_filter_mode: str = "actual",
    ) -> dict[str, object]:
        period = str(period or "").strip()
        if not period:
            periods = self.distinct_periods()
            period = periods[-1] if periods else ""
        if not period:
            return {}
        if not self.has_office_detail_for_period(period):
            return _empty_cross_branch_kpis()
        sql, params = _cross_branch_candidate_sql(
            period,
            filters or CustomerFilters(),
            minimum_branch_count=minimum_branch_count,
            scope_type=scope_type,
            office_code=office_code,
            office_filter_mode=office_filter_mode,
        )
        with self._database() as database:
            database.execute("DROP TABLE IF EXISTS temp_cross_branch_candidates")
            database.execute(f"CREATE TEMP TABLE temp_cross_branch_candidates AS {sql}", params)
            database.execute(
                """
                CREATE INDEX IF NOT EXISTS temp_idx_cross_branch_candidates_sequence
                ON temp_cross_branch_candidates(period, customer_sequence)
                """
            )
            return _cross_branch_kpis_from_temp(database, period, self.unit_directory)

    def get_cross_branch_customer_detail(
        self,
        period: str,
        customer_sequence: str,
        branch_code: str | None = None,
        office_code: str | None = None,
        scope_type: str = "all",
    ) -> list[dict[str, object]]:
        return self.get_cross_branch_customer_offices(
            customer_sequence,
            period,
            branch_code=branch_code,
            office_code=office_code,
            scope_type=scope_type,
        )

    def get_cross_branch_customer_history(
        self,
        customer_sequence: str,
        period_from: str | None = None,
        period_to: str | None = None,
    ) -> list[dict[str, object]]:
        return self.get_cross_branch_customer_unit_history(
            customer_sequence,
            period_from or "",
            period_to or "",
        )

    def get_cross_branch_customer_offices(
        self,
        customer_sequence: str,
        report_period: str,
        branch_code: str | None = None,
        office_code: str | None = None,
        scope_type: str = "all",
    ) -> list[dict[str, object]]:
        sequence = str(customer_sequence or "").strip()
        period = str(report_period or "").strip()
        if not sequence or not period:
            return []
        all_rows = self._office_rows_for_sequence_period(sequence, period)
        if not all_rows or not _scope_matches_office_rows(all_rows, scope_type):
            return []
        branch_filter = str(branch_code or "").strip()
        office_filter = str(office_code or "").strip()
        rows = [
            row for row in all_rows
            if (not branch_filter or str(row.get("branch_code") or "") == branch_filter)
            and (not office_filter or str(row.get("office_code") or "") == office_filter)
        ]
        return [_finalize_cross_branch_detail_row(dict(row), self.unit_directory) for row in rows]

    def get_cross_branch_customer_filtered_kpis(
        self,
        customer_sequence: str,
        report_period: str,
        branch_code: str | None = None,
        office_code: str | None = None,
        scope_type: str = "all",
    ) -> dict[str, object]:
        sequence = str(customer_sequence or "").strip()
        period = str(report_period or "").strip()
        if not sequence or not period:
            return {}
        all_rows = self._office_rows_for_sequence_period(sequence, period)
        summary_rows = self._summary_rows_for_sequence_period(sequence, period)
        if not all_rows:
            return _missing_office_detail_kpis(period, sequence, summary_rows, self.unit_directory)
        if not _scope_matches_office_rows(all_rows, scope_type):
            return _empty_filtered_kpis(period, sequence, summary_rows, missing=False)
        branch_filter = str(branch_code or "").strip()
        office_filter = str(office_code or "").strip()
        rows = [
            row for row in all_rows
            if (not branch_filter or str(row.get("branch_code") or "") == branch_filter)
            and (not office_filter or str(row.get("office_code") or "") == office_filter)
        ]
        return _office_kpis(period, sequence, rows, summary_rows, missing=False)

    def get_cross_branch_customer_unit_history(
        self,
        customer_sequence: str,
        period_from: str = "",
        period_to: str = "",
        branch_code: str | None = None,
        office_code: str | None = None,
        scope_type: str = "all",
    ) -> list[dict[str, object]]:
        sequence = str(customer_sequence or "").strip()
        if not sequence:
            return []
        clauses = ["customer_sequence = ?", "total_balance > 0"]
        params: list[object] = [sequence]
        if period_from:
            clauses.append("period >= ?")
            params.append(str(period_from).strip())
        if period_to:
            clauses.append("period <= ?")
            params.append(str(period_to).strip())
        where = "WHERE " + " AND ".join(clauses)
        with self._database() as database:
            period_rows = database.execute(
                f"""
                SELECT DISTINCT period
                FROM customer_period_summary
                {where}
                ORDER BY period ASC
                """,
                params,
            ).fetchall()
        output: list[dict[str, object]] = []
        previous_total: float | None = None
        for period_row in period_rows:
            period = str(period_row["period"] or "")
            all_rows = self._office_rows_for_sequence_period(sequence, period)
            summary_rows = self._summary_rows_for_sequence_period(sequence, period)
            if all_rows and not _scope_matches_office_rows(all_rows, scope_type):
                continue
            branch_filter = str(branch_code or "").strip()
            office_filter = str(office_code or "").strip()
            if all_rows:
                rows = [
                    row for row in all_rows
                    if (not branch_filter or str(row.get("branch_code") or "") == branch_filter)
                    and (not office_filter or str(row.get("office_code") or "") == office_filter)
                ]
                item = _office_history_row(period, sequence, rows, summary_rows, missing=False, unit_directory=self.unit_directory)
            elif office_filter:
                item = _office_history_row(period, sequence, [], summary_rows, missing=True, unit_directory=self.unit_directory)
            elif branch_filter:
                filtered_summary = [row for row in summary_rows if str(row.get("branch_code") or "") == branch_filter]
                item = _summary_history_row(period, sequence, filtered_summary, missing=True, unit_directory=self.unit_directory)
            else:
                item = _summary_history_row(period, sequence, summary_rows, missing=True, unit_directory=self.unit_directory)
            total = float(item.get("total_balance") or 0)
            item["difference"] = "" if previous_total is None else total - previous_total
            previous_total = total
            output.append(item)
        return output

    def get_customer_available_branches(
        self,
        customer_sequence: str,
        report_period: str | None = None,
    ) -> list[dict[str, object]]:
        sequence = str(customer_sequence or "").strip()
        period = str(report_period or "").strip()
        if not sequence:
            return []
        clauses = ["customer_sequence = ?", "total_balance > 0"]
        params: list[object] = [sequence]
        if period:
            clauses.append("period = ?")
            params.append(period)
        where = "WHERE " + " AND ".join(clauses)
        with self._database() as database:
            rows = database.execute(
                f"""
                SELECT DISTINCT branch_code
                FROM customer_period_summary
                {where}
                ORDER BY branch_code
                """,
                params,
            ).fetchall()
        return [
            {"branch_code": str(row["branch_code"] or ""), "branch_name": self._branch_display(row["branch_code"])}
            for row in rows
            if str(row["branch_code"] or "")
        ]

    def get_customer_available_offices(
        self,
        customer_sequence: str,
        report_period: str,
        branch_code: str | None = None,
    ) -> list[dict[str, object]]:
        sequence = str(customer_sequence or "").strip()
        period = str(report_period or "").strip()
        if not sequence or not period:
            return []
        clauses = ["customer_sequence = ?", "period = ?", "total_balance > 0"]
        params: list[object] = [sequence, period]
        branch = str(branch_code or "").strip()
        if branch:
            clauses.append("branch_code = ?")
            params.append(branch)
        where = "WHERE " + " AND ".join(clauses)
        with self._database() as database:
            rows = database.execute(
                f"""
                SELECT DISTINCT branch_code, trctcd, office_code, office_name, office_type
                FROM customer_office_period
                {where}
                ORDER BY branch_code,
                    CASE
                        WHEN office_type = 'HEAD_OFFICE' THEN 0
                        WHEN office_type = 'TRANSACTION_OFFICE' THEN 1
                        ELSE 2
                    END,
                    trctcd COLLATE NOCASE,
                    office_code COLLATE NOCASE
                """,
                params,
            ).fetchall()
        return [_finalize_office_option(dict(row), self.unit_directory) for row in rows]

    def customer_sequence_periods(self, customer_sequence: str) -> list[str]:
        sequence = str(customer_sequence or "").strip()
        if not sequence:
            return []
        with self._database() as database:
            rows = database.execute(
                """
                SELECT DISTINCT period
                FROM customer_period_summary
                WHERE customer_sequence = ?
                  AND total_balance > 0
                ORDER BY period
                """,
                (sequence,),
            ).fetchall()
        return [str(row["period"] or "") for row in rows]

    def has_office_detail_for_customer_period(self, customer_sequence: str, period: str) -> bool:
        sequence = str(customer_sequence or "").strip()
        clean_period = str(period or "").strip()
        if not sequence or not clean_period:
            return False
        with self._database() as database:
            return bool(
                database.execute(
                    """
                    SELECT 1
                    FROM customer_office_period
                    WHERE customer_sequence = ?
                      AND period = ?
                      AND total_balance > 0
                    LIMIT 1
                    """,
                    (sequence, clean_period),
                ).fetchone()
            )

    def get_multi_unit_customers(
        self,
        period: str,
        scope_type: str,
        branch_code: str | None = None,
        office_code: str | None = None,
        filters: CustomerFilters | None = None,
        *,
        office_filter_mode: str = "actual",
        minimum_branch_count: int = 2,
        limit: int = 100,
        offset: int = 0,
        sort_by: str = "branch_count",
        sort_order: str = "desc",
    ) -> list[dict[str, object]]:
        base_filters = filters or CustomerFilters()
        if branch_code:
            base_filters = replace(base_filters, branch_code=str(branch_code).strip())
        return self.get_cross_branch_customers(
            period,
            base_filters,
            minimum_branch_count=minimum_branch_count,
            scope_type=scope_type,
            office_code=str(office_code or "").strip(),
            office_filter_mode=office_filter_mode,
            limit=limit,
            offset=offset,
            sort_by=sort_by,
            sort_order=sort_order,
        )

    def _office_rows_for_sequence_period(self, customer_sequence: str, period: str) -> list[dict[str, object]]:
        sequence = str(customer_sequence or "").strip()
        clean_period = str(period or "").strip()
        if not sequence or not clean_period:
            return []
        code_expr = _override_value_sql("s", "officer_code", fallback="o.primary_officer_code")
        name_expr = _override_value_sql("s", "officer_name", fallback="o.primary_officer_name", null_if_empty=True)
        has_override = _has_override_sql("s")
        with self._database() as database:
            rows = database.execute(
                f"""
                SELECT
                    o.period,
                    o.customer_sequence,
                    o.customer_code,
                    o.branch_code,
                    o.trctcd,
                    o.office_code,
                    o.office_name,
                    o.office_type,
                    s.customer_name,
                    s.customer_type,
                    o.primary_officer_code AS imported_officer_code,
                    o.primary_officer_name AS imported_officer_name,
                    {code_expr} AS effective_officer_code,
                    {name_expr} AS effective_officer_name,
                    o.officer_count,
                    o.total_balance,
                    o.short_term_balance,
                    o.medium_long_term_balance,
                    o.other_balance,
                    CASE WHEN o.total_balance <> 0
                        THEN o.medium_long_term_balance / o.total_balance * 100
                        ELSE 0
                    END AS medium_long_ratio,
                    CASE WHEN o.total_balance <> 0
                        THEN o.interest_rate_numerator / o.total_balance
                        ELSE 0
                    END AS average_rate,
                    CASE WHEN o.total_balance <> 0
                        THEN o.nim_before_numerator / o.total_balance
                        ELSE 0
                    END AS nim_before,
                    CASE WHEN o.total_balance <> 0
                        THEN o.nim_after_numerator / o.total_balance
                        ELSE 0
                    END AS nim_after,
                    o.interest_rate_numerator,
                    o.nim_before_numerator,
                    o.nim_after_numerator,
                    o.source_loan_count,
                    CASE WHEN {has_override} THEN 1 ELSE 0 END AS has_override
                FROM customer_office_period o
                JOIN customer_period_summary s
                    ON s.period = o.period
                   AND s.customer_code = o.customer_code
                WHERE o.period = ?
                  AND o.customer_sequence = ?
                  AND o.total_balance > 0
                ORDER BY o.branch_code ASC,
                    CASE
                        WHEN o.office_type = 'HEAD_OFFICE' THEN 0
                        WHEN o.office_type = 'TRANSACTION_OFFICE' THEN 1
                        ELSE 2
                    END,
                    o.trctcd COLLATE NOCASE ASC,
                    o.office_code COLLATE NOCASE ASC
                """,
                (clean_period, sequence),
            ).fetchall()
        return [dict(row) for row in rows]

    def _summary_rows_for_sequence_period(self, customer_sequence: str, period: str) -> list[dict[str, object]]:
        sequence = str(customer_sequence or "").strip()
        clean_period = str(period or "").strip()
        if not sequence or not clean_period:
            return []
        with self._database() as database:
            rows = database.execute(
                """
                SELECT *
                FROM customer_period_summary
                WHERE period = ?
                  AND customer_sequence = ?
                  AND total_balance > 0
                ORDER BY branch_code ASC, customer_code COLLATE NOCASE ASC
                """,
                (clean_period, sequence),
            ).fetchall()
        return [dict(row) for row in rows]

    def _cross_branch_detail_rows_for_keys(
        self,
        keys: list[tuple[str, str]],
    ) -> dict[tuple[str, str], list[dict[str, object]]]:
        with self._database() as database:
            return CustomerRepository._cross_branch_detail_rows_for_keys_from_database(database, keys)

    @staticmethod
    def _cross_branch_detail_rows_for_keys_from_database(
        database: sqlite3.Connection,
        keys: list[tuple[str, str]],
    ) -> dict[tuple[str, str], list[dict[str, object]]]:
        clean_by_period: dict[str, set[str]] = {}
        for period, sequence in keys:
            clean_period = str(period or "").strip()
            clean_sequence = str(sequence or "").strip()
            if clean_period and clean_sequence:
                clean_by_period.setdefault(clean_period, set()).add(clean_sequence)
        if not clean_by_period:
            return {}
        code_expr = _override_value_sql("s", "officer_code", fallback="o.primary_officer_code")
        name_expr = _override_value_sql("s", "officer_name", fallback="o.primary_officer_name", null_if_empty=True)
        has_override = _has_override_sql("s")
        output: dict[tuple[str, str], list[dict[str, object]]] = {}
        for period, sequences in clean_by_period.items():
            ordered_sequences = sorted(sequences)
            placeholders = ", ".join("?" for _sequence in ordered_sequences)
            rows = database.execute(
                f"""
                SELECT
                    o.period,
                    o.customer_sequence,
                    o.customer_code,
                    o.branch_code,
                    o.trctcd,
                    o.office_code,
                    o.office_name,
                    o.office_type,
                    s.customer_name,
                    s.customer_type,
                    o.primary_officer_code AS imported_officer_code,
                    o.primary_officer_name AS imported_officer_name,
                    {code_expr} AS effective_officer_code,
                    {name_expr} AS effective_officer_name,
                    o.officer_count,
                    o.total_balance,
                    o.short_term_balance,
                    o.medium_long_term_balance,
                    o.other_balance,
                    CASE WHEN o.total_balance <> 0
                        THEN o.medium_long_term_balance / o.total_balance * 100
                        ELSE 0
                    END AS medium_long_ratio,
                    CASE WHEN o.total_balance <> 0
                        THEN o.interest_rate_numerator / o.total_balance
                        ELSE 0
                    END AS average_rate,
                    CASE WHEN o.total_balance <> 0
                        THEN o.nim_before_numerator / o.total_balance
                        ELSE 0
                    END AS nim_before,
                    CASE WHEN o.total_balance <> 0
                        THEN o.nim_after_numerator / o.total_balance
                        ELSE 0
                    END AS nim_after,
                    o.interest_rate_numerator,
                    o.nim_before_numerator,
                    o.nim_after_numerator,
                    o.source_loan_count,
                    CASE WHEN {has_override} THEN 1 ELSE 0 END AS has_override
                FROM customer_office_period o
                JOIN customer_period_summary s
                    ON s.period = o.period
                   AND s.customer_code = o.customer_code
                WHERE o.period = ?
                  AND o.customer_sequence IN ({placeholders})
                  AND o.total_balance > 0
                ORDER BY o.customer_sequence COLLATE NOCASE ASC,
                    o.branch_code ASC,
                    CASE
                        WHEN o.office_type = 'HEAD_OFFICE' THEN 0
                        WHEN o.office_type = 'TRANSACTION_OFFICE' THEN 1
                        ELSE 2
                    END,
                    o.trctcd COLLATE NOCASE ASC,
                    o.office_code COLLATE NOCASE ASC
                """,
                (period, *ordered_sequences),
            ).fetchall()
            for row in rows:
                item = dict(row)
                key = (str(item.get("period") or ""), str(item.get("customer_sequence") or ""))
                output.setdefault(key, []).append(item)
        return output

    def explain_cross_branch_query_plan(
        self,
        period: str,
        filters: CustomerFilters | None = None,
        *,
        minimum_branch_count: int = 2,
        scope_type: str = "cross_branch",
        office_code: str = "",
        office_filter_mode: str = "actual",
    ) -> list[str]:
        sql, params = _cross_branch_candidate_sql(
            period,
            filters or CustomerFilters(),
            minimum_branch_count=minimum_branch_count,
            scope_type=scope_type,
            office_code=office_code,
            office_filter_mode=office_filter_mode,
        )
        with self._database() as database:
            rows = database.execute(f"EXPLAIN QUERY PLAN {sql}", params).fetchall()
        return [str(row["detail"] if "detail" in row.keys() else row[-1]) for row in rows]

    def _finalize_cross_branch_summary(
        self,
        row: dict[str, object],
        *,
        detail_rows: list[dict[str, object]] | None = None,
    ) -> dict[str, object]:
        if detail_rows is None:
            detail_rows = self.get_cross_branch_customer_detail(
                str(row.get("period") or ""),
                str(row.get("customer_sequence") or ""),
            )
        else:
            detail_rows = [
                _finalize_cross_branch_detail_row(dict(item), self.unit_directory)
                for item in detail_rows
            ]
        if not detail_rows:
            return row
        branch_codes = sorted({str(item.get("branch_code") or "") for item in detail_rows if str(item.get("branch_code") or "")})
        names = [str(item.get("customer_name") or "").strip() for item in detail_rows]
        normalized_names = {_normalize_customer_name(name) for name in names if _normalize_customer_name(name)}
        types = {str(item.get("customer_type") or "").strip().upper() or "OTHER" for item in detail_rows}
        officer_identities = {
            _officer_identity(item.get("effective_officer_code"), item.get("effective_officer_name"))
            for item in detail_rows
            if _officer_identity(item.get("effective_officer_code"), item.get("effective_officer_name"))
        }
        row["customer_name"] = next((name for name in names if name), str(row.get("customer_name") or ""))
        row["branch_count"] = len(branch_codes)
        status = _office_status(detail_rows)
        row["office_count"] = status["office_count"]
        row["head_office_count"] = status["head_office_count"]
        row["pgd_count"] = status["pgd_count"]
        row["has_head_and_pgd"] = status["has_head_and_pgd"]
        row["has_multi_pgd"] = status["has_multi_pgd"]
        row["head_office_balance"] = status["head_office_balance"]
        row["pgd_balance"] = status["pgd_balance"]
        branch_balances: dict[str, float] = {}
        for item in detail_rows:
            code = str(item.get("branch_code") or "")
            if code:
                branch_balances[code] = branch_balances.get(code, 0.0) + _number(item.get("total_balance"))
        row["branch_list"] = "\n".join(
            f"{self._branch_display(code)}: {format_money_vn(branch_balances.get(code, 0))}"
            for code in branch_codes
        )
        row["office_list"] = "\n".join(
            f"{_dynamic_office_display(item, self.unit_directory)}: {format_money_vn(item.get('total_balance'))}"
            for item in detail_rows
            if _dynamic_office_display(item, self.unit_directory)
        )
        representatives = [
            resolve_representative_office(row.get("period"), row.get("customer_sequence"), branch, detail_rows)
            for branch in branch_codes
        ]
        row["representative_office_list"] = "\n".join(
            f"{self._branch_display(branch)}: "
            f"{rep.representative_office_code}"
            f"{' - ' + rep.representative_office_name if rep.representative_office_name else ''}"
            for branch, rep in zip(branch_codes, representatives)
        )
        row["representative_office_type_list"] = "\n".join(
            OFFICE_TYPE_LABELS.get(rep.representative_office_type, rep.representative_office_type)
            for rep in representatives
        )
        row["representative_reason_list"] = "\n".join(
            REPRESENTATIVE_REASON_LABELS.get(rep.reason, rep.reason)
            for rep in representatives
        )
        scope_status = []
        if int(row.get("branch_count") or 0) >= 2:
            scope_status.append("Vay tại nhiều chi nhánh")
        if int(row.get("has_head_and_pgd") or 0):
            scope_status.append("Hội sở và PGD cùng chi nhánh")
        if int(row.get("has_multi_pgd") or 0):
            scope_status.append("Nhiều PGD cùng chi nhánh")
        row["scope_status"] = "; ".join(scope_status) if scope_status else "Một đơn vị"
        row["has_head_and_pgd_text"] = "Có" if int(row.get("has_head_and_pgd") or 0) else "Không"
        row["has_multi_pgd_text"] = "Có" if int(row.get("has_multi_pgd") or 0) else "Không"
        row["has_override"] = 1 if any(int(item.get("has_override") or 0) for item in detail_rows) else 0
        row["officer_count"] = len(officer_identities)
        row["officer_list"] = "\n".join(
            sorted(
                {
                    str(item.get("effective_officer_name") or item.get("effective_officer_code") or "").strip()
                    for item in detail_rows
                    if str(item.get("effective_officer_name") or item.get("effective_officer_code") or "").strip()
                },
                key=str.casefold,
            )
        )
        row["name_conflict"] = 1 if len(normalized_names) > 1 else 0
        row["customer_type_conflict"] = 1 if len(types) > 1 else 0
        if len(types) == 1:
            code = next(iter(types))
            row["customer_type"] = code
            row["customer_type_display"] = CUSTOMER_TYPE_LABELS.get(code, CUSTOMER_TYPE_LABELS["OTHER"])
        else:
            row["customer_type"] = "MIXED"
            row["customer_type_display"] = "Không thống nhất"
        conflicts = []
        if row["name_conflict"]:
            conflicts.append("Có xung đột tên")
        if row["customer_type_conflict"]:
            conflicts.append("Có xung đột loại KH")
        row["conflict_status"] = "; ".join(conflicts) if conflicts else "Không"
        row["name_detail"] = "\n".join(
            f"{self._branch_display(item.get('branch_code'))}: "
            f"{item.get('customer_name') or ''}"
            for item in detail_rows
        )
        return row

    def dashboard_metrics(self, filters: CustomerFilters | None = None) -> dict[str, object]:
        filters = filters or CustomerFilters()
        return self.get_dashboard_kpis(filters, filters.current_period or filters.period_to)

    def get_dashboard_kpis(
        self,
        filters: CustomerFilters | None = None,
        report_period: str = "",
    ) -> dict[str, object]:
        filters = filters or CustomerFilters()
        period = str(report_period or filters.current_period or filters.period_to or "").strip()
        if not period:
            periods = self.distinct_periods()
            period = periods[-1] if periods else ""
        if period:
            filters = replace(filters, current_period=period, period_from="", period_to="")
        where, params = _summary_where(filters)
        has_override_expr = _has_override_sql("s")
        with self._database() as database:
            row = database.execute(
                f"""
                SELECT
                    COUNT(DISTINCT CASE WHEN s.total_balance > 0 THEN s.customer_code END) AS customer_count,
                    COUNT(DISTINCT CASE WHEN s.total_balance > 0 THEN s.customer_code END) AS active_customer_count,
                    COALESCE(SUM(s.total_balance), 0) AS total_balance,
                    COALESCE(SUM(s.short_term_balance), 0) AS short_term_balance,
                    COALESCE(SUM(s.medium_long_term_balance), 0) AS medium_long_term_balance,
                    COALESCE(SUM(s.other_balance), 0) AS other_balance,
                    CASE WHEN SUM(s.total_balance) <> 0
                        THEN SUM(s.medium_long_term_balance) / SUM(s.total_balance) * 100
                        ELSE 0
                    END AS medium_long_ratio,
                    CASE WHEN SUM(s.total_balance) <> 0
                        THEN SUM(s.interest_rate_numerator) / SUM(s.total_balance)
                        ELSE 0
                    END AS average_rate,
                    CASE WHEN SUM(s.total_balance) <> 0
                        THEN SUM(s.nim_before_numerator) / SUM(s.total_balance)
                        ELSE 0
                    END AS nim_before,
                    CASE WHEN SUM(s.total_balance) <> 0
                        THEN SUM(s.nim_after_numerator) / SUM(s.total_balance)
                        ELSE 0
                    END AS nim_after,
                    COALESCE(SUM(CASE WHEN s.has_multiple_officers = 1 THEN 1 ELSE 0 END), 0) AS multiple_officer_customer_count,
                    COALESCE(SUM(CASE WHEN {has_override_expr} THEN 1 ELSE 0 END), 0) AS override_customer_count
                FROM customer_period_summary s
                {where}
                """,
                params,
            ).fetchone()
        return dict(row) if row is not None else {}

    def dashboard_trends(self, filters: CustomerFilters | None = None) -> list[dict[str, object]]:
        filters = (filters or CustomerFilters()).without_exact_period()
        where, params = _summary_where(filters)
        with self._database() as database:
            rows = database.execute(
                f"""
                SELECT
                    s.period,
                    SUM(s.total_balance) AS total_balance,
                    SUM(s.short_term_balance) AS short_term_balance,
                    SUM(s.medium_long_term_balance) AS medium_long_term_balance,
                    SUM(s.other_balance) AS other_balance,
                    CASE WHEN SUM(s.total_balance) <> 0
                        THEN SUM(s.interest_rate_numerator) / SUM(s.total_balance)
                        ELSE 0
                    END AS average_rate,
                    CASE WHEN SUM(s.total_balance) <> 0
                        THEN SUM(s.nim_before_numerator) / SUM(s.total_balance)
                        ELSE 0
                    END AS nim_before,
                    CASE WHEN SUM(s.total_balance) <> 0
                        THEN SUM(s.nim_after_numerator) / SUM(s.total_balance)
                        ELSE 0
                    END AS nim_after
                FROM customer_period_summary s
                {where}
                GROUP BY s.period
                ORDER BY s.period
                """,
                params,
            ).fetchall()
        return [dict(row) for row in rows]

    def get_total_balance_trend(
        self,
        filters: CustomerFilters | None,
        period_from: str = "",
        period_to: str = "",
        *,
        group_by: str = "total",
    ) -> list[dict[str, object]]:
        filters = _trend_filters(filters or CustomerFilters(), period_from, period_to)
        where, params = _summary_where(filters)
        group_key = str(group_by or "total").casefold()
        if group_key == "branch":
            sql = f"""
                SELECT
                    s.period,
                    s.branch_code AS series_key,
                    SUM(s.total_balance) AS value
                FROM customer_period_summary s
                {where}
                GROUP BY s.period, s.branch_code
                ORDER BY s.period ASC, s.branch_code ASC
            """
        elif group_key == "customer_type":
            sql = f"""
                SELECT
                    s.period,
                    CASE
                        WHEN s.customer_type = 'CN' THEN 'CN'
                        WHEN s.customer_type = 'TC' THEN 'TC'
                        ELSE 'OTHER'
                    END AS series_key,
                    SUM(s.total_balance) AS value
                FROM customer_period_summary s
                {where}
                GROUP BY s.period, series_key
                ORDER BY s.period ASC,
                    CASE series_key WHEN 'CN' THEN 1 WHEN 'TC' THEN 2 ELSE 3 END
            """
        else:
            sql = f"""
                SELECT
                    s.period,
                    'total' AS series_key,
                    SUM(s.total_balance) AS value
                FROM customer_period_summary s
                {where}
                GROUP BY s.period
                ORDER BY s.period ASC
            """
        with self._database() as database:
            rows = database.execute(sql, params).fetchall()
        return [_trend_balance_row(dict(row), group_key, self.unit_directory) for row in rows]

    def get_customer_metric_trend(
        self,
        filters: CustomerFilters | None,
        period_from: str = "",
        period_to: str = "",
        *,
        metric: str = "average_rate",
    ) -> list[dict[str, object]]:
        metric_key = str(metric or "average_rate")
        numerator_column, metric_label = CUSTOMER_METRIC_SQL.get(metric_key, CUSTOMER_METRIC_SQL["average_rate"])
        filters = _trend_filters(filters or CustomerFilters(), period_from, period_to)
        where, params = _summary_where(filters)
        with self._database() as database:
            rows = database.execute(
                f"""
                SELECT
                    s.period,
                    ? AS series_key,
                    ? AS series_name,
                    CASE WHEN SUM(s.total_balance) <> 0
                        THEN SUM(s.{numerator_column}) / SUM(s.total_balance)
                        ELSE 0
                    END AS value
                FROM customer_period_summary s
                {where}
                GROUP BY s.period
                ORDER BY s.period ASC
                """,
                (metric_key, metric_label, *params),
            ).fetchall()
        return [dict(row) for row in rows]

    def get_active_customer_count_trend(
        self,
        filters: CustomerFilters | None,
        period_from: str = "",
        period_to: str = "",
    ) -> list[dict[str, object]]:
        filters = _trend_filters(filters or CustomerFilters(), period_from, period_to)
        where, params = _summary_where(filters)
        with self._database() as database:
            rows = database.execute(
                f"""
                SELECT
                    s.period,
                    COUNT(DISTINCT s.customer_code) AS active_customer_count
                FROM customer_period_summary s
                {where} AND s.total_balance > 0
                GROUP BY s.period
                ORDER BY s.period ASC
                """,
                params,
            ).fetchall()
        return [dict(row) for row in rows]

    def top_customers(
        self,
        filters: CustomerFilters | None = None,
        *,
        metric: str = "total_balance",
        limit: int = 10,
        descending: bool = True,
    ) -> list[dict[str, object]]:
        filters = filters or CustomerFilters()
        if metric not in {"total_balance", "short_term_balance", "medium_long_term_balance", "other_balance"}:
            metric = "total_balance"
        result = self.query_customer_list(
            filters,
            page=1,
            page_size=max(1, min(100, int(limit or 10))),
            sort_by=metric,
            sort_desc=descending,
        )
        return result.rows

    def get_top_customers_by_balance(
        self,
        filters: CustomerFilters | None,
        period: str,
        limit: int,
    ) -> list[dict[str, object]]:
        filters = filters or CustomerFilters()
        period = str(period or filters.current_period or "").strip()
        if period:
            filters = filters.with_current_period(period)
        result = self.query_customer_list(
            filters,
            page=1,
            page_size=max(1, min(50, int(limit or 10))),
            sort_by="total_balance",
            sort_desc=True,
        )
        return result.rows

    def get_top_customer_movements(
        self,
        filters: CustomerFilters | None,
        previous_period: str,
        current_period: str,
        *,
        direction: str = "increase",
        limit: int = 10,
    ) -> list[dict[str, object]]:
        filters = filters or CustomerFilters()
        if not previous_period or not current_period:
            return []
        direction = str(direction or "increase").casefold()
        if direction == "decrease":
            status = MOVEMENT_STATUS_DECREASE
            sort_desc = False
        else:
            status = MOVEMENT_STATUS_INCREASE
            sort_desc = True
        return self.movement_rows(
            previous_period,
            current_period,
            replace(filters, movement_status=status),
            page=1,
            page_size=max(1, min(50, int(limit or 10))),
            sort_by="difference",
            sort_desc=sort_desc,
        ).rows

    def movement_rows(
        self,
        previous_period: str,
        current_period: str,
        filters: CustomerFilters | None = None,
        *,
        page: int = 1,
        page_size: int = 100,
        sort_by: str = "difference",
        sort_desc: bool = True,
    ) -> PageResult:
        filters = filters or CustomerFilters()
        page = max(1, int(page or 1))
        page_size = max(1, min(5000, int(page_size or 100)))
        previous_period = str(previous_period or "").strip()
        current_period = str(current_period or "").strip()
        if not previous_period or not current_period:
            return PageResult(rows=[], total_rows=0, page=page, page_size=page_size)
        stats: list[dict[str, object]] = []
        with self._database() as database:
            database.execute("PRAGMA busy_timeout = 10000")
            database.execute("PRAGMA temp_store = MEMORY")
            _materialize_movement_candidates(
                database,
                previous_period,
                current_period,
                filters,
                stats,
                sort_by=sort_by,
            )
            total_rows = _movement_count_from_temp(database, filters, stats, previous_period, current_period)
            rows = _movement_page_from_temp(
                database,
                filters,
                page=page,
                page_size=page_size,
                sort_by=sort_by,
                sort_desc=sort_desc,
                stats=stats,
                previous_period=previous_period,
                current_period=current_period,
            )
            _resolve_movement_officers_for_rows(database, rows, stats, previous_period, current_period)
        _enrich_movement_unit_rows(rows, self.unit_directory)
        return PageResult(rows=rows, total_rows=total_rows, page=page, page_size=page_size)

    def movement_kpis(
        self,
        previous_period: str,
        current_period: str,
        filters: CustomerFilters | None = None,
    ) -> dict[str, object]:
        filters = filters or CustomerFilters()
        previous_period = str(previous_period or "").strip()
        current_period = str(current_period or "").strip()
        if not previous_period or not current_period:
            return _empty_movement_kpis()
        stats: list[dict[str, object]] = []
        with self._database() as database:
            database.execute("PRAGMA busy_timeout = 10000")
            database.execute("PRAGMA temp_store = MEMORY")
            _materialize_movement_candidates(database, previous_period, current_period, filters, stats)
            return _movement_kpis_from_temp(database, filters, stats, previous_period, current_period)

    def movement_payload(
        self,
        previous_period: str,
        current_period: str,
        filters: CustomerFilters | None = None,
        *,
        page: int = 1,
        page_size: int = 100,
        sort_by: str = "difference",
        sort_desc: bool = True,
        generation: int | None = None,
    ) -> dict[str, object]:
        filters = filters or CustomerFilters()
        page = max(1, int(page or 1))
        page_size = max(1, min(5000, int(page_size or 100)))
        previous_period = str(previous_period or "").strip()
        current_period = str(current_period or "").strip()
        stats: list[dict[str, object]] = []
        total_started = perf_counter()
        validate_started = perf_counter()
        if not previous_period or not current_period:
            _record_movement_stage(
                stats,
                "validate_periods",
                validate_started,
                previous_period,
                current_period,
                generation=generation,
                rows=0,
            )
            result = PageResult(rows=[], total_rows=0, page=page, page_size=page_size)
            _record_movement_stage(
                stats,
                "total_load",
                total_started,
                previous_period,
                current_period,
                generation=generation,
                rows=0,
            )
            return {"kpis": _empty_movement_kpis(), "result": result, "stage_stats": tuple(stats)}
        _record_movement_stage(
            stats,
            "validate_periods",
            validate_started,
            previous_period,
            current_period,
            generation=generation,
        )
        with self._database() as database:
            database.execute("PRAGMA busy_timeout = 10000")
            database.execute("PRAGMA temp_store = MEMORY")
            _materialize_movement_candidates(
                database,
                previous_period,
                current_period,
                filters,
                stats,
                generation=generation,
                sort_by=sort_by,
            )
            kpis = _movement_kpis_from_temp(
                database,
                filters,
                stats,
                previous_period,
                current_period,
                generation=generation,
            )
            total_rows = _movement_count_from_temp(
                database,
                filters,
                stats,
                previous_period,
                current_period,
                generation=generation,
            )
            rows = _movement_page_from_temp(
                database,
                filters,
                page=page,
                page_size=page_size,
                sort_by=sort_by,
                sort_desc=sort_desc,
                stats=stats,
                previous_period=previous_period,
                current_period=current_period,
                generation=generation,
            )
            _resolve_movement_officers_for_rows(
                database,
                rows,
                stats,
                previous_period,
                current_period,
                generation=generation,
            )
        _enrich_movement_unit_rows(
            rows,
            self.unit_directory,
            stats,
            previous_period,
            current_period,
            generation=generation,
        )
        model_started = perf_counter()
        result = PageResult(rows=rows, total_rows=total_rows, page=page, page_size=page_size)
        _record_movement_stage(
            stats,
            "model_rows",
            model_started,
            previous_period,
            current_period,
            generation=generation,
            rows=len(rows),
        )
        _record_movement_stage(
            stats,
            "total_load",
            total_started,
            previous_period,
            current_period,
            generation=generation,
            rows=total_rows,
        )
        return {
            "kpis": kpis,
            "result": result,
            "stage_stats": tuple(stats),
        }

    def top_movement_customers(
        self,
        previous_period: str,
        current_period: str,
        filters: CustomerFilters | None = None,
        *,
        movement_status: str = "",
        limit: int = 10,
    ) -> list[dict[str, object]]:
        filters = filters or CustomerFilters()
        if movement_status:
            filters = CustomerFilters(
                period_from=filters.period_from,
                period_to=filters.period_to,
                current_period=filters.current_period,
                compare_period=filters.compare_period,
                branch_code=filters.branch_code,
                customer_type=filters.customer_type,
                officer=filters.officer,
                loan_term=filters.loan_term,
                search_text=filters.search_text,
                movement_status=movement_status,
                multi_status=filters.multi_status,
                override_status=filters.override_status,
                debt_group=filters.debt_group,
            )
        sort_desc = movement_status not in {MOVEMENT_STATUS_DECREASE, MOVEMENT_STATUS_PAID_OFF}
        result = self.movement_rows(
            previous_period,
            current_period,
            filters,
            page=1,
            page_size=max(1, min(100, int(limit or 10))),
            sort_by="difference",
            sort_desc=sort_desc,
        )
        return result.rows

    def all_movement_rows(
        self,
        previous_period: str,
        current_period: str,
        filters: CustomerFilters | None = None,
        *,
        sort_by: str = "difference",
        sort_desc: bool = True,
    ) -> list[dict[str, object]]:
        filters = filters or CustomerFilters()
        previous_period = str(previous_period or "").strip()
        current_period = str(current_period or "").strip()
        if not previous_period or not current_period:
            return []
        stats: list[dict[str, object]] = []
        with self._database() as database:
            database.execute("PRAGMA busy_timeout = 10000")
            database.execute("PRAGMA temp_store = MEMORY")
            _materialize_movement_candidates(
                database,
                previous_period,
                current_period,
                filters,
                stats,
                sort_by=sort_by,
            )
            where, params = _movement_where(filters)
            order_sql = _movement_order_sql(sort_by, sort_desc)
            rows = [
                dict(row)
                for row in database.execute(
                    f"SELECT * FROM {CUSTOMER_MOVEMENT_TEMP_TABLE} q {where} {order_sql}",
                    params,
                ).fetchall()
            ]
            _resolve_movement_officers_for_rows(database, rows, stats, previous_period, current_period)
        _enrich_movement_unit_rows(rows, self.unit_directory)
        return rows

    def multiple_officer_rows(
        self,
        filters: CustomerFilters | None = None,
        *,
        page: int = 1,
        page_size: int = 100,
        sort_by: str = "total_balance",
        sort_desc: bool = True,
    ) -> PageResult:
        filters = filters or CustomerFilters(multi_status="same_period")
        if not filters.multi_status:
            filters = CustomerFilters(
                period_from=filters.period_from,
                period_to=filters.period_to,
                current_period=filters.current_period,
                compare_period=filters.compare_period,
                branch_code=filters.branch_code,
                customer_type=filters.customer_type,
                officer=filters.officer,
                loan_term=filters.loan_term,
                search_text=filters.search_text,
                movement_status=filters.movement_status,
                multi_status="same_period",
                override_status=filters.override_status,
                debt_group=filters.debt_group,
            )
        page_result = self.query_customer_list(
            filters,
            page=page,
            page_size=page_size,
            sort_by=sort_by,
            sort_desc=sort_desc,
        )
        if not page_result.rows:
            return page_result
        keys = [(str(row["period"]), str(row["customer_code"])) for row in page_result.rows]
        officer_map = self._officer_list_for_customer_periods(keys)
        rows = []
        for row in page_result.rows:
            output = dict(row)
            output["officer_list"] = officer_map.get((str(row["period"]), str(row["customer_code"])), "")
            rows.append(output)
        return PageResult(rows=rows, total_rows=page_result.total_rows, page=page_result.page, page_size=page_result.page_size)

    def customer_detail(self, customer_code: str, period: str = "") -> dict[str, object] | None:
        filters = CustomerFilters(current_period=period, search_text=customer_code)
        rows = [
            row
            for row in self.query_customer_list(filters, page=1, page_size=10, sort_by="period", sort_desc=True).rows
            if str(row.get("customer_code") or "") == str(customer_code or "").strip()
        ]
        return rows[0] if rows else None

    def customer_history(self, customer_code: str) -> list[dict[str, object]]:
        code = str(customer_code or "").strip()
        if not code:
            return []
        filters = CustomerFilters(search_text=code)
        rows = [
            row
            for row in self.all_customer_rows(filters, sort_by="period", sort_desc=False)
            if str(row.get("customer_code") or "") == code
        ]
        previous_balance: float | None = None
        for row in rows:
            current = float(row.get("total_balance") or 0)
            row["difference"] = "" if previous_balance is None else current - previous_balance
            row["growth_rate"] = None if previous_balance in (None, 0) else (current - previous_balance) / previous_balance * 100
            previous_balance = current
        return rows

    def customer_officer_history(self, customer_code: str) -> list[dict[str, object]]:
        code = str(customer_code or "").strip()
        if not code:
            return []
        with self._database() as database:
            rows = database.execute(
                """
                SELECT
                    op.period,
                    op.officer_code AS imported_officer_code,
                    op.officer_name AS imported_officer_name,
                    op.balance_managed,
                    op.source_loan_count,
                    op.is_primary,
                    s.primary_officer_code,
                    s.primary_officer_name
                FROM customer_officer_period op
                LEFT JOIN customer_period_summary s
                    ON s.period = op.period AND s.customer_code = op.customer_code
                WHERE op.customer_code = ?
                ORDER BY op.period, op.is_primary DESC, op.balance_managed DESC, op.id
                """,
                (code,),
            ).fetchall()
            overrides = database.execute(
                """
                SELECT *
                FROM customer_officer_override
                WHERE customer_code = ?
                ORDER BY effective_from_period, id
                """,
                (code,),
            ).fetchall()
        override_rows = [dict(row) for row in overrides]
        output = []
        for row in rows:
            item = dict(row)
            effective = _effective_override_from_rows(override_rows, code, str(item["period"] or ""))
            item["override_officer_code"] = effective.get("officer_code", "") if effective else ""
            item["override_officer_name"] = effective.get("officer_name", "") if effective else ""
            item["override_scope"] = _override_scope_text(effective) if effective else ""
            item["override_reason"] = effective.get("reason", "") if effective else ""
            item["override_created_by"] = effective.get("created_by", "") if effective else ""
            item["override_updated_at"] = effective.get("updated_at", "") if effective else ""
            output.append(item)
        return output

    def explain_customer_query_plans(
        self,
        filters: CustomerFilters | None = None,
        *,
        previous_period: str = "",
        current_period: str = "",
    ) -> dict[str, list[str]]:
        filters = filters or CustomerFilters()
        plans: dict[str, list[str]] = {}
        where, params = _summary_where(filters)
        base_sql = _customer_list_base_sql(where)
        with self._database() as database:
            _record_plan(
                database,
                plans,
                "customer_list_count",
                f"SELECT COUNT(*) FROM ({base_sql}) q",
                params,
            )
            _record_plan(
                database,
                plans,
                "customer_list_page",
                f"SELECT * FROM ({base_sql}) q ORDER BY period DESC, customer_code COLLATE NOCASE DESC LIMIT 100 OFFSET 0",
                params,
            )
            _record_plan(
                database,
                plans,
                "dashboard_metrics",
                f"""
                SELECT SUM(s.total_balance), SUM(s.interest_rate_numerator), SUM(s.nim_before_numerator), SUM(s.nim_after_numerator)
                FROM customer_period_summary s
                {where}
                """,
                params,
            )
            trend_filters = filters.with_current_period("") if filters.current_period else filters
            trend_where, trend_params = _summary_where(trend_filters)
            _record_plan(
                database,
                plans,
                "dashboard_trends",
                f"""
                SELECT s.period, SUM(s.total_balance)
                FROM customer_period_summary s
                {trend_where}
                GROUP BY s.period
                ORDER BY s.period
                """,
                trend_params,
            )
            _record_plan(
                database,
                plans,
                "top_customers",
                f"SELECT * FROM ({base_sql}) q ORDER BY total_balance DESC LIMIT 10",
                params,
            )
            sample_code = str(filters.search_text or "").strip()
            if not sample_code:
                sample = database.execute(
                    "SELECT customer_code FROM customer_period_summary ORDER BY period DESC, customer_code LIMIT 1"
                ).fetchone()
                sample_code = str(sample["customer_code"] or "") if sample else ""
            if sample_code:
                _record_plan(
                    database,
                    plans,
                    "customer_history",
                    """
                    SELECT *
                    FROM customer_period_summary
                    WHERE customer_code = ?
                    ORDER BY period
                    """,
                    [sample_code],
                )
                _record_plan(
                    database,
                    plans,
                    "override_effective",
                    """
                    SELECT *
                    FROM customer_officer_override
                    WHERE customer_code = ?
                      AND is_active = 1
                      AND effective_from_period <= ?
                      AND (effective_to_period = '' OR effective_to_period >= ?)
                    ORDER BY updated_at DESC, id DESC
                    LIMIT 1
                    """,
                    [sample_code, current_period or "9999-99", current_period or "9999-99"],
                )
            if previous_period and current_period:
                movement_sql, movement_params = _movement_base_sql(previous_period, current_period)
                movement_where, movement_where_params = _movement_where(filters)
                movement_params.extend(movement_where_params)
                movement_query = f"SELECT * FROM ({movement_sql}) q {movement_where}"
                _record_plan(database, plans, "movement_count", f"SELECT COUNT(*) FROM ({movement_query}) z", movement_params)
                _record_plan(
                    database,
                    plans,
                    "movement_page",
                    f"{movement_query} ORDER BY difference DESC, customer_code COLLATE NOCASE ASC LIMIT 100 OFFSET 0",
                    movement_params,
                )
                _record_plan(
                    database,
                    plans,
                    "top_movement",
                    f"{movement_query} ORDER BY difference DESC, customer_code COLLATE NOCASE ASC LIMIT 10",
                    movement_params,
                )
            multi_filters = filters if filters.multi_status else CustomerFilters(
                period_from=filters.period_from,
                period_to=filters.period_to,
                current_period=filters.current_period,
                compare_period=filters.compare_period,
                branch_code=filters.branch_code,
                customer_type=filters.customer_type,
                officer=filters.officer,
                loan_term=filters.loan_term,
                search_text=filters.search_text,
                movement_status=filters.movement_status,
                multi_status="same_period",
                override_status=filters.override_status,
                debt_group=filters.debt_group,
            )
            multi_where, multi_params = _summary_where(multi_filters)
            multi_base_sql = _customer_list_base_sql(multi_where)
            _record_plan(
                database,
                plans,
                "multiple_officers",
                f"SELECT * FROM ({multi_base_sql}) q ORDER BY total_balance DESC LIMIT 100 OFFSET 0",
                multi_params,
            )
        return plans

    def action_logs(self, customer_code: str = "") -> list[dict[str, object]]:
        code = str(customer_code or "").strip()
        clause = ""
        params: tuple[object, ...] = ()
        if code:
            clause = "WHERE customer_code = ?"
            params = (code,)
        with self._database() as database:
            rows = database.execute(
                f"""
                SELECT *
                FROM customer_action_log
                {clause}
                ORDER BY created_at DESC, id DESC
                LIMIT 1000
                """,
                params,
            ).fetchall()
        return [dict(row) for row in rows]

    def import_runs(
        self,
        *,
        page: int = 1,
        page_size: int = 100,
        sort_desc: bool = True,
        period: str = "",
    ) -> PageResult:
        page = max(1, int(page or 1))
        page_size = max(1, min(5000, int(page_size or 100)))
        params: list[object] = []
        where = ""
        if period:
            where = "WHERE period = ?"
            params.append(period)
        direction = "DESC" if sort_desc else "ASC"
        with self._database() as database:
            total_rows = int(
                database.execute(f"SELECT COUNT(*) FROM customer_import_runs {where}", params).fetchone()[0] or 0
            )
            rows = database.execute(
                f"""
                SELECT *
                FROM customer_import_runs
                {where}
                ORDER BY started_at {direction}, id {direction}
                LIMIT ? OFFSET ?
                """,
                (*params, page_size, (page - 1) * page_size),
            ).fetchall()
        return PageResult(rows=[dict(row) for row in rows], total_rows=total_rows, page=page, page_size=page_size)

    def import_files(self, run_id: int) -> list[dict[str, object]]:
        with self._database() as database:
            rows = database.execute(
                """
                SELECT *
                FROM customer_import_files
                WHERE run_id = ?
                ORDER BY id
                """,
                (int(run_id),),
            ).fetchall()
        return [dict(row) for row in rows]

    def officer_directory(
        self,
        *,
        search_text: str = "",
        branch_code: str = "",
        status: str = "",
        page: int = 1,
        page_size: int = 100,
        sort_by: str = "officer_name",
        sort_desc: bool = False,
    ) -> PageResult:
        page = max(1, int(page or 1))
        page_size = max(1, min(5000, int(page_size or 100)))
        clauses = ["1 = 1"]
        params: list[object] = []
        text = clean_filter_text(search_text)
        if text:
            pattern = f"%{text}%"
            clauses.append("(officer_code LIKE ? OR officer_name LIKE ?)")
            params.extend([pattern, pattern])
        if branch_code:
            clauses.append("branch_code = ?")
            params.append(branch_code)
        if status == "active":
            clauses.append("is_active = 1")
        elif status == "inactive":
            clauses.append("is_active = 0")
        where = "WHERE " + " AND ".join(clauses)
        order_sql = _order_sql(
            sort_by,
            sort_desc,
            {
                "officer_code": "officer_code",
                "officer_name": "officer_name",
                "branch_code": "branch_code",
                "transaction_office": "transaction_office",
                "is_active": "is_active",
                "updated_at": "updated_at",
            },
            default="officer_name",
        )
        with self._database() as database:
            total_rows = int(database.execute(f"SELECT COUNT(*) FROM customer_officer_directory {where}", params).fetchone()[0] or 0)
            rows = database.execute(
                f"""
                SELECT *
                FROM customer_officer_directory
                {where}
                {order_sql}
                LIMIT ? OFFSET ?
                """,
                (*params, page_size, (page - 1) * page_size),
            ).fetchall()
        return PageResult(rows=[dict(row) for row in rows], total_rows=total_rows, page=page, page_size=page_size)

    def find_officers_by_code_prefix(
        self,
        query: str,
        *,
        branch_code: str = "",
        active_only: bool = True,
        limit: int = 20,
    ) -> list[dict[str, object]]:
        text = normalize_officer_code(query)
        if not text:
            return []
        clauses = ["officer_code LIKE ?"]
        params: list[object] = [f"{text}%"]
        if branch_code:
            clauses.append("branch_code = ?")
            params.append(str(branch_code).strip())
        if active_only:
            clauses.append("is_active = 1")
        with self._database() as database:
            rows = database.execute(
                f"""
                SELECT *
                FROM customer_officer_directory
                WHERE {' AND '.join(clauses)}
                ORDER BY officer_code COLLATE NOCASE, officer_name COLLATE NOCASE
                LIMIT ?
                """,
                (*params, max(1, min(100, int(limit or 20)))),
            ).fetchall()
        return [dict(row) for row in rows]

    def find_officers_by_name(
        self,
        query: str,
        *,
        branch_code: str = "",
        active_only: bool = True,
        limit: int = 20,
    ) -> list[dict[str, object]]:
        text = normalize_officer_name(query)
        if not text:
            return []
        clauses = ["officer_name LIKE ? COLLATE NOCASE"]
        params: list[object] = [f"%{text}%"]
        if branch_code:
            clauses.append("branch_code = ?")
            params.append(str(branch_code).strip())
        if active_only:
            clauses.append("is_active = 1")
        with self._database() as database:
            rows = database.execute(
                f"""
                SELECT *
                FROM customer_officer_directory
                WHERE {' AND '.join(clauses)}
                ORDER BY officer_name COLLATE NOCASE, officer_code COLLATE NOCASE
                LIMIT ?
                """,
                (*params, max(1, min(100, int(limit or 20)))),
            ).fetchall()
            output = [dict(row) for row in rows]
            normalized_query = _strip_vietnamese_accents(text).casefold()
            if len(output) < max(1, int(limit or 20)) and normalized_query != text.casefold():
                fallback_clauses = []
                fallback_params: list[object] = []
                first = text[:1]
                if first:
                    fallback_clauses.append("officer_name LIKE ? COLLATE NOCASE")
                    fallback_params.append(f"{first}%")
                if branch_code:
                    fallback_clauses.append("branch_code = ?")
                    fallback_params.append(str(branch_code).strip())
                if active_only:
                    fallback_clauses.append("is_active = 1")
                fallback_where = " AND ".join(fallback_clauses) if fallback_clauses else "1 = 1"
                fallback_rows = database.execute(
                    f"""
                    SELECT *
                    FROM customer_officer_directory
                    WHERE {fallback_where}
                    ORDER BY officer_name COLLATE NOCASE, officer_code COLLATE NOCASE
                    LIMIT 500
                    """,
                    fallback_params,
                ).fetchall()
                seen = {normalize_officer_code(row.get("officer_code")).casefold() for row in output}
                for row in fallback_rows:
                    candidate = dict(row)
                    code = normalize_officer_code(candidate.get("officer_code")).casefold()
                    if code in seen:
                        continue
                    normalized_name = _strip_vietnamese_accents(candidate.get("officer_name")).casefold()
                    if normalized_query in normalized_name:
                        output.append(candidate)
                        seen.add(code)
                    if len(output) >= max(1, min(100, int(limit or 20))):
                        break
        return output

    def get_officer_by_code(self, officer_code: str, *, active_only: bool = True) -> dict[str, object] | None:
        code = normalize_officer_code(officer_code)
        if not code:
            return None
        clauses = ["officer_code = ?"]
        params: list[object] = [code]
        if active_only:
            clauses.append("is_active = 1")
        with self._database() as database:
            row = database.execute(
                f"""
                SELECT *
                FROM customer_officer_directory
                WHERE {' AND '.join(clauses)}
                LIMIT 1
                """,
                params,
            ).fetchone()
        return dict(row) if row is not None else None

    def get_officer_by_identity(
        self,
        officer_code: str,
        officer_name: str,
        *,
        active_only: bool = True,
    ) -> dict[str, object] | None:
        row = self.get_officer_by_code(officer_code, active_only=active_only)
        if row is None:
            return None
        expected = normalize_officer_name(row.get("officer_name"))
        actual = normalize_officer_name(officer_name)
        return row if expected.casefold() == actual.casefold() else None

    def upsert_officer_directory(
        self,
        *,
        officer_code: str,
        officer_name: str,
        branch_code: str = "",
        transaction_office: str = "",
        is_active: bool = True,
    ) -> None:
        code = normalize_officer_code(officer_code)
        name = normalize_officer_name(officer_name)
        if not code:
            raise CustomerDatabaseError("Mã cán bộ không được để trống.")
        if not name:
            raise CustomerDatabaseError("Tên cán bộ không được để trống.")
        self._assert_write_allowed()
        now = now_text()
        with self._database() as database:
            database.execute(
                """
                INSERT INTO customer_officer_directory(
                    officer_code, officer_name, branch_code, transaction_office, is_active, updated_at
                )
                VALUES (?, ?, ?, ?, ?, ?)
                ON CONFLICT(officer_code) DO UPDATE SET
                    officer_name = excluded.officer_name,
                    branch_code = excluded.branch_code,
                    transaction_office = excluded.transaction_office,
                    is_active = excluded.is_active,
                    updated_at = excluded.updated_at
                """,
                (code, name, branch_code.strip(), transaction_office.strip(), 1 if is_active else 0, now),
            )

    def disable_officer(self, officer_code: str) -> None:
        code = normalize_officer_code(officer_code)
        if not code:
            return
        self._assert_write_allowed()
        with self._database() as database:
            database.execute(
                """
                UPDATE customer_officer_directory
                SET is_active = 0, updated_at = ?
                WHERE officer_code = ?
                """,
                (now_text(), code),
            )

    def create_officer_override(
        self,
        *,
        customer_code: str,
        effective_from_period: str,
        officer_code: str,
        officer_name: str,
        reason: str,
        effective_to_period: str = "",
        created_by: str = "",
        computer_name: str = "",
    ) -> int:
        code = str(customer_code or "").strip()
        from_period = str(effective_from_period or "").strip()
        to_period = str(effective_to_period or "").strip()
        new_officer_code = normalize_officer_code(officer_code)
        new_officer_name = normalize_officer_name(officer_name)
        if not code or not from_period:
            raise CustomerDatabaseError("Thiếu mã khách hàng hoặc kỳ hiệu lực.")
        if to_period and to_period < from_period:
            raise CustomerDatabaseError("Kỳ kết thúc override phải sau hoặc bằng kỳ bắt đầu.")
        if not new_officer_code and not new_officer_name:
            raise CustomerDatabaseError("Thiếu cán bộ override.")
        self._assert_write_allowed()
        now = now_text()
        with self._database() as database:
            old_rows = database.execute(
                """
                SELECT *
                FROM customer_officer_override
                WHERE customer_code = ?
                  AND is_active = 1
                  AND effective_from_period <= CASE WHEN ? = '' THEN '9999-99' ELSE ? END
                  AND (effective_to_period = '' OR effective_to_period >= ?)
                ORDER BY updated_at DESC, id DESC
                """,
                (code, to_period, to_period, from_period),
            ).fetchall()
            if old_rows:
                database.execute(
                    """
                    UPDATE customer_officer_override
                    SET is_active = 0, updated_at = ?
                    WHERE customer_code = ?
                      AND is_active = 1
                      AND effective_from_period <= CASE WHEN ? = '' THEN '9999-99' ELSE ? END
                      AND (effective_to_period = '' OR effective_to_period >= ?)
                    """,
                    (now, code, to_period, to_period, from_period),
                )
            cursor = database.execute(
                """
                INSERT INTO customer_officer_override(
                    customer_code, effective_from_period, effective_to_period,
                    officer_code, officer_name, reason, is_active,
                    created_by, created_at, updated_at
                )
                VALUES (?, ?, ?, ?, ?, ?, 1, ?, ?, ?)
                """,
                (code, from_period, to_period, new_officer_code, new_officer_name, reason.strip(), created_by, now, now),
            )
            old_value = "; ".join(
                f"{row['officer_code']}|{row['officer_name']}|{row['effective_from_period']}->{row['effective_to_period']}"
                for row in old_rows
            )
            database.execute(
                """
                INSERT INTO customer_action_log(
                    action_type, customer_code, period, old_value, new_value,
                    reason, user_name, computer_name, created_at
                )
                VALUES ('OFFICER_OVERRIDE', ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    code,
                    from_period,
                    old_value,
                    f"{new_officer_code}|{new_officer_name}|{from_period}->{to_period}",
                    reason.strip(),
                    created_by,
                    computer_name,
                    now,
                ),
            )
        return int(cursor.lastrowid)

    def restore_imported_officer(
        self,
        *,
        customer_code: str,
        period: str,
        reason: str = "Khôi phục theo dữ liệu import",
        user_name: str = "",
        computer_name: str = "",
    ) -> int:
        code = str(customer_code or "").strip()
        clean_period = str(period or "").strip()
        if not code or not clean_period:
            return 0
        self._assert_write_allowed()
        now = now_text()
        with self._database() as database:
            summary = database.execute(
                """
                SELECT primary_officer_code, primary_officer_name
                FROM customer_period_summary
                WHERE customer_code = ? AND period = ?
                """,
                (code, clean_period),
            ).fetchone()
            rows = database.execute(
                """
                SELECT *
                FROM customer_officer_override
                WHERE customer_code = ?
                  AND is_active = 1
                  AND effective_from_period <= ?
                  AND (effective_to_period = '' OR effective_to_period >= ?)
                ORDER BY updated_at DESC, id DESC
                """,
                (code, clean_period, clean_period),
            ).fetchall()
            ids = [int(row["id"]) for row in rows]
            if ids:
                placeholders = ", ".join("?" for _item in ids)
                database.execute(
                    f"""
                    UPDATE customer_officer_override
                    SET is_active = 0, updated_at = ?
                    WHERE id IN ({placeholders})
                    """,
                    (now, *ids),
                )
            old_value = "; ".join(f"{row['officer_code']}|{row['officer_name']}" for row in rows)
            new_value = ""
            if summary is not None:
                new_value = f"{summary['primary_officer_code']}|{summary['primary_officer_name']}"
            database.execute(
                """
                INSERT INTO customer_action_log(
                    action_type, customer_code, period, old_value, new_value,
                    reason, user_name, computer_name, created_at
                )
                VALUES ('OFFICER_RESTORE', ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (code, clean_period, old_value, new_value, reason.strip(), user_name, computer_name, now),
            )
        return len(ids)

    def customer_period_info(self, period: str) -> dict[str, object]:
        clean_period = str(period or "").strip()
        with self._database() as database:
            summary = database.execute(
                """
                SELECT
                    COUNT(*) AS customer_count,
                    SUM(total_balance) AS total_balance
                FROM customer_period_summary
                WHERE period = ?
                """,
                (clean_period,),
            ).fetchone()
            officer_count = int(
                database.execute(
                    "SELECT COUNT(*) FROM customer_officer_period WHERE period = ?",
                    (clean_period,),
                ).fetchone()[0]
                or 0
            )
            run_count = int(
                database.execute(
                    "SELECT COUNT(*) FROM customer_import_runs WHERE period = ?",
                    (clean_period,),
                ).fetchone()[0]
                or 0
            )
            file_count = int(
                database.execute(
                    "SELECT COUNT(*) FROM customer_import_files WHERE period = ?",
                    (clean_period,),
                ).fetchone()[0]
                or 0
            )
        return {
            "period": clean_period,
            "customer_count": int(summary["customer_count"] or 0) if summary else 0,
            "total_balance": float(summary["total_balance"] or 0) if summary else 0.0,
            "officer_relation_count": officer_count,
            "import_run_count": run_count,
            "import_file_count": file_count,
        }

    def delete_customer_period(
        self,
        period: str,
        *,
        user_name: str = "",
        computer_name: str = "",
        delete_officer_directory: bool = False,
        delete_officer_overrides: bool = False,
        delete_action_log: bool = False,
        backup_before_full_delete: bool = True,
    ) -> dict[str, object]:
        clean_period = str(period or "").strip()
        if not clean_period:
            raise CustomerDatabaseError("Chưa chọn kỳ cần xóa.")
        self._assert_write_allowed()
        info = self.customer_period_info(clean_period)
        periods_before_delete = self.distinct_periods()
        is_last_period = periods_before_delete == [clean_period]
        if (delete_officer_directory or delete_officer_overrides or delete_action_log) and not is_last_period:
            raise CustomerDatabaseError("Chỉ được xóa danh mục/ghi đè/nhật ký khi xóa kỳ dữ liệu cuối cùng.")
        delete_officer_overrides = bool(delete_officer_overrides or delete_officer_directory)
        backup_path = ""
        if is_last_period and backup_before_full_delete:
            backup_path = str(self.backup_database())
        now = now_text()
        deleted_officer_directory_count = 0
        deleted_officer_override_count = 0
        deleted_action_log_count = 0
        with self._database() as database:
            affected = [
                str(row["customer_code"] or "")
                for row in database.execute(
                    "SELECT customer_code FROM customer_period_summary WHERE period = ?",
                    (clean_period,),
                ).fetchall()
            ]
            CustomerRepository._delete_periods(database, [clean_period])
            self._refresh_customer_master_after_period_delete(database, affected, now)
            remaining_period_count = int(
                database.execute(
                    "SELECT COUNT(DISTINCT period) FROM customer_period_summary WHERE period <> ''"
                ).fetchone()[0]
                or 0
            )
            if remaining_period_count == 0:
                database.execute("DELETE FROM customer_master")
                if delete_officer_directory:
                    deleted_officer_directory_count = self._count(database, "customer_officer_directory")
                    database.execute("DELETE FROM customer_officer_directory")
                if delete_officer_overrides:
                    deleted_officer_override_count = self._count(database, "customer_officer_override")
                    database.execute("DELETE FROM customer_officer_override")
                if delete_action_log:
                    deleted_action_log_count = self._count(database, "customer_action_log")
                    database.execute("DELETE FROM customer_action_log")
            if not delete_action_log:
                database.execute(
                    """
                    INSERT INTO customer_action_log(
                        action_type, customer_code, period, old_value, new_value,
                        reason, user_name, computer_name, created_at
                    )
                    VALUES ('DELETE_PERIOD', '', ?, ?, ?, 'Xóa dữ liệu khách hàng theo kỳ', ?, ?, ?)
                    """,
                    (
                        clean_period,
                        json.dumps(info, ensure_ascii=False),
                        json.dumps(
                            {
                                "delete_officer_directory": bool(delete_officer_directory),
                                "delete_officer_overrides": bool(delete_officer_overrides),
                                "deleted_officer_directory_count": deleted_officer_directory_count,
                                "deleted_officer_override_count": deleted_officer_override_count,
                            },
                            ensure_ascii=False,
                        ),
                        user_name,
                        computer_name,
                        now,
                    ),
                )
        info = dict(info)
        info["remaining_period_count"] = remaining_period_count
        info["vacuum_recommended"] = remaining_period_count == 0
        info["backup_path"] = backup_path
        info["delete_officer_directory"] = bool(delete_officer_directory)
        info["delete_officer_overrides"] = bool(delete_officer_overrides)
        info["delete_action_log"] = bool(delete_action_log)
        info["deleted_officer_directory_count"] = deleted_officer_directory_count
        info["deleted_officer_override_count"] = deleted_officer_override_count
        info["deleted_action_log_count"] = deleted_action_log_count
        return info

    def delete_officer_directory(
        self,
        *,
        user_name: str = "",
        computer_name: str = "",
        delete_action_log: bool = False,
        allow_with_period_data: bool = False,
        backup_before_delete: bool = True,
    ) -> dict[str, object]:
        self._assert_write_allowed()
        if self.has_period_data() and not allow_with_period_data:
            raise CustomerDatabaseError("Không thể xóa toàn bộ danh mục CBTD khi còn dữ liệu kỳ.")
        backup_path = str(self.backup_database()) if backup_before_delete else ""
        now = now_text()
        deleted_officer_directory_count = 0
        deleted_officer_override_count = 0
        deleted_action_log_count = 0
        with self._database() as database:
            period_count = self._count_distinct(database, "customer_period_summary", "period")
            if period_count and not allow_with_period_data:
                raise CustomerDatabaseError("Không thể xóa toàn bộ danh mục CBTD khi còn dữ liệu kỳ.")
            deleted_officer_directory_count = self._count(database, "customer_officer_directory")
            deleted_officer_override_count = self._count(database, "customer_officer_override")
            database.execute("DELETE FROM customer_officer_directory")
            database.execute("DELETE FROM customer_officer_override")
            if delete_action_log:
                deleted_action_log_count = self._count(database, "customer_action_log")
                database.execute("DELETE FROM customer_action_log")
            else:
                database.execute(
                    """
                    INSERT INTO customer_action_log(
                        action_type, customer_code, period, old_value, new_value,
                        reason, user_name, computer_name, created_at
                    )
                    VALUES ('DELETE_OFFICER_DIRECTORY', '', '', '', ?, 'Xóa danh mục CBTD độc lập', ?, ?, ?)
                    """,
                    (
                        json.dumps(
                            {
                                "deleted_officer_directory_count": deleted_officer_directory_count,
                                "deleted_officer_override_count": deleted_officer_override_count,
                            },
                            ensure_ascii=False,
                        ),
                        user_name,
                        computer_name,
                        now,
                    ),
                )
        return {
            "backup_path": backup_path,
            "deleted_officer_directory_count": deleted_officer_directory_count,
            "deleted_officer_override_count": deleted_officer_override_count,
            "deleted_action_log_count": deleted_action_log_count,
        }

    def _officer_list_for_customer_periods(self, keys: list[tuple[str, str]]) -> dict[tuple[str, str], str]:
        output: dict[tuple[str, str], str] = {}
        if not keys:
            return output
        with self._database() as database:
            for period, customer_code in keys:
                rows = database.execute(
                    """
                    SELECT officer_code, officer_name, balance_managed
                    FROM customer_officer_period
                    WHERE period = ? AND customer_code = ?
                    ORDER BY is_primary DESC, balance_managed DESC, id
                    """,
                    (period, customer_code),
                ).fetchall()
                output[(period, customer_code)] = "\n".join(
                    f"{str(row['officer_name'] or row['officer_code'] or '').strip()}: {float(row['balance_managed'] or 0):.0f}"
                    for row in rows
                )
        return output

    @staticmethod
    def _refresh_customer_master_after_period_delete(
        database: sqlite3.Connection,
        customer_codes: list[str],
        now: str,
    ) -> None:
        for customer_code in sorted({code for code in customer_codes if code}):
            latest = database.execute(
                """
                SELECT *
                FROM customer_period_summary
                WHERE customer_code = ?
                ORDER BY period DESC, id DESC
                LIMIT 1
                """,
                (customer_code,),
            ).fetchone()
            first = database.execute(
                """
                SELECT MIN(period) AS first_period, MAX(period) AS last_period
                FROM customer_period_summary
                WHERE customer_code = ?
                """,
                (customer_code,),
            ).fetchone()
            if latest is None:
                database.execute(
                    """
                    UPDATE customer_master
                    SET latest_officer_code = '',
                        latest_officer_name = '',
                        last_seen_period = '',
                        is_active = 0,
                        updated_at = ?
                    WHERE customer_code = ?
                    """,
                    (now, customer_code),
                )
                continue
            database.execute(
                """
                UPDATE customer_master
                SET branch_code = ?,
                    customer_sequence = ?,
                    customer_name = ?,
                    customer_type = ?,
                    latest_officer_code = ?,
                    latest_officer_name = ?,
                    first_seen_period = ?,
                    last_seen_period = ?,
                    is_active = 1,
                    updated_at = ?
                WHERE customer_code = ?
                """,
                (
                    latest["branch_code"],
                    latest["customer_sequence"],
                    latest["customer_name"],
                    latest["customer_type"],
                    latest["primary_officer_code"],
                    latest["primary_officer_name"],
                    str(first["first_period"] or "") if first else str(latest["period"] or ""),
                    str(first["last_period"] or "") if first else str(latest["period"] or ""),
                    now,
                    customer_code,
                ),
            )

    def _ensure_result_units(self, result: CustomerAggregationResult, *, updated_by: str = "") -> None:
        seen: set[tuple[str, str]] = set()
        for row in result.office_rows:
            branch_code = str(row.branch_code or "").strip()
            trctcd = normalize_trctcd(row.trctcd)
            if not branch_code:
                continue
            key = (branch_code, trctcd)
            if key in seen:
                continue
            seen.add(key)
            self.unit_directory.ensure_known_unit(
                branch_code,
                trctcd,
                updated_by=updated_by or "customer_import",
            )
        for row in result.summaries:
            branch_code = str(row.branch_code or "").strip()
            if branch_code and (branch_code, "00") not in seen:
                self.unit_directory.ensure_known_unit(
                    branch_code,
                    "00",
                    updated_by=updated_by or "customer_import",
                )
                seen.add((branch_code, "00"))

    @staticmethod
    def _assert_write_allowed() -> None:
        try:
            CustomerDatabaseOperationLock.assert_writable()
        except RuntimeError as exc:
            raise CustomerDatabaseError(str(exc)) from exc

    @staticmethod
    def _pragma_int(database: sqlite3.Connection, pragma_name: str) -> int:
        try:
            return int(database.execute(f"PRAGMA {pragma_name}").fetchone()[0] or 0)
        except Exception:
            LOGGER.exception("Unable to read SQLite PRAGMA %s for Customer.db", pragma_name)
            return 0

    @staticmethod
    def _pragma_text(database: sqlite3.Connection, pragma_name: str) -> str:
        try:
            row = database.execute(f"PRAGMA {pragma_name}").fetchone()
            return str(row[0] or "") if row is not None else ""
        except Exception:
            LOGGER.exception("Unable to read SQLite PRAGMA %s for Customer.db", pragma_name)
            return ""

    @staticmethod
    def _count(database: sqlite3.Connection, table_name: str) -> int:
        return int(database.execute(f"SELECT COUNT(*) FROM {table_name}").fetchone()[0] or 0)

    @staticmethod
    def _table_exists(database: sqlite3.Connection, table_name: str) -> bool:
        return (
            database.execute(
                """
                SELECT 1
                FROM sqlite_master
                WHERE type = 'table' AND name = ?
                LIMIT 1
                """,
                (table_name,),
            ).fetchone()
            is not None
        )

    @staticmethod
    def _count_distinct(database: sqlite3.Connection, table_name: str, column_name: str) -> int:
        return int(
            database.execute(f"SELECT COUNT(DISTINCT {column_name}) FROM {table_name}").fetchone()[0]
            or 0
        )

    @staticmethod
    def _validate_aggregation_balance(result: CustomerAggregationResult) -> None:
        difference = float(result.total_balance or 0) - float(result.source_total_balance or 0)
        if abs(difference) > 0.0001:
            raise CustomerDatabaseError(
                "Tong du no Customer khong khop nguon FTPLN: "
                f"customer={result.total_balance}, source={result.source_total_balance}, diff={difference}."
            )

    @staticmethod
    def _validate_office_aggregation_balance(result: CustomerAggregationResult) -> None:
        office_totals: dict[tuple[str, str], float] = {}
        for row in result.office_rows:
            key = (row.period, row.customer_code)
            office_totals[key] = office_totals.get(key, 0.0) + float(row.total_balance or 0)
        for summary in result.summaries:
            key = (summary.period, summary.customer_code)
            office_total = office_totals.get(key)
            if office_total is None:
                raise CustomerDatabaseError(
                    "Thieu tong hop don vi cho khach hang "
                    f"{summary.customer_code} ky {summary.period}."
                )
            difference = float(summary.total_balance or 0) - office_total
            if abs(difference) > 0.0001:
                raise CustomerDatabaseError(
                    "Tong du no theo don vi khong khop summary: "
                    f"period={summary.period}, customer_code={summary.customer_code}, "
                    f"summary={summary.total_balance}, office={office_total}, diff={difference}."
                )

    @staticmethod
    def _validate_debt_group_aggregation(result: CustomerAggregationResult) -> None:
        for collection_name, rows, total_field in (
            ("customer_period_summary", result.summaries, "total_balance"),
            ("customer_officer_period", result.officer_rows, "balance_managed"),
            ("customer_office_period", result.office_rows, "total_balance"),
        ):
            for row in rows:
                if not bool(getattr(row, "has_debt_group_data", False)):
                    continue
                expected = float(getattr(row, total_field, 0) or 0)
                actual = _debt_group_balance_total(row)
                if abs(expected - actual) > 0.0001:
                    raise CustomerDatabaseError(
                        "Tong du no nhom no khong khop: "
                        f"table={collection_name}, period={getattr(row, 'period', '')}, "
                        f"customer_code={getattr(row, 'customer_code', '')}, "
                        f"total={expected}, debt_groups={actual}, diff={expected - actual}."
                    )

    @staticmethod
    def _periods_with_data(database: sqlite3.Connection, periods: list[str]) -> list[str]:
        if not periods:
            return []
        placeholders = ", ".join("?" for _period in periods)
        rows = database.execute(
            f"""
            SELECT DISTINCT period
            FROM customer_period_summary
            WHERE period IN ({placeholders})
            ORDER BY period
            """,
            periods,
        ).fetchall()
        return [str(row["period"] or "") for row in rows]

    @staticmethod
    def _duplicate_import_files(
        database: sqlite3.Connection,
        files: list[tuple[str, str]],
    ) -> list[dict[str, object]]:
        duplicates: list[dict[str, object]] = []
        for period, file_hash in files:
            if not period or not file_hash:
                continue
            row = database.execute(
                """
                SELECT id, run_id, file_name, period, file_hash, status
                FROM customer_import_files
                WHERE period = ? AND file_hash = ? AND status = 'COMPLETED'
                LIMIT 1
                """,
                (period, file_hash),
            ).fetchone()
            if row is not None:
                duplicates.append(dict(row))
        return duplicates

    @staticmethod
    def _delete_periods(database: sqlite3.Connection, periods: list[str]) -> None:
        if not periods:
            return
        placeholders = ", ".join("?" for _period in periods)
        run_rows = database.execute(
            f"""
            SELECT DISTINCT run_id
            FROM customer_import_files
            WHERE period IN ({placeholders})
            """,
            periods,
        ).fetchall()
        run_ids = [int(row["run_id"]) for row in run_rows if row["run_id"] is not None]
        database.execute(
            f"DELETE FROM customer_period_summary WHERE period IN ({placeholders})",
            periods,
        )
        database.execute(
            f"DELETE FROM customer_officer_period WHERE period IN ({placeholders})",
            periods,
        )
        database.execute(
            f"DELETE FROM customer_office_period WHERE period IN ({placeholders})",
            periods,
        )
        database.execute(
            f"DELETE FROM customer_import_files WHERE period IN ({placeholders})",
            periods,
        )
        if run_ids:
            run_placeholders = ", ".join("?" for _run_id in run_ids)
            database.execute(
                f"DELETE FROM customer_import_runs WHERE id IN ({run_placeholders})",
                run_ids,
            )
        database.execute(
            f"DELETE FROM customer_import_runs WHERE period IN ({placeholders})",
            periods,
        )

    @staticmethod
    def _insert_import_files(
        database: sqlite3.Connection,
        run_id: int,
        result: CustomerAggregationResult,
        now: str,
    ) -> None:
        _ = now
        database.executemany(
            """
            INSERT INTO customer_import_files(
                run_id, file_name, file_path, file_hash, branch_code, period,
                source_row_count, customer_count, status, error_message
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            [
                (
                    run_id,
                    file.file_name,
                    file.file_path,
                    file.file_hash,
                    file.branch_code,
                    file.period,
                    file.source_row_count,
                    file.customer_count,
                    file.status,
                    file.error_message,
                )
                for file in result.files
            ],
        )

    @staticmethod
    def _insert_period_summaries(
        database: sqlite3.Connection,
        run_id: int,
        result: CustomerAggregationResult,
        now: str,
    ) -> None:
        debt_columns_sql = ", ".join(DEBT_GROUP_AGGREGATE_COLUMNS)
        placeholders = ", ".join("?" for _item in range(25 + len(DEBT_GROUP_AGGREGATE_COLUMNS)))
        database.executemany(
            f"""
            INSERT INTO customer_period_summary(
                run_id, period, customer_code, branch_code, customer_sequence,
                customer_name, customer_type, primary_officer_code,
                primary_officer_name, officer_count, has_multiple_officers,
                total_balance, short_term_balance, medium_long_term_balance,
                other_balance, medium_long_ratio, interest_rate_numerator,
                nim_before_numerator, nim_after_numerator, average_rate,
                nim_before, nim_after, source_loan_count, created_at, updated_at,
                {debt_columns_sql}
            )
            VALUES ({placeholders})
            """,
            [
                (
                    run_id,
                    row.period,
                    row.customer_code,
                    row.branch_code,
                    row.customer_sequence,
                    row.customer_name,
                    row.customer_type,
                    row.primary_officer_code,
                    row.primary_officer_name,
                    row.officer_count,
                    1 if row.has_multiple_officers else 0,
                    row.total_balance,
                    row.short_term_balance,
                    row.medium_long_term_balance,
                    row.other_balance,
                    row.medium_long_ratio,
                    row.interest_rate_numerator,
                    row.nim_before_numerator,
                    row.nim_after_numerator,
                    row.average_rate,
                    row.nim_before,
                    row.nim_after,
                    row.source_loan_count,
                    now,
                    now,
                    *_debt_group_values(row),
                )
                for row in result.summaries
            ],
        )

    @staticmethod
    def _insert_officer_rows(
        database: sqlite3.Connection,
        result: CustomerAggregationResult,
        now: str,
    ) -> None:
        debt_columns_sql = ", ".join(DEBT_GROUP_AGGREGATE_COLUMNS)
        placeholders = ", ".join("?" for _item in range(16 + len(DEBT_GROUP_AGGREGATE_COLUMNS)))
        database.executemany(
            f"""
            INSERT INTO customer_officer_period(
                period, customer_code, officer_code, officer_name,
                branch_code, transaction_office, balance_managed,
                short_term_balance, medium_long_term_balance, other_balance,
                source_loan_count, interest_rate_numerator, nim_before_numerator,
                nim_after_numerator, is_primary, created_at,
                {debt_columns_sql}
            )
            VALUES ({placeholders})
            """,
            [
                (
                    row.period,
                    row.customer_code,
                    row.officer_code,
                    row.officer_name,
                    row.branch_code,
                    row.transaction_office,
                    row.balance_managed,
                    row.short_term_balance,
                    row.medium_long_term_balance,
                    row.other_balance,
                    row.source_loan_count,
                    row.interest_rate_numerator,
                    row.nim_before_numerator,
                    row.nim_after_numerator,
                    1 if row.is_primary else 0,
                    now,
                    *_debt_group_values(row),
                )
                for row in result.officer_rows
            ],
        )

    @staticmethod
    def _insert_office_rows(
        database: sqlite3.Connection,
        run_id: int,
        result: CustomerAggregationResult,
        now: str,
    ) -> None:
        debt_columns_sql = ", ".join(DEBT_GROUP_AGGREGATE_COLUMNS)
        placeholders = ", ".join("?" for _item in range(22 + len(DEBT_GROUP_AGGREGATE_COLUMNS)))
        database.executemany(
            f"""
            INSERT INTO customer_office_period(
                run_id, period, customer_code, customer_sequence, branch_code,
                trctcd, office_code, office_name, office_type,
                primary_officer_code, primary_officer_name, officer_count,
                total_balance, short_term_balance, medium_long_term_balance,
                other_balance, interest_rate_numerator, nim_before_numerator,
                nim_after_numerator, source_loan_count, created_at, updated_at,
                {debt_columns_sql}
            )
            VALUES ({placeholders})
            """,
            [
                (
                    run_id,
                    row.period,
                    row.customer_code,
                    row.customer_sequence,
                    row.branch_code,
                    row.trctcd,
                    row.office_code,
                    row.office_name,
                    row.office_type,
                    row.primary_officer_code,
                    row.primary_officer_name,
                    row.officer_count,
                    row.total_balance,
                    row.short_term_balance,
                    row.medium_long_term_balance,
                    row.other_balance,
                    row.interest_rate_numerator,
                    row.nim_before_numerator,
                    row.nim_after_numerator,
                    row.source_loan_count,
                    now,
                    now,
                    *_debt_group_values(row),
                )
                for row in result.office_rows
            ],
        )

    @staticmethod
    def _upsert_customer_master(
        database: sqlite3.Connection,
        result: CustomerAggregationResult,
        now: str,
    ) -> None:
        database.executemany(
            """
            INSERT INTO customer_master(
                customer_code, branch_code, customer_sequence, customer_name,
                customer_type, latest_officer_code, latest_officer_name,
                first_seen_period, last_seen_period, is_active, created_at, updated_at
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, 1, ?, ?)
            ON CONFLICT(customer_code) DO UPDATE SET
                branch_code = excluded.branch_code,
                customer_sequence = excluded.customer_sequence,
                customer_name = CASE
                    WHEN excluded.customer_name <> '' THEN excluded.customer_name
                    ELSE customer_master.customer_name
                END,
                customer_type = CASE
                    WHEN excluded.customer_type <> 'OTHER' THEN excluded.customer_type
                    WHEN customer_master.customer_type = '' THEN excluded.customer_type
                    ELSE customer_master.customer_type
                END,
                first_seen_period = CASE
                    WHEN customer_master.first_seen_period = ''
                      OR excluded.first_seen_period < customer_master.first_seen_period
                    THEN excluded.first_seen_period
                    ELSE customer_master.first_seen_period
                END,
                latest_officer_code = CASE
                    WHEN customer_master.last_seen_period = ''
                      OR excluded.last_seen_period >= customer_master.last_seen_period
                    THEN excluded.latest_officer_code
                    ELSE customer_master.latest_officer_code
                END,
                latest_officer_name = CASE
                    WHEN customer_master.last_seen_period = ''
                      OR excluded.last_seen_period >= customer_master.last_seen_period
                    THEN excluded.latest_officer_name
                    ELSE customer_master.latest_officer_name
                END,
                last_seen_period = CASE
                    WHEN customer_master.last_seen_period = ''
                      OR excluded.last_seen_period > customer_master.last_seen_period
                    THEN excluded.last_seen_period
                    ELSE customer_master.last_seen_period
                END,
                is_active = 1,
                updated_at = excluded.updated_at
            """,
            [
                (
                    row.customer_code,
                    row.branch_code,
                    row.customer_sequence,
                    row.customer_name,
                    row.customer_type,
                    row.primary_officer_code,
                    row.primary_officer_name,
                    row.period,
                    row.period,
                    now,
                    now,
                )
                for row in result.summaries
            ],
        )

    @staticmethod
    def _upsert_officer_directory(
        database: sqlite3.Connection,
        result: CustomerAggregationResult,
        now: str,
    ) -> None:
        rows = [row for row in result.officer_rows if row.officer_code]
        database.executemany(
            """
            INSERT INTO customer_officer_directory(
                officer_code, officer_name, branch_code, transaction_office,
                is_active, updated_at
            )
            VALUES (?, ?, ?, ?, 1, ?)
            ON CONFLICT(officer_code) DO UPDATE SET
                officer_name = CASE
                    WHEN excluded.officer_name <> '' THEN excluded.officer_name
                    ELSE customer_officer_directory.officer_name
                END,
                branch_code = CASE
                    WHEN excluded.branch_code <> '' THEN excluded.branch_code
                    ELSE customer_officer_directory.branch_code
                END,
                transaction_office = CASE
                    WHEN excluded.transaction_office <> '' THEN excluded.transaction_office
                    ELSE customer_officer_directory.transaction_office
                END,
                is_active = 1,
                updated_at = excluded.updated_at
            """,
            [
                (
                    row.officer_code,
                    row.officer_name,
                    row.branch_code,
                    row.transaction_office,
                    now,
                )
                for row in rows
            ],
        )


def _override_lookup_sql(alias: str) -> str:
    return (
        "SELECT o.{column} "
        "FROM customer_officer_override o "
        f"WHERE o.customer_code = {alias}.customer_code "
        "AND o.is_active = 1 "
        f"AND o.effective_from_period <= {alias}.period "
        f"AND (o.effective_to_period = '' OR o.effective_to_period >= {alias}.period) "
        "ORDER BY o.updated_at DESC, o.id DESC "
        "LIMIT 1"
    )


def _override_value_sql(
    alias: str,
    column: str,
    *,
    fallback: str,
    null_if_empty: bool = False,
) -> str:
    lookup = f"({_override_lookup_sql(alias).format(column=column)})"
    if null_if_empty:
        lookup = f"NULLIF({lookup}, '')"
    return f"COALESCE({lookup}, {fallback})"


def _has_override_sql(alias: str) -> str:
    return (
        "EXISTS ("
        "SELECT 1 FROM customer_officer_override o "
        f"WHERE o.customer_code = {alias}.customer_code "
        "AND o.is_active = 1 "
        f"AND o.effective_from_period <= {alias}.period "
        f"AND (o.effective_to_period = '' OR o.effective_to_period >= {alias}.period)"
        ")"
    )


def _override_value_for_expr_sql(
    customer_expr: str,
    period_expr: str,
    column: str,
    *,
    fallback: str,
    null_if_empty: bool = False,
) -> str:
    lookup = (
        "(SELECT o.{column} "
        "FROM customer_officer_override o "
        f"WHERE o.customer_code = {customer_expr} "
        "AND o.is_active = 1 "
        f"AND o.effective_from_period <= {period_expr} "
        f"AND (o.effective_to_period = '' OR o.effective_to_period >= {period_expr}) "
        "ORDER BY o.updated_at DESC, o.id DESC "
        "LIMIT 1)"
    ).format(column=column)
    if null_if_empty:
        lookup = f"NULLIF({lookup}, '')"
    return f"COALESCE({lookup}, {fallback})"


def _has_override_for_expr_sql(customer_expr: str, period_expr: str) -> str:
    return (
        "EXISTS ("
        "SELECT 1 FROM customer_officer_override o "
        f"WHERE o.customer_code = {customer_expr} "
        "AND o.is_active = 1 "
        f"AND o.effective_from_period <= {period_expr} "
        f"AND (o.effective_to_period = '' OR o.effective_to_period >= {period_expr})"
        ")"
    )


def _normalize_cross_branch_scopes(scope_type: object) -> tuple[str, ...]:
    if isinstance(scope_type, (list, tuple, set, frozenset)):
        values = tuple(str(item or "").strip().casefold() for item in scope_type if str(item or "").strip())
    else:
        text = str(scope_type or "cross_branch").strip().casefold()
        values = tuple(part.strip() for part in text.split("|") if part.strip())
    return values or ("cross_branch",)


def _cross_branch_scope_clause(
    scope_type: object,
    minimum_branch_count: int,
    *,
    alias: str = "us",
) -> tuple[str, list[object]]:
    minimum = max(2, int(minimum_branch_count or 2))
    prefix = f"{alias}." if alias else ""
    clauses: list[str] = []
    params: list[object] = []
    for scope in _normalize_cross_branch_scopes(scope_type):
        if scope == "__none__":
            clauses.append("0 = 1")
        elif scope == "cross_branch":
            clauses.append(f"COALESCE({prefix}branch_count, 0) >= ?")
            params.append(minimum)
        elif scope == "head_and_pgd":
            clauses.append(f"COALESCE({prefix}has_head_and_pgd, 0) = 1")
        elif scope == "multi_pgd":
            clauses.append(f"COALESCE({prefix}has_multi_pgd, 0) = 1")
        elif scope == "only_head_office":
            clauses.append(
                f"COALESCE({prefix}branch_count, 0) = 1 "
                f"AND COALESCE({prefix}head_office_count, 0) = 1 "
                f"AND COALESCE({prefix}pgd_count, 0) = 0"
            )
        elif scope == "only_one_pgd":
            clauses.append(
                f"COALESCE({prefix}branch_count, 0) = 1 "
                f"AND COALESCE({prefix}head_office_count, 0) = 0 "
                f"AND COALESCE({prefix}pgd_count, 0) = 1"
            )
        else:
            clauses.append(
                f"COALESCE({prefix}branch_count, 0) >= 2 "
                f"OR COALESCE({prefix}has_head_and_pgd, 0) = 1 "
                f"OR COALESCE({prefix}has_multi_pgd, 0) = 1"
            )
    if not clauses:
        return "0 = 1", []
    if len(clauses) == 1:
        return f"({clauses[0]})", params
    return "(" + " OR ".join(f"({clause})" for clause in clauses) + ")", params


def _cross_branch_candidate_sql(
    period: str,
    filters: CustomerFilters,
    *,
    minimum_branch_count: int = 2,
    scope_type: object = "cross_branch",
    office_code: str = "",
    office_filter_mode: str = "actual",
) -> tuple[str, list[object]]:
    period = str(period or "").strip()
    minimum = max(2, int(minimum_branch_count or 2))
    office_filter = str(office_code or "").strip()
    office_mode = str(office_filter_mode or "actual").strip().casefold()
    params: list[object] = [period, period]
    filters = filters or CustomerFilters()
    clauses = ["1 = 1"]
    scope_clause, scope_params = _cross_branch_scope_clause(scope_type, minimum, alias="c")
    clauses.append(scope_clause)
    params.extend(scope_params)
    if filters.branch_code:
        clauses.append(
            """
            EXISTS (
                SELECT 1
                FROM office_rows selected_branch
                WHERE selected_branch.period = c.period
                  AND selected_branch.customer_sequence = c.customer_sequence
                  AND selected_branch.branch_code = ?
            )
            """
        )
        params.append(filters.branch_code)
    if office_filter:
        if office_mode == "representative":
            clauses.append(
                """
                EXISTS (
                    SELECT 1
                    FROM representative_rows selected_office
                    WHERE selected_office.period = c.period
                      AND selected_office.customer_sequence = c.customer_sequence
                      AND selected_office.office_code = ?
                )
                """
            )
        else:
            clauses.append(
                """
                EXISTS (
                    SELECT 1
                    FROM office_rows selected_office
                    WHERE selected_office.period = c.period
                      AND selected_office.customer_sequence = c.customer_sequence
                      AND selected_office.office_code = ?
                )
                """
            )
        params.append(office_filter)
    if filters.customer_type:
        clauses.append(
            """
            EXISTS (
                SELECT 1
                FROM customer_period_summary selected_type
                WHERE selected_type.period = c.period
                  AND selected_type.customer_sequence = c.customer_sequence
                  AND selected_type.total_balance > 0
                  AND selected_type.customer_type = ?
            )
            """
        )
        params.append(filters.customer_type)
    if filters.officer:
        officer = clean_filter_text(filters.officer)
        code_expr = _override_value_sql("selected_officer", "officer_code", fallback="selected_officer.primary_officer_code")
        name_expr = _override_value_sql(
            "selected_officer",
            "officer_name",
            fallback="selected_officer.primary_officer_name",
            null_if_empty=True,
        )
        clauses.append(
            f"""
            EXISTS (
                SELECT 1
                FROM customer_period_summary selected_officer
                WHERE selected_officer.period = c.period
                  AND selected_officer.customer_sequence = c.customer_sequence
                  AND selected_officer.total_balance > 0
                  AND (
                    {code_expr} = ?
                    OR {name_expr} = ?
                    OR {name_expr} LIKE ?
                  )
            )
            """
        )
        params.extend([officer, officer, f"%{officer}%"])
    if filters.search_text:
        pattern = f"%{clean_filter_text(filters.search_text)}%"
        clauses.append(
            """
            (
                c.customer_sequence LIKE ?
                OR EXISTS (
                    SELECT 1
                    FROM customer_period_summary selected_search
                    WHERE selected_search.period = c.period
                      AND selected_search.customer_sequence = c.customer_sequence
                      AND selected_search.total_balance > 0
                      AND (
                        selected_search.customer_code LIKE ?
                        OR selected_search.customer_name LIKE ?
                      )
                )
            )
            """
        )
        params.extend([pattern, pattern, pattern])
    where = "WHERE " + " AND ".join(clauses)
    representative_cte = ""
    if office_filter and office_mode == "representative":
        representative_cte = """
        ,
        ranked_representative AS (
            SELECT
                o.*,
                bo.has_head_office,
                bo.pgd_count,
                ROW_NUMBER() OVER (
                    PARTITION BY o.period, o.customer_sequence, o.branch_code
                    ORDER BY
                        CASE
                            WHEN COALESCE(bo.has_head_office, 0) = 1 AND o.office_type = 'HEAD_OFFICE' THEN 0
                            WHEN COALESCE(bo.has_head_office, 0) = 0 AND o.office_type = 'TRANSACTION_OFFICE' THEN 0
                            ELSE 1
                        END,
                        CASE
                            WHEN COALESCE(bo.has_head_office, 0) = 0
                             AND COALESCE(bo.pgd_count, 0) > 1
                             AND o.office_type = 'TRANSACTION_OFFICE'
                            THEN -o.total_balance
                            ELSE 0
                        END,
                        o.trctcd COLLATE NOCASE ASC,
                        o.office_code COLLATE NOCASE ASC
                ) AS rn
            FROM office_rows o
            LEFT JOIN branch_office bo
                ON bo.period = o.period
               AND bo.customer_sequence = o.customer_sequence
               AND bo.branch_code = o.branch_code
        ),
        representative_rows AS (
            SELECT *
            FROM ranked_representative
            WHERE rn = 1
        )
        """
    sql = f"""
        WITH office_rows AS (
            SELECT
                o.period,
                o.customer_sequence,
                o.customer_code,
                o.branch_code,
                o.trctcd,
                o.office_code,
                o.office_name,
                o.office_type,
                o.primary_officer_code,
                o.primary_officer_name,
                o.total_balance,
                o.short_term_balance,
                o.medium_long_term_balance,
                o.other_balance,
                o.interest_rate_numerator,
                o.nim_before_numerator,
                o.nim_after_numerator
            FROM customer_office_period o
            WHERE o.period = ?
              AND o.total_balance > 0
              AND o.customer_sequence <> ''
        ),
        branch_office AS (
            SELECT
                period,
                customer_sequence,
                branch_code,
                COUNT(DISTINCT office_code) AS office_count,
                COUNT(DISTINCT CASE WHEN office_type = 'HEAD_OFFICE' THEN office_code END) AS head_office_count,
                COUNT(DISTINCT CASE WHEN office_type = 'TRANSACTION_OFFICE' THEN office_code END) AS pgd_count,
                SUM(CASE WHEN office_type = 'HEAD_OFFICE' THEN total_balance ELSE 0 END) AS head_office_balance,
                SUM(CASE WHEN office_type = 'TRANSACTION_OFFICE' THEN total_balance ELSE 0 END) AS pgd_balance,
                MAX(CASE WHEN office_type = 'HEAD_OFFICE' THEN 1 ELSE 0 END) AS has_head_office,
                MAX(CASE WHEN office_type = 'TRANSACTION_OFFICE' THEN 1 ELSE 0 END) AS has_transaction_office
            FROM office_rows
            GROUP BY period, customer_sequence, branch_code
        ),
        customer_summary AS (
            SELECT
                bo.period,
                bo.customer_sequence,
                COUNT(DISTINCT bo.branch_code) AS branch_count,
                COALESCE(SUM(bo.office_count), 0) AS office_count,
                COALESCE(SUM(bo.head_office_count), 0) AS head_office_count,
                COALESCE(SUM(bo.pgd_count), 0) AS pgd_count,
                COALESCE(SUM(bo.head_office_balance), 0) AS head_office_balance,
                COALESCE(SUM(bo.pgd_balance), 0) AS pgd_balance,
                MAX(CASE WHEN COALESCE(bo.has_head_office, 0) = 1
                          AND COALESCE(bo.has_transaction_office, 0) = 1
                    THEN 1 ELSE 0 END) AS has_head_and_pgd,
                MAX(CASE WHEN COALESCE(bo.pgd_count, 0) >= 2 THEN 1 ELSE 0 END) AS has_multi_pgd
            FROM branch_office bo
            GROUP BY bo.period, bo.customer_sequence
        ),
        balance_summary AS (
            SELECT
                period,
                customer_sequence,
                SUM(total_balance) AS total_balance,
                SUM(short_term_balance) AS short_term_balance,
                SUM(medium_long_term_balance) AS medium_long_term_balance,
                SUM(other_balance) AS other_balance,
                CASE WHEN SUM(total_balance) <> 0
                    THEN SUM(medium_long_term_balance) / SUM(total_balance) * 100
                    ELSE 0
                END AS medium_long_ratio,
                CASE WHEN SUM(total_balance) <> 0
                    THEN SUM(interest_rate_numerator) / SUM(total_balance)
                    ELSE 0
                END AS average_rate,
                CASE WHEN SUM(total_balance) <> 0
                    THEN SUM(nim_before_numerator) / SUM(total_balance)
                    ELSE 0
                END AS nim_before,
                CASE WHEN SUM(total_balance) <> 0
                    THEN SUM(nim_after_numerator) / SUM(total_balance)
                    ELSE 0
                END AS nim_after,
                COUNT(DISTINCT CASE
                    WHEN primary_officer_code <> '' THEN 'C:' || primary_officer_code
                    WHEN primary_officer_name <> '' THEN 'N:' || UPPER(TRIM(primary_officer_name))
                    ELSE NULL
                END) AS officer_count,
                0 AS has_override
            FROM office_rows
            GROUP BY period, customer_sequence
        ),
        summary_info AS (
            SELECT
                s.period,
                s.customer_sequence,
                MIN(NULLIF(TRIM(s.customer_name), '')) AS customer_name
            FROM customer_period_summary s
            WHERE s.period = ?
              AND s.total_balance > 0
              AND s.customer_sequence <> ''
            GROUP BY s.period, s.customer_sequence
        ),
        candidate_rows AS (
            SELECT
                cs.period,
                cs.customer_sequence,
                COALESCE(si.customer_name, '') AS customer_name,
                cs.branch_count,
                cs.office_count,
                cs.head_office_count,
                cs.pgd_count,
                cs.has_head_and_pgd,
                cs.has_multi_pgd,
                cs.head_office_balance,
                cs.pgd_balance,
                bs.total_balance,
                bs.short_term_balance,
                bs.medium_long_term_balance,
                bs.other_balance,
                bs.medium_long_ratio,
                bs.average_rate,
                bs.nim_before,
                bs.nim_after,
                bs.officer_count,
                bs.has_override
            FROM customer_summary cs
            JOIN balance_summary bs
                ON bs.period = cs.period
               AND bs.customer_sequence = cs.customer_sequence
            LEFT JOIN summary_info si
                ON si.period = cs.period
               AND si.customer_sequence = cs.customer_sequence
        )
        {representative_cte},
        matched AS (
            SELECT
                c.period,
                c.customer_sequence,
                c.customer_name,
                '' AS customer_type,
                '' AS customer_type_display,
                c.branch_count,
                c.office_count,
                c.head_office_count,
                c.pgd_count,
                c.has_head_and_pgd,
                c.has_multi_pgd,
                c.head_office_balance,
                c.pgd_balance,
                '' AS branch_list,
                '' AS office_list,
                '' AS representative_office_list,
                '' AS representative_office_type_list,
                '' AS representative_reason_list,
                c.total_balance,
                c.short_term_balance,
                c.medium_long_term_balance,
                c.other_balance,
                c.medium_long_ratio,
                c.average_rate,
                c.nim_before,
                c.nim_after,
                c.officer_count,
                '' AS officer_list,
                c.has_override,
                0 AS name_conflict,
                0 AS customer_type_conflict,
                'Không' AS conflict_status
            FROM candidate_rows c
            {where}
        )
        SELECT *
        FROM matched
    """
    return sql, params


def _cross_branch_select_sql(
    period: str,
    filters: CustomerFilters,
    *,
    minimum_branch_count: int = 2,
    scope_type: object = "cross_branch",
    office_code: str = "",
    office_filter_mode: str = "actual",
) -> tuple[str, list[object]]:
    return _cross_branch_candidate_sql(
        period,
        filters,
        minimum_branch_count=minimum_branch_count,
        scope_type=scope_type,
        office_code=office_code,
        office_filter_mode=office_filter_mode,
    )


def _cross_branch_order_sql(sort_by: str, sort_order: str) -> str:
    allowed = {
        "branch_count": "branch_count",
        "office_count": "office_count",
        "pgd_count": "pgd_count",
        "total_balance": "total_balance",
        "nim_after": "nim_after",
        "medium_long_ratio": "medium_long_ratio",
        "customer_name": "customer_name COLLATE NOCASE",
    }
    column = allowed.get(str(sort_by or ""), "branch_count")
    direction = "ASC" if str(sort_order or "").casefold() == "asc" else "DESC"
    if str(sort_by or "") == "customer_name":
        return f"ORDER BY {column} {direction}, branch_count DESC, total_balance DESC, customer_sequence COLLATE NOCASE ASC"
    if str(sort_by or "") == "branch_count" and direction == "DESC":
        return "ORDER BY branch_count DESC, total_balance DESC, customer_sequence COLLATE NOCASE ASC"
    return f"ORDER BY {column} {direction}, branch_count DESC, total_balance DESC, customer_sequence COLLATE NOCASE ASC"


def _empty_cross_branch_kpis() -> dict[str, object]:
    return {
        "cross_customer_count": 0,
        "total_balance": 0,
        "branch_occurrence_count": 0,
        "office_occurrence_count": 0,
        "head_office_occurrence_count": 0,
        "pgd_occurrence_count": 0,
        "head_and_pgd_customer_count": 0,
        "multi_pgd_customer_count": 0,
        "head_office_balance": 0,
        "pgd_balance": 0,
        "three_branch_customer_count": 0,
        "multiple_officer_customer_count": 0,
        "override_customer_count": 0,
        "top_branch_code": "",
        "top_branch_name": "",
        "top_branch_customer_count": 0,
    }


def _cross_branch_kpis_from_temp(
    database: sqlite3.Connection,
    period: str,
    unit_directory: UnitDirectoryService | None = None,
) -> dict[str, object]:
    row = database.execute(
        """
        SELECT
            COUNT(*) AS cross_customer_count,
            COALESCE(SUM(total_balance), 0) AS total_balance,
            COALESCE(SUM(branch_count), 0) AS branch_occurrence_count,
            COALESCE(SUM(office_count), 0) AS office_occurrence_count,
            COALESCE(SUM(head_office_count), 0) AS head_office_occurrence_count,
            COALESCE(SUM(pgd_count), 0) AS pgd_occurrence_count,
            COALESCE(SUM(CASE WHEN has_head_and_pgd = 1 THEN 1 ELSE 0 END), 0) AS head_and_pgd_customer_count,
            COALESCE(SUM(CASE WHEN has_multi_pgd = 1 THEN 1 ELSE 0 END), 0) AS multi_pgd_customer_count,
            COALESCE(SUM(head_office_balance), 0) AS head_office_balance,
            COALESCE(SUM(pgd_balance), 0) AS pgd_balance,
            COALESCE(SUM(CASE WHEN branch_count >= 3 THEN 1 ELSE 0 END), 0) AS three_branch_customer_count
        FROM temp_cross_branch_candidates
        """
    ).fetchone()
    output = dict(row) if row is not None else _empty_cross_branch_kpis()
    code_expr = _override_value_sql("s", "officer_code", fallback="s.primary_officer_code")
    name_expr = _override_value_sql("s", "officer_name", fallback="s.primary_officer_name", null_if_empty=True)
    has_override = _has_override_sql("s")
    officer_row = database.execute(
        f"""
        WITH officer_status AS (
            SELECT
                c.customer_sequence,
                COUNT(DISTINCT CASE
                    WHEN {code_expr} <> '' THEN 'C:' || {code_expr}
                    WHEN {name_expr} <> '' THEN 'N:' || UPPER(TRIM({name_expr}))
                    ELSE NULL
                END) AS effective_officer_count,
                MAX(CASE WHEN {has_override} THEN 1 ELSE 0 END) AS has_override
            FROM temp_cross_branch_candidates c
            JOIN customer_period_summary s
                ON s.period = c.period
               AND s.customer_sequence = c.customer_sequence
               AND s.total_balance > 0
            GROUP BY c.customer_sequence
        )
        SELECT
            COALESCE(SUM(CASE WHEN effective_officer_count >= 2 THEN 1 ELSE 0 END), 0) AS multiple_officer_customer_count,
            COALESCE(SUM(CASE WHEN has_override = 1 THEN 1 ELSE 0 END), 0) AS override_customer_count
        FROM officer_status
        """
    ).fetchone()
    if officer_row is not None:
        output["multiple_officer_customer_count"] = int(officer_row["multiple_officer_customer_count"] or 0)
        output["override_customer_count"] = int(officer_row["override_customer_count"] or 0)
    top = database.execute(
        """
        SELECT ar.branch_code, COUNT(DISTINCT ar.customer_sequence) AS customer_count
        FROM temp_cross_branch_candidates c
        JOIN customer_period_summary ar
            ON ar.period = c.period
           AND ar.customer_sequence = c.customer_sequence
           AND ar.total_balance > 0
        GROUP BY ar.branch_code
        ORDER BY customer_count DESC, ar.branch_code ASC
        LIMIT 1
        """
    ).fetchone()
    if top is not None and top["branch_code"]:
        branch_code = str(top["branch_code"] or "")
        output["top_branch_code"] = branch_code
        output["top_branch_name"] = _branch_display(branch_code, unit_directory)
        output["top_branch_customer_count"] = int(top["customer_count"] or 0)
    else:
        output["top_branch_code"] = ""
        output["top_branch_name"] = ""
        output["top_branch_customer_count"] = 0
    for key, value in _empty_cross_branch_kpis().items():
        output.setdefault(key, value)
    return output


def _finalize_cross_branch_detail_row(
    row: dict[str, object],
    unit_directory: UnitDirectoryService | None = None,
) -> dict[str, object]:
    branch_code = str(row.get("branch_code") or "")
    row["branch_name"] = _branch_display(branch_code, unit_directory)
    row["office_type_display"] = OFFICE_TYPE_LABELS.get(
        str(row.get("office_type") or "").upper(),
        OFFICE_TYPE_LABELS[CustomerOfficeType.UNKNOWN.value],
    )
    office_display = _dynamic_office_display(row, unit_directory)
    row["office_display"] = office_display
    row["office_name"] = _dynamic_office_name(row, unit_directory)
    row["customer_type_display"] = CUSTOMER_TYPE_LABELS.get(str(row.get("customer_type") or "").upper(), CUSTOMER_TYPE_LABELS["OTHER"])
    row["imported_officer_display"] = str(row.get("imported_officer_name") or row.get("imported_officer_code") or "")
    row["effective_officer_display"] = str(row.get("effective_officer_name") or row.get("effective_officer_code") or "")
    row["override_status"] = "Có override" if int(row.get("has_override") or 0) else "Không override"
    return row


def _finalize_office_option(
    row: dict[str, object],
    unit_directory: UnitDirectoryService | None = None,
) -> dict[str, object]:
    row["office_type_display"] = OFFICE_TYPE_LABELS.get(
        str(row.get("office_type") or "").upper(),
        OFFICE_TYPE_LABELS[CustomerOfficeType.UNKNOWN.value],
    )
    row["office_display"] = _dynamic_office_display(row, unit_directory)
    return row


def _branch_display(branch_code: object, unit_directory: UnitDirectoryService | None = None) -> str:
    code = str(branch_code or "").strip()
    if not code:
        return ""
    if unit_directory is not None:
        return unit_directory.get_branch_display_name(code)
    return code


def _dynamic_office_display(
    row: dict[str, object],
    unit_directory: UnitDirectoryService | None = None,
) -> str:
    branch_code = str(row.get("branch_code") or "").strip()
    trctcd = normalize_trctcd(row.get("trctcd"))
    if unit_directory is not None and branch_code and trctcd:
        office = unit_directory.get_office(branch_code, trctcd)
        if office is not None:
            return unit_directory.get_office_display_name(branch_code, trctcd)
    return _office_display(row)


def _dynamic_office_name(
    row: dict[str, object],
    unit_directory: UnitDirectoryService | None = None,
) -> str:
    branch_code = str(row.get("branch_code") or "").strip()
    trctcd = normalize_trctcd(row.get("trctcd"))
    if unit_directory is not None and branch_code and trctcd:
        office = unit_directory.get_office(branch_code, trctcd)
        if office is not None:
            return office.short_name or office.office_name or str(row.get("office_name") or "").strip()
    return str(row.get("office_name") or "").strip()


def _office_display(row: dict[str, object]) -> str:
    code = str(row.get("office_code") or "").strip()
    name = str(row.get("office_name") or "").strip()
    if code and name:
        return f"{code} - {name}"
    return code or name


def _number(value: object) -> float:
    try:
        return float(value or 0)
    except (TypeError, ValueError):
        return 0.0


def _table_columns(database: sqlite3.Connection, table_name: str) -> set[str]:
    return {
        str(row[1] or "")
        for row in database.execute(f"PRAGMA table_info({table_name})").fetchall()
    }


def _debt_group_values(row: object) -> tuple[object, ...]:
    values: list[object] = []
    for column in DEBT_GROUP_AGGREGATE_COLUMNS:
        value = getattr(row, column, 0)
        if column == "has_debt_group_data":
            values.append(1 if bool(value) else 0)
        else:
            values.append(value)
    return tuple(values)


def _debt_group_balance_total(row: object) -> float:
    return sum(
        float(getattr(row, f"debt_group_{suffix}_balance", 0) or 0)
        for suffix in ("1", "2", "3", "4", "5", "unknown")
    )


def _ratio(numerator: float, denominator: float) -> float:
    return numerator / denominator if denominator else 0.0


def _office_status(rows: list[dict[str, object]]) -> dict[str, object]:
    positive = [row for row in rows if _number(row.get("total_balance")) > 0]
    branch_codes = {str(row.get("branch_code") or "") for row in positive if str(row.get("branch_code") or "")}
    office_codes = {str(row.get("office_code") or "") for row in positive if str(row.get("office_code") or "")}
    head_codes = {
        str(row.get("office_code") or "")
        for row in positive
        if str(row.get("office_type") or "") == CustomerOfficeType.HEAD_OFFICE.value
        and str(row.get("office_code") or "")
    }
    pgd_codes = {
        str(row.get("office_code") or "")
        for row in positive
        if str(row.get("office_type") or "") == CustomerOfficeType.TRANSACTION_OFFICE.value
        and str(row.get("office_code") or "")
    }
    has_head_and_pgd = False
    has_multi_pgd = False
    for branch in branch_codes:
        branch_rows = [row for row in positive if str(row.get("branch_code") or "") == branch]
        has_head = any(str(row.get("office_type") or "") == CustomerOfficeType.HEAD_OFFICE.value for row in branch_rows)
        pgd_count = len(
            {
                str(row.get("office_code") or "")
                for row in branch_rows
                if str(row.get("office_type") or "") == CustomerOfficeType.TRANSACTION_OFFICE.value
                and str(row.get("office_code") or "")
            }
        )
        has_head_and_pgd = has_head_and_pgd or (has_head and pgd_count > 0)
        has_multi_pgd = has_multi_pgd or pgd_count >= 2
    return {
        "branch_count": len(branch_codes),
        "office_count": len(office_codes),
        "head_office_count": len(head_codes),
        "pgd_count": len(pgd_codes),
        "has_head_and_pgd": 1 if has_head_and_pgd else 0,
        "has_multi_pgd": 1 if has_multi_pgd else 0,
        "head_office_balance": sum(
            _number(row.get("total_balance"))
            for row in positive
            if str(row.get("office_type") or "") == CustomerOfficeType.HEAD_OFFICE.value
        ),
        "pgd_balance": sum(
            _number(row.get("total_balance"))
            for row in positive
            if str(row.get("office_type") or "") == CustomerOfficeType.TRANSACTION_OFFICE.value
        ),
    }


def _scope_matches_office_rows(rows: list[dict[str, object]], scope_type: object, *, minimum_branch_count: int = 2) -> bool:
    scopes = _normalize_cross_branch_scopes(scope_type) if scope_type not in ("", None, "all") else ("all",)
    status = _office_status(rows)
    for scope in scopes:
        if scope in {"", "all"}:
            return True
        if scope == "cross_branch" and int(status["branch_count"] or 0) >= max(2, int(minimum_branch_count or 2)):
            return True
        if scope == "head_and_pgd" and int(status["has_head_and_pgd"] or 0) == 1:
            return True
        if scope == "multi_pgd" and int(status["has_multi_pgd"] or 0) == 1:
            return True
        if (
            scope == "only_head_office"
            and int(status["branch_count"] or 0) == 1
            and int(status["head_office_count"] or 0) == 1
            and int(status["pgd_count"] or 0) == 0
        ):
            return True
        if (
            scope == "only_one_pgd"
            and int(status["branch_count"] or 0) == 1
            and int(status["head_office_count"] or 0) == 0
            and int(status["pgd_count"] or 0) == 1
        ):
            return True
        if scope not in {"cross_branch", "head_and_pgd", "multi_pgd", "only_head_office", "only_one_pgd"} and (
            int(status["branch_count"] or 0) >= 2
            or int(status["has_head_and_pgd"] or 0) == 1
            or int(status["has_multi_pgd"] or 0) == 1
        ):
            return True
    return False


def _office_kpis(
    period: str,
    sequence: str,
    rows: list[dict[str, object]],
    summary_rows: list[dict[str, object]],
    *,
    missing: bool,
) -> dict[str, object]:
    total_balance = sum(_number(row.get("total_balance")) for row in rows)
    interest_numerator = sum(_number(row.get("interest_rate_numerator")) for row in rows)
    nim_before_numerator = sum(_number(row.get("nim_before_numerator")) for row in rows)
    nim_after_numerator = sum(_number(row.get("nim_after_numerator")) for row in rows)
    status = _office_status(rows)
    types = {str(row.get("customer_type") or "").strip().upper() or "OTHER" for row in rows}
    if not types:
        types = {str(row.get("customer_type") or "").strip().upper() or "OTHER" for row in summary_rows}
    names = [str(row.get("customer_name") or "").strip() for row in rows]
    if not any(names):
        names = [str(row.get("customer_name") or "").strip() for row in summary_rows]
    return {
        "period": period,
        "customer_sequence": sequence,
        "customer_name": next((name for name in names if name), ""),
        "customer_type": "MIXED" if len(types) > 1 else next(iter(types), ""),
        "customer_type_display": "Không thống nhất" if len(types) > 1 else CUSTOMER_TYPE_LABELS.get(next(iter(types), ""), CUSTOMER_TYPE_LABELS["OTHER"]),
        "branch_count": status["branch_count"],
        "office_count": status["office_count"],
        "head_office_count": status["head_office_count"],
        "pgd_count": status["pgd_count"],
        "has_head_and_pgd": status["has_head_and_pgd"],
        "has_multi_pgd": status["has_multi_pgd"],
        "total_balance": total_balance,
        "head_office_balance": status["head_office_balance"],
        "pgd_balance": status["pgd_balance"],
        "average_rate": _ratio(interest_numerator, total_balance),
        "nim_before": _ratio(nim_before_numerator, total_balance),
        "nim_after": _ratio(nim_after_numerator, total_balance),
        "office_detail_missing": 1 if missing else 0,
    }


def _empty_filtered_kpis(
    period: str,
    sequence: str,
    summary_rows: list[dict[str, object]],
    *,
    missing: bool,
) -> dict[str, object]:
    return _office_kpis(period, sequence, [], summary_rows, missing=missing)


def _missing_office_detail_kpis(
    period: str,
    sequence: str,
    summary_rows: list[dict[str, object]],
    unit_directory: UnitDirectoryService | None = None,
) -> dict[str, object]:
    item = _summary_history_row(period, sequence, summary_rows, missing=True, unit_directory=unit_directory)
    names = [str(row.get("customer_name") or "").strip() for row in summary_rows]
    types = {str(row.get("customer_type") or "").strip().upper() or "OTHER" for row in summary_rows}
    return {
        "period": period,
        "customer_sequence": sequence,
        "customer_name": next((name for name in names if name), ""),
        "customer_type": "MIXED" if len(types) > 1 else next(iter(types), ""),
        "customer_type_display": "Không thống nhất" if len(types) > 1 else CUSTOMER_TYPE_LABELS.get(next(iter(types), ""), CUSTOMER_TYPE_LABELS["OTHER"]),
        "branch_count": item["branch_count"],
        "office_count": 0,
        "head_office_count": 0,
        "pgd_count": 0,
        "has_head_and_pgd": 0,
        "has_multi_pgd": 0,
        "total_balance": item["total_balance"],
        "head_office_balance": 0,
        "pgd_balance": 0,
        "average_rate": item["average_rate"],
        "nim_before": item["nim_before"],
        "nim_after": item["nim_after"],
        "office_detail_missing": 1,
    }


def _office_history_row(
    period: str,
    sequence: str,
    rows: list[dict[str, object]],
    summary_rows: list[dict[str, object]],
    *,
    missing: bool,
    unit_directory: UnitDirectoryService | None = None,
) -> dict[str, object]:
    kpis = _office_kpis(period, sequence, rows, summary_rows, missing=missing)
    branch_codes = sorted({str(row.get("branch_code") or "") for row in rows if str(row.get("branch_code") or "")})
    office_displays = [_dynamic_office_display(row, unit_directory) for row in rows if _dynamic_office_display(row, unit_directory)]
    return {
        **kpis,
        "branch_list": "\n".join(_branch_display(code, unit_directory) for code in branch_codes),
        "office_list": "\n".join(office_displays),
        "difference": "",
    }


def _summary_history_row(
    period: str,
    sequence: str,
    summary_rows: list[dict[str, object]],
    *,
    missing: bool,
    unit_directory: UnitDirectoryService | None = None,
) -> dict[str, object]:
    total_balance = sum(_number(row.get("total_balance")) for row in summary_rows)
    interest_numerator = sum(_number(row.get("interest_rate_numerator")) for row in summary_rows)
    nim_before_numerator = sum(_number(row.get("nim_before_numerator")) for row in summary_rows)
    nim_after_numerator = sum(_number(row.get("nim_after_numerator")) for row in summary_rows)
    branch_codes = sorted({str(row.get("branch_code") or "") for row in summary_rows if str(row.get("branch_code") or "")})
    return {
        "period": period,
        "customer_sequence": sequence,
        "branch_count": len(branch_codes),
        "office_count": 0,
        "head_office_count": 0,
        "pgd_count": 0,
        "has_head_and_pgd": 0,
        "has_multi_pgd": 0,
        "branch_list": "\n".join(_branch_display(code, unit_directory) for code in branch_codes),
        "office_list": "",
        "total_balance": total_balance,
        "head_office_balance": 0,
        "pgd_balance": 0,
        "difference": "",
        "average_rate": _ratio(interest_numerator, total_balance),
        "nim_before": _ratio(nim_before_numerator, total_balance),
        "nim_after": _ratio(nim_after_numerator, total_balance),
        "office_detail_missing": 1 if missing else 0,
    }


def _split_codes(value: object) -> list[str]:
    text = "" if value is None else str(value)
    separators = ["\n", ";"]
    for separator in separators:
        text = text.replace(separator, ",")
    return sorted({part.strip() for part in text.split(",") if part.strip()})


def _normalize_customer_name(value: object) -> str:
    return " ".join(str(value or "").strip().split()).casefold()


def _officer_identity(code: object, name: object) -> str:
    clean_code = str(code or "").strip()
    if clean_code:
        return f"C:{clean_code}"
    clean_name = normalize_officer_name(name).casefold()
    return f"N:{clean_name}" if clean_name else ""


def _summary_where(
    filters: CustomerFilters,
    *,
    exclude: set[str] | None = None,
) -> tuple[str, list[object]]:
    exclude = exclude or set()
    clauses = ["1 = 1"]
    params: list[object] = []
    if filters.current_period and "current_period" not in exclude:
        clauses.append("s.period = ?")
        params.append(filters.current_period)
    else:
        if filters.period_from and "period_from" not in exclude:
            clauses.append("s.period >= ?")
            params.append(filters.period_from)
        if filters.period_to and "period_to" not in exclude:
            clauses.append("s.period <= ?")
            params.append(filters.period_to)
    if filters.branch_code and "branch_code" not in exclude:
        clauses.append("s.branch_code = ?")
        params.append(filters.branch_code)
    if filters.customer_type and "customer_type" not in exclude:
        clauses.append("s.customer_type = ?")
        params.append(filters.customer_type)
    if filters.loan_term and "loan_term" not in exclude:
        if filters.loan_term == "SHORT_TERM":
            clauses.append("s.short_term_balance > 0")
        elif filters.loan_term == "MEDIUM_LONG_TERM":
            clauses.append("s.medium_long_term_balance > 0")
        elif filters.loan_term == "OTHER":
            clauses.append("s.other_balance > 0")
    if filters.search_text and "search_text" not in exclude:
        pattern = f"%{clean_filter_text(filters.search_text)}%"
        clauses.append("(s.customer_code LIKE ? OR s.customer_name LIKE ?)")
        params.extend([pattern, pattern])
    if filters.officer and "officer" not in exclude:
        code_expr = _override_value_sql("s", "officer_code", fallback="s.primary_officer_code")
        name_expr = _override_value_sql("s", "officer_name", fallback="s.primary_officer_name", null_if_empty=True)
        officer = clean_filter_text(filters.officer)
        clauses.append(f"({code_expr} = ? OR {name_expr} = ? OR {name_expr} LIKE ?)")
        params.extend([officer, officer, f"%{officer}%"])
    override_status = filters.override_status
    if filters.multi_status in {"override", "no_override"} and not override_status:
        override_status = filters.multi_status
    if override_status and "override_status" not in exclude:
        has_override = _has_override_sql("s")
        if override_status == "override":
            clauses.append(has_override)
        elif override_status == "no_override":
            clauses.append(f"NOT {has_override}")
    if filters.multi_status and "multi_status" not in exclude:
        if filters.multi_status in {"multiple", "same_period"}:
            clauses.append("s.has_multiple_officers = 1")
        elif filters.multi_status == "one":
            clauses.append("s.has_multiple_officers = 0")
        elif filters.multi_status == "changed_period":
            clauses.append(
                """
                EXISTS (
                    SELECT 1
                    FROM customer_period_summary x
                    WHERE x.customer_code = s.customer_code
                      AND x.period <> s.period
                      AND (
                        COALESCE(NULLIF(x.primary_officer_code, ''), x.primary_officer_name)
                        <> COALESCE(NULLIF(s.primary_officer_code, ''), s.primary_officer_name)
                      )
                )
                """
            )
    if filters.debt_group and "debt_group" not in exclude:
        _append_debt_group_filter(
            clauses,
            params,
            filters.debt_group,
            alias="s",
            total_column="total_balance",
        )
    return "WHERE " + " AND ".join(clauses), params


def _debt_report_filters(filters: CustomerFilters | None, report_period: str) -> CustomerFilters:
    base = filters or CustomerFilters()
    period = str(report_period or base.current_period or "").strip()
    return replace(base, current_period=period) if period else base


def _append_debt_group_filter(
    clauses: list[str],
    params: list[object],
    filter_key: object,
    *,
    alias: str,
    total_column: str,
) -> None:
    key = str(filter_key or "").strip().upper()
    if key in {"", DEBT_GROUP_ALL}:
        return
    prefix = f"{alias}."
    if key == DEBT_GROUP_HAS_GROUP_1:
        clauses.append(f"COALESCE({prefix}debt_group_1_balance, 0) > 0")
    elif key in {DEBT_GROUP_HAS_GROUP_2, DEBT_GROUP_ATTENTION}:
        clauses.append(f"COALESCE({prefix}debt_group_2_balance, 0) > 0")
    elif key == DEBT_GROUP_HAS_GROUP_3:
        clauses.append(f"COALESCE({prefix}debt_group_3_balance, 0) > 0")
    elif key == DEBT_GROUP_HAS_GROUP_4:
        clauses.append(f"COALESCE({prefix}debt_group_4_balance, 0) > 0")
    elif key == DEBT_GROUP_HAS_GROUP_5:
        clauses.append(f"COALESCE({prefix}debt_group_5_balance, 0) > 0")
    elif key == DEBT_GROUP_BAD_DEBT:
        clauses.append(
            f"(COALESCE({prefix}debt_group_3_balance, 0) "
            f"+ COALESCE({prefix}debt_group_4_balance, 0) "
            f"+ COALESCE({prefix}debt_group_5_balance, 0)) > 0"
        )
    elif key == DEBT_GROUP_UNKNOWN:
        clauses.append(f"COALESCE({prefix}debt_group_unknown_balance, 0) > 0")
    elif key in {
        DEBT_GROUP_WORST_1,
        DEBT_GROUP_WORST_2,
        DEBT_GROUP_WORST_3,
        DEBT_GROUP_WORST_4,
        DEBT_GROUP_WORST_5,
    }:
        code = key.rsplit("_", 1)[-1]
        clauses.append(f"{prefix}worst_debt_group = ?")
        params.append(f"{int(code):02d}")
    else:
        clauses.append(f"COALESCE({prefix}{total_column}, 0) >= 0")


def _debt_group_aggregate_select_sql(
    *,
    alias: str,
    total_column: str,
    customer_count_expr: str = "COUNT(DISTINCT {alias}.customer_code)",
) -> str:
    prefix = f"{alias}."
    count_expr = customer_count_expr.format(alias=alias)
    return f"""
        MAX(COALESCE({prefix}has_debt_group_data, 0)) AS has_debt_group_data,
        {count_expr} AS customer_count,
        COALESCE(SUM({prefix}{total_column}), 0) AS total_balance,
        COALESCE(SUM({prefix}debt_group_1_balance), 0) AS debt_group_1_balance,
        COALESCE(SUM({prefix}debt_group_2_balance), 0) AS debt_group_2_balance,
        COALESCE(SUM({prefix}debt_group_3_balance), 0) AS debt_group_3_balance,
        COALESCE(SUM({prefix}debt_group_4_balance), 0) AS debt_group_4_balance,
        COALESCE(SUM({prefix}debt_group_5_balance), 0) AS debt_group_5_balance,
        COALESCE(SUM({prefix}debt_group_unknown_balance), 0) AS debt_group_unknown_balance,
        COALESCE(SUM(
            {prefix}debt_group_3_balance + {prefix}debt_group_4_balance + {prefix}debt_group_5_balance
        ), 0) AS bad_debt_balance,
        COALESCE(SUM({prefix}interest_rate_numerator), 0) AS interest_rate_numerator,
        COALESCE(SUM({prefix}nim_before_numerator), 0) AS nim_before_numerator,
        COALESCE(SUM({prefix}nim_after_numerator), 0) AS nim_after_numerator,
        CASE WHEN COALESCE(SUM({prefix}{total_column}), 0) <> 0
            THEN COALESCE(SUM({prefix}debt_group_2_balance), 0) / SUM({prefix}{total_column}) * 100
            ELSE NULL
        END AS attention_ratio,
        CASE WHEN COALESCE(SUM({prefix}{total_column}), 0) <> 0
            THEN COALESCE(SUM(
                {prefix}debt_group_3_balance + {prefix}debt_group_4_balance + {prefix}debt_group_5_balance
            ), 0) / SUM({prefix}{total_column}) * 100
            ELSE NULL
        END AS bad_debt_ratio,
        CASE WHEN COALESCE(SUM({prefix}{total_column}), 0) <> 0
            THEN COALESCE(SUM({prefix}interest_rate_numerator), 0) / SUM({prefix}{total_column})
            ELSE NULL
        END AS average_rate,
        CASE WHEN COALESCE(SUM({prefix}{total_column}), 0) <> 0
            THEN COALESCE(SUM({prefix}nim_before_numerator), 0) / SUM({prefix}{total_column})
            ELSE NULL
        END AS nim_before,
        CASE WHEN COALESCE(SUM({prefix}{total_column}), 0) <> 0
            THEN COALESCE(SUM({prefix}nim_after_numerator), 0) / SUM({prefix}{total_column})
            ELSE NULL
        END AS nim_after,
        SUM(CASE WHEN {prefix}debt_group_2_balance > 0 THEN 1 ELSE 0 END) AS attention_customer_count,
        SUM(CASE WHEN (
            {prefix}debt_group_3_balance + {prefix}debt_group_4_balance + {prefix}debt_group_5_balance
        ) > 0 THEN 1 ELSE 0 END) AS bad_debt_customer_count,
        SUM(CASE WHEN {prefix}debt_group_unknown_balance > 0 THEN 1 ELSE 0 END) AS unknown_customer_count
    """


def _debt_group_by_branch_base_sql(filters: CustomerFilters) -> tuple[str, list[object]]:
    where, params = _summary_where(filters)
    sql = f"""
        SELECT
            s.branch_code,
            {_debt_group_aggregate_select_sql(alias="s", total_column="total_balance")}
        FROM customer_period_summary s
        {where}
        GROUP BY s.branch_code
    """
    return sql, params


def _debt_group_by_officer_base_sql(filters: CustomerFilters) -> tuple[str, list[object]]:
    clauses = ["1 = 1"]
    params: list[object] = []
    if filters.current_period:
        clauses.append("op.period = ?")
        params.append(filters.current_period)
    else:
        if filters.period_from:
            clauses.append("op.period >= ?")
            params.append(filters.period_from)
        if filters.period_to:
            clauses.append("op.period <= ?")
            params.append(filters.period_to)
    if filters.branch_code:
        clauses.append("(op.branch_code = ? OR s.branch_code = ?)")
        params.extend([filters.branch_code, filters.branch_code])
    if filters.customer_type:
        clauses.append("s.customer_type = ?")
        params.append(filters.customer_type)
    if filters.loan_term:
        if filters.loan_term == "SHORT_TERM":
            clauses.append("s.short_term_balance > 0")
        elif filters.loan_term == "MEDIUM_LONG_TERM":
            clauses.append("s.medium_long_term_balance > 0")
        elif filters.loan_term == "OTHER":
            clauses.append("s.other_balance > 0")
    if filters.search_text:
        pattern = f"%{clean_filter_text(filters.search_text)}%"
        clauses.append("(s.customer_code LIKE ? OR s.customer_name LIKE ?)")
        params.extend([pattern, pattern])
    if filters.officer:
        officer = clean_filter_text(filters.officer)
        clauses.append("(op.officer_code = ? OR op.officer_name = ? OR op.officer_name LIKE ?)")
        params.extend([officer, officer, f"%{officer}%"])
    if filters.debt_group:
        _append_debt_group_filter(clauses, params, filters.debt_group, alias="op", total_column="balance_managed")
    where = "WHERE " + " AND ".join(clauses)
    sql = f"""
        SELECT
            op.officer_code,
            op.officer_name,
            COALESCE(NULLIF(op.branch_code, ''), MIN(s.branch_code), '') AS branch_code,
            op.transaction_office,
            {_debt_group_aggregate_select_sql(alias="op", total_column="balance_managed", customer_count_expr="COUNT(DISTINCT {alias}.customer_code)")}
        FROM customer_officer_period op
        LEFT JOIN customer_period_summary s
            ON s.period = op.period AND s.customer_code = op.customer_code
        {where}
        GROUP BY op.officer_code, op.officer_name, op.branch_code, op.transaction_office
    """
    return sql, params


def _debt_group_by_office_base_sql(filters: CustomerFilters) -> tuple[str, list[object]]:
    clauses = ["1 = 1"]
    params: list[object] = []
    if filters.current_period:
        clauses.append("o.period = ?")
        params.append(filters.current_period)
    else:
        if filters.period_from:
            clauses.append("o.period >= ?")
            params.append(filters.period_from)
        if filters.period_to:
            clauses.append("o.period <= ?")
            params.append(filters.period_to)
    if filters.branch_code:
        clauses.append("o.branch_code = ?")
        params.append(filters.branch_code)
    if filters.customer_type:
        clauses.append("s.customer_type = ?")
        params.append(filters.customer_type)
    if filters.loan_term:
        if filters.loan_term == "SHORT_TERM":
            clauses.append("o.short_term_balance > 0")
        elif filters.loan_term == "MEDIUM_LONG_TERM":
            clauses.append("o.medium_long_term_balance > 0")
        elif filters.loan_term == "OTHER":
            clauses.append("o.other_balance > 0")
    if filters.search_text:
        pattern = f"%{clean_filter_text(filters.search_text)}%"
        clauses.append("(s.customer_code LIKE ? OR s.customer_name LIKE ?)")
        params.extend([pattern, pattern])
    if filters.officer:
        officer = clean_filter_text(filters.officer)
        clauses.append("(o.primary_officer_code = ? OR o.primary_officer_name = ? OR o.primary_officer_name LIKE ?)")
        params.extend([officer, officer, f"%{officer}%"])
    if filters.debt_group:
        _append_debt_group_filter(clauses, params, filters.debt_group, alias="o", total_column="total_balance")
    where = "WHERE " + " AND ".join(clauses)
    sql = f"""
        SELECT
            o.branch_code,
            o.office_code,
            o.office_name,
            o.office_type,
            {_debt_group_aggregate_select_sql(alias="o", total_column="total_balance", customer_count_expr="COUNT(DISTINCT {alias}.customer_code)")}
        FROM customer_office_period o
        LEFT JOIN customer_period_summary s
            ON s.period = o.period AND s.customer_code = o.customer_code
        {where}
        GROUP BY o.branch_code, o.office_code, o.office_name, o.office_type
    """
    return sql, params


def _debt_group_customer_base_sql(filters: CustomerFilters) -> tuple[str, list[object]]:
    where, params = _summary_where(filters)
    return _debt_group_customer_select_sql(where), params


def _debt_group_customer_select_sql(where: str) -> str:
    code_expr = _override_value_sql("s", "officer_code", fallback="s.primary_officer_code")
    name_expr = _override_value_sql("s", "officer_name", fallback="s.primary_officer_name", null_if_empty=True)
    return f"""
        SELECT
            s.period,
            s.customer_code,
            s.customer_name,
            s.customer_type,
            s.branch_code,
            {code_expr} AS effective_officer_code,
            {name_expr} AS effective_officer_name,
            s.total_balance,
            s.has_debt_group_data,
            s.worst_debt_group,
            s.debt_group_unknown_row_count,
            s.debt_group_1_balance,
            s.debt_group_2_balance,
            s.debt_group_3_balance,
            s.debt_group_4_balance,
            s.debt_group_5_balance,
            s.debt_group_unknown_balance,
            s.interest_rate_numerator,
            s.nim_before_numerator,
            s.nim_after_numerator,
            s.average_rate,
            s.nim_before,
            s.nim_after,
            CASE WHEN s.total_balance <> 0
                THEN s.debt_group_2_balance / s.total_balance * 100
                ELSE NULL
            END AS attention_ratio,
            CASE WHEN s.total_balance <> 0
                THEN (s.debt_group_3_balance + s.debt_group_4_balance + s.debt_group_5_balance) / s.total_balance * 100
                ELSE NULL
            END AS bad_debt_ratio,
            (s.debt_group_3_balance + s.debt_group_4_balance + s.debt_group_5_balance) AS bad_debt_balance
        FROM customer_period_summary s
        {where}
    """


def _debt_group_order_sql(sort_by: str, sort_desc: bool, *, default: str) -> str:
    allowed = {
        "period": "period",
        "branch_code": "branch_code",
        "branch_name": "branch_code",
        "office_code": "office_code",
        "officer_code": "officer_code",
        "officer_name": "officer_name",
        "customer_code": "customer_code",
        "customer_name": "customer_name",
        "total_balance": "total_balance",
        "debt_group_1_balance": "debt_group_1_balance",
        "debt_group_2_balance": "debt_group_2_balance",
        "debt_group_3_balance": "debt_group_3_balance",
        "debt_group_4_balance": "debt_group_4_balance",
        "debt_group_5_balance": "debt_group_5_balance",
        "bad_debt_balance": "bad_debt_balance",
        "attention_ratio": "attention_ratio",
        "bad_debt_ratio": "bad_debt_ratio",
        "average_rate": "average_rate",
        "nim_before": "nim_before",
        "nim_after": "nim_after",
        "customer_count": "customer_count",
        "attention_customer_count": "attention_customer_count",
        "bad_debt_customer_count": "bad_debt_customer_count",
    }
    column = allowed.get(str(sort_by or ""), allowed.get(default, "bad_debt_ratio"))
    direction = "DESC" if sort_desc else "ASC"
    return f"ORDER BY {column} {direction}, total_balance DESC"


def _finalize_debt_group_metrics(row: dict[str, object]) -> dict[str, object]:
    output = dict(row)
    total_balance = _number(output.get("total_balance"))
    attention_balance = _number(output.get("debt_group_2_balance"))
    bad_debt_balance = (
        _number(output.get("debt_group_3_balance"))
        + _number(output.get("debt_group_4_balance"))
        + _number(output.get("debt_group_5_balance"))
    )
    output["attention_balance"] = attention_balance
    output["bad_debt_balance"] = bad_debt_balance
    output["attention_ratio"] = _ratio(attention_balance * 100, total_balance) if total_balance else None
    output["bad_debt_ratio"] = _ratio(bad_debt_balance * 100, total_balance) if total_balance else None
    output["average_rate"] = _ratio(_number(output.get("interest_rate_numerator")), total_balance) if total_balance else None
    output["nim_before"] = _ratio(_number(output.get("nim_before_numerator")), total_balance) if total_balance else None
    output["nim_after"] = _ratio(_number(output.get("nim_after_numerator")), total_balance) if total_balance else None
    output["has_debt_group_data"] = bool(output.get("has_debt_group_data"))
    return output


def _finalize_debt_group_rows(rows: list[dict[str, object]]) -> None:
    for row in rows:
        finalized = _finalize_debt_group_metrics(row)
        row.update(finalized)
        worst = str(row.get("worst_debt_group") or "")
        row["worst_debt_group_label"] = DEBT_GROUP_CODE_LABELS.get(worst, worst)
        row["debt_group_unknown_balance"] = _number(row.get("debt_group_unknown_balance"))


def _enrich_debt_group_branch_rows(rows: list[dict[str, object]], unit_directory: UnitDirectoryService | None) -> None:
    for row in rows:
        branch_code = str(row.get("branch_code") or "")
        row["branch_name"] = _branch_display(branch_code, unit_directory) if branch_code else ""


def _trend_filters(filters: CustomerFilters, period_from: str = "", period_to: str = "") -> CustomerFilters:
    from_period = str(period_from or filters.period_from or "").strip()
    to_period = str(period_to or filters.period_to or "").strip()
    return replace(filters, current_period="", period_from=from_period, period_to=to_period)


def _trend_balance_row(
    row: dict[str, object],
    group_by: str,
    unit_directory: UnitDirectoryService | None = None,
) -> dict[str, object]:
    key = str(row.get("series_key") or "total").strip() or "total"
    if group_by == "branch":
        name = _branch_display(key, unit_directory)
    elif group_by == "customer_type":
        name = CUSTOMER_TYPE_LABELS.get(key, CUSTOMER_TYPE_LABELS["OTHER"])
    else:
        key = "total"
        name = "Tổng dư nợ"
    return {
        "period": str(row.get("period") or ""),
        "series_key": key,
        "series_name": name,
        "value": float(row.get("value") or 0),
    }


def _customer_list_base_sql(where: str) -> str:
    code_expr = _override_value_sql("s", "officer_code", fallback="s.primary_officer_code")
    name_expr = _override_value_sql("s", "officer_name", fallback="s.primary_officer_name", null_if_empty=True)
    has_override = _has_override_sql("s")
    return f"""
        SELECT
            s.period,
            s.customer_code,
            s.customer_name,
            s.customer_type,
            s.branch_code,
            {code_expr} AS effective_officer_code,
            {name_expr} AS effective_officer_name,
            s.primary_officer_code AS imported_officer_code,
            s.primary_officer_name AS imported_officer_name,
            s.officer_count,
            s.has_multiple_officers,
            s.total_balance,
            s.short_term_balance,
            s.medium_long_term_balance,
            s.other_balance,
            s.medium_long_ratio,
            s.interest_rate_numerator,
            s.nim_before_numerator,
            s.nim_after_numerator,
            s.average_rate,
            s.nim_before,
            s.nim_after,
            s.source_loan_count,
            CASE WHEN {has_override} THEN 1 ELSE 0 END AS has_override,
            CASE WHEN {has_override} THEN 'Có override' ELSE 'Không override' END AS override_status
        FROM customer_period_summary s
        {where}
    """


def _order_sql(
    sort_by: str,
    sort_desc: bool,
    allowed: dict[str, str],
    *,
    default: str,
) -> str:
    column = allowed.get(str(sort_by or ""), allowed[default])
    direction = "DESC" if sort_desc else "ASC"
    if "customer_code" not in allowed.values() or column == "customer_code":
        return f"ORDER BY {column} {direction}"
    tie_direction = "DESC" if sort_desc and column == "period" else "ASC"
    return f"ORDER BY {column} {direction}, customer_code COLLATE NOCASE {tie_direction}"


def _movement_base_sql(
    previous_period: str,
    current_period: str,
    *,
    resolve_officer: bool = True,
) -> tuple[str, list[object]]:
    customer_expr = "COALESCE(c.customer_code, p.customer_code)"
    period_expr = "COALESCE(c.period, p.period)"
    imported_code_expr = "COALESCE(c.primary_officer_code, p.primary_officer_code, '')"
    imported_name_expr = "COALESCE(c.primary_officer_name, p.primary_officer_name, '')"
    if resolve_officer:
        effective_code_expr = _override_value_for_expr_sql(
            customer_expr,
            period_expr,
            "officer_code",
            fallback=imported_code_expr,
        )
        effective_name_expr = _override_value_for_expr_sql(
            customer_expr,
            period_expr,
            "officer_name",
            fallback=imported_name_expr,
            null_if_empty=True,
        )
        has_override_expr = _has_override_for_expr_sql(customer_expr, period_expr)
    else:
        effective_code_expr = imported_code_expr
        effective_name_expr = imported_name_expr
        has_override_expr = "0"
    select_body = f"""
        COALESCE(p.period, '') AS previous_period,
        COALESCE(c.period, '') AS current_period,
        {period_expr} AS effective_period,
        {customer_expr} AS customer_code,
        COALESCE(c.customer_sequence, p.customer_sequence, '') AS customer_sequence,
        COALESCE(c.customer_name, p.customer_name, '') AS customer_name,
        COALESCE(c.customer_type, p.customer_type, '') AS customer_type,
        COALESCE(c.branch_code, p.branch_code, '') AS branch_code,
        {effective_code_expr} AS effective_officer_code,
        {effective_name_expr} AS effective_officer_name,
        {imported_code_expr} AS imported_officer_code,
        {imported_name_expr} AS imported_officer_name,
        COALESCE(p.total_balance, 0) AS previous_balance,
        COALESCE(c.total_balance, 0) AS current_balance,
        COALESCE(c.short_term_balance, p.short_term_balance, 0) AS short_term_balance,
        COALESCE(c.medium_long_term_balance, p.medium_long_term_balance, 0) AS medium_long_term_balance,
        COALESCE(c.other_balance, p.other_balance, 0) AS other_balance,
        CASE WHEN {has_override_expr} THEN 1 ELSE 0 END AS has_override
    """
    sql = f"""
        WITH previous AS (
            SELECT
                period,
                customer_code,
                customer_sequence,
                customer_name,
                customer_type,
                branch_code,
                primary_officer_code,
                primary_officer_name,
                total_balance,
                short_term_balance,
                medium_long_term_balance,
                other_balance
            FROM customer_period_summary
            WHERE period = ?
        ),
        current AS (
            SELECT
                period,
                customer_code,
                customer_sequence,
                customer_name,
                customer_type,
                branch_code,
                primary_officer_code,
                primary_officer_name,
                total_balance,
                short_term_balance,
                medium_long_term_balance,
                other_balance
            FROM customer_period_summary
            WHERE period = ?
        ),
        paired AS (
            SELECT {select_body}
            FROM current c
            LEFT JOIN previous p ON p.customer_code = c.customer_code
            UNION ALL
            SELECT {select_body}
            FROM previous p
            LEFT JOIN current c ON c.customer_code = p.customer_code
            WHERE c.customer_code IS NULL
        )
        SELECT
            paired.*,
            current_balance - previous_balance AS difference,
            CASE WHEN previous_balance <= 0
                THEN NULL
                ELSE (current_balance - previous_balance) / previous_balance * 100
            END AS growth_rate,
            CASE
                WHEN previous_balance <= 0 AND current_balance > 0 THEN '{MOVEMENT_STATUS_NEW}'
                WHEN previous_balance > 0 AND current_balance <= 0 THEN '{MOVEMENT_STATUS_PAID_OFF}'
                WHEN current_balance > previous_balance THEN '{MOVEMENT_STATUS_INCREASE}'
                WHEN current_balance < previous_balance THEN '{MOVEMENT_STATUS_DECREASE}'
                ELSE '{MOVEMENT_STATUS_UNCHANGED}'
            END AS movement_status
        FROM paired
    """
    return sql, [previous_period, current_period]


def _movement_where(
    filters: CustomerFilters,
    *,
    exclude: set[str] | None = None,
) -> tuple[str, list[object]]:
    exclude = exclude or set()
    clauses = ["1 = 1"]
    params: list[object] = []
    if filters.branch_code and "branch_code" not in exclude:
        clauses.append("q.branch_code = ?")
        params.append(filters.branch_code)
    if filters.customer_type and "customer_type" not in exclude:
        clauses.append("q.customer_type = ?")
        params.append(filters.customer_type)
    if filters.loan_term and "loan_term" not in exclude:
        if filters.loan_term == "SHORT_TERM":
            clauses.append("q.short_term_balance > 0")
        elif filters.loan_term == "MEDIUM_LONG_TERM":
            clauses.append("q.medium_long_term_balance > 0")
        elif filters.loan_term == "OTHER":
            clauses.append("q.other_balance > 0")
    if filters.search_text and "search_text" not in exclude:
        pattern = f"%{clean_filter_text(filters.search_text)}%"
        clauses.append("(q.customer_code LIKE ? OR q.customer_name LIKE ?)")
        params.extend([pattern, pattern])
    if filters.officer and "officer" not in exclude:
        officer = clean_filter_text(filters.officer)
        clauses.append(
            "(q.effective_officer_code = ? OR q.effective_officer_name = ? OR q.effective_officer_name LIKE ?)"
        )
        params.extend([officer, officer, f"%{officer}%"])
    if filters.override_status and "override_status" not in exclude:
        if filters.override_status == "override":
            clauses.append("q.has_override = 1")
        elif filters.override_status == "no_override":
            clauses.append("q.has_override = 0")
    if filters.movement_status and "movement_status" not in exclude:
        clauses.append("q.movement_status = ?")
        params.append(filters.movement_status)
    return "WHERE " + " AND ".join(clauses), params


def _movement_order_sql(sort_by: str, sort_desc: bool) -> str:
    return _order_sql(
        sort_by,
        sort_desc,
        {
            "customer_code": "customer_code",
            "customer_name": "customer_name COLLATE NOCASE",
            "customer_type": "customer_type",
            "branch_code": "branch_code",
            "branch_display": "branch_code",
            "effective_officer_name": "effective_officer_name COLLATE NOCASE",
            "previous_balance": "previous_balance",
            "current_balance": "current_balance",
            "difference": "difference",
            "growth_rate": "growth_rate",
            "movement_status": "movement_status",
        },
        default="difference",
    )


def _empty_movement_kpis() -> dict[str, object]:
    return {
        "new_customer_count": 0,
        "paid_off_customer_count": 0,
        "increased_customer_count": 0,
        "decreased_customer_count": 0,
        "unchanged_customer_count": 0,
        "new_customer_balance": 0,
        "paid_off_customer_balance": 0,
        "unchanged_customer_balance": 0,
        "total_increase": 0,
        "total_decrease": 0,
        "net_difference": 0,
    }


def _movement_requires_effective_officer_in_candidate(filters: CustomerFilters, sort_by: str = "") -> bool:
    return bool(
        (filters.officer or "").strip()
        or (filters.override_status or "").strip()
        or str(sort_by or "") in {"effective_officer_code", "effective_officer_name"}
    )


def _record_movement_stage(
    stats: list[dict[str, object]],
    stage: str,
    started: float,
    previous_period: str,
    current_period: str,
    *,
    generation: int | None = None,
    rows: int | None = None,
) -> None:
    elapsed_ms = (perf_counter() - started) * 1000
    item = {
        "generation": generation,
        "stage": stage,
        "previous_period": previous_period,
        "current_period": current_period,
        "elapsed_ms": elapsed_ms,
    }
    if rows is not None:
        item["rows"] = rows
    stats.append(item)
    LOGGER.info(
        "Customer movement stage: generation=%s stage=%s previous_period=%s current_period=%s elapsed_ms=%.1f rows=%s",
        generation,
        stage,
        previous_period,
        current_period,
        elapsed_ms,
        "" if rows is None else rows,
    )


def _materialize_movement_candidates(
    database: sqlite3.Connection,
    previous_period: str,
    current_period: str,
    filters: CustomerFilters,
    stats: list[dict[str, object]],
    *,
    generation: int | None = None,
    sort_by: str = "",
) -> None:
    period_started = perf_counter()
    previous_count = int(
        database.execute(
            "SELECT COUNT(*) FROM customer_period_summary WHERE period = ?",
            (previous_period,),
        ).fetchone()[0]
        or 0
    )
    _record_movement_stage(
        stats,
        "previous_period_scope",
        period_started,
        previous_period,
        current_period,
        generation=generation,
        rows=previous_count,
    )
    current_started = perf_counter()
    current_count = int(
        database.execute(
            "SELECT COUNT(*) FROM customer_period_summary WHERE period = ?",
            (current_period,),
        ).fetchone()[0]
        or 0
    )
    _record_movement_stage(
        stats,
        "current_period_scope",
        current_started,
        previous_period,
        current_period,
        generation=generation,
        rows=current_count,
    )
    join_started = perf_counter()
    base_sql, params = _movement_base_sql(
        previous_period,
        current_period,
        resolve_officer=_movement_requires_effective_officer_in_candidate(filters, sort_by),
    )
    database.execute(f"DROP TABLE IF EXISTS {CUSTOMER_MOVEMENT_TEMP_TABLE}")
    database.execute(
        f"""
        CREATE TEMP TABLE {CUSTOMER_MOVEMENT_TEMP_TABLE} AS
        SELECT *
        FROM ({base_sql}) q
        """,
        params,
    )
    candidate_count = int(
        database.execute(f"SELECT COUNT(*) FROM {CUSTOMER_MOVEMENT_TEMP_TABLE}").fetchone()[0]
        or 0
    )
    _record_movement_stage(
        stats,
        "candidate_join",
        join_started,
        previous_period,
        current_period,
        generation=generation,
        rows=candidate_count,
    )
    index_started = perf_counter()
    database.execute(f"CREATE INDEX IF NOT EXISTS idx_temp_movement_customer ON {CUSTOMER_MOVEMENT_TEMP_TABLE}(customer_code)")
    database.execute(f"CREATE INDEX IF NOT EXISTS idx_temp_movement_branch ON {CUSTOMER_MOVEMENT_TEMP_TABLE}(branch_code)")
    database.execute(f"CREATE INDEX IF NOT EXISTS idx_temp_movement_type ON {CUSTOMER_MOVEMENT_TEMP_TABLE}(customer_type)")
    database.execute(f"CREATE INDEX IF NOT EXISTS idx_temp_movement_status ON {CUSTOMER_MOVEMENT_TEMP_TABLE}(movement_status)")
    database.execute(f"CREATE INDEX IF NOT EXISTS idx_temp_movement_difference ON {CUSTOMER_MOVEMENT_TEMP_TABLE}(difference)")
    if _movement_requires_effective_officer_in_candidate(filters, sort_by):
        database.execute(
            f"CREATE INDEX IF NOT EXISTS idx_temp_movement_effective_officer ON {CUSTOMER_MOVEMENT_TEMP_TABLE}(effective_officer_code, effective_officer_name)"
        )
        database.execute(f"CREATE INDEX IF NOT EXISTS idx_temp_movement_override ON {CUSTOMER_MOVEMENT_TEMP_TABLE}(has_override)")
    _record_movement_stage(
        stats,
        "temp_indexes",
        index_started,
        previous_period,
        current_period,
        generation=generation,
        rows=candidate_count,
    )


def _movement_kpis_from_temp(
    database: sqlite3.Connection,
    filters: CustomerFilters,
    stats: list[dict[str, object]],
    previous_period: str,
    current_period: str,
    *,
    generation: int | None = None,
) -> dict[str, object]:
    started = perf_counter()
    where, params = _movement_where(filters, exclude={"movement_status"})
    row = database.execute(
        f"""
        SELECT
            SUM(CASE WHEN movement_status = ? THEN 1 ELSE 0 END) AS new_customer_count,
            SUM(CASE WHEN movement_status = ? THEN 1 ELSE 0 END) AS paid_off_customer_count,
            SUM(CASE WHEN movement_status = ? THEN 1 ELSE 0 END) AS increased_customer_count,
            SUM(CASE WHEN movement_status = ? THEN 1 ELSE 0 END) AS decreased_customer_count,
            SUM(CASE WHEN movement_status = ? THEN 1 ELSE 0 END) AS unchanged_customer_count,
            SUM(CASE WHEN movement_status = ? THEN current_balance ELSE 0 END) AS new_customer_balance,
            SUM(CASE WHEN movement_status = ? THEN previous_balance ELSE 0 END) AS paid_off_customer_balance,
            SUM(CASE WHEN movement_status = ? THEN current_balance ELSE 0 END) AS unchanged_customer_balance,
            SUM(CASE WHEN difference > 0 THEN difference ELSE 0 END) AS total_increase,
            SUM(CASE WHEN difference < 0 THEN ABS(difference) ELSE 0 END) AS total_decrease,
            SUM(difference) AS net_difference
        FROM {CUSTOMER_MOVEMENT_TEMP_TABLE} q
        {where}
        """,
        (
            MOVEMENT_STATUS_NEW,
            MOVEMENT_STATUS_PAID_OFF,
            MOVEMENT_STATUS_INCREASE,
            MOVEMENT_STATUS_DECREASE,
            MOVEMENT_STATUS_UNCHANGED,
            MOVEMENT_STATUS_NEW,
            MOVEMENT_STATUS_PAID_OFF,
            MOVEMENT_STATUS_UNCHANGED,
            *params,
        ),
    ).fetchone()
    output = dict(row) if row is not None else _empty_movement_kpis()
    for key, value in _empty_movement_kpis().items():
        if output.get(key) is None:
            output[key] = value
        else:
            output.setdefault(key, value)
    _record_movement_stage(
        stats,
        "kpi",
        started,
        previous_period,
        current_period,
        generation=generation,
    )
    return output


def _movement_count_from_temp(
    database: sqlite3.Connection,
    filters: CustomerFilters,
    stats: list[dict[str, object]],
    previous_period: str,
    current_period: str,
    *,
    generation: int | None = None,
) -> int:
    started = perf_counter()
    where, params = _movement_where(filters)
    total_rows = int(
        database.execute(
            f"SELECT COUNT(*) FROM {CUSTOMER_MOVEMENT_TEMP_TABLE} q {where}",
            params,
        ).fetchone()[0]
        or 0
    )
    _record_movement_stage(
        stats,
        "count",
        started,
        previous_period,
        current_period,
        generation=generation,
        rows=total_rows,
    )
    return total_rows


def _movement_page_from_temp(
    database: sqlite3.Connection,
    filters: CustomerFilters,
    *,
    page: int,
    page_size: int,
    sort_by: str,
    sort_desc: bool,
    stats: list[dict[str, object]],
    previous_period: str,
    current_period: str,
    generation: int | None = None,
) -> list[dict[str, object]]:
    started = perf_counter()
    offset = (max(1, int(page or 1)) - 1) * max(1, int(page_size or 100))
    where, params = _movement_where(filters)
    order_sql = _movement_order_sql(sort_by, sort_desc)
    rows = [
        dict(row)
        for row in database.execute(
            f"""
            SELECT *
            FROM {CUSTOMER_MOVEMENT_TEMP_TABLE} q
            {where}
            {order_sql}
            LIMIT ? OFFSET ?
            """,
            (*params, page_size, offset),
        ).fetchall()
    ]
    _record_movement_stage(
        stats,
        "page",
        started,
        previous_period,
        current_period,
        generation=generation,
        rows=len(rows),
    )
    return rows


def _resolve_movement_officers_for_rows(
    database: sqlite3.Connection,
    rows: list[dict[str, object]],
    stats: list[dict[str, object]],
    previous_period: str,
    current_period: str,
    *,
    generation: int | None = None,
) -> None:
    started = perf_counter()
    if not rows:
        _record_movement_stage(
            stats,
            "officer_resolution",
            started,
            previous_period,
            current_period,
            generation=generation,
            rows=0,
        )
        return
    customer_codes = sorted({str(row.get("customer_code") or "") for row in rows if str(row.get("customer_code") or "")})
    periods = [str(row.get("effective_period") or current_period or previous_period) for row in rows]
    max_period = max(periods) if periods else current_period
    min_period = min(periods) if periods else previous_period
    placeholders = ", ".join("?" for _code in customer_codes)
    override_rows: list[dict[str, object]] = []
    if customer_codes:
        override_rows = [
            dict(row)
            for row in database.execute(
                f"""
                SELECT
                    id,
                    customer_code,
                    effective_from_period,
                    effective_to_period,
                    officer_code,
                    officer_name,
                    reason,
                    is_active,
                    updated_at
                FROM customer_officer_override
                WHERE is_active = 1
                  AND customer_code IN ({placeholders})
                  AND effective_from_period <= ?
                  AND (effective_to_period = '' OR effective_to_period >= ?)
                ORDER BY customer_code COLLATE NOCASE ASC, updated_at DESC, id DESC
                """,
                (*customer_codes, max_period, min_period),
            ).fetchall()
        ]
    for row in rows:
        customer_code = str(row.get("customer_code") or "")
        period = str(row.get("effective_period") or current_period or previous_period)
        override = _effective_override_from_rows(override_rows, customer_code, period)
        if override is not None:
            row["effective_officer_code"] = str(override.get("officer_code") or "")
            row["effective_officer_name"] = str(override.get("officer_name") or "") or str(row.get("imported_officer_name") or "")
            row["has_override"] = 1
        else:
            row["effective_officer_code"] = str(row.get("imported_officer_code") or "")
            row["effective_officer_name"] = str(row.get("imported_officer_name") or "")
            row["has_override"] = 0
        row["override_status"] = "Có override" if int(row.get("has_override") or 0) else "Không override"
    _record_movement_stage(
        stats,
        "officer_resolution",
        started,
        previous_period,
        current_period,
        generation=generation,
        rows=len(rows),
    )


def _enrich_movement_unit_rows(
    rows: list[dict[str, object]],
    unit_directory: UnitDirectoryService | None,
    stats: list[dict[str, object]] | None = None,
    previous_period: str = "",
    current_period: str = "",
    *,
    generation: int | None = None,
) -> None:
    started = perf_counter()
    for row in rows:
        branch_code = str(row.get("branch_code") or "")
        row["branch_display"] = _branch_display(branch_code, unit_directory)
    if stats is not None:
        _record_movement_stage(
            stats,
            "unit_enrich",
            started,
            previous_period,
            current_period,
            generation=generation,
            rows=len(rows),
        )
    LOGGER.debug("Customer movement unit enrich rows=%s elapsed_ms=%.1f", len(rows), (perf_counter() - started) * 1000)


def _effective_override_from_rows(
    rows: list[dict[str, object]],
    customer_code: str,
    period: str,
) -> dict[str, object] | None:
    candidates = [
        row
        for row in rows
        if str(row.get("customer_code") or "") == customer_code
        and int(row.get("is_active") or 0) == 1
        and str(row.get("effective_from_period") or "") <= period
        and (not str(row.get("effective_to_period") or "") or str(row.get("effective_to_period") or "") >= period)
    ]
    if not candidates:
        return None
    candidates.sort(key=lambda row: (str(row.get("updated_at") or ""), int(row.get("id") or 0)), reverse=True)
    return candidates[0]


def _override_scope_text(row: dict[str, object] | None) -> str:
    if not row:
        return ""
    start = str(row.get("effective_from_period") or "")
    end = str(row.get("effective_to_period") or "")
    if not end:
        return f"Từ {start} trở đi"
    if start == end:
        return start
    return f"{start} đến {end}"


def _strip_vietnamese_accents(value: object) -> str:
    text = "" if value is None else str(value)
    decomposed = unicodedata.normalize("NFD", text)
    return "".join(char for char in decomposed if unicodedata.category(char) != "Mn").replace("đ", "d").replace("Đ", "D")


def _record_plan(
    database: sqlite3.Connection,
    plans: dict[str, list[str]],
    name: str,
    sql: str,
    params: list[object] | tuple[object, ...],
) -> None:
    rows = database.execute(f"EXPLAIN QUERY PLAN {sql}", tuple(params)).fetchall()
    plans[name] = [str(row["detail"] if "detail" in row.keys() else row[-1]) for row in rows]
