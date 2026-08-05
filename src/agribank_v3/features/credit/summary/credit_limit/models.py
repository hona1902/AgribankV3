from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime
from enum import StrEnum
from pathlib import Path


CREDIT_LIMIT_BATCH_SCHEMA_VERSION = "2"
DATA_SHEET_NAME = "DuLieuHanMuc"
META_SHEET_NAME = "ThongTinBatch"
STORAGE_FOLDER_NAME = "HMHETHAN"


@dataclass(frozen=True)
class CreditLimitBatchMetadata:
    batch_id: str
    batch_name: str
    file_path: Path
    file_name: str
    source_file_name: str
    source_file_sha256: str
    source_file_size: int
    imported_at: datetime | None
    imported_by: str
    app_version: str
    source_row_count: int
    accepted_row_count: int
    rejected_row_count: int
    warning_count: int
    reference_date_at_import: date | None
    minimum_limit_at_import: float
    warning_days_at_import: int
    expired_count_at_import: int
    expiring_count_at_import: int
    status: str
    notes: str
    period: str = ""
    branch_code: str = ""
    office_code: str = ""
    is_active: bool = True
    replaced_batch_id: str = ""
    previous_version: str = ""
    period_missing: bool = False
    schema_version: str = CREDIT_LIMIT_BATCH_SCHEMA_VERSION
    file_size: int = 0
    modified_at: datetime | None = None


class CreditLimitBatchLookupState(StrEnum):
    NOT_FOUND = "NOT_FOUND"
    FOUND_VALID = "FOUND_VALID"
    FOUND_INVALID = "FOUND_INVALID"


@dataclass(frozen=True)
class CreditLimitBatchLookup:
    state: CreditLimitBatchLookupState
    metadata: CreditLimitBatchMetadata | None = None
    invalid_file_path: Path | None = None
    error_message: str = ""


@dataclass(frozen=True)
class CreditLimitStorageStatus:
    storage_path: Path
    valid_files: int
    invalid_files: int
    temporary_files: int
    total_size_bytes: int
    last_checked_at: datetime
    batches: tuple[CreditLimitBatchMetadata, ...]
    invalid_entries: tuple[tuple[str, str], ...]
