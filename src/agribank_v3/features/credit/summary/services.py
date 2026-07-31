from __future__ import annotations

from collections.abc import Callable
from contextlib import closing, contextmanager
from dataclasses import dataclass, replace
from datetime import date, datetime
import csv
import getpass
import hashlib
from io import StringIO
import os
from pathlib import Path
import re
import sqlite3
import tempfile
from time import perf_counter
from typing import Iterable

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill

from agribank_v3.features.credit.summary.models import (
    CreditLimitRow,
    ImportResult,
    LoanSnapshotRow,
    NIM_DN_CONFIG,
    NIM_NV_CONFIG,
    NormalizedLoanRow,
    OfficerHistory,
    OfficerHistoryPoint,
    NimConfig,
    NimRow,
    SummaryDataType,
    SummaryError,
)
from agribank_v3.features.credit.summary.reports import (
    export_credit_limit_vba,
    export_loan_compare_vba,
    export_nim_vba,
)
from agribank_v3.features.credit.summary.repository import SummaryRepository
from agribank_v3.features.credit.summary.customer.repository import CustomerRepository
from agribank_v3.features.credit.summary.customer.services import (
    CustomerAggregationService,
    build_customer_code,
    build_office_code,
    classify_office_type,
    map_customer_type_code,
    normalize_trctcd,
    normalize_customer_sequence,
    split_officer as split_customer_officer,
)
from agribank_v3.features.settings.unit_directory.service import (
    UnitDirectoryService,
    get_unit_directory_service,
)


ProgressCallback = Callable[[str], None]


NIM_COMMON_HEADERS = (
    "BRCD",
    "FTP",
    "INTRT",
    "MUCFTPDC",
    "LDRBAL",
)

CUSTOMER_FTPLN_REQUIRED_HEADERS = (
    "BRCD",
    "CUSTSEQ",
    "CUSTNM",
    "CUSTTP",
    "FTPCD",
    "LDRBAL",
    "FTP",
    "INTRT",
    "MUCFTPDC",
    "CBTD",
)

CUSTTP_TOTAL = "[Tất cả KH]"
BRANCH_TOTAL = "[Tổng Chi Nhánh]"
CREDIT_LIMIT_EXPIRED_NOTE = "Hợp đồng hạn mức tín dụng ã quá hạn đến thời đểm hiện tại"
CREDIT_LIMIT_SOON_NOTE_TEMPLATE = "Hợp đồng tín dụng dến hạn trong vòng {warn_days} ngày tới theo thời đểm hiện tại"
OFFICER_HISTORY_HEADERS = (
    "Kỳ",
    "Dư nợ",
    "Lãi suất bình quân",
    "NIM trước ĐC",
    "NIM sau ĐC",
)


@dataclass(frozen=True, slots=True)
class ParsedNimFile:
    rows: list[NimRow]
    period: str
    source_hash: str
    customer_rows: list[NormalizedLoanRow]
    customer_missing_headers: tuple[str, ...]
    customer_source_total_balance: float
    customer_invalid_row_count: int
    customer_warning_count: int
    customer_warnings: tuple[str, ...]
    unit_warnings: tuple[str, ...]


def import_nim_folder(
    repository: SummaryRepository,
    folder_path: Path,
    config: NimConfig,
    *,
    credit_card_rate: float = 0.0,
    export_path: Path | None = None,
    imported_by: str | None = None,
    progress: ProgressCallback | None = None,
    replace_existing_periods: bool = False,
) -> ImportResult:
    folder_path = Path(folder_path)
    if not folder_path.is_dir():
        raise SummaryError("Thư mục import NIM không tồn tại.")
    files = sorted(
        path for path in folder_path.iterdir()
        if path.is_file()
        and path.suffix.casefold() == ".csv"
        and config.file_pattern_token in path.name.casefold()
    )
    if not files:
        raise SummaryError(f"Không tìm thấy file CSV phù hợp {config.file_pattern_token} trong thư mục đã chọn.")

    parsed_files: list[tuple[Path, ParsedNimFile, int]] = []
    user = imported_by or _current_user()
    customer_aggregation = (
        CustomerAggregationService(folder_path)
        if config.data_type == SummaryDataType.NIM_DN
        else None
    )
    unit_directory = get_unit_directory_service(repository.main_database_path)
    customer_missing_messages: list[str] = []
    unit_warning_messages: list[str] = []
    parse_started = perf_counter()
    for file_index, file_path in enumerate(files, start=1):
        if progress:
            progress(f"Đang đọc file {file_index}/{len(files)}: {file_path.name}")
        started = perf_counter()
        parsed = _parse_nim_file(
            file_path,
            config,
            credit_card_rate=credit_card_rate,
            unit_directory=unit_directory,
        )
        duration_ms = int((perf_counter() - started) * 1000)
        parsed_files.append((file_path, parsed, duration_ms))
        unit_warning_messages.extend(parsed.unit_warnings)
        if customer_aggregation is not None:
            if parsed.customer_missing_headers:
                customer_missing_messages.append(
                    f"{file_path.name} thiếu cột Customer.db: {', '.join(parsed.customer_missing_headers)}"
                )
            else:
                customer_aggregation.add_file(
                    parsed.customer_rows,
                    period=parsed.period,
                    file_path=file_path,
                    file_hash=parsed.source_hash,
                    source_row_count=len(parsed.rows),
                    source_total_balance=parsed.customer_source_total_balance,
                    invalid_row_count=parsed.customer_invalid_row_count,
                    warning_count=parsed.customer_warning_count,
                    warnings=parsed.customer_warnings,
                )
        if progress:
            customer_count = len({row.customer_code for row in parsed.customer_rows})
            progress(
                f"Đã đọc {len(parsed.rows):,} dòng nguồn, "
                f"{customer_count:,} khách hàng từ {file_path.name}"
            )

    customer_result = None
    if customer_aggregation is not None and not customer_missing_messages:
        customer_result = customer_aggregation.build_result()
        if customer_result.source_row_count <= 0:
            customer_result = None

    customer_repository = CustomerRepository(repository.main_database_path) if customer_result is not None else None
    snapshot_paths = [repository.database_path]
    if customer_repository is not None:
        snapshot_paths.append(customer_repository.database_path)

    total_rows = 0
    first_batch_id = 0
    export_rows_buffer: list[NimRow] = []
    total_summary_rows = 0
    customer_run_id = 0
    write_started = perf_counter()
    with _sqlite_snapshots(snapshot_paths):
        if replace_existing_periods and config.data_type == SummaryDataType.NIM_DN:
            for period in sorted({parsed.period for _file_path, parsed, _duration_ms in parsed_files}):
                try:
                    repository.delete_nim_period(config.data_type, period, created_by=user)
                except SummaryError as exc:
                    if "Không có dữ liệu" not in str(exc):
                        raise
        for file_path, parsed, duration_ms in parsed_files:
            if progress:
                progress(f"Đang ghi CreditSummary.db: {file_path.name}")
            batch_id = repository.create_batch(
                config.data_type,
                period=parsed.period,
                source_path=file_path,
                imported_by=user,
                row_count=len(parsed.rows),
                duration_ms=duration_ms,
                message=f"Import {file_path.name}",
                source_hash=parsed.source_hash,
            )
            rows = [replace(row, batch_id=batch_id) for row in parsed.rows]
            summary_rows = aggregate_nim_rows(rows)
            repository.save_nim_rows(summary_rows)
            export_rows_buffer.extend(rows)
            if not first_batch_id:
                first_batch_id = batch_id
            total_rows += len(rows)
            total_summary_rows += len(summary_rows)
            if progress:
                progress(f"Đã lưu {len(summary_rows):,} dòng tổng hợp NIM từ {len(rows):,} dòng nguồn")
        if customer_repository is not None and customer_result is not None:
            if progress:
                progress("Đang ghi Customer.db")
            customer_run_id = customer_repository.save_aggregation(
                customer_result,
                replace_periods=replace_existing_periods,
                created_by=user,
                duration_ms=int((perf_counter() - write_started) * 1000),
            )
    output_path = None
    if export_path is not None and export_rows_buffer:
        detail_rows, branch_rows = build_nim_vba_tables(export_rows_buffer, config.data_type)
        output_path = export_nim_vba(detail_rows, branch_rows, Path(export_path), data_type=config.data_type)
    customer_message = ""
    if customer_result is not None:
        unknown_count = len(customer_result.unknown_ftp_codes)
        customer_message = (
            f" Customer.db: run {customer_run_id}, {customer_result.customer_count:,} khách hàng, "
            f"tổng dư nợ {customer_result.total_balance:,.0f}, "
            f"ngắn hạn {customer_result.short_term_balance:,.0f}, "
            f"trung/dài hạn {customer_result.medium_long_term_balance:,.0f}, "
            f"chưa phân loại {customer_result.other_balance:,.0f}, "
            f"KH nhiều cán bộ {customer_result.multiple_officer_customer_count:,}, "
            f"FTPCD lạ {unknown_count:,}, lỗi dòng {customer_result.invalid_row_count:,}."
        )
    elif customer_missing_messages:
        customer_message = " Customer.db chưa cập nhật vì " + "; ".join(customer_missing_messages[:3]) + "."
    if progress:
        progress("Hoàn thành")
    unit_message = ""
    if unit_warning_messages:
        unique_warnings = list(dict.fromkeys(unit_warning_messages))
        unit_message = " Cài đặt đơn vị: " + "; ".join(unique_warnings[:5])
        if len(unique_warnings) > 5:
            unit_message += f"; còn {len(unique_warnings) - 5} cảnh báo khác"
        unit_message += "."
    return ImportResult(
        batch_id=first_batch_id,
        row_count=total_rows,
        message=(
            f"Đã import {len(files)} file, {total_rows:,} dòng nguồn, "
            f"{total_summary_rows:,} dòng tổng hợp trong {int((perf_counter() - parse_started) * 1000):,} ms."
            f"{customer_message}"
            f"{unit_message}"
        ),
        output_path=output_path,
    )


def import_nim_dn(
    repository: SummaryRepository,
    folder_path: Path,
    *,
    credit_card_rate: float = 0.0,
    export_path: Path | None = None,
    imported_by: str | None = None,
    progress: ProgressCallback | None = None,
    replace_existing_periods: bool = False,
) -> ImportResult:
    return import_nim_folder(
        repository,
        folder_path,
        NIM_DN_CONFIG,
        credit_card_rate=credit_card_rate,
        export_path=export_path,
        imported_by=imported_by,
        progress=progress,
        replace_existing_periods=replace_existing_periods,
    )


def import_nim_nv(
    repository: SummaryRepository,
    folder_path: Path,
    *,
    credit_card_rate: float = 0.0,
    export_path: Path | None = None,
    imported_by: str | None = None,
    progress: ProgressCallback | None = None,
    replace_existing_periods: bool = False,
) -> ImportResult:
    _ = replace_existing_periods
    return import_nim_folder(
        repository,
        folder_path,
        NIM_NV_CONFIG,
        credit_card_rate=credit_card_rate,
        export_path=export_path,
        imported_by=imported_by,
        progress=progress,
    )


def build_officer_history(
    repository: SummaryRepository,
    data_type: SummaryDataType,
    *,
    officer: str,
    branch: str = "",
    transaction_office: str = "",
    customer_type: str = "",
) -> OfficerHistory:
    rows = repository.get_officer_history(
        data_type,
        officer=officer,
        branch=branch,
        transaction_office=transaction_office,
        customer_type=customer_type,
    )
    points = tuple(
        OfficerHistoryPoint(
            period=str(row.get("period") or ""),
            balance=float(row.get("balance") or 0),
            average_rate=float(row.get("average_rate") or 0),
            nim_before=float(row.get("nim_before") or 0),
            nim_after=float(row.get("nim_after") or 0),
        )
        for row in rows
    )
    current = points[-1] if points else None
    return OfficerHistory(
        data_type=data_type,
        officer=officer,
        branch=branch,
        transaction_office=transaction_office,
        customer_type=customer_type,
        current_period=current.period if current else "",
        current_balance=current.balance if current else 0.0,
        current_average_rate=current.average_rate if current else 0.0,
        current_nim_before=current.nim_before if current else 0.0,
        current_nim_after=current.nim_after if current else 0.0,
        points=points,
    )


def export_officer_history_excel(history: OfficerHistory, destination: Path) -> Path:
    destination = Path(destination)
    destination.parent.mkdir(parents=True, exist_ok=True)
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "LichSu_NIM_CBTD"
    worksheet["A1"] = "Lịch sử NIM CBTD"
    worksheet["A1"].font = Font(bold=True, size=14)
    worksheet["A3"] = "CBTD"
    worksheet["B3"] = history.officer
    worksheet["A4"] = "Chi nhánh"
    worksheet["B4"] = history.branch
    worksheet["A5"] = "Phòng GD"
    worksheet["B5"] = history.transaction_office
    worksheet["A6"] = "Loại khách hàng"
    worksheet["B6"] = history.customer_type or "Tất cả"
    for column_index, value in enumerate(OFFICER_HISTORY_HEADERS, start=1):
        cell = worksheet.cell(8, column_index, value)
        cell.font = Font(bold=True)
        cell.fill = PatternFill("solid", fgColor="D9EAF7")
    for row_index, point in enumerate(history.points, start=9):
        worksheet.cell(row_index, 1, point.period)
        worksheet.cell(row_index, 2, point.balance)
        worksheet.cell(row_index, 3, point.average_rate)
        worksheet.cell(row_index, 4, point.nim_before)
        worksheet.cell(row_index, 5, point.nim_after)
    for cell in worksheet["B"]:
        if cell.row >= 9:
            cell.number_format = "#,##0"
    for column_letter in ("C", "D", "E"):
        for cell in worksheet[column_letter]:
            if cell.row >= 9:
                cell.number_format = '0.00"%"'
    widths = {
        "A": 14,
        "B": 18,
        "C": 20,
        "D": 16,
        "E": 16,
    }
    for column_letter, width in widths.items():
        worksheet.column_dimensions[column_letter].width = width
    workbook.save(destination)
    return destination


def build_nim_vba_tables(
    rows: Iterable[NimRow],
    data_type: SummaryDataType,
) -> tuple[list[list[object]], list[list[object]]]:
    detail: dict[str, list[float]] = {}
    branch: dict[str, list[float]] = {}
    branch_has_pgd: dict[str, bool] = {}

    for row in rows:
        detail_key = "|".join(
            [
                row.period,
                row.branch_name,
                row.transaction_office,
                row.customer_type,
                row.officer,
            ]
        )
        detail_total_key = "|".join(
            [
                row.period,
                row.branch_name,
                row.transaction_office,
                CUSTTP_TOTAL,
                row.officer,
            ]
        )
        old_trctcd = row.trctcd.replace("'", "")
        if old_trctcd != "00":
            branch_has_pgd[row.branch_name] = True
        elif row.branch_name not in branch_has_pgd:
            branch_has_pgd[row.branch_name] = False

        _add_nim_values(detail, detail_key, row, data_type)
        _add_nim_values(detail, detail_total_key, row, data_type)

        branch_key = "|".join([row.period, row.branch_name, BRANCH_TOTAL, row.customer_type])
        branch_total_key = "|".join([row.period, row.branch_name, BRANCH_TOTAL, CUSTTP_TOTAL])
        branch_pgd_key = "|".join([row.period, row.branch_name, row.transaction_office, row.customer_type])
        branch_pgd_total_key = "|".join([row.period, row.branch_name, row.transaction_office, CUSTTP_TOTAL])
        _add_nim_values(branch, branch_key, row, data_type)
        _add_nim_values(branch, branch_total_key, row, data_type)
        _add_nim_values(branch, branch_pgd_key, row, data_type)
        _add_nim_values(branch, branch_pgd_total_key, row, data_type)

    detail_rows = [_nim_detail_output(key, values, data_type) for key, values in detail.items()]
    branch_rows: list[list[object]] = []
    for key, values in branch.items():
        period, branch_name, transaction_office, _customer_type = key.split("|")
        if transaction_office == "Hội sở" and branch_has_pgd.get(branch_name) is False:
            continue
        branch_rows.append(_nim_branch_output(key, values, data_type))
    return detail_rows, branch_rows


def aggregate_nim_rows(rows: Iterable[NimRow]) -> list[NimRow]:
    grouped: dict[tuple[object, ...], dict[str, object]] = {}
    for row in rows:
        officer_code, officer_name = _split_officer(row.officer)
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
                "officer_code": officer_code,
                "officer_name": officer_name,
                "balance": 0.0,
                "average_rate_numerator": 0.0,
                "numerator_before": 0.0,
                "numerator_after": 0.0,
                "source_row_count": 0,
            }
            grouped[key] = item
        item["balance"] = float(item["balance"]) + float(row.balance or 0)
        item["average_rate_numerator"] = float(item["average_rate_numerator"]) + float(row.average_rate_numerator or 0)
        item["numerator_before"] = float(item["numerator_before"]) + float(row.numerator_before or 0)
        item["numerator_after"] = float(item["numerator_after"]) + float(row.numerator_after or 0)
        item["source_row_count"] = int(item["source_row_count"]) + int(row.source_row_count or 1)

    summary_rows: list[NimRow] = []
    for item in grouped.values():
        base = item["row"]
        if not isinstance(base, NimRow):
            continue
        balance = float(item["balance"] or 0)
        average_rate_numerator = float(item["average_rate_numerator"] or 0)
        summary_rows.append(
            replace(
                base,
                officer=_format_officer(str(item["officer_code"] or ""), str(item["officer_name"] or "")),
                balance=balance,
                interest_rate=average_rate_numerator / balance if balance else 0.0,
                ftp_rate=0.0,
                adjustment_rate=0.0,
                numerator_before=float(item["numerator_before"] or 0),
                numerator_after=float(item["numerator_after"] or 0),
                average_rate_numerator=average_rate_numerator,
                source_row_count=int(item["source_row_count"] or 0),
            )
        )
    return summary_rows


def compare_loan_balances(
    repository: SummaryRepository,
    previous_file: Path,
    current_file: Path,
    *,
    export_path: Path | None = None,
    imported_by: str | None = None,
    progress: ProgressCallback | None = None,
) -> ImportResult:
    previous_file = Path(previous_file)
    current_file = Path(current_file)
    if progress:
        progress(f"Đang đọc kỳ trước: {previous_file.name}")
    previous, info = _load_loan_file(previous_file)
    if progress:
        progress(f"Đang đọc kỳ này: {current_file.name}")
    current, info_current = _load_loan_file(current_file)
    for key, value in info_current.items():
        info.setdefault(key, value)
    rows = build_loan_compare_rows(previous, current, info)
    started = perf_counter()
    period = f"{previous_file.stem} -> {current_file.stem}"
    batch_id = repository.create_batch(
        SummaryDataType.LOAN_COMPARE,
        period=period,
        source_path=current_file,
        imported_by=imported_by or _current_user(),
        row_count=len(rows),
        duration_ms=int((perf_counter() - started) * 1000),
        message=f"So sánh {previous_file.name} và {current_file.name}",
        source_hash=_combined_file_hash((previous_file, current_file)),
    )
    repository.save_loan_compare_rows(batch_id, rows)
    output_path = export_loan_compare_vba(rows, Path(export_path)) if export_path is not None and rows else None
    if progress:
        progress(f"Đã lưu {len(rows):,} dòng đối chiếu")
    return ImportResult(
        batch_id=batch_id,
        row_count=len(rows),
        message=f"Đã đối chiếu {len(rows):,} khách hàng.",
        output_path=output_path,
    )


def compare_loan_folder(
    repository: SummaryRepository,
    folder_path: Path,
    *,
    export_path: Path | None = None,
    imported_by: str | None = None,
    progress: ProgressCallback | None = None,
) -> ImportResult:
    files = sorted(
        path for path in Path(folder_path).iterdir()
        if path.is_file() and path.suffix.casefold() in {".csv", ".txt", ".xlsx"}
    )
    if len(files) < 2:
        raise SummaryError("Không tìm thấy đủ 2 file dữ liệu trong thư mục đã chọn.")
    return compare_loan_balances(
        repository,
        files[0],
        files[1],
        export_path=export_path,
        imported_by=imported_by,
        progress=progress,
    )


def build_loan_compare_rows(
    previous: dict[str, float],
    current: dict[str, float],
    info: dict[str, tuple[str, str, str]],
) -> list[LoanSnapshotRow]:
    all_keys = list(previous)
    for key in current:
        if key not in previous:
            all_keys.append(key)
    rows: list[LoanSnapshotRow] = []
    for key in all_keys:
        old_value = previous.get(key, 0.0)
        new_value = current.get(key, 0.0)
        if old_value <= 0 and new_value > 0:
            category = "Khach hang vay moi"
        elif old_value > 0 and new_value <= 0:
            category = "Khach hang tat toan"
        elif new_value > old_value:
            category = "Khach hang vay tang"
        elif new_value < old_value:
            category = "Khach hang vay giam"
        else:
            category = "Khong thay doi"
        customer_name, address, officer = info.get(key, ("", "", ""))
        rows.append(
            LoanSnapshotRow(
                customer_code=key,
                customer_name=customer_name,
                address=address,
                officer=officer,
                previous_balance=old_value,
                current_balance=new_value,
                category=category,
            )
        )
    return rows


def import_credit_limit_file(
    repository: SummaryRepository,
    file_path: Path,
    *,
    min_limit: float = 0.0,
    warn_days: int = 30,
    reference_date: date | None = None,
    export_path: Path | None = None,
    imported_by: str | None = None,
    progress: ProgressCallback | None = None,
) -> ImportResult:
    file_path = Path(file_path)
    if file_path.suffix.casefold() != ".csv":
        raise SummaryError("Hạn mức tín dụng chỉ hỗ trợ file CSV LN01.")
    reference_date = reference_date or date.today()
    if progress:
        progress(f"Đang đọc file LN01: {file_path.name}")
    started = perf_counter()
    agreements = _load_credit_limit_agreements(file_path)
    rows = filter_credit_limit_rows(
        agreements,
        min_limit=float(min_limit),
        warn_days=int(warn_days),
        reference_date=reference_date,
    )
    duration_ms = int((perf_counter() - started) * 1000)
    batch_id = repository.create_batch(
        SummaryDataType.CREDIT_LIMIT,
        period=reference_date.isoformat(),
        source_path=file_path,
        imported_by=imported_by or _current_user(),
        row_count=len(rows),
        duration_ms=duration_ms,
        message=f"Lọc hạn mức từ {file_path.name}",
        source_hash=_file_hash(file_path),
    )
    repository.save_credit_limit_rows(
        batch_id,
        rows,
        reference_date=reference_date.isoformat(),
        warn_days=warn_days,
        min_limit=min_limit,
    )
    output_path = export_credit_limit_vba(rows, Path(export_path)) if export_path is not None and rows else None
    if progress:
        progress(f"Đã lưu {len(rows):,} hợp đồng cảnh báo")
    return ImportResult(
        batch_id=batch_id,
        row_count=len(rows),
        message=f"Đã lưu {len(rows):,} hợp đồng hạn mức.",
        output_path=output_path,
    )


def filter_credit_limit_rows(
    agreements: Iterable[CreditLimitRow],
    *,
    min_limit: float,
    warn_days: int,
    reference_date: date,
) -> list[CreditLimitRow]:
    selected: list[CreditLimitRow] = []
    for row in agreements:
        if min_limit != 0 and not (row.approved_amount > min_limit):
            continue
        if row.expiry_date is None:
            continue
        days = (row.expiry_date - reference_date).days
        if days < 0:
            selected.append(
                replace(
                    row,
                    days_to_expiry=days,
                    status="Đã hết hạn",
                    note=CREDIT_LIMIT_EXPIRED_NOTE,
                )
            )
        elif 0 <= days <= warn_days:
            selected.append(
                replace(
                    row,
                    days_to_expiry=days,
                    status="Sắp hết hạn",
                    note=CREDIT_LIMIT_SOON_NOTE_TEMPLATE.format(warn_days=warn_days),
                )
            )
    return selected


def _parse_nim_file(
    file_path: Path,
    config: NimConfig,
    *,
    credit_card_rate: float,
    unit_directory: UnitDirectoryService | None = None,
) -> ParsedNimFile:
    unit_directory = unit_directory or get_unit_directory_service()
    data = Path(file_path).read_bytes()
    source_hash = hashlib.sha256(data).hexdigest()
    text = data.decode("utf-8-sig")
    reader = csv.reader(StringIO(text))
    try:
        headers = next(reader)
    except StopIteration as exc:
        raise SummaryError(f"File {file_path.name} không có dữ liệu.") from exc
    header_map = {clean_header(header): index for index, header in enumerate(headers)}
    required = set(NIM_COMMON_HEADERS)
    missing = [header for header in required if header not in header_map]
    if missing:
        raise SummaryError(f"File {file_path.name} thiếu cột: {', '.join(missing)}")
    period = parse_period_from_filename(file_path.name)
    rows: list[NimRow] = []
    customer_rows: list[NormalizedLoanRow] = []
    collect_customer = config.data_type == SummaryDataType.NIM_DN
    customer_missing_headers: tuple[str, ...] = ()
    if collect_customer:
        customer_missing_headers = tuple(
            header for header in CUSTOMER_FTPLN_REQUIRED_HEADERS if header not in header_map
        )
    customer_source_total_balance = 0.0
    customer_invalid_row_count = 0
    customer_warning_count = 0
    customer_warnings: list[str] = []
    unit_warnings: list[str] = []
    checked_units: set[tuple[str, str]] = set()
    missing_trctcd = "TRCTCD" not in header_map
    if collect_customer and missing_trctcd:
        customer_warning_count += 1
        customer_warnings.append(
            f"{file_path.name} thieu cot TRCTCD; ghi chi tiet don vi UNKNOWN, khong tu gan Hoi so."
        )
    for source_row_number, raw in enumerate(reader, start=2):
        if not raw or not any(str(item).strip() for item in raw):
            continue
        value = lambda header: _field(raw, header_map[header])
        optional_value = lambda header: _field(raw, header_map[header]) if header in header_map else ""
        trref = optional_value("TRREF")
        intrt_text = value("INTRT")
        is_glx = trref.upper() == "GLX" or intrt_text == ""
        intrt = float(credit_card_rate) if is_glx else safe_rate(intrt_text)
        ftp = safe_rate(value("FTP"))
        adjustment = safe_rate(value("MUCFTPDC"))
        balance = safe_rate(value("LDRBAL"))
        row_has_numeric_warning = False
        if collect_customer and not customer_missing_headers:
            customer_source_total_balance += balance
            for header in ("LDRBAL", "FTP", "INTRT", "MUCFTPDC"):
                text_value = value(header)
                if text_value and not _has_vba_numeric_prefix(str(text_value).replace(",", ".")):
                    row_has_numeric_warning = True
                    customer_warning_count += 1
                    if len(customer_warnings) < 50:
                        customer_warnings.append(
                            f"{file_path.name}:{source_row_number} gia tri so khong hop le o cot {header}: {text_value}"
                        )
            if row_has_numeric_warning:
                customer_invalid_row_count += 1
        if config.data_type == SummaryDataType.NIM_DN:
            numerator_after = (intrt - ftp - adjustment) * balance
            numerator_before = (intrt - ftp) * balance
            average_rate_numerator = intrt * balance
        else:
            numerator_after = (ftp - intrt + adjustment) * balance
            numerator_before = (ftp - intrt) * balance
            average_rate_numerator = 0.0
        branch_code = value("BRCD")
        trctcd = normalize_trctcd(optional_value("TRCTCD"))
        unit_key = (branch_code, trctcd)
        if unit_key not in checked_units:
            checked_units.add(unit_key)
            unit_warnings.extend(
                unit_directory.ensure_known_unit(
                    branch_code,
                    trctcd,
                    updated_by="import_nim",
                )
            )
        office_code = build_office_code(branch_code, trctcd)
        office_type = classify_office_type(trctcd).value
        transaction_office = unit_directory.get_office_name(branch_code, trctcd)
        officer = optional_value(config.officer_header)
        if config.data_type == SummaryDataType.NIM_DN:
            if not officer or officer == "Không xác định":
                officer = "Thẻ tín dụng"
        elif not officer or officer == "Không xác định":
            officer = "Không gắn mã CB"
        rows.append(
            NimRow(
                batch_id=0,
                data_type=config.data_type,
                period=period,
                branch_code=branch_code,
                branch_name=unit_directory.get_branch_display_name(branch_code),
                trctcd=trctcd,
                transaction_office=transaction_office,
                customer_type=map_customer_type(optional_value("CUSTTP")),
                officer=officer,
                balance=balance,
                interest_rate=intrt,
                ftp_rate=ftp,
                adjustment_rate=adjustment,
                numerator_before=numerator_before,
                numerator_after=numerator_after,
                average_rate_numerator=average_rate_numerator,
                source_file=file_path.name,
            )
        )
        if collect_customer and not customer_missing_headers:
            branch_code = value("BRCD")
            customer_sequence = normalize_customer_sequence(value("CUSTSEQ"))
            customer_code = build_customer_code(branch_code, customer_sequence) if branch_code and customer_sequence else ""
            if not branch_code or not customer_sequence or not customer_code:
                customer_invalid_row_count += 1
                customer_warning_count += 1
                if len(customer_warnings) < 50:
                    customer_warnings.append(
                        f"{file_path.name}:{source_row_number} thieu BRCD/CUSTSEQ, khong tao customer_code."
                    )
                continue
            officer_identity = split_customer_officer(officer)
            customer_rows.append(
                NormalizedLoanRow(
                    period=period,
                    source_file=file_path.name,
                    source_row_number=source_row_number,
                    branch_code=branch_code,
                    trctcd=trctcd,
                    transaction_office=transaction_office,
                    customer_sequence=customer_sequence,
                    customer_code=customer_code,
                    customer_name=clean_cell(value("CUSTNM")),
                    customer_type=map_customer_type_code(value("CUSTTP")).value,
                    ftp_code=clean_cell(value("FTPCD")).upper(),
                    balance=balance,
                    ftp=ftp,
                    interest_rate=intrt,
                    ftp_adjustment=adjustment,
                    officer_code=officer_identity.officer_code,
                    officer_name=officer_identity.officer_name,
                    office_code=office_code,
                    office_name=transaction_office,
                    office_type=office_type,
                )
            )
    return ParsedNimFile(
        rows=rows,
        period=period,
        source_hash=source_hash,
        customer_rows=customer_rows,
        customer_missing_headers=customer_missing_headers,
        customer_source_total_balance=customer_source_total_balance,
        customer_invalid_row_count=customer_invalid_row_count,
        customer_warning_count=customer_warning_count,
        customer_warnings=tuple(customer_warnings),
        unit_warnings=tuple(unit_warnings),
    )


def _load_loan_file(file_path: Path) -> tuple[dict[str, float], dict[str, tuple[str, str, str]]]:
    suffix = file_path.suffix.casefold()
    if suffix in {".csv", ".txt"}:
        rows = _read_vba_simple_split_rows(file_path)
        use_vba_val = True
    elif suffix == ".xlsx":
        rows = _read_excel_rows(file_path)
        use_vba_val = False
    else:
        raise SummaryError(f"Không hỗ trợ định dạng file so sánh tăng giảm khách hàng: {file_path.name}")
    if not rows:
        return {}, {}
    headers = [clean_header(item) for item in rows[0]]
    try:
        col_customer = headers.index("CUSTSEQ")
        col_balance = headers.index("DU_NO")
    except ValueError as exc:
        raise SummaryError(f"File {file_path.name} thiếu cột CUSTSEQ hoặc DU_NO.") from exc
    col_name = _optional_index(headers, "CUSTNM")
    col_address = _optional_index(headers, "ADDR1")
    col_officer = _optional_index(headers, "OFFICER_NAME")
    balances: dict[str, float] = {}
    info: dict[str, tuple[str, str, str]] = {}
    for raw in rows[1:]:
        if len(raw) <= max(col_customer, col_balance):
            continue
        customer_code = clean_cell(raw[col_customer])
        if not customer_code:
            continue
        balance = _vba_loan_amount(raw[col_balance]) if use_vba_val else safe_amount(raw[col_balance])
        balances[customer_code] = balances.get(customer_code, 0.0) + balance
        if customer_code not in info:
            info[customer_code] = (
                clean_cell(raw[col_name]) if col_name is not None and len(raw) > col_name else "",
                clean_cell(raw[col_address]) if col_address is not None and len(raw) > col_address else "",
                clean_cell(raw[col_officer]) if col_officer is not None and len(raw) > col_officer else "",
            )
    return balances, info


def _load_credit_limit_agreements(file_path: Path) -> list[CreditLimitRow]:
    rows = _read_delimited_rows(file_path)
    if not rows:
        raise SummaryError("File LN01 không có dữ liệu.")
    headers = [clean_header(item) for item in rows[0]]
    if len(headers) < 63:
        raise SummaryError("File LN01 không đủ số cột tới CREDIT_LINE_YPE.")
    h0, h1, h2, h3, hbk = headers[0], headers[1], headers[2], headers[3], headers[62]
    if h0 != "BRCD" or h1 not in {"CUSTSED", "CUSTSEQ"} or h2 != "CUSTNM" or h3 != "TAI_KHOAN" or hbk != "CREDIT_LINE_YPE":
        raise SummaryError("Bạn chọn không đúng file xxxx_ln01_yyyymmdd.csv xuất từ mssr98.")
    by_contract: dict[str, CreditLimitRow] = {}
    for raw in rows[1:]:
        if len(raw) < 63:
            continue
        if clean_cell(raw[62]).upper() != "LINE OF CREDIT":
            continue
        contract_number = clean_cell(raw[14])
        if not contract_number:
            continue
        outstanding = safe_amount(raw[5])
        if contract_number in by_contract:
            existing = by_contract[contract_number]
            by_contract[contract_number] = replace(
                existing,
                outstanding_balance=existing.outstanding_balance + outstanding,
            )
            continue
        by_contract[contract_number] = CreditLimitRow(
            customer_code=clean_cell(raw[1]),
            customer_name=clean_cell(raw[2]),
            contract_number=contract_number,
            approved_date=parse_date(raw[15]),
            approved_amount=safe_amount(raw[17]),
            outstanding_balance=outstanding,
            expiry_date=parse_date(raw[18]),
            address=clean_cell(raw[35]),
            officer=clean_cell(raw[27]),
            note="",
            days_to_expiry=None,
            status="",
        )
    return list(by_contract.values())


def _read_delimited_rows(file_path: Path) -> list[list[str]]:
    text = file_path.read_text(encoding="utf-8-sig")
    first_line = text.splitlines()[0] if text.splitlines() else ""
    delimiter = ","
    if "\t" in first_line:
        delimiter = "\t"
    elif ";" in first_line:
        delimiter = ";"
    return [row for row in csv.reader(StringIO(text), delimiter=delimiter)]


def _read_vba_simple_split_rows(file_path: Path) -> list[list[str]]:
    text = file_path.read_text(encoding="utf-8-sig")
    text = text.replace("\r\n", "\n")
    lines = text.split("\n")
    first_line = lines[0] if lines else ""
    delimiter = ","
    if "\t" in first_line:
        delimiter = "\t"
    elif ";" in first_line:
        delimiter = ";"
    return [line.split(delimiter) for line in lines]


def _read_excel_rows(file_path: Path) -> list[list[object]]:
    workbook = load_workbook(file_path, read_only=True, data_only=True)
    try:
        worksheet = workbook.worksheets[0]
        header_row = None
        for row_index, row in enumerate(worksheet.iter_rows(values_only=True), start=1):
            values = list(row)
            if any(clean_header(value) == "CUSTSEQ" for value in values):
                header_row = row_index
                break
        if header_row is None:
            raise SummaryError(f"File {file_path.name} không có cột CUSTSEQ.")
        rows: list[list[object]] = []
        for row in worksheet.iter_rows(min_row=header_row, values_only=True):
            rows.append(list(row))
        return rows
    finally:
        workbook.close()


def parse_period_from_filename(file_name: str) -> str:
    stem = Path(file_name).stem
    date_part = stem[-8:]
    if len(date_part) == 8 and date_part.isdigit():
        return f"{date_part[:4]}-{date_part[4:6]}"
    return "Không Rõ"


def get_branch_name(brcd: str) -> str:
    return get_unit_directory_service().get_branch_display_name(brcd)


def get_trctcd_name(brcd: str, trctcd: str) -> str:
    return get_unit_directory_service().get_office_name(brcd, trctcd)


def map_customer_type(value: str) -> str:
    clean = str(value).strip()
    if clean == "CN":
        return "Cá nhân (CN)"
    if clean == "TC":
        return "Tổ chức (TC)"
    if clean:
        return "Khách hàng khác"
    return "Không rõ"


def safe_rate(value: object) -> float:
    text = clean_cell(value)
    if not text:
        return 0.0
    text = text.replace(",", ".")
    return _vba_val(text)


def safe_amount(value: object) -> float:
    text = clean_cell(value)
    if not text:
        return 0.0
    text = text.replace(",", "")
    try:
        return float(text)
    except ValueError:
        return 0.0


def parse_date(value: object) -> date | None:
    text = clean_cell(value)
    if not text:
        return None
    if " " in text:
        text = text.split(" ", 1)[0]
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%Y%m%d"):
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            pass
    try:
        return datetime.fromisoformat(text).date()
    except ValueError:
        return None


def clean_header(value: object) -> str:
    text = str(value or "").replace("\ufeff", "").strip().replace('"', "")
    upper = text.upper()
    if upper.endswith("CUSTSEQ"):
        return "CUSTSEQ"
    return upper


def clean_cell(value: object) -> str:
    if value is None:
        return ""
    return str(value).strip().replace('"', "")


def _split_officer(raw_name: object) -> tuple[str, str]:
    text = str(raw_name or "").strip()
    match = re.match(r"^\s*\[([^\]]+)\]\s*(.*)$", text)
    if match:
        return match.group(1).strip(), match.group(2).strip()
    return "", text


def _format_officer(officer_code: str, officer_name: str) -> str:
    code = str(officer_code or "").strip()
    name = str(officer_name or "").strip()
    if code:
        return f"[{code}] {name}".strip()
    return name


def _field(row: list[str], index: int) -> str:
    if index < 0 or index >= len(row):
        return ""
    return clean_cell(row[index])


def _optional_index(headers: list[str], header: str) -> int | None:
    try:
        return headers.index(header)
    except ValueError:
        return None


def _add_nim_values(
    target: dict[str, list[float]],
    key: str,
    row: NimRow,
    data_type: SummaryDataType,
) -> None:
    if data_type == SummaryDataType.NIM_DN:
        values = [
            row.numerator_after,
            row.balance,
            row.numerator_before,
            row.average_rate_numerator,
        ]
    else:
        values = [
            row.numerator_after,
            row.balance,
            row.numerator_before,
        ]
    if key not in target:
        target[key] = values
        return
    current = target[key]
    for index, value in enumerate(values):
        current[index] += value


def _nim_detail_output(
    key: str,
    values: list[float],
    data_type: SummaryDataType,
) -> list[object]:
    period, branch_name, transaction_office, customer_type, officer = key.split("|")
    if data_type == SummaryDataType.NIM_DN:
        nim_after, balance, nim_before, average_rate = _nim_rates(values, include_average=True)
        return [
            period,
            branch_name,
            transaction_office,
            customer_type,
            officer,
            average_rate,
            nim_before,
            nim_after,
        ]
    nim_after, _balance, nim_before = _nim_rates(values, include_average=False)
    return [
        period,
        branch_name,
        transaction_office,
        customer_type,
        officer,
        nim_before,
        nim_after,
    ]


def _nim_branch_output(
    key: str,
    values: list[float],
    data_type: SummaryDataType,
) -> list[object]:
    period, branch_name, transaction_office, customer_type = key.split("|")
    if data_type == SummaryDataType.NIM_DN:
        nim_after, balance, nim_before, average_rate = _nim_rates(values, include_average=True)
        _ = balance
        return [
            period,
            branch_name,
            transaction_office,
            customer_type,
            average_rate,
            nim_before,
            nim_after,
        ]
    nim_after, _balance, nim_before = _nim_rates(values, include_average=False)
    return [
        period,
        branch_name,
        transaction_office,
        customer_type,
        nim_before,
        nim_after,
    ]


def _nim_rates(values: list[float], *, include_average: bool) -> tuple[float, ...]:
    numerator_after = values[0]
    balance = values[1]
    numerator_before = values[2]
    nim_after = numerator_after / balance if balance else 0.0
    nim_before = numerator_before / balance if balance else 0.0
    if include_average:
        average_rate_numerator = values[3]
        average_rate = average_rate_numerator / balance if balance else 0.0
        return nim_after, balance, nim_before, average_rate
    return nim_after, balance, nim_before


def _vba_val(value: object) -> float:
    text = clean_cell(value)
    if not text:
        return 0.0
    match = re.match(r"^[\s]*([+-]?(?:\d+(?:\.\d*)?|\.\d+))", text)
    if match is None:
        return 0.0
    try:
        return float(match.group(1))
    except ValueError:
        return 0.0


def _has_vba_numeric_prefix(value: object) -> bool:
    text = clean_cell(value)
    return bool(re.match(r"^[\s]*([+-]?(?:\d+(?:\.\d*)?|\.\d+))", text))


def _vba_loan_amount(value: object) -> float:
    return _vba_val(clean_cell(value).replace(",", ""))


@contextmanager
def _sqlite_snapshots(database_paths: list[Path]) -> Iterable[None]:
    paths = [Path(path) for path in database_paths]
    with tempfile.TemporaryDirectory(prefix="agribank-import-snapshot-") as temporary_directory:
        temporary_root = Path(temporary_directory)
        snapshots: list[tuple[Path, Path | None]] = []
        for index, path in enumerate(paths):
            snapshot_path = temporary_root / f"{index}-{path.name}"
            if path.is_file():
                _checkpoint_sqlite(path)
                with closing(sqlite3.connect(path, timeout=30)) as source:
                    with closing(sqlite3.connect(snapshot_path, timeout=30)) as target:
                        source.backup(target)
                snapshots.append((path, snapshot_path))
            else:
                snapshots.append((path, None))
        try:
            yield
        except Exception:
            for path, snapshot_path in reversed(snapshots):
                _restore_sqlite_snapshot(path, snapshot_path)
            raise


def _checkpoint_sqlite(path: Path) -> None:
    if not path.is_file():
        return
    try:
        with closing(sqlite3.connect(path, timeout=30)) as connection:
            connection.execute("PRAGMA wal_checkpoint(FULL)")
    except sqlite3.Error:
        pass


def _restore_sqlite_snapshot(path: Path, snapshot_path: Path | None) -> None:
    for suffix in ("-wal", "-shm"):
        Path(f"{path}{suffix}").unlink(missing_ok=True)
    if snapshot_path is None:
        path.unlink(missing_ok=True)
        return
    path.parent.mkdir(parents=True, exist_ok=True)
    replacement = path.with_name(f".restore-{os.getpid()}-{path.name}")
    if replacement.exists():
        replacement.unlink()
    with closing(sqlite3.connect(snapshot_path, timeout=30)) as source:
        with closing(sqlite3.connect(replacement, timeout=30)) as target:
            source.backup(target)
    os.replace(replacement, path)


def _file_hash(path: Path) -> str:
    digest = hashlib.sha256()
    with Path(path).open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _combined_file_hash(paths: Iterable[Path]) -> str:
    digest = hashlib.sha256()
    for path in paths:
        file_path = Path(path)
        digest.update(file_path.name.encode("utf-8", errors="ignore"))
        digest.update(b"\0")
        digest.update(_file_hash(file_path).encode("ascii"))
        digest.update(b"\0")
    return digest.hexdigest()


def _current_user() -> str:
    try:
        return getpass.getuser()
    except Exception:
        return ""
