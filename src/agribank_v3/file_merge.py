from __future__ import annotations

import csv
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable

from openpyxl import Workbook


class FileMergeError(RuntimeError):
    """Raised when same-structure files cannot be merged."""


@dataclass(frozen=True, slots=True)
class MergeResult:
    output_path: Path
    source_count: int
    row_count: int
    column_count: int


def merge_same_structure_csv_to_xlsx(
    source_paths: Iterable[Path],
    output_path: Path,
    *,
    sheet_name: str = "Data",
) -> MergeResult:
    sources = tuple(Path(path) for path in source_paths)
    if not sources:
        raise FileMergeError("Chưa chọn file CSV để nối.")
    for source in sources:
        if source.suffix.casefold() != ".csv":
            raise FileMergeError(f"File không phải CSV: {source.name}")
        if not source.is_file():
            raise FileMergeError(f"Không tìm thấy file: {source}")

    output = Path(output_path)
    if output.suffix.casefold() != ".xlsx":
        output = output.with_suffix(".xlsx")
    output.parent.mkdir(parents=True, exist_ok=True)

    workbook = Workbook()
    sheet = workbook.active
    sheet.title = sheet_name

    try:
        expected_header: tuple[str, ...] | None = None
        total_rows = 0
        column_count = 0
        for source in sources:
            header, rows = _read_csv(source)
            if not header:
                raise FileMergeError(f"File không có dòng tiêu đề: {source.name}")
            if expected_header is None:
                expected_header = header
                column_count = len(header)
                sheet.append(list(header))
            elif header != expected_header:
                raise FileMergeError(
                    f"File {source.name} không cùng cấu trúc với file {sources[0].name}."
                )
            for row in rows:
                if not any(value != "" for value in row):
                    continue
                normalized = list(row[:column_count])
                if len(normalized) < column_count:
                    normalized.extend([""] * (column_count - len(normalized)))
                sheet.append(normalized)
                total_rows += 1

        workbook.save(output)
    except Exception:
        workbook.close()
        if output.exists():
            output.unlink()
        raise
    workbook.close()
    return MergeResult(
        output_path=output,
        source_count=len(sources),
        row_count=total_rows,
        column_count=column_count,
    )


def merge_same_structure_csv_to_csv(
    source_paths: Iterable[Path],
    output_path: Path,
) -> MergeResult:
    sources = tuple(Path(path) for path in source_paths)
    if not sources:
        raise FileMergeError("Chưa chọn file CSV để nối.")
    output = Path(output_path)
    output.parent.mkdir(parents=True, exist_ok=True)

    expected_header: tuple[str, ...] | None = None
    total_rows = 0
    column_count = 0
    try:
        with output.open("w", encoding="utf-8-sig", newline="") as stream:
            writer = csv.writer(stream)
            for source in sources:
                if source.suffix.casefold() != ".csv":
                    raise FileMergeError(f"File không phải CSV: {source.name}")
                if not source.is_file():
                    raise FileMergeError(f"Không tìm thấy file: {source}")
                header, rows = _read_csv(source)
                if not header:
                    raise FileMergeError(f"File không có dòng tiêu đề: {source.name}")
                if expected_header is None:
                    expected_header = header
                    column_count = len(header)
                    writer.writerow(header)
                elif header != expected_header:
                    raise FileMergeError(
                        f"File {source.name} không cùng cấu trúc với file {sources[0].name}."
                    )
                for row in rows:
                    if not any(value != "" for value in row):
                        continue
                    normalized = list(row[:column_count])
                    if len(normalized) < column_count:
                        normalized.extend([""] * (column_count - len(normalized)))
                    writer.writerow(normalized)
                    total_rows += 1
    except Exception:
        if output.exists():
            output.unlink()
        raise
    return MergeResult(
        output_path=output,
        source_count=len(sources),
        row_count=total_rows,
        column_count=column_count,
    )


def _read_csv(path: Path) -> tuple[tuple[str, ...], list[list[str]]]:
    last_error: Exception | None = None
    for encoding in ("utf-8-sig", "cp1258", "mbcs"):
        try:
            text = path.read_text(encoding=encoding)
            dialect = _sniff_dialect(text)
            reader = csv.reader(text.splitlines(), dialect)
            rows = [list(row) for row in reader]
            if not rows:
                return (), []
            return tuple(cell.strip() for cell in rows[0]), rows[1:]
        except (OSError, UnicodeError, csv.Error) as exc:
            last_error = exc
    raise FileMergeError(f"Không thể đọc file CSV {path.name}: {last_error}")


def _sniff_dialect(text: str) -> csv.Dialect:
    sample = text[:8192]
    try:
        return csv.Sniffer().sniff(sample, delimiters=",;\t|")
    except csv.Error:
        return csv.excel
