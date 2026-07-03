from __future__ import annotations

import ctypes
import sys

from agribank_v3.excel import ExcelConnectionError, ExcelService


def is_administrator() -> bool:
    try:
        return bool(ctypes.windll.shell32.IsUserAnAdmin())
    except (AttributeError, OSError):
        return False


def main() -> int:
    print("AgribankV3 - Excel connection diagnostic")
    print(f"Python: {sys.version.split()[0]} ({sys.executable})")
    print(f"AgribankV3 chạy Administrator: {'Có' if is_administrator() else 'Không'}")
    service = ExcelService()
    try:
        context = service.connect()
    except ExcelConnectionError as exc:
        print(f"KẾT NỐI THẤT BẠI: {exc}")
        print(
            "Kiểm tra: Excel phải có ít nhất một workbook; không để hộp thoại "
            "modal đang mở; Excel và AgribankV3 phải chạy cùng mức quyền."
        )
        return 1

    print("KẾT NỐI THÀNH CÔNG")
    print(f"Phiên bản: {context.excel_name} (COM {context.excel_version})")
    print(f"Workbook: {context.workbook}")
    print(f"Worksheet: {context.worksheet}")
    print(f"Selection: {context.selection} ({context.cell_count:,} ô)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
