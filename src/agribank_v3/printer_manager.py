from __future__ import annotations

from dataclasses import dataclass
import ipaddress
import logging
import os
import subprocess
import sys

try:  # pragma: no cover - import availability depends on the platform.
    import win32print
except Exception:  # pragma: no cover
    win32print = None


LOGGER = logging.getLogger(__name__)


class PrinterManagerError(RuntimeError):
    """Raised when a Windows printer operation cannot be completed."""


@dataclass(frozen=True, slots=True)
class PrinterInfo:
    """Normalized printer information used by the UI."""

    name: str
    is_default: bool
    status: str
    driver_name: str | None
    port_name: str | None
    connection_type: str | None


def is_windows() -> bool:
    """Return True when the current runtime is Windows."""

    return sys.platform.startswith("win")


def powershell_single_quote(value: str) -> str:
    """Escape a value for safe use inside a PowerShell single-quoted string."""

    return "'" + value.replace("'", "''") + "'"


def get_default_printer() -> str | None:
    """Return the Windows default printer name, if available."""

    if not is_windows():
        return None
    if win32print is not None:
        try:
            return win32print.GetDefaultPrinter()
        except Exception as exc:
            LOGGER.exception("Không lấy được máy in mặc định qua pywin32: %s", exc)
    result = _run_powershell("(Get-CimInstance Win32_Printer | Where-Object Default).Name", check=False)
    text = result.stdout.strip()
    return text or None


def get_installed_printers() -> tuple[PrinterInfo, ...]:
    """Return installed Windows printers with normalized metadata."""

    _ensure_windows()
    default_name = get_default_printer()
    printers = _get_printers_with_pywin32(default_name)
    if printers is not None:
        return printers
    return _get_printers_with_powershell(default_name)


def set_default_printer(printer_name: str) -> None:
    """Set a printer as the Windows default printer."""

    _ensure_printer_name(printer_name)
    if win32print is not None:
        try:
            win32print.SetDefaultPrinter(printer_name)
            return
        except Exception as exc:
            LOGGER.exception("Không đặt được máy in mặc định qua pywin32: %s", exc)
    _run_powershell(f"Set-Printer -Name {powershell_single_quote(printer_name)} -IsDefault $true")


def open_printer_queue(printer_name: str) -> None:
    """Open the Windows print queue window for a printer."""

    _ensure_printer_name(printer_name)
    _run_detached(["rundll32.exe", "printui.dll,PrintUIEntry", "/o", "/n", printer_name])


def open_printer_properties(printer_name: str) -> None:
    """Open the Windows printer properties dialog."""

    _ensure_printer_name(printer_name)
    _run_detached(["rundll32.exe", "printui.dll,PrintUIEntry", "/p", "/n", printer_name])


def print_test_page(printer_name: str) -> None:
    """Ask Windows to print a test page for the selected printer."""

    _ensure_printer_name(printer_name)
    _run_detached(["rundll32.exe", "printui.dll,PrintUIEntry", "/k", "/n", printer_name])


def open_windows_printer_settings() -> None:
    """Open the Windows printers settings page with Control Panel fallback."""

    _ensure_windows()
    try:
        os.startfile("ms-settings:printers")  # type: ignore[attr-defined]
        return
    except Exception as exc:
        LOGGER.exception("Không mở được Windows Settings máy in: %s", exc)
    _run_detached(["control.exe", "printers"])


def remove_printer(printer_name: str) -> None:
    """Remove a printer without deleting its driver or port."""

    _ensure_printer_name(printer_name)
    _run_powershell(f"Remove-Printer -Name {powershell_single_quote(printer_name)}")


def _get_printers_with_pywin32(default_name: str | None) -> tuple[PrinterInfo, ...] | None:
    if win32print is None:
        return None
    try:
        flags = win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS
        raw_printers = win32print.EnumPrinters(flags, None, 2)
    except Exception as exc:
        LOGGER.exception("Không đọc được danh sách máy in qua pywin32: %s", exc)
        return None
    printers: list[PrinterInfo] = []
    for printer in raw_printers:
        name = _printer_field(printer, "pPrinterName", 1)
        if not name:
            continue
        port_name = _printer_field(printer, "pPortName", 3)
        driver_name = _printer_field(printer, "pDriverName", 4)
        status_value = _printer_field(printer, "Status", 18)
        attributes = _printer_field(printer, "Attributes", 13)
        printers.append(
            PrinterInfo(
                name=name,
                is_default=(name == default_name),
                status=_status_text(status_value),
                driver_name=driver_name or None,
                port_name=port_name or None,
                connection_type=_connection_type(name, port_name, driver_name, attributes),
            )
        )
    return tuple(sorted(printers, key=lambda item: (not item.is_default, item.name.casefold())))


def _get_printers_with_powershell(default_name: str | None) -> tuple[PrinterInfo, ...]:
    script = (
        "Get-Printer | Select-Object Name,PrinterStatus,DriverName,PortName "
        "| ConvertTo-Csv -NoTypeInformation"
    )
    result = _run_powershell(script)
    import csv
    from io import StringIO

    printers: list[PrinterInfo] = []
    for row in csv.DictReader(StringIO(result.stdout)):
        name = (row.get("Name") or "").strip()
        if not name:
            continue
        port_name = (row.get("PortName") or "").strip() or None
        driver_name = (row.get("DriverName") or "").strip() or None
        printers.append(
            PrinterInfo(
                name=name,
                is_default=(name == default_name),
                status=(row.get("PrinterStatus") or "Không rõ").strip(),
                driver_name=driver_name,
                port_name=port_name,
                connection_type=_connection_type(name, port_name, driver_name, None),
            )
        )
    return tuple(sorted(printers, key=lambda item: (not item.is_default, item.name.casefold())))


def _connection_type(
    name: str,
    port_name: str | None,
    driver_name: str | None,
    attributes: object,
) -> str:
    text = " ".join(value or "" for value in (name, port_name, driver_name)).casefold()
    if any(token in text for token in ("pdf", "xps", "onenote", "fax")):
        return "Máy in ảo"
    if port_name and port_name.startswith("\\\\"):
        return "Máy in chia sẻ"
    if port_name and (port_name.upper().startswith("IP_") or _looks_like_ip_port(port_name)):
        return "Máy in mạng theo IP"
    if port_name and any(token in port_name.upper() for token in ("USB", "LPT", "COM", "PORTPROMPT")):
        return "USB/local"
    if isinstance(attributes, int) and win32print is not None:
        if attributes & getattr(win32print, "PRINTER_ATTRIBUTE_NETWORK", 0):
            return "Máy in chia sẻ"
    return "Không rõ"


def _looks_like_ip_port(port_name: str) -> bool:
    try:
        ipaddress.ip_address(port_name)
        return True
    except ValueError:
        return False


def _status_text(value: object) -> str:
    try:
        status = int(value or 0)
    except (TypeError, ValueError):
        return str(value or "Không rõ")
    if status == 0:
        return "Sẵn sàng"
    labels = {
        0x00000001: "Tạm dừng",
        0x00000002: "Lỗi",
        0x00000004: "Đang xóa",
        0x00000010: "Kẹt giấy",
        0x00000020: "Hết giấy",
        0x00000080: "Ngoại tuyến",
        0x00000200: "Đang in",
        0x00000400: "Đang xuất dữ liệu",
        0x00400000: "Cần chú ý",
    }
    matched = [label for bit, label in labels.items() if status & bit]
    return ", ".join(matched) if matched else f"Trạng thái {status}"


def _printer_field(printer: object, key: str, index: int) -> object:
    if isinstance(printer, dict):
        return printer.get(key)
    try:
        return printer[index]
    except (IndexError, TypeError):
        return None


def _run_powershell(script: str, *, check: bool = True) -> subprocess.CompletedProcess[str]:
    _ensure_windows()
    command = [
        "powershell.exe",
        "-NoProfile",
        "-ExecutionPolicy",
        "Bypass",
        "-Command",
        script,
    ]
    try:
        result = subprocess.run(
            command,
            capture_output=True,
            text=True,
            timeout=90,
            check=False,
        )
    except FileNotFoundError as exc:
        raise PrinterManagerError("Không tìm thấy PowerShell trên Windows.") from exc
    if check and result.returncode != 0:
        LOGGER.error("PowerShell lỗi: %s", result.stderr.strip())
        raise PrinterManagerError(_friendly_command_error(result.stderr or result.stdout))
    return result


def _run_detached(arguments: list[str]) -> None:
    _ensure_windows()
    try:
        subprocess.Popen(arguments, close_fds=True)
    except OSError as exc:
        LOGGER.exception("Không chạy được lệnh Windows: %s", arguments)
        raise PrinterManagerError(f"Không mở được cửa sổ Windows: {exc}") from exc


def _friendly_command_error(text: str) -> str:
    detail = (text or "").strip()
    if not detail:
        return "Lệnh Windows không thực hiện được."
    if "Access is denied" in detail or "denied" in detail.casefold():
        return "Windows từ chối thao tác. Có thể cần chạy bằng quyền Administrator."
    return detail[-1000:]


def _ensure_windows() -> None:
    if not is_windows():
        raise PrinterManagerError("Chức năng quản lý máy in hiện chỉ hỗ trợ Windows.")


def _ensure_printer_name(printer_name: str) -> None:
    _ensure_windows()
    if not printer_name.strip():
        raise PrinterManagerError("Chưa chọn máy in.")
