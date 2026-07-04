from __future__ import annotations

from dataclasses import dataclass
import os
from pathlib import Path
import struct
import subprocess
import tempfile
from typing import Iterable
from uuid import uuid4
import winreg
import zipfile

import win32api


@dataclass(frozen=True, slots=True)
class ExcelInstallation:
    path: Path
    version: str
    major_version: int
    architecture: str

    @property
    def display_name(self) -> str:
        names = {
            12: "Microsoft Excel 2007",
            14: "Microsoft Excel 2010",
            15: "Microsoft Excel 2013",
            16: "Microsoft Excel 2016/2019/2021/2024/365",
        }
        product = names.get(self.major_version, f"Microsoft Excel {self.version}")
        return f"{product} ({self.architecture})"


@dataclass(frozen=True, slots=True)
class ExcelLaunchHandle:
    process: subprocess.Popen[bytes]
    bootstrap_workbook: Path


def _candidate_paths() -> Iterable[Path]:
    for environment_name in ("ProgramFiles", "ProgramFiles(x86)"):
        root_value = os.environ.get(environment_name)
        if not root_value:
            continue
        root = Path(root_value) / "Microsoft Office"
        for office_directory in (
            "Office12",
            "Office14",
            "Office15",
            "Office16",
            "root/Office16",
        ):
            yield root / office_directory / "EXCEL.EXE"

    registry_locations = (
        (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\excel.exe"),
        (winreg.HKEY_CURRENT_USER, r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\excel.exe"),
    )
    access_modes = (
        winreg.KEY_READ | winreg.KEY_WOW64_32KEY,
        winreg.KEY_READ | winreg.KEY_WOW64_64KEY,
    )
    for hive, key_name in registry_locations:
        for access in access_modes:
            try:
                with winreg.OpenKey(hive, key_name, 0, access) as key:
                    value, _ = winreg.QueryValueEx(key, None)
                    yield Path(value)
            except OSError:
                continue


def _file_version(path: Path) -> tuple[str, int]:
    try:
        info = win32api.GetFileVersionInfo(str(path), "\\")
        major = info["FileVersionMS"] >> 16
        minor = info["FileVersionMS"] & 0xFFFF
        return f"{major}.{minor}", major
    except (OSError, KeyError, TypeError):
        return "unknown", 0


def _pe_architecture(path: Path) -> str:
    try:
        with path.open("rb") as executable:
            executable.seek(0x3C)
            pe_offset = struct.unpack("<I", executable.read(4))[0]
            executable.seek(pe_offset + 4)
            machine = struct.unpack("<H", executable.read(2))[0]
        return {0x014C: "32-bit", 0x8664: "64-bit"}.get(machine, "không rõ")
    except (OSError, struct.error):
        return "không rõ"


def discover_excel_installations() -> list[ExcelInstallation]:
    installations: list[ExcelInstallation] = []
    seen: set[str] = set()
    for candidate in _candidate_paths():
        try:
            path = candidate.resolve(strict=True)
        except OSError:
            continue
        key = str(path).casefold()
        if key in seen:
            continue
        seen.add(key)
        version, major = _file_version(path)
        installations.append(
            ExcelInstallation(
                path=path,
                version=version,
                major_version=major,
                architecture=_pe_architecture(path),
            )
        )
    return sorted(
        installations,
        key=lambda item: (item.major_version, item.architecture),
        reverse=True,
    )


def _bootstrap_workbook() -> Path:
    directory = Path(tempfile.gettempdir()) / "AgribankV3"
    directory.mkdir(parents=True, exist_ok=True)
    workbook = directory / f"AgribankV3-New-{uuid4().hex[:8]}.xlsx"
    with zipfile.ZipFile(workbook, "w", zipfile.ZIP_DEFLATED) as archive:
        archive.writestr(
            "[Content_Types].xml",
            """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>""",
        )
        archive.writestr(
            "_rels/.rels",
            """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>""",
        )
        archive.writestr(
            "xl/workbook.xml",
            """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>
</workbook>""",
        )
        archive.writestr(
            "xl/_rels/workbook.xml.rels",
            """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>""",
        )
        archive.writestr(
            "xl/worksheets/sheet1.xml",
            """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData/>
</worksheet>""",
        )
    return workbook


def launch_excel(installation: ExcelInstallation) -> ExcelLaunchHandle:
    # Opening an actual workbook makes older/partially activated Excel releases
    # publish a complete COM Application object. /x keeps the chosen release in
    # its own process.
    bootstrap = _bootstrap_workbook()
    process = subprocess.Popen(
        [str(installation.path), "/x", str(bootstrap)]
    )
    return ExcelLaunchHandle(process=process, bootstrap_workbook=bootstrap)
