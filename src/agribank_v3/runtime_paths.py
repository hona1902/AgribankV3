from __future__ import annotations

from pathlib import Path
import sys


def application_root() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parents[2]


def access_database_path() -> Path:
    if getattr(sys, "frozen", False):
        return application_root() / "Data" / "AgribankMenuData.mdb"
    return application_root().parent / "Data" / "AgribankMenuData.mdb"
