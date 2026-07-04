from __future__ import annotations

from importlib.resources import files

from PySide6.QtGui import QIcon


APP_ICON_NAME = "Logo-HNA.png"


def icon_path(name: str) -> str:
    return str(files("agribank_v3").joinpath("resources", "icons", *name.split("/")))


def app_icon() -> QIcon:
    return QIcon(icon_path(APP_ICON_NAME))
