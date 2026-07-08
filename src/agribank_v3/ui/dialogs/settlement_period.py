from __future__ import annotations

from PySide6.QtCore import QSettings


SETTINGS_KEY = "settlement/output_prefix"


def normalize_output_prefix(value: object) -> str:
    return "BN" if str(value or "").strip().upper() == "BN" else "QT"


def load_output_prefix() -> str:
    return normalize_output_prefix(QSettings().value(SETTINGS_KEY, "QT"))


def save_output_prefix(prefix: str) -> None:
    QSettings().setValue(SETTINGS_KEY, normalize_output_prefix(prefix))
