from __future__ import annotations

from pathlib import Path

from PySide6.QtCore import Qt
from PySide6.QtWidgets import QWidget

from agribank_v3.features.credit.summary.customer.officer_center_window import (
    OFFICER_CENTER_TITLE,
    OfficerCenterWindow,
)


def open_officer_center_window(
    parent: QWidget,
    main_database_path: Path,
    *,
    initial_tab: str | int | None = None,
) -> OfficerCenterWindow:
    host = _window_host(parent)
    existing = getattr(host, "_officer_center_window", None)
    if _is_live_window(existing):
        if initial_tab is not None and hasattr(existing, "select_tab"):
            existing.select_tab(initial_tab)
        _raise_window(existing)
        return existing
    window = OfficerCenterWindow(Path(main_database_path), parent=parent)
    if initial_tab is not None:
        window.select_tab(initial_tab)
    setattr(host, "_officer_center_window", window)
    window.finished.connect(lambda _result, item=window, owner=host: _clear_window(owner, item))
    window.destroyed.connect(lambda _object=None, item=window, owner=host: _clear_window(owner, item))
    window.show()
    _raise_window(window)
    return window


def _window_host(parent: QWidget) -> QWidget:
    current: QWidget | None = parent
    fallback = parent
    while current is not None:
        if hasattr(current, "_officer_center_window"):
            return current
        fallback = current
        current = current.parent()
    return fallback


def _is_live_window(window: object) -> bool:
    if window is None:
        return False
    try:
        return bool(window.isVisible() or not window.isHidden())
    except RuntimeError:
        return False


def _raise_window(window: OfficerCenterWindow) -> None:
    if window.windowState() & Qt.WindowState.WindowMinimized:
        window.setWindowState(window.windowState() & ~Qt.WindowState.WindowMinimized)
        window.showNormal()
    else:
        window.show()
    window.raise_()
    window.activateWindow()


def _clear_window(owner: QWidget, window: OfficerCenterWindow) -> None:
    try:
        current = getattr(owner, "_officer_center_window", None)
    except RuntimeError:
        return
    if current is window:
        setattr(owner, "_officer_center_window", None)
