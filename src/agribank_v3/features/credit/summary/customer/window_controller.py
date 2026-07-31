from __future__ import annotations

from collections.abc import Callable
from pathlib import Path

from PySide6.QtCore import Qt
from PySide6.QtWidgets import QWidget

from agribank_v3.features.credit.summary.customer.management_window import CustomerManagementWindow
from agribank_v3.features.credit.summary.customer.routes import (
    CUSTOMER_DATA_DESCRIPTION,
    CUSTOMER_DATA_ROUTE,
    CUSTOMER_DATA_TITLE,
)


def open_customer_management_window(
    parent: QWidget,
    main_database_path: Path,
    *,
    open_nim_dn_callback: Callable[[], None] | None = None,
    initial_tab: str | int | None = None,
) -> CustomerManagementWindow:
    host = _window_host(parent)
    existing = getattr(host, "_customer_management_window", None)
    if _is_live_window(existing):
        if initial_tab is not None and hasattr(existing, "select_tab"):
            existing.select_tab(initial_tab)
        _raise_window(existing)
        return existing
    dialog = CustomerManagementWindow(Path(main_database_path), parent=parent)
    if open_nim_dn_callback is not None:
        dialog.openNimDnRequested.connect(open_nim_dn_callback)
    if initial_tab is not None:
        dialog.select_tab(initial_tab)
    setattr(host, "_customer_management_window", dialog)
    dialog.finished.connect(lambda _result, item=dialog, owner=host: _clear_window(owner, item))
    dialog.destroyed.connect(lambda _object=None, item=dialog, owner=host: _clear_window(owner, item))
    dialog.show()
    _raise_window(dialog)
    return dialog


def _window_host(parent: QWidget) -> QWidget:
    current: QWidget | None = parent
    fallback = parent
    while current is not None:
        if hasattr(current, "_customer_management_window"):
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


def _raise_window(window: CustomerManagementWindow) -> None:
    if window.windowState() & Qt.WindowState.WindowMinimized:
        window.setWindowState(window.windowState() & ~Qt.WindowState.WindowMinimized)
        window.showNormal()
    else:
        window.show()
    window.raise_()
    window.activateWindow()


def _clear_window(owner: QWidget, window: CustomerManagementWindow) -> None:
    try:
        current = getattr(owner, "_customer_management_window", None)
    except RuntimeError:
        return
    if current is window:
        setattr(owner, "_customer_management_window", None)
