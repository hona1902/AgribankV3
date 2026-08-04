from __future__ import annotations

from pathlib import Path

from PySide6.QtWidgets import QWidget

from agribank_v3.features.credit.summary.customer.officer_center_repository import OfficerCenterFilters
from agribank_v3.features.credit.summary.customer.officer_detail_window import OfficerDetailWindow


def open_shared_officer_detail(
    owner: QWidget,
    main_database_path: Path,
    row: dict[str, object],
    *,
    filters: OfficerCenterFilters | None = None,
    initial_tab: int | None = None,
) -> OfficerDetailWindow | None:
    code = str(row.get("officer_code") or row.get("imported_officer_code") or "").strip()
    name = str(row.get("officer_name") or row.get("imported_officer_name") or row.get("officer_display") or "").strip()
    officer_key = str(row.get("officer_key") or "").strip() or _stable_identity(code, name)
    if not code and not name and officer_key == "UNRESOLVED":
        return None
    registry = _registry(owner)
    existing = registry.get(officer_key)
    if existing is not None:
        try:
            if initial_tab is not None:
                existing.tabs.setCurrentIndex(max(0, min(existing.tabs.count() - 1, int(initial_tab))))
            existing.showNormal()
            existing.raise_()
            existing.activateWindow()
            return existing
        except RuntimeError:
            registry.pop(officer_key, None)
    dialog = OfficerDetailWindow(
        Path(main_database_path),
        officer_code=code,
        officer_name=name,
        officer_key=officer_key,
        branch_code=str(row.get("branch_code") or "").strip(),
        transaction_office=str(row.get("transaction_office") or row.get("trctcd") or "").strip(),
        filters=filters,
        parent=owner,
    )
    if initial_tab is not None:
        dialog.tabs.setCurrentIndex(max(0, min(dialog.tabs.count() - 1, int(initial_tab))))
    registry[officer_key] = dialog
    dialog.finished.connect(lambda _result, key=officer_key: registry.pop(key, None))
    dialog.show()
    dialog.raise_()
    dialog.activateWindow()
    return dialog


def _registry(owner: QWidget) -> dict[str, OfficerDetailWindow]:
    registry = getattr(owner, "_shared_officer_detail_windows", None)
    if not isinstance(registry, dict):
        registry = {}
        setattr(owner, "_shared_officer_detail_windows", registry)
    return registry


def _stable_identity(code: str, name: str) -> str:
    code = str(code or "").strip()
    name = str(name or "").strip()
    if code:
        return f"CODE:{code}"
    if name:
        return f"NAME:{name.upper()}"
    return "UNRESOLVED"
