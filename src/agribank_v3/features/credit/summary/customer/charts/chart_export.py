from __future__ import annotations

from datetime import datetime
from pathlib import Path

from PySide6.QtWidgets import QWidget


def save_widget_png(widget: QWidget, destination: Path | str) -> Path:
    path = Path(destination)
    path.parent.mkdir(parents=True, exist_ok=True)
    widget.grab().save(str(path), "PNG")
    return path


def default_chart_png_name(chart_name: str) -> str:
    safe_name = "".join(char for char in chart_name if char.isalnum() or char in {"_", "-"}).strip() or "Chart"
    return f"CustomerChart_{safe_name}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png"
