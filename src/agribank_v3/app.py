from __future__ import annotations

import sys

from PySide6.QtCore import Qt
from PySide6.QtGui import QFont
from PySide6.QtWidgets import QApplication

from agribank_v3.ui.icons import app_icon
from agribank_v3.ui.main_window import MainWindow
from agribank_v3.ui.styles import APP_STYLESHEET


def _set_windows_app_id() -> None:
    if sys.platform != "win32":
        return
    try:
        from ctypes import windll

        windll.shell32.SetCurrentProcessExplicitAppUserModelID(
            "HNASoft.AgribankV3"
        )
    except (AttributeError, OSError):
        pass


def create_application() -> QApplication:
    _set_windows_app_id()
    app = QApplication(sys.argv)
    app.setApplicationName("AgribankV3")
    app.setApplicationDisplayName("AgribankV3")
    app.setOrganizationName("Agribank")
    app.setWindowIcon(app_icon())
    app.setFont(QFont("Segoe UI", 10))
    app.setStyleSheet(APP_STYLESHEET)
    app.setAttribute(Qt.ApplicationAttribute.AA_DontShowIconsInMenus, False)
    return app


def main() -> int:
    app = create_application()
    window = MainWindow()
    window.show()
    return app.exec()
