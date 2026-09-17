#!/usr/bin/env python3
"""자동 내역서식 프로그램 진입점."""

from __future__ import annotations

import sys

from PyQt6.QtGui import QIcon
from PyQt6.QtWidgets import QApplication

from app.main_window import MainWindow
from app.paths import app_icon_path
from app.version import APP_TITLE


def main() -> int:
    app = QApplication(sys.argv)
    app.setApplicationName(APP_TITLE)
    app.setOrganizationName("전기공사공량산출")
    icon_file = app_icon_path()
    if icon_file.is_file():
        app.setWindowIcon(QIcon(str(icon_file)))
    window = MainWindow()
    window.show()
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
