#!/usr/bin/env python3
"""제작자_박남석 자동 내역서식 프로그램 진입점."""

from __future__ import annotations

import sys

from PyQt6.QtWidgets import QApplication

from app.main_window import MainWindow
from app.version import APP_TITLE


def main() -> int:
    app = QApplication(sys.argv)
    app.setApplicationName(APP_TITLE)
    app.setOrganizationName("전기공사공량산출")
    window = MainWindow()
    window.show()
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
