#!/usr/bin/env python3
"""전기공사 견적·공량 산출 데스크톱 앱 진입점."""

from __future__ import annotations

import sys

from PyQt6.QtWidgets import QApplication

from app.main_window import MainWindow


def main() -> int:
    app = QApplication(sys.argv)
    app.setApplicationName("전기공사 공량 산출")
    app.setOrganizationName("전기공사공량산출")
    window = MainWindow()
    window.show()
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
