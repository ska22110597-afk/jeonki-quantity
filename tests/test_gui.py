from __future__ import annotations

import os

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PyQt6.QtWidgets import QApplication

from app.main_window import MainWindow
from app.paths import display_result_directory


def test_run_button_requires_file_and_confirm(tmp_path) -> None:
    app = QApplication.instance() or QApplication([])
    window = MainWindow()
    try:
        assert window.run_button.isEnabled() is False
        assert window.dest_edit.text() == display_result_directory()

        source = tmp_path / "단가대비표.xlsx"
        source.write_bytes(b"unused")
        window._on_file_dropped(str(source))
        assert window.run_button.isEnabled() is False
        assert window.source_edit.text() == str(source)

        window.confirm_box.setChecked(True)
        assert window.run_button.isEnabled() is True
    finally:
        window.close()
        if app is not None:
            app.processEvents()
