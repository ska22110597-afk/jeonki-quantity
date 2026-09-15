from __future__ import annotations

import os

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PyQt6.QtCore import QSettings
from PyQt6.QtWidgets import QApplication, QFrame, QLabel

from app.main_window import MainWindow
from app.paths import display_result_directory


def test_run_button_requires_file_and_confirm(tmp_path) -> None:
    QSettings("전기공사공량산출", "GongryangCalc").clear()
    app = QApplication.instance() or QApplication([])
    window = MainWindow()
    try:
        assert window.run_button.isEnabled() is False
        assert window.dest_edit.text() == display_result_directory()
        assert window.dest_edit.isReadOnly() is False
        assert window.browse_button.text() == "폴더 찾기"
        assert window.source_edit.minimumHeight() >= 36
        assert window.dest_edit.minimumHeight() >= 36
        assert window.dest_edit.height() <= 44
        assert window.minimumWidth() >= 980
        assert window.windowTitle().startswith("제작자_박남석")
        assert "v1." in window.windowTitle()
        assert window.reset_button.text() == "새로고침"
        assert window.reset_button.minimumWidth() >= 120
        assert window.findChild(QLabel, "badge") is None
        assert window.findChild(QFrame, "laneReverse") is not None
        assert "내역서_결과_" in window.confirm_note.text()
        assert window.ilwidae_drop is not None
        assert window.forward_estimate_drop is not None
        assert window.drop_zone is not None
        assert window.reverse_ilwidae_drop is not None
        assert window.quantity_drop is not None
        assert window.electric_button.isChecked() is True
        assert window.telecom_button.isChecked() is False
        window.telecom_button.click()
        assert window.telecom_button.isChecked() is True
        assert window.electric_button.isChecked() is False
        assert window._selected_discipline() == "통신"

        source = tmp_path / "단가대비표.xlsx"
        source.write_bytes(b"unused")
        window._on_unit_price_dropped(str(source))
        assert window.run_button.isEnabled() is False
        assert "단가대비표" in window.source_edit.text()

        chosen = tmp_path / "저장위치"
        window.dest_edit.setText(str(chosen))
        window.confirm_box.setChecked(True)
        assert window.run_button.isEnabled() is True
        assert window._chosen_dest_dir() == chosen
        assert "내역서_결과_날짜시간.xlsx" in window.confirm_note.text()

        window._on_reset()
        assert window._fwd_compare_path is None
        assert window.source_edit.text() == ""
        assert window.confirm_box.isChecked() is False
        assert window.run_button.isEnabled() is False
        assert window.compare_drop._title.text() == "단가대비표"
    finally:
        window.close()
        if app is not None:
            app.processEvents()
