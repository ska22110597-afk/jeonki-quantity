from __future__ import annotations

import os

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PyQt6.QtCore import QSettings
from PyQt6.QtWidgets import QApplication, QFrame, QLabel, QProgressBar

from app.main_window import BusyDialog, DoneDialog, MainWindow, SETTINGS_APP, SETTINGS_APP_LEGACY, SETTINGS_ORG
from app.paths import display_result_directory


def _clear_settings() -> None:
    QSettings(SETTINGS_ORG, SETTINGS_APP).clear()
    QSettings(SETTINGS_ORG, SETTINGS_APP_LEGACY).clear()


def test_run_button_requires_file_and_confirm(tmp_path) -> None:
    _clear_settings()
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
        assert "일위대가목록_결과_" in window.confirm_note.text()
        assert window.ilwidae_drop is not None
        assert window.forward_estimate_drop is None
        assert window.drop_zone is not None
        assert window.reverse_ilwidae_drop is not None
        assert window.quantity_drop is not None
        assert window.quantity_drop._idle_title == "일위대가목록"
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
        assert "일위대가목록_결과_날짜시간.xlsx" in window.confirm_note.text()

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


def test_busy_dialog_tells_user_not_to_click_again() -> None:
    _clear_settings()
    app = QApplication.instance() or QApplication([])
    window = MainWindow()
    try:
        dialog = window._open_busy_dialog()
        try:
            assert isinstance(dialog, BusyDialog)
            assert dialog.windowTitle() == "작업 중"
            assert dialog.findChild(QLabel, "doneTitle").text() == "작업 중"
            assert "만드는 중" in dialog.labelText
            assert "다시 누르지" in dialog.labelText
            gauge = dialog.findChild(QProgressBar, "busyGauge")
            assert gauge is not None
            assert gauge.minimum() == 0
            assert gauge.maximum() == 100
            assert gauge.value() >= 0
            assert gauge.isTextVisible() is True
        finally:
            dialog.complete_and_close()
        window._lock_run_ui()
        try:
            assert window._busy is True
            assert window.run_button.isEnabled() is False
            assert "누르지 마세요" in window.run_button.text()
            window._refresh_run_enabled()
            assert window.run_button.isEnabled() is False
        finally:
            window._unlock_run_ui()
        assert window._busy is False
        assert window.run_button.text() == "산출 및 저장"
    finally:
        window.close()
        if app is not None:
            app.processEvents()


def test_done_dialog_keeps_short_electrician_copy(tmp_path) -> None:
    _clear_settings()
    app = QApplication.instance() or QApplication([])
    window = MainWindow()
    try:
        dest = tmp_path / "공량산출_결과_20260916.xlsx"
        dialog = DoneDialog(window, dest, "전기_표준품셈.xlsx")
        try:
            assert dialog.windowTitle() == "서식 생성 완료"
            assert dialog.findChild(QLabel, "doneTitle").text() == "서식 생성 완료!"
            assert dialog.findChild(QLabel, "donePath").text() == str(dest)
            assert dialog.findChild(QLabel, "doneRef").text() == "품셈표, 노임단가"
            assert dialog.findChild(QLabel, "donePumsam").text() == "전기_표준품셈.xlsx"
            assert dialog.findChild(QLabel, "doneFoot").text() == (
                "표준품셈·노임단가는 저장 폴더의 데이터베이스에서 고칠 수 있습니다."
            )
            captions = [
                label.text()
                for label in dialog.findChildren(QLabel)
                if label.objectName() == "doneCaption"
            ]
            assert captions == ["파일 저장경로 :", "참고 데이터 :", "사용한 표준품셈 :"]
            assert "원본은 그대로" not in dialog.findChild(QLabel, "doneFoot").text()
        finally:
            dialog.close()
    finally:
        window.close()
        if app is not None:
            app.processEvents()
