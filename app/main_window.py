"""제작자_박남석 자동 내역서식 프로그램 화면."""

from __future__ import annotations

from pathlib import Path

from PyQt6.QtCore import Qt, QSettings
from PyQt6.QtGui import QIcon
from PyQt6.QtWidgets import (
    QButtonGroup,
    QCheckBox,
    QFileDialog,
    QFrame,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QSizePolicy,
    QStatusBar,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from app.discipline import ELECTRIC, TELECOM, normalize_discipline, pumsam_filename

from app.drop_zone import DropZone
from app.excel_io import (
    COMPARE_SHEET_NAME,
    ESTIMATE_SHEET_NAME,
    ILWIDAE_SHEET_NAME,
    QUANTITY_SHEET_NAME,
    save_result_workbook,
)
from app.paths import ResultDirectoryError, WINDOWS_RESULT_DIR, app_icon_path, display_result_directory, is_windows
from app.pumsam import PUMSAM_SHEET_NAME
from app.version import APP_TAGLINE, APP_TITLE
from app.wages import WAGES_SHEET_NAME

APP_STYLESHEET = """
QMainWindow, QWidget#root {
    background: #F3F4F6;
    color: #243040;
    font-family: "Malgun Gothic", "Apple SD Gothic Neo", "Noto Sans CJK KR", sans-serif;
    font-size: 13px;
}
QFrame#hero {
    background: #1C2B3A;
    border: none;
}
QLabel#appTitle {
    color: #F7F3EA;
    font-size: 20px;
    font-weight: 600;
    letter-spacing: 0.4px;
}
QLabel#appSubtitle {
    color: #B4C2CF;
    font-size: 12px;
    font-weight: 400;
    letter-spacing: 0.2px;
}
QFrame#laneForward {
    background: #C5D6E6;
    border: 2px solid #4A7190;
    border-radius: 12px;
}
QFrame#laneReverse {
    background: #E8D0C9;
    border: 2px solid #8B4A42;
    border-radius: 12px;
}
QFrame#laneQty {
    background: #D5E4D0;
    border: 2px solid #4F7F62;
    border-radius: 12px;
}
QLabel#laneQtyTitle {
    color: #2F5D3F;
}
QLabel#laneTitle {
    font-size: 13px;
    font-weight: 700;
    letter-spacing: 0.3px;
}
QLabel#laneForwardTitle {
    color: #2C4A63;
}
QLabel#laneReverseTitle {
    color: #6B322C;
}
QFrame#card {
    background: #FFFFFF;
    border: 1px solid #D5DDE4;
    border-radius: 12px;
}
QLabel#sectionLabel {
    color: #5A6A78;
    font-size: 12px;
    font-weight: 700;
}
QLineEdit {
    background: #F7F9FB;
    border: 1px solid #D5DDE4;
    border-radius: 8px;
    padding: 6px 10px;
    min-height: 22px;
    color: #243040;
    font-size: 13px;
    selection-background-color: #1C2B3A;
}
QLineEdit#pathEdit {
    font-size: 13px;
    padding: 6px 10px;
}
QLineEdit:read-only {
    color: #334155;
}
QCheckBox {
    color: #243040;
    spacing: 10px;
    font-size: 13px;
}
QCheckBox::indicator {
    width: 18px;
    height: 18px;
}
QLabel#confirmNote {
    color: #5A6A78;
    font-size: 12px;
}
QPushButton#runButton {
    background: #C45911;
    color: #FFFFFF;
    border: none;
    border-radius: 8px;
    padding: 14px 22px;
    font-size: 15px;
    font-weight: 700;
}
QPushButton#runButton:hover {
    background: #A3470C;
}
QPushButton#runButton:disabled {
    background: #C9B8AE;
    color: #F4EDE8;
}
QPushButton#browseButton {
    background: #1C2B3A;
    color: #FFFFFF;
    border: none;
    border-radius: 8px;
    padding: 6px 16px;
    font-size: 13px;
    font-weight: 700;
    min-height: 34px;
    min-width: 104px;
}
QPushButton#browseButton:hover {
    background: #2A4054;
}
QPushButton#resetButton {
    background: #EEF3F7;
    color: #2C4A63;
    border: 1px solid #8AA0B3;
    border-radius: 8px;
    padding: 6px 18px;
    font-size: 13px;
    font-weight: 700;
    min-height: 34px;
    min-width: 120px;
}
QPushButton#resetButton:hover {
    background: #D9E4EE;
}
QTextEdit#log {
    background: #121C28;
    color: #D3DFEA;
    border: none;
    border-radius: 8px;
    padding: 12px;
    font-family: "Malgun Gothic", "Noto Sans CJK KR", sans-serif;
    font-size: 12px;
}
QFrame#dropZoneForward {
    background: #D4E3EE;
    border: 2px dashed #4A7190;
    border-radius: 10px;
}
QFrame#dropZoneForward[hover="true"] {
    background: #C0D4E4;
    border: 2px dashed #2C4A63;
}
QFrame#dropZoneForward[loaded="true"] {
    background: #E7F2EA;
    border: 2px solid #4F7F62;
}
QFrame#dropZoneReverse {
    background: #EDD4CE;
    border: 2px dashed #8B4A42;
    border-radius: 10px;
}
QFrame#dropZoneReverse[hover="true"] {
    background: #E4C4BC;
    border: 2px dashed #6B322C;
}
QFrame#dropZoneReverse[loaded="true"] {
    background: #E7F2EA;
    border: 2px solid #4F7F62;
}
QFrame#dropZoneQty {
    background: #E3EFE0;
    border: 2px dashed #4F7F62;
    border-radius: 10px;
}
QFrame#dropZoneQty[hover="true"] {
    background: #D5E4D0;
    border: 2px dashed #2F5D3F;
}
QFrame#dropZoneQty[loaded="true"] {
    background: #E7F2EA;
    border: 2px solid #4F7F62;
}
QFrame#dropZoneQty QLabel#dropTitle {
    color: #2F5D3F;
}
QLabel#dropTitle {
    font-size: 13px;
    font-weight: 700;
}
QFrame#dropZoneForward QLabel#dropTitle {
    color: #2C4A63;
}
QFrame#dropZoneReverse QLabel#dropTitle {
    color: #6B322C;
}
QLabel#dropHint {
    color: #667888;
    font-size: 12px;
}
QLabel#badge {
    background: #2A4054;
    color: #FFFFFF;
    border-radius: 4px;
    padding: 4px 10px;
    font-size: 11px;
    font-weight: 700;
}
QFrame#partCard {
    background: #FFFFFF;
    border: 1px solid #D5DDE4;
    border-radius: 12px;
}
QLabel#partCaption {
    color: #5A6A78;
    font-size: 12px;
    font-weight: 700;
}
QLabel#partFileHint {
    color: #667888;
    font-size: 12px;
}
QPushButton#partElectric, QPushButton#partTelecom {
    min-height: 48px;
    border-radius: 8px;
    font-size: 16px;
    font-weight: 700;
    letter-spacing: 1px;
}
QPushButton#partElectric {
    background: #EEF3F7;
    border: 2px solid #8AA0B3;
    color: #3E5B70;
}
QPushButton#partElectric:checked {
    background: #3E5B70;
    color: #FFFFFF;
    border: 2px solid #3E5B70;
}
QPushButton#partTelecom {
    background: #F6F1EE;
    border: 2px solid #C4A199;
    color: #7A534C;
}
QPushButton#partTelecom:checked {
    background: #7A534C;
    color: #FFFFFF;
    border: 2px solid #7A534C;
}
QStatusBar {
    background: #E8EDF2;
    color: #5A6A78;
}
"""


class MainWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle(APP_TITLE)
        self.setMinimumSize(1180, 980)
        self.resize(1260, 1080)
        self.setStyleSheet(APP_STYLESHEET)
        icon_file = app_icon_path()
        if icon_file.is_file():
            self.setWindowIcon(QIcon(str(icon_file)))

        self._settings = QSettings("전기공사공량산출", "GongryangCalc")
        self._fwd_compare_path: Path | None = None
        self._fwd_ilwidae_path: Path | None = None
        self._fwd_estimate_path: Path | None = None
        self._rev_estimate_path: Path | None = None
        self._rev_ilwidae_path: Path | None = None
        self._qty_estimate_path: Path | None = None
        self._build_ui()
        self._refresh_run_enabled()

    def _saved_dest_dir(self) -> str:
        stored = self._settings.value("dest_dir", "")
        if isinstance(stored, str) and stored.strip():
            return stored.strip()
        return display_result_directory()

    def _build_ui(self) -> None:
        root = QWidget()
        root.setObjectName("root")
        self.setCentralWidget(root)
        outer = QVBoxLayout(root)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.setSpacing(0)

        hero = QFrame()
        hero.setObjectName("hero")
        hero_layout = QVBoxLayout(hero)
        hero_layout.setContentsMargins(36, 22, 36, 20)
        hero_layout.setSpacing(8)
        title = QLabel(APP_TITLE)
        title.setObjectName("appTitle")
        subtitle = QLabel(APP_TAGLINE)
        subtitle.setObjectName("appSubtitle")
        subtitle.setWordWrap(True)
        hero_layout.addWidget(title)
        hero_layout.addWidget(subtitle)
        outer.addWidget(hero)

        body = QWidget()
        body_layout = QVBoxLayout(body)
        body_layout.setContentsMargins(28, 20, 28, 16)
        body_layout.setSpacing(14)

        part_card = QFrame()
        part_card.setObjectName("partCard")
        part_layout = QVBoxLayout(part_card)
        part_layout.setContentsMargins(22, 14, 22, 14)
        part_layout.setSpacing(8)
        part_caption = QLabel("표준품셈 파트")
        part_caption.setObjectName("partCaption")
        part_layout.addWidget(part_caption)
        part_buttons = QHBoxLayout()
        part_buttons.setSpacing(10)
        self.electric_button = QPushButton("전기")
        self.electric_button.setObjectName("partElectric")
        self.electric_button.setCheckable(True)
        self.electric_button.setCursor(Qt.CursorShape.PointingHandCursor)
        self.telecom_button = QPushButton("통신")
        self.telecom_button.setObjectName("partTelecom")
        self.telecom_button.setCheckable(True)
        self.telecom_button.setCursor(Qt.CursorShape.PointingHandCursor)
        self.part_group = QButtonGroup(self)
        self.part_group.setExclusive(True)
        self.part_group.addButton(self.electric_button)
        self.part_group.addButton(self.telecom_button)
        part_buttons.addWidget(self.electric_button, 1)
        part_buttons.addWidget(self.telecom_button, 1)
        part_layout.addLayout(part_buttons)
        self.part_file_hint = QLabel("")
        self.part_file_hint.setObjectName("partFileHint")
        part_layout.addWidget(self.part_file_hint)
        self.electric_button.clicked.connect(self._on_part_changed)
        self.telecom_button.clicked.connect(self._on_part_changed)
        body_layout.addWidget(part_card)
        self._restore_discipline()

        lanes = QHBoxLayout()
        lanes.setSpacing(14)

        forward_lane = QFrame()
        forward_lane.setObjectName("laneForward")
        forward_layout = QVBoxLayout(forward_lane)
        forward_layout.setContentsMargins(12, 10, 12, 12)
        forward_layout.setSpacing(8)
        forward_title = QLabel("정방향  ·  단가대비표 → 내역서")
        forward_title.setObjectName("laneTitle")
        forward_title.setStyleSheet("color: #2C4A63;")
        forward_title.setWordWrap(True)
        forward_layout.addWidget(forward_title)

        self.compare_drop = DropZone(
            title="단가대비표",
            hint="물량 품목의 자재 단가  ·  놓으면 이후 시트를 작성",
            dialog_title="단가대비표 엑셀 선택",
            tone="forward",
        )
        self.compare_drop.file_dropped.connect(self._on_fwd_compare_dropped)
        forward_layout.addWidget(self.compare_drop)

        self.ilwidae_drop = DropZone(
            title="일위대가",
            hint="이미 만든 호표가 있으면 놓습니다  ·  없으면 프로그램이 작성",
            dialog_title="일위대가 엑셀 선택",
            tone="forward",
        )
        self.ilwidae_drop.file_dropped.connect(self._on_fwd_ilwidae_dropped)
        forward_layout.addWidget(self.ilwidae_drop)

        self.forward_estimate_drop = DropZone(
            title="내역서",
            hint="이미 있는 내역서가 있으면 놓습니다  ·  없으면 프로그램이 작성",
            dialog_title="내역서 엑셀 선택",
            tone="forward",
        )
        self.forward_estimate_drop.file_dropped.connect(self._on_fwd_estimate_dropped)
        forward_layout.addWidget(self.forward_estimate_drop)
        for zone in (self.compare_drop, self.ilwidae_drop, self.forward_estimate_drop):
            zone.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        forward_layout.addStretch(1)
        lanes.addWidget(forward_lane, 1)

        reverse_lane = QFrame()
        reverse_lane.setObjectName("laneReverse")
        reverse_layout = QVBoxLayout(reverse_lane)
        reverse_layout.setContentsMargins(12, 10, 12, 12)
        reverse_layout.setSpacing(8)
        reverse_title = QLabel("역방향  ·  내역서 → 단가대비표")
        reverse_title.setObjectName("laneTitle")
        reverse_title.setStyleSheet("color: #6B322C;")
        reverse_title.setWordWrap(True)
        reverse_layout.addWidget(reverse_title)

        self.drop_zone = DropZone(
            title="내역서",
            hint="이미 있는 내역서를 놓으면 단가대비표를 만듭니다",
            dialog_title="내역서 엑셀 선택",
            tone="reverse",
        )
        self.drop_zone.file_dropped.connect(self._on_rev_estimate_dropped)
        reverse_layout.addWidget(self.drop_zone)

        self.reverse_ilwidae_drop = DropZone(
            title="일위대가",
            hint="이미 만든 호표가 있으면 함께 넣습니다",
            dialog_title="일위대가 엑셀 선택",
            tone="reverse",
        )
        self.reverse_ilwidae_drop.file_dropped.connect(self._on_rev_ilwidae_dropped)
        reverse_layout.addWidget(self.reverse_ilwidae_drop)
        for zone in (self.drop_zone, self.reverse_ilwidae_drop):
            zone.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        reverse_layout.addStretch(1)
        lanes.addWidget(reverse_lane, 1)
        body_layout.addLayout(lanes)

        qty_lane = QFrame()
        qty_lane.setObjectName("laneQty")
        qty_layout = QVBoxLayout(qty_lane)
        qty_layout.setContentsMargins(12, 10, 12, 10)
        qty_layout.setSpacing(8)
        qty_title = QLabel("공량산출  ·  내역서 → 공량산출서")
        qty_title.setObjectName("laneTitle")
        qty_title.setStyleSheet("color: #2F5D3F;")
        qty_title.setWordWrap(True)
        qty_layout.addWidget(qty_title)
        self.quantity_drop = DropZone(
            title="내역서",
            hint="1. 전열설비공사처럼 파트를 나눈 내역서를 놓으면 그 구분 그대로 공량산출서를 만듭니다",
            dialog_title="공량산출용 내역서 엑셀 선택",
            tone="quantity",
        )
        self.quantity_drop.setMaximumHeight(64)
        self.quantity_drop.setMinimumHeight(56)
        self.quantity_drop.file_dropped.connect(self._on_qty_estimate_dropped)
        qty_layout.addWidget(self.quantity_drop)
        body_layout.addWidget(qty_lane)

        path_card = QFrame()
        path_card.setObjectName("card")
        path_layout = QVBoxLayout(path_card)
        path_layout.setContentsMargins(20, 16, 20, 18)
        path_layout.setSpacing(10)

        source_head = QHBoxLayout()
        source_head.setContentsMargins(0, 0, 0, 0)
        source_label = QLabel("선택한 파일  (읽기 전용 · 원본은 수정하지 않습니다)")
        source_label.setObjectName("sectionLabel")
        source_head.addWidget(source_label)
        source_head.addStretch(1)
        self.reset_button = QPushButton("새로고침")
        self.reset_button.setObjectName("resetButton")
        self.reset_button.setCursor(Qt.CursorShape.PointingHandCursor)
        self.reset_button.setMinimumSize(120, 34)
        self.reset_button.clicked.connect(self._on_reset)
        source_head.addWidget(self.reset_button)
        path_layout.addLayout(source_head)
        self.source_edit = QLineEdit()
        self.source_edit.setObjectName("pathEdit")
        self.source_edit.setReadOnly(True)
        self.source_edit.setFixedHeight(36)
        self.source_edit.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        self.source_edit.setPlaceholderText("아직 파일이 없습니다. 정방향·역방향·공량 칸에 엑셀을 놓아 주세요.")
        path_layout.addWidget(self.source_edit)

        dest_label = QLabel("결과 저장 폴더")
        dest_label.setObjectName("sectionLabel")
        path_layout.addWidget(dest_label)

        dest_row = QHBoxLayout()
        dest_row.setContentsMargins(0, 0, 0, 0)
        dest_row.setSpacing(10)
        self.dest_edit = QLineEdit(self._saved_dest_dir())
        self.dest_edit.setObjectName("pathEdit")
        self.dest_edit.setReadOnly(False)
        self.dest_edit.setFixedHeight(36)
        self.dest_edit.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        self.dest_edit.setPlaceholderText(str(WINDOWS_RESULT_DIR))
        dest_row.addWidget(self.dest_edit, 1)
        self.browse_button = QPushButton("폴더 찾기")
        self.browse_button.setObjectName("browseButton")
        self.browse_button.setCursor(Qt.CursorShape.PointingHandCursor)
        self.browse_button.setMinimumSize(104, 36)
        self.browse_button.clicked.connect(self._on_browse_dest)
        dest_row.addWidget(self.browse_button, 0)
        path_layout.addLayout(dest_row)

        self.confirm_note = QLabel("")
        self.confirm_note.setObjectName("confirmNote")
        self.confirm_note.setWordWrap(True)
        path_layout.addWidget(self.confirm_note)
        path_layout.addSpacing(12)
        self.confirm_box = QCheckBox("이 폴더에 새 파일로 저장")
        self.confirm_box.setSizePolicy(QSizePolicy.Policy.Preferred, QSizePolicy.Policy.Fixed)
        self.confirm_box.toggled.connect(self._refresh_run_enabled)
        path_layout.addWidget(self.confirm_box)
        body_layout.addWidget(path_card)

        self.run_button = QPushButton("산출 및 저장")
        self.run_button.setObjectName("runButton")
        self.run_button.setMinimumHeight(50)
        self.run_button.setCursor(Qt.CursorShape.PointingHandCursor)
        self.run_button.clicked.connect(self._on_run)
        body_layout.addWidget(self.run_button)

        log_label = QLabel("처리 기록")
        log_label.setObjectName("sectionLabel")
        self.log = QTextEdit()
        self.log.setObjectName("log")
        self.log.setReadOnly(True)
        self.log.setMinimumHeight(120)
        body_layout.addWidget(log_label)
        body_layout.addWidget(self.log, 1)

        outer.addWidget(body, 1)

        status = QStatusBar()
        status.showMessage("대기 — 정방향 또는 역방향 칸에 엑셀을 놓고 저장 폴더를 확인하세요.")
        self.setStatusBar(status)
        self._write_startup_log()
        self._refresh_filename_hint()

    def _selected_discipline(self) -> str:
        if self.telecom_button.isChecked():
            return TELECOM
        return ELECTRIC

    def _restore_discipline(self) -> None:
        stored = self._settings.value("discipline", ELECTRIC)
        disc = normalize_discipline(str(stored) if stored is not None else ELECTRIC)
        self.electric_button.setChecked(disc == ELECTRIC)
        self.telecom_button.setChecked(disc == TELECOM)
        self._refresh_part_hint()

    def _refresh_part_hint(self) -> None:
        disc = self._selected_discipline()
        filename = pumsam_filename(disc)
        self.part_file_hint.setText(
            f"지금 불러오는 파일: 데이터베이스\\{filename}  ·  같은 자재라도 전공·품셈은 이 파일만 봅니다."
        )

    def _on_part_changed(self) -> None:
        disc = self._selected_discipline()
        self._settings.setValue("discipline", disc)
        self._refresh_part_hint()
        self.statusBar().showMessage(f"{disc} 표준품셈을 사용합니다.")
        self._append_log(f"표준품셈 파트: {disc} → {pumsam_filename(disc)}")

    def _write_startup_log(self) -> None:
        self._append_log("원본 엑셀은 읽기만 합니다. 병합 셀은 메모리에서 채웁니다.")
        self._append_log(f"저장 폴더: {self.dest_edit.text()}")
        self._append_log("왼쪽(정방향): 단가대비표 · 일위대가 · 내역서 → 내역서")
        self._append_log("오른쪽(역방향): 내역서 · 일위대가 → 단가대비표")
        self._append_log("아래(공량산출): 파트별로 나눈 내역서 → 공량산출서")
        self._append_log(f"표준품셈: {pumsam_filename(self._selected_discipline())}")
        self._append_log("노임단가는 2026년 하반기 시중노임(2026.9.1)을 넣어 두었습니다. 저장 폴더의 데이터베이스에서 고칠 수 있습니다.")

    def _refresh_filename_hint(self) -> None:
        if self._has_forward() and not self._has_reverse() and not self._has_quantity():
            name = "내역서_결과_날짜시간.xlsx"
        elif self._has_reverse() and not self._has_forward() and not self._has_quantity():
            name = "단가대비표_결과_날짜시간.xlsx"
        elif self._has_quantity() and not self._has_forward() and not self._has_reverse():
            name = "공량산출_결과_날짜시간.xlsx"
        else:
            name = "내역서_결과_날짜시간.xlsx"
        text = f"원본은 그대로 두고, {name} 새 파일로만 저장합니다."
        if not is_windows():
            text += f"  (이 환경 기본 폴더: {display_result_directory()})"
        self.confirm_note.setText(text)

    def _on_reset(self) -> None:
        self._fwd_compare_path = None
        self._fwd_ilwidae_path = None
        self._fwd_estimate_path = None
        self._rev_estimate_path = None
        self._rev_ilwidae_path = None
        self._qty_estimate_path = None
        for zone in (
            self.compare_drop,
            self.ilwidae_drop,
            self.forward_estimate_drop,
            self.drop_zone,
            self.reverse_ilwidae_drop,
            self.quantity_drop,
        ):
            zone.reset()
        self.source_edit.clear()
        self.confirm_box.setChecked(False)
        self.log.clear()
        self._write_startup_log()
        self._refresh_filename_hint()
        self._refresh_run_enabled()
        self.statusBar().showMessage("대기 — 파일을 모두 비웠습니다. 처음 켠 상태로 돌아갑니다.")
        self._append_log("새로고침: 선택한 파일을 모두 비웠습니다.")

    def _append_log(self, message: str) -> None:
        self.log.append(message)

    def _refresh_run_enabled(self) -> None:
        ready = self._has_input() and self.confirm_box.isChecked()
        self.run_button.setEnabled(ready)

    def _has_input(self) -> bool:
        return self._has_forward() or self._has_reverse() or self._has_quantity()

    def _has_forward(self) -> bool:
        return any([self._fwd_compare_path, self._fwd_ilwidae_path, self._fwd_estimate_path])

    def _has_reverse(self) -> bool:
        return any([self._rev_estimate_path, self._rev_ilwidae_path])

    def _has_quantity(self) -> bool:
        return self._qty_estimate_path is not None

    def _sync_source_edit(self) -> None:
        parts: list[str] = []
        if self._fwd_compare_path is not None:
            parts.append(f"정·단가대비표: {self._fwd_compare_path}")
        if self._fwd_ilwidae_path is not None:
            parts.append(f"정·일위대가: {self._fwd_ilwidae_path}")
        if self._fwd_estimate_path is not None:
            parts.append(f"정·내역서: {self._fwd_estimate_path}")
        if self._rev_estimate_path is not None:
            parts.append(f"역·내역서: {self._rev_estimate_path}")
        if self._rev_ilwidae_path is not None:
            parts.append(f"역·일위대가: {self._rev_ilwidae_path}")
        if self._qty_estimate_path is not None:
            parts.append(f"공량·내역서: {self._qty_estimate_path}")
        self.source_edit.setText("   |   ".join(parts))
        self._refresh_filename_hint()

    def _on_browse_dest(self) -> None:
        start = self.dest_edit.text().strip() or display_result_directory()
        chosen = QFileDialog.getExistingDirectory(self, "결과 저장 폴더 선택", start)
        if not chosen:
            return
        self.dest_edit.setText(chosen)
        self._settings.setValue("dest_dir", chosen)
        self._append_log(f"저장 폴더 변경: {chosen}")

    def _chosen_dest_dir(self) -> Path | None:
        text = self.dest_edit.text().strip()
        if not text:
            return None
        return Path(text)

    def _on_fwd_compare_dropped(self, path_text: str) -> None:
        path = Path(path_text)
        self._fwd_compare_path = path
        self.compare_drop.set_loaded(path.name)
        self._sync_source_edit()
        self.statusBar().showMessage(f"정방향 단가대비표 선택됨 (읽기 전용): {path.name}")
        self._append_log(f"정방향 단가대비표 로드 대기: {path}")
        self._refresh_run_enabled()

    def _on_fwd_ilwidae_dropped(self, path_text: str) -> None:
        path = Path(path_text)
        self._fwd_ilwidae_path = path
        self.ilwidae_drop.set_loaded(path.name)
        self._sync_source_edit()
        self.statusBar().showMessage(f"정방향 일위대가 선택됨 (읽기 전용): {path.name}")
        self._append_log(f"정방향 일위대가 로드 대기: {path}")
        self._refresh_run_enabled()

    def _on_fwd_estimate_dropped(self, path_text: str) -> None:
        path = Path(path_text)
        self._fwd_estimate_path = path
        self.forward_estimate_drop.set_loaded(path.name)
        self._sync_source_edit()
        self.statusBar().showMessage(f"정방향 내역서 선택됨 (읽기 전용): {path.name}")
        self._append_log(f"정방향 내역서 로드 대기: {path}")
        self._refresh_run_enabled()

    def _on_rev_estimate_dropped(self, path_text: str) -> None:
        path = Path(path_text)
        self._rev_estimate_path = path
        self.drop_zone.set_loaded(path.name)
        self._sync_source_edit()
        self.statusBar().showMessage(f"역방향 내역서 선택됨 (읽기 전용): {path.name}")
        self._append_log(f"역방향 내역서 로드 대기: {path}")
        self._refresh_run_enabled()

    def _on_rev_ilwidae_dropped(self, path_text: str) -> None:
        path = Path(path_text)
        self._rev_ilwidae_path = path
        self.reverse_ilwidae_drop.set_loaded(path.name)
        self._sync_source_edit()
        self.statusBar().showMessage(f"역방향 일위대가 선택됨 (읽기 전용): {path.name}")
        self._append_log(f"역방향 일위대가 로드 대기: {path}")
        self._refresh_run_enabled()

    def _on_unit_price_dropped(self, path_text: str) -> None:
        self._on_fwd_compare_dropped(path_text)

    def _on_ilwidae_dropped(self, path_text: str) -> None:
        self._on_fwd_ilwidae_dropped(path_text)

    def _on_file_dropped(self, path_text: str) -> None:
        self._on_rev_estimate_dropped(path_text)

    def _on_qty_estimate_dropped(self, path_text: str) -> None:
        path = Path(path_text)
        self._qty_estimate_path = path
        self.quantity_drop.set_loaded(path.name)
        self._sync_source_edit()
        self.statusBar().showMessage(f"공량산출용 내역서 선택됨 (읽기 전용): {path.name}")
        self._append_log(f"공량산출 내역서 로드 대기: {path}")
        self._refresh_run_enabled()

    def _on_run(self) -> None:
        if not self._has_input():
            QMessageBox.warning(self, "파일 없음", "단가대비표 또는 내역서를 먼저 놓아 주세요.")
            return
        if not self.confirm_box.isChecked():
            QMessageBox.warning(self, "저장 경로 미확인", "저장 확인란을 선택해 주세요.")
            return
        lane_count = sum(
            [self._has_forward(), self._has_reverse(), self._has_quantity()]
        )
        if lane_count > 1:
            QMessageBox.warning(
                self,
                "칸이 여러 개 선택됨",
                "정방향·역방향·공량산출에 파일이 같이 들어 있습니다.\n"
                "새로고침으로 비운 뒤, 지금 만들 칸에만 엑셀을 놓아 주세요.",
            )
            return

        dest_dir = self._chosen_dest_dir()
        if dest_dir is not None:
            self._settings.setValue("dest_dir", str(dest_dir))
            if "onedrive" in str(dest_dir).lower():
                self._append_log("알림: 선택한 폴더가 OneDrive 경로로 보입니다. 가능하면 로컬 폴더를 쓰세요.")

        self.run_button.setEnabled(False)
        self.statusBar().showMessage("산출 파일을 생성하는 중…")
        discipline = self._selected_discipline()
        self._append_log(f"원본 읽기 전용 · {pumsam_filename(discipline)} · 노임단가 결합 · 결과 엑셀 생성")

        if self._has_forward():
            mode = "forward"
            unit_price_path = self._fwd_compare_path
            ilwidae_path = self._fwd_ilwidae_path
            estimate_path = self._fwd_estimate_path
        elif self._has_reverse():
            mode = "reverse"
            unit_price_path = None
            ilwidae_path = self._rev_ilwidae_path
            estimate_path = self._rev_estimate_path
        else:
            mode = "quantity"
            unit_price_path = None
            ilwidae_path = None
            estimate_path = self._qty_estimate_path

        try:
            dest = save_result_workbook(
                dest_dir=dest_dir,
                unit_price_path=unit_price_path,
                ilwidae_path=ilwidae_path,
                estimate_path=estimate_path,
                discipline=discipline,
                mode=mode,
            )
        except ResultDirectoryError as exc:
            self._append_log(f"저장 폴더 오류: {exc}")
            self.statusBar().showMessage("저장 폴더 오류")
            QMessageBox.critical(self, "저장 폴더 오류", str(exc))
            self._refresh_run_enabled()
            return
        except Exception as exc:  # noqa: BLE001 — GUI에서 사용자 메시지로 보여 준다.
            self._append_log(f"실패: {exc}")
            self.statusBar().showMessage("실패")
            QMessageBox.critical(self, "처리 실패", str(exc))
            self._refresh_run_enabled()
            return

        originals = [
            p
            for p in (
                self._fwd_compare_path,
                self._fwd_ilwidae_path,
                self._fwd_estimate_path,
                self._rev_estimate_path,
                self._rev_ilwidae_path,
                self._qty_estimate_path,
            )
            if p is not None
        ]
        for original in originals:
            self._append_log(f"원본 보존 확인: {original}")
        self._append_log(f"새 파일 저장: {dest}")
        self.statusBar().showMessage(f"저장 완료 — {dest.name}")
        if mode == "forward":
            sheets_line = f"{COMPARE_SHEET_NAME} · {ILWIDAE_SHEET_NAME} · {ESTIMATE_SHEET_NAME}"
        elif mode == "reverse":
            sheets_line = f"{COMPARE_SHEET_NAME} · {ESTIMATE_SHEET_NAME}"
        else:
            sheets_line = f"{ESTIMATE_SHEET_NAME} · {QUANTITY_SHEET_NAME}"
        QMessageBox.information(
            self,
            "저장 완료",
            (
                "원본은 그대로 두었습니다.\n\n"
                f"결과: {dest}\n\n"
                f"{sheets_line}\n"
                f"참고 시트: {PUMSAM_SHEET_NAME}, {WAGES_SHEET_NAME}\n"
                f"사용한 표준품셈: {pumsam_filename(discipline)}\n"
                "표준품셈·노임단가는 저장 폴더의 데이터베이스에서 고칠 수 있습니다."
            ),
        )
        self._refresh_run_enabled()
