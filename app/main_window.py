"""제작자_박남석 자동 내역서식 프로그램 화면."""

from __future__ import annotations

from pathlib import Path

from PyQt6.QtCore import Qt, QSettings
from PyQt6.QtWidgets import (
    QCheckBox,
    QFileDialog,
    QFrame,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QStatusBar,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from app.drop_zone import DropZone
from app.excel_io import (
    COMPARE_SHEET_NAME,
    ESTIMATE_SHEET_NAME,
    ILWIDAE_SHEET_NAME,
    QUANTITY_SHEET_NAME,
    save_result_workbook,
)
from app.paths import ResultDirectoryError, WINDOWS_RESULT_DIR, display_result_directory, is_windows
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
    background: #E7EEF4;
    border: 1px solid #B7C7D4;
    border-radius: 12px;
}
QFrame#laneReverse {
    background: #F3EBE8;
    border: 1px solid #D2B8B1;
    border-radius: 12px;
}
QLabel#laneTitle {
    font-size: 13px;
    font-weight: 700;
    letter-spacing: 0.3px;
}
QLabel#laneForwardTitle {
    color: #3E5B70;
}
QLabel#laneReverseTitle {
    color: #7A534C;
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
    padding: 10px 14px;
    min-height: 26px;
    color: #243040;
    font-size: 13px;
    selection-background-color: #1C2B3A;
}
QLineEdit#pathEdit {
    min-height: 34px;
    font-size: 13px;
    padding: 12px 14px;
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
    padding: 12px 18px;
    font-size: 13px;
    font-weight: 700;
    min-height: 44px;
    min-width: 96px;
}
QPushButton#browseButton:hover {
    background: #2A4054;
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
    background: #F4F8FB;
    border: 2px dashed #8AA0B3;
    border-radius: 10px;
}
QFrame#dropZoneForward[hover="true"] {
    background: #E4EDF4;
    border: 2px dashed #5D7A90;
}
QFrame#dropZoneForward[loaded="true"] {
    background: #E7F2EA;
    border: 2px solid #4F7F62;
}
QFrame#dropZoneReverse {
    background: #F8F3F1;
    border: 2px dashed #C4A199;
    border-radius: 10px;
}
QFrame#dropZoneReverse[hover="true"] {
    background: #F0E4E0;
    border: 2px dashed #A0756C;
}
QFrame#dropZoneReverse[loaded="true"] {
    background: #E7F2EA;
    border: 2px solid #4F7F62;
}
QLabel#dropTitle {
    font-size: 16px;
    font-weight: 700;
}
QFrame#dropZoneForward QLabel#dropTitle {
    color: #3E5B70;
}
QFrame#dropZoneReverse QLabel#dropTitle {
    color: #7A534C;
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
QStatusBar {
    background: #E8EDF2;
    color: #5A6A78;
}
"""


class MainWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle(APP_TITLE)
        self.setMinimumSize(1080, 860)
        self.resize(1180, 920)
        self.setStyleSheet(APP_STYLESHEET)

        self._settings = QSettings("전기공사공량산출", "GongryangCalc")
        self._unit_price_path: Path | None = None
        self._ilwidae_path: Path | None = None
        self._source_path: Path | None = None
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

        lanes = QHBoxLayout()
        lanes.setSpacing(14)

        forward_lane = QFrame()
        forward_lane.setObjectName("laneForward")
        forward_layout = QVBoxLayout(forward_lane)
        forward_layout.setContentsMargins(16, 14, 16, 16)
        forward_layout.setSpacing(10)
        forward_title = QLabel("정방향  ·  단가대비표 → 일위대가 → 내역서 → 공량")
        forward_title.setObjectName("laneTitle")
        forward_title.setProperty("class", "laneForwardTitle")
        forward_title.setStyleSheet("color: #3E5B70;")
        forward_layout.addWidget(forward_title)

        self.compare_drop = DropZone(
            title="단가대비표",
            hint="물량 품목의 자재 단가  ·  놓으면 이후 시트를 작성",
            dialog_title="단가대비표 엑셀 선택",
            tone="forward",
        )
        self.compare_drop.setMinimumHeight(86)
        self.compare_drop.file_dropped.connect(self._on_unit_price_dropped)
        forward_layout.addWidget(self.compare_drop)

        self.ilwidae_drop = DropZone(
            title="일위대가 (선택)",
            hint="이미 만든 호표가 있을 때만 놓습니다",
            dialog_title="일위대가 엑셀 선택",
            tone="forward",
        )
        self.ilwidae_drop.setMinimumHeight(86)
        self.ilwidae_drop.file_dropped.connect(self._on_ilwidae_dropped)
        forward_layout.addWidget(self.ilwidae_drop)
        lanes.addWidget(forward_lane, 1)

        reverse_lane = QFrame()
        reverse_lane.setObjectName("laneReverse")
        reverse_layout = QVBoxLayout(reverse_lane)
        reverse_layout.setContentsMargins(16, 14, 16, 16)
        reverse_layout.setSpacing(10)
        reverse_title = QLabel("역방향  ·  내역서 → 단가대비표")
        reverse_title.setObjectName("laneTitle")
        reverse_title.setStyleSheet("color: #7A534C;")
        reverse_layout.addWidget(reverse_title)

        self.drop_zone = DropZone(
            title="내역서",
            hint="이미 있는 내역서를 놓으면 단가대비표와 공량산출서를 만듭니다",
            dialog_title="내역서 엑셀 선택",
            tone="reverse",
        )
        self.drop_zone.setMinimumHeight(188)
        self.drop_zone.file_dropped.connect(self._on_file_dropped)
        reverse_layout.addWidget(self.drop_zone, 1)
        lanes.addWidget(reverse_lane, 1)
        body_layout.addLayout(lanes)

        path_card = QFrame()
        path_card.setObjectName("card")
        path_layout = QVBoxLayout(path_card)
        path_layout.setContentsMargins(22, 18, 22, 18)
        path_layout.setSpacing(14)

        source_label = QLabel("선택한 파일  (읽기 전용 · 원본은 수정하지 않습니다)")
        source_label.setObjectName("sectionLabel")
        self.source_edit = QLineEdit()
        self.source_edit.setObjectName("pathEdit")
        self.source_edit.setReadOnly(True)
        self.source_edit.setMinimumHeight(46)
        self.source_edit.setPlaceholderText("아직 파일이 없습니다. 왼쪽 또는 오른쪽에 엑셀을 놓아 주세요.")
        path_layout.addWidget(source_label)
        path_layout.addWidget(self.source_edit)

        dest_label = QLabel("결과 저장 폴더")
        dest_label.setObjectName("sectionLabel")
        dest_head = QHBoxLayout()
        dest_head.setContentsMargins(0, 4, 0, 0)
        dest_head.addWidget(dest_label)
        dest_head.addStretch(1)
        badge = QLabel("가능하면 OneDrive 제외")
        badge.setObjectName("badge")
        badge.setAlignment(Qt.AlignmentFlag.AlignCenter)
        dest_head.addWidget(badge)
        path_layout.addLayout(dest_head)

        dest_pick = QHBoxLayout()
        dest_pick.setSpacing(10)
        self.dest_edit = QLineEdit(self._saved_dest_dir())
        self.dest_edit.setObjectName("pathEdit")
        self.dest_edit.setReadOnly(False)
        self.dest_edit.setMinimumHeight(46)
        self.dest_edit.setPlaceholderText(str(WINDOWS_RESULT_DIR))
        self.browse_button = QPushButton("폴더 찾기")
        self.browse_button.setObjectName("browseButton")
        self.browse_button.setCursor(Qt.CursorShape.PointingHandCursor)
        self.browse_button.clicked.connect(self._on_browse_dest)
        dest_pick.addWidget(self.dest_edit, 1)
        dest_pick.addWidget(self.browse_button)
        path_layout.addLayout(dest_pick)

        confirm_row = QHBoxLayout()
        confirm_row.setSpacing(10)
        confirm_row.setContentsMargins(2, 6, 2, 2)
        self.confirm_box = QCheckBox("이 폴더에 새 파일로 저장")
        self.confirm_box.toggled.connect(self._refresh_run_enabled)
        confirm_note = QLabel("파일명 공량산출_결과_날짜시간.xlsx  ·  원본은 그대로 둡니다.")
        confirm_note.setObjectName("confirmNote")
        confirm_note.setWordWrap(True)
        confirm_row.addWidget(self.confirm_box, 0, Qt.AlignmentFlag.AlignTop)
        confirm_row.addWidget(confirm_note, 1)
        if not is_windows():
            confirm_note.setText(
                confirm_note.text() + f"  (이 환경 기본 폴더: {display_result_directory()})"
            )
        path_layout.addLayout(confirm_row)
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
        self.log.setMinimumHeight(140)
        body_layout.addWidget(log_label)
        body_layout.addWidget(self.log, 1)

        outer.addWidget(body, 1)

        status = QStatusBar()
        status.showMessage("대기 — 정방향 또는 역방향 칸에 엑셀을 놓고 저장 폴더를 확인하세요.")
        self.setStatusBar(status)
        self._append_log("원본 엑셀은 읽기만 합니다. 병합 셀은 메모리에서 채웁니다.")
        self._append_log(f"저장 폴더: {self.dest_edit.text()}")
        self._append_log("왼쪽(정방향): 단가대비표 → 일위대가 → 내역서 → 공량산출서")
        self._append_log("오른쪽(역방향): 내역서 → 단가대비표 · 공량산출서")
        self._append_log("노임단가는 2026년 하반기 시중노임(2026.9.1)을 넣어 두었습니다. 저장 폴더의 데이터베이스에서 고칠 수 있습니다.")

    def _append_log(self, message: str) -> None:
        self.log.append(message)

    def _refresh_run_enabled(self) -> None:
        ready = self._has_input() and self.confirm_box.isChecked()
        self.run_button.setEnabled(ready)

    def _has_input(self) -> bool:
        return any([self._unit_price_path, self._ilwidae_path, self._source_path])

    def _sync_source_edit(self) -> None:
        parts: list[str] = []
        if self._unit_price_path is not None:
            parts.append(f"단가대비표: {self._unit_price_path}")
        if self._ilwidae_path is not None:
            parts.append(f"일위대가: {self._ilwidae_path}")
        if self._source_path is not None:
            parts.append(f"내역서: {self._source_path}")
        self.source_edit.setText("   |   ".join(parts))

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

    def _on_unit_price_dropped(self, path_text: str) -> None:
        path = Path(path_text)
        self._unit_price_path = path
        self.compare_drop.set_loaded(path.name)
        self._sync_source_edit()
        self.statusBar().showMessage(f"단가대비표 선택됨 (읽기 전용): {path.name}")
        self._append_log(f"단가대비표 로드 대기: {path}")
        self._refresh_run_enabled()

    def _on_ilwidae_dropped(self, path_text: str) -> None:
        path = Path(path_text)
        self._ilwidae_path = path
        self.ilwidae_drop.set_loaded(path.name)
        self._sync_source_edit()
        self.statusBar().showMessage(f"일위대가 선택됨 (읽기 전용): {path.name}")
        self._append_log(f"일위대가 로드 대기: {path}")
        self._refresh_run_enabled()

    def _on_file_dropped(self, path_text: str) -> None:
        path = Path(path_text)
        self._source_path = path
        self.drop_zone.set_loaded(path.name)
        self._sync_source_edit()
        self.statusBar().showMessage(f"내역서 선택됨 (읽기 전용): {path.name}")
        self._append_log(f"내역서 로드 대기: {path}")
        self._refresh_run_enabled()

    def _on_run(self) -> None:
        if not self._has_input():
            QMessageBox.warning(self, "파일 없음", "단가대비표 또는 내역서를 먼저 놓아 주세요.")
            return
        if not self.confirm_box.isChecked():
            QMessageBox.warning(self, "저장 경로 미확인", "저장 확인란을 선택해 주세요.")
            return

        dest_dir = self._chosen_dest_dir()
        if dest_dir is not None:
            self._settings.setValue("dest_dir", str(dest_dir))
            if "onedrive" in str(dest_dir).lower():
                self._append_log("알림: 선택한 폴더가 OneDrive 경로로 보입니다. 가능하면 로컬 폴더를 쓰세요.")

        self.run_button.setEnabled(False)
        self.statusBar().showMessage("산출 파일을 생성하는 중…")
        self._append_log("원본 읽기 전용 · 표준품셈·노임단가 결합 · 결과 엑셀 생성")

        try:
            dest = save_result_workbook(
                dest_dir=dest_dir,
                unit_price_path=self._unit_price_path,
                ilwidae_path=self._ilwidae_path,
                estimate_path=self._source_path,
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

        originals = [p for p in (self._unit_price_path, self._ilwidae_path, self._source_path) if p is not None]
        for original in originals:
            self._append_log(f"원본 보존 확인: {original}")
        self._append_log(f"새 파일 저장: {dest}")
        self.statusBar().showMessage(f"저장 완료 — {dest.name}")
        QMessageBox.information(
            self,
            "저장 완료",
            (
                "원본은 그대로 두었습니다.\n\n"
                f"결과: {dest}\n\n"
                f"{COMPARE_SHEET_NAME} · {ILWIDAE_SHEET_NAME} · {ESTIMATE_SHEET_NAME} · {QUANTITY_SHEET_NAME}\n"
                f"참고 시트: {PUMSAM_SHEET_NAME}, {WAGES_SHEET_NAME}\n"
                "표준품셈·노임단가는 저장 폴더의 데이터베이스에서 고칠 수 있습니다."
            ),
        )
        self._refresh_run_enabled()
