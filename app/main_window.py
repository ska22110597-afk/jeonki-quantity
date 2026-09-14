"""전기공사 공량 산출 — 1단계 GUI 뼈대."""

from __future__ import annotations

from pathlib import Path

from PyQt6.QtCore import Qt
from PyQt6.QtWidgets import (
    QCheckBox,
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
from app.excel_io import QUANTITY_SHEET_NAME, UNIT_PRICE_SHEET_NAME, save_result_workbook
from app.paths import WINDOWS_RESULT_DIR, display_result_directory, is_windows

APP_STYLESHEET = """
QMainWindow, QWidget#root {
    background: #F4F6F8;
    color: #1B2430;
    font-family: "Malgun Gothic", "Apple SD Gothic Neo", "Noto Sans CJK KR", sans-serif;
    font-size: 13px;
}
QFrame#hero {
    background: #16324F;
    border: none;
}
QLabel#appTitle {
    color: #FFFFFF;
    font-size: 20px;
    font-weight: 700;
}
QLabel#appSubtitle {
    color: #C5D4E3;
    font-size: 12px;
}
QFrame#card {
    background: #FFFFFF;
    border: 1px solid #D7DEE6;
    border-radius: 10px;
}
QLabel#sectionLabel {
    color: #4A5A6A;
    font-size: 11px;
    font-weight: 700;
}
QLineEdit {
    background: #F7F9FB;
    border: 1px solid #D7DEE6;
    border-radius: 6px;
    padding: 8px 10px;
    color: #1B2430;
    selection-background-color: #16324F;
}
QLineEdit:read-only {
    color: #2C3E50;
}
QCheckBox {
    color: #1B2430;
    spacing: 8px;
}
QCheckBox::indicator {
    width: 16px;
    height: 16px;
}
QPushButton#runButton {
    background: #C45911;
    color: #FFFFFF;
    border: none;
    border-radius: 8px;
    padding: 12px 18px;
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
QTextEdit#log {
    background: #0F1C2A;
    color: #D6E4F0;
    border: none;
    border-radius: 8px;
    padding: 10px;
    font-family: "Consolas", "D2Coding", monospace;
    font-size: 12px;
}
QFrame#dropZone {
    background: #F8FBFF;
    border: 2px dashed #7F98B0;
    border-radius: 10px;
}
QFrame#dropZone[hover="true"] {
    background: #E8F1FA;
    border: 2px dashed #16324F;
}
QFrame#dropZone[loaded="true"] {
    background: #EEF7F0;
    border: 2px solid #2E7D4F;
}
QLabel#dropTitle {
    font-size: 16px;
    font-weight: 700;
    color: #16324F;
}
QLabel#dropHint {
    color: #5C6F82;
    font-size: 12px;
}
QLabel#badge {
    background: #1F4E79;
    color: #FFFFFF;
    border-radius: 4px;
    padding: 2px 8px;
    font-size: 10px;
    font-weight: 700;
}
QStatusBar {
    background: #E9EEF3;
    color: #4A5A6A;
}
"""


class MainWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("전기공사 공량 산출")
        self.setMinimumSize(780, 680)
        self.resize(860, 740)
        self.setStyleSheet(APP_STYLESHEET)

        self._source_path: Path | None = None
        self._build_ui()
        self._refresh_run_enabled()

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
        hero_layout.setContentsMargins(28, 20, 28, 20)
        title = QLabel("전기공사 견적 · 공량 산출")
        title.setObjectName("appTitle")
        subtitle = QLabel("단가대비표를 읽기 전용으로 불러와 C드라이브 로컬 폴더에 새 결과 파일만 생성합니다.")
        subtitle.setObjectName("appSubtitle")
        hero_layout.addWidget(title)
        hero_layout.addWidget(subtitle)
        outer.addWidget(hero)

        body = QWidget()
        body_layout = QVBoxLayout(body)
        body_layout.setContentsMargins(24, 20, 24, 16)
        body_layout.setSpacing(14)

        self.drop_zone = DropZone()
        self.drop_zone.file_dropped.connect(self._on_file_dropped)
        body_layout.addWidget(self.drop_zone)

        path_card = QFrame()
        path_card.setObjectName("card")
        path_layout = QVBoxLayout(path_card)
        path_layout.setContentsMargins(16, 14, 16, 14)
        path_layout.setSpacing(10)

        source_row = QVBoxLayout()
        source_label = QLabel("원본 파일 경로 (읽기 전용 · 수정·덮어쓰기 없음)")
        source_label.setObjectName("sectionLabel")
        self.source_edit = QLineEdit()
        self.source_edit.setReadOnly(True)
        self.source_edit.setPlaceholderText("아직 단가대비표가 없습니다. 위에서 엑셀을 놓아 주세요.")
        source_row.addWidget(source_label)
        source_row.addWidget(self.source_edit)
        path_layout.addLayout(source_row)

        dest_row = QVBoxLayout()
        dest_head = QHBoxLayout()
        dest_label = QLabel("C드라이브 저장 경로")
        dest_label.setObjectName("sectionLabel")
        badge = QLabel("OneDrive 제외")
        badge.setObjectName("badge")
        badge.setAlignment(Qt.AlignmentFlag.AlignCenter)
        dest_head.addWidget(dest_label)
        dest_head.addStretch(1)
        dest_head.addWidget(badge)
        self.dest_edit = QLineEdit(display_result_directory())
        self.dest_edit.setReadOnly(True)
        dest_row.addLayout(dest_head)
        dest_row.addWidget(self.dest_edit)
        path_layout.addLayout(dest_row)

        confirm_text = (
            f"결과를 {WINDOWS_RESULT_DIR} 규칙의 로컬 폴더에 "
            "단가대비_공량산출_결과_YYYYMMDD_HHMMSS.xlsx 로 새로 저장합니다."
        )
        if not is_windows():
            confirm_text += f" (이 환경 실제 저장: {display_result_directory()})"
        self.confirm_box = QCheckBox(confirm_text)
        self.confirm_box.toggled.connect(self._refresh_run_enabled)
        path_layout.addWidget(self.confirm_box)
        body_layout.addWidget(path_card)

        self.run_button = QPushButton("공량 산출 및 C드라이브 저장")
        self.run_button.setObjectName("runButton")
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
        status.showMessage("대기 — 단가대비표를 드롭한 뒤 C드라이브 저장 경로를 확인하세요.")
        self.setStatusBar(status)
        self._append_log("원본 엑셀은 읽기 전용으로만 엽니다. 저장은 새 타임스탬프 파일만 생성합니다.")
        self._append_log(f"지정 저장 폴더: {display_result_directory()}")
        self._append_log(f"결과 시트: {UNIT_PRICE_SHEET_NAME}, {QUANTITY_SHEET_NAME}")

    def _append_log(self, message: str) -> None:
        self.log.append(message)

    def _refresh_run_enabled(self) -> None:
        ready = self._source_path is not None and self.confirm_box.isChecked()
        self.run_button.setEnabled(ready)

    def _on_file_dropped(self, path_text: str) -> None:
        path = Path(path_text)
        self._source_path = path
        self.source_edit.setText(str(path))
        self.drop_zone.set_loaded(path.name)
        self.statusBar().showMessage(f"원본 선택됨 (읽기 전용): {path.name}")
        self._append_log(f"원본 로드 대기: {path}")
        self._refresh_run_enabled()

    def _on_run(self) -> None:
        if self._source_path is None:
            QMessageBox.warning(self, "원본 없음", "단가대비표 엑셀을 먼저 놓아 주세요.")
            return
        if not self.confirm_box.isChecked():
            QMessageBox.warning(self, "저장 경로 미확인", "C드라이브 저장 경로 확인란을 선택해 주세요.")
            return

        source = self._source_path
        self.run_button.setEnabled(False)
        self.statusBar().showMessage("공량 산출 파일을 생성하는 중…")
        self._append_log("원본 읽기 전용 접근 시작")

        try:
            dest = save_result_workbook(source)
        except Exception as exc:  # noqa: BLE001 — GUI에서 사용자 메시지로 보여 준다.
            self._append_log(f"실패: {exc}")
            self.statusBar().showMessage("실패")
            QMessageBox.critical(self, "처리 실패", str(exc))
            self._refresh_run_enabled()
            return

        self._append_log(f"원본 보존 확인: {source}")
        self._append_log(f"새 파일 저장: {dest}")
        self.statusBar().showMessage(f"저장 완료 — {dest.name}")
        QMessageBox.information(
            self,
            "저장 완료",
            (
                "원본은 그대로 두었습니다.\n\n"
                f"원본: {source}\n"
                f"결과: {dest}\n\n"
                f"시트1 {UNIT_PRICE_SHEET_NAME} — 원본 데이터 정리\n"
                f"시트2 {QUANTITY_SHEET_NAME} — 공량 산출 기본 구조"
            ),
        )
        self._refresh_run_enabled()
