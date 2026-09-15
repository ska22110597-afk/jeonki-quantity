"""전기공사 공량 산출 — 1단계 GUI 뼈대."""

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
from app.wages import WAGES_SHEET_NAME

APP_STYLESHEET = """
QMainWindow, QWidget#root {
    background: #F4F6F8;
    color: #1B2430;
    font-family: "Malgun Gothic", "Apple SD Gothic Neo", "Noto Sans CJK KR", sans-serif;
    font-size: 14px;
}
QFrame#hero {
    background: #16324F;
    border: none;
}
QLabel#appTitle {
    color: #FFFFFF;
    font-size: 22px;
    font-weight: 700;
}
QLabel#appSubtitle {
    color: #C5D4E3;
    font-size: 13px;
}
QFrame#card {
    background: #FFFFFF;
    border: 1px solid #D7DEE6;
    border-radius: 10px;
}
QLabel#sectionLabel {
    color: #4A5A6A;
    font-size: 12px;
    font-weight: 700;
}
QLineEdit {
    background: #F7F9FB;
    border: 1px solid #D7DEE6;
    border-radius: 8px;
    padding: 12px 14px;
    min-height: 28px;
    color: #1B2430;
    font-size: 14px;
    selection-background-color: #16324F;
}
QLineEdit#pathEdit {
    min-height: 36px;
    font-size: 14px;
    padding: 14px 16px;
}
QLineEdit:read-only {
    color: #2C3E50;
}
QCheckBox {
    color: #1B2430;
    spacing: 8px;
}
QCheckBox::indicator {
    width: 18px;
    height: 18px;
}
QPushButton#runButton {
    background: #C45911;
    color: #FFFFFF;
    border: none;
    border-radius: 8px;
    padding: 16px 22px;
    font-size: 16px;
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
    background: #16324F;
    color: #FFFFFF;
    border: none;
    border-radius: 8px;
    padding: 12px 18px;
    font-size: 13px;
    font-weight: 700;
    min-height: 44px;
}
QPushButton#browseButton:hover {
    background: #1F4E79;
}
QTextEdit#log {
    background: #0F1C2A;
    color: #D6E4F0;
    border: none;
    border-radius: 8px;
    padding: 12px;
    font-family: "Consolas", "D2Coding", monospace;
    font-size: 13px;
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
    font-size: 17px;
    font-weight: 700;
    color: #16324F;
}
QLabel#dropHint {
    color: #5C6F82;
    font-size: 13px;
}
QLabel#badge {
    background: #1F4E79;
    color: #FFFFFF;
    border-radius: 4px;
    padding: 4px 10px;
    font-size: 11px;
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
        self.setMinimumSize(1000, 900)
        self.resize(1100, 980)
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
        hero_layout.setContentsMargins(32, 24, 32, 24)
        title = QLabel("전기공사 견적 · 공량 산출")
        title.setObjectName("appTitle")
        subtitle = QLabel(
            "단가대비표를 놓으면 일위대가·내역서·공량산출서까지 만듭니다. "
            "일위대가나 내역서만 놓아도 그다음 단계를 이어서 작성합니다. "
            "저장 폴더는 아래에서 고를 수 있습니다."
        )
        subtitle.setObjectName("appSubtitle")
        subtitle.setWordWrap(True)
        hero_layout.addWidget(title)
        hero_layout.addWidget(subtitle)
        outer.addWidget(hero)

        body = QWidget()
        body_layout = QVBoxLayout(body)
        body_layout.setContentsMargins(28, 22, 28, 18)
        body_layout.setSpacing(16)

        self.compare_drop = DropZone(
            title="1) 단가대비표",
            hint="물량 품목의 자재 단가  ·  놓으면 일위대가부터 공량산출서까지 작성",
            dialog_title="단가대비표 엑셀 선택",
        )
        self.compare_drop.setMinimumHeight(88)
        self.compare_drop.file_dropped.connect(self._on_unit_price_dropped)
        body_layout.addWidget(self.compare_drop)

        self.ilwidae_drop = DropZone(
            title="2) 일위대가",
            hint="호표가 있는 일위대가  ·  선택. 단가대비표가 있으면 새로 작성합니다",
            dialog_title="일위대가 엑셀 선택",
        )
        self.ilwidae_drop.setMinimumHeight(88)
        self.ilwidae_drop.file_dropped.connect(self._on_ilwidae_dropped)
        body_layout.addWidget(self.ilwidae_drop)

        self.drop_zone = DropZone(
            title="3) 내역서",
            hint="이미 있는 내역서만 놓고 공량산출서를 만들 때도 사용",
            dialog_title="내역서 엑셀 선택",
        )
        self.drop_zone.setMinimumHeight(88)
        self.drop_zone.file_dropped.connect(self._on_file_dropped)
        body_layout.addWidget(self.drop_zone)

        path_card = QFrame()
        path_card.setObjectName("card")
        path_layout = QVBoxLayout(path_card)
        path_layout.setContentsMargins(18, 16, 18, 16)
        path_layout.setSpacing(12)

        source_row = QVBoxLayout()
        source_label = QLabel("선택한 파일 경로 (읽기 전용 · 원본은 수정하지 않습니다)")
        source_label.setObjectName("sectionLabel")
        self.source_edit = QLineEdit()
        self.source_edit.setObjectName("pathEdit")
        self.source_edit.setReadOnly(True)
        self.source_edit.setMinimumHeight(48)
        self.source_edit.setPlaceholderText("아직 파일이 없습니다. 단가대비표·일위대가·내역서 중 하나를 놓아 주세요.")
        source_row.addWidget(source_label)
        source_row.addWidget(self.source_edit)
        path_layout.addLayout(source_row)

        dest_row = QVBoxLayout()
        dest_head = QHBoxLayout()
        dest_label = QLabel("결과 저장 폴더")
        dest_label.setObjectName("sectionLabel")
        badge = QLabel("가능하면 OneDrive 제외")
        badge.setObjectName("badge")
        badge.setAlignment(Qt.AlignmentFlag.AlignCenter)
        dest_head.addWidget(dest_label)
        dest_head.addStretch(1)
        dest_head.addWidget(badge)

        dest_pick = QHBoxLayout()
        dest_pick.setSpacing(10)
        self.dest_edit = QLineEdit(self._saved_dest_dir())
        self.dest_edit.setObjectName("pathEdit")
        self.dest_edit.setReadOnly(False)
        self.dest_edit.setMinimumHeight(48)
        self.dest_edit.setPlaceholderText(str(WINDOWS_RESULT_DIR))
        self.browse_button = QPushButton("폴더 찾기")
        self.browse_button.setObjectName("browseButton")
        self.browse_button.setCursor(Qt.CursorShape.PointingHandCursor)
        self.browse_button.clicked.connect(self._on_browse_dest)
        dest_pick.addWidget(self.dest_edit, 1)
        dest_pick.addWidget(self.browse_button)

        dest_row.addLayout(dest_head)
        dest_row.addLayout(dest_pick)
        path_layout.addLayout(dest_row)

        confirm_text = (
            "지정한 폴더에 공량산출_결과_YYYYMMDD_HHMMSS.xlsx 로 새로 저장합니다. "
            "원본 내역서는 그대로 둡니다."
        )
        if not is_windows():
            confirm_text += f" (이 환경 기본 폴더: {display_result_directory()})"
        self.confirm_box = QCheckBox(confirm_text)
        self.confirm_box.toggled.connect(self._refresh_run_enabled)
        path_layout.addWidget(self.confirm_box)
        body_layout.addWidget(path_card)

        self.run_button = QPushButton("산출 및 저장")
        self.run_button.setObjectName("runButton")
        self.run_button.setMinimumHeight(52)
        self.run_button.setCursor(Qt.CursorShape.PointingHandCursor)
        self.run_button.clicked.connect(self._on_run)
        body_layout.addWidget(self.run_button)

        log_label = QLabel("처리 기록")
        log_label.setObjectName("sectionLabel")
        self.log = QTextEdit()
        self.log.setObjectName("log")
        self.log.setReadOnly(True)
        self.log.setMinimumHeight(168)
        body_layout.addWidget(log_label)
        body_layout.addWidget(self.log, 1)

        outer.addWidget(body, 1)

        status = QStatusBar()
        status.showMessage("대기 — 단가대비표·일위대가·내역서 중 하나를 놓고 저장 폴더를 확인하세요.")
        self.setStatusBar(status)
        self._append_log("원본 엑셀은 읽기만 합니다. 병합 셀은 메모리에서 채웁니다.")
        self._append_log(f"저장 폴더: {self.dest_edit.text()}")
        self._append_log(
            "단가대비표를 놓으면 일위대가 → 내역서 → 공량산출서 순으로 새 파일을 만듭니다."
        )
        self._append_log(
            "표준품셈과 노임단가는 프로그램 data 폴더에 들어 있고, "
            "저장 폴더의 데이터베이스 안에서 계속 고칠 수 있습니다."
        )

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
            QMessageBox.warning(self, "파일 없음", "단가대비표, 일위대가, 내역서 중 하나를 먼저 놓아 주세요.")
            return
        if not self.confirm_box.isChecked():
            QMessageBox.warning(self, "저장 경로 미확인", "저장 폴더 확인란을 선택해 주세요.")
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
                f"{COMPARE_SHEET_NAME} → {ILWIDAE_SHEET_NAME} → {ESTIMATE_SHEET_NAME} → {QUANTITY_SHEET_NAME}\n"
                f"참고 시트: {PUMSAM_SHEET_NAME}, {WAGES_SHEET_NAME}\n"
                "표준품셈·노임단가는 저장 폴더의 데이터베이스에서 고칠 수 있습니다."
            ),
        )
        self._refresh_run_enabled()
