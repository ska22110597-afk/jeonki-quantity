"""내역서 작성 화면: 단가대비표 → 일위대가목록 → 체크박스로 고른 품목 → 공정별 내역서."""

from __future__ import annotations

from pathlib import Path
from typing import Any, Callable

from PyQt6.QtCore import Qt, QThread, pyqtSignal
from PyQt6.QtGui import QIcon
from PyQt6.QtWidgets import (
    QAbstractItemView,
    QApplication,
    QButtonGroup,
    QFileDialog,
    QFrame,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QListWidget,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QSizePolicy,
    QStatusBar,
    QTableWidget,
    QTableWidgetItem,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from app.discipline import ELECTRIC, TELECOM, normalize_discipline, pumsam_filename
from app.drop_zone import DropZone
from app.estimate_build import (
    EstimateLine,
    EstimateSection,
    EstimateSession,
    generate_session,
    save_session,
)
from app.main_window import (
    APP_STYLESHEET,
    BusyDialog,
    CaptionBar,
    DoneDialog,
    MainWindow as LegacyWindow,
    PaperRoot,
    _load_settings,
)
from app.paths import ResultDirectoryError, WINDOWS_RESULT_DIR, app_icon_path, display_result_directory
from app.version import APP_CONTACT, APP_MAKER, APP_NOTICE, APP_TITLE, DROP_HINT

EXTRA_STYLESHEET = """
QTableWidget#itemTable {
    background: rgba(255, 252, 246, 0.96);
    alternate-background-color: #F7F0E2;
    gridline-color: #E4D6BE;
    border: 1px solid #D4B896;
    border-radius: 6px;
    selection-background-color: #E8D9BC;
    selection-color: #2C281F;
    font-size: 13px;
}
QHeaderView::section {
    background: #5C4A32;
    color: #F7F3EA;
    border: none;
    border-right: 1px solid #7A6A52;
    padding: 6px 4px;
    font-weight: 700;
}
QTableWidget::indicator {
    width: 16px;
    height: 16px;
    border: 1px solid #8A7349;
    border-radius: 3px;
    background: #FFFcf7;
}
QTableWidget::indicator:checked {
    background: qlineargradient(x1:0, y1:0, x2:1, y2:1, stop:0 #B8893A, stop:1 #5C4A32);
    border: 1px solid #5C4A32;
}
QScrollBar:vertical { background: #EFE6D4; width: 12px; margin: 0; }
QScrollBar::handle:vertical { background: #C4A574; border-radius: 5px; min-height: 28px; }
QScrollBar:horizontal { background: #EFE6D4; height: 12px; margin: 0; }
QScrollBar::handle:horizontal { background: #C4A574; border-radius: 5px; min-width: 28px; }
QScrollBar::add-line, QScrollBar::sub-line { width: 0; height: 0; }
QScrollBar::add-page, QScrollBar::sub-page { background: transparent; }
QListWidget#sectionList {
    background: rgba(255, 252, 246, 0.96);
    border: 1px solid #D4B896;
    border-radius: 6px;
    font-size: 13px;
}
QListWidget#sectionList::item { padding: 6px 8px; }
QListWidget#sectionList::item:selected { background: #E8D9BC; color: #2C281F; }
QPushButton#smallButton {
    background: #FFFcf7;
    color: #5C4A32;
    border: 1px solid #D4B896;
    border-radius: 6px;
    padding: 6px 12px;
    font-size: 12px;
    font-weight: 700;
    min-height: 28px;
}
QPushButton#smallButton:hover { background: #F0E6D6; }
QPushButton#smallButton:disabled { color: #B5A68E; }
QPushButton#addButton {
    background: #8A7349;
    color: #FFFFFF;
    border: none;
    border-radius: 6px;
    padding: 10px 16px;
    font-size: 14px;
    font-weight: 700;
}
QPushButton#addButton:hover { background: #6F5C39; }
QPushButton#addButton:disabled { background: #D9CDB8; color: #F4EDE8; }
QLabel#stepLabel {
    color: #5C4A32;
    font-size: 14px;
    font-weight: 800;
}
QLabel#countLabel { color: #7A6A52; font-size: 12px; font-weight: 700; }
"""

COL_CHECK, COL_HO, COL_NAME, COL_SPEC, COL_UNIT, COL_QTY, COL_MAT, COL_LAB, COL_NOTE = range(9)
HEADERS = ["선택", "호표", "명  칭", "규  격", "단위", "수량", "재료비 단가", "노무비 단가", "비  고"]


class FuncWorker(QThread):
    """엑셀 작업은 별도 줄에서 돌려 화면과 게이지가 멈추지 않게 한다."""

    succeeded = pyqtSignal(object)
    failed = pyqtSignal(str, str)

    def __init__(self, func: Callable[[], Any]) -> None:
        super().__init__()
        self._func = func

    def run(self) -> None:
        try:
            self.succeeded.emit(self._func())
        except ResultDirectoryError as exc:
            self.failed.emit("저장 폴더 오류", str(exc))
        except Exception as exc:  # noqa: BLE001 — 사용자 메시지로 보여 준다.
            self.failed.emit("처리 실패", str(exc))


def _money(value: float) -> str:
    return f"{value:,.0f}" if value else ""


class EstimateWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowFlags(
            Qt.WindowType.Window
            | Qt.WindowType.FramelessWindowHint
            | Qt.WindowType.WindowSystemMenuHint
            | Qt.WindowType.WindowMinimizeButtonHint
            | Qt.WindowType.WindowMaximizeButtonHint
        )
        self.setWindowTitle(APP_TITLE)
        self.setMinimumSize(1180, 820)
        self.resize(1320, 900)
        self.setStyleSheet(APP_STYLESHEET + EXTRA_STYLESHEET)
        icon_file = app_icon_path()
        if icon_file.is_file():
            self.setWindowIcon(QIcon(str(icon_file)))
        self._settings = _load_settings()
        self._source_path: Path | None = None
        self._session: EstimateSession | None = None
        self._sections: list[EstimateSection] = []
        self._updating = False
        self._busy_dialog: BusyDialog | None = None
        self._worker: FuncWorker | None = None
        self._legacy: LegacyWindow | None = None
        self._build_ui()
        self._refresh_state()

    # ------------------------------------------------------------------ UI
    def _build_ui(self) -> None:
        root = PaperRoot()
        self.setCentralWidget(root)
        outer = QVBoxLayout(root)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.setSpacing(0)

        self.caption = CaptionBar()
        self.caption.min_button.clicked.connect(self.showMinimized)
        self.caption.max_button.clicked.connect(self._toggle_maximized)
        self.caption.close_button.clicked.connect(self.close)
        outer.addWidget(self.caption)

        hero = QFrame()
        hero.setObjectName("hero")
        hero_layout = QVBoxLayout(hero)
        hero_layout.setContentsMargins(36, 14, 36, 12)
        hero_layout.setSpacing(6)
        title = QLabel(APP_TITLE)
        title.setObjectName("appTitle")
        hero_layout.addWidget(title)
        gold = QFrame()
        gold.setObjectName("doneGoldBar")
        gold.setFixedHeight(3)
        hero_layout.addWidget(gold)
        credits = QHBoxLayout()
        left = QVBoxLayout()
        left.setSpacing(2)
        maker = QLabel(APP_MAKER)
        maker.setObjectName("appMaker")
        contact = QLabel(APP_CONTACT)
        contact.setObjectName("appContact")
        left.addWidget(maker)
        left.addWidget(contact)
        notice = QLabel(APP_NOTICE)
        notice.setObjectName("appNotice")
        credits.addLayout(left, 1)
        credits.addWidget(notice, 0, Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        hero_layout.addLayout(credits)
        outer.addWidget(hero)

        body = QWidget()
        body_layout = QVBoxLayout(body)
        body_layout.setContentsMargins(24, 14, 24, 10)
        body_layout.setSpacing(10)

        # --- 1단계: 파트 · 단가대비표 · 저장 폴더 · 생성
        top = QFrame()
        top.setObjectName("card")
        top_layout = QHBoxLayout(top)
        top_layout.setContentsMargins(16, 12, 16, 12)
        top_layout.setSpacing(14)

        part_box = QVBoxLayout()
        step1 = QLabel("① 표준품셈 파트")
        step1.setObjectName("stepLabel")
        part_box.addWidget(step1)
        part_buttons = QHBoxLayout()
        self.electric_button = QPushButton("전기")
        self.electric_button.setObjectName("partElectric")
        self.telecom_button = QPushButton("통신")
        self.telecom_button.setObjectName("partTelecom")
        self.part_group = QButtonGroup(self)
        self.part_group.setExclusive(True)
        for button in (self.electric_button, self.telecom_button):
            button.setCheckable(True)
            button.setCursor(Qt.CursorShape.PointingHandCursor)
            self.part_group.addButton(button)
            part_buttons.addWidget(button)
            button.clicked.connect(self._on_part_changed)
        part_box.addLayout(part_buttons)
        self.part_file_hint = QLabel("")
        self.part_file_hint.setObjectName("partFileHint")
        part_box.addWidget(self.part_file_hint)
        top_layout.addLayout(part_box, 2)

        src_box = QVBoxLayout()
        step2 = QLabel("② 단가대비표")
        step2.setObjectName("stepLabel")
        src_box.addWidget(step2)
        self.compare_drop = DropZone(
            title="단가대비표",
            hint=DROP_HINT,
            dialog_title="단가대비표 엑셀 선택",
            tone="forward",
        )
        self.compare_drop.file_dropped.connect(self._on_source_dropped)
        src_box.addWidget(self.compare_drop)
        top_layout.addLayout(src_box, 4)

        dest_box = QVBoxLayout()
        dest_title = QLabel("③ 결과 저장 폴더")
        dest_title.setObjectName("stepLabel")
        dest_box.addWidget(dest_title)
        dest_row = QHBoxLayout()
        self.dest_edit = QLineEdit(self._saved_dest_dir())
        self.dest_edit.setObjectName("pathEdit")
        self.dest_edit.setFixedHeight(36)
        self.dest_edit.setPlaceholderText(str(WINDOWS_RESULT_DIR))
        dest_row.addWidget(self.dest_edit, 1)
        self.browse_button = QPushButton("폴더 찾기")
        self.browse_button.setObjectName("browseButton")
        self.browse_button.setMinimumSize(104, 36)
        self.browse_button.clicked.connect(self._on_browse_dest)
        dest_row.addWidget(self.browse_button)
        dest_box.addLayout(dest_row)
        gen_row = QHBoxLayout()
        self.generate_button = QPushButton("일위대가 생성")
        self.generate_button.setObjectName("runButton")
        self.generate_button.setMinimumHeight(40)
        self.generate_button.setCursor(Qt.CursorShape.PointingHandCursor)
        self.generate_button.clicked.connect(self._on_generate)
        gen_row.addWidget(self.generate_button, 1)
        self.reset_button = QPushButton("새로고침")
        self.reset_button.setObjectName("resetButton")
        self.reset_button.setMinimumHeight(40)
        self.reset_button.clicked.connect(self._on_reset)
        gen_row.addWidget(self.reset_button)
        dest_box.addLayout(gen_row)
        top_layout.addLayout(dest_box, 4)
        body_layout.addWidget(top)

        # --- 2단계: 일위대가목록(체크) + 공정 패널
        main_row = QHBoxLayout()
        main_row.setSpacing(12)

        list_card = QFrame()
        list_card.setObjectName("card")
        list_layout = QVBoxLayout(list_card)
        list_layout.setContentsMargins(14, 10, 14, 12)
        list_layout.setSpacing(8)
        list_head = QHBoxLayout()
        step4 = QLabel("④ 일위대가목록 — 품목 체크 · 수량 입력")
        step4.setObjectName("stepLabel")
        list_head.addWidget(step4)
        list_head.addStretch(1)
        self.count_label = QLabel("")
        self.count_label.setObjectName("countLabel")
        self.count_label.setMinimumWidth(170)
        self.count_label.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        list_head.addWidget(self.count_label)
        list_layout.addLayout(list_head)

        tools = QHBoxLayout()
        self.search_edit = QLineEdit()
        self.search_edit.setPlaceholderText("품명·규격 검색 (띄어쓰기로 여러 낱말)")
        self.search_edit.textChanged.connect(self._apply_filter)
        tools.addWidget(self.search_edit, 1)
        self.check_visible_button = QPushButton("보이는 항목 모두 체크")
        self.check_visible_button.setObjectName("smallButton")
        self.check_visible_button.clicked.connect(lambda: self._set_visible_checked(True))
        self.uncheck_all_button = QPushButton("체크 모두 해제")
        self.uncheck_all_button.setObjectName("smallButton")
        self.uncheck_all_button.clicked.connect(lambda: self._set_all_checked(False))
        tools.addWidget(self.check_visible_button)
        tools.addWidget(self.uncheck_all_button)
        list_layout.addLayout(tools)

        self.table = QTableWidget(0, len(HEADERS))
        self.table.setObjectName("itemTable")
        self.table.setHorizontalHeaderLabels(HEADERS)
        self.table.setAlternatingRowColors(True)
        self.table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.table.verticalHeader().setVisible(False)
        self.table.verticalHeader().setDefaultSectionSize(30)
        header = self.table.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        for col, width in zip(range(9), (46, 46, 190, 190, 46, 78, 92, 92, 140)):
            self.table.setColumnWidth(col, width)
        header.setStretchLastSection(True)
        self.table.itemChanged.connect(self._on_item_changed)
        list_layout.addWidget(self.table, 1)
        main_row.addWidget(list_card, 1)

        side = QFrame()
        side.setObjectName("card")
        side.setFixedWidth(340)
        side_layout = QVBoxLayout(side)
        side_layout.setContentsMargins(14, 10, 14, 12)
        side_layout.setSpacing(8)
        step5 = QLabel("⑤ 공정 제목")
        step5.setObjectName("stepLabel")
        side_layout.addWidget(step5)
        self.title_edit = QLineEdit()
        self.title_edit.setObjectName("sectionTitle")
        self.title_edit.setPlaceholderText("예) 1.아파트전기공사::1-1.단위세대설비공사")
        self.title_edit.textChanged.connect(self._refresh_state)
        side_layout.addWidget(self.title_edit)
        self.add_button = QPushButton("체크한 품목을 이 공정에 담기")
        self.add_button.setObjectName("addButton")
        self.add_button.setCursor(Qt.CursorShape.PointingHandCursor)
        self.add_button.clicked.connect(self._on_add_section)
        side_layout.addWidget(self.add_button)

        step6 = QLabel("⑥ 담은 공정 (내역서 시트가 공정마다 하나씩 만들어집니다)")
        step6.setObjectName("stepLabel")
        step6.setWordWrap(True)
        side_layout.addWidget(step6)
        self.section_list = QListWidget()
        self.section_list.setObjectName("sectionList")
        self.section_list.setWordWrap(True)
        self.section_list.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.section_list.itemSelectionChanged.connect(self._refresh_state)
        side_layout.addWidget(self.section_list, 1)
        section_buttons = QHBoxLayout()
        self.edit_section_button = QPushButton("불러와 고치기")
        self.edit_section_button.setObjectName("smallButton")
        self.edit_section_button.clicked.connect(self._on_edit_section)
        self.remove_section_button = QPushButton("공정 삭제")
        self.remove_section_button.setObjectName("smallButton")
        self.remove_section_button.clicked.connect(self._on_remove_section)
        section_buttons.addWidget(self.edit_section_button)
        section_buttons.addWidget(self.remove_section_button)
        side_layout.addLayout(section_buttons)
        self.save_button = QPushButton("내역서 저장")
        self.save_button.setObjectName("runButton")
        self.save_button.setMinimumHeight(46)
        self.save_button.setCursor(Qt.CursorShape.PointingHandCursor)
        self.save_button.clicked.connect(self._on_save)
        side_layout.addWidget(self.save_button)
        self.legacy_button = QPushButton("기존 도구 (역방향 · 공량산출)")
        self.legacy_button.setObjectName("smallButton")
        self.legacy_button.clicked.connect(self._open_legacy)
        side_layout.addWidget(self.legacy_button)
        main_row.addWidget(side)
        body_layout.addLayout(main_row, 1)

        self.log = QTextEdit()
        self.log.setObjectName("log")
        self.log.setReadOnly(True)
        self.log.setFixedHeight(84)
        body_layout.addWidget(self.log)
        outer.addWidget(body, 1)

        status = QStatusBar()
        status.setSizeGripEnabled(True)
        status.showMessage("대기 — 단가대비표를 놓고 [일위대가 생성]을 누르세요.")
        self.setStatusBar(status)
        self._restore_discipline()
        self._append_log("원본 엑셀은 읽기만 합니다. 결과는 새 파일로만 저장합니다.")

    # ------------------------------------------------------------- helpers
    def _saved_dest_dir(self) -> str:
        stored = self._settings.value("dest_dir", "")
        if isinstance(stored, str) and stored.strip():
            return stored.strip()
        return display_result_directory()

    def _chosen_dest_dir(self) -> Path | None:
        text = self.dest_edit.text().strip()
        return Path(text) if text else None

    def _selected_discipline(self) -> str:
        return TELECOM if self.telecom_button.isChecked() else ELECTRIC

    def _restore_discipline(self) -> None:
        stored = self._settings.value("discipline", ELECTRIC)
        disc = normalize_discipline(str(stored) if stored is not None else ELECTRIC)
        self.electric_button.setChecked(disc == ELECTRIC)
        self.telecom_button.setChecked(disc == TELECOM)
        self._refresh_part_hint()

    def _refresh_part_hint(self) -> None:
        self.part_file_hint.setText(f"적용 품셈 : {pumsam_filename(self._selected_discipline())}")

    def _on_part_changed(self) -> None:
        self._settings.setValue("discipline", self._selected_discipline())
        self._refresh_part_hint()
        if self._session is not None:
            self._append_log("파트를 바꿨습니다. 새 파트로 만들려면 [일위대가 생성]을 다시 누르세요.")

    def _append_log(self, message: str) -> None:
        self.log.append(message)

    def _toggle_maximized(self) -> None:
        if self.isMaximized():
            self.showNormal()
            self.caption.max_button.setText("□")
        else:
            self.showMaximized()
            self.caption.max_button.setText("❐")

    def _on_browse_dest(self) -> None:
        start = self.dest_edit.text().strip() or display_result_directory()
        chosen = QFileDialog.getExistingDirectory(self, "결과 저장 폴더 선택", start)
        if chosen:
            self.dest_edit.setText(chosen)
            self._settings.setValue("dest_dir", chosen)

    def _open_legacy(self) -> None:
        if self._legacy is None:
            self._legacy = LegacyWindow()
        self._legacy.show()
        self._legacy.raise_()

    # --------------------------------------------------------------- state
    def _checked_rows(self) -> list[int]:
        rows = []
        for row in range(self.table.rowCount()):
            item = self.table.item(row, COL_CHECK)
            if item is not None and item.checkState() == Qt.CheckState.Checked:
                rows.append(row)
        return rows

    def _refresh_state(self) -> None:
        has_session = self._session is not None
        checked = len(self._checked_rows())
        busy = self._worker is not None and self._worker.isRunning()
        self.generate_button.setEnabled(self._source_path is not None and not busy)
        self.search_edit.setEnabled(has_session)
        self.check_visible_button.setEnabled(has_session)
        self.uncheck_all_button.setEnabled(has_session)
        self.add_button.setEnabled(has_session and checked > 0 and bool(self.title_edit.text().strip()) and not busy)
        selected = bool(self.section_list.selectedItems())
        self.edit_section_button.setEnabled(selected)
        self.remove_section_button.setEnabled(selected)
        self.save_button.setEnabled(has_session and bool(self._sections) and not busy)
        total = self.table.rowCount()
        self.count_label.setText(f"전체 {total}건 · 체크 {checked}건" if has_session else "")

    # ------------------------------------------------------------- source
    def _on_source_dropped(self, path_text: str) -> None:
        self._source_path = Path(path_text)
        self.compare_drop.set_loaded(self._source_path.name)
        self._append_log(f"단가대비표 선택: {self._source_path.name}")
        self.statusBar().showMessage("단가대비표가 선택되었습니다. [일위대가 생성]을 누르세요.")
        self._refresh_state()

    def _on_reset(self) -> None:
        self._source_path = None
        self._session = None
        self._sections = []
        self.compare_drop.reset()
        self.table.setRowCount(0)
        self.section_list.clear()
        self.title_edit.clear()
        self.search_edit.clear()
        self.log.clear()
        self.statusBar().showMessage("대기 — 단가대비표를 놓고 [일위대가 생성]을 누르세요.")
        self._refresh_state()

    # ---------------------------------------------------------- generation
    def _open_busy(self) -> BusyDialog:
        dialog = BusyDialog(self)
        dialog.setWindowModality(Qt.WindowModality.ApplicationModal)
        dialog.show()
        dialog.raise_()
        dialog.start_gauge()
        QApplication.processEvents()
        return dialog

    def _run_background(self, func: Callable[[], Any], on_done: Callable[[Any], None]) -> None:
        self._busy_dialog = self._open_busy()
        self._worker = FuncWorker(func)

        def done(result: Any) -> None:
            self._close_busy()
            on_done(result)

        def failed(title: str, text: str) -> None:
            self._close_busy()
            self._append_log(f"{title}: {text}")
            self.statusBar().showMessage(title)
            QMessageBox.critical(self, title, text)

        self._worker.succeeded.connect(done)
        self._worker.failed.connect(failed)
        self._worker.start()
        self._refresh_state()

    def _close_busy(self) -> None:
        if self._worker is not None:
            self._worker.wait(3000)
        if self._busy_dialog is not None:
            self._busy_dialog.complete_and_close()
            self._busy_dialog = None
        QApplication.processEvents()

    def _on_generate(self) -> None:
        if self._source_path is None:
            QMessageBox.warning(self, "파일 없음", "단가대비표를 먼저 놓아 주세요.")
            return
        dest = self._chosen_dest_dir()
        if dest is not None:
            self._settings.setValue("dest_dir", str(dest))
        disc = self._selected_discipline()
        source = self._source_path
        self._append_log(f"일위대가 생성 중… ({disc} · {pumsam_filename(disc)})")
        self.statusBar().showMessage("일위대가를 생성하는 중…")
        self._run_background(lambda: generate_session(source, disc, dest), self._on_generated)

    def _on_generated(self, session: EstimateSession) -> None:
        self._load_session(session)
        self._append_log(f"일위대가 {len(session.entries)}건 생성. 내역서에 넣을 품목을 체크하세요.")
        self.statusBar().showMessage("일위대가목록이 만들어졌습니다.")

    def _load_session(self, session: EstimateSession) -> None:
        self._session = session
        self._sections = []
        self.section_list.clear()
        self._updating = True
        try:
            self.table.setRowCount(0)
            self.table.setRowCount(len(session.entries))
            for row, entry in enumerate(session.entries):
                check = QTableWidgetItem()
                check.setFlags(Qt.ItemFlag.ItemIsUserCheckable | Qt.ItemFlag.ItemIsEnabled)
                check.setCheckState(Qt.CheckState.Unchecked)
                check.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                self.table.setItem(row, COL_CHECK, check)
                self._put(row, COL_HO, str(entry.ho_no), center=True)
                self._put(row, COL_NAME, entry.name)
                self._put(row, COL_SPEC, entry.spec)
                self._put(row, COL_UNIT, entry.unit, center=True)
                self._put(row, COL_QTY, "", editable=True, right=True)
                self._put(row, COL_MAT, _money(entry.material_unit), right=True)
                self._put(row, COL_LAB, _money(entry.labor_unit), right=True)
                self._put(row, COL_NOTE, "", editable=True)
        finally:
            self._updating = False
        self._apply_filter()
        self._refresh_state()

    def _put(self, row: int, col: int, text: str, *, editable: bool = False, center: bool = False, right: bool = False) -> None:
        item = QTableWidgetItem(text)
        flags = Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsSelectable
        if editable:
            flags |= Qt.ItemFlag.ItemIsEditable
        item.setFlags(flags)
        if center:
            item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
        elif right:
            item.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        self.table.setItem(row, col, item)

    # ---------------------------------------------------------- selection
    def _on_item_changed(self, item: QTableWidgetItem) -> None:
        if self._updating:
            return
        if item.column() == COL_CHECK:
            qty_item = self.table.item(item.row(), COL_QTY)
            if qty_item is not None:
                self._updating = True
                try:
                    if item.checkState() == Qt.CheckState.Checked and not qty_item.text().strip():
                        qty_item.setText("1")
                finally:
                    self._updating = False
        self._refresh_state()

    def _apply_filter(self) -> None:
        tokens = [t for t in self.search_edit.text().lower().split() if t]
        for row in range(self.table.rowCount()):
            haystack = " ".join(
                (self.table.item(row, c).text() if self.table.item(row, c) else "") for c in (COL_NAME, COL_SPEC)
            ).lower()
            self.table.setRowHidden(row, not all(t in haystack for t in tokens))

    def _set_visible_checked(self, state: bool) -> None:
        self._updating = True
        try:
            for row in range(self.table.rowCount()):
                if self.table.isRowHidden(row):
                    continue
                self._set_check(row, state)
        finally:
            self._updating = False
        self._refresh_state()

    def _set_all_checked(self, state: bool) -> None:
        self._updating = True
        try:
            for row in range(self.table.rowCount()):
                self._set_check(row, state)
        finally:
            self._updating = False
        self._refresh_state()

    def _set_check(self, row: int, state: bool) -> None:
        item = self.table.item(row, COL_CHECK)
        qty_item = self.table.item(row, COL_QTY)
        if item is None:
            return
        item.setCheckState(Qt.CheckState.Checked if state else Qt.CheckState.Unchecked)
        if qty_item is not None:
            if state and not qty_item.text().strip():
                qty_item.setText("1")
            if not state:
                qty_item.setText("")
        note = self.table.item(row, COL_NOTE)
        if note is not None and not state:
            note.setText("")

    # ------------------------------------------------------------ sections
    @staticmethod
    def _parse_qty(text: str) -> float | None:
        try:
            value = float(text.replace(",", "").strip())
        except ValueError:
            return None
        return value if value > 0 else None

    def _collect_lines(self) -> list[EstimateLine] | None:
        assert self._session is not None
        lines: list[EstimateLine] = []
        bad: list[str] = []
        for row in self._checked_rows():
            entry = self._session.entries[row]
            qty = self._parse_qty(self.table.item(row, COL_QTY).text())
            if qty is None:
                bad.append(entry.label)
                continue
            lines.append(EstimateLine(entry, qty, self.table.item(row, COL_NOTE).text().strip()))
        if bad:
            QMessageBox.warning(
                self,
                "수량 확인",
                "수량이 비었거나 0 이하인 품목이 있습니다.\n\n" + "\n".join(bad[:12]) + ("\n…" if len(bad) > 12 else ""),
            )
            return None
        return lines

    def _on_add_section(self) -> None:
        if self._session is None:
            return
        title = self.title_edit.text().strip()
        if not title:
            QMessageBox.warning(self, "공정 제목", "공정 제목을 입력해 주세요.")
            return
        lines = self._collect_lines()
        if not lines:
            return
        self._sections.append(EstimateSection(title, lines))
        self.section_list.addItem(f"{len(self._sections)}. {title}  ({len(lines)}품목)")
        self._append_log(f"공정 담기: {title} — {len(lines)}품목")
        self._set_all_checked(False)
        self.title_edit.clear()
        self.statusBar().showMessage("공정을 담았습니다. 다른 공정을 이어서 담거나 [내역서 저장]을 누르세요.")
        self._refresh_state()

    def _selected_section_index(self) -> int | None:
        rows = self.section_list.selectedIndexes()
        return rows[0].row() if rows else None

    def _renumber_sections(self) -> None:
        self.section_list.clear()
        for number, section in enumerate(self._sections, start=1):
            self.section_list.addItem(f"{number}. {section.title}  ({len(section.lines)}품목)")

    def _on_remove_section(self) -> None:
        index = self._selected_section_index()
        if index is None:
            return
        removed = self._sections.pop(index)
        self._renumber_sections()
        self._append_log(f"공정 삭제: {removed.title}")
        self._refresh_state()

    def _on_edit_section(self) -> None:
        index = self._selected_section_index()
        if index is None or self._session is None:
            return
        section = self._sections.pop(index)
        self._renumber_sections()
        self._set_all_checked(False)
        self._updating = True
        try:
            for line in section.lines:
                row = line.entry.index
                self.table.item(row, COL_CHECK).setCheckState(Qt.CheckState.Checked)
                self.table.item(row, COL_QTY).setText(f"{line.qty:g}")
                self.table.item(row, COL_NOTE).setText(line.remark)
        finally:
            self._updating = False
        self.title_edit.setText(section.title)
        self._append_log(f"공정을 불러왔습니다: {section.title}. 고친 뒤 다시 [담기]를 누르세요.")
        self._refresh_state()

    # --------------------------------------------------------------- save
    def _on_save(self) -> None:
        if self._session is None:
            return
        if self._checked_rows() and self.title_edit.text().strip():
            answer = QMessageBox.question(
                self,
                "담지 않은 품목",
                "체크해 둔 품목이 아직 공정에 담기지 않았습니다.\n이 품목도 담고 저장할까요?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No | QMessageBox.StandardButton.Cancel,
            )
            if answer == QMessageBox.StandardButton.Cancel:
                return
            if answer == QMessageBox.StandardButton.Yes:
                self._on_add_section()
        if not self._sections:
            QMessageBox.warning(self, "공정 없음", "내역서에 담을 공정이 없습니다.")
            return
        dest = self._chosen_dest_dir()
        if dest is not None:
            self._settings.setValue("dest_dir", str(dest))
        session, sections = self._session, list(self._sections)
        self.statusBar().showMessage("내역서를 저장하는 중…")
        self._run_background(lambda: save_session(session, sections, dest), self._on_saved)

    def _on_saved(self, dest: Path) -> None:
        self._append_log(f"새 파일 저장: {dest}")
        self.statusBar().showMessage(f"저장 완료 — {Path(dest).name}")
        DoneDialog(self, Path(dest), pumsam_filename(self._selected_discipline())).exec()
        self._refresh_state()
