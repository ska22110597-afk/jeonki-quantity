"""내역서 작성 화면: 단가대비표 → 일위대가목록 → 체크박스로 고른 품목 → 공정별 내역서."""

from __future__ import annotations

from pathlib import Path
from typing import Any, Callable

from PyQt6.QtCore import Qt, QThread, pyqtSignal
from PyQt6.QtGui import QBrush, QColor, QGuiApplication, QIcon
from PyQt6.QtWidgets import (
    QAbstractItemView,
    QApplication,
    QButtonGroup,
    QCheckBox,
    QComboBox,
    QDialog,
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
    QSizeGrip,
    QSizePolicy,
    QStatusBar,
    QTableWidget,
    QTableWidgetItem,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from app.sound import play_done
from app.numbers import parse_number
from app.link_dialog import ColumnMapDialog, LinkDialog
from app.pumsam import load_pumsam_database
from app.pumsam_links import save_link
from app.discipline import ELECTRIC, TELECOM, normalize_discipline, pumsam_filename
from app.drop_zone import DropZone
from app.estimate_build import (
    EstimateLine,
    EstimateSection,
    strip_section_number,
    EstimateSession,
    generate_session,
    set_material_price,
    open_session,
    save_session,
    MODE_ILWIDAE,
    MODE_QUANTITY,
    autosave_path,
    convert_with_columns,
    duplicate_rows,
    layout_ok,
    read_preview,
    read_autosave,
    write_autosave,
    is_result_workbook,
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
from app.paths import workspace_dir, ResultDirectoryError, WINDOWS_RESULT_DIR, app_icon_path, display_result_directory
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
    outline: 0;
}
QTableWidget#itemTable::item:focus { border: none; outline: none; }
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

COL_CHECK, COL_HO, COL_NAME, COL_SPEC, COL_UNIT, COL_QTY, COL_MAT, COL_LAB, COL_NOTE, COL_STATE = range(10)
HEADERS = ["선택", "호표", "명  칭", "규  격", "단위", "수량", "재료비 단가", "노무비 단가", "비  고", "품셈"]


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


_RESIZE_MARGIN = 5


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
        self.setMinimumSize(900, 620)
        screen = QGuiApplication.primaryScreen()
        avail = screen.availableGeometry() if screen is not None else None
        width = min(1560, avail.width() - 40) if avail else 1400
        height = min(920, avail.height() - 40) if avail else 900
        self.resize(width, height)
        self.setMouseTracking(True)
        self.setContentsMargins(_RESIZE_MARGIN, _RESIZE_MARGIN, _RESIZE_MARGIN, _RESIZE_MARGIN)
        self.setStyleSheet(APP_STYLESHEET + EXTRA_STYLESHEET)
        icon_file = app_icon_path()
        if icon_file.is_file():
            self.setWindowIcon(QIcon(str(icon_file)))
        self._settings = _load_settings()
        self._source_path: Path | None = None
        self._session: EstimateSession | None = None
        self._sections: list[EstimateSection] = []
        self._dup_rows: set[int] = set()
        self._hide_dups = False
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
        step2 = QLabel("② 단가대비표 / 일위대가목록 파일")
        step2.setObjectName("stepLabel")
        src_box.addWidget(step2)
        self.compare_drop = DropZone(
            title="단가대비표 또는 일위대가목록",
            hint="단가대비표를 넣으면 일위대가부터 새로 만들고, 저장해 둔 일위대가목록 파일을 넣으면 이어서 내역서를 만듭니다.",
            dialog_title="단가대비표 또는 일위대가목록 엑셀 선택",
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
        self.unmatched_only = QCheckBox("미매칭만 보기")
        self.unmatched_only.stateChanged.connect(self._apply_filter)
        self.no_price_only = QCheckBox("재료비 단가 공란만 보기")
        self.no_price_only.stateChanged.connect(self._apply_filter)
        self.link_button = QPushButton("선택 품목 품셈 연결")
        self.link_button.setObjectName("smallButton")
        self.link_button.clicked.connect(self._on_link)
        tools.addWidget(self.unmatched_only)
        tools.addWidget(self.no_price_only)
        tools.addWidget(self.link_button)
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
        for col, width in zip(range(10), (44, 44, 170, 170, 44, 66, 86, 86, 110, 110)):
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
        self.title_edit.setPlaceholderText("공정 제목만 입력 (번호는 자동) 예) 내선설비공사")
        self.title_edit.textChanged.connect(self._refresh_state)
        side_layout.addWidget(self.title_edit)
        self.add_button = QPushButton("체크한 품목을 이 공정에 담기")
        self.add_button.setObjectName("addButton")
        self.add_button.setCursor(Qt.CursorShape.PointingHandCursor)
        self.add_button.clicked.connect(self._on_add_section)
        side_layout.addWidget(self.add_button)

        step6 = QLabel("⑥ 담은 공정 (내역서 한 시트에 순서대로 이어서 작성됩니다)")
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
        mode_label = QLabel("내역서 방식")
        mode_label.setObjectName("stepLabel")
        side_layout.addWidget(mode_label)
        self.mode_combo = QComboBox()
        self.mode_combo.addItem("일위대가 방식 (품목별 재료비+노무비 단가)", MODE_ILWIDAE)
        self.mode_combo.addItem("공량산출 방식 (재료비 + 직종별 공량·노무비, 공량산출서 시트 추가)", MODE_QUANTITY)
        side_layout.addWidget(self.mode_combo)
        self.save_button = QPushButton("저장")
        self.save_button.setObjectName("runButton")
        self.save_button.setMinimumHeight(46)
        self.save_button.setCursor(Qt.CursorShape.PointingHandCursor)
        self.save_button.clicked.connect(self._on_save)
        side_layout.addWidget(self.save_button)
        look_row = QHBoxLayout()
        self.bg_button = QPushButton("배경 그림 변경")
        self.bg_button.setObjectName("smallButton")
        self.bg_button.clicked.connect(self._on_change_background)
        self.bg_reset_button = QPushButton("기본 배경")
        self.bg_reset_button.setObjectName("smallButton")
        self.bg_reset_button.clicked.connect(self._on_reset_background)
        look_row.addWidget(self.bg_button, 1)
        look_row.addWidget(self.bg_reset_button)
        side_layout.addLayout(look_row)
        self.sound_check = QCheckBox("완료 알림음")
        self.sound_check.setChecked(str(self._settings.value("sound", "1")) != "0")
        self.sound_check.stateChanged.connect(lambda _s: self._settings.setValue("sound", "1" if self.sound_check.isChecked() else "0"))
        side_layout.addWidget(self.sound_check)
        self._apply_saved_background()
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
        return workspace_dir(Path(text)) if text else None

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

    def _edges_at(self, pos) -> Qt.Edge:
        edges = Qt.Edge(0)
        m = _RESIZE_MARGIN + 1
        if pos.x() <= m:
            edges |= Qt.Edge.LeftEdge
        if pos.x() >= self.width() - m:
            edges |= Qt.Edge.RightEdge
        if pos.y() <= m:
            edges |= Qt.Edge.TopEdge
        if pos.y() >= self.height() - m:
            edges |= Qt.Edge.BottomEdge
        return edges

    def mouseMoveEvent(self, event) -> None:  # noqa: N802
        if not self.isMaximized():
            edges = self._edges_at(event.position().toPoint())
            horizontal = bool(edges & (Qt.Edge.LeftEdge | Qt.Edge.RightEdge))
            vertical = bool(edges & (Qt.Edge.TopEdge | Qt.Edge.BottomEdge))
            if horizontal and vertical:
                diagonal = (edges & Qt.Edge.LeftEdge and edges & Qt.Edge.TopEdge) or (edges & Qt.Edge.RightEdge and edges & Qt.Edge.BottomEdge)
                self.setCursor(Qt.CursorShape.SizeFDiagCursor if diagonal else Qt.CursorShape.SizeBDiagCursor)
            elif horizontal:
                self.setCursor(Qt.CursorShape.SizeHorCursor)
            elif vertical:
                self.setCursor(Qt.CursorShape.SizeVerCursor)
            else:
                self.unsetCursor()
        super().mouseMoveEvent(event)

    def mousePressEvent(self, event) -> None:  # noqa: N802
        if event.button() == Qt.MouseButton.LeftButton and not self.isMaximized():
            edges = self._edges_at(event.position().toPoint())
            handle = self.windowHandle()
            if edges and handle is not None:
                handle.startSystemResize(edges)
                return
        super().mousePressEvent(event)

    # ----------------------------------------------------- 배경 그림·알림음
    def _background_store(self) -> Path:
        from app.activation import marker_path

        return marker_path().parent

    def _apply_saved_background(self) -> None:
        stored = str(self._settings.value("background", "") or "")
        if stored and Path(stored).is_file():
            self.centralWidget().set_background(stored)

    def _on_change_background(self) -> None:
        start = str(Path(str(self._settings.value("background", "") or "")).parent) if self._settings.value("background", "") else str(Path.home())
        chosen, _ = QFileDialog.getOpenFileName(self, "배경으로 쓸 그림 선택", start, "그림 파일 (*.png *.jpg *.jpeg *.bmp *.gif *.webp)")
        if not chosen:
            return
        try:  # 원본을 옮기거나 지워도 유지되게 사본을 프로그램 설정 폴더에 둔다.
            import shutil

            store = self._background_store()
            store.mkdir(parents=True, exist_ok=True)
            copy = store / f"background{Path(chosen).suffix.lower()}"
            shutil.copyfile(chosen, copy)
        except OSError:
            copy = Path(chosen)
        if not self.centralWidget().set_background(str(copy)):
            QMessageBox.warning(self, "배경 그림", "그림 파일을 열 수 없습니다. 다른 파일을 골라 주세요.")
            return
        self._settings.setValue("background", str(copy))
        self._append_log(f"배경 그림을 바꿨습니다: {Path(chosen).name}")

    def _on_reset_background(self) -> None:
        self.centralWidget().set_background(None)
        self._settings.setValue("background", "")
        self._append_log("기본 배경으로 되돌렸습니다.")

    def _notify_done(self) -> None:
        if self.sound_check.isChecked():
            play_done()

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

    def _on_link(self) -> None:
        """선택한 미매칭 품목에 품셈 항목을 직접 연결하고, 다시 일위대가를 만든다."""
        if self._session is None:
            return
        rows = sorted({i.row() for i in self.table.selectedItems()})
        if not rows:
            QMessageBox.information(self, "품목 선택", "연결할 품목 줄을 먼저 눌러 선택하세요.")
            return
        if self._source_path is None or is_result_workbook(self._source_path):
            QMessageBox.information(
                self,
                "단가대비표에서 적용",
                "품셈 연결은 단가대비표에서 일위대가를 새로 만들 때 적용됩니다.\n단가대비표를 넣고 다시 생성해 주세요. (연결은 저장해 둘 수 있습니다.)",
            )
        entry = self._session.entries[rows[0]]
        book = self._session.pumsam_rows or load_pumsam_database(None, self._session.discipline)
        dialog = LinkDialog(self, entry.label, book)
        if dialog.exec() != QDialog.DialogCode.Accepted or dialog.chosen() is None:
            return
        book_name, book_spec = dialog.chosen()
        for row in rows:
            e = self._session.entries[row]
            save_link(self._chosen_dest_dir(), e.name, e.spec, book_name, book_spec)
        self._append_log(f"품셈 연결 저장: {len(rows)}건 → {book_name} {book_spec}")
        if self._source_path is not None and not is_result_workbook(self._source_path):
            self._on_generate()

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
        self.save_button.setEnabled(has_session and not busy)
        total = self.table.rowCount()
        unmatched = sum(1 for e in self._session.entries if not e.matched) if has_session else 0
        self.link_button.setEnabled(has_session and not busy)
        self.unmatched_only.setEnabled(has_session)
        self.count_label.setText(f"전체 {total}건 · 미매칭 {unmatched}건 · 체크 {checked}건" if has_session else "")

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
            self._settings.setValue("dest_dir", self.dest_edit.text().strip())
        disc = self._selected_discipline()
        source = self._source_path
        if not is_result_workbook(source) and not layout_ok(source):
            source = self._ask_columns(source, dest)
            if source is None:
                return
        duplicate_sheet = True
        if not is_result_workbook(source):
            from app.pipeline import scan_duplicates

            found = scan_duplicates(source)
            if found:
                names = ", ".join(f"{n} {s}".strip() + f"({c}회)" for n, s, c in found[:5])
                box = QMessageBox(self)
                box.setWindowTitle("중복 품목")
                box.setIcon(QMessageBox.Icon.Question)
                box.setText(
                    f"같은 품명·규격이 중복된 품목이 {len(found)}종 있습니다.\n({names} …)\n\n"
                    "단가대비표·일위대가·일위대가목록에는 한 줄만 남깁니다.\n"
                    "중복 품목 전체를 '중복항목' 시트(빨간 탭)에 모아 둘까요?"
                )
                make_btn = box.addButton("중복항목 시트 만들기", QMessageBox.ButtonRole.YesRole)
                skip_btn = box.addButton("시트 없이 한 줄로만", QMessageBox.ButtonRole.NoRole)
                box.addButton("취소", QMessageBox.ButtonRole.RejectRole)
                box.setDefaultButton(make_btn)
                box.exec()
                clicked = box.clickedButton()
                if clicked is skip_btn:
                    duplicate_sheet = False
                elif clicked is not make_btn:
                    self._append_log("중복 품목 확인에서 취소했습니다.")
                    return
                self._append_log(
                    f"중복 품목 {len(found)}종: "
                    + ("중복항목 시트로 모으고 한 줄로 정리" if duplicate_sheet else "한 줄로만 정리")
                )
        self._append_log(f"일위대가 생성 중… ({disc} · {pumsam_filename(disc)})")
        self.statusBar().showMessage("일위대가를 생성하는 중…")
        self._run_background(lambda: open_session(source, disc, dest, duplicate_sheet), self._on_generated)

    def _ask_columns(self, source: Path, dest: Path | None) -> Path | None:
        """양식을 못 읽으면 열을 직접 고르게 하고, 이 프로그램 양식으로 바꾼 파일을 돌려준다."""
        try:
            preview = read_preview(source)
        except Exception as exc:  # noqa: BLE001
            QMessageBox.critical(self, "파일 읽기 실패", f"엑셀 파일을 읽지 못했습니다.\n{exc}")
            return None
        dialog = ColumnMapDialog(self, preview)
        if dialog.exec() != QDialog.DialogCode.Accepted:
            return None
        mapping = dialog.mapping()
        try:
            from app.paths import user_database_dir

            converted = convert_with_columns(source, user_database_dir(dest), **mapping)
        except Exception as exc:  # noqa: BLE001
            QMessageBox.critical(self, "변환 실패", str(exc))
            return None
        if not layout_ok(converted):
            QMessageBox.warning(self, "열 지정 확인", "지정한 열에서 품목을 읽지 못했습니다. 행 번호와 열을 다시 확인해 주세요.")
            return None
        self._append_log(f"열 지정으로 변환했습니다: {converted.name}")
        return converted

    def _on_generated(self, session: EstimateSession) -> None:
        if session.warnings:
            shown = "\n".join(session.warnings[:8]) + ("\n…" if len(session.warnings) > 8 else "")
            answer = QMessageBox.question(
                self,
                "불러온 파일 확인",
                f"읽는 중 문제가 {len(session.warnings)}건 있었습니다.\n\n{shown}\n\n인식된 호표만으로 계속할까요?",
            )
            if answer != QMessageBox.StandardButton.Yes:
                self._append_log("불러오기를 취소했습니다. 파일을 고친 뒤 다시 넣어 주세요.")
                self.statusBar().showMessage("불러오기 취소")
                self._refresh_state()
                return
        self._load_session(session)
        self._notify_done()
        unmatched = sum(1 for e in session.entries if not e.matched)
        if unmatched:
            self._append_log(f"품셈 미매칭 {unmatched}건 — '미매칭만 보기'로 확인하고 [선택 품목 품셈 연결]을 쓰세요.")
        if self._dup_rows:
            names = ", ".join(session.entries[i].label for i in sorted(self._dup_rows)[:5])
            answer = QMessageBox.question(
                self,
                "중복 품목",
                f"같은 품명·규격이 {len(self._dup_rows)}건 중복되어 있습니다.\n({names} …)\n\n중복 줄을 선택 목록에서 숨길까요?",
            )
            if answer == QMessageBox.StandardButton.Yes:
                self._hide_dups = True
                self._apply_filter()
        for warning in session.warnings:
            self._append_log(f"주의: {warning}")
        self._offer_restore(session)
        for notice in session.notices[:10]:
            self._append_log(f"정리: {notice}")
        if len(session.notices) > 10:
            self._append_log(f"정리: 문자로 적힌 단가 {len(session.notices)}건을 숫자로 읽었습니다.")
        from app.pumsam import DB_NOTICES

        while DB_NOTICES:
            self._append_log(DB_NOTICES.pop(0))
        self._append_log(f"일위대가 {len(session.entries)}건을 불러왔습니다. 내역서에 넣을 품목을 체크하세요.")
        self.statusBar().showMessage("일위대가목록이 만들어졌습니다.")

    def _load_session(self, session: EstimateSession) -> None:
        self._session = session
        self._sections = []
        duplicates = duplicate_rows(session.entries)
        self._dup_rows = set(duplicates)
        self._hide_dups = False
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
                self._put(row, COL_MAT, _money(entry.material_unit), editable=True, right=True)
                self._put(row, COL_LAB, _money(entry.labor_unit), right=True)
                self._put(row, COL_NOTE, "", editable=True)
                state_text = entry.pumsam_ref if entry.matched else "미매칭"
                if row in duplicates:
                    state_text += " · 중복"
                self._put(row, COL_STATE, state_text)
                if not entry.matched:
                    self.table.item(row, COL_STATE).setForeground(QBrush(QColor("#B3261E")))
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
        if item.column() == COL_QTY and item.text().strip():
            number, cleaned = parse_number(item.text())
            if cleaned and number is not None:
                self._updating = True
                try:
                    self._append_log(f"수량 「{item.text().strip()}」 → {number:g} 로 정리했습니다.")
                    item.setText(f"{number:g}")
                finally:
                    self._updating = False
        if item.column() == COL_MAT and self._session is not None:
            self._on_price_edited(item)
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

    def _on_price_edited(self, item: QTableWidgetItem) -> None:
        entry = self._session.entries[item.row()]
        number, _cleaned = parse_number(item.text())
        self._updating = True
        try:
            if number is None or number < 0:
                item.setText(_money(entry.material_unit))
                if item.text().strip():
                    self._append_log("재료비 단가는 숫자로 입력해 주세요.")
                return
            set_material_price(self._session, entry, number)
            item.setText(_money(entry.material_unit))
            self._append_log(f"「{entry.label}」 재료비 단가를 {number:,.0f} 으로 넣었습니다. (일위대가 시트에 반영)")
        finally:
            self._updating = False

    def _apply_filter(self) -> None:
        tokens = [t for t in self.search_edit.text().lower().split() if t]
        for row in range(self.table.rowCount()):
            haystack = " ".join(
                (self.table.item(row, c).text() if self.table.item(row, c) else "") for c in (COL_NAME, COL_SPEC)
            ).lower()
            hide = not all(t in haystack for t in tokens) or (self._hide_dups and row in self._dup_rows)
            if not hide and self.unmatched_only.isChecked() and self._session is not None:
                hide = self._session.entries[row].matched
            if not hide and self.no_price_only.isChecked() and self._session is not None:
                hide = self._session.entries[row].material_unit > 0
            self.table.setRowHidden(row, hide)

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
        value, _cleaned = parse_number(text)
        return value if value is not None and value > 0 else None

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
        title = strip_section_number(self.title_edit.text())
        if not title:
            QMessageBox.warning(self, "공정 제목", "공정 제목을 입력해 주세요.")
            return
        lines = self._collect_lines()
        if not lines:
            return
        self._sections.append(EstimateSection(title, lines))
        self.section_list.addItem(f"{len(self._sections)}. {title}  ({len(lines)}품목)")
        self._autosave()
        self._append_log(f"공정 담기: {title} — {len(lines)}품목")
        self._set_all_checked(False)
        self.title_edit.clear()
        self.statusBar().showMessage("공정을 담았습니다. 다른 공정을 이어서 담거나 [내역서 저장]을 누르세요.")
        self._refresh_state()

    def _selected_section_index(self) -> int | None:
        rows = self.section_list.selectedIndexes()
        return rows[0].row() if rows else None

    def _autosave(self) -> None:
        if self._session is None:
            return
        try:
            write_autosave(autosave_path(self._chosen_dest_dir()), self._session, self._sections)
        except OSError:
            pass

    def _offer_restore(self, session: EstimateSession) -> None:
        try:
            path = autosave_path(self._chosen_dest_dir())
        except OSError:
            return
        if not path.exists():
            return
        restored = read_autosave(path, session)
        if not restored:
            return
        count = sum(len(sec.lines) for sec in restored)
        answer = QMessageBox.question(
            self,
            "이전 작업 복구",
            f"저장하지 못한 공정 {len(restored)}개({count}품목)가 임시저장되어 있습니다.\n불러올까요?",
        )
        if answer == QMessageBox.StandardButton.Yes:
            self._sections = restored
            self._renumber_sections()
            self._append_log(f"임시저장된 공정 {len(restored)}개를 불러왔습니다.")
            self._refresh_state()
        else:
            path.unlink(missing_ok=True)

    def _renumber_sections(self) -> None:
        self._autosave()
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
            answer = QMessageBox.question(
                self,
                "일위대가목록만 저장",
                "담아 둔 공정이 없습니다.\n내역서 없이 일위대가·일위대가목록까지만 저장할까요?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            )
            if answer != QMessageBox.StandardButton.Yes:
                return
        dest = self._chosen_dest_dir()
        if dest is not None:
            self._settings.setValue("dest_dir", self.dest_edit.text().strip())
        session, sections = self._session, list(self._sections)
        self.statusBar().showMessage("내역서를 저장하는 중…")
        mode = self.mode_combo.currentData()
        self._run_background(lambda: save_session(session, sections, dest, mode), self._on_saved)

    def _on_saved(self, dest: Path) -> None:
        try:
            autosave_path(self._chosen_dest_dir()).unlink(missing_ok=True)
        except OSError:
            pass
        self._append_log(f"새 파일 저장: {dest}")
        self.statusBar().showMessage(f"저장 완료 — {Path(dest).name}")
        self._notify_done()
        DoneDialog(self, Path(dest), pumsam_filename(self._selected_discipline())).exec()
        self._refresh_state()
