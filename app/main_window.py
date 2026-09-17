"""자동 내역서식 프로그램 화면."""

from __future__ import annotations

import math
from pathlib import Path

from PyQt6.QtCore import QEventLoop, QPoint, Qt, QSettings, QThread, QTimer, pyqtSignal
from PyQt6.QtGui import QColor, QIcon, QMouseEvent, QPainter, QPaintEvent, QPen, QPixmap
from PyQt6.QtWidgets import (
    QApplication,
    QButtonGroup,
    QCheckBox,
    QDialog,
    QFileDialog,
    QFrame,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QProgressBar,
    QPushButton,
    QSizePolicy,
    QStatusBar,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from app.discipline import ELECTRIC, TELECOM, normalize_discipline, pumsam_filename
from app.estimate_parse import list_sheet_titles

from app.drop_zone import DropZone
from app.excel_io import (
    ILWIDAE_LIST_SHEET_NAME,
    save_result_workbook,
)
from app.paths import (
    ResultDirectoryError,
    WINDOWS_RESULT_DIR,
    app_icon_path,
    display_result_directory,
    is_windows,
    ui_background_path,
)
from app.pumsam import PUMSAM_SHEET_NAME
from app.version import APP_CONTACT, APP_EXE_NAME, APP_MAKER, APP_NOTICE, APP_TITLE, DROP_HINT
from app.wages import WAGES_SHEET_NAME

APP_STYLESHEET = """
QMainWindow {
    background: #F4EFE4;
    color: #2C281F;
    font-family: "Malgun Gothic", "Apple SD Gothic Neo", "Noto Sans CJK KR", sans-serif;
    font-size: 13px;
}
QWidget#root {
    background: transparent;
    color: #2C281F;
    font-family: "Malgun Gothic", "Apple SD Gothic Neo", "Noto Sans CJK KR", sans-serif;
    font-size: 13px;
}
QFrame#hero {
    background: rgba(44, 40, 31, 0.88);
    border: none;
}
QLabel#appTitle {
    color: #F7F3EA;
    font-size: 18px;
    font-weight: 600;
    letter-spacing: 0.3px;
}
QLabel#appMaker, QLabel#appContact {
    color: #E8DFD0;
    font-size: 12px;
    font-weight: 500;
}
QLabel#appNotice {
    color: #E8C98A;
    font-size: 12px;
    font-weight: 700;
}
QFrame#captionBar {
    background: #1A1A1A;
    border: none;
    min-height: 34px;
    max-height: 34px;
}
QLabel#captionTitle {
    color: #FFFFFF;
    font-size: 12px;
    font-weight: 600;
    letter-spacing: 0.4px;
}
QPushButton#captionMin, QPushButton#captionMax, QPushButton#captionClose {
    background: transparent;
    color: #FFFFFF;
    border: none;
    min-width: 42px;
    max-width: 42px;
    min-height: 34px;
    font-size: 12px;
    font-weight: 700;
}
QPushButton#captionMin:hover, QPushButton#captionMax:hover {
    background: #3A3A3A;
}
QPushButton#captionClose:hover {
    background: #5A5A5A;
    color: #FFFFFF;
}
QFrame#laneForward, QFrame#laneReverse, QFrame#laneQty, QFrame#card, QFrame#partCard {
    background: rgba(255, 252, 246, 0.92);
    border: 1px solid #D4B896;
    border-radius: 8px;
}
QLabel#laneTitle {
    font-size: 13px;
    font-weight: 700;
    color: #5C4A32;
}
QLabel#sectionLabel, QLabel#partCaption {
    color: #7A6A52;
    font-size: 12px;
    font-weight: 700;
}
QLineEdit {
    background: #FFFcf7;
    border: 1px solid #D4B896;
    border-radius: 6px;
    padding: 6px 10px;
    min-height: 22px;
    color: #2C281F;
    font-size: 13px;
    selection-background-color: #5C4A32;
}
QLineEdit#pathEdit {
    font-size: 13px;
    padding: 6px 10px;
}
QLineEdit:read-only {
    color: #4A4338;
}
QCheckBox {
    color: #2C281F;
    spacing: 10px;
    font-size: 13px;
}
QCheckBox::indicator {
    width: 18px;
    height: 18px;
}
QLabel#confirmNote, QLabel#partFileHint, QLabel#dropHint {
    color: #7A6A52;
    font-size: 12px;
}
QPushButton#runButton {
    background: #5C4A32;
    color: #F7F3EA;
    border: none;
    border-radius: 6px;
    padding: 12px 20px;
    font-size: 15px;
    font-weight: 700;
}
QPushButton#runButton:hover {
    background: #3F3424;
}
QPushButton#runButton:disabled {
    background: #C9B8A0;
    color: #F4EDE8;
}
QPushButton#browseButton {
    background: #2C281F;
    color: #F7F3EA;
    border: none;
    border-radius: 6px;
    padding: 6px 16px;
    font-size: 13px;
    font-weight: 700;
    min-height: 34px;
    min-width: 104px;
}
QPushButton#browseButton:hover {
    background: #4A4338;
}
QPushButton#resetButton {
    background: #FFFcf7;
    color: #5C4A32;
    border: 1px solid #D4B896;
    border-radius: 6px;
    padding: 6px 18px;
    font-size: 13px;
    font-weight: 700;
    min-height: 34px;
    min-width: 120px;
}
QPushButton#resetButton:hover {
    background: #F0E6D6;
}
QTextEdit#log {
    background: rgba(44, 40, 31, 0.90);
    color: #E8DFD0;
    border: 1px solid #D4B896;
    border-radius: 6px;
    padding: 12px;
    font-family: "Malgun Gothic", "Noto Sans CJK KR", sans-serif;
    font-size: 12px;
}
QFrame#dropZoneForward, QFrame#dropZoneReverse, QFrame#dropZoneQty {
    background: rgba(255, 252, 246, 0.75);
    border: 1px dashed #C4A574;
    border-radius: 6px;
}
QFrame#dropZoneForward[hover="true"], QFrame#dropZoneReverse[hover="true"], QFrame#dropZoneQty[hover="true"] {
    background: #F3E7D4;
    border: 1px dashed #5C4A32;
}
QFrame#dropZoneForward[loaded="true"], QFrame#dropZoneReverse[loaded="true"], QFrame#dropZoneQty[loaded="true"] {
    background: #EFE6D4;
    border: 1px solid #8A7349;
}
QLabel#dropTitle {
    font-size: 13px;
    font-weight: 700;
    color: #5C4A32;
}
QPushButton#partElectric, QPushButton#partTelecom {
    min-height: 44px;
    border-radius: 6px;
    font-size: 16px;
    font-weight: 700;
    letter-spacing: 1px;
    background: #FFFcf7;
    border: 1px solid #D4B896;
    color: #5C4A32;
}
QPushButton#partElectric:checked, QPushButton#partTelecom:checked {
    background: #5C4A32;
    color: #F7F3EA;
    border: 1px solid #5C4A32;
}
QStatusBar {
    background: #EBE3D4;
    color: #7A6A52;
}
QDialog#doneDialog {
    background: #F4EFE4;
    border: 1px solid #1A1A1A;
}
QFrame#doneHeader {
    background: #2C281F;
    border: none;
}
QLabel#doneTitle {
    color: #E8C98A;
    font-size: 20px;
    font-weight: 700;
    letter-spacing: 2px;
}
QFrame#doneGoldBar {
    background: #D4B896;
    border: none;
    max-height: 3px;
}
QFrame#doneBody {
    background: #F4EFE4;
}
QLabel#doneCaption {
    color: #8A7349;
    font-size: 11px;
    font-weight: 700;
    letter-spacing: 1px;
}
QLabel#donePath, QLabel#doneRef, QLabel#donePumsam {
    color: #2C281F;
    font-size: 13px;
    font-weight: 600;
}
QLabel#doneFoot {
    color: #7A6A52;
    font-size: 12px;
}
QPushButton#doneOk {
    background: #2C281F;
    color: #F7F3EA;
    border: none;
    border-radius: 6px;
    padding: 8px 28px;
    font-size: 13px;
    font-weight: 700;
    min-height: 34px;
    min-width: 96px;
}
QPushButton#doneOk:hover {
    background: #4A4338;
}
QProgressBar#busyGauge {
    border: 1px solid #C9A45C;
    border-radius: 8px;
    background: #E8DFD0;
    color: #2C281F;
    font-size: 12px;
    font-weight: 700;
    min-height: 22px;
    max-height: 22px;
    text-align: center;
}
QProgressBar#busyGauge::chunk {
    background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
        stop:0 #B8893A, stop:0.55 #E8C98A, stop:1 #F3DEAE);
    border-radius: 7px;
    margin: 1px;
}
QLabel#busyMessage, QLabel#busyStatus {
    color: #2C281F;
    font-size: 13px;
    font-weight: 600;
}
QLabel#busyFoot {
    color: #7A6A52;
    font-size: 12px;
}
"""


class DragHeader(QFrame):
    """검정 머리 줄을 잡고 창을 옮긴다. 흰색 윈도우 막대 대신 쓴다."""

    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self._drag: QPoint | None = None

    def paintEvent(self, event: QPaintEvent) -> None:  # noqa: N802
        super().paintEvent(event)
        painter = QPainter(self)
        painter.fillRect(self.rect(), QColor("#1A1A1A"))
        painter.setPen(QPen(QColor(255, 255, 255, 28), 1))
        step = 6
        width = self.width()
        height = self.height()
        for x in range(-height, width + height, step):
            painter.drawLine(x, 0, x + height, height)
        painter.fillRect(0, height - 1, width, 1, QColor("#EDEDED"))
        painter.end()

    def mousePressEvent(self, event: QMouseEvent) -> None:  # noqa: N802
        if event.button() == Qt.MouseButton.LeftButton:
            self._drag = event.globalPosition().toPoint() - self.window().frameGeometry().topLeft()
        super().mousePressEvent(event)

    def mouseMoveEvent(self, event: QMouseEvent) -> None:  # noqa: N802
        if self._drag is not None and event.buttons() & Qt.MouseButton.LeftButton:
            self.window().move(event.globalPosition().toPoint() - self._drag)
        super().mouseMoveEvent(event)

    def mouseReleaseEvent(self, event: QMouseEvent) -> None:  # noqa: N802
        self._drag = None
        super().mouseReleaseEvent(event)


class CaptionBar(DragHeader):
    """메인 창 맨 위 검은 표제. 빗금 무늬를 깔고 검정·흰색만 쓴다."""

    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setObjectName("captionBar")
        self.setFixedHeight(34)
        layout = QHBoxLayout(self)
        layout.setContentsMargins(14, 0, 0, 0)
        layout.setSpacing(0)
        self._title = QLabel(APP_TITLE)
        self._title.setObjectName("captionTitle")
        layout.addWidget(self._title, 1)
        self.min_button = QPushButton("—")
        self.min_button.setObjectName("captionMin")
        self.min_button.setCursor(Qt.CursorShape.PointingHandCursor)
        self.max_button = QPushButton("□")
        self.max_button.setObjectName("captionMax")
        self.max_button.setCursor(Qt.CursorShape.PointingHandCursor)
        self.close_button = QPushButton("×")
        self.close_button.setObjectName("captionClose")
        self.close_button.setCursor(Qt.CursorShape.PointingHandCursor)
        layout.addWidget(self.min_button)
        layout.addWidget(self.max_button)
        layout.addWidget(self.close_button)

    def mouseDoubleClickEvent(self, event: QMouseEvent) -> None:  # noqa: N802
        if event.button() == Qt.MouseButton.LeftButton:
            window = self.window()
            if window.isMaximized():
                window.showNormal()
            else:
                window.showMaximized()
        super().mouseDoubleClickEvent(event)


def _frameless_dialog_flags() -> Qt.WindowType:
    return (
        Qt.WindowType.Dialog
        | Qt.WindowType.FramelessWindowHint
        | Qt.WindowType.WindowStaysOnTopHint
    )


class DoneDialog(QDialog):
    """산출이 끝났을 때 띄우는 완료 창. 배전함 명판처럼 검정이랑 놋쇠색만 쓴다."""

    def __init__(self, parent: QWidget | None, dest: Path, pumsam_name: str) -> None:
        super().__init__(parent)
        self.setObjectName("doneDialog")
        self.setWindowTitle("서식 생성 완료")
        self.setModal(True)
        self.setMinimumWidth(560)
        self.setStyleSheet(APP_STYLESHEET)
        self.setWindowFlags(_frameless_dialog_flags())

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        header = DragHeader()
        header.setObjectName("doneHeader")
        header_layout = QVBoxLayout(header)
        header_layout.setContentsMargins(28, 22, 28, 16)
        header_layout.setSpacing(12)
        title = QLabel("서식 생성 완료!")
        title.setObjectName("doneTitle")
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        header_layout.addWidget(title)
        gold = QFrame()
        gold.setObjectName("doneGoldBar")
        gold.setFixedHeight(3)
        header_layout.addWidget(gold)
        layout.addWidget(header)

        body = QFrame()
        body.setObjectName("doneBody")
        body_layout = QVBoxLayout(body)
        body_layout.setContentsMargins(28, 22, 28, 20)
        body_layout.setSpacing(6)

        def add_row(caption: str, value: str, value_name: str) -> None:
            cap = QLabel(f"{caption} :")
            cap.setObjectName("doneCaption")
            val = QLabel(value)
            val.setObjectName(value_name)
            val.setWordWrap(True)
            val.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
            body_layout.addWidget(cap)
            body_layout.addWidget(val)
            body_layout.addSpacing(10)

        add_row("파일 저장경로", str(dest), "donePath")
        add_row("참고 데이터", f"{PUMSAM_SHEET_NAME}, {WAGES_SHEET_NAME}", "doneRef")
        add_row("사용한 표준품셈", pumsam_name, "donePumsam")

        foot = QLabel("표준품셈·노임단가는 저장 폴더의 데이터베이스에서 고칠 수 있습니다.")
        foot.setObjectName("doneFoot")
        foot.setWordWrap(True)
        body_layout.addWidget(foot)
        body_layout.addSpacing(12)

        ok = QPushButton("확인")
        ok.setObjectName("doneOk")
        ok.setDefault(True)
        ok.clicked.connect(self.accept)
        btn_row = QHBoxLayout()
        btn_row.addStretch()
        btn_row.addWidget(ok)
        body_layout.addLayout(btn_row)
        layout.addWidget(body)


class BusyDialog(QDialog):
    """작업 중 창. 완료 창과 같은 배전함 무늬에, 프로그램 받을 때처럼 게이지를 채운다."""

    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setObjectName("doneDialog")
        self.setWindowTitle("작업 중")
        self.setModal(True)
        self.setMinimumWidth(560)
        self.setStyleSheet(APP_STYLESHEET)
        self.setWindowFlags(_frameless_dialog_flags())
        self._allow_close = False
        self._ticks = 0
        self._finishing = False

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        header = DragHeader()
        header.setObjectName("doneHeader")
        header_layout = QVBoxLayout(header)
        header_layout.setContentsMargins(28, 22, 28, 16)
        header_layout.setSpacing(12)
        title = QLabel("작업 중")
        title.setObjectName("doneTitle")
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        header_layout.addWidget(title)
        gold = QFrame()
        gold.setObjectName("doneGoldBar")
        gold.setFixedHeight(3)
        header_layout.addWidget(gold)
        layout.addWidget(header)

        body = QFrame()
        body.setObjectName("doneBody")
        body_layout = QVBoxLayout(body)
        body_layout.setContentsMargins(28, 22, 28, 24)
        body_layout.setSpacing(10)

        cap = QLabel("진행 상황 :")
        cap.setObjectName("doneCaption")
        body_layout.addWidget(cap)
        self._message = QLabel("산출 파일을 만드는 중입니다.")
        self._message.setObjectName("busyMessage")
        self._message.setWordWrap(True)
        body_layout.addWidget(self._message)

        self._status = QLabel(f"{APP_EXE_NAME}  ·  서식 작성")
        self._status.setObjectName("busyStatus")
        body_layout.addWidget(self._status)

        self._gauge = QProgressBar()
        self._gauge.setObjectName("busyGauge")
        self._gauge.setRange(0, 100)
        self._gauge.setValue(0)
        self._gauge.setTextVisible(True)
        self._gauge.setFormat("%p%")
        self._gauge.setMinimumHeight(22)
        body_layout.addWidget(self._gauge)

        self._foot = QLabel("끝날 때까지 버튼을 다시 누르지 마세요.")
        self._foot.setObjectName("busyFoot")
        self._foot.setWordWrap(True)
        body_layout.addWidget(self._foot)
        layout.addWidget(body)

        self._timer = QTimer(self)
        self._timer.setInterval(120)
        self._timer.timeout.connect(self._tick_gauge)

    @property
    def labelText(self) -> str:  # noqa: N802 — 예전 QProgressDialog 검사와 맞춤
        return f"{self._message.text()}\n{self._foot.text()}"

    def start_gauge(self) -> None:
        self._gauge.setValue(0)
        self._ticks = 0
        self._finishing = False
        self._timer.start()

    def _tick_gauge(self) -> None:
        if self._finishing:
            return
        self._ticks += 1
        elapsed = self._ticks * (self._timer.interval() / 1000.0)
        eased = int(100 * (1.0 - math.exp(-elapsed / 16.0)))
        current = self._gauge.value()
        target = min(95, eased)
        if target > current:
            self._gauge.setValue(target)
            return
        if current < 95 and self._ticks % 10 == 0:
            self._gauge.setValue(current + 1)

    def complete_and_close(self) -> None:
        self._timer.stop()
        self._finishing = True
        self._message.setText("산출 파일을 만들었습니다.")
        QApplication.processEvents()
        loop = QEventLoop(self)
        anim = QTimer(self)
        anim.setInterval(45)

        def step() -> None:
            value = self._gauge.value()
            if value >= 100:
                anim.stop()
                QTimer.singleShot(480, loop.quit)
                return
            bump = 1 if value >= 88 else max(1, (100 - value) // 14)
            self._gauge.setValue(min(100, value + bump))
            QApplication.processEvents()

        anim.timeout.connect(step)
        if self._gauge.value() >= 100:
            QTimer.singleShot(480, loop.quit)
        else:
            anim.start()
            step()
        loop.exec()
        self._allow_close = True
        self.close()

    def closeEvent(self, event) -> None:  # noqa: N802
        if not self._allow_close:
            event.ignore()
            return
        self._timer.stop()
        event.accept()


class PipelineWorker(QThread):
    """엑셀 산출은 별도 줄에서 돌려, 작업 중 게이지가 멈추지 않게 한다."""

    succeeded = pyqtSignal(str)
    failed = pyqtSignal(str, str)

    def __init__(self, kwargs: dict) -> None:
        super().__init__()
        self._kwargs = kwargs

    def run(self) -> None:
        try:
            dest = save_result_workbook(**self._kwargs)
            self.succeeded.emit(str(dest))
        except ResultDirectoryError as exc:
            self.failed.emit("저장 폴더 오류", str(exc))
        except Exception as exc:  # noqa: BLE001 — GUI에서 사용자 메시지로 보여 준다.
            self.failed.emit("처리 실패", str(exc))


SETTINGS_ORG = "전기공사공량산출"
SETTINGS_APP = "전기통신서식생성"
SETTINGS_APP_LEGACY = "GongryangCalc"


def _load_settings() -> QSettings:
    settings = QSettings(SETTINGS_ORG, SETTINGS_APP)
    if not str(settings.value("dest_dir", "") or "").strip():
        legacy = QSettings(SETTINGS_ORG, SETTINGS_APP_LEGACY)
        stored = legacy.value("dest_dir", "")
        if isinstance(stored, str) and stored.strip():
            settings.setValue("dest_dir", stored.strip())
    return settings


class PaperRoot(QWidget):
    """금색 문장을 베이지 바탕 위에 옅게 깔아 둔 화면."""

    def __init__(self, parent=None) -> None:
        super().__init__(parent)
        self.setObjectName("root")
        bg = ui_background_path()
        self._bg = QPixmap(str(bg)) if bg.is_file() else QPixmap()

    def paintEvent(self, event: QPaintEvent) -> None:  # noqa: N802
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.SmoothPixmapTransform)
        painter.fillRect(self.rect(), QColor("#F4EFE4"))
        if not self._bg.isNull():
            pix = self._bg.scaled(
                self.size(),
                Qt.AspectRatioMode.KeepAspectRatio,
                Qt.TransformationMode.SmoothTransformation,
            )
            x = (self.width() - pix.width()) // 2
            y = (self.height() - pix.height()) // 2
            painter.setOpacity(0.26)
            painter.drawPixmap(x, y, pix)
        painter.end()
        super().paintEvent(event)


class MainWindow(QMainWindow):
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
        self.setMinimumSize(1180, 980)
        self.resize(1260, 1080)
        self.setStyleSheet(APP_STYLESHEET)
        icon_file = app_icon_path()
        if icon_file.is_file():
            self.setWindowIcon(QIcon(str(icon_file)))

        self._settings = _load_settings()
        self._fwd_compare_path: Path | None = None
        self._fwd_ilwidae_path: Path | None = None
        self._fwd_estimate_path: Path | None = None
        self._rev_estimate_path: Path | None = None
        self._rev_ilwidae_path: Path | None = None
        self._qty_estimate_path: Path | None = None
        self._busy = False
        self._busy_dialog: BusyDialog | None = None
        self._worker: PipelineWorker | None = None
        self._build_ui()
        self._refresh_run_enabled()

    def _saved_dest_dir(self) -> str:
        stored = self._settings.value("dest_dir", "")
        if isinstance(stored, str) and stored.strip():
            return stored.strip()
        return display_result_directory()

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
        hero_layout.setContentsMargins(36, 22, 36, 20)
        hero_layout.setSpacing(10)
        title = QLabel(APP_TITLE)
        title.setObjectName("appTitle")
        hero_layout.addWidget(title)
        gold = QFrame()
        gold.setObjectName("doneGoldBar")
        gold.setFixedHeight(3)
        hero_layout.addWidget(gold)
        credits = QHBoxLayout()
        credits.setContentsMargins(0, 4, 0, 0)
        credits.setSpacing(16)
        left = QVBoxLayout()
        left.setSpacing(4)
        maker = QLabel(APP_MAKER)
        maker.setObjectName("appMaker")
        contact = QLabel(APP_CONTACT)
        contact.setObjectName("appContact")
        left.addWidget(maker)
        left.addWidget(contact)
        notice = QLabel(APP_NOTICE)
        notice.setObjectName("appNotice")
        notice.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        credits.addLayout(left, 1)
        credits.addWidget(notice, 0, Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        hero_layout.addLayout(credits)
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
        forward_title = QLabel("정방향  ·  단가대비표 → 일위대가목록")
        forward_title.setObjectName("laneTitle")
        forward_title.setWordWrap(True)
        forward_layout.addWidget(forward_title)

        self.compare_drop = DropZone(
            title="단가대비표",
            hint=DROP_HINT,
            dialog_title="단가대비표 엑셀 선택",
            tone="forward",
        )
        self.compare_drop.file_dropped.connect(self._on_fwd_compare_dropped)
        forward_layout.addWidget(self.compare_drop)
        self.ilwidae_drop = None
        self.forward_estimate_drop = None
        self.compare_drop.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        forward_layout.addStretch(1)
        lanes.addWidget(forward_lane, 1)

        reverse_lane = QFrame()
        reverse_lane.setObjectName("laneReverse")
        reverse_layout = QVBoxLayout(reverse_lane)
        reverse_layout.setContentsMargins(12, 10, 12, 12)
        reverse_layout.setSpacing(8)
        reverse_title = QLabel("역방향  ·  일위대가목록 → 단가대비표")
        reverse_title.setObjectName("laneTitle")
        reverse_title.setWordWrap(True)
        reverse_layout.addWidget(reverse_title)

        self.drop_zone = DropZone(
            title="일위대가목록",
            hint=DROP_HINT,
            dialog_title="일위대가목록 엑셀 선택",
            tone="reverse",
        )
        self.drop_zone.file_dropped.connect(self._on_rev_estimate_dropped)
        reverse_layout.addWidget(self.drop_zone)
        self.reverse_ilwidae_drop = None
        self.drop_zone.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        reverse_layout.addStretch(1)
        lanes.addWidget(reverse_lane, 1)
        body_layout.addLayout(lanes)

        qty_lane = QFrame()
        qty_lane.setObjectName("laneQty")
        qty_layout = QVBoxLayout(qty_lane)
        qty_layout.setContentsMargins(12, 10, 12, 10)
        qty_layout.setSpacing(8)
        qty_title = QLabel("공량산출  ·  일위대가목록 → 공량산출서")
        qty_title.setObjectName("laneTitle")
        qty_title.setWordWrap(True)
        qty_layout.addWidget(qty_title)
        self.quantity_drop = DropZone(
            title="일위대가목록",
            hint=DROP_HINT,
            dialog_title="공량산출용 일위대가목록 엑셀 선택",
            tone="quantity",
        )
        self.quantity_drop.setMaximumHeight(88)
        self.quantity_drop.setMinimumHeight(72)
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
        status.setSizeGripEnabled(True)
        status.showMessage("대기 — 정방향 또는 역방향 칸에 엑셀을 놓고 저장 폴더를 확인하세요.")
        self.setStatusBar(status)
        self._write_startup_log()
        self._refresh_filename_hint()

    def _toggle_maximized(self) -> None:
        if self.isMaximized():
            self.showNormal()
            self.caption.max_button.setText("□")
        else:
            self.showMaximized()
            self.caption.max_button.setText("❐")

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
        self.part_file_hint.setText(f"현재 선택된 적용 품셈 : {filename}")

    def _on_part_changed(self) -> None:
        disc = self._selected_discipline()
        self._settings.setValue("discipline", disc)
        self._refresh_part_hint()
        self.statusBar().showMessage(f"{disc} 표준품셈을 사용합니다.")
        self._append_log(f"표준품셈 파트: {disc} → {pumsam_filename(disc)}")

    def _write_startup_log(self) -> None:
        self._append_log("원본 엑셀은 읽기만 합니다. 병합 셀은 메모리에서 채웁니다.")
        self._append_log(f"저장 폴더: {self.dest_edit.text()}")
        self._append_log("왼쪽(정방향): 단가대비표 → 일위대가목록")
        self._append_log("오른쪽(역방향): 일위대가목록 → 단가대비표")
        self._append_log("아래(공량산출): 일위대가목록 시트가 있는 엑셀 → 공량산출서")
        self._append_log(f"표준품셈: {pumsam_filename(self._selected_discipline())}")
        self._append_log("노임단가는 2026년 하반기 시중노임(2026.9.1)을 넣어 두었습니다. 저장 폴더의 데이터베이스에서 고칠 수 있습니다.")

    def _refresh_filename_hint(self) -> None:
        if self._has_forward() and not self._has_reverse() and not self._has_quantity():
            name = "일위대가목록_결과_날짜시간.xlsx"
        elif self._has_reverse() and not self._has_forward() and not self._has_quantity():
            name = "단가대비표_결과_날짜시간.xlsx"
        elif self._has_quantity() and not self._has_forward() and not self._has_reverse():
            name = "공량산출_결과_날짜시간.xlsx"
        else:
            name = "일위대가목록_결과_날짜시간.xlsx"
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
        for zone in (self.compare_drop, self.drop_zone, self.quantity_drop):
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
        if self._busy:
            self.run_button.setEnabled(False)
            return
        ready = self._has_input() and self.confirm_box.isChecked()
        self.run_button.setEnabled(ready)

    def _lock_run_ui(self) -> None:
        self._busy = True
        self.run_button.setEnabled(False)
        self.run_button.setText("작업 중 — 누르지 마세요")
        self.reset_button.setEnabled(False)
        self.browse_button.setEnabled(False)
        self.confirm_box.setEnabled(False)
        QApplication.setOverrideCursor(Qt.CursorShape.WaitCursor)

    def _unlock_run_ui(self) -> None:
        QApplication.restoreOverrideCursor()
        self.run_button.setText("산출 및 저장")
        self.reset_button.setEnabled(True)
        self.browse_button.setEnabled(True)
        self.confirm_box.setEnabled(True)
        self._busy = False
        self._refresh_run_enabled()

    def _has_input(self) -> bool:
        return self._has_forward() or self._has_reverse() or self._has_quantity()

    def _has_forward(self) -> bool:
        return self._fwd_compare_path is not None

    def _has_reverse(self) -> bool:
        return self._rev_estimate_path is not None

    def _has_quantity(self) -> bool:
        return self._qty_estimate_path is not None

    def _sync_source_edit(self) -> None:
        parts: list[str] = []
        if self._fwd_compare_path is not None:
            parts.append(f"정·단가대비표: {self._fwd_compare_path}")
        if self._fwd_ilwidae_path is not None:
            parts.append(f"정·일위대가: {self._fwd_ilwidae_path}")
        if self._rev_estimate_path is not None:
            parts.append(f"역·일위대가목록: {self._rev_estimate_path}")
        if self._rev_ilwidae_path is not None:
            parts.append(f"역·일위대가: {self._rev_ilwidae_path}")
        if self._qty_estimate_path is not None:
            parts.append(f"공량·일위대가목록: {self._qty_estimate_path}")
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
        if self.ilwidae_drop is not None:
            self.ilwidae_drop.set_loaded(path.name)
        self._sync_source_edit()
        self.statusBar().showMessage(f"정방향 일위대가 선택됨 (읽기 전용): {path.name}")
        self._append_log(f"정방향 일위대가 로드 대기: {path}")
        self._refresh_run_enabled()

    def _on_fwd_estimate_dropped(self, path_text: str) -> None:
        """정방향은 일위대가목록을 입력으로 받지 않는다."""
        return

    def _on_rev_estimate_dropped(self, path_text: str) -> None:
        path = Path(path_text)
        self._rev_estimate_path = path
        self.drop_zone.set_loaded(path.name)
        self._sync_source_edit()
        self.statusBar().showMessage(f"역방향 일위대가목록 선택됨 (읽기 전용): {path.name}")
        self._append_log(f"역방향 일위대가목록 로드 대기: {path}")
        self._refresh_run_enabled()

    def _on_rev_ilwidae_dropped(self, path_text: str) -> None:
        path = Path(path_text)
        self._rev_ilwidae_path = path
        if self.reverse_ilwidae_drop is not None:
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
        try:
            titles = list_sheet_titles(path)
        except Exception as exc:  # noqa: BLE001 — 잘못된 파일은 창으로 알린다.
            QMessageBox.warning(self, "파일을 열 수 없음", str(exc))
            return
        if ILWIDAE_LIST_SHEET_NAME not in titles:
            QMessageBox.warning(
                self,
                "시트 없음",
                "일위대가목록 시트가 없습니다.\n해당 시트가 있는 엑셀을 놓아 주세요.",
            )
            self._append_log(f"공량산출 거부: 일위대가목록 시트 없음 ({path.name})")
            return
        self._qty_estimate_path = path
        self.quantity_drop.set_loaded(path.name)
        self._sync_source_edit()
        self.statusBar().showMessage(f"공량산출용 일위대가목록 선택됨 (읽기 전용): {path.name}")
        self._append_log(f"공량산출 일위대가목록 로드 대기: {path}")
        self._refresh_run_enabled()

    def _open_busy_dialog(self) -> BusyDialog:
        dialog = BusyDialog(self)
        dialog.setWindowModality(Qt.WindowModality.ApplicationModal)
        dialog.show()
        dialog.raise_()
        dialog.activateWindow()
        dialog.start_gauge()
        QApplication.processEvents()
        return dialog

    def _on_run(self) -> None:
        if self._busy:
            return
        if not self._has_input():
            QMessageBox.warning(self, "파일 없음", "단가대비표 또는 일위대가목록을 먼저 놓아 주세요.")
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

        self._lock_run_ui()
        self.statusBar().showMessage("산출 파일을 생성하는 중…")
        discipline = self._selected_discipline()
        self._append_log(f"원본 읽기 전용 · {pumsam_filename(discipline)} · 노임단가 결합 · 결과 엑셀 생성")

        if self._has_forward():
            mode = "forward"
            unit_price_path = self._fwd_compare_path
            ilwidae_path = self._fwd_ilwidae_path
            estimate_path = None
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

        self._busy_dialog = self._open_busy_dialog()
        self._worker = PipelineWorker(
            {
                "dest_dir": dest_dir,
                "unit_price_path": unit_price_path,
                "ilwidae_path": ilwidae_path,
                "estimate_path": estimate_path,
                "discipline": discipline,
                "mode": mode,
            }
        )
        self._worker.succeeded.connect(self._on_pipeline_succeeded)
        self._worker.failed.connect(self._on_pipeline_failed)
        self._worker.start()

    def _finish_busy_dialog(self) -> None:
        if self._busy_dialog is not None:
            self._busy_dialog.complete_and_close()
            self._busy_dialog = None
        QApplication.processEvents()
        self._unlock_run_ui()

    def _on_pipeline_succeeded(self, dest_text: str) -> None:
        self._finish_busy_dialog()
        dest = Path(dest_text)
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
        DoneDialog(self, dest, pumsam_filename(self._selected_discipline())).exec()
        self._refresh_run_enabled()

    def _on_pipeline_failed(self, title: str, text: str) -> None:
        self._append_log(f"{title}: {text}")
        self.statusBar().showMessage("실패" if title == "처리 실패" else title)
        self._finish_busy_dialog()
        QMessageBox.critical(self, title, text)
