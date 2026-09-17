"""엑셀 파일을 끌어다 놓는 드롭 존."""

from __future__ import annotations

from pathlib import Path

from PyQt6.QtCore import Qt, pyqtSignal
from PyQt6.QtGui import QDragEnterEvent, QDragLeaveEvent, QDropEvent, QMouseEvent
from PyQt6.QtWidgets import QFileDialog, QFrame, QLabel, QVBoxLayout

from app.paths import is_allowed_excel


class DropZone(QFrame):
    file_dropped = pyqtSignal(str)

    def __init__(
        self,
        parent=None,
        title: str = "내역서 엑셀 파일을 여기에 놓으세요",
        hint: str = "xlsx · xlsm  ·  클릭하면 파일 선택  ·  원본은 읽기만 합니다",
        dialog_title: str = "엑셀 선택",
        tone: str = "forward",
    ) -> None:
        super().__init__(parent)
        names = {
            "forward": "dropZoneForward",
            "reverse": "dropZoneReverse",
            "quantity": "dropZoneQty",
        }
        self.setObjectName(names.get(tone, "dropZoneForward"))
        self.setAcceptDrops(True)
        self.setCursor(Qt.CursorShape.PointingHandCursor)
        self.setMinimumHeight(72)
        self.setMaximumHeight(96)
        self._idle_title = title
        self._idle_hint = hint
        self._dialog_title = dialog_title

        self._title = QLabel(title)
        self._title.setObjectName("dropTitle")
        self._title.setAlignment(Qt.AlignmentFlag.AlignCenter)

        self._hint = QLabel(hint)
        self._hint.setObjectName("dropHint")
        self._hint.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self._hint.setWordWrap(True)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(10, 6, 10, 6)
        layout.setSpacing(2)
        layout.addWidget(self._title)
        layout.addWidget(self._hint)

    def set_loaded(self, filename: str) -> None:
        self._title.setText(filename)
        self._hint.setText("다른 파일을 놓으면 교체됩니다. 원본은 수정하지 않습니다.")
        self.setProperty("loaded", True)
        self.style().unpolish(self)
        self.style().polish(self)

    def reset(self) -> None:
        self._title.setText(self._idle_title)
        self._hint.setText(self._idle_hint)
        self.setProperty("loaded", False)
        self.style().unpolish(self)
        self.style().polish(self)

    def _accept_url(self, path: Path) -> bool:
        return path.is_file() and is_allowed_excel(path)

    def dragEnterEvent(self, event: QDragEnterEvent) -> None:  # noqa: N802
        if self._first_excel(event) is not None:
            event.acceptProposedAction()
            self.setProperty("hover", True)
            self.style().unpolish(self)
            self.style().polish(self)
        else:
            event.ignore()

    def dragLeaveEvent(self, event: QDragLeaveEvent) -> None:  # noqa: N802
        self.setProperty("hover", False)
        self.style().unpolish(self)
        self.style().polish(self)
        event.accept()

    def dropEvent(self, event: QDropEvent) -> None:  # noqa: N802
        self.setProperty("hover", False)
        self.setProperty("loaded", False)
        self.style().unpolish(self)
        self.style().polish(self)
        path = self._first_excel(event)
        if path is None:
            event.ignore()
            return
        event.acceptProposedAction()
        self.file_dropped.emit(str(path))

    def mouseReleaseEvent(self, event: QMouseEvent) -> None:  # noqa: N802
        if event.button() == Qt.MouseButton.LeftButton:
            chosen, _ = QFileDialog.getOpenFileName(
                self,
                self._dialog_title,
                "",
                "Excel (*.xlsx *.xlsm)",
            )
            if chosen:
                self.file_dropped.emit(chosen)
        super().mouseReleaseEvent(event)

    def _first_excel(self, event: QDragEnterEvent | QDropEvent) -> Path | None:
        mime = event.mimeData()
        if mime is None or not mime.hasUrls():
            return None
        for url in mime.urls():
            path = Path(url.toLocalFile())
            if self._accept_url(path):
                return path
        return None
