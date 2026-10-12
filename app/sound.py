"""완료 알림음. Windows 는 winsound 로 메모리의 짧은 차임을 재생하고, 그 밖에는 시스템 벨소리."""

from __future__ import annotations

import io
import math
import struct
import sys
import wave

_RATE = 22050


def chime_wav() -> bytes:
    """도-미-솔 세 음이 이어지는 0.6초 알림음(WAV)."""
    frames = bytearray()
    for freq, length in ((784.0, 0.16), (988.0, 0.16), (1318.5, 0.30)):
        count = int(_RATE * length)
        for i in range(count):
            envelope = min(1.0, i / 200) * math.exp(-3.0 * i / count)
            value = int(9000 * envelope * math.sin(2 * math.pi * freq * i / _RATE))
            frames += struct.pack("<h", value)
    buffer = io.BytesIO()
    with wave.open(buffer, "wb") as out:
        out.setnchannels(1)
        out.setsampwidth(2)
        out.setframerate(_RATE)
        out.writeframes(bytes(frames))
    return buffer.getvalue()


def play_done() -> None:
    try:
        if sys.platform == "win32":
            import winsound

            winsound.PlaySound(chime_wav(), winsound.SND_MEMORY | winsound.SND_ASYNC)
            return
    except Exception:  # noqa: BLE001 — 소리가 안 나도 작업에는 영향이 없다.
        pass
    try:
        from PyQt6.QtWidgets import QApplication

        QApplication.beep()
    except Exception:  # noqa: BLE001
        pass
