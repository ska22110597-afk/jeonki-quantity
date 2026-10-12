"""최초 실행 인증 — 숫자 코드를 한 번 입력하면 그 PC에서는 다시 묻지 않는다.

코드 원문은 프로그램 안에 두지 않고 해시만 둔다. 인증 표식은 그 PC의 고유값(MachineGuid)과 묶어서
저장하므로, 표식 파일만 다른 PC로 복사해서는 통과되지 않는다(그 PC에서 코드를 다시 입력해야 한다).
※ 코드를 아는 사람은 누구나 인증할 수 있으므로 '가벼운 복제·공유 억제' 용도이며 완전한 차단은 아니다.
"""

from __future__ import annotations

import hashlib
import hmac
import os
import sys
import uuid
from pathlib import Path

_SALT = "jeonki-quantity/activation/v1"
# 코드 "373302" 의 SHA-256 (코드 원문은 저장하지 않는다)
CODE_HASH = hashlib.sha256((_SALT + "373302").encode("utf-8")).hexdigest()


def _hash_code(code: str) -> str:
    return hashlib.sha256((_SALT + code.strip()).encode("utf-8")).hexdigest()


def check_code(code: str) -> bool:
    return hmac.compare_digest(_hash_code(code), CODE_HASH)


def machine_id() -> str:
    if sys.platform == "win32":
        try:
            import winreg

            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Cryptography", 0, winreg.KEY_READ | winreg.KEY_WOW64_64KEY) as key:
                return str(winreg.QueryValueEx(key, "MachineGuid")[0])
        except OSError:
            pass
    return f"mac-{uuid.getnode():x}"


def marker_path() -> Path:
    base = os.environ.get("LOCALAPPDATA") or os.environ.get("XDG_CONFIG_HOME") or str(Path.home() / ".config")
    return Path(base) / "전기통신서식생성" / "activation.dat"


def _token(machine: str) -> str:
    return hmac.new(CODE_HASH.encode(), (machine + _SALT).encode("utf-8"), hashlib.sha256).hexdigest()


def is_activated(path: Path | None = None) -> bool:
    target = path or marker_path()
    try:
        return hmac.compare_digest(target.read_text(encoding="utf-8").strip(), _token(machine_id()))
    except OSError:
        return False


def activate(code: str, path: Path | None = None) -> bool:
    if not check_code(code):
        return False
    target = path or marker_path()
    try:
        target.parent.mkdir(parents=True, exist_ok=True)
        target.write_text(_token(machine_id()), encoding="utf-8")
    except OSError:
        return False
    return True


def ask_activation(parent=None) -> bool:
    """인증 창을 띄운다. 통과하면 True, 취소하면 False."""
    from PyQt6.QtWidgets import QInputDialog, QLineEdit, QMessageBox

    message = "처음 실행입니다. 인증 번호를 입력해 주세요."
    while True:
        code, ok = QInputDialog.getText(parent, "인증", message, QLineEdit.EchoMode.Password)
        if not ok:
            return False
        if activate(code):
            return True
        QMessageBox.warning(parent, "인증 실패", "인증 번호가 맞지 않습니다. 다시 입력해 주세요.")
        message = "인증 번호를 다시 입력해 주세요."
