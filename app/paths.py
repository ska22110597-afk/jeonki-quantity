"""C드라이브 전용 결과 저장 경로.

원드라이브(OneDrive) 동기화 경로를 피하기 위해, Windows에서는
항상 ``C:\\전기공사_공량산출_결과`` 로컬 폴더만 사용한다.
원본 엑셀 경로는 결과 경로로 쓰지 않는다.
"""

from __future__ import annotations

import os
import sys
from datetime import datetime
from pathlib import Path

RESULT_FOLDER_NAME = "전기공사_공량산출_결과"
WINDOWS_RESULT_DIR = Path(r"C:\전기공사_공량산출_결과")
RESULT_FILENAME_PREFIX = "단가대비_공량산출_결과_"
ALLOWED_EXCEL_SUFFIXES = (".xlsx", ".xlsm")


def is_windows() -> bool:
    return sys.platform == "win32"


def get_result_directory() -> Path:
    """결과 파일을 둘 로컬 폴더를 반환한다.

    우선순위:
    1. 테스트용 환경변수 ``JEONKI_RESULT_DIR`` (있으면 그대로 사용)
    2. Windows: ``C:\\전기공사_공량산출_결과`` 고정
    3. 그 외 OS: 홈 디렉터리 아래 동일 폴더명 (개발·검증용)
    """
    override = os.environ.get("JEONKI_RESULT_DIR", "").strip()
    if override:
        return Path(override)
    if is_windows():
        return Path("C:/") / RESULT_FOLDER_NAME
    return Path.home() / RESULT_FOLDER_NAME


def display_result_directory() -> str:
    """UI에 보여줄 저장 경로 문자열."""
    return str(get_result_directory())


def ensure_result_directory(directory: Path | None = None) -> Path:
    """결과 폴더가 없으면 생성하고, 생성된 경로를 반환한다."""
    target = Path(directory) if directory is not None else get_result_directory()
    target.mkdir(parents=True, exist_ok=True)
    return target


def build_result_filename(now: datetime | None = None) -> str:
    stamp = (now or datetime.now()).strftime("%Y%m%d_%H%M%S")
    return f"{RESULT_FILENAME_PREFIX}{stamp}.xlsx"


def build_result_path(directory: Path | None = None, now: datetime | None = None) -> Path:
    return ensure_result_directory(directory) / build_result_filename(now)


def is_allowed_excel(path: Path) -> bool:
    return path.suffix.lower() in ALLOWED_EXCEL_SUFFIXES


def assert_safe_save(source: Path, dest: Path) -> None:
    """원본을 덮어쓰거나 같은 파일을 가리키면 즉시 중단한다."""
    source_resolved = source.expanduser().resolve()
    dest_resolved = dest.expanduser().resolve()
    if dest_resolved.suffix.lower() != ".xlsx":
        raise ValueError("결과 파일은 .xlsx 만 허용합니다.")
    if source_resolved == dest_resolved:
        raise ValueError("원본 파일을 덮어쓸 수 없습니다. 새 결과 파일만 생성합니다.")
    if dest_resolved.parent == source_resolved.parent and dest_resolved.name == source_resolved.name:
        raise ValueError("원본과 동일한 파일명으로는 저장하지 않습니다.")
