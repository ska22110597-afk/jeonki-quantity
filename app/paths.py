"""C드라이브 전용 결과 저장 경로.

원드라이브(OneDrive) 동기화 경로를 피하기 위해, Windows에서는
항상 ``C:\\전기공사_공량산출_결과`` 로컬 폴더만 사용한다.

PyInstaller --onefile 로 묶이면 작업 폴더가 임시 해제 경로(sys._MEIPASS)가
되므로, 결과 경로는 실행 파일 위치나 번들 경로가 아닌 C드라이브 절대 경로만 쓴다.
"""

from __future__ import annotations

import os
import sys
import tempfile
from datetime import datetime
from pathlib import Path

RESULT_FOLDER_NAME = "전기공사_공량산출_결과"
WINDOWS_RESULT_DIR = Path(r"C:\전기공사_공량산출_결과")
RESULT_FILENAME_PREFIX = "단가대비_공량산출_결과_"
ALLOWED_EXCEL_SUFFIXES = (".xlsx", ".xlsm")


class ResultDirectoryError(OSError):
    """결과 폴더를 만들거나 쓸 수 없을 때."""


def is_windows() -> bool:
    return sys.platform == "win32"


def is_frozen() -> bool:
    """PyInstaller 등으로 패키징된 실행 파일인지 여부."""
    return bool(getattr(sys, "frozen", False))


def bundle_extract_dir() -> Path | None:
    """onefile 임시 해제 폴더. 결과 저장에는 절대 사용하지 않는다."""
    meipass = getattr(sys, "_MEIPASS", None)
    if not meipass:
        return None
    return Path(meipass)


def get_result_directory() -> Path:
    """결과 파일을 둘 로컬 폴더를 반환한다.

    우선순위:
    1. 테스트용 환경변수 ``JEONKI_RESULT_DIR`` (있으면 그대로 사용)
    2. Windows: ``C:\\전기공사_공량산출_결과`` 고정 (exe 임시폴더 무시)
    3. 그 외 OS: 홈 디렉터리 아래 동일 폴더명 (개발·검증용)
    """
    override = os.environ.get("JEONKI_RESULT_DIR", "").strip()
    if override:
        return Path(override)

    if is_windows():
        # Path("C:/") 는 Windows에서 C:\ 로 정규화되는 절대 경로.
        # sys._MEIPASS / exe 상대 경로는 쓰지 않는다.
        return Path("C:/") / RESULT_FOLDER_NAME
    return Path.home() / RESULT_FOLDER_NAME


def display_result_directory() -> str:
    """UI에 보여줄 저장 경로 문자열."""
    if is_windows():
        return str(WINDOWS_RESULT_DIR)
    return str(get_result_directory())


def _reject_bundle_path(target: Path) -> Path:
    """상대 경로나 PyInstaller 임시 폴더로 떨어지지 않게 절대 경로로 고정한다."""
    extract = bundle_extract_dir()
    if extract is not None:
        try:
            target.expanduser().resolve().relative_to(extract.resolve())
        except (ValueError, OSError):
            pass
        else:
            raise ResultDirectoryError(
                "결과 폴더가 실행 파일 임시 경로로 잡혔습니다. "
                f"저장은 {WINDOWS_RESULT_DIR} 만 사용합니다."
            )

    if not target.is_absolute():
        if is_windows():
            return Path("C:/") / RESULT_FOLDER_NAME
        return Path.home() / RESULT_FOLDER_NAME
    return target


def ensure_result_directory(directory: Path | None = None) -> Path:
    """결과 폴더가 없으면 생성하고, 쓰기 가능한지 확인한 뒤 경로를 반환한다."""
    target = Path(directory) if directory is not None else get_result_directory()
    target = _reject_bundle_path(target.expanduser())

    try:
        target.mkdir(parents=True, exist_ok=True)
    except PermissionError as exc:
        raise ResultDirectoryError(
            f"저장 폴더를 만들 권한이 없습니다: {target}\n"
            "C드라이브 로컬 폴더 쓰기 권한을 확인하거나, "
            "관리자 권한으로 한 번 실행해 주세요."
        ) from exc
    except OSError as exc:
        raise ResultDirectoryError(
            f"저장 폴더를 만들 수 없습니다: {target}\n{exc}"
        ) from exc

    probe: Path | None = None
    try:
        fd, probe_name = tempfile.mkstemp(prefix=".jq_write_", suffix=".tmp", dir=str(target))
        os.close(fd)
        probe = Path(probe_name)
    except OSError as exc:
        raise ResultDirectoryError(
            f"저장 폴더에 파일을 쓸 수 없습니다: {target}\n{exc}"
        ) from exc
    finally:
        if probe is not None:
            try:
                probe.unlink(missing_ok=True)
            except OSError:
                pass

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
