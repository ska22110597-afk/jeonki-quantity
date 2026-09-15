from __future__ import annotations

import sys
from datetime import datetime
from pathlib import Path
from unittest.mock import patch

import pytest

from app.paths import (
    RESULT_FOLDER_NAME,
    WINDOWS_RESULT_DIR,
    ResultDirectoryError,
    assert_safe_save,
    build_result_filename,
    ensure_result_directory,
    get_result_directory,
    is_allowed_excel,
    is_windows,
)


def test_windows_result_dir_is_c_drive_local() -> None:
    assert RESULT_FOLDER_NAME == "전기공사_공량산출_결과"
    assert str(WINDOWS_RESULT_DIR) == r"C:\전기공사_공량산출_결과"
    if is_windows():
        resolved = get_result_directory()
        assert resolved.drive == "C:"
        assert resolved.name == RESULT_FOLDER_NAME


def test_result_filename_uses_timestamp() -> None:
    name = build_result_filename(datetime(2026, 9, 15, 9, 30, 7))
    assert name == "공량산출_결과_20260915_093007.xlsx"


def test_allowed_excel_suffixes() -> None:
    assert is_allowed_excel(Path("단가대비표.xlsx"))
    assert is_allowed_excel(Path("표.xlsm"))
    assert not is_allowed_excel(Path("표.xls"))
    assert not is_allowed_excel(Path("표.csv"))


def test_result_dir_env_override(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    monkeypatch.setenv("JEONKI_RESULT_DIR", str(tmp_path / "out"))
    assert get_result_directory() == tmp_path / "out"


def test_assert_safe_save_rejects_overwrite(tmp_path: Path) -> None:
    source = tmp_path / "원본.xlsx"
    source.write_bytes(b"dummy")
    with pytest.raises(ValueError, match="덮어쓸 수 없습니다"):
        assert_safe_save(source, source)


def test_frozen_result_dir_ignores_meipass(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    monkeypatch.setattr(sys, "frozen", True, raising=False)
    monkeypatch.setattr(sys, "_MEIPASS", str(tmp_path / "_MEIabc123"), raising=False)
    directory = get_result_directory()
    assert "_MEIabc123" not in str(directory)
    assert RESULT_FOLDER_NAME in str(directory)


def test_ensure_result_directory_creates_and_is_writable(tmp_path: Path) -> None:
    target = tmp_path / "전기공사_공량산출_결과"
    created = ensure_result_directory(target)
    assert created == target
    assert target.is_dir()
    leftover = list(target.glob(".jq_write_*"))
    assert leftover == []


def test_ensure_rejects_pyinstaller_extract_dir(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    meipass = tmp_path / "_MEIxxxx"
    meipass.mkdir()
    monkeypatch.setattr(sys, "frozen", True, raising=False)
    monkeypatch.setattr(sys, "_MEIPASS", str(meipass), raising=False)
    with pytest.raises(ResultDirectoryError, match="임시 경로"):
        ensure_result_directory(meipass / "결과")


def test_ensure_wraps_permission_error(tmp_path: Path) -> None:
    target = tmp_path / "locked"
    with patch.object(Path, "mkdir", side_effect=PermissionError("denied")):
        with pytest.raises(ResultDirectoryError, match="권한이 없습니다"):
            ensure_result_directory(target)
