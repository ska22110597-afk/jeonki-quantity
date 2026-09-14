from __future__ import annotations

from datetime import datetime
from pathlib import Path

import pytest

from app.paths import (
    RESULT_FOLDER_NAME,
    WINDOWS_RESULT_DIR,
    assert_safe_save,
    build_result_filename,
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
    assert name == "단가대비_공량산출_결과_20260915_093007.xlsx"


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
