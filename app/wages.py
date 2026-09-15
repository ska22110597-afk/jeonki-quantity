"""직종별 노임단가. 프로그램 data 씨앗 + 사용자가 고치는 데이터베이스."""

from __future__ import annotations

from pathlib import Path
from typing import Any

from openpyxl import Workbook, load_workbook

from app.estimate_parse import find_header_row, normalize_header
from app.merge_parse import fill_merged_values, trim_grid
from app.paths import bundled_data_dir, ensure_result_directory, is_allowed_excel, user_database_dir

WAGES_SHEET_NAME = "노임단가"
WAGES_DB_FILENAME = "노임단가.xlsx"
WAGES_HEADERS = ["직종", "노임단가", "비고"]

WageRow = dict[str, Any]


def default_wage_rows() -> list[WageRow]:
    """일위대가 샘플에 있던 직종 단가. 없는 직종은 0으로 두고 직접 입력한다."""
    return [
        {"직종": "내선전공", "노임단가": 276108, "비고": "샘플 일위대가 값. 최신 노임으로 고치세요."},
        {"직종": "보통인부", "노임단가": 0, "비고": "노임단가를 입력하세요."},
        {"직종": "저압케이블전공", "노임단가": 306274, "비고": "샘플 일위대가 값. 최신 노임으로 고치세요."},
        {"직종": "케이블전공", "노임단가": 306274, "비고": "저압케이블전공과 같게 시작. 필요하면 수정."},
        {"직종": "고압케이블전공", "노임단가": 0, "비고": "노임단가를 입력하세요."},
    ]


def wages_db_path(directory: Path | None = None) -> Path:
    return user_database_dir(directory) / WAGES_DB_FILENAME


def bundled_wages_path() -> Path:
    return bundled_data_dir() / WAGES_DB_FILENAME


def _rows_from_sheet(path: Path) -> list[WageRow]:
    workbook = load_workbook(path, data_only=True)
    try:
        sheet = workbook.active
        for candidate in workbook.worksheets:
            if "노임" in str(candidate.title) or "직종" in str(candidate.title):
                sheet = candidate
                break
        grid = fill_merged_values(sheet)
    finally:
        workbook.close()
    table = trim_grid(grid)
    if not table:
        return []
    header_idx = find_header_row(table)
    header = [normalize_header(c) for c in table[header_idx]]
    job_idx = 0
    wage_idx = 1
    note_idx = 2 if len(header) > 2 else None
    for i, token in enumerate(header):
        if token in {"직종", "노무명칭", "명칭", "공사인원"}:
            job_idx = i
        elif token in {"노임단가", "단가", "노임"}:
            wage_idx = i
        elif token == "비고":
            note_idx = i
    rows: list[WageRow] = []
    for source in table[header_idx + 1 :]:
        job = source[job_idx] if job_idx < len(source) else None
        if job is None or normalize_header(job) in {"직종", "노무명칭"}:
            continue
        wage = source[wage_idx] if wage_idx < len(source) else None
        note = source[note_idx] if note_idx is not None and note_idx < len(source) else None
        rows.append({"직종": str(job).strip(), "노임단가": wage, "비고": note})
    return rows


def merge_wage_rows(*groups: list[WageRow]) -> list[WageRow]:
    merged: dict[str, WageRow] = {}
    for group in groups:
        for row in group:
            key = normalize_header(row.get("직종"))
            if not key:
                continue
            merged[key] = dict(row)
    return list(merged.values())


def write_wages_workbook(rows: list[WageRow], path: Path) -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = WAGES_SHEET_NAME
    sheet.append(WAGES_HEADERS)
    for row in rows:
        sheet.append([row.get("직종"), row.get("노임단가"), row.get("비고")])
    workbook.save(path)
    workbook.close()
    return path


def ensure_wages_database(directory: Path | None = None) -> Path:
    """사용자 폴더에 노임단가 파일이 없으면 프로그램 씨앗(또는 기본값)을 복사한다."""
    dest = wages_db_path(directory)
    if dest.exists():
        return dest
    bundled = bundled_wages_path()
    if bundled.exists():
        dest.parent.mkdir(parents=True, exist_ok=True)
        dest.write_bytes(bundled.read_bytes())
        return dest
    write_wages_workbook(default_wage_rows(), dest)
    return dest


def load_wages(directory: Path | None = None) -> list[WageRow]:
    path = ensure_wages_database(directory)
    parsed = _rows_from_sheet(path) if path.exists() else []
    return merge_wage_rows(default_wage_rows(), parsed)


def save_wages(rows: list[WageRow], directory: Path | None = None) -> Path:
    ensure_result_directory(directory)
    return write_wages_workbook(rows, wages_db_path(directory))


def import_wages_file(source_path: Path) -> list[WageRow]:
    path = Path(source_path)
    if not is_allowed_excel(path):
        raise ValueError("xlsx 또는 xlsm 파일만 읽을 수 있습니다.")
    return _rows_from_sheet(path)


def wage_lookup(job_name: Any, rows: list[WageRow]) -> float:
    token = normalize_header(job_name)
    if not token:
        return 0.0
    exact = {normalize_header(r.get("직종")): r.get("노임단가") for r in rows}
    if token in exact:
        return _as_number(exact[token])
    for key, value in exact.items():
        if token in key or key in token:
            return _as_number(value)
    return 0.0


def _as_number(value: Any) -> float:
    if value is None or value == "":
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    try:
        return float(str(value).replace(",", ""))
    except ValueError:
        return 0.0
