"""품셈표 데이터베이스. C드라이브 로컬 파일에 쌓고, 결과 엑셀에도 시트로 넣는다."""

from __future__ import annotations

from pathlib import Path
from typing import Any

from openpyxl import Workbook, load_workbook

from app.estimate_parse import find_column_index, find_header_row, lookup_key, normalize_header
from app.merge_parse import SheetRows, fill_merged_values, trim_grid
from app.paths import ensure_result_directory, get_result_directory, is_allowed_excel

PUMSAM_SHEET_NAME = "품셈표"
PUMSAM_DB_FILENAME = "품셈표_데이터베이스.xlsx"
PUMSAM_HEADERS = ["검색키", "명칭", "규격", "단위", "노무명칭", "품셈", "할증%", "품셈근거"]

PumsamRow = dict[str, Any]


def pumsam_db_path(directory: Path | None = None) -> Path:
    folder = Path(directory) if directory is not None else get_result_directory()
    return folder / PUMSAM_DB_FILENAME


def _hi_spec(mm: int) -> str:
    return f"HI {mm} mm"


def default_pumsam_rows() -> list[PumsamRow]:
    """사용자가 보여 준 경질비닐전선관 품셈 초기값."""
    buried = {
        16: 0.050,
        22: 0.060,
        28: 0.080,
        36: 0.100,
        42: 0.130,
        54: 0.190,
        70: 0.280,
        82: 0.370,
        92: 0.450,
        104: 0.460,
    }
    exposed = {
        16: 0.060,
        22: 0.072,
        28: 0.096,
        36: 0.120,
        42: 0.156,
        54: 0.228,
        70: 0.336,
        82: 0.444,
        92: 0.540,
        104: 0.552,
    }
    rows: list[PumsamRow] = []
    for mm, value in buried.items():
        name = "경질비닐전선관_지중"
        spec = _hi_spec(mm)
        rows.append(
            {
                "검색키": lookup_key(name, spec),
                "명칭": name,
                "규격": spec,
                "단위": "M",
                "노무명칭": "내선전공",
                "품셈": value,
                "할증%": 100,
                "품셈근거": "전기5-1",
            }
        )
    for mm, value in exposed.items():
        name = "경질비닐전선관_노출"
        spec = _hi_spec(mm)
        rows.append(
            {
                "검색키": lookup_key(name, spec),
                "명칭": name,
                "규격": spec,
                "단위": "M",
                "노무명칭": "내선전공",
                "품셈": value,
                "할증%": 120,
                "품셈근거": "전기5-1",
            }
        )
    return rows


def _row_to_dict(values: list[Any]) -> PumsamRow | None:
    name = values[1] if len(values) > 1 else None
    spec = values[2] if len(values) > 2 else None
    if name is None and spec is None:
        return None
    return {
        "검색키": values[0] or lookup_key(name, spec),
        "명칭": name,
        "규격": spec,
        "단위": values[3] if len(values) > 3 else None,
        "노무명칭": values[4] if len(values) > 4 else None,
        "품셈": values[5] if len(values) > 5 else None,
        "할증%": values[6] if len(values) > 6 else None,
        "품셈근거": values[7] if len(values) > 7 else None,
    }


def rows_from_grid(table: SheetRows) -> list[PumsamRow]:
    if not table:
        return []
    header = table[0]
    # 품셈표는 명칭이 두 번(품목 / 노무) 나올 수 있다.
    name_idx = find_column_index(header, "명칭")
    spec_idx = find_column_index(header, "규격")
    unit_idx = find_column_index(header, "단위")
    labor_idx = None
    pumsam_idx = None
    rate_idx = None
    ref_idx = None
    seen_name = False
    for i, cell in enumerate(header):
        token = normalize_header(cell)
        if token in {"명칭", "품명", "품목"}:
            if not seen_name:
                name_idx = i
                seen_name = True
            else:
                labor_idx = i
        elif token in {"노무명칭", "직종", "공사인원"}:
            labor_idx = i
        elif token == "품셈":
            pumsam_idx = i
        elif token in {"할증%", "할증", "노무할증%", "할증률"}:
            rate_idx = i
        elif token in {"품셈근거", "근거"}:
            ref_idx = i

    parsed: list[PumsamRow] = []
    for source in table[1:]:
        def pick(index: int | None) -> Any:
            if index is None or index >= len(source):
                return None
            return source[index]

        name = pick(name_idx)
        spec = pick(spec_idx)
        if name is None and spec is None:
            continue
        parsed.append(
            {
                "검색키": lookup_key(name, spec),
                "명칭": name,
                "규격": spec,
                "단위": pick(unit_idx),
                "노무명칭": pick(labor_idx),
                "품셈": pick(pumsam_idx),
                "할증%": pick(rate_idx),
                "품셈근거": pick(ref_idx),
            }
        )
    return parsed


def merge_pumsam_rows(*groups: list[PumsamRow]) -> list[PumsamRow]:
    merged: dict[str, PumsamRow] = {}
    for group in groups:
        for row in group:
            key = str(row.get("검색키") or lookup_key(row.get("명칭"), row.get("규격")))
            if not key:
                continue
            row = dict(row)
            row["검색키"] = key
            merged[key] = row
    return list(merged.values())


def load_pumsam_database(directory: Path | None = None) -> list[PumsamRow]:
    path = pumsam_db_path(directory)
    if not path.exists():
        return default_pumsam_rows()
    workbook = load_workbook(path, data_only=True)
    try:
        sheet = workbook.active
        grid = fill_merged_values(sheet)
    finally:
        workbook.close()
    table = trim_grid(grid)
    if not table:
        return default_pumsam_rows()
    header_idx = find_header_row(table)
    parsed = rows_from_grid(trim_grid(table[header_idx:]))
    if not parsed:
        return default_pumsam_rows()
    return merge_pumsam_rows(default_pumsam_rows(), parsed)


def save_pumsam_database(rows: list[PumsamRow], directory: Path | None = None) -> Path:
    folder = ensure_result_directory(directory)
    path = pumsam_db_path(folder)
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = PUMSAM_SHEET_NAME
    sheet.append(PUMSAM_HEADERS)
    for row in rows:
        sheet.append([row.get(col) for col in PUMSAM_HEADERS])
    workbook.save(path)
    workbook.close()
    return path


def import_pumsam_file(source_path: Path) -> list[PumsamRow]:
    path = Path(source_path)
    if not is_allowed_excel(path):
        raise ValueError("xlsx 또는 xlsm 파일만 읽을 수 있습니다.")
    workbook = load_workbook(path, data_only=True)
    try:
        sheet = workbook.active
        for candidate in workbook.worksheets:
            if "품셈" in str(candidate.title):
                sheet = candidate
                break
        grid = fill_merged_values(sheet)
    finally:
        workbook.close()
    table = trim_grid(grid)
    if not table:
        return []
    header_idx = find_header_row(table)
    return rows_from_grid(trim_grid(table[header_idx:]))
