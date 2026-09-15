"""품셈표 데이터베이스. C드라이브 로컬 파일에 쌓고, 결과 엑셀에도 시트로 넣는다."""

from __future__ import annotations

from pathlib import Path
from typing import Any

from openpyxl import Workbook, load_workbook

from app.estimate_parse import (
    find_column_index,
    find_header_row,
    header_row_span,
    lookup_key,
    normalize_header,
)
from app.merge_parse import SheetRows, fill_merged_values, trim_grid
from app.paths import (
    bundled_data_dir,
    ensure_result_directory,
    get_result_directory,
    is_allowed_excel,
    user_database_dir,
)

PUMSAM_SHEET_NAME = "품셈표"
PUMSAM_DB_FILENAME = "표준품셈.xlsx"
LEGACY_PUMSAM_DB_FILENAME = "품셈표_데이터베이스.xlsx"
PUMSAM_HEADERS = ["검색키", "명칭", "규격", "단위", "노무명칭", "품셈", "할증%", "품셈근거"]

PumsamRow = dict[str, Any]


def pumsam_db_path(directory: Path | None = None) -> Path:
    return user_database_dir(directory) / PUMSAM_DB_FILENAME


def bundled_pumsam_path() -> Path:
    return bundled_data_dir() / PUMSAM_DB_FILENAME


def _load_rows_from_path(path: Path) -> list[PumsamRow]:
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


def ensure_pumsam_database(directory: Path | None = None) -> Path:
    dest = pumsam_db_path(directory)
    if dest.exists():
        return dest
    dest.parent.mkdir(parents=True, exist_ok=True)
    bundled = bundled_pumsam_path()
    if bundled.exists():
        dest.write_bytes(bundled.read_bytes())
        return dest
    legacy = get_result_directory() / LEGACY_PUMSAM_DB_FILENAME
    if directory is not None:
        legacy_alt = Path(directory) / LEGACY_PUMSAM_DB_FILENAME
        if legacy_alt.exists():
            dest.write_bytes(legacy_alt.read_bytes())
            return dest
    if legacy.exists():
        dest.write_bytes(legacy.read_bytes())
        return dest
    save_pumsam_database(default_pumsam_rows(), directory)
    return dest


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
    rows.append(
        {
            "검색키": lookup_key("경질비닐전선관_노출", _hi_spec(104)),
            "명칭": "경질비닐전선관_노출",
            "규격": _hi_spec(104),
            "단위": "M",
            "노무명칭": "보통인부",
            "품셈": 0.001,
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


_HEADER_LIKE_NAMES = {"명칭", "품명", "품목", "검색키", "품목명"}
_GROUP_TITLES = {"공량산출", "품목", "비고"}


def _combined_header(table: SheetRows) -> tuple[list[Any], int]:
    """1~2단 헤더를 한 줄로 합치고, 데이터 시작 인덱스를 반환한다."""
    if not table:
        return [], 0
    header_idx = find_header_row(table)
    span = header_row_span(table, header_idx)
    row1 = table[header_idx]
    row2 = table[header_idx + 1] if span == 2 and header_idx + 1 < len(table) else []
    width = max(len(row1), len(row2))
    combined: list[Any] = []
    for col in range(width):
        top = row1[col] if col < len(row1) else None
        bottom = row2[col] if col < len(row2) else None
        bottom_token = normalize_header(bottom)
        top_token = normalize_header(top)
        if bottom_token and bottom_token not in _GROUP_TITLES:
            combined.append(bottom)
        elif top_token and top_token not in _GROUP_TITLES:
            combined.append(top)
        else:
            combined.append(bottom if bottom is not None else top)
    return combined, header_idx + span


def rows_from_grid(table: SheetRows) -> list[PumsamRow]:
    if not table:
        return []
    header, data_start = _combined_header(table)
    spec_idx = find_column_index(header, "규격")
    unit_idx = find_column_index(header, "단위")
    name_idx = None
    labor_idx = None
    pumsam_idx = None
    rate_idx = None
    ref_idx = None
    seen_name = False
    for i, cell in enumerate(header):
        token = normalize_header(cell)
        if token in {"명칭", "품명"}:
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
    if name_idx is None:
        name_idx = find_column_index(header, "명칭")

    parsed: list[PumsamRow] = []
    prev_name: Any = None
    prev_spec: Any = None
    prev_unit: Any = None
    for source in table[data_start:]:
        def pick(index: int | None) -> Any:
            if index is None or index >= len(source):
                return None
            return source[index]

        name = pick(name_idx)
        spec = pick(spec_idx)
        labor = pick(labor_idx)
        pumsam_value = pick(pumsam_idx)
        if name is None and spec is None:
            if (labor or pumsam_value is not None) and prev_name is not None:
                name, spec = prev_name, prev_spec
                unit_value = prev_unit
            else:
                continue
        else:
            unit_value = pick(unit_idx)
            prev_name, prev_spec, prev_unit = name, spec, unit_value
        if normalize_header(name) in _HEADER_LIKE_NAMES:
            continue
        if normalize_header(pumsam_value) == "품셈":
            continue
        parsed.append(
            {
                "검색키": lookup_key(name, spec),
                "명칭": name,
                "규격": spec,
                "단위": unit_value if name is not None else prev_unit,
                "노무명칭": labor,
                "품셈": pumsam_value,
                "할증%": pick(rate_idx),
                "품셈근거": pick(ref_idx),
            }
        )
    return parsed


def pumsam_identity(row: PumsamRow) -> str:
    labor = str(row.get("노무명칭") or "").strip()
    return f"{lookup_key(row.get('명칭'), row.get('규격'))}|{labor}"


def merge_pumsam_rows(*groups: list[PumsamRow]) -> list[PumsamRow]:
    merged: dict[str, PumsamRow] = {}
    for group in groups:
        for row in group:
            key = pumsam_identity(row)
            if not key or key == "|":
                continue
            row = dict(row)
            row["검색키"] = lookup_key(row.get("명칭"), row.get("규격"))
            merged[key] = row
    return list(merged.values())


def load_pumsam_database(directory: Path | None = None) -> list[PumsamRow]:
    path = ensure_pumsam_database(directory)
    parsed = _load_rows_from_path(path) if path.exists() else []
    if not parsed:
        return default_pumsam_rows()
    return merge_pumsam_rows(default_pumsam_rows(), parsed)


def save_pumsam_database(rows: list[PumsamRow], directory: Path | None = None) -> Path:
    path = pumsam_db_path(directory)
    path.parent.mkdir(parents=True, exist_ok=True)
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
