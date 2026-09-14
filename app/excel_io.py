"""단가대비표 읽기(Read-Only) 및 결과 엑셀 생성.

원본 워크북은 항상 read_only로만 열고, 저장은 새 경로에만 수행한다.
"""

from __future__ import annotations

from pathlib import Path
from typing import Any

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

from app.paths import assert_safe_save, build_result_path, is_allowed_excel

SheetRows = list[list[Any]]

UNIT_PRICE_SHEET_NAME = "단가대비표"
QUANTITY_SHEET_NAME = "공량산출표"

QUANTITY_HEADERS = [
    "번호",
    "공종",
    "품명",
    "규격",
    "단위",
    "산출근거",
    "수량",
    "비고",
]

HEADER_ALIASES = {
    "공종": ("공종", "공사종별", "종별"),
    "품명": ("품명", "품목", "명칭", "자재명", "항목"),
    "규격": ("규격", "사양", "규격/사양"),
    "단위": ("단위", "단위명"),
}

HEADER_FILL = PatternFill("solid", fgColor="1F4E79")
HEADER_FONT = Font(name="맑은 고딕", bold=True, color="FFFFFF", size=11)
BODY_FONT = Font(name="맑은 고딕", size=10)
THIN = Border(
    left=Side(style="thin", color="BFBFBF"),
    right=Side(style="thin", color="BFBFBF"),
    top=Side(style="thin", color="BFBFBF"),
    bottom=Side(style="thin", color="BFBFBF"),
)
ALT_FILL = PatternFill("solid", fgColor="F2F2F2")
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)


def _cell_value(value: Any) -> Any:
    if value is None:
        return None
    if isinstance(value, str):
        stripped = value.strip()
        return stripped if stripped else None
    return value


def _row_is_empty(row: list[Any]) -> bool:
    return all(value is None or value == "" for value in row)


def _trim_sheet(rows: SheetRows) -> SheetRows:
    """앞뒤 빈 행·끝 빈 열을 제거해 원본 표를 정리한다. 값은 바꾸지 않는다."""
    cleaned: SheetRows = []
    for row in rows:
        cleaned.append([_cell_value(v) for v in row])

    while cleaned and _row_is_empty(cleaned[0]):
        cleaned.pop(0)
    while cleaned and _row_is_empty(cleaned[-1]):
        cleaned.pop()
    if not cleaned:
        return []

    max_col = 0
    for row in cleaned:
        for idx, value in enumerate(row, start=1):
            if value is not None and value != "":
                max_col = max(max_col, idx)
    return [row[:max_col] for row in cleaned]


def read_unit_price_table(source_path: Path) -> SheetRows:
    """원본 엑셀을 읽기 전용으로 열어 첫 시트를 정리된 행 목록으로 반환한다."""
    path = Path(source_path)
    if not path.exists():
        raise FileNotFoundError(f"원본 파일을 찾을 수 없습니다: {path}")
    if not is_allowed_excel(path):
        raise ValueError("xlsx 또는 xlsm 파일만 읽을 수 있습니다.")

    # data_only=True: 수식 대신 캐시된 값만 읽는다. 원본은 열기만 하고 저장하지 않는다.
    workbook = load_workbook(filename=path, read_only=True, data_only=True, keep_vba=False)
    try:
        sheet = workbook.worksheets[0]
        raw_rows: SheetRows = [list(row) for row in sheet.iter_rows(values_only=True)]
    finally:
        workbook.close()

    cleaned = _trim_sheet(raw_rows)
    if not cleaned:
        raise ValueError("단가대비표에 읽을 수 있는 데이터가 없습니다.")
    return cleaned


def _normalize_header(value: Any) -> str:
    if value is None:
        return ""
    return str(value).replace(" ", "").replace("\n", "").strip()


def _find_column_index(header_row: list[Any], field: str) -> int | None:
    aliases = HEADER_ALIASES.get(field, (field,))
    normalized_aliases = {_normalize_header(alias) for alias in aliases}
    for idx, cell in enumerate(header_row):
        if _normalize_header(cell) in normalized_aliases:
            return idx
    return None


def build_quantity_rows(unit_price_rows: SheetRows) -> SheetRows:
    """공량산출표 기본 골격. 원본에서 공종/품명/규격/단위를 옮기고 산출 칸은 비운다."""
    rows: SheetRows = [list(QUANTITY_HEADERS)]
    if len(unit_price_rows) < 2:
        return rows

    header = unit_price_rows[0]
    col_kind = _find_column_index(header, "공종")
    col_name = _find_column_index(header, "품명")
    col_spec = _find_column_index(header, "규격")
    col_unit = _find_column_index(header, "단위")

    serial = 1
    for source in unit_price_rows[1:]:
        def pick(index: int | None) -> Any:
            if index is None or index >= len(source):
                return None
            return source[index]

        name = pick(col_name)
        spec = pick(col_spec)
        unit = pick(col_unit)
        kind = pick(col_kind)
        if name is None and spec is None and kind is None:
            continue
        rows.append([serial, kind, name, spec, unit, None, None, None])
        serial += 1
    return rows


def _style_sheet(sheet: Worksheet, rows: SheetRows, header_fill: PatternFill) -> None:
    for r_idx, row in enumerate(rows, start=1):
        for c_idx, value in enumerate(row, start=1):
            cell = sheet.cell(row=r_idx, column=c_idx, value=value)
            cell.font = HEADER_FONT if r_idx == 1 else BODY_FONT
            cell.border = THIN
            cell.alignment = CENTER if r_idx == 1 or c_idx == 1 else LEFT
            if r_idx == 1:
                cell.fill = header_fill
            elif r_idx % 2 == 0:
                cell.fill = ALT_FILL

    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = sheet.dimensions
    sheet.row_dimensions[1].height = 22

    for c_idx in range(1, (len(rows[0]) if rows else 1) + 1):
        max_len = 8
        for r_idx in range(1, min(len(rows), 80) + 1):
            value = rows[r_idx - 1][c_idx - 1] if c_idx - 1 < len(rows[r_idx - 1]) else None
            if value is None:
                continue
            max_len = max(max_len, min(len(str(value)), 36))
        sheet.column_dimensions[get_column_letter(c_idx)].width = max_len + 4


def create_result_workbook(unit_price_rows: SheetRows) -> Workbook:
    """시트1 단가대비표 + 시트2 공량산출표 기본 구조를 만든다."""
    quantity_rows = build_quantity_rows(unit_price_rows)
    workbook = Workbook()

    sheet1 = workbook.active
    sheet1.title = UNIT_PRICE_SHEET_NAME
    _style_sheet(sheet1, unit_price_rows, HEADER_FILL)

    sheet2 = workbook.create_sheet(QUANTITY_SHEET_NAME)
    qty_fill = PatternFill("solid", fgColor="C45911")
    _style_sheet(sheet2, quantity_rows, qty_fill)

    sheet1.sheet_properties.tabColor = "1F4E79"
    sheet2.sheet_properties.tabColor = "C45911"
    return workbook


def save_result_workbook(
    source_path: Path,
    dest_dir: Path | None = None,
    unit_price_rows: SheetRows | None = None,
) -> Path:
    """원본을 읽기 전용으로 반영한 뒤, 타임스탬프 결과 파일만 새로 저장한다."""
    source = Path(source_path)
    rows = unit_price_rows if unit_price_rows is not None else read_unit_price_table(source)
    dest = build_result_path(dest_dir)
    assert_safe_save(source, dest)

    workbook = create_result_workbook(rows)
    try:
        workbook.save(dest)
    finally:
        workbook.close()
    return dest
