"""단가대비표 읽기(원본 미수정) 및 공량산출표 수식 엑셀 생성."""

from __future__ import annotations

from pathlib import Path
from typing import Any

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

from app.merge_parse import SheetRows, fill_merged_values, trim_grid
from app.paths import assert_safe_save, build_result_path, is_allowed_excel

UNIT_PRICE_SHEET_NAME = "단가대비표"
QUANTITY_SHEET_NAME = "공량산출표"

QTY_TITLE_ROW = 1
QTY_INFO_ROW = 2
QTY_NOTE_ROW = 3
QTY_HEADER_ROW = 4
QTY_DATA_START_ROW = 5

QUANTITY_HEADERS = [
    "번호",
    "명칭",
    "규격",
    "단위",
    "결정수량",
    "수량할증 %",
    "산출수량",
    "노무 명칭",
    "품셈",
    "노무 할증 %",
    "공량",
]

HEADER_ALIASES = {
    "공종": ("공종", "공사종별", "종별"),
    "품명": ("품명", "품목", "명칭", "자재명", "항목"),
    "규격": ("규격", "사양", "규격/사양"),
    "단위": ("단위", "단위명"),
    "수량": ("수량", "설계수량", "물량", "결정수량", "계약수량", "설계물량"),
}

HEADER_HINTS = ("품명", "명칭", "규격", "단위", "단가", "수량", "공종")

NUMBER_FORMAT = "#,##0.000"
SURCHARGE_FORMAT = "0.000"
PERCENT_FORMAT = "0.000"

HEADER_FILL_BLUE = PatternFill("solid", fgColor="BDD7EE")
HEADER_FILL_GRAY = PatternFill("solid", fgColor="D9E2F3")
TITLE_FILL = PatternFill("solid", fgColor="D6EAF8")
SUM_FILL = PatternFill("solid", fgColor="FFF2CC")
ALT_FILL = PatternFill("solid", fgColor="F7F9FB")

HEADER_FONT = Font(name="맑은 고딕", bold=True, color="1F4E79", size=11)
TITLE_FONT = Font(name="맑은 고딕", bold=True, color="1F4E79", size=16)
BODY_FONT = Font(name="맑은 고딕", size=10)
SUM_FONT = Font(name="맑은 고딕", bold=True, size=11)
THIN = Border(
    left=Side(style="thin", color="7F8C8D"),
    right=Side(style="thin", color="7F8C8D"),
    top=Side(style="thin", color="7F8C8D"),
    bottom=Side(style="thin", color="7F8C8D"),
)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)
RIGHT = Alignment(horizontal="right", vertical="center")


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


def find_quantity_column(header_row: list[Any]) -> int | None:
    idx = _find_column_index(header_row, "수량")
    if idx is not None:
        return idx
    for i, cell in enumerate(header_row):
        name = _normalize_header(cell)
        if "수량" in name and "단가" not in name and "금액" not in name:
            return i
    return None


def _find_header_row(rows: SheetRows) -> int:
    for i, row in enumerate(rows):
        normalized = [_normalize_header(c) for c in row]
        hits = sum(1 for hint in HEADER_HINTS if any(hint in cell for cell in normalized))
        if hits >= 2:
            return i
    return 0


def read_unit_price_table(source_path: Path) -> SheetRows:
    """원본 엑셀을 이진 읽기로만 열어 병합을 채운 뒤 정리된 표로 반환한다."""
    path = Path(source_path)
    if not path.exists():
        raise FileNotFoundError(f"원본 파일을 찾을 수 없습니다: {path}")
    if not is_allowed_excel(path):
        raise ValueError("xlsx 또는 xlsm 파일만 읽을 수 있습니다.")

    with path.open("rb") as handle:
        workbook = load_workbook(filename=handle, data_only=True, keep_vba=False)
        try:
            if not workbook.worksheets:
                raise ValueError("단가대비표에 시트가 없습니다.")
            sheet = workbook.worksheets[0]
            grid = fill_merged_values(sheet)
        finally:
            workbook.close()

    cleaned = trim_grid(grid)
    if not cleaned:
        raise ValueError("단가대비표에 읽을 수 있는 데이터가 없습니다.")
    header_idx = _find_header_row(cleaned)
    table = trim_grid(cleaned[header_idx:])
    if not table:
        raise ValueError("단가대비표 헤더를 찾지 못했습니다.")
    return table


def decided_qty_formula(source_col_letter: str, source_row: int) -> str:
    return f"='{UNIT_PRICE_SHEET_NAME}'!{source_col_letter}{source_row}"


def calc_qty_formula(row: int) -> str:
    return f"=E{row}*(1+F{row})"


def gongryang_formula(row: int) -> str:
    return f'=IF(G{row}*I{row}=0,"",G{row}*I{row}*(J{row}/100))'


def gongryang_sum_formula(end_row: int) -> str:
    return f"=SUM(K{QTY_DATA_START_ROW}:K{end_row})"


def _display_width(text: str) -> int:
    width = 0
    for char in text:
        width += 2 if ord(char) > 127 else 1
    return width


def _autosize_columns(sheet: Worksheet, min_width: float = 10, max_width: float = 28) -> None:
    max_col = sheet.max_column or 1
    max_row = sheet.max_row or 1
    for col_idx in range(1, max_col + 1):
        longest = min_width
        for row_idx in range(1, max_row + 1):
            value = sheet.cell(row=row_idx, column=col_idx).value
            if value is None:
                continue
            text = str(value)
            if text.startswith("="):
                continue
            longest = max(longest, min(_display_width(text) + 3, max_width))
        sheet.column_dimensions[get_column_letter(col_idx)].width = longest


def _apply_thin_range(sheet: Worksheet, min_row: int, max_row: int, min_col: int, max_col: int) -> None:
    for r in range(min_row, max_row + 1):
        for c in range(min_col, max_col + 1):
            sheet.cell(row=r, column=c).border = THIN


def _write_unit_price_sheet(sheet: Worksheet, rows: SheetRows) -> None:
    header_fill = HEADER_FILL_BLUE
    for r_idx, row in enumerate(rows, start=1):
        for c_idx, value in enumerate(row, start=1):
            cell = sheet.cell(row=r_idx, column=c_idx, value=value)
            cell.border = THIN
            cell.font = HEADER_FONT if r_idx == 1 else BODY_FONT
            cell.alignment = CENTER if r_idx == 1 else LEFT
            if r_idx == 1:
                cell.fill = header_fill
            elif r_idx % 2 == 0:
                cell.fill = ALT_FILL
            if r_idx > 1 and isinstance(value, (int, float)):
                cell.number_format = NUMBER_FORMAT
                cell.alignment = RIGHT

    if rows:
        sheet.freeze_panes = "A2"
        sheet.auto_filter.ref = sheet.dimensions
        sheet.row_dimensions[1].height = 22
    sheet.sheet_properties.tabColor = "5B9BD5"
    _autosize_columns(sheet)


def _write_quantity_sheet(
    sheet: Worksheet,
    unit_price_rows: SheetRows,
    source_name: str = "",
) -> None:
    header = unit_price_rows[0] if unit_price_rows else []
    name_idx = _find_column_index(header, "품명")
    spec_idx = _find_column_index(header, "규격")
    unit_idx = _find_column_index(header, "단위")
    qty_idx = find_quantity_column(header)
    qty_letter = get_column_letter(qty_idx + 1) if qty_idx is not None else None

    last_col = len(QUANTITY_HEADERS)
    sheet.merge_cells(start_row=QTY_TITLE_ROW, start_column=1, end_row=QTY_TITLE_ROW, end_column=last_col)
    title = sheet.cell(row=QTY_TITLE_ROW, column=1, value="공량산출표")
    title.font = TITLE_FONT
    title.fill = TITLE_FILL
    title.alignment = CENTER

    sheet.merge_cells(start_row=QTY_INFO_ROW, start_column=1, end_row=QTY_INFO_ROW, end_column=last_col)
    info = sheet.cell(
        row=QTY_INFO_ROW,
        column=1,
        value=f"원본: {source_name}    |    결정수량(E)은 단가대비표 수량 셀 수식 참조",
    )
    info.font = BODY_FONT
    info.alignment = LEFT

    sheet.merge_cells(start_row=QTY_NOTE_ROW, start_column=1, end_row=QTY_NOTE_ROW, end_column=last_col)
    note = sheet.cell(
        row=QTY_NOTE_ROW,
        column=1,
        value="F열 수량할증은 0 또는 0.05(5%)처럼 소수로 입력 · J열 노무 할증은 기본 100(%)",
    )
    note.font = BODY_FONT
    note.alignment = LEFT

    for c_idx, header_name in enumerate(QUANTITY_HEADERS, start=1):
        cell = sheet.cell(row=QTY_HEADER_ROW, column=c_idx, value=header_name)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL_GRAY
        cell.alignment = CENTER
        cell.border = THIN
    sheet.row_dimensions[QTY_HEADER_ROW].height = 22

    def pick(row: list[Any], index: int | None) -> Any:
        if index is None or index >= len(row):
            return None
        return row[index]

    dest_row = QTY_DATA_START_ROW
    serial = 1
    data_rows = unit_price_rows[1:] if len(unit_price_rows) > 1 else []
    for source_row_number, source in enumerate(data_rows, start=2):
        name = pick(source, name_idx)
        spec = pick(source, spec_idx)
        unit = pick(source, unit_idx)
        if name is None and spec is None:
            continue

        values = {
            1: serial,
            2: name,
            3: spec,
            4: unit,
            5: decided_qty_formula(qty_letter, source_row_number) if qty_letter else None,
            6: 0,
            7: calc_qty_formula(dest_row),
            8: None,
            9: None,
            10: 100,
            11: gongryang_formula(dest_row),
        }
        for col, value in values.items():
            cell = sheet.cell(row=dest_row, column=col, value=value)
            cell.font = BODY_FONT
            cell.border = THIN
            cell.alignment = CENTER if col in (1, 4) else LEFT
            if dest_row % 2 == 0:
                cell.fill = ALT_FILL
            if col in (5, 7, 9, 11):
                cell.number_format = NUMBER_FORMAT
                cell.alignment = RIGHT
            elif col == 6:
                cell.number_format = SURCHARGE_FORMAT
                cell.alignment = RIGHT
            elif col == 10:
                cell.number_format = PERCENT_FORMAT
                cell.alignment = RIGHT
        serial += 1
        dest_row += 1

    if dest_row == QTY_DATA_START_ROW:
        # 데이터 없을 때도 SUM 범위가 유효하도록 빈 데이터 행 1개를 둔다.
        for col in range(1, last_col + 1):
            cell = sheet.cell(row=QTY_DATA_START_ROW, column=col, value=None)
            cell.border = THIN
            if col in (5, 7, 9, 11):
                cell.number_format = NUMBER_FORMAT
        last_data_row = QTY_DATA_START_ROW
        dest_row = QTY_DATA_START_ROW + 1
    else:
        last_data_row = dest_row - 1

    sum_row = last_data_row + 1
    for col in range(1, last_col + 1):
        cell = sheet.cell(row=sum_row, column=col, value=None)
        cell.border = THIN
        cell.fill = SUM_FILL
        cell.font = SUM_FONT
    sheet.cell(row=sum_row, column=1, value="합계")
    sum_cell = sheet.cell(row=sum_row, column=11, value=gongryang_sum_formula(last_data_row))
    sum_cell.number_format = NUMBER_FORMAT
    sum_cell.alignment = RIGHT
    sum_cell.font = SUM_FONT
    sum_cell.fill = SUM_FILL

    _apply_thin_range(sheet, QTY_TITLE_ROW, QTY_HEADER_ROW, 1, last_col)
    sheet.freeze_panes = "A5"
    sheet.auto_filter.ref = f"A{QTY_HEADER_ROW}:{get_column_letter(last_col)}{last_data_row}"
    sheet.sheet_properties.tabColor = "C45911"
    _autosize_columns(sheet, min_width=12, max_width=22)
    sheet.column_dimensions["B"].width = 22
    sheet.column_dimensions["C"].width = 18
    sheet.column_dimensions["H"].width = 16
    sheet.column_dimensions["K"].width = 14


def create_result_workbook(unit_price_rows: SheetRows, source_name: str = "") -> Workbook:
    """시트1 단가대비표 + 시트2 공량산출표(수식 연동)를 만든다."""
    workbook = Workbook()
    sheet1 = workbook.active
    sheet1.title = UNIT_PRICE_SHEET_NAME
    _write_unit_price_sheet(sheet1, unit_price_rows)

    sheet2 = workbook.create_sheet(QUANTITY_SHEET_NAME)
    _write_quantity_sheet(sheet2, unit_price_rows, source_name=source_name)
    return workbook


def save_result_workbook(
    source_path: Path,
    dest_dir: Path | None = None,
    unit_price_rows: SheetRows | None = None,
) -> Path:
    """원본을 읽기만 한 뒤, 타임스탬프 결과 파일만 새로 저장한다."""
    source = Path(source_path)
    rows = unit_price_rows if unit_price_rows is not None else read_unit_price_table(source)
    dest = build_result_path(dest_dir)
    assert_safe_save(source, dest)

    workbook = create_result_workbook(rows, source_name=source.name)
    try:
        workbook.save(dest)
    finally:
        workbook.close()
    return dest
