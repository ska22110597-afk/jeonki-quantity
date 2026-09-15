"""내역서 읽기(원본 미수정) 및 3시트 결과 엑셀 생성."""

from __future__ import annotations

from pathlib import Path
from typing import Any

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

from app.estimate_parse import (
    SheetRows,
    find_column_index,
    find_quantity_column,
    is_section_row,
    read_named_sheet_rows,
    read_workbook_first_sheet,
)
from app.paths import assert_safe_save, build_result_path
from app.pumsam import (
    PUMSAM_HEADERS,
    PUMSAM_SHEET_NAME,
    PumsamRow,
    import_pumsam_file,
    load_pumsam_database,
    merge_pumsam_rows,
    rows_from_grid,
    save_pumsam_database,
)

ESTIMATE_SHEET_NAME = "내역서"
QUANTITY_SHEET_NAME = "공량산출서"
UNIT_PRICE_SHEET_NAME = ESTIMATE_SHEET_NAME  # 이전 테스트 이름 호환

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
    "품셈근거",
]

NUMBER_FORMAT = "#,##0.000"
SURCHARGE_FORMAT = "0.00"
PERCENT_FORMAT = "0"

HEADER_FILL_BLUE = PatternFill("solid", fgColor="BDD7EE")
HEADER_FILL_GRAY = PatternFill("solid", fgColor="D9E2F3")
HEADER_FILL_GREEN = PatternFill("solid", fgColor="C6EFCE")
TITLE_FILL = PatternFill("solid", fgColor="D6EAF8")
SUM_FILL = PatternFill("solid", fgColor="FFF2CC")
ALT_FILL = PatternFill("solid", fgColor="F7F9FB")
SECTION_FILL = PatternFill("solid", fgColor="9BC2E6")

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

KEY_FORMULA = '=SUBSTITUTE(SUBSTITUTE(B{row}&C{row}," ",""),"　","")'
VLOOKUP = (
    '=IFERROR(VLOOKUP(SUBSTITUTE(SUBSTITUTE($B{row}&$C{row}," ",""),"　",""),'
    f"'{PUMSAM_SHEET_NAME}'!$A:$H,{{col}},FALSE),\"\")"
)


def calc_qty_formula(row: int) -> str:
    return f"=E{row}*(1+F{row}/100)"


def gongryang_formula(row: int) -> str:
    return f'=IF(OR(G{row}="",I{row}="",G{row}*I{row}=0),"",G{row}*I{row}*(IF(J{row}="",100,J{row})/100))'


def gongryang_sum_formula(end_row: int) -> str:
    return f"=SUM(K{QTY_DATA_START_ROW}:K{end_row})"


def decided_qty_formula(source_col_letter: str, source_row: int) -> str:
    return f"='{ESTIMATE_SHEET_NAME}'!{source_col_letter}{source_row}"


def labor_formula(row: int) -> str:
    return VLOOKUP.format(row=row, col=5)


def pumsam_formula(row: int) -> str:
    return VLOOKUP.format(row=row, col=6)


def labor_rate_formula(row: int) -> str:
    return VLOOKUP.format(row=row, col=7)


def ref_formula(row: int) -> str:
    return VLOOKUP.format(row=row, col=8)


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


def _write_estimate_sheet(sheet: Worksheet, rows: SheetRows) -> None:
    header = rows[0] if rows else []
    name_idx = find_column_index(header, "명칭")
    spec_idx = find_column_index(header, "규격")
    unit_idx = find_column_index(header, "단위")

    for r_idx, row in enumerate(rows, start=1):
        section = False
        if r_idx > 1:
            name = row[name_idx] if name_idx is not None and name_idx < len(row) else None
            spec = row[spec_idx] if spec_idx is not None and spec_idx < len(row) else None
            unit = row[unit_idx] if unit_idx is not None and unit_idx < len(row) else None
            section = is_section_row(name, spec, unit)
        for c_idx, value in enumerate(row, start=1):
            cell = sheet.cell(row=r_idx, column=c_idx, value=value)
            cell.border = THIN
            cell.font = HEADER_FONT if r_idx == 1 else BODY_FONT
            cell.alignment = CENTER if r_idx == 1 else LEFT
            if r_idx == 1:
                cell.fill = HEADER_FILL_BLUE
            elif section:
                cell.fill = SECTION_FILL
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


def _write_pumsam_sheet(sheet: Worksheet, rows: list[PumsamRow]) -> None:
    for c_idx, header_name in enumerate(PUMSAM_HEADERS, start=1):
        cell = sheet.cell(row=1, column=c_idx, value=header_name)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL_GREEN
        cell.alignment = CENTER
        cell.border = THIN
    for offset, row in enumerate(rows):
        excel_row = offset + 2
        values = [
            KEY_FORMULA.format(row=excel_row),
            row.get("명칭"),
            row.get("규격"),
            row.get("단위"),
            row.get("노무명칭"),
            row.get("품셈"),
            row.get("할증%"),
            row.get("품셈근거"),
        ]
        for c_idx, value in enumerate(values, start=1):
            cell = sheet.cell(row=excel_row, column=c_idx, value=value)
            cell.font = BODY_FONT
            cell.border = THIN
            cell.alignment = CENTER if c_idx in (3, 4, 7) else LEFT
            if excel_row % 2 == 0:
                cell.fill = ALT_FILL
            if c_idx == 6:
                cell.number_format = NUMBER_FORMAT
                cell.alignment = RIGHT
            elif c_idx == 7:
                cell.number_format = PERCENT_FORMAT
                cell.alignment = RIGHT
    sheet.freeze_panes = "A2"
    if rows:
        sheet.auto_filter.ref = f"A1:H{len(rows) + 1}"
    sheet.sheet_properties.tabColor = "548235"
    _autosize_columns(sheet, min_width=12, max_width=24)
    sheet.column_dimensions["B"].width = 24
    note = sheet.cell(
        row=len(rows) + 3,
        column=1,
        value="행을 추가해 품셈을 계속 쌓으면 됩니다. 검색키(A열) 수식은 위 행을 복사하세요. 명칭+규격이 내역서와 같아야 공량산출서에 붙습니다.",
    )
    note.font = BODY_FONT


def _write_quantity_sheet(
    sheet: Worksheet,
    estimate_rows: SheetRows,
    source_name: str = "",
) -> None:
    header = estimate_rows[0] if estimate_rows else []
    name_idx = find_column_index(header, "명칭")
    spec_idx = find_column_index(header, "규격")
    unit_idx = find_column_index(header, "단위")
    qty_idx = find_quantity_column(header)
    qty_letter = get_column_letter(qty_idx + 1) if qty_idx is not None else None

    last_col = len(QUANTITY_HEADERS)
    sheet.merge_cells(start_row=QTY_TITLE_ROW, start_column=1, end_row=QTY_TITLE_ROW, end_column=last_col)
    title = sheet.cell(row=QTY_TITLE_ROW, column=1, value="공량산출서")
    title.font = TITLE_FONT
    title.fill = TITLE_FILL
    title.alignment = CENTER

    sheet.merge_cells(start_row=QTY_INFO_ROW, start_column=1, end_row=QTY_INFO_ROW, end_column=last_col)
    info = sheet.cell(
        row=QTY_INFO_ROW,
        column=1,
        value=f"원본: {source_name}  |  E열=내역서 수량  |  H/I/J/L열=품셈표 VLOOKUP  |  공량=산출수량×품셈×(노무할증/100)",
    )
    info.font = BODY_FONT
    info.alignment = LEFT

    sheet.merge_cells(start_row=QTY_NOTE_ROW, start_column=1, end_row=QTY_NOTE_ROW, end_column=last_col)
    note = sheet.cell(
        row=QTY_NOTE_ROW,
        column=1,
        value="F열 수량할증은 % 숫자입니다. 예: 0 또는 5 (5%). 산출수량 = 결정수량×(1+수량할증/100)",
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
    data_rows = estimate_rows[1:] if len(estimate_rows) > 1 else []
    for source_row_number, source in enumerate(data_rows, start=2):
        name = pick(source, name_idx)
        spec = pick(source, spec_idx)
        unit = pick(source, unit_idx)
        if is_section_row(name, spec, unit):
            continue
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
            8: labor_formula(dest_row),
            9: pumsam_formula(dest_row),
            10: labor_rate_formula(dest_row),
            11: gongryang_formula(dest_row),
            12: ref_formula(dest_row),
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
        for col in range(1, last_col + 1):
            cell = sheet.cell(row=QTY_DATA_START_ROW, column=col, value=None)
            cell.border = THIN
            if col in (5, 7, 9, 11):
                cell.number_format = NUMBER_FORMAT
        last_data_row = QTY_DATA_START_ROW
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
    sheet.column_dimensions["B"].width = 24
    sheet.column_dimensions["C"].width = 16
    sheet.column_dimensions["H"].width = 14
    sheet.column_dimensions["K"].width = 14


def collect_pumsam_rows(
    source_path: Path | None,
    dest_dir: Path | None,
    extra_pumsam_path: Path | None = None,
) -> list[PumsamRow]:
    groups = [load_pumsam_database(dest_dir)]
    if source_path is not None:
        embedded = read_named_sheet_rows(source_path, "품셈")
        if embedded:
            groups.append(rows_from_grid(embedded))
    if extra_pumsam_path is not None:
        groups.append(import_pumsam_file(extra_pumsam_path))
    merged = merge_pumsam_rows(*groups)
    save_pumsam_database(merged, dest_dir)
    return merged


def create_result_workbook(
    estimate_rows: SheetRows,
    pumsam_rows: list[PumsamRow] | None = None,
    source_name: str = "",
) -> Workbook:
    workbook = Workbook()
    sheet1 = workbook.active
    sheet1.title = ESTIMATE_SHEET_NAME
    _write_estimate_sheet(sheet1, estimate_rows)

    sheet2 = workbook.create_sheet(PUMSAM_SHEET_NAME)
    _write_pumsam_sheet(sheet2, pumsam_rows or [])

    sheet3 = workbook.create_sheet(QUANTITY_SHEET_NAME)
    _write_quantity_sheet(sheet3, estimate_rows, source_name=source_name)
    return workbook


def save_result_workbook(
    source_path: Path,
    dest_dir: Path | None = None,
    estimate_rows: SheetRows | None = None,
    extra_pumsam_path: Path | None = None,
) -> Path:
    """원본 내역서는 읽기만 하고, 3시트 결과 파일만 새로 저장한다."""
    source = Path(source_path)
    rows = estimate_rows if estimate_rows is not None else read_workbook_first_sheet(source)
    pumsam_rows = collect_pumsam_rows(source, dest_dir, extra_pumsam_path)
    dest = build_result_path(dest_dir)
    assert_safe_save(source, dest)

    workbook = create_result_workbook(rows, pumsam_rows=pumsam_rows, source_name=source.name)
    try:
        workbook.save(dest)
    finally:
        workbook.close()
    return dest


# 이전 이름 호환
read_unit_price_table = read_workbook_first_sheet
