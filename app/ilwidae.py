"""일위대가 호표 작성. 표준품셈 인부와 전선관 부가항목을 넣는다."""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any

from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

from app.discipline import default_labor_name
from app.estimate_parse import normalize_header
from app.excel_io import (
    BODY_FONT,
    CENTER,
    FORM_ROW_HEIGHT,
    HEADER_FONT,
    ILWIDAE_SHEET_NAME,
    LEFT,
    MONEY_FORMAT,
    RIGHT,
    write_title_banner,
    _apply_sheet_look,
    _set_cell,
)
from app.items import LineItem
from app.pumsam import (
    PumsamRow,
    format_pumsam_ref,
    match_pumsam as match_pumsam_rows,
    pumsam_qty_value,
    pumsam_rate_value,
)
from app.wages import WageRow

CONDUIT_FITTING_RATE = 0.15
SUNDRY_RATE = 0.02
TOOL_RATE = 0.03
PIPE_WASTE_QTY = 1.1


@dataclass
class IlwidaeBlock:
    item: LineItem
    ho_no: int
    title_row: int
    material_row: int
    sum_row: int
    labor_rows: list[int] = field(default_factory=list)
    pumsam_ref: str = ""


def is_conduit_name(name: Any) -> bool:
    return "전선관" in str(name or "")


def is_cable_name(name: Any) -> bool:
    text = str(name or "")
    if "전선관" in text:
        return False
    return "케이블" in text or "전선" in text


def is_meter_unit(unit: Any) -> bool:
    token = normalize_header(unit).lower()
    return token in {"m", "ｍ", "미터", "metre", "meter"}


def material_qty(item: LineItem) -> float:
    if is_meter_unit(item.unit) or is_conduit_name(item.name) or is_cable_name(item.name):
        return PIPE_WASTE_QTY
    return 1


def match_pumsam(item: LineItem, rows: list[PumsamRow]) -> list[PumsamRow]:
    return match_pumsam_rows(item.name, item.spec, rows)


def labor_qty_formula(row: PumsamRow) -> str:
    qty = pumsam_qty_value(row)
    rate = pumsam_rate_value(row)
    return f"={qty}*{rate}"


def _write_ilwidae_header(sheet: Worksheet) -> None:
    write_title_banner(sheet, "일 위 대 가", 13)
    labels = {
        1: "품      명",
        2: "규      격",
        3: "단위",
        4: "수량",
        5: "재  료  비",
        7: "노  무  비",
        9: "경      비",
        11: "합      계",
        13: "비  고",
    }
    for col in range(1, 14):
        _set_cell(sheet, 3, col, labels.get(col), font=HEADER_FONT, align=CENTER)
        _set_cell(sheet, 4, col, None, font=HEADER_FONT, align=CENTER)
    for col, text in (
        (5, "단  가"),
        (6, "금  액"),
        (7, "단  가"),
        (8, "금  액"),
        (9, "단  가"),
        (10, "금  액"),
        (11, "단  가"),
        (12, "금  액"),
    ):
        _set_cell(sheet, 4, col, text, font=HEADER_FONT, align=CENTER)
    sheet.merge_cells("A3:A4")
    sheet.merge_cells("B3:B4")
    sheet.merge_cells("C3:C4")
    sheet.merge_cells("D3:D4")
    sheet.merge_cells("E3:F3")
    sheet.merge_cells("G3:H3")
    sheet.merge_cells("I3:J3")
    sheet.merge_cells("K3:L3")
    sheet.merge_cells("M3:M4")


def _money(sheet: Worksheet, row: int, col: int, value: Any, align=RIGHT) -> None:
    _set_cell(sheet, row, col, value, font=BODY_FONT, align=align, number_format=MONEY_FORMAT)


def _price_ref(item: LineItem, compare_sheet: str) -> str:
    col_index = (item.price_col if item.price_col is not None else 4) + 1
    return f"='{compare_sheet}'!{get_column_letter(col_index)}{item.excel_row}"


def write_ilwidae_sheet(
    sheet: Worksheet,
    items: list[LineItem],
    pumsam_rows: list[PumsamRow],
    wage_rows: list[WageRow],
    compare_sheet: str = "단가대비표",
    discipline: str | None = None,
) -> list[IlwidaeBlock]:
    sheet.title = ILWIDAE_SHEET_NAME
    _write_ilwidae_header(sheet)
    cursor = 5
    blocks: list[IlwidaeBlock] = []
    ho_no = 0
    fallback_labor = default_labor_name(discipline)
    for item in items:
        if item.section or (item.spec is None and item.unit is None):
            continue
        ho_no += 1
        labors = match_pumsam_rows(item.name, item.spec, pumsam_rows)
        if not labors:
            labors = [
                {
                    "명칭": item.name,
                    "규격": item.spec,
                    "노무명칭": fallback_labor,
                    "품셈": 0,
                    "할증%": 100,
                    "품셈근거": "",
                }
            ]
        ref = format_pumsam_ref(labors[0].get("품셈근거") or "")
        title = f"{item.name} {item.spec or ''}  ( 호표 {ho_no} )".strip()
        title_row = cursor
        _set_cell(sheet, cursor, 1, title, font=BODY_FONT, align=LEFT)
        for col in range(2, 13):
            _set_cell(sheet, cursor, col, None)
        _set_cell(sheet, cursor, 13, ref or None, font=BODY_FONT, align=CENTER)
        cursor += 1

        material_row = cursor
        qty = material_qty(item)
        _set_cell(sheet, cursor, 1, item.name, font=BODY_FONT)
        _set_cell(sheet, cursor, 2, item.spec, font=BODY_FONT)
        _set_cell(sheet, cursor, 3, item.unit or "개", font=BODY_FONT, align=CENTER)
        _money(sheet, cursor, 4, qty)
        _money(sheet, cursor, 5, _price_ref(item, compare_sheet))
        _money(sheet, cursor, 6, f"=D{cursor}*E{cursor}")
        _money(sheet, cursor, 7, 0)
        _money(sheet, cursor, 8, 0)
        _money(sheet, cursor, 9, 0)
        _money(sheet, cursor, 10, 0)
        _money(sheet, cursor, 11, f"=TRUNC(E{cursor}+G{cursor}+I{cursor},2)")
        _money(sheet, cursor, 12, f"=TRUNC(F{cursor}+H{cursor}+J{cursor},1)")
        _set_cell(sheet, cursor, 13, ref or None, font=BODY_FONT, align=CENTER)
        cursor += 1

        fitting_row = None
        if is_conduit_name(item.name):
            fitting_row = cursor
            _set_cell(sheet, cursor, 1, "전선관부속품비", font=BODY_FONT)
            _set_cell(sheet, cursor, 2, "전선관의 15%", font=BODY_FONT)
            _set_cell(sheet, cursor, 3, "식", font=BODY_FONT, align=CENTER)
            _money(sheet, cursor, 4, 1)
            _money(sheet, cursor, 5, 0)
            _money(sheet, cursor, 6, f"=F{material_row}*{CONDUIT_FITTING_RATE}")
            _money(sheet, cursor, 7, 0)
            _money(sheet, cursor, 8, 0)
            _money(sheet, cursor, 9, 0)
            _money(sheet, cursor, 10, 0)
            _money(sheet, cursor, 11, f"=TRUNC(E{cursor}+G{cursor}+I{cursor},2)")
            _money(sheet, cursor, 12, f"=TRUNC(F{cursor}+H{cursor}+J{cursor},1)")
            _set_cell(sheet, cursor, 13, None)
            cursor += 1

        sundry_row = None
        if is_conduit_name(item.name) or is_cable_name(item.name):
            sundry_row = cursor
            material_sum = f"F{material_row}"
            if fitting_row is not None:
                material_sum = f"(F{material_row}+F{fitting_row})"
            _set_cell(sheet, cursor, 1, "잡재료비", font=BODY_FONT)
            _set_cell(sheet, cursor, 2, "배관배선의 2%", font=BODY_FONT)
            _set_cell(sheet, cursor, 3, "식", font=BODY_FONT, align=CENTER)
            _money(sheet, cursor, 4, 1)
            _money(sheet, cursor, 5, 0)
            _money(sheet, cursor, 6, f"={material_sum}*{SUNDRY_RATE}")
            _money(sheet, cursor, 7, 0)
            _money(sheet, cursor, 8, 0)
            _money(sheet, cursor, 9, 0)
            _money(sheet, cursor, 10, 0)
            _money(sheet, cursor, 11, f"=TRUNC(E{cursor}+G{cursor}+I{cursor},2)")
            _money(sheet, cursor, 12, f"=TRUNC(F{cursor}+H{cursor}+J{cursor},1)")
            _set_cell(sheet, cursor, 13, None)
            cursor += 1
            _ = sundry_row

        labor_rows: list[int] = []
        for labor in labors:
            job = labor.get("노무명칭") or "내선전공"
            _set_cell(sheet, cursor, 1, job, font=BODY_FONT)
            _set_cell(sheet, cursor, 2, "일반공사 직종", font=BODY_FONT)
            _set_cell(sheet, cursor, 3, "인", font=BODY_FONT, align=CENTER)
            _set_cell(sheet, cursor, 4, labor_qty_formula(labor), font=BODY_FONT, align=RIGHT, number_format="0.000")
            _money(sheet, cursor, 5, 0)
            _money(sheet, cursor, 6, 0)
            _money(sheet, cursor, 7, f'=IFERROR(VLOOKUP("{job}",노임단가!A:B,2,FALSE),0)')
            _money(sheet, cursor, 8, f"=TRUNC(G{cursor}*D{cursor},1)")
            _money(sheet, cursor, 9, 0)
            _money(sheet, cursor, 10, 0)
            _money(sheet, cursor, 11, f"=TRUNC(E{cursor}+G{cursor}+I{cursor},2)")
            _money(sheet, cursor, 12, f"=TRUNC(F{cursor}+H{cursor}+J{cursor},1)")
            _set_cell(sheet, cursor, 13, None)
            labor_rows.append(cursor)
            cursor += 1

        first_labor = labor_rows[0]
        last_labor = labor_rows[-1]
        _set_cell(sheet, cursor, 1, "공구손료", font=BODY_FONT)
        _set_cell(sheet, cursor, 2, "인력품의 3%", font=BODY_FONT)
        _set_cell(sheet, cursor, 3, "식", font=BODY_FONT, align=CENTER)
        _money(sheet, cursor, 4, 1)
        _money(sheet, cursor, 5, 0)
        _money(sheet, cursor, 6, 0)
        _money(sheet, cursor, 7, 0)
        _money(sheet, cursor, 8, f"=SUM(H{first_labor}:H{last_labor})*{TOOL_RATE}")
        _money(sheet, cursor, 9, 0)
        _money(sheet, cursor, 10, 0)
        _money(sheet, cursor, 11, f"=TRUNC(E{cursor}+G{cursor}+I{cursor},2)")
        _money(sheet, cursor, 12, f"=TRUNC(F{cursor}+H{cursor}+J{cursor},1)")
        _set_cell(sheet, cursor, 13, None)
        cursor += 1

        first_data = material_row
        last_data = cursor - 1
        sum_row = cursor
        _set_cell(sheet, cursor, 1, " [ 합          계 ]", font=BODY_FONT)
        for col in range(2, 6):
            _set_cell(sheet, cursor, col, None)
        _money(sheet, cursor, 6, f"=SUM(F{first_data}:F{last_data})")
        _set_cell(sheet, cursor, 7, None)
        _money(sheet, cursor, 8, f"=SUM(H{first_data}:H{last_data})")
        _set_cell(sheet, cursor, 9, None)
        _money(sheet, cursor, 10, f"=SUM(J{first_data}:J{last_data})")
        _set_cell(sheet, cursor, 11, None)
        _money(sheet, cursor, 12, f"=SUM(L{first_data}:L{last_data})")
        _set_cell(sheet, cursor, 13, None)
        blocks.append(
            IlwidaeBlock(
                item=item,
                ho_no=ho_no,
                title_row=title_row,
                material_row=material_row,
                sum_row=sum_row,
                labor_rows=labor_rows,
                pumsam_ref=ref,
            )
        )
        cursor += 2

    _apply_sheet_look(sheet, max(cursor, 4), 13, row_height=FORM_ROW_HEIGHT)
    sheet.column_dimensions["A"].width = 40
    sheet.column_dimensions["B"].width = 18
    sheet.column_dimensions["C"].width = 8
    sheet.column_dimensions["D"].width = 11
    for col in range(5, 14):
        sheet.column_dimensions[get_column_letter(col)].width = 14
    return blocks
