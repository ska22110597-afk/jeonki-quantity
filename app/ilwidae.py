"""일위대가 호표 작성. 품명과 표준품셈 인부만 넣고, 부가세는 내역서 아래에 모은다."""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any

from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

from app.discipline import default_labor_name
from app.estimate_parse import normalize_header
from app.excel_io import (
    AMOUNT_FORMAT,
    BODY_FONT,
    CENTER,
    FORM_ROW_HEIGHT,
    HEADER_FONT,
    ILWIDAE_SHEET_NAME,
    LEFT,
    PRICE_FORMAT,
    RIGHT,
    write_title_banner,
    _apply_sheet_look,
    _set_cell,
)
from app.items import LineItem
from app.pumsam import (
    PumsamRow,
    format_pumsam_ref,
    labor_kind_text,
    match_pumsam as match_pumsam_rows,
    pumsam_qty_value,
    pumsam_rate_value,
)
from app.wages import WAGES_SHEET_NAME, WageRow

PIPE_WASTE_QTY = 1.1
CD_FITTING_RATE = 0.40
CONDUIT_FITTING_RATE = 0.15
SUNDRY_RATE = 0.02
TOOL_RATE = 0.03


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


def _price(sheet: Worksheet, row: int, col: int, value: Any, align=RIGHT) -> None:
    _set_cell(sheet, row, col, value, font=BODY_FONT, align=align, number_format=PRICE_FORMAT)


def _amount(sheet: Worksheet, row: int, col: int, value: Any, align=RIGHT) -> None:
    _set_cell(sheet, row, col, value, font=BODY_FONT, align=align, number_format=AMOUNT_FORMAT)


def _cost_totals(sheet: Worksheet, row: int) -> None:
    """단가=재료+노무+경비(소수 2자리), 금액=재료+노무+경비(소수 1자리)."""
    _price(sheet, row, 11, f"=TRUNC(E{row}+G{row}+I{row},2)")
    _amount(sheet, row, 12, f"=TRUNC(F{row}+H{row}+J{row},1)")


def _zero_price() -> str:
    return "=0"


def _qty_times_price(row: int, price_col: str) -> str:
    """금액 = 단가×수량, 소수 1자리."""
    return f"=TRUNC(D{row}*{price_col}{row},1)"


def _qty_times_price_trunc(row: int, price_col: str) -> str:
    """노무·경비 금액. 소수 1자리."""
    return f"=TRUNC({price_col}{row}*D{row},1)"


def _idle_labor_expense(sheet: Worksheet, row: int) -> None:
    """재료 행에서 쓰지 않는 노무·경비 칸도 수식으로 둔다."""
    _price(sheet, row, 7, _zero_price())
    _amount(sheet, row, 8, _qty_times_price_trunc(row, "G"))
    _price(sheet, row, 9, _zero_price())
    _amount(sheet, row, 10, _qty_times_price_trunc(row, "I"))


def _idle_material_expense(sheet: Worksheet, row: int) -> None:
    """노무 행에서 쓰지 않는 재료·경비 칸도 수식으로 둔다."""
    _price(sheet, row, 5, _zero_price())
    _amount(sheet, row, 6, _qty_times_price(row, "E"))
    _price(sheet, row, 9, _zero_price())
    _amount(sheet, row, 10, _qty_times_price_trunc(row, "I"))


def _price_ref(item: LineItem, compare_sheet: str) -> str:
    col_index = (item.price_col if item.price_col is not None else 11) + 1
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
        for col in range(2, 14):
            _set_cell(sheet, cursor, col, None)
        sheet.merge_cells(start_row=cursor, start_column=1, end_row=cursor, end_column=13)
        cursor += 1

        material_row = cursor
        qty = material_qty(item)
        _set_cell(sheet, cursor, 1, item.name, font=BODY_FONT)
        _set_cell(sheet, cursor, 2, item.spec, font=BODY_FONT)
        _set_cell(sheet, cursor, 3, item.unit or "개", font=BODY_FONT, align=CENTER)
        _set_cell(sheet, cursor, 4, qty, font=BODY_FONT, align=RIGHT, number_format="0.000")
        _price(sheet, cursor, 5, _price_ref(item, compare_sheet))
        _amount(sheet, cursor, 6, _qty_times_price(cursor, "E"))
        _idle_labor_expense(sheet, cursor)
        _cost_totals(sheet, cursor)
        _set_cell(sheet, cursor, 13, ref or None, font=BODY_FONT, align=CENTER)
        cursor += 1

        labor_rows: list[int] = []
        for labor in labors:
            job = labor.get("노무명칭") or fallback_labor
            _set_cell(sheet, cursor, 1, job, font=BODY_FONT)
            _set_cell(sheet, cursor, 2, labor_kind_text(labor), font=BODY_FONT)
            _set_cell(sheet, cursor, 3, "인", font=BODY_FONT, align=CENTER)
            _set_cell(sheet, cursor, 4, labor_qty_formula(labor), font=BODY_FONT, align=RIGHT, number_format="0.000")
            job_lit = str(job).replace('"', '""')
            _idle_material_expense(sheet, cursor)
            _price(
                sheet,
                cursor,
                7,
                f"=IFERROR(VLOOKUP(\"{job_lit}\",'{WAGES_SHEET_NAME}'!A:B,2,FALSE),0)",
            )
            _amount(sheet, cursor, 8, f"=TRUNC(G{cursor}*D{cursor},1)")
            _cost_totals(sheet, cursor)
            _set_cell(sheet, cursor, 13, None)
            labor_rows.append(cursor)
            cursor += 1

        first_data = material_row
        last_data = cursor - 1
        sum_row = cursor
        _set_cell(sheet, cursor, 1, " [ 합          계 ]", font=BODY_FONT)
        for col in range(2, 14):
            _set_cell(sheet, cursor, col, None)
        _amount(sheet, cursor, 6, f"=SUM(F{first_data}:F{last_data})")
        _amount(sheet, cursor, 8, f"=SUM(H{first_data}:H{last_data})")
        _amount(sheet, cursor, 10, f"=SUM(J{first_data}:J{last_data})")
        _amount(sheet, cursor, 12, f"=TRUNC(F{cursor}+H{cursor}+J{cursor},1)")
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
    sheet.column_dimensions["B"].width = 28
    sheet.column_dimensions["C"].width = 8
    sheet.column_dimensions["D"].width = 11
    for col in range(5, 13):
        sheet.column_dimensions[get_column_letter(col)].width = 14
    sheet.column_dimensions["M"].width = 24
    return blocks
