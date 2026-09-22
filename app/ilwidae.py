"""일위대가 호표 작성. 품명과 표준품셈 인부만 넣고, 부가세는 내역서 아래에 모은다."""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import Any

from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

from app.discipline import default_labor_name
from app.estimate_parse import normalize_header
from app.pumsam_aliases import alias_group
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
HO_TITLE_RE = re.compile(r"\(\s*호표\s*\d+\s*\)")


@dataclass
class IlwidaeBlock:
    item: LineItem
    ho_no: int
    title_row: int
    material_row: int
    sum_row: int
    labor_rows: list[int] = field(default_factory=list)
    pumsam_ref: str = ""


_CONDUIT_ANCHORS = ("강제전선관", "경질비닐전선관", "CD전선관", "금속제가요전선관", "박강전선관")
_WIRE_ANCHORS = ("옥내배선", "전력케이블", "제어용케이블", "UTP케이블", "광케이블")


def _in_anchor_group(name: Any, anchors: tuple[str, ...]) -> bool:
    group = alias_group(name)
    return any(anchor in group for anchor in anchors)


def is_conduit_name(name: Any) -> bool:
    """전선관과 같은 자재. HI관·후강관처럼 다른 말도 전선관으로 본다."""
    if _in_anchor_group(name, _CONDUIT_ANCHORS):
        return True
    text = str(name or "")
    return "전선관" in text and "부속" not in text


def is_cable_name(name: Any) -> bool:
    """전선·케이블. CV·HIV처럼 다른 말도 전선류로 본다. 전선관·트레이는 빼 둔다."""
    if is_conduit_name(name):
        return False
    if _in_anchor_group(name, _WIRE_ANCHORS):
        return True
    text = str(name or "")
    if any(word in text for word in ("트레이", "덕트", "몰딩")):
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


def quantity_ilwidae_refs(blocks: list[IlwidaeBlock]) -> list[tuple[str, int, list[int]]]:
    """공량산출서가 일위대가 호표를 찾을 때 쓰는 (키, 자재행, 인부행들)."""
    return [(block.item.key, block.material_row, list(block.labor_rows)) for block in blocks]


def renumber_ilwidae_ho(sheet: Worksheet) -> int:
    """시트에 적힌 호표를 위에서부터 1, 2, 3… 순서로 다시 붙인다."""
    count = 0
    for row in range(1, (sheet.max_row or 1) + 1):
        cell = sheet.cell(row, 1)
        value = cell.value
        if not isinstance(value, str) or "호표" not in value:
            continue
        count += 1
        label = f"( 호표 {count} )"
        if HO_TITLE_RE.search(value):
            cell.value = HO_TITLE_RE.sub(label, value, count=1)
        else:
            cell.value = re.sub(r"호표\s*\d+", f"호표 {count}", value, count=1)
    return count


def parse_ilwidae_blocks(sheet: Worksheet) -> list[IlwidaeBlock]:
    """이미 적힌 일위대가 시트에서 호표 구간을 읽는다."""
    blocks: list[IlwidaeBlock] = []
    title_row: int | None = None
    material_row: int | None = None
    labor_rows: list[int] = []
    name: Any = None
    spec: Any = None
    ho_no = 0

    def flush(sum_row: int | None = None) -> None:
        nonlocal title_row, material_row, labor_rows, name, spec, ho_no
        if title_row is None or material_row is None:
            title_row = None
            material_row = None
            labor_rows = []
            name = None
            spec = None
            return
        ho_no += 1
        blocks.append(
            IlwidaeBlock(
                item=LineItem(
                    excel_row=material_row,
                    name=name,
                    spec=spec,
                    unit=None,
                    qty=None,
                    material_price=None,
                ),
                ho_no=ho_no,
                title_row=title_row,
                material_row=material_row,
                sum_row=sum_row or (labor_rows[-1] + 1 if labor_rows else material_row + 1),
                labor_rows=list(labor_rows),
            )
        )
        title_row = None
        material_row = None
        labor_rows = []
        name = None
        spec = None

    for row in range(1, (sheet.max_row or 1) + 1):
        a = sheet.cell(row, 1).value
        b = sheet.cell(row, 2).value
        c = sheet.cell(row, 3).value
        a_text = str(a).strip() if a not in (None, "") else ""
        if not a_text and b in (None, ""):
            continue
        compact = a_text.replace(" ", "")
        if "호표" in a_text:
            if title_row is not None:
                flush()
            title_row = row
            continue
        if title_row is None:
            continue
        if "합계" in compact:
            flush(sum_row=row)
            continue
        unit = str(c).strip() if c not in (None, "") else ""
        if unit == "인":
            labor_rows.append(row)
            continue
        if material_row is None and a_text:
            material_row = row
            name = a
            spec = b
    if title_row is not None:
        flush()
    return blocks


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


def _write_percent_material_row(
    sheet: Worksheet,
    row: int,
    name: str,
    spec: str,
    material_row: int,
    rate: float,
) -> None:
    """부속품·잡재료. 규격에 비율을 적고, 수식은 재료비 금액(F)에 둔다."""
    _set_cell(sheet, row, 1, name, font=BODY_FONT)
    _set_cell(sheet, row, 2, spec, font=BODY_FONT)
    _set_cell(sheet, row, 3, "식", font=BODY_FONT, align=CENTER)
    _set_cell(sheet, row, 4, 1, font=BODY_FONT, align=RIGHT, number_format="0.000")
    _price(sheet, row, 5, f"=IF(D{row}=0,0,TRUNC(F{row}/D{row},2))")
    _amount(sheet, row, 6, f"=TRUNC(F{material_row}*{rate},1)")
    _idle_labor_expense(sheet, row)
    _cost_totals(sheet, row)
    _set_cell(sheet, row, 13, None)


def _write_tool_loss_row(sheet: Worksheet, row: int, labor_rows: list[int]) -> None:
    """기본 행. 직접노무비의 3%를 경비 금액에 넣는다."""
    _set_cell(sheet, row, 1, "공구손료", font=BODY_FONT)
    _set_cell(sheet, row, 2, "직접노무비의 3%", font=BODY_FONT)
    _set_cell(sheet, row, 3, "식", font=BODY_FONT, align=CENTER)
    _set_cell(sheet, row, 4, 1, font=BODY_FONT, align=RIGHT, number_format="0.000")
    _price(sheet, row, 5, _zero_price())
    _amount(sheet, row, 6, _qty_times_price(row, "E"))
    _price(sheet, row, 7, _zero_price())
    _amount(sheet, row, 8, _qty_times_price_trunc(row, "G"))
    if labor_rows:
        parts = "+".join(f"H{labor}" for labor in labor_rows)
        tool_amount = f"=TRUNC(({parts})*{TOOL_RATE},1)"
    else:
        tool_amount = "=0"
    _price(sheet, row, 9, f"=IF(D{row}=0,0,TRUNC(J{row}/D{row},2))")
    _amount(sheet, row, 10, tool_amount)
    _cost_totals(sheet, row)
    _set_cell(sheet, row, 13, None)


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

        conduit = is_conduit_name(item.name)
        wire = is_cable_name(item.name)
        if conduit:
            _write_percent_material_row(
                sheet, cursor, "전선관부속품비", "전선관의 15%", material_row, CONDUIT_FITTING_RATE
            )
            cursor += 1
        if conduit or wire:
            sundry_spec = "배관의 2%" if conduit else "배선의 2%"
            _write_percent_material_row(
                sheet, cursor, "잡재료비", sundry_spec, material_row, SUNDRY_RATE
            )
            cursor += 1

        labor_rows: list[int] = []
        for labor in labors:
            job = labor.get("노무명칭") or fallback_labor
            _set_cell(sheet, cursor, 1, job, font=BODY_FONT)
            _set_cell(sheet, cursor, 2, labor_kind_text(labor), font=BODY_FONT)
            _set_cell(sheet, cursor, 3, "인", font=BODY_FONT, align=CENTER)
            _set_cell(sheet, cursor, 4, None, font=BODY_FONT, align=RIGHT, number_format="0.000")
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

        _write_tool_loss_row(sheet, cursor, labor_rows)
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
