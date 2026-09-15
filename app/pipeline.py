"""단가대비표 → 일위대가 → 일위대가목록 → 공량산출서 파이프라인."""

from __future__ import annotations

from pathlib import Path
from typing import Any

from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

from app.discipline import normalize_discipline
from app.estimate_parse import (
    EstimateSheet,
    list_sheet_titles,
    load_estimate_sheet,
    load_named_sheet,
    normalize_header,
)
from app.excel_io import (
    AMOUNT_FORMAT,
    BODY_FONT,
    CENTER,
    COMPARE_DATA_START,
    COMPARE_PRICE_COL,
    COMPARE_SHEET_NAME,
    ESTIMATE_SHEET_NAME,
    FORM_ROW_HEIGHT,
    HEADER_FONT,
    ILWIDAE_LIST_SHEET_NAME,
    ILWIDAE_SHEET_NAME,
    PAGE_FORMAT,
    PRICE_FORMAT,
    PUMSAM_DATA_START,
    QUANTITY_SHEET_NAME,
    RIGHT,
    _apply_sheet_look,
    _set_cell,
    _write_estimate_sheet,
    _write_pumsam_sheet,
    _write_quantity_sheet,
    collect_pumsam_rows,
    write_title_banner,
)
from app.ilwidae import (
    CD_FITTING_RATE,
    CONDUIT_FITTING_RATE,
    IlwidaeBlock,
    SUNDRY_RATE,
    TOOL_RATE,
    write_ilwidae_sheet,
)
from app.items import LineItem, first_filled, parse_line_items
from app.paths import assert_safe_save, build_result_path
from app.pumsam import PUMSAM_SHEET_NAME, PumsamRow
from app.wages import WAGES_SHEET_NAME, WageRow, load_wages, save_wages

COMPARE_LAST_COL = 21


def remap_items_to_compare(items: list[LineItem]) -> None:
    """다시 쓴 단가대비표 행에 맞춰 단가 열을 L(적용단가)로 고정한다. 코드 열은 없다."""
    excel_row = COMPARE_DATA_START
    for item in items:
        if item.section:
            continue
        item.excel_row = excel_row
        item.price_col = COMPARE_PRICE_COL
        item.qty_col = None
        excel_row += 1


def _write_price_cell(sheet: Worksheet, row: int, col: int, value: Any, *, page: bool = False) -> None:
    if value in (None, ""):
        _set_cell(sheet, row, col, None)
        return
    if page:
        _set_cell(sheet, row, col, value, font=BODY_FONT, align=CENTER, number_format=PAGE_FORMAT)
        return
    _set_cell(
        sheet,
        row,
        col,
        value,
        font=BODY_FONT,
        align=RIGHT,
        number_format=PRICE_FORMAT,
    )


def _write_compare_from_items(sheet: Worksheet, items: list[LineItem]) -> None:
    """단가대비표는 원본 서식을 복사하지 않는다. 코드 열은 만들지 않는다."""
    sheet.title = COMPARE_SHEET_NAME
    last_col = COMPARE_LAST_COL
    write_title_banner(sheet, "단 가 대 비 표", last_col)

    top = {
        1: "품      명",
        2: "규격",
        3: "단위",
        4: "재  료  비",
        13: "노 무 비",
        14: "경    비",
        20: "번  호",
        21: "비  고",
    }
    for col in range(1, last_col + 1):
        _set_cell(sheet, 3, col, top.get(col), font=HEADER_FONT, align=CENTER)
        _set_cell(sheet, 4, col, None, font=HEADER_FONT, align=CENTER)
    material_subs = {
        4: "물가정보",
        5: "PAGE",
        6: "조달청",
        7: "PAGE",
        8: "조사가격2",
        9: "PAGE",
        10: "조사가격3",
        11: "PAGE",
        12: "적용단가",
    }
    labor_subs = {
        14: "조달청가격",
        15: "거래가격",
        16: "유통물가",
        17: "조사가격1",
        18: "조사가격2",
        19: "적용단가",
    }
    for col, text in {**material_subs, **labor_subs}.items():
        _set_cell(sheet, 4, col, text, font=HEADER_FONT, align=CENTER)
    sheet.merge_cells("A3:A4")
    sheet.merge_cells("B3:B4")
    sheet.merge_cells("C3:C4")
    sheet.merge_cells("D3:L3")
    sheet.merge_cells("M3:M4")
    sheet.merge_cells("N3:S3")
    sheet.merge_cells("T3:T4")
    sheet.merge_cells("U3:U4")

    excel_row = COMPARE_DATA_START
    serial = 0
    for item in items:
        if item.section:
            continue
        serial += 1
        fields = item.fields or {}
        _set_cell(sheet, excel_row, 1, item.name, font=BODY_FONT)
        _set_cell(sheet, excel_row, 2, item.spec, font=BODY_FONT)
        _set_cell(sheet, excel_row, 3, item.unit, font=BODY_FONT, align=CENTER)
        for col in range(4, last_col + 1):
            _set_cell(sheet, excel_row, col, None)
        info = first_filled(fields.get("물가정보"))
        pps = first_filled(fields.get("조달청"))
        survey2 = first_filled(fields.get("조사가격2"))
        survey3 = first_filled(fields.get("조사가격3"))
        if info is None and pps is None and survey2 is None and survey3 is None:
            info = first_filled(fields.get("단가"), fields.get("재료비"), item.material_price)
        applied = first_filled(
            fields.get("적용단가"),
            info,
            pps,
            survey2,
            survey3,
            item.material_price,
        )
        _write_price_cell(sheet, excel_row, 4, info)
        _write_price_cell(sheet, excel_row, 5, fields.get("물가정보_PAGE"), page=True)
        _write_price_cell(sheet, excel_row, 6, pps)
        _write_price_cell(sheet, excel_row, 7, fields.get("조달청_PAGE"), page=True)
        _write_price_cell(sheet, excel_row, 8, survey2)
        _write_price_cell(sheet, excel_row, 9, fields.get("조사가격2_PAGE"), page=True)
        _write_price_cell(sheet, excel_row, 10, survey3)
        _write_price_cell(sheet, excel_row, 11, fields.get("조사가격3_PAGE"), page=True)
        _write_price_cell(sheet, excel_row, 12, applied)
        _set_cell(sheet, excel_row, 20, f"자재 {serial}", font=BODY_FONT, align=CENTER)
        excel_row += 1

    _apply_sheet_look(sheet, max(excel_row - 1, 4), last_col, row_height=FORM_ROW_HEIGHT)
    sheet.column_dimensions["A"].width = 32
    sheet.column_dimensions["B"].width = 16
    sheet.column_dimensions["C"].width = 8
    page_cols = {5, 7, 9, 11}  # E, G, I, K — PAGE 열너비 5
    for col in range(4, last_col + 1):
        sheet.column_dimensions[get_column_letter(col)].width = 5 if col in page_cols else 12


def _write_wages_sheet(sheet: Worksheet, rows: list[WageRow]) -> None:
    sheet.title = WAGES_SHEET_NAME
    write_title_banner(sheet, "노 임 단 가", 3)
    _set_cell(sheet, 3, 1, "직종", font=HEADER_FONT, align=CENTER)
    _set_cell(sheet, 3, 2, "노임단가", font=HEADER_FONT, align=CENTER)
    _set_cell(sheet, 3, 3, "비고", font=HEADER_FONT, align=CENTER)
    last = 3
    for offset, row in enumerate(rows):
        excel_row = offset + 4
        last = excel_row
        _set_cell(sheet, excel_row, 1, row.get("직종"), font=BODY_FONT)
        _set_cell(
            sheet,
            excel_row,
            2,
            row.get("노임단가"),
            font=BODY_FONT,
            align=RIGHT,
            number_format=PRICE_FORMAT,
        )
        _set_cell(sheet, excel_row, 3, row.get("비고"), font=BODY_FONT)
    _apply_sheet_look(sheet, last, 3)
    sheet.row_dimensions[1].height = FORM_ROW_HEIGHT
    sheet.row_dimensions[2].height = FORM_ROW_HEIGHT
    sheet.column_dimensions["A"].width = 18
    sheet.column_dimensions["B"].width = 14
    sheet.column_dimensions["C"].width = 40


SUNDRY_LABOR_JOBS = ("내선전공", "저압케이블전공", "통신케이블공")


def _sumproduct_amount(item_first: int, item_last: int, *search_flags: str) -> str:
    """내역서 품목 구간의 재료비 금액(F)을 품명 조건으로 더한다."""
    cells_a = f"$A${item_first}:$A${item_last}"
    cells_f = f"$F${item_first}:$F${item_last}"
    terms = [f"(ISNUMBER(SEARCH(\"{token}\",{cells_a})))" for token in search_flags]
    joined = "*".join(terms)
    return f"SUMPRODUCT({joined}*({cells_f}))"


def _write_estimate_sundry_form(
    sheet: Worksheet,
    excel_row: int,
    *,
    item_first_row: int,
    item_last_row: int,
    last_col: int,
    filled: list[list[Any]],
) -> int:
    """내역서 아래에 부속재·잡자재·노무비·공구손료 통합 양식을 넣는다."""
    first = item_first_row
    last = max(item_last_row, item_first_row)
    cd_base = _sumproduct_amount(first, last, "전선관", "CD")
    conduit_only = (
        f"SUMPRODUCT((ISNUMBER(SEARCH(\"전선관\",$A${first}:$A${last})))*"
        f"(NOT(ISNUMBER(SEARCH(\"CD\",$A${first}:$A${last}))))*"
        f"($F${first}:$F${last}))"
    )
    wire_base = (
        f"SUMPRODUCT(((ISNUMBER(SEARCH(\"전선\",$A${first}:$A${last})))+"
        f"(ISNUMBER(SEARCH(\"케이블\",$A${first}:$A${last}))))*"
        f"($F${first}:$F${last}))"
    )

    def _push(row: list[Any]) -> None:
        padded = list(row) + [None] * (last_col - len(row))
        filled.append(padded[:last_col])

    def _unit_price(row: int, amount_col: str) -> str:
        return f'=IF(D{row}=0,0,TRUNC({amount_col}{row}/D{row},2))'

    def _line_total(row: int) -> str:
        return f"=TRUNC(F{row}+H{row}+J{row},1)"

    rows: list[tuple[str, str, str, Any, str | None, str | None]] = [
        (
            "[ 배관 부속재 ]",
            "CD 전선관의 40 %",
            "식",
            1,
            "F",
            f"=TRUNC({cd_base}*{CD_FITTING_RATE},1)",
        ),
        (
            "[ 배관 부속재 ]",
            "전선관의 15 %",
            "식",
            1,
            "F",
            f"=TRUNC({conduit_only}*{CONDUIT_FITTING_RATE},1)",
        ),
        (
            "[ 소모 잡자재 ]",
            "전선, 전선관의 2 %",
            "식",
            1,
            "F",
            f"=TRUNC({wire_base}*{SUNDRY_RATE},1)",
        ),
    ]
    labor_start = excel_row + len(rows)
    for job in SUNDRY_LABOR_JOBS:
        job_lit = str(job).replace('"', '""')
        qty = f"=TRUNC(SUMIF('{ILWIDAE_SHEET_NAME}'!$A:$A,\"{job_lit}\",'{ILWIDAE_SHEET_NAME}'!$D:$D),0)"
        amount = f"=SUMIF('{ILWIDAE_SHEET_NAME}'!$A:$A,\"{job_lit}\",'{ILWIDAE_SHEET_NAME}'!$H:$H)"
        rows.append(("노 무 비", job, "인", qty, "H", amount))

    labor_end = labor_start + len(SUNDRY_LABOR_JOBS) - 1
    tool_labor_sum = "+".join(f"H{r}" for r in range(labor_start, labor_end + 1))
    rows.append(
        (
            "[ 공 구 손 료 ]",
            "직접노무비의 3 %",
            "식",
            1,
            "J",
            f"=TRUNC(({tool_labor_sum})*{TOOL_RATE},1)",
        )
    )

    form_first = excel_row
    for name, spec, unit, qty, amount_col, amount_formula in rows:
        _set_cell(sheet, excel_row, 1, name, font=BODY_FONT)
        _set_cell(sheet, excel_row, 2, spec, font=BODY_FONT)
        _set_cell(sheet, excel_row, 3, unit, font=BODY_FONT, align=CENTER)
        qty_format = "#,##0" if unit == "인" else AMOUNT_FORMAT
        _set_cell(sheet, excel_row, 4, qty, font=BODY_FONT, align=RIGHT, number_format=qty_format)
        for col in range(5, last_col + 1):
            _set_cell(sheet, excel_row, col, None)
        if amount_col == "F":
            _set_cell(sheet, excel_row, 6, amount_formula, font=BODY_FONT, align=RIGHT, number_format=AMOUNT_FORMAT)
            _set_cell(sheet, excel_row, 5, _unit_price(excel_row, "F"), font=BODY_FONT, align=RIGHT, number_format=PRICE_FORMAT)
        elif amount_col == "H":
            job_lit = str(spec).replace('"', '""')
            _set_cell(
                sheet,
                excel_row,
                7,
                f"=IFERROR(VLOOKUP(\"{job_lit}\",'{WAGES_SHEET_NAME}'!A:B,2,FALSE),0)",
                font=BODY_FONT,
                align=RIGHT,
                number_format=PRICE_FORMAT,
            )
            _set_cell(sheet, excel_row, 8, amount_formula, font=BODY_FONT, align=RIGHT, number_format=AMOUNT_FORMAT)
        elif amount_col == "J":
            _set_cell(sheet, excel_row, 10, amount_formula, font=BODY_FONT, align=RIGHT, number_format=AMOUNT_FORMAT)
            _set_cell(sheet, excel_row, 9, _unit_price(excel_row, "J"), font=BODY_FONT, align=RIGHT, number_format=PRICE_FORMAT)
        _set_cell(
            sheet,
            excel_row,
            12,
            _line_total(excel_row),
            font=BODY_FONT,
            align=RIGHT,
            number_format=AMOUNT_FORMAT,
        )
        _push([name, spec, unit, qty])
        excel_row += 1

    form_last = excel_row - 1
    _set_cell(sheet, excel_row, 1, "( 합 계 )", font=BODY_FONT)
    for col in range(2, last_col + 1):
        _set_cell(sheet, excel_row, col, None)
    _set_cell(
        sheet,
        excel_row,
        6,
        f"=SUM(F{form_first}:F{form_last})+SUM(F{first}:F{last})",
        font=BODY_FONT,
        align=RIGHT,
        number_format=AMOUNT_FORMAT,
    )
    _set_cell(
        sheet,
        excel_row,
        8,
        f"=SUM(H{form_first}:H{form_last})",
        font=BODY_FONT,
        align=RIGHT,
        number_format=AMOUNT_FORMAT,
    )
    _set_cell(
        sheet,
        excel_row,
        10,
        f"=SUM(J{form_first}:J{form_last})",
        font=BODY_FONT,
        align=RIGHT,
        number_format=AMOUNT_FORMAT,
    )
    _set_cell(
        sheet,
        excel_row,
        12,
        f"=TRUNC(F{excel_row}+H{excel_row}+J{excel_row},1)",
        font=BODY_FONT,
        align=RIGHT,
        number_format=AMOUNT_FORMAT,
    )
    _push(["( 합 계 )"])
    return excel_row + 1


def _write_generated_estimate(
    sheet: Worksheet,
    items: list[LineItem],
    blocks: list[IlwidaeBlock],
) -> EstimateSheet:
    """일위대가목록 품목은 일위대가 재료비 금액만 연결한다. 부가세·노무는 아래 양식에서 합친다."""
    sheet.title = ILWIDAE_LIST_SHEET_NAME
    last_col = 13
    write_title_banner(sheet, "[일위대가목록]", last_col)
    headers = {
        1: "명칭",
        2: "규격",
        3: "단위",
        4: "수량",
        5: "재료비",
        7: "노무비",
        9: "경비",
        11: "합계",
        13: "비고",
    }
    for col in range(1, last_col + 1):
        _set_cell(sheet, 3, col, headers.get(col), font=HEADER_FONT, align=CENTER)
        _set_cell(sheet, 4, col, None, font=HEADER_FONT, align=CENTER)
    for col, text in (
        (5, "단가"),
        (6, "금액"),
        (7, "단가"),
        (8, "금액"),
        (9, "단가"),
        (10, "금액"),
        (11, "단가"),
        (12, "금액"),
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

    block_by_key = {block.item.key: block for block in blocks}
    filled: list[list[Any]] = [[None] * last_col for _ in range(4)]
    filled[0][0] = "[일위대가목록]"
    filled[2][0] = "명칭"
    filled[2][1] = "규격"
    filled[2][2] = "단위"
    filled[2][3] = "수량"
    filled[2][4] = "재료비"
    filled[3][4] = "단가"
    filled[3][5] = "금액"

    excel_row = 5
    for item in items:
        if item.section:
            continue
        block = block_by_key.get(item.key)
        _set_cell(sheet, excel_row, 1, item.name, font=BODY_FONT)
        _set_cell(sheet, excel_row, 2, item.spec, font=BODY_FONT)
        _set_cell(sheet, excel_row, 3, item.unit, font=BODY_FONT, align=CENTER)
        qty_value: Any = item.qty if item.qty not in (None, "") else None
        _set_cell(
            sheet,
            excel_row,
            4,
            qty_value,
            font=BODY_FONT,
            align=RIGHT,
            number_format=AMOUNT_FORMAT,
        )
        for price_col in (5, 7, 9, 11):
            _set_cell(sheet, excel_row, price_col, None, font=BODY_FONT, align=RIGHT, number_format=PRICE_FORMAT)
        if block:
            material_amt = f"='{ILWIDAE_SHEET_NAME}'!F{block.material_row}"
        else:
            material_amt = None
        _set_cell(sheet, excel_row, 6, material_amt, font=BODY_FONT, align=RIGHT, number_format=AMOUNT_FORMAT)
        _set_cell(sheet, excel_row, 8, None, font=BODY_FONT, align=RIGHT, number_format=AMOUNT_FORMAT)
        _set_cell(sheet, excel_row, 10, None, font=BODY_FONT, align=RIGHT, number_format=AMOUNT_FORMAT)
        _set_cell(
            sheet,
            excel_row,
            12,
            f"=TRUNC(F{excel_row}+H{excel_row}+J{excel_row},1)",
            font=BODY_FONT,
            align=RIGHT,
            number_format=AMOUNT_FORMAT,
        )
        _set_cell(sheet, excel_row, 13, None)
        filled.append(
            [item.name, item.spec, item.unit, item.qty, None, None, None, None, None, None, None, None, None]
        )
        excel_row += 1

    item_last_row = excel_row - 1
    for _ in range(2):
        for col in range(1, last_col + 1):
            _set_cell(sheet, excel_row, col, None)
        filled.append([None] * last_col)
        excel_row += 1
    excel_row = _write_estimate_sundry_form(
        sheet,
        excel_row,
        item_first_row=5,
        item_last_row=max(item_last_row, 5),
        last_col=last_col,
        filled=filled,
    )

    _apply_sheet_look(sheet, max(excel_row - 1, 4), last_col, row_height=FORM_ROW_HEIGHT)
    sheet.column_dimensions["A"].width = 32
    sheet.column_dimensions["B"].width = 18
    sheet.column_dimensions["C"].width = 8
    sheet.column_dimensions["D"].width = 10
    for col in range(5, last_col + 1):
        sheet.column_dimensions[get_column_letter(col)].width = 12
    return EstimateSheet(filled=filled, raw=filled, merges=[], title=ILWIDAE_LIST_SHEET_NAME)


def _load_optional(path: Path | None) -> EstimateSheet | None:
    if path is None:
        return None
    return load_estimate_sheet(path)


def _is_ilwidae_extra_row(item: LineItem) -> bool:
    name = normalize_header(item.name)
    unit = normalize_header(item.unit)
    if "합계" in name or "부속품" in name:
        return True
    if name in {"잡재료비", "공구손료"}:
        return True
    if "배관부속재" in name or "소모잡자재" in name:
        return True
    if name == "노무비":
        return True
    if unit in {"인", "식"}:
        return True
    return False


def _items_for_compare(items: list[LineItem], *, from_ilwidae: bool) -> list[LineItem]:
    if not from_ilwidae:
        return items
    return [item for item in items if not item.section and not _is_ilwidae_extra_row(item)]


class _SheetFactory:
    def __init__(self, workbook: Workbook) -> None:
        self.workbook = workbook
        self._first = workbook.active
        self._used_first = False

    def take(self, title: str) -> Worksheet:
        if not self._used_first:
            self._used_first = True
            self._first.title = title
            return self._first
        return self.workbook.create_sheet(title)


def build_result_workbook(
    *,
    compare: EstimateSheet | None,
    estimate: EstimateSheet | None,
    pumsam_rows: list[PumsamRow],
    wage_rows: list[WageRow],
    discipline: str | None = None,
    mode: str = "forward",
    dropped_ilwidae: EstimateSheet | None = None,
    companion_sheets: dict[str, EstimateSheet] | None = None,
) -> Workbook:
    workbook = Workbook()
    sheets = _SheetFactory(workbook)
    items: list[LineItem] = []
    blocks: list[IlwidaeBlock] = []

    if mode == "reverse":
        source_items: list[LineItem] = []
        from_ilwidae = False
        if estimate is not None:
            source_items = parse_line_items(estimate)
        elif dropped_ilwidae is not None:
            source_items = parse_line_items(dropped_ilwidae)
            from_ilwidae = True
        compare_items = _items_for_compare(source_items, from_ilwidae=from_ilwidae)
        compare_sheet = sheets.take(COMPARE_SHEET_NAME)
        _write_compare_from_items(compare_sheet, compare_items)
        if estimate is not None:
            copied = sheets.take(ILWIDAE_LIST_SHEET_NAME)
            _write_estimate_sheet(copied, estimate)
            copied.title = ILWIDAE_LIST_SHEET_NAME
        if dropped_ilwidae is not None:
            copied_ilwidae = sheets.take(ILWIDAE_SHEET_NAME)
            _write_estimate_sheet(copied_ilwidae, dropped_ilwidae)
            copied_ilwidae.title = ILWIDAE_SHEET_NAME
    elif mode == "quantity":
        if estimate is None:
            raise ValueError("공량산출을 하려면 내역서 또는 일위대가목록이 있는 엑셀을 놓아 주세요.")
        companions = companion_sheets or {}
        copied_names: list[str] = []
        for name in (
            COMPARE_SHEET_NAME,
            ILWIDAE_SHEET_NAME,
            ILWIDAE_LIST_SHEET_NAME,
            ESTIMATE_SHEET_NAME,
        ):
            data = companions.get(name)
            if data is None:
                continue
            copied = sheets.take(name)
            _write_estimate_sheet(copied, data)
            copied.title = name
            copied_names.append(name)
        qty_title = estimate.title or ESTIMATE_SHEET_NAME
        if qty_title not in copied_names:
            copied = sheets.take(qty_title)
            _write_estimate_sheet(copied, estimate)
            copied.title = qty_title
        pumsam_last = PUMSAM_DATA_START + len(pumsam_rows) - 1 if pumsam_rows else PUMSAM_DATA_START
        qty = sheets.take(QUANTITY_SHEET_NAME)
        _write_quantity_sheet(
            qty,
            estimate,
            pumsam_last,
            pumsam_rows,
            source_sheet_name=qty_title,
        )
    else:
        if compare is not None:
            items = parse_line_items(compare)
            compare_sheet = sheets.take(COMPARE_SHEET_NAME)
            _write_compare_from_items(compare_sheet, items)
            remap_items_to_compare(items)
        if dropped_ilwidae is not None:
            ilwidae_sheet = sheets.take(ILWIDAE_SHEET_NAME)
            _write_estimate_sheet(ilwidae_sheet, dropped_ilwidae)
            ilwidae_sheet.title = ILWIDAE_SHEET_NAME
        elif items:
            ilwidae_sheet = sheets.take(ILWIDAE_SHEET_NAME)
            blocks = write_ilwidae_sheet(
                ilwidae_sheet,
                items,
                pumsam_rows,
                wage_rows,
                compare_sheet=COMPARE_SHEET_NAME,
                discipline=discipline,
            )
        if estimate is not None:
            copied = sheets.take(ILWIDAE_LIST_SHEET_NAME)
            _write_estimate_sheet(copied, estimate)
            copied.title = ILWIDAE_LIST_SHEET_NAME
        elif items:
            generated = sheets.take(ILWIDAE_LIST_SHEET_NAME)
            _write_generated_estimate(generated, items, blocks)

    pumsam_sheet = sheets.take(PUMSAM_SHEET_NAME)
    _write_pumsam_sheet(pumsam_sheet, pumsam_rows)
    wages_sheet = sheets.take(WAGES_SHEET_NAME)
    _write_wages_sheet(wages_sheet, wage_rows)
    return workbook


def run_pipeline(
    *,
    dest_dir: Path | None = None,
    unit_price_path: Path | None = None,
    ilwidae_path: Path | None = None,
    estimate_path: Path | None = None,
    extra_pumsam_path: Path | None = None,
    db_dir: Path | None = None,
    estimate_rows: list[list[Any]] | None = None,
    discipline: str | None = None,
    mode: str | None = None,
) -> Path:
    database_dir = db_dir if db_dir is not None else dest_dir
    disc = normalize_discipline(discipline)
    primary = unit_price_path or ilwidae_path or estimate_path
    if primary is None and estimate_rows is None:
        raise ValueError("단가대비표, 일위대가, 일위대가목록, 내역서 중 하나를 놓아 주세요.")

    if mode is None:
        if unit_price_path is not None:
            mode = "forward"
        else:
            mode = "reverse"

    compare = _load_optional(Path(unit_price_path) if unit_price_path else None)
    dropped_ilwidae = _load_optional(Path(ilwidae_path) if ilwidae_path else None)
    if estimate_rows is not None:
        estimate = EstimateSheet(filled=estimate_rows, raw=estimate_rows, merges=[])
    else:
        estimate = _load_optional(Path(estimate_path) if estimate_path else None)

    companion_sheets: dict[str, EstimateSheet] = {}
    if mode == "quantity" and estimate_path is not None:
        source = Path(estimate_path)
        titles = list_sheet_titles(source)
        for name in (
            COMPARE_SHEET_NAME,
            ILWIDAE_SHEET_NAME,
            ILWIDAE_LIST_SHEET_NAME,
            ESTIMATE_SHEET_NAME,
        ):
            if name not in titles:
                continue
            loaded = load_named_sheet(source, name)
            if loaded is not None:
                companion_sheets[name] = loaded
        if ESTIMATE_SHEET_NAME in companion_sheets:
            estimate = companion_sheets[ESTIMATE_SHEET_NAME]
        elif ILWIDAE_LIST_SHEET_NAME in companion_sheets:
            estimate = companion_sheets[ILWIDAE_LIST_SHEET_NAME]

    if compare is None and dropped_ilwidae is not None and estimate is None and mode == "forward":
        compare = dropped_ilwidae
        dropped_ilwidae = None

    source_for_pumsam = unit_price_path or estimate_path or ilwidae_path
    pumsam_rows = collect_pumsam_rows(
        Path(source_for_pumsam) if source_for_pumsam else None,
        database_dir,
        extra_pumsam_path,
        db_dir=database_dir,
        discipline=disc,
    )
    wage_rows = load_wages(database_dir)
    save_wages(wage_rows, database_dir)

    dest = build_result_path(dest_dir, mode=mode)
    if primary is not None:
        assert_safe_save(Path(primary), dest)

    workbook = build_result_workbook(
        compare=compare,
        estimate=estimate,
        pumsam_rows=pumsam_rows,
        wage_rows=wage_rows,
        discipline=disc,
        mode=mode,
        dropped_ilwidae=dropped_ilwidae,
        companion_sheets=companion_sheets or None,
    )
    try:
        workbook.save(dest)
    finally:
        workbook.close()
    return dest
