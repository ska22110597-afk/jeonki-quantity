from __future__ import annotations

from pathlib import Path

from openpyxl import Workbook, load_workbook

from app.excel_io import (
    COMPARE_SHEET_NAME,
    ESTIMATE_SHEET_NAME,
    ILWIDAE_SHEET_NAME,
    QUANTITY_SHEET_NAME,
    save_result_workbook,
)
from app.ilwidae import is_conduit_name, labor_qty_formula, match_pumsam
from app.items import LineItem
from app.pumsam import PUMSAM_SHEET_NAME, default_pumsam_rows
from app.wages import WAGES_SHEET_NAME, default_wage_rows


def _write_compare(path: Path) -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "단가대비표"
    sheet["A1"] = "단 가 대 비 표"
    sheet["A2"] = "명칭"
    sheet["B2"] = "규격"
    sheet["C2"] = "단위"
    sheet["D2"] = "수량"
    sheet["E2"] = "재료비"
    sheet["E3"] = "단가"
    sheet["F3"] = "금액"
    sheet.merge_cells("A2:A3")
    sheet.merge_cells("B2:B3")
    sheet.merge_cells("C2:C3")
    sheet.merge_cells("D2:D3")
    sheet.merge_cells("E2:F2")
    sheet["A4"] = "경질비닐전선관_지중"
    sheet["B4"] = "HI 16 mm"
    sheet["C4"] = "M"
    sheet["D4"] = 100
    sheet["E4"] = 200
    sheet["A5"] = "경질비닐전선관_노출"
    sheet["B5"] = "HI 104 mm"
    sheet["C5"] = "M"
    sheet["D5"] = 40
    sheet["E5"] = 151
    workbook.save(path)
    workbook.close()


def test_conduit_extras_and_two_labors(tmp_path: Path) -> None:
    source = tmp_path / "단가대비표.xlsx"
    _write_compare(source)
    dest = save_result_workbook(unit_price_path=source, dest_dir=tmp_path / "out")
    result = load_workbook(dest, data_only=False)
    try:
        assert result.sheetnames[0] == COMPARE_SHEET_NAME
        assert ILWIDAE_SHEET_NAME in result.sheetnames
        assert ESTIMATE_SHEET_NAME in result.sheetnames
        assert QUANTITY_SHEET_NAME in result.sheetnames
        assert PUMSAM_SHEET_NAME in result.sheetnames
        assert WAGES_SHEET_NAME in result.sheetnames

        compare = result[COMPARE_SHEET_NAME]
        assert compare["A4"].value == "경질비닐전선관_지중"
        assert compare["E4"].value == 200

        ilwidae = result[ILWIDAE_SHEET_NAME]
        names = [ilwidae.cell(r, 1).value for r in range(1, 40)]
        assert "전선관부속품비" in names
        assert "잡재료비" in names
        assert "공구손료" in names
        assert "내선전공" in names
        assert "보통인부" in names
        assert ilwidae["E6"].value == "='단가대비표'!E4"
        assert "0.15" in str(ilwidae["F7"].value)
        assert any(str(ilwidae.cell(r, 7).value or "").find("노임단가") >= 0 for r in range(5, 20))

        estimate = result[ESTIMATE_SHEET_NAME]
        assert estimate["A5"].value == "경질비닐전선관_지중"
        assert "단가대비표" in str(estimate["D5"].value)
        assert "일위대가" in str(estimate["G5"].value)

        qty = result[QUANTITY_SHEET_NAME]
        assert qty["B5"].value == "경질비닐전선관_지중"
        assert qty["G5"].value == "='내역서'!D5"
    finally:
        result.close()

    db = tmp_path / "out" / "데이터베이스"
    assert (db / "표준품셈.xlsx").exists()
    assert (db / "노임단가.xlsx").exists()


def test_match_pumsam_keeps_two_labors() -> None:
    item = LineItem(excel_row=5, name="경질비닐전선관_노출", spec="HI 104 mm", unit="M", qty=1, material_price=1)
    matched = match_pumsam(item, default_pumsam_rows())
    jobs = [row.get("노무명칭") for row in matched]
    assert "내선전공" in jobs
    assert "보통인부" in jobs
    assert is_conduit_name(item.name)
    assert "*1.2" in labor_qty_formula(matched[0]) or "*120" in labor_qty_formula(matched[0])


def test_default_wages_include_trades() -> None:
    jobs = {row["직종"] for row in default_wage_rows()}
    assert "내선전공" in jobs
    assert "보통인부" in jobs
    assert "저압케이블전공" in jobs
