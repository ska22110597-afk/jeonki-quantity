from __future__ import annotations

from pathlib import Path

from openpyxl import Workbook, load_workbook

from app.excel_io import (
    ESTIMATE_SHEET_NAME,
    NUMBER_FORMAT,
    QUANTITY_SHEET_NAME,
    calc_qty_formula,
    gongryang_formula,
    gongryang_sum_formula,
    read_unit_price_table,
    save_result_workbook,
)
from app.estimate_parse import is_section_row, lookup_key
from app.pumsam import PUMSAM_SHEET_NAME


def _write_estimate(path: Path, *, with_merge: bool = False) -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "BHU기존내역서"
    sheet.append(["BHU 기존내역서"])
    sheet.append([None, None, None, None, "재료비", None, "노무비"])
    sheet.append(["명칭", "규격", "단위", "수량", "단가", "금액", "단가"])
    sheet.append(["1. 옥외전기공사", None, None, None, None, None, None])
    sheet.append(["경질비닐전선관_지중", "HI 16 mm", "M", 10, 0, 0, 0])
    sheet.append([None if with_merge else "경질비닐전선관_지중", "HI 22 mm", "M", 8, 0, 0, 0])
    sheet.append(["경질비닐전선관_노출", "HI 16 mm", "M", 4, 0, 0, 0])
    if with_merge:
        sheet.merged_cells.add("A5:A6")
    workbook.save(path)
    workbook.close()


def test_section_row_detection() -> None:
    assert is_section_row("1. 옥외전기공사", None, None)
    assert not is_section_row("경질비닐전선관_지중", "HI 16 mm", "M")


def test_read_only_skips_title_keeps_items(tmp_path: Path) -> None:
    source = tmp_path / "내역서.xlsx"
    _write_estimate(source)
    before = source.read_bytes()
    rows = read_unit_price_table(source)
    assert source.read_bytes() == before
    assert rows[0][0] == "명칭"
    assert rows[1][0] == "1. 옥외전기공사"
    assert rows[2][0] == "경질비닐전선관_지중"


def test_merged_name_fills_without_rewriting_source(tmp_path: Path) -> None:
    source = tmp_path / "내역서_병합.xlsx"
    _write_estimate(source, with_merge=True)
    before = source.read_bytes()
    rows = read_unit_price_table(source)
    assert source.read_bytes() == before
    assert rows[2][0] == "경질비닐전선관_지중"
    assert rows[3][0] == "경질비닐전선관_지중"


def test_three_sheets_formulas_and_pumsam_lookup(tmp_path: Path) -> None:
    source = tmp_path / "내역서.xlsx"
    _write_estimate(source)
    dest_dir = tmp_path / "결과"
    dest = save_result_workbook(source, dest_dir=dest_dir)

    assert dest.name.startswith("공량산출_결과_")
    assert source.read_bytes()  # still exists

    result = load_workbook(dest, data_only=False)
    try:
        assert result.sheetnames == [ESTIMATE_SHEET_NAME, PUMSAM_SHEET_NAME, QUANTITY_SHEET_NAME]
        estimate = result[ESTIMATE_SHEET_NAME]
        assert estimate["A1"].value == "명칭"
        assert estimate["D3"].value == 10

        pumsam = result[PUMSAM_SHEET_NAME]
        assert pumsam["A1"].value == "검색키"
        assert pumsam["B2"].value == "경질비닐전선관_지중"
        assert str(pumsam["A2"].value).startswith("=SUBSTITUTE")

        qty = result[QUANTITY_SHEET_NAME]
        assert [qty.cell(4, c).value for c in range(1, 13)][4] == "결정수량"
        assert qty["B5"].value == "경질비닐전선관_지중"
        assert qty["C5"].value == "HI 16 mm"
        assert qty["E5"].value == "='내역서'!D3"
        assert qty["F5"].value == 0
        assert qty["G5"].value == calc_qty_formula(5)
        assert "품셈표" in str(qty["H5"].value)
        assert "VLOOKUP" in str(qty["I5"].value)
        assert qty["K5"].value == gongryang_formula(5)
        assert qty["E5"].number_format == NUMBER_FORMAT
        assert qty["K8"].value == gongryang_sum_formula(7)
        assert qty["A8"].value == "합계"
        # 공종 제목 행은 공량산출서에 안 들어간다.
        names = [qty.cell(r, 2).value for r in range(5, 8)]
        assert "1. 옥외전기공사" not in names
        assert lookup_key("경질비닐전선관_지중", "HI 16 mm") == "경질비닐전선관_지중HI16mm"
    finally:
        result.close()
