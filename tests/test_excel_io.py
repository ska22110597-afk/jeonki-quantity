from __future__ import annotations

from pathlib import Path

from openpyxl import Workbook, load_workbook

from app.excel_io import (
    NUMBER_FORMAT,
    QUANTITY_SHEET_NAME,
    UNIT_PRICE_SHEET_NAME,
    calc_qty_formula,
    gongryang_formula,
    gongryang_sum_formula,
    read_unit_price_table,
    save_result_workbook,
)


def _write_source(path: Path, *, with_merge: bool = False) -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "원본시트"
    sheet.append(["단가대비표"])
    sheet.append(["번호", "공종", "품명", "규격", "단위", "수량", "설계단가"])
    sheet.append([1, "배관", "HI-PVC 전선관", "16mm", "m", 12.5, 1200])
    sheet.append([2, None, "CD관", "22mm", "m", 8, 850])
    sheet.append([None, None, None, None, None, None, None])
    if with_merge:
        sheet.merge_cells("B3:B4")
    workbook.save(path)
    workbook.close()


def test_read_only_trims_empty_rows(tmp_path: Path) -> None:
    source = tmp_path / "단가대비표.xlsx"
    _write_source(source)
    before = source.read_bytes()

    rows = read_unit_price_table(source)

    assert source.read_bytes() == before
    assert rows[0][2] == "품명"
    assert len(rows) == 3
    assert rows[-1][2] == "CD관"


def test_merged_cells_fill_top_left_without_rewriting_source(tmp_path: Path) -> None:
    source = tmp_path / "단가대비_병합.xlsx"
    _write_source(source, with_merge=True)
    before = source.read_bytes()

    rows = read_unit_price_table(source)

    assert source.read_bytes() == before
    assert rows[1][1] == "배관"
    assert rows[2][1] == "배관"

    original = load_workbook(source)
    try:
        merged = original.active.merged_cells.ranges
        assert any(str(item) == "B3:B4" for item in merged)
        assert original.active["B4"].value is None
    finally:
        original.close()


def test_save_creates_formulas_formats_and_sum(tmp_path: Path) -> None:
    source = tmp_path / "단가대비표.xlsx"
    _write_source(source, with_merge=True)
    original = source.read_bytes()
    dest_dir = tmp_path / "결과"

    dest = save_result_workbook(source, dest_dir=dest_dir)

    assert dest.exists()
    assert dest.name.startswith("단가대비_공량산출_결과_")
    assert source.read_bytes() == original

    result = load_workbook(dest, data_only=False)
    try:
        assert result.sheetnames == [UNIT_PRICE_SHEET_NAME, QUANTITY_SHEET_NAME]
        sheet1 = result[UNIT_PRICE_SHEET_NAME]
        assert sheet1["C1"].value == "품명"
        assert sheet1["B3"].value == "배관"
        fg = sheet1["C1"].fill.fgColor
        rgb = str(getattr(fg, "rgb", "")).upper()
        assert rgb.endswith("BDD7EE")
        assert sheet1["C1"].border.left.style == "thin"
        assert sheet1["F2"].number_format == NUMBER_FORMAT

        sheet2 = result[QUANTITY_SHEET_NAME]
        headers = [sheet2.cell(4, c).value for c in range(1, 12)]
        assert headers[1] == "명칭"
        assert headers[4] == "결정수량"
        assert headers[10] == "공량"

        assert sheet2["B5"].value == "HI-PVC 전선관"
        assert sheet2["E5"].value == "='단가대비표'!F2"
        assert sheet2["F5"].value == 0
        assert sheet2["G5"].value == calc_qty_formula(5)
        assert sheet2["J5"].value == 100
        assert sheet2["K5"].value == gongryang_formula(5)
        assert sheet2["E5"].number_format == NUMBER_FORMAT
        assert sheet2["G5"].number_format == NUMBER_FORMAT
        assert sheet2["I5"].number_format == NUMBER_FORMAT
        assert sheet2["K5"].number_format == NUMBER_FORMAT

        assert sheet2["B6"].value == "CD관"
        assert sheet2["E6"].value == "='단가대비표'!F3"
        assert sheet2["G6"].value == "=E6*(1+F6)"
        assert sheet2["K7"].value == gongryang_sum_formula(6)
        assert sheet2["A7"].value == "합계"
        assert sheet2.column_dimensions["B"].width >= 10
    finally:
        result.close()
