from __future__ import annotations

from pathlib import Path

from openpyxl import Workbook, load_workbook

from app.excel_io import (
    QUANTITY_SHEET_NAME,
    UNIT_PRICE_SHEET_NAME,
    read_unit_price_table,
    save_result_workbook,
)


def _write_source(path: Path) -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "원본시트"
    sheet.append(["번호", "공종", "품명", "규격", "단위", "설계단가"])
    sheet.append([1, "배관", "HI-PVC 전선관", "16mm", "m", 1200])
    sheet.append([2, "배선", "IV전선", "2.5㎟", "m", 850])
    sheet.append([None, None, None, None, None, None])
    workbook.save(path)
    workbook.close()


def test_read_only_trims_empty_rows(tmp_path: Path) -> None:
    source = tmp_path / "단가대비표.xlsx"
    _write_source(source)
    before = source.read_bytes()

    rows = read_unit_price_table(source)

    assert source.read_bytes() == before
    assert rows[0][1] == "공종"
    assert len(rows) == 3
    assert rows[-1][2] == "IV전선"


def test_save_creates_timestamped_two_sheet_file(tmp_path: Path) -> None:
    source = tmp_path / "단가대비표.xlsx"
    _write_source(source)
    original = source.read_bytes()
    dest_dir = tmp_path / "결과"

    dest = save_result_workbook(source, dest_dir=dest_dir)

    assert dest.exists()
    assert dest.parent == dest_dir
    assert dest.name.startswith("단가대비_공량산출_결과_")
    assert dest.suffix == ".xlsx"
    assert dest.resolve() != source.resolve()
    assert source.read_bytes() == original

    result = load_workbook(dest, read_only=True, data_only=True)
    try:
        assert result.sheetnames == [UNIT_PRICE_SHEET_NAME, QUANTITY_SHEET_NAME]
        sheet1 = result[UNIT_PRICE_SHEET_NAME]
        values = [row[2] for row in sheet1.iter_rows(min_row=2, max_col=3, values_only=True)]
        assert "HI-PVC 전선관" in values
        sheet2 = result[QUANTITY_SHEET_NAME]
        headers = next(sheet2.iter_rows(min_row=1, max_row=1, values_only=True))
        assert headers[0] == "번호"
        assert "산출근거" in headers
        first_qty = next(sheet2.iter_rows(min_row=2, max_row=2, values_only=True))
        assert first_qty[2] == "HI-PVC 전선관"
        assert first_qty[5] is None
    finally:
        result.close()
