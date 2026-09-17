from __future__ import annotations

from pathlib import Path

from openpyxl import load_workbook

from app.wages import (
    WAGE_COL_WIDTHS,
    WAGE_NUMBER_FORMAT,
    WAGE_ROW_HEIGHT,
    default_wage_rows,
    write_wages_workbook,
)


def test_wages_workbook_layout(tmp_path: Path) -> None:
    path = write_wages_workbook(default_wage_rows(), tmp_path / "노임단가.xlsx")
    workbook = load_workbook(path)
    try:
        sheet = workbook.active
        assert sheet["A1"].value == "직종"
        assert sheet["B1"].value == "노임단가"
        assert sheet["C1"].value == "비고"
        assert sheet.column_dimensions["A"].width == WAGE_COL_WIDTHS[0]
        assert sheet.column_dimensions["B"].width == WAGE_COL_WIDTHS[1]
        assert sheet.column_dimensions["C"].width == WAGE_COL_WIDTHS[2]
        assert sheet.row_dimensions[1].height == WAGE_ROW_HEIGHT
        assert sheet.row_dimensions[2].height == WAGE_ROW_HEIGHT
        assert sheet["B2"].value == 276108
        assert sheet["B2"].number_format == WAGE_NUMBER_FORMAT
        assert sheet["A2"].font.name == "굴림"
    finally:
        workbook.close()
