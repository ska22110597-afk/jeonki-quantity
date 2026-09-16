from __future__ import annotations

from pathlib import Path

from openpyxl import load_workbook

from app.pumsam import (
    PUMSAM_COLUMN_WIDTHS,
    PUMSAM_DISPLAY_HEADERS,
    PUMSAM_ROW_HEIGHT,
    default_pumsam_rows,
    save_pumsam_database,
)


def _fill_rgb(cell) -> str:
    color = cell.fill.fgColor
    rgb = getattr(color, "rgb", None)
    return str(rgb).upper() if rgb is not None else ""


def test_pumsam_workbook_layout_electric_and_telecom(tmp_path: Path) -> None:
    for disc in ("전기", "통신"):
        rows = default_pumsam_rows(disc)[:40]
        path = save_pumsam_database(rows, directory=tmp_path, discipline=disc)
        workbook = load_workbook(path)
        try:
            sheet = workbook["품셈표"]
            assert [sheet.cell(1, col).value for col in range(1, 9)] == list(PUMSAM_DISPLAY_HEADERS)
            assert sheet["A1"].value == "키워드"
            assert sheet.row_dimensions[1].height == PUMSAM_ROW_HEIGHT
            assert sheet.row_dimensions[2].height == PUMSAM_ROW_HEIGHT
            assert sheet["A1"].font.name == "굴림"
            assert sheet["B2"].font.name == "굴림"
            assert sheet["A1"].alignment.horizontal == "center"
            assert sheet["B1"].alignment.horizontal == "center"
            assert "B7DEE8" in _fill_rgb(sheet["A1"])
            assert "FFFFFF" in _fill_rgb(sheet["A2"]) or _fill_rgb(sheet["A2"]).endswith("FFFFFF")
            assert sheet["B2"].alignment.horizontal == "left"
            assert sheet["C2"].alignment.horizontal == "left"
            assert sheet["D2"].alignment.horizontal == "center"
            assert sheet["E2"].alignment.horizontal == "center"
            assert sheet["F2"].alignment.horizontal == "right"
            assert sheet["G2"].alignment.horizontal == "center"
            assert sheet["H2"].alignment.horizontal == "center"
            letters = "ABCDEFGH"
            for letter, width in zip(letters, PUMSAM_COLUMN_WIDTHS):
                assert sheet.column_dimensions[letter].width == width
            assert "적용기준" in workbook.sheetnames
            assert "공종별인부" in workbook.sheetnames
        finally:
            workbook.close()
