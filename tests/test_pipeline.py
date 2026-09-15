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
        assert qty["H5"].value == "내선전공"
        assert "전기" in str(qty["L5"].value or "")
        titles = [
            ilwidae.cell(r, 1).value
            for r in range(1, 40)
            if "호표" in str(ilwidae.cell(r, 1).value or "")
        ]
        assert titles
        assert all("전기" not in str(title) for title in titles)
        remarks = [ilwidae.cell(r, 13).value for r in range(5, 40)]
        assert any(str(value or "").replace(" ", "") == "전기5-1" for value in remarks)
        assert ilwidae["A5"].font.name == "굴림"
        assert ilwidae["A5"].font.bold is not True
    finally:
        result.close()

    db = tmp_path / "out" / "데이터베이스"
    assert (db / "전기_표준품셈.xlsx").exists()
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
    rows = {row["직종"]: row for row in default_wage_rows()}
    assert rows["내선전공"]["노임단가"] == 276108
    assert rows["보통인부"]["노임단가"] == 172698
    assert rows["저압케이블전공"]["노임단가"] == 306274
    assert "2026" in str(rows["내선전공"]["비고"])


def test_sample_unit_price_skips_header_and_empty_qty(tmp_path: Path) -> None:
    sample = Path("/home/ubuntu/.cursor/projects/workspace/uploads/_________797b.xlsx")
    if not sample.exists():
        return
    dest = save_result_workbook(unit_price_path=sample, dest_dir=tmp_path / "out")
    result = load_workbook(dest, data_only=False)
    try:
        ilwidae = result[ILWIDAE_SHEET_NAME]
        titles = [
            str(ilwidae.cell(r, 1).value)
            for r in range(1, 30)
            if "호표" in str(ilwidae.cell(r, 1).value or "")
        ]
        assert titles
        assert all("품 명" not in title and "품명" not in title.replace(" ", "") for title in titles)
        assert "강제전선관" in titles[0]
        assert all("전기" not in title for title in titles)
        estimate = result[ESTIMATE_SHEET_NAME]
        assert estimate["A5"].value == "강제전선관"
        assert estimate["D5"].value in (None, "")
        assert estimate["C5"].value == "M"
        qty = result[QUANTITY_SHEET_NAME]
        assert qty["B5"].value == "강제전선관"
        assert qty["H8"].value == "내선전공" or qty["H5"].value in (None, "", "내선전공")
        hi_rows = [
            r
            for r in range(4, 20)
            if qty.cell(r, 2).value == "경질비닐전선관" and str(qty.cell(r, 3).value or "").startswith("HI 16")
        ]
        assert hi_rows
        assert qty.cell(hi_rows[0], 8).value == "내선전공"
    finally:
        result.close()


def test_reverse_estimate_builds_unit_price(tmp_path: Path) -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "내역서"
    sheet["A1"] = "[내역서 ]"
    sheet["A2"] = "명칭"
    sheet["B2"] = "규격"
    sheet["C2"] = "단위"
    sheet["D2"] = "수량"
    sheet["E2"] = "재료비"
    sheet["E3"] = "단가"
    sheet.merge_cells("A2:A3")
    sheet.merge_cells("B2:B3")
    sheet.merge_cells("C2:C3")
    sheet.merge_cells("D2:D3")
    sheet.merge_cells("E2:F2")
    sheet["A4"] = "1. 전기공사"
    sheet["A5"] = "강제전선관"
    sheet["B5"] = "아연도 16 mm"
    sheet["C5"] = "M"
    sheet["D5"] = 80
    sheet["E5"] = 3374
    source = tmp_path / "내역서.xlsx"
    workbook.save(source)
    workbook.close()

    dest = save_result_workbook(estimate_path=source, dest_dir=tmp_path / "out")
    result = load_workbook(dest, data_only=False)
    try:
        assert result.sheetnames[0] == ESTIMATE_SHEET_NAME
        assert COMPARE_SHEET_NAME in result.sheetnames
        compare = result[COMPARE_SHEET_NAME]
        assert compare["A1"].value == "단 가 대 비 표"
        assert compare["B3"].value is not None
        assert "품" in str(compare["B3"].value).replace(" ", "")
        assert compare["A5"].value in (None, "")
        assert compare["B5"].value == "강제전선관"
        assert compare["E5"].value == 3374
        assert compare["M5"].value == 3374
        qty = result[QUANTITY_SHEET_NAME]
        assert qty["B5"].value == "강제전선관"
    finally:
        result.close()


def test_telecom_pumsam_does_not_pull_electric_labor(tmp_path: Path) -> None:
    from app.discipline import TELECOM
    from app.pumsam import load_pumsam_database, pumsam_filename

    electric = {row.get("노무명칭") for row in default_pumsam_rows("전기")}
    telecom = {row.get("노무명칭") for row in default_pumsam_rows(TELECOM)}
    assert "내선전공" in electric
    assert "내선전공" not in telecom
    assert "통신내선공" in telecom
    assert "통신케이블공" in telecom

    source = tmp_path / "단가대비표.xlsx"
    _write_compare(source)
    dest = save_result_workbook(
        unit_price_path=source,
        dest_dir=tmp_path / "out",
        discipline=TELECOM,
    )
    result = load_workbook(dest, data_only=False)
    try:
        ilwidae = result[ILWIDAE_SHEET_NAME]
        names = [ilwidae.cell(r, 1).value for r in range(1, 40)]
        assert "통신내선공" in names
        assert "내선전공" not in names
        qty = result[QUANTITY_SHEET_NAME]
        assert qty["H5"].value == "통신내선공"
        remarks = [ilwidae.cell(r, 13).value for r in range(5, 40)]
        assert any(str(value or "").replace(" ", "") == "전기5-1" for value in remarks)
    finally:
        result.close()

    db = tmp_path / "out" / "데이터베이스"
    assert (db / pumsam_filename(TELECOM)).exists()
    assert not (db / "전기_표준품셈.xlsx").exists()
    loaded = load_pumsam_database(tmp_path / "out", TELECOM)
    jobs = {row.get("노무명칭") for row in loaded if "경질비닐전선관" in str(row.get("명칭") or "")}
    assert "통신내선공" in jobs
    assert "내선전공" not in jobs
