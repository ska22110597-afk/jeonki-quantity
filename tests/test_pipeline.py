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
        assert dest.name.startswith("내역서_결과_")
        assert result.sheetnames[0] == COMPARE_SHEET_NAME
        assert ILWIDAE_SHEET_NAME in result.sheetnames
        assert ESTIMATE_SHEET_NAME in result.sheetnames
        assert QUANTITY_SHEET_NAME not in result.sheetnames
        assert PUMSAM_SHEET_NAME in result.sheetnames
        assert WAGES_SHEET_NAME in result.sheetnames

        compare = result[COMPARE_SHEET_NAME]
        assert compare["A1"].value == "단 가 대 비 표"
        assert "품" in str(compare["A3"].value or "").replace(" ", "")
        assert all(
            "코드" not in str(compare.cell(3, col).value or "").replace(" ", "")
            for col in range(1, 22)
        )
        assert compare["A5"].value == "경질비닐전선관_지중"
        assert compare["D5"].value == 200
        assert compare["L5"].value == 200
        assert compare.column_dimensions["E"].width == 5
        assert compare.column_dimensions["K"].width == 5
        assert compare["A6"].value == "경질비닐전선관_노출"
        name_merges = [
            str(range_)
            for range_ in compare.merged_cells.ranges
            if range_.min_col == 1 and range_.min_row >= 5
        ]
        assert name_merges == []

        ilwidae = result[ILWIDAE_SHEET_NAME]
        names = [ilwidae.cell(r, 1).value for r in range(1, 40)]
        assert "전선관부속품비" not in names
        assert "잡재료비" not in names
        assert "공구손료" not in names
        assert "내선전공" in names
        assert "보통인부" in names
        assert ilwidae["A1"].value == "일 위 대 가"
        assert ilwidae["A5"].alignment.horizontal == "left"
        assert ilwidae["E6"].value == "='단가대비표'!L5"
        assert ilwidae["G6"].value == "=0"
        assert "TRUNC" in str(ilwidae["H6"].value)
        assert "G6" in str(ilwidae["H6"].value)
        labor_row = next(r for r in range(5, 30) if ilwidae.cell(r, 1).value == "내선전공")
        assert ilwidae.cell(labor_row, 5).value == "=0"
        assert f"E{labor_row}" in str(ilwidae.cell(labor_row, 6).value)
        assert labor_row == 7
        sum_rows = [
            r
            for r in range(5, 25)
            if "합계" in str(ilwidae.cell(r, 1).value or "").replace(" ", "")
        ]
        assert sum_rows
        assert "SUM" in str(ilwidae.cell(sum_rows[0], 6).value)
        assert ilwidae["E6"].number_format == "#,##0.00"
        assert ilwidae["F6"].number_format == "#,##0.0"
        assert any("'노임단가'" in str(ilwidae.cell(r, 7).value or "") for r in range(5, 20))
        assert ilwidae.row_dimensions[1].height == 30

        estimate = result[ESTIMATE_SHEET_NAME]
        assert estimate["A1"].value == "[내역서 ]"
        assert estimate["A3"].value == "명칭"
        assert estimate["A5"].value == "1. 전기공사"
        assert estimate["A6"].value == "경질비닐전선관_지중"
        assert estimate["D6"].value == 100
        assert estimate["E6"].value in (None, "")
        assert "일위대가" in str(estimate["F6"].value)
        assert "F6" in str(estimate["F6"].value)
        assert estimate["G6"].value in (None, "")
        assert estimate["H6"].value in (None, "")
        assert "TRUNC" in str(estimate["L6"].value)
        assert estimate["L6"].number_format == "#,##0.0"
        assert estimate.row_dimensions[6].height == 30
        estimate_names = [estimate.cell(r, 1).value for r in range(1, 40)]
        assert any("배관" in str(value or "") and "부속" in str(value or "") for value in estimate_names)
        assert any("소모" in str(value or "") and "잡자" in str(value or "") for value in estimate_names)
        assert any("공구" in str(value or "").replace(" ", "") for value in estimate_names)
        assert any(str(value or "").replace(" ", "") == "(합계)" for value in estimate_names)
        sundry_row = next(
            r for r in range(5, 40) if "부속" in str(estimate.cell(r, 1).value or "") and "CD" in str(estimate.cell(r, 2).value or "")
        )
        assert "SUMPRODUCT" in str(estimate.cell(sundry_row, 6).value)
        assert "0.4" in str(estimate.cell(sundry_row, 6).value)
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
        compare = result[COMPARE_SHEET_NAME]
        assert all(
            "코드" not in str(compare.cell(3, col).value or "").replace(" ", "")
            for col in range(1, 22)
        )
        assert compare["A5"].value == "강제전선관"
        assert compare["A6"].value == "강제전선관"
        name_merges = [
            str(range_)
            for range_ in compare.merged_cells.ranges
            if range_.min_col == 1 and range_.min_row >= 5
        ]
        assert name_merges == []
        assert compare["E5"].value == 1202
        assert estimate["A6"].value == "강제전선관"
        assert estimate["D6"].value in (None, "")
        assert estimate["C6"].value == "M"
        assert estimate["E6"].value in (None, "")
        assert "일위대가" in str(estimate["F6"].value or "")
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
        assert dest.name.startswith("단가대비표_결과_")
        assert result.sheetnames[0] == COMPARE_SHEET_NAME
        assert ESTIMATE_SHEET_NAME in result.sheetnames
        assert QUANTITY_SHEET_NAME not in result.sheetnames
        compare = result[COMPARE_SHEET_NAME]
        assert compare["A1"].value == "단 가 대 비 표"
        assert "품" in str(compare["A3"].value or "").replace(" ", "")
        assert all(
            "코드" not in str(compare.cell(3, col).value or "").replace(" ", "")
            for col in range(1, 22)
        )
        assert compare["A5"].value == "강제전선관"
        assert compare["D5"].value == 3374
        assert compare["L5"].value == 3374
        assert compare.row_dimensions[5].height == 30
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
        remarks = [ilwidae.cell(r, 13).value for r in range(5, 40)]
        assert any(str(value or "").replace(" ", "") == "전기5-1" for value in remarks)
        assert QUANTITY_SHEET_NAME not in result.sheetnames
    finally:
        result.close()

    db = tmp_path / "out" / "데이터베이스"
    assert (db / pumsam_filename(TELECOM)).exists()
    assert not (db / "전기_표준품셈.xlsx").exists()
    loaded = load_pumsam_database(tmp_path / "out", TELECOM)
    jobs = {row.get("노무명칭") for row in loaded if "경질비닐전선관" in str(row.get("명칭") or "")}
    assert "통신내선공" in jobs
    assert "내선전공" not in jobs


def test_quantity_mode_keeps_estimate_parts(tmp_path: Path) -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "내역서"
    sheet["A1"] = "[내역서 ]"
    sheet["A3"] = "명칭"
    sheet["B3"] = "규격"
    sheet["C3"] = "단위"
    sheet["D3"] = "수량"
    sheet["E3"] = "재료비"
    sheet["E4"] = "단가"
    sheet.merge_cells("A3:A4")
    sheet.merge_cells("B3:B4")
    sheet.merge_cells("C3:C4")
    sheet.merge_cells("D3:D4")
    sheet.merge_cells("E3:F3")
    sheet["A5"] = "1. 전기공사"
    sheet["A6"] = "강제전선관"
    sheet["B6"] = "아연도 16 mm"
    sheet["C6"] = "M"
    sheet["A7"] = "1. 전열설비공사"
    sheet["A8"] = "경질비닐전선관"
    sheet["B8"] = "HI 16 mm"
    sheet["C8"] = "M"
    source = tmp_path / "내역서_파트.xlsx"
    workbook.save(source)
    workbook.close()

    dest = save_result_workbook(estimate_path=source, dest_dir=tmp_path / "out", mode="quantity")
    result = load_workbook(dest, data_only=False)
    try:
        assert dest.name.startswith("공량산출_결과_")
        assert COMPARE_SHEET_NAME not in result.sheetnames
        assert QUANTITY_SHEET_NAME in result.sheetnames
        qty = result[QUANTITY_SHEET_NAME]
        names = [qty.cell(r, 2).value for r in range(4, 10)]
        assert "1. 전기공사" in names
        assert "1. 전열설비공사" in names
        assert names.index("1. 전기공사") < names.index("강제전선관") < names.index("1. 전열설비공사")
        assert qty["B8"].value == "경질비닐전선관"
        assert "VLOOKUP" in str(qty["H8"].value)
        assert "SUMPRODUCT" in str(qty["K8"].value)
        assert qty["A3"].value == "품목"
    finally:
        result.close()


def test_compare_keeps_page_and_pps_prices(tmp_path: Path) -> None:
    sample = Path("/home/ubuntu/.cursor/projects/workspace/uploads/_________8473.xlsx")
    if not sample.exists():
        return
    dest = save_result_workbook(unit_price_path=sample, dest_dir=tmp_path / "out")
    result = load_workbook(dest, data_only=False)
    try:
        compare = result[COMPARE_SHEET_NAME]
        assert compare["A5"].value == "강제전선관"
        assert compare["D5"].value == 3374
        assert compare["E5"].value == 1202
        assert compare.column_dimensions["E"].width == 5
        assert compare.column_dimensions["G"].width == 5
        assert compare.column_dimensions["I"].width == 5
        assert compare.column_dimensions["K"].width == 5
        assert compare["L5"].value == 3374
        rows = {
            (compare.cell(r, 1).value, str(compare.cell(r, 2).value or "")): r
            for r in range(5, 40)
        }
        d30 = rows.get(("관로구방수장치", "D30"))
        assert d30
        assert compare.cell(d30, 4).value in (None, "")
        assert compare.cell(d30, 6).value == 14350
        assert compare.cell(d30, 12).value == 14350
        estimate = result[ESTIMATE_SHEET_NAME]
        assert estimate["E6"].value in (None, "")
        assert "일위대가" in str(estimate["F6"].value)
    finally:
        result.close()


def _merged_overlap(left, right) -> bool:
    return not (
        left.max_row < right.min_row
        or left.min_row > right.max_row
        or left.max_col < right.min_col
        or left.min_col > right.max_col
    )


def test_quantity_keeps_sheet_with_overlapping_source_merges(tmp_path: Path) -> None:
    from app.excel_io import _without_overlapping_merges
    from app.estimate_parse import EstimateSheet
    from app.pipeline import build_result_workbook
    from app.pumsam import default_pumsam_rows
    from app.wages import default_wage_rows

    kept = _without_overlapping_merges([(1, 1, 1, 13), (1, 1, 2, 1), (2, 1, 3, 1)])
    assert kept[0] == (1, 1, 1, 13)
    assert (1, 1, 2, 1) not in kept

    filled = [
        ["[내역서 ]"] + [None] * 12,
        ["명칭", "규격", "단위", "수량", "재료비"] + [None] * 8,
        [None, None, None, None, "단가", "금액"] + [None] * 7,
        ["1. 옥외전기공사"] + [None] * 12,
        ["경질비닐전선관_지중", "HI 16 mm", "M", 10, 200, "=D5*E5"] + [None] * 7,
    ]
    estimate = EstimateSheet(
        filled=filled,
        raw=filled,
        merges=[(1, 1, 1, 13), (1, 1, 2, 1), (2, 1, 3, 1), (2, 5, 2, 6)],
    )
    workbook = build_result_workbook(
        compare=None,
        estimate=estimate,
        pumsam_rows=default_pumsam_rows(),
        wage_rows=default_wage_rows(),
        mode="quantity",
    )
    dest = tmp_path / "out.xlsx"
    workbook.save(dest)
    workbook.close()
    result = load_workbook(dest, data_only=False)
    try:
        assert QUANTITY_SHEET_NAME in result.sheetnames
        assert ESTIMATE_SHEET_NAME in result.sheetnames
        qty = result[QUANTITY_SHEET_NAME]
        assert qty["B5"].value == "경질비닐전선관_지중"
        for sheet in result.worksheets:
            ranges = list(sheet.merged_cells.ranges)
            for i, left in enumerate(ranges):
                for right in ranges[i + 1 :]:
                    assert not _merged_overlap(left, right), f"{sheet.title}: {left} vs {right}"
    finally:
        result.close()


def test_quantity_strips_external_workbook_formulas(tmp_path: Path) -> None:
    sample = Path("/home/ubuntu/.cursor/projects/workspace/uploads/_______d25a.xlsx")
    if not sample.exists():
        return
    dest = save_result_workbook(estimate_path=sample, dest_dir=tmp_path / "out", mode="quantity")
    result = load_workbook(dest, data_only=False)
    try:
        estimate = result[ESTIMATE_SHEET_NAME]
        assert QUANTITY_SHEET_NAME in result.sheetnames
        for row in estimate.iter_rows():
            for cell in row:
                value = cell.value
                if isinstance(value, str) and value.startswith("="):
                    assert "[" not in value
                    assert "노임산출서" not in value
        assert estimate["F5"].value == "=D5*E5"
        assert estimate["D214"].value in (0, None)
    finally:
        result.close()


def test_quantity_skips_sundry_form_formulas(tmp_path: Path) -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "내역서"
    sheet["A1"] = "[내역서 ]"
    sheet["A3"] = "명칭"
    sheet["B3"] = "규격"
    sheet["C3"] = "단위"
    sheet["D3"] = "수량"
    sheet["E3"] = "재료비"
    sheet["E4"] = "단가"
    sheet["A5"] = "1. 전기공사"
    sheet["A6"] = "강제전선관"
    sheet["B6"] = "아연도 16 mm"
    sheet["C6"] = "M"
    sheet["D6"] = 10
    sheet["A8"] = "[ 배관 부속재 ]"
    sheet["B8"] = "전선관의 15 %"
    sheet["C8"] = "식"
    sheet["D8"] = 1
    sheet["A9"] = "노 무 비"
    sheet["B9"] = "내선전공"
    sheet["C9"] = "인"
    sheet["A10"] = "( 합 계 )"
    source = tmp_path / "내역서_양식.xlsx"
    workbook.save(source)
    workbook.close()

    dest = save_result_workbook(estimate_path=source, dest_dir=tmp_path / "out", mode="quantity")
    result = load_workbook(dest, data_only=False)
    try:
        qty = result[QUANTITY_SHEET_NAME]
        assert qty["B6"].value == "강제전선관"
        assert "VLOOKUP" in str(qty["H6"].value)
        assert qty["B8"].value == "[ 배관 부속재 ]"
        assert qty["C8"].value == "전선관의 15 %"
        assert qty["H8"].value in (None, "")
        assert qty["K8"].value in (None, "")
        assert qty["B9"].value == "노 무 비"
        assert qty["H9"].value in (None, "")
        assert qty["B10"].value == "( 합 계 )"
        assert qty["H10"].value in (None, "")
    finally:
        result.close()

