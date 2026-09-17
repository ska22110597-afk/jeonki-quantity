from __future__ import annotations

from app.electric_pumsam_data import ELECTRIC_TRADE_GUIDE, _cd_qty, electric_pumsam_rows
from app.pumsam import (
    default_pumsam_rows,
    has_pumsam_qty,
    match_pumsam,
    pumsam_rate_value,
    pumsam_surcharge_note,
)
from app.ilwidae import labor_qty_formula
from app.items import LineItem


def test_electric_pumsam_covers_places_and_trades() -> None:
    rows = electric_pumsam_rows()
    assert len(rows) > 8000
    names = {str(row.get("명칭")) for row in rows}
    jobs = {str(row.get("노무명칭")) for row in rows}
    assert "경질비닐전선관" in names
    assert "경질비닐전선관_매입" in names
    assert "경질비닐전선관_노출" in names
    assert "경질비닐전선관_지중" in names
    assert "배선용단자함" in names
    assert "HIV전선" in names
    assert "CV케이블" in names
    assert "내선전공" in jobs
    assert "보통인부" in jobs
    assert "저압케이블전공" in jobs
    assert "배전전공" in jobs
    assert "변전전공" in jobs
    assert "송전전공" in jobs
    assert "계장공" in jobs
    assert any("콘크리트전주" in n for n in names)
    assert any("변압기" in n for n in names)
    assert any(str(row.get("품셈근거") or "").startswith("전기2-") for row in rows)
    assert any(str(row.get("품셈근거") or "").startswith("전기3-") for row in rows)
    assert any(str(row.get("품셈근거") or "").startswith("전기4-") for row in rows)
    assert any(str(row.get("품셈근거") or "").startswith("전기6-") for row in rows)
    assert any(str(row.get("품셈근거") or "").startswith("전기7-") for row in rows)
    assert all(row[0] for row in ELECTRIC_TRADE_GUIDE)


def test_distribution_pole_and_transformer_labors() -> None:
    rows = default_pumsam_rows()
    pole = match_pumsam("콘크리트전주", "8 m 이하", rows)
    jobs = {row.get("노무명칭") for row in pole}
    assert "배전전공" in jobs
    assert "보통인부" in jobs
    tr = match_pumsam("22 kV 변압기 설치", "100 kVA 이하", rows)
    tr_jobs = {row.get("노무명칭") for row in tr}
    assert "변전전공" in tr_jobs
    assert "특별인부" in tr_jobs


def test_cd_is_80_percent_of_resin_under_100mm() -> None:
    assert _cd_qty(16) == 0.04
    assert _cd_qty(104) == 0.46
    matched = match_pumsam("CD전선관", "CD 16 mm", default_pumsam_rows())
    assert len(matched) == 1
    assert matched[0]["품셈"] == 0.04
    assert matched[0]["노무명칭"] == "내선전공"


def test_direct_burial_cable_uses_80_percent() -> None:
    rows = default_pumsam_rows()
    buried = match_pumsam("CV케이블_직매", "60 mm2×1C", rows)
    open_run = match_pumsam("CV케이블", "60 mm2×1C", rows)
    assert len(buried) == 1
    assert len(open_run) == 1
    assert buried[0]["할증%"] == 80
    assert open_run[0]["할증%"] == 100
    assert buried[0]["노무명칭"] == "저압케이블전공"
    assert pumsam_rate_value(buried[0]) == 0.8


def test_exposed_conduit_keeps_base_qty_and_120_rate() -> None:
    item = LineItem(excel_row=5, name="경질비닐전선관_노출", spec="HI 104 mm", unit="M", qty=1, material_price=1)
    matched = match_pumsam(item.name, item.spec, default_pumsam_rows())
    assert [row.get("노무명칭") for row in matched] == ["내선전공"]
    assert matched[0]["품셈"] == 0.46
    assert matched[0]["할증%"] == 120
    assert labor_qty_formula(matched[0]) == "=0.46*1.2"


def test_square_mm_lookup_matches_mm2_and_mm2_display_is_unified() -> None:
    rows = default_pumsam_rows()
    by_mark = match_pumsam("HIV전선", "14 ㎟", rows)
    by_ascii = match_pumsam("HIV전선", "14 mm2", rows)
    assert by_mark
    assert by_ascii
    assert {row["노무명칭"] for row in by_mark} == {row["노무명칭"] for row in by_ascii}
    assert all("mm2" not in str(row.get("규격") or "").lower() for row in by_mark)


def test_ceiling_이하_uses_smallest_band_and_keeps_all_labors() -> None:
    rows = default_pumsam_rows()
    wire = match_pumsam("HIV전선", "4 ㎟", rows)
    assert len(wire) == 1
    assert wire[0]["노무명칭"] == "내선전공"
    assert "6" in str(wire[0]["규격"])
    assert "이하" in str(wire[0]["규격"])
    assert "㎟" in str(wire[0]["규격"])
    assert wire[0]["품셈"] == 0.010

    exact = match_pumsam("HIV전선", "14 ㎟", rows)
    assert exact
    assert "이하" not in str(exact[0]["규격"])
    assert "14" in str(exact[0]["규격"])

    pole = match_pumsam("콘크리트전주", "7 m", rows)
    jobs = {row.get("노무명칭") for row in pole}
    assert "배전전공" in jobs
    assert "보통인부" in jobs
    assert all("8" in str(row.get("규격") or "") and "이하" in str(row.get("규격") or "") for row in pole)

    power = match_pumsam("CV케이블", "14 mm2×1C", rows)
    assert power
    assert power[0]["노무명칭"] == "저압케이블전공"
    assert "1C" in str(power[0]["규격"]).replace(" ", "").upper()
    assert "3C" not in str(power[0]["규격"]).replace(" ", "").upper()
    buried = match_pumsam("CV케이블_직매", "14 mm2×1C", rows)
    assert buried[0]["할증%"] == 80
    assert power[0]["할증%"] == 100

    hfix = match_pumsam("HFIX전선", "4 ㎟", rows)
    assert hfix
    assert hfix[0]["노무명칭"] == "내선전공"
    assert "6" in str(hfix[0]["규격"])
    assert "이하" in str(hfix[0]["규격"])


def test_galvanized_spec_matches_plain_or_g_size() -> None:
    """단가대비표 아연도 16 mm, 품셈 16 mm / G 16 mm 처럼 앞말만 달라도 맞춘다."""
    rows = [
        {"명칭": "강제전선관", "규격": "16 mm", "노무명칭": "내선전공", "품셈": 0.08, "할증%": 100, "품셈근거": "전기5-1"},
        {"명칭": "강제전선관", "규격": "G 16 mm", "노무명칭": "내선전공", "품셈": 0.08, "할증%": 100, "품셈근거": "전기5-1"},
        {"명칭": "강제전선관", "규격": "22 mm", "노무명칭": "내선전공", "품셈": 0.11, "할증%": 100, "품셈근거": "전기5-1"},
        {"명칭": "가요전선관", "규격": "16 mm", "노무명칭": "내선전공", "품셈": 0.044, "할증%": 100, "품셈근거": "전기5-1"},
    ]
    matched = match_pumsam("강제전선관", "아연도 16 mm", rows)
    assert len(matched) == 1
    assert matched[0]["품셈"] == 0.08
    assert matched[0]["노무명칭"] == "내선전공"
    assert "22" not in str(matched[0]["규격"])
    other = match_pumsam("강제전선관", "G-16 mm", rows)
    assert other[0]["품셈"] == 0.08
    flex = match_pumsam("강제전선관", "아연도 16 mm", rows)
    assert flex[0]["명칭"] == "강제전선관"
    seed = match_pumsam("강제전선관", "아연도 16 mm", default_pumsam_rows())
    assert seed
    assert seed[0]["품셈"] == 0.08
    assert seed[0]["노무명칭"] == "내선전공"


def test_name_aliases_treat_steel_and_resin_as_same_family() -> None:
    """후강전선관 = 강제전선관, HI관 = 경질비닐전선관. 지중/노출은 그대로 구분."""
    steel_only = [
        {"명칭": "강제전선관", "규격": "16 mm", "노무명칭": "내선전공", "품셈": 0.08, "할증%": 100, "품셈근거": "전기5-1"},
        {"명칭": "강제전선관_지중", "규격": "16 mm", "노무명칭": "내선전공", "품셈": 0.056, "할증%": 70, "품셈근거": "전기5-1"},
        {"명칭": "가요전선관", "규격": "16 mm", "노무명칭": "내선전공", "품셈": 0.044, "할증%": 100, "품셈근거": "전기5-1"},
    ]
    by_alias = match_pumsam("후강전선관", "아연도 16 mm", steel_only)
    assert len(by_alias) == 1
    assert by_alias[0]["명칭"] == "강제전선관"
    assert by_alias[0]["품셈"] == 0.08
    buried = match_pumsam("후강전선관_지중", "16 mm", steel_only)
    assert buried[0]["품셈"] == 0.056
    assert buried[0]["할증%"] == 70
    assert match_pumsam("후강전선관_지중", "16 mm", steel_only[:1]) == []

    resin_only = [
        {"명칭": "경질비닐전선관", "규격": "HI 28 mm", "노무명칭": "내선전공", "품셈": 0.08, "할증%": 100},
        {"명칭": "경질비닐전선관_노출", "규격": "HI 28 mm", "노무명칭": "내선전공", "품셈": 0.08, "할증%": 120},
    ]
    hi = match_pumsam("HI관", "28 mm", resin_only)
    assert hi[0]["품셈"] == 0.08
    assert hi[0]["할증%"] == 100
    exposed = match_pumsam("합성수지전선관_노출", "HI 28 mm", resin_only)
    assert exposed[0]["할증%"] == 120

    hiv_only = [
        {"명칭": "HIV전선", "규격": "6 ㎟ 이하", "노무명칭": "내선전공", "품셈": 0.010, "할증%": 100},
        {"명칭": "CV케이블", "규격": "16 ㎟ 이하×1C", "노무명칭": "저압케이블전공", "품셈": 0.023, "할증%": 100},
    ]
    hfix = match_pumsam("HFIX전선", "4 ㎟", hiv_only)
    assert hfix[0]["품셈"] == 0.010
    power = match_pumsam("전력케이블", "14 ㎟×1C", hiv_only)
    assert power[0]["노무명칭"] == "저압케이블전공"
    assert power[0]["품셈"] == 0.023
    assert match_pumsam("후강전선관", "16 mm", resin_only) == []


def test_pe_cd_paren_place_and_tray_size_match() -> None:
    rows = default_pumsam_rows()
    buried = match_pumsam("폴리에틸렌 전선관(지중)", "PE 16mm", rows)
    assert buried
    assert buried[0]["노무명칭"] == "내선전공"
    assert buried[0]["품셈"] == 0.04
    assert buried[0]["할증%"] == 70

    cd_flex = match_pumsam("합성수지제 가요전선관", "CD-난연성 16mm", rows)
    assert cd_flex
    assert cd_flex[0]["노무명칭"] == "내선전공"
    assert cd_flex[0]["품셈"] == 0.04
    assert cd_flex[0]["할증%"] == 100
    metal = match_pumsam("가요전선관", "16 mm", rows)
    assert metal[0]["품셈"] == 0.044

    tray = match_pumsam("Hi-Tec Tray부속(분체도장)", "JOINER SET, W300 × H100", rows)
    assert tray
    assert tray[0]["노무명칭"] == "내선전공"
    assert tray[0]["품셈"] == 0.23
    assert "30000" in str(tray[0]["규격"]).replace(",", "")
    angle = match_pumsam("Hi-Tec Tray부속(분체도장)", "BEARING ANGLE, L370", rows)
    assert angle
    assert angle[0]["품셈"] == 0.18


def test_ceiling_3mm2_and_6mm2_bands_pick_nearest_이상_이하() -> None:
    rows = [
        {"명칭": "시험전선", "규격": "3 ㎟ 이하", "노무명칭": "내선전공", "품셈": 0.01, "할증%": 100},
        {"명칭": "시험전선", "규격": "3 ㎟ 이하", "노무명칭": "보통인부", "품셈": 0.02, "할증%": 100},
        {"명칭": "시험전선", "규격": "6 ㎟ 이하", "노무명칭": "내선전공", "품셈": 0.03, "할증%": 100},
        {"명칭": "시험전선", "규격": "6 ㎟ 이하", "노무명칭": "보통인부", "품셈": 0.04, "할증%": 100},
    ]
    four = match_pumsam("시험전선", "4 ㎟", rows)
    assert {row.get("노무명칭") for row in four} == {"내선전공", "보통인부"}
    assert all("6" in str(row.get("규격") or "") and "이하" in str(row.get("규격") or "") for row in four)
    three = match_pumsam("시험전선", "3 ㎟", rows)
    assert {row.get("노무명칭") for row in three} == {"내선전공", "보통인부"}
    assert all(str(row.get("규격") or "").startswith("3") and "이하" in str(row.get("규격") or "") for row in three)
    assert all("6" not in str(row.get("규격") or "") for row in three)


def test_pumsam_qty_is_raw_and_surcharge_is_separate() -> None:
    rows = default_pumsam_rows()
    exposed = match_pumsam("경질비닐전선관_노출", "HI 104 mm", rows)
    assert exposed[0]["품셈"] == 0.46
    assert exposed[0]["할증%"] == 120
    assert pumsam_rate_value(exposed[0]) == 1.2
    assert labor_qty_formula(exposed[0]) == "=0.46*1.2"
    assert pumsam_surcharge_note(exposed[0]) == "품셈 0.460 · 할증 120%"


def test_horizontal_conductor_is_single_book_item() -> None:
    rows = default_pumsam_rows()
    matched = match_pumsam("수평도체", "일반", rows)
    jobs = {row.get("노무명칭"): row.get("품셈") for row in matched}
    assert jobs["내선전공"] == 0.017
    assert jobs["보통인부"] == 0.008
    names = {
        row.get("명칭")
        for row in rows
        if "수평도체" in str(row.get("명칭") or "") or "수평도체" in str(row.get("짧은명칭") or "")
    }
    assert names == {"수평도체 설치"}
    specs = {
        row.get("규격")
        for row in rows
        if row.get("명칭") == "수평도체 설치" or row.get("짧은명칭") == "수평도체"
    }
    assert specs <= {"", None}


def test_horizontal_conductor_keyword_is_single() -> None:
    from app.estimate_parse import display_keyword

    rows = default_pumsam_rows()
    keys = {
        display_keyword(row.get("명칭"), row.get("규격"))
        for row in rows
        if "수평도체" in str(row.get("명칭") or "")
    }
    assert keys == {"수평도체설치"}


def test_generator_spec_is_kva_not_qty() -> None:
    rows = default_pumsam_rows()
    matched = match_pumsam("자가발전기", "20 kVA", rows)
    jobs = {row.get("노무명칭"): row.get("품셈") for row in matched}
    assert jobs["전기공사기사"] == 10.5
    assert jobs["플랜트전공"] == 6.3
    assert jobs["기계설비공"] == 6.3
    assert jobs["특별인부"] == 5.3
    gen_specs = {
        str(row.get("규격") or "")
        for row in rows
        if row.get("명칭") == "자가발전기 설치"
    }
    assert any("20" in spec and "kVA" in spec for spec in gen_specs)
    assert not any(spec.replace(".", "", 1).isdigit() for spec in gen_specs if spec)
    commission = match_pumsam("자가발전기 시운전 및 조정", "20 kVA", rows)
    comm_jobs = {row.get("노무명칭"): row.get("품셈") for row in commission}
    assert comm_jobs["전기공사기사"] == 3.2
    assert comm_jobs["플랜트전공"] == 3.2


def test_match_pumsam_skips_dash_labor() -> None:
    rows = [
        {"명칭": "지중 케이블", "규격": "OF 400 ㎟ 이하", "노무명칭": "전기공사기사", "품셈": 3.49, "할증%": 100},
        {"명칭": "지중 케이블", "규격": "OF 400 ㎟ 이하", "노무명칭": "특별인부", "품셈": None, "할증%": 100},
        {"명칭": "지중 케이블", "규격": "OF 400 ㎟ 이하", "노무명칭": "특고압케이블전공", "품셈": "-", "할증%": 100},
    ]
    matched = match_pumsam("지중 케이블", "OF 400 ㎟ 이하", rows)
    assert [row.get("노무명칭") for row in matched] == ["전기공사기사"]
    assert has_pumsam_qty(rows[0]) is True
    assert has_pumsam_qty(rows[1]) is False
    assert has_pumsam_qty(rows[2]) is False


def test_ditto_and_dash_are_interpreted_from_book() -> None:
    rows = default_pumsam_rows()
    buried = match_pumsam("지중 케이블 인력 설치", "154 kV OF 케이블 1200 ㎟ 이하", rows)
    jobs = {row.get("노무명칭"): row for row in buried}
    assert jobs["전기공사기사"]["품셈"] == 4.21
    assert jobs["특고압케이블전공"]["품셈"] == 78.75
    assert jobs["특별인부"]["품셈"] == 81.13
    assert all("〃" not in str(row.get("규격")) for row in buried)
    oil = match_pumsam("OF 케이블 급유장치 설치", "급유관 설치", rows)
    assert oil
    assert {row.get("노무명칭") for row in oil} == {"특고압케이블전공"}
    assert oil[0]["단위"] == "m"
    assert oil[0]["품셈"] == 0.19
    assert not str(oil[0]["규격"]).endswith("-")


def test_telecom_book_and_conduit_aliases() -> None:
    rows = default_pumsam_rows("통신")
    assert len(rows) > 7000
    jobs = {str(row.get("노무명칭")) for row in rows}
    assert "통신내선공" in jobs
    assert "통신케이블공" in jobs
    names = {str(row.get("명칭")) for row in rows}
    assert "합성수지 전선관" in names or "경질비닐전선관" in names
    buried = match_pumsam("경질비닐전선관_지중", "HI 16 mm", rows)
    assert buried
    assert buried[0]["노무명칭"] == "통신내선공"
    assert buried[0]["할증%"] == 70
    utp = match_pumsam("UTP케이블", "CAT.6", rows)
    assert utp
    assert utp[0]["노무명칭"] == "통신내선공"
