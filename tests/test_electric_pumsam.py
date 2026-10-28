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
    assert "CD전선관" in names
    assert "경질비닐전선관_매입" not in names
    assert "경질비닐전선관_노출" not in names
    assert "경질비닐전선관_지중" not in names
    assert "CD관_노출" not in names
    assert "세대분전반_노출" in names
    assert "LED등기구_매입" in names
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
    assert any(str(row.get("품셈근거") or "").startswith("전기10-") for row in rows)
    refs = [str(row.get("품셈근거") or "") for row in rows if row.get("품셈근거")]
    first_ch10 = next(i for i, ref in enumerate(refs) if ref.startswith("전기10-"))
    first_ch2 = next(i for i, ref in enumerate(refs) if ref.startswith("전기2-"))
    assert first_ch2 < first_ch10
    none_qty = [row for row in rows if row.get("품셈") in (None, "")]
    assert len(none_qty) > 400
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
    exposed = match_pumsam("CD전선관_노출", "CD 16 mm", default_pumsam_rows())
    assert exposed[0]["품셈"] == 0.04
    assert exposed[0]["할증%"] == 120


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
    from_base = match_pumsam("후강전선관_지중", "16 mm", steel_only[:1])
    assert len(from_base) == 1
    assert from_base[0]["품셈"] == 0.08
    assert from_base[0]["할증%"] == 70

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
    # 규격(폭×높이)이 없는 트레이 지지·부속품은 표준품셈 5-8 해설상 트레이 설치품에 포함 → 별도 품 없음
    assert not angle


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


def test_household_panel_exposed_keeps_book_ninety_percent() -> None:
    rows = default_pumsam_rows()
    names = {str(row.get("명칭")) for row in rows}
    assert "세대분전반_노출" in names
    assert "세대분전반_매입" not in names
    exposed = match_pumsam("세대분전반_노출", "3회로", rows)
    assert exposed
    assert exposed[0]["품셈"] == 0.59
    assert exposed[0]["할증%"] == 90


def test_merge_drops_place_clones_but_keeps_distinct_book_rows() -> None:
    from app.pumsam import merge_pumsam_rows

    merged = merge_pumsam_rows(
        [
            {"명칭": "경질비닐전선관", "규격": "HI 16 mm", "노무명칭": "내선전공", "품셈": 0.05, "할증%": 100},
            {"명칭": "경질비닐전선관_노출", "규격": "HI 16 mm", "노무명칭": "내선전공", "품셈": 0.05, "할증%": 120},
            {"명칭": "경질비닐전선관_매입", "규격": "HI 16 mm", "노무명칭": "내선전공", "품셈": 0.05, "할증%": 100},
            {"명칭": "세대분전반", "규격": "3회로", "노무명칭": "내선전공", "품셈": 0.59, "할증%": 100},
            {"명칭": "세대분전반_노출", "규격": "3회로", "노무명칭": "내선전공", "품셈": 0.59, "할증%": 90},
        ]
    )
    names = {row.get("명칭") for row in merged}
    assert names == {"경질비닐전선관", "세대분전반", "세대분전반_노출"}


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
    assert keys == {"수평도체 설치"}


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
    assert "경질비닐전선관_노출" not in names
    assert "경질비닐전선관_지중" not in names
    buried = match_pumsam("경질비닐전선관_지중", "HI 16 mm", rows)
    assert buried
    assert buried[0]["노무명칭"] == "통신내선공"
    assert buried[0]["할증%"] == 70
    # 통신 4-3-1: 구내 UTP 4P 0.15 /10m, 통신케이블공 → 1m 당 0.015
    utp = match_pumsam("UTP 케이블", "Cat.6 0.5mm 4P", rows)
    assert [(row["노무명칭"], row["품셈"]) for row in utp] == [("통신케이블공", 0.015)]
    fiber = match_pumsam("광케이블(옥내용)", "SM 12c", rows)
    assert {(row["노무명칭"], row["품셈"]) for row in fiber} == {("광케이블설치사", 0.0092), ("특별인부", 0.0046)}
    jack = match_pumsam("모듈라짹", "매입형 8P 2구", rows)
    assert [(row["노무명칭"], row["품셈"], row["할증%"]) for row in jack] == [("통신내선공", 0.028, 120)]
    none_qty = [row for row in rows if row.get("품셈") in (None, "")]
    assert len(none_qty) > 800


def test_electric_5472_vehicle_names_are_split() -> None:
    """5-47-2 교통신호 정기점검은 장치별로 명칭을 나눈다.

    원표는 차량자동인식장치·차량검지시스템·전자 교통신호제어기·가변정보표지판이다.
    """
    rows = default_pumsam_rows()
    section = [row for row in rows if row.get("품셈근거") == "전기5-47-2"]
    names = {str(row.get("명칭")) for row in section}
    assert "교통신호 시스템 정기점검" not in names
    assert names == {
        "차량자동인식장치 정기점검",
        "차량검지시스템 정기점검",
        "전자 교통신호제어기 정기점검",
        "가변정보표지판 정기점검",
    }

    avi = match_pumsam("차량자동인식장치", "카메라부 조명장치", rows)
    assert [(row.get("노무명칭"), row.get("품셈")) for row in avi] == [("내선전공", 0.19)]
    assert avi[0]["품셈근거"] == "전기5-47-2"

    controller = match_pumsam("차량자동인식장치 정기점검", "카메라부 카메라 컨트롤러", rows)
    jobs = {row.get("노무명칭"): row.get("품셈") for row in controller}
    assert jobs == {"내선전공": 0.17, "전기공사산업기사": 0.17}

    avi_main = match_pumsam("차량자동인식장치", "서브랙 메인 컨트롤러", rows)
    avi_jobs = {row.get("노무명칭"): row.get("품셈") for row in avi_main}
    assert avi_jobs == {"내선전공": 0.04, "전기공사산업기사": 0.25}

    vds_main = match_pumsam("차량검지시스템", "서브랙 메인 컨트롤러", rows)
    vds_jobs = {row.get("노무명칭"): row.get("품셈") for row in vds_main}
    assert vds_jobs == {"내선전공": 0.04, "전기공사산업기사": 0.31}

    signal = match_pumsam("전자 교통신호제어기", "구동부 신호구동기(LSU)", rows)
    signal_jobs = {row.get("노무명칭"): row.get("품셈") for row in signal}
    assert signal_jobs == {"내선전공": 0.06, "전기공사산업기사": 0.13}

    vms = match_pumsam("가변정보표지판", "LED 출력 모듈 2단 10열", rows)
    vms_jobs = {row.get("노무명칭"): row.get("품셈") for row in vms}
    assert vms_jobs == {"내선전공": 0.08, "전기공사산업기사": 0.13}


def test_electric_wire_and_box_adapter() -> None:
    from app.pumsam import load_pumsam_database, match_pumsam

    rows = load_pumsam_database(None, "전기")

    def first(name, spec):
        found = match_pumsam(name, spec, rows)
        return [(r["노무명칭"], r["품셈"]) for r in found]

    assert first("접지용비닐절연전선(F-GV)", "16㎟") == [("내선전공", 0.023)]
    assert first("450/750V 내열비닐절연전선", "HFIX 1.78mm(2.5㎟)") == [("내선전공", 0.01)]
    assert first("난연제어케이블", "F-CVV-S, 2Cx1.5㎟") == [("저압케이블전공", 0.014)]
    assert first("아우트렛박스", "8각 54㎜") == [("내선전공", 0.2)]
    assert first("스위치박스", "1 개용 54 mm") == [("내선전공", 0.2)]
    assert first("아우트렛박스 커버", "커버, 8각, 둥근구멍(평)") == []


def test_workspace_dir_and_qty_format(tmp_path) -> None:
    from app.estimate_build import qty_format
    from app.paths import workspace_dir

    assert workspace_dir(tmp_path) == tmp_path / "자동내역서"
    assert workspace_dir(tmp_path / "자동내역서") == tmp_path / "자동내역서"
    assert qty_format(1) == "#,##0" and qty_format(2.0) == "#,##0"
    assert qty_format(1.5) == "#,##0.0##"


def test_multicore_power_cable_uses_5_11_core_rate():
    from app.pumsam import load_pumsam_database, match_pumsam

    rows = load_pumsam_database(None, "전기")
    hit = match_pumsam("F-CV", "3Cx16㎟", rows)
    assert hit and hit[0]["할증%"] == 200 and abs(hit[0]["품셈"] - 0.023) < 1e-9
    small = match_pumsam("F-CV", "2Cx4㎟", rows)
    assert small and abs(small[0]["품셈"] - 0.016) < 1e-9  # 10㎟ 이하는 제어용케이블 준용
    assert match_pumsam("F-CV", "5Cx16㎟", rows) == []


def test_2026_supplement_rows_present():
    keys = {(r["명칭"], r["규격"]) for r in electric_pumsam_rows()}
    assert ("LED가로등기구", "150 W 이하") in keys
    assert ("커버나이프스위치", "100 A") in keys


def test_misc_unmatched_items_now_match() -> None:
    from app.pumsam import load_pumsam_database, match_pumsam

    rows = load_pumsam_database(None, "전기")

    def first(name, spec):
        m = match_pumsam(name, spec, rows)
        assert m, (name, spec)
        return m[0]

    rows30 = match_pumsam("파상형 폴리에틸렌 전선관", "30㎜(가로등,신호,구내)", rows)
    assert {r["노무명칭"] for r in rows30} == {"배전전공", "보통인부"}
    assert all(r["품셈근거"].startswith("전기4-31") and r["할증%"] == 150 for r in rows30)  # 4-31 해설⑧ 50% 가산
    assert {r["품셈"] for r in rows30} == {0.006, 0.014}
    assert match_pumsam("파상형 폴리에틸렌 전선관", "65㎜", rows)[0]["규격"] == "80 mm 이하"
    assert first("매입스위치", "15A, 250V 3로1구")["명칭"] == "텀블러스위치 3로용"
    two = first("매입스위치", "15A, 250V 1로2구")
    assert two["할증%"] == 200
    assert first("콘센트", "매입-접지형, 15A 250V 1구")["명칭"] == "콘센트 접지극부 15 A"
    assert first("러그단자(IEC 90)", "동관단자 2 HOLE 240 ㎟")["품셈근거"].startswith("전기4-37")
    assert first("직선접속제", "95㎟")["품셈근거"].startswith("전기4-36")
    assert first("가로등철주", "11m POLE, 2.8m ARM(분체도장)")["규격"] == "12 m 이하 1등용 인력"
    assert first("무정전전원장치(UPS)", "3상 20kVA")["품셈근거"].startswith("전기5-45")
    assert not match_pumsam("분전반", "L-1(휴게소)세종방향", rows)  # 용량 규격이 없으면 고르지 않는다


def test_tunnel_through_connector_uses_5_11_1() -> None:
    from app.pumsam import load_pumsam_database, match_pumsam

    rows = load_pumsam_database(None, "전기")
    m = match_pumsam("분기접속제(관통형/커버)", "95㎟", rows)
    assert m and m[0]["품셈"] == 0.036 and m[0]["할증%"] == 35 and m[0]["품셈근거"].startswith("전기5-11-1")


def test_duct_laying_rows_and_street_light_surcharge() -> None:
    import pytest
    from app.pumsam import load_pumsam_database, match_pumsam

    rows = load_pumsam_database(None, "전기")
    steel = {r["노무명칭"]: r for r in match_pumsam("강관 포설", "Φ200x2열", rows)}
    assert steel["배전전공"]["품셈"] == 0.061 and steel["배전전공"]["할증%"] == 180
    lamp = {r["노무명칭"]: r for r in match_pumsam("강관 포설(가로등)", "Φ200x2열", rows)}
    assert lamp["보통인부"]["품셈"] == 0.122 and lamp["보통인부"]["할증%"] == pytest.approx(180 * 1.3)
    corr = {r["노무명칭"]: r for r in match_pumsam("전선관 포설(가로등)", "Φ30x1열", rows)}
    assert corr["배전전공"]["품셈"] == 0.006 and corr["배전전공"]["할증%"] == pytest.approx(150)
    combo = {r["노무명칭"]: r for r in match_pumsam("전선관 포설", "Φ80x3열,Φ50x2열", rows)}  # 5열
    assert combo["배전전공"]["품셈"] == pytest.approx((0.009 * 3 + 0.007 * 2) / 5)
    assert combo["배전전공"]["할증%"] == pytest.approx(380) and "보간" in combo["배전전공"]["품셈근거"]
    ten = {r["노무명칭"]: r for r in match_pumsam("전선관 포설", "Φ80x6열,Φ50x4열", rows)}
    assert ten["배전전공"]["할증%"] == 580


def test_tray_control_cable_and_pvc_conduit_hyphen_underground() -> None:
    """공내역서 표기(트레이용 제어용케이블, 경질비닐PVC전선관-지중, HI, 28mm)가 5-13/5-1로 매칭된다."""
    from app.pumsam import load_pumsam_database, match_pumsam

    rows = load_pumsam_database(None, "전기")

    def first(name, spec):
        return [(r["노무명칭"], r["품셈"], r["할증%"]) for r in match_pumsam(name, spec, rows)]

    cvv = "0.6/1kV 트레이용 절연난연 제어용케이블(F-CVV)"
    assert first(cvv, "3C*1.5㎟") == [("저압케이블전공", 0.019, 100)]
    assert first(cvv, "4C*1.5㎟") == [("저압케이블전공", 0.026, 100)]
    assert first(cvv, "5C*1.5㎟") == [("저압케이블전공", 0.032, 100)]
    assert first("0.6/1kv 트레이용 제어용케이블(F-CVV-S)", "2C*1.5㎟") == [("저압케이블전공", 0.014, 120)]  # 해설⑥ 쉴드 120%
    # 하이픈 지중 표기 = 밑줄 지중(할증 70%)
    assert first("경질비닐PVC전선관-지중", "HI, 28mm") == [("내선전공", 0.08, 70)]
    assert first("경질비닐전선관_지중", "HI 28 mm") == [("내선전공", 0.08, 70)]
    # 트레이 자재 자체는 여전히 케이블 품으로 잡히지 않는다
    assert first("케이블트레이 부속", "행거") == []


def test_grounding_and_level_switch_researched_values() -> None:
    """나동선 3-38 0.008, 접지봉커넥터는 접지봉 품에 포함(공란), 수위스위치 6-2 0.22(계장공, 품셈 원본)."""
    from app.pumsam import load_pumsam_database, match_pumsam

    rows = load_pumsam_database(None, "전기")

    def first(name, spec):
        return [(r["노무명칭"], r["품셈"], r["할증%"], r["품셈근거"]) for r in match_pumsam(name, spec, rows)]

    assert first("접지봉", "14∮*1000mm") == [("내선전공", 0.11, 100, "전기3-38"), ("보통인부", 0.08, 100, "전기3-38")]
    # 접지봉-3본(1조) = 3-38 '3본 연결' 줄
    assert first("접지봉-3본", "14∮*1000mm(3본)") == [("내선전공", 0.24, 100, "전기3-38"), ("보통인부", 0.2, 100, "전기3-38")]
    assert first("접지봉커넥터", "U-BLOT형 14∮") == []  # 접지봉 품에 포함(3-38 해설①)
    # 쉴드케이블 5-13 해설⑥ 120%
    assert first("0.6/1kv 트레이용 제어용케이블(F-CVV-S)", "2C*1.5㎟")[0][:3] == ("저압케이블전공", 0.014, 120)
    assert first("전기용 나동선", "BC 50㎟")[0][:3] == ("내선전공", 0.008, 100)
    assert first("수위스위치", "오뚜기식")[0][:3] == ("계장공", 0.22, 100)
    # 대표님 자료 반영: F-GV 직매 6㎟ = 3-38 접지선매설 0.006, 압착단자 6㎟ = 3-38 압축단자 0.016, 현장조작반 = 5-39 플랜트전공 1.85
    assert first("난연 비닐절연접지용전선(F-GV)-직매", "0.6/1kv, 6㎟")[0][:3] == ("내선전공", 0.006, 100)
    assert first("난연 비닐절연접지용전선(F-GV)", "0.6/1kv, 2.5㎟")[0][:3] == ("내선전공", 0.01, 100)
    assert first("압착단자", "압착단자, 터미널, 6㎟")[0][:3] == ("내선전공", 0.016, 100)
    # 1.5·2.5·4·10㎟ 는 케이블 설치품(5-13 해설①(마))에 포함 - 공란
    for sz in ("1.5", "2.5", "4", "10"):
        assert first("압착단자", f"압착단자, 터미널, {sz}㎟") == []
    assert first("자가펌프 현장조작반", "옥외자립형")[0][:3] == ("플랜트전공", 1.85, 100)
    assert first("스위치", "1로 250V 15A")[0][:3] == ("내선전공", 0.085, 100)
    # 슬리브 접속(C TYPE) = 3-38 접속 및 단자설치 압축(압축슬리브 사용) 0.097
    assert first("슬리브 접속", "C TYPE SLEEVE 50㎟")[0][:3] == ("내선전공", 0.097, 100)


def test_pipe_inlet_waterproofing_uses_2_18() -> None:
    """관로구방수 ∮30·∮40 = 전기 표준품셈 해당 품 없음(대표님 자료) → 공란. 씰링가스켓 줄은 유지."""
    from app.pumsam import load_pumsam_database, match_pumsam

    rows = load_pumsam_database(None, "전기")
    for spec in ("∮30", "∮40"):
        assert match_pumsam("관로구방수", spec, rows) == []
    # 씰링가스켓 200 mm 이하 줄(2-18, 케이블 금구류 부착)은 그대로 유지된다
    gasket = {r["노무명칭"]: r["품셈"] for r in match_pumsam("케이블 금구류 부착", "200 mm 이하", rows)}
    assert gasket == {"특고압케이블전공": 0.026, "보통인부": 0.12}


def test_daepyo_notes() -> None:
    from app.ilwidae import daepyo_note

    assert daepyo_note("0.6/1kV 난연 전력케이블(F-CV)", "2C*2.5㎟") == "5-11 해설② 600V 10㎟ 이하는 5-13 준용, 단자처리 포함"
    assert daepyo_note("0.6/1kV 난연 전력케이블(F-CV)", "4C*4㎟") == "5-11 해설② 준용, 단자처리 포함"
    assert daepyo_note("0.6/1kv 트레이용 제어용케이블(F-CVV-S)", "2C*1.5㎟") == "쉴드케이블 120% (해설⑥)"
    assert daepyo_note("압착단자", "압착단자, 터미널, 4㎟").startswith("케이블 설치품")
    assert daepyo_note("압착단자", "압착단자, 터미널, 6㎟") == "접지선(F-GV 6㎟) 단자 설치"
    assert daepyo_note("슬리브 접속", "C TYPE SLEEVE 50㎟") is None
