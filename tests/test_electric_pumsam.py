from __future__ import annotations

from app.electric_pumsam_data import ELECTRIC_TRADE_GUIDE, _cd_qty, electric_pumsam_rows
from app.pumsam import default_pumsam_rows, match_pumsam, pumsam_rate_value
from app.ilwidae import labor_qty_formula
from app.items import LineItem


def test_electric_pumsam_covers_places_and_trades() -> None:
    rows = electric_pumsam_rows()
    assert len(rows) > 1000
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
