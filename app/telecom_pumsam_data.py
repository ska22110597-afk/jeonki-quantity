"""정보통신공사 표준품셈 베이스. 원표 JSON + 단가대비표용 배관 별칭."""

from __future__ import annotations

import json

from app.electric_pumsam_data import (
    FLEX_METAL,
    HEAVY_STEEL,
    RESIN_CONDUIT,
    THIN_STEEL,
    _book_qty,
    _cd_qty,
    _hi_specs,
    _place_rows,
    _row,
)
from app.estimate_parse import lookup_key
from app.paths import bundled_data_dir
from app.pumsam_text import clean_spec_unit

PumsamRow = dict[str, object]

TELECOM_RULES = [
    "한국정보통신산업연구원 정보통신공사 표준품셈 적용 기준 (2025)",
    "",
    "장 구성: 1장 공통, 2장 관로·전봇대, 3장 배관, 4장 통신케이블, 5장 교환, 6장 전송, 7장 무선·방송, 8장 네트워크, 9장 정보설비, 10장 기계경비, 11장 전원, 12장 지능형 홈, 13장 유지보수.",
    "",
    "배관(3-1-1)",
    "1. 원표 단위는 10m 입니다. 단가대비표가 M 이면 같은 품을 1m 기준으로 나눠 둔 별칭(경질비닐전선관 등)을 씁니다.",
    "2. 접미사 없는 명칭은 콘크리트 매입 기준, 할증 100%.",
    "3. 명칭 끝 _매입 = 콘크리트 매입, 할증 100%.",
    "4. 명칭 끝 _노출 = 철근콘크리트 노출, 할증 120%.",
    "5. 명칭 끝 _지중 = 지중 매설, 할증 70%.",
    "6. CD관·폴리에틸렌관은 합성수지 품의 80%입니다.",
    "7. 전선관·박스·덕트·트레이 → 통신내선공. 광·시내 케이블 → 통신케이블공. 가공·전봇대 → 통신외선공.",
    "8. 명칭이 같고 규격이 같으면 그대로 붙습니다. 후강전선관 = 강제전선관, HI관 = 경질비닐전선관처럼 같은 품 묶음이면 그 이름으로도 찾습니다. 아연도 16 mm 와 16 mm / G 16 mm 처럼 앞말만 다르면 같은 크기로 맞춥니다. 딱 맞는 규격이 없고 「이하」 구간만 있으면 품목 규격 이상인 가장 작은 이하 구간을 씁니다. 예: 4㎟ → 6㎟ 이하, 3㎟ → 3㎟ 이하. 지중/노출처럼 다른 말은 끌어오지 않습니다.",
    "9. 품셈 칸은 표준품셈 원표 숫자입니다. 할증을 곱하지 않은 값입니다. 할증은 할증% 칸에 따로 적습니다. 일위대가 인부 규격은 일반공사 직종만 적고, 비고에는 전기 5-1 같은 적용품만 적습니다. 일위대가 인부 수량(품)은 비워 두고 엑셀에서 직접 채웁니다.",
    "10. 표에 직종이 두 개면 일위대가 호표에도 둘 다 넣습니다. 표에 없으면 통신내선공 1명만 넣습니다.",
    "11. 같은 명칭·규격에 인부가 여러 명이면 품셈표에서 아래 칸의 키워드·명칭만 비웁니다. 규격은 그대로 둡니다. 인부 행은 지우지 않습니다.",
]

TELECOM_TRADE_GUIDE = [
    ("공종", "품셈근거", "인부", "매입", "노출", "지중", "비고"),
    ("구내통신배관", "통신 3-1-1", "통신내선공", "100%", "120%", "70%", "원표 10m. CD·PE는 합성수지 80%"),
    ("박스·풀박스", "통신 3-2-1", "통신내선공", "100%", "120%", "-", "콘크리트 매입 기준"),
    ("단자함", "통신 3-3-1", "통신내선공, 보통인부", "100%", "120%", "-", "두 직종"),
    ("배선반·종말단자", "통신 3-3-2", "통신케이블공, 보통인부", "-", "-", "-", "피뢰탄기반은 통신내선공"),
    ("케이블랙·트레이", "통신 3-4", "통신내선공", "100%", "-", "-", ""),
    ("광섬유케이블", "통신 4-1", "통신케이블공", "-", "-", "-", "구내는 4-1-3"),
    ("꼬임케이블(UTP)", "통신 4-3", "통신내선공", "-", "-", "-", "구내 포설"),
    ("제어케이블", "통신 4-4", "통신케이블공", "-", "-", "-", "실드는 120%"),
    ("전봇대·지지선", "통신 2-4", "통신외선공, 보통인부", "-", "가공", "-", ""),
]


def _spec_variants(spec: str) -> list[str]:
    text, _ = clean_spec_unit(spec, "")
    return [text] if text else []


def book_pumsam_rows() -> list[PumsamRow]:
    path = bundled_data_dir() / "통신품셈_원표.json"
    if not path.exists():
        return []
    payload = json.loads(path.read_text(encoding="utf-8"))
    rows: list[PumsamRow] = []
    for item in payload:
        names = [str(item.get("명칭") or "").strip()]
        short = str(item.get("짧은명칭") or "").strip()
        if short and short not in names:
            names.append(short)
        unit = str(item.get("단위") or "").strip() or "식"
        labor = str(item.get("노무명칭") or "").strip()
        qty_f = _book_qty(item.get("품셈"))
        if qty_f is ValueError:
            continue
        rate = int(item.get("할증%") or 100)
        ref = str(item.get("품셈근거") or "")
        spec_text, unit = clean_spec_unit(str(item.get("규격") or ""), unit)
        for name in names:
            if not name or not labor:
                continue
            for spec in _spec_variants(spec_text):
                rows.append(_row(name, spec, unit, labor, qty_f, rate, ref))
    return rows


def _conduit_block() -> list[PumsamRow]:
    """3-1-1 을 단가대비표 M 단위에 맞게 1m 기준으로 풀어 둔다."""
    rows: list[PumsamRow] = []
    labor = "통신내선공"
    ref = "통신3-1-1"
    resin_names = ("경질비닐전선관", "합성수지전선관", "합성수지제전선관", "HIVP전선관", "HI관")
    for mm, qty in RESIN_CONDUIT.items():
        for spec in _hi_specs(mm):
            for name in resin_names:
                rows.extend(_place_rows(name, spec, "M", labor, qty, ref))
        cd_qty = _cd_qty(mm)
        for spec in (f"PE {mm} mm", f"CD {mm} mm", f"PE-{mm} mm", f"CD-{mm} mm"):
            for name in ("CD전선관", "CD관", "폴리에틸렌전선관"):
                rows.extend(_place_rows(name, spec, "M", labor, cd_qty, ref))
    for mm, qty in HEAVY_STEEL.items():
        for spec in (f"아연도 {mm} mm", f"G {mm} mm", f"G-{mm} mm", f"{mm} mm"):
            for name in ("강제전선관", "후강전선관", "아연도강전선관"):
                rows.extend(_place_rows(name, spec, "M", labor, qty, ref))
    flex_names = ("금속제가요전선관", "가요전선관", "EF전선관")
    for mm, qty in FLEX_METAL.items():
        spec = f"{mm} mm" if mm != 16 else "16 mm 이하"
        for name in flex_names:
            rows.extend(_place_rows(name, spec, "M", labor, qty, ref, include_buried=False))
        if mm == 16:
            for name in flex_names:
                rows.extend(_place_rows(name, "16 mm", "M", labor, qty, ref, include_buried=False))
    for mm, qty in THIN_STEEL.items():
        spec = f"{mm} mm"
        for name in ("박강전선관", "나사없는전선관", "E전선관"):
            rows.extend(_place_rows(name, spec, "M", labor, qty, ref))
    return rows


def _cable_aliases() -> list[PumsamRow]:
    rows: list[PumsamRow] = []
    for spec, qty in (
        ("CAT.5e", 0.040),
        ("CAT.6", 0.040),
        ("CAT.6A", 0.045),
        ("CAT.7", 0.050),
    ):
        for name in ("UTP케이블", "꼬임케이블", "LAN케이블"):
            rows.append(_row(name, spec, "M", "통신내선공", qty, 100, "통신4-3-1"))
    for spec, qty in (
        ("SM 4C", 0.080),
        ("SM 8C", 0.090),
        ("SM 12C", 0.100),
        ("SM 24C", 0.120),
        ("MM 4C", 0.080),
        ("MM 12C", 0.100),
    ):
        for name in ("광케이블", "광섬유케이블"):
            rows.append(_row(name, spec, "M", "통신케이블공", qty, 100, "통신4-1-1"))
    return rows


def telecom_pumsam_rows() -> list[PumsamRow]:
    merged: dict[str, PumsamRow] = {}
    for row in (*book_pumsam_rows(), *_conduit_block(), *_cable_aliases()):
        key = f"{lookup_key(row.get('명칭'), row.get('규격'))}|{row.get('노무명칭')}"
        merged[key] = row
    rows = list(merged.values())
    rows.sort(
        key=lambda item: (
            str(item.get("품셈근거") or ""),
            str(item.get("명칭") or ""),
            str(item.get("규격") or ""),
            str(item.get("노무명칭") or ""),
        )
    )
    return rows
