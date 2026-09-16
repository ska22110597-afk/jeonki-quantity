"""대한전기협회 전기공사 표준품셈 제5장 내선설비 해석 베이스.

표 숫자는 전기공사 표준품셈 제5장(2026 적용, 내선 전선관·배선 표)을
프로그램이 붙일 수 있게 풀어 넣은 것이다.

- 매입: 콘크리트 매입 기준(할증 100%)
- 노출: 철근콘크리트 노출(할증 120%)
- 지중: 해당 품의 70%(할증 70%)
접미사 없는 명칭은 매입 기준이다. 단가대비표에 _노출/_매입/_지중을
붙이면 그 할증으로 맞춘다.
"""

from __future__ import annotations

import json

from app.estimate_parse import lookup_key
from app.paths import bundled_data_dir

PumsamRow = dict[str, object]


def _row(
    name: str,
    spec: str,
    unit: str,
    labor: str,
    qty: float,
    rate: int,
    ref: str,
) -> PumsamRow:
    return {
        "검색키": lookup_key(name, spec),
        "명칭": name,
        "규격": spec,
        "단위": unit,
        "노무명칭": labor,
        "품셈": qty,
        "할증%": rate,
        "품셈근거": ref,
    }


def _place_rows(
    name: str,
    spec: str,
    unit: str,
    labor: str,
    base_qty: float,
    ref: str,
    *,
    include_buried: bool = True,
    include_bare: bool = True,
) -> list[PumsamRow]:
    """매입 기준 품을 노출·지중 할증 행으로 펼친다."""
    rows: list[PumsamRow] = []
    if include_bare:
        rows.append(_row(name, spec, unit, labor, base_qty, 100, ref))
    rows.append(_row(f"{name}_매입", spec, unit, labor, base_qty, 100, ref))
    rows.append(_row(f"{name}_노출", spec, unit, labor, base_qty, 120, ref))
    if include_buried:
        rows.append(_row(f"{name}_지중", spec, unit, labor, base_qty, 70, ref))
    return rows


def _hi_specs(mm: int) -> tuple[str, ...]:
    return (f"HI {mm} mm", f"{mm} mm", f"HI-{mm} mm")


def _area_specs(label: str) -> tuple[str, ...]:
    """mm2 표기를 단가대비표에서 자주 쓰는 ㎟·mm² 로도 넣는다."""
    specs = [label]
    if " mm2" in label:
        specs.append(label.replace(" mm2", "㎟"))
        specs.append(label.replace(" mm2", " mm²"))
    return tuple(specs)


# 5-1 전선관 배관. 콘크리트 매입, 내선전공 인/m.
RESIN_CONDUIT = {  # 합성수지(경질비닐 HI)
    14: 0.04,
    16: 0.05,
    22: 0.06,
    28: 0.08,
    36: 0.10,
    42: 0.13,
    54: 0.19,
    70: 0.28,
    82: 0.37,
    92: 0.45,
    104: 0.46,
    125: 0.51,
}
HEAVY_STEEL = {  # 후강전선관(G)
    16: 0.08,
    22: 0.11,
    28: 0.14,
    36: 0.20,
    42: 0.25,
    54: 0.34,
    70: 0.44,
    82: 0.54,
    92: 0.60,
    104: 0.71,
}
FLEX_METAL = {  # 금속제 가요전선관
    16: 0.044,
    22: 0.059,
    28: 0.072,
    36: 0.087,
    42: 0.104,
    54: 0.136,
    70: 0.156,
    82: 0.176,
    92: 0.196,
    104: 0.216,
}
THIN_STEEL = {  # 나사 없는 전선관(E) / 박강전선관(C)
    19: 0.05,
    25: 0.06,
    31: 0.08,
    39: 0.10,
    51: 0.13,
    63: 0.19,
    75: 0.28,
}


def _cd_qty(mm: int) -> float:
    """CD·폴리에틸렌관은 합성수지 품의 80%. 100 mm 이상은 100%."""
    base = RESIN_CONDUIT[mm]
    return round(base if mm >= 100 else base * 0.8, 3)


def _conduit_block() -> list[PumsamRow]:
    rows: list[PumsamRow] = []
    resin_names = ("경질비닐전선관", "합성수지전선관", "합성수지제전선관", "HIVP전선관", "HI관")
    for mm, qty in RESIN_CONDUIT.items():
        for spec in _hi_specs(mm):
            for name in resin_names:
                rows.extend(_place_rows(name, spec, "M", "내선전공", qty, "전기5-1"))
        cd_qty = _cd_qty(mm)
        for spec in (f"PE {mm} mm", f"CD {mm} mm", f"PE-{mm} mm", f"CD-{mm} mm"):
            for name in ("CD전선관", "CD관", "폴리에틸렌전선관"):
                rows.extend(_place_rows(name, spec, "M", "내선전공", cd_qty, "전기5-1"))
    for mm, qty in HEAVY_STEEL.items():
        for spec in (f"아연도 {mm} mm", f"G {mm} mm", f"G-{mm} mm", f"{mm} mm"):
            for name in ("강제전선관", "후강전선관", "아연도강전선관"):
                rows.extend(_place_rows(name, spec, "M", "내선전공", qty, "전기5-1"))
    flex_names = ("금속제가요전선관", "가요전선관", "EF전선관")
    for mm, qty in FLEX_METAL.items():
        spec = f"{mm} mm" if mm != 16 else "16 mm 이하"
        for name in flex_names:
            rows.extend(_place_rows(name, spec, "M", "내선전공", qty, "전기5-1", include_buried=False))
        if mm == 16:
            for name in flex_names:
                rows.extend(_place_rows(name, "16 mm", "M", "내선전공", qty, "전기5-1", include_buried=False))
    for mm, qty in THIN_STEEL.items():
        spec = f"{mm} mm"
        for name in ("박강전선관", "나사없는전선관", "E전선관"):
            rows.extend(_place_rows(name, spec, "M", "내선전공", qty, "전기5-1"))
    return rows


def _indoor_block() -> list[PumsamRow]:
    rows: list[PumsamRow] = []
    # 5-3 박스. 콘크리트 매입 기준. 노출형은 표에 이미 노출 품.
    boxes = [
        ("콘크리트박스", "일반", 0.12),
        ("Outlet박스", "일반", 0.20),
        ("스위치박스", "2개용 이하", 0.20),
        ("스위치박스", "3개용 이상", 0.25),
        ("플로어박스", "일반", 0.20),
        ("연결용박스", "일반", 0.04),
    ]
    for name, spec, qty in boxes:
        rows.extend(_place_rows(name, spec, "개", "내선전공", qty, "전기5-3", include_buried=False))
    rows.append(_row("노출형박스", "콘크리트 노출기준", "개", "내선전공", 0.29, 100, "전기5-3"))

    pulls = [
        ("100 mm×100 mm×100 mm 이하", 0.04, 0.17),
        ("250 mm×250 mm×200 mm", 0.22, 0.55),
        ("400 mm×400 mm×300 mm", 0.35, 0.66),
        ("700 mm×700 mm×400 mm", 0.66, 0.95),
        ("1000 mm×1000 mm×150 mm", 0.95, 1.23),
        ("1200 mm×1200 mm×150 mm", 1.30, 1.56),
        ("1500 mm×1500 mm×250 mm", 2.50, 3.00),
        ("2000 mm×2000 mm×300 mm", 4.70, 5.64),
    ]
    for spec, ceil_qty, wall_qty in pulls:
        rows.append(_row("풀박스_천장", spec, "개", "내선전공", ceil_qty, 100, "전기5-4"))
        rows.append(_row("풀박스_벽면", spec, "개", "내선전공", wall_qty, 100, "전기5-4"))
        rows.extend(_place_rows("풀박스", spec, "개", "내선전공", wall_qty, "전기5-4", include_buried=False))

    headers = [("150×40", 0.30), ("200×40", 0.40), ("300×40", 0.54)]
    for spec, qty in headers:
        rows.append(_row("헤더덕트", spec, "M", "내선전공", qty, 100, "전기5-5"))
    rows.append(_row("시스템박스", "콘크리트매입 전선관용", "개", "내선전공", 0.63, 100, "전기5-5"))
    rows.append(_row("시스템박스", "콘크리트매입 데크플레이트용", "개", "내선전공", 0.41, 100, "전기5-5"))
    rows.append(_row("시스템박스", "액세스 플로어용", "개", "내선전공", 0.25, 100, "전기5-5"))

    floors = [
        ("F4 35×41", "M", 0.60),
        ("F7 35×73", "M", 0.70),
        ("F5 25×51", "M", 0.50),
        ("F6 노스타드 25×51", "M", 0.50),
        ("F6 23×60", "M", 0.60),
        ("F6 노스타드 25×55", "M", 0.50),
        ("F8 23×80", "M", 0.60),
        ("Junction박스 대형", "개", 1.00),
        ("Junction박스 중형", "개", 0.90),
        ("Junction박스 소형", "개", 0.80),
        ("노출 Insert Cap", "개", 0.10),
    ]
    for spec, unit, qty in floors:
        rows.append(_row("플로어덕트", spec, unit, "내선전공", qty, 100, "전기5-6"))

    ducts = [
        ("60 mm×30 mm 이하", 0.15),
        ("100 mm×50 mm", 0.20),
        ("100 mm×100 mm", 0.30),
        ("150 mm×100 mm", 0.40),
        ("200 mm×100 mm", 0.45),
        ("300 mm×100 mm", 0.50),
        ("400 mm×150 mm", 0.60),
        ("500 mm×200 mm", 1.50),
        ("600 mm×300 mm", 2.00),
        ("700 mm×400 mm", 2.50),
        ("1000 mm×400 mm", 3.00),
        ("1200 mm×450 mm", 3.70),
    ]
    for spec, qty in ducts:
        rows.append(_row("금속덕트", spec, "M", "내선전공", qty, 100, "전기5-7"))

    trays = [
        ("10000 mm2 이하", 0.18, 0.13),
        ("30000 mm2 이하", 0.23, 0.16),
        ("50000 mm2 이하", 0.30, 0.20),
        ("60000 mm2 이하", 0.36, 0.25),
        ("80000 mm2 이하", 0.48, 0.34),
        ("90000 mm2 이하", 0.54, 0.38),
        ("120000 mm2 이하", 0.72, 0.50),
        ("150000 mm2 이하", 0.90, 0.63),
    ]
    for spec, steel, alu in trays:
        rows.append(_row("케이블트레이", spec, "M", "내선전공", steel, 100, "전기5-8"))
        rows.append(_row("케이블트레이_철제", spec, "M", "내선전공", steel, 100, "전기5-8"))
        rows.append(_row("케이블트레이_알루미늄", spec, "M", "내선전공", alu, 100, "전기5-8"))
    rows.append(_row("케이블트레이내진버팀대", "전산볼트 13 mm 이하", "식", "내선전공", 0.16, 100, "전기5-8-1"))

    rows.append(_row("금속몰딩", "소형 210 mm2 이하", "M", "내선전공", 0.16, 100, "전기5-9"))
    rows.append(_row("금속몰딩", "중형 595 mm2 이하", "M", "내선전공", 0.18, 100, "전기5-9"))
    rows.append(_row("금속몰딩", "대형 600 mm2 초과", "M", "내선전공", 0.22, 100, "전기5-9"))
    rows.append(_row("PVC몰딩", "바닥", "M", "내선전공", 0.025, 100, "전기5-9"))
    rows.append(_row("PVC몰딩", "벽면", "M", "내선전공", 0.025, 110, "전기5-9"))
    rows.append(_row("PVC몰딩", "천장", "M", "내선전공", 0.025, 130, "전기5-9"))
    rows.append(_row("레이스웨이", "40×40 1600 mm2 이하", "M", "내선전공", 0.20, 100, "전기5-9-1"))
    rows.append(_row("레이스웨이", "70×40 2800 mm2 이하", "M", "내선전공", 0.30, 100, "전기5-9-1"))
    rows.append(_row("레이스웨이", "110×50 5500 mm2 이하", "M", "내선전공", 0.51, 100, "전기5-9-1"))

    bus = [
        (100, 0.18, 0.21, 0.15, 0.18, 0.21, 0.24),
        (200, 0.24, 0.28, 0.21, 0.25, 0.26, 0.29),
        (400, 0.33, 0.38, 0.24, 0.29, 0.30, 0.34),
        (600, 0.51, 0.59, 0.41, 0.50, 0.53, 0.65),
        (800, 0.92, 1.06, 0.73, 0.84, 0.90, 1.00),
        (1000, 1.00, 1.15, 0.76, 0.87, 0.93, 1.02),
        (1200, 1.80, 2.10, 1.50, 1.70, 2.00, 2.10),
        (1500, 2.00, 2.30, 1.60, 1.90, 2.10, 2.30),
        (2000, 3.30, 3.60, 2.50, 2.90, 3.00, 3.40),
        (2500, 4.60, 5.30, 3.50, 4.10, 4.00, 4.80),
        (3000, 6.00, 6.90, 4.50, 5.30, 5.40, 6.20),
    ]
    kinds = ("Cu-Fe 3W", "Cu-Fe 4W", "Al-Al 3W", "Al-Al 4W", "Al-Fe 3W", "Al-Fe 4W")
    for amp, *vals in bus:
        for kind, qty in zip(kinds, vals):
            spec = f"{amp} A 이하 {kind}"
            rows.append(_row("모선덕트", spec, "M", "내선전공", qty, 100, "전기5-16"))
            rows.append(_row("버스덕트", spec, "M", "내선전공", qty, 100, "전기5-16"))
    return rows


def _wire_cable_block() -> list[PumsamRow]:
    rows: list[PumsamRow] = []
    indoor = [
        ("6 mm2 이하", 0.010),
        ("16 mm2 이하", 0.023),
        ("38 mm2 이하", 0.031),
        ("50 mm2 이하", 0.043),
        ("60 mm2 이하", 0.052),
        ("70 mm2 이하", 0.061),
        ("100 mm2 이하", 0.064),
        ("120 mm2 이하", 0.077),
        ("150 mm2 이하", 0.088),
        ("200 mm2 이하", 0.107),
        ("250 mm2 이하", 0.130),
        ("300 mm2 이하", 0.148),
        ("325 mm2 이하", 0.160),
        ("400 mm2 이하", 0.197),
    ]
    indoor_names = ("옥내배선", "IV전선", "HIV전선", "HFIX전선", "600V IV", "600V HIV")
    indoor_sizes = {
        0.010: (1.6, 2.0, 2.6, 3.5, 5.5, 6),
        0.023: (8, 14, 16),
        0.031: (22, 25, 38),
        0.043: (50,),
        0.052: (60,),
        0.061: (70,),
        0.064: (80, 100),
        0.077: (120,),
        0.088: (150,),
        0.107: (185, 200),
        0.130: (240, 250),
        0.148: (300,),
        0.160: (325,),
        0.197: (400,),
    }
    for spec, qty in indoor:
        for label in _area_specs(spec):
            for name in indoor_names:
                rows.append(_row(name, label, "M", "내선전공", qty, 100, "전기5-10"))
        for size in indoor_sizes.get(qty, ()):
            for label in _area_specs(f"{size} mm2"):
                for name in indoor_names:
                    rows.append(_row(name, label, "M", "내선전공", qty, 100, "전기5-10"))

    power = [
        ("16 mm2 이하×1C", 0.023),
        ("25 mm2×1C", 0.030),
        ("38 mm2×1C", 0.036),
        ("50 mm2×1C", 0.043),
        ("60 mm2×1C", 0.049),
        ("70 mm2×1C", 0.057),
        ("80 mm2×1C", 0.060),
        ("100 mm2×1C", 0.071),
        ("125 mm2×1C", 0.084),
        ("150 mm2×1C", 0.097),
        ("185 mm2×1C", 0.108),
        ("200 mm2×1C", 0.117),
        ("240 mm2×1C", 0.136),
        ("250 mm2×1C", 0.142),
        ("300 mm2×1C", 0.159),
        ("325 mm2×1C", 0.172),
        ("400 mm2×1C", 0.205),
        ("500 mm2×1C", 0.240),
        ("630 mm2×1C", 0.285),
        ("1000 mm2×1C", 0.415),
    ]
    for spec, qty in power:
        for label in _area_specs(spec):
            for name in ("전력케이블", "CV케이블", "VV케이블", "F-CV케이블", "600V CV"):
                rows.append(_row(name, label, "M", "저압케이블전공", qty, 100, "전기5-11"))
                rows.append(_row(f"{name}_직매", label, "M", "저압케이블전공", qty, 80, "전기5-11"))

    cores = ["1C", "2C", "3C", "4C", "5C", "6C", "7C", "8C", "10C", "12C", "14C", "15C", "19C", "20C", "24C", "30C", "50C"]
    ctrl_25 = [0.010, 0.014, 0.019, 0.026, 0.032, 0.035, 0.039, 0.042, 0.048, 0.054, 0.059, 0.062, 0.072, 0.074, 0.084, 0.098, 0.112]
    ctrl_4 = [0.011, 0.016, 0.022, 0.029, 0.034, 0.038, 0.042, 0.046, 0.052, 0.058, 0.064, 0.067, 0.078, 0.08, 0.09, None, None]
    ctrl_6 = [0.013, 0.018, 0.026, 0.034, 0.039, 0.044, 0.048, 0.052, 0.059, 0.066, 0.073, 0.076, 0.089, 0.092, 0.103, None, None]
    ctrl_8 = [0.014, 0.020, 0.029, 0.039, 0.044, 0.050, 0.054, 0.058, 0.067, None, None, None, None, None, None, None, None]
    ctrl_10 = [0.018, 0.025, 0.036, 0.049, 0.055, 0.063, 0.068, 0.073, 0.084, None, None, None, None, None, None, None, None]
    for size, col in (
        ("2.5 mm2 이하", ctrl_25),
        ("4 mm2 이하", ctrl_4),
        ("6 mm2 이하", ctrl_6),
        ("8 mm2 이하", ctrl_8),
        ("10 mm2 이하", ctrl_10),
    ):
        for core, qty in zip(cores, col):
            if qty is None:
                continue
            for label in _area_specs(f"{size} {core}"):
                for name in ("제어용케이블", "CVV케이블", "CCV케이블"):
                    rows.append(_row(name, label, "M", "저압케이블전공", qty, 100, "전기5-13"))

    vvf = [
        ("1.6 mm-2C", 0.020, 0.026, 0.010),
        ("2.0 mm-2C", 0.025, 0.033, 0.013),
        ("2.6 mm-2C", 0.031, 0.042, 0.017),
        ("1.6 mm-3C", 0.025, 0.033, 0.013),
        ("2.0 mm-3C", 0.030, 0.041, 0.017),
        ("2.6 mm-3C", 0.038, 0.051, 0.021),
    ]
    for spec, wood, conc, ceil in vvf:
        rows.append(_row("VVF케이블", spec, "M", "저압케이블전공", wood, 100, "전기5-14"))
        rows.append(_row("VVF케이블_목조", spec, "M", "저압케이블전공", wood, 100, "전기5-14"))
        rows.append(_row("VVF케이블_콘크리트", spec, "M", "저압케이블전공", conc, 100, "전기5-14"))
        rows.append(_row("VVF케이블_천장", spec, "M", "저압케이블전공", ceil, 100, "전기5-14"))

    vvr_2 = [("1.6 mm", 0.026), ("2.0 mm", 0.041), ("5.5 mm2", 0.047), ("8 mm2", 0.052), ("14 mm2", 0.063), ("38 mm2", 0.100), ("60 mm2", 0.147), ("100 mm2", 0.190), ("150 mm2", 0.239)]
    vvr_3 = [("1.6 mm", 0.038), ("2.0 mm", 0.046), ("5.5 mm2", 0.067), ("8 mm2", 0.070), ("14 mm2", 0.080), ("38 mm2", 0.147), ("60 mm2", 0.189), ("100 mm2", 0.234), ("150 mm2", 0.306)]
    for spec, qty in vvr_2:
        rows.append(_row("VVR케이블", f"{spec} 2C", "M", "저압케이블전공", qty, 100, "전기5-15"))
    for spec, qty in vvr_3:
        rows.append(_row("VVR케이블", f"{spec} 3C", "M", "저압케이블전공", qty, 100, "전기5-15"))
    return rows


def _gear_light_block() -> list[PumsamRow]:
    rows: list[PumsamRow] = []
    mccb = [("30 AF 이하", 0.19), ("50 AF", 0.26), ("100 AF", 0.36), ("225 AF", 0.47), ("400 AF 이하", 0.68), ("600 AF", 0.78), ("800 AF", 0.89)]
    for spec, qty in mccb:
        compact = spec.replace(" ", "")
        for label in (spec, compact):
            for name in ("배선용차단기", "누전차단기", "MCCB", "ELB"):
                rows.append(_row(name, label, "개", "내선전공", qty, 100, "전기5-19"))
    safety = [("30 A 이하", 0.20), ("50 A", 0.30), ("100 A", 0.40), ("225 A", 0.55), ("300 A", 0.70), ("400 A", 0.87), ("600 A", 1.15), ("800 A", 1.50)]
    magnet = [("30 A 이하", 0.30), ("50 A", 0.45), ("100 A", 0.60), ("225 A", 0.80), ("300 A", 1.05), ("400 A", 1.25), ("600 A", 1.70), ("800 A", 2.20)]
    for spec, qty in safety:
        rows.append(_row("안전개폐기", spec, "개", "내선전공", qty, 100, "전기5-19"))
    for spec, qty in magnet:
        rows.append(_row("마그넷스위치", spec, "개", "내선전공", qty, 100, "전기5-19"))
    acb = [("1500 A 이하", 2.3), ("1500 A 초과~3000 A", 2.6), ("3000 A 초과~5000 A", 3.0)]
    for spec, qty in acb:
        rows.append(_row("저압기중차단기", spec, "개", "내선전공", qty, 100, "전기5-20"))

    for n, qty in enumerate([0.59, 0.65, 0.71, 0.77, 0.83, 0.89, 0.95, 1.01], start=3):
        spec = f"{n}회로"
        rows.append(_row("세대분전반", spec, "식", "내선전공", qty, 100, "전기5-18-1"))
        rows.append(_row("주택용분전반", spec, "식", "내선전공", qty, 100, "전기5-18-1"))
        rows.append(_row("세대분전반_매입", spec, "식", "내선전공", qty, 100, "전기5-18-1"))
        rows.append(_row("세대분전반_노출", spec, "식", "내선전공", qty, 90, "전기5-18"))
    rows.append(_row("가로등분전반", "4회로", "대", "내선전공", 0.86, 100, "전기5-18-2"))
    rows.append(_row("가로등분전반", "6회로", "대", "내선전공", 1.02, 100, "전기5-18-2"))
    rows.append(_row("가로등분전반", "8회로", "대", "내선전공", 1.23, 100, "전기5-18-2"))

    board = [
        ("10 P 이하", 0.65, 0.45),
        ("20 P 이하", 0.68, 0.46),
        ("50 P 이하", 0.72, 0.48),
        ("100 P 이하", 1.29, 0.86),
        ("150 P 이하", 1.87, 1.24),
        ("200 P 이하", 2.44, 1.63),
        ("250 P 이하", 3.02, 2.01),
    ]
    for spec, a, b in board:
        for name in ("배선용단자함", "단자함"):
            rows.append(_row(name, spec, "대", "내선전공", a, 100, "전기5-41"))
            rows.append(_row(name, spec, "대", "보통인부", b, 100, "전기5-41"))

    outlets = [
        ("콘센트 15 A", "2P", 0.065),
        ("콘센트 접지극부 15 A", "2P", 0.08),
        ("콘센트 접지극부 20 A", "2P", 0.085),
        ("콘센트 접지극부 30 A", "2P", 0.11),
        ("플로어콘센트 15 A", "2P", 0.096),
        ("플로어콘센트 20 A", "2P", 0.096),
        ("콘센트 15 A", "3P", 0.095),
        ("콘센트 접지극부 30 A", "3P", 0.145),
        ("콘센트 15 A", "4P", 0.10),
        ("콘센트 접지극부 30 A", "4P", 0.15),
    ]
    for name, spec, qty in outlets:
        rows.extend(_place_rows(name, spec, "개", "내선전공", qty, "전기5-23", include_buried=False))

    switches = [
        ("텀블러스위치 단로용", 0.085),
        ("텀블러스위치 3로용", 0.085),
        ("텀블러스위치 4로용", 0.10),
        ("풀스위치", 0.10),
        ("푸시버튼", 0.065),
        ("리모콘스위치", 0.07),
        ("표시등", 0.10),
        ("자동점멸기 광전식", 0.19),
        ("타임스위치", 0.20),
        ("조명등용센서스위치", 0.063),
    ]
    for name, qty in switches:
        rows.extend(_place_rows(name, "일반", "개", "내선전공", qty, "전기5-23", include_buried=False))

    led = [
        ("15 W 이하", 0.117, 0.158, 0.155, None),
        ("25 W 이하", 0.138, 0.163, 0.182, None),
        ("35 W 이하", 0.163, 0.213, 0.208, 0.242),
        ("45 W 이하", 0.221, 0.249, None, 0.263),
        ("55 W 이하", 0.254, None, None, 0.306),
    ]
    for spec, wall, pend, down, rec in led:
        for label in (spec, spec.replace(" ", "")):
            rows.append(_row("LED등기구_직부", label, "개", "내선전공", wall, 100, "전기5-25-3"))
            rows.append(_row("LED등기구", label, "개", "내선전공", wall, 100, "전기5-25-3"))
            if pend is not None:
                rows.append(_row("LED등기구_펜던트", label, "개", "내선전공", pend, 100, "전기5-25-3"))
            if down is not None:
                rows.append(_row("LED등기구_다운라이트", label, "개", "내선전공", down, 100, "전기5-25-3"))
            if rec is not None:
                rows.append(_row("LED등기구_매입", label, "개", "내선전공", rec, 100, "전기5-25-3"))
                rows.append(_row("LED등기구_반매입", label, "개", "내선전공", rec, 100, "전기5-25-3"))

    fl = [
        ("20 W 이하 ×1", 0.141, 0.168, 0.214),
        ("40 W 이하 ×1", 0.223, 0.268, 0.340),
        ("40 W 이하 ×2", 0.277, 0.332, 0.418),
    ]
    for spec, a, b, c in fl:
        rows.append(_row("형광등기구_직부", spec, "등", "내선전공", a, 100, "전기5-25"))
        rows.append(_row("형광등기구_펜던트", spec, "등", "내선전공", b, 100, "전기5-25"))
        rows.append(_row("형광등기구_매입", spec, "등", "내선전공", c, 100, "전기5-25"))

    meters = [
        ("전력량계 1ø 2W", 0.14),
        ("전력량계 1ø 3W", 0.21),
        ("전력량계 3ø 4W", 0.32),
        ("전류변성기 CT", 0.40),
        ("전압변성기 PT", 0.40),
        ("계기함", 0.30),
    ]
    for spec, qty in meters:
        rows.append(_row("전력량계", spec, "대", "내선전공", qty, 100, "전기5-21"))

    rods = [
        ("7.5 m 이하", 0.66),
        ("10 m 이하", 0.84),
        ("15 m 이하", 1.14),
        ("20 m 이하", 1.50),
        ("25 m 이하", 1.80),
        ("30 m 이하", 2.11),
    ]
    for spec, qty in rods:
        rows.append(_row("피뢰침", spec, "본", "내선전공", qty, 100, "전기5-42"))
    rows.append(_row("수평도체", "일반", "M", "내선전공", 0.017, 100, "전기5-42-2"))
    rows.append(_row("수평도체", "일반", "M", "보통인부", 0.008, 100, "전기5-42-2"))

    rows.append(_row("박스커버", "일반", "장", "내선전공", 0.03, 100, "전기5-29"))
    rows.append(_row("칼블럭", "9 mm 이하", "개", "내선전공", 0.028, 100, "전기5-29"))
    rows.append(_row("칼블럭", "12 mm 이하", "개", "내선전공", 0.036, 100, "전기5-29"))
    rows.append(_row("앵커볼트", "13 mm 이하", "개", "내선전공", 0.036, 100, "전기5-29"))
    rows.append(_row("앵커볼트", "16~19 mm", "개", "내선전공", 0.12, 100, "전기5-29"))
    for mm, qty in ((22, 0.08), (28, 0.12), (36, 0.16), (42, 0.20), (54, 0.30), (70, 0.45), (82, 0.55)):
        rows.append(_row("배관용홈파기", f"{mm} mm 이하", "M", "보통인부", qty, 100, "전기5-29"))
    rows.append(_row("가로등기초", "높이 12 m 이하", "개소", "내선전공", 0.08, 100, "전기5-27-1"))
    rows.append(_row("가로등기초", "높이 12 m 이하", "개소", "보통인부", 0.19, 100, "전기5-27-1"))
    return rows


ELECTRIC_RULES = [
    "대한전기협회 전기공사 표준품셈 전문 적용 기준 (2026, 내선 표는 단가대비표 명칭에 맞게 풀음)",
    "",
    "장 구성: 1장 적용기준, 2장 송전, 3장 변전, 4장 배전, 5장 내선, 6장 계측·자동제어, 7장 전기철도.",
    "",
    "내선(5장) 배관",
    "1. 접미사 없는 명칭은 콘크리트 매입 기준입니다. 할증 100%.",
    "2. 명칭 끝 _매입 = 콘크리트 매입, 할증 100%.",
    "3. 명칭 끝 _노출 = 철근콘크리트 노출, 할증 120%. 품셈 숫자는 매입 기준값 그대로 두고 할증만 올립니다.",
    "4. 명칭 끝 _지중 = 지중 배관, 할증 70%.",
    "5. CD관·폴리에틸렌관은 합성수지(경질비닐) 품의 80%입니다. 100 mm 이상은 100%.",
    "6. 전력케이블 _직매는 할증 80%입니다. 인부는 저압케이블전공입니다.",
    "",
    "인부",
    "7. 전선관·박스·덕트·트레이·등기구·차단기 → 내선전공.",
    "8. 전력·제어·VVF·VVR 케이블 → 저압케이블전공.",
    "9. 송전 철탑·가선 → 송전전공(활선은 송전활선전공), 지중 특고압 케이블 → 특고압케이블전공.",
    "10. 변압기·GIS·배전반 → 변전전공(+ 특별인부·비계공·인력운반공 등 표에 있는 직종).",
    "11. 전주·가선·주상변압기 → 배전전공(+ 보통인부). 활선은 배전활선전공.",
    "12. 계기반·계기 → 계장공. 전철 강체·전차선 → 배전전공(표 기준).",
    "13. 배선용단자함·가로등기초·수평도체는 내선전공과 보통인부를 함께 넣습니다.",
    "14. 단가대비표의 명칭+규격이 이 표와 글자 그대로 같아야 붙습니다.",
    "15. 표에 없는 품목은 전기 내선전공 1명만 넣습니다.",
]

ELECTRIC_TRADE_GUIDE = [
    ("공종", "품셈근거", "인부", "매입", "노출", "지중", "비고"),
    ("송전 철탑·가선", "전기 2장", "송전전공, 특별인부, 전기공사기사", "-", "가공", "지중 케이블은 특고압케이블전공", "활선 작업은 송전활선전공"),
    ("변전 변압기·GIS·배전반", "전기 3장", "변전전공, 특별인부, 비계공 등", "-", "-", "-", "용량·종별 계 행 기준"),
    ("배전 전주·가선", "전기 4장", "배전전공, 보통인부", "-", "가공", "지중 케이블·관로", "활선은 배전활선전공"),
    ("전선관 배관", "전기 5-1", "내선전공", "100%", "120%", "70%", "콘크리트 매입 기준. CD·PE는 합성수지 80%(100mm 이상 100%)"),
    ("박스", "전기 5-3", "내선전공", "100%", "120%", "-", "노출형박스는 표에 노출 품이 따로 있음"),
    ("풀박스", "전기 5-4", "내선전공", "100%", "120%", "-", "천장·벽면 품이 다름. 접미사 없는 풀박스는 벽면 품"),
    ("헤더덕트·시스템박스", "전기 5-5", "내선전공", "100%", "-", "-", "설치 위치별 규격"),
    ("플로어덕트", "전기 5-6", "내선전공", "100%", "-", "-", ""),
    ("금속덕트", "전기 5-7", "내선전공", "100%", "-", "-", ""),
    ("케이블트레이", "전기 5-8", "내선전공", "100%", "-", "-", "철제·알루미늄 단면적별"),
    ("몰딩·레이스웨이", "전기 5-9", "내선전공", "100%", "-", "-", "PVC몰딩 벽면 110%, 천장 130%"),
    ("옥내배선(IV·HIV·HFIX)", "전기 5-10", "내선전공", "100%", "-", "-", "단면적 이하 구간"),
    ("전력케이블(CV·VV)", "전기 5-11", "저압케이블전공", "100%", "-", "직매 80%", "1C 기준"),
    ("제어용케이블(CVV)", "전기 5-13", "저압케이블전공", "100%", "-", "-", "심선 수·단면적"),
    ("VVF케이블", "전기 5-14", "저압케이블전공", "-", "-", "-", "목조·콘크리트·천장 품이 다름"),
    ("VVR케이블", "전기 5-15", "저압케이블전공", "100%", "-", "-", "2C·3C"),
    ("모선덕트(버스덕트)", "전기 5-16", "내선전공", "-", "-", "-", "재질·심선(3W/4W)·전류"),
    ("세대·가로등 분전반", "전기 5-18", "내선전공", "100%", "90%", "-", "노출형은 매입의 90%"),
    ("차단기·개폐기", "전기 5-19", "내선전공", "100%", "-", "-", "MCCB·ELB·안전개폐기·마그넷"),
    ("저압기중차단기", "전기 5-20", "내선전공", "100%", "-", "-", "ACB"),
    ("전력량계", "전기 5-21", "내선전공", "100%", "-", "-", ""),
    ("콘센트·스위치", "전기 5-23", "내선전공", "100%", "120%", "-", ""),
    ("형광등기구", "전기 5-25", "내선전공", "매입 별도", "직부·펜던트", "-", "설치 방식별 품"),
    ("LED등기구", "전기 5-25-3", "내선전공", "매입 별도", "직부·펜던트·다운", "-", "설치 방식별 품"),
    ("가로등기초", "전기 5-27-1", "내선전공, 보통인부", "100%", "-", "-", "두 직종을 같이 넣음"),
    ("옥내잡공사", "전기 5-29", "내선전공 또는 보통인부", "100%", "-", "-", "홈파기는 보통인부"),
    ("배선용단자함", "전기 5-41", "내선전공, 보통인부", "100%", "-", "-", "두 직종을 같이 넣음"),
    ("피뢰침·수평도체", "전기 5-42", "내선전공 / 수평도체는 보통인부 추가", "100%", "-", "-", ""),
    ("계측·자동제어", "전기 6장", "계장공, 보통인부", "-", "-", "-", "계기반·계기"),
    ("전기철도", "전기 7장", "배전전공, 특별인부, 도장공 등", "-", "-", "-", "강체·전차선·조가선"),
]


def _spec_variants(spec: str) -> list[str]:
    text = str(spec or "").strip()
    if not text:
        return []
    alts = [text]
    swapped = (
        text.replace("ｍ", "m")
        .replace("㎜", "mm")
        .replace("㎟", "mm2")
        .replace("㎞", "km")
    )
    if swapped not in alts:
        alts.append(swapped)
    return alts


def book_pumsam_rows() -> list[PumsamRow]:
    """2~7장 원표. 단가대비표에 붙이도록 짧은 이름·단위 글자만 풀어 둔다."""
    path = bundled_data_dir() / "전기품셈_원표.json"
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
        qty = item.get("품셈")
        rate = int(item.get("할증%") or 100)
        ref = str(item.get("품셈근거") or "")
        try:
            qty_f = float(qty)
        except (TypeError, ValueError):
            continue
        for name in names:
            if not name or not labor:
                continue
            for spec in _spec_variants(str(item.get("규격") or "")):
                rows.append(_row(name, spec, unit, labor, qty_f, rate, ref))
    return rows


def electric_pumsam_rows() -> list[PumsamRow]:
    """전기 표준품셈 베이스. 같은 명칭·규격·직종은 한 줄만 남긴다."""
    merged: dict[str, PumsamRow] = {}
    for row in (
        *book_pumsam_rows(),
        *_conduit_block(),
        *_indoor_block(),
        *_wire_cable_block(),
        *_gear_light_block(),
    ):
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
