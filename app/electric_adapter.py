"""전기 전선·케이블: 단가대비표의 긴 품명·규격을 표준품셈 5-10/5-11/5-13 줄로 바꿔 찾는다.

예) 접지용비닐절연전선(F-GV) 16㎟, 450/750V 내열비닐절연전선 HFIX 2.5㎟ → 옥내배선(5-10).
    F-CVV-S 2Cx1.5㎟ → 제어용케이블(5-13).
다심 전력케이블은 5-11 해설에 따라 10㎟ 이하는 제어용케이블, 초과는 1C 품의 2심 140%·3심 200%·4심 260%.
"""

from __future__ import annotations

import re

_AREA = re.compile(r"(\d+(?:\.\d+)?)\s*(?:㎟|mm2|mm²|sq)", re.I)
_CORES = re.compile(r"(?<![\d.])(\d{1,2})\s*C(?:(?=[xX×*])|(?![A-Za-z]))", re.I)

INDOOR_BUCKETS = (6, 16, 38, 50, 60, 70, 100, 120, 150, 200, 250, 300, 325, 400)
POWER_SIZES = (25, 38, 50, 60, 70, 80, 100, 125, 150, 185, 200, 240, 250, 300, 325, 400, 500, 630, 1000)
CONTROL_BUCKETS = (2.5, 4, 6, 8, 10)


def _fmt(value: float) -> str:
    return str(int(value)) if float(value) == int(value) else str(value)


def _bucket(area: float, buckets) -> float | None:
    for limit in buckets:
        if area <= limit + 1e-9:
            return limit
    return None


_PULL_BOXES = (
    (100, 100, 100, "100 mm×100 mm×100 mm 이하"),
    (250, 250, 200, "250 mm×250 mm×200 mm"),
    (400, 400, 300, "400 mm×400 mm×300 mm"),
    (700, 700, 400, "700 mm×700 mm×400 mm"),
    (1000, 1000, 150, "1000 mm×1000 mm×150 mm"),
    (1200, 1200, 150, "1200 mm×1200 mm×150 mm"),
    (1500, 1500, 250, "1500 mm×1500 mm×250 mm"),
    (2000, 2000, 300, "2000 mm×2000 mm×300 mm"),
)
_ACCESSORY = ("커버", "COVER", "부품", "행거", "앵커", "볼트", "커넥터", "브라켓", "BRACKET", "고정", "서포트")
_MM = re.compile(r"(\d+)\s*(?:mm|㎜)", re.I)


def _box_candidates(name: str, spec: str) -> list[tuple[str, str, int]] | None:
    """박스 류. 해당 품목이 아니면 None."""
    text = f"{name} {spec}"
    compact = text.replace(" ", "")
    if "박스" not in compact and "BOX" not in compact.upper():
        return None
    if any(word in compact.upper() for word in _ACCESSORY):
        return []
    if "풀박스" in compact:
        dims = [int(x) for x in re.findall(r"\d+", spec.replace(",", ""))[:3]]
        if len(dims) < 3:
            return []
        for w, d, h, label in _PULL_BOXES:
            if dims[0] <= w and dims[1] <= d and dims[2] <= h:
                return [("풀박스", label, 100)]
        return []
    if "아우트렛" in compact or "아웃렛" in compact or "OUTLET" in compact.upper():
        return [("Outlet박스", "일반", 100)]
    if "스위치박스" in compact:
        count = re.search(r"(\d+)\s*개용", text)
        n = int(count.group(1)) if count else 1
        return [("스위치박스", "2개용 이하" if n <= 2 else "3개용 이상", 100)]
    if "정션" in compact or "연결박스" in compact or "연결용박스" in compact:
        return [("연결용박스", "일반", 100)]
    if "콘크리트박스" in compact:
        return [("콘크리트박스", "일반", 100)]
    if "플로어박스" in compact:
        return [("플로어박스", "일반", 100)]
    return None


_LED_BUCKETS = (15, 25, 35, 45, 55)
_LED_TYPES = (
    ("다운라이트", "LED등기구_다운라이트"),
    ("펜던트", "LED등기구_펜던트"),
    ("펜단트", "LED등기구_펜던트"),
    ("반매입", "LED등기구_반매입"),
    ("매입", "LED등기구_매입"),
    ("직부", "LED등기구_직부"),
    ("벽부", "LED등기구_직부"),
    ("레이스웨이", "LED등기구_직부"),
)


def _lamp_candidates(name: str, spec: str) -> list[tuple[str, str, int]] | None:
    """조명기구(A) LED 40W(슬림매입등) 같은 표기를 5-25-3 LED등기구 줄로."""
    compact = f"{name} {spec}".replace(" ", "")
    if "LED" not in compact.upper() or not any(w in compact for w in ("조명기구", "등기구", "LED등")):
        return None
    watts = re.search(r"(\d+(?:\.\d+)?)\s*W", f"{name} {spec}", re.I)
    if not watts:
        return []
    bucket = _bucket(float(watts.group(1)), _LED_BUCKETS)
    if bucket is None:
        return []
    for word, book in _LED_TYPES:
        if word in compact:
            return [(book, f"{bucket} W 이하", 100)]
    return []


def _conduit_size_candidates(name: str, spec: str) -> list[tuple[str, str, int]] | None:
    """나사없는전선관 E31 → 31 mm, 폴리에틸렌 전선관 CD-P 42㎜ → CD 42 mm."""
    compact = f"{name}{spec}".replace(" ", "")
    if any(w in compact for w in _ACCESSORY) or "노말밴드" in compact or "포설" in compact:
        return None
    if "나사없는전선관" in compact or "박강" in compact:
        m = re.search(r"[EC]\s*-?\s*(\d+)", spec or "") or _MM.search(spec or "")
        if m:
            return [("나사없는전선관", f"{m.group(1)} mm", 100)]
    if "폴리에틸렌전선관" in compact and "CD" in compact.upper():
        m = _MM.search(spec or "")
        if m:
            return [("CD전선관", f"CD {m.group(1)} mm", 100)]
    # 경질비닐PVC전선관 HI, 28mm → 경질비닐전선관 HI 28 mm (5-1). 가요·CD·파형·폴리에틸렌 관은 따로 처리한다.
    if (
        ("경질비닐" in compact or "PVC" in compact.upper() or "합성수지" in compact)
        and "전선관" in compact
        and not any(w in compact for w in ("가요", "CD", "파상", "파형", "폴리에틸렌", "PE"))
    ):
        m = _MM.search(spec or "")
        if m:
            label = f"HI {m.group(1)} mm" if "HI" in (spec or "").upper() else f"{m.group(1)} mm"
            return [("경질비닐전선관", label, 100)]
    return None


_CORRUGATED_SIZES = (16, 30, 50, 80, 100, 125, 150, 175, 200)
_CD_SIZES = (14, 16, 22, 28, 36, 42, 54, 70, 82, 92, 104, 125)
_TERMINAL_SIZES = (10, 16, 25, 35, 50, 70, 95, 120, 185, 240, 300, 400, 500, 630, 800, 1000, 2000)
_POLE_HEIGHTS = (5, 6, 7, 8, 9, 10, 12, 14)
_UPS_BANDS = ((3, "3 kVA 이하"), (10, "3 kVA 초과 ~ 10 kVA 이하"), (20, "10 kVA 초과 ~ 20 kVA 이하"), (30, "20 kVA 초과 ~ 30 kVA 이하"),
              (100, "30 kVA 초과 ~ 100 kVA 이하"), (250, "100 kVA 초과 ~ 250 kVA 이하"), (500, "250 kVA 초과 ~ 500 kVA 이하"))
_NOT_SWITCH = ("박스", "BOX", "센서", "제어", "콘트롤러", "컨트롤러", "차단기", "나이프", "마그넷", "소등", "수위", "FLOAT", "플로트", "리밋", "리미트")


def _first_number(text: str, pattern: str) -> float | None:
    m = re.search(pattern, text, re.I)
    return float(m.group(1)) if m else None


_ROWS_RATE = {1: 100, 2: 180, 3: 260, 4: 340, 6: 420, 8: 500, 10: 580, 12: 660, 14: 740, 16: 820}  # 4-29·4-31 해설 동시 배열
_STEEL_SIZES = (16, 22, 36, 54, 75, 100, 125, 150, 175, 200, 250, 300)
_DUCT_PART = re.compile(r"[ΦφøØ∅]\s*(\d+(?:\.\d+)?)\s*[xX×*]\s*(\d+)\s*열")


def _duct_candidates(name: str, spec: str) -> list[tuple] | None:
    """강관 포설 Φ200x2열 → 4-29 강관 설치, 전선관 포설 Φ30x1열 → 4-31 합성수지 파형관 설치.

    열 수는 해설의 동시 배열 할증(2열 180%, 3열 260%, 4열 340%, 6열 420%…), 서로 다른 관을 함께 적으면
    (기본품×열수의 합)×동시적용율÷총열수. 가로등·신호등·보안등·구내는 강관 30%, 파형관 50% 가산.
    """
    compact = f"{name}{spec}".replace(" ", "")
    if "포설" not in compact or not ("강관" in compact or "전선관" in compact) or "케이블" in compact:
        return None
    parts_raw = _DUCT_PART.findall(spec or "")
    if not parts_raw:
        return None
    steel = "강관" in compact
    parts: list[tuple[str, str, int]] = []
    for size_text, rows in parts_raw:
        size = float(size_text)
        sizes = _STEEL_SIZES if steel else _CORRUGATED_SIZES
        bucket = next((b for b in sizes if size <= b), None)
        if bucket is None:
            return []
        if steel:
            parts.append(("강관 설치", f"ø{bucket} mm 이하", int(rows)))
        else:
            parts.append(("합성수지 파형관 설치", f"{bucket} mm 이하", int(rows)))
    total = sum(n for _, _, n in parts)
    base = _ROWS_RATE.get(total)
    interpolated = False
    if base is None:  # 해설 표에 없는 열 수(5, 7, 9열 등)는 앞뒤 값을 직선으로 이어 계산하고 비고에 밝힌다.
        keys = sorted(_ROWS_RATE)
        lower = max((k for k in keys if k < total), default=None)
        upper = min((k for k in keys if k > total), default=None)
        if lower is None or upper is None:
            return []
        base = _ROWS_RATE[lower] + (_ROWS_RATE[upper] - _ROWS_RATE[lower]) * (total - lower) / (upper - lower)
        interpolated = True
    if steel and total > 10:
        return []
    extra_rate = 0.30 if steel else 0.50
    extra = any(w in compact for w in ("가로등", "신호", "보안등", "구내"))
    rate = base * (1 + extra_rate) if extra else base
    note = f"{'4-29 강관' if steel else '4-31 파형관(전선관 포설)'} {total}열 {base:g}%" + ("(표에 없어 보간)" if interpolated else "") + (f" ×가산 {round(extra_rate*100)}%" if extra else "")
    return [("__combo__", parts, rate, note)]


def _misc_candidates(name: str, spec: str) -> list[tuple] | None:
    """5-1 파상관, 5-23 스위치·콘센트, 5-26 LED 등기구, 5-27 폴, 5-42 피뢰침, 5-45 UPS, 3-38 접지, 4-36/4-37 접속·단말, 5-9-1 레이스웨이 등.

    해당 품목이 아니면 None. 부속품·자재 단품은 [] (품셈 없음).
    """
    text = f"{name} {spec}"
    compact = text.replace(" ", "")
    upper = compact.upper()
    # --- 파상형(파형) 전선관: 4-31 합성수지 파형관 설치(배전전공+보통인부, 지중). 가로등·신호등·보안등·구내는 50% 가산(해설⑧)
    if "파상형" in compact or "파형관" in compact:
        n = _first_number(spec, r"(\d+(?:\.\d+)?)\s*(?:mm|㎜)")
        size = next((s for s in _CORRUGATED_SIZES if n is not None and s >= n), None)
        if size is None:
            return []
        extra = any(w in compact for w in ("가로등", "신호", "보안등", "구내"))
        if extra:
            return [("합성수지 파형관 설치", f"{size} mm 이하", 150, "해설⑧ 가로등·신호등·보안등·구내 50% 가산")]
        return [("합성수지 파형관 설치", f"{size} mm 이하", 100, "지중 파형관 설치 기준")]
    duct = _duct_candidates(name, spec)
    if duct is not None:
        return duct
    # --- 레이스웨이: 본체만 5-9-1, 부속(END CAP·COVER 등)은 품셈 없음
    if "레이스웨이" in compact:
        if "BODY" in upper:
            dims = re.findall(r"\d+", spec or "")[-2:]
            table = {("40", "40"): "(40×40) 1,600 ㎟ 이하", ("70", "40"): "(70×40) 2,800 ㎟ 이하", ("110", "50"): "(110×50) 5,500 ㎟ 이하"}
            label = table.get(tuple(dims))
            return [("레이스웨이 설치", label, 100)] if label else []
        return []
    # --- 케이블트레이 부속·행거·볼트 류는 모체(트레이/관) 설치품에 포함
    if "트레이부속" in compact:
        return []
    # --- 스위치: 1로=단로용, 3로, 4로. N구는 구수만큼 적용(품셈에 구수 규정 없음)
    if "스위치" in compact and not any(w in upper for w in _NOT_SWITCH):
        ways = _first_number(spec, r"(\d)\s*로")
        gangs = int(_first_number(spec, r"(\d)\s*구") or 1)
        kind = {None: "단로용", 1: "단로용", 3: "3로용", 4: "4로용"}.get(int(ways) if ways else None)
        if kind is None:
            return []
        rate = 100 * gangs
        note = f"{gangs}구 = 1구 품 × {gangs}(구수 규정 없음)" if gangs > 1 else "해설 매입 기준"
        return [(f"텀블러스위치 {kind}", "일반", rate, note)]
    # --- 콘센트: 매입 기준. 접지형은 접지극부, 시스템박스용은 하이텐션
    if "콘센트" in compact and not any(w in upper for w in _NOT_SWITCH if w != "소등"):
        gangs = int(_first_number(spec, r"(\d)\s*구") or 1)
        amp = _first_number(spec, r"(\d+)\s*A")
        base_rate = 120 if "노출" in compact else 100
        rate = base_rate * gangs
        note = f"{'노출 120% ' if base_rate == 120 else ''}{gangs}구 = 1구 품 × {gangs}(구수 규정 없음)".strip() if (gangs > 1 or base_rate == 120) else "해설 매입 기준"
        if "SYSTEMBOX" in upper:
            return [("하이텐션콘센트", "2P", rate, "해설③ 시스템 박스 내 콘센트는 하이텐션 적용" + (f" {gangs}구" if gangs > 1 else ""))]
        a = 15 if amp is None or amp <= 15 else (20 if amp <= 20 else 30)
        if "접지" in compact:
            return [(f"콘센트 접지극부 {a} A", "2P", rate, note)]
        return [("콘센트 15 A", "2P", rate, note)] if a == 15 else []
    # --- LED 가로등/터널/보안등기구
    if "LED" in upper and any(w in compact for w in ("가로등기구", "터널등기구", "보안등")) and "설치" not in name:
        watts = _first_number(text, r"(\d+(?:\.\d+)?)\s*W")
        if watts is None:
            return []
        if "터널" in compact:
            band = next((b for b in (100, 150, 200, 250) if watts <= b), None)
            return [("LED 터널등기구 설치", f"{band} W 이하", 100)] if band else []
        if "보안등" in compact:
            band = next((b for b in (50, 100) if watts <= b), None)
            return [("LED 보안등기구 설치", f"{band} W 이하", 100)] if band else []
        band = next((b for b in (100, 150, 200, 250) if watts <= b), None)
        return [("LED 가로등기구 설치", f"{band} W 이하", 100)] if band else []
    # --- 가로등철주(Pole): 5-27 인력 설치
    if "가로등철주" in compact or "가로등주" in compact and "기계경비" not in compact:
        height = _first_number(text, r"(\d+(?:\.\d+)?)\s*[mM](?!m)")
        band = next((h for h in _POLE_HEIGHTS if height is not None and height <= h), None)
        if band is None:
            return []
        lamps = "2등용" if re.search(r"2\s*등", text) else "1등용"
        return [("PoleLight", f"{band} m 이하 {lamps} 인력", 100, "전기 5-27 Pole 인력설치(1등용 기준)")]
    # --- 관로구방수 ∮30·∮40: 2-18 삽입형 관로구 방수장치(200 mm 이하)
    if "관로구" in compact and "방수" in compact:
        n = _first_number(spec, r"[∮ΦφøØ∅]\s*(\d+(?:\.\d+)?)") or _first_number(spec, r"(\d+(?:\.\d+)?)\s*(?:mm|㎜)")
        if n is not None and n > 200:
            return []
        return [("삽입형 관로구방수장치", "200 mm 이하", 100, "삽입형 관로구 방수장치")]
    # --- 접지 (전기 3-38)
    # 전기용 나동선 BC 50㎟ → 접지선 매설 95 ㎟ 이하(0.008)
    if "나동선" in compact or ("BC" in upper and "동선" in compact):
        n = _first_number(spec or text, r"(\d+(?:\.\d+)?)\s*(?:㎟|mm2|mm²|sq)")
        if n is None:
            return []
        band = next((b for b in (10, 35, 95, 150, 240) if n <= b), None)
        return [("접지선매설", f"{band} ㎟ 이하" if band else "240 ㎟ 초과", 100, "접지선 매설")]
    if "접지봉" in compact and "천공" not in compact:
        if "커넥터" in compact:  # 접지봉커넥터 = 접지선 접속 볼트체결형(3-38)
            return [("접지선접속및단자설치", "볼트체결형", 100, "접지선 접속 볼트체결형")]
        if any(w in compact for w in _ACCESSORY):  # 그 밖의 부속은 접지봉 설치품에 포함
            return []
        rods = re.search(r"(\d+)\s*본", text)
        count = int(rods.group(1)) if rods else 1
        if count in (2, 3):  # 접지봉-3본(1조) = 3본 연결 품(원표에 줄이 따로 있음)
            return [("접지봉", f"지하 0.75 m 길이 1~2 m {count}본 연결", 100)]
        return [("접지봉", "지하 0.75 m 길이 1~2 m 1본", 100)]
    # --- 슬리브 접속(C TYPE SLEEVE 50㎟): 3-38 접속 및 단자설치 압축(압축슬리브 사용) 0.097
    if "슬리브접속" in compact or ("SLEEVE" in upper and "TYPE" in upper):
        return [("접지선접속및단자설치", "압축(압축슬리브 사용)", 100, "압축슬리브 사용")]
    # --- 수위스위치(플로트): 6-2 계기 설치 일반각종 0.22, 인부는 내선전공으로 적용
    if "수위스위치" in compact:
        return [("수위스위치", "일반", 100, "계기 설치 일반각종 적용")]
    if "접지" in compact and "단자함" in compact:
        return [("접지시험단자함", "일반", 100)]
    # --- 피뢰침(일반형, 높이 1~1.5 m)
    if "피뢰침" in compact:
        return [("피뢰침 설치", "길이 7.5 m 이하", 100)]
    # --- UPS
    if "무정전" in compact or "UPS" in upper:
        kva = _first_number(text, r"(\d+(?:\.\d+)?)\s*kVA")
        label = next((l for b, l in _UPS_BANDS if kva is not None and kva <= b), None)
        return [("무정전 전원장치(UPS, CVCF) 설치", label, 100)] if label else []
    # --- 케이블 헤드 지지금구
    if "케이블헤드지지금구" in compact:
        return [("케이블 헤드 지지금구 설치", "지지금구 설치", 100)]
    # --- 압착단자·러그: 4-37 전력케이블 단말처리 1개소 (러그 1개 = 1개소로 계산)
    if "러그" in compact or "압착단자" in compact:
        n = _first_number(spec, r"(\d+(?:\.\d+)?)\s*(?:㎟|mm2|mm²|sq)")
        size = next((s for s in _TERMINAL_SIZES if n is not None and n <= s), None)
        if size is None or size < 16:  # 4-37 1 kV 이하 1C 의 10 ㎟ 이하 칸은 원문에 값이 없음("-")
            return []
        label = f"{size:,} ㎟ 이하" if size >= 1000 else f"{size} ㎟ 이하"
        return [("전력케이블 단말처리", label, 100, "단말처리 1개소당 기준(러그 1개=1개소)")]
    # --- 터널 관통형 커넥터(커버 포함): 5-11-1 은 개소당 3개 기준 → 1개 = 1/3, 커버 동시설치 5% 가산
    if "분기접속제" in compact and "관통형" in compact:
        return [("저압 케이블 관통형 커넥터 설치", "터널 내부", 35, "개소당 3개 기준 1/3, 커버 동시 5% 가산(개당)")]
    # --- 직선접속제: 4-36 전력케이블 직선접속
    if "직선접속" in compact:
        n = _first_number(spec, r"(\d+(?:\.\d+)?)\s*(?:㎟|mm2|mm²|sq)")
        size = next((s for s in _TERMINAL_SIZES if n is not None and n <= s), None)
        if size is None or size < 16:  # 4-36 도 1 kV 이하 1C 의 10 ㎟ 이하 칸이 "-" (해설④: 10㎟ 이하는 5-13 준용)
            return []
        label = f"{size:,} ㎟ 이하" if size >= 1000 else f"{size} ㎟ 이하"
        return [("전력케이블 직선접속", label, 100)]
    return None


def electric_candidates(name: str, spec: str) -> list[tuple]:
    text = f"{name} {spec}"
    misc = _misc_candidates(name, spec)
    if misc is not None:
        return misc
    boxed = _box_candidates(name, spec)
    if boxed is not None:
        return boxed
    lamp = _lamp_candidates(name, spec)
    if lamp is not None:
        return lamp
    sized = _conduit_size_candidates(name, spec)
    if sized is not None:
        return sized
    if "가요" in text and "전선관" in text and not any(w in text for w in _ACCESSORY):
        mm = _MM.search(spec or "")
        if mm:
            return [("금속제가요전선관", f"{mm.group(1)} mm", 100)]
    upper = text.upper()
    # 「트레이용 제어용케이블」의 트레이는 포설 장소 설명이지 트레이 자재가 아니다.
    guard_text = text.replace("트레이용", "")
    if any(word in guard_text for word in ("전선관", "접지봉", "단자", "덕트", "트레이", "커넥터")):
        return []
    areas = _AREA.findall(text)
    if not areas:
        return []
    area = float(areas[-1])
    cable = "케이블" in text or re.search(r"\bF?-?C(?:V|VV|CV)", upper) is not None
    cores_match = _CORES.search(spec or "") or _CORES.search(name or "")
    cores = int(cores_match.group(1)) if cores_match else None
    out: list[tuple[str, str, int]] = []
    is_control = any(word in upper for word in ("CVV", "CCV", "제어"))
    if is_control:
        if cores is None:
            return []
        for label in ("제어용케이블",):
            b = _bucket(area, CONTROL_BUCKETS)
            if b is not None:
                out.append((label, f"{_fmt(b)} ㎟ 이하 {cores}C", 100))
        return out
    if cable:
        if cores not in (None, 1):
            if area <= 10:  # 5-11 해설②: 10㎟ 이하는 제어용케이블 설치 준용
                b = _bucket(area, CONTROL_BUCKETS)
                return [("제어용케이블", f"{_fmt(b)} ㎟ 이하 {cores}C", 100)] if b is not None and cores <= 50 else []
            rate = {2: 140, 3: 200, 4: 260}.get(cores)  # 5-11 해설④ 2심 140%, 3심 200%, 4심 260%
            if rate is None:
                return []
            b = _bucket(area, POWER_SIZES)
            if b is None:
                return []
            if area <= 16:
                return [("전력케이블", "16 ㎟ 이하×1C", rate)]
            return [("전력케이블", f"{_fmt(b)} ㎟×1C", rate)]
        if area <= 16:
            out.append(("전력케이블", "16 ㎟ 이하×1C", 100))
        else:
            b = _bucket(area, POWER_SIZES)
            if b is not None:
                out.append(("전력케이블", f"{_fmt(b)} ㎟×1C", 100))
        return out
    if "전선" in text or re.search(r"\b(IV|HIV|HFIX|GV|F-GV|HFIX)\b", upper) or "절연" in text:
        b = _bucket(area, INDOOR_BUCKETS)
        if b is not None:
            out.append(("옥내배선", f"{_fmt(b)} ㎟ 이하", 100))
    return out
