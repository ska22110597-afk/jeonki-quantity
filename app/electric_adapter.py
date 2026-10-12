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
    return None


_CD_SIZES = (14, 16, 22, 28, 36, 42, 54, 70, 82, 92, 104, 125)
_TERMINAL_SIZES = (10, 16, 25, 35, 50, 70, 95, 120, 185, 240, 300, 400, 500, 630, 800, 1000, 2000)
_POLE_HEIGHTS = (5, 6, 7, 8, 9, 10, 12, 14)
_UPS_BANDS = ((3, "3 kVA 이하"), (10, "3 kVA 초과 ~ 10 kVA 이하"), (20, "10 kVA 초과 ~ 20 kVA 이하"), (30, "20 kVA 초과 ~ 30 kVA 이하"),
              (100, "30 kVA 초과 ~ 100 kVA 이하"), (250, "100 kVA 초과 ~ 250 kVA 이하"), (500, "250 kVA 초과 ~ 500 kVA 이하"))
_NOT_SWITCH = ("박스", "BOX", "센서", "제어", "콘트롤러", "컨트롤러", "차단기", "나이프", "마그넷", "소등")


def _first_number(text: str, pattern: str) -> float | None:
    m = re.search(pattern, text, re.I)
    return float(m.group(1)) if m else None


def _misc_candidates(name: str, spec: str) -> list[tuple] | None:
    """5-1 파상관, 5-23 스위치·콘센트, 5-26 LED 등기구, 5-27 폴, 5-42 피뢰침, 5-45 UPS, 3-38 접지, 4-36/4-37 접속·단말, 5-9-1 레이스웨이 등.

    해당 품목이 아니면 None. 부속품·자재 단품은 [] (품셈 없음).
    """
    text = f"{name} {spec}"
    compact = text.replace(" ", "")
    upper = compact.upper()
    # --- 파상형(파형) 폴리에틸렌 전선관 → 5-1 CD(PE) 전선관, 같거나 큰 호칭
    if "파상형" in compact and "전선관" in compact:
        n = _first_number(spec, r"(\d+(?:\.\d+)?)\s*(?:mm|㎜)")
        size = next((s for s in _CD_SIZES if n is not None and s >= n), None)
        return [("CD전선관", f"CD {size} mm", 100, "파상형관은 CD관 호칭 기준 준용")] if size else []
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
    # --- 접지
    if "접지봉" in compact and "천공" not in compact:
        return [("접지봉", "지하 0.75 m 길이 1~2 m 1본", 100)]
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
        if size is None:
            return []
        label = f"{size:,} ㎟ 이하" if size >= 1000 else f"{size} ㎟ 이하"
        return [("전력케이블 단말처리", label, 100, "단말처리 1개소당 기준(러그 1개=1개소)")]
    # --- 직선접속제: 4-36 전력케이블 직선접속
    if "직선접속" in compact:
        n = _first_number(spec, r"(\d+(?:\.\d+)?)\s*(?:㎟|mm2|mm²|sq)")
        size = next((s for s in _TERMINAL_SIZES if n is not None and n <= s), None)
        if size is None:
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
    if any(word in text for word in ("전선관", "접지봉", "단자", "덕트", "트레이", "커넥터")):
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
