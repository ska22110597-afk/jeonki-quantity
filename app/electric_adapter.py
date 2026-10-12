"""전기 전선·케이블: 단가대비표의 긴 품명·규격을 표준품셈 5-10/5-11/5-13 줄로 바꿔 찾는다.

예) 접지용비닐절연전선(F-GV) 16㎟, 450/750V 내열비닐절연전선 HFIX 2.5㎟ → 옥내배선(5-10).
    F-CVV-S 2Cx1.5㎟ → 제어용케이블(5-13).
다심 전력케이블(F-CV 2C×4㎟ 등)은 5-11 표가 1C 기준이라 임의로 환산하지 않고 비워 둔다.
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


def electric_candidates(name: str, spec: str) -> list[tuple[str, str, int]]:
    text = f"{name} {spec}"
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
            return []
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
