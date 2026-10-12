"""단가대비표의 통신 자재 이름·규격을 표준품셈 원표 항목으로 바꿔 준다.

예) UTP 케이블 Cat.5E 0.5mm 4P → 꼬임케이블 4P, 광케이블(옥내용) SM 12c → 광케이블 12코어 이하.
원표에 없는 자재(부속품·커버·분배기 등)는 돌려주지 않는다. 이름이 비슷하다고 임의로 맞추지 않는다.
돌려주는 값: (원표 명칭, 원표 규격, 할증%) 목록. 호출하는 쪽에서 실제 원표 행에 있을 때만 사용한다.
"""

from __future__ import annotations

import re

Candidate = tuple[str, str, int]

_PAIR = re.compile(r"(\d+)\s*P(?![A-Za-z])", re.I)
_CORE = re.compile(r"(\d+)\s*(?:c|core|코어)\b", re.I)
_COAX = re.compile(r"(\d+)\s*C\b", re.I)
_WH_D = re.compile(r"(\d+(?:\.\d+)?)\s*[x×*]\s*(\d+(?:\.\d+)?)\s*[x×*]\s*(\d+(?:\.\d+)?)", re.I)

TRAY_WIDTHS = (200, 300, 400, 500, 600, 800, 1000)
BOX_AREAS = ((500, 10), (1800, 13), (5250, 15), (11000, 15), (18200, 18), (27200, 25))


def _compact(text: str) -> str:
    return "".join(str(text or "").split()).lower()


def _fmt(n: int) -> str:
    return f"{n:,}"


def telecom_candidates(name: str, spec: str) -> list[Candidate]:
    n = _compact(name)
    s = str(spec or "")
    sc = _compact(s)
    out: list[Candidate] = []

    if "커버" in n or "부속" in n:
        return []

    if any(k in n for k in ("아우트렛박스", "아웃렛박스", "아울렛박스")) and "스위치" not in n:
        return [("박스", "Outlet Box", 100)]
    if "스위치박스" in n:
        m = re.search(r"(\d+)\s*개용", s)
        count = int(m.group(1)) if m else 1
        return [("박스", "Switch Box (3개용이하)" if count <= 3 else "Switch Box (4개용이상)", 100)]
    if "콘크리트박스" in n:
        return [("박스", "Concrete Box", 100)]
    if "정션박스" in n or "접속박스" in n:
        return [("박스", "연결용 박스", 100)]

    if any(k in n for k in ("단자함", "분배기함", "통신함체")):
        m = _WH_D.search(s)
        if m:
            area = float(m.group(1)) * float(m.group(2)) / 100.0
            depth = float(m.group(3)) / 10.0
            for limit, dmax in BOX_AREAS:
                if area <= limit and depth <= dmax:
                    return [("단자함", f"단자함 단면적 {_fmt(limit)}cm2 이하(깊이 {dmax}cm 이하)", 100)]
        return []

    if "utp" in n or "stp" in n or "ftp" in n or "꼬임케이블" in n or "랜케이블" in n or "lan케이블" in n:
        m = _PAIR.search(s)
        if not m:
            return []
        pairs = int(m.group(1))
        for bucket in (4, 25, 50, 100):
            if pairs <= bucket:
                kind = "STP케이블" if "stp" in n else "FTP케이블" if "ftp" in n else "UTP케이블"
                return [(kind, f"{bucket}P", 100)]
        return []

    if "광케이블" in n or "광섬유케이블" in n:
        m = _CORE.search(s)
        if not m:
            return []
        cores = int(m.group(1))
        if cores <= 12:
            return [("광케이블", "12C", 100)]
        if cores <= 24:
            return [("광케이블", "24C", 100)]
        return []

    if "동축케이블" in n or "동축" in n and "케이블" in n:
        m = _COAX.search(s)
        if m and m.group(1) in ("5", "7", "10"):
            return [("동축케이블", f"{m.group(1)}C", 100)]
        return []

    if "모듈라짹" in n or "모듈라잭" in n or "모듈러잭" in n:
        m = re.search(r"(\d)\s*구", s)
        gang = int(m.group(1)) if m else 1
        gang = min(max(gang, 1), 4)
        return [("모듈라짹", "1구", 100 + 20 * (gang - 1))]

    if n.startswith("fdf") or "광분배반" in n:
        return [("광분배함 및 성단", "광분배반(FDF)", 100)]

    if "국내성단" in n:
        m = _PAIR.search(s)
        if m and "utp" in sc:
            pairs = int(m.group(1))
            if pairs == 25:
                return [("Patch Panel 및 성단", "성단 110 Block 25P", 100)]
            if pairs == 4:
                return [("Patch Panel 및 성단", "성단 110 Block 4P", 100)]
        return []

    if "회선시험" in n and "utp" in sc:
        return [("Patch Panel 및 성단", "회선시험", 100)]

    if "tray" in n or "트레이" in n:
        if "straight" in sc:
            m = re.search(r"w\s*(\d+)", s, re.I)
            if m:
                width = int(m.group(1))
                for bucket in TRAY_WIDTHS:
                    if width <= bucket:
                        return [("케이블트레이", f"폭 {_fmt(bucket)}mm 이하 철재", 100)]
        return []
    return out
