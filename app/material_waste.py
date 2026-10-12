"""자재 수량 할증(재료의 할증률). 일위대가 자재 수량 칸에 곱하고, 비고에 근거를 적는다.

통신: 정보통신공사 표준품셈 1-1-6 「재료의 할증률 및 철거손실률」 표(옥내/옥외 구분).
전기: 전기공사 표준품셈 1-6 「전기재료의 할증률 및 철거손실률」(옥내전선 10, 옥외전선 5, 케이블 옥내 5·옥외 3, 전선관 옥내 10·옥외 5, 트레이·덕트 5, 파형전선관 3).
"""

from __future__ import annotations

from typing import Any

from app.pumsam_aliases import split_place_name

OUTDOOR_WORDS = ("옥외", "지중", "직매", "관로", "매설", "파형", "ELP", "맨홀", "핸드홀", "가공")


def _outdoor(name: Any, spec: Any) -> bool:
    text = f"{name or ''} {spec or ''}"
    suffix = split_place_name(name)[1]
    return suffix in ("_지중", "_직매") or any(word in text.upper() for word in OUTDOOR_WORDS)


def _pct(value: float) -> str:
    return f"{value:g}"


def material_waste(
    name: Any,
    spec: Any,
    unit: Any,
    discipline: str,
    *,
    is_conduit: bool,
    is_cable: bool,
) -> tuple[float, str]:
    """(자재 수량 배율, 비고 문구). 할증이 없으면 (1.0, "")."""
    text = f"{name or ''} {spec or ''}"
    compact = text.replace(" ", "")
    if any(w in compact.upper() for w in ("부속", "철물", "표지", "시트", "JUMPER", "CLAMP", "ANGLE", "BRACKET", "커버", "COVER")):
        return 1.0, ""
    outdoor = _outdoor(name, spec)
    is_meter = str(unit or "").strip().lower() in {"m", "ｍ", "미터"}
    if discipline == "통신":
        base = "통신 1-1-6"
        if any(w in compact for w in ("트레이", "덕트", "레이스웨이", "케이블랙")):
            return 1.05, f"{base} 트레이·덕트 자재할증 5%"
        if is_conduit:
            if outdoor:
                return 1.03, f"{base} 옥외 합성수지관 자재할증 3%"
            return 1.10, f"{base} 옥내 전선관 자재할증 10%"
        if is_cable or (is_meter and any(w in compact for w in ("케이블", "UTP", "STP", "동축", "광", "통신선"))):
            if "직매" in compact or "관로" in compact or "지중" in compact:
                return 1.03, f"{base} 지하관로·직매 케이블 자재할증 3%"
            if outdoor:
                return 1.03, f"{base} 옥외 케이블 자재할증 3%"
            return 1.075, f"{base} 옥내 케이블 자재할증 7.5%"
        if "열수축관" in compact or "점퍼선" in compact:
            return 1.05, f"{base} 자재할증 5%"
        return 1.0, ""
    # 전기: 표준품셈 1-6 「전기재료의 할증률 및 철거손실률」
    if any(w in compact.upper() for w in ("트레이", "덕트", "레이스웨이", "케이블랙", "TRAY")):
        return 1.05, "전기 1-6 트레이·덕트 자재할증 5%"
    if is_conduit:
        if any(w in compact.upper() for w in ("파형", "파상", "ELP")):
            return 1.03, "전기 1-6 합성수지 파형전선관 자재할증 3%"
        if outdoor:
            return 1.05, "전기 1-6 옥외 전선관 자재할증 5%"
        return 1.10, "전기 1-6 옥내 전선관 자재할증 10%"
    if is_cable or is_meter:
        cable = "케이블" in compact or any(w in compact.upper() for w in ("CV", "CCV", "VVF", "VVR", "UTP", "STP"))
        if cable:
            return (1.03, "전기 1-6 옥외 케이블 자재할증 3%") if outdoor else (1.05, "전기 1-6 옥내 케이블 자재할증 5%")
        if "전선" in compact or any(w in compact.upper() for w in ("IV", "HIV", "HFIX", "GV", "연동선")):
            return (1.05, "전기 1-6 옥외 전선 자재할증 5%") if outdoor else (1.10, "전기 1-6 옥내 전선 자재할증 10%")
    return 1.0, ""
