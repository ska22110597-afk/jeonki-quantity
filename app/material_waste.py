"""자재 수량 할증(재료의 할증률). 일위대가 자재 수량 칸에 곱하고, 비고에 근거를 적는다.

통신: 정보통신공사 표준품셈 1-1-6 「재료의 할증률 및 철거손실률」 표(옥내/옥외 구분).
전기: 이 프로그램은 지금까지 전선관·전선·케이블 10% 를 써 왔다. 전기 표준품셈 총칙 표를 확인하면 이 파일만 고치면 된다.
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
    # 전기: 기존 규칙(전선관·전선·케이블·m 단위 10%) 유지
    if is_conduit or is_cable or is_meter:
        return 1.10, "전기 자재할증 10% (전선관·전선·케이블)"
    return 1.0, ""
