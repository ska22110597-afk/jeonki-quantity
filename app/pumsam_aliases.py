"""단가대비표와 품셈이 다른 말로 같은 품을 적을 때 쓰는 이름 묶음.

후강전선관 = 강제전선관처럼 현장에서 섞어 쓰는 말을 먼저 풀어 둔다.
_지중 / _노출 / _매입 / _직매 접미사는 묶지 않고 그대로 구분한다.
"""

from __future__ import annotations

from typing import Any

from app.estimate_parse import lookup_key

PLACE_SUFFIXES = ("_지중", "_노출", "_매입", "_직매", "_천장", "_벽면")

# 한 묶음은 같은 표준품셈을 쓰는 이름. 다른 자재와 섞지 않는다.
NAME_ALIAS_GROUPS: tuple[tuple[str, ...], ...] = (
    (
        "강제전선관",
        "후강전선관",
        "아연도강전선관",
        "아연도전선관",
        "후강관",
        "강제관",
        "아연도관",
        "G관",
        "G전선관",
    ),
    (
        "경질비닐전선관",
        "합성수지전선관",
        "합성수지제전선관",
        "HIVP전선관",
        "HI전선관",
        "HI관",
        "HI-P전선관",
        "PVC전선관",
        "경질비닐관",
        "합성수지관",
        "HIVP관",
    ),
    (
        "CD전선관",
        "CD관",
        "폴리에틸렌전선관",
        "PE전선관",
        "PE관",
        "폴리에틸렌관",
    ),
    (
        "금속제가요전선관",
        "가요전선관",
        "EF전선관",
        "후렉서블전선관",
        "플렉시블전선관",
        "후렉시블전선관",
        "금속제가요관",
        "가요관",
    ),
    (
        "박강전선관",
        "나사없는전선관",
        "E전선관",
        "C관",
        "박강관",
    ),
    (
        "옥내배선",
        "IV전선",
        "HIV전선",
        "HFIX전선",
        "600V IV",
        "600V HIV",
        "600V HFIX",
        "IV",
        "HIV",
        "HFIX",
    ),
    (
        "전력케이블",
        "CV케이블",
        "VV케이블",
        "F-CV케이블",
        "600V CV",
        "CV",
        "FCV케이블",
    ),
    (
        "제어용케이블",
        "CVV케이블",
        "CCV케이블",
        "CVV",
    ),
    (
        "UTP케이블",
        "꼬임케이블",
        "LAN케이블",
        "UTP",
    ),
    (
        "광케이블",
        "광섬유케이블",
        "광섬유",
    ),
    (
        "모선덕트",
        "버스덕트",
        "버스웨이",
    ),
    (
        "케이블트레이",
        "케이블 트레이",
        "Cable Tray",
    ),
    (
        "배선용단자함",
        "전기단자함",
    ),
    (
        "배선용차단기",
        "MCCB",
        "배선차단기",
        "몰드차단기",
    ),
    (
        "누전차단기",
        "ELB",
        "누전전로차단기",
    ),
    (
        "세대분전반",
        "주택용분전반",
    ),
    (
        "콘크리트박스",
        "아웃렛박스",
        "Outlet박스",
        "아웃렛 박스",
    ),
)


def split_place_name(name: Any) -> tuple[str, str]:
    text = str(name or "").strip()
    for suffix in PLACE_SUFFIXES:
        if text.endswith(suffix):
            return text[: -len(suffix)], suffix
    return text, ""


def _group_for(base: str) -> tuple[str, ...]:
    key = lookup_key(base, "")
    if not key:
        return (base,)
    for group in NAME_ALIAS_GROUPS:
        if key in {lookup_key(item, "") for item in group}:
            return group
    return (base,)


def alias_names(name: Any) -> list[str]:
    """원래 이름을 앞에 두고, 같은 품 묶음의 다른 이름을 이어서 돌려준다."""
    text = str(name or "").strip()
    if not text:
        return []
    base, suffix = split_place_name(text)
    names = [text]
    seen = {lookup_key(text, "")}
    for alt in _group_for(base):
        candidate = f"{alt}{suffix}"
        key = lookup_key(candidate, "")
        if key and key not in seen:
            seen.add(key)
            names.append(candidate)
    return names
