"""문자로 적힌 수량·단가("1,200원", "약 3", "12㎡")를 숫자로 정리한다."""

from __future__ import annotations

import re
from typing import Any

_NUMBER = re.compile(r"-?\d[\d,]*(?:\.\d+)?|-?\.\d+")
NOTICES: list[str] = []


def parse_number(value: Any) -> tuple[float | None, bool]:
    """(숫자, 글자를 걷어 냈는지). 읽을 숫자가 없으면 (None, False)."""
    if isinstance(value, bool):
        return None, False
    if isinstance(value, (int, float)):
        return float(value), False
    text = str(value or "").strip()
    if not text:
        return None, False
    try:
        return float(text.replace(",", "")), False
    except ValueError:
        pass
    found = _NUMBER.search(text)
    if not found:
        return None, False
    try:
        return float(found.group(0).replace(",", "")), True
    except ValueError:
        return None, False


def clean_price(value: Any, where: str = "") -> Any:
    """단가 칸 값. 문자로 적힌 숫자는 숫자로 바꾸고, 안내를 남긴다. 숫자가 아닌 문자는 그대로 둔다."""
    if value is None or isinstance(value, (int, float)):
        return value
    number, cleaned = parse_number(value)
    if number is None:
        return value
    if cleaned:
        NOTICES.append(f"{where} 단가 「{value}」 → {number:g} 로 읽었습니다.".strip())
    return number
