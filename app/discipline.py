"""전기·통신 파트. 표준품셈 파일을 나눠 불러온다."""

from __future__ import annotations

ELECTRIC = "전기"
TELECOM = "통신"
DISCIPLINES = (ELECTRIC, TELECOM)

LEGACY_PUMSAM_FILENAME = "표준품셈.xlsx"


def normalize_discipline(value: str | None) -> str:
    text = str(value or "").replace(" ", "").replace("_", "")
    if "통신" in text:
        return TELECOM
    return ELECTRIC


def pumsam_filename(discipline: str | None = None) -> str:
    return f"{normalize_discipline(discipline)}_표준품셈.xlsx"


def default_labor_name(discipline: str | None = None) -> str:
    if normalize_discipline(discipline) == TELECOM:
        return "통신내선공"
    return "내선전공"


def discipline_label(discipline: str | None = None) -> str:
    disc = normalize_discipline(discipline)
    return f"{disc} 표준품셈"
