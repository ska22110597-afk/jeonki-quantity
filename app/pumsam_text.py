"""품셈 규격·단위 글자를 원표 뜻대로 고친다. 〃·제곱·하이픈·셈 숫자 누설."""

from __future__ import annotations

import re
from typing import Any

DITTO = "〃"
DITTO_CHARS = str.maketrans(
    {
        "＂": DITTO,
        "″": DITTO,
        "〻": DITTO,
        '"': DITTO,
    }
)
QUALIFIERS = ("이하", "초과", "미만", "내외", "이상", "단심")
UNIT_HINTS = ("㎟", "㎠", "㎢", "㎥", "㎜", "mm", "cm", "m", "km", "kV", "kVA", "A", "P", "C")
TRAILING_UNITS = (
    "급유구간",
    "개소",
    "km",
    "㎞",
    "mm",
    "㎜",
    "m",
    "ｍ",
    "M",
    "식",
    "개",
    "조",
    "톤",
    "본",
    "대",
    "선",
    "회",
    "장",
    "기",
)
LEAKED_QTY = re.compile(r"\s+\d{1,4}\.\d{2,}$")
AREA_MM = re.compile(
    r"(?i)(?:㎜|mm)\s*(?:²|2|\^\s*2)|mm²|㎜²"
)
BROKEN_CM2 = re.compile(
    r"(?:\ue0e7\ue0f1\ue035||㎝\s*²|cm\s*²|cm2)"
)
BROKEN_CM = re.compile(r"(?:\ue0e7\ue0f1|)")
SQUARE_M = re.compile(r"(?i)(?<![k㎝])m\s*(?:²|2|\^\s*2)|㎡")
THIN_SPACES = re.compile(r"[\u00a0\u2000-\u200b\ufeff]+")
MULTI_SPACE = re.compile(r"\s+")
LEADING_DITTO = re.compile(rf"^{DITTO}\s*")
KIND_SPLIT = re.compile(r"(Ø|∅|\d)")
DASH_ONLY = re.compile(r"^[-－—–]$")


def normalize_spaces(text: str) -> str:
    text = THIN_SPACES.sub(" ", str(text or "").replace("\n", " "))
    return MULTI_SPACE.sub(" ", text).strip()


def unify_square_display(text: str) -> str:
    """단면적 제곱밀리미터는 ㎟ 로 통일한다. 제곱센티는 ㎠."""
    text = str(text or "")
    text = BROKEN_CM2.sub("㎠", text)
    text = BROKEN_CM.sub("㎝", text)
    text = AREA_MM.sub("㎟", text)
    text = text.replace("mm²", "㎟").replace("㎜²", "㎟")
    text = re.sub(r"(?i)(?<=\d)\s*mm2", " ㎟", text)
    text = re.sub(r"(?i)(?<=\d)mm2", "㎟", text)
    text = re.sub(r"㎟이(?!하)", "㎟ 이하", text)
    text = re.sub(r"(\d)\s*㎟", r"\1 ㎟", text)
    text = re.sub(r"㎟(?=\S)", "㎟ ", text)
    return MULTI_SPACE.sub(" ", text).strip()


def unify_length_display(text: str) -> str:
    text = (
        str(text or "")
        .replace("ｍ", "m")
        .replace("Ｍ", "M")
        .replace("㎞", "km")
        .replace("㎜", "mm")
    )
    return text


def display_spec(text: str) -> str:
    text = unify_square_display(unify_length_display(normalize_spaces(text)))
    text = text.replace("''", DITTO).replace("´´", DITTO)
    text = text.translate(DITTO_CHARS)
    return MULTI_SPACE.sub(" ", text).strip()


def display_unit(text: str) -> str:
    unit = display_spec(text)
    unit = re.sub(r"^단위\s*[:：]\s*", "", unit)
    unit = unit.split(",")[0].strip()
    unit = re.sub(r"적용직종.*", "", unit).strip()
    if unit in {"ｍ", "M"}:
        return "m"
    if unit.lower() == "10m":
        return "10m"
    return unit


def lookup_measure(text: str) -> str:
    """검색키용. ㎟·mm2·mm² 를 같은 글자로 본다."""
    text = unify_square_display(unify_length_display(str(text or "")))
    text = text.replace("㎟", "mm2").replace("㎠", "cm2").replace("㎝", "cm")
    return text


def _copy_qualifier(prev: str) -> str:
    for word in QUALIFIERS:
        if word in prev:
            return word
    return ""


def _copy_unit_token(prev: str, hint: str = "") -> str:
    if hint:
        return hint
    for token in UNIT_HINTS:
        if token and token in prev:
            return token
    return ""


def apply_unit_hint(spec: str, hint: str) -> str:
    """규격(㎟) 표는 숫자 뒤에 ㎟ 를 붙인다. 전압 kV 와 섞지 않는다."""
    spec = display_spec(spec)
    hint = display_spec(hint)
    if not hint or not spec or hint in spec:
        return spec
    if hint == "㎟":
        replaced = re.sub(
            r"(\d+(?:[.,]\d+)?)\s*(이하|초과|미만|이상)?\s*$",
            lambda match: f"{match.group(1)} ㎟ {match.group(2) or ''}".strip(),
            spec,
        )
        if replaced != spec:
            return replaced
    elif re.search(rf"\d+\s*{re.escape(hint)}\b", spec):
        return spec
    else:
        replaced = re.sub(
            r"(\d+(?:[.,]\d+)?)\s*(이하|초과|미만|이상)?\s*$",
            lambda match: f"{match.group(1)} {hint} {match.group(2) or ''}".strip(),
            spec,
        )
        if replaced != spec:
            return replaced
    return spec


def _prefix_before_size(prev: str) -> str:
    matched = KIND_SPLIT.search(prev)
    if not matched or matched.start() == 0:
        return ""
    return prev[: matched.start()].strip()


def expand_ditto(spec: str, prev: str, unit_hint: str = "") -> str:
    """〃 는 위 칸과 같다. 단위(㎟)·이하·종류 앞말을 이어 받는다."""
    spec = display_spec(spec)
    prev = display_spec(prev)
    if DITTO not in spec:
        return spec
    if not prev:
        spec = spec.replace(DITTO, "").strip()
        if unit_hint and unit_hint not in spec:
            spec = f"{spec} {unit_hint}".strip()
        return MULTI_SPACE.sub(" ", spec)

    if spec.strip() == DITTO:
        return prev

    qualifier = _copy_qualifier(prev)
    unit_token = _copy_unit_token(prev, unit_hint)

    if spec.startswith(DITTO):
        rest = LEADING_DITTO.sub("", spec).strip()
        prefix = _prefix_before_size(prev)
        pieces = [prefix, rest]
        spec = " ".join(part for part in pieces if part).strip()
    else:
        spec = spec.replace(DITTO, "").strip()

    if unit_token and unit_token not in spec:
        spec = f"{spec} {unit_token}".strip()
    if qualifier and qualifier not in spec:
        spec = f"{spec} {qualifier}".strip()
    return MULTI_SPACE.sub(" ", spec)


def strip_leaked_qty(spec: str) -> str:
    """규격 끝에 붙은 품셈 숫자(4.21 같은 소수 둘째자리)를 뗀다."""
    spec = display_spec(spec)
    return LEAKED_QTY.sub("", spec).strip()


def strip_trailing_unit_dash(spec: str, unit: str = "") -> tuple[str, str]:
    """규격에 섞인 단위·빈칸 표시 '-' 를 떼고, 단위 칸으로 돌린다."""
    spec = display_spec(spec)
    unit = display_unit(unit)
    if DASH_ONLY.match(spec):
        return "", unit
    spec = re.sub(r"\s*[-－—–]\s*$", "", spec).strip()
    spec = re.sub(r"\s+\d{1,4}\.\d{2,}\s*$", "", spec).strip()
    pattern = r"\s+(" + "|".join(re.escape(item) for item in TRAILING_UNITS) + r")\s*$"
    matched = re.search(pattern, spec, flags=re.I)
    if matched:
        found = display_unit(matched.group(1))
        if found.lower() in {"mm", "㎟", "㎠", "cm", "㎝", "㎥"}:
            return spec, unit
        spec = spec[: matched.start()].strip()
        if not unit or unit in {"식", "개"} or found.lower() in {"m", "km", "10m", "톤", "조", "선", "급유구간"}:
            unit = found or unit
    spec = re.sub(r"\s*[-－—–]\s*$", "", spec).strip()
    spec = re.sub(r"\s+\d{1,4}\.\d{2,}\s*$", "", spec).strip()
    return spec, unit


CABLE_AREA_AS_KV = re.compile(
    r"(?<!\d)(200|250|300|400|500|600|800|1000|1200|1600|2000|2500)\s*kV\s*(이하|초과|미만|이상)?"
)


def fix_cable_area_kv(spec: str, context: str = "") -> str:
    blob = f"{context} {spec}"
    if "케이블" not in blob:
        return spec
    return CABLE_AREA_AS_KV.sub(
        lambda match: f"{match.group(1)} ㎟ {match.group(2) or ''}".strip(),
        spec,
    )


def clean_title(title: str) -> str:
    text = normalize_spaces(title)
    text = re.sub(r"[·⋅．.]{2,}.*$", "", text)
    text = re.sub(r"([가-힣])\d{2,3}$", r"\1", text)
    text = re.sub(r"\s+\d{1,3}$", "", text)
    return text.strip()


def join_kind(kind: str, spec: str) -> str:
    kind = display_spec(kind)
    spec = display_spec(spec)
    if not kind:
        return spec
    if not spec:
        return kind
    compact_kind = kind.replace(" ", "")
    compact_spec = spec.replace(" ", "")
    if compact_spec.startswith(compact_kind) or compact_kind in compact_spec:
        return spec
    return f"{kind} {spec}".strip()


def clean_spec_unit(spec: Any, unit: Any, prev_spec: str = "", unit_hint: str = "", kind: str = "") -> tuple[str, str]:
    spec_text = expand_ditto(str(spec or ""), prev_spec, unit_hint=unit_hint)
    spec_text = apply_unit_hint(spec_text, unit_hint)
    spec_text = strip_leaked_qty(spec_text)
    spec_text, unit_text = strip_trailing_unit_dash(spec_text, str(unit or ""))
    spec_text = join_kind(kind, spec_text)
    spec_text = apply_unit_hint(spec_text, unit_hint)
    spec_text = fix_cable_area_kv(spec_text, f"{kind} {prev_spec}")
    spec_text = strip_leaked_qty(spec_text)
    spec_text = display_spec(spec_text)
    unit_text = display_unit(unit_text)
    return spec_text, unit_text


def is_empty_qty(value: Any) -> bool:
    text = normalize_spaces(str(value or ""))
    return text in {"", "-", "－", "—", "–", "None"}
