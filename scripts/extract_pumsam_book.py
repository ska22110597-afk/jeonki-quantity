#!/usr/bin/env python3
"""전기·통신 표준품셈 PDF 표를 품셈 행으로 푼다. 〃·'-'·제곱·단위를 원표 뜻대로 읽는다."""

from __future__ import annotations

import argparse
import json
import re
from collections import Counter, defaultdict
from pathlib import Path

import pdfplumber

from app.pumsam_text import (
    DASH_ONLY,
    DITTO,
    clean_spec_unit,
    clean_title,
    display_spec,
    display_unit,
    expand_ditto,
    is_empty_qty,
    is_qty_like_spec,
    join_kind,
    normalize_spaces,
    strip_leaked_qty,
)

ELECTRIC_JOBS: list[tuple[str, str]] = [
    ("특고압케이블전공", "특고압케이블전공"),
    ("고압케이블전공", "고압케이블전공"),
    ("저압케이블전공", "저압케이블전공"),
    ("송전활선전공", "송전활선전공"),
    ("배전활선전공", "배전활선전공"),
    ("전기공사산업기사", "전기공사산업기사"),
    ("전기공사기사", "전기공사기사"),
    ("중급기술자(엔지니어링)", "중급기술자(엔지니어링)"),
    ("중급기술자(측량)", "중급기술자(측량)"),
    ("초급기술자(측량)", "초급기술자(측량)"),
    ("인력운반공", "인력운반공"),
    ("기계설비공", "기계설비공"),
    ("작업반장", "작업반장"),
    ("케이블전공", "케이블전공"),
    ("플랜트전공", "플랜트전공"),
    ("송전전공", "송전전공"),
    ("배전전공", "배전전공"),
    ("변전전공", "변전전공"),
    ("내선전공", "내선전공"),
    ("전철전공", "전철전공"),
    ("특별인부", "특별인부"),
    ("보통인부", "보통인부"),
    ("비계공", "비계공"),
    ("도장공", "도장공"),
    ("용접공", "용접공"),
    ("철골공", "철골공"),
    ("철근공", "철근공"),
    ("콘크리트공", "콘크리트공"),
    ("형틀목공", "형틀목공"),
    ("계장공", "계장공"),
    ("배관공", "배관공"),
    ("조력공", "조력공"),
    ("착암공", "착암공"),
    ("화약취급공", "화약취급공"),
    ("건설기계운전사", "건설기계운전사"),
]

ELECTRIC_SPACED = [
    ("특고압 케이블전공", "특고압케이블전공"),
    ("고압 케이블전공", "고압케이블전공"),
    ("저압 케이블전공", "저압케이블전공"),
    ("송전 활선전공", "송전활선전공"),
    ("배전 활선전공", "배전활선전공"),
    ("전기공사 산업기사", "전기공사산업기사"),
    ("전기공사 기사", "전기공사기사"),
    ("중급 기술자 (엔지니어링)", "중급기술자(엔지니어링)"),
    ("중급 기술자 (측량)", "중급기술자(측량)"),
    ("초급 기술자 (측량)", "초급기술자(측량)"),
    ("인력 운반공", "인력운반공"),
    ("기계 설비공", "기계설비공"),
    ("변전 전공", "변전전공"),
    ("송전 전공", "송전전공"),
    ("배전 전공", "배전전공"),
    ("특별 인부", "특별인부"),
    ("보통 인부", "보통인부"),
]

TELECOM_JOBS: list[tuple[str, str]] = [
    ("통신관련산업기사", "통신관련산업기사"),
    ("광케이블설치사", "광케이블설치사"),
    ("통신케이블공", "통신케이블공"),
    ("통신내선공", "통신내선공"),
    ("통신외선공", "통신외선공"),
    ("통신설비공", "통신설비공"),
    ("통신기사", "통신기사"),
    ("특별인부", "특별인부"),
    ("보통인부", "보통인부"),
    ("비계공", "비계공"),
    ("도장공", "도장공"),
    ("용접공", "용접공"),
    ("철골공", "철골공"),
    ("콘크리트공", "콘크리트공"),
    ("조력공", "조력공"),
]

TELECOM_SPACED = [
    ("통신 관련 산업기사", "통신관련산업기사"),
    ("통신관련 산업기사", "통신관련산업기사"),
    ("광케이블 설치사", "광케이블설치사"),
    ("통신 케이블공", "통신케이블공"),
    ("통신 내선공", "통신내선공"),
    ("통신 외선공", "통신외선공"),
    ("통신 설비공", "통신설비공"),
    ("통신 기사", "통신기사"),
    ("특별 인부", "특별인부"),
    ("보통 인부", "보통인부"),
]

SKIP_SPEC = re.compile(r"^(합계|소계|계|구분|공종|규격|종별|명칭|용량|직종|단위|비고|적용직종|호칭|공정)$")
MANY_FLOATS = re.compile(r"\d+\.\d+")
NUM_TOKEN = re.compile(r"^(?:-|\d{1,4}(?:,\d{3})*(?:\.\d+)?)$")
PERCENT = re.compile(r"^\d+(?:\.\d+)?%$")
PAGE_NOISE = re.compile(
    r"^(제[0-9]+장|전기부문|정보통신|\d{2,4}\s*전기부문|\d+\s*$|【해설】|\[해\s*설\]|주\d+\)|※)"
)
UNIT_LINE = re.compile(r"단위\s*[:：]\s*([^)\n]+)")
APPLY_JOB = re.compile(r"(?:적용)?직종\s*[:：]\s*([가-힣]+)")
SKIP_QTY_HEADER = re.compile(r"(장비|시간\s*\(|비고|참고|기계손료)")
TOC_PAGE_TAIL = re.compile(r"\s+\d{1,3}$")
CAPACITY = re.compile(r"(\d+(?:\.\d+)?\s*(?:kVA|MVA|kV|㎸|㎸A|A)\s*(?:이하|초과|미만)?)", re.I)
PROCESS_ONLY = re.compile(r"^(운반|소운반|OT|점검|설치|배치|조가|교정|도장|높이)")
ELECTRIC_ITEM = re.compile(
    r"^([2-9]|10)-(\d{1,2}(?:-\d{1,2})?)\s+(.+?)(?:\s*\([’'`]?\d{2}년.*)?$"
)
TELECOM_ITEM = re.compile(
    r"^((?:[2-9]|1[0-3])-\d{1,2}(?:-\d{1,2})?)\s+(.+?)(?:\s*\(삭제.*)?$"
)
TOC_DOTS = re.compile(r"[·⋅．\.]{3,}")
HOCHING = re.compile(r"^호칭$")
KIND_HEADERS = {"종류", "구분", "공정", "공종"}
SPEC_COL_HEADERS = ("규격", "용량", "호칭", "종별", "단면적")
SPEC_HEADERS = {"규격", "규격및명칭", "공정및규격"}
UNIT_HEADERS = {"단위"}
SKIP_ELECTRIC = {"2-1", "2-1-1", "2-1-2", "2-1-3", "2-10", "2-10-1", "2-10-2"}


def job_pairs(discipline: str) -> tuple[list[tuple[str, str]], list[tuple[str, str]]]:
    if discipline == "통신":
        return TELECOM_JOBS, TELECOM_SPACED
    return ELECTRIC_JOBS, ELECTRIC_SPACED


def detect_jobs(line: str, discipline: str) -> list[str]:
    aliases, spaced = job_pairs(discipline)
    raw = normalize_spaces(line)
    compact = raw.replace(" ", "")
    found: list[tuple[int, str]] = []
    work = raw
    for alias, canon in spaced + aliases:
        idx = work.find(alias)
        if idx >= 0:
            found.append((idx, canon))
            work = work.replace(alias, " " * len(alias), 1)
    if not found:
        tmp = compact
        for alias, canon in aliases:
            idx = tmp.find(alias)
            if idx >= 0:
                found.append((idx, canon))
                tmp = tmp.replace(alias, " " * len(alias), 1)
    found.sort()
    jobs = [job for _, job in found]
    if not jobs:
        return []
    nums = [tok for tok in raw.split() if NUM_TOKEN.match(tok.replace(",", "")) and tok not in {"-"}]
    if nums:
        return []
    return jobs


def cell_job(text: str, discipline: str) -> str | None:
    jobs = detect_jobs(text or "", discipline)
    if len(jobs) == 1:
        return jobs[0]
    compact = normalize_spaces(text or "").replace(" ", "")
    aliases, _ = job_pairs(discipline)
    for alias, canon in aliases:
        if compact == alias.replace(" ", ""):
            return canon
    return None


def looks_like_header_piece(line: str) -> bool:
    compact = line.replace(" ", "")
    keys = ("규격", "종별", "명칭", "공종", "용량", "직종", "전공", "인부", "기술자", "호칭", "공정")
    if not any(key in compact for key in keys):
        return False
    if MANY_FLOATS.search(line) and not any(key in compact for key in ("전공", "인부", "기술자", "내선공", "외선공")):
        return False
    return True


def parse_qty_cell(text: str) -> tuple[bool, float | None]:
    """칸에 값이 있는지, 숫자는 얼마인지. 대시('-')는 값이 있는 빈 품으로 본다."""
    tok = normalize_spaces(text or "").replace(",", "")
    if not tok:
        return False, None
    if DASH_ONLY.match(tok) or tok in {"-", "－", "—", "–"}:
        return True, None
    if PERCENT.match(tok):
        return False, None
    if NUM_TOKEN.match(tok):
        return True, float(tok)
    return False, None


def single_number(text: str) -> float | None:
    present, value = parse_qty_cell(text)
    if not present:
        return None
    return value


def split_cell_lines(text: str) -> list[str]:
    raw = str(text or "")
    if "\n" not in raw:
        lines = [normalize_spaces(raw)]
    else:
        lines = [normalize_spaces(part) for part in raw.split("\n")]
    collapsed: list[str] = []
    for line in lines:
        if collapsed and re.fullmatch(r"\d+", line) and re.search(r"[가-힣A-Za-z]", collapsed[-1]):
            collapsed[-1] = merge_subscript_pair(collapsed[-1], line)
        else:
            collapsed.append(line)
    return collapsed


def flatten_cell(text: str) -> str:
    return normalize_spaces(str(text or "").replace("\n", " "))


def merge_subscript_pair(first: str, second: str) -> str:
    """B (1.415 ㎡) + 0 → B0 (1.415 ㎡). SF 가스 처리 + 6 → SF6 가스 처리."""
    first = normalize_spaces(first)
    second = normalize_spaces(second)
    if not second:
        return first
    if not first:
        return second
    if re.fullmatch(r"\d+", second):
        if re.match(r"^SF\b", first, flags=re.I):
            return re.sub(r"^SF\b", f"SF{second}", first, count=1, flags=re.I)
        if re.match(r"^B\b", first):
            return re.sub(r"^B\b", f"B{second}", first, count=1)
        return f"{first} {second}".strip()
    return f"{first} {second}".strip()


def pair_subscript_lines(lines: list[str]) -> list[str]:
    if len(lines) < 2 or len(lines) % 2:
        return lines
    paired: list[str] = []
    for index in range(0, len(lines), 2):
        paired.append(merge_subscript_pair(lines[index], lines[index + 1]))
    return paired


def align_cell_lines(lines: list[str], common: int) -> list[str]:
    if len(lines) == common:
        return lines
    if len(lines) == common * 2:
        return pair_subscript_lines(lines)
    if len(lines) == 1:
        return lines
    if common > 1 and len(lines) == common * 2 + 0:
        return pair_subscript_lines(lines)
    return lines


def split_multiline_row(cells: list[str]) -> list[list[str]]:
    parts = [split_cell_lines(cell) for cell in cells]
    counts = [len(part) for part in parts]
    multi = [count for count in counts if count > 1]
    if not multi:
        return [[flatten_cell(cell) for cell in cells]]
    common, freq = Counter(multi).most_common(1)[0]
    if common <= 1 or freq < 2:
        return [[flatten_cell(cell) for cell in cells]]
    aligned = [align_cell_lines(part, common) for part in parts]
    rows: list[list[str]] = []
    for index in range(common):
        row: list[str] = []
        for part in aligned:
            if len(part) == common:
                row.append(part[index])
            elif len(part) == 1:
                row.append(part[0] if index == 0 else "")
            else:
                row.append(part[index] if index < len(part) else "")
        rows.append(row)
    return rows


def header_unit_hint(header: list[str]) -> str:
    blob = " ".join(header)
    matched = re.search(r"(?:규격|구경|용량|호칭|종별|단면적)\s*\(([^)]+)\)", blob)
    if matched:
        hint = display_spec(matched.group(1).split("×")[0].split("x")[0].split(",")[0])
        hint = hint.replace("㎜", "mm")
        if hint in {"㎟", "mm2", "mm²"}:
            return "㎟"
        if "mm" in hint.lower() or "㎜" in matched.group(1):
            return "mm"
        if re.search(r"kVA", hint, re.I):
            return "kVA"
        if re.search(r"kW", hint, re.I):
            return "kW"
        return hint
    if re.search(r"구경", blob) and re.search(r"㎜|mm", blob, re.I):
        return "mm"
    if re.search(r"전동기", blob) and re.search(r"용량", blob):
        return "kW"
    if re.search(r"kVA", blob, re.I):
        return "kVA"
    if re.search(r"kW", blob, re.I):
        return "kW"
    return ""


def short_name(title: str) -> str:
    text = re.sub(r"\s*\(.*\)\s*$", "", title)
    for tail in (
        " 기계화 시공 및 지지선 설치",
        " 인력 이용 설치",
        " 기계장비 이용 설치",
        " 인력 세움",
        " 기계 세움",
        " 백호 세움",
        " 인력 설치",
        " 기계 설치",
        " 기계화 설치",
        " 조립 및 설치",
        " 설치공사",
        " 포설",
        " 설치",
        " 신설",
    ):
        if text.endswith(tail):
            text = text[: -len(tail)]
            break
    return text.strip() or title


def blank_title_spec(spec: str, title: str) -> str:
    """공종 칸이 품명과 같으면 규격으로 쓰지 않는다. 품명 전체가 앞에 반복되면 뗀다."""
    compact_spec = re.sub(r"\s+", "", spec or "")
    if not compact_spec:
        return ""
    compact_title = re.sub(r"\s+", "", title or "")
    compact_short = re.sub(r"\s+", "", short_name(title) or "")
    if compact_spec in {compact_title, compact_short}:
        return ""
    if title and spec.startswith(title) and spec != title:
        return spec[len(title) :].strip()
    short = short_name(title)
    if short and spec.startswith(short + " ") and spec != short:
        return spec[len(short) :].strip()
    return spec


def fill_header_groups(row: list) -> list[str]:
    groups: list[str] = []
    last = ""
    for cell in row:
        text = flatten_cell(cell)
        if text:
            last = text
        groups.append(last)
    return groups


def title_for_group(title: str, group: str) -> str:
    compact = (group or "").replace(" ", "")
    if "시운전" in compact:
        base = re.sub(r"\s*설치\s*$", "", title).strip() or title
        if "시운전" in base.replace(" ", ""):
            return title
        return f"{base} 시운전 및 조정"
    if compact == "조립" or compact.endswith("조립"):
        base = re.sub(r"\s*설치\s*$", "", title).strip() or title
        if "조립" in base.replace(" ", ""):
            return title
        return f"{base} 조립"
    return title


def spec_for_group(spec: str, group: str) -> str:
    compact = (group or "").replace(" ", "").replace("△", "Δ").replace("δ", "Δ")
    if "직입" in compact:
        if "직입" in (spec or "").replace(" ", ""):
            return spec
        return f"{spec} 직입기동".strip()
    if "Y-" in compact or "YΔ" in compact or "Y△" in group or "델타" in compact:
        if "Y-" in (spec or "") or "직입" in (spec or "").replace(" ", ""):
            return spec
        return f"{spec} Y-△ 기동".strip()
    return spec


def canon_job_name(job: str) -> str:
    job = (job or "").strip()
    if job.endswith(")") and "(" not in job:
        job = job[:-1]
    return job


def emit(rows: list[dict], title: str, spec: str, unit: str, job: str, qty: float | None, code: str, prefix: str) -> None:
    title = clean_title(title)
    job = canon_job_name(job)
    if qty is None or qty <= 0:
        return
    spec = strip_leaked_qty(spec)
    spec, unit = clean_spec_unit(spec, unit)
    spec = blank_title_spec(spec, title)
    spec = re.sub(r"\s*[-－—–]\s*$", "", spec).strip()
    if is_qty_like_spec(spec, qty):
        return
    if not job or not title or not code:
        return
    if spec and SKIP_SPEC.match(spec.replace(" ", "")):
        return
    if spec and len(MANY_FLOATS.findall(spec)) >= 2:
        if not re.search(r"[~～]|초과|이하|미만|이상|kW|kVA|mm|㎟|㎜", spec):
            return
    if spec and PROCESS_ONLY.match(spec) and not CAPACITY.search(spec):
        return
    if "삭제" in title:
        return
    rows.append(
        {
            "명칭": title,
            "짧은명칭": short_name(title),
            "규격": spec,
            "단위": unit or "식",
            "노무명칭": job,
            "품셈": round(qty, 4) if isinstance(qty, float) and abs(qty) < 1000 else qty,
            "할증%": 100,
            "품셈근거": f"{prefix}{code}",
        }
    )


def _spec_job_pairs(header: list[str], discipline: str) -> list[tuple[int, int, str, str]]:
    """용량+직종이 두 덩어리로 반복되는 표. 왼쪽·오른쪽 규격을 섞지 않는다."""
    pairs: list[tuple[int, int, str, str]] = []
    index = 0
    while index < len(header) - 1:
        compact = header[index].replace(" ", "")
        job = cell_job(header[index + 1], discipline)
        is_spec = any(token in compact for token in ("호칭", "용량", "규격", "종별", "단면적"))
        if is_spec and job:
            pairs.append((index, index + 1, job, header[index]))
            index += 2
            continue
        index += 1
    return pairs


def parse_paired_size_table(
    table: list[list],
    code: str,
    title: str,
    unit: str,
    rows: list[dict],
    discipline: str,
    prefix: str,
) -> int:
    if not table or len(table) < 2:
        return 0
    top = [flatten_cell(cell) for cell in table[0]]
    mid = [flatten_cell(cell) for cell in table[1]] if len(table) > 1 else []
    hint = header_unit_hint(top) or header_unit_hint(mid)
    pairs = _spec_job_pairs(top, discipline)
    start = 1
    if len(pairs) < 2 and mid:
        for col in range(0, len(mid) - 1):
            if "호칭" not in mid[col].replace(" ", "") and mid[col] != "호칭":
                continue
            job = cell_job(mid[col + 1], discipline)
            if not job:
                continue
            group = top[col] if col < len(top) else ""
            if not group and col:
                group = top[col - 1] if col - 1 < len(top) else ""
            pairs.append((col, col + 1, job, group or title))
        if len(pairs) >= 2:
            start = 2
    if len(pairs) < 2:
        return 0
    before = len(rows)
    last_specs = {spec_col: "" for spec_col, _, _, _ in pairs}
    for raw in table[start:]:
        cells = [cell if cell is not None else "" for cell in raw]
        for spec_col, qty_col, job, group in pairs:
            spec_lines = split_cell_lines(cells[spec_col] if spec_col < len(cells) else "")
            qty_lines = split_cell_lines(cells[qty_col] if qty_col < len(cells) else "")
            count = max(len(spec_lines), len(qty_lines), 1)
            for index in range(count):
                spec = display_spec(spec_lines[index] if index < len(spec_lines) else "")
                present, qty = parse_qty_cell(qty_lines[index] if index < len(qty_lines) else "")
                if spec in {"-"}:
                    continue
                spec = expand_ditto(spec, last_specs[spec_col], unit_hint=hint)
                if re.fullmatch(r"\d+(?:\.\d+)?", spec):
                    if hint in {"mm", "㎜"}:
                        spec = f"{spec} mm"
                    elif hint:
                        spec = f"{spec} {hint}"
                    elif any("호칭" in cell.replace(" ", "") for cell in top + mid):
                        spec = f"{spec} mm"
                if spec:
                    last_specs[spec_col] = spec
                if not present or not spec:
                    continue
                compact_group = (group or "").replace(" ", "")
                use_title = title
                if compact_group and not any(
                    token in compact_group for token in ("호칭", "용량", "규격", "종별", "단면적")
                ):
                    use_title = group
                emit(rows, use_title, spec, unit, job, qty, code, prefix)
    return len(rows) - before


def column_roles(header: list[str], discipline: str) -> tuple[int | None, int | None, int | None, list[tuple[int, str]], str]:
    kind_col = spec_col = unit_col = None
    job_cols: list[tuple[int, str]] = []
    unnamed: list[int] = []
    hint = header_unit_hint(header)
    for index, cell in enumerate(header):
        compact = cell.replace(" ", "")
        job = cell_job(cell, discipline)
        if job:
            job_cols.append((index, job))
            continue
        if any(token in compact for token in SPEC_COL_HEADERS):
            if spec_col is None:
                spec_col = index
            continue
        if compact in UNIT_HEADERS:
            unit_col = index
            continue
        if compact in KIND_HEADERS or compact in SPEC_HEADERS:
            if spec_col is None and compact in SPEC_HEADERS:
                spec_col = index
            elif kind_col is None:
                kind_col = index
            continue
        if not compact:
            unnamed.append(index)
    first_job = job_cols[0][0] if job_cols else None
    if spec_col is not None and first_job is not None and spec_col >= first_job:
        spec_col = None
    if spec_col is None and first_job is not None:
        for index in range(first_job):
            if index != kind_col and index != unit_col:
                spec_col = index
                break
    if spec_col is None and unnamed:
        spec_col = unnamed[0]
        if first_job is not None and spec_col >= first_job:
            spec_col = 0 if first_job > 0 else None
    if spec_col is None and kind_col is not None and unit_col is not None:
        spec_col, kind_col = kind_col, None
    return kind_col, spec_col, unit_col, job_cols, hint


def is_skip_qty_header(text: str) -> bool:
    compact = (text or "").replace(" ", "")
    return bool(SKIP_QTY_HEADER.search(compact))


def attach_apply_jobs(
    header: list[str],
    spec_col: int | None,
    unit_col: int | None,
    kind_col: int | None,
    apply_job: str,
) -> tuple[list[tuple[int, str]], list[str]]:
    """표에 직종 칸이 없고 적용직종만 있을 때, 숫자 칸을 그 직종으로 읽는다."""
    job_cols: list[tuple[int, str]] = []
    groups: list[str] = []
    last = ""
    for index, cell in enumerate(header):
        text = flatten_cell(cell)
        if text:
            last = text
        groups.append(last)
        if index in {spec_col, unit_col, kind_col}:
            continue
        if not text or is_skip_qty_header(text):
            continue
        if SKIP_SPEC.match(text.replace(" ", "")):
            continue
        if cell_job(text, "전기") or cell_job(text, "통신"):
            continue
        job_cols.append((index, apply_job))
    return job_cols, groups


def parse_matrix_table(
    table: list[list],
    code: str,
    title: str,
    unit: str,
    rows: list[dict],
    apply_job: str,
    prefix: str,
) -> int:
    """구경 × 두께처럼 직종 없이 숫자만 있는 표."""
    if not table or not apply_job or not code:
        return 0
    header0 = [flatten_cell(cell) for cell in table[0]]
    blob = "".join(header0)
    if "구경" not in blob and "두께" not in blob:
        return 0
    hint = header_unit_hint(header0)
    col_headers = list(header0)
    start = 1
    if len(table) > 1:
        header1 = [flatten_cell(cell) for cell in table[1]]
        qty_in_header1 = any(parse_qty_cell(cell)[0] for cell in header1 if cell)
        if not qty_in_header1:
            while len(col_headers) < len(header1):
                col_headers.append("")
            for index, cell in enumerate(header1):
                if cell:
                    col_headers[index] = cell
            hint = hint or header_unit_hint(header0 + header1)
            start = 2
    spec_col = 0
    unit_col = None
    qty_cols: list[int] = []
    for index, header in enumerate(col_headers):
        compact = header.replace(" ", "")
        if index == 0 or "구경" in compact:
            spec_col = index
            continue
        if compact in UNIT_HEADERS or compact == "단위":
            unit_col = index
            continue
        if not header or is_skip_qty_header(header):
            continue
        if "두께" in compact and "이하" not in compact and "㎜" not in compact and "mm" not in compact.lower():
            continue
        qty_cols.append(index)
    if not qty_cols:
        return 0
    before = len(rows)
    last_spec = ""
    last_unit = unit
    for raw in table[start:]:
        cells = [cell if cell is not None else "" for cell in raw]
        for split in split_multiline_row(cells):
            while len(split) < max(qty_cols) + 1:
                split.append("")
            spec = flatten_cell(split[spec_col]) if spec_col < len(split) else ""
            spec = expand_ditto(spec, last_spec, unit_hint=hint)
            spec, _ = clean_spec_unit(spec, "", prev_spec=last_spec, unit_hint=hint)
            if hint and spec and hint not in spec and re.fullmatch(r"\d+(?:\.\d+)?", spec):
                spec = f"{spec} {hint}"
            if spec:
                last_spec = spec
            row_unit = unit
            if unit_col is not None and unit_col < len(split) and flatten_cell(split[unit_col]):
                row_unit = display_unit(split[unit_col])
            if row_unit in {"〃", "＂", "“", "-"} or DITTO in (row_unit or ""):
                row_unit = last_unit or unit
            last_unit = row_unit or last_unit
            if not spec or spec in {"계", "합계", "소계"}:
                continue
            if hint in {"mm", "㎜"} and not re.match(r"\d{2,}", spec):
                continue
            for col in qty_cols:
                present, qty = parse_qty_cell(split[col] if col < len(split) else "")
                if not present:
                    continue
                thick = col_headers[col] if col < len(col_headers) else ""
                full_spec = f"{spec} × {thick}".strip(" ×") if thick else spec
                emit(rows, title, full_spec, row_unit or unit, apply_job, qty, code, prefix)
    return len(rows) - before


def parse_labor_table(
    table: list[list],
    code: str,
    title: str,
    unit: str,
    rows: list[dict],
    discipline: str,
    prefix: str,
    apply_job: str = "",
) -> int:
    if not table or len(table) < 2 or not code:
        return 0
    header0 = [flatten_cell(cell) for cell in table[0]]
    kind_col, spec_col, unit_col, job_cols, hint = column_roles(header0, discipline)
    groups = fill_header_groups(table[0])
    start = 1
    if not job_cols and len(table) > 1:
        header1 = [flatten_cell(cell) for cell in table[1]]
        kind1, spec1, unit1, job_cols, hint1 = column_roles(header1, discipline)
        hint = hint or hint1 or header_unit_hint(header0)
        if spec_col is None:
            spec_col = spec1
        if unit_col is None:
            unit_col = unit1
        if kind_col is None:
            kind_col = kind1
        if job_cols and spec_col is not None and spec_col >= job_cols[0][0]:
            spec_col = spec_col if spec_col < job_cols[0][0] else (0 if job_cols[0][0] > 0 else None)
        start = 2
        if job_cols:
            while len(groups) < max(col for col, _ in job_cols) + 1:
                groups.append(groups[-1] if groups else "")
    if not job_cols and apply_job:
        hint = hint or header_unit_hint(header0)
        extra_cols, extra_groups = attach_apply_jobs(header0, spec_col, unit_col, kind_col, apply_job)
        if extra_cols:
            job_cols = extra_cols
            groups = extra_groups
            start = 1
            if len(table) > 1 and any(cell_job(flatten_cell(cell), discipline) for cell in table[1]):
                start = 2
            if spec_col is None:
                spec_col = 0
    if not job_cols:
        return 0
    job_cols = [(col, job) for col, job in job_cols if not is_skip_qty_header(groups[col] if col < len(groups) else "")]
    if not job_cols:
        return 0
    if spec_col is not None and spec_col >= job_cols[0][0]:
        spec_col = 0 if job_cols[0][0] > 0 else None
    while len(groups) < max(col for col, _ in job_cols) + 1:
        groups.append(groups[-1] if groups else "")
    before = len(rows)
    last_kind = ""
    last_spec = ""
    last_unit = unit
    first_job = job_cols[0][0]
    for raw in table[start:]:
        cells = [cell if cell is not None else "" for cell in raw]
        for split in split_multiline_row(cells):
            while len(split) < max(first_job + 1, len(header0), (table[1] and len(table[1]) or 0)):
                split.append("")
            kind = flatten_cell(split[kind_col]) if kind_col is not None and kind_col < len(split) else ""
            if kind:
                last_kind = kind
            else:
                kind = last_kind
            if spec_col is not None and spec_col < len(split):
                spec = flatten_cell(split[spec_col])
            else:
                left = [flatten_cell(split[i]) for i in range(first_job) if i != kind_col]
                spec = " ".join(part for part in left if part).strip()
            if not spec and kind:
                spec = kind
            row_unit = unit
            if unit_col is not None and unit_col < len(split) and flatten_cell(split[unit_col]):
                row_unit = display_unit(split[unit_col])
            if row_unit in {"〃", "＂", "“", "-"} or DITTO in (row_unit or ""):
                row_unit = last_unit or unit
            spec = expand_ditto(spec, last_spec, unit_hint=hint)
            spec, row_unit = clean_spec_unit(spec, row_unit, prev_spec=last_spec, unit_hint=hint, kind=kind)
            spec = blank_title_spec(spec, title)
            if spec in {"계", "합계", "소계"}:
                continue
            if spec:
                last_spec = spec
            last_unit = row_unit or last_unit
            parsed = False
            for col, job in job_cols:
                if col >= len(split):
                    continue
                present, qty = parse_qty_cell(split[col])
                if not present:
                    continue
                group = groups[col] if col < len(groups) else ""
                emit(
                    rows,
                    title_for_group(title, group),
                    spec_for_group(spec, group),
                    row_unit or unit,
                    job,
                    qty,
                    code,
                    prefix,
                )
                parsed = True
            if not parsed:
                continue
    return len(rows) - before


def parse_table(
    table: list[list] | None,
    code: str,
    title: str,
    unit: str,
    rows: list[dict],
    discipline: str,
    prefix: str,
    skip_codes: set[str],
    apply_job: str = "",
) -> int:
    if not table or not code or code in skip_codes:
        return 0
    added = parse_paired_size_table(table, code, title, unit, rows, discipline, prefix)
    if added:
        return added
    added = parse_matrix_table(table, code, title, unit, rows, apply_job, prefix)
    if added:
        return added
    return parse_labor_table(table, code, title, unit, rows, discipline, prefix, apply_job=apply_job)


def looks_like_toc_line(line: str) -> bool:
    if TOC_DOTS.search(line):
        return True
    if re.search(r"[·⋅．\.]{3,}", line):
        return True
    if TOC_PAGE_TAIL.search(line) and not MANY_FLOATS.search(line) and not UNIT_LINE.search(line):
        if "적용직종" not in line.replace(" ", "") and "단위" not in line:
            return True
    return False


def parse_item_line(line: str, discipline: str) -> tuple[str, str] | None:
    if looks_like_toc_line(line):
        return None
    if discipline == "통신":
        matched = TELECOM_ITEM.match(line)
        if not matched:
            return None
        title = normalize_spaces(matched.group(2))
        title = re.sub(r"\s*\(삭제.*$", "", title).strip()
        title = re.sub(r"[·⋅．\.]{2,}.*$", "", title).strip()
        if not title:
            return None
        return matched.group(1), title
    matched = ELECTRIC_ITEM.match(line)
    if not matched:
        return None
    title = normalize_spaces(matched.group(3))
    title = re.sub(r"\s*\([’'`´]?[0-9]{2}년.*$", "", title).strip()
    return f"{matched.group(1)}-{matched.group(2)}", title


def parse_nums(line: str, n: int) -> tuple[str, list[float | None]] | None:
    parts = normalize_spaces(line).split()
    if not parts:
        return None
    nums: list[float | None] = []
    while parts and len(nums) < n:
        tok = parts[-1].replace(",", "")
        if tok in {"-", "－", "—", "–"}:
            nums.append(None)
            parts.pop()
            continue
        if NUM_TOKEN.match(tok):
            nums.append(float(tok))
            parts.pop()
            continue
        break
    nums.reverse()
    if len(nums) != n:
        return None
    spec = " ".join(parts).strip()
    if not spec or spec in {"규격", "종별", "명칭", "공종", "용량", "구분", "호칭"}:
        return None
    if SKIP_SPEC.match(spec.replace(" ", "")) and spec.replace(" ", "") in {"합계", "소계", "계"}:
        return None
    return spec, nums


def extract(pdf_path: Path, discipline: str) -> list[dict]:
    prefix = "통신" if discipline == "통신" else "전기"
    skip_codes = SKIP_ELECTRIC if discipline == "전기" else set()
    pdf = pdfplumber.open(str(pdf_path))
    rows: list[dict] = []
    current_code = ""
    current_title = ""
    current_unit = ""
    jobs: list[str] = []
    apply_job = ""
    last_spec = ""
    last_kind = ""
    header_buf: list[str] = []
    titles_by_code: dict[str, str] = {}
    apply_by_code: dict[str, str] = {}
    unit_by_code: dict[str, str] = {}
    table_codes: set[str] = set()

    for page in pdf.pages:
        text = page.extract_text() or ""
        incoming_code, incoming_title = current_code, current_title
        incoming_unit = current_unit
        word_heads: list[tuple[float, str, str]] = []
        by_top: dict[float, list[str]] = {}
        for word in page.extract_words() or []:
            by_top.setdefault(round(float(word["top"]), 1), []).append(word["text"])
        for top, toks in sorted(by_top.items()):
            line = normalize_spaces(" ".join(toks))
            parsed = parse_item_line(line, discipline)
            if parsed:
                word_heads.append((top, parsed[0], parsed[1]))
                full = clean_title(parsed[1])
                if len(full) > len(titles_by_code.get(parsed[0], "")):
                    titles_by_code[parsed[0]] = full
        page_tables = []
        try:
            page_tables = page.find_tables() or []
        except Exception:
            page_tables = []
        toc_page = len(re.findall(r"[·⋅．\.]{6,}", text)) >= 5 or (
            len(word_heads) >= 6 and not page_tables
        )
        if not toc_page:
            for _top, code, title in word_heads:
                current_code, current_title = code, title
        scan_code = incoming_code if toc_page else incoming_code
        for raw in text.splitlines():
            line = normalize_spaces(raw)
            parsed_item = parse_item_line(line, discipline)
            if parsed_item:
                if toc_page:
                    full = clean_title(parsed_item[1])
                    if len(full) > len(titles_by_code.get(parsed_item[0], "")):
                        titles_by_code[parsed_item[0]] = full
                else:
                    scan_code, current_title = parsed_item
                    current_code = scan_code
                    current_unit = ""
                    full = clean_title(parsed_item[1])
                    if len(full) > len(titles_by_code.get(parsed_item[0], "")):
                        titles_by_code[parsed_item[0]] = full
            if toc_page:
                continue
            unit_match = UNIT_LINE.search(line)
            if unit_match and scan_code:
                current_unit = display_unit(unit_match.group(1))
                unit_by_code[scan_code] = current_unit
            apply_match = APPLY_JOB.search(line)
            if apply_match and scan_code:
                apply_by_code[scan_code] = apply_match.group(1).replace(" ", "")

        heads_for_tables = [] if toc_page else word_heads

        def table_context(table_top: float | None = None) -> tuple[str, str, str, str]:
            code, title = incoming_code, incoming_title
            if table_top is not None:
                for hy, hc, ht in heads_for_tables:
                    if hy <= table_top + 18:
                        code, title = hc, ht
            unit = unit_by_code.get(code) or incoming_unit
            job = apply_by_code.get(code, "")
            return code, title, unit, job

        if not page_tables:
            try:
                extracted = page.extract_tables() or []
            except Exception:
                extracted = []
            for tbl in extracted:
                code, title, unit, job = table_context()
                added = parse_table(
                    tbl, code, title, unit, rows, discipline, prefix, skip_codes, apply_job=job
                )
                if added:
                    table_codes.add(code)
        for tbl in page_tables:
            ty = tbl.bbox[1]
            code, title, unit, job = table_context(ty)
            try:
                extracted = tbl.extract()
            except Exception:
                continue
            added = parse_table(
                extracted, code, title, unit, rows, discipline, prefix, skip_codes, apply_job=job
            )
            if added:
                table_codes.add(code)

        jobs = []
        apply_job = apply_by_code.get(incoming_code, "")
        last_spec = ""
        last_kind = ""
        header_buf = []
        line_code = incoming_code
        line_title = incoming_title
        line_unit = incoming_unit
        if toc_page:
            continue
        for raw in text.splitlines():
            line = normalize_spaces(raw)
            if not line or PAGE_NOISE.match(line):
                continue
            parsed_item = parse_item_line(line, discipline)
            if parsed_item:
                line_code, line_title = parsed_item
                line_unit = ""
                jobs = []
                apply_job = ""
                last_spec = ""
                last_kind = ""
                header_buf = []
                continue
            unit_match = UNIT_LINE.search(line)
            if unit_match:
                line_unit = display_unit(unit_match.group(1))
            apply_match = APPLY_JOB.search(line)
            if apply_match:
                apply_job = apply_match.group(1).replace(" ", "")
            if looks_like_header_piece(line) and not MANY_FLOATS.search(line):
                header_buf.append(line)
                header_buf = header_buf[-4:]
                detected = detect_jobs(" ".join(header_buf), discipline)
                if len(detected) >= 1:
                    jobs = detected
                    header_buf = []
                    continue
            detected = detect_jobs(line, discipline)
            if len(detected) >= 2:
                jobs = detected
                header_buf = []
                last_spec = ""
                continue
            if not line_code or line_code in skip_codes:
                continue
            if line_code in table_codes:
                continue
            use_jobs = jobs if jobs else ([apply_job] if apply_job else [])
            if not use_jobs:
                continue
            parsed = parse_nums(line, len(use_jobs))
            if not parsed:
                continue
            spec, values = parsed
            spec = expand_ditto(spec, last_spec)
            spec, unit = clean_spec_unit(spec, line_unit, prev_spec=last_spec, kind=last_kind)
            spec = blank_title_spec(spec, line_title)
            if spec in {"계", "합계", "소계"}:
                continue
            if spec:
                last_spec = spec
            for job, qty in zip(use_jobs, values):
                emit(rows, line_title, spec, unit, job, qty, line_code, prefix)

    pdf.close()
    return finalize_book_rows(rows, titles_by_code, prefix)


def title_is_bad(name: str) -> bool:
    text = str(name or "").strip()
    if not re.search(r"[가-힣]", text):
        return True
    if text.endswith("(") or text.endswith(","):
        return True
    if len(text) < 3:
        return True
    return False


def finalize_book_rows(rows: list[dict], titles_by_code: dict[str, str], prefix: str) -> list[dict]:
    for row in rows:
        code = str(row.get("품셈근거") or "").replace(prefix, "", 1)
        better = titles_by_code.get(code)
        name = str(row.get("명칭") or "")
        if better and title_is_bad(name) and "조립" not in name and "시운전" not in name:
            row["명칭"] = better
            row["짧은명칭"] = short_name(better)
        if str(row.get("단위") or "") in {"식", ""} and (
            "단면적" in str(row.get("규격") or "") or "Box" in str(row.get("규격") or "")
        ):
            row["단위"] = "개"
    return dedupe_book_rows(cleanup_book_rows(rows))


def cleanup_book_rows(rows: list[dict]) -> list[dict]:
    """쪽번호·해설 찌꺼기·품셈 숫자가 규격으로 들어간 줄을 버린다."""
    cleaned: list[dict] = []
    for row in rows:
        job = canon_job_name(str(row.get("노무명칭") or ""))
        row["노무명칭"] = job
        spec = str(row.get("규격") or "").strip()
        name = str(row.get("명칭") or "")
        qty = row.get("품셈")
        unit = display_unit(str(row.get("단위") or ""))
        row["단위"] = unit
        if not job:
            continue
        if is_qty_like_spec(spec, qty):
            continue
        if "분전반" in name and re.fullmatch(r"\d+", spec):
            spec = f"{spec}회로"
            row["규격"] = spec
        if "분배기" in name:
            if spec.endswith("분기기") and "(" not in spec:
                spec = re.sub(r"\s*분기기\s*$", "", spec).strip()
                row["규격"] = spec
            if spec.count("분기기") > 1:
                spec = re.sub(r"\s*분기기\s*$", "", spec).strip()
                row["규격"] = spec
            if spec.endswith("분배기") and spec.lstrip().startswith("("):
                spec = spec[: spec.rfind("분배기")].strip()
                row["규격"] = spec
            if re.search(r"광\s*송신|광\s*증폭|위성방송|신호변환", spec):
                continue
        cleaned.append(row)
    return cleaned


def spec_score(spec: str) -> int:
    score = min(len(spec), 80)
    if "〃" in spec:
        score -= 80
    if re.search(r"\d+\.\d{2}$", spec):
        score -= 60
    if spec.endswith("-"):
        score -= 60
    if "㎟" in spec or "mm" in spec or "kV" in spec:
        score += 8
    if any(word in spec for word in ("이하", "초과", "미만")):
        score += 6
    return score


def spec_is_dirty(spec: str) -> bool:
    text = str(spec or "")
    if "〃" in text:
        return True
    if re.search(r"\d+\.\d{2}$", text):
        return True
    if text.endswith("-"):
        return True
    return False


def identity_key(row: dict) -> str:
    return f"{row['명칭']}|{row['규격']}|{row['노무명칭']}|{row['품셈근거']}"


def dedupe_book_rows(rows: list[dict]) -> list[dict]:
    """같은 칸을 여러 번 읽은 것만 합친다. 규격이 다른데 품 숫자만 같은 행은 남긴다."""
    by_id: dict[str, dict] = {}
    for row in rows:
        key = identity_key(row)
        prev = by_id.get(key)
        if prev is None or spec_score(str(row["규격"])) > spec_score(str(prev["규격"])):
            by_id[key] = row

    groups: dict[str, list[dict]] = defaultdict(list)
    kept: dict[str, dict] = {}
    for row in by_id.values():
        if row.get("품셈") is None:
            kept[identity_key(row)] = row
            continue
        qty_key = f"{row['명칭']}|{row['노무명칭']}|{row['품셈근거']}|{row['품셈']}"
        groups[qty_key].append(row)

    for group in groups.values():
        clean = [row for row in group if not spec_is_dirty(str(row["규격"]))]
        chosen = clean if clean else [max(group, key=lambda row: spec_score(str(row["규격"])))]
        for row in chosen:
            key = identity_key(row)
            prev = kept.get(key)
            if prev is None or spec_score(str(row["규격"])) > spec_score(str(prev["규격"])):
                kept[key] = row
    return list(kept.values())


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--discipline", choices=("전기", "통신"), required=True)
    parser.add_argument("--pdf", type=Path, required=True)
    parser.add_argument("--out", type=Path, required=True)
    args = parser.parse_args()
    rows = extract(args.pdf, args.discipline)
    args.out.parent.mkdir(parents=True, exist_ok=True)
    args.out.write_text(json.dumps(rows, ensure_ascii=False, indent=0), encoding="utf-8")
    by_ch = Counter(
        str(row["품셈근거"]).split("-")[0].replace("전기", "").replace("통신", "") for row in rows
    )
    jobs = Counter(row["노무명칭"] for row in rows)
    print("rows", len(rows))
    print("by chapter", dict(sorted(by_ch.items(), key=lambda item: int(item[0] or 0) if str(item[0]).isdigit() else 99)))
    print("jobs", dict(jobs.most_common(15)))
    ditto = sum(1 for row in rows if "〃" in str(row["규격"]))
    dash = sum(1 for row in rows if str(row["규격"]).endswith("-"))
    mm2 = sum(1 for row in rows if re.search(r"mm2|mm²", str(row["규격"]), re.I))
    print("remaining 〃", ditto, "dash-end", dash, "mm2", mm2)
    for row in rows[:6]:
        print(row)


if __name__ == "__main__":
    main()
