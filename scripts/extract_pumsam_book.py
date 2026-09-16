#!/usr/bin/env python3
"""전기·통신 표준품셈 PDF 표를 품셈 행으로 푼다. 〃·'-'·제곱·단위를 원표 뜻대로 읽는다."""

from __future__ import annotations

import argparse
import json
import re
from collections import Counter
from pathlib import Path

import pdfplumber

from app.pumsam_text import (
    DASH_ONLY,
    clean_spec_unit,
    clean_title,
    display_spec,
    display_unit,
    expand_ditto,
    is_empty_qty,
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
APPLY_JOB = re.compile(r"적용직종\s*[:：]\s*([가-힣]+)")
CAPACITY = re.compile(r"(\d+(?:\.\d+)?\s*(?:kVA|MVA|kV|㎸|㎸A|A)\s*(?:이하|초과|미만)?)", re.I)
PROCESS_ONLY = re.compile(r"^(운반|소운반|OT|점검|설치|배치|조가|교정|도장|높이)")
ELECTRIC_ITEM = re.compile(
    r"^([2-7])-(\d{1,2}(?:-\d{1,2})?)\s+(.+?)(?:\s*\([’'`]?\d{2}년.*)?$"
)
TELECOM_ITEM = re.compile(
    r"^((?:[2-9]|1[0-3])-\d{1,2}(?:-\d{1,2})?)\s+(.+?)(?:\s*\(삭제.*)?$"
)
TOC_DOTS = re.compile(r"[·⋅．\.]{3,}")
HOCHING = re.compile(r"^호칭$")
KIND_HEADERS = {"종류", "구분", "공정", "공종"}
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


def single_number(text: str) -> float | None:
    tok = normalize_spaces(text or "").replace(",", "")
    if is_empty_qty(tok) or PERCENT.match(tok):
        return None
    if NUM_TOKEN.match(tok):
        value = float(tok)
        return value
    return None


def split_cell_lines(text: str) -> list[str]:
    raw = str(text or "")
    if "\n" not in raw:
        return [normalize_spaces(raw)]
    return [normalize_spaces(part) for part in raw.split("\n")]


def flatten_cell(text: str) -> str:
    return normalize_spaces(str(text or "").replace("\n", " "))


def split_multiline_row(cells: list[str]) -> list[list[str]]:
    parts = [split_cell_lines(cell) for cell in cells]
    counts = [len(part) for part in parts]
    multi = [count for count in counts if count > 1]
    if not multi:
        return [[flatten_cell(cell) for cell in cells]]
    common, freq = Counter(multi).most_common(1)[0]
    if common <= 1 or freq < 2:
        return [[flatten_cell(cell) for cell in cells]]
    rows: list[list[str]] = []
    for index in range(common):
        row: list[str] = []
        for part in parts:
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
    matched = re.search(r"규격\s*\(([^)]+)\)", blob)
    if matched:
        hint = display_spec(matched.group(1))
        if hint in {"㎟", "mm2", "mm²"}:
            return "㎟"
        return hint
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


def emit(rows: list[dict], title: str, spec: str, unit: str, job: str, qty: float, code: str, prefix: str) -> None:
    title = clean_title(title)
    if qty is None or qty <= 0:
        return
    spec = strip_leaked_qty(spec)
    spec, unit = clean_spec_unit(spec, unit)
    if not spec or not job or not title or not code:
        return
    if SKIP_SPEC.match(spec.replace(" ", "")):
        return
    if len(MANY_FLOATS.findall(spec)) >= 2:
        return
    if PROCESS_ONLY.match(spec) and not CAPACITY.search(spec):
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
            "품셈": round(qty, 4) if abs(qty) < 1000 else qty,
            "할증%": 100,
            "품셈근거": f"{prefix}{code}",
        }
    )


def parse_paired_size_table(
    table: list[list],
    code: str,
    title: str,
    unit: str,
    rows: list[dict],
    discipline: str,
    prefix: str,
) -> int:
    if not table or len(table) < 3:
        return 0
    top = [flatten_cell(cell) for cell in table[0]]
    mid = [flatten_cell(cell) for cell in table[1]]
    pairs: list[tuple[int, int, str, str]] = []
    for col in range(0, len(mid) - 1):
        if "호칭" not in mid[col].replace(" ", "") and mid[col] != "호칭":
            continue
        job = cell_job(mid[col + 1], discipline)
        if not job:
            continue
        group = top[col] or (top[col - 1] if col else "")
        group = group or title
        pairs.append((col, col + 1, job, group))
    if len(pairs) < 2:
        return 0
    before = len(rows)
    for raw in table[2:]:
        cells = [cell if cell is not None else "" for cell in raw]
        for split in split_multiline_row(cells):
            while len(split) < len(mid):
                split.append("")
            for spec_col, qty_col, job, group in pairs:
                spec = display_spec(split[spec_col])
                qty = single_number(split[qty_col])
                if qty is None or not spec or spec in {"-"}:
                    continue
                if re.fullmatch(r"\d+(?:\.\d+)?", spec):
                    spec = f"{spec} mm"
                emit(rows, group or title, spec, unit, job, qty, code, prefix)
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
        if compact.startswith("규격") or "규격" in compact:
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
    if spec_col is None and unnamed:
        spec_col = unnamed[0]
    if spec_col is None and kind_col is not None and unit_col is not None:
        spec_col, kind_col = kind_col, None
    return kind_col, spec_col, unit_col, job_cols, hint


def parse_labor_table(
    table: list[list],
    code: str,
    title: str,
    unit: str,
    rows: list[dict],
    discipline: str,
    prefix: str,
) -> int:
    if not table or len(table) < 2 or not code:
        return 0
    header = [flatten_cell(cell) for cell in table[0]]
    kind_col, spec_col, unit_col, job_cols, hint = column_roles(header, discipline)
    start = 1
    if not job_cols and len(table) > 1:
        hint0 = hint
        kind0, spec0, unit0 = kind_col, spec_col, unit_col
        header = [flatten_cell(cell) for cell in table[1]]
        kind_col, spec_col, unit_col, job_cols, hint = column_roles(header, discipline)
        hint = hint or hint0
        kind_col = kind0 if kind0 is not None else kind_col
        spec_col = spec0 if spec0 is not None else spec_col
        unit_col = unit0 if unit0 is not None else unit_col
        start = 2
    if not job_cols:
        return 0
    before = len(rows)
    last_kind = ""
    last_spec = ""
    last_unit = unit
    first_job = job_cols[0][0]
    for raw in table[start:]:
        cells = [cell if cell is not None else "" for cell in raw]
        for split in split_multiline_row(cells):
            while len(split) < max(first_job + 1, len(header)):
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
            row_unit = unit
            if unit_col is not None and unit_col < len(split) and flatten_cell(split[unit_col]):
                row_unit = display_unit(split[unit_col])
            if row_unit in {"〃", "＂", "-"} or "〃" in (row_unit or ""):
                row_unit = last_unit or unit
            spec = expand_ditto(spec, last_spec, unit_hint=hint)
            spec, row_unit = clean_spec_unit(spec, row_unit, prev_spec=last_spec, unit_hint=hint, kind=kind)
            if not spec or spec in {"계", "합계", "소계"}:
                continue
            last_spec = spec
            last_unit = row_unit or last_unit
            parsed = False
            for col, job in job_cols:
                if col >= len(split):
                    continue
                qty = single_number(split[col])
                if qty is None:
                    continue
                emit(rows, title, spec, row_unit or unit, job, qty, code, prefix)
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
) -> int:
    if not table or not code or code in skip_codes:
        return 0
    added = parse_paired_size_table(table, code, title, unit, rows, discipline, prefix)
    if added:
        return added
    return parse_labor_table(table, code, title, unit, rows, discipline, prefix)


def parse_item_line(line: str, discipline: str) -> tuple[str, str] | None:
    if TOC_DOTS.search(line):
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
    table_codes: set[str] = set()

    for page in pdf.pages:
        text = page.extract_text() or ""
        word_heads: list[tuple[float, str, str]] = []
        by_top: dict[float, list[str]] = {}
        for word in page.extract_words() or []:
            by_top.setdefault(round(float(word["top"]), 1), []).append(word["text"])
        for top, toks in sorted(by_top.items()):
            line = normalize_spaces(" ".join(toks))
            parsed = parse_item_line(line, discipline)
            if parsed:
                word_heads.append((top, parsed[0], parsed[1]))
                current_code, current_title = parsed
                full = clean_title(parsed[1])
                if len(full) > len(titles_by_code.get(parsed[0], "")):
                    titles_by_code[parsed[0]] = full
        for raw in text.splitlines():
            line = normalize_spaces(raw)
            parsed_item = parse_item_line(line, discipline)
            if parsed_item:
                current_code, current_title = parsed_item
                current_unit = ""
                full = clean_title(parsed_item[1])
                if len(full) > len(titles_by_code.get(parsed_item[0], "")):
                    titles_by_code[parsed_item[0]] = full
            unit_match = UNIT_LINE.search(line)
            if unit_match:
                current_unit = display_unit(unit_match.group(1))

        page_tables = []
        try:
            page_tables = page.find_tables() or []
        except Exception:
            page_tables = []
        if not page_tables:
            try:
                extracted = page.extract_tables() or []
            except Exception:
                extracted = []
            for tbl in extracted:
                added = parse_table(
                    tbl, current_code, current_title, current_unit, rows, discipline, prefix, skip_codes
                )
                if added:
                    table_codes.add(current_code)
        for tbl in page_tables:
            ty = tbl.bbox[1]
            code, title = current_code, current_title
            for hy, hc, ht in word_heads:
                if hy <= ty + 18:
                    code, title = hc, ht
            try:
                extracted = tbl.extract()
            except Exception:
                continue
            added = parse_table(extracted, code, title, current_unit, rows, discipline, prefix, skip_codes)
            if added:
                table_codes.add(code)

        jobs = []
        apply_job = ""
        last_spec = ""
        last_kind = ""
        header_buf = []
        line_code = current_code
        line_title = current_title
        line_unit = current_unit
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
            if spec in {"계", "합계", "소계"}:
                continue
            last_spec = spec
            for job, qty in zip(use_jobs, values):
                if qty is None:
                    continue
                emit(rows, line_title, spec, unit, job, qty, line_code, prefix)

    pdf.close()

    def title_is_bad(name: str) -> bool:
        text = str(name or "").strip()
        if len(text) < 8:
            return True
        if not re.search(r"[가-힣]", text):
            return True
        if text.endswith(")") and " " not in text:
            return True
        return False

    for row in rows:
        code = str(row.get("품셈근거") or "").replace(prefix, "", 1)
        better = titles_by_code.get(code)
        if better and title_is_bad(str(row.get("명칭") or "")):
            row["명칭"] = better
            row["짧은명칭"] = short_name(better)
        if str(row.get("단위") or "") in {"식", ""} and ("단면적" in str(row.get("규격") or "") or "Box" in str(row.get("규격") or "")):
            row["단위"] = "개"

    merged: dict[str, dict] = {}
    scores: dict[str, int] = {}

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

    for row in rows:
        key = f"{row['명칭']}|{row['규격']}|{row['노무명칭']}|{row['품셈근거']}"
        merged[key] = row
        qty_key = f"{row['명칭']}|{row['노무명칭']}|{row['품셈근거']}|{row['품셈']}"
        prev = scores.get(qty_key)
        score = spec_score(str(row["규격"]))
        if prev is None or score > prev:
            scores[qty_key] = score
            merged[qty_key] = row

    # keep identity keys plus best-spec-per-qty (qty_key overwrites poorer specs)
    by_qty: dict[str, dict] = {}
    for row in merged.values():
        qty_key = f"{row['명칭']}|{row['노무명칭']}|{row['품셈근거']}|{row['품셈']}"
        current = by_qty.get(qty_key)
        if current is None or spec_score(str(row["규격"])) > spec_score(str(current["규격"])):
            by_qty[qty_key] = row
    final: dict[str, dict] = {}
    for row in by_qty.values():
        key = f"{row['명칭']}|{row['규격']}|{row['노무명칭']}|{row['품셈근거']}"
        final[key] = row
    return list(final.values())


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
