#!/usr/bin/env python3
"""전기공사 표준품셈 PDF에서 직종 열이 있는 표를 품셈 행으로 푼다."""

from __future__ import annotations

import json
import re
from pathlib import Path

import pdfplumber

PDF = Path("/tmp/pumsam/2025_elec.pdf")
OUT = Path("/tmp/pumsam/book_rows.json")

# 긴 이름 먼저 맞춘다.
JOB_ALIASES: list[tuple[str, str]] = [
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

# 헤더에서 공백이 끼는 표기
SPACED = [
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

SKIP_SPEC = re.compile(r"^(합계|소계|계|구분|공종|규격|종별|명칭|용량|직종|단위|비고|적용직종)")
MANY_FLOATS = re.compile(r"\d+\.\d+")
CAPACITY = re.compile(r"(\d+(?:\.\d+)?\s*(?:kVA|MVA|kV|㎸|㎸A|A)\s*(?:이하|초과|미만)?)", re.I)
PROCESS_ONLY = re.compile(r"^(운반|소운반|OT|점검|설치|배치|조가|교정|도장|높이)")
ITEM_HEAD = re.compile(
    r"^([2-7])-(\d{1,2}(?:-\d{1,2})?)\s+(.+?)(?:\s*\([’'`]?\d{2}년.*)?$"
)
UNIT_LINE = re.compile(r"단위\s*[:：]\s*([^)\n]+)")
APPLY_JOB = re.compile(r"적용직종\s*[:：]\s*([가-힣]+)")
NUM_TOKEN = re.compile(r"^(?:-|\d{1,4}(?:,\d{3})*(?:\.\d+)?)$")
PAGE_NOISE = re.compile(
    r"^(제[0-9]+장|전기부문|\d{2,4}\s*전기부문|\d+\s*$|【해설】|주\d+\)|※)"
)


def normalize_spaces(text: str) -> str:
    return re.sub(r"\s+", " ", (text or "").replace("\xa0", " ")).strip()


def looks_like_header_piece(line: str) -> bool:
    compact = line.replace(" ", "")
    keys = ("규격", "종별", "명칭", "공종", "용량", "직종", "전공", "인부", "기술자")
    if not any(k in compact for k in keys):
        return False
    if ANY_DATA_NUM.search(line) and not any(k in compact for k in ("전공", "인부", "기술자")):
        return False
    return True


ANY_DATA_NUM = re.compile(r"\d+\.\d+")


def detect_jobs(line: str) -> list[str]:
    raw = normalize_spaces(line)
    compact = raw.replace(" ", "")
    # 직종 헤더는 보통 2개 이상 직종이거나, 직종+계
    found: list[tuple[int, str]] = []
    search_from = 0
    work = raw
    for alias, canon in SPACED + JOB_ALIASES:
        idx = work.find(alias)
        if idx >= 0:
            found.append((idx, canon))
            work = work.replace(alias, " " * len(alias), 1)
    if not found:
        # compact fallback
        pos = 0
        tmp = compact
        for alias, canon in JOB_ALIASES:
            idx = tmp.find(alias)
            if idx >= 0:
                found.append((idx, canon))
                tmp = tmp.replace(alias, " " * len(alias), 1)
    found.sort()
    jobs = [j for _, j in found]
    # 헤더로 보려면 직종이 1개 이상이면서 숫자 데이터가 없어야 함
    if not jobs:
        return []
    if any(NUM_TOKEN.match(tok.replace(",", "")) for tok in raw.split() if tok not in {"-"}):
        # 데이터 행일 수 있음. 다만 직종명만 있는 헤더는 숫자 없음
        nums = [tok for tok in raw.split() if NUM_TOKEN.match(tok.replace(",", ""))]
        if nums:
            return []
    # '계'만 있는 열은 인부가 아님
    jobs = [j for j in jobs if j != "계"]
    return jobs


def parse_nums(line: str, n: int) -> tuple[str, list[float | None]] | None:
    parts = normalize_spaces(line).split()
    if not parts:
        return None
    nums: list[float | None] = []
    while parts and len(nums) < n:
        tok = parts[-1].replace(",", "")
        if tok in {"-", "－", "—"}:
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
    if not spec:
        return None
    if SKIP_SPEC.match(spec.replace(" ", "")) and spec.replace(" ", "") in {"합계", "소계", "계"}:
        return None
    if spec in {"규격", "종별", "명칭", "공종", "용량", "구분"}:
        return None
    return spec, nums


def chapter_of(code: str) -> int:
    return int(code.split("-")[0])


def short_name(title: str) -> str:
    text = title
    text = re.sub(r"\s*\(.*\)\s*$", "", text)
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
        " 설치",
        " 신설",
    ):
        if text.endswith(tail):
            text = text[: -len(tail)]
            break
    return text.strip() or title


def cell_job(text: str) -> str | None:
    jobs = detect_jobs(text or "")
    if len(jobs) == 1:
        return jobs[0]
    compact = normalize_spaces(text or "").replace(" ", "")
    for alias, canon in JOB_ALIASES:
        if compact == alias.replace(" ", ""):
            return canon
    return None


def single_number(text: str) -> float | None:
    tok = normalize_spaces(text or "").replace(",", "")
    if tok in {"", "-", "－", "—"}:
        return None
    if NUM_TOKEN.match(tok):
        return float(tok)
    return None


def expand_ditto(spec: str, prev: str) -> str:
    spec = spec.replace("＂", "〃").replace("''", "〃")
    if "〃" not in spec:
        return spec
    suffix = ""
    m = re.search(r"(이하|초과|내외|미만|단심)$", prev)
    if m:
        suffix = m.group(1)
    spec = spec.replace("〃", suffix).strip()
    return re.sub(r"\s+", " ", spec)


def emit(rows: list[dict], title: str, spec: str, unit: str, job: str, qty: float, code: str) -> None:
    if qty is None or qty <= 0:
        return
    if not spec or not job or not title or not code:
        return
    if len(MANY_FLOATS.findall(spec)) >= 2:
        return
    if PROCESS_ONLY.match(spec) and not CAPACITY.search(spec):
        return
    rows.append(
        {
            "명칭": title,
            "짧은명칭": short_name(title),
            "규격": spec,
            "단위": unit,
            "노무명칭": job,
            "품셈": round(qty, 4) if abs(qty) < 1000 else qty,
            "할증%": 100,
            "품셈근거": f"전기{code}",
        }
    )


SKIP_CODES = {"2-1", "2-1-1", "2-1-2", "2-1-3", "2-10", "2-10-1", "2-10-2"}


def parse_table(table: list[list], code: str, title: str, unit: str, rows: list[dict]) -> int:
    if not table or len(table) < 2 or not code or code in SKIP_CODES:
        return 0
    header = [normalize_spaces(c or "") for c in table[0]]
    job_cols: list[tuple[int, str]] = []
    for i, cell in enumerate(header):
        job = cell_job(cell)
        if job:
            job_cols.append((i, job))
    if len(job_cols) < 2:
        return 0
    added = 0
    last_capacity = ""
    before = len(rows)
    for raw in table[1:]:
        cells = [normalize_spaces(c or "") for c in raw]
        if not cells:
            continue
        left = " ".join(cells[: job_cols[0][0]]).strip()
        cap = CAPACITY.search(left)
        if cap:
            last_capacity = cap.group(1)
        spec = left
        if spec in {"계", "합계", "소계", ""}:
            spec = last_capacity or spec
        if spec in {"계", "합계", "소계", ""}:
            continue
        ok = True
        parsed: list[tuple[str, float]] = []
        for col, job in job_cols:
            if col >= len(cells):
                ok = False
                break
            qty = single_number(cells[col])
            if qty is None:
                if cells[col] in {"", "-", "－"}:
                    continue
                ok = False
                break
            parsed.append((job, qty))
        if not ok or not parsed:
            continue
        for job, qty in parsed:
            emit(rows, title, spec, unit, job, qty, code)
            added += 1
    return len(rows) - before


def extract() -> list[dict]:
    pdf = pdfplumber.open(str(PDF))
    rows: list[dict] = []
    current_code = ""
    current_title = ""
    current_unit = ""
    jobs: list[str] = []
    apply_job = ""
    last_spec = ""
    last_capacity = ""
    header_buf: list[str] = []

    for page in pdf.pages:
        text = page.extract_text() or ""
        page_heads: list[tuple[float, str, str]] = []
        for raw in text.splitlines():
            line = normalize_spaces(raw)
            if not line or PAGE_NOISE.match(line):
                continue
            m = ITEM_HEAD.match(line)
            if m:
                current_code = f"{m.group(1)}-{m.group(2)}"
                current_title = normalize_spaces(m.group(3))
                current_title = re.sub(r"\s*\([’'`´]?[0-9]{2}년.*$", "", current_title).strip()
                page_heads.append((0.0, current_code, current_title))
                current_unit = ""
                jobs = []
                apply_job = ""
                last_spec = ""
                last_capacity = ""
                header_buf = []
                continue
            um = UNIT_LINE.search(line)
            if um:
                unit_blob = um.group(1)
                current_unit = unit_blob.split(",")[0].strip()
                current_unit = re.sub(r"적용직종.*", "", current_unit).strip()
            am = APPLY_JOB.search(line)
            if am:
                apply_job = am.group(1).replace(" ", "")
            if looks_like_header_piece(line) and not MANY_FLOATS.search(line):
                header_buf.append(line)
                header_buf = header_buf[-4:]
                detected = detect_jobs(" ".join(header_buf))
                if len(detected) >= 2 or (len(detected) == 1 and apply_job):
                    jobs = detected
                    header_buf = []
                    continue
            detected = detect_jobs(line)
            if len(detected) >= 2:
                jobs = detected
                header_buf = []
                last_spec = ""
                continue
            if not current_code or current_code in SKIP_CODES or chapter_of(current_code) < 2:
                continue
            use_jobs = jobs if jobs else ([apply_job] if apply_job else [])
            if not use_jobs:
                continue
            parsed = parse_nums(line, len(use_jobs))
            if not parsed:
                cap = CAPACITY.search(line)
                if cap and not MANY_FLOATS.search(line):
                    last_capacity = cap.group(1)
                continue
            spec, values = parsed
            spec = expand_ditto(spec, last_spec)
            if spec in {"계", "합계", "소계"} or spec.startswith("계 "):
                spec = last_capacity or spec
            cap = CAPACITY.search(spec) or CAPACITY.search(line)
            if cap:
                last_capacity = cap.group(1)
                if spec in {"계", "합계", "소계"}:
                    spec = last_capacity
            if spec in {"계", "합계", "소계"}:
                continue
            last_spec = spec
            for job, qty in zip(use_jobs, values):
                if qty is None:
                    continue
                emit(rows, current_title, spec, current_unit, job, qty, current_code)

        by_top: dict[float, list[str]] = {}
        for word in page.extract_words() or []:
            by_top.setdefault(round(float(word["top"]), 1), []).append(word["text"])
        word_heads: list[tuple[float, str, str]] = []
        for top, toks in sorted(by_top.items()):
            line = normalize_spaces(" ".join(toks))
            matched = ITEM_HEAD.match(line)
            if matched:
                title = normalize_spaces(matched.group(3))
                title = re.sub(r"\s*\([’'`´]?[0-9]{2}년.*$", "", title).strip()
                word_heads.append((top, f"{matched.group(1)}-{matched.group(2)}", title))
        try:
            found_tables = page.find_tables() or []
        except Exception:
            found_tables = []
        for tbl in found_tables:
            ty = tbl.bbox[1]
            code, title = current_code, current_title
            for hy, hc, ht in word_heads:
                if hy <= ty + 12:
                    code, title = hc, ht
            parse_table(tbl.extract(), code, title, current_unit, rows)

    pdf.close()
    # 같은 명칭+규격+직종은 한 줄
    merged: dict[str, dict] = {}
    for row in rows:
        key = f"{row['명칭']}|{row['규격']}|{row['노무명칭']}|{row['품셈근거']}"
        merged[key] = row
    return list(merged.values())


def main() -> None:
    rows = extract()
    OUT.write_text(json.dumps(rows, ensure_ascii=False, indent=0), encoding="utf-8")
    from collections import Counter

    by_ch = Counter(r["품셈근거"].split("-")[0].replace("전기", "") for r in rows)
    jobs = Counter(r["노무명칭"] for r in rows)
    print("rows", len(rows))
    print("by chapter", dict(sorted(by_ch.items())))
    print("jobs", dict(jobs.most_common(20)))
    print("sample")
    for row in rows[:8]:
        print(row)


if __name__ == "__main__":
    main()
