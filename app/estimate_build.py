"""체크박스로 고른 일위대가목록 품목으로 공정별 내역서를 만든다.

내역서 양식
- 맨 윗줄: 공정 제목(사용자가 입력)
- 품목 행: 재료비 단가 = 일위대가 자재 행 금액(자재 수량 할증 포함), 노무비 단가 = 일위대가 인부 금액 합
- 아래 집계: [배관 부속재] · [소모 잡자재] · 직종별 노무비 · [공구손료] · 합계
  합계 재료비 = 품목 재료비 + 부속재 + 잡자재, 합계 노무비 = 직종별 노무비 + 공구손료
"""

from __future__ import annotations

import math
import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from openpyxl import Workbook
from openpyxl.styles import Alignment
from openpyxl.utils import get_column_letter

from app.discipline import normalize_discipline
from app.estimate_parse import load_estimate_sheet
from app.excel_io import (
    apply_form_widths,
    BODY_FONT,
    CENTER,
    COMPARE_SHEET_NAME,
    ESTIMATE_SHEET_NAME,
    FORM_ROW_HEIGHT,
    HEADER_FONT,
    ILWIDAE_LIST_SHEET_NAME,
    ILWIDAE_SHEET_NAME,
    LEFT,
    RIGHT,
    _apply_sheet_look,
    _set_cell,
    collect_pumsam_rows,
    write_title_banner,
)
from app.fitting_rules import load_fitting_excludes
from app.ilwidae import CD_FITTING_RATE, SUNDRY_RATE, TOOL_RATE, IlwidaeBlock
from app.paths import assert_safe_save, build_result_path, save_workbook_safely
from app.wages import WAGES_SHEET_NAME, load_wages, save_wages

MONEY_FORMAT = "#,##0;-#,##0;;@"
QTY_INT_FORMAT = "#,##0"
QTY_DEC_FORMAT = "#,##0.0##"


def qty_format(value) -> str:
    """수량 뒤에 점이 남지 않게, 정수는 정수로 소수는 소수로 보인다."""
    try:
        return QTY_INT_FORMAT if float(value) == int(float(value)) else QTY_DEC_FORMAT
    except (TypeError, ValueError):
        return QTY_INT_FORMAT
LAST_COL = 12
BLANK_ROWS_BEFORE_TOTAL = 6
HELPER_START_COL = 14  # N 열부터는 계산용 보조 열(숨김)
_BAD_SHEET_CHARS = re.compile(r"[\[\]:*?/\\]")


def trunc1(value: float) -> float:
    return math.floor(value * 10 + 1e-9) / 10


@dataclass
class IlwidaeEntry:
    """일위대가목록 한 줄(= 일위대가 호표 하나)."""

    index: int
    name: str
    spec: str
    unit: str
    ho_no: int
    material_row: int
    labor_rows: dict[str, list[int]]  # 직종 → 일위대가 시트의 인부 행들
    material_unit: float  # 자재 수량(할증 포함) × 단가
    labor_unit_by_job: dict[str, float]
    fitting_rate: float
    has_sundry: bool
    pumsam_ref: str = ""
    matched: bool = True  # 품셈 인부가 붙었는지(미매칭이면 False)

    @property
    def labor_unit(self) -> float:
        return sum(self.labor_unit_by_job.values())

    @property
    def label(self) -> str:
        return f"{self.name} {self.spec}".strip()


@dataclass
class EstimateLine:
    entry: IlwidaeEntry
    qty: float
    remark: str = ""


@dataclass
class EstimateSection:
    title: str
    lines: list[EstimateLine] = field(default_factory=list)


@dataclass
class EstimateSession:
    workbook: Workbook
    entries: list[IlwidaeEntry]
    discipline: str
    source_path: Path | None = None
    estimate_sheets: list[str] = field(default_factory=list)  # 앞서 저장할 때 붙인 내역서 시트
    pumsam_rows: list = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    notices: list[str] = field(default_factory=list)
    db_dir: Path | None = None


def extract_entries(workbook: Workbook, blocks: list[IlwidaeBlock], wage_rows: list[dict]) -> list[IlwidaeEntry]:
    wages = {str(r.get("직종")): float(r.get("노임단가") or 0) for r in wage_rows}
    compare = workbook[COMPARE_SHEET_NAME]
    entries: list[IlwidaeEntry] = []
    for block in blocks:
        item = block.item
        col = (item.price_col if item.price_col is not None else 11) + 1
        price = compare.cell(item.excel_row, col).value
        price = float(price) if isinstance(price, (int, float)) else 0.0
        labor_rows: dict[str, list[int]] = {}
        labor_unit: dict[str, float] = {}
        for job, qty, row in zip(block.labor_jobs, block.labor_qtys, block.labor_rows):
            labor_rows.setdefault(job, []).append(row)
            labor_unit[job] = labor_unit.get(job, 0.0) + trunc1(wages.get(job, 0.0) * qty)
        entries.append(
            IlwidaeEntry(
                index=len(entries),
                name=str(item.name or ""),
                spec=str(item.spec or ""),
                unit=str(item.unit or ""),
                ho_no=block.ho_no,
                material_row=block.material_row,
                labor_rows=labor_rows,
                material_unit=trunc1(block.material_qty * price),
                labor_unit_by_job=labor_unit,
                fitting_rate=block.fitting_rate,
                has_sundry=block.has_sundry,
                pumsam_ref=block.pumsam_ref,
                matched=bool(block.pumsam_ref) or any(q > 0 for q in block.labor_qtys),
            )
        )
    return entries


def generate_session(
    unit_price_path: Path,
    discipline: str | None = None,
    db_dir: Path | None = None,
) -> EstimateSession:
    """단가대비표 → 일위대가·일위대가목록을 메모리에 만들고, 고를 수 있는 품목 목록을 돌려준다."""
    from app.pipeline import _load_optional, build_result_workbook

    from app.pumsam_links import activate

    disc = normalize_discipline(discipline)
    source = Path(unit_price_path)
    from app import numbers

    activate(db_dir)
    numbers.NOTICES.clear()
    compare = _load_optional(source)
    pumsam_rows = collect_pumsam_rows(source, db_dir, None, db_dir=db_dir, discipline=disc)
    wage_rows = load_wages(db_dir)
    save_wages(wage_rows, db_dir)
    blocks: list[IlwidaeBlock] = []
    workbook = build_result_workbook(
        compare=compare,
        estimate=None,
        pumsam_rows=pumsam_rows,
        wage_rows=wage_rows,
        discipline=disc,
        mode="forward",
        fitting_excludes=load_fitting_excludes(db_dir),
        blocks_out=blocks,
    )
    entries = extract_entries(workbook, blocks, wage_rows)
    session = EstimateSession(
        workbook=workbook, entries=entries, discipline=disc, source_path=source, pumsam_rows=pumsam_rows, db_dir=db_dir
    )
    session.notices = list(numbers.NOTICES)
    return session


_PCT_RE = re.compile(r"(\d+(?:\.\d+)?)\s*%")
_CELL_REF_RE = re.compile(r"^=\s*'?([^'!]+)'?!\$?([A-Z]+)\$?(\d+)$")
_PRODUCT_RE = re.compile(r"^=?\s*\d+(?:\.\d+)?(?:\s*\*\s*\d+(?:\.\d+)?)*$")


def _number(value: Any) -> float | None:
    """숫자 칸 또는 `=0.04*0.7` 같은 단순 곱 수식을 숫자로."""
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str) and _PRODUCT_RE.match(value.strip()):
        result = 1.0
        for part in value.strip().lstrip("=").split("*"):
            result *= float(part)
        return result
    return None


CHECK_SHEET_NAME = "점검목록"
AUTOSAVE_NAME = "임시저장.json"


def autosave_path(dest_dir: Path | None) -> Path:
    from app.paths import ensure_result_directory

    return ensure_result_directory(dest_dir) / AUTOSAVE_NAME


def write_autosave(path: Path, session: EstimateSession, sections: list[EstimateSection]) -> None:
    """담아 둔 공정을 파일로 남겨, 프로그램이 꺼져도 다시 불러올 수 있게 한다."""
    import json

    if not sections:
        path.unlink(missing_ok=True)
        return
    payload = {
        "source": session.source_path.name if session.source_path else "",
        "discipline": session.discipline,
        "sections": [
            {
                "title": sec.title,
                "lines": [{"ho": ln.entry.ho_no, "qty": ln.qty, "remark": ln.remark} for ln in sec.lines],
            }
            for sec in sections
        ],
    }
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=1), encoding="utf-8")


def read_autosave(path: Path, session: EstimateSession) -> list[EstimateSection]:
    """임시저장 파일에서 같은 원본의 공정을 되살린다. 맞지 않으면 빈 목록."""
    import json

    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return []
    if data.get("source") != (session.source_path.name if session.source_path else ""):
        return []
    by_ho = {e.ho_no: e for e in session.entries}
    sections: list[EstimateSection] = []
    for sec in data.get("sections", []):
        lines = [
            EstimateLine(by_ho[int(item["ho"])], float(item["qty"]), str(item.get("remark") or ""))
            for item in sec.get("lines", [])
            if int(item.get("ho", 0)) in by_ho
        ]
        if lines:
            sections.append(EstimateSection(str(sec.get("title") or ""), lines))
    return sections


def duplicate_rows(entries: list[IlwidaeEntry]) -> set[int]:
    """같은 품명·규격이 앞에 이미 있는 줄(0부터 센 표 줄 번호)."""
    from app.estimate_parse import lookup_key

    seen: set[str] = set()
    dup: set[int] = set()
    for index, entry in enumerate(entries):
        key = lookup_key(entry.name, entry.spec)
        if key in seen:
            dup.add(index)
        seen.add(key)
    return dup


def _write_check_sheet(workbook: Workbook, entries: list[IlwidaeEntry]) -> None:
    """품셈 미매칭·중복 품목을 한 시트에 모아 둔다. 문제가 없으면 시트를 만들지 않는다."""
    if CHECK_SHEET_NAME in workbook.sheetnames:
        del workbook[CHECK_SHEET_NAME]
    dup = duplicate_rows(entries)
    problems = [(e, "품셈 미매칭" if not e.matched else "") for e in entries]
    problems = [(e, why) for i, (e, why) in enumerate(problems) if why or i in dup]
    if not problems:
        return
    sheet = workbook.create_sheet(CHECK_SHEET_NAME)
    sheet.append(["호표", "명칭", "규격", "단위", "점검 내용", "조치"])
    index_of = {id(e): i for i, e in enumerate(entries)}
    for entry, why in problems:
        notes = [why] if why else []
        if index_of[id(entry)] in dup:
            notes.append("같은 품명·규격이 중복")
        action = "품셈 연결 기능으로 연결하거나 일위대가의 인부 수량을 직접 입력" if not entry.matched else "하나만 남길지 확인"
        sheet.append([entry.ho_no, entry.name, entry.spec, entry.unit, " · ".join(notes), action])
    for col, width in zip("ABCDEF", (8, 30, 30, 8, 24, 50)):
        sheet.column_dimensions[col].width = width


def layout_ok(path: Path) -> bool:
    """이 프로그램이 읽는 단가대비표 양식인지(품목이 하나 이상 읽히는지)."""
    from app.items import parse_line_items

    try:
        sheet = load_estimate_sheet(Path(path))
        items = [i for i in parse_line_items(sheet) if not i.section and i.name]
        with_form = [i for i in items if i.spec or i.unit]
        priced = [i for i in with_form if isinstance(i.material_price, (int, float))]
        return bool(with_form) and bool(priced)
    except Exception:
        return False


def read_preview(path: Path, rows: int = 14) -> list[list[Any]]:
    from app.numbers import NOTICES

    sheet = load_estimate_sheet(Path(path))
    del NOTICES[:]
    return [list(r) for r in sheet.filled[:rows]]


def convert_with_columns(
    path: Path,
    out_dir: Path,
    *,
    first_data_row: int,
    name_col: int,
    spec_col: int | None,
    unit_col: int | None,
    price_col: int,
) -> Path:
    """사용자가 열을 직접 지정한 낯선 양식을 이 프로그램의 단가대비표 양식으로 바꿔 새 파일로 쓴다. 열 번호는 1부터."""
    from openpyxl import Workbook

    from app.numbers import clean_price

    grid = load_estimate_sheet(Path(path)).filled
    wb = Workbook()
    ws = wb.active
    ws.title = "단가대비표"
    ws.append(["단 가 대 비 표"])
    ws.append([])
    ws.append(["코  드", "품      명", "규격", "단위", "재  료  비"])
    ws.append([None, None, None, None, "적용단가"])

    def cell(row: list[Any], col: int | None) -> Any:
        return row[col - 1] if col and col - 1 < len(row) else None

    for row in grid[first_data_row - 1 :]:
        name = cell(row, name_col)
        if name in (None, ""):
            continue
        ws.append([None, name, cell(row, spec_col), cell(row, unit_col), clean_price(cell(row, price_col))])
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    dest = out_dir / f"변환_{Path(path).stem}.xlsx"
    wb.save(dest)
    return dest


def is_result_workbook(path: Path) -> bool:
    """이미 만들어 둔 일위대가·일위대가목록 파일인지(내역서만 이어서 만들 수 있는 파일)."""
    from openpyxl import load_workbook

    try:
        wb = load_workbook(path, read_only=True)
        try:
            return ILWIDAE_SHEET_NAME in wb.sheetnames and ILWIDAE_LIST_SHEET_NAME in wb.sheetnames
        finally:
            wb.close()
    except Exception:
        return False


def load_session_from_result(result_path: Path, discipline: str | None = None) -> EstimateSession:
    """저장해 둔 일위대가·일위대가목록 파일을 읽어, 내역서만 이어서 만들 수 있게 한다."""
    from openpyxl import load_workbook

    disc = normalize_discipline(discipline)
    path = Path(result_path)
    workbook = load_workbook(path)
    if ILWIDAE_SHEET_NAME not in workbook.sheetnames:
        raise ValueError("일위대가 시트가 없는 파일입니다. 일위대가목록까지 저장한 결과 파일을 넣어 주세요.")
    sheet = workbook[ILWIDAE_SHEET_NAME]
    compare = workbook[COMPARE_SHEET_NAME] if COMPARE_SHEET_NAME in workbook.sheetnames else None
    wages: dict[str, float] = {}
    if WAGES_SHEET_NAME in workbook.sheetnames:
        for row in workbook[WAGES_SHEET_NAME].iter_rows(min_row=2, max_col=2, values_only=True):
            if row[0] and isinstance(row[1], (int, float)):
                wages[str(row[0]).strip()] = float(row[1])
    if not wages:
        wages = {str(r.get("직종")): float(r.get("노임단가") or 0) for r in load_wages(None)}

    def price_of(cell_value: Any) -> float:
        direct = _number(cell_value)
        if direct is not None and not (isinstance(cell_value, str) and cell_value.startswith("=")):
            return direct
        if isinstance(cell_value, str):
            matched = _CELL_REF_RE.match(cell_value.strip())
            if matched and compare is not None and matched.group(1) == COMPARE_SHEET_NAME:
                value = compare[f"{matched.group(2)}{matched.group(3)}"].value
                if isinstance(value, (int, float)):
                    return float(value)
        return 0.0

    entries: list[IlwidaeEntry] = []
    warnings: list[str] = []
    max_row = sheet.max_row
    row = 5
    while row <= max_row:
        title = sheet.cell(row, 1).value
        if not (isinstance(title, str) and "호표" in title and sheet.cell(row, 2).value is None and sheet.cell(row, 3).value is None):
            row += 1
            continue
        # 중간에 호표를 끼워 넣거나 지웠어도 위에서부터 1, 2, 3… 으로 번호를 다시 매긴다.
        ho_no = len(entries) + 1
        new_title = re.sub(r"호표\s*\d*", f"호표 {ho_no}", title, count=1)
        if new_title != title:
            sheet.cell(row, 1).value = new_title
        material_row = row + 1
        end = material_row
        found_sum = False
        while end + 1 <= max_row:
            head = sheet.cell(end + 1, 1).value
            if isinstance(head, str) and head.strip().startswith("[ 합"):
                end += 1
                found_sum = True
                break
            if isinstance(head, str) and "호표" in head and sheet.cell(end + 1, 2).value is None and sheet.cell(end + 1, 3).value is None:
                break  # 다음 호표가 시작되었는데 합계 줄이 없다
            end += 1
        if not found_sum:
            warnings.append(f"일위대가 {row}행 「{str(title).strip()[:30]}」 호표에 [ 합계 ] 줄이 없어 건너뛰었습니다.")
            row += 1
            continue
        sum_row = end
        qty = _number(sheet.cell(material_row, 4).value) or 1.0
        labor_rows: dict[str, list[int]] = {}
        labor_unit: dict[str, float] = {}
        fitting_rate = 0.0
        has_sundry = False
        ref = ""
        for r in range(material_row + 1, sum_row):
            label = str(sheet.cell(r, 1).value or "").strip()
            unit = str(sheet.cell(r, 3).value or "").strip()
            if label == "전선관부속품비":
                found = _PCT_RE.search(str(sheet.cell(r, 2).value or ""))
                fitting_rate = float(found.group(1)) / 100 if found else 0.0
            elif label == "잡재료비":
                has_sundry = True
            elif unit == "인" and label:
                labor_rows.setdefault(label, []).append(r)
                q = _number(sheet.cell(r, 4).value) or 0.0
                labor_unit[label] = labor_unit.get(label, 0.0) + trunc1(wages.get(label, 0.0) * q)
        head_note = str(sheet.cell(material_row, 13).value or "")
        ref = head_note.split(" · ")[0].strip() if head_note else ""
        if not str(sheet.cell(material_row, 1).value or "").strip():
            warnings.append(f"일위대가 {material_row}행: 자재 줄의 품명이 비어 있습니다 (호표 {ho_no}).")
        elif _number(sheet.cell(material_row, 4).value) is None:
            warnings.append(f"일위대가 {material_row}행 「{sheet.cell(material_row, 1).value}」: 자재 수량을 숫자로 읽지 못해 1로 봤습니다.")
        if not labor_rows:
            warnings.append(f"일위대가 {material_row}행 「{sheet.cell(material_row, 1).value}」: 인부 줄이 없습니다.")
        entries.append(
            IlwidaeEntry(
                index=len(entries),
                name=str(sheet.cell(material_row, 1).value or "").strip(),
                spec=str(sheet.cell(material_row, 2).value or "").strip(),
                unit=str(sheet.cell(material_row, 3).value or "").strip(),
                ho_no=ho_no,
                material_row=material_row,
                labor_rows=labor_rows,
                material_unit=trunc1(qty * price_of(sheet.cell(material_row, 5).value)),
                labor_unit_by_job=labor_unit,
                fitting_rate=fitting_rate,
                has_sundry=has_sundry,
                pumsam_ref=ref,
                matched=any(
                    (_number(sheet.cell(r, 4).value) or 0) > 0 for rows_ in labor_rows.values() for r in rows_
                ),
            )
        )
        row = sum_row + 1
    if not entries:
        raise ValueError("일위대가 호표를 찾지 못했습니다. 이 프로그램으로 만든 일위대가목록 파일인지 확인해 주세요.")
    session = EstimateSession(workbook=workbook, entries=entries, discipline=disc, source_path=path)
    session.warnings = warnings
    session.estimate_sheets = [n for n in workbook.sheetnames if n == ESTIMATE_SHEET_NAME or n.startswith(f"{ESTIMATE_SHEET_NAME}(")]
    return session


def open_session(path: Path, discipline: str | None = None, db_dir: Path | None = None) -> EstimateSession:
    """단가대비표면 새로 일위대가를 만들고, 일위대가목록 결과 파일이면 그대로 이어서 읽는다."""
    if is_result_workbook(Path(path)):
        return load_session_from_result(Path(path), discipline)
    return generate_session(Path(path), discipline, db_dir)


def _sheet_name(base: str, taken: set[str]) -> str:
    name = _BAD_SHEET_CHARS.sub("", base)[:31] or ESTIMATE_SHEET_NAME
    candidate = name
    n = 2
    while candidate in taken:
        suffix = f"({n})"
        candidate = name[: 31 - len(suffix)] + suffix
        n += 1
    return candidate


def _ref(row: int, col: str) -> str:
    return f"'{ILWIDAE_SHEET_NAME}'!{col}{row}"


_LEADING_NUMBER = re.compile(r"^\s*\d+\s*[.)．]\s*")


def strip_section_number(title: str) -> str:
    """사용자가 「1. 내선설비공사」처럼 번호를 붙여 적어도 제목만 남긴다. 번호는 순서대로 자동으로 붙인다."""
    text = str(title or "").strip()
    while True:
        cut = _LEADING_NUMBER.sub("", text, count=1)
        if cut == text:
            return text
        text = cut.strip()


def section_label(number: int, title: str) -> str:
    return f"{number}. {strip_section_number(title)}"


def write_estimate_sheet(workbook: Workbook, name: str, sections: "EstimateSection | list[EstimateSection]") -> None:
    """내역서 시트 하나에 공정을 순서대로 이어 쓴다. 제목 줄은 「1. 공정명」으로 번호가 자동으로 붙는다."""
    if isinstance(sections, EstimateSection):
        sections = [sections]
    sheet = workbook.create_sheet(name)
    write_title_banner(sheet, "내   역   서", LAST_COL)
    heads = {1: "명  칭", 2: "규  격", 3: "단위", 4: "수량", 5: "재  료  비", 7: "노  무  비", 9: "경    비", 11: "계", 12: "비  고"}
    for col in range(1, LAST_COL + 1):
        _set_cell(sheet, 3, col, heads.get(col), font=HEADER_FONT, align=CENTER)
        _set_cell(sheet, 4, col, None, font=HEADER_FONT, align=CENTER)
    for col, text in ((5, "단  가"), (6, "금  액"), (7, "단  가"), (8, "금  액"), (9, "단  가"), (10, "금  액")):
        _set_cell(sheet, 4, col, text, font=HEADER_FONT, align=CENTER)
    for rng in ("A3:A4", "B3:B4", "C3:C4", "D3:D4", "E3:F3", "G3:H3", "I3:J3", "K3:K4", "L3:L4"):
        sheet.merge_cells(rng)

    all_lines = [ln for sec in sections for ln in sec.lines]
    rates = sorted({ln.entry.fitting_rate for ln in all_lines if ln.entry.fitting_rate > 0})
    jobs: list[str] = []
    for ln in all_lines:
        for job, amount in ln.entry.labor_unit_by_job.items():
            if amount > 0 and job not in jobs:
                jobs.append(job)
    col_of: dict[Any, int] = {}
    next_col = HELPER_START_COL
    for rate in rates:
        col_of[("rate", rate)] = next_col
        next_col += 1
    col_of["sundry"] = next_col
    next_col += 1
    for job in jobs:
        col_of[("job", job)] = next_col
        next_col += 1
    helper_last = next_col - 1

    def letter(key: Any) -> str:
        return get_column_letter(col_of[key])

    def blank(r: int) -> None:
        for c in range(1, LAST_COL + 1):
            _set_cell(sheet, r, c, None)

    row = 5
    total = row
    for number, section in enumerate(sections, start=1):
        lines = section.lines
        title_row = row
        _set_cell(sheet, title_row, 1, section_label(number, section.title), font=HEADER_FONT, align=LEFT)
        for col in range(2, LAST_COL + 1):
            _set_cell(sheet, title_row, col, None)
        sheet.merge_cells(start_row=title_row, start_column=1, end_row=title_row, end_column=LAST_COL)

        first = title_row + 1
        row = first
        for ln in lines:
            e = ln.entry
            _set_cell(sheet, row, 1, e.name)
            _set_cell(sheet, row, 2, e.spec)
            _set_cell(sheet, row, 3, e.unit, align=CENTER)
            _set_cell(sheet, row, 4, ln.qty, align=RIGHT, number_format=qty_format(ln.qty))
            _set_cell(sheet, row, 5, f"={_ref(e.material_row, 'F')}", align=RIGHT, number_format=MONEY_FORMAT)
            _set_cell(sheet, row, 6, f"=D{row}*E{row}", align=RIGHT, number_format=MONEY_FORMAT)
            job_cols = [letter(("job", j)) for j in jobs if e.labor_unit_by_job.get(j, 0) > 0]
            if job_cols:
                _set_cell(sheet, row, 7, "=" + "+".join(f"{c}{row}" for c in job_cols), align=RIGHT, number_format=MONEY_FORMAT)
                _set_cell(sheet, row, 8, f"=D{row}*G{row}", align=RIGHT, number_format=MONEY_FORMAT)
            else:
                _set_cell(sheet, row, 7, None)
                _set_cell(sheet, row, 8, None)
            _set_cell(sheet, row, 9, None)
            _set_cell(sheet, row, 10, None)
            _set_cell(sheet, row, 11, f"=F{row}+H{row}+J{row}", align=RIGHT, number_format=MONEY_FORMAT)
            _set_cell(sheet, row, 12, ln.remark or None)
            for key, col in col_of.items():
                if key == "sundry":
                    value = 1 if e.has_sundry else None
                elif key[0] == "rate":
                    value = 1 if abs(e.fitting_rate - key[1]) < 1e-9 else None
                else:
                    refs = e.labor_rows.get(key[1], [])
                    value = "=" + "+".join(_ref(r, "H") for r in refs) if refs and e.labor_unit_by_job.get(key[1], 0) > 0 else None
                sheet.cell(row, col, value)
            row += 1
        last = max(row - 1, first)
        sec_rates = sorted({ln.entry.fitting_rate for ln in lines if ln.entry.fitting_rate > 0})
        sec_jobs = [j for j in jobs if any(ln.entry.labor_unit_by_job.get(j, 0) > 0 for ln in lines)]

        material_rows: list[int] = []
        labor_first = labor_last = None
        summary = row
        for rate in sec_rates:
            blank(summary)
            _set_cell(sheet, summary, 1, "[ 배관 부속재 ]")
            label = f"CD 전선관의 {round(rate * 100)} %" if abs(rate - CD_FITTING_RATE) < 1e-9 else f"전선관의 {round(rate * 100)} %"
            _set_cell(sheet, summary, 2, label)
            _set_cell(sheet, summary, 3, "식", align=CENTER)
            _set_cell(sheet, summary, 4, 1, align=RIGHT, number_format=QTY_INT_FORMAT)
            c = letter(("rate", rate))
            _set_cell(sheet, summary, 5, f"=SUMPRODUCT(F{first}:F{last},{c}{first}:{c}{last})", align=RIGHT, number_format=MONEY_FORMAT)
            _set_cell(sheet, summary, 6, f"=D{summary}*E{summary}*{rate}", align=RIGHT, number_format=MONEY_FORMAT)
            _set_cell(sheet, summary, 11, f"=F{summary}+H{summary}+J{summary}", align=RIGHT, number_format=MONEY_FORMAT)
            material_rows.append(summary)
            summary += 1
        if any(ln.entry.has_sundry for ln in lines):
            blank(summary)
            _set_cell(sheet, summary, 1, "[ 소모 잡자재 ]")
            _set_cell(sheet, summary, 2, f"전선, 전선관의 {round(SUNDRY_RATE * 100)} %")
            _set_cell(sheet, summary, 3, "식", align=CENTER)
            _set_cell(sheet, summary, 4, 1, align=RIGHT, number_format=QTY_INT_FORMAT)
            c = letter("sundry")
            _set_cell(sheet, summary, 5, f"=SUMPRODUCT(F{first}:F{last},{c}{first}:{c}{last})", align=RIGHT, number_format=MONEY_FORMAT)
            _set_cell(sheet, summary, 6, f"=D{summary}*E{summary}*{SUNDRY_RATE}", align=RIGHT, number_format=MONEY_FORMAT)
            _set_cell(sheet, summary, 11, f"=F{summary}+H{summary}+J{summary}", align=RIGHT, number_format=MONEY_FORMAT)
            material_rows.append(summary)
            summary += 1
        for job in sec_jobs:
            blank(summary)
            _set_cell(sheet, summary, 1, "노 무 비")
            _set_cell(sheet, summary, 2, job)
            _set_cell(sheet, summary, 3, "인", align=CENTER)
            c = letter(("job", job))
            _set_cell(sheet, summary, 8, f"=SUMPRODUCT(D{first}:D{last},{c}{first}:{c}{last})", align=RIGHT, number_format=MONEY_FORMAT)
            labor_first = labor_first or summary
            labor_last = summary
            summary += 1
        if sec_jobs:
            blank(summary)
            _set_cell(sheet, summary, 1, "[ 공 구 손 료 ]")
            _set_cell(sheet, summary, 2, f"노무비의 {round(TOOL_RATE * 100)} %")
            _set_cell(sheet, summary, 3, "식", align=CENTER)
            _set_cell(sheet, summary, 4, 1, align=RIGHT, number_format=QTY_INT_FORMAT)
            _set_cell(sheet, summary, 7, f"=SUM(H{labor_first}:H{labor_last})", align=RIGHT, number_format=MONEY_FORMAT)
            _set_cell(sheet, summary, 8, f"=D{summary}*G{summary}*{TOOL_RATE}", align=RIGHT, number_format=MONEY_FORMAT)
            tool_row = summary
            summary += 1
        else:
            tool_row = None
        for _ in range(BLANK_ROWS_BEFORE_TOTAL):
            blank(summary)
            summary += 1
        total = summary
        blank(total)
        _set_cell(sheet, total, 1, "( 합       계 )")
        parts = [f"SUM(F{first}:F{last})"] + [f"F{r}" for r in material_rows]
        _set_cell(sheet, total, 6, "=" + "+".join(parts), align=RIGHT, number_format=MONEY_FORMAT)
        if tool_row:
            _set_cell(sheet, total, 8, f"=SUM(H{labor_first}:H{tool_row})", align=RIGHT, number_format=MONEY_FORMAT)
        _set_cell(sheet, total, 11, f"=F{total}+H{total}+J{total}", align=RIGHT, number_format=MONEY_FORMAT)
        row = total + 1

    _apply_sheet_look(sheet, total, LAST_COL, row_height=FORM_ROW_HEIGHT)
    apply_form_widths(sheet, LAST_COL, remark_col=LAST_COL)
    for col in range(HELPER_START_COL, helper_last + 1):
        sheet.column_dimensions[get_column_letter(col)].hidden = True
    sheet.freeze_panes = "A5"
    sheet.page_setup.orientation = "landscape"
    sheet.page_setup.fitToWidth = 1
    sheet.page_setup.fitToHeight = 0
    sheet.sheet_properties.pageSetUpPr.fitToPage = True


def save_session(session: EstimateSession, sections: list[EstimateSection], dest_dir: Path | None) -> Path:
    """원본은 두고, 고른 공정마다 내역서 시트를 붙여 새 엑셀 파일로 저장한다."""
    workbook = session.workbook
    for old in session.estimate_sheets:  # 다시 저장할 때 내역서가 쌓이지 않게 먼저 지운다.
        if old in workbook.sheetnames:
            del workbook[old]
    session.estimate_sheets = []
    taken = set(workbook.sheetnames)
    names: list[str] = []
    if sections:  # 모든 공정을 시트 하나에 순서대로 쓴다.
        name = _sheet_name(ESTIMATE_SHEET_NAME, taken)
        names.append(name)
        write_estimate_sheet(workbook, name, list(sections))
    # 일위대가목록 바로 뒤에 내역서를 둔다.
    anchor = workbook.sheetnames.index(ILWIDAE_LIST_SHEET_NAME) if ILWIDAE_LIST_SHEET_NAME in workbook.sheetnames else len(workbook.sheetnames) - 1
    for offset, name in enumerate(names, start=1):
        current = workbook.sheetnames.index(name)
        workbook.move_sheet(name, offset=anchor + offset - current)
    session.estimate_sheets = list(names)
    _write_check_sheet(workbook, session.entries)
    dest = build_result_path(dest_dir, mode="estimate" if sections else "forward")
    if session.source_path is not None:
        assert_safe_save(session.source_path, dest)
    return save_workbook_safely(workbook, dest)
