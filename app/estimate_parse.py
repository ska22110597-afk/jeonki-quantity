"""내역서 엑셀 읽기. 원본 파일은 저장하지 않는다."""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from app.merge_parse import SheetRows, fill_merged_values, trim_grid
from app.paths import is_allowed_excel

HEADER_ALIASES = {
    "명칭": ("명칭", "품명", "품목", "자재명", "항목"),
    "규격": ("규격", "사양", "규격/사양"),
    "단위": ("단위", "단위명"),
    "수량": ("수량", "설계수량", "물량", "결정수량", "계약수량", "설계물량"),
    "단가": ("단가", "재료비단가", "재료비", "가격"),
}

HEADER_HINTS = ("명칭", "품명", "규격", "단위", "수량", "단가")
SECTION_NAME = re.compile(r"^\s*\d+\s*[\.．]")
SUBHEADER_TOKENS = {
    "단가",
    "금액",
    "할증",
    "할증%",
    "품셈",
    "공량",
    "산출수량",
    "결정수량",
    "물가정보",
    "적용단가",
    "조달청",
    "조달청가격",
    "조사가격",
    "조사가격1",
    "조사가격2",
    "조사가격3",
    "거래가격",
    "유통물가",
    "page",
    "PAGE",
}
HEADER_ITEM_NAMES = {"품명", "명칭", "품목", "코드", "자재명", "항목", "품목명"}
MergeRange = tuple[int, int, int, int]


@dataclass
class EstimateSheet:
    """원본 행 번호를 유지한 내역서·일위대가목록. 병합은 메모리에서만 채운다."""

    filled: SheetRows
    raw: SheetRows
    merges: list[MergeRange] = field(default_factory=list)
    title: str = ""

    @property
    def max_row(self) -> int:
        return len(self.filled)

    @property
    def max_col(self) -> int:
        if not self.filled:
            return 0
        return max(len(row) for row in self.filled)


def normalize_header(value: Any) -> str:
    if value is None:
        return ""
    return str(value).replace(" ", "").replace("\n", "").strip()


def find_column_index(header_row: list[Any], field: str) -> int | None:
    aliases = HEADER_ALIASES.get(field, (field,))
    normalized_aliases = {normalize_header(alias) for alias in aliases}
    for idx, cell in enumerate(header_row):
        if normalize_header(cell) in normalized_aliases:
            return idx
    return None


def find_quantity_column(header_row: list[Any]) -> int | None:
    idx = find_column_index(header_row, "수량")
    if idx is not None:
        return idx
    for i, cell in enumerate(header_row):
        name = normalize_header(cell)
        if "수량" in name and "단가" not in name and "금액" not in name:
            return i
    return None


def find_header_row(rows: SheetRows) -> int:
    for i, row in enumerate(rows):
        normalized = [normalize_header(c) for c in row]
        hits = sum(1 for hint in HEADER_HINTS if any(hint in cell for cell in normalized))
        if hits >= 2:
            return i
    return 0


def header_row_span(rows: SheetRows, header_idx: int) -> int:
    """2단 헤더(명칭 행 + 단가/금액·품셈 행)이면 2, 아니면 1."""
    if header_idx + 1 >= len(rows):
        return 1
    next_tokens = [normalize_header(c) for c in rows[header_idx + 1]]
    if any(token in SUBHEADER_TOKENS or token.startswith("조사가격") for token in next_tokens):
        return 2
    return 1


def is_header_item(name: Any, spec: Any, unit: Any) -> bool:
    """헤더 글자(품명·규격·단위)가 데이터 행으로 섞인 경우."""
    name_token = normalize_header(name)
    if name_token in HEADER_ITEM_NAMES:
        return True
    if name_token == "코드":
        return True
    spec_token = normalize_header(spec)
    unit_token = normalize_header(unit)
    if spec_token == "규격" and unit_token in {"단위", "단위명", ""}:
        return True
    return False


def first_data_row_number(rows: SheetRows) -> int:
    """데이터가 시작되는 엑셀 행 번호(1부터)."""
    header_idx = find_header_row(rows)
    return header_idx + header_row_span(rows, header_idx) + 1


def is_sundry_form_row(name: Any) -> bool:
    """내역서 아래 부속재·잡자재·노무비·공구손료·합계 양식 행."""
    token = normalize_header(name)
    if not token:
        return False
    if "합계" in token and (token.startswith("(") or token.startswith("[")):
        return True
    return any(key in token for key in ("배관부속재", "소모잡자재", "공구손료", "노무비"))


def is_section_row(name: Any, spec: Any, unit: Any) -> bool:
    """'1. 옥외전기공사' 같은 공종 제목 행."""
    name_text = str(name).strip() if name is not None else ""
    if not name_text:
        return True
    if spec is None and unit is None and SECTION_NAME.match(name_text):
        return True
    if "호표" in name_text:
        return True
    if spec is None and unit is None and name_text.endswith("공사") and "_" not in name_text:
        return True
    return False


def lookup_key(name: Any, spec: Any) -> str:
    """품셈표 검색용. 공백을 없애 명칭+규격을 붙인다."""
    left = "" if name is None else str(name)
    right = "" if spec is None else str(spec)
    return "".join(ch for ch in f"{left}{right}" if not ch.isspace())


def concat_key(name: Any, spec: Any) -> str:
    """엑셀 CONCATENATE(B,C) 와 같은 키. 가운데 공백은 유지한다."""
    left = "" if name is None else str(name)
    right = "" if spec is None else str(spec)
    return f"{left}{right}"


def collect_merge_ranges(sheet: Any) -> list[MergeRange]:
    ranges: list[MergeRange] = []
    try:
        merged = getattr(sheet, "merged_cells", None)
        if merged is None:
            return ranges
        for item in list(merged.ranges):
            ranges.append((int(item.min_row), int(item.min_col), int(item.max_row), int(item.max_col)))
    except Exception:
        return ranges
    return ranges


def is_merge_top_left(row: int, col: int, merges: list[MergeRange]) -> bool:
    for min_row, min_col, max_row, max_col in merges:
        if min_row <= row <= max_row and min_col <= col <= max_col:
            return row == min_row and col == min_col
    return True


def _raw_grid_from_sheet(sheet: Any) -> SheetRows:
    max_row = sheet.max_row or 0
    max_col = sheet.max_column or 0
    if max_row <= 0 or max_col <= 0:
        return []
    grid: SheetRows = [[None] * max_col for _ in range(max_row)]
    for row in sheet.iter_rows(min_row=1, max_row=max_row, min_col=1, max_col=max_col):
        for cell in row:
            grid[cell.row - 1][cell.column - 1] = cell.value
    return grid


_SKIP_QTY_SHEETS = {"품셈표", "노임단가", "공량산출서"}


def list_sheet_titles(source_path: Path) -> list[str]:
    workbook = _open_workbook(source_path)
    try:
        return [str(sheet.title) for sheet in workbook.worksheets]
    finally:
        workbook.close()


def _pick_estimate_sheet(workbook: Any, preferred_title: str | None = None) -> Any:
    if not workbook.worksheets:
        raise ValueError("엑셀에 시트가 없습니다.")
    if preferred_title:
        for candidate in workbook.worksheets:
            if str(candidate.title) == preferred_title:
                return candidate
    by_title = {str(sheet.title): sheet for sheet in workbook.worksheets}
    for name in ("내역서", "일위대가목록"):
        if name in by_title:
            return by_title[name]
    for candidate in workbook.worksheets:
        if "내역" in str(candidate.title):
            return candidate
    for candidate in workbook.worksheets:
        title = str(candidate.title)
        if title in _SKIP_QTY_SHEETS:
            continue
        if title in {"단가대비표", "일위대가"}:
            continue
        return candidate
    return workbook.worksheets[0]


def _trim_trailing(grid: SheetRows) -> SheetRows:
    """앞 빈 행은 남겨 행 번호를 유지하고, 뒤 빈 행·열만 자른다."""
    if not grid:
        return []
    last_row = -1
    last_col = 0
    for r_idx, row in enumerate(grid):
        for c_idx, value in enumerate(row, start=1):
            if value is None:
                continue
            if isinstance(value, str) and not str(value).strip():
                continue
            last_row = r_idx
            last_col = max(last_col, c_idx)
    if last_row < 0:
        return []
    clipped = [list(row[:last_col]) for row in grid[: last_row + 1]]
    return clipped


def is_external_formula(value: Any) -> bool:
    """다른 통합문서(`[1]파일명`)를 가리키는 수식. 결과 파일에 그대로 두면 엑셀이 복구 창을 띄운다."""
    if not isinstance(value, str) or not value.startswith("="):
        return False
    return "[" in value


def sanitize_copied_value(value: Any, cached: Any = None) -> Any:
    if is_external_formula(value):
        return cached
    return value


def _open_workbook(source_path: Path, *, data_only: bool = False):
    from openpyxl import load_workbook

    path = Path(source_path)
    if not path.exists():
        raise FileNotFoundError(f"원본 파일을 찾을 수 없습니다: {path}")
    if not is_allowed_excel(path):
        raise ValueError("xlsx 또는 xlsm 파일만 읽을 수 있습니다.")
    with path.open("rb") as handle:
        return load_workbook(filename=handle, data_only=data_only, keep_vba=False)


def _apply_cached_over_external(formula_grid: SheetRows, cached_grid: SheetRows) -> SheetRows:
    height = len(formula_grid)
    width = max((len(row) for row in formula_grid), default=0)
    out: SheetRows = []
    for r_idx in range(height):
        formula_row = formula_grid[r_idx] if r_idx < len(formula_grid) else []
        cached_row = cached_grid[r_idx] if r_idx < len(cached_grid) else []
        row: list[Any] = []
        for c_idx in range(width):
            value = formula_row[c_idx] if c_idx < len(formula_row) else None
            cached = cached_row[c_idx] if c_idx < len(cached_row) else None
            row.append(sanitize_copied_value(value, cached))
        out.append(row)
    return out


def load_named_sheet(source_path: Path, title: str) -> EstimateSheet | None:
    titles = list_sheet_titles(source_path)
    if title not in titles:
        return None
    return load_estimate_sheet(source_path, preferred_title=title)


def load_estimate_sheet(source_path: Path, preferred_title: str | None = None) -> EstimateSheet:
    """병합을 채운 전체 격자. 제목 행을 버리지 않아 원본 행 번호를 유지한다."""
    workbook = _open_workbook(source_path)
    try:
        sheet = _pick_estimate_sheet(workbook, preferred_title)
        sheet_title = str(sheet.title)
        merges = collect_merge_ranges(sheet)
        raw = _trim_trailing(_raw_grid_from_sheet(sheet))
        filled = _trim_trailing(fill_merged_values(sheet))
    finally:
        workbook.close()
    cached_wb = _open_workbook(source_path, data_only=True)
    try:
        cached_sheet = None
        for candidate in cached_wb.worksheets:
            if str(candidate.title) == sheet_title:
                cached_sheet = candidate
                break
        cached_grid = _trim_trailing(_raw_grid_from_sheet(cached_sheet)) if cached_sheet is not None else []
    finally:
        cached_wb.close()
    raw = _apply_cached_over_external(raw, cached_grid)
    filled = _apply_cached_over_external(filled, cached_grid)
    if not filled:
        raise ValueError("내역서에 읽을 수 있는 데이터가 없습니다.")
    width = max(len(row) for row in filled)
    raw_width = max((len(row) for row in raw), default=width)
    width = max(width, raw_width)
    filled = [list(row) + [None] * (width - len(row)) for row in filled]
    raw = [list(row) + [None] * (width - len(row)) for row in raw]
    while len(raw) < len(filled):
        raw.append([None] * width)
    while len(filled) < len(raw):
        filled.append([None] * width)
    return EstimateSheet(filled=filled, raw=raw, merges=merges, title=sheet_title)


def read_full_grid(source_path: Path) -> SheetRows:
    return load_estimate_sheet(source_path).filled


def read_workbook_first_sheet(source_path: Path) -> SheetRows:
    """헤더부터의 표. 테스트·품셈 추출용. 원본은 저장하지 않는다."""
    filled = read_full_grid(source_path)
    cleaned = trim_grid(filled)
    if not cleaned:
        raise ValueError("내역서에 읽을 수 있는 데이터가 없습니다.")
    header_idx = find_header_row(cleaned)
    table = trim_grid(cleaned[header_idx:])
    if not table:
        raise ValueError("내역서 헤더(명칭/규격/단위/수량)를 찾지 못했습니다.")
    return table


def read_named_sheet_rows(source_path: Path, name_contains: str) -> SheetRows | None:
    """통합 엑셀에 품셈표 시트가 있으면 그 표를 반환한다."""
    from openpyxl import load_workbook

    path = Path(source_path)
    with path.open("rb") as handle:
        workbook = load_workbook(filename=handle, data_only=True, keep_vba=False)
        try:
            sheet = None
            for candidate in workbook.worksheets:
                if name_contains in str(candidate.title):
                    sheet = candidate
                    break
            if sheet is None:
                return None
            grid = fill_merged_values(sheet)
        finally:
            workbook.close()
    cleaned = trim_grid(grid)
    if not cleaned:
        return None
    header_idx = find_header_row(cleaned)
    return trim_grid(cleaned[header_idx:])


# 이전 이름 호환
read_unit_price_table = read_workbook_first_sheet
