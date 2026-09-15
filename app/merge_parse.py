"""원본 단가대비표의 병합 셀을 메모리에서만 풀어 대표 값을 채운다.

원본 파일은 이진 읽기만 하며 save/unmerge 를 디스크에 쓰지 않는다.
"""

from __future__ import annotations

import math
from typing import Any, Iterable

import pandas as pd
from openpyxl.worksheet.worksheet import Worksheet

SheetRows = list[list[Any]]


def _has_value(value: Any) -> bool:
    if value is None:
        return False
    try:
        if pd.isna(value):
            return False
    except (TypeError, ValueError):
        pass
    if isinstance(value, float) and math.isnan(value):
        return False
    if isinstance(value, str) and not value.strip():
        return False
    return True


def _cell_value(value: Any) -> Any:
    if value is None:
        return None
    if isinstance(value, str):
        stripped = value.strip()
        return stripped if stripped else None
    return value


def _iter_merged_ranges(sheet: Worksheet) -> Iterable[Any]:
    try:
        ranges = getattr(sheet, "merged_cells", None)
        if ranges is None:
            return []
        return list(ranges.ranges)
    except Exception:
        return []


def _grid_from_sheet(sheet: Worksheet) -> SheetRows:
    max_row = sheet.max_row or 0
    max_col = sheet.max_column or 0
    if max_row <= 0 or max_col <= 0:
        return []
    grid: SheetRows = [[None] * max_col for _ in range(max_row)]
    for row in sheet.iter_rows(min_row=1, max_row=max_row, min_col=1, max_col=max_col):
        for cell in row:
            grid[cell.row - 1][cell.column - 1] = _cell_value(cell.value)
    return grid


def _representative_value(grid: SheetRows, min_row: int, min_col: int, max_row: int, max_col: int) -> Any:
    """병합 범위의 상단·좌측 대표 값을 고른다. 비어 있으면 범위 안·위·왼쪽을 본다."""
    rows = len(grid)
    cols = len(grid[0]) if grid else 0

    def at(r: int, c: int) -> Any:
        if 1 <= r <= rows and 1 <= c <= cols:
            return grid[r - 1][c - 1]
        return None

    top_left = at(min_row, min_col)
    if _has_value(top_left):
        return top_left

    for r in range(min_row, max_row + 1):
        for c in range(min_col, max_col + 1):
            value = at(r, c)
            if _has_value(value):
                return value

    above = at(min_row - 1, min_col)
    if _has_value(above):
        return above

    left = at(min_row, min_col - 1)
    if _has_value(left):
        return left

    return top_left


def fill_merged_values(sheet: Worksheet) -> SheetRows:
    """병합 셀을 해제하지 않고, 값 격자만 상단/좌측 대표 값으로 채운다."""
    grid = _grid_from_sheet(sheet)
    if not grid:
        return []

    for merged in _iter_merged_ranges(sheet):
        try:
            min_row, min_col = int(merged.min_row), int(merged.min_col)
            max_row, max_col = int(merged.max_row), int(merged.max_col)
        except (AttributeError, TypeError, ValueError):
            continue
        if min_row < 1 or min_col < 1:
            continue
        fill_value = _representative_value(grid, min_row, min_col, max_row, max_col)
        for r in range(min_row, max_row + 1):
            if r < 1 or r > len(grid):
                continue
            row = grid[r - 1]
            for c in range(min_col, max_col + 1):
                if c < 1 or c > len(row):
                    continue
                if not _has_value(row[c - 1]):
                    row[c - 1] = fill_value
    return grid


def trim_grid(rows: SheetRows) -> SheetRows:
    """완전히 빈 앞뒤 행과 끝 빈 열만 제거한다. 가운데 빈 열은 유지한다."""
    if not rows:
        return []
    frame = pd.DataFrame(rows, dtype=object)
    frame = frame.dropna(how="all")
    if frame.empty:
        return []
    frame = frame.where(frame.notna(), None)
    values = frame.values.tolist()

    while values and not any(_has_value(v) for v in values[0]):
        values.pop(0)
    while values and not any(_has_value(v) for v in values[-1]):
        values.pop()
    if not values:
        return []

    max_col = 0
    for row in values:
        for idx, value in enumerate(row, start=1):
            if _has_value(value):
                max_col = max(max_col, idx)
    return [list(row[:max_col]) for row in values]
