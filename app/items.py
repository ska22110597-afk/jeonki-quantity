"""단가대비표·내역서에서 품목 행을 읽는다."""

from __future__ import annotations

from dataclasses import dataclass
from typing import Any

from app.estimate_parse import (
    EstimateSheet,
    find_column_index,
    find_header_row,
    find_quantity_column,
    first_data_row_number,
    header_row_span,
    is_section_row,
    lookup_key,
    normalize_header,
)


@dataclass
class LineItem:
    excel_row: int
    name: Any
    spec: Any
    unit: Any
    qty: Any
    material_price: Any
    section: bool = False

    @property
    def key(self) -> str:
        return lookup_key(self.name, self.spec)


def _pick(row: list[Any], index: int | None) -> Any:
    if index is None or index >= len(row):
        return None
    return row[index]


def _price_column(header: list[Any], subheader: list[Any] | None) -> int | None:
    idx = find_column_index(header, "단가")
    if idx is not None:
        return idx
    if subheader:
        idx = find_column_index(subheader, "단가")
        if idx is not None:
            return idx
    for i, cell in enumerate(header):
        token = normalize_header(cell)
        if "재료" in token and "금액" not in token:
            return i
    return 4 if len(header) > 4 else None


def parse_line_items(sheet: EstimateSheet) -> list[LineItem]:
    filled = sheet.filled
    if not filled:
        return []
    header_idx = find_header_row(filled)
    header = filled[header_idx]
    span = header_row_span(filled, header_idx)
    subheader = filled[header_idx + 1] if span == 2 and header_idx + 1 < len(filled) else None
    name_idx = find_column_index(header, "명칭")
    spec_idx = find_column_index(header, "규격")
    unit_idx = find_column_index(header, "단위")
    qty_idx = find_quantity_column(header)
    price_idx = _price_column(header, subheader)
    data_start = first_data_row_number(filled)
    items: list[LineItem] = []
    for excel_row in range(data_start, len(filled) + 1):
        row = filled[excel_row - 1]
        name = _pick(row, name_idx)
        spec = _pick(row, spec_idx)
        unit = _pick(row, unit_idx)
        if name is None and spec is None:
            continue
        section = is_section_row(name, spec, unit)
        items.append(
            LineItem(
                excel_row=excel_row,
                name=name,
                spec=spec,
                unit=unit,
                qty=None if section else _pick(row, qty_idx),
                material_price=None if section else _pick(row, price_idx),
                section=section,
            )
        )
    return items
