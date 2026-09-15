"""단가대비표·내역서에서 품목 행을 읽는다."""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any

from app.estimate_parse import (
    EstimateSheet,
    find_column_index,
    find_header_row,
    find_quantity_column,
    first_data_row_number,
    header_row_span,
    is_header_item,
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
    price_col: int | None = None
    qty_col: int | None = None
    fields: dict[str, Any] = field(default_factory=dict)

    @property
    def key(self) -> str:
        return lookup_key(self.name, self.spec)


def _pick(row: list[Any], index: int | None) -> Any:
    if index is None or index >= len(row):
        return None
    return row[index]


def _combined_labels(header: list[Any], subheader: list[Any] | None) -> list[str]:
    width = max(len(header), len(subheader or []))
    labels: list[str] = []
    for i in range(width):
        top = normalize_header(header[i] if i < len(header) else None)
        bottom = normalize_header(subheader[i] if subheader and i < len(subheader) else None)
        labels.append(bottom or top)
    return labels


def _numeric_count(filled: list[list[Any]], col: int, data_start: int) -> int:
    count = 0
    for row in filled[data_start - 1 :]:
        value = row[col] if col < len(row) else None
        if isinstance(value, (int, float)):
            count += 1
    return count


def find_material_price_column(
    header: list[Any],
    subheader: list[Any] | None,
    filled: list[list[Any]],
    data_start: int,
) -> int | None:
    """재료비 단가 열. 적용단가 → 물가정보 → 재료비 순이되, 숫자가 있는 열을 고른다."""
    labels = _combined_labels(header, subheader)
    ranked = (
        "적용단가",
        "물가정보",
        "조달청",
        "조사가격",
        "재료비단가",
        "재료비",
        "단가",
    )
    candidates: list[tuple[int, int, int]] = []
    for index, label in enumerate(labels):
        if "노무" in label or "경비" in label:
            continue
        for rank, token in enumerate(ranked):
            if token in label:
                filled_n = _numeric_count(filled, index, data_start)
                candidates.append((0 if filled_n else 1, rank, index))
                break
    if candidates:
        candidates.sort()
        return candidates[0][2]
    idx = find_column_index(header, "단가")
    if idx is not None:
        return idx
    if subheader:
        idx = find_column_index(subheader, "단가")
        if idx is not None:
            return idx
    return 4 if len(header) > 4 else None


_SKIP_FIELD_TOKENS = {
    "코드",
    "품명",
    "명칭",
    "규격",
    "단위",
    "번호",
    "비고",
    "수량",
    "품목",
}


def extract_price_fields(row: list[Any], labels: list[str]) -> dict[str, Any]:
    """물가정보·PAGE·조달청처럼 짝이 되는 단가 열을 이름 그대로 담는다."""
    result: dict[str, Any] = {}
    last_key: str | None = None
    for index, label in enumerate(labels):
        token = normalize_header(label)
        if not token or token in _SKIP_FIELD_TOKENS:
            continue
        if token == "PAGE":
            if last_key:
                result[f"{last_key}_PAGE"] = _pick(row, index)
            continue
        result[token] = _pick(row, index)
        last_key = token
    return result


def first_filled(*values: Any) -> Any:
    for value in values:
        if value not in (None, ""):
            return value
    return None


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
    data_start = first_data_row_number(filled)
    price_idx = find_material_price_column(header, subheader, filled, data_start)
    labels = _combined_labels(header, subheader)
    items: list[LineItem] = []
    for excel_row in range(data_start, len(filled) + 1):
        row = filled[excel_row - 1]
        name = _pick(row, name_idx)
        spec = _pick(row, spec_idx)
        unit = _pick(row, unit_idx)
        if name is None and spec is None:
            continue
        if is_header_item(name, spec, unit):
            continue
        section = is_section_row(name, spec, unit)
        qty = None if section or qty_idx is None else _pick(row, qty_idx)
        fields = {} if section else extract_price_fields(row, labels)
        items.append(
            LineItem(
                excel_row=excel_row,
                name=name,
                spec=spec,
                unit=unit,
                qty=qty,
                material_price=None if section else first_filled(
                    fields.get("적용단가"),
                    fields.get("물가정보"),
                    fields.get("조달청"),
                    fields.get("조사가격2"),
                    fields.get("조사가격3"),
                    _pick(row, price_idx),
                ),
                section=section,
                price_col=price_idx,
                qty_col=qty_idx,
                fields=fields,
            )
        )
    return items
