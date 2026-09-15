from __future__ import annotations

from app.merge_parse import _representative_value, trim_grid


def test_representative_value_prefers_top_left() -> None:
    grid = [
        ["배관", None],
        [None, None],
    ]
    assert _representative_value(grid, 1, 1, 2, 1) == "배관"


def test_representative_value_falls_back_to_left_when_range_empty() -> None:
    grid = [
        ["내선전공", None, None],
    ]
    assert _representative_value(grid, 1, 2, 1, 3) == "내선전공"


def test_trim_grid_drops_empty_edges_only() -> None:
    rows = [
        [None, None, None],
        ["품명", None, "수량"],
        ["전선", None, 3],
        [None, None, None],
    ]
    trimmed = trim_grid(rows)
    assert trimmed[0] == ["품명", None, "수량"]
    assert trimmed[-1][0] == "전선"
    assert trimmed[0][1] is None
