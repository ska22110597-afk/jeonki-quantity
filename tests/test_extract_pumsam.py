from __future__ import annotations

from scripts.extract_pumsam_book import dedupe_book_rows


def test_dedupe_keeps_different_specs_with_same_qty() -> None:
    rows = [
        {
            "명칭": "관로 청소",
            "규격": "150 이하",
            "노무명칭": "보통인부",
            "품셈": 9.7,
            "품셈근거": "전기2-11-1",
        },
        {
            "명칭": "관로 청소",
            "규격": "300 이하",
            "노무명칭": "보통인부",
            "품셈": 9.7,
            "품셈근거": "전기2-11-1",
        },
        {
            "명칭": "관로 청소",
            "규격": "150 이하 9.70",
            "노무명칭": "보통인부",
            "품셈": 9.7,
            "품셈근거": "전기2-11-1",
        },
        {
            "명칭": "관로 청소",
            "규격": "150 이하",
            "노무명칭": "특별인부",
            "품셈": None,
            "품셈근거": "전기2-11-1",
        },
    ]
    kept = dedupe_book_rows(rows)
    specs = {(row["규격"], row["노무명칭"]) for row in kept}
    assert ("150 이하", "보통인부") in specs
    assert ("300 이하", "보통인부") in specs
    assert ("150 이하 9.70", "보통인부") not in specs
    assert ("150 이하", "특별인부") in specs
