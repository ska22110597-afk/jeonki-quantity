from __future__ import annotations

from scripts.extract_pumsam_book import (
    dedupe_book_rows,
    merge_subscript_pair,
    parse_item_line,
    parse_labor_table,
    parse_matrix_table,
    parse_paired_size_table,
    parse_table,
    title_is_bad,
)


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


def test_parse_horizontal_conductor_is_one_item() -> None:
    rows: list[dict] = []
    added = parse_labor_table(
        [["공종", "내선전공", "보통인부"], ["수평도체 설치", "0.017", "0.008"]],
        "5-42-2",
        "수평도체 설치",
        "m",
        rows,
        "전기",
        "전기",
    )
    assert added == 2
    names = {row["명칭"] for row in rows}
    specs = {row["규격"] for row in rows}
    jobs = {row["노무명칭"]: row["품셈"] for row in rows}
    assert names == {"수평도체 설치"}
    assert specs == {""}
    assert jobs == {"내선전공": 0.017, "보통인부": 0.008}


def test_parse_generator_keeps_kva_and_splits_commissioning() -> None:
    table = [
        ["발전기\n용량", "설치", None, None, None, "시운전 및 조정", None],
        [None, "전기공사\n기사", "플랜트\n전공", "기계\n설비공", "특별인부", "전기공사\n기사", "플랜트\n전공"],
        [
            "20 kVA\n50 〃",
            "10.5\n15.8",
            "6.3\n8.4",
            "6.3\n8.4",
            "5.3\n6.3",
            "3.2\n3.2",
            "3.2\n4.2",
        ],
    ]
    rows: list[dict] = []
    added = parse_labor_table(table, "5-43", "자가발전기 설치", "대", rows, "전기", "전기")
    assert added == 12
    install = [row for row in rows if row["명칭"] == "자가발전기 설치"]
    commission = [row for row in rows if "시운전" in row["명칭"]]
    install_20 = {row["노무명칭"]: row["품셈"] for row in install if "20" in row["규격"]}
    assert "kVA" in next(row["규격"] for row in install if "20" in row["규격"])
    assert install_20["전기공사기사"] == 10.5
    assert install_20["플랜트전공"] == 6.3
    assert install_20["기계설비공"] == 6.3
    assert install_20["특별인부"] == 5.3
    assert not any(row["규격"] == "6.3" for row in rows)
    comm_20 = {row["노무명칭"]: row["품셈"] for row in commission if "20" in row["규격"]}
    assert comm_20["전기공사기사"] == 3.2
    assert comm_20["플랜트전공"] == 3.2
    spec50 = next(row["규격"] for row in install if row["품셈"] == 15.8)
    assert "50" in spec50
    assert "kVA" in spec50


def test_parse_wall_hole_matrix_keeps_diameter_and_thickness() -> None:
    table = [
        ["구경(㎜)", "콘크리트 두께(㎜)", None, None, None],
        [None, "150 ㎜ 이하", "200 ㎜ 이하", "300 ㎜ 이하", "400 ㎜ 이하"],
        [
            "50\n75",
            "0.13\n0.15",
            "0.21\n0.23",
            "0.42\n0.46",
            "0.52\n0.59",
        ],
    ]
    rows: list[dict] = []
    added = parse_matrix_table(table, "5-29-2", "벽관통 구멍뚫기", "개소", rows, "특별인부", "전기")
    assert added == 8
    specs = {row["규격"] for row in rows}
    assert any("50" in spec and "mm" in spec and "150" in spec for spec in specs)
    assert any("75" in spec and "400" in spec for spec in specs)
    jobs = {row["노무명칭"] for row in rows}
    assert jobs == {"특별인부"}
    qty_50_150 = next(row["품셈"] for row in rows if "50" in row["규격"] and "150" in row["규격"])
    assert qty_50_150 == 0.13


def test_parse_motor_control_uses_kw_and_starting_method() -> None:
    table = [
        ["전동기 용량 (kW)", "직입기동", "Y-△ 기동"],
        ["0.2 ~ 2.2\n3.7\n11", "1.85\n2.05\n2.95", "-\n-\n3.04"],
    ]
    rows: list[dict] = []
    added = parse_labor_table(
        table, "5-39", "전동기 제어반 설치", "대", rows, "전기", "전기", apply_job="플랜트전공"
    )
    assert added == 4
    specs = {row["규격"] for row in rows}
    assert any("kW" in spec and "3.7" in spec and "직입" in spec for spec in specs)
    assert any("11" in spec and "Y-" in spec for spec in specs)
    assert not any(row["규격"] in {"3.7", "2.05", "6.3"} for row in rows)
    three = {row["규격"]: row["품셈"] for row in rows if "3.7" in row["규격"]}
    assert list(three.values()) == [2.05]
    eleven_y = next(row["품셈"] for row in rows if "11" in row["규격"] and "Y-" in row["규격"])
    assert eleven_y == 3.04


def test_parse_paired_motor_keeps_right_column_kva_size() -> None:
    table = [
        ["전동기용량", "플랜트전공", "전동기용량", "플랜트전공"],
        [
            "0.75 kW 이하\n1.5 〃",
            "0.44\n0.55",
            "110 kW 이하\n125 〃",
            "5.70\n6.15",
        ],
    ]
    rows: list[dict] = []
    added = parse_paired_size_table(table, "5-36", "전동기 설치", "대", rows, "전기", "전기")
    assert added == 4
    specs = {row["규격"] for row in rows}
    assert any("0.75" in spec and "kW" in spec and "이하" in spec for spec in specs)
    assert any("1.5" in spec and "kW" in spec and spec.index("kW") < spec.index("이하") for spec in specs)
    assert any("110" in spec and "kW" in spec for spec in specs)
    assert any("125" in spec and "kW" in spec for spec in specs)
    qty_110 = next(row["품셈"] for row in rows if "110" in row["규격"])
    assert qty_110 == 5.7
    assert not any(row["규격"] in {"0.44", "5.70", "6.15"} for row in rows)
    assert {row["명칭"] for row in rows} == {"전동기 설치"}


def test_assembly_title_is_not_treated_as_truncated() -> None:
    assert title_is_bad("고정 빔 조립") is False
    assert title_is_bad("22.9 kV GIS 변압기 2차(Main) 베이(") is True


def test_parse_fixed_beam_splits_assembly_and_keeps_kind() -> None:
    table = [
        ["종별", "단위", "조립", None, "설치", None, "장비사용 시간(분)"],
        [None, None, "배전전공", "보통인부", "배전전공", "보통인부", None],
        [
            "4각형트라스 1선용\n〃 2 〃",
            "본\n〃",
            "1.12\n1.86",
            "0.56\n0.93",
            "2.20\n3.49",
            "0.90\n1.63",
            "78\n93",
        ],
    ]
    rows: list[dict] = []
    added = parse_labor_table(table, "7-24", "고정 빔 설치", "본", rows, "전기", "전기")
    assert added >= 8
    names = {row["명칭"] for row in rows}
    assert "고정 빔 조립" in names
    assert "고정 빔 설치" in names
    specs = {row["규격"] for row in rows}
    assert any("4각형트라스" in spec and "2" in spec for spec in specs)
    assert "2" not in specs
    assemble = next(
        row
        for row in rows
        if row["명칭"] == "고정 빔 조립"
        and "1선용" in row["규격"].replace(" ", "")
        and row["노무명칭"] == "배전전공"
    )
    assert assemble["품셈"] == 1.12
    install = next(
        row
        for row in rows
        if row["명칭"] == "고정 빔 설치"
        and "1선용" in row["규격"].replace(" ", "")
        and row["노무명칭"] == "배전전공"
    )
    assert install["품셈"] == 2.2


def test_merge_subscript_and_skip_toc() -> None:
    assert merge_subscript_pair("B (1.415 ㎡)", "0") == "B0 (1.415 ㎡)"
    assert merge_subscript_pair("SF 가스 처리", "6") == "SF6 가스 처리"
    assert parse_item_line("13-10-1 무정전 전원장치(UPS, CVCF) 점검 503", "통신") is None
    assert parse_item_line("7-12-5 광 송·수신기 등", "통신") == ("7-12-5", "광 송·수신기 등")
    assert parse_item_line("8-1 활주로 등화시설 등기구 설치", "전기") == (
        "8-1",
        "활주로 등화시설 등기구 설치",
    )


def test_parse_table_matrix_prefers_apply_job() -> None:
    table = [
        ["구경(㎜)", "두께(㎜)"],
        [None, "250㎜ 이하"],
        ["100", "0.40"],
    ]
    rows: list[dict] = []
    added = parse_table(
        table, "3-7-2", "벽 관통 구멍뚫기", "개소", rows, "통신", "통신", set(), apply_job="특별인부"
    )
    assert added == 1
    assert "100" in rows[0]["규격"]
    assert "mm" in rows[0]["규격"]
    assert rows[0]["품셈"] == 0.40
    assert rows[0]["노무명칭"] == "특별인부"
