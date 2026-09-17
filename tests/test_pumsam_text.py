from __future__ import annotations

from app.pumsam_text import clean_spec_unit, expand_ditto, strip_trailing_unit_dash


def test_expand_ditto_copies_prefix_and_선용() -> None:
    assert expand_ditto("〃 2 〃", "4각형트라스 1선용") == "4각형트라스 2 선용"
    assert expand_ditto("평면트라스 2 〃", "평면트라스 1 선용") == "평면트라스 2 선용"


def test_expand_ditto_splitter_keeps_분배기_words() -> None:
    assert expand_ditto("3 〃 (2 〃 )", "2분배기(1분기기)").replace(" ", "") == "3분배기(2분기기)"


def test_expand_ditto_keeps_kva_and_이하() -> None:
    assert "kVA" in expand_ditto("50 〃", "20 kVA")
    assert expand_ditto("10 m 〃", "일반용 8 m 이하") == "10 m 이하"
    assert expand_ditto("1.5 〃", "0.75 kW 이하") == "1.5 kW 이하"


def test_expand_ditto_분배기_uses_first_kind_not_paren() -> None:
    six = expand_ditto("6 〃", "5 분배기 (4 분기기 )")
    assert "분배기" in six
    assert "분기기" not in six
    three = expand_ditto("3 〃 (2 〃 )", "2분배기(1분기기)").replace(" ", "")
    assert "3분배기" in three
    assert three.count("분기기") == 1


def test_size_in_meters_stays_in_spec_not_unit() -> None:
    spec, unit = strip_trailing_unit_dash("10 m", "본")
    assert spec == "10 m"
    assert unit == "본"
    spec, unit = clean_spec_unit("10 m", "본")
    assert spec == "10 m"
    assert unit == "본"
    spec, unit = clean_spec_unit("중하중용 14 m", "본")
    assert spec == "중하중용 14 m"
    assert unit == "본"
    spec, unit = clean_spec_unit("Ø 30mm 이 하", "본")
    assert spec == "Ø 30mm 이하"
    assert unit == "본"
    spec, unit = clean_spec_unit("154 kV OF 케이블 1200 OF ㎟ 이하", "km")
    assert spec == "154 kV OF 케이블 1200 ㎟ 이하"


def test_process_name_trailing_m_still_moves_to_unit() -> None:
    spec, unit = strip_trailing_unit_dash("급유관 설치 m", "식")
    assert spec == "급유관 설치"
    assert unit == "m"
