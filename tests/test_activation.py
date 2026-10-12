from app import activation


def test_code_and_marker(tmp_path):
    marker = tmp_path / "a.dat"
    assert not activation.is_activated(marker)
    assert not activation.activate("000000", marker)
    assert not marker.exists()
    assert activation.activate(" 373302 ", marker)
    assert activation.is_activated(marker)
    marker.write_text("copied-from-other-pc")
    assert not activation.is_activated(marker)
