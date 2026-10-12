from app.fitting_rules import fitting_excluded, load_fitting_excludes, EXCLUDE_FILE_NAME
from app.ilwidae import is_conduit_name
from app.pumsam import spec_match_parts


def test_tray_width_height_both_orders():
    assert spec_match_parts("W300 x H100")[0] == 30000
    assert spec_match_parts("STRAIGHT, W300x100Hx1.2t")[0] == 30000
    assert spec_match_parts("STRAIGHT, W600x100Hx1.2t")[0] == 60000


def test_fittings_excluded_but_conduit_kept():
    assert not fitting_excluded("강제전선관", "아연도 16 mm")
    assert not fitting_excluded("경질비닐전선관", "HI 16 mm")
    assert fitting_excluded("강제전선관용 부품", "노말밴드, 아연도 36 mm")
    assert fitting_excluded("1종금속제가요전선관", "커넥터, 16 mm 일반-비방수")
    assert is_conduit_name("강제전선관용 부품")  # 잡재료비 대상은 그대로


def test_user_exclude_file(tmp_path):
    words = load_fitting_excludes(tmp_path)
    assert (tmp_path / EXCLUDE_FILE_NAME).exists()
    (tmp_path / EXCLUDE_FILE_NAME).write_text("# 설명\n파상형\n", encoding="utf-8")
    words = load_fitting_excludes(tmp_path)
    assert fitting_excluded("파상형 폴리에틸렌 전선관", "65㎜", words)
    assert not fitting_excluded("파상형 폴리에틸렌 전선관", "65㎜")


def test_fitting_rate_follows_pumsam_5_2():
    from app.ilwidae import conduit_fitting_rate

    assert conduit_fitting_rate("강제전선관", "아연도 16 mm") == 0.15
    assert conduit_fitting_rate("경질비닐전선관", "HI 16 mm") == 0.15
    assert conduit_fitting_rate("합성수지제 가요전선관", "CD-난연성 16㎜") == 0.40
    assert conduit_fitting_rate("CD관", "22㎜") == 0.40


def test_conduit_aliases_resolve_to_same_item():
    from app.ilwidae import is_conduit_name
    from app.pumsam_aliases import alias_names

    for name in ("ELP전선관", "ELP관", "지중전선관", "파상형 폴리에틸렌 전선관"):
        assert is_conduit_name(name), name
        assert "합성수지 파형관 설치" in alias_names(name), name
    assert "강제전선관" in alias_names("RMC")
    assert "경질비닐전선관" in alias_names("VE전선관")
