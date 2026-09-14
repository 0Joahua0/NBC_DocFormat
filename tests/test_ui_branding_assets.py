"""UI branding and feature-card asset safeguards."""

from pathlib import Path


PROJECT_ROOT = Path(__file__).parent.parent


def _read(relative_path):
    return (PROJECT_ROOT / relative_path).read_text(encoding="utf-8")


def test_application_ui_does_not_show_community_edition_notices():
    source = _read("NBC_DocFormat.py")

    forbidden = (
        "免费开源社区版",
        "此版本为",
        "非商业用途",
        "PolyForm NC",
        "查看使用许可证",
        "CommunityNoticeTicker",
        "CommunityEditionDialog",
        "should_show_community_notice",
        "community_notice",
    )
    for text in forbidden:
        assert text not in source

    assert "AboutDialog" in source


def test_application_ui_uses_nbc_branding_without_visible_version_or_author():
    source = _read("NBC_DocFormat.py")
    build_source = _read("build.py")
    desktop_entry = _read("packaging/appimage/NBC_DocFormat.desktop")

    assert 'self.root.title("NBC_DocFormat")' in source
    assert 'text="NBC_DocFormat"' in source
    assert 'APP_NAME = "NBC_DocFormat"' in build_source
    assert "Name=NBC_DocFormat" in desktop_entry

    for old_text in (
        "公文" + "格式处理工具",
        "KaguraNanaga",
        "v{__version__}",
        "版本:",
    ):
        assert old_text not in source

    assert "# {APP_NAME} v{VERSION}" not in build_source


def test_about_dialog_only_keeps_local_document_notice():
    source = _read("NBC_DocFormat.py")

    assert "所有文档处理均在本地完成，不上传、不收集任何数据。" in source
    assert "处理结果仅供参考，建议人工复核。" in source

    for removed_about_text in (
        "一键将 Word 文档排版为标准公文格式",
        "开发者：",
        "项目地址：",
    ):
        assert removed_about_text not in source


def test_feature_card_png_assets_exist_and_are_loaded_by_the_ui():
    source = _read("NBC_DocFormat.py")

    for asset_name in (
        "feature_smart.png",
        "feature_analyze.png",
        "feature_punctuation.png",
    ):
        asset = PROJECT_ROOT / "assets" / asset_name
        assert asset.exists()
        assert asset.stat().st_size > 0
        assert asset_name in source

    assert "tk.PhotoImage" in source
    assert "resource_path('assets', 'feature_smart.png')" in source
    assert "resource_path('assets', 'feature_analyze.png')" in source
    assert "resource_path('assets', 'feature_punctuation.png')" in source


def test_packaging_includes_feature_card_png_assets():
    build_source = _read("build.py")
    workflow = _read(".github/workflows/build.yml")

    for asset_name in (
        "feature_smart.png",
        "feature_analyze.png",
        "feature_punctuation.png",
    ):
        assert f"--add-data=assets/{asset_name};assets" in build_source
        assert build_source.count(f"--add-data=assets/{asset_name}:assets") == 2
        assert workflow.count(f'--add-data "assets/{asset_name}:assets"') == 2


def test_blue_theme_replaces_the_old_red_and_warm_paper_palette():
    source = _read("NBC_DocFormat.py")

    assert "PRIMARY = '#25B6DB'" in source
    assert "PRIMARY_HOVER = '#1399BF'" in source
    assert "BORDER_SELECTED = '#25B6DB'" in source

    for old_color in (
        "#BC4B26",
        "#A3421F",
        "#F9F0EC",
        "#E7A28A",
        "#F7F4EF",
        "#F2EFE9",
        "#FBF9F6",
    ):
        assert old_color not in source


def test_repository_does_not_ship_commercial_promotion_surfaces():
    assert not (PROJECT_ROOT / ("P" + "RO.md")).exists()
    assert not (PROJECT_ROOT / "assets" / ("xian" + "yu_qr.png")).exists()

    forbidden = (
        "P" + "ro 版",
        "P" + "RO.md",
        "闲" + "鱼",
        "goo" + "fish.com",
        "xian" + "yu_qr",
    )
    for relative_path in (
        "NBC_DocFormat.py",
        "README.md",
        "build.py",
        ".github/workflows/build.yml",
    ):
        source = _read(relative_path)
        for text in forbidden:
            assert text not in source
