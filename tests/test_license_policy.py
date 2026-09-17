"""License migration safeguards for v1.0.0 and later releases."""

from pathlib import Path


PROJECT_ROOT = Path(__file__).parent.parent


def _read(relative_path):
    return (PROJECT_ROOT / relative_path).read_text(encoding="utf-8")


def test_polyform_license_has_project_notices_and_no_foreign_notice():
    license_text = _read("LICENSE")

    assert license_text.startswith("# PolyForm Noncommercial License 1.0.0")
    assert "<https://polyformproject.org/licenses/noncommercial/1.0.0>" in license_text
    assert "Required Notice: Copyright 2025-2026 KaguraNanaga." in license_text
    assert "released in version v1.0.0 and later" in license_text
    assert "Xueyou Luo" not in license_text
    assert "CAD files" not in license_text


def test_license_history_keeps_the_mit_boundary_explicit():
    history = _read("LICENSE-HISTORY.md")

    assert "v1.8.8.2" in history
    assert "MIT License" in history
    assert "v1.0.0" in history
    assert "PolyForm Noncommercial License 1.0.0" in history


def test_application_and_build_versions_match_current_release():
    gui_source = _read("NBC_DocFormat.py")
    build_source = _read("build.py")

    assert "__version__ = '1.0.2'" in gui_source
    assert 'VERSION = "1.0.2"' in build_source


def test_readme_declares_upstream_project_and_accurate_license_section():
    readme = _read("README.md")
    readme_en = _read("README_EN.md")

    assert "# NBC_DocFormat" in readme
    assert "基于原项目 [KaguraNanaga/docformat-gui](https://github.com/KaguraNanaga/docformat-gui) 修改开发" in readme
    assert "原作者：[KaguraNanaga](https://github.com/KaguraNanaga)" in readme
    assert "本改版在原项目基础上进行了 NBC 格式适配、界面品牌调整和多平台客户端构建" in readme
    assert "License-PolyForm%20Noncommercial%201.0.0" in readme
    assert "License-PolyForm%20Noncommercial%201.0.0" in readme_en
    assert "仅限个人和非商业用途" in readme
    assert "free for personal and noncommercial purposes" in readme_en


def test_packaged_apps_include_the_license_and_history_files():
    build_source = _read("build.py")
    workflow = _read(".github/workflows/build.yml")

    assert "--add-data=LICENSE;." in build_source
    assert build_source.count("--add-data=LICENSE:.") == 2
    assert "--add-data=LICENSE-HISTORY.md;." in build_source
    assert build_source.count("--add-data=LICENSE-HISTORY.md:.") == 2
    assert workflow.count('--add-data "LICENSE:."') == 2
    assert workflow.count('--add-data "LICENSE-HISTORY.md:."') == 2
