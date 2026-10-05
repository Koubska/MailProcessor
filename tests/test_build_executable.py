from __future__ import annotations

import importlib.util
from pathlib import Path


def _load_build_module():
    module_path = Path("scripts/build_executable.py")
    spec = importlib.util.spec_from_file_location("build_executable", module_path)
    assert spec is not None
    assert spec.loader is not None
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def test_resolve_executable_name_by_platform() -> None:
    module = _load_build_module()
    assert module._resolve_executable_name("windows") == "mailprocessor.exe"
    assert module._resolve_executable_name("macos") == "mailprocessor"


def test_icon_args_point_to_existing_icon_files() -> None:
    module = _load_build_module()
    for platform_label, suffix in (("windows", ".ico"), ("macos", ".icns")):
        flag, icon_path = module._icon_args(platform_label)
        assert flag == "--icon"
        assert icon_path.endswith(suffix)
        assert Path(icon_path).is_file()
    assert module._icon_args("linux") == []
    assert (module.ASSETS_DIR / "icon.png").is_file()


def test_write_bundle_readme_mentions_config_and_command(tmp_path: Path) -> None:
    module = _load_build_module()
    module._write_bundle_readme(tmp_path, "mailprocessor")
    readme = (tmp_path / "RUNNING.md").read_text(encoding="utf-8")

    assert "config.toml" in readme
    assert "parsing_rules.toml" in readme
    assert "./mailprocessor --config ./config.toml --rules ./parsing_rules.toml" in readme

