import subprocess
import sys
from pathlib import Path

from typer.testing import CliRunner

import mailprocessor.main as main_module
from mailprocessor.config import (
    AppConfig,
    AppSection,
    EmlSourceConfig,
    FieldRule,
    ImapSourceConfig,
    ParsingRules,
    SourceConfig,
)
from mailprocessor.processor import RunSummary


def _build_config() -> AppConfig:
    return AppConfig(
        app=AppSection(
            log_level="INFO",
            sqlite_path="./data/ledger.db",
            output_xlsx="./out/mail_export.xlsx",
            sheet_data="daten",
            sheet_errors="fehler",
            dry_run=False,
            max_messages=0,
            max_age_days=0,
        ),
        source=SourceConfig(
            type="eml",
            eml=EmlSourceConfig(folder="./inbox", glob="*.eml"),
        ),
    )


def _build_rules() -> ParsingRules:
    return ParsingRules(fields=[FieldRule(column="Name", pattern=r"(?im)^Name:\s*(.+)$", required=True)])


def test_cli_runs_pipeline_with_loaded_config_and_rules(monkeypatch) -> None:
    runner = CliRunner()
    config_obj = _build_config()
    rules_obj = _build_rules()
    calls: dict[str, object] = {}

    def fake_load_app_config(path: Path) -> AppConfig:
        calls["config_path"] = path
        return config_obj

    def fake_load_parsing_rules(path: Path) -> ParsingRules:
        calls["rules_path"] = path
        return rules_obj

    def fake_run_pipeline(cfg: AppConfig, rules: ParsingRules) -> RunSummary:
        calls["config"] = cfg
        calls["rules"] = rules
        return RunSummary(seen=2, processed=1, skipped=1, failed=0)

    monkeypatch.setattr(main_module, "load_app_config", fake_load_app_config)
    monkeypatch.setattr(main_module, "load_parsing_rules", fake_load_parsing_rules)
    monkeypatch.setattr(main_module, "run_pipeline", fake_run_pipeline)

    result = runner.invoke(main_module.app, ["--config", "config.toml", "--rules", "rules.toml"])

    assert result.exit_code == 0
    assert "seen=2 processed=1 skipped=1 failed=0" in result.stdout
    assert calls["config"] is config_obj
    assert calls["rules"] is rules_obj


def test_cli_applies_dry_run_and_max_age_days_overrides(monkeypatch) -> None:
    runner = CliRunner()
    config_obj = _build_config()
    rules_obj = _build_rules()

    def fake_load_app_config(path: Path) -> AppConfig:
        return config_obj

    def fake_load_parsing_rules(path: Path) -> ParsingRules:
        return rules_obj

    def fake_run_pipeline(cfg: AppConfig, rules: ParsingRules) -> RunSummary:
        assert cfg.app.dry_run is True
        assert cfg.app.max_age_days == 5
        return RunSummary(seen=0, processed=0, skipped=0, failed=0)

    monkeypatch.setattr(main_module, "load_app_config", fake_load_app_config)
    monkeypatch.setattr(main_module, "load_parsing_rules", fake_load_parsing_rules)
    monkeypatch.setattr(main_module, "run_pipeline", fake_run_pipeline)

    result = runner.invoke(
        main_module.app,
        ["--config", "config.toml", "--rules", "rules.toml", "--dry-run", "--max-age-days", "5"],
    )

    assert result.exit_code == 0


def test_cli_returns_error_exit_code_on_failure(monkeypatch) -> None:
    runner = CliRunner()

    def fake_load_app_config(path: Path) -> AppConfig:
        raise ValueError("broken config")

    monkeypatch.setattr(main_module, "load_app_config", fake_load_app_config)

    result = runner.invoke(main_module.app, ["--config", "config.toml", "--rules", "rules.toml"])

    assert result.exit_code == 1
    assert "Error: broken config" in result.stdout


def test_main_module_runs_as_script() -> None:
    """PyInstaller executes main.py as a script; without the __main__ guard the binary does nothing."""

    script = Path(main_module.__file__)
    result = subprocess.run([sys.executable, str(script), "--help"], capture_output=True, text=True, check=False)

    assert result.returncode == 0
    assert "--config" in result.stdout


def test_imap_password_is_read_from_environment(monkeypatch) -> None:
    config = _build_config()
    config.source.type = "imap"
    config.source.imap = ImapSourceConfig(host="h", username="u")
    monkeypatch.setenv(main_module.PASSWORD_ENV_VAR, "from-env")

    main_module._resolve_imap_password(config)

    assert config.source.imap.password == "from-env"


def test_unexpected_errors_print_a_short_message(monkeypatch) -> None:
    runner = CliRunner()
    monkeypatch.setattr(main_module, "load_app_config", lambda path: _build_config())
    monkeypatch.setattr(main_module, "load_parsing_rules", lambda path: _build_rules())

    def boom(cfg, rules):
        raise KeyError("internal")

    monkeypatch.setattr(main_module, "run_pipeline", boom)

    result = runner.invoke(main_module.app, ["--config", "config.toml", "--rules", "rules.toml"])

    assert result.exit_code == 1
    assert "unexpected KeyError" in result.stdout
    assert "Traceback" not in result.stdout


def test_main_without_arguments_opens_gui_with_files_next_to_executable(monkeypatch, tmp_path: Path) -> None:
    calls = {}
    monkeypatch.setattr(sys, "argv", [str(tmp_path / "mailprocessor")])
    monkeypatch.setattr(sys, "frozen", True, raising=False)
    monkeypatch.setattr(sys, "executable", str(tmp_path / "mailprocessor"))
    monkeypatch.setattr("mailprocessor.gui.launch_gui", lambda config_path, rules_path: calls.update(cfg=config_path, rules=rules_path))

    main_module.main()

    assert calls == {"cfg": tmp_path.resolve() / "config.toml", "rules": tmp_path.resolve() / "parsing_rules.toml"}


def test_main_with_arguments_runs_cli(monkeypatch) -> None:
    called = []
    monkeypatch.setattr(sys, "argv", ["mailprocessor", "--help"])
    monkeypatch.setattr(main_module, "app", lambda: called.append("cli"))

    main_module.main()

    assert called == ["cli"]


def test_main_falls_back_to_cli_help_when_gui_cannot_open(monkeypatch) -> None:
    def no_display(config_path, rules_path):
        raise RuntimeError("no display name")

    called = []
    monkeypatch.setattr(sys, "argv", ["mailprocessor"])
    monkeypatch.setattr("mailprocessor.gui.launch_gui", no_display)
    monkeypatch.setattr(main_module, "app", lambda: called.append(list(sys.argv)))

    main_module.main()

    assert called == [["mailprocessor", "--help"]]


def test_cli_launches_gui_without_requiring_config_and_rules(monkeypatch) -> None:
    runner = CliRunner()
    calls: dict[str, object] = {}

    def fake_launch_gui(config: Path | None = None, rules: Path | None = None) -> None:
        calls["config"] = config
        calls["rules"] = rules

    def fail_run_pipeline(_cfg: AppConfig, _rules: ParsingRules) -> RunSummary:
        raise AssertionError("Pipeline must not run when --gui is set")

    monkeypatch.setattr(main_module, "launch_gui", fake_launch_gui)
    monkeypatch.setattr(main_module, "run_pipeline", fail_run_pipeline)

    result = runner.invoke(main_module.app, ["--gui"])

    assert result.exit_code == 0
    assert calls["config"] is None
    assert calls["rules"] is None


def test_cli_requires_config_and_rules_without_gui() -> None:
    runner = CliRunner()

    result = runner.invoke(main_module.app, [])

    assert result.exit_code == 1
    assert "--config and --rules are required unless --gui is used" in result.stdout
