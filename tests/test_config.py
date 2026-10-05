from pathlib import Path

import pytest

from mailprocessor.config import create_missing_files, load_app_config, load_parsing_rules


def test_load_app_config_for_eml_source(tmp_path: Path) -> None:
    cfg_file = tmp_path / "config.toml"
    cfg_file.write_text(
        """
[app]
log_level = "INFO"
sqlite_path = "./data/ledger.db"
output_xlsx = "./out/mail_export.xlsx"
sheet_data = "daten"
sheet_errors = "fehler"
dry_run = false
max_messages = 0

[source]
type = "eml"

[source.eml]
folder = "./samples/eml"
glob = "*.eml"
""".strip()
    )

    config = load_app_config(cfg_file)

    assert config.source.type == "eml"
    assert config.source.eml is not None
    assert config.source.imap is None


def test_load_parsing_rules_rejects_duplicate_columns(tmp_path: Path) -> None:
    rules_file = tmp_path / "rules.toml"
    rules_file.write_text(
        """
[[fields]]
column = "Name"
pattern = "(?im)^Name:\\s*(.+)$"
required = true

[[fields]]
column = "Name"
pattern = "(?im)^Kind:\\s*(.+)$"
required = true
""".strip()
    )

    with pytest.raises(ValueError):
        load_parsing_rules(rules_file)


def test_load_app_config_for_imap_source_with_sender_filter(tmp_path: Path) -> None:
    cfg_file = tmp_path / "config.toml"
    cfg_file.write_text(
        """
[app]
log_level = "INFO"
sqlite_path = "./data/ledger.db"
output_xlsx = "./out/mail_export.xlsx"
sheet_data = "daten"
sheet_errors = "fehler"
dry_run = false
max_messages = 0

[source]
type = "imap"

[source.imap]
host = "imap.example.com"
port = 993
username = "user@example.com"
password = "plaintext-password"
mailbox = "INBOX"
use_ssl = true
sender_filter = "schule@example.com"
""".strip()
    )

    config = load_app_config(cfg_file)

    assert config.source.type == "imap"
    assert config.source.imap is not None
    assert config.source.imap.sender_filter == "schule@example.com"


def _write_minimal_config(path: Path, extra_app: str = "") -> None:
    path.write_text(
        f"""
[app]
sqlite_path = "./data/ledger.db"
output_xlsx = "out/mail_export.xlsx"
{extra_app}

[source]
type = "eml"

[source.eml]
folder = "./mails"
""".strip()
    )


def test_relative_paths_resolve_against_config_directory(tmp_path: Path, monkeypatch) -> None:
    cfg_dir = tmp_path / "bundle"
    cfg_dir.mkdir()
    _write_minimal_config(cfg_dir / "config.toml")
    monkeypatch.chdir(tmp_path)

    config = load_app_config(cfg_dir / "config.toml")

    assert config.app.sqlite_path == str(cfg_dir.resolve() / "data" / "ledger.db")
    assert config.app.output_xlsx == str(cfg_dir.resolve() / "out" / "mail_export.xlsx")
    assert config.source.eml is not None
    assert config.source.eml.folder == str(cfg_dir.resolve() / "mails")


def test_invalid_regex_is_rejected_at_load_time(tmp_path: Path) -> None:
    rules_file = tmp_path / "rules.toml"
    rules_file.write_text('[[fields]]\ncolumn = "Name"\npattern = "(unclosed"\n')

    with pytest.raises(ValueError, match="Invalid regex pattern for column 'Name'"):
        load_parsing_rules(rules_file)


def test_invalid_toml_is_a_value_error(tmp_path: Path) -> None:
    rules_file = tmp_path / "rules.toml"
    rules_file.write_text("[[fields]\n")

    with pytest.raises(ValueError, match="not valid TOML"):
        load_parsing_rules(rules_file)


@pytest.mark.parametrize("extra", ['log_level = "LOUD"', "max_messages = -1", 'sheet_errors = "daten"'])
def test_invalid_app_values_are_rejected(tmp_path: Path, extra: str) -> None:
    _write_minimal_config(tmp_path / "config.toml", extra)

    with pytest.raises(ValueError):
        load_app_config(tmp_path / "config.toml")


def test_create_missing_files_writes_defaults_and_mail_folder(tmp_path: Path) -> None:
    created = create_missing_files(tmp_path / "config.toml", tmp_path / "parsing_rules.toml")

    assert created == [tmp_path / "config.toml", tmp_path / "parsing_rules.toml"]
    assert (tmp_path / "mails").is_dir()
    assert load_parsing_rules(tmp_path / "parsing_rules.toml").fields
    assert load_app_config(tmp_path / "config.toml").source.type == "eml"


def test_create_missing_files_never_overwrites(tmp_path: Path) -> None:
    (tmp_path / "config.toml").write_text("mine", encoding="utf-8")

    created = create_missing_files(tmp_path / "config.toml", tmp_path / "parsing_rules.toml")

    assert created == [tmp_path / "parsing_rules.toml"]
    assert (tmp_path / "config.toml").read_text(encoding="utf-8") == "mine"
    assert not (tmp_path / "mails").exists()
