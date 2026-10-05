from pathlib import Path

import pytest

from mailprocessor.config import (
    DEFAULT_PROFILE_NAME,
    FieldRule,
    ParsingRules,
    Profile,
    create_missing_files,
    load_app_config,
    load_parsing_rules,
    profile_name_problem,
)


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
    assert config.filter.sender == ["schule@example.com"]


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


@pytest.mark.parametrize(
    "extra", ['log_level = "LOUD"', "max_messages = -1", 'sheet_errors = "daten"', 'profile_sheets = "each"']
)
def test_invalid_app_values_are_rejected(tmp_path: Path, extra: str) -> None:
    _write_minimal_config(tmp_path / "config.toml", extra)

    with pytest.raises(ValueError):
        load_app_config(tmp_path / "config.toml")


def test_create_missing_files_writes_defaults_and_mail_folder(tmp_path: Path) -> None:
    created = create_missing_files(tmp_path / "config.toml", tmp_path / "parsing_rules.toml")

    assert created == [tmp_path / "config.toml", tmp_path / "parsing_rules.toml"]
    assert (tmp_path / "mails").is_dir()
    assert load_parsing_rules(tmp_path / "parsing_rules.toml").profiles[0].fields
    assert load_app_config(tmp_path / "config.toml").source.type == "eml"


def test_create_missing_files_never_overwrites(tmp_path: Path) -> None:
    (tmp_path / "config.toml").write_text("mine", encoding="utf-8")

    created = create_missing_files(tmp_path / "config.toml", tmp_path / "parsing_rules.toml")

    assert created == [tmp_path / "parsing_rules.toml"]
    assert (tmp_path / "config.toml").read_text(encoding="utf-8") == "mine"
    assert not (tmp_path / "mails").exists()


@pytest.mark.parametrize("column", ["E-Mail-Inhalt", "Eingegangen am", "Übertragen am", "Profil"])
def test_columns_the_app_fills_itself_are_reserved(column: str) -> None:
    with pytest.raises(ValueError, match="is reserved"):
        ParsingRules(fields=[FieldRule(column=column, type="label", label="X")])


def _field(column: str) -> FieldRule:
    return FieldRule(column=column, type="label", label=f"{column}:")


def test_profile_sheets_defaults_to_one_shared_sheet(tmp_path: Path) -> None:
    _write_minimal_config(tmp_path / "config.toml", "")

    assert load_app_config(tmp_path / "config.toml").app.profile_sheets == "shared"


def test_load_parsing_rules_reads_profiles(tmp_path: Path) -> None:
    path = tmp_path / "parsing_rules.toml"
    path.write_text(
        """
[[profiles]]
name = "Anmeldung"

[[profiles.fields]]
column = "Name"
type = "label"
label = "Name:"

[[profiles.fields]]
column = "Kurs"
type = "label"
label = "Angebot:"

[[profiles]]
name = "Abmeldung"

[[profiles.fields]]
column = "Name"
type = "label"
label = "Name:"

[[profiles.fields]]
column = "Grund"
type = "label"
label = "Grund:"
required = false
""",
        encoding="utf-8",
    )

    rules = load_parsing_rules(path)

    assert [profile.name for profile in rules.profiles] == ["Anmeldung", "Abmeldung"]
    # The same column may appear in several profiles; `columns` lists each once.
    assert rules.columns == ["Name", "Kurs", "Grund"]


def test_fields_without_profiles_are_one_default_profile() -> None:
    rules = ParsingRules.model_validate({"fields": [{"column": "Name", "type": "label", "label": "Name:"}]})

    assert [(profile.name, len(profile.fields)) for profile in rules.profiles] == [(DEFAULT_PROFILE_NAME, 1)]


def test_fields_and_profiles_together_are_rejected() -> None:
    with pytest.raises(ValueError, match="either"):
        ParsingRules.model_validate(
            {
                "fields": [{"column": "A", "pattern": "a"}],
                "profiles": [{"name": "P", "fields": [{"column": "B", "pattern": "b"}]}],
            }
        )


def test_profile_names_must_be_unique_ignoring_case() -> None:
    with pytest.raises(ValueError, match="Duplicate profile names"):
        ParsingRules(
            profiles=[Profile(name="Anmeldung", fields=[_field("A")]), Profile(name="anmeldung", fields=[_field("B")])]
        )


def test_columns_are_unique_within_a_profile() -> None:
    with pytest.raises(ValueError, match="Duplicate parsing column"):
        Profile(name="P", fields=[_field("A"), _field("A")])


def test_a_profile_needs_fields() -> None:
    with pytest.raises(ValueError):
        Profile(name="P", fields=[])


@pytest.mark.parametrize("name", ["", " P", "a/b", "Kurs [neu]", "x" * 32, "'P'"])
def test_profile_names_must_be_valid_sheet_names(name: str) -> None:
    assert profile_name_problem(name) is not None
    with pytest.raises(ValueError):
        Profile(name=name, fields=[_field("A")])


def test_profile_name_may_contain_spaces_and_umlauts() -> None:
    assert profile_name_problem("Anmeldung Förderkurs") is None
