"""Reading and writing config.toml and parsing_rules.toml from the GUI."""

from pathlib import Path

import pytest
from gui_helpers import SIMPLE_RULES

from mailprocessor.config import FieldRule, Profile
from mailprocessor.gui.config_files import (
    output_file_path,
    parse_config_text,
    parse_rules_text,
    path_setting,
    render_config_text,
    render_rules_text,
)


def _profile(*fields: FieldRule, name: str = "Standard") -> Profile:
    # Not validated, so tests can also write invalid rules files.
    return Profile.model_construct(name=name, fields=list(fields))


def test_parse_rules_text_reads_fields() -> None:
    rules_text = '[[fields]]\ncolumn = "Name"\npattern = "(?im)^Name:\\\\s*(.+)$"\nrequired = true\n'

    profiles = parse_rules_text(rules_text)

    # A file from before profiles existed is one profile.
    assert profiles == [
        Profile(name="Standard", fields=[FieldRule(column="Name", pattern=r"(?im)^Name:\s*(.+)$", required=True)])
    ]


def test_parse_rules_text_reads_profiles() -> None:
    rules_text = (
        '[[profiles]]\nname = "Anmeldung"\n\n'
        '[[profiles.fields]]\ncolumn = "Kurs"\ntype = "label"\nlabel = "Angebot:"\n\n'
        '[[profiles]]\nname = "Abmeldung"\n\n'
        '[[profiles.fields]]\ncolumn = "Grund"\ntype = "label"\nlabel = "Grund:"\n'
    )

    profiles = parse_rules_text(rules_text)

    assert [profile.name for profile in profiles] == ["Anmeldung", "Abmeldung"]
    assert [field.column for field in profiles[1].fields] == ["Grund"]
    assert parse_rules_text("") == []


def test_render_rules_text_roundtrip() -> None:
    original = [
        Profile(
            name="Anmeldung",
            fields=[
                FieldRule(column="Name", pattern=r"(?im)^Name:\s*(.+)$", required=True),
                FieldRule(column="Telefonnummer", pattern=r"(?im)^Telefonnummer:\s*(.+)$", required=False),
            ],
        ),
        Profile(name="Abmeldung", fields=[FieldRule(column="Name", type="label", label="Name:")]),
    ]

    rendered = render_rules_text(original)
    reparsed = parse_rules_text(rendered)

    assert reparsed == original


def test_render_config_text_roundtrip() -> None:
    config_text = """
[app]
log_level = "INFO"
sqlite_path = "./data/ledger.db"
output_xlsx = "./out/mail_export.xlsx"
sheet_data = "daten"
sheet_errors = "fehler"
dry_run = false
max_messages = 0
max_age_days = 4

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

[filter]
subject = ["Kontaktformular", "Anmeldung \\"neu\\""]
""".strip()

    parsed = parse_config_text(config_text)
    rendered = render_config_text(parsed)
    reparsed = parse_config_text(rendered)

    # Everything round-trips except the password, which is never written to disk.
    assert "password" not in rendered
    assert parsed.source.imap is not None
    parsed.source.imap.password = None
    assert reparsed == parsed
    # The old IMAP-only sender filter is written as part of the filter for every source.
    assert "sender_filter" not in rendered
    assert reparsed.filter.sender == ["schule@example.com"]
    assert reparsed.filter.subject == ["Kontaktformular", 'Anmeldung "neu"']


def test_render_config_text_leaves_out_an_empty_filter() -> None:
    config = parse_config_text(
        '[app]\nsqlite_path = "l.db"\noutput_xlsx = "o.xlsx"\n[source]\ntype = "eml"\n[source.eml]\nfolder = "m"\n'
    )

    assert "[filter]" not in render_config_text(config)


def test_render_rules_text_escapes_control_characters() -> None:
    original = [
        Profile(name='Say "hi" now', fields=[FieldRule(column='Say "hi"\tnow', pattern="a\\b\nc\x7f", required=True)])
    ]

    assert parse_rules_text(render_rules_text(original)) == original


def test_mail_text_column_name_is_reserved() -> None:
    text = render_rules_text([_profile(FieldRule(column="E-Mail-Inhalt", pattern="a", required=True))])

    with pytest.raises(ValueError, match="is reserved"):
        parse_rules_text(text)


def test_parse_rules_text_rejects_duplicate_columns() -> None:
    text = render_rules_text([_profile(FieldRule(column="Name", pattern="a"), FieldRule(column="Name", pattern="b"))])

    with pytest.raises(ValueError, match="Duplicate parsing column"):
        parse_rules_text(text)


def test_render_rules_text_roundtrips_all_rule_types() -> None:
    text = render_rules_text([Profile(name="Standard", fields=SIMPLE_RULES)])

    assert parse_rules_text(text)[0].fields == SIMPLE_RULES
    assert 'label = ["Von", "From"]' in text
    assert 'label = "Angebot:"' in text
    # Simple rules are stored by their inputs only; regex rules keep the original format.
    assert text.count("pattern =") == 1
    assert text.count("type =") == len(SIMPLE_RULES) - 1


def test_path_setting_is_relative_inside_config_folder(tmp_path: Path) -> None:
    (tmp_path / "mails" / "2026").mkdir(parents=True)

    assert path_setting(tmp_path / "mails", tmp_path) == "./mails"
    assert path_setting(tmp_path / "mails" / "2026", tmp_path) == "./mails/2026"
    assert path_setting(tmp_path, tmp_path) == "."


def test_path_setting_is_absolute_outside_config_folder(tmp_path: Path) -> None:
    (tmp_path / "app").mkdir()
    (tmp_path / "elsewhere").mkdir()

    assert path_setting(tmp_path / "elsewhere", tmp_path / "app") == str((tmp_path / "elsewhere").resolve())


def test_path_setting_works_for_files(tmp_path: Path) -> None:
    assert path_setting(tmp_path / "out" / "export.xlsx", tmp_path) == "./out/export.xlsx"


def test_output_file_path_is_relative_to_config_folder(tmp_path: Path) -> None:
    assert output_file_path("./out/x.xlsx", tmp_path) == (tmp_path / "out" / "x.xlsx").resolve()
    absolute = tmp_path / "elsewhere.xlsx"
    assert output_file_path(str(absolute), Path("/unused")) == absolute
