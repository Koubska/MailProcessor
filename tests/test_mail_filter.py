from pathlib import Path

import pytest

from mailprocessor.config import DEFAULT_FILES, MailFilter, load_app_config

CONFIG_START = """
[app]
sqlite_path = "./data/ledger.db"
output_xlsx = "./out/mail_export.xlsx"
"""


def _load(tmp_path: Path, text: str):
    path = tmp_path / "config.toml"
    path.write_text(CONFIG_START + text, encoding="utf-8")
    return load_app_config(path)


def test_without_filter_section_every_mail_matches(tmp_path: Path) -> None:
    config = _load(tmp_path, '[source]\ntype = "eml"\n[source.eml]\nfolder = "./mails"\n')

    assert not config.filter.active
    assert config.filter.matches(subject="Newsletter", sender="news@example.com")


def test_filter_section_accepts_a_text_or_a_list(tmp_path: Path) -> None:
    config = _load(
        tmp_path,
        '[source]\ntype = "eml"\n[source.eml]\nfolder = "./mails"\n'
        '[filter]\nsubject = "Kontaktformular"\nsender = ["schule@example.com", " ", "info@example.com"]\n',
    )

    assert config.filter.subject == ["Kontaktformular"]
    assert config.filter.sender == ["schule@example.com", "info@example.com"]
    assert config.filter.active


def test_old_imap_sender_filter_becomes_the_sender_filter(tmp_path: Path) -> None:
    config = _load(
        tmp_path,
        '[source]\ntype = "imap"\n[source.imap]\nhost = "imap.example.com"\nusername = "u"\n'
        'sender_filter = "schule@example.com"\n',
    )

    assert config.filter.sender == ["schule@example.com"]


def test_old_imap_sender_filter_does_not_override_the_new_one(tmp_path: Path) -> None:
    config = _load(
        tmp_path,
        '[source]\ntype = "imap"\n[source.imap]\nhost = "imap.example.com"\nusername = "u"\n'
        'sender_filter = "alt@example.com"\n[filter]\nsender = "neu@example.com"\n',
    )

    assert config.filter.sender == ["neu@example.com"]


def test_subject_matches_any_entry_ignoring_case_and_line_breaks() -> None:
    mail_filter = MailFilter(subject=["kontaktformular", "Anmeldung"])

    assert mail_filter.matches(subject="Neue Nachricht: KONTAKTFORMULAR", sender="")
    assert mail_filter.matches(subject="Anmeldung  für\n den Kurs", sender="")
    assert not mail_filter.matches(subject="Newsletter Oktober", sender="")


def test_entries_with_spaces_match_across_folded_headers() -> None:
    mail_filter = MailFilter(subject=["Anmeldung für den Kurs"])

    assert mail_filter.matches(subject="Anmeldung für\r\n den Kurs", sender="")


def test_sender_matches_name_or_address() -> None:
    mail_filter = MailFilter(sender=["schule@example.com", "Sekretariat"])

    assert mail_filter.matches(subject="", sender="Schule <SCHULE@example.com>")
    assert mail_filter.matches(subject="", sender="Sekretariat Muster <info@example.org>")
    assert not mail_filter.matches(subject="", sender="Max Mustermann <max.mustermann@mail.com>")


def test_subject_and_sender_must_both_match() -> None:
    mail_filter = MailFilter(subject=["Kontaktformular"], sender=["schule@example.com"])

    assert mail_filter.matches(subject="Kontaktformular", sender="schule@example.com")
    assert not mail_filter.matches(subject="Kontaktformular", sender="spam@example.com")
    assert not mail_filter.matches(subject="Werbung", sender="schule@example.com")


@pytest.mark.parametrize("entries", [[""], ["  "], []])
def test_blank_entries_are_no_filter(entries: list[str]) -> None:
    assert not MailFilter(subject=entries, sender=entries).active


def test_the_example_in_the_default_config_works_when_uncommented(tmp_path: Path) -> None:
    text = (DEFAULT_FILES / "config.toml").read_text(encoding="utf-8")
    example = text
    for commented in ("[filter]", "subject =", "sender ="):
        example = example.replace(f"# {commented}", commented)
    path = tmp_path / "config.toml"
    path.write_text(example, encoding="utf-8")

    assert load_app_config(path).filter == MailFilter(subject=["Kontaktformular"], sender=["schule@example.com"])
