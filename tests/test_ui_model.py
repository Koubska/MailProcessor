from pathlib import Path

import pytest
from openpyxl import Workbook

from mailprocessor.config import FieldRule, MailFilter, Profile
from mailprocessor.gui import _default_config
from mailprocessor.parser import ParseResult
from mailprocessor.preview import RulePreview
from mailprocessor.processor import Problem
from mailprocessor.ui_model import (
    CardStatus,
    ProfileList,
    RuleList,
    column_note,
    complete_column,
    config_from_form,
    excel_status,
    fields_status,
    form_values,
    mails_status,
    problem_text,
    profiles_status,
    shortcuts,
)


def test_form_values_roundtrip_to_the_same_config() -> None:
    config = _default_config()

    rebuilt, errors = config_from_form(form_values(config), "de")

    assert errors == {}
    assert rebuilt == config


def test_form_keeps_the_profile_sheets_choice() -> None:
    config = _default_config()
    config.app.profile_sheets = "per_profile"

    rebuilt, _errors = config_from_form(form_values(config), "de")

    assert rebuilt is not None and rebuilt.app.profile_sheets == "per_profile"


@pytest.mark.parametrize(
    ("changes", "key", "message"),
    [
        ({"eml_folder": " "}, "eml_folder", "Ordner auswählen"),
        ({"max_age_days": "drei"}, "max_age_days", "ganze Zahl"),
        ({"max_messages": "-1"}, "max_messages", "ganze Zahl"),
        ({"output_xlsx": "out/export.csv"}, "output_xlsx", ".xlsx"),
        ({"sheet_errors": "daten"}, "sheet_errors", "verschiedene Namen"),
        ({"source_type": "imap", "imap_host": ""}, "imap_host", "E-Mail-Server"),
        ({"source_type": "imap", "imap_port": "99999"}, "imap_port", "Portnummer"),
    ],
)
def test_invalid_inputs_are_reported_per_field(changes: dict, key: str, message: str) -> None:
    values = {**form_values(_default_config()), **changes}

    config, errors = config_from_form(values, "de")

    assert config is None
    assert message in errors[key]


def test_invalid_port_of_the_inactive_source_does_not_block_saving() -> None:
    values = {**form_values(_default_config()), "imap_port": "abc"}

    config, errors = config_from_form(values, "de")

    assert errors == {}
    assert config is not None and config.source.imap is not None and config.source.imap.port == 993


def test_mails_status_for_folder(tmp_path: Path) -> None:
    values = {**form_values(_default_config()), "eml_folder": "./mails"}

    assert mails_status(values, tmp_path, False, "de").ok is False  # missing
    (tmp_path / "mails").mkdir()
    assert "ist leer" in mails_status(values, tmp_path, False, "de").text
    (tmp_path / "mails" / "a.eml").write_text("x", encoding="utf-8")
    status = mails_status(values, tmp_path, False, "de")
    assert status.ok and status.text == "Ordner „mails“ – 1 E-Mail(s)"


def test_filter_entries_are_typed_separated_by_semicolons() -> None:
    values = {
        **form_values(_default_config()),
        "filter_subject": "Kontaktformular;  Anmeldung ;",
        "filter_sender": "schule@example.com",
    }

    config, errors = config_from_form(values, "de")

    assert errors == {}
    assert config.filter == MailFilter(subject=["Kontaktformular", "Anmeldung"], sender=["schule@example.com"])
    assert form_values(config)["filter_subject"] == "Kontaktformular; Anmeldung"
    assert config_from_form(form_values(config), "de")[0] == config


def test_mails_status_names_an_active_filter(tmp_path: Path) -> None:
    (tmp_path / "mails").mkdir()
    (tmp_path / "mails" / "a.eml").write_text("x", encoding="utf-8")
    values = {**form_values(_default_config()), "eml_folder": "./mails", "filter_subject": "Kontakt; Anmeldung"}

    assert mails_status(values, tmp_path, False, "de").text == (
        "Ordner „mails“ – 1 E-Mail(s) · nur mit Betreff „Kontakt“ oder „Anmeldung“"
    )
    values = {**values, "source_type": "imap", "filter_subject": "", "filter_sender": "schule@example.com"}
    assert mails_status(values, tmp_path, True, "en").text == (
        "Mailbox user@example.com on imap.example.com · only from “schule@example.com”"
    )


def test_mails_status_for_imap_asks_for_the_password() -> None:
    values = {**form_values(_default_config()), "source_type": "imap"}

    assert "Passwort" in mails_status(values, Path("."), False, "de").text
    assert mails_status(values, Path("."), True, "de").text == "Postfach user@example.com auf imap.example.com"


def test_fields_status() -> None:
    rules = [FieldRule(column="Name", type="label", label="Name"), FieldRule(column="Tel", type="label", label="Tel")]

    assert fields_status([], None, "de").ok is False
    assert fields_status(rules, None, "de").ok is None
    assert fields_status(rules, [RulePreview("a"), RulePreview("b")], "de").ok is True
    missing = fields_status(rules, [RulePreview("a"), RulePreview(None)], "de")
    assert missing.ok is False and missing.text == "2 Feld(er) – in der Beispiel-Mail fehlt „Tel“"


def test_excel_status_counts_rows(tmp_path: Path) -> None:
    path = tmp_path / "out.xlsx"
    assert "wird beim ersten Übertragen angelegt" in excel_status(path, ["daten"], "fehler", "de").text

    workbook = Workbook()
    workbook.active.title = "daten"
    workbook.active.append(["Name"])
    workbook.active.append(["A"])
    workbook.active.append(["B"])
    workbook.create_sheet("fehler").append(["reason"])
    workbook["fehler"].append(["x"])
    workbook.save(path)

    assert excel_status(path, ["daten"], "fehler", "de").text == "out.xlsx – 2 Zeile(n), 1 E-Mail(s) mit Problemen"
    (tmp_path / "broken.xlsx").write_text("not excel", encoding="utf-8")
    assert excel_status(tmp_path / "broken.xlsx", ["daten"], "fehler", "de").ok is None


def test_problem_text() -> None:
    missing = Problem(name="a.eml", reason="Required fields missing: Tel, Kurs", missing=("Tel", "Kurs"), body="x")
    unreadable = Problem(name="file:b.eml", reason="Could not read file (OSError)")

    assert problem_text(missing, "de") == "nicht gefunden: „Tel“, „Kurs“"
    assert problem_text(unreadable, "de") == "E-Mail konnte nicht gelesen werden"
    assert problem_text(Problem(name="c", reason="Internal parser error", body="x"), "en") == "Internal parser error"


def _rule(column: str, label: str | None = None) -> FieldRule:
    return FieldRule(column=column, type="label", label=label or column)


def test_rule_list_add_validates_like_the_rules_file() -> None:
    rules = RuleList([_rule("Name")])

    assert rules.add(_rule("Kurs")) == 1
    with pytest.raises(ValueError, match="Duplicate"):
        rules.add(_rule("Kurs"))
    with pytest.raises(ValueError, match="reserved"):
        rules.add(_rule("E-Mail-Inhalt"))
    assert [rule.column for rule in rules] == ["Name", "Kurs"]


def test_rule_list_replace_reports_changes_and_keeps_invalid_out() -> None:
    rules = RuleList([_rule("Name"), _rule("Kurs")])

    assert rules.replace(0, _rule("Name")) is False
    assert rules.replace(0, _rule("Name", "Vorname:")) is True
    with pytest.raises(ValueError):
        rules.replace(0, _rule("Kurs"))
    assert rules[0].label == ["Vorname:"]


def test_rule_list_remove_can_be_undone_once() -> None:
    rules = RuleList([_rule("Name"), _rule("Kurs"), _rule("Zeit")])

    removed = rules.remove(1)

    assert removed.column == "Kurs"
    assert [rule.column for rule in rules] == ["Name", "Zeit"]
    assert rules.can_undo
    assert rules.undo_remove() == 1
    assert [rule.column for rule in rules] == ["Name", "Kurs", "Zeit"]
    assert rules.undo_remove() is None


def test_rule_list_other_changes_end_the_undo() -> None:
    rules = RuleList([_rule("Name"), _rule("Kurs")])
    rules.remove(0)
    rules.add(_rule("Zeit"))

    assert not rules.can_undo
    assert rules.undo_remove() is None


def test_rule_list_undo_does_not_create_a_duplicate() -> None:
    rules = RuleList([_rule("Name"), _rule("Kurs")])
    rules.remove(1)
    rules.replace(0, _rule("Kurs"))  # the remaining field now uses the removed name

    assert rules.undo_remove() is None
    assert [rule.column for rule in rules] == ["Kurs"]


def test_rule_list_move() -> None:
    rules = RuleList([_rule("A"), _rule("B"), _rule("C")])

    assert rules.move(0, 1) == 1
    assert [rule.column for rule in rules] == ["B", "A", "C"]
    assert rules.move(0, -1) is None
    assert rules.move(2, 1) is None
    assert [rule.column for rule in rules] == ["B", "A", "C"]


def test_rule_list_columns_and_parsing_rules() -> None:
    rules = RuleList([_rule("A"), _rule("B")])

    assert rules.columns(except_index=0) == ["B"]
    assert rules.parsing_rules().profiles[0].fields == rules.rules
    assert not RuleList([])


@pytest.mark.parametrize(
    ("platform", "lang", "labels", "run_sequence"),
    [
        ("darwin", "de", ("⌘↩", "⇧⌘↩", "⌘E"), "<Command-Return>"),
        ("win32", "de", ("Strg+Enter", "Strg+Umschalt+Enter", "Strg+E"), "<Control-Return>"),
        ("linux", "en", ("Ctrl+Enter", "Ctrl+Shift+Enter", "Ctrl+E"), "<Control-Return>"),
    ],
)
def test_shortcuts_follow_the_platform(platform: str, lang: str, labels: tuple[str, ...], run_sequence: str) -> None:
    result = shortcuts(platform, lang)

    assert tuple(result[action].label for action in ("run", "test_run", "open_excel")) == labels
    assert result["run"].sequences == (run_sequence,)
    assert result["test_run"].sequences[0].endswith("Shift-Return>")
    # Caps Lock must not break the letter shortcut.
    assert {sequence[-2] for sequence in result["open_excel"].sequences} == {"e", "E"}


def test_excel_status_adds_up_the_profile_sheets(tmp_path: Path) -> None:
    path = tmp_path / "out.xlsx"
    workbook = Workbook()
    workbook.active.title = "Anmeldung"
    workbook.active.append(["Name"])
    workbook.active.append(["A"])
    abmeldung = workbook.create_sheet("Abmeldung")
    abmeldung.append(["Name"])
    abmeldung.append(["B"])
    abmeldung.append(["C"])
    workbook.save(path)

    assert excel_status(path, ["Anmeldung", "Abmeldung"], "fehler", "de").text == "out.xlsx – 3 Zeile(n)"


def test_problem_text_names_the_closest_profile() -> None:
    problem = Problem(name="a.eml", reason="", missing=("Kurs",), body="x", profile="Abmeldung")

    assert problem_text(problem, "de") == "nicht gefunden: „Kurs“ (am ähnlichsten: Profil „Abmeldung“)"


def test_profiles_status() -> None:
    fits = ParseResult({"Kurs": "Judo"}, [], None, "Anmeldung")
    misses = ParseResult({}, ["Kurs"], "missing", "Anmeldung")

    assert profiles_status(2, 7, None, "de") == CardStatus(None, "2 Profile, 7 Feld(er)")
    assert profiles_status(2, 7, fits, "de") == CardStatus(
        True, "2 Profile, 7 Feld(er) – die Beispiel-Mail passt zu „Anmeldung“"
    )
    assert profiles_status(2, 7, misses, "de").ok is False


def _profile(name: str, *columns: str) -> Profile:
    return Profile(name=name, fields=[_rule(column) for column in columns])


def test_profile_list_starts_with_one_empty_default_profile() -> None:
    profiles = ProfileList([])

    assert profiles.names == ["Standard"]
    assert not profiles.current
    assert profiles.profiles() == []
    with pytest.raises(ValueError):
        profiles.parsing_rules()


def test_profile_list_switches_adds_renames_and_removes() -> None:
    profiles = ProfileList([_profile("Anmeldung", "Name", "Kurs"), _profile("Abmeldung", "Name")])
    assert (profiles.current_name, profiles.current.columns(), profiles.field_count) == (
        "Anmeldung",
        ["Name", "Kurs"],
        3,
    )

    profiles.select(1)
    assert profiles.current.columns() == ["Name"]

    assert profiles.add(" Warteliste ", "de") == 2
    assert profiles.current_name == "Warteliste"
    # Saved and run without the new profile until it has fields.
    assert [profile.name for profile in profiles.profiles()] == ["Anmeldung", "Abmeldung"]
    profiles.current.add(_rule("Kurs"))
    assert [profile.name for profile in profiles.parsing_rules().profiles] == ["Anmeldung", "Abmeldung", "Warteliste"]

    profiles.rename("Nachrücker", "de")
    assert profiles.names == ["Anmeldung", "Abmeldung", "Nachrücker"]

    assert profiles.remove() == "Nachrücker"
    assert (profiles.names, profiles.index) == (["Anmeldung", "Abmeldung"], 1)
    profiles.remove()
    with pytest.raises(ValueError):
        profiles.remove()  # the last profile stays


@pytest.mark.parametrize(
    ("name", "message"),
    [
        ("  ", "Bitte einen Namen"),
        ("x" * 32, "höchstens 31 Zeichen"),
        ("Kurs/Tag", "Zeichen"),
        ("anmeldung", "gibt es schon"),
    ],
)
def test_profile_names_are_checked_with_readable_messages(name: str, message: str) -> None:
    profiles = ProfileList([_profile("Anmeldung", "Name")])

    with pytest.raises(ValueError, match=message):
        profiles.add(name, "de")
    assert profiles.names == ["Anmeldung"]


def test_renaming_a_profile_may_change_only_its_case() -> None:
    profiles = ProfileList([_profile("anmeldung", "Name")])

    profiles.rename("Anmeldung", "de")

    assert profiles.names == ["Anmeldung"]


def _two_profiles() -> ProfileList:
    profiles = ProfileList(
        [
            Profile(
                name="Anmeldung", fields=[_rule("Mail-Adresse"), _rule("Kurs", "Angebot:"), _rule("Telefonnummer")]
            ),
            Profile(name="Abmeldung", fields=[_rule("Kurs", "Angebot:"), _rule("Grund")]),
        ]
    )
    profiles.add("Warteliste", "de")
    profiles.current.add(_rule("Kurs", "Angebot:"))
    return profiles


def test_column_suggestions_come_from_the_other_profiles() -> None:
    profiles = _two_profiles()

    # "Kurs" is already used in the shown profile, so it is not offered again.
    assert profiles.column_suggestions("") == ["Mail-Adresse", "Telefonnummer", "Grund"]
    # Matches anywhere in the name, ignoring case.
    assert profiles.column_suggestions("ADR") == ["Mail-Adresse"]
    assert profiles.column_suggestions("nummer") == ["Telefonnummer"]
    # While editing "Kurs" itself, its name stays available.
    assert "Kurs" in profiles.column_suggestions("", except_index=0)


def test_complete_column_continues_what_was_typed() -> None:
    suggestions = ["Mail-Adresse", "Telefonnummer", "Grund"]

    assert complete_column("tel", suggestions) == "Telefonnummer"
    assert complete_column("Grund", suggestions) is None  # nothing left to add
    assert complete_column("", suggestions) is None
    assert complete_column("x", suggestions) is None


def test_rule_for_column_comes_from_the_first_other_profile() -> None:
    profiles = _two_profiles()

    rule = profiles.rule_for_column("Kurs")

    assert rule is not None and rule.label == ["Angebot:"]
    assert profiles.rule_for_column("Unbekannt") is None


def test_column_note_explains_shared_and_own_columns() -> None:
    profiles = _two_profiles()

    shared = column_note(profiles, "Kurs", False, "de")
    assert (
        shared.text
        == "Gleiche Spalte wie in den Profilen „Anmeldung“, „Abmeldung“ – in Excel stehen die Werte untereinander."
    )
    assert column_note(profiles, "Grund", False, "de").text.startswith("Gleiche Spalte wie im Profil „Abmeldung“")
    assert column_note(profiles, "Wunsch", False, "de").text.startswith("Eigene Spalte dieses Profils")
    assert column_note(profiles, "Kurs", True, "de").text == "Spalte im Blatt „Warteliste“ (ein Blatt pro Profil)."


def test_column_note_warns_about_near_duplicates() -> None:
    note = column_note(_two_profiles(), "telefon-nummer", False, "de")

    assert note.warning
    assert "„Telefonnummer“ (Profil „Anmeldung“)" in note.text and "zweite Spalte" in note.text


def test_column_note_is_empty_with_one_profile() -> None:
    assert column_note(ProfileList([Profile(name="A", fields=[_rule("Kurs")])]), "Kurs", False, "de").text == ""
