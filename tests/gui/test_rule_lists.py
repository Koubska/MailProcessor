"""The profiles and fields edited in the Felder tab."""

import pytest

from mailprocessor.config import FieldRule, Profile
from mailprocessor.gui.rule_lists import ProfileList, RuleList, column_note, complete_column


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
