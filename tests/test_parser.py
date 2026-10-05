from mailprocessor.config import FieldRule, ParsingRules, Profile
from mailprocessor.parser import extract_field, normalize_body, parse_mail, parse_profiles

SAMPLE_BODY = """Von: Max Mustermann <max.mustermann@mail.com>
Betreff: Schnuppernachmittag

Sehr geehrter Herr Test,

hiermit möchte ich mein Kind Jan Must+ für folgendes Angebot anmelden:
Tag: Mo, 23.11.2026 (14:15-16:00 Uhr)
Angebot: Experimente
Telefonnummer: 1234 567890
"""


def build_rules() -> ParsingRules:
    return ParsingRules(
        fields=[
            FieldRule(
                column="Mail-Adresse",
                pattern=r"(?im)^\s*Von:\s*.*?<([^>]+)>\s*$",
                required=True,
            ),
            FieldRule(
                column="Name",
                pattern=r"(?im)mein\s+Kind\s+(.+?)\s+f[uü]r\s+folgendes\s+Angebot",
                required=True,
            ),
            FieldRule(
                column="Kurs",
                pattern=r"(?im)^\s*Angebot:\s*(.+?)\s*$",
                required=True,
            ),
            FieldRule(
                column="Zeit",
                pattern=r"(?im)^\s*Tag:\s*(.+?)\s*$",
                required=True,
            ),
            FieldRule(
                column="Telefonnummer",
                pattern=r"(?im)^\s*Telefonnummer:\s*(.+?)\s*$",
                required=True,
            ),
        ]
    )


def test_parse_mail_successfully_extracts_all_required_fields() -> None:
    result = parse_mail(SAMPLE_BODY, build_rules())

    assert result.missing_required == []
    assert result.error_reason is None
    assert result.values == {
        "Mail-Adresse": "max.mustermann@mail.com",
        "Name": "Jan Must+",
        "Kurs": "Experimente",
        "Zeit": "Mo, 23.11.2026 (14:15-16:00 Uhr)",
        "Telefonnummer": "1234 567890",
    }


def test_parse_mail_reports_missing_required_fields() -> None:
    text_without_phone = SAMPLE_BODY.replace("Telefonnummer: 1234 567890", "")

    result = parse_mail(text_without_phone, build_rules())

    assert "Telefonnummer" in result.missing_required
    assert result.error_reason is not None


def test_normalize_body_makes_line_endings_deterministic() -> None:
    normalized = normalize_body("a\r\nb\rc  \n\n")

    assert normalized == "a\nb\nc"


def test_extract_field_treats_unmatched_optional_group_as_missing() -> None:
    assert extract_field(r"Name:(?:\s*(\w+))?|Kurs:(\w+)", "Kurs:Chemie") is None


def test_extract_field_without_group_collapses_whitespace() -> None:
    assert extract_field(r"Tel\S*\s+\d+\s+\d+", "Tel:  1234   5678") == "Tel: 1234 5678"


def _label(column: str, label: str, required: bool = True) -> FieldRule:
    return FieldRule(column=column, type="label", label=label, required=required)


ANMELDUNG = Profile(name="Anmeldung", fields=[_label("Kurs", "Angebot:"), _label("Zeit", "Tag:")])
ABMELDUNG = Profile(
    name="Abmeldung",
    fields=[_label("Kurs", "Angebot:"), _label("Grund", "Grund:"), _label("Telefon", "Telefonnummer:", required=False)],
)


def test_complete_profile_with_most_fields_found_wins() -> None:
    rules = ParsingRules(profiles=[ANMELDUNG, ABMELDUNG])
    text = "Angebot: Judo\nTag: Montag\nGrund: Umzug\nTelefonnummer: 123"

    result = parse_mail(text, rules)

    # Both profiles are complete; Abmeldung finds 3 fields, Anmeldung only 2.
    assert result.profile == "Abmeldung"
    assert result.values == {"Kurs": "Judo", "Grund": "Umzug", "Telefon": "123"}
    assert result.error_reason is None


def test_a_profile_with_a_missing_required_field_never_wins() -> None:
    rules = ParsingRules(profiles=[ABMELDUNG, ANMELDUNG])
    text = "Angebot: Judo\nTag: Montag\nTelefonnummer: 123"

    result = parse_mail(text, rules)

    # Abmeldung finds more (Kurs, Telefon) but misses its required "Grund".
    assert result.profile == "Anmeldung"
    assert result.values == {"Kurs": "Judo", "Zeit": "Montag"}


def test_tie_goes_to_the_first_profile() -> None:
    other = Profile(name="Andere", fields=[_label("Kurs", "Angebot:"), _label("Tag", "Tag:")])

    assert parse_mail("Angebot: Judo\nTag: Mo", ParsingRules(profiles=[other, ANMELDUNG])).profile == "Andere"
    assert parse_mail("Angebot: Judo\nTag: Mo", ParsingRules(profiles=[ANMELDUNG, other])).profile == "Anmeldung"


def test_without_a_complete_profile_the_closest_one_explains_the_failure() -> None:
    rules = ParsingRules(profiles=[ANMELDUNG, ABMELDUNG])

    result = parse_mail("Grund: Umzug\nTelefonnummer: 123", rules)

    # Anmeldung misses 2 required fields, Abmeldung only "Kurs".
    assert result.profile == "Abmeldung"
    assert result.missing_required == ["Kurs"]
    assert result.error_reason == "Required fields missing (closest profile 'Abmeldung'): Kurs"


def test_parse_profiles_returns_one_result_per_profile_in_order() -> None:
    results = parse_profiles("Angebot: Judo", ParsingRules(profiles=[ANMELDUNG, ABMELDUNG]))

    assert [(result.profile, result.missing_required) for result in results] == [
        ("Anmeldung", ["Zeit"]),
        ("Abmeldung", ["Grund"]),
    ]


def test_single_profile_keeps_the_plain_error_reason() -> None:
    result = parse_mail("Angebot: Judo", ParsingRules(profiles=[ANMELDUNG]))

    assert result.error_reason == "Required fields missing: Zeit"
