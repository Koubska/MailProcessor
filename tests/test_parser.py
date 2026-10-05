from mailprocessor.config import FieldRule, ParsingRules
from mailprocessor.parser import extract_field, normalize_body, parse_mail

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
