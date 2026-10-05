"""The shipped default rules must extract every field from the documented example email."""

from importlib import resources
from pathlib import Path

import pytest

from mailprocessor.config import DEFAULT_FILES, ParsingRules, load_parsing_rules
from mailprocessor.parser import parse_mail
from mailprocessor.sources.email_content import parse_message_bytes

REPO = Path(__file__).resolve().parents[1]
EXAMPLE = REPO / "docs" / "example.eml"

EXPECTED = {
    "Mail-Adresse": "max.mustermann@mail.com",
    "Name": "Jan Mustermann",
    "Kurs": "Experimente",
    "Zeit": "Mo, 23.11.2026 (14:15-16:00 Uhr)",
    "Telefonnummer": "1234 567890",
}


@pytest.fixture
def rules() -> ParsingRules:
    with resources.as_file(DEFAULT_FILES / "parsing_rules.toml") as path:
        return load_parsing_rules(path)


def _parse(raw: bytes, rules: ParsingRules) -> dict[str, str]:
    mail = parse_message_bytes(raw, "eml", "test")
    result = parse_mail(mail.body_text, rules, mail.header_text)
    assert result.missing_required == []
    return result.values


def test_example_file_as_is(rules: ParsingRules) -> None:
    # The example starts with "Von:"/"Betreff:" lines, which an email parser reads as headers.
    assert _parse(EXAMPLE.read_bytes(), rules) == EXPECTED


def test_example_with_windows_line_endings(rules: ParsingRules) -> None:
    raw = EXAMPLE.read_bytes().replace(b"\r\n", b"\n").replace(b"\n", b"\r\n")

    assert _parse(raw, rules) == EXPECTED


def test_example_text_inside_a_real_email(rules: ParsingRules) -> None:
    # Saved from a mail program: real headers, the example text (with its "Von:" line) in the body.
    raw = b"From: Kontaktformular <noreply@test.de>\nSubject: Anmeldung\n\n" + EXAMPLE.read_bytes()

    # The "Von:" line in the text wins over the form's From header.
    assert _parse(raw, rules) == EXPECTED


def test_display_name_hides_fallback_hash() -> None:
    mail = parse_message_bytes(EXAMPLE.read_bytes(), "eml", "test", origin="example.eml")

    assert mail.message_identity.startswith("fallback:")
    assert mail.display_name == "example.eml"


def test_sender_from_header_when_text_has_no_von_line(rules: ParsingRules) -> None:
    body = b"\n".join(line for line in EXAMPLE.read_bytes().splitlines() if not line.startswith(b"Von:"))
    raw = b"From: eltern@example.org\nSubject: Anmeldung\n\n" + body

    assert _parse(raw, rules)["Mail-Adresse"] == "eltern@example.org"
