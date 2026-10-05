import re

import pytest

from mailprocessor.config import FieldRule, ParsingRules
from mailprocessor.parser import parse_mail

TEXT = "\n".join(
    [
        "Von: Max Mustermann <max.mustermann@mail.com>",
        "hiermit möchte ich mein Kind Jan Must+ für folgendes Angebot anmelden:",
        "Tag: Mo, 23.11.2026 (14:15-16:00 Uhr)",
        "Tagesordnung: nicht gemeint",
        "Angebot:   Experimente  ",
        "telefonnummer 1234 567890",
        "Kosten (EUR): 12,50",
        "Bemerkung:",
        "",
        "  Bitte Turnschuhe mitbringen",
        "Leer:",
        "Nächste Zeile",
    ]
)


def _value(rule: FieldRule, text: str = TEXT) -> str | None:
    result = parse_mail(text, ParsingRules(fields=[rule.model_copy(update={"required": False})]))
    return result.values.get(rule.column)


@pytest.mark.parametrize(
    ("labels", "expected"),
    [
        (["Tag:"], "Mo, 23.11.2026 (14:15-16:00 Uhr)"),
        (["Angebot"], "Experimente"),  # colon optional, surrounding spaces dropped
        (["Telefonnummer:"], "1234 567890"),  # case-insensitive, colon missing in the text
        (["Kosten (EUR):"], "12,50"),  # special characters are literal
        (["Kosten  (eur)"], "12,50"),  # extra spaces in the label
        (["Gibt es nicht", "Angebot"], "Experimente"),  # alternatives
    ],
)
def test_label_finds_rest_of_line(labels: list[str], expected: str) -> None:
    assert _value(FieldRule(column="X", type="label", label=labels)) == expected


def test_label_is_not_a_word_prefix() -> None:
    # "Tag" must not match "Tagesordnung:"; the first real "Tag:" line is found.
    text = "Tagesordnung: nicht gemeint\nTag: Montag"
    assert _value(FieldRule(column="X", type="label", label="Tag"), text) == "Montag"


def test_label_must_start_the_line() -> None:
    assert _value(FieldRule(column="X", type="label", label="Turnschuhe"), TEXT) is None


def test_label_without_value_does_not_take_the_next_line() -> None:
    assert _value(FieldRule(column="X", type="label", label="Leer:")) is None


def test_next_line_skips_blank_lines() -> None:
    assert _value(FieldRule(column="X", type="next_line", label="Bemerkung")) == "Bitte Turnschuhe mitbringen"
    assert _value(FieldRule(column="X", type="next_line", label="Leer")) == "Nächste Zeile"


def test_between_spans_line_breaks_and_spacing() -> None:
    rule = FieldRule(column="X", type="between", start="mein  Kind", end="für folgendes\nAngebot")
    assert _value(rule) == "Jan Must+"
    wrapped = "mein Kind Jan\nMust+ für\nfolgendes Angebot"
    assert _value(rule, wrapped) == "Jan Must+"


def test_email_with_and_without_label() -> None:
    assert _value(FieldRule(column="X", type="email", label=["Von", "From"])) == "max.mustermann@mail.com"
    assert _value(FieldRule(column="X", type="email"), "Kontakt bitte an info@schule.de\n") == "info@schule.de"
    assert _value(FieldRule(column="X", type="email", label="Von"), "Von: Max ohne Adresse\n") is None


def test_regex_rules_keep_working() -> None:
    rule = FieldRule(column="X", pattern=r"(?im)^\s*Angebot:\s*(.+?)\s*$")
    assert rule.type == "regex"
    assert _value(rule) == "Experimente"


@pytest.mark.parametrize(
    ("kwargs", "message"),
    [
        ({"type": "label"}, "a label is required"),
        ({"type": "next_line", "label": ["  "]}, "a label is required"),
        ({"type": "between", "start": "a"}, "start and end text are required"),
        ({"type": "regex"}, "a pattern is required"),
        ({"type": "regex", "pattern": "("}, "Invalid regex pattern"),
    ],
)
def test_missing_inputs_are_rejected(kwargs: dict, message: str) -> None:
    with pytest.raises(ValueError, match=message):
        FieldRule(column="X", **kwargs)


def test_generated_patterns_compile_for_all_types() -> None:
    for rule in (
        FieldRule(column="a", type="label", label=["(", "a+b", "[x]"]),
        FieldRule(column="b", type="next_line", label="*"),
        FieldRule(column="c", type="between", start="?", end="\\"),
        FieldRule(column="d", type="email", label="$"),
    ):
        re.compile(rule.regex)
