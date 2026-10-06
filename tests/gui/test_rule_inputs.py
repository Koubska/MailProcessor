"""The rule editor's inputs and the description of a rule."""

import pytest
from gui_helpers import SIMPLE_RULES

from mailprocessor.config import FieldRule
from mailprocessor.gui.rule_inputs import describe_rule, rule_from_inputs, split_labels


def test_describe_rule_in_words() -> None:
    described = [describe_rule(rule, "de") for rule in SIMPLE_RULES]

    assert described == [
        "E-Mail-Adresse in der Zeile „Von“ oder „From“",
        "Erste E-Mail-Adresse im Text",
        "Zwischen „mein Kind“ und „für folgendes Angebot“",
        "Zeile nach „Angebot:“",
        'Zeile unter „Say "hi"“',
        r"(?m)^X:\s*(.+)$",
    ]
    assert describe_rule(SIMPLE_RULES[0], "en") == "Email address in the line “Von” or “From”"


def test_split_labels() -> None:
    assert split_labels(" Von ; From;; ") == ["Von", "From"]


def test_rule_from_inputs_keeps_only_inputs_of_the_chosen_type() -> None:
    rule = rule_from_inputs(
        column=" Kurs ", rule_type="label", labels_text="Angebot:; Kurs", start="left over", pattern="left over"
    )

    assert rule == FieldRule(column="Kurs", type="label", label=["Angebot:", "Kurs"])


@pytest.mark.parametrize(
    ("kwargs", "message"),
    [
        ({"column": " ", "rule_type": "label", "labels_text": "A"}, "Spaltennamen"),
        ({"column": "Kurs", "rule_type": "label", "other_columns": ["Kurs"]}, "gibt es schon"),
        ({"column": "Kurs", "rule_type": "next_line", "labels_text": " ; "}, "Bezeichnung"),
        ({"column": "Name", "rule_type": "between", "start": "mein Kind"}, "Anfang und Ende"),
        ({"column": "X", "rule_type": "regex"}, "regulären Ausdruck"),
        ({"column": "X", "rule_type": "regex", "pattern": "("}, "Invalid regex pattern for column 'X'"),
    ],
)
def test_rule_from_inputs_explains_what_is_missing(kwargs: dict, message: str) -> None:
    with pytest.raises(ValueError, match=message):
        rule_from_inputs(**kwargs)
