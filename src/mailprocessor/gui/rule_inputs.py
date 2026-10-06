"""The rule editor's inputs: building a `FieldRule` from them and describing a rule in words."""

from __future__ import annotations

from pydantic import ValidationError

from mailprocessor.config import FieldRule
from mailprocessor.i18n import quote, t
from mailprocessor.rule_patterns import LABEL_TYPES, RuleType


def split_labels(text: str) -> list[str]:
    """Several labels are typed separated by ";"."""
    return [part.strip() for part in text.split(";") if part.strip()]


def describe_rule(rule: FieldRule, lang: str) -> str:
    """What a rule looks for, in words, for the rule list."""
    if rule.type == "regex":
        return rule.pattern or ""
    if rule.type == "between":
        start, end = quote(rule.start or "", lang), quote(rule.end or "", lang)
        return t("rule.describe.between", lang).format(start=start, end=end)
    labels = t("rule.or", lang).join(quote(label, lang) for label in rule.label)
    if rule.type == "email" and not rule.label:
        return t("rule.describe.email_anywhere", lang)
    return t(f"rule.describe.{rule.type}", lang).format(labels=labels)


def rule_from_inputs(
    *,
    column: str,
    rule_type: RuleType,
    labels_text: str = "",
    start: str = "",
    end: str = "",
    pattern: str = "",
    required: bool = True,
    other_columns: list[str] | None = None,
    lang: str = "de",
) -> FieldRule:
    """Build a rule from the editor inputs, with messages users understand. Only this type's inputs are kept."""
    column = column.strip()
    labels = split_labels(labels_text)
    if not column:
        raise ValueError(t("error.rule.column", lang))
    if column in (other_columns or []):
        raise ValueError(t("error.rule.duplicate", lang).format(column=column))
    if rule_type in LABEL_TYPES - {"email"} and not labels:
        raise ValueError(t("error.rule.label", lang))
    if rule_type == "between" and not (start.strip() and end.strip()):
        raise ValueError(t("error.rule.between", lang))
    if rule_type == "regex" and not pattern.strip():
        raise ValueError(t("error.rule.pattern", lang))
    try:
        return FieldRule(
            column=column,
            type=rule_type,
            label=labels if rule_type in LABEL_TYPES else [],
            start=start.strip() if rule_type == "between" else None,
            end=end.strip() if rule_type == "between" else None,
            pattern=pattern.strip() if rule_type == "regex" else None,
            required=required,
        )
    except ValidationError as exc:
        # e.g. an invalid regex; pydantic's message without its decoration
        raise ValueError(exc.errors()[0]["msg"].removeprefix("Value error, ")) from None
