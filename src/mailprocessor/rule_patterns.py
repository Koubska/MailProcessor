"""Regular expressions for the simple rule types, so users never have to write regex themselves.

Label matching is tolerant: case-insensitive, any spacing, optional colon, label at the start of a line.
User input is always escaped, so characters like "+", "(" or "." in a label are matched literally.
"""

from __future__ import annotations

import re
from typing import Literal

RuleType = Literal["label", "next_line", "between", "email", "regex"]
RULE_TYPES: tuple[RuleType, ...] = ("label", "next_line", "between", "email", "regex")
LABEL_TYPES: frozenset[str] = frozenset({"label", "next_line", "email"})

EMAIL_PATTERN = r"[a-z0-9._%+-]+@[a-z0-9.-]+\.[a-z]{2,}"
# Spacing within a line: any whitespace except a line break, so also the non-breaking space (U+00A0)
# that HTML forms ("&nbsp;") and Outlook put between label and value. The text has no "\r" (normalize_body).
_SPACE = r"[^\S\n]"
_LINE_START = rf"^{_SPACE}*"
# After a label: a colon, or at least a space (so "Tag" does not match "Tagesordnung").
_LABEL_END = rf"(?:{_SPACE}*:|{_SPACE}){_SPACE}*"


def _flexible(text: str) -> str:
    """Literal text where any run of whitespace in it matches any whitespace (also line breaks)."""
    return r"\s+".join(re.escape(word) for word in text.split())


def _labels(labels: list[str]) -> str:
    """Alternatives for the label text; a trailing colon typed by the user is optional when matching."""
    alternatives = []
    for label in labels:
        words = label.strip().removesuffix(":").split()
        alternatives.append(f"{_SPACE}+".join(re.escape(word) for word in words))
    return "(?:" + "|".join(alternatives) + ")"


def build_pattern(
    rule_type: RuleType,
    *,
    labels: list[str] | None = None,
    start: str | None = None,
    end: str | None = None,
    pattern: str | None = None,
) -> str:
    """The regex for a rule. Group 1 is the value, as for hand-written patterns."""
    if rule_type == "regex":
        return pattern or ""
    if rule_type == "between":
        return rf"(?is){_flexible(start or '')}\s*(\S.*?)\s*{_flexible(end or '')}"
    if rule_type == "email" and not labels:
        return rf"(?i)({EMAIL_PATTERN})"
    label = _LINE_START + _labels(labels or [])
    if rule_type == "label":
        return rf"(?im){label}{_LABEL_END}(\S.*?){_SPACE}*$"
    if rule_type == "next_line":
        # The label stands alone on its line; the value is the next non-empty line.
        return rf"(?im){label}{_SPACE}*:?{_SPACE}*\n(?:{_SPACE}*\n)*{_SPACE}*(\S.*?){_SPACE}*$"
    if rule_type == "email":
        return rf"(?im){label}{_LABEL_END}.*?({EMAIL_PATTERN})"
    raise ValueError(f"Unknown rule type: {rule_type}")
