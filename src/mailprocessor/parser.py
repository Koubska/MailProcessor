"""Mail body parsing."""

from __future__ import annotations

import re
from dataclasses import dataclass

from mailprocessor.config import ParsingRules


@dataclass(frozen=True)
class ParseResult:
    values: dict[str, str]
    missing_required: list[str]
    error_reason: str | None


def normalize_body(text: str) -> str:
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    lines = [line.rstrip() for line in text.split("\n")]
    return "\n".join(lines).strip()


def extract_field(pattern: str, text: str) -> str | None:
    match = re.search(pattern, text)
    if not match:
        return None
    # Group 1 if the pattern has groups (it may not have participated in the match), else the whole match.
    value = match.group(1) if match.re.groups else match.group(0)
    if value is None:
        return None
    return " ".join(value.split()) or None


def parse_mail(body_text: str, rules: ParsingRules, header_text: str = "") -> ParseResult:
    """Match each rule against the body followed by the header lines.

    The body comes first, so a label in the text (e.g. "Von:" in a forwarded mail) wins over the
    mail's own header; headers are the fallback (e.g. the sender address in "From:").
    """
    normalized = normalize_body(f"{body_text}\n\n{header_text}" if header_text else body_text)
    values: dict[str, str] = {}
    missing_required: list[str] = []

    for field_rule in rules.fields:
        extracted = extract_field(field_rule.regex, normalized)
        if extracted is None:
            if field_rule.required:
                missing_required.append(field_rule.column)
            continue
        values[field_rule.column] = extracted

    reason = f"Required fields missing: {', '.join(missing_required)}" if missing_required else None
    return ParseResult(values=values, missing_required=missing_required, error_reason=reason)
