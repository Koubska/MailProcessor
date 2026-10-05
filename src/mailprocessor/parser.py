"""Mail body parsing."""

from __future__ import annotations

import re
from collections.abc import Sequence
from dataclasses import dataclass

from mailprocessor.config import ParsingRules, Profile


@dataclass(frozen=True)
class ParseResult:
    values: dict[str, str]
    missing_required: list[str]
    error_reason: str | None
    # The profile the values were read with: the best-fitting one (see `best_result`).
    profile: str = ""


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


def _parse_profile(text: str, profile: Profile) -> ParseResult:
    values: dict[str, str] = {}
    missing_required: list[str] = []
    for field_rule in profile.fields:
        extracted = extract_field(field_rule.regex, text)
        if extracted is None:
            if field_rule.required:
                missing_required.append(field_rule.column)
            continue
        values[field_rule.column] = extracted
    reason = f"Required fields missing: {', '.join(missing_required)}" if missing_required else None
    return ParseResult(values=values, missing_required=missing_required, error_reason=reason, profile=profile.name)


def parse_profiles(body_text: str, rules: ParsingRules, header_text: str = "") -> list[ParseResult]:
    """Match every profile against the body followed by the header lines; one result per profile, in order.

    The body comes first, so a label in the text (e.g. "Von:" in a forwarded mail) wins over the
    mail's own header; headers are the fallback (e.g. the sender address in "From:").
    """
    normalized = normalize_body(f"{body_text}\n\n{header_text}" if header_text else body_text)
    return [_parse_profile(normalized, profile) for profile in rules.profiles]


def best_result(results: Sequence[ParseResult]) -> ParseResult:
    """The best-fitting profile's result.

    Only profiles with every required field found can succeed; of those, the one with the most fields found wins.
    If none succeeds, the closest one (fewest required fields missing, then most found) explains the failure.
    Ties go to the profile listed first.
    """
    # Fewest missing first (so complete profiles beat all others), then most found; max() keeps the first of equals.
    best = max(results, key=lambda result: (-len(result.missing_required), len(result.values)))
    if best.missing_required and len(results) > 1:
        reason = f"Required fields missing (closest profile '{best.profile}'): {', '.join(best.missing_required)}"
        return ParseResult(best.values, best.missing_required, reason, best.profile)
    return best


def parse_mail(body_text: str, rules: ParsingRules, header_text: str = "") -> ParseResult:
    """Parse with every profile and keep the best-fitting result."""
    return best_result(parse_profiles(body_text, rules, header_text))
