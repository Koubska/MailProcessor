"""Rule preview against a sample mail, for the GUI's live test. No tkinter here, so it is testable.

Values are computed exactly like a run (`parser.extract_field` on the normalized text plus headers);
the position is only used to highlight the value in the displayed mail text.
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path

from mailprocessor.config import FieldRule, ParsingRules
from mailprocessor.i18n import t
from mailprocessor.parser import ParseResult, extract_field, normalize_body, parse_mail
from mailprocessor.sources.email_content import parse_message_bytes


@dataclass(frozen=True)
class SampleMail:
    title: str  # file name or a note such as "pasted text"
    body: str  # shown and editable in the GUI
    header_text: str = ""  # searched as a fallback, like in a run


@dataclass(frozen=True)
class RulePreview:
    value: str | None
    # Position of the value in the body text, or None if it was found in a header (or not at all).
    span: tuple[int, int] | None = None

    @property
    def found(self) -> bool:
        return self.value is not None

    @property
    def from_header(self) -> bool:
        return self.value is not None and self.span is None


def load_sample(path: Path) -> SampleMail:
    """Read a .eml file (read-only) the same way the eml source does."""
    mail = parse_message_bytes(path.read_bytes(), "eml", str(path.parent), origin=path.name)
    return SampleMail(title=path.name, body=normalize_body(mail.body_text), header_text=mail.header_text)


def sample_files(folder: Path, glob_pattern: str = "*.eml") -> list[Path]:
    """The mails a user can step through; empty if the folder does not exist."""
    if not folder.is_dir():
        return []
    return [path for path in sorted(folder.glob(glob_pattern)) if path.is_file()]


def preview_rule(rule: FieldRule, sample: SampleMail) -> RulePreview:
    text = f"{sample.body}\n\n{sample.header_text}" if sample.header_text else sample.body
    value = extract_field(rule.regex, normalize_body(text))
    if value is None:
        return RulePreview(value=None)
    match = re.search(rule.regex, sample.body)
    if match is None:
        return RulePreview(value=value)
    group = 1 if match.re.groups else 0
    if match.group(group) is None:
        return RulePreview(value=value)
    return RulePreview(value=value, span=match.span(group))


def _quoted_list(columns: list[str], lang: str) -> str:
    quote = ("„", "“") if lang == "de" else ("“", "”")
    return ", ".join(f"{quote[0]}{column}{quote[1]}" for column in columns)


def summary_text(rules: list[FieldRule], previews: list[RulePreview], error_sheet: str, lang: str) -> str:
    """What would happen to the sample mail in a run, in plain words."""
    missing = [(rule.column, rule.required) for rule, preview in zip(rules, previews, strict=True) if not preview.found]
    missing_required = [column for column, required in missing if required]
    missing_optional = [column for column, required in missing if not required]
    if missing_required:
        key = "preview.summary.missing_one" if len(missing_required) == 1 else "preview.summary.missing_many"
        return t(key, lang).format(columns=_quoted_list(missing_required, lang), sheet=error_sheet)
    if missing_optional:
        return t("preview.summary.ok_optional_missing", lang).format(columns=_quoted_list(missing_optional, lang))
    return t("preview.summary.ok", lang).format(count=len(rules))


def best_profile(rules: ParsingRules, sample: SampleMail) -> ParseResult:
    """The profile a run would choose for the sample mail (see `parser.best_result`)."""
    return parse_mail(sample.body, rules, sample.header_text)


def profile_summary_text(best: ParseResult, error_sheet: str, lang: str) -> str:
    """With several profiles: which one the sample mail fits, or which one comes closest and what it misses."""
    quote = ("„", "“") if lang == "de" else ("“", "”")
    profile = f"{quote[0]}{best.profile}{quote[1]}"
    if best.missing_required:
        return t("preview.summary.no_profile", lang).format(
            profile=profile, columns=_quoted_list(best.missing_required, lang), sheet=error_sheet
        )
    return t("preview.summary.profile", lang).format(profile=profile, count=len(best.values))


def result_text(preview: RulePreview | None, lang: str) -> str:
    """Short text for the rule list's result column."""
    if preview is None:
        return "–"
    if not preview.found:
        return t("preview.not_found", lang)
    return f"✓ {preview.value}"
