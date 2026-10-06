"""Configuration loading and validation."""

from __future__ import annotations

import re
import tomllib
from collections.abc import Callable
from importlib import resources
from pathlib import Path
from typing import Literal

from pydantic import BaseModel, Field, ValidationError, field_validator, model_validator

from mailprocessor.columns import FIXED_DATA_COLUMNS, PROFILE_COLUMN
from mailprocessor.rule_patterns import LABEL_TYPES, RuleType, build_pattern


class AppSection(BaseModel):
    log_level: Literal["DEBUG", "INFO", "WARNING", "ERROR"] = "INFO"
    sqlite_path: str
    output_xlsx: str
    sheet_data: str = "daten"
    sheet_errors: str = "fehler"
    dry_run: bool = False
    max_messages: int = Field(default=0, ge=0)
    max_age_days: int = Field(default=0, ge=0)
    # Where the rows of several profiles go: "shared" = all in sheet_data with a "Profil" column,
    # "per_profile" = one sheet per profile, named like the profile.
    profile_sheets: Literal["shared", "per_profile"] = "shared"

    @model_validator(mode="after")
    def validate_sheet_names(self) -> AppSection:
        for setting in ("sheet_data", "sheet_errors"):
            problem = sheet_name_problem(getattr(self, setting), setting)
            if problem:
                raise ValueError(problem)
        # Excel compares sheet names ignoring case, so "Daten" and "daten" would be the same sheet.
        if self.sheet_data.casefold() == self.sheet_errors.casefold():
            raise ValueError("sheet_data and sheet_errors must be different sheet names")
        return self


class EmlSourceConfig(BaseModel):
    folder: str
    glob: str = "*.eml"


class ImapSourceConfig(BaseModel):
    host: str
    port: int = 993
    username: str
    # Optional on purpose: prefer the MAILPROCESSOR_IMAP_PASSWORD env var or the interactive prompt.
    password: str | None = None
    mailbox: str = "INBOX"
    use_ssl: bool = True


class SourceConfig(BaseModel):
    type: Literal["imap", "eml"]
    eml: EmlSourceConfig | None = None
    imap: ImapSourceConfig | None = None

    @model_validator(mode="after")
    def validate_source_details(self) -> SourceConfig:
        if self.type == "eml" and self.eml is None:
            raise ValueError("source.eml is required when source.type is 'eml'")
        if self.type == "imap" and self.imap is None:
            raise ValueError("source.imap is required when source.type is 'imap'")
        return self


def _collapse_spaces(text: str) -> str:
    return " ".join(text.split())


def _text_entries(value: object, clean: Callable[[str], str] = str.strip) -> object:
    """Accept a single text or a list of texts; clean each and drop the empty ones."""
    if isinstance(value, str):
        value = [value]
    if isinstance(value, list):
        cleaned = (clean(item) for item in value if isinstance(item, str))
        return [item for item in cleaned if item]
    return value


class MailFilter(BaseModel):
    """Which mails a run reads at all. Others are left out without an error and are read once the filter changes.

    Each list matches if the mail's subject (or From header: name and address) contains one of its entries,
    ignoring case and line breaks. An empty list matches every mail; both lists must match.
    """

    subject: list[str] = Field(default_factory=list)
    sender: list[str] = Field(default_factory=list)

    @field_validator("subject", "sender", mode="before")
    @classmethod
    def entries_as_list(cls, value: object) -> object:
        return _text_entries(value, _collapse_spaces)

    @property
    def active(self) -> bool:
        return bool(self.subject or self.sender)

    def matches(self, subject: str, sender: str) -> bool:
        return _contains_any(subject, self.subject) and _contains_any(sender, self.sender)


def _contains_any(text: str, entries: list[str]) -> bool:
    if not entries:
        return True
    folded = _collapse_spaces(text).casefold()
    return any(entry.casefold() in folded for entry in entries)


class AppConfig(BaseModel):
    app: AppSection
    source: SourceConfig
    filter: MailFilter = Field(default_factory=MailFilter)

    @model_validator(mode="before")
    @classmethod
    def move_old_sender_filter(cls, data: object) -> object:
        """Before the filter existed, IMAP had its own `source.imap.sender_filter`; it is now `filter.sender`."""
        # Only raw data from a file has it; already built sections are passed through.
        source = data.get("source") if isinstance(data, dict) else None
        imap = source.get("imap") if isinstance(source, dict) else None
        if not isinstance(imap, dict) or "sender_filter" not in imap:
            return data
        old = imap["sender_filter"]
        source = {**source, "imap": {key: value for key, value in imap.items() if key != "sender_filter"}}
        mail_filter = dict(data.get("filter") or {})
        mail_filter.setdefault("sender", old)
        return {**data, "source": source, "filter": mail_filter}


class FieldRule(BaseModel):
    """One Excel column. Simple types are described by labels/texts; `regex` takes a hand-written pattern."""

    column: str = Field(min_length=1)
    # Rules written before the simple types existed have only a pattern, so "regex" is the default.
    type: RuleType = "regex"
    label: list[str] = Field(default_factory=list)
    start: str | None = None
    end: str | None = None
    pattern: str | None = None
    required: bool = True

    @field_validator("label", mode="before")
    @classmethod
    def label_as_list(cls, value: object) -> object:
        return _text_entries(value)

    @model_validator(mode="after")
    def validate_inputs_for_type(self) -> FieldRule:
        if self.type in LABEL_TYPES - {"email"} and not self.label:
            raise ValueError(f"Column '{self.column}': a label is required")
        if self.type == "between" and not ((self.start or "").strip() and (self.end or "").strip()):
            raise ValueError(f"Column '{self.column}': start and end text are required")
        if self.type == "regex" and not self.pattern:
            raise ValueError(f"Column '{self.column}': a pattern is required")
        try:
            re.compile(self.regex)
        except re.error as exc:
            raise ValueError(f"Invalid regex pattern for column '{self.column}': {exc}") from None
        return self

    @property
    def regex(self) -> str:
        """The pattern the parser runs; generated for the simple types."""
        return build_pattern(self.type, labels=self.label, start=self.start, end=self.end, pattern=self.pattern)


# The name of the single profile in a rules file that only has [[fields]] (written before profiles existed).
DEFAULT_PROFILE_NAME = "Standard"
# Profile names can become sheet names, so they follow Excel's rules for those.
SHEET_NAME_MAX_LENGTH = 31
SHEET_NAME_FORBIDDEN = "[]:*?/\\"


def sheet_name_problem(name: str, what: str = "sheet name") -> str | None:
    """Why `name` cannot be an Excel sheet name, or None if it can. `what` names the setting in the message."""
    if not name.strip():
        return f"a {what} is required"
    if name != name.strip():
        return f"{what} '{name}' must not start or end with spaces"
    if len(name) > SHEET_NAME_MAX_LENGTH:
        return f"{what} '{name}' is longer than {SHEET_NAME_MAX_LENGTH} characters"
    forbidden = sorted({char for char in name if char in SHEET_NAME_FORBIDDEN})
    if forbidden:
        return f"{what} '{name}' must not contain {' '.join(forbidden)}"
    if name.startswith("'") or name.endswith("'"):
        return f"{what} '{name}' must not start or end with an apostrophe"
    return None


def profile_name_problem(name: str) -> str | None:
    """Why `name` cannot be a profile (and sheet) name, or None if it can."""
    return sheet_name_problem(name, "profile name")


class Profile(BaseModel):
    """One kind of mail (e.g. a registration form): its own fields. Each mail is assigned to the best-fitting one."""

    name: str
    fields: list[FieldRule] = Field(min_length=1)

    @model_validator(mode="after")
    def validate_profile(self) -> Profile:
        problem = profile_name_problem(self.name)
        if problem:
            raise ValueError(problem[0].upper() + problem[1:])
        columns = [field.column for field in self.fields]
        duplicates = sorted({column for column in columns if columns.count(column) > 1})
        if duplicates:
            raise ValueError(f"Duplicate parsing column names are not allowed: {', '.join(duplicates)}")
        for reserved in (*FIXED_DATA_COLUMNS, PROFILE_COLUMN):
            if reserved in columns:
                raise ValueError(f"The column name '{reserved}' is reserved: the app fills that column itself")
        return self


class ParsingRules(BaseModel):
    profiles: list[Profile] = Field(min_length=1)

    @model_validator(mode="before")
    @classmethod
    def fields_as_single_profile(cls, data: object) -> object:
        """A file with only [[fields]] (written before profiles existed) is one profile."""
        if isinstance(data, dict) and "fields" in data:
            if "profiles" in data:
                raise ValueError("Use either [[fields]] or [[profiles]], not both")
            rest = {key: value for key, value in data.items() if key != "fields"}
            return {**rest, "profiles": [{"name": DEFAULT_PROFILE_NAME, "fields": data["fields"]}]}
        return data

    @model_validator(mode="after")
    def validate_unique_names(self) -> ParsingRules:
        # Case-insensitive, like Excel's sheet names.
        names = [profile.name.casefold() for profile in self.profiles]
        duplicates = sorted({profile.name for profile in self.profiles if names.count(profile.name.casefold()) > 1})
        if duplicates:
            raise ValueError(f"Duplicate profile names are not allowed: {', '.join(duplicates)}")
        return self

    @property
    def columns(self) -> list[str]:
        """All field columns of all profiles, each once, in the order they first appear."""
        return list(dict.fromkeys(field.column for profile in self.profiles for field in profile.fields))


DEFAULT_FILES = resources.files("mailprocessor") / "defaults"


def create_missing_files(config_path: Path, rules_path: Path) -> list[Path]:
    """First start: write the default config/rules where missing, plus the default .eml folder.

    Existing files are never touched. Returns the files that were created.
    """
    created = []
    for target in (config_path, rules_path):
        if not target.exists():
            target.parent.mkdir(parents=True, exist_ok=True)
            default_name = "config.toml" if target is config_path else "parsing_rules.toml"
            target.write_text((DEFAULT_FILES / default_name).read_text(encoding="utf-8"), encoding="utf-8")
            created.append(target)
    if config_path in created:
        eml = load_app_config(config_path).source.eml
        if eml is not None:
            Path(eml.folder).mkdir(parents=True, exist_ok=True)
    return created


def _load_toml(path: Path) -> dict:
    with path.open("rb") as file_obj:
        try:
            return tomllib.load(file_obj)
        except tomllib.TOMLDecodeError as exc:
            raise ValueError(f"{path} is not valid TOML: {exc}") from None


def _resolve(base_dir: Path, raw: str) -> str:
    path = Path(raw).expanduser()
    return str(path if path.is_absolute() else (base_dir / path).resolve())


def resolve_relative_paths(config: AppConfig, base_dir: Path) -> AppConfig:
    """Make relative paths independent of the working directory (e.g. double-clicked executables)."""
    config.app.sqlite_path = _resolve(base_dir, config.app.sqlite_path)
    config.app.output_xlsx = _resolve(base_dir, config.app.output_xlsx)
    if config.source.eml is not None:
        config.source.eml.folder = _resolve(base_dir, config.source.eml.folder)
    return config


def load_app_config(path: Path) -> AppConfig:
    data = _load_toml(path)
    try:
        config = AppConfig.model_validate(data)
    except ValidationError as exc:
        raise ValueError(f"Invalid app config: {exc}") from exc
    return resolve_relative_paths(config, path.resolve().parent)


def load_parsing_rules(path: Path) -> ParsingRules:
    data = _load_toml(path)
    try:
        return ParsingRules.model_validate(data)
    except ValidationError as exc:
        raise ValueError(f"Invalid parsing rules: {exc}") from exc
