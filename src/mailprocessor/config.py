"""Configuration loading and validation."""

from __future__ import annotations

import re
import tomllib
from importlib import resources
from pathlib import Path
from typing import Literal

from pydantic import BaseModel, Field, ValidationError, model_validator

from mailprocessor.excel_writer import CONTENT_COLUMN


class AppSection(BaseModel):
    log_level: Literal["DEBUG", "INFO", "WARNING", "ERROR"] = "INFO"
    sqlite_path: str
    output_xlsx: str
    sheet_data: str = "daten"
    sheet_errors: str = "fehler"
    dry_run: bool = False
    max_messages: int = Field(default=0, ge=0)
    max_age_days: int = Field(default=0, ge=0)

    @model_validator(mode="after")
    def validate_sheet_names(self) -> AppSection:
        if self.sheet_data == self.sheet_errors:
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
    sender_filter: str | None = None


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


class AppConfig(BaseModel):
    app: AppSection
    source: SourceConfig


class FieldRule(BaseModel):
    column: str = Field(min_length=1)
    pattern: str = Field(min_length=1)
    required: bool = True

    @model_validator(mode="after")
    def validate_pattern_compiles(self) -> FieldRule:
        try:
            re.compile(self.pattern)
        except re.error as exc:
            raise ValueError(f"Invalid regex pattern for column '{self.column}': {exc}") from None
        return self


class ParsingRules(BaseModel):
    fields: list[FieldRule] = Field(min_length=1)

    @model_validator(mode="after")
    def validate_unique_columns(self) -> ParsingRules:
        columns = [field.column for field in self.fields]
        duplicates = sorted({column for column in columns if columns.count(column) > 1})
        if duplicates:
            raise ValueError(f"Duplicate parsing column names are not allowed: {', '.join(duplicates)}")
        if CONTENT_COLUMN in columns:
            raise ValueError(f"The column name '{CONTENT_COLUMN}' is reserved for the mail text")
        return self


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
