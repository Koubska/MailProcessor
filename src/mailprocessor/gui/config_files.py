"""The two settings files as the GUI reads and writes them, and the paths inside them.

The GUI rewrites config.toml and parsing_rules.toml itself (never the IMAP password); paths in the
config are relative to the config file, like in a run.
"""

from __future__ import annotations

import json
import tomllib
from pathlib import Path

from mailprocessor.config import (
    AppConfig,
    AppSection,
    EmlSourceConfig,
    FieldRule,
    ImapSourceConfig,
    ParsingRules,
    Profile,
    SourceConfig,
)
from mailprocessor.rule_patterns import LABEL_TYPES


def read_text_or_empty(path: Path) -> str:
    if not path.exists():
        return ""
    return path.read_text(encoding="utf-8")


def parse_rules_text(text: str) -> list[Profile]:
    """The profiles of a rules file; a file with only [[fields]] is one profile."""
    if not text.strip():
        return []
    return ParsingRules.model_validate(tomllib.loads(text)).profiles


def parse_config_text(text: str) -> AppConfig:
    if not text.strip():
        raise ValueError("config.toml is empty")
    return AppConfig.model_validate(tomllib.loads(text))


def _toml_escape(value: str) -> str:
    # JSON string escaping is valid TOML basic-string escaping, except that TOML also forbids raw DEL.
    return json.dumps(value, ensure_ascii=False).replace("\x7f", "\\u007F")


def render_rules_text(profiles: list[Profile]) -> str:
    """Write each profile with its fields; only the inputs that belong to each rule's type."""
    lines: list[str] = []
    for profile in profiles:
        lines.extend(["[[profiles]]", f"name = {_toml_escape(profile.name)}", ""])
        lines.extend(_render_fields(profile.fields))
    return "\n".join(lines).rstrip() + "\n"


def _render_fields(fields: list[FieldRule]) -> list[str]:
    """Regex rules keep the original format (no type line)."""
    lines: list[str] = []
    for field in fields:
        lines.extend(["[[profiles.fields]]", f"column = {_toml_escape(field.column)}"])
        if field.type != "regex":
            lines.append(f"type = {_toml_escape(field.type)}")
        if field.type in LABEL_TYPES and field.label:
            if len(field.label) == 1:
                lines.append(f"label = {_toml_escape(field.label[0])}")
            else:
                lines.append(f"label = [{', '.join(_toml_escape(label) for label in field.label)}]")
        if field.type == "between":
            lines.extend([f"start = {_toml_escape(field.start or '')}", f"end = {_toml_escape(field.end or '')}"])
        if field.type == "regex":
            lines.append(f"pattern = {_toml_escape(field.pattern or '')}")
        lines.extend([f"required = {'true' if field.required else 'false'}", ""])
    return lines


def render_config_text(config: AppConfig) -> str:
    app = config.app
    source = config.source
    lines = [
        "[app]",
        f"log_level = {_toml_escape(app.log_level)}",
        f"sqlite_path = {_toml_escape(app.sqlite_path)}",
        f"output_xlsx = {_toml_escape(app.output_xlsx)}",
        f"sheet_data = {_toml_escape(app.sheet_data)}",
        f"sheet_errors = {_toml_escape(app.sheet_errors)}",
        f"profile_sheets = {_toml_escape(app.profile_sheets)}",
        f"dry_run = {'true' if app.dry_run else 'false'}",
        f"max_messages = {app.max_messages}",
        f"max_age_days = {app.max_age_days}",
        "",
        "[source]",
        f"type = {_toml_escape(source.type)}",
        "",
    ]
    if source.eml is not None:
        lines.extend(
            [
                "[source.eml]",
                f"folder = {_toml_escape(source.eml.folder)}",
                f"glob = {_toml_escape(source.eml.glob)}",
                "",
            ]
        )
    if source.imap is not None:
        lines.extend(
            [
                "[source.imap]",
                f"host = {_toml_escape(source.imap.host)}",
                f"port = {source.imap.port}",
                f"username = {_toml_escape(source.imap.username)}",
                # The password is intentionally never written to disk.
                f"mailbox = {_toml_escape(source.imap.mailbox)}",
                f"use_ssl = {'true' if source.imap.use_ssl else 'false'}",
                "",
            ]
        )
    if config.filter.active:
        lines.append("[filter]")
        for key, entries in (("subject", config.filter.subject), ("sender", config.filter.sender)):
            if entries:
                lines.append(f"{key} = [{', '.join(_toml_escape(entry) for entry in entries)}]")
        lines.append("")
    return "\n".join(lines).rstrip() + "\n"


def default_config() -> AppConfig:
    return AppConfig(
        app=AppSection(
            log_level="INFO",
            sqlite_path="./data/ledger.db",
            output_xlsx="./out/mail_export.xlsx",
            sheet_data="daten",
            sheet_errors="fehler",
            dry_run=False,
            max_messages=0,
            max_age_days=0,
        ),
        source=SourceConfig(
            type="eml",
            eml=EmlSourceConfig(folder="./mails", glob="*.eml"),
            imap=ImapSourceConfig(
                host="imap.example.com",
                port=993,
                username="user@example.com",
                mailbox="INBOX",
                use_ssl=True,
                sender_filter=None,
            ),
        ),
    )


def path_setting(chosen: Path, config_dir: Path) -> str:
    """Config value for a picked folder or file.

    Relative if inside the config folder (keeps the setup portable), else absolute.
    """
    chosen, config_dir = chosen.resolve(), config_dir.resolve()
    if chosen.is_relative_to(config_dir):
        return f"./{chosen.relative_to(config_dir).as_posix()}".removesuffix("/.")
    return str(chosen)


def setting_path(setting: str, config_dir: Path) -> Path:
    """A path from the config; relative settings are relative to the config file, like in a run."""
    path = Path(setting).expanduser()
    return path if path.is_absolute() else (config_dir / path).resolve()


def output_file_path(output_setting: str, config_dir: Path) -> Path:
    """The Excel file a run writes to."""
    return setting_path(output_setting or default_config().app.output_xlsx, config_dir)
