"""Desktop GUI for editing config/rules and running the pipeline."""

from __future__ import annotations

import json
import logging
import os
import queue
import socket
import ssl
import subprocess
import sys
import tomllib
from pathlib import Path

from pydantic import ValidationError

from mailprocessor.config import (
    AppConfig,
    AppSection,
    EmlSourceConfig,
    FieldRule,
    ImapSourceConfig,
    ParsingRules,
    SourceConfig,
)
from mailprocessor.errors import (
    ImapLoginError,
    MailFolderNotFoundError,
    MissingPasswordError,
    SheetColumnsError,
    WorkbookLockedError,
)
from mailprocessor.i18n import t
from mailprocessor.processor import RunSummary
from mailprocessor.rule_patterns import LABEL_TYPES, RuleType

LOG_FORMAT = "%(asctime)s %(levelname)-7s %(message)s"


class QueueLogHandler(logging.Handler):
    """Forwards formatted log lines from the worker thread to the UI thread as ("log", text) items."""

    def __init__(self, sink: queue.Queue) -> None:
        super().__init__()
        self.sink = sink
        self.setFormatter(logging.Formatter(LOG_FORMAT, datefmt="%H:%M:%S"))

    def emit(self, record: logging.LogRecord) -> None:
        try:
            self.sink.put(("log", self.format(record)))
        except Exception:
            self.handleError(record)


def _read_text_or_empty(path: Path) -> str:
    if not path.exists():
        return ""
    return path.read_text(encoding="utf-8")


def parse_rules_text(text: str) -> list[FieldRule]:
    raw = tomllib.loads(text) if text.strip() else {"fields": []}
    return ParsingRules.model_validate(raw).fields


def split_labels(text: str) -> list[str]:
    """Several labels are typed separated by ";"."""
    return [part.strip() for part in text.split(";") if part.strip()]


def _quote(text: str, lang: str) -> str:
    return f"„{text}“" if lang == "de" else f"“{text}”"


def describe_rule(rule: FieldRule, lang: str) -> str:
    """What a rule looks for, in words, for the rule list."""
    if rule.type == "regex":
        return rule.pattern or ""
    if rule.type == "between":
        start, end = _quote(rule.start or "", lang), _quote(rule.end or "", lang)
        return t("rule.describe.between", lang).format(start=start, end=end)
    labels = t("rule.or", lang).join(_quote(label, lang) for label in rule.label)
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


def parse_config_text(text: str) -> AppConfig:
    if not text.strip():
        raise ValueError("config.toml is empty")
    return AppConfig.model_validate(tomllib.loads(text))


def _toml_escape(value: str) -> str:
    # JSON string escaping is valid TOML basic-string escaping, except that TOML also forbids raw DEL.
    return json.dumps(value, ensure_ascii=False).replace("\x7f", "\\u007F")


def render_rules_text(fields: list[FieldRule]) -> str:
    """Write only the inputs that belong to each rule's type; regex rules keep the original format."""
    lines: list[str] = []
    for field in fields:
        lines.extend(["[[fields]]", f"column = {_toml_escape(field.column)}"])
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
    return "\n".join(lines).rstrip() + "\n"


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
            ]
        )
        if source.imap.sender_filter:
            lines.append(f"sender_filter = {_toml_escape(source.imap.sender_filter)}")
        lines.append("")
    return "\n".join(lines).rstrip() + "\n"


def path_setting(chosen: Path, config_dir: Path) -> str:
    """Config value for a picked folder or file.

    Relative if inside the config folder (keeps the setup portable), else absolute.
    """
    chosen, config_dir = chosen.resolve(), config_dir.resolve()
    if chosen.is_relative_to(config_dir):
        return f"./{chosen.relative_to(config_dir).as_posix()}".removesuffix("/.")
    return str(chosen)


def run_summary_text(summary: RunSummary, output_name: str, error_sheet: str, dry_run: bool, lang: str) -> str:
    """Plain-language result of a run for the output box and status bar."""
    parts = [t("summary.cancelled", lang)] if summary.cancelled else []
    if dry_run:
        parts.append(
            t("summary.dry_run", lang).format(
                new=summary.processed + summary.failed, ok=summary.processed, failed=summary.failed
            )
        )
        return " ".join(parts)
    if summary.processed:
        parts.append(t("summary.processed", lang).format(count=summary.processed, file=output_name))
    if summary.failed:
        parts.append(t("summary.failed", lang).format(count=summary.failed, sheet=error_sheet))
    if not summary.processed and not summary.failed:
        parts.append(t("summary.nothing_new", lang))
    if summary.skipped:
        parts.append(t("summary.skipped", lang).format(count=summary.skipped))
    return " ".join(parts)


# Most specific first: e.g. ImapLoginError and ssl.SSLError are OSErrors too.
_FRIENDLY_ERRORS: tuple[tuple[type[BaseException] | tuple[type[BaseException], ...], str], ...] = (
    (WorkbookLockedError, "error.workbook_locked"),
    (SheetColumnsError, "error.sheet_columns"),
    (MailFolderNotFoundError, "error.mail_folder_missing"),
    (MissingPasswordError, "error.imap_password_missing"),
    (ImapLoginError, "error.imap_login"),
    (ssl.SSLError, "error.imap_tls"),
    ((socket.gaierror, ConnectionError, TimeoutError), "error.imap_connection"),
)


def friendly_error(exc: BaseException, lang: str) -> str:
    """Explain a failed run without jargon; the technical message follows as details."""
    for error_types, key in _FRIENDLY_ERRORS:
        if isinstance(exc, error_types):
            return f"{t(key, lang)}\n\n{t('error.details', lang)}: {exc}"
    if isinstance(exc, (ValueError, OSError)):
        return str(exc)
    return t("error.unexpected", lang).format(name=type(exc).__name__)


def setting_path(setting: str, config_dir: Path) -> Path:
    """A path from the config; relative settings are relative to the config file, like in a run."""
    path = Path(setting).expanduser()
    return path if path.is_absolute() else (config_dir / path).resolve()


def output_file_path(output_setting: str, config_dir: Path) -> Path:
    """The Excel file a run writes to."""
    return setting_path(output_setting or _default_config().app.output_xlsx, config_dir)


def open_in_default_app(path: Path) -> None:
    """Open a file with the program the system uses for it (Excel, LibreOffice, Numbers, ...)."""
    if sys.platform == "win32":
        os.startfile(path)
    elif sys.platform == "darwin":
        subprocess.Popen(["open", str(path)])
    else:
        subprocess.Popen(["xdg-open", str(path)])


def _default_config() -> AppConfig:
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


def launch_gui(config_path: Path | None = None, rules_path: Path | None = None) -> None:
    # tkinter is imported only here, so the CLI and the tests of the functions above never need it.
    from mailprocessor.gui_app import run_app

    run_app(config_path or Path("config.toml"), rules_path or Path("parsing_rules.toml"))
