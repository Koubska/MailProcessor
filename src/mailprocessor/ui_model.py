"""Form handling and status texts for the GUI, without tkinter (testable).

The GUI keeps every input as a plain value in a dict (`FormValues`). `config_from_form` turns it into an
`AppConfig` and reports problems per input, so the GUI can show them next to the field and save
automatically only when everything is valid.
"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from openpyxl import load_workbook
from pydantic import ValidationError

from mailprocessor.config import (
    AppConfig,
    AppSection,
    EmlSourceConfig,
    FieldRule,
    ImapSourceConfig,
    SourceConfig,
)
from mailprocessor.gui import setting_path
from mailprocessor.i18n import t
from mailprocessor.preview import RulePreview, sample_files

FormValues = dict[str, str | bool]

DEFAULT_IMAP_PORTS = {True: "993", False: "143"}


def form_values(config: AppConfig) -> FormValues:
    app, source = config.app, config.source
    eml = source.eml or EmlSourceConfig(folder="", glob="*.eml")
    imap = source.imap or ImapSourceConfig(host="", username="")
    return {
        "source_type": source.type,
        "eml_folder": eml.folder,
        "eml_glob": eml.glob,
        "imap_host": imap.host,
        "imap_port": str(imap.port),
        "imap_username": imap.username,
        "imap_mailbox": imap.mailbox,
        "imap_use_ssl": imap.use_ssl,
        "imap_sender_filter": imap.sender_filter or "",
        "output_xlsx": app.output_xlsx,
        "sheet_data": app.sheet_data,
        "sheet_errors": app.sheet_errors,
        "log_level": app.log_level,
        "sqlite_path": app.sqlite_path,
        "max_age_days": str(app.max_age_days),
        "max_messages": str(app.max_messages),
        "dry_run": app.dry_run,
    }


def _text(values: FormValues, key: str) -> str:
    return str(values.get(key, "")).strip()


def _count(values: FormValues, key: str, errors: dict[str, str], lang: str) -> int:
    raw = _text(values, key) or "0"
    if not raw.isdigit():
        errors[key] = t("error.form.number", lang)
        return 0
    return int(raw)


def config_from_form(values: FormValues, lang: str) -> tuple[AppConfig | None, dict[str, str]]:
    """The config for the inputs, or None plus a message per invalid input (key = input name)."""
    errors: dict[str, str] = {}
    source_type = _text(values, "source_type") or "eml"
    use_imap = source_type == "imap"

    if source_type == "eml" and not _text(values, "eml_folder"):
        errors["eml_folder"] = t("error.form.folder_empty", lang)
    if use_imap and not _text(values, "imap_host"):
        errors["imap_host"] = t("error.form.host_empty", lang)
    if use_imap and not _text(values, "imap_username"):
        errors["imap_username"] = t("error.form.username_empty", lang)
    port_text = _text(values, "imap_port") or DEFAULT_IMAP_PORTS[bool(values.get("imap_use_ssl", True))]
    port = int(port_text) if port_text.isdigit() and 0 < int(port_text) < 65536 else None
    if port is None and use_imap:
        errors["imap_port"] = t("error.form.port", lang)

    output = _text(values, "output_xlsx")
    if not output:
        errors["output_xlsx"] = t("error.form.output_empty", lang)
    elif not output.lower().endswith(".xlsx"):
        errors["output_xlsx"] = t("error.form.output_suffix", lang)
    sheet_data, sheet_errors = _text(values, "sheet_data"), _text(values, "sheet_errors")
    for key, name in (("sheet_data", sheet_data), ("sheet_errors", sheet_errors)):
        if not name:
            errors[key] = t("error.form.sheet_empty", lang)
    if sheet_data and sheet_data == sheet_errors:
        errors["sheet_errors"] = t("error.form.sheets_equal", lang)
    if not _text(values, "sqlite_path"):
        errors["sqlite_path"] = t("error.form.sqlite_empty", lang)
    max_age_days = _count(values, "max_age_days", errors, lang)
    max_messages = _count(values, "max_messages", errors, lang)
    if errors:
        return None, errors

    # The inactive source's settings are kept, so switching back and forth loses nothing.
    eml = None
    if source_type == "eml" or _text(values, "eml_folder"):
        eml = EmlSourceConfig(folder=_text(values, "eml_folder"), glob=_text(values, "eml_glob") or "*.eml")
    imap = None
    if use_imap or _text(values, "imap_host"):
        imap = ImapSourceConfig(
            host=_text(values, "imap_host"),
            port=port or int(DEFAULT_IMAP_PORTS[True]),
            username=_text(values, "imap_username"),
            mailbox=_text(values, "imap_mailbox") or "INBOX",
            use_ssl=bool(values.get("imap_use_ssl", True)),
            sender_filter=_text(values, "imap_sender_filter") or None,
        )
    try:
        config = AppConfig(
            app=AppSection(
                log_level=_text(values, "log_level") or "INFO",
                sqlite_path=_text(values, "sqlite_path"),
                output_xlsx=output,
                sheet_data=sheet_data,
                sheet_errors=sheet_errors,
                dry_run=bool(values.get("dry_run", False)),
                max_messages=max_messages,
                max_age_days=max_age_days,
            ),
            source=SourceConfig(type=source_type, eml=eml, imap=imap),
        )
    except ValidationError as exc:
        return None, {"_": exc.errors()[0]["msg"].removeprefix("Value error, ")}
    return config, {}


@dataclass(frozen=True)
class CardStatus:
    """One status card on the start page. ok: True = ready, False = needs attention, None = neutral."""

    ok: bool | None
    text: str


def _quote(text: str, lang: str) -> str:
    return f"„{text}“" if lang == "de" else f"“{text}”"


def mail_count_in_folder(values: FormValues, config_dir: Path) -> int | None:
    """Number of mail files in the configured folder, or None if the folder does not exist."""
    folder = setting_path(_text(values, "eml_folder") or ".", config_dir)
    if not folder.is_dir():
        return None
    return len(sample_files(folder, _text(values, "eml_glob") or "*.eml"))


def mails_status(values: FormValues, config_dir: Path, password_given: bool, lang: str) -> CardStatus:
    if _text(values, "source_type") == "imap":
        host, user = _text(values, "imap_host"), _text(values, "imap_username")
        if not host or not user:
            return CardStatus(False, t("card.mails.imap_incomplete", lang))
        if not password_given:
            return CardStatus(False, t("card.mails.imap_password", lang).format(user=user))
        return CardStatus(True, t("card.mails.imap_ready", lang).format(user=user, host=host))
    folder_name = _quote(Path(_text(values, "eml_folder") or ".").name or _text(values, "eml_folder"), lang)
    count = mail_count_in_folder(values, config_dir)
    if count is None:
        return CardStatus(False, t("card.mails.folder_missing", lang).format(folder=folder_name))
    if count == 0:
        return CardStatus(False, t("card.mails.folder_empty", lang).format(folder=folder_name))
    return CardStatus(True, t("card.mails.folder_ready", lang).format(folder=folder_name, count=count))


def fields_status(rules: list[FieldRule], previews: list[RulePreview] | None, lang: str) -> CardStatus:
    if not rules:
        return CardStatus(False, t("card.fields.none", lang))
    count = t("card.fields.count", lang).format(count=len(rules))
    if previews is None:
        return CardStatus(None, count)
    missing = [
        rule.column for rule, preview in zip(rules, previews, strict=True) if rule.required and not preview.found
    ]
    if missing:
        columns = ", ".join(_quote(column, lang) for column in missing)
        return CardStatus(False, t("card.fields.missing", lang).format(count=count, columns=columns))
    return CardStatus(True, t("card.fields.ok", lang).format(count=count))


def excel_status(path: Path, data_sheet: str, error_sheet: str, lang: str) -> CardStatus:
    if not path.exists():
        return CardStatus(None, t("card.excel.new", lang).format(file=path.name))
    try:
        workbook = load_workbook(path, read_only=True)
        try:
            rows = _data_rows(workbook, data_sheet)
            problems = _data_rows(workbook, error_sheet)
        finally:
            workbook.close()
    except Exception:  # locked, damaged or not an Excel file: the card is only informational
        return CardStatus(None, path.name)
    text = t("card.excel.rows", lang).format(file=path.name, rows=rows)
    if problems:
        text += t("card.excel.problems", lang).format(count=problems)
    return CardStatus(True, text)


def _data_rows(workbook, sheet_name: str) -> int:
    if sheet_name not in workbook.sheetnames:
        return 0
    return max((workbook[sheet_name].max_row or 1) - 1, 0)
