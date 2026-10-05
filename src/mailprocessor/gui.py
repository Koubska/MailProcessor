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
import threading
import time
import tomllib
from importlib import resources
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
    load_app_config,
    load_parsing_rules,
    resolve_relative_paths,
)
from mailprocessor.errors import (
    ImapLoginError,
    MailFolderNotFoundError,
    MissingPasswordError,
    SheetColumnsError,
    WorkbookLockedError,
)
from mailprocessor.i18n import resolve_language, t
from mailprocessor.logfile import attach_log_file, log_file_path
from mailprocessor.processor import RunSummary, run_pipeline, start_over
from mailprocessor.rule_patterns import LABEL_TYPES, RULE_TYPES, RuleType

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
    """Config value for a picked folder or file: relative if inside the config folder (keeps it portable), else absolute."""
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


def output_file_path(output_setting: str, config_dir: Path) -> Path:
    """The Excel file a run writes to; relative settings are relative to the config file, like in a run."""
    config = _default_config()
    config.app.output_xlsx = output_setting or config.app.output_xlsx
    return Path(resolve_relative_paths(config, config_dir).app.output_xlsx)


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
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk

    cfg_path = config_path or Path("config.toml")
    rules_file = rules_path or Path("parsing_rules.toml")
    current_language = resolve_language()

    def tr(key: str) -> str:
        return t(key, current_language)

    root = tk.Tk()
    root.title(tr("app.title"))
    try:
        with resources.as_file(resources.files("mailprocessor") / "assets" / "icon.png") as icon_path:
            app_icon = tk.PhotoImage(file=str(icon_path))
        root.iconphoto(True, app_icon)
    except (OSError, tk.TclError):
        pass  # a missing icon is cosmetic; never block the app
    root.geometry("1060x760")

    container = ttk.Frame(root, padding=12)
    container.pack(fill=tk.BOTH, expand=True)

    top_row = ttk.Frame(container)
    top_row.grid(row=0, column=0, sticky="ew")

    config_file_label = ttk.Label(top_row, text=f"{tr('label.config_file')}: {cfg_path}")
    config_file_label.pack(side=tk.LEFT)

    language_label = ttk.Label(top_row, text=tr("label.language"))
    language_label.pack(side=tk.RIGHT, padx=(12, 6))
    language_var = tk.StringVar(value=current_language)
    language_selector = ttk.Combobox(top_row, state="readonly", textvariable=language_var, values=["de", "en"], width=10)
    language_selector.pack(side=tk.RIGHT)

    notebook = ttk.Notebook(container)
    notebook.grid(row=1, column=0, sticky="nsew")

    config_tab = ttk.Frame(notebook, padding=8)
    fields_tab = ttk.Frame(notebook, padding=8)
    notebook.add(config_tab, text=tr("tab.config"))
    notebook.add(fields_tab, text=tr("tab.fields"))

    fields: list[FieldRule] = []

    form = ttk.Frame(config_tab)
    form.pack(fill=tk.BOTH, expand=True)

    log_level_var = tk.StringVar(value="INFO")
    sqlite_path_var = tk.StringVar()
    output_xlsx_var = tk.StringVar()
    sheet_data_var = tk.StringVar(value="daten")
    sheet_errors_var = tk.StringVar(value="fehler")
    # Not shown in the UI: the "Testlauf" and "Ausführen" buttons decide per run. Kept so the file value survives saving.
    dry_run_var = tk.BooleanVar(value=False)
    max_age_days_var = tk.StringVar(value="0")
    max_messages_var = tk.StringVar(value="0")
    source_type_var = tk.StringVar(value="eml")

    eml_folder_var = tk.StringVar()
    # Not shown in the UI; kept so a custom glob in config.toml survives saving.
    eml_glob_var = tk.StringVar(value="*.eml")

    imap_host_var = tk.StringVar()
    imap_port_var = tk.StringVar(value="993")
    imap_username_var = tk.StringVar()
    imap_password_var = tk.StringVar()
    imap_mailbox_var = tk.StringVar(value="INBOX")
    imap_use_ssl_var = tk.BooleanVar(value=True)
    imap_sender_filter_var = tk.StringVar()

    form_labels: dict[str, ttk.Label] = {}

    def add_row(parent, row: int, label_key: str, widget) -> None:
        label_widget = ttk.Label(parent, text=tr(label_key))
        label_widget.grid(row=row, column=0, sticky="w", padx=(0, 8), pady=(0, 6))
        form_labels[label_key] = label_widget
        widget.grid(row=row, column=1, sticky="ew", pady=(0, 6))

    # Main settings: where the mails come from and where the Excel file goes.
    add_row(
        form,
        0,
        "label.source_type",
        ttk.Combobox(form, textvariable=source_type_var, values=["eml", "imap"], state="readonly"),
    )

    source_frame = ttk.LabelFrame(form, text=tr("source.details"), padding=8)
    source_frame.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(0, 8))
    source_frame.columnconfigure(1, weight=1)

    output_row = ttk.Frame(form)
    ttk.Entry(output_row, textvariable=output_xlsx_var).pack(side=tk.LEFT, fill=tk.X, expand=True)

    def choose_output_file() -> None:
        current = output_file_path(output_xlsx_var.get().strip(), cfg_path.resolve().parent)
        chosen = filedialog.asksaveasfilename(
            parent=root,
            title=tr("dialog.choose_output_xlsx"),
            initialdir=str(current.parent if current.parent.is_dir() else cfg_path.resolve().parent),
            initialfile=current.name,
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
            # New rows are appended to an existing file; it is never replaced.
            confirmoverwrite=False,
        )
        if chosen:
            output_xlsx_var.set(path_setting(Path(chosen), cfg_path.resolve().parent))

    output_browse_button = ttk.Button(output_row, text=tr("button.browse"), command=choose_output_file)
    output_browse_button.pack(side=tk.LEFT, padx=(8, 0))
    add_row(form, 2, "label.output_xlsx", output_row)

    # Technical settings, collapsed by default.
    advanced_visible = tk.BooleanVar(value=False)
    advanced_toggle = ttk.Button(form, command=lambda: toggle_advanced())
    advanced_toggle.grid(row=3, column=0, columnspan=2, sticky="w", pady=(8, 6))
    advanced_frame = ttk.Frame(form)
    advanced_frame.columnconfigure(1, weight=1)
    add_row(
        advanced_frame,
        0,
        "label.log_level",
        ttk.Combobox(
            advanced_frame, textvariable=log_level_var, values=["DEBUG", "INFO", "WARNING", "ERROR"], state="readonly"
        ),
    )
    add_row(advanced_frame, 1, "label.sqlite_path", ttk.Entry(advanced_frame, textvariable=sqlite_path_var))
    add_row(advanced_frame, 2, "label.sheet_data", ttk.Entry(advanced_frame, textvariable=sheet_data_var))
    add_row(advanced_frame, 3, "label.sheet_errors", ttk.Entry(advanced_frame, textvariable=sheet_errors_var))
    add_row(advanced_frame, 4, "label.max_age_days", ttk.Entry(advanced_frame, textvariable=max_age_days_var))
    add_row(advanced_frame, 5, "label.max_messages", ttk.Entry(advanced_frame, textvariable=max_messages_var))
    advanced_buttons = ttk.Frame(advanced_frame)
    advanced_buttons.grid(row=6, column=1, sticky="w", pady=(4, 0))
    start_over_button = ttk.Button(advanced_buttons, command=lambda: on_start_over())
    start_over_button.pack(side=tk.LEFT)
    open_log_button = ttk.Button(advanced_buttons, command=lambda: open_log())
    open_log_button.pack(side=tk.LEFT, padx=(8, 0))

    def update_advanced_toggle() -> None:
        key = "button.advanced_hide" if advanced_visible.get() else "button.advanced_show"
        advanced_toggle.configure(text=tr(key))

    def toggle_advanced() -> None:
        advanced_visible.set(not advanced_visible.get())
        if advanced_visible.get():
            advanced_frame.grid(row=4, column=0, columnspan=2, sticky="ew")
        else:
            advanced_frame.grid_forget()
        update_advanced_toggle()

    eml_frame = ttk.Frame(source_frame)
    eml_folder_label = ttk.Label(eml_frame, text=tr("source.eml.folder"))
    eml_folder_label.grid(row=0, column=0, sticky="w", pady=(0, 6))
    ttk.Entry(eml_frame, textvariable=eml_folder_var).grid(row=0, column=1, sticky="ew", pady=(0, 6))

    def choose_eml_folder() -> None:
        current = Path(eml_folder_var.get().strip() or ".")
        if not current.is_absolute():
            current = cfg_path.resolve().parent / current
        chosen = filedialog.askdirectory(
            parent=root,
            title=tr("dialog.choose_eml_folder"),
            initialdir=str(current if current.is_dir() else cfg_path.resolve().parent),
            mustexist=True,
        )
        if chosen:
            eml_folder_var.set(path_setting(Path(chosen), cfg_path.resolve().parent))

    eml_browse_button = ttk.Button(eml_frame, text=tr("button.browse"), command=choose_eml_folder)
    eml_browse_button.grid(row=0, column=2, sticky="w", padx=(8, 0), pady=(0, 6))
    eml_frame.columnconfigure(1, weight=1)

    imap_frame = ttk.Frame(source_frame)
    imap_host_label = ttk.Label(imap_frame, text=tr("source.imap.host"))
    imap_host_label.grid(row=0, column=0, sticky="w", pady=(0, 6))
    ttk.Entry(imap_frame, textvariable=imap_host_var).grid(row=0, column=1, sticky="ew", pady=(0, 6))
    imap_port_label = ttk.Label(imap_frame, text=tr("source.imap.port"))
    imap_port_label.grid(row=1, column=0, sticky="w", pady=(0, 6))
    ttk.Entry(imap_frame, textvariable=imap_port_var).grid(row=1, column=1, sticky="ew", pady=(0, 6))
    imap_username_label = ttk.Label(imap_frame, text=tr("source.imap.username"))
    imap_username_label.grid(row=2, column=0, sticky="w", pady=(0, 6))
    ttk.Entry(imap_frame, textvariable=imap_username_var).grid(row=2, column=1, sticky="ew", pady=(0, 6))
    imap_password_label = ttk.Label(imap_frame, text=tr("source.imap.password"))
    imap_password_label.grid(row=3, column=0, sticky="w", pady=(0, 6))
    ttk.Entry(imap_frame, textvariable=imap_password_var, show="*").grid(row=3, column=1, sticky="ew", pady=(0, 6))
    imap_mailbox_label = ttk.Label(imap_frame, text=tr("source.imap.mailbox"))
    imap_mailbox_label.grid(row=4, column=0, sticky="w", pady=(0, 6))
    ttk.Entry(imap_frame, textvariable=imap_mailbox_var).grid(row=4, column=1, sticky="ew", pady=(0, 6))
    imap_sender_filter_label = ttk.Label(imap_frame, text=tr("source.imap.sender_filter"))
    imap_sender_filter_label.grid(row=5, column=0, sticky="w", pady=(0, 6))
    ttk.Entry(imap_frame, textvariable=imap_sender_filter_var).grid(row=5, column=1, sticky="ew", pady=(0, 6))
    imap_ssl_checkbox = ttk.Checkbutton(imap_frame, text=tr("source.imap.use_ssl"), variable=imap_use_ssl_var)
    imap_ssl_checkbox.grid(row=6, column=1, sticky="w")
    imap_frame.columnconfigure(1, weight=1)

    def toggle_source_frame() -> None:
        eml_frame.grid_forget()
        imap_frame.grid_forget()
        if source_type_var.get() == "imap":
            imap_frame.grid(row=0, column=0, sticky="ew")
        else:
            eml_frame.grid(row=0, column=0, sticky="ew")

    rules_file_label = ttk.Label(fields_tab, text=f"{tr('label.rules_file')}: {rules_file}")
    rules_file_label.pack(anchor="w")
    view_columns = ("column", "type", "description", "required")
    fields_view = ttk.Treeview(fields_tab, columns=view_columns, show="headings", height=10)
    view_layout = (("column", 160, "w"), ("type", 190, "w"), ("description", 480, "w"), ("required", 70, "center"))
    for name, width, anchor in view_layout:
        fields_view.column(name, width=width, anchor=anchor)
    fields_view.pack(fill=tk.BOTH, expand=True, pady=(6, 8))

    # Editor: column, type and required on top; below only the inputs the chosen type needs.
    editor = ttk.Frame(fields_tab)
    editor.pack(fill=tk.X, pady=(0, 8))
    field_column_var = tk.StringVar()
    field_labels_var = tk.StringVar()
    field_start_var = tk.StringVar()
    field_end_var = tk.StringVar()
    field_pattern_var = tk.StringVar()
    field_required_var = tk.BooleanVar(value=True)
    field_type: list[RuleType] = ["label"]  # current editor type; a list so nested functions can change it

    field_column_label = ttk.Label(editor, text=tr("label.field_column"))
    field_column_label.grid(row=0, column=0, sticky="w", padx=(0, 8))
    ttk.Entry(editor, textvariable=field_column_var, width=24).grid(row=0, column=1, sticky="w")
    field_type_label = ttk.Label(editor, text=tr("label.field_type"))
    field_type_label.grid(row=0, column=2, sticky="w", padx=(16, 8))
    field_type_box = ttk.Combobox(editor, state="readonly", width=28)
    field_type_box.grid(row=0, column=3, sticky="w")
    field_required_checkbox = ttk.Checkbutton(editor, text=tr("label.field_required"), variable=field_required_var)
    field_required_checkbox.grid(row=0, column=4, sticky="w", padx=(16, 0))

    inputs = ttk.Frame(editor)
    inputs.grid(row=1, column=0, columnspan=6, sticky="ew", pady=(8, 0))
    inputs.columnconfigure(1, weight=1)
    labels_label = ttk.Label(inputs)
    labels_entry = ttk.Entry(inputs, textvariable=field_labels_var)
    start_label = ttk.Label(inputs, text=tr("label.field_start"))
    start_entry = ttk.Entry(inputs, textvariable=field_start_var)
    end_label = ttk.Label(inputs, text=tr("label.field_end"))
    end_entry = ttk.Entry(inputs, textvariable=field_end_var)
    pattern_label = ttk.Label(inputs, text=tr("label.field_pattern"))
    pattern_entry = ttk.Entry(inputs, textvariable=field_pattern_var)
    rule_hint = ttk.Label(editor, foreground="gray", wraplength=900, justify=tk.LEFT)
    rule_hint.grid(row=2, column=0, columnspan=5, sticky="w", pady=(4, 0))
    as_regex_button = ttk.Button(editor, command=lambda: convert_to_regex())
    editor.columnconfigure(5, weight=1)

    def show_type_inputs() -> None:
        for widget in inputs.winfo_children():
            widget.grid_forget()
        rule_type = field_type[0]
        if rule_type in LABEL_TYPES:
            key = "label.field_labels_optional" if rule_type == "email" else "label.field_labels"
            labels_label.configure(text=tr(key))
            labels_label.grid(row=0, column=0, sticky="w", padx=(0, 8))
            labels_entry.grid(row=0, column=1, sticky="ew")
        elif rule_type == "between":
            start_label.grid(row=0, column=0, sticky="w", padx=(0, 8))
            start_entry.grid(row=0, column=1, sticky="ew")
            end_label.grid(row=1, column=0, sticky="w", padx=(0, 8), pady=(6, 0))
            end_entry.grid(row=1, column=1, sticky="ew", pady=(6, 0))
        else:
            pattern_label.grid(row=0, column=0, sticky="w", padx=(0, 8))
            pattern_entry.grid(row=0, column=1, sticky="ew")
        rule_hint.configure(text=tr(f"rule.hint.{rule_type}"))
        if rule_type == "regex":
            as_regex_button.grid_forget()
        else:
            as_regex_button.grid(row=2, column=5, sticky="e", pady=(4, 0))

    def set_field_type(rule_type: RuleType) -> None:
        field_type[0] = rule_type
        field_type_box.current(RULE_TYPES.index(rule_type))
        show_type_inputs()

    def on_type_selected() -> None:
        field_type[0] = RULE_TYPES[field_type_box.current()]
        show_type_inputs()

    status_var = tk.StringVar(value=tr("label.status.ready"))
    output_frame = ttk.Frame(container)
    output_frame.grid(row=3, column=0, sticky="nsew", pady=(12, 0))
    output = tk.Text(output_frame, height=10, wrap="word")
    output_scrollbar = ttk.Scrollbar(output_frame, orient=tk.VERTICAL, command=output.yview)
    output.configure(yscrollcommand=output_scrollbar.set, state=tk.DISABLED)
    output.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
    output_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
    status_row = ttk.Frame(container)
    status_row.grid(row=4, column=0, sticky="ew", pady=(8, 0))
    ttk.Label(status_row, textvariable=status_var).pack(side=tk.LEFT, fill=tk.X, expand=True)
    stop_button = ttk.Button(status_row, text=tr("button.stop"), state=tk.DISABLED, command=lambda: stop_run())
    stop_button.pack(side=tk.RIGHT, padx=(8, 0))
    progress_bar = ttk.Progressbar(status_row, mode="determinate", length=220)
    progress_bar.pack(side=tk.RIGHT)

    def append_output(text: str) -> None:
        output.configure(state=tk.NORMAL)
        output.insert(tk.END, text + "\n")
        output.see(tk.END)
        output.configure(state=tk.DISABLED)

    def refresh_fields_view() -> None:
        for item_id in fields_view.get_children():
            fields_view.delete(item_id)
        for index, field in enumerate(fields):
            fields_view.insert(
                "",
                tk.END,
                iid=str(index),
                values=(
                    field.column,
                    tr(f"rule.type.{field.type}"),
                    describe_rule(field, current_language),
                    tr("view.field.required_yes") if field.required else tr("view.field.required_no"),
                ),
            )

    def apply_language(selected_language: str) -> None:
        nonlocal current_language
        current_language = selected_language
        root.title(tr("app.title"))
        language_label.configure(text=tr("label.language"))
        config_file_label.configure(text=f"{tr('label.config_file')}: {cfg_path}")
        rules_file_label.configure(text=f"{tr('label.rules_file')}: {rules_file}")
        notebook.tab(0, text=tr("tab.config"))
        notebook.tab(1, text=tr("tab.fields"))
        for key, label_widget in form_labels.items():
            label_widget.configure(text=tr(key))
        update_advanced_toggle()
        output_browse_button.configure(text=tr("button.browse"))
        source_frame.configure(text=tr("source.details"))
        eml_folder_label.configure(text=tr("source.eml.folder"))
        eml_browse_button.configure(text=tr("button.browse"))
        imap_host_label.configure(text=tr("source.imap.host"))
        imap_port_label.configure(text=tr("source.imap.port"))
        imap_username_label.configure(text=tr("source.imap.username"))
        imap_password_label.configure(text=tr("source.imap.password"))
        imap_mailbox_label.configure(text=tr("source.imap.mailbox"))
        imap_sender_filter_label.configure(text=tr("source.imap.sender_filter"))
        imap_ssl_checkbox.configure(text=tr("source.imap.use_ssl"))
        for name in view_columns:
            fields_view.heading(name, text=tr(f"view.field.{name}"))
        field_column_label.configure(text=tr("label.field_column"))
        field_type_label.configure(text=tr("label.field_type"))
        field_type_box.configure(values=[tr(f"rule.type.{rule_type}") for rule_type in RULE_TYPES])
        start_label.configure(text=tr("label.field_start"))
        end_label.configure(text=tr("label.field_end"))
        pattern_label.configure(text=tr("label.field_pattern"))
        as_regex_button.configure(text=tr("button.as_regex"))
        set_field_type(field_type[0])
        field_required_checkbox.configure(text=tr("label.field_required"))
        save_button.configure(text=tr("button.save"))
        reload_button.configure(text=tr("button.reload"))
        run_button.configure(text=tr("button.run"))
        test_run_button.configure(text=tr("button.test_run"))
        stop_button.configure(text=tr("button.stop"))
        start_over_button.configure(text=tr("button.start_over"))
        open_log_button.configure(text=tr("button.open_log"))
        open_excel_button.configure(text=tr("button.open_excel"))
        add_update_button.configure(text=tr("button.field_add_update"))
        remove_button.configure(text=tr("button.field_remove"))
        up_button.configure(text=tr("button.field_up"))
        down_button.configure(text=tr("button.field_down"))
        if status_var.get() in {t("label.status.ready", "de"), t("label.status.ready", "en")}:
            status_var.set(tr("label.status.ready"))
        refresh_fields_view()

    def config_to_form(config: AppConfig) -> None:
        log_level_var.set(config.app.log_level)
        sqlite_path_var.set(config.app.sqlite_path)
        output_xlsx_var.set(config.app.output_xlsx)
        sheet_data_var.set(config.app.sheet_data)
        sheet_errors_var.set(config.app.sheet_errors)
        dry_run_var.set(config.app.dry_run)
        max_age_days_var.set(str(config.app.max_age_days))
        max_messages_var.set(str(config.app.max_messages))
        source_type_var.set(config.source.type)

        eml = config.source.eml or EmlSourceConfig(folder="", glob="*.eml")
        eml_folder_var.set(eml.folder)
        eml_glob_var.set(eml.glob)

        imap = config.source.imap or ImapSourceConfig(
            host="",
            port=993,
            username="",
            mailbox="INBOX",
            use_ssl=True,
            sender_filter=None,
        )
        imap_host_var.set(imap.host)
        imap_port_var.set(str(imap.port))
        imap_username_var.set(imap.username)
        imap_password_var.set(imap.password or "")
        imap_mailbox_var.set(imap.mailbox)
        imap_use_ssl_var.set(imap.use_ssl)
        imap_sender_filter_var.set(imap.sender_filter or "")
        toggle_source_frame()

    def form_to_config() -> AppConfig:
        max_age_days = int(max_age_days_var.get().strip())
        max_messages = int(max_messages_var.get().strip())
        app_section = AppSection(
            log_level=log_level_var.get().strip(),
            sqlite_path=sqlite_path_var.get().strip(),
            output_xlsx=output_xlsx_var.get().strip(),
            sheet_data=sheet_data_var.get().strip(),
            sheet_errors=sheet_errors_var.get().strip(),
            dry_run=dry_run_var.get(),
            max_messages=max_messages,
            max_age_days=max_age_days,
        )
        source_type = source_type_var.get().strip()
        # Keep the settings of the inactive source too, so switching type and saving loses nothing.
        eml = None
        if source_type == "eml" or eml_folder_var.get().strip():
            eml = EmlSourceConfig(folder=eml_folder_var.get().strip(), glob=eml_glob_var.get().strip() or "*.eml")
        imap = None
        if source_type == "imap" or imap_host_var.get().strip():
            sender_filter = imap_sender_filter_var.get().strip()
            imap = ImapSourceConfig(
                host=imap_host_var.get().strip(),
                port=int(imap_port_var.get().strip()),
                username=imap_username_var.get().strip(),
                mailbox=imap_mailbox_var.get().strip(),
                use_ssl=imap_use_ssl_var.get(),
                sender_filter=sender_filter or None,
            )
        return AppConfig(app=app_section, source=SourceConfig(type=source_type, eml=eml, imap=imap))

    def load_from_disk() -> None:
        nonlocal fields
        config_text = _read_text_or_empty(cfg_path)
        if config_text.strip():
            config_to_form(parse_config_text(config_text))
        else:
            config_to_form(_default_config())

        rules_text = _read_text_or_empty(rules_file)
        fields = parse_rules_text(rules_text) if rules_text.strip() else []
        refresh_fields_view()

    def save_to_disk() -> None:
        cfg_path.write_text(render_config_text(form_to_config()), encoding="utf-8")
        rules_file.write_text(render_rules_text(fields), encoding="utf-8")

    def selected_index() -> int | None:
        selection = fields_view.selection()
        return int(selection[0]) if selection else None

    def load_selected_field() -> None:
        index = selected_index()
        if index is None:
            return
        selected = fields[index]
        field_column_var.set(selected.column)
        field_labels_var.set("; ".join(selected.label))
        field_start_var.set(selected.start or "")
        field_end_var.set(selected.end or "")
        field_pattern_var.set(selected.pattern or "")
        field_required_var.set(selected.required)
        set_field_type(selected.type)

    def rule_from_editor() -> FieldRule:
        index = selected_index()
        return rule_from_inputs(
            column=field_column_var.get(),
            rule_type=field_type[0],
            labels_text=field_labels_var.get(),
            start=field_start_var.get(),
            end=field_end_var.get(),
            pattern=field_pattern_var.get(),
            required=field_required_var.get(),
            other_columns=[field.column for position, field in enumerate(fields) if position != index],
            lang=current_language,
        )

    def convert_to_regex() -> None:
        """Turn the rule in the editor into a hand-editable pattern (one way: regex is not turned back)."""
        try:
            rule = rule_from_editor()
        except ValueError as exc:
            messagebox.showerror(tr("app.title"), str(exc))
            return
        field_pattern_var.set(rule.regex)
        set_field_type("regex")

    def upsert_current_field() -> None:
        candidate = rule_from_editor()
        index = selected_index()
        updated = [*fields, candidate] if index is None else [*fields[:index], candidate, *fields[index + 1 :]]
        # Same validation as loading the rules file (e.g. the reserved column name).
        ParsingRules(fields=updated)
        fields[:] = updated
        refresh_fields_view()
        fields_view.selection_set(str(updated.index(candidate)))

    def add_field() -> None:
        try:
            upsert_current_field()
            status_var.set(tr("label.status.field_updated"))
        except ValueError as exc:
            messagebox.showerror(tr("app.title"), str(exc))

    def remove_field() -> None:
        index = selected_index()
        if index is None:
            messagebox.showerror(tr("app.title"), tr("error.select_field_remove"))
            return
        fields.pop(index)
        refresh_fields_view()
        status_var.set(tr("label.status.field_removed"))

    def move_field(offset: int) -> None:
        index = selected_index()
        if index is None:
            messagebox.showerror(tr("app.title"), tr("error.select_field_move"))
            return
        new_index = index + offset
        if new_index < 0 or new_index >= len(fields):
            return
        fields[index], fields[new_index] = fields[new_index], fields[index]
        refresh_fields_view()
        fields_view.selection_set(str(new_index))
        load_selected_field()
        status_var.set(tr("label.status.field_reordered"))

    def on_save() -> None:
        try:
            save_to_disk()
            status_var.set(tr("label.status.saved"))
        except (OSError, ValueError) as exc:
            status_var.set(tr("label.status.save_failed"))
            messagebox.showerror(tr("app.title"), str(exc))

    def on_reload() -> None:
        try:
            load_from_disk()
            status_var.set(tr("label.status.reload"))
        except ValueError as exc:
            status_var.set(tr("label.status.reload_failed"))
            messagebox.showerror(tr("app.title"), str(exc))

    run_results: queue.Queue = queue.Queue()
    # Pipeline log records are shown in the output box; the level follows log_level from the config.
    package_logger = logging.getLogger("mailprocessor")
    log_handler = QueueLogHandler(run_results)
    package_logger.addHandler(log_handler)

    def open_output() -> None:
        path = output_file_path(output_xlsx_var.get().strip(), cfg_path.resolve().parent)
        if not path.exists():
            messagebox.showinfo(tr("app.title"), tr("info.no_excel_yet").format(file=path))
            return
        try:
            open_in_default_app(path)
        except OSError as exc:
            messagebox.showerror(tr("app.title"), tr("error.open_excel").format(file=path, error=exc))

    def open_log() -> None:
        path = log_file_path(cfg_path.resolve().parent)
        if not path.exists():
            messagebox.showinfo(tr("app.title"), tr("info.no_log_yet").format(file=path))
            return
        try:
            open_in_default_app(path)
        except OSError as exc:
            messagebox.showerror(tr("app.title"), tr("error.open_file").format(file=path, error=exc))

    cancel_event = threading.Event()

    def set_busy(busy: bool) -> None:
        state = tk.DISABLED if busy else tk.NORMAL
        for button in (run_button, test_run_button, save_button, reload_button, start_over_button):
            button.configure(state=state)
        stop_button.configure(state=tk.NORMAL if busy else tk.DISABLED)
        progress_bar.configure(value=0, maximum=1)

    def stop_run() -> None:
        cancel_event.set()
        stop_button.configure(state=tk.DISABLED)
        status_var.set(tr("label.status.stopping"))

    def show_progress(current: int, total: int) -> None:
        progress_bar.configure(maximum=max(total, 1), value=current)
        if not cancel_event.is_set():
            status_var.set(tr("label.status.progress").format(current=current, total=total))

    def run_worker(app_config: AppConfig, parsing_rules: ParsingRules) -> None:
        # Runs off the UI thread; must not touch tkinter objects.
        def report_progress(current: int, total: int) -> None:
            run_results.put(("progress", (current, total)))

        try:
            summary = run_pipeline(app_config, parsing_rules, progress=report_progress, cancel=cancel_event)
            run_results.put(("ok", (app_config, summary)))
        except Exception as exc:
            if not isinstance(exc, (ValueError, OSError)):
                package_logger.debug("Unexpected error", exc_info=True)
            run_results.put(("error", exc))

    def poll_run_result() -> None:
        # Show all queued log lines and progress; the run's result ("ok"/"error") is always the last item.
        while True:
            try:
                kind, payload = run_results.get_nowait()
            except queue.Empty:
                root.after(100, poll_run_result)
                return
            if kind == "log":
                append_output(payload)
            elif kind == "progress":
                show_progress(*payload)
            else:
                break
        set_busy(False)
        if kind == "ok":
            app_config, summary = payload
            text = run_summary_text(
                summary,
                Path(app_config.app.output_xlsx).name,
                app_config.app.sheet_errors,
                app_config.app.dry_run,
                current_language,
            )
            append_output(text)
            status_var.set(text)
        else:
            message = friendly_error(payload, current_language)
            append_output(f"ERROR   {message}")
            status_var.set(tr("label.status.run_failed"))
            messagebox.showerror(tr("app.title"), message)

    def execute_run(dry_run: bool) -> None:
        try:
            if not fields:
                raise ValueError("At least one parsing field is required")
            save_to_disk()
            app_config = load_app_config(cfg_path)
            parsing_rules = load_parsing_rules(rules_file)
            if app_config.source.imap is not None:
                app_config.source.imap.password = imap_password_var.get() or None
            app_config.app.dry_run = dry_run
        except (ValueError, OSError) as exc:
            status_var.set(tr("label.status.run_failed"))
            messagebox.showerror(tr("app.title"), friendly_error(exc, current_language))
            return
        package_logger.setLevel(app_config.app.log_level)
        attach_log_file(cfg_path.resolve().parent)
        append_output(f"--- {time.strftime('%H:%M:%S')} ---")
        cancel_event.clear()
        set_busy(True)
        status_var.set(tr("label.status.running"))
        threading.Thread(target=run_worker, args=(app_config, parsing_rules), daemon=True).start()
        root.after(100, poll_run_result)

    def on_start_over() -> None:
        if not messagebox.askyesno(tr("app.title"), tr("confirm.start_over"), icon="warning"):
            return
        try:
            save_to_disk()
            app_config = load_app_config(cfg_path)
            attach_log_file(cfg_path.resolve().parent)
            backup = start_over(app_config)
        except (ValueError, OSError) as exc:
            messagebox.showerror(tr("app.title"), friendly_error(exc, current_language))
            return
        if backup is not None:
            append_output(tr("info.backup_created").format(file=backup.name))
        execute_run(dry_run=False)

    button_row = ttk.Frame(container)
    button_row.grid(row=2, column=0, sticky="ew", pady=(12, 0))
    save_button = ttk.Button(button_row, text=tr("button.save"), command=on_save)
    save_button.pack(side="left")
    reload_button = ttk.Button(button_row, text=tr("button.reload"), command=on_reload)
    reload_button.pack(side="left", padx=(8, 0))
    run_button = ttk.Button(button_row, text=tr("button.run"), command=lambda: execute_run(dry_run=False))
    run_button.pack(side="right")
    test_run_button = ttk.Button(button_row, text=tr("button.test_run"), command=lambda: execute_run(dry_run=True))
    test_run_button.pack(side="right", padx=(0, 8))
    open_excel_button = ttk.Button(button_row, text=tr("button.open_excel"), command=open_output)
    open_excel_button.pack(side="right", padx=(0, 8))

    field_buttons = ttk.Frame(fields_tab)
    field_buttons.pack(fill=tk.X)
    add_update_button = ttk.Button(field_buttons, text=tr("button.field_add_update"), command=add_field)
    add_update_button.pack(side=tk.LEFT)
    remove_button = ttk.Button(field_buttons, text=tr("button.field_remove"), command=remove_field)
    remove_button.pack(side=tk.LEFT, padx=(8, 0))
    up_button = ttk.Button(field_buttons, text=tr("button.field_up"), command=lambda: move_field(-1))
    up_button.pack(side=tk.LEFT, padx=(8, 0))
    down_button = ttk.Button(field_buttons, text=tr("button.field_down"), command=lambda: move_field(1))
    down_button.pack(side=tk.LEFT, padx=(8, 0))
    fields_view.bind("<<TreeviewSelect>>", lambda _event: load_selected_field())
    field_type_box.bind("<<ComboboxSelected>>", lambda _event: on_type_selected())
    source_type_var.trace_add("write", lambda *_args: toggle_source_frame())
    language_selector.bind("<<ComboboxSelected>>", lambda _event: apply_language(language_var.get()))

    form.columnconfigure(1, weight=1)
    container.columnconfigure(0, weight=1)
    container.rowconfigure(1, weight=4)
    container.rowconfigure(3, weight=2)

    apply_language(language_var.get())
    try:
        load_from_disk()
    except (OSError, ValueError) as exc:
        # A broken config must not prevent the editor from opening; show defaults and the problem.
        config_to_form(_default_config())
        status_var.set(tr("label.status.reload_failed"))
        # Bind the message now: `exc` is unbound once the except block ends.
        root.after(0, lambda message=str(exc): messagebox.showerror(tr("app.title"), message))
    try:
        root.mainloop()
    finally:
        package_logger.removeHandler(log_handler)
