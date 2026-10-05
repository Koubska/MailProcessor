"""The desktop window. Imported lazily by `gui.launch_gui`, so the CLI and the tests never need tkinter.

Layout: Start (status and run) · E-Mails (source) · Felder (rules with live test) · Einstellungen.
Every change is saved automatically as soon as it is valid; invalid inputs are marked next to the field.
"""

from __future__ import annotations

import logging
import queue
import threading
import time
import tkinter as tk
from collections.abc import Callable
from importlib import resources
from pathlib import Path
from tkinter import filedialog, font, messagebox, ttk

from mailprocessor.config import FieldRule, ParsingRules, load_app_config
from mailprocessor.gui import (
    QueueLogHandler,
    _default_config,
    _read_text_or_empty,
    describe_rule,
    friendly_error,
    open_in_default_app,
    output_file_path,
    parse_config_text,
    parse_rules_text,
    path_setting,
    render_config_text,
    render_rules_text,
    rule_from_inputs,
    run_summary_text,
    setting_path,
)
from mailprocessor.i18n import resolve_language, t
from mailprocessor.logfile import attach_log_file, log_file_path
from mailprocessor.parser import normalize_body
from mailprocessor.preview import (
    RulePreview,
    SampleMail,
    load_sample,
    preview_rule,
    result_text,
    sample_files,
    summary_text,
)
from mailprocessor.processor import run_pipeline, start_over
from mailprocessor.rule_patterns import LABEL_TYPES, RULE_TYPES, RuleType
from mailprocessor.sources.imap_source import check_imap_connection
from mailprocessor.ui_model import (
    DEFAULT_IMAP_PORTS,
    CardStatus,
    config_from_form,
    excel_status,
    fields_status,
    form_values,
    mail_count_in_folder,
    mails_status,
)

GREEN, RED, AMBER, GRAY = "#15803d", "#b91c1c", "#b45309", "gray"
HIGHLIGHT = "#fde68a"
LANGUAGES = {"de": "Deutsch", "en": "English"}
# Which tab holds which input, to jump to the first problem.
INPUT_TABS = {
    **dict.fromkeys(("eml_folder", "imap_host", "imap_port", "imap_username"), 1),
    **dict.fromkeys(("output_xlsx", "sheet_data", "sheet_errors", "sqlite_path", "max_age_days", "max_messages"), 3),
}


class App:
    def __init__(self, root: tk.Tk, config_path: Path, rules_path: Path) -> None:
        self.root = root
        self.cfg_path = config_path
        self.rules_path = rules_path
        self.config_dir = config_path.resolve().parent
        self.lang = resolve_language()

        self.form: dict[str, tk.Variable] = {}
        self.error_labels: dict[str, ttk.Label] = {}
        self.translated: list[tuple[tk.Widget, str, str]] = []
        self.fields: list[FieldRule] = []
        self.saved_config_text = ""
        self.form_valid = True
        self.save_job: str | None = None
        self.preview_job: str | None = None
        self.apply_job: str | None = None
        self.loading = True  # no autosave while the window is built and filled
        self.loading_editor = False
        self.editing_index: int | None = None
        self.running = False

        self.password = tk.StringVar()
        self.run_results: queue.Queue = queue.Queue()
        self.cancel_event = threading.Event()
        self.package_logger = logging.getLogger("mailprocessor")
        self.log_handler = QueueLogHandler(self.run_results)
        self.package_logger.addHandler(self.log_handler)

        self.sample: dict = {"files": [], "index": -1, "title": "", "header_text": ""}
        self.previews: list[RulePreview] | None = None
        self.editor_mode = "edit"  # "edit": changes apply to the selected field; "new": "Hinzufügen" adds it
        self.field_type: RuleType = "label"

        self._setup_window()
        self._setup_styles()
        self._build()
        self._load()
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

    # ------------------------------------------------------------------ basics

    def tr(self, key: str) -> str:
        return t(key, self.lang)

    def _translate(self, widget: tk.Widget, key: str, option: str = "text") -> tk.Widget:
        """Set a widget's text and remember it, so a language switch can update it."""
        widget.configure(**{option: self.tr(key)})
        self.translated.append((widget, option, key))
        return widget

    def label(self, parent, key: str, style: str | None = None, **options) -> ttk.Label:
        widget = ttk.Label(parent, style=style or "TLabel", **options)
        return self._translate(widget, key)

    def hint(self, parent, key: str, wrap: int = 560) -> ttk.Label:
        return self.label(parent, key, style="Hint.TLabel", wraplength=wrap, justify=tk.LEFT)

    def button(self, parent, key: str, command: Callable[[], None], style: str | None = None, **options) -> ttk.Button:
        widget = ttk.Button(parent, command=command, style=style or "TButton", **options)
        return self._translate(widget, key)

    def var(self, key: str, value: str | bool) -> tk.Variable:
        variable = tk.BooleanVar(value=value) if isinstance(value, bool) else tk.StringVar(value=value)
        variable.trace_add("write", lambda *_args: self.on_form_changed())
        self.form[key] = variable
        return variable

    def entry(self, parent, key: str, width: int = 40, **options) -> ttk.Entry:
        return ttk.Entry(parent, textvariable=self.form[key], width=width, **options)

    def error_label(self, parent, key: str) -> ttk.Label:
        widget = ttk.Label(parent, style="Error.TLabel", wraplength=520, justify=tk.LEFT)
        self.error_labels[key] = widget
        return widget

    def _setup_window(self) -> None:
        self.root.title(self.tr("app.title"))
        try:
            with resources.as_file(resources.files("mailprocessor") / "assets" / "icon.png") as icon_path:
                self.app_icon = tk.PhotoImage(file=str(icon_path))
            self.root.iconphoto(True, self.app_icon)
        except (OSError, tk.TclError):
            pass  # a missing icon is cosmetic
        width = min(1240, self.root.winfo_screenwidth() - 80)
        height = min(860, self.root.winfo_screenheight() - 120)
        self.root.geometry(f"{width}x{height}")
        self.root.minsize(min(900, width), min(620, height))

    def _setup_styles(self) -> None:
        style = ttk.Style(self.root)
        base = font.nametofont("TkDefaultFont")
        size = abs(int(base.cget("size"))) or 10
        self.title_font = base.copy()
        self.title_font.configure(size=size + 7, weight="bold")
        self.heading_font = base.copy()
        self.heading_font.configure(size=size + 2, weight="bold")
        self.bold_font = base.copy()
        self.bold_font.configure(weight="bold")
        style.configure("Title.TLabel", font=self.title_font)
        style.configure("Heading.TLabel", font=self.heading_font)
        style.configure("Bold.TLabel", font=self.bold_font)
        style.configure("Hint.TLabel", foreground=GRAY)
        style.configure("Error.TLabel", foreground=RED)
        style.configure("Big.TButton", font=self.heading_font, padding=(24, 10))
        style.configure("CardIcon.TLabel", font=self.heading_font)
        style.configure("Heading.TRadiobutton", font=self.heading_font)
        style.configure("TNotebook.Tab", padding=(14, 6))

    # ------------------------------------------------------------------ layout

    def _build(self) -> None:
        outer = ttk.Frame(self.root, padding=(16, 12, 16, 8))
        outer.pack(fill=tk.BOTH, expand=True)
        self.notebook = ttk.Notebook(outer)
        self.notebook.pack(fill=tk.BOTH, expand=True)
        self.tabs = [ttk.Frame(self.notebook, padding=16) for _ in range(4)]
        for tab in self.tabs:
            self.notebook.add(tab)
        self.notebook.bind("<<NotebookTabChanged>>", lambda _event: self.on_tab_changed())

        footer = ttk.Frame(outer)
        footer.pack(fill=tk.X, pady=(8, 0))
        self.save_status = ttk.Label(footer, style="Hint.TLabel")
        self.save_status.pack(side=tk.LEFT)

        self._build_start_tab(self.tabs[0])
        self._build_mails_tab(self.tabs[1])
        self._build_fields_tab(self.tabs[2])
        self._build_settings_tab(self.tabs[3])
        self._update_tab_titles()

    def _update_tab_titles(self) -> None:
        for index, key in enumerate(("tab.start", "tab.mails", "tab.fields", "tab.settings")):
            self.notebook.tab(index, text=self.tr(key))

    # --- Start

    def _build_start_tab(self, tab: ttk.Frame) -> None:
        self.label(tab, "start.title", style="Title.TLabel").pack(anchor="w")
        self.hint(tab, "start.intro", wrap=760).pack(anchor="w", pady=(4, 16))

        cards = ttk.Frame(tab)
        cards.pack(fill=tk.X)
        cards.columnconfigure(2, weight=1)
        self.cards: dict[str, tuple[ttk.Label, ttk.Label]] = {}
        card_rows = (
            ("mails", "card.mails.title", "button.change", lambda: self.notebook.select(1)),
            ("fields", "card.fields.title", "button.edit", lambda: self.notebook.select(2)),
            ("excel", "card.excel.title", "button.open_excel", self.open_output),
        )
        for row, (name, title_key, action_key, action) in enumerate(card_rows):
            icon = ttk.Label(cards, style="CardIcon.TLabel", width=2)
            icon.grid(row=row, column=0, sticky="w", pady=6)
            self.label(cards, title_key, style="Bold.TLabel").grid(row=row, column=1, sticky="w", padx=(4, 16))
            text = ttk.Label(cards, wraplength=640, justify=tk.LEFT)
            text.grid(row=row, column=2, sticky="w")
            self.button(cards, action_key, action).grid(row=row, column=3, sticky="e", padx=(12, 0))
            self.cards[name] = (icon, text)

        ttk.Separator(tab).pack(fill=tk.X, pady=16)
        actions = ttk.Frame(tab)
        actions.pack(fill=tk.X)
        self.run_button = self.button(
            actions, "button.run", lambda: self.execute_run(dry_run=False), style="Big.TButton"
        )
        self.run_button.pack(side=tk.LEFT)
        self.test_run_button = self.button(actions, "button.test_run", lambda: self.execute_run(dry_run=True))
        self.test_run_button.pack(side=tk.LEFT, padx=(12, 0))
        self.hint(tab, "start.test_run_hint", wrap=760).pack(anchor="w", pady=(6, 0))

        self.progress_row = ttk.Frame(tab)
        self.progress_bar = ttk.Progressbar(self.progress_row, mode="determinate", length=320)
        self.progress_bar.pack(side=tk.LEFT)
        self.progress_label = ttk.Label(self.progress_row)
        self.progress_label.pack(side=tk.LEFT, padx=(12, 0))
        self.stop_button = self.button(self.progress_row, "button.stop", self.stop_run)
        self.stop_button.pack(side=tk.LEFT, padx=(12, 0))

        self.result_label = ttk.Label(tab, wraplength=900, justify=tk.LEFT, font=self.bold_font)
        self.result_label.pack(anchor="w", pady=(16, 0))

        details_row = ttk.Frame(tab)
        details_row.pack(fill=tk.X, pady=(12, 0))
        self.details_button = ttk.Button(details_row, command=self.toggle_details)
        self.details_button.pack(side=tk.LEFT)
        self.button(details_row, "button.open_log", self.open_log).pack(side=tk.LEFT, padx=(8, 0))
        self.details_frame = ttk.Frame(tab)
        self.log_text = tk.Text(self.details_frame, height=10, wrap="word", state=tk.DISABLED)
        log_scroll = ttk.Scrollbar(self.details_frame, orient=tk.VERTICAL, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=log_scroll.set)
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        log_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.details_visible = False
        self._update_details_button()

    # --- E-Mails

    def _build_mails_tab(self, tab: ttk.Frame) -> None:
        self.label(tab, "mails.title", style="Title.TLabel").pack(anchor="w")
        self.hint(tab, "mails.intro", wrap=760).pack(anchor="w", pady=(4, 12))
        source = self.var("source_type", "eml")

        eml_choice = ttk.Radiobutton(tab, variable=source, value="eml", style="Heading.TRadiobutton")
        self._translate(eml_choice, "mails.eml.choice").pack(anchor="w")
        self.eml_frame = ttk.Frame(tab, padding=(28, 4, 0, 12))
        self.eml_frame.pack(fill=tk.X)
        self.hint(self.eml_frame, "mails.eml.hint").grid(row=0, column=0, columnspan=4, sticky="w")
        self.var("eml_glob", "*.eml")  # not shown; kept so a custom pattern in config.toml survives
        self.var("eml_folder", "")
        self.entry(self.eml_frame, "eml_folder", width=50).grid(row=1, column=0, sticky="ew", pady=(6, 0))
        self.button(self.eml_frame, "button.browse", self.choose_eml_folder).grid(
            row=1, column=1, padx=(8, 0), pady=(6, 0)
        )
        self.button(self.eml_frame, "button.open_folder", self.open_eml_folder).grid(
            row=1, column=2, padx=(8, 0), pady=(6, 0)
        )
        self.eml_count = ttk.Label(self.eml_frame)
        self.eml_count.grid(row=2, column=0, columnspan=4, sticky="w", pady=(4, 0))
        self.error_label(self.eml_frame, "eml_folder").grid(row=3, column=0, columnspan=4, sticky="w")
        self.eml_frame.columnconfigure(0, weight=1)

        imap_choice = ttk.Radiobutton(tab, variable=source, value="imap", style="Heading.TRadiobutton")
        self._translate(imap_choice, "mails.imap.choice").pack(anchor="w", pady=(8, 0))
        self.imap_frame = ttk.Frame(tab, padding=(28, 4, 0, 0))
        self.imap_frame.pack(fill=tk.X)
        self.hint(self.imap_frame, "mails.imap.hint").grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 6))
        rows = (
            ("imap_host", "mails.imap.host", "mails.imap.host_hint"),
            ("imap_username", "mails.imap.username", "mails.imap.username_hint"),
            (None, "mails.imap.password", "mails.imap.password_hint"),
            ("imap_mailbox", "mails.imap.mailbox", "mails.imap.mailbox_hint"),
            ("imap_sender_filter", "mails.imap.sender_filter", "mails.imap.sender_filter_hint"),
        )
        for key, default in (
            ("imap_host", ""),
            ("imap_username", ""),
            ("imap_mailbox", "INBOX"),
            ("imap_sender_filter", ""),
        ):
            self.var(key, default)
        self.var("imap_port", "993")
        self.var("imap_use_ssl", True)
        for row, (key, label_key, hint_key) in enumerate(rows, start=1):
            self.label(self.imap_frame, label_key).grid(
                row=row * 2 - 1, column=0, sticky="w", padx=(0, 12), pady=(4, 0)
            )
            if key is None:
                self.password_entry = ttk.Entry(self.imap_frame, textvariable=self.password, show="•", width=40)
                widget = self.password_entry
            else:
                widget = self.entry(self.imap_frame, key)
            widget.grid(row=row * 2 - 1, column=1, sticky="ew", pady=(4, 0))
            self.hint(self.imap_frame, hint_key, wrap=440).grid(row=row * 2 - 1, column=2, sticky="w", padx=(12, 0))
            if key is not None:
                self.error_label(self.imap_frame, key).grid(row=row * 2, column=1, columnspan=2, sticky="w")
        security = ttk.Frame(self.imap_frame)
        security.grid(row=12, column=1, columnspan=2, sticky="w", pady=(8, 0))
        ssl_box = ttk.Checkbutton(security, variable=self.form["imap_use_ssl"], command=self.on_ssl_toggled)
        self._translate(ssl_box, "mails.imap.ssl").pack(side=tk.LEFT)
        self.label(security, "mails.imap.port").pack(side=tk.LEFT, padx=(16, 6))
        self.entry(security, "imap_port", width=7).pack(side=tk.LEFT)
        self.error_label(self.imap_frame, "imap_port").grid(row=13, column=1, columnspan=2, sticky="w")
        test_row = ttk.Frame(self.imap_frame)
        test_row.grid(row=14, column=1, columnspan=2, sticky="w", pady=(12, 0))
        self.connection_button = self.button(test_row, "button.test_connection", self.test_connection)
        self.connection_button.pack(side=tk.LEFT)
        self.connection_result = ttk.Label(test_row, wraplength=560, justify=tk.LEFT)
        self.connection_result.pack(side=tk.LEFT, padx=(12, 0))
        self.imap_frame.columnconfigure(1, weight=1)
        self.password.trace_add("write", lambda *_args: self.refresh_cards())

    # --- Felder

    def _build_fields_tab(self, tab: ttk.Frame) -> None:
        self.label(tab, "fields.title", style="Title.TLabel").pack(anchor="w")
        self.hint(tab, "fields.intro", wrap=900).pack(anchor="w", pady=(4, 10))
        paned = ttk.PanedWindow(tab, orient=tk.HORIZONTAL)
        paned.pack(fill=tk.BOTH, expand=True)
        rules_pane = ttk.Frame(paned, padding=(0, 0, 12, 0))
        sample_pane = ttk.Frame(paned, padding=(12, 0, 0, 0))
        paned.add(rules_pane, weight=3)
        paned.add(sample_pane, weight=2)

        columns = ("column", "type", "description", "required", "result")
        self.fields_view = ttk.Treeview(rules_pane, columns=columns, show="headings", height=7, selectmode="browse")
        for name, width, anchor in (
            ("column", 120, "w"),
            ("type", 150, "w"),
            ("description", 220, "w"),
            ("required", 50, "center"),
            ("result", 190, "w"),
        ):
            self.fields_view.column(name, width=width, anchor=anchor)
        self.fields_view.tag_configure("missing_required", foreground=RED)
        self.fields_view.tag_configure("missing_optional", foreground=GRAY)
        self.fields_view.pack(fill=tk.BOTH, expand=True)
        self.fields_view.bind("<<TreeviewSelect>>", lambda _event: self.on_field_selected())

        list_buttons = ttk.Frame(rules_pane)
        list_buttons.pack(fill=tk.X, pady=(6, 12))
        self.button(list_buttons, "button.field_new", self.new_field).pack(side=tk.LEFT)
        self.button(list_buttons, "button.field_remove", self.remove_field).pack(side=tk.LEFT, padx=(8, 0))
        self.button(list_buttons, "button.field_up", lambda: self.move_field(-1)).pack(side=tk.LEFT, padx=(8, 0))
        self.button(list_buttons, "button.field_down", lambda: self.move_field(1)).pack(side=tk.LEFT, padx=(8, 0))

        self.editor = ttk.LabelFrame(rules_pane, padding=10)
        self.editor.pack(fill=tk.X)
        self.editor.columnconfigure(1, weight=1)
        self.field_vars = {name: tk.StringVar() for name in ("column", "labels", "start", "end", "pattern")}
        self.field_required = tk.BooleanVar(value=True)
        for variable in (*self.field_vars.values(), self.field_required):
            variable.trace_add("write", lambda *_args: self.on_editor_changed())

        self.label(self.editor, "label.field_column").grid(row=0, column=0, sticky="w", padx=(0, 8))
        self.column_entry = ttk.Entry(self.editor, textvariable=self.field_vars["column"])
        self.column_entry.grid(row=0, column=1, sticky="ew")
        required_box = ttk.Checkbutton(self.editor, variable=self.field_required)
        self._translate(required_box, "label.field_required").grid(row=0, column=2, sticky="w", padx=(12, 0))
        self.label(self.editor, "label.field_type").grid(row=1, column=0, sticky="w", padx=(0, 8), pady=(8, 0))
        self.type_box = ttk.Combobox(self.editor, state="readonly")
        self.type_box.grid(row=1, column=1, sticky="ew", pady=(8, 0))
        self.type_box.bind("<<ComboboxSelected>>", lambda _event: self.on_type_selected())

        self.inputs = ttk.Frame(self.editor)
        self.inputs.grid(row=2, column=0, columnspan=3, sticky="ew", pady=(8, 0))
        self.inputs.columnconfigure(1, weight=1)
        self.input_widgets = {
            "labels": (ttk.Label(self.inputs), ttk.Entry(self.inputs, textvariable=self.field_vars["labels"])),
            "start": (
                self.label(self.inputs, "label.field_start"),
                ttk.Entry(self.inputs, textvariable=self.field_vars["start"]),
            ),
            "end": (
                self.label(self.inputs, "label.field_end"),
                ttk.Entry(self.inputs, textvariable=self.field_vars["end"]),
            ),
            "pattern": (
                self.label(self.inputs, "label.field_pattern"),
                ttk.Entry(self.inputs, textvariable=self.field_vars["pattern"]),
            ),
        }
        self.type_hint = ttk.Label(self.editor, style="Hint.TLabel", wraplength=560, justify=tk.LEFT)
        self.type_hint.grid(row=3, column=0, columnspan=3, sticky="w", pady=(6, 0))
        self.editor_result = ttk.Label(self.editor, wraplength=560, justify=tk.LEFT, font=self.bold_font)
        self.editor_result.grid(row=4, column=0, columnspan=3, sticky="w", pady=(8, 0))
        self.editor_problem = ttk.Label(self.editor, wraplength=560, justify=tk.LEFT, foreground=AMBER)
        self.editor_problem.grid(row=5, column=0, columnspan=3, sticky="w")
        editor_buttons = ttk.Frame(self.editor)
        editor_buttons.grid(row=6, column=0, columnspan=3, sticky="ew", pady=(10, 0))
        self.add_button = self.button(editor_buttons, "button.field_add", self.add_field)
        self.cancel_new_button = self.button(editor_buttons, "button.cancel", self.cancel_new_field)
        self.as_regex_button = self.button(editor_buttons, "button.as_regex", self.convert_to_regex)

        # Sample mail
        self.label(sample_pane, "preview.title", style="Heading.TLabel").pack(anchor="w")
        nav = ttk.Frame(sample_pane)
        nav.pack(fill=tk.X, pady=(6, 0))
        self.prev_button = ttk.Button(nav, text="◀", width=3, command=lambda: self.step_sample(-1))
        self.prev_button.pack(side=tk.LEFT)
        self.next_button = ttk.Button(nav, text="▶", width=3, command=lambda: self.step_sample(1))
        self.next_button.pack(side=tk.LEFT, padx=(4, 0))
        self.sample_title = ttk.Label(nav)
        self.sample_title.pack(side=tk.LEFT, padx=(8, 0))
        sample_actions = ttk.Frame(sample_pane)
        sample_actions.pack(fill=tk.X, pady=(6, 6))
        self.button(sample_actions, "button.sample_load", self.choose_sample_file).pack(side=tk.LEFT)
        self.button(sample_actions, "button.sample_paste", self.paste_sample).pack(side=tk.LEFT, padx=(8, 0))
        self.sample_summary = ttk.Label(sample_pane, wraplength=440, justify=tk.LEFT, font=self.bold_font)
        self.sample_summary.pack(side=tk.BOTTOM, fill=tk.X, pady=(8, 0))
        self.hint(sample_pane, "preview.header_note", wrap=440).pack(side=tk.BOTTOM, fill=tk.X, pady=(6, 0))
        text_frame = ttk.Frame(sample_pane)
        text_frame.pack(fill=tk.BOTH, expand=True)
        self.sample_text = tk.Text(text_frame, wrap="word", height=12, undo=True)
        sample_scroll = ttk.Scrollbar(text_frame, orient=tk.VERTICAL, command=self.sample_text.yview)
        self.sample_text.configure(yscrollcommand=sample_scroll.set)
        self.sample_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sample_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.sample_text.tag_configure("match", background=HIGHLIGHT, foreground="black")
        self.sample_text.bind("<<Modified>>", lambda _event: self.on_sample_edited())

    # --- Einstellungen

    def _build_settings_tab(self, tab: ttk.Frame) -> None:
        self.label(tab, "settings.title", style="Title.TLabel").pack(anchor="w", pady=(0, 12))
        columns = ttk.Frame(tab)
        columns.pack(fill=tk.BOTH, expand=True)
        left, right = ttk.Frame(columns), ttk.Frame(columns)
        left.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 12))
        right.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(12, 0))

        excel = self._section(left, "settings.excel")
        self.var("output_xlsx", "")
        file_row = ttk.Frame(excel)
        file_row.grid(row=0, column=0, columnspan=2, sticky="ew")
        self.entry(file_row, "output_xlsx").pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.button(file_row, "button.browse", self.choose_output_file).pack(side=tk.LEFT, padx=(8, 0))
        self.error_label(excel, "output_xlsx").grid(row=1, column=0, columnspan=2, sticky="w")
        self.hint(excel, "settings.excel_hint", wrap=440).grid(row=2, column=0, columnspan=2, sticky="w", pady=(2, 8))
        self._setting_row(excel, 3, "sheet_data", "settings.sheet_data", "daten")
        self._setting_row(excel, 5, "sheet_errors", "settings.sheet_errors", "fehler")

        language = self._section(left, "settings.language")
        self.language_box = ttk.Combobox(language, state="readonly", values=list(LANGUAGES.values()), width=20)
        self.language_box.set(LANGUAGES[self.lang])
        self.language_box.grid(row=0, column=0, sticky="w")
        self.language_box.bind("<<ComboboxSelected>>", lambda _event: self.on_language_selected())

        limits = self._section(left, "settings.limits")
        self._setting_row(limits, 0, "max_age_days", "settings.max_age_days", "0", width=8)
        self._setting_row(limits, 2, "max_messages", "settings.max_messages", "0", width=8)

        maintenance = self._section(right, "settings.maintenance")
        self.start_over_button = self.button(maintenance, "button.start_over", self.on_start_over)
        self.start_over_button.grid(row=0, column=0, sticky="w")
        self.hint(maintenance, "settings.start_over_hint", wrap=420).grid(row=1, column=0, sticky="w", pady=(2, 10))
        self.button(maintenance, "button.open_log", self.open_log).grid(row=2, column=0, sticky="w")
        self.button(maintenance, "button.open_config_folder", lambda: self.open_path(self.config_dir)).grid(
            row=3, column=0, sticky="w", pady=(6, 0)
        )

        technical = self._section(right, "settings.technical")
        self.var("log_level", "INFO")
        self.label(technical, "label.log_level").grid(row=0, column=0, sticky="w", padx=(0, 8))
        ttk.Combobox(
            technical,
            textvariable=self.form["log_level"],
            values=["DEBUG", "INFO", "WARNING", "ERROR"],
            state="readonly",
            width=12,
        ).grid(row=0, column=1, sticky="w")
        self._setting_row(technical, 1, "sqlite_path", "label.sqlite_path", "")
        self.files_label = ttk.Label(technical, style="Hint.TLabel", wraplength=440, justify=tk.LEFT)
        self.files_label.grid(row=3, column=0, columnspan=2, sticky="w", pady=(8, 0))
        self.var("dry_run", False)  # not shown: "Testlauf" decides per run; kept so the file value survives

    def _section(self, parent, key: str) -> ttk.LabelFrame:
        section = ttk.LabelFrame(parent, padding=10)
        self._translate(section, key)
        section.pack(fill=tk.X, pady=(0, 12))
        section.columnconfigure(1, weight=1)
        return section

    def _setting_row(self, parent, row: int, key: str, label_key: str, default: str, width: int = 30) -> None:
        self.var(key, default)
        self.label(parent, label_key).grid(row=row, column=0, sticky="w", padx=(0, 8), pady=(2, 0))
        self.entry(parent, key, width=width).grid(row=row, column=1, sticky="w", pady=(2, 0))
        self.error_label(parent, key).grid(row=row + 1, column=1, sticky="w")

    # ------------------------------------------------------------------ loading and saving

    def _load(self) -> None:
        self.loading = True
        problem = None
        try:
            config_text = _read_text_or_empty(self.cfg_path)
            config = parse_config_text(config_text) if config_text.strip() else _default_config()
        except (OSError, ValueError) as exc:
            config, problem = _default_config(), str(exc)
        self.password.set((config.source.imap.password if config.source.imap else None) or "")
        for key, value in form_values(config).items():
            if key in self.form:
                self.form[key].set(value)
        try:
            rules_text = _read_text_or_empty(self.rules_path)
            self.fields = parse_rules_text(rules_text) if rules_text.strip() else []
        except (OSError, ValueError) as exc:
            self.fields, problem = [], str(exc)
        self.loading = False
        # Only a real change rewrites config.toml (opening the app keeps comments in a hand-edited file).
        self.saved_config_text = render_config_text(config) if not problem else ""
        self.apply_language()
        self.refresh_fields_view()
        self.select_field(0 if self.fields else None)
        self.refresh_sample_files()
        self.on_source_changed()
        self.save_config()
        if problem:
            self.root.after(
                0, lambda: messagebox.showerror(self.tr("app.title"), self.tr("error.load").format(error=problem))
            )

    def values(self) -> dict[str, str | bool]:
        return {key: variable.get() for key, variable in self.form.items()}

    def on_form_changed(self) -> None:
        if self.loading:
            return
        if self.save_job is not None:
            self.root.after_cancel(self.save_job)
        self.save_job = self.root.after(400, self.save_config)
        self.on_source_changed()

    def save_config(self) -> bool:
        """Save the inputs if they are valid; mark invalid ones. Returns whether everything is saved."""
        self.save_job = None
        config, errors = config_from_form(self.values(), self.lang)
        for key, widget in self.error_labels.items():
            widget.configure(text=errors.get(key, ""))
        self.form_valid = config is not None
        if config is None:
            message = errors.get("_", self.tr("status.not_saved"))
            self.save_status.configure(text=f"⚠ {message}", foreground=RED)
            return False
        text = render_config_text(config)
        try:
            if text != self.saved_config_text:
                self.cfg_path.write_text(text, encoding="utf-8")
                self.saved_config_text = text
        except OSError as exc:
            self.save_status.configure(text=f"⚠ {friendly_error(exc, self.lang)}", foreground=RED)
            return False
        self.save_status.configure(text=self.tr("status.saved"), foreground=GRAY)
        self.update_mail_count()
        if self.notebook.index("current") == 0:
            self.refresh_cards()
        return True

    def save_rules(self) -> None:
        try:
            self.rules_path.write_text(render_rules_text(self.fields), encoding="utf-8")
            self.save_status.configure(text=self.tr("status.saved"), foreground=GRAY)
        except OSError as exc:
            self.save_status.configure(text=f"⚠ {friendly_error(exc, self.lang)}", foreground=RED)

    def flush_saves(self) -> bool:
        if self.save_job is not None:
            self.root.after_cancel(self.save_job)
        self.flush_editor()
        return self.save_config()

    def flush_editor(self) -> None:
        if self.apply_job is not None:
            self.root.after_cancel(self.apply_job)
            self.apply_editor()

    def jump_to_first_error(self) -> None:
        for key, widget in self.error_labels.items():
            if widget.cget("text"):
                self.notebook.select(INPUT_TABS.get(key, 3))
                return

    def on_close(self) -> None:
        if self.running and not messagebox.askyesno(self.tr("app.title"), self.tr("confirm.close_running")):
            return
        self.cancel_event.set()
        if not self.flush_saves() and not messagebox.askyesno(self.tr("app.title"), self.tr("confirm.close_invalid")):
            self.jump_to_first_error()
            return
        self.package_logger.removeHandler(self.log_handler)
        self.root.destroy()

    # ------------------------------------------------------------------ language

    def on_language_selected(self) -> None:
        names = {name: code for code, name in LANGUAGES.items()}
        self.lang = names.get(self.language_box.get(), "de")
        self.apply_language()

    def apply_language(self) -> None:
        self.root.title(self.tr("app.title"))
        for widget, option, key in self.translated:
            widget.configure(**{option: self.tr(key)})
        self._update_tab_titles()
        self._update_details_button()
        self.files_label.configure(
            text=f"{self.tr('label.config_file')}: {self.cfg_path}\n{self.tr('label.rules_file')}: {self.rules_path}"
        )
        self.type_box.configure(values=[self.tr(f"rule.type.{rule_type}") for rule_type in RULE_TYPES])
        self.set_field_type(self.field_type)
        self.update_sample_title()
        self.refresh_fields_view(keep_selection=True)
        self.save_config()
        self.on_source_changed()

    # ------------------------------------------------------------------ start tab

    def on_tab_changed(self) -> None:
        if self.notebook.index("current") == 0:
            self.refresh_cards()

    def _set_card(self, name: str, status: CardStatus) -> None:
        icon, text = self.cards[name]
        symbol, color = {True: ("✓", GREEN), False: ("⚠", AMBER), None: ("•", GRAY)}[status.ok]
        icon.configure(text=symbol, foreground=color)
        text.configure(text=status.text)

    def refresh_cards(self) -> None:
        values = self.values()
        self._set_card("mails", mails_status(values, self.config_dir, bool(self.password.get()), self.lang))
        self._set_card("fields", fields_status(self.fields, self.previews, self.lang))
        excel_path = output_file_path(str(values.get("output_xlsx", "")), self.config_dir)
        sheets = str(values.get("sheet_data") or "daten"), str(values.get("sheet_errors") or "fehler")
        self._set_card("excel", excel_status(excel_path, *sheets, self.lang))

    def toggle_details(self) -> None:
        self.details_visible = not self.details_visible
        if self.details_visible:
            self.details_frame.pack(fill=tk.BOTH, expand=True, pady=(8, 0))
        else:
            self.details_frame.pack_forget()
        self._update_details_button()

    def _update_details_button(self) -> None:
        self.details_button.configure(
            text=self.tr("button.details_hide" if self.details_visible else "button.details_show")
        )

    def append_log(self, text: str) -> None:
        self.log_text.configure(state=tk.NORMAL)
        self.log_text.insert(tk.END, text + "\n")
        self.log_text.see(tk.END)
        self.log_text.configure(state=tk.DISABLED)

    def show_result(self, text: str, color: str) -> None:
        self.result_label.configure(text=text, foreground=color)

    # ------------------------------------------------------------------ running

    def set_busy(self, busy: bool) -> None:
        self.running = busy
        state = tk.DISABLED if busy else tk.NORMAL
        for button in (self.run_button, self.test_run_button, self.start_over_button):
            button.configure(state=state)
        if busy:
            self.progress_bar.configure(value=0, maximum=1)
            self.progress_label.configure(text=self.tr("label.status.running"))
            self.stop_button.configure(state=tk.NORMAL)
            self.progress_row.pack(fill=tk.X, pady=(16, 0), before=self.result_label)
        else:
            self.progress_row.pack_forget()

    def execute_run(self, dry_run: bool) -> None:
        if self.running:
            return
        if not self.flush_saves():
            messagebox.showerror(self.tr("app.title"), self.tr("error.fix_inputs"))
            self.jump_to_first_error()
            return
        if not self.fields:
            messagebox.showerror(self.tr("app.title"), self.tr("card.fields.none"))
            self.notebook.select(2)
            return
        try:
            app_config = load_app_config(self.cfg_path)
            parsing_rules = ParsingRules(fields=self.fields)
        except (OSError, ValueError) as exc:
            messagebox.showerror(self.tr("app.title"), friendly_error(exc, self.lang))
            return
        if app_config.source.imap is not None:
            app_config.source.imap.password = self.password.get() or None
        app_config.app.dry_run = dry_run
        self.package_logger.setLevel(app_config.app.log_level)
        attach_log_file(self.config_dir)
        self.notebook.select(0)
        self.append_log(f"--- {time.strftime('%H:%M:%S')} ---")
        self.show_result("", GRAY)
        self.cancel_event.clear()
        self.set_busy(True)
        threading.Thread(target=self._run_worker, args=(app_config, parsing_rules), daemon=True).start()
        self.root.after(100, self.poll_run)

    def _run_worker(self, app_config, parsing_rules: ParsingRules) -> None:
        # Runs off the UI thread; must not touch tkinter objects.
        def report_progress(current: int, total: int) -> None:
            self.run_results.put(("progress", (current, total)))

        try:
            summary = run_pipeline(app_config, parsing_rules, progress=report_progress, cancel=self.cancel_event)
            self.run_results.put(("ok", (app_config, summary)))
        except Exception as exc:
            if not isinstance(exc, (ValueError, OSError)):
                self.package_logger.debug("Unexpected error", exc_info=True)
            self.run_results.put(("error", exc))

    def poll_run(self) -> None:
        # Log lines and progress first; the run's result ("ok"/"error") is always the last item.
        while True:
            try:
                kind, payload = self.run_results.get_nowait()
            except queue.Empty:
                self.root.after(100, self.poll_run)
                return
            if kind == "log":
                self.append_log(payload)
            elif kind == "progress":
                current, total = payload
                self.progress_bar.configure(maximum=max(total, 1), value=current)
                if not self.cancel_event.is_set():
                    self.progress_label.configure(
                        text=self.tr("label.status.progress").format(current=current, total=total)
                    )
            else:
                break
        self.set_busy(False)
        if kind == "ok":
            app_config, summary = payload
            text = run_summary_text(
                summary,
                Path(app_config.app.output_xlsx).name,
                app_config.app.sheet_errors,
                app_config.app.dry_run,
                self.lang,
            )
            self.append_log(text)
            self.show_result(text, AMBER if summary.failed or summary.cancelled else GREEN)
        else:
            message = friendly_error(payload, self.lang)
            self.append_log(f"ERROR   {message}")
            self.show_result(message, RED)
            messagebox.showerror(self.tr("app.title"), message)
        self.refresh_cards()

    def stop_run(self) -> None:
        self.cancel_event.set()
        self.stop_button.configure(state=tk.DISABLED)
        self.progress_label.configure(text=self.tr("label.status.stopping"))

    def on_start_over(self) -> None:
        if self.running or not messagebox.askyesno(self.tr("app.title"), self.tr("confirm.start_over"), icon="warning"):
            return
        if not self.flush_saves():
            messagebox.showerror(self.tr("app.title"), self.tr("error.fix_inputs"))
            self.jump_to_first_error()
            return
        try:
            app_config = load_app_config(self.cfg_path)
            attach_log_file(self.config_dir)
            backup = start_over(app_config)
        except (OSError, ValueError) as exc:
            messagebox.showerror(self.tr("app.title"), friendly_error(exc, self.lang))
            return
        if backup is not None:
            self.append_log(self.tr("info.backup_created").format(file=backup.name))
        self.execute_run(dry_run=False)

    # ------------------------------------------------------------------ files and folders

    def open_path(self, path: Path, missing_key: str | None = None) -> None:
        if not path.exists():
            if missing_key:
                messagebox.showinfo(self.tr("app.title"), self.tr(missing_key).format(file=path))
            return
        try:
            open_in_default_app(path)
        except OSError as exc:
            messagebox.showerror(self.tr("app.title"), self.tr("error.open_file").format(file=path, error=exc))

    def open_output(self) -> None:
        self.open_path(
            output_file_path(str(self.form["output_xlsx"].get()).strip(), self.config_dir), "info.no_excel_yet"
        )

    def open_log(self) -> None:
        self.open_path(log_file_path(self.config_dir), "info.no_log_yet")

    def eml_folder(self) -> Path:
        return setting_path(str(self.form["eml_folder"].get()).strip() or ".", self.config_dir)

    def open_eml_folder(self) -> None:
        self.open_path(self.eml_folder(), "info.folder_missing")

    def choose_eml_folder(self) -> None:
        current = self.eml_folder()
        chosen = filedialog.askdirectory(
            parent=self.root,
            title=self.tr("dialog.choose_eml_folder"),
            initialdir=str(current if current.is_dir() else self.config_dir),
            mustexist=True,
        )
        if chosen:
            self.form["eml_folder"].set(path_setting(Path(chosen), self.config_dir))
            self.refresh_sample_files()

    def choose_output_file(self) -> None:
        current = output_file_path(str(self.form["output_xlsx"].get()).strip(), self.config_dir)
        chosen = filedialog.asksaveasfilename(
            parent=self.root,
            title=self.tr("dialog.choose_output_xlsx"),
            initialdir=str(current.parent if current.parent.is_dir() else self.config_dir),
            initialfile=current.name,
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
            confirmoverwrite=False,  # new rows are appended; the file is never replaced
        )
        if chosen:
            self.form["output_xlsx"].set(path_setting(Path(chosen), self.config_dir))

    # ------------------------------------------------------------------ e-mail source

    def on_source_changed(self) -> None:
        use_imap = self.form["source_type"].get() == "imap"
        self._set_enabled(self.eml_frame, not use_imap)
        self._set_enabled(self.imap_frame, use_imap)

    def update_mail_count(self) -> None:
        count = mail_count_in_folder(self.values(), self.config_dir)
        if count is None:
            self.eml_count.configure(text=self.tr("mails.eml.missing"), foreground=AMBER)
        else:
            self.eml_count.configure(
                text=self.tr("mails.eml.count").format(count=count), foreground=GREEN if count else AMBER
            )

    def _set_enabled(self, frame: tk.Widget, enabled: bool) -> None:
        for child in frame.winfo_children():
            if isinstance(child, (ttk.Entry, ttk.Button, ttk.Checkbutton)):
                child.state(["!disabled"] if enabled else ["disabled"])
            elif isinstance(child, ttk.Frame):
                self._set_enabled(child, enabled)

    def on_ssl_toggled(self) -> None:
        # Follow the default port of the chosen security, unless the user typed a custom one.
        use_ssl = bool(self.form["imap_use_ssl"].get())
        if self.form["imap_port"].get().strip() in ("", DEFAULT_IMAP_PORTS[not use_ssl]):
            self.form["imap_port"].set(DEFAULT_IMAP_PORTS[use_ssl])

    def test_connection(self) -> None:
        if not self.save_config():
            self.connection_result.configure(text=self.tr("error.fix_inputs"), foreground=RED)
            return
        config = load_app_config(self.cfg_path).source.imap
        if config is None:
            return
        config.password = self.password.get() or None
        self.connection_button.state(["disabled"])
        self.connection_result.configure(text=self.tr("mails.imap.testing"), foreground=GRAY)
        results: queue.Queue = queue.Queue()

        def worker() -> None:
            try:
                results.put(("ok", check_imap_connection(config)))
            except Exception as exc:
                results.put(("error", exc))

        def poll() -> None:
            try:
                kind, payload = results.get_nowait()
            except queue.Empty:
                self.root.after(150, poll)
                return
            self.connection_button.state(["!disabled"])
            if kind == "ok":
                key = "mails.imap.connected" if payload is None else "mails.imap.connected_count"
                text = self.tr(key).format(count=payload, mailbox=config.mailbox)
                self.connection_result.configure(text=text, foreground=GREEN)
            else:
                self.connection_result.configure(
                    text=friendly_error(payload, self.lang).split("\n\n")[0], foreground=RED
                )

        threading.Thread(target=worker, daemon=True).start()
        self.root.after(150, poll)

    # ------------------------------------------------------------------ fields

    def refresh_fields_view(self, keep_selection: bool = False) -> None:
        selected = self.selected_index() if keep_selection else None
        self.fields_view.delete(*self.fields_view.get_children())
        for index, rule in enumerate(self.fields):
            self.fields_view.insert("", tk.END, iid=str(index), values=self._row_values(rule))
        if selected is not None and selected < len(self.fields):
            self.fields_view.selection_set(str(selected))
        self.schedule_preview()

    def _row_values(self, rule: FieldRule) -> tuple[str, ...]:
        required = self.tr("view.field.required_yes") if rule.required else self.tr("view.field.required_no")
        return (rule.column, self.tr(f"rule.type.{rule.type}"), describe_rule(rule, self.lang), required, "–")

    def selected_index(self) -> int | None:
        selection = self.fields_view.selection()
        return int(selection[0]) if selection else None

    def select_field(self, index: int | None) -> None:
        if index is None:
            self.new_field()
            return
        self.fields_view.selection_set(str(index))
        self.fields_view.see(str(index))

    def on_field_selected(self) -> None:
        index = self.selected_index()
        if index is None or (index == self.editing_index and self.editor_mode == "edit"):
            return
        self.flush_editor()  # pending changes belong to the previously selected field
        rule = self.fields[index]
        self.editor_mode = "edit"
        self.editing_index = index
        self._fill_editor(rule)
        self._update_editor_buttons()

    def _fill_editor(self, rule: FieldRule | None) -> None:
        self.loading_editor = True
        self.field_vars["column"].set(rule.column if rule else "")
        self.field_vars["labels"].set("; ".join(rule.label) if rule else "")
        self.field_vars["start"].set((rule.start or "") if rule else "")
        self.field_vars["end"].set((rule.end or "") if rule else "")
        self.field_vars["pattern"].set((rule.pattern or "") if rule else "")
        self.field_required.set(rule.required if rule else True)
        self.loading_editor = False
        self.set_field_type(rule.type if rule else "label")

    def _update_editor_buttons(self) -> None:
        new = self.editor_mode == "new"
        self.editor.configure(text=self.tr("fields.editor_new" if new else "fields.editor_edit"))
        for widget in (self.add_button, self.cancel_new_button, self.as_regex_button):
            widget.pack_forget()
        if new:
            self.add_button.pack(side=tk.LEFT)
            if self.fields:
                self.cancel_new_button.pack(side=tk.LEFT, padx=(8, 0))
        if self.field_type != "regex":
            self.as_regex_button.pack(side=tk.RIGHT)

    def set_field_type(self, rule_type: RuleType) -> None:
        self.field_type = rule_type
        self.type_box.current(RULE_TYPES.index(rule_type))
        for label_widget, entry_widget in self.input_widgets.values():
            label_widget.grid_forget()
            entry_widget.grid_forget()
        if rule_type in LABEL_TYPES:
            shown = ["labels"]
            key = "label.field_labels_optional" if rule_type == "email" else "label.field_labels"
            self.input_widgets["labels"][0].configure(text=self.tr(key))
        elif rule_type == "between":
            shown = ["start", "end"]
        else:
            shown = ["pattern"]
        for row, name in enumerate(shown):
            label_widget, entry_widget = self.input_widgets[name]
            label_widget.grid(row=row, column=0, sticky="w", padx=(0, 8), pady=(0, 4))
            entry_widget.grid(row=row, column=1, sticky="ew", pady=(0, 4))
        self.type_hint.configure(text=self.tr(f"rule.hint.{rule_type}"))
        self._update_editor_buttons()
        self.on_editor_changed()

    def on_type_selected(self) -> None:
        self.set_field_type(RULE_TYPES[self.type_box.current()])

    def editor_rule(self, for_preview: bool = False) -> FieldRule:
        index = self.editing_index if self.editor_mode == "edit" else None
        values = {name: variable.get() for name, variable in self.field_vars.items()}
        return rule_from_inputs(
            column=values["column"] or ("?" if for_preview else ""),
            rule_type=self.field_type,
            labels_text=values["labels"],
            start=values["start"],
            end=values["end"],
            pattern=values["pattern"],
            required=self.field_required.get(),
            other_columns=None if for_preview else [rule.column for i, rule in enumerate(self.fields) if i != index],
            lang=self.lang,
        )

    def on_editor_changed(self) -> None:
        if self.loading_editor:
            return
        self.schedule_preview()
        if self.editor_mode == "edit" and self.editing_index is not None:
            if self.apply_job is not None:
                self.root.after_cancel(self.apply_job)
            self.apply_job = self.root.after(500, self.apply_editor)

    def apply_editor(self) -> None:
        """Changes to the selected field take effect as soon as they are valid; no extra button."""
        self.apply_job = None
        index = self.editing_index
        if self.editor_mode != "edit" or index is None or index >= len(self.fields):
            return
        try:
            rule = self.editor_rule()
            ParsingRules(fields=[*self.fields[:index], rule, *self.fields[index + 1 :]])
        except ValueError:
            return  # the problem is shown below the editor; the list keeps the last valid version
        if rule == self.fields[index]:
            return
        self.fields[index] = rule
        self.fields_view.item(str(index), values=self._row_values(rule))
        self.save_rules()
        self.schedule_preview()

    def new_field(self) -> None:
        self.flush_editor()
        self.editor_mode = "new"
        self.editing_index = None
        self.fields_view.selection_remove(*self.fields_view.selection())
        self._fill_editor(None)
        self._update_editor_buttons()
        self.column_entry.focus_set()

    def cancel_new_field(self) -> None:
        self.select_field(0 if self.fields else None)

    def add_field(self) -> None:
        try:
            rule = self.editor_rule()
            ParsingRules(fields=[*self.fields, rule])
        except ValueError as exc:
            messagebox.showerror(self.tr("app.title"), str(exc))
            return
        self.fields.append(rule)
        self.editing_index = None
        self.save_rules()
        self.refresh_fields_view()
        self.select_field(len(self.fields) - 1)
        self.save_status.configure(text=self.tr("label.status.field_added").format(column=rule.column), foreground=GRAY)

    def remove_field(self) -> None:
        self.flush_editor()
        index = self.selected_index()
        if index is None:
            messagebox.showinfo(self.tr("app.title"), self.tr("error.select_field_remove"))
            return
        column = self.fields[index].column
        if not messagebox.askyesno(self.tr("app.title"), self.tr("confirm.remove_field").format(column=column)):
            return
        self.fields.pop(index)
        self.editing_index = None
        self.save_rules()
        self.refresh_fields_view()
        self.select_field(min(index, len(self.fields) - 1) if self.fields else None)

    def move_field(self, offset: int) -> None:
        self.flush_editor()
        index = self.selected_index()
        if index is None:
            messagebox.showinfo(self.tr("app.title"), self.tr("error.select_field_move"))
            return
        new_index = index + offset
        if not 0 <= new_index < len(self.fields):
            return
        self.fields[index], self.fields[new_index] = self.fields[new_index], self.fields[index]
        self.editing_index = None
        self.save_rules()
        self.refresh_fields_view()
        self.select_field(new_index)

    def convert_to_regex(self) -> None:
        try:
            rule = self.editor_rule(for_preview=self.editor_mode == "new")
        except ValueError as exc:
            messagebox.showerror(self.tr("app.title"), str(exc))
            return
        self.loading_editor = True
        self.field_vars["pattern"].set(rule.regex)
        self.loading_editor = False
        self.set_field_type("regex")

    # ------------------------------------------------------------------ sample mail and live preview

    def current_sample(self) -> SampleMail | None:
        body = self.sample_text.get("1.0", "end-1c")
        if not body.strip() and not self.sample["header_text"]:
            return None
        return SampleMail(title=self.sample["title"], body=body, header_text=self.sample["header_text"])

    def update_sample_title(self) -> None:
        files, index = self.sample["files"], self.sample["index"]
        if 0 <= index < len(files):
            text = self.tr("preview.position").format(name=self.sample["title"], index=index + 1, total=len(files))
        else:
            text = self.sample["title"] or self.tr("preview.none")
        self.sample_title.configure(text=text)
        self.prev_button.state(["!disabled"] if index > 0 else ["disabled"])
        self.next_button.state(["!disabled"] if len(files) > 1 and index < len(files) - 1 else ["disabled"])

    def show_sample(self, sample: SampleMail, index: int = -1) -> None:
        self.sample.update(title=sample.title, header_text=sample.header_text, index=index)
        self.sample_text.delete("1.0", tk.END)
        self.sample_text.insert("1.0", sample.body)
        self.sample_text.edit_reset()
        self.update_sample_title()
        self.schedule_preview()

    def show_sample_file(self, path: Path, index: int = -1) -> None:
        try:
            self.show_sample(load_sample(path), index)
        except OSError as exc:
            messagebox.showerror(self.tr("app.title"), self.tr("error.sample_unreadable").format(file=path, error=exc))

    def refresh_sample_files(self) -> None:
        """Offer the mails of the configured folder; show the first one if nothing is shown yet."""
        old_files, old_index = self.sample["files"], self.sample["index"]
        current = old_files[old_index] if 0 <= old_index < len(old_files) else None
        files = sample_files(self.eml_folder(), str(self.form["eml_glob"].get()).strip() or "*.eml")
        self.sample["files"] = files
        if current in files:
            self.sample["index"] = files.index(current)
        elif files and self.current_sample() is None:
            self.show_sample_file(files[0], 0)
            return
        else:
            self.sample["index"] = -1
        self.update_sample_title()
        self.schedule_preview()

    def step_sample(self, offset: int) -> None:
        files, index = self.sample["files"], self.sample["index"]
        new_index = index + offset if index >= 0 else 0
        if 0 <= new_index < len(files):
            self.show_sample_file(files[new_index], new_index)

    def choose_sample_file(self) -> None:
        folder = self.eml_folder()
        chosen = filedialog.askopenfilename(
            parent=self.root,
            title=self.tr("dialog.choose_sample"),
            initialdir=str(folder if folder.is_dir() else self.config_dir),
            filetypes=[(self.tr("dialog.eml_files"), "*.eml"), (self.tr("dialog.all_files"), "*")],
        )
        if chosen:
            path = Path(chosen)
            files = self.sample["files"]
            self.show_sample_file(path, files.index(path) if path in files else -1)

    def paste_sample(self) -> None:
        try:
            text = self.root.clipboard_get()
        except tk.TclError:
            text = ""
        if not text.strip():
            messagebox.showinfo(self.tr("app.title"), self.tr("error.clipboard_empty"))
            return
        self.show_sample(SampleMail(title=self.tr("preview.pasted"), body=normalize_body(text)))

    def on_sample_edited(self) -> None:
        if self.sample_text.edit_modified():
            self.sample_text.edit_modified(False)
            self.schedule_preview()

    def schedule_preview(self) -> None:
        # Typing triggers many updates; compute once the input settles.
        if self.preview_job is not None:
            self.root.after_cancel(self.preview_job)
        self.preview_job = self.root.after(120, self.update_preview)

    def update_preview(self) -> None:
        self.preview_job = None
        sample = self.current_sample()
        self.previews = [preview_rule(rule, sample) for rule in self.fields] if sample else None
        for index, rule in enumerate(self.fields):
            iid = str(index)
            if not self.fields_view.exists(iid):
                continue
            preview = self.previews[index] if self.previews else None
            self.fields_view.set(iid, "result", result_text(preview, self.lang))
            tags: tuple[str, ...] = ()
            if preview is not None and not preview.found:
                tags = ("missing_required",) if rule.required else ("missing_optional",)
            self.fields_view.item(iid, tags=tags)
        if sample is None:
            self.sample_summary.configure(text=self.tr("preview.no_sample_summary"), foreground=GRAY)
        elif self.fields and self.previews is not None:
            sheet = str(self.form["sheet_errors"].get()).strip() or "fehler"
            text = summary_text(self.fields, self.previews, sheet, self.lang)
            self.sample_summary.configure(text=text, foreground=GREEN if text.startswith("✓") else RED)
        else:
            self.sample_summary.configure(text="")
        self._update_editor_preview(sample)
        self._set_card("fields", fields_status(self.fields, self.previews, self.lang))

    def _update_editor_preview(self, sample: SampleMail | None) -> None:
        """Result of the rule being edited, and what is still wrong with its inputs (e.g. a duplicate column)."""
        self.sample_text.tag_remove("match", "1.0", tk.END)
        self.editor_problem.configure(text="")
        inputs = {"between": ("start", "end"), "regex": ("pattern",)}.get(self.field_type, ("labels",))
        if self.field_type != "email" and not any(self.field_vars[name].get().strip() for name in inputs):
            self.editor_result.configure(text=self.tr("preview.editor.waiting"), foreground=GRAY)
            return
        try:
            rule = self.editor_rule(for_preview=True)
        except ValueError as exc:
            self.editor_result.configure(text=f"⚠ {exc}", foreground=AMBER)
            return
        try:
            self.editor_rule()
        except ValueError as exc:
            self.editor_problem.configure(text=f"⚠ {exc}")
        if sample is None:
            self.editor_result.configure(text=self.tr("preview.editor.no_sample"), foreground=GRAY)
            return
        preview = preview_rule(rule, sample)
        if not preview.found:
            self.editor_result.configure(text=self.tr("preview.editor.not_found"), foreground=RED)
            return
        key = "preview.editor.from_header" if preview.from_header else "preview.editor.found"
        self.editor_result.configure(text=self.tr(key).format(value=preview.value), foreground=GREEN)
        if preview.span is not None:
            start, end = (f"1.0 + {offset} chars" for offset in preview.span)
            self.sample_text.tag_add("match", start, end)
            self.sample_text.see(start)


def run_app(config_path: Path, rules_path: Path) -> None:
    root = tk.Tk()
    app = App(root, config_path, rules_path)
    try:
        root.mainloop()
    finally:
        app.package_logger.removeHandler(app.log_handler)
