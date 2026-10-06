"""The main window: combines the tabs, loads and autosaves the settings, and keeps the window state."""

from __future__ import annotations

import logging
import queue
import sys
import threading
import tkinter as tk
from collections.abc import Callable
from importlib import resources
from pathlib import Path
from tkinter import font, ttk

from mailprocessor.gui.config_files import (
    default_config,
    parse_config_text,
    parse_rules_text,
    read_text_or_empty,
    render_config_text,
    render_rules_text,
)
from mailprocessor.gui.desktop import enable_windows_dpi_awareness, ui_scale
from mailprocessor.gui.form import config_from_form, form_values
from mailprocessor.gui.log_handler import QueueLogHandler
from mailprocessor.gui.messages import friendly_error
from mailprocessor.gui.preview import RulePreview
from mailprocessor.gui.rule_lists import ProfileList
from mailprocessor.gui.shortcuts import shortcuts
from mailprocessor.gui.window.base import WindowBase
from mailprocessor.gui.window.field_editor import FieldEditor
from mailprocessor.gui.window.fields_tab import FieldsTab
from mailprocessor.gui.window.mails_tab import MailsTab
from mailprocessor.gui.window.profiles import ProfileActions
from mailprocessor.gui.window.run_control import RunControl
from mailprocessor.gui.window.sample_pane import SamplePane, SampleState
from mailprocessor.gui.window.settings_tab import SettingsTab
from mailprocessor.gui.window.start_tab import StartTab
from mailprocessor.gui.window.widgets import (
    FIELDS_TAB,
    GRAY,
    MAILS_TAB,
    RED,
    SETTINGS_TAB,
    START_TAB,
    TOOLTIP_BACKGROUND,
    TOOLTIP_FOREGROUND,
)
from mailprocessor.gui.window_state import (
    STATE_FILE_NAME,
    WindowState,
    fit_geometry,
    load_window_state,
    save_window_state,
)
from mailprocessor.i18n import resolve_language
from mailprocessor.parser import ParseResult
from mailprocessor.rule_patterns import RULE_TYPES, RuleType
from mailprocessor.run_summary import Problem

# Which tab holds which input, to jump to the first problem.
INPUT_TABS = {
    **dict.fromkeys(("eml_folder", "imap_host", "imap_port", "imap_username"), MAILS_TAB),
    **dict.fromkeys(
        ("output_xlsx", "sheet_data", "sheet_errors", "sqlite_path", "max_age_days", "max_messages"), SETTINGS_TAB
    ),
}


class App(StartTab, RunControl, MailsTab, FieldsTab, FieldEditor, ProfileActions, SamplePane, SettingsTab, WindowBase):
    """The window. Each tab's widgets and handlers live in its mixin; they share state through `self`."""

    def __init__(self, root: tk.Tk, config_path: Path, rules_path: Path) -> None:
        self.root = root
        # Sizes in pixels are multiplied by this; 1.5 on a Windows screen at 150 % (see gui.desktop.ui_scale).
        self.scale = ui_scale(root.winfo_fpixels("1i"))
        self.cfg_path = config_path
        self.rules_path = rules_path
        self.config_dir = config_path.resolve().parent
        self.lang = resolve_language()

        self.form: dict[str, tk.Variable] = {}
        self.error_labels: dict[str, ttk.Label] = {}
        self.translated: list[tuple[tk.Widget, str, str]] = []
        self.profiles = ProfileList([])
        self.shortcuts = shortcuts(sys.platform, self.lang)
        self.window_state_path = self.config_dir / "data" / STATE_FILE_NAME
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

        self.sample = SampleState()
        self.previews: list[RulePreview] | None = None
        self.best: ParseResult | None = None  # the profile a run would choose for the sample (several profiles)
        self.problems: tuple[Problem, ...] = ()
        self.editor_mode = "edit"  # "edit": changes apply to the selected field; "new": "Hinzufügen" adds it
        self.field_type: RuleType = "label"

        self._setup_window()
        self._setup_styles()
        self._build()
        self._bind_shortcuts()
        self._load()
        self._restore_window_state()
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

    def _setup_window(self) -> None:
        self.root.title(self.tr("app.title"))
        try:
            with resources.as_file(resources.files("mailprocessor") / "assets" / "icon.png") as icon_path:
                self.app_icon = tk.PhotoImage(file=str(icon_path))
            self.root.iconphoto(True, self.app_icon)
        except (OSError, tk.TclError):
            pass  # a missing icon is cosmetic
        width = min(self.px(1240), self.root.winfo_screenwidth() - self.px(80))
        height = min(self.px(860), self.root.winfo_screenheight() - self.px(120))
        self.root.geometry(f"{width}x{height}")
        self.root.minsize(min(self.px(900), width), min(self.px(620), height))

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
        # ttk does not grow table rows with the font; without this, rows overlap on scaled screens.
        style.configure("Treeview", rowheight=base.metrics("linespace") + self.px(6))
        style.configure(
            "Tooltip.TLabel",
            background=TOOLTIP_BACKGROUND,
            foreground=TOOLTIP_FOREGROUND,
            relief="solid",
            borderwidth=1,
            padding=(8, 4),
        )

    def _build(self) -> None:
        outer = ttk.Frame(self.root, padding=(16, 12, 16, 8))
        outer.pack(fill=tk.BOTH, expand=True)
        self.notebook = ttk.Notebook(outer)
        self.notebook.pack(fill=tk.BOTH, expand=True)
        self.tabs = [ttk.Frame(self.notebook, padding=16) for _ in range(SETTINGS_TAB + 1)]
        for tab in self.tabs:
            self.notebook.add(tab)
        self.notebook.bind("<<NotebookTabChanged>>", lambda _event: self.on_tab_changed())

        footer = ttk.Frame(outer)
        footer.pack(fill=tk.X, pady=(8, 0))
        self.save_status = ttk.Label(footer, style="Hint.TLabel")
        self.save_status.pack(side=tk.LEFT)
        # Shown only right after a field was deleted.
        self.undo_button = self.button(footer, "button.undo", self.undo_remove_field, tip="tip.undo")

        self._build_start_tab(self.tabs[START_TAB])
        self._build_mails_tab(self.tabs[MAILS_TAB])
        self._build_fields_tab(self.tabs[FIELDS_TAB])
        self._build_settings_tab(self.tabs[SETTINGS_TAB])
        self._update_tab_titles()

    def _update_tab_titles(self) -> None:
        for index, key in enumerate(("tab.start", "tab.mails", "tab.fields", "tab.settings")):
            self.notebook.tab(index, text=self.tr(key))

    def _load(self) -> None:
        self.loading = True
        problem = None
        try:
            config_text = read_text_or_empty(self.cfg_path)
            config = parse_config_text(config_text) if config_text.strip() else default_config()
        except (OSError, ValueError) as exc:
            config, problem = default_config(), str(exc)
        self.password.set((config.source.imap.password if config.source.imap else None) or "")
        for key, value in form_values(config).items():
            if key in self.form:
                self.form[key].set(value)
        try:
            rules_text = read_text_or_empty(self.rules_path)
            self.profiles = ProfileList(parse_rules_text(rules_text))
        except (OSError, ValueError) as exc:
            self.profiles, problem = ProfileList([]), str(exc)
        self.loading = False
        # Only a real change rewrites config.toml (opening the app keeps comments in a hand-edited file).
        self.saved_config_text = render_config_text(config) if not problem else ""
        self.apply_language()
        self.refresh_profile_box()
        self.refresh_fields_view()
        self.select_field(0 if self.fields else None)
        self.refresh_sample_files()
        self.on_source_changed()
        self.save_config()
        if problem:
            self.root.after(0, lambda: self.show_error(self.tr("error.load").format(error=problem)))

    def values(self) -> dict[str, str | bool]:
        return {key: variable.get() for key, variable in self.form.items()}

    def on_form_changed(self) -> None:
        if self.loading:
            return
        if self.save_job is not None:
            self.root.after_cancel(self.save_job)
        self.save_job = self.root.after(400, self.save_config)
        self.on_source_changed()
        self.update_column_note()  # depends on "one sheet per profile"

    def save_config(self) -> bool:
        """Save the inputs if they are valid; mark invalid ones. Returns whether everything is saved."""
        self.save_job = None
        config, errors = config_from_form(self.values(), self.lang)
        for key, widget in self.error_labels.items():
            widget.configure(text=errors.get(key, ""))
        self.form_valid = config is not None
        if config is None:
            message = errors.get("_", self.tr("status.not_saved"))
            self.set_status(f"⚠ {message}", RED)
            return False
        text = render_config_text(config)
        try:
            if text != self.saved_config_text:
                self.cfg_path.write_text(text, encoding="utf-8")
                self.saved_config_text = text
        except OSError as exc:
            self.set_status(f"⚠ {friendly_error(exc, self.lang)}", RED)
            return False
        self.set_status(self.tr("status.saved"))
        self.update_mail_count()
        if self.notebook.index("current") == START_TAB:
            self.refresh_cards()
        return True

    def save_rules(self) -> None:
        try:
            self.rules_path.write_text(render_rules_text(self.profiles.profiles()), encoding="utf-8")
            self.set_status(self.tr("status.saved"))
        except OSError as exc:
            self.set_status(f"⚠ {friendly_error(exc, self.lang)}", RED)

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
                self.notebook.select(INPUT_TABS.get(key, SETTINGS_TAB))
                return

    def on_close(self) -> None:
        if self.running and not self.confirm(self.tr("confirm.close_running")):
            return
        if not self.flush_saves() and not self.confirm(self.tr("confirm.close_invalid")):
            self.jump_to_first_error()
            return
        self.cancel_event.set()  # only now: the window really closes
        save_window_state(
            self.window_state_path, WindowState(geometry=self.root.geometry(), tab=self.notebook.index("current"))
        )
        self.package_logger.removeHandler(self.log_handler)
        self.root.destroy()

    def _restore_window_state(self) -> None:
        state = load_window_state(self.window_state_path)
        geometry = fit_geometry(state.geometry, self.root.winfo_screenwidth(), self.root.winfo_screenheight())
        if geometry is not None:
            self.root.geometry(geometry)
        if 0 <= state.tab < len(self.tabs):
            self.notebook.select(state.tab)

    def _bind_shortcuts(self) -> None:
        actions = {
            "run": lambda: self.execute_run(dry_run=False),
            "test_run": lambda: self.execute_run(dry_run=True),
            "open_excel": self.open_output,
        }
        for name, action in actions.items():
            for sequence in self.shortcuts[name].sequences:
                self.root.bind(sequence, lambda event, action=action: self._on_shortcut(event, action))

    def _on_shortcut(self, event: tk.Event, action: Callable[[], None]) -> str:
        # In the sample mail, Enter belongs to the text (it already inserted a line break); do not run as well.
        if event.widget is not self.sample_text:
            action()
        return "break"

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
        self.show_problems(self.problems)


def run_app(config_path: Path, rules_path: Path) -> None:
    enable_windows_dpi_awareness()  # before the first window; sharp text on scaled Windows screens
    root = tk.Tk()
    app = App(root, config_path, rules_path)
    try:
        root.mainloop()
    finally:
        app.package_logger.removeHandler(app.log_handler)
