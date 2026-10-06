"""Widget helpers shared by the tabs: translated labels and buttons, hover hints, form inputs."""

from __future__ import annotations

import tkinter as tk
from collections.abc import Callable
from pathlib import Path
from tkinter import messagebox, ttk

from mailprocessor.gui.desktop import open_in_default_app
from mailprocessor.gui.rule_lists import RuleList
from mailprocessor.gui.window.widgets import GRAY, Tooltip
from mailprocessor.i18n import quote, t


class WindowBase:
    """The helpers every tab uses to create widgets and talk to the user; `App` combines it with the tabs."""

    def tr(self, key: str) -> str:
        return t(key, self.lang)

    def quote(self, text: str) -> str:
        return quote(text, self.lang)

    @property
    def fields(self) -> RuleList:
        """The fields of the profile shown in the Felder tab."""
        return self.profiles.current

    def px(self, pixels: int) -> int:
        """A size in pixels, scaled like the fonts (Tk scales fonts by itself, plain pixel sizes not)."""
        return round(pixels * self.scale)

    def _translate(self, widget: tk.Widget, key: str, option: str = "text") -> tk.Widget:
        """Set a widget's text and remember it, so a language switch can update it."""
        widget.configure(**{option: self.tr(key)})
        self.translated.append((widget, option, key))
        return widget

    def label(self, parent, key: str, style: str | None = None, **options) -> ttk.Label:
        widget = ttk.Label(parent, style=style or "TLabel", **options)
        return self._translate(widget, key)

    def hint(self, parent, key: str, wrap: int = 560) -> ttk.Label:
        return self.label(parent, key, style="Hint.TLabel", wraplength=self.px(wrap), justify=tk.LEFT)

    def button(
        self,
        parent,
        key: str,
        command: Callable[[], None],
        style: str | None = None,
        tip: str | None = None,
        shortcut: str | None = None,
        **options,
    ) -> ttk.Button:
        widget = ttk.Button(parent, command=command, style=style or "TButton", **options)
        if tip is not None:
            self.tip(widget, tip, shortcut)
        return self._translate(widget, key)

    def tip(self, widget: tk.Widget, key: str, shortcut: str | None = None) -> None:
        """Attach a hover hint; `shortcut` names an action in `self.shortcuts` whose keys are appended."""

        def text() -> str:
            hint = self.tr(key)
            return f"{hint} ({self.shortcuts[shortcut].label})" if shortcut else hint

        Tooltip(widget, text, wraplength=self.px(360))

    def var(self, key: str, value: str | bool) -> tk.Variable:
        variable = tk.BooleanVar(value=value) if isinstance(value, bool) else tk.StringVar(value=value)
        variable.trace_add("write", lambda *_args: self.on_form_changed())
        self.form[key] = variable
        return variable

    def entry(self, parent, key: str, width: int = 40, **options) -> ttk.Entry:
        return ttk.Entry(parent, textvariable=self.form[key], width=width, **options)

    def error_label(self, parent, key: str) -> ttk.Label:
        widget = ttk.Label(parent, style="Error.TLabel", wraplength=self.px(520), justify=tk.LEFT)
        self.error_labels[key] = widget
        return widget

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

    def set_status(self, text: str, color: str = GRAY) -> None:
        """The line at the bottom of the window: saved, added, removed, ..."""
        self.save_status.configure(text=text, foreground=color)

    def show_error(self, message: str) -> None:
        messagebox.showerror(self.tr("app.title"), message)

    def show_info(self, message: str) -> None:
        messagebox.showinfo(self.tr("app.title"), message)

    def confirm(self, question: str, **options) -> bool:
        return messagebox.askyesno(self.tr("app.title"), question, **options)

    def open_path(self, path: Path, missing_key: str | None = None) -> None:
        if not path.exists():
            if missing_key:
                messagebox.showinfo(self.tr("app.title"), self.tr(missing_key).format(file=path))
            return
        try:
            open_in_default_app(path)
        except OSError as exc:
            messagebox.showerror(self.tr("app.title"), self.tr("error.open_file").format(file=path, error=exc))
