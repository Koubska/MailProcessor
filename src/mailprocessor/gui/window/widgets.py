"""Colors, tab positions and the hover hint shared by all tabs."""

from __future__ import annotations

import tkinter as tk
from collections.abc import Callable
from tkinter import ttk

GREEN, RED, AMBER, GRAY = "#15803d", "#b91c1c", "#b45309", "gray"
TOOLTIP_BACKGROUND, TOOLTIP_FOREGROUND = "#fffbe6", "#1f2937"
HIGHLIGHT = "#fde68a"
# The notebook's tabs, in order.
START_TAB, MAILS_TAB, FIELDS_TAB, SETTINGS_TAB = range(4)


class Tooltip:
    """Hover hint that appears after a short delay. The text is looked up when shown, so it follows the language."""

    DELAY_MS = 600

    def __init__(self, widget: tk.Widget, text: Callable[[], str], wraplength: int) -> None:
        self.widget = widget
        self.text = text
        self.wraplength = wraplength
        self.window: tk.Toplevel | None = None
        self.job: str | None = None
        widget.bind("<Enter>", lambda _event: self._schedule(), add="+")
        widget.bind("<Leave>", lambda _event: self.hide(), add="+")
        widget.bind("<ButtonPress>", lambda _event: self.hide(), add="+")

    def _schedule(self) -> None:
        self._cancel()
        self.job = self.widget.after(self.DELAY_MS, self._show)

    def _cancel(self) -> None:
        if self.job is not None:
            self.widget.after_cancel(self.job)
            self.job = None

    def _show(self) -> None:
        self.job = None
        text = self.text()
        if not text or self.window is not None:
            return
        x = self.widget.winfo_rootx() + 12
        y = self.widget.winfo_rooty() + self.widget.winfo_height() + 6
        self.window = tk.Toplevel(self.widget)
        self.window.wm_overrideredirect(True)
        self.window.wm_geometry(f"+{x}+{y}")
        ttk.Label(self.window, text=text, style="Tooltip.TLabel", wraplength=self.wraplength, justify=tk.LEFT).pack()

    def hide(self) -> None:
        self._cancel()
        if self.window is not None:
            self.window.destroy()
            self.window = None
