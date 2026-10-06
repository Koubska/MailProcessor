"""Keyboard shortcuts for the main actions."""

from __future__ import annotations

from dataclasses import dataclass


@dataclass(frozen=True)
class Shortcut:
    sequences: tuple[str, ...]  # tkinter event sequences
    label: str  # shown in hints, e.g. "Strg+Enter"


def shortcuts(platform: str, lang: str) -> dict[str, Shortcut]:
    """Keyboard shortcuts for the main actions; Command on macOS, Control elsewhere."""
    if platform == "darwin":
        return {
            "run": Shortcut(("<Command-Return>",), "⌘↩"),
            "test_run": Shortcut(("<Command-Shift-Return>",), "⇧⌘↩"),
            "open_excel": Shortcut(("<Command-e>", "<Command-E>"), "⌘E"),
        }
    control, shift = ("Strg", "Umschalt") if lang == "de" else ("Ctrl", "Shift")
    return {
        "run": Shortcut(("<Control-Return>",), f"{control}+Enter"),
        "test_run": Shortcut(("<Control-Shift-Return>",), f"{control}+{shift}+Enter"),
        "open_excel": Shortcut(("<Control-e>", "<Control-E>"), f"{control}+E"),
    }
