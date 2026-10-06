"""Keyboard shortcuts for the main actions."""

import pytest

from mailprocessor.gui.shortcuts import shortcuts


@pytest.mark.parametrize(
    ("platform", "lang", "labels", "run_sequence"),
    [
        ("darwin", "de", ("⌘↩", "⇧⌘↩", "⌘E"), "<Command-Return>"),
        ("win32", "de", ("Strg+Enter", "Strg+Umschalt+Enter", "Strg+E"), "<Control-Return>"),
        ("linux", "en", ("Ctrl+Enter", "Ctrl+Shift+Enter", "Ctrl+E"), "<Control-Return>"),
    ],
)
def test_shortcuts_follow_the_platform(platform: str, lang: str, labels: tuple[str, ...], run_sequence: str) -> None:
    result = shortcuts(platform, lang)

    assert tuple(result[action].label for action in ("run", "test_run", "open_excel")) == labels
    assert result["run"].sequences == (run_sequence,)
    assert result["test_run"].sequences[0].endswith("Shift-Return>")
    # Caps Lock must not break the letter shortcut.
    assert {sequence[-2] for sequence in result["open_excel"].sequences} == {"e", "E"}
