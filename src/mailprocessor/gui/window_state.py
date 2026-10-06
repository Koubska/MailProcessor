"""Remember the window's size, position and tab between starts.

The state is a small JSON file next to the ledger. Reading or writing it never fails: a missing or broken
file just means the default window.
"""

from __future__ import annotations

import json
import re
from dataclasses import asdict, dataclass
from pathlib import Path

STATE_FILE_NAME = "window.json"
# Tk geometry: "WIDTHxHEIGHT+X+Y"; X and Y can be negative ("+-8") for windows snapped to a screen edge.
_GEOMETRY = re.compile(r"^(\d+)x(\d+)\+(-?\d+)\+(-?\d+)$")
# How much of the window must stay on the screen to restore it there.
_VISIBLE_WIDTH, _VISIBLE_HEIGHT = 100, 50
_MAX_NEGATIVE_OFFSET = 20


@dataclass(frozen=True)
class WindowState:
    geometry: str | None = None
    tab: int = 0


def load_window_state(path: Path) -> WindowState:
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return WindowState()
    if not isinstance(data, dict):
        return WindowState()
    geometry, tab = data.get("geometry"), data.get("tab", 0)
    if not (geometry is None or isinstance(geometry, str)) or type(tab) is not int:
        return WindowState()
    return WindowState(geometry=geometry, tab=tab)


def save_window_state(path: Path, state: WindowState) -> None:
    try:
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(json.dumps(asdict(state)), encoding="utf-8")
    except OSError:
        pass  # only a convenience; the next start uses the default window


def fit_geometry(saved: str | None, screen_width: int, screen_height: int) -> str | None:
    """The saved geometry if the window fits and is visible on the current screen, else None."""
    match = _GEOMETRY.match(saved or "")
    if match is None:
        return None
    width, height, x, y = (int(value) for value in match.groups())
    fits = width <= screen_width and height <= screen_height
    visible = (
        -_MAX_NEGATIVE_OFFSET <= x <= screen_width - _VISIBLE_WIDTH
        and -_MAX_NEGATIVE_OFFSET <= y <= screen_height - _VISIBLE_HEIGHT
    )
    return saved if fits and visible else None
