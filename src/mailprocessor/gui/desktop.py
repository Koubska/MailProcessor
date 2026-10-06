"""What the GUI needs from the operating system: display scaling and opening files in other programs."""

from __future__ import annotations

import os
import subprocess
import sys
from pathlib import Path

# Windows' reference resolution: 96 pixels per inch is 100 % display scaling.
_REFERENCE_PIXELS_PER_INCH = 96


def enable_windows_dpi_awareness(platform: str = sys.platform, windll: object | None = None) -> bool:
    """Tell Windows that the app scales itself, so text stays sharp on scaled screens (125 %, 150 %).

    Without this, Windows draws the window at 100 % and stretches the picture, which looks blurry.
    Must run before the first Tk window exists. Returns whether Windows accepted it; a failure never stops the app.
    """
    if platform != "win32":
        return False
    if windll is None:
        import ctypes

        windll = ctypes.windll
    try:
        # 1 = system DPI aware (Windows 8.1+). Not per-monitor (2): Tk 8.6 cannot re-layout a window that
        # moves to a monitor with a different scaling, so Windows should handle that case.
        windll.shcore.SetProcessDpiAwareness(1)
        return True
    except (AttributeError, OSError):
        pass
    try:
        windll.user32.SetProcessDPIAware()  # Windows 7
        return True
    except (AttributeError, OSError):
        return False


def ui_scale(pixels_per_inch: float) -> float:
    """Factor for sizes given in pixels: 1.5 at 150 % Windows scaling; never below 1 (macOS reports 72)."""
    return max(1.0, round(pixels_per_inch / _REFERENCE_PIXELS_PER_INCH, 2))


def open_in_default_app(path: Path) -> None:
    """Open a file with the program the system uses for it (Excel, LibreOffice, Numbers, ...)."""
    if sys.platform == "win32":
        os.startfile(path)
    elif sys.platform == "darwin":
        subprocess.Popen(["open", str(path)])
    else:
        subprocess.Popen(["xdg-open", str(path)])
