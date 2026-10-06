"""Desktop GUI for editing the config and rules and running the pipeline.

The modules directly in this package hold the GUI's logic without tkinter, so the CLI and the tests run
without a display. Only `gui.window` imports tkinter, and `launch_gui` imports it lazily.
"""

from __future__ import annotations

from pathlib import Path


def launch_gui(config_path: Path | None = None, rules_path: Path | None = None) -> None:
    from mailprocessor.gui.window import run_app

    run_app(config_path or Path("config.toml"), rules_path or Path("parsing_rules.toml"))
