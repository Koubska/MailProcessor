"""CLI entrypoint."""

from __future__ import annotations

import getpass
import logging
import os
import sys
from pathlib import Path

import typer

from mailprocessor.config import (
    AppConfig,
    create_missing_files,
    load_app_config,
    load_parsing_rules,
)
from mailprocessor.logfile import attach_log_file
from mailprocessor.processor import run_pipeline

PASSWORD_ENV_VAR = "MAILPROCESSOR_IMAP_PASSWORD"

logger = logging.getLogger("mailprocessor")

# Pretty tracebacks are disabled: they are not meant for end users and could expose local variables.
app = typer.Typer(add_completion=False, pretty_exceptions_enable=False)


def _configure_logging(level: str) -> None:
    logging.basicConfig(level=level, format="%(levelname)s %(message)s", stream=sys.stderr, force=True)


def _resolve_imap_password(config: AppConfig) -> None:
    imap = config.source.imap
    if config.source.type != "imap" or imap is None or imap.password:
        return
    password = os.environ.get(PASSWORD_ENV_VAR)
    if not password and sys.stdin.isatty():
        password = getpass.getpass(f"IMAP password for {imap.username}: ")
    if not password:
        raise ValueError(
            f"IMAP password missing: set {PASSWORD_ENV_VAR} or run interactively to be prompted"
        )
    imap.password = password


def app_directory() -> Path:
    """Folder with the shipped config files: next to the executable when frozen, else the working directory."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path.cwd()


def launch_gui(config: Path | None = None, rules: Path | None = None) -> None:
    """Open the GUI; files default to config.toml / parsing_rules.toml in `app_directory()`.

    Missing files are created from the built-in defaults, so a first start in an empty folder works.
    """
    from mailprocessor.gui import launch_gui as _launch_gui

    base_dir = app_directory()
    config_path = config or base_dir / "config.toml"
    rules_path = rules or base_dir / "parsing_rules.toml"
    create_missing_files(config_path, rules_path)
    _launch_gui(config_path=config_path, rules_path=rules_path)


@app.command()
def cli(
    config: Path | None = typer.Option(
        None,
        "--config",
        help="Path to app configuration TOML (required unless --gui is used).",
    ),
    rules: Path | None = typer.Option(
        None,
        "--rules",
        help="Path to parsing rules TOML (required unless --gui is used).",
    ),
    dry_run: bool = typer.Option(
        False,
        "--dry-run",
        help="Parse and report only; do not write the Excel workbook or the SQLite ledger.",
    ),
    max_age_days: int | None = typer.Option(
        None,
        "--max-age-days",
        min=0,
        help="Only process messages that are at most this many days old (0 means unlimited).",
    ),
    gui: bool = typer.Option(
        False,
        "--gui",
        help="Open the desktop GUI instead of running one CLI pipeline execution.",
    ),
) -> None:
    """Process mails from the configured source and export parsed values to Excel."""
    try:
        if gui:
            launch_gui(config=config, rules=rules)
            return
        if config is None or rules is None:
            raise ValueError("--config and --rules are required unless --gui is used")

        app_config = load_app_config(config)
        parsing_rules = load_parsing_rules(rules)
        _configure_logging(app_config.app.log_level)
        attach_log_file(config.resolve().parent)

        if dry_run:
            app_config.app.dry_run = True
        if max_age_days is not None:
            app_config.app.max_age_days = max_age_days
        _resolve_imap_password(app_config)

        summary = run_pipeline(app_config, parsing_rules)
        typer.echo(
            f"seen={summary.seen} processed={summary.processed} "
            f"skipped={summary.skipped} failed={summary.failed}"
            # Only with a filter, so the line stays the same for everyone else.
            + (f" filtered={summary.filtered}" if summary.filtered else "")
        )
        if app_config.app.dry_run:
            typer.echo("Dry run: nothing was written.")
    except (ValueError, OSError) as exc:
        typer.echo(f"Error: {exc}")
        raise typer.Exit(code=1) from exc
    except Exception as exc:
        logger.debug("Unexpected error", exc_info=True)
        typer.echo(f'Error: unexpected {type(exc).__name__}. Set log_level = "DEBUG" in the config for details.')
        raise typer.Exit(code=1) from exc


def main() -> None:
    """Open the GUI when started without arguments (e.g. double-click), otherwise run the CLI."""
    if len(sys.argv) <= 1:
        try:
            launch_gui()
            return
        except Exception as exc:  # no Tk in this build, or no display (e.g. headless Linux)
            typer.echo(f"Could not open the GUI ({type(exc).__name__}); use the command line instead.\n", err=True)
            sys.argv.append("--help")
    app()


if __name__ == "__main__":
    # Required: PyInstaller runs this file as a script.
    main()
