"""Build a standalone executable bundle with default config files."""

from __future__ import annotations

import argparse
import os
import shutil
import sys
from pathlib import Path

DEFAULTS_DIR = Path("src/mailprocessor/defaults")
ASSETS_DIR = Path("src/mailprocessor/assets")


def _detect_platform_label() -> str:
    if sys.platform == "darwin":
        return "macos"
    if sys.platform.startswith("win"):
        return "windows"
    return "linux"


def _resolve_executable_name(platform_label: str) -> str:
    return "mailprocessor.exe" if platform_label == "windows" else "mailprocessor"


def _icon_args(platform_label: str) -> list[str]:
    """Executable icon: Windows embeds the .ico; macOS takes the .icns; Linux binaries have none."""
    icon_file = {"windows": "icon.ico", "macos": "icon.icns"}.get(platform_label)
    return ["--icon", str(ASSETS_DIR / icon_file)] if icon_file else []


def _write_bundle_readme(bundle_dir: Path, executable_name: str) -> None:
    run_command = executable_name if executable_name.endswith(".exe") else f"./{executable_name}"
    (bundle_dir / "RUNNING.md").write_text(
        "\n".join(
            [
                "# Running the executable",
                "",
                "Double-click the executable (or start it without arguments) to open the GUI.",
                "It edits `config.toml` and `parsing_rules.toml` in this folder and runs the processing.",
                "",
                "Command line usage:",
                "",
                "1. Adapt `config.toml` to your paths/source settings.",
                "   Relative paths are resolved against the folder containing `config.toml`.",
                "   For `.eml` input, put the files into the `mails` folder.",
                "   For IMAP, set the password in the `MAILPROCESSOR_IMAP_PASSWORD` environment variable",
                "   or enter it when prompted.",
                "2. Adapt `parsing_rules.toml` to your parsing patterns.",
                "3. Run:",
                "",
                f"```bash\n{run_command} --config ./config.toml --rules ./parsing_rules.toml\n```",
                "",
                "Optional:",
                "- `--max-age-days N` to process only recent mails.",
                "- `--dry-run` to parse and report without writing Excel or the ledger.",
                "",
                "The binaries are not code-signed. Windows SmartScreen may show",
                "\"Windows protected your PC\" (choose \"More info\" → \"Run anyway\");",
                "on macOS, right-click the file and choose \"Open\" the first time.",
            ]
        ),
        encoding="utf-8",
    )


def build_bundle(target_platform: str | None) -> Path:
    platform_label = _detect_platform_label()
    if target_platform is not None and target_platform != platform_label:
        raise ValueError(
            f"Requested target '{target_platform}' does not match current platform '{platform_label}'."
        )

    import PyInstaller.__main__

    PyInstaller.__main__.run(
        [
            "--noconfirm",
            "--clean",
            "--onefile",
            "--name",
            "mailprocessor",
            "--paths",
            "src",
            # Built-in defaults, used to recreate missing config files on first start.
            "--add-data",
            f"src/mailprocessor/defaults{os.pathsep}mailprocessor/defaults",
            # Window icon (the GUI loads assets/icon.png).
            "--add-data",
            f"{ASSETS_DIR / 'icon.png'}{os.pathsep}mailprocessor/assets",
            *_icon_args(platform_label),
            "src/mailprocessor/main.py",
        ]
    )

    dist_root = Path("dist")
    bundle_dir = dist_root / f"mailprocessor-{platform_label}"
    if bundle_dir.exists():
        shutil.rmtree(bundle_dir)
    bundle_dir.mkdir(parents=True, exist_ok=True)

    executable_name = _resolve_executable_name(platform_label)
    shutil.copy2(dist_root / executable_name, bundle_dir / executable_name)
    shutil.copy2(DEFAULTS_DIR / "config.toml", bundle_dir / "config.toml")
    shutil.copy2(DEFAULTS_DIR / "parsing_rules.toml", bundle_dir / "parsing_rules.toml")
    shutil.copy2("README.md", bundle_dir / "README.md")
    shutil.copy2("LICENSE", bundle_dir / "LICENSE")
    # The bundled config.toml points at ./mails; ship the folder so a first run works out of the box.
    (bundle_dir / "mails").mkdir()
    _write_bundle_readme(bundle_dir, executable_name)

    archive_path = shutil.make_archive(str(bundle_dir), "zip", root_dir=dist_root, base_dir=bundle_dir.name)
    return Path(archive_path)


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--target-platform", choices=["macos", "windows", "linux"], default=None)
    args = parser.parse_args()
    archive_path = build_bundle(args.target_platform)
    print(f"Built archive: {archive_path}")


if __name__ == "__main__":
    main()
