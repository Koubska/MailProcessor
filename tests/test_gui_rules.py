import logging
import queue
from pathlib import Path

import pytest

from mailprocessor.gui import (
    QueueLogHandler,
    UiFieldRule,
    folder_setting,
    parse_config_text,
    parse_rules_text,
    render_config_text,
    render_rules_text,
)


def test_parse_rules_text_reads_fields() -> None:
    rules_text = (
        '[[fields]]\n'
        'column = "Name"\n'
        'pattern = "(?im)^Name:\\\\s*(.+)$"\n'
        "required = true\n"
    )

    fields = parse_rules_text(rules_text)

    assert fields == [UiFieldRule(column="Name", pattern=r"(?im)^Name:\s*(.+)$", required=True)]


def test_render_rules_text_roundtrip() -> None:
    original = [
        UiFieldRule(column="Name", pattern=r"(?im)^Name:\s*(.+)$", required=True),
        UiFieldRule(column="Telefonnummer", pattern=r"(?im)^Telefonnummer:\s*(.+)$", required=False),
    ]

    rendered = render_rules_text(original)
    reparsed = parse_rules_text(rendered)

    assert reparsed == original


def test_render_config_text_roundtrip() -> None:
    config_text = """
[app]
log_level = "INFO"
sqlite_path = "./data/ledger.db"
output_xlsx = "./out/mail_export.xlsx"
sheet_data = "daten"
sheet_errors = "fehler"
dry_run = false
max_messages = 0
max_age_days = 4

[source]
type = "imap"

[source.imap]
host = "imap.example.com"
port = 993
username = "user@example.com"
password = "plaintext-password"
mailbox = "INBOX"
use_ssl = true
sender_filter = "schule@example.com"
""".strip()

    parsed = parse_config_text(config_text)
    rendered = render_config_text(parsed)
    reparsed = parse_config_text(rendered)

    # Everything round-trips except the password, which is never written to disk.
    assert "password" not in rendered
    assert parsed.source.imap is not None
    parsed.source.imap.password = None
    assert reparsed == parsed


def test_render_rules_text_escapes_control_characters() -> None:
    original = [UiFieldRule(column='Say "hi"\tnow', pattern="a\\b\nc\x7f", required=True)]

    assert parse_rules_text(render_rules_text(original)) == original


def test_mail_text_column_name_is_reserved() -> None:
    text = render_rules_text([UiFieldRule(column="E-Mail-Inhalt", pattern="a", required=True)])

    with pytest.raises(ValueError, match="reserved for the mail text"):
        parse_rules_text(text)


def test_parse_rules_text_rejects_duplicate_columns() -> None:
    text = render_rules_text(
        [UiFieldRule(column="Name", pattern="a", required=True), UiFieldRule(column="Name", pattern="b", required=True)]
    )

    with pytest.raises(ValueError, match="Duplicate parsing column"):
        parse_rules_text(text)


def test_folder_setting_is_relative_inside_config_folder(tmp_path: Path) -> None:
    (tmp_path / "mails" / "2026").mkdir(parents=True)

    assert folder_setting(tmp_path / "mails", tmp_path) == "./mails"
    assert folder_setting(tmp_path / "mails" / "2026", tmp_path) == "./mails/2026"
    assert folder_setting(tmp_path, tmp_path) == "."


def test_folder_setting_is_absolute_outside_config_folder(tmp_path: Path) -> None:
    (tmp_path / "app").mkdir()
    (tmp_path / "elsewhere").mkdir()

    assert folder_setting(tmp_path / "elsewhere", tmp_path / "app") == str((tmp_path / "elsewhere").resolve())


def test_queue_log_handler_forwards_formatted_lines_respecting_level() -> None:
    sink: queue.Queue = queue.Queue()
    logger = logging.getLogger("mailprocessor.test_gui_handler")
    handler = QueueLogHandler(sink)
    logger.addHandler(handler)
    logger.setLevel(logging.INFO)
    try:
        logger.debug("hidden")
        logger.warning("Failed a.eml: Required fields missing: Name")
    finally:
        logger.removeHandler(handler)

    kind, line = sink.get_nowait()
    assert kind == "log"
    assert line.endswith("WARNING Failed a.eml: Required fields missing: Name")
    assert sink.empty()
