import logging
import queue
import socket
import ssl
from pathlib import Path

import pytest

import mailprocessor.gui as gui_module
from mailprocessor.config import FieldRule, Profile
from mailprocessor.errors import (
    ImapLoginError,
    MailFolderNotFoundError,
    MissingPasswordError,
    RunInProgressError,
    SheetHeaderError,
    WorkbookLockedError,
    WorkbookUnreadableError,
)
from mailprocessor.gui import (
    QueueLogHandler,
    describe_rule,
    friendly_error,
    open_in_default_app,
    output_file_path,
    parse_config_text,
    parse_rules_text,
    path_setting,
    render_config_text,
    render_rules_text,
    rule_from_inputs,
    run_summary_text,
    split_labels,
)
from mailprocessor.i18n import t
from mailprocessor.processor import RunSummary


def _profile(*fields: FieldRule, name: str = "Standard") -> Profile:
    # Not validated, so tests can also write invalid rules files.
    return Profile.model_construct(name=name, fields=list(fields))


def test_parse_rules_text_reads_fields() -> None:
    rules_text = '[[fields]]\ncolumn = "Name"\npattern = "(?im)^Name:\\\\s*(.+)$"\nrequired = true\n'

    profiles = parse_rules_text(rules_text)

    # A file from before profiles existed is one profile.
    assert profiles == [
        Profile(name="Standard", fields=[FieldRule(column="Name", pattern=r"(?im)^Name:\s*(.+)$", required=True)])
    ]


def test_parse_rules_text_reads_profiles() -> None:
    rules_text = (
        '[[profiles]]\nname = "Anmeldung"\n\n'
        '[[profiles.fields]]\ncolumn = "Kurs"\ntype = "label"\nlabel = "Angebot:"\n\n'
        '[[profiles]]\nname = "Abmeldung"\n\n'
        '[[profiles.fields]]\ncolumn = "Grund"\ntype = "label"\nlabel = "Grund:"\n'
    )

    profiles = parse_rules_text(rules_text)

    assert [profile.name for profile in profiles] == ["Anmeldung", "Abmeldung"]
    assert [field.column for field in profiles[1].fields] == ["Grund"]
    assert parse_rules_text("") == []


def test_render_rules_text_roundtrip() -> None:
    original = [
        Profile(
            name="Anmeldung",
            fields=[
                FieldRule(column="Name", pattern=r"(?im)^Name:\s*(.+)$", required=True),
                FieldRule(column="Telefonnummer", pattern=r"(?im)^Telefonnummer:\s*(.+)$", required=False),
            ],
        ),
        Profile(name="Abmeldung", fields=[FieldRule(column="Name", type="label", label="Name:")]),
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

[filter]
subject = ["Kontaktformular", "Anmeldung \\"neu\\""]
""".strip()

    parsed = parse_config_text(config_text)
    rendered = render_config_text(parsed)
    reparsed = parse_config_text(rendered)

    # Everything round-trips except the password, which is never written to disk.
    assert "password" not in rendered
    assert parsed.source.imap is not None
    parsed.source.imap.password = None
    assert reparsed == parsed
    # The old IMAP-only sender filter is written as part of the filter for every source.
    assert "sender_filter" not in rendered
    assert reparsed.filter.sender == ["schule@example.com"]
    assert reparsed.filter.subject == ["Kontaktformular", 'Anmeldung "neu"']


def test_render_config_text_leaves_out_an_empty_filter() -> None:
    config = parse_config_text(
        '[app]\nsqlite_path = "l.db"\noutput_xlsx = "o.xlsx"\n[source]\ntype = "eml"\n[source.eml]\nfolder = "m"\n'
    )

    assert "[filter]" not in render_config_text(config)


def test_render_rules_text_escapes_control_characters() -> None:
    original = [
        Profile(name='Say "hi" now', fields=[FieldRule(column='Say "hi"\tnow', pattern="a\\b\nc\x7f", required=True)])
    ]

    assert parse_rules_text(render_rules_text(original)) == original


def test_mail_text_column_name_is_reserved() -> None:
    text = render_rules_text([_profile(FieldRule(column="E-Mail-Inhalt", pattern="a", required=True))])

    with pytest.raises(ValueError, match="is reserved"):
        parse_rules_text(text)


def test_parse_rules_text_rejects_duplicate_columns() -> None:
    text = render_rules_text([_profile(FieldRule(column="Name", pattern="a"), FieldRule(column="Name", pattern="b"))])

    with pytest.raises(ValueError, match="Duplicate parsing column"):
        parse_rules_text(text)


def test_path_setting_is_relative_inside_config_folder(tmp_path: Path) -> None:
    (tmp_path / "mails" / "2026").mkdir(parents=True)

    assert path_setting(tmp_path / "mails", tmp_path) == "./mails"
    assert path_setting(tmp_path / "mails" / "2026", tmp_path) == "./mails/2026"
    assert path_setting(tmp_path, tmp_path) == "."


def test_path_setting_is_absolute_outside_config_folder(tmp_path: Path) -> None:
    (tmp_path / "app").mkdir()
    (tmp_path / "elsewhere").mkdir()

    assert path_setting(tmp_path / "elsewhere", tmp_path / "app") == str((tmp_path / "elsewhere").resolve())


def test_path_setting_works_for_files(tmp_path: Path) -> None:
    assert path_setting(tmp_path / "out" / "export.xlsx", tmp_path) == "./out/export.xlsx"


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


def test_run_summary_text_reports_new_rows_problems_and_skipped() -> None:
    summary = RunSummary(seen=5, processed=2, skipped=2, failed=1)

    text = run_summary_text(summary, "mail_export.xlsx", "fehler", dry_run=False, lang="de")

    assert text == (
        "2 neue Zeile(n) in mail_export.xlsx eingetragen. "
        "1 E-Mail(s) mit Problemen – Details im Blatt „fehler“. "
        "2 bereits verarbeitete E-Mail(s) übersprungen."
    )


def test_run_summary_text_lists_new_rows_per_profile() -> None:
    per_profile = (("Anmeldung", 2), ("Abmeldung", 1), ("Leer", 0))
    summary = RunSummary(seen=3, processed=3, skipped=0, failed=0, per_profile=per_profile)

    text = run_summary_text(summary, "out.xlsx", "fehler", dry_run=False, lang="de")

    assert text == "3 neue Zeile(n) in out.xlsx eingetragen (Anmeldung: 2, Abmeldung: 1)."


def test_run_summary_text_nothing_new() -> None:
    summary = RunSummary(seen=3, processed=0, skipped=3, failed=0)

    text = run_summary_text(summary, "out.xlsx", "errors", dry_run=False, lang="en")

    assert text == "No new emails found. Skipped 3 email(s) processed earlier."


def test_run_summary_text_dry_run_says_nothing_was_saved() -> None:
    summary = RunSummary(seen=4, processed=3, skipped=0, failed=1)

    text = run_summary_text(summary, "out.xlsx", "fehler", dry_run=True, lang="de")

    assert "4 neue E-Mail(s)" in text and "3 fehlerfrei" in text and "1 mit Problemen" in text
    assert "nichts gespeichert" in text


@pytest.mark.parametrize(
    ("exc", "key"),
    [
        (WorkbookLockedError("Cannot write out.xlsx"), "error.workbook_locked"),
        (RunInProgressError("Another run is using ledger.db"), "error.run_in_progress"),
        (SheetHeaderError("no header row"), "error.sheet_header"),
        (WorkbookUnreadableError("Cannot open out.xlsx"), "error.workbook_unreadable"),
        (MailFolderNotFoundError("EML folder does not exist"), "error.mail_folder_missing"),
        (MissingPasswordError("IMAP password is missing"), "error.imap_password_missing"),
        (ImapLoginError("IMAP login failed"), "error.imap_login"),
        (ssl.SSLCertVerificationError("certificate verify failed"), "error.imap_tls"),
        (socket.gaierror("Name or service not known"), "error.imap_connection"),
        (ConnectionRefusedError("refused"), "error.imap_connection"),
        (TimeoutError("timed out"), "error.imap_connection"),
    ],
)
def test_friendly_error_explains_common_failures_and_keeps_details(exc: Exception, key: str) -> None:
    message = friendly_error(exc, "de")

    assert message.startswith(t(key, "de"))
    assert message.endswith(f"Details: {exc}")


def test_friendly_error_falls_back_to_the_original_message() -> None:
    assert friendly_error(ValueError("Invalid regex pattern for column 'Name'"), "de") == (
        "Invalid regex pattern for column 'Name'"
    )


def test_friendly_error_hides_unexpected_details() -> None:
    message = friendly_error(KeyError("secret value"), "en")

    assert "KeyError" in message
    assert "secret value" not in message


def test_output_file_path_is_relative_to_config_folder(tmp_path: Path) -> None:
    assert output_file_path("./out/x.xlsx", tmp_path) == (tmp_path / "out" / "x.xlsx").resolve()
    absolute = tmp_path / "elsewhere.xlsx"
    assert output_file_path(str(absolute), Path("/unused")) == absolute


@pytest.mark.parametrize(("platform", "command"), [("darwin", "open"), ("linux", "xdg-open")])
def test_open_in_default_app_uses_the_system_opener(monkeypatch, tmp_path: Path, platform: str, command: str) -> None:
    calls = []
    monkeypatch.setattr(gui_module.sys, "platform", platform)
    monkeypatch.setattr(gui_module.subprocess, "Popen", lambda args: calls.append(args))

    open_in_default_app(tmp_path / "out.xlsx")

    assert calls == [[command, str(tmp_path / "out.xlsx")]]


def test_run_summary_text_mentions_mails_left_out_by_the_filter() -> None:
    summary = RunSummary(seen=1, processed=1, skipped=0, failed=0, filtered=4)

    text = run_summary_text(summary, "out.xlsx", "fehler", dry_run=False, lang="de")

    assert text == (
        "1 neue Zeile(n) in out.xlsx eingetragen. "
        "4 E-Mail(s) passten nicht zum Filter (Betreff/Absender) und wurden nicht beachtet."
    )


def test_run_summary_text_dry_run_mentions_the_filter() -> None:
    summary = RunSummary(seen=0, processed=0, skipped=0, failed=0, filtered=2)

    text = run_summary_text(summary, "out.xlsx", "fehler", dry_run=True, lang="en")

    assert text.endswith("2 email(s) did not match the filter (subject/sender) and were left out.")


def test_run_summary_text_mentions_a_stopped_run() -> None:
    summary = RunSummary(seen=2, processed=2, skipped=0, failed=0, cancelled=True)

    text = run_summary_text(summary, "out.xlsx", "fehler", dry_run=False, lang="de")

    assert text == (
        "Angehalten – bei der nächsten Übertragung geht es an dieser Stelle weiter. "
        "2 neue Zeile(n) in out.xlsx eingetragen."
    )


SIMPLE_RULES = [
    FieldRule(column="Mail-Adresse", type="email", label=["Von", "From"]),
    FieldRule(column="Absender", type="email"),
    FieldRule(column="Name", type="between", start="mein Kind", end="für folgendes Angebot"),
    FieldRule(column="Kurs", type="label", label="Angebot:", required=False),
    FieldRule(column="Bemerkung", type="next_line", label='Say "hi"'),
    FieldRule(column="Experte", pattern=r"(?m)^X:\s*(.+)$"),
]


def test_render_rules_text_roundtrips_all_rule_types() -> None:
    text = render_rules_text([Profile(name="Standard", fields=SIMPLE_RULES)])

    assert parse_rules_text(text)[0].fields == SIMPLE_RULES
    assert 'label = ["Von", "From"]' in text
    assert 'label = "Angebot:"' in text
    # Simple rules are stored by their inputs only; regex rules keep the original format.
    assert text.count("pattern =") == 1
    assert text.count("type =") == len(SIMPLE_RULES) - 1


def test_describe_rule_in_words() -> None:
    described = [describe_rule(rule, "de") for rule in SIMPLE_RULES]

    assert described == [
        "E-Mail-Adresse in der Zeile „Von“ oder „From“",
        "Erste E-Mail-Adresse im Text",
        "Zwischen „mein Kind“ und „für folgendes Angebot“",
        "Zeile nach „Angebot:“",
        'Zeile unter „Say "hi"“',
        r"(?m)^X:\s*(.+)$",
    ]
    assert describe_rule(SIMPLE_RULES[0], "en") == "Email address in the line “Von” or “From”"


def test_split_labels() -> None:
    assert split_labels(" Von ; From;; ") == ["Von", "From"]


def test_rule_from_inputs_keeps_only_inputs_of_the_chosen_type() -> None:
    rule = rule_from_inputs(
        column=" Kurs ", rule_type="label", labels_text="Angebot:; Kurs", start="left over", pattern="left over"
    )

    assert rule == FieldRule(column="Kurs", type="label", label=["Angebot:", "Kurs"])


@pytest.mark.parametrize(
    ("kwargs", "message"),
    [
        ({"column": " ", "rule_type": "label", "labels_text": "A"}, "Spaltennamen"),
        ({"column": "Kurs", "rule_type": "label", "other_columns": ["Kurs"]}, "gibt es schon"),
        ({"column": "Kurs", "rule_type": "next_line", "labels_text": " ; "}, "Bezeichnung"),
        ({"column": "Name", "rule_type": "between", "start": "mein Kind"}, "Anfang und Ende"),
        ({"column": "X", "rule_type": "regex"}, "regulären Ausdruck"),
        ({"column": "X", "rule_type": "regex", "pattern": "("}, "Invalid regex pattern for column 'X'"),
    ],
)
def test_rule_from_inputs_explains_what_is_missing(kwargs: dict, message: str) -> None:
    with pytest.raises(ValueError, match=message):
        rule_from_inputs(**kwargs)
