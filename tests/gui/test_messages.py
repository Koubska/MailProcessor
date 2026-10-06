"""Plain-language texts about runs, failed mails and errors."""

import socket
import ssl

import pytest

from mailprocessor.errors import (
    ImapLoginError,
    LedgerUnreadableError,
    MailFolderNotFoundError,
    MissingPasswordError,
    RunInProgressError,
    SheetHeaderError,
    WorkbookLockedError,
    WorkbookUnreadableError,
)
from mailprocessor.gui.messages import friendly_error, problem_text, run_summary_text
from mailprocessor.i18n import t
from mailprocessor.run_summary import Problem, RunSummary


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


@pytest.mark.parametrize(
    ("exc", "key"),
    [
        (WorkbookLockedError("Cannot write out.xlsx"), "error.workbook_locked"),
        (RunInProgressError("Another run is using ledger.db"), "error.run_in_progress"),
        (SheetHeaderError("no header row"), "error.sheet_header"),
        (WorkbookUnreadableError("Cannot open out.xlsx"), "error.workbook_unreadable"),
        (LedgerUnreadableError("The ledger is damaged"), "error.ledger_unreadable"),
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


def test_problem_text() -> None:
    missing = Problem(name="a.eml", reason="Required fields missing: Tel, Kurs", missing=("Tel", "Kurs"), body="x")
    unreadable = Problem(name="file:b.eml", reason="Could not read file (OSError)")

    assert problem_text(missing, "de") == "nicht gefunden: „Tel“, „Kurs“"
    assert problem_text(unreadable, "de") == "E-Mail konnte nicht gelesen werden"
    assert problem_text(Problem(name="c", reason="Internal parser error", body="x"), "en") == "Internal parser error"


def test_problem_text_names_the_closest_profile() -> None:
    problem = Problem(name="a.eml", reason="", missing=("Kurs",), body="x", profile="Abmeldung")

    assert problem_text(problem, "de") == "nicht gefunden: „Kurs“ (am ähnlichsten: Profil „Abmeldung“)"
