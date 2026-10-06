"""Plain-language texts about runs: the result, failed mails and errors."""

from __future__ import annotations

import socket
import ssl

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
from mailprocessor.i18n import quote, t
from mailprocessor.run_summary import Problem, RunSummary


def run_summary_text(summary: RunSummary, output_name: str, error_sheet: str, dry_run: bool, lang: str) -> str:
    """Plain-language result of a run for the output box and status bar."""
    parts = [t("summary.cancelled", lang)] if summary.cancelled else []
    if dry_run:
        parts.append(
            t("summary.dry_run", lang).format(
                new=summary.processed + summary.failed, ok=summary.processed, failed=summary.failed
            )
        )
    else:
        if summary.processed:
            parts.append(t("summary.processed", lang).format(count=summary.processed, file=output_name))
            counts = [f"{name}: {count}" for name, count in summary.per_profile if count]
            if counts:
                parts[-1] = parts[-1].removesuffix(".") + f" ({', '.join(counts)})."
        if summary.failed:
            parts.append(t("summary.failed", lang).format(count=summary.failed, sheet=error_sheet))
        if not summary.processed and not summary.failed:
            parts.append(t("summary.nothing_new", lang))
        if summary.skipped:
            parts.append(t("summary.skipped", lang).format(count=summary.skipped))
    if summary.filtered:
        parts.append(t("summary.filtered", lang).format(count=summary.filtered))
    return " ".join(parts)


# Most specific first: e.g. ImapLoginError and ssl.SSLError are OSErrors too.
_FRIENDLY_ERRORS: tuple[tuple[type[BaseException] | tuple[type[BaseException], ...], str], ...] = (
    (WorkbookLockedError, "error.workbook_locked"),
    (RunInProgressError, "error.run_in_progress"),
    (SheetHeaderError, "error.sheet_header"),
    (WorkbookUnreadableError, "error.workbook_unreadable"),
    (LedgerUnreadableError, "error.ledger_unreadable"),
    (MailFolderNotFoundError, "error.mail_folder_missing"),
    (MissingPasswordError, "error.imap_password_missing"),
    (ImapLoginError, "error.imap_login"),
    (ssl.SSLError, "error.imap_tls"),
    ((socket.gaierror, ConnectionError, TimeoutError), "error.imap_connection"),
)


def friendly_error(exc: BaseException, lang: str) -> str:
    """Explain a failed run without jargon; the technical message follows as details."""
    for error_types, key in _FRIENDLY_ERRORS:
        if isinstance(exc, error_types):
            return f"{t(key, lang)}\n\n{t('error.details', lang)}: {exc}"
    if isinstance(exc, (ValueError, OSError)):
        return str(exc)
    return t("error.unexpected", lang).format(name=type(exc).__name__)


def problem_text(problem: Problem, lang: str) -> str:
    """What is wrong with a failed mail, for the problem list."""
    if problem.missing:
        columns = ", ".join(quote(column, lang) for column in problem.missing)
        if problem.profile:
            return t("problems.missing_profile", lang).format(profile=quote(problem.profile, lang), columns=columns)
        return t("problems.missing", lang).format(columns=columns)
    if problem.body is None:
        return t("problems.unreadable_short", lang)
    return problem.reason
