"""Pipeline orchestration: read the source, parse each mail, write the workbook, then commit the ledger."""

from __future__ import annotations

import logging
import threading
from collections.abc import Callable, Iterator
from contextlib import closing
from datetime import UTC, datetime
from pathlib import Path

from mailprocessor.columns import PROFILE_COLUMN
from mailprocessor.config import AppConfig, AppSection, ParsingRules
from mailprocessor.errors import LedgerUnreadableError, MailFolderNotFoundError, WorkbookLockedError
from mailprocessor.excel_writer import (
    INTERNAL_ERROR_REASON,
    MISSING_FIELDS_REASON,
    UNREADABLE_REASON,
    ErrorEntry,
    ExcelOutput,
    error_key,
)
from mailprocessor.ledger import Ledger, LedgerKey, compute_content_hash
from mailprocessor.mail_dates import is_within_max_age, local_time, mail_datetime
from mailprocessor.models import MailReadError, NormalizedMail
from mailprocessor.parser import ParseResult, parse_mail
from mailprocessor.run_summary import MAX_PROBLEM_DETAILS, Problem, RunSummary
from mailprocessor.sources.eml_folder_source import iter_eml_messages
from mailprocessor.sources.imap_source import iter_imap_messages

__all__ = ["Problem", "RunSummary", "data_sheets", "iter_source_messages", "run_pipeline", "start_over"]

logger = logging.getLogger(__name__)

# Called as progress(current, total) while a run reads messages; current counts from 1.
ProgressCallback = Callable[[int, int], None]


def iter_source_messages(
    app_cfg: AppConfig, now_utc: datetime | None = None, progress: ProgressCallback | None = None
) -> Iterator[NormalizedMail | MailReadError]:
    effective_now_utc = now_utc or datetime.now(UTC)
    max_age_days = app_cfg.app.max_age_days
    total = current = 0

    def set_total(count: int) -> None:
        nonlocal total
        total = count

    def advance() -> None:
        # Counts every message the source returns, including ones filtered out below.
        nonlocal current
        current += 1
        if progress is not None:
            progress(current, total)

    if app_cfg.source.type == "imap":
        assert app_cfg.source.imap is not None
        # IMAP filters by age and, where it can, by subject and sender on the server (SEARCH SINCE/SUBJECT/FROM).
        for message in iter_imap_messages(
            app_cfg.source.imap,
            max_age_days=max_age_days,
            now_utc=effective_now_utc,
            on_total=set_total,
            mail_filter=app_cfg.filter,
        ):
            advance()
            yield message
        return

    assert app_cfg.source.eml is not None
    eml = app_cfg.source.eml
    for message in iter_eml_messages(Path(eml.folder), glob_pattern=eml.glob, on_total=set_total):
        advance()
        if isinstance(message, MailReadError) or is_within_max_age(message, max_age_days, effective_now_utc):
            yield message


def _safe_parse(message: NormalizedMail, rules: ParsingRules) -> ParseResult:
    try:
        return parse_mail(message.body_text, rules, message.header_text)
    except Exception as exc:  # a single message must never abort the batch
        logger.debug("Parser crashed for message %s", message.message_identity, exc_info=True)
        reason = f"Internal parser error ({type(exc).__name__})"
        return ParseResult(values={}, missing_required=[], error_reason=reason)


def _error_entry(
    item: NormalizedMail | MailReadError, result: ParseResult, received: datetime | None, several_profiles: bool
) -> ErrorEntry:
    """The error sheet row: where to find the mail, who sent it, and why it failed, in plain German."""
    key = error_key(item.source_type, item.source_location, item.message_identity)
    if isinstance(item, MailReadError):
        return ErrorEntry(key=key, name=item.display_name, reason=f"{UNREADABLE_REASON} ({item.reason})")
    if result.missing_required and several_profiles:
        reason = f"{MISSING_FIELDS_REASON} (am ähnlichsten: Profil „{result.profile}“)"
    elif result.missing_required:
        reason = MISSING_FIELDS_REASON
    else:
        reason = f"{INTERNAL_ERROR_REASON} ({result.error_reason})"
    return ErrorEntry(
        key=key,
        name=item.origin or item.display_name,
        reason=reason,
        sender=item.from_raw,
        subject=item.subject,
        received=received,
        missing=tuple(result.missing_required),
    )


def data_sheets(app: AppSection, rules: ParsingRules) -> dict[str, list[str]]:
    """The data sheets of a run and their field columns.

    One sheet per profile (named like it) with `profile_sheets = "per_profile"`; otherwise all profiles share
    `sheet_data`, which then starts with a "Profil" column if there are several profiles.
    """
    if app.profile_sheets == "per_profile":
        sheets = {profile.name: [field.column for field in profile.fields] for profile in rules.profiles}
        clash = [name for name in sheets if name.casefold() == app.sheet_errors.casefold()]
        if clash:
            raise ValueError(f"Profile '{clash[0]}' has the same name as the sheet for problems; rename one of them")
        return sheets
    profile_column = [PROFILE_COLUMN] if len(rules.profiles) > 1 else []
    return {app.sheet_data: [*profile_column, *rules.columns]}


def _describe_source(app_cfg: AppConfig) -> str:
    if app_cfg.source.type == "imap" and app_cfg.source.imap is not None:
        imap = app_cfg.source.imap
        return f"IMAP mailbox {imap.mailbox} on {imap.host}"
    assert app_cfg.source.eml is not None
    return f"folder {app_cfg.source.eml.folder}"


def _check_eml_folder(app_cfg: AppConfig) -> None:
    """Validate before anything is opened: the folder must exist and is input only."""
    if app_cfg.source.eml is None or app_cfg.source.type != "eml":
        return
    eml_folder = Path(app_cfg.source.eml.folder)
    if not eml_folder.is_dir():
        raise MailFolderNotFoundError(f"EML folder does not exist: {eml_folder}")
    for setting, value in (("output_xlsx", app_cfg.app.output_xlsx), ("sqlite_path", app_cfg.app.sqlite_path)):
        if Path(value).resolve().is_relative_to(eml_folder.resolve()):
            raise ValueError(f"{setting} must not be inside the mail folder {eml_folder}")


class _Run:
    """The state of one run: the open ledger and workbook, the counters and the problems found so far."""

    def __init__(self, app_cfg: AppConfig, rules: ParsingRules, ledger: Ledger, excel: ExcelOutput) -> None:
        self.app = app_cfg.app
        self.filter = app_cfg.filter
        self.rules = rules
        self.ledger = ledger
        self.excel = excel
        self.dry_run = app_cfg.app.dry_run
        self.several_profiles = len(rules.profiles) > 1
        self.rows_per_profile = dict.fromkeys((profile.name for profile in rules.profiles), 0)
        self.seen = self.processed = self.skipped = self.failed = self.filtered = 0
        self.problems: list[Problem] = []

    def handle(self, item: NormalizedMail | MailReadError) -> bool:
        """Process one message. Returns False when the run must stop (`max_messages` new messages handled)."""
        is_mail = isinstance(item, NormalizedMail)
        if is_mail and item.read_error_identity and not self.dry_run:
            # It was read this time, so an earlier "could not be read" row (e.g. a file still being copied,
            # a failed IMAP fetch) is resolved, whatever happens to the mail now.
            self.excel.remove_errors_for(error_key(item.source_type, item.source_location, item.read_error_identity))
        # Checked before the ledger, so changing the filter later picks these mails up.
        if is_mail and not self.filter.matches(subject=item.subject, sender=item.from_raw):
            logger.debug("Left out %s: subject or sender does not match the filter", item.display_name)
            self.filtered += 1
            if not self.dry_run:
                # It may have failed before the filter was set; it is no longer a problem.
                self.excel.remove_errors_for(error_key(item.source_type, item.source_location, item.message_identity))
            return True
        key = LedgerKey(
            source_type=item.source_type,
            source_location=item.source_location,
            message_identity=item.message_identity,
            content_hash=compute_content_hash(item.body_text) if is_mail else "",
        )
        if is_mail and self.ledger.is_already_processed(key):
            logger.debug("Skipped %s: already processed in an earlier run", item.display_name)
            self.seen += 1
            self.skipped += 1
            return True
        max_messages = self.app.max_messages
        if max_messages > 0 and self.processed + self.failed >= max_messages:
            logger.info("Stopping: reached max_messages (%d new messages per run)", max_messages)
            return False
        self.seen += 1

        result = _safe_parse(item, self.rules) if is_mail else ParseResult({}, [], item.reason)
        received = local_time(mail_datetime(item.date_raw)) if is_mail else None
        if is_mail:
            # Column names only; extracted values are confidential and never logged.
            logger.debug(
                "%s: found %s; missing %s%s",
                item.display_name,
                ", ".join(result.values) or "nothing",
                ", ".join(result.missing_required) or "nothing",
                f" (best profile: {result.profile})" if self.several_profiles else "",
            )
        if result.error_reason is None:
            self._record_success(item, key, result, received)
        else:
            self._record_failure(item, key, result, received)
        return True

    def _record_success(
        self, item: NormalizedMail, key: LedgerKey, result: ParseResult, received: datetime | None
    ) -> None:
        self.processed += 1
        self.rows_per_profile[result.profile] += 1
        if self.several_profiles:
            logger.info("Processed %s with profile %s", item.display_name, result.profile)
        else:
            logger.info("Processed %s", item.display_name)
        if self.dry_run:
            return
        values = result.values
        if self.app.profile_sheets == "per_profile":
            sheet = result.profile
        else:
            sheet = self.app.sheet_data
            if self.several_profiles:
                values = {PROFILE_COLUMN: result.profile, **values}
        self.excel.append_data(values, item.body_text, received=received, sheet=sheet)
        self.excel.remove_errors_for(error_key(key.source_type, key.source_location, key.message_identity))
        self.ledger.mark_processed(key)

    def _record_failure(
        self, item: NormalizedMail | MailReadError, key: LedgerKey, result: ParseResult, received: datetime | None
    ) -> None:
        is_mail = isinstance(item, NormalizedMail)
        self.failed += 1
        logger.warning("Failed %s: %s", item.display_name, result.error_reason)
        if len(self.problems) < MAX_PROBLEM_DETAILS:
            self.problems.append(
                Problem(
                    name=item.display_name,
                    reason=result.error_reason or "",
                    missing=tuple(result.missing_required),
                    body=item.body_text if is_mail else None,
                    header_text=item.header_text if is_mail else "",
                    profile=result.profile if self.several_profiles and result.missing_required else "",
                )
            )
        if not self.dry_run:
            self.ledger.mark_failed(key, result.error_reason)
            self.excel.upsert_error(_error_entry(item, result, received, self.several_profiles))

    def summary(self, cancelled: bool) -> RunSummary:
        return RunSummary(
            seen=self.seen,
            processed=self.processed,
            skipped=self.skipped,
            failed=self.failed,
            cancelled=cancelled,
            filtered=self.filtered,
            problems=tuple(self.problems),
            per_profile=tuple(self.rows_per_profile.items()) if self.several_profiles else (),
        )


def _open_output(app_cfg: AppConfig, sheets: dict[str, list[str]], ledger: Ledger) -> ExcelOutput:
    excel = ExcelOutput(Path(app_cfg.app.output_xlsx), sheets, app_cfg.app.sheet_errors)
    if not app_cfg.app.dry_run:
        excel.check_writable()
    if excel.is_new and ledger.count_processed():
        logger.warning(
            "%s is new but the ledger already lists processed messages; those are not exported again. "
            "Delete the ledger file (%s) to re-export everything.",
            excel.path,
            app_cfg.app.sqlite_path,
        )
    return excel


def run_pipeline(
    app_cfg: AppConfig,
    rules: ParsingRules,
    progress: ProgressCallback | None = None,
    cancel: threading.Event | None = None,
) -> RunSummary:
    """Process new messages. Nothing is written when `dry_run` is set.

    `max_messages` limits how many *new* (not already processed) messages one run handles.
    Setting `cancel` stops before the next message; what was handled until then is saved as usual,
    so the next run continues where this one stopped.
    """
    dry_run = app_cfg.app.dry_run
    dry_note = " (dry run: nothing is written)" if dry_run else ""
    logger.info("Run started: reading from %s%s", _describe_source(app_cfg), dry_note)
    _check_eml_folder(app_cfg)
    sheets = data_sheets(app_cfg.app, rules)
    cancelled = False

    with (
        # Opened first: a writing run holds the ledger's lock from here on, so a second run (another window,
        # a scheduled run) is refused before it loads the workbook and could save an older copy over this one.
        Ledger(Path(app_cfg.app.sqlite_path), read_only=dry_run) as ledger,
        closing(iter_source_messages(app_cfg, progress=progress)) as messages,
    ):
        excel = _open_output(app_cfg, sheets, ledger)
        run = _Run(app_cfg, rules, ledger, excel)
        for item in messages:
            if cancel is not None and cancel.is_set():
                cancelled = True
                logger.info("Stopped by the user; saving what was processed so far")
                break
            if not run.handle(item):
                break

        if not dry_run:
            # Save the workbook first; the ledger only records messages whose rows are on disk.
            excel.save()
            ledger.commit()
            logger.info("Saved %s (%d new row(s), %d error(s))", excel.path, run.processed, run.failed)
        else:
            logger.info("Dry run: %s and the ledger were not changed", excel.path)

    logger.info(
        "Run finished: seen=%d processed=%d skipped=%d failed=%d filtered=%d",
        run.seen,
        run.processed,
        run.skipped,
        run.failed,
        run.filtered,
    )
    if run.failed:
        logger.warning("%d message(s) failed; details are in the '%s' sheet", run.failed, app_cfg.app.sheet_errors)
    return run.summary(cancelled)


def start_over(app_cfg: AppConfig, now: datetime | None = None) -> Path | None:
    """Prepare a full re-export: keep the current workbook as a backup and forget all processed messages.

    The mails themselves are not touched. Returns the backup file, or None if there was no workbook yet.
    The workbook is moved first, so a workbook locked by Excel stops this before anything changes.
    """
    workbook = Path(app_cfg.app.output_xlsx)
    backup = None
    # Local time, as users read it.
    stamp = (now or datetime.now().astimezone()).strftime("%Y-%m-%d_%H-%M-%S")
    if workbook.exists():
        backup = workbook.with_name(f"{workbook.stem}_backup_{stamp}{workbook.suffix}")
        try:
            workbook.rename(backup)
        except PermissionError:
            raise WorkbookLockedError(f"Cannot move {workbook}. Is it open in Excel? Close it and try again.") from None
        logger.info("Moved %s to %s", workbook.name, backup.name)
    ledger_path = Path(app_cfg.app.sqlite_path)
    if ledger_path.exists():
        try:
            with Ledger(ledger_path) as ledger:
                ledger.clear()
                ledger.commit()
            logger.info("Cleared the processing history in %s", ledger_path)
        except LedgerUnreadableError:
            # Damaged: it cannot be cleared, so it is kept aside and the next run starts a new one.
            damaged = ledger_path.with_name(f"{ledger_path.stem}_damaged_{stamp}{ledger_path.suffix}")
            ledger_path.rename(damaged)
            for suffix in ("-wal", "-shm"):  # SQLite's side files belong to the damaged state
                ledger_path.with_name(ledger_path.name + suffix).unlink(missing_ok=True)
            logger.warning("The processing history %s was damaged; moved it to %s", ledger_path.name, damaged.name)
    return backup
