"""Pipeline orchestration."""

from __future__ import annotations

import logging
import threading
from collections.abc import Callable, Iterator
from contextlib import closing
from dataclasses import dataclass
from datetime import UTC, datetime, timedelta
from email.utils import parsedate_to_datetime
from pathlib import Path

from mailprocessor.config import AppConfig, AppSection, ParsingRules
from mailprocessor.errors import MailFolderNotFoundError, WorkbookLockedError
from mailprocessor.excel_writer import (
    INTERNAL_ERROR_REASON,
    MISSING_FIELDS_REASON,
    PROFILE_COLUMN,
    UNREADABLE_REASON,
    ErrorEntry,
    ExcelOutput,
    error_key,
)
from mailprocessor.ledger import Ledger, LedgerKey, compute_content_hash
from mailprocessor.models import MailReadError, NormalizedMail
from mailprocessor.parser import ParseResult, parse_mail
from mailprocessor.sources.eml_folder_source import iter_eml_messages
from mailprocessor.sources.imap_source import iter_imap_messages

logger = logging.getLogger(__name__)

# Called as progress(current, total) while a run reads messages; current counts from 1.
ProgressCallback = Callable[[int, int], None]


# Details are kept for at most this many failed messages per run (they hold the mail text in memory).
MAX_PROBLEM_DETAILS = 500


@dataclass(frozen=True)
class Problem:
    """A message that failed in this run, for the GUI's problem list. Never logged or written anywhere."""

    name: str  # display name: file name or IMAP uid
    reason: str
    missing: tuple[str, ...] = ()
    # The mail text and headers as parsed, so the GUI can show the mail; None if it could not be read.
    body: str | None = None
    header_text: str = ""
    # The profile that came closest, if there are several; its fields are the missing ones.
    profile: str = ""


@dataclass(frozen=True)
class RunSummary:
    seen: int
    processed: int
    skipped: int
    failed: int
    # True if the run was stopped early; everything handled until then was saved.
    cancelled: bool = False
    # Mails left out because subject or sender did not match the filter; not counted in `seen`.
    filtered: int = 0
    problems: tuple[Problem, ...] = ()
    # New rows per profile, in the order of the profiles; empty if there is only one profile.
    per_profile: tuple[tuple[str, int], ...] = ()


def _mail_datetime(date_raw: str) -> datetime | None:
    """The mail's Date header, or None if it is missing or unreadable."""
    try:
        return parsedate_to_datetime(date_raw)
    except (TypeError, ValueError):
        return None


def _local_time(moment: datetime | None) -> datetime | None:
    """As people read it in Excel: local time without a time zone (Excel cells have none)."""
    if moment is None or moment.tzinfo is None:
        return moment
    return moment.astimezone().replace(tzinfo=None)


def _is_within_max_age(message: NormalizedMail, max_age_days: int, now_utc: datetime) -> bool:
    if max_age_days <= 0:
        return True
    message_datetime = _mail_datetime(message.date_raw)
    if message_datetime is None:
        # Without a usable Date header the age is unknown; process it rather than drop it silently.
        logger.warning("Message %s has no valid Date header; ignoring max_age_days for it", message.message_identity)
        return True
    if message_datetime.tzinfo is None:
        message_datetime = message_datetime.replace(tzinfo=UTC)
    oldest_allowed = now_utc - timedelta(days=max_age_days)
    return message_datetime.astimezone(UTC) >= oldest_allowed


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
        if isinstance(message, MailReadError) or _is_within_max_age(message, max_age_days, effective_now_utc):
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
    several_profiles = len(rules.profiles) > 1
    per_profile_sheets = app_cfg.app.profile_sheets == "per_profile"
    rows_per_profile = dict.fromkeys((profile.name for profile in rules.profiles), 0)

    seen = processed = skipped = failed = filtered = 0
    problems: list[Problem] = []
    cancelled = False
    max_messages = app_cfg.app.max_messages

    with (
        # Opened first: a writing run holds the ledger's lock from here on, so a second run (another window,
        # a scheduled run) is refused before it loads the workbook and could save an older copy over this one.
        Ledger(Path(app_cfg.app.sqlite_path), read_only=dry_run) as ledger,
        closing(iter_source_messages(app_cfg, progress=progress)) as messages,
    ):
        excel = ExcelOutput(Path(app_cfg.app.output_xlsx), sheets, app_cfg.app.sheet_errors)
        if not dry_run:
            excel.check_writable()
        if excel.is_new and ledger.count_processed():
            logger.warning(
                "%s is new but the ledger already lists processed messages; those are not exported again. "
                "Delete the ledger file (%s) to re-export everything.",
                excel.path,
                app_cfg.app.sqlite_path,
            )

        for item in messages:
            if cancel is not None and cancel.is_set():
                cancelled = True
                logger.info("Stopped by the user; saving what was processed so far")
                break
            is_mail = isinstance(item, NormalizedMail)
            # Checked before the ledger, so changing the filter later picks these mails up.
            if is_mail and not app_cfg.filter.matches(subject=item.subject, sender=item.from_raw):
                logger.debug("Left out %s: subject or sender does not match the filter", item.display_name)
                filtered += 1
                if not dry_run:
                    # It may have failed before the filter was set; it is no longer a problem.
                    excel.remove_errors_for(error_key(item.source_type, item.source_location, item.message_identity))
                continue
            key = LedgerKey(
                source_type=item.source_type,
                source_location=item.source_location,
                message_identity=item.message_identity,
                content_hash=compute_content_hash(item.body_text) if is_mail else "",
            )
            if is_mail and ledger.is_already_processed(key):
                logger.debug("Skipped %s: already processed in an earlier run", item.display_name)
                seen += 1
                skipped += 1
                continue
            if max_messages > 0 and processed + failed >= max_messages:
                logger.info("Stopping: reached max_messages (%d new messages per run)", max_messages)
                break
            seen += 1

            result = _safe_parse(item, rules) if is_mail else ParseResult({}, [], item.reason)
            received = _local_time(_mail_datetime(item.date_raw)) if is_mail else None
            sheet_key = error_key(key.source_type, key.source_location, key.message_identity)
            if is_mail:
                # Column names only; extracted values are confidential and never logged.
                logger.debug(
                    "%s: found %s; missing %s%s",
                    item.display_name,
                    ", ".join(result.values) or "nothing",
                    ", ".join(result.missing_required) or "nothing",
                    f" (best profile: {result.profile})" if several_profiles else "",
                )
            if result.error_reason is None:
                processed += 1
                rows_per_profile[result.profile] += 1
                if several_profiles:
                    logger.info("Processed %s with profile %s", item.display_name, result.profile)
                else:
                    logger.info("Processed %s", item.display_name)
                if not dry_run:
                    values = result.values
                    if per_profile_sheets:
                        sheet = result.profile
                    else:
                        sheet = app_cfg.app.sheet_data
                        if several_profiles:
                            values = {PROFILE_COLUMN: result.profile, **values}
                    excel.append_data(values, item.body_text, received=received, sheet=sheet)
                    excel.remove_errors_for(sheet_key)
                    ledger.mark_processed(key)
                continue

            failed += 1
            logger.warning("Failed %s: %s", item.display_name, result.error_reason)
            if len(problems) < MAX_PROBLEM_DETAILS:
                problems.append(
                    Problem(
                        name=item.display_name,
                        reason=result.error_reason or "",
                        missing=tuple(result.missing_required),
                        body=item.body_text if is_mail else None,
                        header_text=item.header_text if is_mail else "",
                        profile=result.profile if several_profiles and result.missing_required else "",
                    )
                )
            if not dry_run:
                ledger.mark_failed(key, result.error_reason)
                excel.upsert_error(_error_entry(item, result, received, several_profiles))

        if not dry_run:
            # Save the workbook first; the ledger only records messages whose rows are on disk.
            excel.save()
            ledger.commit()
            logger.info("Saved %s (%d new row(s), %d error(s))", excel.path, processed, failed)
        else:
            logger.info("Dry run: %s and the ledger were not changed", excel.path)

    logger.info(
        "Run finished: seen=%d processed=%d skipped=%d failed=%d filtered=%d",
        seen,
        processed,
        skipped,
        failed,
        filtered,
    )
    if failed:
        logger.warning("%d message(s) failed; details are in the '%s' sheet", failed, app_cfg.app.sheet_errors)

    return RunSummary(
        seen=seen,
        processed=processed,
        skipped=skipped,
        failed=failed,
        cancelled=cancelled,
        filtered=filtered,
        problems=tuple(problems),
        per_profile=tuple(rows_per_profile.items()) if several_profiles else (),
    )


def start_over(app_cfg: AppConfig, now: datetime | None = None) -> Path | None:
    """Prepare a full re-export: keep the current workbook as a backup and forget all processed messages.

    The mails themselves are not touched. Returns the backup file, or None if there was no workbook yet.
    The workbook is moved first, so a workbook locked by Excel stops this before anything changes.
    """
    workbook = Path(app_cfg.app.output_xlsx)
    backup = None
    if workbook.exists():
        # Local time, as users read it.
        stamp = (now or datetime.now().astimezone()).strftime("%Y-%m-%d_%H-%M-%S")
        backup = workbook.with_name(f"{workbook.stem}_backup_{stamp}{workbook.suffix}")
        try:
            workbook.rename(backup)
        except PermissionError:
            raise WorkbookLockedError(f"Cannot move {workbook}. Is it open in Excel? Close it and try again.") from None
        logger.info("Moved %s to %s", workbook.name, backup.name)
    ledger_path = Path(app_cfg.app.sqlite_path)
    if ledger_path.exists():
        with Ledger(ledger_path) as ledger:
            ledger.clear()
            ledger.commit()
        logger.info("Cleared the processing history in %s", ledger_path)
    return backup
