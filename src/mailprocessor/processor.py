"""Pipeline orchestration."""

from __future__ import annotations

import logging
from collections.abc import Iterator
from contextlib import closing
from dataclasses import dataclass
from datetime import UTC, datetime, timedelta
from email.utils import parsedate_to_datetime
from pathlib import Path

from mailprocessor.config import AppConfig, ParsingRules
from mailprocessor.errors import MailFolderNotFoundError
from mailprocessor.excel_writer import ExcelOutput
from mailprocessor.ledger import Ledger, LedgerKey, compute_content_hash
from mailprocessor.models import MailReadError, NormalizedMail
from mailprocessor.parser import ParseResult, parse_mail
from mailprocessor.sources.eml_folder_source import iter_eml_messages
from mailprocessor.sources.imap_source import iter_imap_messages

logger = logging.getLogger(__name__)


@dataclass(frozen=True)
class RunSummary:
    seen: int
    processed: int
    skipped: int
    failed: int


def _is_within_max_age(message: NormalizedMail, max_age_days: int, now_utc: datetime) -> bool:
    if max_age_days <= 0:
        return True
    try:
        message_datetime = parsedate_to_datetime(message.date_raw)
    except (TypeError, ValueError):
        # Without a usable Date header the age is unknown; process it rather than drop it silently.
        logger.warning("Message %s has no valid Date header; ignoring max_age_days for it", message.message_identity)
        return True
    if message_datetime.tzinfo is None:
        message_datetime = message_datetime.replace(tzinfo=UTC)
    oldest_allowed = now_utc - timedelta(days=max_age_days)
    return message_datetime.astimezone(UTC) >= oldest_allowed


def iter_source_messages(
    app_cfg: AppConfig, now_utc: datetime | None = None
) -> Iterator[NormalizedMail | MailReadError]:
    effective_now_utc = now_utc or datetime.now(UTC)
    max_age_days = app_cfg.app.max_age_days
    if app_cfg.source.type == "imap":
        assert app_cfg.source.imap is not None
        # IMAP filters by age on the server (SEARCH SINCE).
        yield from iter_imap_messages(app_cfg.source.imap, max_age_days=max_age_days, now_utc=effective_now_utc)
        return

    assert app_cfg.source.eml is not None
    for message in iter_eml_messages(Path(app_cfg.source.eml.folder), glob_pattern=app_cfg.source.eml.glob):
        if isinstance(message, MailReadError) or _is_within_max_age(message, max_age_days, effective_now_utc):
            yield message


def _safe_parse(message: NormalizedMail, rules: ParsingRules) -> ParseResult:
    try:
        return parse_mail(message.body_text, rules, message.header_text)
    except Exception as exc:  # a single message must never abort the batch
        logger.debug("Parser crashed for message %s", message.message_identity, exc_info=True)
        reason = f"Internal parser error ({type(exc).__name__})"
        return ParseResult(values={}, missing_required=[], error_reason=reason)


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


def run_pipeline(app_cfg: AppConfig, rules: ParsingRules) -> RunSummary:
    """Process new messages. Nothing is written when `dry_run` is set.

    `max_messages` limits how many *new* (not already processed) messages one run handles.
    """
    dry_run = app_cfg.app.dry_run
    dry_note = " (dry run: nothing is written)" if dry_run else ""
    logger.info("Run started: reading from %s%s", _describe_source(app_cfg), dry_note)
    _check_eml_folder(app_cfg)

    excel = ExcelOutput(
        Path(app_cfg.app.output_xlsx),
        app_cfg.app.sheet_data,
        app_cfg.app.sheet_errors,
        [field.column for field in rules.fields],
    )
    if not dry_run:
        excel.check_writable()

    seen = processed = skipped = failed = 0
    max_messages = app_cfg.app.max_messages

    with Ledger(Path(app_cfg.app.sqlite_path), read_only=dry_run) as ledger, closing(
        iter_source_messages(app_cfg)
    ) as messages:
        if excel.is_new and ledger.count_processed():
            logger.warning(
                "%s is new but the ledger already lists processed messages; those are not exported again. "
                "Delete the ledger file (%s) to re-export everything.",
                excel.path,
                app_cfg.app.sqlite_path,
            )

        for item in messages:
            is_mail = isinstance(item, NormalizedMail)
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
            if is_mail:
                # Column names only; extracted values are confidential and never logged.
                logger.debug(
                    "%s: found %s; missing %s",
                    item.display_name,
                    ", ".join(result.values) or "nothing",
                    ", ".join(result.missing_required) or "nothing",
                )
            if result.error_reason is None:
                processed += 1
                logger.info("Processed %s", item.display_name)
                if not dry_run:
                    excel.append_data(result.values, item.body_text)
                    excel.remove_errors_for(key.source_type, key.source_location, key.message_identity)
                    ledger.mark_processed(key)
                continue

            failed += 1
            logger.warning("Failed %s: %s", item.display_name, result.error_reason)
            if not dry_run:
                ledger.mark_failed(key, result.error_reason)
                excel.upsert_error(
                    {
                        "source_type": key.source_type,
                        "source_location": key.source_location,
                        "message_identity": key.message_identity,
                        "missing_columns": ", ".join(result.missing_required),
                        "error_reason": result.error_reason,
                        "processed_at": datetime.now(UTC).isoformat(),
                    }
                )

        if not dry_run:
            # Save the workbook first; the ledger only records messages whose rows are on disk.
            excel.save()
            ledger.commit()
            logger.info("Saved %s (%d new row(s), %d error(s))", excel.path, processed, failed)
        else:
            logger.info("Dry run: %s and the ledger were not changed", excel.path)

    logger.info(
        "Run finished: seen=%d processed=%d skipped=%d failed=%d", seen, processed, skipped, failed
    )
    if failed:
        logger.warning("%d message(s) failed; details are in the '%s' sheet", failed, app_cfg.app.sheet_errors)

    return RunSummary(seen=seen, processed=processed, skipped=skipped, failed=failed)
