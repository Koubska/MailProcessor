"""The mail's Date header as a moment in time: for the age limit and the "Eingegangen am" column."""

from __future__ import annotations

import logging
from datetime import UTC, datetime, timedelta
from email.utils import parsedate_to_datetime

from mailprocessor.models import NormalizedMail

logger = logging.getLogger(__name__)


def mail_datetime(date_raw: str) -> datetime | None:
    """The mail's Date header, or None if it is missing, unreadable or cannot be converted to UTC and local time."""
    try:
        moment = parsedate_to_datetime(date_raw)
        if moment.tzinfo is not None:
            # Fails at the edge of the calendar (e.g. the year 10000 in UTC) and, on Windows, can fail before 1970.
            # Checked here, so one such mail cannot abort the run when its age or local time is needed.
            moment.astimezone(UTC)
            moment.astimezone()
        return moment
    except (TypeError, ValueError, OverflowError, OSError):
        return None


def local_time(moment: datetime | None) -> datetime | None:
    """As people read it in Excel: local time without a time zone (Excel cells have none)."""
    if moment is None or moment.tzinfo is None:
        return moment
    return moment.astimezone().replace(tzinfo=None)


def is_within_max_age(message: NormalizedMail, max_age_days: int, now_utc: datetime) -> bool:
    if max_age_days <= 0:
        return True
    message_datetime = mail_datetime(message.date_raw)
    if message_datetime is None:
        # Without a usable Date header the age is unknown; process it rather than drop it silently.
        logger.warning("Message %s has no valid Date header; ignoring max_age_days for it", message.message_identity)
        return True
    if message_datetime.tzinfo is None:
        message_datetime = message_datetime.replace(tzinfo=UTC)
    oldest_allowed = now_utc - timedelta(days=max_age_days)
    return message_datetime.astimezone(UTC) >= oldest_allowed
