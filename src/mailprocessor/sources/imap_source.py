"""IMAP read-only source."""

from __future__ import annotations

import contextlib
import imaplib
import logging
import ssl
from collections.abc import Callable, Iterator
from datetime import UTC, datetime, timedelta
from typing import Protocol

from mailprocessor.config import ImapSourceConfig
from mailprocessor.models import MailReadError, NormalizedMail
from mailprocessor.sources.email_content import parse_message_bytes

logger = logging.getLogger(__name__)


class ImapClient(Protocol):
    def login(self, user: str, password: str): ...
    def select(self, mailbox: str = "INBOX", readonly: bool = False): ...
    def uid(self, command: str, *args): ...
    def response(self, code: str): ...
    def logout(self): ...


# BODY.PEEK[] returns the full message like RFC822 but, unlike RFC822/BODY[], never sets \Seen.
FETCH_ITEMS = "(BODY.PEEK[])"


class ReadOnlyImapClient:
    """Allow-list wrapper; the only way this module talks to the server.

    It deliberately exposes no generic command access, so STORE, COPY, MOVE, EXPUNGE, CLOSE,
    APPEND, DELETE etc. cannot be sent, and fetches can only use BODY.PEEK[].
    """

    def __init__(self, client: ImapClient) -> None:
        self._client = client

    def login(self, user: str, password: str):
        return self._client.login(user, password)

    def examine(self, mailbox: str) -> None:
        """Open the mailbox with EXAMINE and require the server's [READ-ONLY] confirmation."""
        status, _ = self._client.select(mailbox, readonly=True)
        if status != "OK":
            raise OSError(f"Failed to open mailbox in readonly mode: {mailbox}")
        _code, data = self._client.response("READ-ONLY")
        if not data or data[0] is None:
            raise OSError(f"IMAP server did not confirm read-only access to {mailbox}; aborting without reading mail")

    def uid_search(self, *criteria: str):
        return self._client.uid("search", None, *criteria)

    def uid_fetch(self, uid: bytes):
        return self._client.uid("fetch", uid, FETCH_ITEMS)

    def logout(self) -> None:
        with contextlib.suppress(OSError, imaplib.IMAP4.error):
            self._client.logout()


def default_client_factory(use_ssl: bool) -> Callable[[str, int], ImapClient]:
    """Connect with certificate verification; plain connections must upgrade via STARTTLS."""

    def connect(host: str, port: int) -> ImapClient:
        context = ssl.create_default_context()
        if use_ssl:
            return imaplib.IMAP4_SSL(host, port, ssl_context=context)
        client = imaplib.IMAP4(host, port)
        try:
            client.starttls(ssl_context=context)
        except imaplib.IMAP4.error as exc:
            client.shutdown()
            raise OSError(
                f"IMAP server {host}:{port} does not support STARTTLS; refusing to send the password unencrypted"
            ) from exc
        return client

    return connect


def _extract_message_bytes(fetch_data: object) -> bytes | None:
    """The message literal is the second element of the first (envelope, literal) tuple."""
    if not isinstance(fetch_data, list):
        return None
    for item in fetch_data:
        if isinstance(item, tuple) and len(item) >= 2 and isinstance(item[1], bytes):
            return item[1]
    return None


def _build_since_date_token(max_age_days: int, now_utc: datetime) -> str:
    month_names = ("Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")
    cutoff_date = (now_utc - timedelta(days=max_age_days)).date()
    return f"{cutoff_date.day:02d}-{month_names[cutoff_date.month - 1]}-{cutoff_date.year:04d}"


def _quote_imap_string(value: str) -> str:
    escaped = value.replace("\\", "\\\\").replace('"', '\\"')
    return f'"{escaped}"'


def _build_search_args(sender_filter: str | None, max_age_days: int, now_utc: datetime) -> tuple[str, ...]:
    search_args: list[str] = []
    if sender_filter is not None:
        normalized_filter = sender_filter.strip()
        if not normalized_filter:
            raise ValueError("source.imap.sender_filter must not be empty when provided")
        search_args.extend(["FROM", _quote_imap_string(normalized_filter)])
    if max_age_days > 0:
        search_args.extend(["SINCE", _build_since_date_token(max_age_days, now_utc)])
    return tuple(search_args) if search_args else ("ALL",)


def _read_error(source_location: str, uid: bytes, reason: str) -> MailReadError:
    return MailReadError(
        source_type="imap",
        source_location=source_location,
        message_identity=f"uid:{uid.decode('ascii', errors='replace')}",
        reason=reason,
    )


def iter_imap_messages(
    config: ImapSourceConfig,
    client_factory: Callable[[str, int], ImapClient] | None = None,
    max_age_days: int = 0,
    now_utc: datetime | None = None,
) -> Iterator[NormalizedMail | MailReadError]:
    if not config.password:
        raise ValueError("IMAP password is missing")
    search_args = _build_search_args(config.sender_filter, max_age_days, now_utc or datetime.now(UTC))
    if client_factory is None:
        client_factory = default_client_factory(config.use_ssl)

    source_location = f"imap://{config.host}:{config.port}/{config.mailbox}"
    logger.info(
        "Connecting to %s:%s (%s)", config.host, config.port, "SSL" if config.use_ssl else "STARTTLS"
    )
    client = ReadOnlyImapClient(client_factory(config.host, config.port))
    try:
        try:
            status, _ = client.login(config.username, config.password)
        except imaplib.IMAP4.error:
            status = "NO"
        if status != "OK":
            # Deliberately no server response text and no credentials in the message.
            raise OSError(f"IMAP login failed for user {config.username!r}; check username and password")

        logger.info("Logged in as %s", config.username)
        client.examine(config.mailbox)
        logger.info("Opened mailbox %s read-only", config.mailbox)

        status, search_data = client.uid_search(*search_args)
        if status != "OK":
            raise OSError("IMAP search failed")
        if not search_data:
            return
        uid_bytes = search_data[0]
        if not isinstance(uid_bytes, bytes):
            raise OSError("IMAP search returned invalid response data")

        uids = uid_bytes.split()
        logger.info("Found %d message(s) matching %s", len(uids), " ".join(search_args))
        for uid in uids:
            logger.debug("Fetching IMAP uid %s", uid.decode("ascii", errors="replace"))
            status, fetch_data = client.uid_fetch(uid)
            raw_bytes = _extract_message_bytes(fetch_data) if status == "OK" else None
            if raw_bytes is None:
                yield _read_error(source_location, uid, f"IMAP fetch failed (status={status})")
                continue
            try:
                yield parse_message_bytes(
                    raw_bytes, "imap", source_location, origin=f"IMAP uid {uid.decode('ascii', errors='replace')}"
                )
            except Exception as exc:  # one malformed message must not abort the batch
                yield _read_error(source_location, uid, f"Could not parse message ({type(exc).__name__})")
    except imaplib.IMAP4.error as exc:
        raise OSError(f"IMAP error: {exc}") from None
    finally:
        client.logout()
