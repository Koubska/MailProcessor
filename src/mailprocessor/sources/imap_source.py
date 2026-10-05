"""IMAP read-only source."""

from __future__ import annotations

import contextlib
import imaplib
import logging
import ssl
from collections.abc import Callable, Iterator
from datetime import UTC, datetime, timedelta
from typing import Protocol

from mailprocessor.config import ImapSourceConfig, MailFilter
from mailprocessor.errors import ImapLoginError, MissingPasswordError
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

    def examine(self, mailbox: str) -> int | None:
        """Open the mailbox with EXAMINE and require the server's [READ-ONLY] confirmation.

        Returns the number of messages the server reports for the mailbox, if it is readable.
        """
        status, data = self._client.select(mailbox, readonly=True)
        if status != "OK":
            raise OSError(f"Failed to open mailbox in readonly mode: {mailbox}")
        _code, confirmation = self._client.response("READ-ONLY")
        if not confirmation or confirmation[0] is None:
            raise OSError(f"IMAP server did not confirm read-only access to {mailbox}; aborting without reading mail")
        try:
            return int(data[0])
        except (TypeError, ValueError, IndexError):
            return None

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


def _any_of(key: str, entries: list[str]) -> list[str]:
    """SEARCH criteria matching one of the entries: OR k "a" OR k "b" k "c". Empty if the server can't check them."""
    # imaplib sends commands as ASCII. Mails are filtered again after fetching (processor.run_pipeline),
    # so entries the server can't check are only left out of the search, never ignored.
    if not entries or not all(entry.isascii() for entry in entries):
        return []
    criteria: list[str] = []
    for entry in entries[:-1]:
        criteria.extend(["OR", key, _quote_imap_string(entry)])
    return [*criteria, key, _quote_imap_string(entries[-1])]


def _build_search_args(mail_filter: MailFilter, max_age_days: int, now_utc: datetime) -> tuple[str, ...]:
    search_args = [*_any_of("FROM", mail_filter.sender), *_any_of("SUBJECT", mail_filter.subject)]
    if max_age_days > 0:
        search_args.extend(["SINCE", _build_since_date_token(max_age_days, now_utc)])
    return tuple(search_args) if search_args else ("ALL",)


def _read_error(source_location: str, uid: bytes, reason: str) -> MailReadError:
    return MailReadError(
        source_type="imap",
        source_location=source_location,
        message_identity=f"uid:{uid.decode('ascii', errors='replace')}",
        reason=reason,
        origin=f"IMAP uid {uid.decode('ascii', errors='replace')}",
    )


def _connect_and_login(
    config: ImapSourceConfig, client_factory: Callable[[str, int], ImapClient] | None
) -> ReadOnlyImapClient:
    if not config.password:
        raise MissingPasswordError("IMAP password is missing")
    if client_factory is None:
        client_factory = default_client_factory(config.use_ssl)
    logger.info(
        "Connecting to %s:%s (%s)", config.host, config.port, "SSL" if config.use_ssl else "STARTTLS"
    )
    client = ReadOnlyImapClient(client_factory(config.host, config.port))
    try:
        status, _ = client.login(config.username, config.password)
    except imaplib.IMAP4.error:
        status = "NO"
    if status != "OK":
        client.logout()
        # Deliberately no server response text and no credentials in the message.
        raise ImapLoginError(f"IMAP login failed for user {config.username!r}; check username and password")
    logger.info("Logged in as %s", config.username)
    return client


def check_imap_connection(
    config: ImapSourceConfig, client_factory: Callable[[str, int], ImapClient] | None = None
) -> int | None:
    """Log in and open the mailbox read-only, without reading any message. Returns the server's message count."""
    client = _connect_and_login(config, client_factory)
    try:
        count = client.examine(config.mailbox)
        logger.info("Connection test: opened mailbox %s read-only", config.mailbox)
        return count
    except imaplib.IMAP4.error as exc:
        raise OSError(f"IMAP error: {exc}") from None
    finally:
        client.logout()


def iter_imap_messages(
    config: ImapSourceConfig,
    client_factory: Callable[[str, int], ImapClient] | None = None,
    max_age_days: int = 0,
    now_utc: datetime | None = None,
    on_total: Callable[[int], None] | None = None,
    mail_filter: MailFilter | None = None,
) -> Iterator[NormalizedMail | MailReadError]:
    """Yield the matching messages; `on_total` receives their number before the first one is fetched.

    The server pre-selects by `mail_filter` where it can; the caller still has to check each mail against it.
    """
    search_args = _build_search_args(mail_filter or MailFilter(), max_age_days, now_utc or datetime.now(UTC))
    source_location = f"imap://{config.host}:{config.port}/{config.mailbox}"
    client = _connect_and_login(config, client_factory)
    try:
        client.examine(config.mailbox)
        logger.info("Opened mailbox %s read-only", config.mailbox)

        status, search_data = client.uid_search(*search_args)
        if status != "OK":
            raise OSError("IMAP search failed")
        if not search_data:
            if on_total is not None:
                on_total(0)
            return
        uid_bytes = search_data[0]
        if not isinstance(uid_bytes, bytes):
            raise OSError("IMAP search returned invalid response data")

        uids = uid_bytes.split()
        logger.info("Found %d message(s) matching %s", len(uids), " ".join(search_args))
        if on_total is not None:
            on_total(len(uids))
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
