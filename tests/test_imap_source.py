"""Reading messages from IMAP: search criteria, identities, and errors per message."""

from __future__ import annotations

from datetime import UTC, datetime

import pytest
from imap_helpers import FakeImapClient, _build_raw_message, _cfg

from mailprocessor.config import ImapSourceConfig, MailFilter
from mailprocessor.models import MailReadError
from mailprocessor.sources.imap_source import (
    iter_imap_messages,
)


def test_iter_imap_messages_reads_messages_readonly() -> None:
    first_payload = _build_raw_message("first@example.com", "Telefonnummer: 1234")
    second_payload = _build_raw_message("second@example.com", "Telefonnummer: 5678")
    fake_client = FakeImapClient(
        fetch_payloads={
            b"101": first_payload,
            b"202": second_payload,
        }
    )
    cfg = ImapSourceConfig(
        host="imap.example.com",
        port=993,
        username="user@example.com",
        password="plain-text-password",
        mailbox="INBOX",
        use_ssl=True,
    )

    messages = list(iter_imap_messages(cfg, client_factory=lambda _host, _port: fake_client))

    assert len(messages) == 2
    assert fake_client.selected_readonly is True
    assert fake_client.search_calls == [("search", (None, "ALL"))]
    assert messages[0].source_type == "imap"
    assert messages[0].message_identity == "<first@example.com>"
    assert messages[1].message_identity == "<second@example.com>"
    assert messages[0].source_location == "imap://imap.example.com:993/INBOX"
    assert fake_client.logout_called is True


def test_iter_imap_messages_applies_sender_filter_to_search() -> None:
    payload = _build_raw_message("first@example.com", "Telefonnummer: 1234")
    fake_client = FakeImapClient(fetch_payloads={b"101": payload}, search_uids=b"101")

    list(
        iter_imap_messages(
            _cfg(),
            client_factory=lambda _host, _port: fake_client,
            mail_filter=MailFilter(sender=["schule@example.com"]),
        )
    )

    assert fake_client.search_calls == [("search", (None, "FROM", '"schule@example.com"'))]


@pytest.mark.parametrize(
    ("mail_filter", "expected"),
    [
        (
            MailFilter(sender=["a@example.com", "b@example.com"]),
            ("OR", "FROM", '"a@example.com"', "FROM", '"b@example.com"'),
        ),
        (
            MailFilter(subject=["Kontakt", "Anmeldung", "Frage"]),
            ("OR", "SUBJECT", '"Kontakt"', "OR", "SUBJECT", '"Anmeldung"', "SUBJECT", '"Frage"'),
        ),
        (
            MailFilter(subject=["Kontakt"], sender=["schule@example.com"]),
            ("FROM", '"schule@example.com"', "SUBJECT", '"Kontakt"'),
        ),
        # imaplib sends ASCII only; such entries are checked after fetching instead (run_pipeline filters every mail).
        (MailFilter(subject=["Anmeldung für"], sender=["schule@example.com"]), ("FROM", '"schule@example.com"')),
        (MailFilter(subject=["Kontakt", "Rückfrage"]), ("ALL",)),
    ],
)
def test_filter_entries_are_searched_on_the_server(mail_filter: MailFilter, expected: tuple[str, ...]) -> None:
    fake_client = FakeImapClient(search_uids=b"")

    list(iter_imap_messages(_cfg(), client_factory=lambda _host, _port: fake_client, mail_filter=mail_filter))

    assert fake_client.search_calls == [("search", (None, *expected))]


def test_iter_imap_messages_applies_max_age_days_filter_to_search() -> None:
    payload = _build_raw_message("first@example.com", "Telefonnummer: 1234")
    fake_client = FakeImapClient(fetch_payloads={b"101": payload}, search_uids=b"101")
    cfg = ImapSourceConfig(
        host="imap.example.com",
        port=993,
        username="user@example.com",
        password="plain-text-password",
        mailbox="INBOX",
        use_ssl=True,
    )

    _messages = list(
        iter_imap_messages(
            cfg,
            client_factory=lambda _host, _port: fake_client,
            max_age_days=7,
            now_utc=datetime(2026, 10, 3, 12, 0, tzinfo=UTC),
        )
    )

    assert fake_client.search_calls == [("search", (None, "SINCE", "26-Sep-2026"))]


def test_iter_imap_messages_falls_back_to_content_identity() -> None:
    payload = _build_raw_message(None, "Telefonnummer: 1234")
    fake_client = FakeImapClient(fetch_payloads={b"101": payload}, search_uids=b"101")
    cfg = ImapSourceConfig(
        host="imap.example.com",
        port=993,
        username="user@example.com",
        password="plain-text-password",
        mailbox="INBOX",
        use_ssl=True,
    )

    messages = list(iter_imap_messages(cfg, client_factory=lambda _host, _port: fake_client))

    assert len(messages) == 1
    # Content-based like the .eml source, so it stays stable if UIDs change.
    assert messages[0].message_identity.startswith("fallback:")


def test_iter_imap_messages_raises_on_search_error() -> None:
    class SearchErrorClient(FakeImapClient):
        def uid(self, command: str, *args):
            if command.lower() == "search":
                return "NO", [b"search failed"]
            return super().uid(command, *args)

    fake_client = SearchErrorClient(
        fetch_payloads={b"101": _build_raw_message("first@example.com", "Telefonnummer: 1234")}
    )
    cfg = ImapSourceConfig(
        host="imap.example.com",
        port=993,
        username="user@example.com",
        password="plain-text-password",
        mailbox="INBOX",
        use_ssl=True,
    )

    with pytest.raises(OSError, match="IMAP search failed"):
        list(iter_imap_messages(cfg, client_factory=lambda _host, _port: fake_client))


def test_failed_fetch_yields_error_and_continues() -> None:
    class FlakyFetchClient(FakeImapClient):
        def uid(self, command: str, *args):
            if command.lower() == "fetch" and args[0] == b"101":
                return "NO", [b"fetch failed"]
            return super().uid(command, *args)

    fake_client = FlakyFetchClient(fetch_payloads={b"202": _build_raw_message("ok@example.com", "Body")})

    items = list(iter_imap_messages(_cfg(), client_factory=lambda _host, _port: fake_client))


    assert isinstance(items[0], MailReadError)
    assert items[0].message_identity == "uid:101"
    assert items[1].message_identity == "<ok@example.com>"
    # So the error row of uid 202 from an earlier failed fetch can be removed once it is read.
    assert items[1].read_error_identity == "uid:202"


def test_sender_filter_quotes_are_escaped() -> None:
    fake_client = FakeImapClient(fetch_payloads={b"101": _build_raw_message("a@example.com", "x")}, search_uids=b"101")

    list(
        iter_imap_messages(
            _cfg(), client_factory=lambda _host, _port: fake_client, mail_filter=MailFilter(sender=['a"b'])
        )
    )

    assert fake_client.search_calls == [("search", (None, "FROM", '"a\\"b"'))]


def test_fetch_uses_body_peek_so_messages_stay_unread() -> None:
    fake_client = FakeImapClient(fetch_payloads={b"101": _build_raw_message("a@example.com", "x")}, search_uids=b"101")

    list(iter_imap_messages(_cfg(), client_factory=lambda _host, _port: fake_client))

    assert fake_client.fetch_calls == [("fetch", (b"101", "(BODY.PEEK[])"))]


def test_iter_imap_messages_reports_total_before_fetching() -> None:
    fake_client = FakeImapClient(
        fetch_payloads={
            b"101": _build_raw_message("a@example.com", "x"),
            b"202": _build_raw_message("b@example.com", "y"),
        }
    )
    totals: list[int] = []

    items = iter_imap_messages(_cfg(), client_factory=lambda _host, _port: fake_client, on_total=totals.append)
    next(items)

    assert totals == [2]
    items.close()
