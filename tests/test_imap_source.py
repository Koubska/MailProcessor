from __future__ import annotations

import imaplib
import ssl
from dataclasses import dataclass, field
from datetime import UTC, datetime

import pytest

from mailprocessor.config import ImapSourceConfig, MailFilter
from mailprocessor.models import MailReadError
from mailprocessor.sources.imap_source import (
    ReadOnlyImapClient,
    check_imap_connection,
    default_client_factory,
    iter_imap_messages,
    mailbox_argument,
)


def _build_raw_message(message_id: str | None, body: str) -> bytes:
    lines = [
        f"Message-ID: <{message_id}>" if message_id else "",
        "From: Max Mustermann <max.mustermann@mail.com>",
        "Subject: Schnuppernachmittag",
        "Date: Mon, 23 Nov 2026 14:00:00 +0100",
        "Content-Type: text/plain; charset=utf-8",
        "",
        body,
    ]
    return "\r\n".join(line for line in lines if line).encode("utf-8")


@dataclass
class FakeImapClient:
    search_uids: bytes = b"101 202"
    fetch_payloads: dict[bytes, bytes] = field(default_factory=dict)
    selected_readonly: bool | None = None
    login_calls: list[tuple[str, str]] = field(default_factory=list)
    search_calls: list[tuple[str, tuple[object, ...]]] = field(default_factory=list)
    fetch_calls: list[tuple[str, tuple[object, ...]]] = field(default_factory=list)
    logout_called: bool = False
    confirms_readonly: bool = True

    def login(self, username: str, password: str):
        self.login_calls.append((username, password))
        return "OK", [b"logged in"]

    def select(self, mailbox: str, readonly: bool = False):
        self.selected_readonly = readonly
        return "OK", [b"2"]

    def uid(self, command: str, *args):
        if command.lower() == "search":
            self.search_calls.append((command, args))
            return "OK", [self.search_uids]
        if command.lower() == "fetch":
            self.fetch_calls.append((command, args))
            uid = args[0]
            payload = self.fetch_payloads[uid]
            return "OK", [(b"1 (RFC822 {123})", payload), b")"]
        raise AssertionError(f"Unexpected uid command: {command}")

    def response(self, code: str):
        if code == "READ-ONLY" and self.confirms_readonly and self.selected_readonly:
            return code, [b""]
        return code, [None]

    def logout(self):
        self.logout_called = True
        return "BYE", [b"logged out"]

    def store(self, *_args, **_kwargs):
        raise AssertionError("store must not be called for readonly IMAP processing")

    def copy(self, *_args, **_kwargs):
        raise AssertionError("copy must not be called for readonly IMAP processing")

    def expunge(self):
        raise AssertionError("expunge must not be called for readonly IMAP processing")


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


def _cfg(**overrides) -> ImapSourceConfig:
    values = dict(host="imap.example.com", port=993, username="user@example.com", password="pw", use_ssl=True)
    values.update(overrides)
    return ImapSourceConfig(**values)


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


def test_login_failure_is_os_error_without_password() -> None:
    class RejectingClient(FakeImapClient):
        def login(self, username: str, password: str):
            raise imaplib.IMAP4.error("AUTHENTICATIONFAILED secret-echo")

    fake_client = RejectingClient()

    with pytest.raises(OSError, match="IMAP login failed") as excinfo:
        list(iter_imap_messages(_cfg(password="S3CRET"), client_factory=lambda _host, _port: fake_client))

    assert "S3CRET" not in str(excinfo.value)
    assert fake_client.logout_called is True


def test_missing_password_is_rejected() -> None:
    with pytest.raises(ValueError, match="password is missing"):
        list(iter_imap_messages(_cfg(password=None), client_factory=lambda _host, _port: FakeImapClient()))


def test_sender_filter_quotes_are_escaped() -> None:
    fake_client = FakeImapClient(fetch_payloads={b"101": _build_raw_message("a@example.com", "x")}, search_uids=b"101")

    list(
        iter_imap_messages(
            _cfg(), client_factory=lambda _host, _port: fake_client, mail_filter=MailFilter(sender=['a"b'])
        )
    )

    assert fake_client.search_calls == [("search", (None, "FROM", '"a\\"b"'))]


def test_default_ssl_client_verifies_certificates(monkeypatch) -> None:
    captured = {}

    def fake_imap4_ssl(host, port, ssl_context=None):
        captured["context"] = ssl_context
        return object()

    monkeypatch.setattr(imaplib, "IMAP4_SSL", fake_imap4_ssl)

    default_client_factory(use_ssl=True)("imap.example.com", 993)

    context = captured["context"]
    assert context.verify_mode == ssl.CERT_REQUIRED
    assert context.check_hostname is True


def test_plain_connection_requires_starttls(monkeypatch) -> None:
    class NoStartTls:
        error = imaplib.IMAP4.error

        def __init__(self, host, port):
            self.shut_down = False

        def starttls(self, ssl_context=None):
            raise imaplib.IMAP4.error("STARTTLS not supported")

        def shutdown(self):
            self.shut_down = True

    monkeypatch.setattr(imaplib, "IMAP4", NoStartTls)

    with pytest.raises(OSError, match="does not support STARTTLS"):
        default_client_factory(use_ssl=False)("imap.example.com", 143)


def test_fetch_uses_body_peek_so_messages_stay_unread() -> None:
    fake_client = FakeImapClient(fetch_payloads={b"101": _build_raw_message("a@example.com", "x")}, search_uids=b"101")

    list(iter_imap_messages(_cfg(), client_factory=lambda _host, _port: fake_client))

    assert fake_client.fetch_calls == [("fetch", (b"101", "(BODY.PEEK[])"))]


def test_aborts_before_fetching_when_server_does_not_confirm_read_only() -> None:
    fake_client = FakeImapClient(
        fetch_payloads={b"101": _build_raw_message("a@example.com", "x")}, search_uids=b"101", confirms_readonly=False
    )

    with pytest.raises(OSError, match="did not confirm read-only"):
        list(iter_imap_messages(_cfg(), client_factory=lambda _host, _port: fake_client))

    assert fake_client.search_calls == []
    assert fake_client.fetch_calls == []
    assert fake_client.logout_called is True


def test_read_only_client_exposes_no_mutating_commands() -> None:
    public = {name for name in dir(ReadOnlyImapClient) if not name.startswith("_")}

    assert public == {"login", "examine", "uid_search", "uid_fetch", "logout"}


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


def test_check_imap_connection_opens_mailbox_read_only_without_fetching() -> None:
    fake_client = FakeImapClient()

    count = check_imap_connection(_cfg(), client_factory=lambda _host, _port: fake_client)

    assert count == 2
    assert fake_client.selected_readonly is True
    assert fake_client.search_calls == [] and fake_client.fetch_calls == []
    assert fake_client.logout_called


def test_check_imap_connection_reports_login_failure_and_logs_out() -> None:
    class RejectingClient(FakeImapClient):
        def login(self, username: str, password: str):
            return "NO", [b"denied"]

    fake_client = RejectingClient()
    with pytest.raises(OSError, match="login failed"):
        check_imap_connection(_cfg(), client_factory=lambda _host, _port: fake_client)
    assert fake_client.logout_called


@pytest.mark.parametrize(
    ("mailbox", "argument"),
    [
        ("INBOX", '"INBOX"'),
        ("Sent Items", '"Sent Items"'),
        ("Anfragen/Schüler", '"Anfragen/Sch&APw-ler"'),
        ("Entwürfe", '"Entw&APw-rfe"'),
        ("A & B", '"A &- B"'),
        ("日本語", '"&ZeVnLIqe-"'),  # RFC 3501's example
        ("Entw&APw-rfe", '"Entw&APw-rfe"'),  # already encoded, e.g. copied from another mail program
        ('Say "hi"', '"Say \\"hi\\""'),
    ],
)
def test_mailbox_argument_is_quoted_modified_utf7(mailbox: str, argument: str) -> None:
    assert mailbox_argument(mailbox) == argument


def test_examine_sends_the_encoded_mailbox_name() -> None:
    selected: list[str] = []

    class RecordingClient(FakeImapClient):
        def select(self, mailbox: str, readonly: bool = False):
            selected.append(mailbox)
            return super().select(mailbox, readonly=readonly)

    ReadOnlyImapClient(RecordingClient()).examine("Anfragen Schüler")

    assert selected == ['"Anfragen Sch&APw-ler"']
