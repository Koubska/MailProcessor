"""The IMAP connection: TLS, login, and the read-only client that only examines and fetches."""

from __future__ import annotations

import imaplib
import ssl

import pytest
from imap_helpers import FakeImapClient, _build_raw_message, _cfg

from mailprocessor.sources.imap_source import (
    ReadOnlyImapClient,
    check_imap_connection,
    default_client_factory,
    iter_imap_messages,
    mailbox_argument,
)


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


def test_non_ascii_password_is_sent_as_utf8_plain_and_rejection_is_a_login_error() -> None:
    sent: list[bytes] = []

    class PlainRejectingClient(FakeImapClient):
        def authenticate(self, mechanism: str, authobject):
            assert mechanism == "PLAIN"
            sent.append(authobject(b""))
            raise imaplib.IMAP4.error("AUTHENTICATE failed")  # what imaplib raises for NO

    fake_client = PlainRejectingClient()
    config = _cfg(password="Schlüssel§")
    with pytest.raises(OSError, match="login failed"):
        check_imap_connection(config, client_factory=lambda _host, _port: fake_client)

    assert sent == [f"\0{config.username}\0Schlüssel§".encode()]
    assert fake_client.logout_called
