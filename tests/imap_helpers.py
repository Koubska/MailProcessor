"""A scripted IMAP client and test messages for the IMAP source tests."""

from __future__ import annotations

from dataclasses import dataclass, field

from mailprocessor.config import ImapSourceConfig


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


def _cfg(**overrides) -> ImapSourceConfig:
    values = dict(host="imap.example.com", port=993, username="user@example.com", password="pw", use_ssl=True)
    values.update(overrides)
    return ImapSourceConfig(**values)
