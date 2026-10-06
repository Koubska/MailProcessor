"""Shared helpers for RFC822 message content extraction."""

from __future__ import annotations

import hashlib
from email import policy
from email.message import Message
from email.parser import BytesParser
from html.parser import HTMLParser
from typing import Literal

from mailprocessor.models import NormalizedMail
from mailprocessor.parser import normalize_body

_BLOCK_TAGS = {"br", "p", "div", "tr", "li", "h1", "h2", "h3", "h4", "h5", "h6", "table", "blockquote"}


class _HtmlTextExtractor(HTMLParser):
    def __init__(self) -> None:
        super().__init__(convert_charrefs=True)
        self.parts: list[str] = []
        self._skip_depth = 0

    def handle_starttag(self, tag: str, attrs) -> None:
        if tag in {"script", "style"}:
            self._skip_depth += 1
        elif tag in _BLOCK_TAGS:
            self.parts.append("\n")

    def handle_endtag(self, tag: str) -> None:
        if tag in {"script", "style"} and self._skip_depth:
            self._skip_depth -= 1
        elif tag in _BLOCK_TAGS:
            self.parts.append("\n")

    def handle_data(self, data: str) -> None:
        if not self._skip_depth:
            self.parts.append(data)


def html_to_text(html: str) -> str:
    extractor = _HtmlTextExtractor()
    extractor.feed(html)
    extractor.close()
    return "".join(extractor.parts)


def _decode_part(part: Message) -> str | None:
    payload = part.get_payload(decode=True)
    if payload is None:
        raw_payload = part.get_payload()
        return raw_payload if isinstance(raw_payload, str) else None
    charset = part.get_content_charset() or "utf-8"
    try:
        return payload.decode(charset, errors="replace")
    except (LookupError, UnicodeError):
        # Unknown/bogus charset labels must not abort processing; fall back to UTF-8. This includes names
        # codecs knows but that cannot decode text, e.g. "quoted-printable" or "base64" used as a charset.
        return payload.decode("utf-8", errors="replace")


def extract_text_plain_body(message: Message) -> str:
    """Return the first text/plain body, falling back to text extracted from the first text/html part."""
    html_fallback: str | None = None
    parts = message.walk() if message.is_multipart() else [message]
    for part in parts:
        if part.get_content_maintype() == "multipart":
            continue
        if part.get_content_disposition() == "attachment":
            continue
        content_type = part.get_content_type()
        if content_type == "text/plain":
            text = _decode_part(part)
            if text is not None:
                return text
        elif content_type == "text/html" and html_fallback is None:
            html_fallback = _decode_part(part)
    return html_to_text(html_fallback) if html_fallback else ""


def parse_message_bytes(
    raw_bytes: bytes,
    source_type: Literal["imap", "eml"],
    source_location: str,
    origin: str = "",
    read_error_identity: str = "",
) -> NormalizedMail:
    """Parse an RFC822 message into the application model; shared by all sources."""
    message = BytesParser(policy=policy.default).parsebytes(raw_bytes)
    from_raw = str(message.get("From", ""))
    subject = str(message.get("Subject", ""))
    date_raw = str(message.get("Date", ""))
    body_text = normalize_body(extract_text_plain_body(message))
    message_id = message.get("Message-ID")
    return NormalizedMail(
        source_type=source_type,
        source_location=source_location,
        message_identity=build_message_identity(
            str(message_id) if message_id is not None else None, from_raw, subject, date_raw, body_text
        ),
        from_raw=from_raw,
        subject=subject,
        date_raw=date_raw,
        body_text=body_text,
        origin=origin,
        read_error_identity=read_error_identity,
        header_text="\n".join(f"{name}: {value}" for name, value in message.items()),
    )


def build_message_identity(
    message_id_header: str | None,
    from_raw: str,
    subject: str,
    date_raw: str,
    body_text: str,
) -> str:
    """Message-ID if present, otherwise a content hash that is stable across folders, mailboxes and UIDs.

    Changing this formula changes ledger keys and would re-export already processed messages.
    """
    if message_id_header and message_id_header.strip():
        return message_id_header.strip()
    fallback_input = f"{date_raw}\n{from_raw}\n{subject}\n{normalize_body(body_text)}"
    digest = hashlib.sha256(fallback_input.encode("utf-8")).hexdigest()
    return f"fallback:{digest}"
