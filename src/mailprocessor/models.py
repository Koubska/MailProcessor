"""Domain models."""

from __future__ import annotations

from dataclasses import dataclass
from typing import Literal


@dataclass(frozen=True)
class NormalizedMail:
    source_type: Literal["imap", "eml"]
    source_location: str
    message_identity: str
    from_raw: str
    subject: str
    date_raw: str
    body_text: str
    # Where a user can find the mail: the .eml file name or "IMAP uid 123". Used for logs only.
    origin: str = ""
    # All header lines ("Name: value"), so rules can also match e.g. the From header.
    header_text: str = ""
    # The identity a MailReadError for the same file or IMAP uid gets ("file:a.eml", "uid:123"), so the error
    # row of an earlier failed read can be removed once the mail is read.
    read_error_identity: str = ""

    @property
    def display_name(self) -> str:
        if not self.origin:
            return self.message_identity
        # A content-hash fallback identity means nothing to users; the file name / uid is enough.
        if self.message_identity.startswith("fallback:"):
            return self.origin
        return f"{self.origin} ({self.message_identity})"


@dataclass(frozen=True)
class MailReadError:
    """A single message that could not be read; reported instead of aborting the batch."""

    source_type: Literal["imap", "eml"]
    source_location: str
    message_identity: str
    reason: str
    # Where a user can find the mail: the .eml file name or "IMAP uid 123".
    origin: str = ""

    @property
    def display_name(self) -> str:
        return self.origin or self.message_identity
