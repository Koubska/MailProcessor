"""What a run reports back to the CLI and the GUI."""

from __future__ import annotations

from dataclasses import dataclass

# Details are kept for at most this many failed messages per run (they hold the mail text in memory).
MAX_PROBLEM_DETAILS = 500


@dataclass(frozen=True)
class Problem:
    """A message that failed in this run, for the GUI's problem list. Never logged or written anywhere."""

    name: str  # display name: file name or IMAP uid
    reason: str
    missing: tuple[str, ...] = ()
    # The mail text and headers as parsed, so the GUI can show the mail; None if it could not be read.
    body: str | None = None
    header_text: str = ""
    # The profile that came closest, if there are several; its fields are the missing ones.
    profile: str = ""


@dataclass(frozen=True)
class RunSummary:
    seen: int
    processed: int
    skipped: int
    failed: int
    # True if the run was stopped early; everything handled until then was saved.
    cancelled: bool = False
    # Mails left out because subject or sender did not match the filter; not counted in `seen`.
    filtered: int = 0
    problems: tuple[Problem, ...] = ()
    # New rows per profile, in the order of the profiles; empty if there is only one profile.
    per_profile: tuple[tuple[str, int], ...] = ()
