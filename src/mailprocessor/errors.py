"""Expected failures with a plain-language explanation in the GUI.

They subclass ValueError/OSError, so the CLI keeps handling them like before.
"""

from __future__ import annotations


class MailFolderNotFoundError(FileNotFoundError):
    """The configured .eml folder does not exist."""


class WorkbookLockedError(OSError):
    """The Excel file cannot be written, typically because it is open in Excel."""


class WorkbookUnreadableError(ValueError):
    """The output file exists but cannot be opened as a workbook: damaged, not .xlsx, or protected with a password."""


class SheetHeaderError(ValueError):
    """A sheet the app writes to has data but no header row, so its columns cannot be found."""


class ImapLoginError(OSError):
    """The IMAP server rejected username or password."""


class MissingPasswordError(ValueError):
    """No IMAP password was given."""


class RunInProgressError(OSError):
    """Another run (a second window or a scheduled run) is writing the same ledger and workbook right now."""


class LedgerUnreadableError(ValueError):
    """The ledger file (the list of processed mails) exists but is damaged or not a database."""
