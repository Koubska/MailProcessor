"""SQLite-backed idempotency ledger."""

from __future__ import annotations

import hashlib
import sqlite3
from dataclasses import dataclass
from datetime import UTC, datetime
from pathlib import Path

from mailprocessor.errors import LedgerUnreadableError, RunInProgressError

# How long a writing run waits for another one to finish before it gives up.
LOCK_TIMEOUT_SECONDS = 2

_SCHEMA = """
CREATE TABLE IF NOT EXISTS processed_messages (
    source_type TEXT NOT NULL,
    source_location TEXT NOT NULL,
    message_identity TEXT NOT NULL,
    content_hash TEXT NOT NULL,
    status TEXT NOT NULL,
    processed_at TEXT NOT NULL,
    error_reason TEXT,
    PRIMARY KEY (source_type, source_location, message_identity, content_hash)
)
"""


_INDEX = "CREATE INDEX IF NOT EXISTS idx_identity_hash ON processed_messages (message_identity, content_hash)"


@dataclass(frozen=True)
class LedgerKey:
    source_type: str
    source_location: str
    message_identity: str
    content_hash: str


def compute_content_hash(body_text: str) -> str:
    return hashlib.sha256(body_text.encode("utf-8")).hexdigest()


class Ledger:
    """Processing history. Writes are held in one transaction until `commit()`.

    Opened for writing, it holds the database's write lock until `commit()` or `close()`, so two runs never
    write at the same time; the second one raises `RunInProgressError`.

    The pipeline commits only after the Excel workbook was saved, so a crash or a locked
    workbook leaves ledger and workbook consistent (the run simply repeats next time).
    With `read_only=True` (dry runs) the database file is never created or modified.
    """

    def __init__(self, db_path: Path, *, read_only: bool = False) -> None:
        self.read_only = read_only
        if read_only:
            if db_path.exists():
                self._connection = sqlite3.connect(f"{db_path.resolve().as_uri()}?mode=ro", uri=True)
            else:
                self._connection = sqlite3.connect(":memory:")
                self._connection.execute(_SCHEMA)
        else:
            db_path.parent.mkdir(parents=True, exist_ok=True)
            self._connection = sqlite3.connect(db_path, timeout=LOCK_TIMEOUT_SECONDS)
        try:
            if read_only:
                # Reads the file's header, so damage shows here and not in the middle of a run.
                self._connection.execute("SELECT COUNT(*) FROM sqlite_master").fetchone()
            else:
                self._connection.execute("PRAGMA journal_mode = WAL")
                self._connection.execute(_SCHEMA)
                self._connection.execute(_INDEX)
                self._connection.commit()
                # Take the write lock now and keep it until commit/close: only one writing run at a time.
                self._connection.execute("BEGIN IMMEDIATE")
        except sqlite3.OperationalError as exc:
            self._connection.close()
            if "locked" not in str(exc):
                raise  # e.g. the file cannot be opened at all; not about its content
            raise RunInProgressError(
                f"Another run is using {db_path.name} right now (a second window or a scheduled run). "
                "Wait until it has finished and try again."
            ) from None
        except sqlite3.DatabaseError as exc:
            self._connection.close()
            raise LedgerUnreadableError(
                f"The ledger {db_path} (the list of processed mails) is damaged or not a database ({exc}). "
                "Use 'Export everything again' to replace it, or restore the file from a backup."
            ) from None

    def __enter__(self) -> Ledger:
        return self

    def __exit__(self, *_exc_info) -> None:
        self.close()

    def close(self) -> None:
        """Close without committing; uncommitted writes are rolled back."""
        self._connection.close()

    def commit(self) -> None:
        self._connection.commit()

    def is_already_processed(self, key: LedgerKey) -> bool:
        """True if this message (identity + content) was exported before, from any source location.

        Location is deliberately ignored: moving the .eml folder, moving the bundle, or renaming the
        IMAP mailbox must not export the same messages again.
        """
        cursor = self._connection.execute(
            """
            SELECT 1
            FROM processed_messages
            WHERE message_identity = ?
              AND content_hash = ?
              AND status = 'processed'
            LIMIT 1
            """,
            (key.message_identity, key.content_hash),
        )
        return cursor.fetchone() is not None

    def clear(self) -> None:
        """Forget all processed and failed messages (within the open transaction), for a full re-export."""
        if self.read_only:
            raise RuntimeError("Ledger was opened read-only")
        self._connection.execute("DELETE FROM processed_messages")

    def count_processed(self) -> int:
        cursor = self._connection.execute("SELECT COUNT(*) FROM processed_messages WHERE status = 'processed'")
        return int(cursor.fetchone()[0])

    def _upsert_status(self, key: LedgerKey, status: str, error_reason: str | None) -> None:
        if self.read_only:
            raise RuntimeError("Ledger was opened read-only")
        self._connection.execute(
            """
            INSERT INTO processed_messages (
                source_type, source_location, message_identity, content_hash, status, processed_at, error_reason
            ) VALUES (?, ?, ?, ?, ?, ?, ?)
            ON CONFLICT(source_type, source_location, message_identity, content_hash)
            DO UPDATE SET
                status = excluded.status,
                processed_at = excluded.processed_at,
                error_reason = excluded.error_reason
            """,
            (
                key.source_type,
                key.source_location,
                key.message_identity,
                key.content_hash,
                status,
                datetime.now(UTC).isoformat(),
                error_reason,
            ),
        )

    def mark_processed(self, key: LedgerKey) -> None:
        self._upsert_status(key, "processed", None)

    def mark_failed(self, key: LedgerKey, reason: str) -> None:
        self._upsert_status(key, "failed", reason)
