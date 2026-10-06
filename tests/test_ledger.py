from pathlib import Path

import pytest

from mailprocessor.errors import LedgerUnreadableError
from mailprocessor.ledger import Ledger, LedgerKey, compute_content_hash


def key_for(body: str) -> LedgerKey:
    return LedgerKey(
        source_type="eml",
        source_location="/tmp/inbox",
        message_identity="msg-1",
        content_hash=compute_content_hash(body),
    )


def test_processed_key_is_skipped_on_subsequent_checks(tmp_path: Path) -> None:
    db_path = tmp_path / "ledger.db"
    key = key_for("body")

    with Ledger(db_path) as ledger:
        assert ledger.is_already_processed(key) is False
        ledger.mark_processed(key)
        ledger.commit()

    with Ledger(db_path) as ledger:
        assert ledger.is_already_processed(key) is True


def test_failed_key_is_not_treated_as_processed(tmp_path: Path) -> None:
    key = key_for("body")

    with Ledger(tmp_path / "ledger.db") as ledger:
        ledger.mark_failed(key, "parse failed")
        assert ledger.is_already_processed(key) is False


def test_reprocessed_message_with_new_hash_is_not_skipped(tmp_path: Path) -> None:
    key_v1 = key_for("body v1")
    key_v2 = key_for("body v2")

    with Ledger(tmp_path / "ledger.db") as ledger:
        ledger.mark_processed(key_v1)
        assert ledger.is_already_processed(key_v1) is True
        assert ledger.is_already_processed(key_v2) is False


def test_uncommitted_writes_are_rolled_back(tmp_path: Path) -> None:
    db_path = tmp_path / "ledger.db"
    key = key_for("body")

    with Ledger(db_path) as ledger:
        ledger.mark_processed(key)

    with Ledger(db_path) as ledger:
        assert ledger.is_already_processed(key) is False


def test_read_only_ledger_does_not_create_database(tmp_path: Path) -> None:
    db_path = tmp_path / "data" / "ledger.db"

    with Ledger(db_path, read_only=True) as ledger:
        assert ledger.is_already_processed(key_for("body")) is False

    assert not db_path.exists()
    assert not db_path.parent.exists()


def test_read_only_ledger_reads_existing_database(tmp_path: Path) -> None:
    db_path = tmp_path / "ledger.db"
    key = key_for("body")
    with Ledger(db_path) as ledger:
        ledger.mark_processed(key)
        ledger.commit()

    with Ledger(db_path, read_only=True) as ledger:
        assert ledger.is_already_processed(key) is True


@pytest.mark.parametrize("read_only", [False, True])
@pytest.mark.parametrize("content", [b"not a database" * 100, b"SQLite format 3\x00" + bytes(30)])
def test_a_damaged_ledger_file_is_reported_clearly(tmp_path: Path, read_only: bool, content: bytes) -> None:
    path = tmp_path / "ledger.db"
    path.write_bytes(content)

    with pytest.raises(LedgerUnreadableError, match="ledger.db"):
        Ledger(path, read_only=read_only)
    assert path.read_bytes() == content
