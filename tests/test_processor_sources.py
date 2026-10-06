"""Reading the source (IMAP or .eml folder), the age limit, and the mail's date in the rows."""

from datetime import UTC, datetime, timedelta
from pathlib import Path

import pytest
from pipeline_helpers import GOOD_BODY, _build_config, _build_rules, _data_rows, _write_eml, _write_eml_with_date

from mailprocessor.config import (
    AppConfig,
    AppSection,
    ImapSourceConfig,
    MailFilter,
    SourceConfig,
)
from mailprocessor.excel_writer import RECEIVED_COLUMN, TRANSFERRED_COLUMN
from mailprocessor.models import NormalizedMail
from mailprocessor.processor import iter_source_messages, run_pipeline


def test_iter_source_messages_uses_imap_source(monkeypatch) -> None:
    config = AppConfig(
        app=AppSection(
            log_level="INFO",
            sqlite_path="./data/ledger.db",
            output_xlsx="./out/mail_export.xlsx",
            sheet_data="daten",
            sheet_errors="fehler",
            dry_run=False,
            max_messages=0,
            max_age_days=0,
        ),
        source=SourceConfig(
            type="imap",
            imap=ImapSourceConfig(
                host="imap.example.com",
                port=993,
                username="user@example.com",
                password="plain-text-password",
                mailbox="INBOX",
                use_ssl=True,
            ),
        ),
        filter=MailFilter(sender=["max.mustermann@mail.com"]),
    )

    expected = [
        NormalizedMail(
            source_type="imap",
            source_location="imap://imap.example.com:993/INBOX",
            message_identity="uid:123",
            from_raw="Max Mustermann <max.mustermann@mail.com>",
            subject="Test",
            date_raw="Mon, 23 Nov 2026 14:00:00 +0100",
            body_text="Telefonnummer: 1234",
        )
    ]

    def fake_iter_imap_messages(_imap_cfg, max_age_days: int, now_utc: datetime, on_total=None, mail_filter=None):
        assert max_age_days == 0
        assert now_utc.tzinfo is UTC
        # The filter goes to the server too, so non-matching mails are not even downloaded.
        assert mail_filter == MailFilter(sender=["max.mustermann@mail.com"])
        yield from expected

    monkeypatch.setattr("mailprocessor.processor.iter_imap_messages", fake_iter_imap_messages)

    result = list(iter_source_messages(config))

    assert result == expected


def test_iter_source_messages_filters_eml_messages_by_max_age(tmp_path: Path) -> None:
    inbox = tmp_path / "inbox"
    inbox.mkdir()
    _write_eml_with_date(
        inbox / "recent.eml",
        [
            "Von: Max Mustermann <max.mustermann@mail.com>",
            "hiermit möchte ich mein Kind Jan Must+ für folgendes Angebot anmelden:",
            "Tag: Mo, 23.11.2026 (14:15-16:00 Uhr)",
            "Angebot: Experimente",
            "Telefonnummer: 1234 567890",
        ],
        "recent@example.com",
        "Thu, 01 Oct 2026 12:00:00 +0000",
    )
    _write_eml_with_date(
        inbox / "old.eml",
        [
            "Von: Max Mustermann <max.mustermann@mail.com>",
            "hiermit möchte ich mein Kind Jan Must+ für folgendes Angebot anmelden:",
            "Tag: Mo, 23.11.2026 (14:15-16:00 Uhr)",
            "Angebot: Experimente",
            "Telefonnummer: 1234 567890",
        ],
        "old@example.com",
        "Tue, 01 Sep 2026 12:00:00 +0000",
    )
    config = _build_config(tmp_path)
    config.app.max_age_days = 7

    messages = list(iter_source_messages(config, now_utc=datetime(2026, 10, 3, 12, 0, tzinfo=UTC)))

    assert [message.message_identity for message in messages] == ["<recent@example.com>"]


def test_max_age_keeps_messages_without_date(tmp_path: Path) -> None:
    inbox = tmp_path / "inbox"
    inbox.mkdir()
    _write_eml_with_date(inbox / "nodate.eml", GOOD_BODY, "nodate@example.com", "not a date")
    config = _build_config(tmp_path)
    config.app.max_age_days = 7

    messages = list(iter_source_messages(config, now_utc=datetime(2026, 10, 3, 12, 0, tzinfo=UTC)))

    assert [m.message_identity for m in messages] == ["<nodate@example.com>"]


def test_rows_get_the_mail_date_and_the_transfer_time(tmp_path: Path) -> None:
    inbox = tmp_path / "inbox"
    inbox.mkdir()
    _write_eml_with_date(inbox / "a.eml", GOOD_BODY, "a@example.com", "Mon, 23 Nov 2026 14:00:00 +0100")
    before = datetime.now().replace(microsecond=0)

    run_pipeline(_build_config(tmp_path), _build_rules())

    row = _data_rows(tmp_path)[0]
    expected = datetime(2026, 11, 23, 13, 0, tzinfo=UTC).astimezone().replace(tzinfo=None)
    assert row[RECEIVED_COLUMN] == expected  # shown in local time, as people read it
    assert before <= row[TRANSFERRED_COLUMN] <= datetime.now() + timedelta(seconds=1)


def test_mail_without_date_is_still_transferred(tmp_path: Path) -> None:
    inbox = tmp_path / "inbox"
    inbox.mkdir()
    _write_eml_with_date(inbox / "a.eml", GOOD_BODY, "a@example.com", "kein Datum")

    summary = run_pipeline(_build_config(tmp_path), _build_rules())

    assert summary.processed == 1
    assert _data_rows(tmp_path)[0][RECEIVED_COLUMN] is None


@pytest.mark.parametrize("max_age_days", [0, 30])
def test_mail_with_a_date_out_of_range_does_not_abort_the_run(tmp_path: Path, max_age_days: int) -> None:
    # In UTC this is the year 10000, which datetime cannot represent (OverflowError). One such mail,
    # e.g. spam, must not stop every run; it is treated like a mail without a usable date.
    inbox = tmp_path / "inbox"
    inbox.mkdir()
    _write_eml_with_date(inbox / "a.eml", GOOD_BODY, "a@example.com", "Fri, 31 Dec 9999 23:59:59 -2359")
    _write_eml(inbox / "b.eml", GOOD_BODY, "b@example.com")
    config = _build_config(tmp_path)
    config.app.max_age_days = max_age_days

    summary = run_pipeline(config, _build_rules())

    # Like a mail without a date: transferred (max_age_days cannot judge it), "Eingegangen am" stays empty.
    assert (summary.processed, summary.failed) == (2, 0)
    assert [row[RECEIVED_COLUMN] is None for row in _data_rows(tmp_path)] == [True, False]
