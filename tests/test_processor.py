import logging
import threading
from datetime import UTC, datetime, timedelta
from pathlib import Path

import pytest
from openpyxl import load_workbook

import mailprocessor.processor as processor
import mailprocessor.sources.eml_folder_source as eml_source
from mailprocessor.config import (
    AppConfig,
    AppSection,
    EmlSourceConfig,
    FieldRule,
    ImapSourceConfig,
    MailFilter,
    ParsingRules,
    Profile,
    SourceConfig,
)
from mailprocessor.errors import RunInProgressError
from mailprocessor.excel_writer import ERROR_COLUMNS, RECEIVED_COLUMN, TRANSFERRED_COLUMN
from mailprocessor.ledger import Ledger
from mailprocessor.models import MailReadError, NormalizedMail
from mailprocessor.processor import RunSummary, iter_source_messages, run_pipeline, start_over


def _build_config(tmp_path: Path, max_messages: int = 0) -> AppConfig:
    return AppConfig(
        app=AppSection(
            log_level="INFO",
            sqlite_path=str(tmp_path / "data" / "ledger.db"),
            output_xlsx=str(tmp_path / "out" / "mail_export.xlsx"),
            sheet_data="daten",
            sheet_errors="fehler",
            dry_run=False,
            max_messages=max_messages,
            max_age_days=0,
        ),
        source=SourceConfig(
            type="eml",
            eml=EmlSourceConfig(folder=str(tmp_path / "inbox"), glob="*.eml"),
        ),
    )


def _build_rules() -> ParsingRules:
    return ParsingRules(
        fields=[
            FieldRule(column="Mail-Adresse", pattern=r"(?im)^\s*Von:\s*.*?<([^>]+)>\s*$", required=True),
            FieldRule(column="Name", pattern=r"(?im)mein\s+Kind\s+(.+?)\s+f[uü]r\s+folgendes\s+Angebot", required=True),
            FieldRule(column="Kurs", pattern=r"(?im)^\s*Angebot:\s*(.+?)\s*$", required=True),
            FieldRule(column="Zeit", pattern=r"(?im)^\s*Tag:\s*(.+?)\s*$", required=True),
            FieldRule(column="Telefonnummer", pattern=r"(?im)^\s*Telefonnummer:\s*(.+?)\s*$", required=True),
        ]
    )


def _write_eml(path: Path, body_lines: list[str], message_id: str, subject: str = "Schnuppernachmittag") -> None:
    path.write_text(
        "\n".join(
            [
                f"Message-ID: <{message_id}>",
                "From: Max Mustermann <max.mustermann@mail.com>",
                f"Subject: {subject}",
                "Date: Mon, 23 Nov 2026 14:00:00 +0100",
                "Content-Type: text/plain; charset=utf-8",
                "",
                *body_lines,
            ]
        ),
        encoding="utf-8",
    )


def _write_eml_with_date(path: Path, body_lines: list[str], message_id: str, date_header: str) -> None:
    path.write_text(
        "\n".join(
            [
                f"Message-ID: <{message_id}>",
                "From: Max Mustermann <max.mustermann@mail.com>",
                "Subject: Schnuppernachmittag",
                f"Date: {date_header}",
                "Content-Type: text/plain; charset=utf-8",
                "",
                *body_lines,
            ]
        ),
        encoding="utf-8",
    )


def test_run_pipeline_writes_data_and_error_sheets(tmp_path: Path) -> None:
    inbox = tmp_path / "inbox"
    inbox.mkdir()
    _write_eml(
        inbox / "good.eml",
        [
            "Von: Max Mustermann <max.mustermann@mail.com>",
            "hiermit möchte ich mein Kind Jan Must+ für folgendes Angebot anmelden:",
            "Tag: Mo, 23.11.2026 (14:15-16:00 Uhr)",
            "Angebot: Experimente",
            "Telefonnummer: 1234 567890",
        ],
        "good@example.com",
    )
    _write_eml(
        inbox / "bad.eml",
        [
            "Von: Max Mustermann <max.mustermann@mail.com>",
            "hiermit möchte ich mein Kind Jan Must+ für folgendes Angebot anmelden:",
            "Tag: Mo, 23.11.2026 (14:15-16:00 Uhr)",
            "Angebot: Experimente",
        ],
        "bad@example.com",
    )

    summary = run_pipeline(_build_config(tmp_path), _build_rules())

    assert summary.seen == 2
    assert summary.processed == 1
    assert summary.failed == 1
    assert summary.skipped == 0

    workbook = load_workbook(tmp_path / "out" / "mail_export.xlsx")
    assert workbook["daten"].max_row == 2
    assert workbook["fehler"].max_row == 2


def test_run_pipeline_is_idempotent_for_successful_messages(tmp_path: Path) -> None:
    inbox = tmp_path / "inbox"
    inbox.mkdir()
    _write_eml(
        inbox / "good.eml",
        [
            "Von: Max Mustermann <max.mustermann@mail.com>",
            "hiermit möchte ich mein Kind Jan Must+ für folgendes Angebot anmelden:",
            "Tag: Mo, 23.11.2026 (14:15-16:00 Uhr)",
            "Angebot: Experimente",
            "Telefonnummer: 1234 567890",
        ],
        "good@example.com",
    )
    config = _build_config(tmp_path)
    rules = _build_rules()

    first = run_pipeline(config, rules)
    second = run_pipeline(config, rules)

    assert first.processed == 1
    assert second.skipped == 1
    assert second.processed == 0

    workbook = load_workbook(tmp_path / "out" / "mail_export.xlsx")
    assert workbook["daten"].max_row == 2


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


GOOD_BODY = [
    "Von: Max Mustermann <max.mustermann@mail.com>",
    "hiermit möchte ich mein Kind Jan Must+ für folgendes Angebot anmelden:",
    "Tag: Mo, 23.11.2026 (14:15-16:00 Uhr)",
    "Angebot: Experimente",
    "Telefonnummer: 1234 567890",
]


def test_dry_run_writes_neither_workbook_nor_ledger(tmp_path: Path) -> None:
    inbox = tmp_path / "inbox"
    inbox.mkdir()
    _write_eml(inbox / "good.eml", GOOD_BODY, "good@example.com")
    config = _build_config(tmp_path)
    config.app.dry_run = True

    summary = run_pipeline(config, _build_rules())

    assert summary.processed == 1
    assert not (tmp_path / "out").exists()
    assert not (tmp_path / "data").exists()


def test_unreadable_message_does_not_abort_batch(tmp_path: Path) -> None:
    inbox = tmp_path / "inbox"
    inbox.mkdir()
    _write_eml(inbox / "a-good.eml", GOOD_BODY, "good@example.com")
    (inbox / "b-bogus-charset.eml").write_bytes(
        b"Message-ID: <bogus@example.com>\nContent-Type: text/plain; charset=x-unknown-8bit\n\n"
        + "\n".join(GOOD_BODY).encode("utf-8")
    )
    _write_eml(inbox / "c-good.eml", GOOD_BODY, "good2@example.com")

    summary = run_pipeline(_build_config(tmp_path), _build_rules())

    assert summary.processed == 3
    assert summary.failed == 0


def test_eml_read_errors_are_reported_and_batch_continues(tmp_path: Path, monkeypatch) -> None:
    inbox = tmp_path / "inbox"
    inbox.mkdir()
    _write_eml(inbox / "a-broken.eml", GOOD_BODY, "broken@example.com")
    _write_eml(inbox / "b-good.eml", GOOD_BODY, "good@example.com")

    original = eml_source.parse_eml_file

    def flaky_parse(path: Path, source_location: str):
        if path.name == "a-broken.eml":
            raise OSError("disk error")
        return original(path, source_location)

    monkeypatch.setattr(eml_source, "parse_eml_file", flaky_parse)

    summary = run_pipeline(_build_config(tmp_path), _build_rules())

    assert (summary.processed, summary.failed) == (1, 1)
    errors = load_workbook(tmp_path / "out" / "mail_export.xlsx")["fehler"]
    row = dict(zip(ERROR_COLUMNS, next(errors.iter_rows(min_row=2, values_only=True)), strict=True))
    assert row["E-Mail"] == "a-broken.eml"
    assert row["Grund"].startswith("E-Mail konnte nicht gelesen werden")
    assert "disk error" in row["Grund"]


def test_max_messages_counts_only_new_messages(tmp_path: Path) -> None:
    inbox = tmp_path / "inbox"
    inbox.mkdir()
    for index in range(5):
        _write_eml(inbox / f"mail-{index}.eml", GOOD_BODY, f"m{index}@example.com")
    config = _build_config(tmp_path, max_messages=2)
    rules = _build_rules()

    results = [run_pipeline(config, rules) for _ in range(3)]

    assert [r.processed for r in results] == [2, 2, 1]
    assert load_workbook(tmp_path / "out" / "mail_export.xlsx")["daten"].max_row == 6


def test_fixed_failure_removes_its_error_row(tmp_path: Path) -> None:
    inbox = tmp_path / "inbox"
    inbox.mkdir()
    _write_eml(inbox / "mail.eml", GOOD_BODY, "m@example.com")
    config = _build_config(tmp_path)
    strict_rules = ParsingRules(
        fields=[*_build_rules().profiles[0].fields, FieldRule(column="Extra", pattern=r"^Extra:(.+)$", required=True)]
    )
    run_pipeline(config, strict_rules)
    assert load_workbook(tmp_path / "out" / "mail_export.xlsx")["fehler"].max_row == 2

    relaxed_rules = ParsingRules(
        fields=[*_build_rules().profiles[0].fields, FieldRule(column="Extra", pattern=r"^Extra:(.+)$", required=False)]
    )
    summary = run_pipeline(config, relaxed_rules)

    workbook = load_workbook(tmp_path / "out" / "mail_export.xlsx")
    assert summary.processed == 1
    assert workbook["daten"].max_row == 2
    assert workbook["fehler"].max_row == 1


def test_locked_workbook_leaves_ledger_untouched(tmp_path: Path, monkeypatch) -> None:
    inbox = tmp_path / "inbox"
    inbox.mkdir()
    _write_eml(inbox / "mail.eml", GOOD_BODY, "m@example.com")
    config = _build_config(tmp_path)

    def failing_save(self) -> None:
        raise OSError("locked")

    monkeypatch.setattr("mailprocessor.excel_writer.ExcelOutput.save", failing_save)
    with pytest.raises(OSError, match="locked"):
        run_pipeline(config, _build_rules())
    monkeypatch.undo()

    summary = run_pipeline(config, _build_rules())

    assert summary.processed == 1
    assert load_workbook(tmp_path / "out" / "mail_export.xlsx")["daten"].max_row == 2


def test_missing_eml_folder_is_an_error(tmp_path: Path) -> None:
    with pytest.raises(FileNotFoundError, match="EML folder does not exist"):
        run_pipeline(_build_config(tmp_path), _build_rules())


def test_max_age_keeps_messages_without_date(tmp_path: Path) -> None:
    inbox = tmp_path / "inbox"
    inbox.mkdir()
    _write_eml_with_date(inbox / "nodate.eml", GOOD_BODY, "nodate@example.com", "not a date")
    config = _build_config(tmp_path)
    config.app.max_age_days = 7

    messages = list(iter_source_messages(config, now_utc=datetime(2026, 10, 3, 12, 0, tzinfo=UTC)))

    assert [m.message_identity for m in messages] == ["<nodate@example.com>"]


def _snapshot(folder: Path) -> dict[str, tuple[bytes, int, int]]:
    return {
        path.name: (path.read_bytes(), path.stat().st_mtime_ns, path.stat().st_mode)
        for path in sorted(folder.iterdir())
    }


def test_eml_files_remain_completely_unchanged(tmp_path: Path) -> None:
    inbox = tmp_path / "inbox"
    inbox.mkdir()
    _write_eml(inbox / "good.eml", GOOD_BODY, "good@example.com")
    _write_eml(inbox / "bad.eml", GOOD_BODY[:2], "bad@example.com")
    before = _snapshot(inbox)

    run_pipeline(_build_config(tmp_path), _build_rules())
    run_pipeline(_build_config(tmp_path), _build_rules())

    assert _snapshot(inbox) == before


def test_moved_mail_folder_is_not_exported_again(tmp_path: Path) -> None:
    inbox = tmp_path / "inbox"
    inbox.mkdir()
    _write_eml(inbox / "good.eml", GOOD_BODY, "good@example.com")
    (inbox / "no-id.eml").write_text("From: a@example.com\n\n" + "\n".join(GOOD_BODY), encoding="utf-8")
    config = _build_config(tmp_path)
    assert run_pipeline(config, _build_rules()).processed == 2

    inbox.rename(tmp_path / "moved")
    assert config.source.eml is not None
    config.source.eml.folder = str(tmp_path / "moved")
    second = run_pipeline(config, _build_rules())

    assert (second.processed, second.skipped) == (0, 2)
    assert load_workbook(tmp_path / "out" / "mail_export.xlsx")["daten"].max_row == 3


def test_outputs_inside_mail_folder_are_rejected(tmp_path: Path) -> None:
    inbox = tmp_path / "inbox"
    inbox.mkdir()
    config = _build_config(tmp_path)
    config.app.output_xlsx = str(inbox / "export.xlsx")

    with pytest.raises(ValueError, match="must not be inside the mail folder"):
        run_pipeline(config, _build_rules())
    assert list(inbox.iterdir()) == []


def test_run_logs_progress_and_failure_reasons_without_mail_content(tmp_path: Path, caplog) -> None:
    inbox = tmp_path / "inbox"
    inbox.mkdir()
    _write_eml(inbox / "good.eml", GOOD_BODY, "good@example.com")
    _write_eml(inbox / "missing-phone.eml", GOOD_BODY[:4], "bad@example.com")
    config = _build_config(tmp_path)

    with caplog.at_level(logging.DEBUG, logger="mailprocessor"):
        run_pipeline(config, _build_rules())
        run_pipeline(config, _build_rules())

    messages = [record.getMessage() for record in caplog.records]
    assert any(m.startswith("Found 2 file(s) matching *.eml") for m in messages)
    assert "Processed good.eml (<good@example.com>)" in messages
    assert any(
        r.levelno == logging.WARNING
        and r.getMessage() == "Failed missing-phone.eml (<bad@example.com>): Required fields missing: Telefonnummer"
        for r in caplog.records
    )
    assert "Skipped good.eml (<good@example.com>): already processed in an earlier run" in messages
    assert any("found Mail-Adresse, Name, Kurs, Zeit; missing Telefonnummer" in m for m in messages)
    assert any(m.startswith("Run finished: seen=2 processed=0 skipped=1 failed=1") for m in messages)
    # Confidential content (extracted values, body text) never appears in logs.
    log_text = "\n".join(messages)
    assert "Experimente" not in log_text
    assert "1234 567890" not in log_text


def test_info_level_hides_per_field_details(tmp_path: Path, caplog) -> None:
    inbox = tmp_path / "inbox"
    inbox.mkdir()
    _write_eml(inbox / "good.eml", GOOD_BODY, "good@example.com")

    with caplog.at_level(logging.INFO, logger="mailprocessor"):
        run_pipeline(_build_config(tmp_path), _build_rules())

    messages = [record.getMessage() for record in caplog.records]
    assert "Processed good.eml (<good@example.com>)" in messages
    assert not any("found " in m for m in messages)


def test_data_rows_include_the_full_mail_text(tmp_path: Path) -> None:
    inbox = tmp_path / "inbox"
    inbox.mkdir()
    _write_eml(inbox / "good.eml", GOOD_BODY, "good@example.com")

    run_pipeline(_build_config(tmp_path), _build_rules())

    sheet = load_workbook(tmp_path / "out" / "mail_export.xlsx")["daten"]
    assert sheet.cell(row=1, column=8).value == "E-Mail-Inhalt"
    assert sheet.cell(row=2, column=8).value == "\n".join(GOOD_BODY)


def test_control_characters_in_a_mail_do_not_abort_the_run(tmp_path: Path) -> None:
    inbox = tmp_path / "inbox"
    inbox.mkdir()
    body = [line.replace("Jan Must+", "Jan\x01 Must+") for line in GOOD_BODY]
    _write_eml(inbox / "control-char.eml", body, "ctrl@example.com")

    summary = run_pipeline(_build_config(tmp_path), _build_rules())

    assert summary.processed == 1
    assert load_workbook(tmp_path / "out" / "mail_export.xlsx")["daten"]["B2"].value == "Jan Must+"


def _inbox_with_good_mails(tmp_path: Path, count: int) -> Path:
    inbox = tmp_path / "inbox"
    inbox.mkdir()
    for index in range(count):
        _write_eml(inbox / f"mail-{index}.eml", GOOD_BODY, f"mail-{index}@example.com")
    return inbox


def test_progress_reports_each_message_with_total(tmp_path: Path) -> None:
    inbox = _inbox_with_good_mails(tmp_path, 3)
    _write_eml_with_date(inbox / "old.eml", GOOD_BODY, "old@example.com", "Mon, 01 Jan 2001 10:00:00 +0000")
    config = _build_config(tmp_path)
    config.app.max_age_days = 30
    calls: list[tuple[int, int]] = []

    summary = run_pipeline(config, _build_rules(), progress=lambda current, total: calls.append((current, total)))

    # The old mail is filtered out but still counts, so the progress reaches the total.
    assert calls == [(1, 4), (2, 4), (3, 4), (4, 4)]
    assert summary.processed == 3


def test_cancel_saves_what_was_processed_and_next_run_continues(tmp_path: Path) -> None:
    _inbox_with_good_mails(tmp_path, 5)
    config = _build_config(tmp_path)
    cancel = threading.Event()

    def stop_after_two(current: int, _total: int) -> None:
        if current == 3:  # mails 1 and 2 are done when the third one is read
            cancel.set()

    first = run_pipeline(config, _build_rules(), progress=stop_after_two, cancel=cancel)
    assert (first.processed, first.cancelled) == (2, True)
    assert load_workbook(tmp_path / "out" / "mail_export.xlsx")["daten"].max_row == 3

    second = run_pipeline(config, _build_rules())
    assert (second.processed, second.skipped, second.cancelled) == (3, 2, False)
    assert load_workbook(tmp_path / "out" / "mail_export.xlsx")["daten"].max_row == 6


def test_start_over_keeps_backup_and_exports_everything_again(tmp_path: Path) -> None:
    inbox = _inbox_with_good_mails(tmp_path, 2)
    config = _build_config(tmp_path)
    run_pipeline(config, _build_rules())
    before = _snapshot(inbox)

    backup = start_over(config, now=datetime(2026, 10, 5, 21, 30, 0, tzinfo=UTC))

    assert backup == tmp_path / "out" / "mail_export_backup_2026-10-05_21-30-00.xlsx"
    assert load_workbook(backup)["daten"].max_row == 3
    assert not (tmp_path / "out" / "mail_export.xlsx").exists()
    with Ledger(tmp_path / "data" / "ledger.db", read_only=True) as ledger:
        assert ledger.count_processed() == 0
    assert _snapshot(inbox) == before

    again = run_pipeline(config, _build_rules())
    assert (again.processed, again.skipped) == (2, 0)
    assert load_workbook(tmp_path / "out" / "mail_export.xlsx")["daten"].max_row == 3


def test_start_over_without_previous_run_does_nothing(tmp_path: Path) -> None:
    assert start_over(_build_config(tmp_path)) is None
    assert not (tmp_path / "data").exists()


def test_problems_name_the_mail_and_missing_fields_also_in_a_test_run(tmp_path: Path) -> None:
    inbox = tmp_path / "inbox"
    inbox.mkdir()
    _write_eml(inbox / "good.eml", GOOD_BODY, "good@example.com")
    _write_eml(inbox / "no-phone.eml", GOOD_BODY[:-1], "no-phone@example.com")
    (inbox / "broken.eml").write_bytes(b"")
    config = _build_config(tmp_path)
    config.app.dry_run = True

    summary = run_pipeline(config, _build_rules())

    problems = {problem.name.split(" ")[0]: problem for problem in summary.problems}
    assert summary.failed == len(summary.problems)
    phone = problems["no-phone.eml"]
    assert phone.missing == ("Telefonnummer",)
    assert phone.body is not None and "Angebot: Experimente" in phone.body
    assert "Max Mustermann" in phone.header_text


def _data_rows(tmp_path: Path) -> list[dict]:
    sheet = load_workbook(tmp_path / "out" / "mail_export.xlsx")["daten"]
    header, *rows = sheet.iter_rows(values_only=True)
    return [dict(zip(header, row, strict=True)) for row in rows]


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


def test_error_sheet_names_file_sender_subject_and_missing_fields(tmp_path: Path) -> None:
    inbox = tmp_path / "inbox"
    inbox.mkdir()
    _write_eml(inbox / "ohne-telefon.eml", GOOD_BODY[:-1], "no-phone@example.com")

    run_pipeline(_build_config(tmp_path), _build_rules())

    errors = load_workbook(tmp_path / "out" / "mail_export.xlsx")["fehler"]
    row = dict(zip(ERROR_COLUMNS, next(errors.iter_rows(min_row=2, values_only=True)), strict=True))
    assert row["E-Mail"] == "ohne-telefon.eml"
    assert row["Absender"] == "Max Mustermann <max.mustermann@mail.com>"
    assert row["Betreff"] == "Schnuppernachmittag"
    assert isinstance(row["Eingegangen am"], datetime)
    assert row["Fehlende Felder"] == "Telefonnummer"
    assert row["Grund"] == "Pflichtfelder nicht gefunden"


def test_fixed_mail_removes_its_error_row(tmp_path: Path) -> None:
    inbox = tmp_path / "inbox"
    inbox.mkdir()
    _write_eml(inbox / "a.eml", GOOD_BODY[:-1], "a@example.com")
    run_pipeline(_build_config(tmp_path), _build_rules())
    _write_eml(inbox / "a.eml", GOOD_BODY, "a@example.com")  # e.g. the form was fixed and the mail sent again

    summary = run_pipeline(_build_config(tmp_path), _build_rules())

    assert summary.processed == 1

def test_mails_not_matching_the_filter_are_left_out_without_an_error(tmp_path: Path) -> None:
    inbox = tmp_path / "inbox"
    inbox.mkdir()
    _write_eml(inbox / "form.eml", GOOD_BODY, "form@example.com", subject="Kontaktformular: Anmeldung")
    _write_eml(inbox / "news.eml", ["Unser Newsletter im Oktober"], "news@example.com", subject="Newsletter")
    config = _build_config(tmp_path)
    config.filter = MailFilter(subject=["kontaktformular"])
    calls: list[tuple[int, int]] = []

    summary = run_pipeline(config, _build_rules(), progress=lambda current, total: calls.append((current, total)))

    assert (summary.seen, summary.processed, summary.failed, summary.filtered) == (1, 1, 0, 1)
    assert summary.problems == ()
    assert calls[-1] == (2, 2)
    workbook = load_workbook(tmp_path / "out" / "mail_export.xlsx")
    assert workbook["daten"].max_row == 2
    assert workbook["fehler"].max_row == 1


def test_mails_left_out_by_the_filter_are_read_once_the_filter_changes(tmp_path: Path) -> None:
    inbox = tmp_path / "inbox"
    inbox.mkdir()
    _write_eml(inbox / "other.eml", GOOD_BODY, "other@example.com", subject="Rückfrage")
    config = _build_config(tmp_path)
    config.filter = MailFilter(subject=["Kontaktformular"])
    assert run_pipeline(config, _build_rules()).filtered == 1

    config.filter = MailFilter()
    summary = run_pipeline(config, _build_rules())

    assert (summary.processed, summary.filtered) == (1, 0)


def test_unreadable_mails_are_reported_even_with_a_filter(tmp_path: Path, monkeypatch) -> None:
    (tmp_path / "inbox").mkdir()
    config = _build_config(tmp_path)
    config.filter = MailFilter(subject=["Kontaktformular"])
    unreadable = MailReadError(source_type="eml", source_location="x", message_identity="file:x", reason="kaputt")

    def source(*_args, **_kwargs):
        yield unreadable

    monkeypatch.setattr("mailprocessor.processor.iter_source_messages", source)

    summary = run_pipeline(config, _build_rules())

    # Subject and sender of an unreadable mail are unknown, so it may well be one of the wanted ones.
    assert (summary.failed, summary.filtered) == (1, 0)


def test_setting_a_filter_clears_old_error_rows_of_mails_it_leaves_out(tmp_path: Path) -> None:
    inbox = tmp_path / "inbox"
    inbox.mkdir()
    _write_eml(inbox / "news.eml", ["Unser Newsletter im Oktober"], "news@example.com", subject="Newsletter")
    config = _build_config(tmp_path)
    assert run_pipeline(config, _build_rules()).failed == 1

    config.filter = MailFilter(subject=["Kontaktformular"])
    run_pipeline(config, _build_rules())

    assert load_workbook(tmp_path / "out" / "mail_export.xlsx")["fehler"].max_row == 1


# --- several profiles

ABMELDUNG_BODY = [
    "Von: Eva Muster <eva@example.com>",
    "hiermit melde ich mein Kind ab.",
    "Angebot: Experimente",
    "Grund: Umzug",
]


def _two_profiles() -> ParsingRules:
    return ParsingRules(
        profiles=[
            Profile(name="Anmeldung", fields=_build_rules().profiles[0].fields),
            Profile(
                name="Abmeldung",
                fields=[
                    FieldRule(column="Mail-Adresse", type="email", label=["Von", "From"]),
                    FieldRule(column="Kurs", type="label", label="Angebot:"),
                    FieldRule(column="Grund", type="label", label="Grund:"),
                ],
            ),
        ]
    )


def _write_profile_mails(inbox: Path) -> None:
    inbox.mkdir()
    _write_eml(inbox / "1_anmeldung.eml", GOOD_BODY, "an@example.com")
    _write_eml(inbox / "2_abmeldung.eml", ABMELDUNG_BODY, "ab@example.com")
    _write_eml(inbox / "3_unklar.eml", ["Grund: keiner"], "unklar@example.com")


def test_profiles_share_one_sheet_with_a_profile_column(tmp_path: Path) -> None:
    _write_profile_mails(tmp_path / "inbox")

    summary = run_pipeline(_build_config(tmp_path), _two_profiles())

    assert (summary.processed, summary.failed) == (2, 1)
    assert summary.per_profile == (("Anmeldung", 1), ("Abmeldung", 1))
    workbook = load_workbook(tmp_path / "out" / "mail_export.xlsx")
    assert workbook.sheetnames == ["daten", "fehler"]
    rows = list(workbook["daten"].iter_rows(values_only=True))
    assert rows[0][:7] == ("Profil", "Mail-Adresse", "Name", "Kurs", "Zeit", "Telefonnummer", "Grund")
    assert rows[1][:4] == ("Anmeldung", "max.mustermann@mail.com", "Jan Must+", "Experimente")
    assert rows[2][:7] == ("Abmeldung", "eva@example.com", None, "Experimente", None, None, "Umzug")


def test_profiles_can_write_one_sheet_each(tmp_path: Path) -> None:
    _write_profile_mails(tmp_path / "inbox")
    config = _build_config(tmp_path)
    config.app.profile_sheets = "per_profile"

    run_pipeline(config, _two_profiles())

    workbook = load_workbook(tmp_path / "out" / "mail_export.xlsx")
    assert workbook.sheetnames == ["Anmeldung", "Abmeldung", "fehler"]
    anmeldung = list(workbook["Anmeldung"].iter_rows(values_only=True))
    abmeldung = list(workbook["Abmeldung"].iter_rows(values_only=True))
    assert len(anmeldung) == len(abmeldung) == 2
    assert abmeldung[0][:3] == ("Mail-Adresse", "Kurs", "Grund")
    assert abmeldung[1][:3] == ("eva@example.com", "Experimente", "Umzug")


def test_failed_mail_names_the_closest_profile(tmp_path: Path) -> None:
    _write_profile_mails(tmp_path / "inbox")

    summary = run_pipeline(_build_config(tmp_path), _two_profiles())

    # "Grund: keiner" + the From header fit Abmeldung best; only its "Kurs" is missing.
    (problem,) = summary.problems
    assert (problem.profile, problem.missing) == ("Abmeldung", ("Kurs",))
    errors = list(load_workbook(tmp_path / "out" / "mail_export.xlsx")["fehler"].iter_rows(values_only=True))
    row = dict(zip(errors[0], errors[1], strict=False))
    assert row["Fehlende Felder"] == "Kurs"
    assert row["Grund"] == "Pflichtfelder nicht gefunden (am ähnlichsten: Profil „Abmeldung“)"


def test_single_profile_has_no_profile_column(tmp_path: Path) -> None:
    inbox = tmp_path / "inbox"
    inbox.mkdir()
    _write_eml(inbox / "good.eml", GOOD_BODY, "good@example.com")

    summary = run_pipeline(_build_config(tmp_path), _build_rules())

    header = next(load_workbook(tmp_path / "out" / "mail_export.xlsx")["daten"].iter_rows(values_only=True))
    assert "Profil" not in header
    assert summary.per_profile == ()


def test_profile_sheet_must_not_be_the_error_sheet(tmp_path: Path) -> None:
    (tmp_path / "inbox").mkdir()
    config = _build_config(tmp_path)
    config.app.profile_sheets = "per_profile"
    rules = ParsingRules(profiles=[Profile(name="Fehler", fields=_build_rules().profiles[0].fields)])

    with pytest.raises(ValueError, match="same name as the sheet for problems"):
        run_pipeline(config, rules)


def _error_rows(tmp_path: Path) -> list[tuple]:
    return list(load_workbook(tmp_path / "out" / "mail_export.xlsx")["fehler"].iter_rows(min_row=2, values_only=True))


def test_error_row_of_an_unreadable_file_goes_once_the_file_is_read(tmp_path: Path, monkeypatch) -> None:
    # E.g. on Windows a file that is still being copied cannot be opened; the next run reads it fine.
    inbox = tmp_path / "inbox"
    inbox.mkdir()
    _write_eml(inbox / "a.eml", GOOD_BODY, "a@example.com")
    config = _build_config(tmp_path)
    read_bytes = Path.read_bytes

    def still_copying(path: Path) -> bytes:
        if path.name == "a.eml":
            raise PermissionError("in use by another process")
        return read_bytes(path)

    monkeypatch.setattr(Path, "read_bytes", still_copying)
    assert run_pipeline(config, _build_rules()).failed == 1
    assert len(_error_rows(tmp_path)) == 1
    monkeypatch.setattr(Path, "read_bytes", read_bytes)

    assert run_pipeline(config, _build_rules()).processed == 1
    assert _error_rows(tmp_path) == []
def test_overlapping_runs_never_lose_rows(tmp_path: Path, monkeypatch) -> None:
    # Run B has loaded the workbook and waits for its mail source (e.g. a slow IMAP login) while run A
    # (a second window, or a scheduled CLI run) exports the mail. B must not save its older copy over A's rows.
    inbox = tmp_path / "inbox"
    inbox.mkdir()
    _write_eml(inbox / "x.eml", GOOD_BODY, "x@example.com")
    config = _build_config(tmp_path)
    b_waits, release_b = threading.Event(), threading.Event()
    read_mails = processor.iter_eml_messages

    def slow_for_b(*args, **kwargs):
        if threading.current_thread().name == "B":
            b_waits.set()
            release_b.wait(10)
        yield from read_mails(*args, **kwargs)

    monkeypatch.setattr(processor, "iter_eml_messages", slow_for_b)
    monkeypatch.setattr("mailprocessor.ledger.LOCK_TIMEOUT_SECONDS", 0.1)
    outcomes: dict[str, object] = {}

    def run(name: str) -> None:
        try:
            outcomes[name] = run_pipeline(config, _build_rules())
        except Exception as exc:
            outcomes[name] = exc

    run_b = threading.Thread(target=run, args=("B",), name="B")
    run_b.start()
    assert b_waits.wait(10)
    run("A")
    release_b.set()
    run_b.join(10)

    assert [row["Name"] for row in _data_rows(tmp_path)] == ["Jan Must+"]
    # The run that started second is refused with a clear message instead of racing the first.
    assert isinstance(outcomes["A"], RunInProgressError)
    assert isinstance(outcomes["B"], RunSummary) and outcomes["B"].processed == 1
