import logging
import threading
from datetime import UTC, datetime
from pathlib import Path

import pytest
from openpyxl import load_workbook

import mailprocessor.sources.eml_folder_source as eml_source
from mailprocessor.config import (
    AppConfig,
    AppSection,
    EmlSourceConfig,
    FieldRule,
    ImapSourceConfig,
    ParsingRules,
    SourceConfig,
)
from mailprocessor.ledger import Ledger
from mailprocessor.models import NormalizedMail
from mailprocessor.processor import iter_source_messages, run_pipeline, start_over


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


def _write_eml(path: Path, body_lines: list[str], message_id: str) -> None:
    path.write_text(
        "\n".join(
            [
                f"Message-ID: <{message_id}>",
                "From: Max Mustermann <max.mustermann@mail.com>",
                "Subject: Schnuppernachmittag",
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

    def fake_iter_imap_messages(_imap_cfg, max_age_days: int, now_utc: datetime, on_total=None):
        assert max_age_days == 0
        assert now_utc.tzinfo is UTC
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
    assert errors["C2"].value == "file:a-broken.eml"
    assert "disk error" in errors["E2"].value


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
        fields=[*_build_rules().fields, FieldRule(column="Extra", pattern=r"^Extra:(.+)$", required=True)]
    )
    run_pipeline(config, strict_rules)
    assert load_workbook(tmp_path / "out" / "mail_export.xlsx")["fehler"].max_row == 2

    relaxed_rules = ParsingRules(
        fields=[*_build_rules().fields, FieldRule(column="Extra", pattern=r"^Extra:(.+)$", required=False)]
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
    assert sheet.cell(row=1, column=6).value == "E-Mail-Inhalt"
    assert sheet.cell(row=2, column=6).value == "\n".join(GOOD_BODY)


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
