"""Failed mails: one row per mail on the error sheet, and the problem list for the GUI."""

from datetime import datetime
from pathlib import Path

from openpyxl import load_workbook
from pipeline_helpers import GOOD_BODY, _build_config, _build_rules, _error_rows, _write_eml

from mailprocessor.config import (
    FieldRule,
    ParsingRules,
)
from mailprocessor.excel_writer import ERROR_COLUMNS
from mailprocessor.processor import run_pipeline


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


def test_moving_the_mail_folder_keeps_one_error_row_per_mail(tmp_path: Path) -> None:
    # The ledger ignores the folder's location (moving it must not export again); the error sheet must
    # likewise keep one row per failing mail instead of adding a second one for the new location.
    inbox = tmp_path / "inbox"
    inbox.mkdir()
    _write_eml(inbox / "no-phone.eml", GOOD_BODY[:-1], "no-phone@example.com")
    config = _build_config(tmp_path)
    run_pipeline(config, _build_rules())

    moved = inbox.rename(tmp_path / "moved")
    assert config.source.eml is not None
    config.source.eml.folder = str(moved)
    run_pipeline(config, _build_rules())

    rows = load_workbook(tmp_path / "out" / "mail_export.xlsx")["fehler"].iter_rows(min_row=2, values_only=True)
    assert [row[0] for row in rows] == ["no-phone.eml"]
