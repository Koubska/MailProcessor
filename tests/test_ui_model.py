from pathlib import Path

import pytest
from openpyxl import Workbook

from mailprocessor.config import FieldRule
from mailprocessor.gui import _default_config
from mailprocessor.preview import RulePreview
from mailprocessor.processor import Problem
from mailprocessor.ui_model import (
    config_from_form,
    excel_status,
    fields_status,
    form_values,
    mails_status,
    problem_text,
)


def test_form_values_roundtrip_to_the_same_config() -> None:
    config = _default_config()

    rebuilt, errors = config_from_form(form_values(config), "de")

    assert errors == {}
    assert rebuilt == config


@pytest.mark.parametrize(
    ("changes", "key", "message"),
    [
        ({"eml_folder": " "}, "eml_folder", "Ordner auswählen"),
        ({"max_age_days": "drei"}, "max_age_days", "ganze Zahl"),
        ({"max_messages": "-1"}, "max_messages", "ganze Zahl"),
        ({"output_xlsx": "out/export.csv"}, "output_xlsx", ".xlsx"),
        ({"sheet_errors": "daten"}, "sheet_errors", "verschiedene Namen"),
        ({"source_type": "imap", "imap_host": ""}, "imap_host", "E-Mail-Server"),
        ({"source_type": "imap", "imap_port": "99999"}, "imap_port", "Portnummer"),
    ],
)
def test_invalid_inputs_are_reported_per_field(changes: dict, key: str, message: str) -> None:
    values = {**form_values(_default_config()), **changes}

    config, errors = config_from_form(values, "de")

    assert config is None
    assert message in errors[key]


def test_invalid_port_of_the_inactive_source_does_not_block_saving() -> None:
    values = {**form_values(_default_config()), "imap_port": "abc"}

    config, errors = config_from_form(values, "de")

    assert errors == {}
    assert config is not None and config.source.imap is not None and config.source.imap.port == 993


def test_mails_status_for_folder(tmp_path: Path) -> None:
    values = {**form_values(_default_config()), "eml_folder": "./mails"}

    assert mails_status(values, tmp_path, False, "de").ok is False  # missing
    (tmp_path / "mails").mkdir()
    assert "ist leer" in mails_status(values, tmp_path, False, "de").text
    (tmp_path / "mails" / "a.eml").write_text("x", encoding="utf-8")
    status = mails_status(values, tmp_path, False, "de")
    assert status.ok and status.text == "Ordner „mails“ – 1 E-Mail(s)"


def test_mails_status_for_imap_asks_for_the_password() -> None:
    values = {**form_values(_default_config()), "source_type": "imap"}

    assert "Passwort" in mails_status(values, Path("."), False, "de").text
    assert mails_status(values, Path("."), True, "de").text == "Postfach user@example.com auf imap.example.com"


def test_fields_status() -> None:
    rules = [FieldRule(column="Name", type="label", label="Name"), FieldRule(column="Tel", type="label", label="Tel")]

    assert fields_status([], None, "de").ok is False
    assert fields_status(rules, None, "de").ok is None
    assert fields_status(rules, [RulePreview("a"), RulePreview("b")], "de").ok is True
    missing = fields_status(rules, [RulePreview("a"), RulePreview(None)], "de")
    assert missing.ok is False and missing.text == "2 Feld(er) – in der Beispiel-Mail fehlt „Tel“"


def test_excel_status_counts_rows(tmp_path: Path) -> None:
    path = tmp_path / "out.xlsx"
    assert "wird beim ersten Übertragen angelegt" in excel_status(path, "daten", "fehler", "de").text

    workbook = Workbook()
    workbook.active.title = "daten"
    workbook.active.append(["Name"])
    workbook.active.append(["A"])
    workbook.active.append(["B"])
    workbook.create_sheet("fehler").append(["reason"])
    workbook["fehler"].append(["x"])
    workbook.save(path)

    assert excel_status(path, "daten", "fehler", "de").text == "out.xlsx – 2 Zeile(n), 1 E-Mail(s) mit Problemen"
    (tmp_path / "broken.xlsx").write_text("not excel", encoding="utf-8")
    assert excel_status(tmp_path / "broken.xlsx", "daten", "fehler", "de").ok is None


def test_problem_text() -> None:
    missing = Problem(name="a.eml", reason="Required fields missing: Tel, Kurs", missing=("Tel", "Kurs"), body="x")
    unreadable = Problem(name="file:b.eml", reason="Could not read file (OSError)")

    assert problem_text(missing, "de") == "nicht gefunden: „Tel“, „Kurs“"
    assert problem_text(unreadable, "de") == "E-Mail konnte nicht gelesen werden"
    assert problem_text(Problem(name="c", reason="Internal parser error", body="x"), "en") == "Internal parser error"
