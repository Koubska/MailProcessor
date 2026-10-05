from datetime import datetime
from pathlib import Path

import pytest
from openpyxl import Workbook, load_workbook

from mailprocessor.errors import SheetHeaderError
from mailprocessor.excel_writer import (
    CONTENT_COLUMN,
    CONTENT_COLUMN_WIDTH,
    DATE_COLUMN_WIDTH,
    DATE_FORMAT,
    ERROR_COLUMNS,
    EXCEL_CELL_LIMIT,
    KEY_COLUMN,
    LEGACY_ERROR_COLUMNS,
    MAX_COLUMN_WIDTH,
    MIN_COLUMN_WIDTH,
    RECEIVED_COLUMN,
    TRANSFERRED_COLUMN,
    ErrorEntry,
    ExcelOutput,
)

DATA_COLUMNS = ["Mail-Adresse", "Name", "Kurs", "Zeit", "Telefonnummer"]
FIXED_COLUMNS = [RECEIVED_COLUMN, TRANSFERRED_COLUMN, CONTENT_COLUMN]
RUN_AT = datetime(2026, 10, 6, 9, 30)
RECEIVED_AT = datetime(2026, 10, 5, 14, 0)


def _output(path: Path, columns: list[str]) -> ExcelOutput:
    return ExcelOutput(path, "daten", "fehler", columns, now=RUN_AT)


def _rows(path: Path, sheet: str = "daten") -> list[tuple]:
    return list(load_workbook(path)[sheet].iter_rows(values_only=True))


def _error(key: str, reason: str = "Pflichtfelder nicht gefunden", missing: tuple[str, ...] = ("Telefonnummer",)):
    return ErrorEntry(
        key=key,
        name=f"{key}.eml",
        sender="Max Mustermann <max@example.com>",
        subject="Anmeldung",
        received=RECEIVED_AT,
        missing=missing,
        reason=reason,
    )


# --- data sheet


def test_new_workbook_has_rule_columns_then_dates_then_mail_text(tmp_path: Path) -> None:
    path = tmp_path / "output.xlsx"

    _output(path, DATA_COLUMNS).save()

    workbook = load_workbook(path)
    assert workbook.sheetnames == ["daten", "fehler"]
    assert [cell.value for cell in workbook["daten"][1]] == [*DATA_COLUMNS, *FIXED_COLUMNS]
    assert [cell.value for cell in workbook["fehler"][1]] == list(ERROR_COLUMNS)
    assert workbook["daten"].freeze_panes == "A2"


def test_append_data_writes_values_dates_and_mail_text(tmp_path: Path) -> None:
    path = tmp_path / "output.xlsx"
    output = _output(path, DATA_COLUMNS)

    output.append_data(
        {"Name": "Jan", "Kurs": "Experimente", "Telefonnummer": "1234", "Zeit": "Mo", "Mail-Adresse": "max@mail.com"},
        "Mail text",
        received=RECEIVED_AT,
    )
    output.save()

    sheet = load_workbook(path)["daten"]
    values = [cell.value for cell in sheet[2]]
    assert values == ["max@mail.com", "Jan", "Experimente", "Mo", "1234", RECEIVED_AT, RUN_AT, "Mail text"]
    # Real Excel dates, so they sort and filter correctly.
    assert sheet["F2"].number_format == DATE_FORMAT
    assert sheet["G2"].number_format == DATE_FORMAT


def test_mail_without_usable_date_leaves_the_cell_empty(tmp_path: Path) -> None:
    path = tmp_path / "output.xlsx"
    output = _output(path, ["Name"])

    output.append_data({"Name": "Jan"}, "Text", received=None)
    output.save()

    assert _rows(path)[1] == ("Jan", None, RUN_AT, "Text")


def test_values_starting_with_equals_are_stored_as_text_not_formulas(tmp_path: Path) -> None:
    path = tmp_path / "output.xlsx"
    output = _output(path, ["Name"])
    payload = '=HYPERLINK("http://evil.example","klick")'

    output.append_data({"Name": payload}, payload)
    output.upsert_error(_error("k", reason=payload))
    output.save()

    workbook = load_workbook(path)
    assert workbook["daten"]["A2"].value == payload
    assert workbook["daten"]["A2"].data_type == "s"
    reason_cell = workbook["fehler"].cell(row=2, column=ERROR_COLUMNS.index("Grund") + 1)
    assert reason_cell.value == payload and reason_cell.data_type == "s"


def test_existing_workbook_is_appended_to(tmp_path: Path) -> None:
    path = tmp_path / "output.xlsx"
    first = _output(path, ["Name"])
    first.append_data({"Name": "A"}, "text A")
    first.save()

    second = _output(path, ["Name"])
    second.append_data({"Name": "B"}, "text B")
    second.save()

    assert [row[0] for row in _rows(path)] == ["Name", "A", "B"]


def test_reordered_rules_still_write_each_value_to_its_column(tmp_path: Path) -> None:
    path = tmp_path / "output.xlsx"
    _output(path, ["Name", "Kurs"]).save()

    output = _output(path, ["Kurs", "Name"])
    output.append_data({"Name": "Jan", "Kurs": "Experimente"}, "Text")
    output.save()

    rows = _rows(path)
    assert rows[0][:2] == ("Name", "Kurs")  # existing columns never move
    assert rows[1][:2] == ("Jan", "Experimente")


def test_new_rule_gets_a_new_column_at_the_end(tmp_path: Path) -> None:
    path = tmp_path / "output.xlsx"
    first = _output(path, ["Name"])
    first.append_data({"Name": "Alt"}, "alter Text")
    first.save()

    second = _output(path, ["Name", "Betreff"])
    second.append_data({"Name": "Neu", "Betreff": "Anmeldung"}, "neuer Text")
    second.save()

    rows = _rows(path)
    assert rows[0] == ("Name", *FIXED_COLUMNS, "Betreff")
    assert rows[1] == ("Alt", None, RUN_AT, "alter Text", None)
    assert rows[2] == ("Neu", None, RUN_AT, "neuer Text", "Anmeldung")


def test_own_columns_and_notes_in_excel_are_kept(tmp_path: Path) -> None:
    path = tmp_path / "output.xlsx"
    first = _output(path, ["Name"])
    first.append_data({"Name": "Anna"}, "Text A")
    first.save()
    workbook = load_workbook(path)
    sheet = workbook["daten"]
    sheet.cell(row=1, column=sheet.max_column + 1, value="Bestätigt")
    sheet.cell(row=2, column=sheet.max_column, value="ja")
    workbook.save(path)

    second = _output(path, ["Name"])
    second.append_data({"Name": "Ben"}, "Text B")
    second.save()

    rows = _rows(path)
    assert rows[0] == ("Name", *FIXED_COLUMNS, "Bestätigt")
    assert rows[1][-1] == "ja"
    assert rows[2] == ("Ben", None, RUN_AT, "Text B", None)


def test_removed_rule_keeps_its_column_and_new_rows_leave_it_empty(tmp_path: Path) -> None:
    path = tmp_path / "output.xlsx"
    first = _output(path, ["Name", "Kurs"])
    first.append_data({"Name": "Anna", "Kurs": "Chemie"}, "Text A")
    first.save()

    second = _output(path, ["Name"])
    second.append_data({"Name": "Ben"}, "Text B")
    second.save()

    rows = _rows(path)
    assert rows[0][:2] == ("Name", "Kurs")
    assert rows[1][:2] == ("Anna", "Chemie")
    assert rows[2][:2] == ("Ben", None)


def test_columns_added_next_to_unlabelled_notes_do_not_overwrite_them(tmp_path: Path) -> None:
    path = tmp_path / "output.xlsx"
    first = _output(path, ["Name"])
    first.append_data({"Name": "Anna"}, "Text")
    first.save()
    workbook = load_workbook(path)
    workbook["daten"].cell(row=2, column=6, value="Notiz ohne Überschrift")
    workbook.save(path)

    second = _output(path, ["Name", "Kurs"])
    second.save()

    rows = _rows(path)
    assert rows[1][5] == "Notiz ohne Überschrift"
    assert rows[0][6] == "Kurs"


def test_sheet_with_data_but_without_header_row_is_rejected(tmp_path: Path) -> None:
    path = tmp_path / "output.xlsx"
    workbook = Workbook()
    workbook.active.title = "daten"
    workbook.active.cell(row=2, column=1, value="Daten ohne Kopfzeile")
    workbook.save(path)

    with pytest.raises(SheetHeaderError, match="daten"):
        _output(path, ["Name"])


def test_workbook_from_before_the_date_and_text_columns_is_upgraded(tmp_path: Path) -> None:
    path = tmp_path / "output.xlsx"
    old = Workbook()
    old.active.title = "daten"
    old.active.append(["Name"])
    old.active.append(["Alt"])
    old.save(path)

    output = _output(path, ["Name"])
    output.append_data({"Name": "Neu"}, "Text")
    output.save()

    assert _rows(path) == [("Name", *FIXED_COLUMNS), ("Alt", None, None, None), ("Neu", None, RUN_AT, "Text")]


def test_mail_text_is_cleaned_and_truncated_to_excel_limit(tmp_path: Path) -> None:
    path = tmp_path / "output.xlsx"
    output = _output(path, ["Name"])

    output.append_data({"Name": "Jan\x01"}, "Zeile 1\x00\nZeile 2")
    output.append_data({"Name": "Lang"}, "x" * (EXCEL_CELL_LIMIT + 10))
    output.save()

    rows = _rows(path)[1:]
    assert (rows[0][0], rows[0][-1]) == ("Jan", "Zeile 1\nZeile 2")
    assert len(rows[1][-1]) == EXCEL_CELL_LIMIT
    assert rows[1][-1].endswith("gekürzt: Excel erlaubt höchstens 32.767 Zeichen pro Zelle]")


def test_check_writable_reports_locked_workbook(tmp_path: Path, monkeypatch) -> None:
    path = tmp_path / "output.xlsx"
    _output(path, ["Name"]).save()
    output = _output(path, ["Name"])

    def locked_open(self, *args, **kwargs):
        raise PermissionError("locked")

    monkeypatch.setattr(Path, "open", locked_open)

    with pytest.raises(OSError, match="open in Excel"):
        output.check_writable()


# --- error sheet


def test_error_rows_are_readable(tmp_path: Path) -> None:
    path = tmp_path / "output.xlsx"
    output = _output(path, DATA_COLUMNS)

    output.upsert_error(_error("eml|/mails|<a@example.com>", missing=("Telefonnummer", "Kurs")))
    output.save()

    sheet = load_workbook(path)["fehler"]
    row = dict(zip(ERROR_COLUMNS, next(sheet.iter_rows(min_row=2, values_only=True)), strict=True))
    assert row == {
        "E-Mail": "eml|/mails|<a@example.com>.eml",
        "Absender": "Max Mustermann <max@example.com>",
        "Betreff": "Anmeldung",
        "Eingegangen am": RECEIVED_AT,
        "Fehlende Felder": "Telefonnummer, Kurs",
        "Grund": "Pflichtfelder nicht gefunden",
        "Geprüft am": RUN_AT,
        KEY_COLUMN: "eml|/mails|<a@example.com>",
    }
    key_letter = sheet.cell(row=1, column=ERROR_COLUMNS.index(KEY_COLUMN) + 1).column_letter
    assert sheet.column_dimensions[key_letter].hidden  # internal; only used to find the row again


def test_upsert_error_keeps_one_row_per_message(tmp_path: Path) -> None:
    path = tmp_path / "output.xlsx"
    output = _output(path, DATA_COLUMNS)

    output.upsert_error(_error("msg-1", "erster Grund"))
    output.upsert_error(_error("msg-2"))
    output.upsert_error(_error("msg-1", "zweiter Grund"))
    output.save()

    reason, key = ERROR_COLUMNS.index("Grund"), ERROR_COLUMNS.index(KEY_COLUMN)
    rows = _rows(path, "fehler")[1:]
    assert [(row[key], row[reason]) for row in rows] == [
        ("msg-2", "Pflichtfelder nicht gefunden"),
        ("msg-1", "zweiter Grund"),
    ]


def test_remove_errors_for_drops_the_row_of_a_mail_that_now_works(tmp_path: Path) -> None:
    path = tmp_path / "output.xlsx"
    output = _output(path, DATA_COLUMNS)
    output.upsert_error(_error("msg-1"))

    output.remove_errors_for("msg-1")
    output.save()

    assert load_workbook(path)["fehler"].max_row == 1


def test_old_english_error_sheet_is_converted(tmp_path: Path) -> None:
    path = tmp_path / "output.xlsx"
    old = Workbook()
    old.active.title = "daten"
    old.active.append(["Name", CONTENT_COLUMN])
    errors = old.create_sheet("fehler")
    errors.append(LEGACY_ERROR_COLUMNS)
    checked = "2026-10-03T20:00:00+00:00"
    missing = "Required fields missing: Telefonnummer"
    errors.append(["eml", "/mails", "<a@example.com>", "Telefonnummer", missing, checked])
    errors.append(["eml", "/mails", "file:kaputt.eml", "", "Could not read file (OSError): disk", checked])
    old.save(path)

    _output(path, ["Name"]).save()

    header, *data = _rows(path, "fehler")
    rows = [dict(zip(ERROR_COLUMNS, row, strict=True)) for row in data]
    assert list(header) == list(ERROR_COLUMNS)
    assert rows[0]["E-Mail"] == "<a@example.com>"
    assert rows[0]["Fehlende Felder"] == "Telefonnummer"
    assert rows[0]["Grund"] == "Pflichtfelder nicht gefunden"
    assert rows[0][KEY_COLUMN] == "eml|/mails|<a@example.com>"
    assert rows[1]["E-Mail"] == "kaputt.eml"
    assert rows[1]["Grund"] == "Could not read file (OSError): disk"
    assert isinstance(rows[0]["Geprüft am"], datetime)

    # The converted row is found again: the next failure of that mail replaces it.
    again = _output(path, ["Name"])
    again.upsert_error(_error("eml|/mails|<a@example.com>"))
    again.save()
    assert len(_rows(path, "fehler")) == 3


# --- formatting


def test_save_formats_header_and_column_widths(tmp_path: Path) -> None:
    path = tmp_path / "output.xlsx"
    output = _output(path, DATA_COLUMNS)
    output.append_data(
        {"Mail-Adresse": "max.mustermann@mail.com", "Name": "x" * 200}, "Text\n" * 50, received=RECEIVED_AT
    )
    output.save()

    sheet = load_workbook(path)["daten"]
    assert all(cell.font.bold for cell in sheet[1])
    widths = {letter: sheet.column_dimensions[letter].width for letter in "ABCFGH"}
    assert widths["A"] == len("max.mustermann@mail.com") + 2
    assert widths["B"] == MAX_COLUMN_WIDTH
    assert widths["C"] == MIN_COLUMN_WIDTH  # "Kurs", empty
    assert widths["F"] == widths["G"] == DATE_COLUMN_WIDTH
    assert widths["H"] == CONTENT_COLUMN_WIDTH
    assert not sheet["H2"].alignment.wrap_text  # one row per mail, however long the text


def test_save_keeps_column_widths_set_by_the_user(tmp_path: Path) -> None:
    path = tmp_path / "output.xlsx"
    _output(path, DATA_COLUMNS).save()
    workbook = load_workbook(path)
    workbook["daten"].column_dimensions["B"].width = 33
    workbook.save(path)

    output = _output(path, DATA_COLUMNS)
    output.append_data({"Name": "x" * 200}, "Text")
    output.save()

    assert load_workbook(path)["daten"].column_dimensions["B"].width == 33
