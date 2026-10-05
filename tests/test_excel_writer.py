from pathlib import Path

import pytest
from openpyxl import Workbook, load_workbook

from mailprocessor.excel_writer import CONTENT_COLUMN, ERROR_COLUMNS, EXCEL_CELL_LIMIT, ExcelOutput

DATA_COLUMNS = ["Mail-Adresse", "Name", "Kurs", "Zeit", "Telefonnummer"]


def _error(identity: str, reason: str = "required fields missing") -> dict[str, str]:
    return {
        "source_type": "eml",
        "source_location": "/tmp/eml",
        "message_identity": identity,
        "missing_columns": "Telefonnummer",
        "error_reason": reason,
        "processed_at": "2026-10-03T22:00:00Z",
    }


def test_new_workbook_has_expected_sheets_and_headers(tmp_path: Path) -> None:
    path = tmp_path / "output.xlsx"

    ExcelOutput(path, "daten", "fehler", DATA_COLUMNS).save()

    workbook = load_workbook(path)
    assert workbook.sheetnames == ["daten", "fehler"]
    assert [cell.value for cell in workbook["daten"][1]] == [*DATA_COLUMNS, CONTENT_COLUMN]
    assert workbook["daten"].freeze_panes == "A2"


def test_append_data_writes_in_column_order(tmp_path: Path) -> None:
    path = tmp_path / "output.xlsx"
    output = ExcelOutput(path, "daten", "fehler", DATA_COLUMNS)

    output.append_data(
        {
            "Name": "Jan Must+",
            "Kurs": "Experimente",
            "Telefonnummer": "1234 567890",
            "Zeit": "Mo",
            "Mail-Adresse": "max.mustermann@mail.com",
        },
        "Mail text",
    )
    output.save()

    workbook = load_workbook(path)
    assert [cell.value for cell in workbook["daten"][2]] == [
        "max.mustermann@mail.com",
        "Jan Must+",
        "Experimente",
        "Mo",
        "1234 567890",
        "Mail text",
    ]


def test_values_starting_with_equals_are_stored_as_text_not_formulas(tmp_path: Path) -> None:
    path = tmp_path / "output.xlsx"
    output = ExcelOutput(path, "daten", "fehler", ["Name"])
    payload = '=HYPERLINK("http://evil.example","klick")'

    output.append_data({"Name": payload}, payload)
    output.upsert_error(_error(payload))
    output.save()

    workbook = load_workbook(path)
    assert workbook["daten"]["A2"].value == payload
    assert workbook["daten"]["A2"].data_type == "s"
    assert workbook["fehler"]["C2"].data_type == "s"


def test_existing_workbook_with_different_columns_is_rejected(tmp_path: Path) -> None:
    path = tmp_path / "output.xlsx"
    ExcelOutput(path, "daten", "fehler", ["Name", "Kurs"]).save()

    with pytest.raises(ValueError, match="has columns"):
        ExcelOutput(path, "daten", "fehler", ["Kurs", "Name"])


def test_existing_workbook_is_appended_to(tmp_path: Path) -> None:
    path = tmp_path / "output.xlsx"
    first = ExcelOutput(path, "daten", "fehler", ["Name"])
    first.append_data({"Name": "A"}, "text A")
    first.save()

    second = ExcelOutput(path, "daten", "fehler", ["Name"])
    second.append_data({"Name": "B"}, "text B")
    second.save()

    assert [row[0].value for row in load_workbook(path)["daten"].iter_rows()] == ["Name", "A", "B"]


def test_upsert_error_keeps_one_row_per_message(tmp_path: Path) -> None:
    path = tmp_path / "output.xlsx"
    output = ExcelOutput(path, "daten", "fehler", DATA_COLUMNS)

    output.upsert_error(_error("msg-1", "first"))
    output.upsert_error(_error("msg-2"))
    output.upsert_error(_error("msg-1", "second"))
    output.save()

    rows = [[cell.value for cell in row] for row in load_workbook(path)["fehler"].iter_rows(min_row=2)]
    assert [(row[2], row[4]) for row in rows] == [("msg-2", "required fields missing"), ("msg-1", "second")]


def test_remove_errors_for_drops_stale_error_rows(tmp_path: Path) -> None:
    path = tmp_path / "output.xlsx"
    output = ExcelOutput(path, "daten", "fehler", DATA_COLUMNS)
    output.upsert_error(_error("msg-1"))

    output.remove_errors_for("eml", "/tmp/eml", "msg-1")
    output.save()

    assert load_workbook(path)["fehler"].max_row == 1


def test_check_writable_reports_locked_workbook(tmp_path: Path, monkeypatch) -> None:
    path = tmp_path / "output.xlsx"
    ExcelOutput(path, "daten", "fehler", ["Name"]).save()
    output = ExcelOutput(path, "daten", "fehler", ["Name"])

    def locked_open(self, *args, **kwargs):
        raise PermissionError("locked")

    monkeypatch.setattr(Path, "open", locked_open)

    with pytest.raises(OSError, match="open in Excel"):
        output.check_writable()


def test_mail_text_is_cleaned_and_truncated_to_excel_limit(tmp_path: Path) -> None:
    path = tmp_path / "output.xlsx"
    output = ExcelOutput(path, "daten", "fehler", ["Name"])

    output.append_data({"Name": "Jan\x01"}, "Zeile 1\x00\nZeile 2")
    output.append_data({"Name": "Lang"}, "x" * (EXCEL_CELL_LIMIT + 10))
    output.save()

    rows = list(load_workbook(path)["daten"].iter_rows(min_row=2, values_only=True))
    assert rows[0] == ("Jan", "Zeile 1\nZeile 2")
    assert len(rows[1][1]) == EXCEL_CELL_LIMIT
    assert rows[1][1].endswith("gekürzt: Excel erlaubt höchstens 32.767 Zeichen pro Zelle]")


def test_workbook_without_content_column_is_upgraded(tmp_path: Path) -> None:
    path = tmp_path / "output.xlsx"
    old = Workbook()
    old.active.title = "daten"
    old.active.append(["Name"])
    old.active.append(["Alt"])
    old.create_sheet("fehler").append(ERROR_COLUMNS)
    old.save(path)

    output = ExcelOutput(path, "daten", "fehler", ["Name"])
    output.append_data({"Name": "Neu"}, "Text")
    output.save()

    rows = list(load_workbook(path)["daten"].iter_rows(values_only=True))
    assert rows == [("Name", CONTENT_COLUMN), ("Alt", None), ("Neu", "Text")]
