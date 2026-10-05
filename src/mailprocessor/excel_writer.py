"""Excel output helpers."""

from __future__ import annotations

import os
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
from openpyxl.worksheet.worksheet import Worksheet

# Always the last column of the data sheet: the full text of the mail.
CONTENT_COLUMN = "E-Mail-Inhalt"
EXCEL_CELL_LIMIT = 32_767
_TRUNCATION_NOTE = "\n[… gekürzt: Excel erlaubt höchstens 32.767 Zeichen pro Zelle]"

ERROR_COLUMNS = [
    "source_type",
    "source_location",
    "message_identity",
    "missing_columns",
    "error_reason",
    "processed_at",
]
_ERROR_IDENTITY_COLUMNS = 3  # source_type, source_location, message_identity


def _header(sheet: Worksheet) -> list[object]:
    values = [cell.value for cell in sheet[1]]
    while values and values[-1] is None:
        values.pop()
    return values


def _cell_text(value: object) -> object:
    """Make untrusted text storable: drop control characters openpyxl rejects, respect Excel's cell limit."""
    if not isinstance(value, str):
        return value
    value = ILLEGAL_CHARACTERS_RE.sub("", value)
    if len(value) > EXCEL_CELL_LIMIT:
        value = value[: EXCEL_CELL_LIMIT - len(_TRUNCATION_NOTE)] + _TRUNCATION_NOTE
    return value


def _append_text_row(sheet: Worksheet, values: list[object]) -> None:
    sheet.append([_cell_text(value) for value in values])
    for cell in sheet[sheet.max_row]:
        # openpyxl turns any string starting with "=" into a live formula. Values come from
        # untrusted email content, so always store them as plain text (formula injection).
        if isinstance(cell.value, str) and cell.value.startswith("="):
            cell.data_type = "s"


class ExcelOutput:
    """In-memory workbook that is written to disk once, atomically, via `save()`."""

    def __init__(self, path: Path, data_sheet: str, error_sheet: str, data_columns: list[str]) -> None:
        self.path = path
        self.data_sheet = data_sheet
        self.error_sheet = error_sheet
        self.data_columns = data_columns
        self.is_new = not path.exists()

        if self.is_new:
            self.workbook = Workbook()
            self.workbook.active.title = data_sheet
        else:
            self.workbook = load_workbook(path)
        self._prepare_sheet(data_sheet, [*data_columns, CONTENT_COLUMN])
        self._prepare_sheet(error_sheet, ERROR_COLUMNS)

    def _prepare_sheet(self, name: str, columns: list[str]) -> None:
        if name not in self.workbook.sheetnames:
            self.workbook.create_sheet(name)
        sheet = self.workbook[name]
        header = _header(sheet)
        if not header and sheet.max_row <= 1:
            for column_index, column in enumerate(columns, start=1):
                sheet.cell(row=1, column=column_index, value=column)
            return
        if name == self.data_sheet and header == columns[:-1]:
            # Workbook from before the content column existed: add its header; old rows stay empty.
            sheet.cell(row=1, column=len(columns), value=CONTENT_COLUMN)
            return
        if header != columns:
            raise ValueError(
                f"Sheet '{name}' in {self.path} has columns {header}, but the parsing rules expect {columns}. "
                "Use a new output file (output_xlsx) or restore the previous parsing rules."
            )

    def check_writable(self) -> None:
        """Fail early (before any processing) if the workbook is locked, e.g. open in Excel on Windows."""
        if self.is_new:
            return
        try:
            with self.path.open("r+b"):
                pass
        except PermissionError:
            raise OSError(f"Cannot write {self.path}. Is it open in Excel? Close it and run again.") from None

    def append_data(self, values: dict[str, str], content: str) -> None:
        row = [values.get(column, "") for column in self.data_columns]
        _append_text_row(self.workbook[self.data_sheet], [*row, content])

    def remove_errors_for(self, source_type: str, source_location: str, message_identity: str) -> None:
        sheet = self.workbook[self.error_sheet]
        wanted = (source_type, source_location, message_identity)
        for row_index in range(sheet.max_row, 1, -1):
            row = tuple(sheet.cell(row=row_index, column=col).value for col in range(1, _ERROR_IDENTITY_COLUMNS + 1))
            if row == wanted:
                sheet.delete_rows(row_index)

    def upsert_error(self, error: dict[str, str]) -> None:
        """Keep exactly one error row per message, so retried failures don't pile up."""
        self.remove_errors_for(error["source_type"], error["source_location"], error["message_identity"])
        _append_text_row(self.workbook[self.error_sheet], [error.get(column, "") for column in ERROR_COLUMNS])

    def save(self) -> None:
        for name in (self.data_sheet, self.error_sheet):
            sheet = self.workbook[name]
            sheet.freeze_panes = "A2"
            sheet.auto_filter.ref = sheet.dimensions
        self.path.parent.mkdir(parents=True, exist_ok=True)
        temp_path = self.path.with_name(f".{self.path.name}.tmp")
        try:
            self.workbook.save(temp_path)
            os.replace(temp_path, self.path)
        except PermissionError:
            raise OSError(f"Cannot write {self.path}. Is it open in Excel? Close it and run again.") from None
        finally:
            temp_path.unlink(missing_ok=True)
