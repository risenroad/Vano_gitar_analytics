from __future__ import annotations

from datetime import date, datetime
from typing import Any, List, Sequence, Tuple

import gspread
from google.oauth2.service_account import Credentials


SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]


def authorize_service_account(credentials_path: str, scopes: Sequence[str] = SCOPES) -> gspread.Client:
    creds = Credentials.from_service_account_file(credentials_path, scopes=scopes)
    return gspread.authorize(creds)


def read_worksheet_as_table(worksheet: gspread.Worksheet) -> Tuple[List[str], List[List[str]]]:
    """
    Returns:
      - headers: list of column names
      - rows: list of rows (each row is a list of strings) padded to len(headers)
    """
    values = worksheet.get_all_values()
    if not values:
        return [], []

    headers = [str(h).strip() for h in values[0]]
    rows: List[List[str]] = []
    for raw_row in values[1:]:
        row = [str(x) for x in raw_row]
        if len(row) < len(headers):
            row = row + [""] * (len(headers) - len(row))
        elif len(row) > len(headers):
            row = row[: len(headers)]
        rows.append(row)

    return headers, rows


def get_or_create_worksheet(
    spreadsheet: gspread.Spreadsheet,
    title: str,
    rows: int,
    cols: int,
) -> gspread.Worksheet:
    try:
        return spreadsheet.worksheet(title)
    except gspread.WorksheetNotFound:
        # gspread requires initial size; we overwrite all contents right after.
        return spreadsheet.add_worksheet(title=title, rows=max(rows, 1), cols=max(cols, 1))


def _to_sheet_cell(value: Any) -> Any:
    if value is None:
        return ""
    if isinstance(value, (datetime, date)):
        # ISO date is safely understood by Sheets and keeps "python date" semantics.
        return value.isoformat()
    return value


def write_table(
    worksheet: gspread.Worksheet,
    headers: List[str],
    rows: List[List[Any]],
) -> None:
    worksheet.clear()
    values: List[List[Any]] = [[_to_sheet_cell(x) for x in headers]]
    values.extend([[_to_sheet_cell(cell) for cell in row] for row in rows])

    # Full table update (small tables are the target of this project).
    worksheet.update("A1", values, value_input_option="USER_ENTERED")

