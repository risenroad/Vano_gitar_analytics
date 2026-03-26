"""
Vano Guitar Analytics — ETL-процесс для расчёта производных полей.
"""

from __future__ import annotations

import os
from pathlib import Path

from dotenv import load_dotenv

from sheets_client import (
    authorize_service_account,
    get_or_create_worksheet,
    read_worksheet_as_table,
    write_table,
)
from lesson_dates import build_lesson_dates_table
from students import build_students_table
from transform import transform_table
from subscriptions_analysis import build_subscriptions_analysis_table


def _require_env(name: str) -> str:
    value = os.getenv(name)
    if not value:
        raise RuntimeError(f"Missing required env var: {name}")
    return value


def _get_worksheet_case_insensitive(spreadsheet, title: str):
    """
    gspread ищет лист по точному совпадению названия.
    Чтобы уменьшить трение при вводе в .env, ищем без учёта регистра.
    """
    wanted = title.strip().lower()
    for ws in spreadsheet.worksheets():
        if ws.title.strip().lower() == wanted:
            return ws
    available = [ws.title for ws in spreadsheet.worksheets()]
    raise RuntimeError(
        f"Worksheet not found: `{title}`. Available worksheets: {available}"
    )


def main() -> None:
    # Загружаем .env относительно текущего файла, а не CWD.
    # Это предотвращает проблемы при запуске из другой папки.
    dotenv_path = Path(__file__).resolve().parent / ".env"
    load_dotenv(dotenv_path=dotenv_path)

    credentials_path = _require_env("GOOGLE_CREDENTIALS_PATH")
    sheet_id = _require_env("GOOGLE_SHEET_ID")
    raw_sheet_name = _require_env("RAW_SHEET_NAME")
    processed_sheet_name = _require_env("PROCESSED_SHEET_NAME")
    students_sheet_name = (os.getenv("STUDENTS_SHEET_NAME") or "students").strip()
    lesson_dates_sheet_name = (os.getenv("LESSON_DATES_SHEET_NAME") or "lesson_dates").strip()
    subscriptions_analysis_sheet_name = (
        os.getenv("SUBSCRIPTIONS_ANALYSIS_SHEET_NAME") or "abonements_analysis"
    ).strip()

    client = authorize_service_account(credentials_path)
    spreadsheet = client.open_by_key(sheet_id)

    raw_ws = _get_worksheet_case_insensitive(spreadsheet, raw_sheet_name)
    raw_headers, raw_rows = read_worksheet_as_table(raw_ws)
    if not raw_headers:
        print(f"[INFO] RAW sheet `{raw_sheet_name}` is empty. Nothing to process.")
        return

    out_headers, out_rows = transform_table(raw_headers, raw_rows)

    processed_ws = get_or_create_worksheet(
        spreadsheet=spreadsheet,
        title=processed_sheet_name,
        rows=len(out_rows) + 10,
        cols=len(out_headers) + 10,
    )
    write_table(processed_ws, out_headers, out_rows)

    print(
        f"[OK] Processed sheet updated: `{processed_sheet_name}` "
        f"(rows={len(out_rows)}, cols={len(out_headers)})"
    )

    stu_headers, stu_rows = build_students_table(out_headers, out_rows)
    students_ws = get_or_create_worksheet(
        spreadsheet=spreadsheet,
        title=students_sheet_name,
        rows=max(len(stu_rows) + 10, 20),
        cols=max(len(stu_headers) + 5, 10),
    )
    write_table(students_ws, stu_headers, stu_rows)
    print(
        f"[OK] Students sheet updated: `{students_sheet_name}` "
        f"(rows={len(stu_rows)}, cols={len(stu_headers)})"
    )

    ld_headers, ld_rows, lesson_outliers = build_lesson_dates_table(out_headers, out_rows)
    ld_ws = get_or_create_worksheet(
        spreadsheet=spreadsheet,
        title=lesson_dates_sheet_name,
        rows=max(len(ld_rows) + 10, 50),
        cols=max(len(ld_headers) + 3, 10),
    )
    write_table(ld_ws, ld_headers, ld_rows)
    print(
        f"[OK] Lesson dates sheet updated: `{lesson_dates_sheet_name}` "
        f"(rows={len(ld_rows)}, cols={len(ld_headers)})"
    )
    if lesson_outliers:
        print(
            "\n=== Занятия вне правила "
            "(после даты покупки — не больше 2 календарных месяцев; "
            "раньше оплаты = «в долг», не считается ошибкой) ==="
        )
        for line in lesson_outliers:
            print(line)
    else:
        print(
            "[OK] По всем распознанным датам занятий: правило "
            "(не больше 2 мес. после покупки для занятий после оплаты) не нарушено."
        )

    sa_headers, sa_rows = build_subscriptions_analysis_table(out_headers, out_rows)
    sa_ws = get_or_create_worksheet(
        spreadsheet=spreadsheet,
        title=subscriptions_analysis_sheet_name,
        rows=max(len(sa_rows) + 10, 50),
        cols=max(len(sa_headers) + 3, 10),
    )
    write_table(sa_ws, sa_headers, sa_rows)
    print(
        f"[OK] Subscriptions analysis sheet updated: `{subscriptions_analysis_sheet_name}` "
        f"(rows={len(sa_rows)}, cols={len(sa_headers)})"
    )


if __name__ == "__main__":
    main()