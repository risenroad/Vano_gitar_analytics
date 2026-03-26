"""
Лист для анализа абонементов по покупкам:
- дата покупки
- кол-во купленных занятий
- студент
- стоимость 1 занятия в лари
- плановая дата оплаты этого абонемента (берём из предыдущей покупки студента)
- количество дней опоздания покупки (actual - planned)
"""

from __future__ import annotations

from datetime import date
from typing import Any, Dict, List, Optional, Tuple

from transform import _parse_date_cell, _parse_number, _parse_int

ANALYSIS_HEADERS: List[str] = [
    "Дата покупки",
    "Кол-во занятий",
    "Ученик",
    "Стоимость 1 занятия в лари",
    "Плановая дата оплаты этого абонемента",
    "Количество дней опоздания покупки",
]


def _cell(row: List[Any], col: Dict[str, int], name: str) -> Any:
    i = col.get(name)
    if i is None or i >= len(row):
        return ""
    return row[i]


def build_subscriptions_analysis_table(
    processed_headers: List[str], processed_rows: List[List[Any]]
) -> Tuple[List[str], List[List[Any]]]:
    col = {h: i for i, h in enumerate(processed_headers)}
    needed = [
        "Ученик",
        "Дата покупки",
        "Кол-во занятий",
        "Стоимость 1 занятия в лари",
        "Дат след. платежа (РАСЧЕТ)",
    ]
    for name in needed:
        if name not in col:
            raise RuntimeError(
                f"subscriptions_analysis: в processed нет столбца «{name}»."
            )

    purchases_by_student: Dict[str, List[Tuple[date, int, List[Any]]]] = {}

    for row_i, row in enumerate(processed_rows):
        student = str(_cell(row, col, "Ученик")).strip()
        if not student:
            continue
        purchase_dt = _parse_date_cell(_cell(row, col, "Дата покупки"))
        if purchase_dt is None:
            continue
        purchases_by_student.setdefault(student, []).append(
            (purchase_dt, row_i, row)
        )

    out_rows: List[List[Any]] = []

    for student in sorted(purchases_by_student.keys()):
        items = sorted(purchases_by_student[student], key=lambda x: (x[0], x[1]))

        for idx, (purchase_dt, row_i, row) in enumerate(items):
            planned_dt: Optional[date] = None
            if idx > 0:
                prev_row = items[idx - 1][2]
                planned_dt = _parse_date_cell(
                    _cell(prev_row, col, "Дат след. платежа (РАСЧЕТ)")
                )

            planned_sessions = _parse_int(_cell(row, col, "Кол-во занятий"))
            cost_per_session_gel = _parse_number(
                _cell(row, col, "Стоимость 1 занятия в лари")
            )

            delay_days: Any = ""
            if planned_dt is not None:
                delay_days = (purchase_dt - planned_dt).days

            out_rows.append(
                [
                    purchase_dt,
                    planned_sessions if planned_sessions is not None else "",
                    student,
                    cost_per_session_gel if cost_per_session_gel is not None else "",
                    planned_dt if planned_dt is not None else "",
                    delay_days,
                ]
            )

    return ANALYSIS_HEADERS, out_rows

