"""
Агрегация по ученикам для листа `students` (последний абонемент по дате покупки).
"""

from __future__ import annotations

from datetime import date
from typing import Any, Dict, List, Optional, Set, Tuple

from transform import _parse_date_cell, _parse_int, _parse_number, resolve_attendance_date_for_subscription

STUDENTS_HEADERS: List[str] = [
    "Ученик",
    "Статус абонемента",
    "Количество занятий оставшихся по текущему абонементу",
    "Тип последнего купленного абонемента",
    "Дата последнего платежа",
    "Дата последнего занятия",
    "Дата планового следующего платежа",
    "Количество купленных занятий за всю историю",
    "Количество завершенных занятий",
    "Стоимость 1 занятия в лари",
]


def _cell(row: List[Any], col: Dict[str, int], name: str) -> Any:
    j = col.get(name)
    if j is None or j >= len(row):
        return ""
    return row[j]


def _max_attendance_date(row: List[Any], col: Dict[str, int], purchase_dt: Optional[date]) -> Optional[date]:
    """Максимальная дата среди столбцов 1..16 — то же разрешение даты, что и в transform."""
    best: Optional[date] = None
    for k in range(1, 17):
        key = str(k)
        if key not in col:
            continue
        val = _cell(row, col, key)
        if val is None or str(val).strip() == "":
            continue
        d = resolve_attendance_date_for_subscription(val, purchase_dt)
        if d is not None:
            if best is None or d > best:
                best = d
    return best


def _attendance_dates_set(
    row: List[Any], col: Dict[str, int], purchase_dt: Optional[date]
) -> Set[date]:
    """
    Набор уникальных распознанных дат занятий для текущей строки покупки
    (аналогично logic в `lesson_dates`).
    """
    dates: Set[date] = set()
    for k in range(1, 17):
        key = str(k)
        if key not in col:
            continue
        val = _cell(row, col, key)
        if val is None or str(val).strip() == "":
            continue
        d = resolve_attendance_date_for_subscription(val, purchase_dt)
        if d is not None:
            dates.add(d)
    return dates


def _parse_next_payment_date(value: Any) -> Optional[date]:
    if value is None or value == "":
        return None
    if isinstance(value, date):
        return value
    return _parse_date_cell(value)


def _subscription_status(type_value: Any, remaining_value: Any) -> str:
    """
    - если тип "раз в неделю" и остаток <= 1 -> "абонемент заканчивается"
    - если тип "дважды в неделю" и остаток <= 2 -> "абонемент заканчивается"
    """
    rem = _parse_int(remaining_value)
    if rem is None:
        return ""

    t = str(type_value).strip().lower().replace("ё", "е")
    is_twice = ("дваж" in t or "2 раза" in t or "два раза" in t) and "нед" in t
    is_once = ("раз в неделю" in t) and not is_twice

    if is_once and rem <= 1:
        return "абонемент заканчивается"
    if is_twice and rem <= 2:
        return "абонемент заканчивается"
    return ""


def build_students_table(
    processed_headers: List[str], processed_rows: List[List[Any]]
) -> Tuple[List[str], List[List[Any]]]:
    """
    Одна строка на ученика: берётся покупка с максимальной «Дата покупки»
    (при равенстве — нижняя строка в таблице).
    """
    col = {h: i for i, h in enumerate(processed_headers)}
    needed = [
        "Ученик",
        "Дата покупки",
        "Тип",
        "Кол-во занятий",
        "Отхожено",
        "Сумма",
        "Дат след. платежа (РАСЧЕТ)",
        "Стоимость 1 занятия в лари",
    ]
    for name in needed:
        if name not in col:
            raise RuntimeError(
                f"Лист students: в данных processed нет столбца «{name}». "
                "Проверьте порядок колонок после transform."
            )

    # Собираем все абонементы (строки) студента, чтобы выбрать самый новый (последний по дате покупки).
    # Правило: если один и тот же день встречается в нескольких абонементах — относим его к более новому.
    purchases_by_student: Dict[str, List[Tuple[date, int, List[Any]]]] = {}
    for row_idx, row in enumerate(processed_rows):
        student = str(_cell(row, col, "Ученик")).strip()
        if not student:
            continue
        purchase_dt = _parse_date_cell(_cell(row, col, "Дата покупки"))
        if purchase_dt is None:
            continue
        purchases_by_student.setdefault(student, []).append((purchase_dt, row_idx, row))

    rows_with_sort: List[Tuple[date, str, List[Any]]] = []

    for student in purchases_by_student.keys():
        purchases = sorted(purchases_by_student[student], key=lambda x: (x[0], x[1]))
        current_purchase_dt, current_row_idx, current_row = purchases[-1]

        # Кол-во купленных занятий за всю историю (суммируем «Кол-во занятий» по всем покупкам).
        purchased_total_sessions = 0
        for _purchase_dt, _row_idx, row in purchases:
            planned_sessions = _parse_int(_cell(row, col, "Кол-во занятий"))
            if planned_sessions is not None:
                purchased_total_sessions += planned_sessions

        # Все уникальные даты занятий студента (между абонементами дедуплицируются по дате).
        all_dates: Set[date] = set()
        for purchase_dt, row_idx, row in purchases:
            all_dates.update(_attendance_dates_set(row, col, purchase_dt))

        # Для самого нового абонемента считаем занятия только по его собственной строке:
        # если день присутствует и раньше, и сейчас — он всё равно относится к более новому.
        attendance_dates = _attendance_dates_set(current_row, col, current_purchase_dt)
        attended_for_current = len(attendance_dates)
        # Дата последнего занятия должна быть по всей истории студента,
        # даже если на текущем абонементе занятий ещё нет.
        last_lesson = max(all_dates) if all_dates else None
        next_pay = _parse_next_payment_date(
            _cell(current_row, col, "Дат след. платежа (РАСЧЕТ)")
        )

        planned = _parse_int(_cell(current_row, col, "Кол-во занятий"))
        remaining: Any = ""
        if planned is not None:
            remaining = max(0, planned - attended_for_current)
        # purchased_total_sessions уже посчитано выше

        # Количество завершенных занятий со студентом (уникальные даты по дедупликации из lesson_dates).
        completed_sessions = len(all_dates)

        cost_per_session_gel = _parse_number(
            _cell(current_row, col, "Стоимость 1 занятия в лари")
        )
        row = [
            student,
            _subscription_status(_cell(current_row, col, "Тип"), remaining),
            remaining,
            str(_cell(current_row, col, "Тип")).strip(),
            current_purchase_dt if current_purchase_dt is not None else "",
            last_lesson if last_lesson is not None else "",
            next_pay if next_pay is not None else "",
            purchased_total_sessions if purchased_total_sessions != 0 else "",
            completed_sessions,
            cost_per_session_gel if cost_per_session_gel is not None else "",
        ]

        sort_last = last_lesson if last_lesson is not None else date(1, 1, 1)
        rows_with_sort.append((sort_last, student, row))

    # Сортируем: от самой поздней даты последнего занятия к самой ранней.
    rows_with_sort.sort(key=lambda x: (x[0], x[1]), reverse=True)
    out_rows: List[List[Any]] = [r for _, __, r in rows_with_sort]

    return STUDENTS_HEADERS, out_rows
