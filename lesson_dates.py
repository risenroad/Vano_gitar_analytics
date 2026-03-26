"""
Таблица всех дат занятий по строкам processed (каждая покупка абонемента — свой контекст).

Правило 99 дней: при |занятие−оплата| ≥ 99 по стандартному году дд.мм переносятся на
**предыдущий** календарный год относительно покупки (иначе год покупки), см. resolve_attendance_date_for_subscription.

«В долг» — только когда после разбора дата занятия *раньше* даты покупки.

Проверка в консоли: для занятий *после* оплаты — не более 2 календарных месяцев; сюда строки не фильтруем.
"""

from __future__ import annotations

from datetime import date
from typing import Any, Dict, List, Set, Tuple

from transform import _parse_date_cell, _parse_number, resolve_attendance_date_for_subscription

LESSON_DATES_HEADERS: List[str] = [
    "Ученик",
    "Дата покупки абонемента",
    "Дата занятия по абонементу",
    "Стоимость 1 занятия в лари",
    "Был ли урок оплачен в долг (занятие раньше оплаты)",
]


def _calendar_months_after(purchase: date, lesson: date) -> int:
    return (lesson.year - purchase.year) * 12 + (lesson.month - purchase.month)


def build_lesson_dates_table(
    processed_headers: List[str], processed_rows: List[List[Any]]
) -> Tuple[List[str], List[List[Any]], List[str]]:
    """
    Возвращает заголовки, строки (даты как date-значения)
    и список сообщений о нарушениях правил (для вывода в консоль).
    """
    col = {h: i for i, h in enumerate(processed_headers)}
    if "Ученик" not in col or "Дата покупки" not in col:
        raise RuntimeError(
            "lesson_dates: в processed должны быть столбцы «Ученик» и «Дата покупки»."
        )

    att_cols: List[int] = []
    for h, i in col.items():
        if h.isdigit():
            att_cols.append(i)
    att_cols.sort(key=lambda j: int(processed_headers[j]))

    # chosen[(student, lesson_dt)] -> запись, относящая эту дату занятия к более новому абонементу.
    chosen: Dict[Tuple[str, date], Dict[str, Any]] = {}
    outliers: List[str] = []

    for row_i, row in enumerate(processed_rows):
        i_stu = col["Ученик"]
        student = str(row[i_stu] if i_stu < len(row) else "").strip()
        if not student:
            continue
        pd_raw = row[col["Дата покупки"]] if col["Дата покупки"] < len(row) else ""
        purchase_dt = _parse_date_cell(pd_raw)
        if purchase_dt is None:
            continue

        cost_per_session_gel = ""
        if "Стоимость 1 занятия в лари" in col:
            cost_per_session_gel = _parse_number(
                row[col["Стоимость 1 занятия в лари"]]
            )

        seen: Set[date] = set()
        for ac in att_cols:
            if ac >= len(row):
                continue
            raw = row[ac]
            if raw is None or str(raw).strip() == "":
                continue
            sraw = str(raw).strip()
            lesson_dt = resolve_attendance_date_for_subscription(raw, purchase_dt)

            if lesson_dt is None:
                outliers.append(
                    f"[строка processed ~{row_i + 2}] {student} | покупка {pd_raw} | "
                    f"ячейка «{processed_headers[ac]}»={sraw!r} — не удалось распознать дату занятия"
                )
                continue

            if lesson_dt in seen:
                continue
            seen.add(lesson_dt)

            key = (student, lesson_dt)
            candidate = {
                "purchase_dt": purchase_dt,
                "purchase_row_i": row_i,
                "purchase_raw": pd_raw,
                "cell_header": processed_headers[ac],
                "cell_raw": sraw,
                "cost_per_session_gel": cost_per_session_gel,
            }

            if key not in chosen:
                chosen[key] = candidate
            else:
                prev = chosen[key]
                # Оставляем более новый абонемент.
                # Если даты покупки равны — берем тот, что ниже в RAW (row_i больше).
                if candidate["purchase_dt"] > prev["purchase_dt"] or (
                    candidate["purchase_dt"] == prev["purchase_dt"]
                    and candidate["purchase_row_i"] > prev["purchase_row_i"]
                ):
                    chosen[key] = candidate

    out_rows: List[List[Any]] = []
    # Сначала сортируем выбранные записи, затем формируем таблицу и сообщения о нарушениях.
    chosen_items = [
        (student_lesson[0], student_lesson[1], meta)
        for student_lesson, meta in chosen.items()
    ]
    chosen_items.sort(key=lambda x: (x[1], x[0]), reverse=True)

    for student, lesson_dt, meta in chosen_items:
        purchase_dt = meta["purchase_dt"]
        out_rows.append(
            [
                student,
                purchase_dt,
                lesson_dt,
                meta["cost_per_session_gel"] if meta["cost_per_session_gel"] is not None else "",
                bool(lesson_dt < purchase_dt),
            ]
        )

        # Проверка нарушения >2 календарных месяцев выполняется для выбранного (оставленного) абонемента.
        if lesson_dt < purchase_dt:
            continue
        cm = _calendar_months_after(purchase_dt, lesson_dt)
        if cm > 2:
            outliers.append(
                f"[строка processed ~{meta['purchase_row_i'] + 2}] {student} | покупка {purchase_dt.isoformat()} | "
                f"ячейка «{meta['cell_header']}»={meta['cell_raw']!r} -> занятие {lesson_dt.isoformat()} "
                f"— между покупкой и занятием больше 2 месяцев "
                f"(разница по календарным месяцам: {cm})"
            )

    return LESSON_DATES_HEADERS, out_rows, outliers
