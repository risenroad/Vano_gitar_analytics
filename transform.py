from __future__ import annotations

from datetime import date, timedelta
from typing import Any, Dict, List, Optional, Sequence, Tuple
import re

import pandas as pd

from fx_rates import get_gel_rate_on_date


def _parse_date_cell(value: Any, default_year: Optional[int] = None) -> Optional[date]:
    if value is None:
        return None

    s = str(value).strip()
    if not s:
        return None

    # Google Sheets может отдавать даты в колонках посещений в формате dd.mm без года.
    # Примеры: "17.02", "21.02", "31.03".
    m = re.match(r"^(\d{1,2})[.\-/](\d{1,2})$", s)
    if m:
        day = int(m.group(1))
        month = int(m.group(2))
        year = default_year if default_year is not None else date.today().year
        try:
            return date(year, month, day)
        except Exception:
            return None

    # Полная текстовая дата со второй парой как месяц: dd.mm.yyyy / dd-mm-yyyy / dd/mm/yyyy.
    m_full = re.match(r"^(\d{1,2})[.\-/](\d{1,2})[.\-/](\d{2,4})$", s)
    if m_full:
        day = int(m_full.group(1))
        month = int(m_full.group(2))
        year_raw = int(m_full.group(3))
        year = year_raw + 2000 if year_raw < 100 else year_raw
        try:
            return date(year, month, day)
        except Exception:
            return None

    # ISO-подобная полная дата: yyyy-mm-dd / yyyy.mm.dd / yyyy/mm/dd.
    # Здесь вторая группа — месяц, третья — день.
    m_iso = re.match(r"^(\d{4})[.\-/](\d{1,2})[.\-/](\d{1,2})$", s)
    if m_iso:
        year = int(m_iso.group(1))
        month = int(m_iso.group(2))
        day = int(m_iso.group(3))
        try:
            return date(year, month, day)
        except Exception:
            return None

    # Google Sheets иногда отдаёт "дату" как число (serial).
    # Для современных дат serial обычно порядка 40000-50000.
    # Если число маленькое (например 1..31), это чаще всего не дата.
    try:
        if re.match(r"^\d+(\.\d+)?$", s):
            serial = float(s)
            if serial > 10000:
                base = date(1899, 12, 30)  # Excel-compatible epoch
                return base + timedelta(days=serial)
            # Не похоже на serial-дату
            return None
    except Exception:
        # Переходим к текстовому парсингу
        pass

    # Pandas хорошо вытягивает разные форматы (например: 25.03.2026 и 2026-03-25),
    # но для неоднозначных записей лучше попробовать 2 режима.
    ts = pd.to_datetime(s, dayfirst=True, errors="coerce")
    if ts is pd.NaT:
        ts = pd.to_datetime(s, dayfirst=False, errors="coerce")
    if ts is pd.NaT:
        return None
    return ts.to_pydatetime().date()


# Если |занятие − оплата| по «стандартному» году >= столько дней — перекидываем год дд.мм:
# для подавляющего большинства случаев реальный год занятия совпадает с годом оплаты.
# 98 и меньше — оставляем стандартный год.
ATTENDANCE_DDMM_REFINE_MIN_DAYS = 99


def resolve_attendance_date_for_subscription(
    raw_att: Any, purchase_dt: Optional[date]
) -> Optional[date]:
    """
    Для дд.мм без года: сначала год как при «занятии после оплаты»
    (месяц занятия < месяца покупки → следующий календарный год).
    Если |занятие − оплата| >= ATTENDANCE_DDMM_REFINE_MIN_DAYS — переносим дд.мм
    на год даты покупки (при невалидной дате — оставляем стандартный год).
    """
    if purchase_dt is None:
        return _parse_date_cell(raw_att)
    s = str(raw_att).strip()
    m = re.match(r"^(\d{1,2})[.\-/](\d{1,2})$", s)
    if not m:
        return _parse_date_cell(raw_att, default_year=purchase_dt.year)
    day = int(m.group(1))
    month = int(m.group(2))
    year_std = purchase_dt.year + (1 if month < purchase_dt.month else 0)
    d_std = _parse_date_cell(raw_att, default_year=year_std)
    if d_std is None:
        return None
    if abs((purchase_dt - d_std).days) < ATTENDANCE_DDMM_REFINE_MIN_DAYS:
        d_final = d_std
    else:
        # При срабатывании порога пересчитываем год на год даты покупки.
        # Это исправляет неверный год у dd.mm.
        try:
            d_final = date(purchase_dt.year, month, day)
        except ValueError:
            d_final = d_std

    # Доп. условие: если между датой покупки и занятием есть смена года,
    # то год занятия должен быть -1 (возвращаем в год покупки).
    if d_final.year == purchase_dt.year + 1:
        try:
            return date(purchase_dt.year, month, day)
        except ValueError:
            return d_final

    # Отдельный кейс на смене года:
    # если занятие в декабре, а абонемент оплачен в январе/феврале,
    # то год занятия должен быть предыдущим (purchase.year - 1).
    if month == 12 and purchase_dt.month in (1, 2):
        try:
            return date(purchase_dt.year - 1, month, day)
        except ValueError:
            return d_final

    return d_final


def _parse_number(value: Any) -> Optional[float]:
    if value is None:
        return None
    s = str(value).strip()
    if not s:
        return None

    # Убираем пробелы-разделители и меняем запятую на точку.
    s = s.replace("\u00a0", " ").replace(" ", "")
    s = s.replace(",", ".")

    try:
        return float(s)
    except ValueError:
        return None


def _parse_int(value: Any) -> Optional[int]:
    n = _parse_number(value)
    if n is None:
        return None
    if n == 0:
        return 0
    return int(round(n))


def _normalize_type(value: Any) -> str:
    s = str(value).strip().lower()
    s_norm = s.replace("ё", "е").strip()

    # Возможны вариации написания: "раз в неделю", "дважды в неделю",
    # "два раза в неделю", "2 раза в неделю", и т.д.
    if "разовое" in s_norm or "разово" in s_norm:
        return "one_time"

    two_week_signals = [
        "дважд",  # "дважды"
        "два раза",  # "два раза"
        "2 раза",  # "2 раза"
        "два раза в неделю",
        "2 раза в неделю",
    ]
    if ("нед" in s_norm) and any(sig in s_norm for sig in two_week_signals):
        return "two_week"

    # Также: "два в неделю", "2 раза/нед" и т.п.
    if ("нед" in s_norm) and ("два" in s_norm or "2" in s_norm) and ("раз" in s_norm):
        return "two_week"

    if ("раз" in s_norm) and ("нед" in s_norm):
        return "once_week"

    return s_norm


def _normalize_currency(value: Any) -> str:
    s = str(value).strip().lower()
    if not s:
        return ""

    # В таблице ожидаются: "Лари", "Рубли", "Доллары".
    # Важно: "доллары" содержит подстроку "лар", поэтому проверяем точнее.
    if "лар" in s and "дол" not in s and "руб" not in s:
        # чаще всего это "лари"
        return "GEL"
    if "руб" in s:
        return "RUB"
    if "дол" in s:
        return "USD"
    return s.upper()


def _is_non_empty_cell(value: Any) -> bool:
    if value is None:
        return False
    s = str(value).strip()
    return bool(s)


# Порядок столбцов в итоговом листе `processed` (как в ТЗ).
PROCESSED_COLUMN_ORDER: List[str] = [
    "Дата покупки",
    "Ученик",
    "Тип",
    "Кол-во занятий",
    "Отхожено",
    "Разница между кол-вом оплаченных и отхоженных занятий",
    "Дат след. платежа (РАСЧЕТ)",
    "Дат след. платежа (ФАКТ)",
    "Дней между ФАКТ и РАСЧЕТ",
    "Сумма",
    "Валюта",
    "Стоимость 1 занятия в валюте платежа",
    "Курс лари на день платежа",
    "Сумма в лари",
    "Стоимость 1 занятия в лари",
    "Банк",
    "Комментарий",
] + [str(i) for i in range(1, 17)]

_LEGACY_DIFF_COL = "Разница между датой ФАКТ и датой РАСЧЕТ платежа (дней)"

# Схема исходного листа (RAW): только вводные данные, без расчётных столбцов.
# Все производные поля добавляются при сборке `processed`.
RAW_INPUT_COLUMNS: List[str] = [
    "Дата покупки",
    "Ученик",
    "Тип",
    "Кол-во занятий",
    "Сумма",
    "Валюта",
    "Банк",
    "Комментарий",
] + [str(i) for i in range(1, 17)]

# Минимум столбцов, без которых расчёт невозможен (остальные из RAW_INPUT_COLUMNS могут отсутствовать — тогда в processed будут пустые ячейки).
REQUIRED_RAW_FOR_ETL: List[str] = [
    "Дата покупки",
    "Ученик",
    "Тип",
    "Кол-во занятий",
    "Сумма",
    "Валюта",
]


def _reorder_processed_table(
    headers: List[str], rows: List[List[Any]]
) -> Tuple[List[str], List[List[Any]]]:
    """Приводит заголовки и строки к фиксированному порядку столбцов."""
    col_index = {h: i for i, h in enumerate(headers)}
    # Старый длинный заголовок колонки разницы дат -> новый.
    if _LEGACY_DIFF_COL in col_index and "Дней между ФАКТ и РАСЧЕТ" not in col_index:
        col_index["Дней между ФАКТ и РАСЧЕТ"] = col_index[_LEGACY_DIFF_COL]

    new_headers = list(PROCESSED_COLUMN_ORDER)
    new_rows: List[List[Any]] = []
    for row in rows:
        new_row: List[Any] = []
        for h in PROCESSED_COLUMN_ORDER:
            if h in col_index:
                j = col_index[h]
                new_row.append(row[j] if j < len(row) else "")
            else:
                new_row.append("")
        new_rows.append(new_row)
    return new_headers, new_rows


def transform_table(headers: List[str], rows: List[List[str]]) -> Tuple[List[str], List[List[Any]]]:
    """
    Преобразует RAW-таблицу в таблицу для листа processed.

    Исходный лист не содержит расчётных колонок — только поля из `RAW_INPUT_COLUMNS`
    (см. также `REQUIRED_RAW_FOR_ETL`). Производные столбцы добавляются к строке и
    выводятся в порядке `PROCESSED_COLUMN_ORDER`.
    """
    headers = [str(h).strip() for h in headers]

    def _norm_col_name(s: str) -> str:
        return str(s).strip().lower().replace("ё", "е")

    header_norms = {_norm_col_name(h) for h in headers}
    missing_req = [c for c in REQUIRED_RAW_FOR_ETL if _norm_col_name(c) not in header_norms]
    if missing_req:
        raise RuntimeError(
            "В исходном листе не хватает обязательных столбцов: "
            f"{missing_req}. Ожидается набор без расчётных полей (см. RAW_INPUT_COLUMNS в transform.py)."
        )

    computed_columns = [
        "Отхожено",
        "Разница между кол-вом оплаченных и отхоженных занятий",
        "Дат след. платежа (РАСЧЕТ)",
        "Дат след. платежа (ФАКТ)",
        "Дней между ФАКТ и РАСЧЕТ",
        "Стоимость 1 занятия в валюте платежа",
        "Курс лари на день платежа",
        "Сумма в лари",
        "Стоимость 1 занятия в лари",
    ]

    # Если какой-то столбец не нашёлся в RAW — добавим его в конец.
    output_headers = list(headers)
    existing_norm = {_norm_col_name(h) for h in output_headers}
    for col in computed_columns:
        if _norm_col_name(col) not in existing_norm:
            output_headers.append(col)
            existing_norm.add(_norm_col_name(col))
    output_col_count = len(output_headers)

    def ensure_row_len(r: List[Any]) -> List[Any]:
        if len(r) < output_col_count:
            return r + [""] * (output_col_count - len(r))
        return r[:output_col_count]

    # Индексы по RAW.
    idx = {h: i for i, h in enumerate(output_headers)}
    att_cols = [i for i, h in enumerate(output_headers) if h.isdigit()]

    def _norm_col(s: str) -> str:
        return _norm_col_name(s)

    headers_norm = [_norm_col(h) for h in output_headers]

    def _find_col(target: str) -> Optional[int]:
        t = _norm_col(target)
        # 1) точное совпадение
        for i, hn in enumerate(headers_norm):
            if hn == t:
                return i
        # 2) fallback по вхождению
        for i, hn in enumerate(headers_norm):
            if t and t in hn:
                # минимизируем риск перехватить "Сумма в лари" вместо "Сумма"
                if t == "сумма" and hn != "сумма":
                    continue
                return i
        return None

    col_purchase_date = _find_col("Дата покупки")
    col_student = _find_col("Ученик")
    col_type = _find_col("Тип")
    col_planned_sessions = _find_col("Кол-во занятий")
    col_sum = _find_col("Сумма")
    col_currency = _find_col("Валюта")

    # Индексы вычисляемых полей.
    i_attended = idx.get("Отхожено")
    i_next_calc = idx.get("Дат след. платежа (РАСЧЕТ)")
    i_next_fact = idx.get("Дат след. платежа (ФАКТ)")
    i_cost_per_curr = idx.get("Стоимость 1 занятия в валюте платежа")
    i_fx_rate = idx.get("Курс лари на день платежа")
    i_sum_gel = idx.get("Сумма в лари")
    i_cost_per_gel = idx.get("Стоимость 1 занятия в лари")
    i_diff_fact_minus_calc = idx.get("Дней между ФАКТ и РАСЧЕТ")
    if i_diff_fact_minus_calc is None:
        i_diff_fact_minus_calc = idx.get(_LEGACY_DIFF_COL)
    i_paid_minus_attended = idx.get(
        "Разница между кол-вом оплаченных и отхоженных занятий"
    )
    first_session_col_idx = idx.get("1")

    # Подготовим данные построчно, чтобы потом вычислить "ФАКТ"
    # через следующую покупку того же ученика.
    parsed: List[Dict[str, Any]] = []
    debug_bad_date_printed = 0
    for row_i, raw_row in enumerate(rows):
        r = ensure_row_len(list(raw_row))

        purchase_dt = _parse_date_cell(r[col_purchase_date]) if col_purchase_date is not None else None
        student = str(r[col_student]).strip() if col_student is not None else ""
        type_norm = _normalize_type(r[col_type]) if col_type is not None else ""
        planned_sessions = _parse_int(r[col_planned_sessions]) if col_planned_sessions is not None else None
        sum_amount = _parse_number(r[col_sum]) if col_sum is not None else None
        currency_norm = _normalize_currency(r[col_currency]) if col_currency is not None else ""

        attended_count = 0
        attendance_dates: List[date] = []
        attendance_by_col: Dict[int, date] = {}
        first_bad_attendance_value: Optional[str] = None
        for ac in att_cols:
            if _is_non_empty_cell(r[ac]):
                attended_count += 1
                raw_att = r[ac]
                d = resolve_attendance_date_for_subscription(raw_att, purchase_dt)
                if d is not None:
                    attendance_dates.append(d)
                    attendance_by_col[ac] = d
                else:
                    # Запомним пример, чтобы понять формат значений.
                    if first_bad_attendance_value is None:
                        first_bad_attendance_value = str(raw_att).strip()

        # Дата первого занятия берётся строго из колонки "1".
        first_attendance_date: Optional[date] = None
        first_session_raw_value: Optional[str] = None
        first_session_ddmm: Optional[Tuple[int, int]] = None
        if first_session_col_idx is not None:
            first_session_raw = r[first_session_col_idx] if first_session_col_idx < len(r) else None
            if _is_non_empty_cell(first_session_raw):
                first_session_raw_value = str(first_session_raw).strip()
                m = re.match(
                    r"^(\d{1,2})[.\-/](\d{1,2})$", str(first_session_raw).strip()
                )
                if m:
                    first_session_ddmm = (int(m.group(1)), int(m.group(2)))
                first_attendance_date = resolve_attendance_date_for_subscription(
                    first_session_raw, purchase_dt
                )

        if (
            attended_count > 0
            and first_attendance_date is None
            and debug_bad_date_printed < 10
        ):
            print(
                f"[WARN] Attendance dates failed to parse "
                f"(attended_count={attended_count}, purchase_dt={purchase_dt}, "
                f"student={student}, bad_example={first_bad_attendance_value}).",
                flush=True,
            )
            debug_bad_date_printed += 1

        # Если дата “уехала” в 1900-е, это почти всегда признак неправильного
        # парсинга содержимого колонок посещений.
        if first_attendance_date is not None and first_attendance_date.year < 2000:
            print(
                f"[WARN] Suspicious attendance date parsed: {first_attendance_date} "
                f"(student={student}, purchase_dt={purchase_dt}).",
                flush=True,
            )

        parsed.append(
            {
                "row_index": row_i,
                "purchase_dt": purchase_dt,
                "student": student,
                "type_norm": type_norm,
                "planned_sessions": planned_sessions,
                "sum_amount": sum_amount,
                "currency_norm": currency_norm,
                "attended_count": attended_count,
                "first_attendance_date": first_attendance_date,
                "first_session_raw_value": first_session_raw_value,
                "first_session_ddmm": first_session_ddmm,
                "attendance_by_col": attendance_by_col,
                "row": r,
            }
        )

    # next purchase mapping for "Дат след. платежа (ФАКТ)"
    next_purchase_dt_by_row: Dict[int, Optional[date]] = {}
    # Для каждого ученика сортируем покупки по дате, а при равенстве по индексу в RAW.
    purchases_by_student: Dict[str, List[Tuple[date, int]]] = {}
    for item in parsed:
        if not item["student"]:
            continue
        if item["purchase_dt"] is None:
            continue
        purchases_by_student.setdefault(item["student"], []).append((item["purchase_dt"], item["row_index"]))

    for student, items in purchases_by_student.items():
        items_sorted = sorted(items, key=lambda x: (x[0], x[1]))
        for pos, (_dt, row_index) in enumerate(items_sorted):
            next_dt = items_sorted[pos + 1][0] if pos + 1 < len(items_sorted) else None
            next_purchase_dt_by_row[row_index] = next_dt

    # Теперь посчитаем все вычисляемые поля
    out_rows: List[List[Any]] = []
    for item in parsed:
        r = list(item["row"])
        for ac, d in item.get("attendance_by_col", {}).items():
            if ac < len(r):
                r[ac] = d

        paid_sessions = item["planned_sessions"]
        next_fact_date: Optional[date] = next_purchase_dt_by_row.get(
            item["row_index"], None
        )

        # 1) Отхожено
        if i_attended is not None:
            r[i_attended] = item["attended_count"]

        # 1.1) Разница (оплачено - отхожено)
        if i_paid_minus_attended is not None:
            if paid_sessions in (None, 0):
                r[i_paid_minus_attended] = ""
            else:
                r[i_paid_minus_attended] = float(paid_sessions) - float(item["attended_count"])

        # 2) Дат след. платежа (РАСЧЕТ)
        if i_next_calc is not None:
            next_calc: Optional[date] = None
            if item["planned_sessions"] not in (None, 0):
                if item["type_norm"] == "once_week":
                    weeks = float(item["planned_sessions"])
                elif item["type_norm"] == "two_week":
                    weeks = float(item["planned_sessions"]) / 2.0
                elif item["type_norm"] == "one_time":
                    weeks = None
                else:
                    weeks = None

                if weeks is not None:
                    # weeks может быть X.5 => добавляем X*7 дней, включая 3.5 дня для полу-недели.
                    if item.get("first_session_ddmm") is not None and item["purchase_dt"] is not None:
                        day, month = item["first_session_ddmm"]
                        candidate_first_dates: List[date] = []
                        for year in (
                            item["purchase_dt"].year - 1,
                            item["purchase_dt"].year,
                            item["purchase_dt"].year + 1,
                        ):
                            try:
                                candidate_first_dates.append(date(year, month, day))
                            except Exception:
                                continue

                        if candidate_first_dates and item["purchase_dt"] is not None:
                            # Основное ограничение: плановый платеж должен быть
                            # близко (примерно 1-2 месяца) к дате покупки абонемента.
                            purchase_dt = item["purchase_dt"]
                            MAX_OFFSET_DAYS = 75  # ~ 1.5-2.5 месяца

                            def _score(first_date: date) -> int:
                                cand_next = first_date + timedelta(days=weeks * 7.0)
                                d_purchase_next = abs((cand_next - purchase_dt).days)
                                # Сильный штраф, если плановый платеж слишком далеко от покупки.
                                penalty = max(0, d_purchase_next - MAX_OFFSET_DAYS)
                                # `ФАКТ` используется только как небольшой tie-breaker.
                                d_fact = (
                                    abs((next_fact_date - cand_next).days)
                                    if next_fact_date is not None
                                    else 0
                                )
                                return int(d_purchase_next + penalty * 10 + d_fact * 0.05)

                            best_first = min(candidate_first_dates, key=_score)
                            next_calc = best_first + timedelta(days=weeks * 7.0)
                        elif candidate_first_dates:
                            # ФАКТ нет -> используем текущий (эвристический) first_attendance_date.
                            if item["first_attendance_date"] is not None:
                                next_calc = (
                                    item["first_attendance_date"]
                                    + timedelta(days=weeks * 7.0)
                                )
                            else:
                                # запасной вариант: берем первую кандидату по году
                                next_calc = (
                                    candidate_first_dates[0]
                                    + timedelta(days=weeks * 7.0)
                                )
                    elif item["first_attendance_date"] is not None:
                        next_calc = item["first_attendance_date"] + timedelta(
                            days=weeks * 7.0
                        )

            r[i_next_calc] = next_calc if next_calc is not None else ""
            next_calc_date = next_calc
        else:
            next_calc_date = None

        # 3) Дат след. платежа (ФАКТ)
        if i_next_fact is not None:
            next_fact = next_purchase_dt_by_row.get(item["row_index"], None)
            r[i_next_fact] = next_fact if next_fact is not None else ""
            next_fact_date = next_fact

        # 3.5) Разница между плановым и фактическим платежом (дней)
        if i_diff_fact_minus_calc is not None:
            if next_calc_date is not None and next_fact_date is not None:
                r[i_diff_fact_minus_calc] = (next_fact_date - next_calc_date).days
            else:
                r[i_diff_fact_minus_calc] = ""

        # 4) Стоимость 1 занятия в валюте платежа
        if i_cost_per_curr is not None:
            if item["sum_amount"] is None or item["planned_sessions"] in (None, 0):
                r[i_cost_per_curr] = ""
            else:
                r[i_cost_per_curr] = item["sum_amount"] / float(item["planned_sessions"])

        # 5) Сумма в лари
        if i_sum_gel is not None:
            sum_gel: Optional[float] = None
            fx_rate: Optional[float] = None
            if item["sum_amount"] is not None:
                if item["currency_norm"] == "GEL":
                    sum_gel = item["sum_amount"]
                elif item["currency_norm"] in ("USD", "RUB"):
                    if item["purchase_dt"] is None:
                        sum_gel = None
                    else:
                        fx_rate = get_gel_rate_on_date(item["currency_norm"], item["purchase_dt"])
                        if fx_rate is None:
                            sum_gel = None
                        else:
                            sum_gel = item["sum_amount"] * fx_rate
                else:
                    sum_gel = None

            r[i_sum_gel] = sum_gel if sum_gel is not None else ""

            # 5.1) Курс лари на день платежа (по дате следующего платежа РАСЧЕТ)
            if i_fx_rate is not None:
                if item["currency_norm"] == "GEL":
                    r[i_fx_rate] = 1.0
                elif item["currency_norm"] not in ("RUB", "USD"):
                    # Неизвестная валюта (пусто/мусор) -> курс не запрашиваем.
                    r[i_fx_rate] = ""
                elif next_calc_date is None:
                    r[i_fx_rate] = ""
                    print(
                        "[WARN] FX rate for payment date skipped: "
                        f"payment_date is empty (student={item['student']}, row={item['row_index']}).",
                        flush=True,
                    )
                elif next_calc_date.year < 2000:
                    r[i_fx_rate] = ""
                    print(
                        "[WARN] FX rate for payment date skipped: "
                        f"payment_date suspicious ({next_calc_date}, student={item['student']}, row={item['row_index']}).",
                        flush=True,
                    )
                else:
                    rate_on_payment = get_gel_rate_on_date(
                        item["currency_norm"], next_calc_date
                    )
                    r[i_fx_rate] = (
                        rate_on_payment if rate_on_payment is not None else ""
                    )

        # 6) Стоимость 1 занятия в лари
        if i_cost_per_gel is not None:
            if r[i_sum_gel] == "" or item["planned_sessions"] in (None, 0):
                r[i_cost_per_gel] = ""
            else:
                # r[i_sum_gel] может быть float, но безопасно приводим.
                sum_gel_val = float(r[i_sum_gel])
                r[i_cost_per_gel] = sum_gel_val / float(item["planned_sessions"])

        out_rows.append(r)

    # Фиксированный порядок столбцов в `processed`.
    return _reorder_processed_table(output_headers, out_rows)

