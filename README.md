# Vano Guitar Analytics

ETL-скрипт для аналитики абонементов в Google Sheets.
Берет исходные данные из RAW-листа, рассчитывает производные поля и обновляет несколько аналитических вкладок.

## Что делает ETL

При запуске `python main.py` скрипт:

- читает исходный лист (`RAW_SHEET_NAME`);
- пересчитывает данные по абонементам и платежам;
- полностью перезаписывает итоговые листы:
  - `processed` — расчетный слой с нормализованными полями;
  - `students` — срез по студентам (статус, остаток, последние даты, агрегаты);
  - `lesson_dates` — занятия по датам с дедупликацией по правилу "в пользу более нового абонемента";
  - `abonements_analysis` — анализ покупок и опозданий по плановой дате оплаты.

Исходный RAW-лист не изменяется.

## Структура проекта

```
.
├── main.py                      # orchestration ETL
├── transform.py                 # ключевая логика расчетов и нормализации
├── students.py                  # сборка листа students
├── lesson_dates.py              # сборка листа lesson_dates
├── subscriptions_analysis.py    # сборка листа abonements_analysis
├── sheets_client.py             # чтение/запись Google Sheets
├── fx_rates.py                  # получение валютных курсов
├── requirements.txt
├── .env.example
└── README.md
```

## Установка

```bash
python -m venv .venv
source .venv/bin/activate   # Windows: .venv\Scripts\activate
pip install -r requirements.txt
```

## Настройка

1. Скопируйте шаблон:

```bash
cp .env.example .env
```

2. Заполните переменные в `.env`:

- `GOOGLE_CREDENTIALS_PATH` — путь к service account JSON;
- `GOOGLE_SHEET_ID` — ID Google таблицы;
- `RAW_SHEET_NAME` — имя исходного листа;
- `PROCESSED_SHEET_NAME` — имя листа processed;
- `STUDENTS_SHEET_NAME` — имя листа students;
- `LESSON_DATES_SHEET_NAME` — имя листа lesson_dates;
- `SUBSCRIPTIONS_ANALYSIS_SHEET_NAME` — имя листа анализа абонементов.

## Запуск

```bash
python main.py
```

## Ключевые правила в расчетах

- даты посещений нормализуются с учетом неоднозначных форматов;
- даты занятий дедуплицируются (один студент, один день, при конфликте привязка к более новому абонементу);
- для students считаются остатки и статусы абонементов;
- для abonements_analysis считается плановая дата оплаты и дни опоздания покупки.
