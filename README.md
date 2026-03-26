# Vano Guitar Analytics

## Структура проекта

```
.
├── main.py             # Точка входа (ETL)
├── sheets_client.py   # Чтение/запись Google Sheets
├── transform.py       # Все правила расчётов
├── fx_rates.py        # Внешний API для курсов валют
├── requirements.txt   # Зависимости Python
├── .env.example       # Шаблон переменных окружения
├── README.md
```

## Установка

```bash
python -m venv .venv
source .venv/bin/activate   # Windows: .venv\Scripts\activate
pip install -r requirements.txt
```

## Настройка

Скопируйте `.env.example` в `.env` и при необходимости заполните переменные:

```bash
cp .env.example .env
```

## Запуск
```bash
python main.py
```

## Что обновляется в Google Sheets
При каждом запуске ETL полностью пересчитывает данные и перезаписывает лист `processed` в той же таблице (лист `raw` не изменяется).
