from __future__ import annotations

from datetime import date
from typing import Dict, Optional, Tuple

import requests


_CACHE: Dict[Tuple[str, str], Optional[float]] = {}


def _format_date(d: date) -> str:
    return d.isoformat()  # YYYY-MM-DD


def get_gel_rate_on_date(base_currency: str, purchase_date: date) -> Optional[float]:
    """
    Возвращает курс: 1 * base_currency = X * GEL на purchase_date.
    В первую очередь пытается взять историю с National Bank of Georgia
    через `api.businessonline.ge/api/rates/nbg/...`.
    """
    base = base_currency.upper().strip()
    cache_key = (base, _format_date(purchase_date))
    if cache_key in _CACHE:
        return _CACHE[cache_key]

    if base == "GEL":
        _CACHE[cache_key] = 1.0
        return 1.0

    # 1) NBG historical rates (to GEL)
    def _request_nbg(for_base: str) -> Optional[float]:
        # Исторические курсы: https://api.businessonline.ge/api/rates/nbg/{currency}/{startDate}/{endDate}
        url = (
            f"https://api.businessonline.ge/api/rates/nbg/{for_base}/"
            f"{_format_date(purchase_date)}/{_format_date(purchase_date)}"
        )

        try:
            resp = requests.get(url, timeout=20)
            resp.raise_for_status()
            payload = resp.json()
        except Exception as e:  # noqa: BLE001 - намеренно перехватываем для ETL-потока
            print(f"[WARN] NBG rate request failed ({for_base} -> GEL, {purchase_date}): {e}")
            return None

        try:
            if not isinstance(payload, list) or not payload:
                return None
            rate = payload[0].get("Rate")
            return float(rate) if rate is not None else None
        except Exception:
            return None

    # NBG endpoint для рубля иногда возвращает Currency=RUR, поэтому пробуем обе строки.
    if base == "RUB":
        value = _request_nbg("RUB")
        if value is None:
            value = _request_nbg("RUR")
    else:
        value = _request_nbg(base)

    if value is not None:
        _CACHE[cache_key] = value
        return value

    # 2) Fallback: exchangerate.host
    def _request_fallback(for_base: str) -> Optional[float]:
        url = (
            f"https://api.exchangerate.host/{_format_date(purchase_date)}"
            f"?base={for_base}&symbols=GEL"
        )
        try:
            resp = requests.get(url, timeout=20)
            resp.raise_for_status()
            payload = resp.json()
        except Exception as e:  # noqa: BLE001
            print(f"[WARN] FX fallback request failed ({for_base} -> GEL, {purchase_date}): {e}")
            return None

        rates = payload.get("rates") or {}
        rate = rates.get("GEL")
        if rate is None:
            return None
        try:
            return float(rate)
        except Exception:
            return None

    value = _request_fallback(base)
    if value is None and base == "RUB":
        value = _request_fallback("RUR")

    if value is None:
        print(f"[WARN] FX rate not found ({base} -> GEL, {purchase_date}).")

    _CACHE[cache_key] = value
    return value

