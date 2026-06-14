"""Helpers for converting the source IDUs workbook."""

from __future__ import annotations

import pandas as pd

def parse_idus_source_date_series(series: pd.Series, swap_excel_dates: bool = False) -> pd.Series:
    """Parse the source IDUs sheet dates safely.

    The source workbook mixes text dates written as DD/MM/YYYY with Excel date
    cells.  In the main visit date column, some Excel datetime cells were stored
    as MM/DD/YYYY while the intended input was DD/MM/YYYY, so that column needs
    day/month swapping.  Follow-up date columns are already stored correctly.
    """
    import datetime as _dt
    import re as _re

    def _parse_one(value):
        if pd.isna(value):
            return pd.NaT

        if isinstance(value, pd.Timestamp):
            value = value.to_pydatetime()

        if isinstance(value, (_dt.datetime, _dt.date)):
            if swap_excel_dates:
                try:
                    return pd.Timestamp(value.year, value.day, value.month)
                except Exception:
                    return pd.Timestamp(value)
            return pd.Timestamp(value)

        text = str(value).strip()
        m = _re.match(r"^(\d{1,2})[/-](\d{1,2})[/-](\d{4})$", text)
        if m:
            day, month, year = map(int, m.groups())
            try:
                return pd.Timestamp(year, month, day)
            except Exception:
                return pd.NaT

        return pd.to_datetime(text, errors='coerce', dayfirst=True)

    return series.apply(_parse_one)


def format_idus_source_date(series: pd.Series, swap_excel_dates: bool = False) -> pd.Series:
    return parse_idus_source_date_series(series, swap_excel_dates=swap_excel_dates).dt.strftime('%Y-%m-%d')


# ════════════════════════════════════════
