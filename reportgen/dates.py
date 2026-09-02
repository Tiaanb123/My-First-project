"""Date parsing and formatting helpers shared across the pipeline.

Everything here is locale-independent on purpose: the exports come from systems in
several countries and we can't rely on the machine's locale for month names.
"""
from __future__ import annotations

import datetime as dt
import re

MONTHS: dict[str, int] = {
    name: i
    for i, name in enumerate(
        (
            "january", "february", "march", "april", "may", "june",
            "july", "august", "september", "october", "november", "december",
        ),
        start=1,
    )
}
MONTHS_ABBR: dict[str, int] = {name[:3]: i for name, i in MONTHS.items()}
WEEKDAYS = ("Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday")
_MONTH_NAMES = (
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December",
)

# Excel stores dates as a day count from this epoch (the 1900 date system, with its
# historical leap-year-1900 quirk already baked into the offset when we use 1899-12-30).
_EXCEL_EPOCH = dt.datetime(1899, 12, 30)


def ordinal(n: int) -> str:
    """1 -> '1st', 2 -> '2nd', 11 -> '11th', 23 -> '23rd'."""
    if 11 <= (n % 100) <= 13:
        suffix = "th"
    else:
        suffix = {1: "st", 2: "nd", 3: "rd"}.get(n % 10, "th")
    return f"{n}{suffix}"


def format_subtitle(start: dt.date, end: dt.date) -> str:
    """Part A date-range subtitle, e.g. 'Monday 24th August to Sunday 30th August 2026'.

    The year is shown once, on the end date, matching the established report format.
    """
    left = f"{WEEKDAYS[start.weekday()]} {ordinal(start.day)} {_MONTH_NAMES[start.month - 1]}"
    right = (
        f"{WEEKDAYS[end.weekday()]} {ordinal(end.day)} "
        f"{_MONTH_NAMES[end.month - 1]} {end.year}"
    )
    return f"{left} to {right}"


def filename_tag(start: dt.date, end: dt.date) -> str:
    """Filesystem-safe range token for output filenames: '24.08.2026_to_30.08.2026'."""
    return f"{start:%d.%m.%Y}_to_{end:%d.%m.%Y}"


def parse_english_date(text: str) -> dt.date:
    """Parse 'August 30, 2026' or '30 August 2026' (comma optional, trailing dot ok)."""
    cleaned = text.strip().rstrip(".").replace(",", " ")
    parts = cleaned.split()
    if len(parts) != 3:
        raise ValueError(f"unrecognised date: {text!r}")
    a, b, c = parts
    try:
        year = int(c)
    except ValueError as exc:
        raise ValueError(f"unrecognised date: {text!r}") from exc
    if a.lower() in MONTHS:            # "August 30 2026"
        month, day = MONTHS[a.lower()], b
    elif b.lower() in MONTHS:          # "30 August 2026"
        day, month = a, MONTHS[b.lower()]
    else:
        raise ValueError(f"unrecognised month in date: {text!r}")
    try:
        return dt.date(year, month, int(day))
    except ValueError as exc:
        raise ValueError(f"invalid calendar date: {text!r}") from exc


def parse_us_date(text: str) -> dt.date:
    """Parse a US-format 'M/D/YYYY' date (as BagAnalytics writes its filter cell)."""
    m = re.fullmatch(r"\s*(\d{1,2})/(\d{1,2})/(\d{4})\s*", text)
    if not m:
        raise ValueError(f"unrecognised US date: {text!r}")
    month, day, year = (int(g) for g in m.groups())
    try:
        return dt.date(year, month, day)
    except ValueError as exc:
        raise ValueError(f"invalid calendar date: {text!r}") from exc


def parse_dmy_abbr(text: str) -> dt.date:
    """Parse '24-Aug-2026' (the form used inside phone-log 'E-mailed on ...' notes)."""
    m = re.fullmatch(r"\s*(\d{1,2})-([A-Za-z]{3})-(\d{4})\s*", text)
    if not m:
        raise ValueError(f"unrecognised date: {text!r}")
    day, mon, year = m.group(1), m.group(2).lower(), m.group(3)
    if mon not in MONTHS_ABBR:
        raise ValueError(f"unrecognised month: {text!r}")
    return dt.date(int(year), MONTHS_ABBR[mon], int(day))


def excel_serial_to_date(serial: float) -> dt.date:
    """Convert an Excel 1900-system serial number to a date."""
    return (_EXCEL_EPOCH + dt.timedelta(days=float(serial))).date()


def coerce_cell_date(value) -> dt.date | None:
    """Best-effort date out of an openpyxl cell value (datetime, serial, or string)."""
    if isinstance(value, dt.datetime):
        return value.date()
    if isinstance(value, dt.date):
        return value
    if isinstance(value, (int, float)):
        # Values below 1 are pure times; a real date serial is well above 20000.
        if value >= 20000:
            return excel_serial_to_date(value)
        return None
    if isinstance(value, str):
        for parser in (parse_english_date, parse_us_date, parse_dmy_abbr):
            try:
                return parser(value)
            except ValueError:
                continue
    return None
