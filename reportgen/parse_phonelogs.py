"""Parse the PhonePlus weekly call-log workbook."""
from __future__ import annotations

import dataclasses
import datetime as dt
import re

from .discover import DiscoveryError, _load_workbook

_AFTER_HOURS_RE = re.compile(r"(\d+(?:\.\d+)?)\s*%[^.]*after hours", re.IGNORECASE)


@dataclasses.dataclass(frozen=True)
class Call:
    number: str
    entered: dt.datetime | None
    location_id: str | None
    transferred_to: str | None
    operator: str | None
    comments: str
    inquiry: str


@dataclasses.dataclass
class PhoneData:
    calls: list[Call]
    after_hours_pct: float | None
    stats_lines: list[str]
    site_tally: dict[str, int]


def _headers(row) -> dict[str, int]:
    return {str(v).strip().lower(): i for i, v in enumerate(row) if v is not None}


def _cell(row, i):
    return row[i] if i is not None and i < len(row) else None


def _text(value) -> str:
    """Cell text with Excel's `_x000D_` (escaped carriage return) stripped."""
    return str(value or "").replace("_x000D_", "").strip()


_TIME_FORMATS = ("%d/%m/%Y %H:%M:%S", "%d/%m/%Y %H:%M")


def _parse_entered(value, number: str) -> dt.datetime | None:
    """`Time Entered` is a real datetime in some rows and a `DD/MM/YYYY HH:MM:SS` string in
    others (same export). Blank -> None; anything present but unreadable is a hard error,
    since a silent None drops the call out of ticket matching."""
    if isinstance(value, dt.datetime):
        return value
    if isinstance(value, (int, float)):
        return dt.datetime(1899, 12, 30) + dt.timedelta(days=float(value))
    if isinstance(value, str) and value.strip():
        text = value.strip()
        for fmt in _TIME_FORMATS:
            try:
                return dt.datetime.strptime(text, fmt)
            except ValueError:
                continue
        raise DiscoveryError(
            f"Phone log call {number}: unreadable 'Time Entered' value {value!r} "
            f"(expected DD/MM/YYYY HH:MM:SS)"
        )
    return None


def parse(path) -> PhoneData:
    wb = _load_workbook(path)
    try:
        if "Phone Logs" not in wb.sheetnames:
            raise DiscoveryError(f"{path.name!r}: no 'Phone Logs' sheet")
        rows = wb["Phone Logs"].iter_rows(values_only=True)
        try:
            idx = _headers(next(rows))
        except StopIteration:
            raise DiscoveryError(f"{path.name!r}: 'Phone Logs' sheet is empty") from None

        c_num = idx.get("call number")
        c_time = idx.get("time entered")
        c_loc = idx.get("location of the caller")
        c_who = idx.get("who did you transfer the call to")
        c_op = idx.get("operator")
        c_com = idx.get("comments")
        c_inq = idx.get("inquiry")

        calls: list[Call] = []
        for row in rows:
            number = _cell(row, c_num)
            if number is None or str(number).strip() == "":
                continue
            call_number = str(number).strip().split(".")[0]
            entered = _parse_entered(_cell(row, c_time), call_number)
            loc = _cell(row, c_loc)
            calls.append(
                Call(
                    number=call_number,
                    entered=entered,
                    location_id=str(loc).strip() if loc else None,
                    transferred_to=(str(_cell(row, c_who)).strip() or None) if _cell(row, c_who) else None,
                    operator=(str(_cell(row, c_op)).strip() or None) if _cell(row, c_op) else None,
                    comments=_text(_cell(row, c_com)),
                    inquiry=_text(_cell(row, c_inq)),
                )
            )

        stats_lines: list[str] = []
        after_hours: float | None = None
        if "Phone Stats" in wb.sheetnames:
            for row in wb["Phone Stats"].iter_rows(values_only=True):
                for value in row:
                    if isinstance(value, str) and value.strip():
                        line = re.sub(r"^[\s•\-\*]+", "", value).strip()
                        stats_lines.append(line)
                        m = _AFTER_HOURS_RE.search(line)
                        if m and after_hours is None:
                            after_hours = float(m.group(1))

        site_tally: dict[str, int] = {}
        if "Data" in wb.sheetnames:
            for row in wb["Data"].iter_rows(values_only=True):
                if len(row) >= 2 and row[0] and isinstance(row[1], (int, float)):
                    site_tally[str(row[0]).strip()] = int(row[1])

        return PhoneData(
            calls=calls,
            after_hours_pct=after_hours,
            stats_lines=stats_lines,
            site_tally=site_tally,
        )
    finally:
        wb.close()
