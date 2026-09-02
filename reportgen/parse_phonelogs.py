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
            entered = _cell(row, c_time)
            if isinstance(entered, (int, float)):
                entered = dt.datetime(1899, 12, 30) + dt.timedelta(days=float(entered))
            elif not isinstance(entered, dt.datetime):
                entered = None
            loc = _cell(row, c_loc)
            calls.append(
                Call(
                    number=str(number).strip().split(".")[0],
                    entered=entered,
                    location_id=str(loc).strip() if loc else None,
                    transferred_to=(str(_cell(row, c_who)).strip() or None) if _cell(row, c_who) else None,
                    operator=(str(_cell(row, c_op)).strip() or None) if _cell(row, c_op) else None,
                    comments=str(_cell(row, c_com) or "").strip(),
                    inquiry=str(_cell(row, c_inq) or "").strip(),
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
