"""Number formatting, percentages and period labels for Part B - one place, so the table,
the prose and the chart titles always agree."""
from __future__ import annotations

import datetime as dt

from .stats import BhsStats


def n(value: int | float) -> str:
    return f"{int(round(value)):,}"


def pct(value: float, places: int = 2) -> str:
    return f"{value:.{places}f}%"


def share(part: int, whole: int, places: int = 1) -> str:
    return pct(100 * part / whole, places) if whole else "n/a"


def l1_pct(st: BhsStats) -> float | None:
    if st.l1_flagged is None or not st.l1_denominator:
        return None
    return 100 * st.l1_flagged / st.l1_denominator


def l2_pct(st: BhsStats) -> float | None:
    if st.l2_reject is None or not st.main_throughput:
        return None
    return 100 * st.l2_reject / st.main_throughput


def _day(d: dt.date) -> str:
    return f"{d:%A} {d.day} {d:%B %Y}"


def week_range(start: dt.date, end: dt.date) -> str:
    """'Monday 19 - Sunday 25 October 2026' (month repeated only when it changes)."""
    if (start.year, start.month) == (end.year, end.month):
        return f"{start:%A} {start.day} - {end:%A} {end.day} {end:%B %Y}"
    if start.year == end.year:
        return f"{start:%A} {start.day} {start:%B} - {end:%A} {end.day} {end:%B %Y}"
    return f"{_day(start)} - {_day(end)}"


def days_label(days: list[dt.date]) -> str:
    if len(days) == 1:
        return _day(days[0])
    if days and (days[-1] - days[0]).days == len(days) - 1:
        return week_range(days[0], days[-1])
    return ", ".join(f"{d:%a} {d.day} {d:%b}" for d in days) + f" {days[-1]:%Y}"


def short_range(st: BhsStats) -> str:
    """For chart titles: '19-25 Oct 2026' / '11 Oct 2026'."""
    days = st.days_covered or [st.week_start, st.week_end]
    a, b = days[0], days[-1]
    if a == b:
        return f"{a.day} {a:%b %Y}"
    if a.month == b.month:
        return f"{a.day}-{b.day} {b:%b %Y}"
    return f"{a.day} {a:%b} - {b.day} {b:%b %Y}"


def scope(st: BhsStats) -> str:
    """Prose qualifier: 'over the week' / 'on Sunday 11 October 2026 only'."""
    if st.full_week:
        return "over the week"
    return f"on {days_label(st.days_covered)} only"


def source_line(st: BhsStats) -> str:
    if st.full_week:
        return f"{week_range(st.week_start, st.week_end)}, sourced from {st.source}."
    return (f"{days_label(st.days_covered)} only (see data scope note above) - "
            f"sourced from {st.source}.")


def review_period(st: BhsStats) -> str:
    if st.full_week:
        return f"{week_range(st.week_start, st.week_end)} (reporting week)"
    kind = "single-day" if st.single_day else "partial-week"
    return (f"{days_label(st.days_covered)} only ({kind} {st.source} export - see data "
            f"scope note on page 1)")


def _with_label(count: int, label: str) -> str:
    return f"{n(count)} bags" + (f" ({label})" if label else "")


def table_rows(st: BhsStats) -> list[tuple[str, str, str]]:
    """The fixed BHS Statistics rows (spec section 7): Metric | Value | Note."""
    rows = [("Total throughput (main line)", _with_label(st.main_throughput, st.main_label),
             st.main_note)]
    rows += st.secondary
    rows.append(("Oversize line volume", _with_label(st.oversize, st.oversize_label),
                 "Separate system"))
    if st.reject_withheld:
        rows.append(("L1 / L2 reject figures", "Provisional - withheld",
                     "See note below the table"))
    else:
        rows.append(("L1 flagged for re-screen",
                     f"{n(st.l1_flagged)} bags ({pct(l1_pct(st), 1)})",
                     "Sent to L2, not a final reject"))
        note = "The operational reject figure"
        if st.l2_timeouts:
            note += (f"; {n(st.l2_timeouts)} further L2 timeouts not counted as confirmed "
                     f"rejects")
        rows.append(("L2 confirmed reject (final)",
                     f"{n(st.l2_reject)} bags ({pct(l2_pct(st))} of throughput)", note))
    rows.append(("Lost-bag events", f"{n(st.lost_total)} ({pct(st.lost_pct)} of "
                 f"{st.lost_pct_basis})", st.lost_note))
    return rows
