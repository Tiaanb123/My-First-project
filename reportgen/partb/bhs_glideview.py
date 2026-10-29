"""GlideView extractors - Queenstown (ZQN) and Christchurch (CHC). Spec sections 5.1-5.2.

Both read the workbook **after** `recalc.recalculated_copy()`; every SUMMARY value is
found by label. Each key total is re-derived from an underlying sheet and a disagreement
is a hard stop - surface it, never pick one.
"""
from __future__ import annotations

import datetime as dt
from pathlib import Path

from ..discover import DiscoveryError, _load_workbook
from .grid import Grid, as_int, num
from .stats import BhsStats


def _fmt(n: int | float) -> str:
    return f"{int(n):,}"


def _check(ok: bool, path: Path, what: str) -> None:
    if not ok:
        raise DiscoveryError(f"{Path(path).name!r}: cross-check failed - {what}. "
                             f"Check the export before building the report.")


# --------------------------------------------------------------------------------------
# shared: Fault Table + 5-minute throughput series
# --------------------------------------------------------------------------------------


def fault_table(wb, path: Path, exclude: dict[str, str]) -> tuple[
        list[tuple[str, int]], list[tuple[str, int, str]], list[str]]:
    """Sum every fault column over the device rows of the 'Fault Table' sheet.

    Header row = the first row with a blank column A and fault names across; device rows
    follow (column A = device). Rows and columns whose label contains 'total' are skipped.
    Returns (genuine faults descending, excluded counters, console flags)."""
    g = Grid.from_sheet(wb, "Fault Table", path)
    header_row = next(
        (i for i, row in enumerate(g.rows)
         if row and row[0] in (None, "") and sum(isinstance(v, str) for v in row[1:]) >= 3),
        None,
    )
    if header_row is None:
        raise DiscoveryError(f"{Path(path).name!r}: can't find the Fault Table header row")
    headers = g.rows[header_row]
    totals: dict[str, int] = {}
    for row in g.rows[header_row + 1:]:
        device = row[0] if row else None
        if device in (None, "") or "total" in str(device).lower():
            continue
        for name, value in zip(headers[1:], row[1:]):
            if not isinstance(name, str) or not name.strip() or "total" in name.lower():
                continue
            n = num(value)
            if n:
                key = name.strip()
                totals[key] = totals.get(key, 0) + int(round(n))

    excluded = [(name, totals.pop(name), reason) for name, reason in exclude.items()
                if totals.get(name)]
    faults = sorted(((k, v) for k, v in totals.items() if v), key=lambda kv: (-kv[1], kv[0]))
    flags: list[str] = []
    grand = sum(v for _, v in faults)
    if faults and grand and faults[0][1] > 0.5 * grand:
        flags.append(
            f"fault category {faults[0][0]!r} is {faults[0][1] / grand:.0%} of all faults - "
            f"if it's a heartbeat/status counter, add it to the site's exclusions"
        )
    return faults, excluded, flags


def throughput_series(wb, path: Path, tag: str, week_start: dt.date, week_end: dt.date
                      ) -> tuple[dict[dt.date, int], list[str]]:
    """Daily totals of `tag` from the 5-minute 'System Throughput Rates' sheet. Each row is
    stamped with its interval *end*, so 00:00 belongs to the previous day."""
    g = Grid.from_sheet(wb, "System Throughput Rates", path)
    r, _ = g.find("LocalTime")
    header = g.rows[r]
    if tag not in header:
        raise DiscoveryError(f"{Path(path).name!r}: no {tag!r} column on System Throughput Rates")
    i_time, i_tag = header.index("LocalTime"), header.index(tag)
    daily: dict[dt.date, int] = {}
    outside = 0
    for row in g.rows[r + 1:]:
        when = row[i_time] if i_time < len(row) else None
        if not isinstance(when, dt.datetime):
            continue
        day = (when - dt.timedelta(minutes=1)).date()
        if not week_start <= day <= week_end:
            outside += 1
            continue
        daily[day] = daily.get(day, 0) + (as_int(row[i_tag]) or 0)
    flags = [f"{outside} System Throughput Rates rows fall outside the reporting week (ignored)"
             ] if outside else []
    return daily, flags


def _days_note(days: list[dt.date], week_start: dt.date, week_end: dt.date) -> list[str]:
    if len(days) == 7:
        return []
    have = ", ".join(f"{d:%a %d %b}" for d in days) or "none"
    return [f"!! export covers only {len(days)} day(s) of the week {week_start:%d.%m}-"
            f"{week_end:%d.%m.%Y} ({have}) - every figure is labelled as such; ask for a "
            f"full-week re-export"]


# --------------------------------------------------------------------------------------
# ZQN
# --------------------------------------------------------------------------------------

ZQN_CHECKPOINTS = ("T109PC", "T110PC", "T111PC", "T112PC", "T1DV1PD", "T1DV1PU", "T110APC")
ZQN_EXCLUDE = {"Dieback Status": "a heartbeat/status counter, not a fault"}


def extract_zqn(path: Path, week_start: dt.date, week_end: dt.date,
                recalculated: bool = False) -> BhsStats:
    path = Path(path)
    wb = _load_workbook(path, read_only=False)
    try:
        s = Grid.from_sheet(wb, "SUMMARY", path)
        t101, t114, os104 = (as_int(s.number(k)) for k in ("T101", "T114", "OS104"))
        l1_clear, l1_reject, l1_noscan, l1_timeout = (as_int(s.number(k)) for k in (
            "L1 Clear Count", "L1 Reject Bag Count", "L1 No Scan Count", "L1 Timeout Count"))
        l2_reject, l2_timeout = (as_int(s.number(k)) for k in (
            "L2 Reject Bag Count", "L2 Timeout Count"))

        hdr_r, hdr_c = s.find("Tracked Bag Count")
        checkpoints: list[tuple[str, int, int]] = []
        total_r = None
        for r in range(hdr_r + 1, len(s.rows)):
            label = s.cell(r, hdr_c - 1)
            if isinstance(label, str) and label.strip().upper() == "TOTAL":
                total_r = r
                break
            if label in (None, ""):
                break
            checkpoints.append((str(label).strip(), as_int(s.cell(r, hdr_c)) or 0,
                                as_int(s.cell(r, hdr_c + 1)) or 0))
        if total_r is None:
            raise DiscoveryError(f"{path.name!r}: no TOTAL row under the SUMMARY tracking block")
        lost_total = as_int(s.right_of("TOTAL", start=total_r, end=total_r + 1))
        _check(sum(c[2] for c in checkpoints) == lost_total, path,
               f"checkpoint lost bags sum to {sum(c[2] for c in checkpoints)} but SUMMARY "
               f"TOTAL says {lost_total}")
        tracked = sum(c[1] for c in checkpoints)
        missing = [c for c in ZQN_CHECKPOINTS if c not in {name for name, _, _ in checkpoints}]

        faults, excluded, flags = fault_table(wb, path, ZQN_EXCLUDE)
        daily, more = throughput_series(wb, path, "T101", week_start, week_end)
        flags += more
    finally:
        wb.close()

    days = sorted(d for d in daily)
    _check(sum(daily.values()) == t101, path,
           f"System Throughput Rates T101 sums to {sum(daily.values()):,} but SUMMARY T101 "
           f"is {t101:,}")
    if missing:
        flags.append(f"ZQN checkpoints missing from SUMMARY: {', '.join(missing)}")
    flags += _days_note(days, week_start, week_end)

    trend = [(d, daily[d]) for d in days] if len(days) >= 2 else None
    return BhsStats(
        site="ZQN", source="GlideView", week_start=week_start, week_end=week_end,
        days_covered=days,
        main_throughput=t101, main_label="T101",
        main_note=f"Excludes oversize; T114 secondary tag read {_fmt(t114)}",
        secondary_tags=f"The T114 secondary tag read {_fmt(t114)} bags",
        oversize=os104, oversize_label="OS104",
        l1_flagged=l1_reject, l1_denominator=l1_clear + l1_reject + l1_noscan + l1_timeout,
        l2_reject=l2_reject, l2_timeouts=l2_timeout,
        lost_total=lost_total,
        lost_pct=100 * lost_total / tracked if tracked else 0.0,
        lost_pct_basis="tracked scans",
        lost_note="Flat sum across ZQN's checkpoint list",
        lost_granularity="Checkpoint",
        lost_breakdown=[(name, lost) for name, _, lost in checkpoints],
        faults=faults, faults_total=sum(v for _, v in faults), excluded=excluded,
        trend=trend,
        trend_note="" if trend else (
            "No daily trend chart: the export covers a single day, so there is no "
            "day-by-day series to plot."),
        recalculated=recalculated, flags=flags,
    )


# --------------------------------------------------------------------------------------
# CHC
# --------------------------------------------------------------------------------------

CHC_EDS_LINES = ("X1X0100 EDS", "X1X0200 EDS", "X1X0300 EDS", "X1X0400 EDS")
CHC_ZONES = ("TC1 & TC2 prior to ME", "Domestic CBS", "International CBS", "Exiting ME",
             "Between ME and L4T01", "S1 Sort Line", "S2 Sort Line", "Level 4 Line")
# PM testing drives these counts - quoted with their share so the chart isn't read as
# all unplanned faults (spec 5.2).
CHC_PM_FAULTS = ("Emergency Stop Fault", "Emergency Stop Relay Fault",
                 "Motor Isolator Fault", "Manual Mode Active")


def _raw_screening(wb, path: Path) -> tuple[int, int]:
    """(sum of all L1 decisions, sum of L2 rejects) from 'Security Screening Statistics'."""
    g = Grid.from_sheet(wb, "Security Screening Statistics", path)
    r1, c1 = g.find("Level 1 Clear Count")
    l1_cols = [c1 + i for i in range(4)]  # clear, reject, timeout, no-scan
    l1 = 0
    r = r1 + 1
    while r < len(g.rows) and isinstance(g.cell(r, 0), str) and g.cell(r, 0).startswith("X1X0"):
        l1 += sum(as_int(g.cell(r, c)) or 0 for c in l1_cols)
        r += 1
    r2, c2 = g.find("Level 2 Reject Count")
    l2 = 0
    r = r2 + 1
    while r < len(g.rows) and isinstance(g.cell(r, 0), str) and g.cell(r, 0).startswith("X1X0"):
        l2 += as_int(g.cell(r, c2)) or 0
        r += 1
    return l1, l2


def extract_chc(path: Path, week_start: dt.date, week_end: dt.date,
                recalculated: bool = False) -> BhsStats:
    path = Path(path)
    wb = _load_workbook(path, read_only=False)
    try:
        s = Grid.from_sheet(wb, "SUMMARY", path)
        lines = {}
        for line in CHC_EDS_LINES:
            r, _ = s.find(line)
            lines[line] = {k: as_int(s.number(k, start=r + 1, end=r + 8)) for k in (
                "Bags Screened", "Level 1 Clear", "No Scans", "Level 2 Reject",
                "Level 2 Timeout")}
        screened = sum(v["Bags Screened"] for v in lines.values())
        screening_total = as_int(s.number("SCREENING TOTAL"))
        sort_system = as_int(s.number("SORT SYSTEM"))
        oog = as_int(s.number("OOG Total"))

        tr, _ = s.find("TRACKING STATS")
        zones = [(z, as_int(s.number(z, start=tr, end=tr + 12))) for z in CHC_ZONES]
        tot_r, _ = s.find("TOTAL", start=tr)
        lost_total = as_int(s.number("TOTAL", start=tot_r, end=tot_r + 1))
        lost_pct = 100 * s.number("Percentage*", start=tot_r, end=tot_r + 2)

        raw_l1, raw_l2 = _raw_screening(wb, path)
        faults, excluded, flags = fault_table(wb, path, {})
        has_series = any(n.strip().lower() == "system throughput rates" for n in wb.sheetnames)
    finally:
        wb.close()

    l2_reject = sum(v["Level 2 Reject"] for v in lines.values())
    _check(screened == screening_total, path,
           f"the four EDS lines' Bags Screened sum to {screened:,} but SCREENING TOTAL is "
           f"{screening_total:,}")
    _check(raw_l1 == screened, path,
           f"Security Screening Statistics L1 decisions sum to {raw_l1:,} but SUMMARY "
           f"Bags Screened is {screened:,}")
    _check(raw_l2 == l2_reject, path,
           f"Security Screening Statistics L2 rejects sum to {raw_l2:,} but SUMMARY says "
           f"{l2_reject:,}")
    _check(sum(n for _, n in zones) == lost_total, path,
           f"the 8 tracking zones sum to {sum(n for _, n in zones):,} but TRACKING STATS "
           f"TOTAL is {lost_total:,}")
    _check(abs(100 * lost_total / sort_system - lost_pct) < 0.01 if sort_system else False,
           path, f"tracking Percentage* {lost_pct:.2f}% doesn't equal TOTAL / SORT SYSTEM")

    faults_total = sum(v for _, v in faults)
    by_name = dict(faults)
    pm = sum(by_name.get(n, 0) for n in CHC_PM_FAULTS)
    if has_series:
        flags.append("CHC export now has a System Throughput Rates sheet - the trend chart "
                     "could be enabled for CHC")

    return BhsStats(
        site="CHC", source="GlideView", week_start=week_start, week_end=week_end,
        days_covered=[week_start + dt.timedelta(days=i) for i in range(7)],
        main_throughput=screened, main_label="X1X0100-X1X0400 EDS",
        main_note="Sum of 4 parallel EDS lines, excl. oversize",
        secondary=[(
            "Check-in-side total (cross-check)", f"{_fmt(sort_system)} bags (C1+C2+C3+Tranship)",
            "Secondary figure - will not match EDS total exactly due to re-screening, expected",
        )],
        secondary_tags=(
            f"The check-in-side total (C1+C2+C3+Tranship) was {_fmt(sort_system)} bags; "
            f"it does not match the EDS aggregate exactly because some bags are re-screened, "
            f"which is expected"),
        oversize=oog, oversize_label="OOG Total",
        l1_flagged=screened - sum(v["Level 1 Clear"] for v in lines.values())
        - sum(v["No Scans"] for v in lines.values()),
        l1_denominator=screened,
        l2_reject=l2_reject, l2_timeouts=sum(v["Level 2 Timeout"] for v in lines.values()),
        lost_total=lost_total, lost_pct=lost_pct,
        lost_pct_basis="bags entering sort system",
        lost_note="8-zone Tracking Stats SUMMARY aggregate, not a flat checkpoint sum",
        lost_granularity="Zone",
        lost_breakdown=zones,
        faults=faults, faults_total=faults_total, excluded=excluded,
        pm_testing=(pm, faults_total), pm_categories=list(CHC_PM_FAULTS),
        trend=None,
        trend_note=("No daily trend chart: the CHC GlideView export has no 5-minute "
                    "time-series sheet to build one from."),
        recalculated=recalculated, flags=flags,
    )
