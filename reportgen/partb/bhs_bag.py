"""BAGanalytics extractor - Adelaide (ADL). Spec section 5.3.

Numbers arrive as text ('59,302', '6.28%'); labels sit above their values. L1/L2 reject
figures are **always withheld** for now: early exports had an all-zero EDS table, and
later ones carry Level 1/2 counts whose mapping to the GlideView L1-flag / L2-confirmed-
reject definitions hasn't been confirmed. The raw EDS table goes to the console flags so
the mapping can be checked by hand.
"""
from __future__ import annotations

import datetime as dt
from pathlib import Path

from .. import dates
from ..discover import _BAGANALYTICS_DATE_RE, DiscoveryError, _load_workbook
from .grid import Grid, as_int, num
from .stats import BhsStats

ADL_EXCLUDE = {"Heartbeat Fault": "a heartbeat counter, not a fault"}


def _week(g: Grid, path: Path) -> tuple[dt.date, dt.date]:
    for row in g.rows[:40]:
        for value in row:
            if isinstance(value, str) and (m := _BAGANALYTICS_DATE_RE.search(value)):
                return dates.parse_us_date(m.group(1)), dates.parse_us_date(m.group(2))
    raise DiscoveryError(f"{path.name!r}: no 'Date: LAST 7 DAYS' cell on the Summary sheet")


def _faults(wb, path: Path) -> tuple[list[tuple[str, int]], list[tuple[str, int, str]]]:
    """Aggregate 'Faults List' (alarm name col D, count col E) by name across areas."""
    g = Grid.from_sheet(wb, "Faults List", path)
    r, c = g.find("Alarm")
    count_c = c + 1
    totals: dict[str, int] = {}
    for row in g.rows[r + 1:]:
        name = row[c] if c < len(row) else None
        if not isinstance(name, str) or not name.strip() or name.strip().lower().startswith("total"):
            continue
        n = as_int(row[count_c] if count_c < len(row) else None)
        if n:
            totals[name.strip()] = totals.get(name.strip(), 0) + n
    excluded = [(name, totals.pop(name), reason) for name, reason in ADL_EXCLUDE.items()
                if totals.get(name)]
    return sorted(totals.items(), key=lambda kv: (-kv[1], kv[0])), excluded


def _eds_table(wb, path: Path) -> list[tuple[str, str, dict[str, float]]]:
    """Rows of the 'EDS Statistics' decision table: (level, component, {metric: value})."""
    g = Grid.from_sheet(wb, "EDS Statistics", path)
    r, c = g.find("Component")
    headers = [str(h).strip() if h is not None else "" for h in g.rows[r]]
    out = []
    level = ""
    for row in g.rows[r + 1:]:
        if not any(v not in (None, "") for v in row):
            break
        if isinstance(row[c - 1] if c >= 1 else None, str) and row[c - 1].strip():
            level = row[c - 1].strip()
        comp = row[c]
        if not isinstance(comp, str) or not comp.strip():
            continue
        metrics = {h: num(v) or 0.0 for h, v in zip(headers[c + 1:], row[c + 1:])
                   if h and "%" not in h}
        out.append((level, comp.strip(), metrics))
    return out


def extract_adl(path: Path, week_start: dt.date, week_end: dt.date,
                recalculated: bool = False) -> BhsStats:
    path = Path(path)
    wb = _load_workbook(path, read_only=False)
    try:
        s = Grid.from_sheet(wb, "Summary", path)
        got = _week(s, path)
        if got != (week_start, week_end):
            raise DiscoveryError(
                f"{path.name!r} covers {got[0]:%d.%m.%Y}-{got[1]:%d.%m.%Y}, not the "
                f"reporting week {week_start:%d.%m.%Y}-{week_end:%d.%m.%Y}")
        in_gauge = as_int(s.number("Total Check-in Bags (In Gauge)", how="below"))
        oversize = as_int(s.number("Total Check-in Bags (Oversize)", how="below"))
        processed = as_int(s.number("Total Bags Processed", how="below"))
        with_recirc = as_int(s.number("Total Recircs plus Check-in Bags (In Gauge)", how="below"))
        makeup = as_int(s.number("Total Make-up Output", how="below"))
        lost_pct = s.number("Lost Track Percentage", how="below")
        tracking_rate = s.number("Tracking Rate", how="below")
        lost_total = as_int(s.number("Total Lost Bags", how="below"))

        zr, zc = s.find("Zone")
        count_c = next(i for i, v in enumerate(s.rows[zr]) if i > zc and isinstance(v, str)
                       and v.strip().lower() == "lost count")
        zones: list[tuple[str, int]] = []
        for row in s.rows[zr + 1:]:
            name = row[zc] if zc < len(row) else None
            n = as_int(row[count_c] if count_c < len(row) else None)
            if not isinstance(name, str) or n is None:
                break
            zones.append((name.strip(), n))

        faults, excluded = _faults(wb, path)
        eds = _eds_table(wb, path)
    finally:
        wb.close()

    flags: list[str] = []
    if in_gauge + oversize != processed:
        flags.append(f"in-gauge {in_gauge:,} + oversize {oversize:,} != Total Bags Processed "
                     f"{processed:,}")
    if sum(n for _, n in zones) != lost_total:
        flags.append(f"lost-bag zone table sums to {sum(n for _, n in zones):,} but Total Lost "
                     f"Bags is {lost_total:,} - the chart shows only the listed zones")
    if abs(lost_pct + tracking_rate - 100) > 0.02:
        flags.append(f"Lost Track Percentage {lost_pct}% + Tracking Rate {tracking_rate}% != 100%")

    all_zero = all(not any(m.values()) for _, _, m in eds)
    if all_zero:
        reason = ("The EDS Statistics export reads zero for every line although "
                  f"{in_gauge:,} bags went through the system - a data/export problem, not 0% "
                  "rejects")
    else:
        reason = ("the BAGanalytics EDS levels have not yet been reconciled to the L1 "
                  "auto-flag / L2 confirmed-reject definitions")
    flags.append("ADL L1/L2 withheld - raw EDS Statistics table for manual reconciliation:")
    for level, comp, metrics in eds:
        flags.append("    " + f"{level:<8} {comp:<6} " + "  ".join(
            f"{k} {v:,.0f}" for k, v in metrics.items()))

    return BhsStats(
        site="ADL", source="BAGanalytics", week_start=week_start, week_end=week_end,
        days_covered=[week_start + dt.timedelta(days=i) for i in range(7)],
        main_throughput=in_gauge, main_label="",
        main_note="Total Check-in Bags (In Gauge), excl. oversize",
        secondary_tags=(
            f"Total bags processed (in-gauge plus oversize) were {processed:,}; including "
            f"recirculations the system handled {with_recirc:,}, and make-up output was "
            f"{makeup:,}"),
        oversize=oversize, oversize_label="",
        reject_withheld=reason,
        lost_total=lost_total, lost_pct=lost_pct, lost_pct_basis="tracked scans",
        lost_note=f"BAGanalytics Lost Track Percentage (tracking rate {tracking_rate:.2f}%)",
        lost_granularity="Zone",
        lost_breakdown=zones,
        faults=faults, faults_total=sum(v for _, v in faults), excluded=excluded,
        trend=None,
        trend_note=("No daily trend chart: the authoritative daily throughput column in "
                    "ADL's BAGanalytics export has not been confirmed yet."),
        limitations=[
            "L1/L2 reject figures are withheld: " + reason + ".",
            "No day-by-day throughput trend is shown until the authoritative daily "
            "throughput column in the BAGanalytics export is confirmed.",
        ],
        recalculated=recalculated, flags=flags,
    )
