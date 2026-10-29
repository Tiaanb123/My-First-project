"""Rule-based wording for the BHS Performance Review (deterministic; NZ English).

Every number comes from `BhsStats`; the fixed sentences are module-level constants so the
wording can be edited without touching the logic. `entries` are the site's Part A ticket
entries from work/bundle.json.
"""
from __future__ import annotations

from . import metrics as m
from .stats import BhsStats

L1_L2_EXPLAINER = (
    "L1 is an automated first-pass flag, NOT a final reject: the L1 machine sends any bag "
    "it cannot clear on its own for a second look. L2 (thorough second screening) is the "
    "figure that reflects actual operational reject volume."
)
TIMEOUT_NOTE = (
    "A further {t} L2 timeouts were recorded; a timeout is not a confirmed reject, so these "
    "are reported separately."
)
COMPARABILITY_CAVEAT = (
    "Lost-bag percentages are not directly comparable between sites: each site's export "
    "aggregates tracking differently (for example, a flat checkpoint list at Queenstown "
    "versus system-zone groupings at Christchurch)."
)
RECALC_NOTE = (
    "The export's SUMMARY sheet carried stale values and was recalculated before any "
    "figure was read; key totals were cross-checked against the underlying sheets."
)
TREND_LATER = (
    "Reintroduce a week-over-week trend and a 'most recent full week' snapshot once "
    "several consecutive weeks of exports are on file for {site}."
)
FULL_WEEK_REQUEST = (
    "Request a full Monday-Sunday {source} export for {site} - this week's file covers "
    "{days} only, so these figures are not a weekly view."
)


def _cap(text: str) -> str:
    return text[:1].upper() + text[1:]


def _uncap(text: str) -> str:
    """Lower-case a leading capital for mid-sentence use, but leave acronyms ('BAGanalytics')."""
    return text[:1].lower() + text[1:] if text[1:2].islower() else text


def _tickets(entries: list[dict]) -> list[str]:
    ids: list[str] = []
    for e in entries:
        if e.get("kind") in ("ticket", "direct"):
            ids += [e["ticket_id"], *e.get("merged_ticket_ids", [])]
    return list(dict.fromkeys(ids))


def _top_lost(st: BhsStats) -> tuple[str, int] | None:
    data = [kv for kv in st.lost_breakdown if kv[1]]
    return max(data, key=lambda kv: kv[1]) if data else None


def _top_faults(st: BhsStats, k: int = 3) -> str:
    return ", ".join(f"{name} ({m.n(v)})" for name, v in st.faults[:k])


def scope_note(st: BhsStats) -> str | None:
    """Page-1 data-scope note when the export doesn't cover the whole reporting week."""
    if st.full_week:
        return None
    return (
        f"Data scope: the {st.source} export supplied for {st.site} covers "
        f"{m.days_label(st.days_covered)} only, not the full reporting week "
        f"({m.week_range(st.week_start, st.week_end)}). Every BHS figure in this report is "
        f"labelled accordingly. Please re-upload a full-week export so a weekly view can "
        f"be given."
    )


def exec_summary(st: BhsStats, entries: list[dict]) -> list[str]:
    """Exactly five bullets (spec section 8)."""
    sc = m.scope(st)
    label = f" ({st.main_label})" if st.main_label else ""
    out = [f"Main-line throughput {sc} was {m.n(st.main_throughput)} bags{label}, with a "
           f"further {m.n(st.oversize)} bags on the separate oversize line."]
    if st.reject_withheld:
        out.append("L1/L2 reject figures are withheld this week as provisional: "
                   + st.reject_withheld + ".")
        out.append(f"The largest fault categories {sc} were {_top_faults(st)}.")
    else:
        out.append(f"{m.n(st.l2_reject)} bags were confirmed as rejects at L2 - "
                   f"{m.pct(m.l2_pct(st))} of main-line throughput. This is the operational "
                   f"reject figure.")
        out.append(f"{m.n(st.l1_flagged)} bags ({m.pct(m.l1_pct(st), 1)} of those screened) "
                   f"were flagged at L1 and sent for second-level screening; an L1 flag is an "
                   f"automated first pass, not a final reject.")
    top = _top_lost(st)
    lost = (f"{m.n(st.lost_total)} lost-bag events were recorded ({m.pct(st.lost_pct)} of "
            f"{st.lost_pct_basis})")
    if top:
        lost += (f", the largest share at {top[0]} ({m.n(top[1])}, "
                 f"{m.share(top[1], st.lost_total)} of the total)")
    out.append(lost + ".")

    if not st.full_week:
        out.append(f"Data scope: the export covers {m.days_label(st.days_covered)} only, so "
                   f"none of these figures is a full-week view.")
    elif st.reject_withheld:
        out.append(st.trend_note or "No day-by-day trend is shown this week.")
    elif st.pm_testing and st.pm_testing[0]:
        pm, total = st.pm_testing
        out.append(f"{m.share(pm, total)} of recorded faults ({m.n(pm)} of {m.n(total)}) "
                   f"fall in categories driven partly by planned PM testing, so the fault "
                   f"chart overstates unplanned faults.")
    elif st.recalculated:
        out.append(RECALC_NOTE)
    else:
        out.append(f"The most frequent fault categories were {_top_faults(st)}.")
    return out


def throughput(st: BhsStats) -> str:
    label = f" ({st.main_label})" if st.main_label else ""
    text = (f"{_cap(m.scope(st))}, the main line handled {m.n(st.main_throughput)} "
            f"bags{label}, excluding oversize. ")
    if st.secondary_tags:
        text += st.secondary_tags + ". "
    olabel = f" ({st.oversize_label})" if st.oversize_label else ""
    text += f"The oversize line{olabel} handled {m.n(st.oversize)} bags separately."
    if st.trend:
        busiest = max(st.trend, key=lambda dv: dv[1])
        quietest = min(st.trend, key=lambda dv: dv[1])
        text += (f" Daily totals are charted below: the busiest day was {busiest[0]:%A} "
                 f"({m.n(busiest[1])} bags) and the quietest {quietest[0]:%A} "
                 f"({m.n(quietest[1])} bags).")
    elif st.trend_note:
        text += " " + st.trend_note
    return text


def rejects(st: BhsStats) -> list[str]:
    if st.reject_withheld:
        return [L1_L2_EXPLAINER,
                f"For {st.site} these figures are withheld this week: {st.reject_withheld}. "
                f"No reject rate is published rather than risk a misleading one."]
    paras = [L1_L2_EXPLAINER]
    text = (f"{_cap(m.scope(st))}, {m.n(st.l1_flagged)} of {m.n(st.l1_denominator)} "
            f"{st.l1_basis} ({m.pct(m.l1_pct(st), 1)}) were flagged for re-screening. "
            f"{m.n(st.l2_reject)} bags were confirmed as rejects at L2 - "
            f"{m.pct(m.l2_pct(st))} of main-line throughput.")
    if st.l2_timeouts:
        text += " " + TIMEOUT_NOTE.format(t=m.n(st.l2_timeouts))
    paras.append(text)
    return paras


def faults(st: BhsStats) -> str:
    shown = min(len(st.faults), 10)
    text = (f"The chart shows the top {shown} of {len(st.faults)} fault categories recorded "
            f"{m.scope(st)}, summed across all devices ({m.n(st.faults_total)} fault events "
            f"in total).")
    for name, count, reason in st.excluded:
        text += (f" {name} ({m.n(count)} hits) is excluded from the chart and totals as "
                 f"{reason}.")
    if st.pm_testing and st.pm_testing[0]:
        pm, total = st.pm_testing
        text += (f" Note: {', '.join(st.pm_categories)} are driven partly by planned PM "
                 f"testing - together {m.n(pm)} of the {m.n(total)} faults recorded "
                 f"({m.share(pm, total)}) - so the chart should not be read as all "
                 f"unplanned faults.")
    return text


def lost(st: BhsStats) -> str:
    text = (f"{m.n(st.lost_total)} lost-bag events were recorded {m.scope(st)} - "
            f"{m.pct(st.lost_pct)} of {st.lost_pct_basis} ({_uncap(st.lost_note)}).")
    top = _top_lost(st)
    if top:
        text += (f" The chart breaks this down by {st.lost_granularity.lower()}; the largest "
                 f"was {top[0]} with {m.n(top[1])} ({m.share(top[1], st.lost_total)} of "
                 f"the total).")
    return text + " " + COMPARABILITY_CAVEAT


def observations(st: BhsStats, entries: list[dict]) -> list[str]:
    out = []
    top = _top_lost(st)
    if top:
        out.append(f"{top[0]} accounted for the largest share of lost-bag events "
                   f"({m.share(top[1], st.lost_total)}, {m.n(top[1])} of "
                   f"{m.n(st.lost_total)}).")
    if st.faults:
        out.append(f"The largest genuine fault categories were {_top_faults(st)}.")
    tickets = _tickets(entries)
    if tickets:
        out.append(f"{len(tickets)} support ticket(s) were logged for {st.site} this week "
                   f"(WO {', '.join(tickets)}) - see Support Tickets above.")
    else:
        out.append(f"No support tickets were logged for {st.site} this week.")
    if not st.full_week:
        out.append(f"Open data-quality item: the export covers {m.days_label(st.days_covered)} "
                   f"only.")
    elif st.reject_withheld:
        out.append("Open data-quality item: L1/L2 reject figures remain provisional until "
                   "the EDS levels are reconciled.")
    elif st.recalculated:
        out.append("Open data-quality item: the SUMMARY sheet needed recalculating before "
                   "it could be read.")
    return out[:4]


def next_steps(st: BhsStats) -> list[str]:
    out = []
    if not st.full_week:
        out.append(FULL_WEEK_REQUEST.format(source=st.source, site=st.site,
                                            days=m.days_label(st.days_covered)))
    if st.reject_withheld:
        out.append(f"Confirm with the site how the {st.source} EDS levels map to the L1 "
                   f"flag / L2 confirmed-reject definitions so reject figures can be "
                   f"published.")
    top = _top_lost(st)
    if top and len(out) < 2:
        out.append(f"Review tracking performance at {top[0]}, which accounted for "
                   f"{m.share(top[1], st.lost_total)} of lost-bag events {m.scope(st)}.")
    out.append(TREND_LATER.format(site=st.site))
    return out[:3]
