"""Link CarlSource tickets to phone-log calls and classify every Part A entry.

Entry kinds:
  - "ticket"     a CarlSource ticket with >=1 phone call in its region + time window
  - "direct"     a CarlSource ticket with no matching call ("Direct Email" / "Direct Contact")
  - "call_only"  a phone call that matched no ticket

A CarlSource ticket's `begin_date` is, in practice, the timestamp of the first call
(truncated to the minute), so matching is: same Part A region, call entered from a little
before the ticket begins to a couple of hours after. Calls are **not** consumed — a call
can support more than one ticket (duplicate tickets for one incident are common).
"""
from __future__ import annotations

import dataclasses
import datetime as dt

from .parse_carlsource import Ticket
from .parse_phonelogs import Call
from .regions import Regions

WINDOW_BEFORE = dt.timedelta(minutes=15)
WINDOW_AFTER = dt.timedelta(hours=2)
# second pass: a leftover call is grouped with the most recent same-SITE ticket that
# began within this long before it (captures same-incident follow-up calls hours later)
FOLLOWUP_BEFORE = dt.timedelta(hours=24)

_EMAIL_HINT = ("email", "e-mail", "e-mailed", "emailed")


@dataclasses.dataclass
class Entry:
    kind: str
    site_code: str | None
    region: str
    sort_key: tuple
    data: dict


@dataclasses.dataclass
class LinkResult:
    entries: list[Entry]
    flags: list[str]
    unmatched_tickets: list[str]
    unmatched_calls: list[str]


def _matches(call: Call, ticket: Ticket, regions: Regions) -> bool:
    if call.entered is None or ticket.begin_date is None:
        return False
    if not (ticket.begin_date - WINDOW_BEFORE <= call.entered <= ticket.begin_date + WINDOW_AFTER):
        return False
    cs, ts = regions.resolve(call.location_id), regions.resolve(ticket.client_customer)
    return bool(cs and ts and cs.region == ts.region) or (
        bool(cs and ts) and cs.code == ts.code
    )


def link(calls: list[Call], tickets: list[Ticket], regions: Regions) -> LinkResult:
    flags: list[str] = []
    call_by_number = {c.number: c for c in calls}
    calls_for_ticket: dict[str, list[str]] = {}
    tickets_for_call: dict[str, list[str]] = {c.number: [] for c in calls}

    for ticket in tickets:
        matched = [c for c in calls if _matches(c, ticket, regions)]
        if ticket.call_ref_hint:
            locked = [c for c in matched if c.number.startswith(ticket.call_ref_hint)]
            if len(locked) == 1:
                matched = locked
                flags.append(
                    f"ticket {ticket.internal_id} linked to call {locked[0].number} "
                    f"on 'call ref {ticket.call_ref_hint}' hint"
                )
        numbers = sorted(
            (c.number for c in matched),
            key=lambda n: call_by_number[n].entered or dt.datetime.max,
        )
        calls_for_ticket[ticket.internal_id] = numbers
        for n in numbers:
            tickets_for_call[n].append(ticket.internal_id)

    # second pass: attach leftover follow-up calls to the most recent same-site ticket
    ticket_by_id = {t.internal_id: t for t in tickets}
    for call in calls:
        if tickets_for_call[call.number] or call.entered is None:
            continue
        c_ref = regions.resolve(call.location_id)
        same_site = [
            t for t in tickets
            if t.begin_date is not None
            and (r := regions.resolve(t.client_customer)) and c_ref and r.code == c_ref.code
            and call.entered - FOLLOWUP_BEFORE <= t.begin_date <= call.entered + WINDOW_BEFORE
        ]
        if not same_site:
            continue
        parent = max(same_site, key=lambda t: t.begin_date)
        calls_for_ticket[parent.internal_id].append(call.number)
        calls_for_ticket[parent.internal_id].sort(
            key=lambda n: call_by_number[n].entered or dt.datetime.max
        )
        tickets_for_call[call.number].append(parent.internal_id)
        flags.append(
            f"call {call.number} grouped with ticket {parent.internal_id} as a same-site "
            f"follow-up ({call.entered:%d %b %H:%M})"
        )

    for number, tids in tickets_for_call.items():
        if len(tids) > 1:
            flags.append(f"call {number} matched {len(tids)} tickets: {', '.join(tids)}")

    entries: list[Entry] = []
    unmatched_tickets: list[str] = []

    for ticket in tickets:
        ref = regions.resolve(ticket.client_customer)
        region = ref.region if ref else "Other"
        code = ref.code if ref else None
        if not ref:
            flags.append(
                f"ticket {ticket.internal_id}: site {ticket.client_customer!r} not in "
                f"sites.yaml -> filed under Other"
            )
        numbers = calls_for_ticket[ticket.internal_id]
        common = dict(
            ticket_id=ticket.internal_id, title=ticket.title,
            begin_date=ticket.begin_date.isoformat() if ticket.begin_date else None,
            end_date=ticket.end_date.isoformat() if ticket.end_date else None,
            status=ticket.status, engineer=ticket.engineer,
            reported=ticket.reported, findings=ticket.findings, remedy=ticket.remedy,
            freeform=ticket.freeform,
        )
        first_ts = _first_ts(numbers, call_by_number, ticket.begin_date)
        if numbers:
            common.update(
                call_numbers=numbers,
                timestamps=[call_by_number[n].entered.isoformat()
                            for n in numbers if call_by_number[n].entered],
                confidence="high" if len(numbers) == 1 or ticket.call_ref_hint else "medium",
            )
            entries.append(Entry("ticket", code, region, ((code or "ZZZ"), first_ts), common))
        else:
            blob = f"{ticket.freeform}\n{ticket.reported}".lower()
            common["label"] = "Direct Email" if any(h in blob for h in _EMAIL_HINT) else "Direct Contact"
            common["timestamp"] = common["begin_date"]
            entries.append(Entry("direct", code, region, ((code or "ZZZ"), first_ts), common))
            unmatched_tickets.append(ticket.internal_id)

    unmatched_calls = [n for n, tids in tickets_for_call.items() if not tids]
    for number in unmatched_calls:
        call = call_by_number[number]
        ref = regions.resolve(call.location_id)
        region = ref.region if ref else "Other"
        code = ref.code if ref else None
        if not ref:
            flags.append(f"call {number}: location {call.location_id!r} not in sites.yaml -> Other")
        entries.append(Entry(
            "call_only", code, region,
            ((code or "ZZZ"), call.entered or dt.datetime.max),
            dict(call_number=number,
                 timestamp=call.entered.isoformat() if call.entered else None,
                 location_id=call.location_id, comments=call.comments, inquiry=call.inquiry,
                 transferred_to=call.transferred_to),
        ))

    return LinkResult(entries, flags, unmatched_tickets, unmatched_calls)


def _first_ts(numbers, call_by_number, fallback):
    stamps = [call_by_number[n].entered for n in numbers if call_by_number[n].entered]
    return min(stamps) if stamps else (fallback or dt.datetime.max)
