"""Link CarlSource tickets to phone-log calls and classify every Part A entry.

Entry kinds:
  - "ticket"     a CarlSource ticket with >=1 phone call in its region + time window
  - "direct"     a CarlSource ticket with no matching call ("Direct Email" / "Direct Contact")
  - "call_only"  a phone call that matched no ticket

A ticket's `Call Reference` column names its calls, so those are linked directly by call
number (every number in the field, however separated). Only tickets with no resolvable
reference fall back to time-window matching: a CarlSource `begin_date` is, in practice, the
timestamp of the first call (truncated to the minute), so the fallback is: same Part A
region, call entered from a little before the ticket begins to a couple of hours after.
Calls are **not** consumed — a call can support more than one ticket (duplicate tickets for
one incident are common).
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
_DIRECT_LABELS = {"direct email": "Direct Email", "direct call": "Direct Contact"}


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
    cs, ts = regions.resolve(call.location_id), regions.resolve(ticket.client_code)
    return bool(cs and ts and cs.region == ts.region) or (
        bool(cs and ts) and cs.code == ts.code
    )


def link(calls: list[Call], tickets: list[Ticket], regions: Regions) -> LinkResult:
    flags: list[str] = []
    call_by_number = {c.number: c for c in calls}
    calls_for_ticket: dict[str, list[str]] = {}
    tickets_for_call: dict[str, list[str]] = {c.number: [] for c in calls}

    # calls a ticket names outright are not up for grabs by another ticket's window match
    claimed = {n for t in tickets for n in t.call_refs if n in call_by_number}

    for ticket in tickets:
        named = list(dict.fromkeys(n for n in ticket.call_refs if n in call_by_number))
        if named:
            _flag_named_call_issues(ticket, named, call_by_number, regions, flags)
            numbers = sorted(named, key=lambda n: call_by_number[n].entered or dt.datetime.max)
            calls_for_ticket[ticket.internal_id] = numbers
            for n in numbers:
                tickets_for_call[n].append(ticket.internal_id)
            continue

        matched = [c for c in calls if c.number not in claimed and _matches(c, ticket, regions)]
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
            and (r := regions.resolve(t.client_code)) and c_ref and r.code == c_ref.code
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

    entries: list[Entry] = []
    unmatched_tickets: list[str] = []

    for ticket in tickets:
        ref = regions.resolve(ticket.client_code)
        region = ref.region if ref else "Other"
        code = ref.code if ref else None
        if not ref:
            flags.append(
                f"ticket {ticket.internal_id}: site {ticket.client_code!r} not in "
                f"sites.yaml -> filed under Other"
            )
        numbers = calls_for_ticket[ticket.internal_id]
        common = dict(
            ticket_id=ticket.internal_id, title=ticket.title, client_name=ticket.client_name,
            begin_date=ticket.begin_date.isoformat() if ticket.begin_date else None,
            status=ticket.status,
            reported=ticket.reported, findings=ticket.findings, remedy=ticket.remedy,
            freeform=ticket.freeform, summary=ticket.summary,
        )
        first_ts = _first_ts(numbers, call_by_number, ticket.begin_date)
        if numbers:
            common.update(
                call_numbers=numbers,
                timestamps=[call_by_number[n].entered.isoformat()
                            for n in numbers if call_by_number[n].entered],
                # Forward-provisioned for the not-yet-built AI narrative step (skill
                # step 4) - no current consumer reads this field.
                confidence="high" if len(numbers) == 1 or ticket.call_ref_hint
                or ticket.call_refs else "medium",
            )
            entries.append(Entry("ticket", code, region, ((code or "ZZZ"), first_ts), common))
        else:
            ref_key = (ticket.call_reference or "").strip().lower()
            if ref_key in _DIRECT_LABELS:
                common["label"] = _DIRECT_LABELS[ref_key]
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

    entries = _merge_duplicate_tickets(entries, flags)
    return LinkResult(entries, flags, unmatched_tickets, unmatched_calls)


def _flag_named_call_issues(ticket: Ticket, named, call_by_number, regions, flags) -> None:
    missing = [n for n in ticket.call_refs if n not in call_by_number]
    if missing:
        flags.append(
            f"ticket {ticket.internal_id}: Call Reference lists {', '.join(missing)} "
            f"which is not in the phone log"
        )
    t_ref = regions.resolve(ticket.client_code)
    for n in named:
        c_ref = regions.resolve(call_by_number[n].location_id)
        if t_ref and c_ref and t_ref.region != c_ref.region:
            flags.append(
                f"ticket {ticket.internal_id} ({t_ref.code}) names call {n} "
                f"from a different region ({c_ref.code}) - linked anyway"
            )


def _merge_duplicate_tickets(entries: list[Entry], flags: list[str]) -> list[Entry]:
    """Two CarlSource tickets for the same site linked to the exact same call(s) are one
    incident logged twice - collapse them into a single entry."""
    buckets: dict[tuple, list[Entry]] = {}
    passthrough: list[Entry] = []
    for e in entries:
        calls = tuple(sorted(e.data.get("call_numbers", [])))
        if e.kind == "ticket" and calls:
            buckets.setdefault((e.site_code, calls), []).append(e)
        else:
            passthrough.append(e)

    for (_, calls), group in buckets.items():
        if len(group) == 1:
            passthrough.append(group[0])
            continue
        group.sort(key=lambda e: e.data.get("begin_date") or "")
        primary = group[0]
        ids = [g.data["ticket_id"] for g in group]
        merged = dict(primary.data)
        merged["merged_ticket_ids"] = ids[1:]
        for field in ("reported", "findings", "remedy", "freeform", "summary"):
            seen: list[str] = []
            for g in group:
                v = (g.data.get(field) or "").strip()
                if v and v not in seen:
                    seen.append(v)
            merged[field] = "\n\n".join(seen)
        passthrough.append(Entry(primary.kind, primary.site_code, primary.region,
                                 primary.sort_key, merged))
        flags.append(
            f"merged tickets {' + '.join(ids)} - same site, same call(s) {', '.join(calls)}"
        )
    return passthrough


def _first_ts(numbers, call_by_number, fallback):
    stamps = [call_by_number[n].entered for n in numbers if call_by_number[n].entered]
    return min(stamps) if stamps else (fallback or dt.datetime.max)
