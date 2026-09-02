from __future__ import annotations

import datetime as dt
from pathlib import Path

import pytest

from reportgen import parse_carlsource, parse_phonelogs
from reportgen.bundle import build_part_a
from reportgen.link import link
from reportgen.regions import Regions
from tests.conftest import carlsource_xml, make_carlsource, make_phonelogs

REPO_ROOT = Path(__file__).resolve().parent.parent
KNOWLEDGE = REPO_ROOT / "knowledge"
DROPBOX = REPO_ROOT / "DropBox"

WEEK = {"start": "2026-08-24", "end": "2026-08-30",
        "subtitle": "Monday 24th August to Sunday 30th August 2026"}


@pytest.fixture
def regions() -> Regions:
    return Regions.from_yaml(KNOWLEDGE)


# ---------------------------------------------------------------- regions

def test_regions_resolve_by_location_code_and_contract(regions):
    for token in ("YLW_APAC", "YLW", "BHSYLW"):
        ref = regions.resolve(token)
        assert ref and ref.code == "YLW" and ref.region == "Canada"
    assert regions.resolve("ANZ_APAC").region == "New Zealand"
    assert regions.resolve("BOI_APAC").region == "United States"
    assert regions.resolve("AMX_APAC").region == "Parcels"


def test_regions_obscure_region_is_other(regions):
    # Nadi (Fiji) is in the lookup but Fiji has no Part A heading.
    assert regions.resolve("NAN_APAC").region == "Other"
    assert regions.region_for("NOPE_APAC") == "Other"
    assert regions.resolve("NOPE_APAC") is None


# ---------------------------------------------------------------- parse_carlsource

def test_parse_carlsource_fields_and_sections(tmp_path):
    path = tmp_path / "cs.xls"
    make_carlsource(path, [
        {"id": "176973", "title": "TC2 was down", "customer": "YLW_APAC",
         "contract": "BHSYLW", "begin": "2026-08-24T03:31:00.000",
         "end": "2026-08-24T15:07:27.752", "engineer": "Venthan Premadas",
         "status": "Closed", "reported": "Whole TC2 line was off.",
         "findings": "PLC I/O looked good.", "remedy": "Replaced the fuse.",
         "freeform": "Don reported TC2 went down."},
        {"id": "177547", "title": "BagSort failures", "customer": "YEG_APAC",
         "contract": "BHSYEG", "begin": "2026-08-28T16:31:00.000",
         "end": "2026-08-31T18:00:00.000", "engineer": "V P", "status": "In progress",
         "body": "Reported: BagSort comms failure\nRemedy: Escalated\n"
                 "28/08:\nWe received a call (Call ref 1519) about it.\n"},
    ])
    tickets = parse_carlsource.parse(path)
    assert [t.internal_id for t in tickets] == ["176973", "177547"]

    t0 = tickets[0]
    assert t0.client_customer == "YLW_APAC" and t0.contract == "BHSYLW"
    assert t0.begin_date == dt.datetime(2026, 8, 24, 3, 31)
    assert t0.status == "Closed" and t0.engineer == "Venthan Premadas"
    assert t0.reported == "Whole TC2 line was off."
    assert t0.findings == "PLC I/O looked good."
    assert t0.remedy == "Replaced the fuse."
    assert "Don reported" in t0.freeform

    assert tickets[1].call_ref_hint == "1519"
    assert tickets[1].findings == ""


def test_parse_carlsource_splits_run_together_labels(tmp_path):
    path = tmp_path / "cs.xls"
    make_carlsource(path, [{
        "id": "177122", "title": "Critical fault", "customer": "WLG_APAC",
        "begin": "2026-08-25T05:56:00.000",
        "body": "Reported: Relay tripped.\nFindings: Just a monitor relay.Remedy: Site "
                "electrician to reset.\n25/08:\nIssue with baggage system\n",
    }])
    t = parse_carlsource.parse(path)[0]
    assert t.findings == "Just a monitor relay."
    assert t.remedy == "Site electrician to reset."
    assert t.freeform.startswith("25/08")


# ---------------------------------------------------------------- parse_phonelogs

def test_parse_phonelogs(tmp_path):
    path = tmp_path / "pl.xlsx"
    make_phonelogs(
        path,
        [{"number": 15178, "entered": dt.datetime(2026, 8, 24, 1, 35), "location_id": "BOI_APAC"},
         {"number": 15179, "entered": dt.datetime(2026, 8, 24, 3, 31), "location_id": "YLW_APAC"}],
        stats_lines=["For hotline support:", "23 Calls received", "69.6% of calls received were after hours"],
    )
    data = parse_phonelogs.parse(path)
    assert [c.number for c in data.calls] == ["15178", "15179"]
    assert data.calls[0].location_id == "BOI_APAC"
    assert data.calls[1].entered == dt.datetime(2026, 8, 24, 3, 31)
    assert data.after_hours_pct == 69.6


# ---------------------------------------------------------------- link

def _calls(regions, *specs):
    from reportgen.parse_phonelogs import Call
    return [Call(number=n, entered=e, location_id=loc, transferred_to=None,
                 operator=None, comments="", inquiry="") for n, e, loc in specs]


def _ticket(**kw):
    from reportgen.parse_carlsource import Ticket
    base = dict(internal_id="1", title="t", client_customer="YLW_APAC", contract=None,
                begin_date=None, end_date=None, engineer=None, status=None, downtime=None,
                reported="", findings="", remedy="", freeform="", call_ref_hint=None)
    base.update(kw)
    return Ticket(**base)


def test_link_window_match_and_direct(regions):
    calls = _calls(regions, ("15179", dt.datetime(2026, 8, 24, 3, 31), "YLW_APAC"))
    t_match = _ticket(internal_id="176973", begin_date=dt.datetime(2026, 8, 24, 3, 31))
    t_direct = _ticket(internal_id="177611", client_customer="YLW_APAC",
                       begin_date=dt.datetime(2026, 8, 28, 8, 14),
                       freeform="Marcus emailed about it")
    res = link(calls, [t_match, t_direct], regions)
    kinds = {e.data.get("ticket_id"): e.kind for e in res.entries}
    assert kinds["176973"] == "ticket"
    assert kinds["177611"] == "direct"
    direct = next(e for e in res.entries if e.data.get("ticket_id") == "177611")
    assert direct.data["label"] == "Direct Email"


def test_link_call_ref_hint_locks(regions):
    calls = _calls(
        regions,
        ("15197", dt.datetime(2026, 8, 28, 16, 31), "YEG_APAC"),
        ("15192", dt.datetime(2026, 8, 27, 23, 6), "YEG_APAC"),  # earlier, different incident
    )
    t = _ticket(internal_id="177547", client_customer="YEG_APAC",
                begin_date=dt.datetime(2026, 8, 28, 16, 31), call_ref_hint="1519")
    res = link(calls, [t], regions)
    entry = next(e for e in res.entries if e.data.get("ticket_id") == "177547")
    assert entry.data["call_numbers"] == ["15197"]
    assert any("call ref 1519" in f for f in res.flags)


def test_link_cross_site_same_region(regions):
    # call from ANZ_APAC, ticket filed against AKL_APAC — same region (NZ) should match.
    calls = _calls(regions, ("15196", dt.datetime(2026, 8, 28, 13, 3), "ANZ_APAC"))
    t = _ticket(internal_id="177545", client_customer="AKL_APAC",
                begin_date=dt.datetime(2026, 8, 28, 13, 3))
    res = link(calls, [t], regions)
    entry = next(e for e in res.entries if e.data.get("ticket_id") == "177545")
    assert entry.kind == "ticket" and entry.data["call_numbers"] == ["15196"]


def test_link_merges_duplicate_tickets(regions):
    calls = _calls(regions, ("15191", dt.datetime(2026, 8, 27, 14, 42), "WLG_APAC"))
    t1 = _ticket(internal_id="177542", client_customer="WLG_APAC",
                 begin_date=dt.datetime(2026, 8, 27, 14, 42), reported="Relay fault A",
                 remedy="Adjusted voltage")
    t2 = _ticket(internal_id="177543", client_customer="WLG_APAC",
                 begin_date=dt.datetime(2026, 8, 27, 14, 22), reported="Relay fault A",
                 findings="No active alarms")
    res = link(calls, [t1, t2], regions)
    tickets = [e for e in res.entries if e.kind == "ticket"]
    assert len(tickets) == 1
    merged = tickets[0].data
    assert merged["ticket_id"] == "177543"  # earlier begin_date is primary
    assert merged["merged_ticket_ids"] == ["177542"]
    assert merged["reported"] == "Relay fault A"          # deduped
    assert "No active alarms" in merged["findings"] and "Adjusted voltage" in merged["remedy"]


def test_link_followup_call_grouped(regions):
    calls = _calls(
        regions,
        ("15194", dt.datetime(2026, 8, 28, 4, 36), "BDA_APAC"),
        ("15195", dt.datetime(2026, 8, 28, 11, 51), "BDA_APAC"),
    )
    t = _ticket(internal_id="177538", client_customer="BDA_APAC",
                begin_date=dt.datetime(2026, 8, 28, 4, 36))
    res = link(calls, [t], regions)
    entry = next(e for e in res.entries if e.data.get("ticket_id") == "177538")
    assert entry.data["call_numbers"] == ["15194", "15195"]
    assert not res.unmatched_calls


# ---------------------------------------------------------------- end to end

needs_samples = pytest.mark.skipif(
    not (DROPBOX / "PartA" / "Phonelogs - 24.08 to 30.08.2026.xlsx").is_file(),
    reason="committed sample drop not present",
)


@needs_samples
def test_sample_region_grid_matches_finished_example(regions):
    phone = parse_phonelogs.parse(DROPBOX / "PartA" / "Phonelogs - 24.08 to 30.08.2026.xlsx")
    tickets = parse_carlsource.parse(DROPBOX / "PartA" / "Carlsource - 24.08 to 30.08.2026.xls")
    res = link(phone.calls, tickets, regions)
    data = build_part_a(WEEK, phone, res, regions)["part_a"]

    assert data["totals"] == {"calls": 23, "direct_contacts": 2}
    assert data["region_counts"] == {
        "Africa": 0, "Australia": 1, "Bermuda": 2, "Canada": 10, "Jamaica": 0,
        "New Zealand": 8, "United States": 4, "Parcels": 0, "Other": 0,
    }
    assert len(tickets) == 18
    ids = set()
    for e in res.entries:
        if e.kind in ("ticket", "direct"):
            ids.add(e.data["ticket_id"])
            ids.update(e.data.get("merged_ticket_ids", []))
    assert ids == {t.internal_id for t in tickets}
    assert not res.unmatched_calls
    assert not any(r["name"] == "Other" for r in data["regions"])
    # 177542 / 177543 are one WLG incident (call 15191) - collapsed to one entry
    assert any(f.startswith("merged tickets 177543 + 177542") for f in res.flags)
