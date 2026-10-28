"""Regressions found comparing the 7-13 Sep 2026 Part A output with its source workbooks."""
from __future__ import annotations

import datetime as dt

import pytest

from reportgen import parse_carlsource, parse_phonelogs
from reportgen.discover import DiscoveryError
from reportgen.link import link
from tests.conftest import make_carlsource, make_phonelogs
from tests.test_parse_parta import _calls, _ticket, regions  # noqa: F401  (fixture)


def _one_ticket(tmp_path, **fields):
    path = tmp_path / "cs.xlsx"
    make_carlsource(path, [{"title": "t", "begin": "07/09/2026 12:53:00", **fields}])
    return parse_carlsource.parse(path)[0]


# ---------------------------------------------------------------- phone log

def test_phone_times_accept_string_and_datetime(tmp_path):
    path = tmp_path / "pl.xlsx"
    make_phonelogs(path, [
        {"number": 15222, "entered": dt.datetime(2026, 9, 9, 13, 35, 6), "location_id": "YLW_APAC"},
        {"number": "15223", "entered": "11/09/2026 08:43:46", "location_id": "ADL_APAC",
         "comments": "E-mailed on 11-Sep-2026 at 08:45 by Op_x000D_\nCalling in"},
    ])
    calls = parse_phonelogs.parse(path).calls
    assert [c.entered for c in calls] == [
        dt.datetime(2026, 9, 9, 13, 35, 6), dt.datetime(2026, 9, 11, 8, 43, 46)]
    assert "_x000D_" not in calls[1].comments


def test_unreadable_phone_time_is_an_error(tmp_path):
    path = tmp_path / "pl.xlsx"
    make_phonelogs(path, [{"number": 15300, "entered": "yesterday-ish", "comments": "x"}])
    with pytest.raises(DiscoveryError, match="15300"):
        parse_phonelogs.parse(path)


# ---------------------------------------------------------------- call references / linking

def test_call_refs_extracted_whatever_the_separator(tmp_path):
    path = tmp_path / "cs.xlsx"
    make_carlsource(path, [
        {"id": "1", "title": "a", "client_code": "ZQN_APAC", "begin": "12/09/2026 11:55:00",
         "call_reference": "15230. 15231", "reported": "x"},
        {"id": "2", "title": "b", "client_code": "SFO_APAC", "begin": "13/09/2026 01:25:00",
         "call_reference": "15232,15233, 15236", "reported": "x"},
        {"id": "3", "title": "c", "client_code": "SFO_APAC", "begin": "13/09/2026 01:25:00",
         "call_reference": "Direct Email", "reported": "x"},
    ])
    assert [t.call_refs for t in parse_carlsource.parse(path)] == [
        ["15230", "15231"], ["15232", "15233", "15236"], []]


def test_link_locks_named_calls_outside_the_time_window(regions):  # noqa: F811
    calls = _calls(
        regions,
        ("15230", dt.datetime(2026, 9, 12, 11, 55), "ZQN_APAC"),
        ("15231", dt.datetime(2026, 9, 12, 18, 0), "ZQN_APAC"),
    )
    t = _ticket(internal_id="178566", client_code="ZQN_APAC",
                begin_date=dt.datetime(2026, 9, 12, 11, 55), call_refs=["15230", "15231"])
    res = link(calls, [t], regions)
    (entry,) = res.entries
    assert entry.kind == "ticket" and entry.data["call_numbers"] == ["15230", "15231"]
    assert not res.unmatched_calls


def test_link_flags_call_ref_missing_from_phone_log(regions):  # noqa: F811
    calls = _calls(regions, ("15230", dt.datetime(2026, 9, 12, 11, 55), "ZQN_APAC"))
    t = _ticket(internal_id="178566", client_code="ZQN_APAC",
                begin_date=dt.datetime(2026, 9, 12, 11, 55), call_refs=["15230", "15999"])
    res = link(calls, [t], regions)
    assert res.entries[0].data["call_numbers"] == ["15230"]
    assert any("15999" in f and "not in the phone log" in f for f in res.flags)


# ---------------------------------------------------------------- ticket text

def test_technical_dump_is_cut_and_capped(tmp_path):
    dump = '{ "autoInduct": { "enable": false } }\n' * 200
    t = _one_ticket(
        tmp_path, id="178240", client_code="SFO_APAC",
        comments="Reported: Bags rejected.\nRemedy: Config corrected and SITA to confirm.\n"
                 "Resolution\nAccess the SBD via ssh host\n" + dump,
    )
    assert t.remedy.startswith("Config corrected and SITA to confirm.")
    assert "omitted here for readability" in t.remedy and "WO 178240" in t.remedy
    assert "autoInduct" not in t.remedy and "ssh" not in t.remedy
    assert len(t.remedy) < 900


def test_long_plain_bullet_is_capped_with_pointer(tmp_path):
    t = _one_ticket(tmp_path, id="555", client_code="SFO_APAC",
                    comments="Reported: " + "It was slow. " * 200)
    assert len(t.reported) < 800 and t.reported.endswith("(see WO 555 for full detail)")


def test_unlabelled_ticket_gets_a_summary(tmp_path):
    t = _one_ticket(
        tmp_path, id="178374", client_code="ST_PERP_APAC",
        comments="***CHG0250397 -Startrack Perth Road\n\nBOTH sites have a planned power outage "
                 "on Saturday. Request Alstef shut the servers down.\nStep 1) ssh in\n",
    )
    assert not (t.reported or t.findings or t.remedy)
    assert t.summary.startswith("CHG0250397 -Startrack Perth Road.")
    assert "planned power outage" in t.summary and "ssh" not in t.summary


def test_glued_labels_and_duplicate_block(tmp_path):
    t = _one_ticket(
        tmp_path, id="178180", client_code="ZNE_APAC",
        comments="Call Reference:\n\n709:\nReportedAll bags stopped.\nFindingsTracking was lost.\n"
                 "Remedy\nRe-tune TC106.\n\nReported: raw duplicate\nRemedy: raw duplicate\n",
    )
    assert (t.reported, t.findings, t.remedy) == (
        "All bags stopped.", "Tracking was lost.", "Re-tune TC106.")


def test_placeholder_only_section_is_empty(tmp_path):
    t = _one_ticket(tmp_path, id="178566", client_code="ZQN_APAC",
                    comments="Reported:\nFindings:\nRemedy: xx/09:\n")
    assert (t.reported, t.findings, t.remedy, t.summary) == ("", "", "", "")


def test_start_date_accepts_text_datetime_and_iso(tmp_path):
    path = tmp_path / "cs.xlsx"
    make_carlsource(path, [
        {"id": "1", "title": "a", "client_code": "ZQN_APAC", "begin": "12/09/2026 11:55:00"},
        {"id": "2", "title": "b", "client_code": "ZQN_APAC", "begin": dt.datetime(2026, 9, 9, 5, 23)},
        {"id": "3", "title": "c", "client_code": "ZQN_APAC", "begin": "2026-09-09 05:23:00"},
    ])
    assert [t.begin_date for t in parse_carlsource.parse(path)] == [
        dt.datetime(2026, 9, 12, 11, 55), dt.datetime(2026, 9, 9, 5, 23),
        dt.datetime(2026, 9, 9, 5, 23)]


def test_unreadable_start_date_is_an_error(tmp_path):
    path = tmp_path / "cs.xlsx"
    make_carlsource(path, [{"id": "9", "title": "a", "client_code": "ZQN_APAC", "begin": "soon"}])
    with pytest.raises(DiscoveryError, match="'9'"):
        parse_carlsource.parse(path)
