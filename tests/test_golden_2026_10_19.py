"""Golden figures for the week Mon 19 - Sun 25 Oct 2026, from the real exports in
DropBox/PartB. Skipped when those files aren't present (they're not committed).

Figures computed by the pipeline on 28 Oct 2026 and listed for Tiaan to confirm; CHC is
from the Excel-recalculated copy.
"""
from __future__ import annotations

import datetime as dt
from pathlib import Path

import pytest

from reportgen.discover import DiscoveryError
from reportgen.partb import metrics
from reportgen.partb.bhs_bag import extract_adl
from reportgen.partb.bhs_glideview import extract_chc, extract_zqn
from reportgen.partb.recalc import recalculated_copy

PART_B = Path(__file__).resolve().parent.parent / "DropBox" / "PartB"
WS, WE = dt.date(2026, 10, 19), dt.date(2026, 10, 25)
ZQN = PART_B / "Queenstown International Airport - GlideView Reports Week Ending Sunday, 25 October 2026.xlsx"
CHC = PART_B / "Christchurch International Airport - GlideView Reports Week Ending Sunday, October 25, 2026.xlsx"
ADL = PART_B / "BAGanalytics Report.xlsx"


def _need(path: Path):
    return pytest.mark.skipif(not path.is_file(), reason=f"{path.name} not in DropBox/PartB")


@_need(ZQN)
def test_zqn_golden():
    st = extract_zqn(ZQN, WS, WE)
    assert (st.main_throughput, st.oversize) == (20573, 1582)
    assert "20,551" in st.main_note  # T114
    assert (st.l1_flagged, st.l1_denominator) == (7240, 19707)
    assert metrics.pct(metrics.l1_pct(st), 1) == "36.7%"
    assert (st.l2_reject, st.l2_timeouts) == (507, 157)
    assert metrics.pct(metrics.l2_pct(st)) == "2.46%"
    assert st.lost_total == 68 and metrics.pct(st.lost_pct) == "0.06%"
    assert dict(st.lost_breakdown) == {"T109PC": 0, "T110PC": 35, "T111PC": 1, "T112PC": 1,
                                       "T1DV1PD": 11, "T1DV1PU": 2, "T110APC": 18}
    assert st.excluded[0][:2] == ("Dieback Status", 2864)
    assert st.faults[:3] == [("Desk has a Fault", 45), ("Emergency Stop Fault", 40),
                             ("PE1 PE Sensor Fault", 27)]
    assert st.full_week
    assert [n for _, n in st.trend] == [3058, 2377, 3014, 2670, 3238, 3001, 3215]


@_need(CHC)
def test_chc_golden(tmp_path):
    try:
        rc = recalculated_copy(CHC, tmp_path)
    except DiscoveryError as exc:
        if "couldn't recalculate" in str(exc):
            pytest.skip("no Excel / LibreOffice to recalculate with")
        raise
    st = extract_chc(rc.path, WS, WE, recalculated=rc.needed)
    assert st.main_throughput == 42785
    assert st.secondary[0][1] == "40,719 bags (C1+C2+C3+Tranship)"
    assert st.oversize == 3861
    assert st.l1_flagged == 16037 and metrics.pct(metrics.l1_pct(st), 1) == "37.5%"
    assert st.l2_reject == 853 and metrics.pct(metrics.l2_pct(st)) == "1.99%"
    assert st.l2_timeouts == 243
    assert st.lost_total == 2231 and metrics.pct(st.lost_pct) == "5.48%"
    assert dict(st.lost_breakdown) == {
        "TC1 & TC2 prior to ME": 176, "Domestic CBS": 867, "International CBS": 666,
        "Exiting ME": 149, "Between ME and L4T01": 20, "S1 Sort Line": 48,
        "S2 Sort Line": 230, "Level 4 Line": 75}
    assert st.faults_total == 946
    assert st.pm_testing == (238, 946)
    assert st.faults[:3] == [("ME Waiting Bag Warning", 218), ("Emergency Stop Fault", 139),
                             ("Intrusion System Health Fault", 98)]
    assert st.trend is None


@_need(ADL)
def test_adl_golden():
    st = extract_adl(ADL, WS, WE)
    assert (st.main_throughput, st.oversize) == (59302, 3339)
    assert st.lost_total == 4134 and st.lost_pct == 6.28
    assert st.lost_breakdown[:3] == [("XT1", 1992), ("SL1", 1031), ("XT2", 369)]
    assert sum(n for _, n in st.lost_breakdown) == 4134
    assert st.excluded[0][:2] == ("Heartbeat Fault", 10)
    assert st.faults[:3] == [("VFD Not Ready Fault", 498), ("MU4 Not Ready Fault", 424),
                             ("Communication Fault", 355)]
    assert st.reject_withheld and st.l2_reject is None
    assert st.trend is None
