from __future__ import annotations

import datetime as dt
from pathlib import Path

import pytest

from reportgen import dates
from reportgen.discover import (
    DiscoveryError,
    _validate_week,
    baganalytics_week,
    discover_inputs,
    parse_glideview,
    parse_parta_week,
    phone_log_date_span,
    resolve_baganalytics_site,
    resolve_glideview_site,
)

REPO_ROOT = Path(__file__).resolve().parent.parent
DROPBOX = REPO_ROOT / "DropBox"

WEEK_START = dt.date(2026, 8, 24)
WEEK_END = dt.date(2026, 8, 30)


# --------------------------------------------------------------------------------------
# dates helpers
# --------------------------------------------------------------------------------------


def test_format_subtitle_year_on_end_only():
    assert (
        dates.format_subtitle(WEEK_START, WEEK_END)
        == "Monday 24th August to Sunday 30th August 2026"
    )


def test_filename_tag():
    assert dates.filename_tag(WEEK_START, WEEK_END) == "24.08.2026_to_30.08.2026"


@pytest.mark.parametrize(
    "text",
    ["August 30, 2026", "30 August 2026", "August 30 2026", "30 August, 2026", "30 August 2026."],
)
def test_parse_english_date_forms(text):
    assert dates.parse_english_date(text) == WEEK_END


def test_parse_us_date():
    assert dates.parse_us_date("8/24/2026") == WEEK_START


def test_parse_dmy_abbr():
    assert dates.parse_dmy_abbr("24-Aug-2026") == WEEK_START


def test_coerce_cell_date_serial_and_datetime():
    assert dates.coerce_cell_date(dt.datetime(2026, 8, 24, 1, 38)) == WEEK_START
    # 46258 is the Excel 1900-system serial for 2026-08-24.
    assert dates.coerce_cell_date(46258) == WEEK_START
    assert dates.coerce_cell_date(0.25) is None  # time-only fraction


# --------------------------------------------------------------------------------------
# Part A filename parsing
# --------------------------------------------------------------------------------------


def test_parse_parta_week_ok():
    got = parse_parta_week(Path("Phonelogs - 24.08 to 30.08.2026.xlsx"))
    assert (got.kind, got.start, got.end) == ("phonelogs", WEEK_START, WEEK_END)
    assert parse_parta_week(Path("Carlsource - 24.08 to 30.08.2026.xls")).kind == "carlsource"


def test_parse_parta_week_december_january_wrap():
    got = parse_parta_week(Path("Phonelogs - 28.12 to 03.01.2027.xlsx"))
    assert got.start == dt.date(2026, 12, 28)
    assert got.end == dt.date(2027, 1, 3)


@pytest.mark.parametrize(
    "name",
    [
        "Phonelogs 24.08 to 30.08.2026.xlsx",       # missing " - "
        "Phonelogs - 24.08 to 30.08.xlsx",          # missing year
        "Phonelogs - 24-08 to 30-08-2026.xlsx",     # wrong separators
        "Randomfile - 24.08 to 30.08.2026.xlsx",    # wrong static token
    ],
)
def test_parse_parta_week_rejects_bad_names(name):
    with pytest.raises(DiscoveryError):
        parse_parta_week(Path(name))


def test_parse_parta_week_rejects_impossible_date():
    with pytest.raises(DiscoveryError):
        parse_parta_week(Path("Phonelogs - 31.02 to 30.08.2026.xlsx"))


# --------------------------------------------------------------------------------------
# week validation
# --------------------------------------------------------------------------------------


def test_validate_week_ok():
    _validate_week(WEEK_START, WEEK_END)


def test_validate_week_start_not_monday():
    with pytest.raises(DiscoveryError, match="Monday"):
        _validate_week(dt.date(2026, 8, 25), dt.date(2026, 8, 30))


def test_validate_week_end_not_sunday():
    with pytest.raises(DiscoveryError, match="Sunday"):
        _validate_week(dt.date(2026, 8, 24), dt.date(2026, 8, 29))


def test_validate_week_wrong_length():
    with pytest.raises(DiscoveryError, match="7 days"):
        _validate_week(dt.date(2026, 8, 24), dt.date(2026, 9, 6))


# --------------------------------------------------------------------------------------
# GlideView filename parsing
# --------------------------------------------------------------------------------------


def test_parse_glideview_both_date_forms():
    airport, week = parse_glideview(
        Path("Christchurch International Airport - GlideView Reports Week Ending Sunday, August 30, 2026.xlsx")
    )
    assert airport == "Christchurch International Airport"
    assert week == WEEK_END
    _, week2 = parse_glideview(
        Path("Queenstown International Airport - GlideView Reports Week Ending Sunday, 30 August 2026.xlsx")
    )
    assert week2 == WEEK_END


# --------------------------------------------------------------------------------------
# site resolution
# --------------------------------------------------------------------------------------


SITES = {
    "ZQN": {"airport_name": "Queenstown International Airport", "bhs_source": "glideview"},
    "ADL": {"airport_name": "Adelaide Airport", "bhs_source": "baganalytics"},
}


def test_resolve_glideview_site_is_punctuation_insensitive():
    assert resolve_glideview_site("queenstown  international   airport", SITES) == "ZQN"


def test_resolve_glideview_site_unknown():
    with pytest.raises(DiscoveryError, match="sites.yaml"):
        resolve_glideview_site("Auckland Airport", SITES)


def test_resolve_baganalytics_site():
    assert resolve_baganalytics_site(SITES) == "ADL"
    with pytest.raises(DiscoveryError):
        resolve_baganalytics_site({"A": {"bhs_source": "baganalytics"}, "B": {"bhs_source": "baganalytics"}})


# --------------------------------------------------------------------------------------
# discover_inputs end to end, on a synthetic well-formed drop
# --------------------------------------------------------------------------------------

from tests.conftest import make_baganalytics, make_glideview, make_phonelogs  # noqa: E402

KNOWLEDGE = lambda db: db.parent / "knowledge"  # noqa: E731


def test_discover_inputs_clean_drop(dropbox):
    manifest = discover_inputs(dropbox, KNOWLEDGE(dropbox))
    assert (manifest.start, manifest.end) == (WEEK_START, WEEK_END)
    d = manifest.to_dict()
    assert d["week"]["subtitle"] == "Monday 24th August to Sunday 30th August 2026"
    assert Path(d["part_a"]["phonelogs"]).name == "Phonelogs - 24.08 to 30.08.2026.xlsx"
    assert d["part_b"] == [] and d["ignored"] == []


def test_discover_inputs_full_part_b(dropbox):
    make_glideview(dropbox / "PartB" / "Queenstown International Airport - GlideView Reports Week Ending Sunday, 30 August 2026.xlsx")
    make_glideview(dropbox / "PartB" / "Christchurch International Airport - GlideView Reports Week Ending Sunday, August 30, 2026.xlsx")
    make_baganalytics(dropbox / "PartB" / "ADL - BAGanalytics Report.xlsx", WEEK_START, WEEK_END)
    (dropbox / "PartB" / "BAGanalytics Report.xlsx").unlink(missing_ok=True)

    manifest = discover_inputs(dropbox, KNOWLEDGE(dropbox))
    by_site = {pb.site: pb.source for pb in manifest.part_b}
    assert by_site == {"ADL": "baganalytics", "CHC": "glideview", "ZQN": "glideview"}


def test_discover_inputs_baganalytics_no_prefix_uses_sites_yaml(dropbox):
    make_baganalytics(dropbox / "PartB" / "BAGanalytics Report.xlsx", WEEK_START, WEEK_END)
    manifest = discover_inputs(dropbox, KNOWLEDGE(dropbox))
    assert [(pb.site, pb.source) for pb in manifest.part_b] == [("ADL", "baganalytics")]


def test_discover_inputs_unknown_code_prefix(dropbox):
    make_baganalytics(dropbox / "PartB" / "XYZ - BAGanalytics Report.xlsx", WEEK_START, WEEK_END)
    with pytest.raises(DiscoveryError, match="XYZ"):
        discover_inputs(dropbox, KNOWLEDGE(dropbox))


def test_discover_inputs_part_b_wrong_week(dropbox):
    make_glideview(dropbox / "PartB" / "Queenstown International Airport - GlideView Reports Week Ending Sunday, 23 August 2026.xlsx")
    with pytest.raises(DiscoveryError, match="blending two weeks|week ending"):
        discover_inputs(dropbox, KNOWLEDGE(dropbox))


def test_discover_inputs_phone_log_row_outside_week_hard_stops(dropbox):
    make_phonelogs(
        dropbox / "PartA" / "Phonelogs - 24.08 to 30.08.2026.xlsx",
        [dt.datetime(2026, 8, 24, 1, 0), dt.datetime(2026, 8, 31, 1, 30)],
    )
    with pytest.raises(DiscoveryError, match="outside its named week"):
        discover_inputs(dropbox, KNOWLEDGE(dropbox))


def test_discover_inputs_mismatched_part_a_files(dropbox):
    (dropbox / "PartA" / "Carlsource - 24.08 to 30.08.2026.xlsx").unlink()
    (dropbox / "PartA" / "Carlsource - 17.08 to 23.08.2026.xlsx").write_text("dummy", "utf-8")
    with pytest.raises(DiscoveryError, match="different weeks|missing"):
        discover_inputs(dropbox, KNOWLEDGE(dropbox))


def test_discover_inputs_unknown_glideview_airport(dropbox):
    make_glideview(dropbox / "PartB" / "Gotham International Airport - GlideView Reports Week Ending Sunday, 30 August 2026.xlsx")
    with pytest.raises(DiscoveryError, match="sites.yaml"):
        discover_inputs(dropbox, KNOWLEDGE(dropbox))


def test_parse_parta_week_accepts_phone_logs_spelling():
    got = parse_parta_week(Path("Phone Logs - 19.10 to 25.10.2026.xlsx"))
    assert (got.kind, got.start, got.end) == (
        "phonelogs", dt.date(2026, 10, 19), dt.date(2026, 10, 25))


def test_discover_inputs_phone_logs_spelling(dropbox):
    part_a = dropbox / "PartA"
    (part_a / "Phonelogs - 24.08 to 30.08.2026.xlsx").rename(
        part_a / "Phone Logs - 24.08 to 30.08.2026.xlsx")
    manifest = discover_inputs(dropbox, KNOWLEDGE(dropbox))
    assert manifest.phonelogs.name == "Phone Logs - 24.08 to 30.08.2026.xlsx"
    assert manifest.notices == []


def test_discover_inputs_swapped_part_a_files_detected_by_content(dropbox):
    part_a = dropbox / "PartA"
    phone = part_a / "Phonelogs - 24.08 to 30.08.2026.xlsx"
    carl = part_a / "Carlsource - 24.08 to 30.08.2026.xlsx"
    tmp = part_a / "swap.tmp"
    phone.rename(tmp)
    carl.rename(phone)
    tmp.rename(carl)

    manifest = discover_inputs(dropbox, KNOWLEDGE(dropbox))
    assert manifest.phonelogs == carl   # holds the 'Phone Logs' sheet
    assert manifest.carlsource == phone  # holds the 'WO Report' sheet
    assert len(manifest.notices) == 2
    assert all("using it as the" in n for n in manifest.notices)


def test_discover_inputs_ignores_unrecognised_part_b_file(dropbox):
    (dropbox / "PartB" / "notes.txt").write_text("random", "utf-8")
    manifest = discover_inputs(dropbox, KNOWLEDGE(dropbox))
    assert [Path(p).name for p in manifest.to_dict()["ignored"]] == ["notes.txt"]


# --------------------------------------------------------------------------------------
# smoke tests against the committed sample drop (if present)
# --------------------------------------------------------------------------------------

needs_samples = pytest.mark.skipif(
    not (DROPBOX / "PartA" / "Phonelogs - 24.08 to 30.08.2026.xlsx").is_file(),
    reason="committed sample drop not present",
)


@needs_samples
def test_sample_discovers_cleanly():
    manifest = discover_inputs(DROPBOX, REPO_ROOT / "knowledge")
    assert (manifest.start, manifest.end) == (WEEK_START, WEEK_END)
    assert {pb.site: pb.source for pb in manifest.part_b} == {
        "ADL": "baganalytics", "CHC": "glideview", "ZQN": "glideview"
    }
    assert manifest.ignored == []
    lo, hi = phone_log_date_span(DROPBOX / "PartA" / "Phonelogs - 24.08 to 30.08.2026.xlsx")
    assert WEEK_START <= lo <= hi <= WEEK_END


@needs_samples
def test_sample_baganalytics_week():
    bag = next((DROPBOX / "PartB").glob("*BAGanalytics Report.xlsx"))
    assert baganalytics_week(bag) == (WEEK_START, WEEK_END)


@needs_samples
def test_sample_glideview_names_resolve():
    for gv in (DROPBOX / "PartB").glob("*GlideView Reports*.xlsx"):
        airport, week = parse_glideview(gv)
        assert week == WEEK_END and airport
