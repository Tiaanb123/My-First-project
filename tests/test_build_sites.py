from __future__ import annotations

from reportgen.build_sites import finalise, normalise_region, parse_contact, split_name_and_code


def test_normalise_region():
    assert normalise_region("USA") == "United States"
    assert normalise_region("New Zealand") == "New Zealand"
    assert normalise_region("Fiji") == "Other"
    assert normalise_region("") == "Other"


def test_split_name_and_code():
    assert split_name_and_code("Kelowna Airport YLW") == "Kelowna Airport"
    assert split_name_and_code("Air NZ Domestic ANZ") == "Air NZ Domestic"
    assert split_name_and_code("Union Station (Toronto) VIA UST") == "Union Station (Toronto) VIA"
    assert split_name_and_code("General") == "General"


def test_parse_contact():
    assert parse_contact("Sean Whittaker - sean.whittaker@qa.co.nz\nCC: someone else") == {
        "contact_name": "Sean Whittaker", "contact_email": "sean.whittaker@qa.co.nz"}
    assert parse_contact("cmadeiros@skyport.bm") == {"contact_email": "cmadeiros@skyport.bm"}
    assert parse_contact("Send feedback form to caller") == {
        "contact_raw": "Send feedback form to caller"}
    assert parse_contact("") == {}


def test_finalise_merges_overlay_and_keeps_region_raw():
    raw = {
        "ZQN": {"location_id": "ZQN_APAC", "region": "New Zealand", "contracts": ["MNTZQN"],
                "status": "Active", "_names": {"Queenstown Airport"}},
        "NAN": {"location_id": "NAN_APAC", "region": "Other", "region_raw": "Fiji",
                "contracts": [], "_names": {"Nadi Airport"}},
    }
    overlay = {"overrides": {"ZQN": {"bhs_source": "glideview", "airport_name": "Queenstown International Airport"}}}
    out = finalise(raw, overlay)
    assert out["ZQN"]["bhs_source"] == "glideview"
    assert out["ZQN"]["airport_name"] == "Queenstown International Airport"
    assert out["ZQN"]["name"] == "Queenstown Airport"
    assert out["NAN"]["region"] == "Other" and out["NAN"]["region_raw"] == "Fiji"
