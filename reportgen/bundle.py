"""Assemble `work/bundle.json` — the deterministic data contract for the report build."""
from __future__ import annotations

import datetime as dt

from .link import LinkResult
from .parse_phonelogs import PhoneData
from .regions import PART_A_REGIONS, REGION_ORDER, Regions


def build_part_a(
    week: dict,
    phone: PhoneData,
    links: LinkResult,
    regions: Regions,
) -> dict:
    # region grid = one per phone call (by caller's region) + one per direct-contact ticket
    region_counts = {r: 0 for r in REGION_ORDER}
    for call in phone.calls:
        region_counts[regions.region_for(call.location_id)] += 1
    for entry in links.entries:
        if entry.kind == "direct":
            region_counts[entry.region] += 1

    by_region: dict[str, dict[str, dict]] = {}
    for entry in links.entries:
        region = entry.region
        site_code = entry.site_code or "UNKNOWN"
        site_name = regions.by_code[site_code].name if site_code in regions.by_code else (
            entry.data.get("location_id") or site_code
        )
        sites = by_region.setdefault(region, {})
        site = sites.setdefault(site_code, {"code": site_code, "name": site_name,
                                            "country": None, "_entries": []})
        site["_entries"].append(entry)

    regions_out = []
    for region in REGION_ORDER:
        sites = by_region.get(region, {})
        if region not in PART_A_REGIONS and not sites:
            continue  # "Other" only shows when it has content
        site_list = []
        for site in sorted(sites.values(), key=lambda s: s["name"].lower()):
            entries = [
                {"kind": e.kind, **e.data}
                for e in sorted(site["_entries"], key=lambda e: e.sort_key)
            ]
            site_list.append({"code": site["code"], "name": site["name"],
                              "country": site["country"], "entries": entries})
        regions_out.append({"name": region, "count": region_counts[region], "sites": site_list})

    return {
        "week": week,
        "generated_at": dt.datetime.now().isoformat(timespec="seconds"),
        "part_a": {
            "totals": {
                "calls": len(phone.calls),
                "direct_contacts": sum(1 for e in links.entries if e.kind == "direct"),
            },
            "phone_stats": {
                "after_hours_pct": phone.after_hours_pct,
                "lines": phone.stats_lines,
            },
            "region_counts": region_counts,
            "regions": regions_out,
            "flags": links.flags,
            "unmatched_tickets": links.unmatched_tickets,
            "unmatched_calls": links.unmatched_calls,
        },
    }
