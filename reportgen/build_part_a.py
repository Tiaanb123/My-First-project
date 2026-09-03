"""Assemble the Part A .docx from work/bundle.json."""
from __future__ import annotations

import datetime as dt
import json
from pathlib import Path

from . import chart_pie, docx_base
from .dates import filename_tag
from .regions import REGION_ORDER


def build(bundle_path: Path, out_dir: Path) -> Path:
    bundle = json.loads(Path(bundle_path).read_text(encoding="utf-8"))
    week = bundle["week"]
    pa = bundle["part_a"]

    doc = docx_base.open_report()
    docx_base.add_title(doc, "Weekly Support Tickets Report:")
    docx_base.add_subtitle(doc, week["subtitle"])
    docx_base.add_kpi(doc, f"Total Amount of Calls: {pa['totals']['calls']}")
    docx_base.add_kpi(doc, f"Total Amount of Direct Contacts: {pa['totals']['direct_contacts']}")
    docx_base.add_region_list(doc, pa["region_counts"], list(REGION_ORDER))

    for region in pa["regions"]:
        docx_base.add_region_heading(doc, region["name"], region["count"])
        for site in region["sites"]:
            docx_base.add_site_heading(doc, site["name"], site["code"])
            for entry in site["entries"]:
                docx_base.add_entry(doc, entry)

    docx_base.add_signoff(doc)

    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    start = dt.date.fromisoformat(week["start"])
    end = dt.date.fromisoformat(week["end"])
    out = out_dir / f"Weekly_Support_Tickets_Report_{filename_tag(start, end)}.docx"
    doc.save(str(out))

    data = chart_pie.calls_by_site(bundle)
    chart_pie.inject(out, data, f"Weekly Call Logs {end:%d/%m/%Y}")
    return out
