"""Generate `knowledge/sites.yaml` from the site lookup workbook + `sites_overlay.yaml`.

    python -m reportgen.build_sites [--lookup PATH] [--out PATH]

The lookup workbook ("LookUp - Sites and Emails.xlsx") is the authoritative table of
sites, contracts, regions and facilities contacts. It does not need to live in the repo -
see `resolve_lookup_path()` for where we look. The generated `sites.yaml` is checked in;
re-run this whenever the lookup changes.
"""
from __future__ import annotations

import argparse
import os
import re
import sys
from pathlib import Path

import yaml

from .discover import _load_workbook
from .regions import PART_A_REGIONS

_REPO_ROOT = Path(__file__).resolve().parent.parent
_LOOKUP_NAME = "LookUp - Sites and Emails.xlsx"
_SHEET = "Site Names"

# lookup Region value -> Part A region. Anything not mapped and not already a Part A
# region collapses to "Other".
_REGION_ALIASES = {
    "usa": "United States",
    "united states": "United States",
}

_NEEDS_BACKFILL = {"need bf", "", "n/a", "tbc"}


def resolve_lookup_path(explicit: str | None) -> Path:
    candidates = []
    if explicit:
        candidates.append(Path(explicit))
    if os.environ.get("ALSTEF_SITES_LOOKUP"):
        candidates.append(Path(os.environ["ALSTEF_SITES_LOOKUP"]))
    onedrive = os.environ.get("OneDrive") or os.environ.get("OneDriveCommercial")
    if onedrive:
        candidates.append(Path(onedrive) / "Support Reports" / _LOOKUP_NAME)
    candidates.append(_REPO_ROOT / "Intructions" / "Templates" / _LOOKUP_NAME)
    for path in candidates:
        if path.is_file():
            return path
    tried = "\n  ".join(str(c) for c in candidates)
    raise SystemExit(f"site lookup workbook not found. Tried:\n  {tried}")


def normalise_region(raw: str) -> str:
    value = (raw or "").strip()
    alias = _REGION_ALIASES.get(value.lower())
    if alias:
        return alias
    if value in PART_A_REGIONS:
        return value
    return "Other"


def split_name_and_code(site_field: str) -> str:
    """'Kelowna Airport YLW' -> 'Kelowna Airport'; leave the name if the tail isn't a code."""
    name = (site_field or "").strip()
    m = re.match(r"^(.*\S)\s+([A-Z0-9]{2,5})$", name)
    return m.group(1) if m else name


def parse_contact(emails_field: str) -> dict:
    """Best-effort {contact_name, contact_email} or {contact_raw} from the Emails column."""
    text = (emails_field or "").strip()
    if not text:
        return {}
    first = text.splitlines()[0].strip().rstrip(",")
    email = re.search(r"[\w.+-]+@[\w-]+\.[\w.-]+", first)
    if not email:
        return {"contact_raw": first}
    out = {"contact_email": email.group(0)}
    before = first[: email.start()].rstrip(" -–\t")
    if before and before.lower() not in {"cc:", "email"}:
        out["contact_name"] = before
    return out


def load_lookup(path: Path) -> dict[str, dict]:
    wb = _load_workbook(path)
    try:
        if _SHEET not in wb.sheetnames:
            raise SystemExit(f"{path.name}: no '{_SHEET}' sheet")
        rows = wb[_SHEET].iter_rows(values_only=True)
        header = next(rows)
        # Column layout (a leading blank column A): CS Client, CS Cost Centre, Site,
        # WB Code, Status, Region, Emails.
        idx = {str(h).strip().lower(): i for i, h in enumerate(header) if h}
        c_client = idx.get("cs client", 1)
        c_cost = idx.get("cs cost centre", 2)
        c_site = idx.get("site", 3)
        c_status = idx.get("status", 5)
        c_region = idx.get("region", 6)
        c_email = idx.get("emails", 7)

        sites: dict[str, dict] = {}
        for row in rows:
            client = (row[c_client] or "").strip() if c_client < len(row) else ""
            if not client:
                continue
            code = client[:-5] if client.upper().endswith("_APAC") else client
            entry = sites.setdefault(
                code,
                {"location_id": client, "contracts": [], "_names": set()},
            )
            name = split_name_and_code(row[c_site] if c_site < len(row) else "")
            if name:
                entry["_names"].add(name)
            region_raw = (row[c_region] or "").strip() if c_region < len(row) else ""
            entry["region"] = normalise_region(region_raw)
            if region_raw and normalise_region(region_raw) == "Other":
                entry["region_raw"] = region_raw
            status = (row[c_status] or "").strip() if c_status < len(row) else ""
            if status:
                entry["status"] = status
            cost = (row[c_cost] or "").strip() if c_cost < len(row) else ""
            if cost and cost.lower() not in _NEEDS_BACKFILL and cost not in entry["contracts"]:
                entry["contracts"].append(cost)
            contact = parse_contact(row[c_email] if c_email < len(row) else "")
            for k, v in contact.items():
                entry.setdefault(k, v)
        return sites
    finally:
        wb.close()


def finalise(sites: dict[str, dict], overlay: dict) -> dict:
    overrides = (overlay or {}).get("overrides", {}) or {}
    out: dict[str, dict] = {}
    for code in sorted(sites):
        raw = sites[code]
        names = sorted(raw.pop("_names", []) or [], key=len)
        entry = {"location_id": raw["location_id"]}
        if names:
            entry["name"] = names[-1]
        entry["region"] = raw.get("region", "Other")
        if "region_raw" in raw:
            entry["region_raw"] = raw["region_raw"]
        if raw.get("contracts"):
            entry["contracts"] = raw["contracts"]
        for key in ("status", "contact_name", "contact_email", "contact_raw"):
            if raw.get(key):
                entry[key] = raw[key]
        entry.update(overrides.get(code, {}))
        out[code] = entry
    return out


def render_yaml(sites: dict[str, dict]) -> str:
    header = (
        "# GENERATED by `python -m reportgen.build_sites` - do not hand-edit.\n"
        "# Source: LookUp - Sites and Emails.xlsx  +  knowledge/sites_overlay.yaml\n"
    )
    doc = {
        "part_a_region_order": list(PART_A_REGIONS) + ["Other"],
        "sites": sites,
    }
    return header + yaml.safe_dump(doc, sort_keys=False, allow_unicode=True, width=100)


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(prog="reportgen.build_sites")
    parser.add_argument("--lookup", default=None)
    parser.add_argument("--out", type=Path, default=_REPO_ROOT / "knowledge" / "sites.yaml")
    parser.add_argument("--overlay", type=Path,
                        default=_REPO_ROOT / "knowledge" / "sites_overlay.yaml")
    args = parser.parse_args(argv)

    lookup_path = resolve_lookup_path(args.lookup)
    overlay = yaml.safe_load(args.overlay.read_text(encoding="utf-8")) if args.overlay.is_file() else {}
    sites = finalise(load_lookup(lookup_path), overlay)

    args.out.write_text(render_yaml(sites), encoding="utf-8")
    regions: dict[str, int] = {}
    for info in sites.values():
        regions[info["region"]] = regions.get(info["region"], 0) + 1
    print(f"Wrote {args.out.relative_to(_REPO_ROOT).as_posix()}  ({len(sites)} sites)")
    print(f"Source: {lookup_path}")
    print("By region: " + "  ".join(f"{k} {v}" for k, v in sorted(regions.items())))
    return 0


if __name__ == "__main__":
    sys.exit(main())
