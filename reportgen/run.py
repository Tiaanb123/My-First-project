"""CLI orchestrator for the weekly report pipeline.

Implemented stages:
  discover  - resolve the reporting week + input manifest, print it, write work/manifest.json
  parse     - parse Part A (phone logs + CarlSource) + link -> work/bundle.json

Planned stages (not yet implemented): build.
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

from . import bundle as bundle_mod
from . import parse_carlsource, parse_phonelogs
from .discover import DiscoveryError, Manifest, discover_inputs
from .link import link
from .regions import PART_A_REGIONS, Regions

_REPO_ROOT = Path(__file__).resolve().parent.parent


def _print_manifest(manifest: Manifest) -> None:
    d = manifest.to_dict()
    print(f"Reporting week:  {d['week']['subtitle']}")
    print(f"                 ({d['week']['start']} to {d['week']['end']})")
    print()
    print("PartA:")
    print(f"  phonelogs   {Path(d['part_a']['phonelogs']).name}")
    print(f"  carlsource  {Path(d['part_a']['carlsource']).name}")
    print()
    print("PartB:")
    if not d["part_b"]:
        print("  (none - no Part B reports this week)")
    for row in d["part_b"]:
        print(
            f"  {row['site']:<4} {row['source']:<12} "
            f"week ending {row['week_ending']}   {Path(row['file']).name}"
        )
    if d["ignored"]:
        print()
        print("Ignored (not a recognised input):")
        for name in d["ignored"]:
            print(f"  {Path(name).name}")


def _stage_discover(args: argparse.Namespace) -> int:
    try:
        manifest = discover_inputs(args.dropbox, args.knowledge)
    except DiscoveryError as exc:
        print(f"discover: {exc}", file=sys.stderr)
        return 2

    _print_manifest(manifest)

    out = Path(args.work) / "manifest.json"
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(json.dumps(manifest.to_dict(), indent=2) + "\n", encoding="utf-8")
    print()
    print(f"Wrote {out.relative_to(_REPO_ROOT).as_posix()}")
    return 0


def _stage_parse(args: argparse.Namespace) -> int:
    try:
        manifest = discover_inputs(args.dropbox, args.knowledge)
    except DiscoveryError as exc:
        print(f"parse: {exc}", file=sys.stderr)
        return 2

    regions = Regions.from_yaml(Path(args.knowledge))
    phone = parse_phonelogs.parse(manifest.phonelogs)
    tickets = parse_carlsource.parse(manifest.carlsource)
    links = link(phone.calls, tickets, regions)
    data = bundle_mod.build_part_a(manifest.to_dict()["week"], phone, links, regions)

    out = Path(args.work) / "bundle.json"
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(json.dumps(data, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")

    pa = data["part_a"]
    print(f"Part A:  {pa['totals']['calls']} calls, "
          f"{pa['totals']['direct_contacts']} direct contacts   "
          f"({len(tickets)} CarlSource tickets)")
    grid = pa["region_counts"]
    print("Regions: " + "  ".join(
        f"{r} {grid[r]}" for r in PART_A_REGIONS
    ) + (f"  Other {grid['Other']}" if grid.get("Other") else ""))
    if pa["flags"]:
        print(f"Flags ({len(pa['flags'])}):")
        for f in pa["flags"]:
            print(f"  - {f}")
    if pa["unmatched_calls"]:
        print(f"Unmatched calls: {', '.join(pa['unmatched_calls'])}")
    print(f"\nWrote {out.relative_to(_REPO_ROOT).as_posix()}")
    return 0


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(prog="reportgen.run")
    parser.add_argument("--stage", required=True, choices=["discover", "parse"])
    parser.add_argument("--dropbox", type=Path, default=_REPO_ROOT / "DropBox")
    parser.add_argument("--knowledge", type=Path, default=_REPO_ROOT / "knowledge")
    parser.add_argument("--work", type=Path, default=_REPO_ROOT / "work")
    parser.add_argument("--partb", help="reserved for later stages", default=None)
    args = parser.parse_args(argv)

    if args.stage == "discover":
        return _stage_discover(args)
    if args.stage == "parse":
        return _stage_parse(args)
    parser.error(f"unknown stage {args.stage!r}")
    return 2


if __name__ == "__main__":
    raise SystemExit(main())
