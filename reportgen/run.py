"""CLI orchestrator for the weekly report pipeline.

  discover  - resolve the reporting week + input manifest -> work/manifest.json
  parse     - parse Part A (phone logs + CarlSource) + link -> work/bundle.json
  build     - render Part A -> Output/Weekly_Support_Tickets_Report_<week>.docx
  partb     - render Part B -> Output/<SITE-OR-GROUP>_Weekly_Phone_Logs_and_BHS_Statistics.docx
              (default: every site with a Part B export; or --sites ZQN CHC "Canadian Sites")

See `reportgen.demo` for the one-shot drop-files-and-run entry point.
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

from . import build_part_a, config
from . import bundle as bundle_mod
from . import parse_carlsource, parse_phonelogs
from .discover import DiscoveryError, Manifest, discover_inputs
from .link import link
from .regions import PART_A_REGIONS, Regions


def _rel(path: Path) -> str:
    try:
        return path.resolve().relative_to(config.REPO_ROOT).as_posix()
    except ValueError:
        return str(path)


def print_manifest(manifest: Manifest) -> None:
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
        print(f"  {row['site']:<4} {row['source']:<12} "
              f"week ending {row['week_ending']}   {Path(row['file']).name}")
    if d["ignored"]:
        print("\nIgnored (not a recognised input):")
        for name in d["ignored"]:
            print(f"  {Path(name).name}")
    for notice in d["notices"]:
        print(f"\n!! NOTICE: {notice}")


def run_parse(dropbox: Path, knowledge: Path, work: Path) -> Path:
    """discover + parse Part A -> work/bundle.json (returns the path). Raises DiscoveryError."""
    manifest = discover_inputs(dropbox, knowledge)
    regions = Regions.from_yaml(Path(knowledge))
    phone = parse_phonelogs.parse(manifest.phonelogs)
    tickets = parse_carlsource.parse(manifest.carlsource)
    links = link(phone.calls, tickets, regions)
    data = bundle_mod.build_part_a(manifest.to_dict()["week"], phone, links, regions)

    out = Path(work) / "bundle.json"
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(json.dumps(data, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")

    pa = data["part_a"]
    print(f"Part A:  {pa['totals']['calls']} calls, "
          f"{pa['totals']['direct_contacts']} direct contacts   "
          f"({len(tickets)} CarlSource tickets)")
    grid = pa["region_counts"]
    print("Regions: " + "  ".join(f"{r} {grid[r]}" for r in PART_A_REGIONS)
          + (f"  Other {grid['Other']}" if grid.get("Other") else ""))
    for flag in pa["flags"]:
        print(f"  - {flag}")
    return out


def _stage_discover(args) -> int:
    try:
        manifest = discover_inputs(args.dropbox, args.knowledge)
    except DiscoveryError as exc:
        print(f"discover: {exc}", file=sys.stderr)
        return 2
    print_manifest(manifest)
    out = Path(args.work) / "manifest.json"
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(json.dumps(manifest.to_dict(), indent=2) + "\n", encoding="utf-8")
    print(f"\nWrote {_rel(out)}")
    return 0


def _stage_parse(args) -> int:
    try:
        out = run_parse(args.dropbox, args.knowledge, args.work)
    except DiscoveryError as exc:
        print(f"parse: {exc}", file=sys.stderr)
        return 2
    print(f"\nWrote {_rel(out)}")
    return 0


def _stage_build(args) -> int:
    bundle_path = Path(args.work) / "bundle.json"
    if not bundle_path.is_file():
        print(f"build: {_rel(bundle_path)} not found - run `--stage parse` first",
              file=sys.stderr)
        return 2
    out, data = build_part_a.build(bundle_path, args.output)
    print(f"Wrote {_rel(out)}")
    print("Pie (calls per site): " + ", ".join(f"{c} {n}" for c, n in data))
    print("Open it in Word to confirm the chart renders under the date.")
    return 0


def run_part_b(args_or_paths, names: list[str] | None = None):
    """discover + Part B build; prints notices/flags. Returns PartBResult."""
    from .partb.build import build_part_b

    dropbox, knowledge, work, output, history = args_or_paths
    manifest = discover_inputs(dropbox, knowledge)
    bundle_path = Path(work) / "bundle.json"
    if not bundle_path.is_file():
        raise DiscoveryError(f"{_rel(bundle_path)} not found - run `--stage parse` first")
    result = build_part_b(bundle_path, manifest, knowledge, output, work, history, names)
    for notice in manifest.notices + result.notices:
        print(f"!! NOTICE: {notice}")
    for flag in result.flags:
        print(f"  - {flag}")
    for out in result.outputs:
        print(f"Wrote {_rel(out)}")
    return result


def _stage_partb(args) -> int:
    try:
        run_part_b((args.dropbox, args.knowledge, args.work, args.output, args.history),
                   args.sites or None)
    except DiscoveryError as exc:
        print(f"partb: {exc}", file=sys.stderr)
        return 2
    return 0


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(prog="reportgen.run")
    parser.add_argument("--stage", required=True,
                        choices=["discover", "parse", "build", "partb"])
    parser.add_argument("--dropbox", type=Path, default=config.DROPBOX)
    parser.add_argument("--knowledge", type=Path, default=config.KNOWLEDGE)
    parser.add_argument("--work", type=Path, default=config.WORK)
    parser.add_argument("--output", type=Path, default=config.OUTPUT)
    parser.add_argument("--history", type=Path, default=config.HISTORY)
    parser.add_argument("--sites", nargs="*", default=None,
                        help="Part B reports to build (site codes or 'Canadian Sites' / "
                             "'StarTrack'); default = every site with a Part B export")
    args = parser.parse_args(argv)

    stages = {"discover": _stage_discover, "parse": _stage_parse, "build": _stage_build,
              "partb": _stage_partb}
    return stages[args.stage](args)


if __name__ == "__main__":
    raise SystemExit(main())
