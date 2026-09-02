"""CLI orchestrator for the weekly report pipeline.

Implemented stages:
  discover  - resolve the reporting week + input manifest, print it, write work/manifest.json

Planned stages (not yet implemented): parse, build.
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

from .discover import DiscoveryError, Manifest, discover_inputs

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


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(prog="reportgen.run")
    parser.add_argument("--stage", required=True, choices=["discover"])
    parser.add_argument("--dropbox", type=Path, default=_REPO_ROOT / "DropBox")
    parser.add_argument("--knowledge", type=Path, default=_REPO_ROOT / "knowledge")
    parser.add_argument("--work", type=Path, default=_REPO_ROOT / "work")
    parser.add_argument("--partb", help="reserved for later stages", default=None)
    args = parser.parse_args(argv)

    if args.stage == "discover":
        return _stage_discover(args)
    parser.error(f"unknown stage {args.stage!r}")
    return 2


if __name__ == "__main__":
    raise SystemExit(main())
