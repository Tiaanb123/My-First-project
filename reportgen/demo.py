"""One-shot demo entry point: drop files in DropBox/, run this, get the Word report.

Packaged as WeeklyReport.exe (see build_exe.bat). Chains discover -> (pause) -> parse ->
build, then opens the Output folder.
"""
from __future__ import annotations

import os
import sys
import traceback

from . import build_part_a, config
from .discover import DiscoveryError, discover_inputs
from .run import print_manifest, run_parse

for _stream in (sys.stdout, sys.stderr):
    try:
        _stream.reconfigure(encoding="utf-8")  # type: ignore[attr-defined]
    except (AttributeError, ValueError):
        pass


def _pause(msg: str = "Press Enter to exit... ") -> None:
    try:
        input(msg)
    except EOFError:
        pass


def main() -> int:
    print("=" * 60)
    print("  Alstef  |  Weekly Support Report  |  V1")
    print("=" * 60)
    print(f"\nInput folder:  {config.DROPBOX}")
    print(f"Output folder: {config.OUTPUT}\n")

    try:
        manifest = discover_inputs(config.DROPBOX, config.KNOWLEDGE)
    except DiscoveryError as exc:
        print(f"\n[stopped] {exc}\n")
        _pause()
        return 1

    print_manifest(manifest)

    try:
        input("\nPress Enter to build the report (Ctrl+C to abort)... ")
    except (EOFError, KeyboardInterrupt):
        print("\nAborted.")
        return 1

    try:
        print()
        bundle_path = run_parse(config.DROPBOX, config.KNOWLEDGE, config.WORK)
        out, _chart_data = build_part_a.build(bundle_path, config.OUTPUT)
    except DiscoveryError as exc:
        print(f"\n[stopped] {exc}\n")
        _pause()
        return 1
    except Exception:  # noqa: BLE001 - demo entry point, show the trace and pause
        print("\n[error] something went wrong:\n")
        traceback.print_exc()
        _pause()
        return 1

    print(f"\n[OK] Report written to:\n  {out}\n")
    if sys.platform == "win32":
        try:
            os.startfile(config.OUTPUT)  # noqa: S606
        except OSError:
            pass
    _pause("Press Enter to close... ")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
