"""Part B orchestration: bundle.json + manifest -> one .docx per requested report.

Ticket entries come from Part A's work/bundle.json (never re-parsed). Each site with a
BHS export in the manifest is extracted (GlideView after recalculation), its BhsStats is
archived to history/<site>/<week_start>.json, and every requested report is rendered.
"""
from __future__ import annotations

import dataclasses
import datetime as dt
import json
from pathlib import Path

from ..discover import DiscoveryError, Manifest, load_sites
from . import recalc, render, sites as sites_mod
from .bhs_bag import extract_adl
from .bhs_glideview import extract_chc, extract_zqn
from .stats import BhsStats

EXTRACTORS = {"ZQN": extract_zqn, "CHC": extract_chc, "ADL": extract_adl}


@dataclasses.dataclass
class PartBResult:
    outputs: list[Path]
    stats: dict[str, BhsStats]
    notices: list[str]           # shown prominently (recalc needed, partial week, ...)
    flags: list[str]             # console detail


def entries_by_code(bundle: dict) -> dict[str, list[dict]]:
    out: dict[str, list[dict]] = {}
    for region in bundle["part_a"]["regions"]:
        for site in region["sites"]:
            out.setdefault(site["code"], []).extend(site["entries"])
    return out


def default_reports(manifest: Manifest) -> list[str]:
    """Every site with a Part B export this week (the 'pre-ticked' set)."""
    return [pb.site for pb in manifest.part_b]


def extract(manifest: Manifest, codes: set[str], work_dir: Path) -> tuple[
        dict[str, BhsStats], list[str]]:
    stats: dict[str, BhsStats] = {}
    notices: list[str] = []
    for pb in manifest.part_b:
        if pb.site not in codes:
            continue
        fn = EXTRACTORS.get(pb.site)
        if fn is None:
            raise DiscoveryError(
                f"{pb.path.name!r}: no BHS extractor for {pb.site} yet - a new site's "
                f"conventions (main-line tag, lost-bag basis, ...) must be worked out from "
                f"its export and added to reportgen/partb before its report can be built"
            )
        path, needed = pb.path, False
        if pb.source == "glideview":
            rc = recalc.recalculated_copy(pb.path, work_dir)
            path, needed = rc.path, rc.needed
            if needed:
                notices.append(f"{pb.site}: {pb.path.name!r} SUMMARY read all zeros until "
                               f"recalculated ({rc.engine}) - figures are from the "
                               f"recalculated copy")
        st = fn(path, manifest.start, manifest.end, recalculated=needed)
        notices += [f"{pb.site}: {f[3:]}" for f in st.flags if f.startswith("!! ")]
        stats[pb.site] = st
    return stats, notices


def archive(stats: dict[str, BhsStats], history_dir: Path) -> None:
    """Spec gap 9: keep every week's figures so week-over-week trends can return later."""
    for code, st in stats.items():
        out = Path(history_dir) / code / f"{st.week_start.isoformat()}.json"
        out.parent.mkdir(parents=True, exist_ok=True)
        out.write_text(st.to_json() + "\n", encoding="utf-8")


def build_part_b(bundle_path: Path, manifest: Manifest, knowledge: Path, out_dir: Path,
                 work_dir: Path, history_dir: Path, names: list[str] | None = None
                 ) -> PartBResult:
    bundle = json.loads(Path(bundle_path).read_text(encoding="utf-8"))
    week = bundle["week"]
    if (dt.date.fromisoformat(week["start"]), dt.date.fromisoformat(week["end"])) != (
            manifest.start, manifest.end):
        raise DiscoveryError(
            f"work/bundle.json is for {week['start']}..{week['end']} but DropBox holds "
            f"{manifest.start}..{manifest.end} - re-run `--stage parse` first")

    by_code = entries_by_code(bundle)
    sites = load_sites(Path(knowledge))
    specs = sites_mod.resolve(names or default_reports(manifest), sites,
                              {c for c, e in by_code.items() if e})
    if not specs:
        raise DiscoveryError("no Part B reports requested and no Part B exports in DropBox/PartB")

    codes = {s.code for spec in specs for s in spec.sites}
    stats, notices = extract(manifest, codes, Path(work_dir))
    archive(stats, history_dir)

    flags = [f"{code}: {f}" for code, st in stats.items() for f in st.flags
             if not f.startswith("!! ")]
    chart_dir = Path(work_dir) / "charts"
    outputs = [
        render.render(spec, week["subtitle"], by_code, stats,
                      Path(out_dir) / spec.filename, chart_dir)
        for spec in specs
    ]
    return PartBResult(outputs, stats, notices, flags)
