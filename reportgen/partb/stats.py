"""`BhsStats` - the one data contract between the BHS extractors and the renderer.

All site-specific logic (which tag is main-line throughput, what the lost-bag % is a
percentage *of*, which counters are excluded and why) lives in the extractors and lands
here as values + note strings. The renderer and wording builders read only this.
"""
from __future__ import annotations

import dataclasses
import datetime as dt
import json


@dataclasses.dataclass
class BhsStats:
    site: str
    source: str                         # "GlideView" | "BAGanalytics"
    week_start: dt.date
    week_end: dt.date
    # days the export actually covers inside the reporting week (all 7 for a full week)
    days_covered: list[dt.date]

    main_throughput: int
    main_label: str                     # "T101" / "X1X0100-X1X0400 EDS" / ""
    main_note: str
    oversize: int
    oversize_label: str
    lost_total: int
    lost_pct: float                     # percent, e.g. 0.07 means 0.07%
    lost_pct_basis: str                 # "tracked scans" / "bags entering sort system"
    lost_note: str
    lost_granularity: str               # "Checkpoint" | "Zone"
    lost_breakdown: list[tuple[str, int]]
    faults: list[tuple[str, int]]       # every genuine category, descending
    faults_total: int                   # sum of `faults` (excluded counters not included)

    # L1/L2 - None when the export can't support them (ADL); see `reject_withheld`
    l1_flagged: int | None = None
    l1_denominator: int | None = None
    l1_basis: str = "bags screened at L1"
    l2_reject: int | None = None
    l2_timeouts: int | None = None
    reject_withheld: str | None = None  # reason, when L1/L2 are withheld

    secondary: list[tuple[str, str, str]] = dataclasses.field(default_factory=list)
    secondary_tags: str = ""            # e.g. "T114 secondary tag read 20,551"
    excluded: list[tuple[str, int, str]] = dataclasses.field(default_factory=list)
    pm_testing: tuple[int, int] | None = None   # (PM-driven fault count, faults_total)
    pm_categories: list[str] = dataclasses.field(default_factory=list)
    trend: list[tuple[dt.date, int]] | None = None
    trend_note: str = ""                # why there's no trend chart, when there isn't
    limitations: list[str] = dataclasses.field(default_factory=list)  # standing caveats
    recalculated: bool = False          # SUMMARY needed recalculation before reading
    flags: list[str] = dataclasses.field(default_factory=list)        # console-only

    @property
    def single_day(self) -> bool:
        return len(self.days_covered) == 1

    @property
    def full_week(self) -> bool:
        return len(self.days_covered) == 7

    def to_json(self) -> str:
        def default(o):
            if isinstance(o, dt.date):
                return o.isoformat()
            raise TypeError(type(o))
        return json.dumps(dataclasses.asdict(self), indent=2, default=default,
                          ensure_ascii=False)

    @classmethod
    def from_json(cls, text: str) -> "BhsStats":
        data = json.loads(text)
        for key in ("week_start", "week_end"):
            data[key] = dt.date.fromisoformat(data[key])
        data["days_covered"] = [dt.date.fromisoformat(d) for d in data["days_covered"]]
        if data.get("trend") is not None:
            data["trend"] = [(dt.date.fromisoformat(d), n) for d, n in data["trend"]]
        for key in ("lost_breakdown", "faults", "secondary", "excluded"):
            data[key] = [tuple(x) for x in data.get(key) or []]
        if data.get("pm_testing") is not None:
            data["pm_testing"] = tuple(data["pm_testing"])
        return cls(**data)
