"""Part B report registry (spec section 4): which reports exist and who's in a group.

Individual reports are one site each; group reports (Canadian Sites, StarTrack) put every
member site under its own subheading. Membership comes from knowledge/sites.yaml so a new
Canadian / ST_* site joins its group automatically. Facilities contacts ("Prepared for")
come from the `prepared_for` / `prepared_for_cc` overlay keys - never guessed.
"""
from __future__ import annotations

import dataclasses
import re

from ..discover import DiscoveryError

STANDING_INDIVIDUAL = ("ADL", "AMX", "BDA", "PER", "SFO", "ZQN", "DUD", "CHC")
GROUPS = {
    "Canadian Sites": lambda code, info: info.get("region") == "Canada",
    "StarTrack": lambda code, info: code.startswith("ST_"),
}
# group members shown even with no entries that week
_LISTED_STATUSES = {"active"}


@dataclasses.dataclass(frozen=True)
class SiteInfo:
    code: str
    name: str
    prepared_for: str | None = None
    prepared_for_cc: str | None = None


@dataclasses.dataclass(frozen=True)
class ReportSpec:
    key: str                    # "ZQN" / "Canadian Sites" - the title and filename stem
    is_group: bool
    sites: tuple[SiteInfo, ...]

    @property
    def title(self) -> str:
        return f"{self.key} Weekly Phone Logs Report & BHS Statistics"

    @property
    def filename(self) -> str:
        stem = re.sub(r"[^A-Za-z0-9]+", "_", self.key).strip("_")
        return f"{stem}_Weekly_Phone_Logs_and_BHS_Statistics.docx"


def _info(code: str, sites: dict[str, dict]) -> SiteInfo:
    raw = sites.get(code, {})
    return SiteInfo(code=code, name=raw.get("name", code),
                    prepared_for=raw.get("prepared_for"),
                    prepared_for_cc=raw.get("prepared_for_cc"))


def resolve(names: list[str], sites: dict[str, dict],
            codes_with_entries: set[str] = frozenset()) -> list[ReportSpec]:
    """Turn the requested report names into ReportSpecs. Anything that isn't a standing
    report is a hard stop - the user decides whether it's a one-off or a new standing one."""
    groups_by_norm = {g.lower(): g for g in GROUPS}
    out: list[ReportSpec] = []
    unknown: list[str] = []
    for raw in names:
        name = raw.strip()
        if name.lower() in groups_by_norm:
            group = groups_by_norm[name.lower()]
            members = sorted(
                (code for code, info in sites.items()
                 if GROUPS[group](code, info)
                 and (str(info.get("status", "")).lower() in _LISTED_STATUSES
                      or code in codes_with_entries)),
                key=lambda c: sites[c].get("name", c).lower(),
            )
            out.append(ReportSpec(group, True, tuple(_info(c, sites) for c in members)))
        elif name.upper() in STANDING_INDIVIDUAL:
            out.append(ReportSpec(name.upper(), False, (_info(name.upper(), sites),)))
        else:
            unknown.append(name)
    if unknown:
        raise DiscoveryError(
            f"{', '.join(repr(u) for u in unknown)} is not a standing Part B report "
            f"(individual: {', '.join(STANDING_INDIVIDUAL)}; groups: {', '.join(GROUPS)}). "
            f"Is it a one-off, or should it be added as a standing report? Add it to "
            f"reportgen/partb/sites.py once decided."
        )
    return out
