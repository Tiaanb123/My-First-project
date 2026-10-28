"""Resolve a site token (phone-log location id, bare code, or CarlSource contract) to a
site record, and know the Part A region order.
"""
from __future__ import annotations

import dataclasses
from pathlib import Path

from .discover import _norm, load_sites

# Spec order for Part A region headings. These 8 always get a heading (even at 0).
PART_A_REGIONS: tuple[str, ...] = (
    "Africa", "Australia", "Bermuda", "Canada",
    "Jamaica", "New Zealand", "United States", "Parcels",
)
# "Other" catches sites whose real region has no Part A heading; it renders only when
# it actually has entries.
REGION_ORDER: tuple[str, ...] = PART_A_REGIONS + ("Other",)


@dataclasses.dataclass(frozen=True)
class SiteRef:
    code: str
    name: str
    region: str
    location_id: str | None = None
    contact_name: str | None = None
    contact_email: str | None = None
    contact_raw: str | None = None
    country: str | None = None  # Parcels sub-heading (default Australia; override in sites_overlay)


class Regions:
    def __init__(self, sites: dict[str, dict]):
        self._sites = sites
        self.by_code: dict[str, SiteRef] = {}
        self.by_location_id: dict[str, str] = {}
        self.by_contract: dict[str, str] = {}
        self.by_norm_name: dict[str, str] = {}
        for code, info in sites.items():
            region = info.get("region", "Other")
            ref = SiteRef(
                code=code,
                name=info.get("name", code),
                region=region,
                country=info.get("country") or ("Australia" if region == "Parcels" else None),
                location_id=info.get("location_id"),
                contact_name=info.get("contact_name"),
                contact_email=info.get("contact_email"),
                contact_raw=info.get("contact_raw"),
            )
            self.by_code[code] = ref
            if info.get("location_id"):
                self.by_location_id[info["location_id"].upper()] = code
            for contract in info.get("contracts", []) or []:
                self.by_contract.setdefault(contract.upper(), code)
            if info.get("name"):
                self.by_norm_name.setdefault(_norm(info["name"]), code)

    @classmethod
    def from_yaml(cls, knowledge_dir: Path) -> "Regions":
        return cls(load_sites(knowledge_dir))

    def resolve(self, token: str | None) -> SiteRef | None:
        if not token:
            return None
        t = str(token).strip()
        code = (
            self.by_location_id.get(t.upper())
            or (t.upper() if t.upper() in self.by_code else None)
            or (t[:-5].upper() if t.upper().endswith("_APAC") and t[:-5].upper() in self.by_code else None)
            or self.by_contract.get(t.upper())
            or self.by_norm_name.get(_norm(t))
        )
        return self.by_code.get(code) if code else None

    def region_for(self, token: str | None) -> str:
        ref = self.resolve(token)
        return ref.region if ref else "Other"
