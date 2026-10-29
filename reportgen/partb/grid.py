"""Label-addressed access to export sheets (never fixed row numbers - spec section 5).

GlideView / BAGanalytics exports shift rows between weeks and store numbers as text
('59,302', '0.00', '6.28%'). `Grid` loads a sheet into memory once and finds values by
the label next to (right of) or above (below) them.
"""
from __future__ import annotations

import re
from pathlib import Path

from ..discover import DiscoveryError


def num(value) -> float | None:
    """A cell value as a number: ints/floats as-is; text with thousands separators,
    a trailing '%' or surrounding spaces parsed; blanks -> None."""
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip().replace(",", "").replace("\xa0", "")
    if not text:
        return None
    text = text.rstrip("%").strip()
    try:
        return float(text)
    except ValueError:
        return None


def as_int(value) -> int | None:
    n = num(value)
    return None if n is None else int(round(n))


def _norm(text) -> str:
    return re.sub(r"\s+", " ", str(text)).strip().lower()


class Grid:
    def __init__(self, rows: list[tuple], name: str, source: Path):
        self.rows = [list(r) for r in rows]
        self.name = name
        self.source = Path(source)

    @classmethod
    def from_sheet(cls, wb, name: str, source: Path) -> "Grid":
        matches = [s for s in wb.sheetnames if s.strip().lower() == name.strip().lower()]
        if not matches:
            raise DiscoveryError(f"{Path(source).name!r}: no {name!r} sheet")
        return cls(list(wb[matches[0]].iter_rows(values_only=True)), matches[0], source)

    def cell(self, r: int, c: int):
        if 0 <= r < len(self.rows) and 0 <= c < len(self.rows[r]):
            return self.rows[r][c]
        return None

    def find(self, label: str, start: int = 0, end: int | None = None,
             required: bool = True) -> tuple[int, int] | None:
        """(row, col) of the first cell whose whitespace-normalised text equals `label`."""
        want = _norm(label)
        stop = len(self.rows) if end is None else min(end, len(self.rows))
        for r in range(start, stop):
            for c, value in enumerate(self.rows[r]):
                if isinstance(value, str) and _norm(value) == want:
                    return r, c
        if required:
            raise DiscoveryError(
                f"{self.source.name!r}: couldn't find the label {label!r} on the "
                f"{self.name!r} sheet - has the export layout changed?"
            )
        return None

    def right_of(self, label: str, start: int = 0, end: int | None = None):
        """First non-blank value to the right of `label` on the same row."""
        r, c = self.find(label, start, end)
        for value in self.rows[r][c + 1:]:
            if value not in (None, ""):
                return value
        return None

    def below(self, label: str, start: int = 0, end: int | None = None):
        """First non-blank value in the same column under `label` (BAGanalytics layout)."""
        r, c = self.find(label, start, end)
        for rr in range(r + 1, min(r + 4, len(self.rows))):
            value = self.cell(rr, c)
            if value not in (None, ""):
                return value
        return None

    def number(self, label: str, how: str = "right", start: int = 0,
               end: int | None = None) -> float:
        raw = (self.right_of if how == "right" else self.below)(label, start, end)
        value = num(raw)
        if value is None:
            raise DiscoveryError(
                f"{self.source.name!r}: {label!r} on the {self.name!r} sheet has no "
                f"numeric value (got {raw!r})"
            )
        return value
