"""Parse the CarlSource bulk ticket export.

The file claims to be `.xls` but is SpreadsheetML XML (a BIRT export). It is valid UTF-8
and parses with the stdlib; no `lxml` / `xlrd` needed. One `Report` worksheet holds one
block per ticket, delimited by a row `<internal id> - <title>`.
"""
from __future__ import annotations

import dataclasses
import datetime as dt
import re
import xml.etree.ElementTree as ET
from pathlib import Path

_SS = "urn:schemas-microsoft-com:office:spreadsheet"
_NS = {"s": _SS}
_INDEX_ATTR = f"{{{_SS}}}Index"

# Real header: "176973 - TC2 was down…" — 4-6 digit id, spaces around the dash, title
# starts with a non-digit. (Excludes ISO date values like "2026-08-24T03:31:00.000".)
_TICKET_HEADER_RE = re.compile(r"^\s*(\d{4,6})\s+-\s+(\D.*?)\s*$")
_LABELS = {
    "customer": "client",           # "Client/Customer :"
    "contract": "contract",         # "Contrat / Contract:"
    "begin": "begin date",
    "end": "end date",
    "engineer": "reception call",
    "status": "ticket status",
    "downtime": "downtime",
}
_RFR_RE = {k: re.compile(rf"(?i)\b{k}\s*:") for k in ("reported", "findings", "remedy")}
_FREEFORM_START_RE = re.compile(
    r"(?im)^\s*(?:\d{1,2}/\d{1,2}\s*:?(?:\s|$)|\d{1,2}(?:st|nd|rd|th)?\s+"
    r"(?:jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*\b)"
)
_CALL_REF_RE = re.compile(r"call\s*ref\.?\s*#?\s*(\d{3,6})", re.IGNORECASE)


@dataclasses.dataclass
class Ticket:
    internal_id: str
    title: str
    client_customer: str | None
    contract: str | None
    begin_date: dt.datetime | None
    end_date: dt.datetime | None
    engineer: str | None
    status: str | None
    downtime: str | None
    reported: str
    findings: str
    remedy: str
    freeform: str
    call_ref_hint: str | None


def _clean(text: str) -> str:
    """Normalise a cell: NBSP -> space, drop stray replacement chars / leading bullets."""
    text = text.replace("\xa0", " ").replace("�", "")
    text = re.sub(r"(?m)^[ \t]*[•·▪‣◦]\s*", "", text)
    return text.strip()


def _row_cells(row) -> list[tuple[int, str]]:
    """(column index, text) for each cell, honouring ss:Index gaps."""
    out: list[tuple[int, str]] = []
    col = 0
    for cell in row.findall("s:Cell", _NS):
        if cell.get(_INDEX_ATTR):
            col = int(cell.get(_INDEX_ATTR))
        else:
            col += 1
        data = cell.find("s:Data", _NS)
        text = _clean("".join(data.itertext())) if data is not None else ""
        out.append((col, text))
    return out


def _iso(value: str) -> dt.datetime | None:
    value = value.strip()
    if not value:
        return None
    try:
        return dt.datetime.fromisoformat(value)
    except ValueError:
        return None


def _label_value(cells: list[tuple[int, str]], fragment: str) -> str | None:
    for pos, (_, text) in enumerate(cells):
        if fragment in text.lower():
            for _, later in cells[pos + 1:]:
                if later.strip():
                    return later.strip()
            return None
    return None


def _split_sections(detail: str) -> tuple[str, str, str, str]:
    if not detail.strip():
        return "", "", "", ""
    found = {k: m.start() for k, r in _RFR_RE.items() if (m := r.search(detail))}
    if not found:
        return "", "", "", detail.strip()

    first_label = min(found.values())
    ff = _FREEFORM_START_RE.search(detail, first_label)
    rfr_end = ff.start() if ff else len(detail)
    freeform = detail[rfr_end:].strip()

    ordered = sorted((p, k) for k, p in found.items() if p < rfr_end)
    out = {"reported": "", "findings": "", "remedy": ""}
    for i, (pos, key) in enumerate(ordered):
        seg_start = _RFR_RE[key].search(detail, pos).end()
        seg_end = ordered[i + 1][0] if i + 1 < len(ordered) else rfr_end
        out[key] = detail[seg_start:seg_end].strip()
    return out["reported"], out["findings"], out["remedy"], freeform


def _finish(block: list[list[tuple[int, str]]]) -> Ticket | None:
    header = None
    for cells in block:
        for _, text in cells:
            m = _TICKET_HEADER_RE.match(text)
            if m and not text.lower().startswith("date"):
                header = m
                break
        if header:
            break
    if not header:
        return None

    detail_lines: list[str] = []
    in_detail = False
    fields: dict[str, str | None] = {}
    for cells in block:
        joined_low = " ".join(t.lower() for _, t in cells)
        if "work details" in joined_low:
            in_detail = True
            continue
        if in_detail:
            for _, text in cells:
                if text.strip():
                    detail_lines.append(text.strip())
            continue
        for key, fragment in _LABELS.items():
            if key not in fields:
                value = _label_value(cells, fragment)
                if value:
                    fields[key] = value

    detail = "\n".join(detail_lines)
    reported, findings, remedy, freeform = _split_sections(detail)
    ref = _CALL_REF_RE.search(detail)

    return Ticket(
        internal_id=header.group(1),
        title=header.group(2),
        client_customer=fields.get("customer"),
        contract=fields.get("contract"),
        begin_date=_iso(fields.get("begin") or ""),
        end_date=_iso(fields.get("end") or ""),
        engineer=fields.get("engineer"),
        status=fields.get("status"),
        downtime=fields.get("downtime"),
        reported=reported,
        findings=findings,
        remedy=remedy,
        freeform=freeform,
        call_ref_hint=ref.group(1) if ref else None,
    )


def parse(path) -> list[Ticket]:
    root = ET.fromstring(Path(path).read_bytes())
    table = root.find(".//s:Worksheet/s:Table", _NS)
    if table is None:
        return []

    tickets: list[Ticket] = []
    block: list[list[tuple[int, str]]] = []
    for row in table.findall("s:Row", _NS):
        cells = _row_cells(row)
        first = next((t for _, t in cells if t.strip()), "")
        if _TICKET_HEADER_RE.match(first) and not first.lower().startswith("date"):
            if block:
                t = _finish(block)
                if t:
                    tickets.append(t)
            block = [cells]
        else:
            block.append(cells)
    if block:
        t = _finish(block)
        if t:
            tickets.append(t)
    return tickets
