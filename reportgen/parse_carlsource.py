"""Parse the CarlSource ticket export.

A real `.xlsx` with one `WO Report` worksheet: a couple of "Filter criteria" rows, a
header row, one row per ticket (`Name, WO, Call Reference, Status code, Start date,
Client Code, Client name, Comments`), then a `Lines number` footer row. `Comments` holds
the same `Reported:/Findings:/Remedy:` + dated-freeform-notes text as before.
"""
from __future__ import annotations

import dataclasses
import datetime as dt
import re

from . import textclean
from .discover import DiscoveryError, _load_workbook

_HEADER = ("name", "wo", "call reference", "status code", "start date",
           "client code", "client name", "comments")
# A label is `Reported:` anywhere, or - because some exports drop the colon and run the text
# on (`ReportedAll bags...`, `Remedy` alone on a line) - the bare word at the start of a line.
_RFR_RE = {
    k: re.compile(rf"(?i:\b{k}\s*:)|(?m:^)[ \t]*(?i:{k})(?=[A-Z]|(?m:[ \t]*$))")
    for k in ("reported", "findings", "remedy")
}
_CALL_REF_NUM_RE = re.compile(r"\d{4,}")
_FREEFORM_START_RE = re.compile(
    r"(?im)^\s*(?:\d{1,2}/\d{1,2}\s*:?(?:\s|$)|\d{1,2}(?:st|nd|rd|th)?\s+"
    r"(?:jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*\b)"
)


@dataclasses.dataclass
class Ticket:
    internal_id: str
    title: str
    client_code: str | None
    client_name: str | None
    status: str | None
    begin_date: dt.datetime | None
    call_ref_hint: str | None
    call_reference: str | None
    reported: str
    findings: str
    remedy: str
    freeform: str
    call_refs: list[str] = dataclasses.field(default_factory=list)
    summary: str = ""


def _clean(text: str) -> str:
    """Normalise a cell: NBSP -> space, drop stray replacement chars / leading bullets."""
    text = text.replace("\xa0", " ").replace("�", "")
    text = re.sub(r"(?m)^[ \t]*[•·▪‣◦]\s*", "", text)
    return text.strip()


def _text(value) -> str:
    return _clean(str(value)) if value is not None else ""


def _split_sections(detail: str) -> tuple[str, str, str, str]:
    if not detail.strip():
        return "", "", "", ""
    labels = sorted(
        (m.start(), m.end(), k) for k, r in _RFR_RE.items() for m in r.finditer(detail)
    )
    if not labels:
        return "", "", "", detail.strip()

    ff = _FREEFORM_START_RE.search(detail, labels[0][0])
    rfr_end = ff.start() if ff else len(detail)
    freeform = detail[rfr_end:].strip()

    labels = [lab for lab in labels if lab[0] < rfr_end]
    out = {"reported": "", "findings": "", "remedy": ""}
    seen: set[str] = set()
    for i, (_, seg_start, key) in enumerate(labels):
        if key in seen:
            continue  # a repeated label starts a duplicate block - keep the first
        seen.add(key)
        seg_end = labels[i + 1][0] if i + 1 < len(labels) else rfr_end
        out[key] = detail[seg_start:seg_end].strip()
    return out["reported"], out["findings"], out["remedy"], freeform


def _find_header_row(rows: list[tuple]) -> int:
    for i, row in enumerate(rows):
        cells = tuple(_text(v).lower() for v in row[: len(_HEADER)])
        if cells == _HEADER:
            return i
    raise DiscoveryError(
        "CarlSource export: couldn't find the ticket table header row "
        f"({', '.join(_HEADER)}) on the 'WO Report' sheet"
    )


_BEGIN_FORMATS = ("%d/%m/%Y %H:%M:%S", "%d/%m/%Y %H:%M", "%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M")


def _parse_begin_date(raw, ticket_id: str) -> dt.datetime:
    """`Start date` is `DD/MM/YYYY HH:MM:SS` text in the export, but a cell re-saved from
    Excel can hold a real date (or its ISO text) - accept all of those."""
    if isinstance(raw, dt.datetime):
        return raw
    text = _text(raw)
    for fmt in _BEGIN_FORMATS:
        try:
            return dt.datetime.strptime(text, fmt)
        except ValueError:
            continue
    raise DiscoveryError(
        f"CarlSource ticket {ticket_id!r}: unreadable Start date {text!r} "
        f"(expected DD/MM/YYYY HH:MM:SS)"
    )


def parse(path) -> list[Ticket]:
    wb = _load_workbook(path, read_only=False)
    try:
        if "WO Report" not in wb.sheetnames:
            raise DiscoveryError(f"{path!s}: no 'WO Report' sheet")
        rows = list(wb["WO Report"].iter_rows(values_only=True))
    finally:
        wb.close()

    idx = {name: i for i, name in enumerate(_HEADER)}
    header_row = _find_header_row(rows)

    tickets: list[Ticket] = []
    for row in rows[header_row + 1:]:
        name = _text(row[idx["name"]]) if idx["name"] < len(row) else ""
        if name.lower() == "lines number":
            break  # footer row
        internal_id = _text(row[idx["wo"]]) if idx["wo"] < len(row) else ""
        if not internal_id.isdigit():
            break  # a blank trailing row, or anything else unexpected

        call_reference = _text(row[idx["call reference"]]) or None
        detail = _text(row[idx["comments"]])
        reported, findings, remedy, freeform = _split_sections(detail)
        reported, findings, remedy = (
            "" if textclean.is_placeholder(v) else textclean.limit_text(v, internal_id)
            for v in (reported, findings, remedy)
        )
        summary = ""
        if not (reported or findings or remedy):
            summary = textclean.summarise_freeform(freeform, internal_id)

        tickets.append(Ticket(
            internal_id=internal_id,
            title=_text(row[idx["name"]]),
            client_code=_text(row[idx["client code"]]) or None,
            client_name=_text(row[idx["client name"]]) or None,
            status=_text(row[idx["status code"]]) or None,
            begin_date=_parse_begin_date(row[idx["start date"]], internal_id),
            call_ref_hint=call_reference if call_reference and call_reference.isdigit() else None,
            call_reference=call_reference,
            reported=reported, findings=findings, remedy=remedy, freeform=freeform,
            call_refs=_CALL_REF_NUM_RE.findall(call_reference or ""),
            summary=summary,
        ))
    return tickets
