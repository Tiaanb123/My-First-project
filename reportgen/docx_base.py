"""Build the Part A document on the Alstef letterhead with python-docx.

`open_report()` clones the letterhead template and strips its placeholder body text; the
`add_*` helpers append report content. The pie chart is spliced in afterwards by
`chart_pie.inject()` at the `__PIE_CHART__` sentinel paragraph that `add_subtitle` lays.
"""
from __future__ import annotations

import datetime as dt
import re

from docx import Document
from docx.shared import Pt, RGBColor

from . import config

NAVY = RGBColor(0x0B, 0x25, 0x45)
GREY = RGBColor(0x59, 0x59, 0x59)

# Placeholder body text in the letterhead template, whitespace-normalised.
_PLACEHOLDER_TEXT = {
    "heading:", "xnd april", "warm regards,", "alstef group",
    "alstef group support engineering team", "support engineering team",
}


def open_report() -> Document:
    doc = Document(str(config.LETTERHEAD))
    for para in list(doc.paragraphs):
        if para._p.findall(".//" + _qn("w:drawing")):
            continue  # keep the banner
        norm = re.sub(r"\s+", " ", para.text.strip().lower())
        if norm in _PLACEHOLDER_TEXT:
            para._p.getparent().remove(para._p)
    return doc


def _qn(tag: str) -> str:
    from docx.oxml.ns import qn
    return qn(tag)


def _navy(run, size: int | None = None, bold: bool = True):
    run.font.color.rgb = NAVY
    run.font.bold = bold
    if size:
        run.font.size = Pt(size)
    return run


def add_title(doc: Document, text: str):
    p = doc.add_paragraph(style="Title")
    _navy(p.add_run(text), size=20)
    return p


def add_subtitle(doc: Document, text: str):
    p = doc.add_paragraph()
    run = p.add_run(text)
    run.font.size = Pt(12)
    run.font.color.rgb = GREY
    sentinel = doc.add_paragraph("__PIE_CHART__")
    sentinel.runs[0].font.size = Pt(1)
    return p


def add_kpi(doc: Document, text: str):
    p = doc.add_paragraph(style="Heading 1")
    _navy(p.add_run(text), size=13)
    return p


def add_region_list(doc: Document, counts: dict, order: list[str]):
    for name in order:
        if name == "Other" and not counts.get("Other"):
            continue
        doc.add_paragraph(f"{name}: {counts.get(name, 0)}")


def add_region_heading(doc: Document, name: str, n: int):
    label = "ticket" if n == 1 else "tickets"
    p = doc.add_paragraph(style="Heading 1")
    _navy(p.add_run(f"{name}: {n} {label}"), size=13)
    return p


def add_site_heading(doc: Document, name: str, code: str):
    p = doc.add_paragraph()
    run = p.add_run(f"{name} {code}".strip())
    run.bold = True
    return p


def _fmt_ts(iso: str | None) -> str:
    if not iso:
        return ""
    try:
        return dt.datetime.fromisoformat(iso).strftime("%d/%m/%Y %H:%M")
    except ValueError:
        return iso


def _bullet(doc: Document, label: str, text: str):
    try:
        p = doc.add_paragraph(style="List Bullet")
    except KeyError:
        p = doc.add_paragraph(style="List Paragraph")
    p.add_run(f"{label}: ").bold = True
    p.add_run(text)
    return p


def add_entry(doc: Document, entry: dict):
    kind = entry["kind"]
    if kind == "ticket":
        pairs = list(zip(entry.get("call_numbers", []), entry.get("timestamps", [])))
        head = " & ".join(f"{num} ({_fmt_ts(ts)})" for num, ts in pairs) or entry["ticket_id"]
        if entry.get("merged_ticket_ids"):
            head += f"  (tickets {', '.join([entry['ticket_id'], *entry['merged_ticket_ids']])})"
    elif kind == "direct":
        head = f"{entry.get('label', 'Direct Contact')} ({_fmt_ts(entry.get('timestamp'))})"
    else:  # call_only
        head = f"{entry['call_number']} ({_fmt_ts(entry.get('timestamp'))})"

    p = doc.add_paragraph()
    p.add_run(head)

    if kind == "call_only":
        comment = (entry.get("comments") or entry.get("inquiry") or "").strip()
        if comment:
            _bullet(doc, "Call", comment)
        return

    for field in ("reported", "findings", "remedy"):
        value = (entry.get(field) or "").strip()
        if value:
            _bullet(doc, field.capitalize(), value)


def add_signoff(doc: Document):
    for line in ("Warm regards,", "ALSTEF Group", "Support Services Team"):
        doc.add_paragraph(line)
