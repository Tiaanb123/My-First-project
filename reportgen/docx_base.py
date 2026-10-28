"""Build the Part A document on the Alstef letterhead with python-docx.

`open_report()` clones the letterhead template and strips its placeholder body text; the
`add_*` helpers append report content. The pie chart is spliced in afterwards by
`chart_pie.inject()` at the `__PIE_CHART__` sentinel paragraph that `add_subtitle` lays.
"""
from __future__ import annotations

import datetime as dt
import re

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
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
    _collapse_blank_runs(doc)
    return doc


def _collapse_blank_runs(doc: Document) -> None:
    """The letterhead template leaves several empty paragraphs behind once the
    placeholder body text above is stripped out - keep at most one of each run."""
    prev_blank = False
    for para in list(doc.paragraphs):
        blank = not para.text.strip() and not para._p.findall(".//" + _qn("w:drawing"))
        if blank and prev_blank:
            para._p.getparent().remove(para._p)
        else:
            prev_blank = blank


def _qn(tag: str) -> str:
    from docx.oxml.ns import qn
    return qn(tag)


def _navy(run, size: int | None = None, bold: bool = True):
    run.font.color.rgb = NAVY
    run.font.bold = bold
    if size:
        run.font.size = Pt(size)
    return run


def _tight(paragraph):
    """Word's 'Remove Space After Paragraph', applied to every paragraph we generate -
    visible gaps come only from the explicit `_spacer()` blanks below, not style defaults."""
    paragraph.paragraph_format.space_after = Pt(0)
    return paragraph


def _spacer(doc: Document, loose: bool = False):
    """A blank line between blocks. `loose` leaves the style's own (larger) spacing in
    place for a clear break (e.g. between region headings); otherwise it's tight, giving
    just a single blank line of breathing room (e.g. between calls of a ticket entry)."""
    p = doc.add_paragraph()
    if not loose:
        _tight(p)
    return p


def add_title(doc: Document, text: str):
    p = doc.add_paragraph(style="Title")
    _navy(p.add_run(text), size=20)
    return _tight(p)


def add_subtitle(doc: Document, text: str):
    p = doc.add_paragraph()
    run = p.add_run(text)
    run.font.size = Pt(12)
    run.font.color.rgb = GREY
    _tight(p)
    sentinel = doc.add_paragraph("__PIE_CHART__")
    sentinel.runs[0].font.size = Pt(1)
    return p


def add_kpi(doc: Document, text: str):
    p = doc.add_paragraph(style="Heading 1")
    _navy(p.add_run(text), size=13)
    return _tight(p)


GRID_COLUMNS = 3


def _hide_borders(cell) -> None:
    """White borders, as in the finished example's region grid."""
    borders = OxmlElement("w:tcBorders")
    for side in ("top", "left", "bottom", "right"):
        edge = OxmlElement(f"w:{side}")
        for attr, val in (("val", "single"), ("sz", "4"), ("space", "0"), ("color", "FFFFFF")):
            edge.set(_qn(f"w:{attr}"), val)
        borders.append(edge)
    cell._tc.get_or_add_tcPr().append(borders)


def _distribute_evenly(cell) -> None:
    """A pct-width cell, matching Word's 'Distribute Columns Evenly'."""
    tcPr = cell._tc.get_or_add_tcPr()
    tcW = tcPr.find(_qn("w:tcW"))
    if tcW is None:
        tcW = OxmlElement("w:tcW")
        tcPr.append(tcW)
    tcW.set(_qn("w:w"), str(5000 // GRID_COLUMNS))
    tcW.set(_qn("w:type"), "pct")


def add_region_list(doc: Document, counts: dict, order: list[str]):
    """Region summary: a 3-per-row grid spanning the full page width, one `Region: N`
    cell each, columns distributed evenly."""
    names = [n for n in order if not (n == "Other" and not counts.get("Other"))]
    rows = -(-len(names) // GRID_COLUMNS)
    table = doc.add_table(rows=rows, cols=GRID_COLUMNS)
    try:
        table.style = "Table Grid"
    except KeyError:
        pass
    table.autofit = False
    tblPr = table._tbl.tblPr
    tblW = tblPr.find(_qn("w:tblW"))
    if tblW is None:
        tblW = OxmlElement("w:tblW")
        tblPr.append(tblW)
    tblW.set(_qn("w:w"), "5000")
    tblW.set(_qn("w:type"), "pct")
    for i, name in enumerate(names):
        cell = table.cell(i // GRID_COLUMNS, i % GRID_COLUMNS)
        para = cell.paragraphs[0]
        para.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        # Cell paragraphs have no pStyle, so they'd otherwise inherit the letterhead's
        # Normal style - a 0.5" left/right indent and 18pt space-after meant for letter
        # body text, which ate into the columns and inflated the row gaps in this grid.
        pf = para.paragraph_format
        pf.left_indent = pf.right_indent = 0
        pf.space_before = pf.space_after = Pt(0)
        _navy(para.add_run(f"{name}: {counts.get(name, 0)}"), size=14)
    for row in table.rows:
        for cell in row.cells:
            _hide_borders(cell)
            _distribute_evenly(cell)
    return table


def _call_label(n: int) -> str:
    return "Call" if n == 1 else "Calls"


def add_region_heading(doc: Document, name: str, n: int):
    p = doc.add_paragraph(style="Heading 1")
    _navy(p.add_run(f"{name}: {n} {_call_label(n)}"), size=13)
    return _tight(p)


def add_country_heading(doc: Document, name: str, n: int):
    p = doc.add_paragraph()
    _navy(p.add_run(f"{name}: {n} {_call_label(n)}"), size=12)
    return _tight(p)


def add_site_heading(doc: Document, name: str, code: str):
    p = doc.add_paragraph()
    run = p.add_run(f"{name} {code}".strip())
    run.bold = True
    return _tight(p)


def _fmt_ts(iso: str | None) -> str:
    if not iso:
        return ""
    try:
        return dt.datetime.fromisoformat(iso).strftime("%d/%m/%Y %H:%M")
    except ValueError:
        return iso


def _bullet_num_id(doc: Document) -> int | None:
    """numId of a bullet list defined in the letterhead's numbering part (cached)."""
    cached = getattr(doc, "_bullet_num_id", False)
    if cached is not False:
        return cached
    found = None
    try:
        root = doc.part.numbering_part.element
        bullet_abstracts = {
            a.get(_qn("w:abstractNumId"))
            for a in root.findall(_qn("w:abstractNum"))
            if (lvl := a.find(_qn("w:lvl"))) is not None
            and (fmt := lvl.find(_qn("w:numFmt"))) is not None
            and fmt.get(_qn("w:val")) == "bullet"
        }
        for num in root.findall(_qn("w:num")):
            ref = num.find(_qn("w:abstractNumId"))
            if ref is not None and ref.get(_qn("w:val")) in bullet_abstracts:
                found = int(num.get(_qn("w:numId")))
                break
    except (KeyError, NotImplementedError):
        pass
    doc._bullet_num_id = found
    return found


def _has_style(doc: Document, name: str) -> bool:
    """Look up a style by name without mutating the document - `add_paragraph(style=...)`
    inserts the new (then-orphaned) paragraph *before* raising on an unknown style, which
    left a stray blank paragraph in front of every bullet when the letterhead template
    (which has no "List Bullet" style) hit the except branch below."""
    cached = getattr(doc, "_has_style_cache", None)
    if cached is None:
        cached = doc._has_style_cache = {}
    if name not in cached:
        try:
            doc.styles[name]
            cached[name] = True
        except KeyError:
            cached[name] = False
    return cached[name]


def _bullet(doc: Document, label: str | None, text: str):
    """A real Word bullet (`numPr` -> a bullet list), not just an indented paragraph."""
    if _has_style(doc, "List Bullet"):
        p = doc.add_paragraph(style="List Bullet")
    else:
        p = doc.add_paragraph(style="List Paragraph")
        num_id = _bullet_num_id(doc)
        if num_id is not None:
            num_pr = p._p.get_or_add_pPr().get_or_add_numPr()
            num_pr.get_or_add_ilvl().val = 0
            num_pr.get_or_add_numId().val = num_id
    if label:
        p.add_run(f"{label}: ").bold = True
    p.add_run(text)
    return _tight(p)


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
    _tight(p)

    if kind == "call_only":
        comment = (entry.get("comments") or entry.get("inquiry") or "").strip()
        if comment:
            _bullet(doc, "Call", comment)
        return

    wrote = False
    for field in ("reported", "findings", "remedy"):
        value = (entry.get(field) or "").strip()
        if value:
            _bullet(doc, field.capitalize(), value)
            wrote = True
    if not wrote:
        summary = (entry.get("summary") or "").strip()
        _bullet(doc, "Summary" if summary else None,
                summary or f"No detail recorded in WO {entry['ticket_id']}.")


def add_signoff(doc: Document):
    for line in ("Warm regards,", "ALSTEF Group", "Support Services Team"):
        _tight(doc.add_paragraph(line))
