"""The calls-per-site pie chart for Part A.

Rendered as a native Office chart cloned from the finished template (see
`reportgen/assets/chart_pie/`), so it stays editable in Word. `inject()` splices the
chart parts into an already-saved `.docx` at the `__PIE_CHART__` sentinel paragraph.
"""
from __future__ import annotations

import io
import re
import zipfile
from collections import Counter
from pathlib import Path

import openpyxl

from . import config

_TITLE_PLACEHOLDER = "Weekly Call Logs 31/08/2026"

# The chart's original design size (EMU) - its aspect ratio is preserved when the chart
# is resized to fill the page's content width.
_DEFAULT_CX = 5486400
_DEFAULT_CY = 3200400


def _drawing_xml(rid: str, cx: int, cy: int) -> str:
    return (
        '<w:p><w:pPr><w:spacing w:before="120" w:after="120"/></w:pPr>'
        '<w:r><w:rPr><w:noProof/></w:rPr><w:drawing>'
        '<wp:inline distT="0" distB="0" distL="0" distR="0" '
        'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">'
        f'<wp:extent cx="{cx}" cy="{cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>'
        '<wp:docPr id="100" name="Calls by site"/><wp:cNvGraphicFramePr/>'
        '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
        '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">'
        f'<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
        f'r:id="{rid}"/></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>'
    )

_SENTINEL = "__PIE_CHART__"
_SENTINEL_PARA_RE = re.compile(
    r"<w:p\b(?:(?!</w:p>).)*?" + _SENTINEL + r".*?</w:p>", re.S
)

_CT_ADDITIONS = (
    '<Default Extension="xlsx" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"/>'
    '<Override PartName="/word/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>'
    '<Override PartName="/word/charts/style1.xml" ContentType="application/vnd.ms-office.chartstyle+xml"/>'
    '<Override PartName="/word/charts/colors1.xml" ContentType="application/vnd.ms-office.chartcolorstyle+xml"/>'
    '<Override PartName="/word/theme/themeOverride1.xml" ContentType="application/vnd.openxmlformats-officedocument.themeOverride+xml"/>'
)

_CHART_RELS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2011/relationships/chartStyle" Target="style1.xml"/>'
    '<Relationship Id="rId2" Type="http://schemas.microsoft.com/office/2011/relationships/chartColorStyle" Target="colors1.xml"/>'
    '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/themeOverride" Target="../theme/themeOverride1.xml"/>'
    '<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/package" Target="../embeddings/pie.xlsx"/>'
    '</Relationships>'
)


def calls_by_site(bundle: dict) -> list[tuple[str, int]]:
    counts = Counter(
        c["location_id"] or "UNKNOWN" for c in bundle["part_a"].get("phone_calls", [])
    )
    return sorted(counts.items(), key=lambda kv: (-kv[1], kv[0]))


def _esc(text: str) -> str:
    return text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def render_chart_xml(data: list[tuple[str, int]], title: str) -> bytes:
    x = (config.CHART_DIR / "chart1.xml").read_text(encoding="utf-8")
    x = x.replace(_TITLE_PLACEHOLDER, _esc(title))
    x = re.sub(r"<c:dPt>.*?</c:dPt>", "", x, flags=re.S)

    n = len(data)
    cats = "".join(
        f'<c:pt idx="{i}"><c:v>{_esc(code)}</c:v></c:pt>' for i, (code, _) in enumerate(data)
    )
    vals = "".join(
        f'<c:pt idx="{i}"><c:v>{cnt}</c:v></c:pt>' for i, (_, cnt) in enumerate(data)
    )
    new_cat = (
        f'<c:cat><c:strRef><c:f>Sheet1!$A$1:$A${n}</c:f>'
        f'<c:strCache><c:ptCount val="{n}"/>{cats}</c:strCache></c:strRef></c:cat>'
    )
    new_val = (
        f'<c:val><c:numRef><c:f>Sheet1!$B$1:$B${n}</c:f>'
        f'<c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="{n}"/>'
        f'{vals}</c:numCache></c:numRef></c:val>'
    )
    x = re.sub(r"<c:cat>.*?</c:cat>", lambda _m: new_cat, x, flags=re.S)
    x = re.sub(r"<c:val>.*?</c:val>", lambda _m: new_val, x, flags=re.S)
    return x.encode("utf-8")


def embedded_xlsx(data: list[tuple[str, int]]) -> bytes:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    for code, count in data:
        ws.append([code, count])
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def inject(
    docx_path: Path, data: list[tuple[str, int]], title: str, width_emu: int | None = None,
) -> None:
    docx_path = Path(docx_path)
    with zipfile.ZipFile(docx_path) as z:
        parts: dict[str, bytes] = {name: z.read(name) for name in z.namelist()}

    ct = parts["[Content_Types].xml"].decode("utf-8")
    parts["[Content_Types].xml"] = ct.replace("</Types>", _CT_ADDITIONS + "</Types>").encode()

    rels = parts["word/_rels/document.xml.rels"].decode("utf-8")
    used = [int(m) for m in re.findall(r'Id="rId(\d+)"', rels)]
    rid = f"rId{max(used) + 1}"
    rels = rels.replace(
        "</Relationships>",
        f'<Relationship Id="{rid}" '
        f'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" '
        f'Target="charts/chart1.xml"/></Relationships>',
    )
    parts["word/_rels/document.xml.rels"] = rels.encode()

    parts["word/charts/chart1.xml"] = render_chart_xml(data, title)
    parts["word/charts/_rels/chart1.xml.rels"] = _CHART_RELS.encode()
    parts["word/charts/colors1.xml"] = (config.CHART_DIR / "colors1.xml").read_bytes()
    parts["word/charts/style1.xml"] = (config.CHART_DIR / "style1.xml").read_bytes()
    parts["word/theme/themeOverride1.xml"] = (config.CHART_DIR / "themeOverride1.xml").read_bytes()
    parts["word/embeddings/pie.xlsx"] = embedded_xlsx(data)

    cx = width_emu if width_emu else _DEFAULT_CX
    cy = round(cx * _DEFAULT_CY / _DEFAULT_CX)
    doc = parts["word/document.xml"].decode("utf-8")
    drawing = _drawing_xml(rid, cx, cy)
    doc, count = _SENTINEL_PARA_RE.subn(lambda _m: drawing, doc, count=1)
    if count != 1:
        raise RuntimeError(f"pie-chart sentinel {_SENTINEL!r} not found in document.xml")
    parts["word/document.xml"] = doc.encode()

    with zipfile.ZipFile(docx_path, "w", zipfile.ZIP_DEFLATED) as z:
        for name, blob in parts.items():
            z.writestr(name, blob)
