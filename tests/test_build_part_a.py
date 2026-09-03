from __future__ import annotations

import json
import zipfile
from pathlib import Path

import pytest
from docx import Document

from reportgen import build_part_a, chart_pie

REPO_ROOT = Path(__file__).resolve().parent.parent
SAMPLE_BUNDLE = REPO_ROOT / "work" / "bundle.json"

pytestmark = pytest.mark.skipif(
    not SAMPLE_BUNDLE.is_file(),
    reason="run `python -m reportgen.run --stage parse` first",
)


@pytest.fixture(scope="module")
def report(tmp_path_factory) -> Path:
    out_dir = tmp_path_factory.mktemp("out")
    return build_part_a.build(SAMPLE_BUNDLE, out_dir)


def test_calls_by_site_sorted_desc():
    bundle = json.loads(SAMPLE_BUNDLE.read_text(encoding="utf-8"))
    data = chart_pie.calls_by_site(bundle)
    assert data[0] == ("YLW_APAC", 5)
    assert sum(n for _, n in data) == bundle["part_a"]["totals"]["calls"]
    assert data == sorted(data, key=lambda kv: (-kv[1], kv[0]))


def test_report_filename(report):
    assert report.name == "Weekly_Support_Tickets_Report_24.08.2026_to_30.08.2026.docx"


def test_report_body(report):
    doc = Document(str(report))
    texts = [p.text.strip() for p in doc.paragraphs]
    joined = "\n".join(texts)

    assert "Weekly Support Tickets Report:" in texts
    assert "Monday 24th August to Sunday 30th August 2026" in texts
    assert "Total Amount of Calls: 23" in texts
    assert "Total Amount of Direct Contacts: 2" in texts

    headings = [t for t in texts if t.endswith(" ticket") or t.endswith(" tickets")]
    assert headings == [
        "Africa: 0 tickets", "Australia: 1 ticket", "Bermuda: 2 tickets",
        "Canada: 10 tickets", "Jamaica: 0 tickets", "New Zealand: 8 tickets",
        "United States: 4 tickets", "Parcels: 0 tickets",
    ]
    assert "Kelowna Airport YLW" in texts
    assert "Reported:" in joined and "Direct Email" in joined
    assert texts[-3:] == ["Warm regards,", "ALSTEF Group", "Support Services Team"]


def test_report_has_native_chart(report):
    with zipfile.ZipFile(report) as z:
        names = z.namelist()
        assert "word/charts/chart1.xml" in names
        assert "word/embeddings/pie.xlsx" in names
        ct = z.read("[Content_Types].xml").decode("utf-8")
        assert "drawingml.chart+xml" in ct and 'Extension="xlsx"' in ct
        doc = z.read("word/document.xml").decode("utf-8")
        assert "<c:chart" in doc and "__PIE_CHART__" not in doc
        chart = z.read("word/charts/chart1.xml").decode("utf-8")
        assert "<c:v>YLW_APAC</c:v>" in chart
        assert "<c:dPt>" not in chart
        assert "Weekly Call Logs 30/08/2026" in chart
