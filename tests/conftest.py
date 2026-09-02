from __future__ import annotations

import datetime as dt
import shutil
from pathlib import Path

import openpyxl
import pytest

REPO_ROOT = Path(__file__).resolve().parent.parent


def make_phonelogs(path: Path, calls, stats_lines: list[str] | None = None) -> None:
    """Build a phone-log workbook.

    `calls` items are either a `datetime` (location defaults to ZQN_APAC) or a dict with
    keys number / entered / location_id / comments / transferred_to / inquiry.
    """
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    stats = wb.create_sheet("Phone Stats")
    for i, line in enumerate(stats_lines or ["placeholder"], start=1):
        stats.cell(row=i, column=1, value=line)
    logs = wb.create_sheet("Phone Logs")
    logs.append(["Call Number", "Call Type", "Comments", "Operator", "Time Entered",
                 "Location of the caller", "Who did you transfer the call to", "Inquiry"])
    for i, item in enumerate(calls, start=15000):
        if isinstance(item, dt.datetime):
            item = {"entered": item}
        when = item.get("entered")
        logs.append([
            item.get("number", i), "Divert",
            item.get("comments", f"E-mailed on {when:%d-%b-%Y} by Op" if when else ""),
            "Op", when, item.get("location_id", "ZQN_APAC"),
            item.get("transferred_to", "Divert #2"), item.get("inquiry", ""),
        ])
    wb.create_sheet("Data")
    wb.save(path)


def make_glideview(path: Path) -> None:
    openpyxl.Workbook().save(path)


def make_baganalytics(path: Path, start: dt.date, end: dt.date) -> None:
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    ws = wb.create_sheet("Summary")
    ws["A1"] = "BAGview Reports | Summary"
    ws["A2"] = "FILTERS (2)"
    ws["A3"] = (
        f"Date: LAST 7 DAYS: {start.month}/{start.day}/{start.year} - "
        f"{end.month}/{end.day}/{end.year}"
    )
    wb.save(path)


_SS_HEAD = (
    '<?xml version="1.0" encoding="UTF-8"?>\n'
    '<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"'
    ' xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet">'
    '<Worksheet ss:Name="Report"><Table>'
)
_SS_TAIL = "</Table></Worksheet></Workbook>"


def _row(*cells: tuple[int, str]) -> str:
    out = ["<Row>"]
    for idx, text in cells:
        text = (text or "").replace("&", "&amp;").replace("<", "&lt;")
        out.append(f'<Cell ss:Index="{idx}"><Data ss:Type="String">{text}</Data></Cell>')
    out.append("</Row>")
    return "".join(out)


def carlsource_xml(tickets: list[dict]) -> str:
    """`tickets`: dicts with id, title, customer, contract, begin, end, engineer, status,
    reported, findings, remedy, freeform."""
    rows = [_row((1, "INTERVENTION REPORT"))]
    for t in tickets:
        rows.append(_row((1, f"{t['id']} - {t['title']}")))
        rows.append(_row((1, "GENERAL INFORMATION")))
        rows.append(_row((1, "Client/Customer :"), (7, t.get("customer", "")),
                         (8, "Contrat / Contract: "), (12, t.get("contract", ""))))
        rows.append(_row((1, "Date de debut / Begin date : "), (7, t.get("begin", "")),
                         (8, "Date de fin / End date : "), (12, t.get("end", ""))))
        rows.append(_row((1, "Reception appel / Reception call : "), (7, t.get("engineer", ""))))
        rows.append(_row((1, "Etat du ticket / Ticket status : "), (7, t.get("status", ""))))
        rows.append(_row((1, "Temps d'arret / Downtime :  "), (7, "00:00:00")))
        rows.append(_row((1, "DETAIL DE L'INTERVENTION / WORK DETAILS")))
        rows.append(_row((1, "Help:")))
        body = t.get("body")
        if body is None:
            body = (f"Reported: {t.get('reported', '')}\n"
                    f"Findings: {t.get('findings', '')}\n"
                    f"Remedy: {t.get('remedy', '')}\n")
            if t.get("freeform"):
                body += f"24/08:\n{t['freeform']}\n"
        for line in body.split("\n"):
            rows.append(_row((1, line)))
    return _SS_HEAD + "".join(rows) + _SS_TAIL


def make_carlsource(path: Path, tickets: list[dict]) -> None:
    path.write_text(carlsource_xml(tickets), encoding="utf-8")


@pytest.fixture
def dropbox(tmp_path: Path) -> Path:
    """A well-formed DropBox for the week Mon 24 Aug - Sun 30 Aug 2026."""
    root = tmp_path / "DropBox"
    (root / "PartA").mkdir(parents=True)
    (root / "PartB").mkdir(parents=True)

    make_phonelogs(
        root / "PartA" / "Phonelogs - 24.08 to 30.08.2026.xlsx",
        [dt.datetime(2026, 8, 24, 1, 35), dt.datetime(2026, 8, 27, 9, 20),
         dt.datetime(2026, 8, 30, 22, 5)],
    )
    (root / "PartA" / "Carlsource - 24.08 to 30.08.2026.xls").write_text("<Workbook/>", "utf-8")

    shutil.copytree(REPO_ROOT / "knowledge", tmp_path / "knowledge")
    return root
