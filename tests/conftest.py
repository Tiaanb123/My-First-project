from __future__ import annotations

import datetime as dt
import shutil
from pathlib import Path

import openpyxl
import pytest

REPO_ROOT = Path(__file__).resolve().parent.parent


def make_phonelogs(path: Path, call_dates: list[dt.datetime]) -> None:
    """Minimal workbook with the columns reportgen.discover.phone_log_date_span reads."""
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    ws = wb.create_sheet("Phone Stats")
    ws["A1"] = "placeholder"
    logs = wb.create_sheet("Phone Logs")
    logs.append(["Call Number", "Call Type", "Comments", "Operator", "Time Entered",
                 "Location of the caller"])
    for i, when in enumerate(call_dates, start=15000):
        logs.append([i, "Divert", f"E-mailed on {when:%d-%b-%Y} at {when:%H:%M} by Op",
                     "Op", when, "ZQN_APAC"])
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
