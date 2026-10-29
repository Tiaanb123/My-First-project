"""Path resolution that works both from a source checkout and a PyInstaller `--onefile` exe.

Static assets (site data, letterhead + chart templates) are bundled into the exe and read
from `sys._MEIPASS`; the working folders (`DropBox/`, `work/`, `Output/`) live next to the
exe so the user drops files beside it.
"""
from __future__ import annotations

import sys
from pathlib import Path

FROZEN = getattr(sys, "frozen", False)


def _assets() -> Path:
    if FROZEN:
        return Path(sys._MEIPASS)  # type: ignore[attr-defined]
    return Path(__file__).resolve().parent.parent


def _data() -> Path:
    if FROZEN:
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent.parent


REPO_ROOT = Path(__file__).resolve().parent.parent

DROPBOX = _data() / "DropBox"
WORK = _data() / "work"
OUTPUT = _data() / "Output"
# Part B: each week's BhsStats per site, kept so week-over-week trends can return later
HISTORY = _data() / "history"

KNOWLEDGE = _assets() / "knowledge"
TEMPLATES = _assets() / ("templates" if FROZEN else "Intructions/Templates")
CHART_DIR = _assets() / ("chart_pie" if FROZEN else "reportgen/assets/chart_pie")

LETTERHEAD = TEMPLATES / "ALSTEF Group Letterhead - Template.docx"
