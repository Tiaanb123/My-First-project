"""Recalculate a workbook's formulas before trusting cached values (spec section 6).

GlideView SUMMARY sheets are formula-driven and the export sometimes ships with every
cached value at 0 (CHC, every week so far). openpyxl can't evaluate formulas, so a copy of
the workbook is recalculated by Excel (COM, via pywin32) - or LibreOffice headless if Excel
isn't available - and the copy is what gets read. The original in DropBox is never touched.
"""
from __future__ import annotations

import dataclasses
import shutil
import subprocess
import tempfile
from pathlib import Path

from ..discover import DiscoveryError, _load_workbook


@dataclasses.dataclass(frozen=True)
class Recalculated:
    path: Path          # the recalculated copy to read with data_only=True
    needed: bool        # the original's SUMMARY cached values were all zero/blank
    engine: str         # "excel" | "libreoffice"


def summary_looks_stale(path: Path, sheet: str = "SUMMARY") -> bool:
    """True when every non-text formula on `sheet` has a cached value of 0 or nothing - the
    CHC symptom. (Text-valued formulas such as the title row are ignored.) A sheet with no
    numeric formulas is never stale."""
    formulas = _load_workbook(path, read_only=False, data_only=False)
    try:
        values = _load_workbook(path, read_only=False)
    except Exception:
        formulas.close()
        raise
    try:
        if sheet not in formulas.sheetnames:
            return False
        f_ws, v_ws = formulas[sheet], values[sheet]
        n_formulas = 0
        for f_row, v_row in zip(f_ws.iter_rows(), v_ws.iter_rows()):
            for f_cell, v_cell in zip(f_row, v_row):
                if isinstance(f_cell.value, str) and f_cell.value.startswith("="):
                    if isinstance(v_cell.value, str) and v_cell.value.strip():
                        continue
                    n_formulas += 1
                    if v_cell.value not in (None, 0, 0.0, ""):
                        return False
        return n_formulas > 0
    finally:
        formulas.close()
        values.close()


def _excel(path: Path) -> None:
    import pythoncom  # noqa: F401  (pywin32; import error -> caller falls back)
    import win32com.client

    pythoncom.CoInitialize()
    app = win32com.client.DispatchEx("Excel.Application")
    try:
        app.Visible = False
        app.DisplayAlerts = False
        app.AskToUpdateLinks = False
        wb = app.Workbooks.Open(str(path.resolve()), UpdateLinks=0)
        try:
            app.CalculateFullRebuild()
            wb.Save()
        finally:
            wb.Close(SaveChanges=False)
    finally:
        app.Quit()
        pythoncom.CoUninitialize()


def _soffice() -> str | None:
    found = shutil.which("soffice")
    if found:
        return found
    for candidate in (r"C:\Program Files\LibreOffice\program\soffice.exe",
                      r"C:\Program Files (x86)\LibreOffice\program\soffice.exe"):
        if Path(candidate).is_file():
            return candidate
    return None


def _libreoffice(path: Path) -> None:
    exe = _soffice()
    if not exe:
        raise FileNotFoundError("soffice not found")
    out_dir = Path(tempfile.mkdtemp(prefix="recalc_"))
    # --convert-to xlsx recalculates on load (RecalcMode from the profile defaults to
    # "always" for files not saved by LibreOffice) and writes a fresh copy.
    subprocess.run(
        [exe, "--headless", "--norestore", "--calc", "--convert-to", "xlsx",
         "--outdir", str(out_dir), str(path)],
        check=True, capture_output=True, timeout=300,
    )
    shutil.move(str(out_dir / path.name), str(path))


def recalculated_copy(src: Path, work_dir: Path) -> Recalculated:
    """Copy `src` into `work_dir/recalc/` and recalculate the copy. Raises DiscoveryError
    (a hard stop) when no engine is available - stale SUMMARY values must never be read."""
    src = Path(src)
    needed = summary_looks_stale(src)
    dest_dir = Path(work_dir) / "recalc"
    dest_dir.mkdir(parents=True, exist_ok=True)
    dest = dest_dir / src.name
    shutil.copy2(src, dest)

    errors: list[str] = []
    for engine, fn in (("excel", _excel), ("libreoffice", _libreoffice)):
        try:
            fn(dest)
        except Exception as exc:  # noqa: BLE001 - try the next engine, report all at the end
            errors.append(f"{engine}: {exc}")
            continue
        if summary_looks_stale(dest):
            raise DiscoveryError(
                f"{src.name!r}: SUMMARY still reads all zeros after recalculating with "
                f"{engine} - check the export before building the report"
            )
        return Recalculated(dest, needed, engine)

    raise DiscoveryError(
        f"{src.name!r}: couldn't recalculate the workbook (needs Excel or LibreOffice) - "
        + "; ".join(errors)
    )
