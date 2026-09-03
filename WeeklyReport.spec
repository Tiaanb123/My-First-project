# PyInstaller spec for the one-shot demo exe.  Build:  pyinstaller WeeklyReport.spec
from PyInstaller.utils.hooks import collect_submodules

a = Analysis(
    ["weeklyreport.py"],
    pathex=["."],
    datas=[
        ("knowledge", "knowledge"),
        ("Intructions/Templates", "templates"),
        ("reportgen/assets/chart_pie", "chart_pie"),
    ],
    hiddenimports=collect_submodules("openpyxl"),
    excludes=["tkinter", "matplotlib", "numpy", "pandas", "PIL"],
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name="WeeklyReport",
    console=True,
    upx=False,
)
