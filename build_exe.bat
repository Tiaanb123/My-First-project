@echo off
REM Build the standalone demo exe -> dist\WeeklyReport.exe
cd /d "%~dp0"

if not exist ".venv\Scripts\python.exe" (
  echo No .venv found. Create one and:  .venv\Scripts\pip install -r requirements-dev.txt
  pause & exit /b 1
)

".venv\Scripts\python.exe" -m reportgen.build_sites
".venv\Scripts\python.exe" -m PyInstaller --noconfirm WeeklyReport.spec

echo.
echo Built: dist\WeeklyReport.exe
echo Put it in a folder next to  DropBox\PartA  DropBox\PartB  Output\  and run it.
pause
