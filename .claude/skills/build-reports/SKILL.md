---
name: build-reports
description: >-
  Generate the weekly Alstef support reports (Part A master ticket report + per-site
  Part B phone-log & BHS reports) from the exports dropped in DropBox/PartA and
  DropBox/PartB. Use when the user says "build the reports", "run the weekly reports",
  or drops new weekly exports and asks for the write-ups.
---

# build-reports

Full spec: `Intructions/Intructions.md` (authoritative); Part B: `docs/PART_B_SPEC.md`.
Design + repo layout:
`~/.claude/plans/my-goal-is-to-goofy-sprout.md`. Step 1 detail:
`~/.claude/plans/for-line-113-in-serene-melody.md`.

Weekly job: drop files → run this skill → review the `.docx` files in `Output/`.

## Step 1 — Determine the reporting week & list inputs  *(implemented)*

Run:

```
python -m reportgen.run --stage discover
```

It reads `DropBox/PartA` + `DropBox/PartB`, resolves the Monday–Sunday reporting week
from the Part A filenames, and writes `work/manifest.json`. Expected output shape:

```
Reporting week:  Monday 24th August to Sunday 30th August 2026
                 (2026-08-24 to 2026-08-30)

PartA:
  phonelogs   Phonelogs - 24.08 to 30.08.2026.xlsx   ("Phone Logs - ..." also accepted)
  carlsource  Carlsource - 24.08 to 30.08.2026.xls

PartB:
  ADL  baganalytics  week ending 2026-08-30   BAGanalytics Report.xlsx
  CHC  glideview      week ending 2026-08-30   Christchurch International Airport - GlideView ...
  ZQN  glideview      week ending 2026-08-30   Queenstown International Airport - GlideView ...
```

**Input filename conventions** (see the step-1 plan for the full grammar):

- `DropBox/PartA/Phonelogs - DD.MM to DD.MM.YYYY.xlsx`
- `DropBox/PartA/Carlsource - DD.MM to DD.MM.YYYY.xls`
- `DropBox/PartB/<Airport full name> - GlideView Reports Week Ending Sunday, <date>.xlsx`
- `DropBox/PartB/<CODE> - BAGanalytics Report.xlsx`  (e.g. `ADL - …`; week read from inside)

**If `--stage discover` exits non-zero:** relay the printed message to the user verbatim
and STOP. Do not continue to step 2 until the files are fixed. Every failure is a
deliberate hard stop (wrong day-of-week, mismatched Part A files, a Part B export for a
different week, an unknown GlideView airport, phone-log dates outside the named week, …).

**On success:** show the user the week + file table and get explicit confirmation before
step 2. If it printed a `!! NOTICE:` (e.g. the Phonelogs/Carlsource files were swapped and
were identified by content instead), relay it to the user too.

## Site / region data  *(implemented)*

`knowledge/sites.yaml` is **generated** from the site lookup workbook + `sites_overlay.yaml`:

```
python -m reportgen.build_sites          # re-run whenever the lookup changes
```

The lookup ("LookUp - Sites and Emails.xlsx") need not be in the repo — `build_sites`
also checks `$ALSTEF_SITES_LOOKUP` and `$OneDrive/Support Reports/`. Do not hand-edit
`sites.yaml`; put overrides in `sites_overlay.yaml`.

## Step 2 — Choose Part B sites/groups  *(implemented)*

Pre-tick every site with a Part B file in the manifest (that's the default when
`--sites` is omitted) and ask the user which other reports they want this week.
Standing individual reports: ADL, AMX, BDA, PER, SFO, ZQN, DUD, CHC. Standing groups:
`"Canadian Sites"` (region Canada) and `StarTrack` (`ST_*` codes) — one document, each
Active member (or any member with entries) under its own subheading. Registry:
`reportgen/partb/sites.py`. If the user names anything else, `--stage partb` stops with
"one-off or standing?" — ask the user, don't guess.

## Step 3 — Parse Part A  *(implemented, Part A only)*

```
python -m reportgen.run --stage parse
```

Parses the phone log + CarlSource export, links calls to tickets, and writes
`work/bundle.json` (regions → sites → entries; `region_counts`, `totals`, `flags`).
Surface the printed flag list to the user. BHS (GlideView / BagAnalytics) is not built yet.

## Step 5 — Build Part A .docx  *(implemented — V1, raw text)*

```
python -m reportgen.run --stage build
```

Renders `Output/Weekly_Support_Tickets_Report_<week>.docx`: letterhead, title, date, the
calls-per-site **native pie chart** right after the date, totals, region grid, region
headings + ticket/direct/call blocks with raw Reported/Findings/Remedy. No AI prose yet
(step 4) and no Part B / BHS yet.

## One-shot demo

`python -m reportgen.demo` (or the packaged `dist/WeeklyReport.exe` — see `build_exe.bat`
and README) chains discover → pause → parse → build and opens `Output/`.

## Step 5b — Build Part B .docx  *(implemented)*

```
python -m reportgen.run --stage partb                       # sites with a BHS export
python -m reportgen.run --stage partb --sites ZQN CHC ADL "Canadian Sites" StarTrack DUD
```

Needs `work/bundle.json` from step 3 (tickets are never re-parsed). Per site with a BHS
export: GlideView files are recalculated on a copy through **Excel** (COM; LibreOffice
fallback) before SUMMARY is read, key totals are cross-checked against the underlying
sheets (a mismatch is a hard stop), and the figures are archived to
`history/<SITE>/<week_start>.json`. Writes
`Output/<SITE-OR-GROUP>_Weekly_Phone_Logs_and_BHS_Statistics.docx`.

Relay to the user: every `!! NOTICE:` line (SUMMARY needed recalculation, export covers
only part of the week, swapped Part A files) and the ADL raw EDS table (L1/L2 stay
"Provisional – withheld" until Tiaan confirms the BAGanalytics level mapping).

## Steps 4, 6  *(not yet implemented)*

4. Claude reads `work/bundle.json` and writes `work/narrative.json` (prose + ambiguous
   links) which `--stage build` will merge.
6. Report what landed in `Output/`.
