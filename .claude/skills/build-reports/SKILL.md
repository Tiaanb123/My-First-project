---
name: build-reports
description: >-
  Generate the weekly Alstef support reports (Part A master ticket report + per-site
  Part B phone-log & BHS reports) from the exports dropped in DropBox/PartA and
  DropBox/PartB. Use when the user says "build the reports", "run the weekly reports",
  or drops new weekly exports and asks for the write-ups.
---

# build-reports

Full spec: `Intructions/Intructions.md` (authoritative). Design + repo layout:
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
  phonelogs   Phonelogs - 24.08 to 30.08.2026.xlsx
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
- `DropBox/PartB/BAGanalytics Report.xlsx`  (Adelaide's BHS source; week read from inside)

**If `--stage discover` exits non-zero:** relay the printed message to the user verbatim
and STOP. Do not continue to step 2 until the files are fixed. Every failure is a
deliberate hard stop (wrong day-of-week, mismatched Part A files, a Part B export for a
different week, an unknown GlideView airport, phone-log dates outside the named week, …).

**On success:** show the user the week + file table and get explicit confirmation before
step 2.

## Step 2 — Choose Part B sites/groups  *(not yet implemented)*

Pre-tick any site with a Part B file present in the manifest. Resolve standing groups
from `knowledge/groups.yaml`. Standing individual Part B reports: Adelaide (ADL), Aramex
(AMX), Bermuda (BDA), Canadian Sites (group), Perth (PER), San Francisco (SFO), StarTrack
(AUSPOS), Queenstown (ZQN), Dunedin (DUD), Christchurch (CHC). If the user names a site
not on that list, ask "one-off or new standing group?".

## Steps 3–6  *(not yet implemented)*

3. `python -m reportgen.run --stage parse` — surface all flags.
4. Claude writes `work/narrative.json` (prose + ambiguous-match decisions).
5. `python -m reportgen.run --stage build`.
6. Report what landed in `Output/`.
