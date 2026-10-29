# CLAUDE.md

Weekly Alstef support-report automation (Part A master ticket report + per-site Part B
phone-log & BHS reports). Python 3.12 in `.venv`; package `reportgen/`.

- Part A spec: `Intructions/Intructions.md`. Part B spec: `docs/PART_B_SPEC.md` — read it
  before touching anything under `reportgen/partb/`.
- Never reconstruct the letterhead banner in code — always clone it from
  `Intructions/Templates/ALSTEF Group Letterhead - Template.docx` (`docx_base.open_report`).
- Never fabricate trends or figures. A site with no BHS upload gets no BHS section; a chart
  that isn't backed by real data is omitted with a note saying why.
- Fail fast on bad inputs: raise with a plain-English message naming exactly what's wrong
  rather than guessing or silently adjusting.
- NZ English throughout (labelled, organised, colour, …).
- Run tests: `.venv\Scripts\python -m pytest`.
