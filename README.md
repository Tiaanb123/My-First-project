# Weekly Support Report automation

Generates the Alstef weekly support reports from the raw exports:

- **Part A** — master *Weekly Support Tickets Report* (all regions/sites).
- **Part B** — per-site *Phone Logs & BHS Statistics* reports.

Authoritative spec: [`Intructions/Intructions.md`](Intructions/Intructions.md); Part B detail:
[`docs/PART_B_SPEC.md`](docs/PART_B_SPEC.md).

## Status

| Stage | State |
|---|---|
| `discover` — reporting week + input manifest | ✅ |
| `parse` — Part A: phone logs + CarlSource + call↔ticket link → `work/bundle.json` | ✅ |
| `build` — Part A `.docx` (letterhead, pie chart, region blocks, **raw** ticket text) | ✅ V1 |
| `partb` — per-site Part B `.docx` (tickets + BHS table + Performance Review, charts) | ✅ ZQN / CHC / ADL extractors, groups |
| narrative (AI prose) | ⬜ |

## Demo (standalone exe)

1. `build_exe.bat` → `dist/WeeklyReport.exe` (needs `.venv` with
   `pip install -r requirements-dev.txt`).
2. Put `WeeklyReport.exe` in a folder alongside:
   ```
   DropBox/PartA/    Phonelogs - DD.MM to DD.MM.YYYY.xlsx
                     Carlsource - DD.MM to DD.MM.YYYY.xls
   DropBox/PartB/    <site> GlideView / BagAnalytics workbooks (optional)
   Output/
   ```
3. Run it. It prints the detected week + files, waits for **Enter**, builds the report,
   and opens `Output/`.

## Dev

```
py -m venv .venv
.venv\Scripts\pip install -r requirements-dev.txt
.venv\Scripts\python -m reportgen.build_sites          # regenerate knowledge/sites.yaml
.venv\Scripts\python -m reportgen.run --stage parse
.venv\Scripts\python -m reportgen.run --stage build
.venv\Scripts\python -m reportgen.run --stage partb [--sites ZQN CHC ADL "Canadian Sites"]
.venv\Scripts\python -m pytest
```

`knowledge/sites.yaml` is generated from `LookUp - Sites and Emails.xlsx` (found via
`--lookup`, `$ALSTEF_SITES_LOOKUP`, `$OneDrive\Support SharedDrive\LookUp Table\`, or the
repo copy) plus `knowledge/sites_overlay.yaml`.
