# Part B — Site Weekly Phone Logs Report & BHS Statistics
## Full build spec for a standalone script (Claude Code / Python)

Purpose: everything the Claude Project does today for **Part B**, written down so it can be rebuilt as a
local, repeatable script with no Claude Pro dependency. Sources: the Project Instructions, the Project
memory file `reporting-method.md` (which takes precedence where it disagrees with the written
instructions), and the three Part B reports actually built for 5–11 Oct 2026 (ZQN, CHC, ADL).

Where something is an **inference** rather than a stated rule, it is marked *(inferred – verify against
the golden values in §12)*.

> **Repo decisions (Oct 2026 build)** — where the implementation deliberately differs from the text below:
> - Golden tests use the **19–25 Oct 2026** exports (the 5–11 Oct files aren't on hand); §12's values remain
>   as the historical reference.
> - Recalculation (§6) uses **Excel via COM** (pywin32) first, LibreOffice headless only as a fallback.
> - ADL L1/L2 stays **"Provisional – withheld"** even though later exports carry Level 1/2 counts; the run
>   prints the raw Level 1/2/3 table so the mapping can be confirmed before it's published.
> - Ticket meta lines keep the hand-made Part A format (`15292 (10/10/2026 06:43)`) — open item §14.4 resolved.

---

## 0. How to use this with Claude Code

1. Put this file in the repo root as `docs/PART_B_SPEC.md` and add a `CLAUDE.md` containing:
   *"Read docs/PART_B_SPEC.md before touching Part B code. Never reconstruct the letterhead in code.
   Never fabricate trends. Fail fast on bad inputs. NZ English."*
2. If you already have the `reportgen` package (parse_carlsource, parse_phonelogs, link, build_part_a,
   bundle, regions, docx_base…), **Part B should reuse it**: Part B = (Part A's grouped entries filtered to
   one site) + (BHS extraction) + (renderer). Do not re-parse tickets in a second place.
3. Build order: §13. Verify each stage against §12 before moving on.

---

## 1. What Part B is

For each site named that week, one document titled
**"[SITE / SITE GROUP] Weekly Phone Logs Report & BHS Statistics"**, containing:

1. Letterhead banner → title → date subtitle
2. Per-site **Support Tickets** section (filtered Part A data for that site)
3. **Weekly BHS Statistics** table (only if GlideView / BAGanalytics data supplied for that site)
4. Page break → letterhead banner again → **BHS Performance Review** (only if BHS data supplied)
5. Sign-off

**Standing individual reports** (one file each): ADL (Adelaide), AMX (Aramex), BDA (Bermuda), PER (Perth),
SFO (San Francisco), ZQN (Queenstown), DUD (Dunedin), CHC (Christchurch).
**Standing group reports** (one combined file, each site under its own subheading): *Canadian Sites*
(~19 supported Canadian sites) and *StarTrack* (the `ST_*` parcels sites).
The old combined "South Island" report is **retired** (DUD/ZQN/CHC are separate).
A site not on the list → ask whether it's one-off or should be added as standing; don't guess.

Output filename: `[SITE-OR-GROUP]_Weekly_Phone_Logs_and_BHS_Statistics.docx`
(e.g. `ZQN_Weekly_Phone_Logs_and_BHS_Statistics.docx`).

---

## 2. Inputs and validation (fail fast, never guess)

| Input | Name pattern | Notes |
|---|---|---|
| Phone log workbook | `Phonelogs - DD.MM to DD.MM.YYYY.xlsx` | Sheets: `Phone Logs` (Call Number, Time Entered, Location of the Caller, Who did you transfer the call to, Operator, Comments, Inquiry), optional `Phone Stats`, `Data`, `Charts` |
| CarlSource export | `Carlsource - DD.MM to DD.MM.YYYY.xls(x)` | Sheet `WO Report`: Name, WO, Call Reference, Status code, Start date, Client Code, Client name, Comments |
| GlideView (ZQN, CHC) | `<Airport full name> - GlideView Reports Week Ending Sunday, <date>.xlsx` | see §5 |
| BAGanalytics (ADL) | `<CODE> - BAGanalytics Report.xlsx` (week lives inside Summary sheet: `Date: LAST 7 DAYS: M/D/YYYY - M/D/YYYY`) | see §5 |
| Lookup | `LookUp - Sites and Emails.xlsx`, sheet `Site Names` | authoritative CS-client → site code, full name, region, contact email |

Checks, in order — abort with a plain-English message naming exactly what's wrong:

1. Exactly one phone log + one CarlSource file; both name the **identical Monday→Sunday** week (start is a
   Monday, end is a Sunday, exactly 7 days).
2. **Detect by content, not filename.** On 5–11 Oct the two files were swapped (the "Phonelogs" file held
   `WO Report`, the "Carlsource" file held `Phone Logs`). Identify by sheet names, **report the swap to the
   user**, then proceed using the content. (The project rule is "tell me, don't silently adjust"; the swap
   was unambiguous so processing continued after flagging.)
3. Every call's `Time Entered` is inside the named week (a single outlier = mislabelled file).
4. Every GlideView/BAG file resolves to a site the user named, and no two files share a site.
5. Each BHS file's week-ending Sunday equals the Part A week. **If it doesn't, flag it — never blend.**
   - ZQN's 5–11 Oct export was **single-day (Sun 11 Oct only)**: `System Throughput Rates` LocalTime ran
     2026-10-11 00:05 → 2026-10-12 00:00. Detect this from the time-series range (and from the title row
     `"... Sunday, 11 October 2026"` vs `"Week Ending"`) and label every figure as single-day.
6. Recalculate every GlideView workbook before trusting SUMMARY (§6).

---

## 3. Support-ticket section (shared with Part A)

Part B does **not** have its own ticket logic. It takes Part A's resolved entries and filters to the site.

Entry kinds, chronological within a site:

| Kind | Meta line (grey, bold-italic, 10pt) | Body |
|---|---|---|
| Ticket matched to call(s) | `WO 180537 — Call(s) 15292 (10/10/2026 06:43)`; multiple: `Call(s) 15287, 15288 (08/10/2026 06:08)` | bullets: Reported / Findings / Remedy |
| Ticket, no call | `WO 180539 — Direct Email (09/10/2026 17:09)` or `Direct Contact` | bullets R/F/R |
| Call matching no ticket | `Call 15281 (05/10/2026 10:18)` | single bullet `Call: <comment or inquiry>` |

Rules (all inherited from Part A):
- Match ticket↔call: same region; call entered from ~15 min **before** to ~2 h **after** ticket start.
  **But first** use CarlSource `Call Reference`: extract every 5-digit number with regex `\d{5}` (delimiters are
  inconsistent — commas, periods, "&"), look each up directly in the phone log, mark as *consumed*. Only
  un-consumed calls fall back to the time window. (This was the headline bug in an earlier script — see
  `Part_A_Bug_Report_and_Build_Spec.md`.)
- Leftover call within 24 h of a same-site ticket that has no call of its own → grouped under that ticket.
- Two tickets for the same site linked to the exact same call(s) = one incident → merge (union R/F/R text,
  list extra WO numbers).
- Split `Comments` on `Reported:` / `Findings:` / `Remedy:`; **always bullets**; drop stray bullet characters
  and `\xa0`; replace `_x000D_` with a space. Omit a bullet whose section is empty (don't print an empty
  Findings). Trailing dated follow-ups (`07/09:`, `xx/09:`…) are optional supplementary text — trim or
  summarise; **never paste code/JSON/config dumps** (one ticket leaked ~23 occurrences of `autoInduct`
  config JSON into a client report). Heuristic: reject follow-up text with high density of `{ } " :` or long
  unbroken tokens or `ssh`.
  Date-marker regex that worked: `(?:\d{1,2}|xx)/(?:\d{1,2}|xx)(?:/\d{4})?\s*(?:[:\-–]|\n)`.
- Count **tickets** (distinct WO), not printed blocks.
- Site line: **bold** `Site name + code` (e.g. `Adelaide Airport ADL`), always bold.
- No entries for the site → exactly: `No support tickets were logged for [SITE] during this period.`
  (italic grey). This happened for CHC on 5–11 Oct even though CHC BHS data existed.

Site-block heading in Part B: `"[Site full name] [CODE] — Support Tickets"` (Heading1-level).

---

## 4. Site registry and per-site conventions

| Site | Report type | BHS source | Prepared for | Notes |
|---|---|---|---|---|
| ZQN | individual | GlideView | Sean Whittaker, Queenstown Airport (CC: Blair Poland, National Support Services Manager, Alstef Group) | exclude `Dieback Status` |
| CHC | individual | GlideView | Blair Poland, National Support Services Manager | PM-testing note, 4 EDS lines |
| ADL | individual | BAGanalytics | *(none known → omit "Prepared for"; ask the user once)* | reject figures provisional |
| DUD, PER, SFO, BDA, AMX | individual | none uploaded so far | unknown → omit line, ask once | tickets only; no BHS section |
| Canadian Sites | group | none so far | — | one subheading per site |
| StarTrack | group | none so far | — | `ST_*` codes, Parcels/Australia |

General: "Prepared by: Tiaan Botha, Alstef Group" is included. If a site's facilities contact is not known,
**omit** "Prepared for" (never invent). The lookup workbook has contact emails; names above come from project
notes.

New site appearing with BHS data for the first time → work out its convention from the export structure
(same logic as below), **state the assumption in the report and add it to the registry**.

---

## 5. Extraction recipes (exact sheet/cell layouts seen on 5–11 Oct 2026)

Always load with `openpyxl.load_workbook(path, data_only=True)` **after recalculation** (§6).

### 5.1 ZQN — GlideView (`Queenstown International Airport - GlideView Reports …xlsx`)
Sheets: Throughput Statistics, Maintenance Statistics, Security Screening Statistics, Tracking Statistics,
Plough Statistics, Fault Table, Downtime, System Throughput Rates, SUMMARY, Throughput Trends.

- **SUMMARY** is a label→value list (col A label, col B value). Read by label: `T101`, `T114`, `OS104`,
  `L1 Clear Count`, `L1 Reject Bag Count`, `L1 No Scan Count`, `L1 Timeout Count`, `L2 Clear Count`,
  `L2 Reject Bag Count`, `L2 Timeout Count`; tracking block: header `Zone | Tracked Bag Count | Lost Bag Count`,
  rows `T109PC, T110PC, T111PC, T112PC, T1DV1PD, T1DV1PU, T110APC`, then `TOTAL` (lost).
- Main-line throughput = **T101** (T114 secondary tag, quote it). Oversize = **OS104**.
- L1 flagged = `L1 Reject Bag Count`; denominator for L1 % = L1 Clear + L1 Reject + L1 No Scan (+Timeout)
  = bags screened at L1 (5–11 Oct: 2231+1208+12 = **3,451**; 1208/3451 = 35.0%).
- L2 confirmed reject = `L2 Reject Bag Count`; % of **main-line throughput** (90/3543 = 2.54%).
- Lost bags = TOTAL lost (flat sum across the 7 checkpoints); % of **tracked scans** = lost ÷ Σ(Tracked Bag Count)
  (14 / 20,515 = 0.07%).
- **Fault Table**: row 4 = headers (col B onward = fault names), data from row 5 (col A = device), final row
  `TOTAL`. Sum each column over device rows (skip None rows and any row whose col A contains "total",
  case-insensitive), sort descending, take top ~10. **Exclude `Dieback Status`** (heartbeat counter, 617 hits
  that week) and say so in a footnote.
- **Trend**: `System Throughput Rates` row 4 headers (`UTCTime`, `LocalTime`, then tags), 5-minute rows from
  row 5. Daily total = sum of the **`T101`** column grouped by `LocalTime.date()`. Only chart if the sheet
  spans ≥2 days of the reporting week; if single-day, **no chart** + explicit note.

### 5.2 CHC — GlideView (`Christchurch International Airport - GlideView Reports Week Ending …xlsx`)
Sheets: Throughput Statistics, Maintenance Statistics, Diverter Statistics, Security Screening Statistics,
Automatic Tag Reader Statistics, Tracking Statistics, Fault Table, Downtime, Routing Results, Energy Usage,
Intrusion Statistics, SUMMARY.

**SUMMARY cached values were all 0 until recalculated** (flag this to the user when it happens).
After recalculation (col A label, col B value; row numbers as seen, search by label instead):
- No single main-line tag. **Main-line throughput = Σ `Bags Screened` for the four EDS lines**
  `X1X0100 EDS, X1X0200 EDS, X1X0300 EDS, X1X0400 EDS` (rows 65/73/81/89 → 12,495 / 13,255 / 12,360 / 13,220;
  `SCREENING TOTAL` row 96 = **51,330** — cross-check).
- Secondary cross-check: check-in-side total `SORT SYSTEM` (row 13) = C1+C2+C3+Tranship = **49,450**.
  Will not equal the EDS aggregate (re-screening) — **expected, say so, not an error**.
- Oversize = `OOG Total` (row 19) = 4,156. (`OVERALL TOTAL` row 20 = 53,606 is bags in incl. OOG.)
- Per-EDS rows under each header: `Bags Screened, Level 1 Clear, Level 1 Clear %, No Scans, Level 2 Reject,
  Level 2 Timeout`.
- **L1 flagged** = screened − L1 Clear − No Scans (summed over 4 lines) *(inferred – verify)*:
  51,330 − 31,059 − 89 = **20,182 (39.3%)**.
- **L2 confirmed reject** = Σ `Level 2 Reject` = 251+255+233+286 = **1,025 (2.00% of 51,330)**.
  **L2 timeouts** Σ = 245 — report separately ("a timeout is not a confirmed reject").
- **Lost bags = the 8-zone `TRACKING STATS` aggregate** (rows 125–133, `TOTAL` = **2,038**, and row 134
  `Percentage*` = **4.12%** of bags entering the sort system). Zones: TC1 & TC2 prior to ME, Domestic CBS,
  International CBS, Exiting ME, Between ME and L4T01, S1 Sort Line, S2 Sort Line, Level 4 Line. **Not** a
  flat checkpoint sum.
- **Fault Table**: row 4 headers, row 5+ data, 128 columns; same summing rule. Top ~10 categories.
  **Mandatory CHC note**: PM-testing drives some counts — `Emergency Stop Fault`, `Emergency Stop Relay Fault`,
  `Motor Isolator Fault`, `Manual Mode Active`: give count and % of total faults shown
  (5–11 Oct: 177+54+50+4 = 285 of 1,166 = 24.4%) so the chart isn't read as all unplanned faults.
- No 5-minute time-series sheet → **no trend chart**; note it (unlike ZQN's `System Throughput Rates`).

### 5.3 ADL — BAGanalytics (`ADL - BAGanalytics Report.xlsx`)
Sheets: Summary, System Throughput, Maintenance Statistics, Diverter Statistics, `EDS Statistics ` (note
trailing space), ATR Statistics, Tracking Statistics (Sortation…), Tracking Statistics (All), Faults List,
Faults List (With Timestamp), Downtime, … 15-min sheets.
**Numbers are stored as text with thousands separators** (`'63,562'`, `'0.00'`) → strip commas, convert.

- Week: Summary row 4 `Date: LAST 7 DAYS: 10/5/2026 - 10/11/2026` (US M/D/YYYY!) — parse and compare to Part A.
- Summary label→value (labels in col D/B; value on the **following row** or in far-right cols — search by
  label): `Total Check-in Bags (In Gauge)` = **63,562** (main line), `Total Check-in Bags (Oversize)` = **3,010**,
  `Total Recircs plus Check-in Bags` 68,747, `Total Make-up Output` 63,335, `Total Bags Processed` 66,572.
- Lost bags: Summary `Lost Track Percentage` (**3.10%**), `Tracking Rate` (**96.90%**), `Total Lost Bags`
  (**2,128**) each on the row after the label (D137, D140, D143). Zone table below (`Zone | Lost Count`,
  D147…): SL1 935, XT1 331, XT3 324, … — use for the lost-bag chart.
- **Faults List**: from row 10, alarm name col D (leading spaces — strip), count col E; aggregate by name across
  areas. **Exclude `Heartbeat Fault`** (1,132 hits) as a heartbeat counter, with footnote.
- **EDS Statistics** (Level 3 Pass/Fail/Unknown/Error per EDS OS1X1, OS2X1, XT1X1, XT2X1, XT3X1) returned
  **all zeros** on 5–11 Oct although 63,562 bags were screened. Rule: if every EDS count is zero while
  throughput is large, **withhold L1/L2 figures** and flag a data/export problem — do **not** publish 0%.
  ADL reject figures stay "provisional pending L1/L2 reconciliation".
- Trend chart withheld: the "authoritative daily throughput column" for ADL is unconfirmed. Say so.

---

## 6. Recalculation (mandatory)

GlideView SUMMARY formulas can carry stale/zero cached values (CHC: every total read 0; 2,392 formulas).
Before reading: recalculate with LibreOffice headless — the working approach is a Basic macro
`ThisComponent.calculateAll(): ThisComponent.store(): ThisComponent.close(True)` run via
`soffice --headless` against a **copy** of the workbook (the Claude skill's `recalc.py` does this, ~450 lines;
re-implement minimally). Then reload with `data_only=True`.
If a file needed this, say so in the run output. **Belt and braces:** also recompute the key totals from the
underlying sheets (e.g. Σ per-EDS Bags Screened vs `SCREENING TOTAL`; Σ tracking zones vs `TOTAL`) and abort
on disagreement. Always cross-check totals against SUMMARY; if they differ, surface it — don't pick one.

---

## 7. The BHS Statistics table (inside the site's section)

Heading: `"[Site name] — Weekly BHS Statistics"`, then a one-line italic source line
(`Monday 5 - Sunday 11 October 2026, sourced from GlideView.` / `…BAGanalytics.`; for single-day data
`Sunday 11 October 2026 only (see data scope note above) — sourced from GlideView.`).
Columns **Metric | Value | Note**. Fixed rows, in order:

| Metric | Value format | Note column |
|---|---|---|
| Total throughput (main line) | `3,543 bags (T101)` / `51,330 bags (X1X0100-X1X0400 EDS)` / `63,562 bags` | ZQN: "Excludes oversize; T114 secondary tag read 3,554". CHC: "Sum of 4 parallel EDS lines, excl. oversize". ADL: "Total Check-in Bags (In Gauge), excl. oversize" |
| *(CHC only)* Check-in-side total (cross-check) | `49,450 bags (C1+C2+C3+Tranship)` | "Secondary figure - will not match EDS total exactly due to re-screening, expected" |
| Oversize line volume | `309 bags (OS104)` / `4,156 bags (OOG Total)` / `3,010 bags` | "Separate system" |
| L1 flagged for re-screen | `1,208 bags (35.0%)` | "Sent to L2, not a final reject" |
| L2 confirmed reject (final) | `90 bags (2.54% of throughput)` | "The operational reject figure" (+ CHC: "245 further L2 timeouts not counted as confirmed rejects") |
| Lost-bag events | `14 (0.07% of tracked scans)` / `2,038 (4.12% of bags entering sort system)` / `2,128 (3.10% of tracked scans)` | ZQN "Flat sum across ZQN's checkpoint list"; CHC "8-zone Tracking Stats SUMMARY aggregate, not a flat checkpoint sum" |

Omit the whole BHS section (table + Performance Review) for a site with no BHS upload — never invent figures.
ADL replaces the L1/L2 row with `L1 / L2 reject figures | Provisional - withheld | <why>` and adds a paragraph
under the table about the two standing limitations (EDS zero-data; no confirmed trend column).
Always include the **lost-bag comparability caveat** somewhere in the review: percentages are not directly
comparable between sites because the exports aggregate differently.

---

## 8. BHS Performance Review (new page, banner repeated)

Start with a page break, the **letterhead banner again**, then:

`"[Site name] — BHS Performance Review"` (Heading1) then three plain lines:
- *Period covered:* `Monday 5 October - Sunday 11 October 2026 (reporting week)` — or, single-day:
  `Sunday 11 October 2026 only (single-day GlideView export — see data scope note on page 1)`. (Italic.)
- `Prepared for: …` (only if known)
- `Prepared by: Tiaan Botha, Alstef Group`

Then sub-headings (Heading2) in this order:

1. **Executive Summary** — exactly **5 bullets**: (a) main-line volume + oversize, (b) L2 final reject count/%,
   (c) L1 flag share and explanation, (d) tracking/lost-bag reliability, (e) one data-quality/scope caveat
   (or the top fault finding). Plain-English, numbers inline.
2. **1. Total Throughput** — paragraph; secondary tags/cross-checks; trend chart **only if genuine** (§5), else
   a sentence saying why no chart.
3. **2. Reject Bags** — explain L1 vs L2 clearly: *L1 is an automated first-pass flag, NOT a final reject; L2
   (thorough second screening) is the figure that reflects actual operational reject volume.* Report timeouts
   separately where available.
4. **3. Fault Subsections** — one paragraph (what was excluded and why, PM-testing note for CHC) + **bar chart**
   of top ~10 fault categories (summed over all devices).
5. **Lost-bag Events** — paragraph + **bar chart** by whatever granularity that site's export has (ZQN: flat
   7-checkpoint list; CHC: 8 system zones; ADL: tracking zones SL1/XT1/XT3…). Don't force one site's zones on another.
6. **Observations & Discussion Points** — 3–4 bullets (largest lost-bag zone/share, largest genuine fault
   categories, link to any ticket that week, open data-quality items).
7. **Suggested Next Steps** — 2–3 bullets (e.g. request full-week export; confirm ADL EDS config; reintroduce
   week-over-week trend + "most recent full week" snapshot **once several consecutive weeks are on file**).

Dropped for now (do **not** build): "Comparative Percentage/trend" week-over-week and "Snapshot: most recent
full week". No monthly framing anywhere.

Data-scope note (when the BHS file doesn't match the report week, e.g. ZQN single day): italic grey 10pt
paragraph at the top of page 1 stating exactly what the export covers, that figures are labelled as such
throughout, and asking for a full-week re-upload. Label the Exec Summary, throughput, observations and next
steps accordingly.

---

## 9. Charts

Currently matplotlib PNGs inserted at ~6 in wide, centred. Keep it simple and consistent:
- Bar charts, bar colour **navy `#17406D`** for faults, **cyan `#009DD9`** for lost-bag charts; value labels on
  bars; title `"<SITE> — Top Fault Categories (5-11 Oct 2026)"` / `"<SITE> — Lost-bag Events by <Checkpoint|Zone>
  (5-11 Oct 2026)"`; y-label "Event count" / "Lost-bag events"; light/no gridlines; x labels rotated.
- Part A's pie is Tiaan's **native Excel chart** (from the `Charts` sheet of the phone-log workbook) grafted as
  an OOXML chart part — Part B has no pie. (If Part B ever needs a native chart, same graft: copy
  `xl/charts/chartN.xml` + `colorsN.xml` + `styleN.xml` into `word/charts/`, add rels + content types, insert a
  `<w:drawing><wp:inline>…<c:chart r:id="…"/>` paragraph.)
- Trend chart: only per §5; line/bar of daily totals; otherwise omit with a note.
- Excluded counters (Dieback Status, Heartbeat Fault) are named in a **footnote/paragraph**, never silently dropped.

---

## 10. Document styling (confirmed against Tiaan's hand-made 28.09–04.10.2026 report)

| Element | Spec |
|---|---|
| Body | **Aptos** 11pt, colour as before (near-black `#000000` for body; grey `#767676` for meta lines) |
| Report title | Franklin Gothic Medium, bold, `#0B2545`, **20pt** (`sz=40`) |
| Date subtitle | Aptos 11pt, grey `#595959`, not bold: `Monday 5th October to Sunday 11th October 2026` (ordinal day, year only on end date) |
| Heading1 (site/section: "Support Tickets", "Weekly BHS Statistics", "BHS Performance Review", region headings) | Franklin Gothic Medium bold `#0B2545`, **15pt** (`sz=30`), **not all-caps**, space before 260 / after 120 twips |
| Heading2 (Executive Summary, "1. Total Throughput" … "Suggested Next Steps", Parcels country sub-heading) | Franklin Gothic Medium bold **`#112F51`**, **14pt** (`sz=28`) |
| Site line | Aptos 11pt **bold**, black |
| Ticket meta line | Aptos 10pt bold-italic grey `#767676`, indent ~360 |
| Bullets | Aptos 11pt; hanging indent; `Reported:/Findings:/Remedy:/Call:` labels inline, same weight |
| KPI table | full width; header row fill **`#0B2545`**, white bold Aptos 10pt; body Aptos 10pt, alternate-row fill `#F2F2F2`, thin grey borders, cell margins 60–80/100 twips |
| Data-scope / methodology notes | italic grey 10pt |
| Summary grid (Part A only) | borderless (white 4-eighth-pt borders), 3 equal cols, row height 567 twips, col1 right / col2 centre / col3 left aligned, bottom-aligned, Franklin Gothic Medium bold **ALL CAPS** `#0B2545` **14pt**, one run `AFRICA: 0` |
| Sign-off | `Warm regards,` / **`ALSTEF Group`** (bold, navy) / `Support Services Team` (plain) |
| Language | NZ English (`en-NZ`), spelling: labelled, organised, colour… |
| Page | A4 portrait, 1440-twip margins |

Colours in play: **`#17406D`** brand navy (banner, chart bars), **`#009DD9`** brand cyan, **`#0B2545`** headings,
**`#112F51`** sub-headings, **`#595959`** subtitle grey, **`#767676`** meta grey.
The old Project Instructions text still says Calibri / `#0B2545` / no Executive Summary / "United States" — all
superseded (memory file wins). **Region label is "USA"**, not "United States".

---

## 11. Letterhead — the part most likely to go wrong

- The banner is **not** a flat rectangle. It is two freeform vector waves with gradient fills (`a:custGeom` +
  `a:gradFill`; theme accent1 `#17406D` → lighter, accent2 `#009DD9` → darker) plus a white "ALSTEF GROUP"
  logo PNG, all floating drawings with wrap, followed by plain address text (`ALSTEF Group` /
  `17-19 Patiki Rd, Avondale,` / `Auckland, 1026, New Zealand`, style *Contact Info*: Franklin Gothic Medium
  bold 12pt).
- **Never reconstruct it in code.** Source of truth = `ALSTEF Group Letterhead - Template.docx` (project doc) or
  any finished report. Approach: copy the template as `base.docx`, delete everything after the banner
  paragraphs (it contains placeholder junk: `Heading:`, `x^nd April`, a stray cursor mark, and the wrong
  sign-off "Support Engineering Team"), then append content with python-docx.
- Keep the template's `theme1.xml` (accent colours drive the gradients) and the logo relationship.
- **Docdefaults trap:** the banner anchor paragraphs depend on the default font size. If `docDefaults` has no
  `<w:sz w:val="24"/>` the empty anchor paragraphs render shorter and the banner overlaps the title. Ensure
  `<w:sz w:val="24"/><w:szCs w:val="24"/>` in `w:docDefaults/w:rPrDefault`, or give the anchor paragraphs an
  explicit size.
- Banner appears **twice** per Part B report with BHS data (top of page 1; top of the Performance Review
  page). When duplicating the XML, give each copy a **unique `wp:docPr/@id`** (and keep the same image rId).
- Output must open cleanly in Word: validate XML, check `[Content_Types].xml`, and render via
  `soffice --headless --convert-to pdf` for a visual check (LibreOffice substitutes Aptos/Franklin Gothic, so
  fonts will look different there — layout is what you're checking).

---

## 12. Golden test values (week 5–11 Oct 2026) — your acceptance tests

**ZQN** (single-day export, Sun 11 Oct): T101 **3,543**; T114 3,554; OS104 **309**; L1 clear 2,231 / reject
**1,208** / no-scan 12 / timeout 0 → L1 denominator **3,451**, **35.0%**; L2 clear 1,083 / reject **90** /
timeout 35 → **2.54%** of 3,543; lost **14** over Σ tracked **20,515** = **0.07%**; checkpoints T109PC 0, T110PC 4,
T111PC 0, T112PC 1, T1DV1PD 4, T1DV1PU 1, T110APC 4; Dieback Status excluded (617). Trend chart **withheld**
(single day). Prepared for Sean Whittaker (CC Blair Poland).

**CHC** (full week): EDS bags screened 12,495 / 13,255 / 12,360 / 13,220 → **51,330**; check-in total **49,450**;
OOG **4,156**; L1 flagged **20,182 (39.3%)**; L2 reject **1,025 (2.00%)**; L2 timeouts 245; lost **2,038
(4.12%)** (Domestic CBS 751, S2 Sort Line 260, International CBS 407, TC1&TC2 prior to ME 238, Exiting ME 168,
Level 4 Line 111, S1 Sort Line 77, Between ME and L4T01 26); PM-testing faults 285 of 1,166 = **24.4%**; top
faults Intrusion System Health Fault 234, ME Waiting Bag Warning 208, Emergency Stop Fault 177. SUMMARY needed
recalculation. No trend chart. **No support tickets** for CHC → the explicit "No support tickets were logged…"
line. Prepared for Blair Poland.

**ADL** (full week): main line **63,562**; oversize **3,010**; total processed 66,572 / with recircs 68,747;
make-up output 63,335; lost **2,128 = 3.10%** (tracking rate 96.90%); SL1 935, XT1 331, XT3 324; Heartbeat
Fault 1,132 excluded; top faults Bag Waiting Alarm 339, Security Intrusion Fault 333, Emergency Stop Fault 184;
EDS all zeros → reject figures **withheld**; one ticket **WO 180537 — Call(s) 15292 (10/10/2026 06:43)**.
No "Prepared for".

Part A cross-check same week: 14 calls, 3 direct contacts, grid Africa 0 / Australia 1 / Bermuda 0 / Canada 3 /
Jamaica 0 / New Zealand 3 / USA 3 / Parcels 3; Executive Summary bullets copied verbatim from the phone-log
workbook's `Phone Stats` sheet (`14 Calls received in the month of October`, `50% … after hours`, `SFO_APAC
accounted for 24% of calls` — wording ("month") is Tiaan's own, kept as-is).

---

## 13. Suggested architecture & build order

```
reportgen/
  partb/
    sites.py          # registry (§4): code, long name, type, bhs_source, prepared_for
    validate.py       # §2 checks, content-based file detection, week parsing
    recalc.py         # §6 LibreOffice macro recalc + cross-checks
    bhs_glideview.py  # ZQN/CHC extractors (§5.1-5.2) -> BhsStats dataclass
    bhs_bag.py        # ADL extractor (§5.3)
    metrics.py        # %s, L1/L2 maths, exclusions, scope labelling
    charts.py         # matplotlib bars/trend (§9)
    text.py           # Exec-summary / observations / next-steps wording builders
    render.py         # python-docx on base.docx (banner), styles (§10-11)
    cli.py            # reportgen partb --week 05.10-11.10.2026 --sites ZQN CHC ADL "Canadian Sites"
assets/
  base.docx           # cleaned letterhead template
tests/
  test_golden_2026_10_05.py   # §12
```

`BhsStats` fields: `site, source, period_label, single_day:bool, main_throughput, main_label, secondary_note,
oversize, oversize_label, l1_flagged, l1_denominator, l2_reject, l2_timeouts, lost_total, lost_pct,
lost_pct_basis, lost_breakdown:[(name,count)], faults:[(name,count)], excluded:[(name,count,reason)],
trend:[(date,total)]|None, flags:[str]`. Renderer reads only this + the ticket entries; **all
site-specific logic lives in the extractors**, so adding a site = one new extractor + one registry row.

Build order: (1) validate + week parsing → (2) recalc + CHC cross-checks → (3) three extractors against §12 →
(4) metrics → (5) charts → (6) renderer with banner → (7) wording builders → (8) group reports (Canadian /
StarTrack: one subheading per site, no BHS) → (9) golden tests + visual check → (10) CLI.

Prompting guidance for Claude Code: give it §12 first and say "make the tests pass", then let it write the
extractors; keep the wording templates in `text.py` as data you can edit without touching logic. A rule-based
`text.py` is deterministic; if you later want LLM-written Observations, call the API for **only** that
paragraph set and keep every number sourced from `BhsStats`.

---

## 14. Known gaps / decisions still open (carry into the repo's TODO)

1. **Executive Summary wording for Part B bullet 5** and Observations are Claude-written prose; there is no
   hand-made reference — decide on templates.
2. **Aramex region/country**: lookup has none; currently nested under *New Zealand* as a flagged best guess.
3. **`AUS_POST_APAC`** generic client code: resolve facility via the matched call's location, else flag
   "facility unconfirmed" (WO 180298 on 5–11 Oct).
4. **Ticket meta format**: generated reports print `WO 180537 — Call(s) 15292 (date)`; the hand-made file
   prints `15268 (28/09/2026 10:39)` (no "WO", no label). Not yet decided which wins — confirm with the user.
   *(Resolved Oct 2026: hand-made format wins.)*
5. **ADL**: EDS all-zero export, and no confirmed "true daily throughput" column.
6. **Body-text colour**: hand-made file's Normal style is grey `#595959`; pipeline kept near-black. Not confirmed.
7. **Part A pie placement**: hand-made file floats it beside the title; pipeline stacks it below the subtitle.
8. **Weekly files for BHS must be full Mon–Sun** — ZQN's single-day export recurred as a data problem; the script
   should print a loud warning, not just a footnote.
9. Trend + "most recent full week" sections return once ≥ several consecutive weeks per site are archived —
   consider persisting each week's `BhsStats` as JSON in `history/<site>/<week>.json` **now** so that data exists.
10. Static Project Instructions text is out of date (see §10) — update it or delete it so the repo spec is the
    single source of truth.

---

## 15. Things that have already bitten (don't repeat)

- Reconstructing the banner as a flat navy rectangle; using `#0B2545` for the banner (it's `#17406D`).
- Trusting filenames (phone/CarlSource swapped) and SUMMARY cached values (CHC zeros).
- Presenting ADL's all-zero EDS as "0% rejects"; presenting ZQN's single day as a week.
- `_x000D_` leaking into comments; JSON config dumps in ticket remedies.
- Parcels country dict keyed by raw CarlSource code instead of resolved site code.
- Counting printed blocks rather than distinct tickets; calls printed twice after being matched.
- Region heading made all-caps (only the Part A summary grid is caps).
- Putting a lost-bag % from ZQN beside CHC's as if comparable (always add the caveat).
