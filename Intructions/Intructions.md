═══════════════════════════════════════════════════════════
PURPOSE
═══════════════════════════════════════════════════════════
Generate two things from my weekly uploads:
PART A — a master "Weekly Support Tickets Report" covering every 
region/site.
PART B — individual or grouped "[SITE] Weekly Phone Logs Report & BHS 
Statistics" documents, built only for the sites I name each week, as 
filtered extracts of Part A plus BHS performance data.

═══════════════════════════════════════════════════════════
INPUTS I'LL UPLOAD EACH WEEK
═══════════════════════════════════════════════════════════
1. PhonePlus call log export(s) — daily files or one combined weekly 
   pull, Monday to Sunday.
2. CarlSource bulk ticket export (CSV or PDF) for that same date range, 
   status = all (open + closed) — ticket number, timestamp, site/
   location, and Reported / Findings / Remedy text.
3. (Optional, per site) GlideView export(s) — Throughput, Screening, 
   Tracking CSVs, and/or daily/weekly xlsx reports — for any site 
   needing a BHS Statistics section that week.

═══════════════════════════════════════════════════════════
PART A — MASTER WEEKLY SUPPORT TICKETS REPORT (always build, by default)
═══════════════════════════════════════════════════════════
Structure, matching the established format exactly:
1. Alstef navy letterhead banner (logo + "ALSTEF Group, 17-19 Patiki Rd, 
   Avondale, Auckland, 1026, New Zealand") at the top.
2. Title: "Weekly Support Tickets Report:"
3. Date range subtitle: "[Monday Dth Month] to [Sunday Dth Month YYYY]"
4. "Total Amount of Calls: [N]"
5. "Total Amount of Direct Contacts: [N]"
6. Summary grid table — one cell per region, three per row, count only: 
   Africa, Australia, Bermuda / Canada, Jamaica, New Zealand / United 
   States, Parcels.
7. One heading per region, in that order, even if the count is 0 (keep 
   the heading, no ticket blocks underneath): "Africa: [N] ticket(s)", 
   "Australia: [N] ticket(s)", etc.
8. Under each region, one block per site: site name + airport/site code 
   (bold), then ticket number(s) + timestamp, then Reported / Findings 
   / Remedy as bullets (see TICKET FORMATTING rule below).
9. Parcels is a special region: it nests a country sub-heading one 
   level down (e.g. "Parcels: 1 ticket" → "Australia: 1 ticket" → the 
   ticket block), matching the real example on file.
10. Sort within each region: Location (A→Z) → Call Number (oldest to 
    newest).
11. Sign-off: "Warm regards, / ALSTEF Group / Support Services Team"
12. Output filename: "Weekly_Support_Tickets_Report_[startdate]_to_[enddate].docx"

═══════════════════════════════════════════════════════════
PART B — SITE-SPECIFIC REPORTS (build only the ones I name each week)
═══════════════════════════════════════════════════════════
Standing groups — when I name one of these, build it as described:
  - ADL (Adelaide) — individual report
  - BDA — individual report
  - South Island — ONE combined report covering DUD (Dunedin), ZQN 
    (Queenstown), and CHC (Christchurch), each under its own subheading
  - Canada — ONE combined report covering the ~19 supported Canadian 
    sites, each under its own subheading

If I name a site not on this list, ask whether it's a one-off or should 
be added as a standing group for future weeks.

Title: "[SITE / SITE GROUP] Weekly Phone Logs Report & BHS Statistics"

For each site's subheading within a report:
- Filter Part A's data down to just that site's ticket(s), same bullet 
  format. If none that week, say so explicitly: "No support tickets 
  were logged for [SITE] during this period."
- Add a "[SITE] – Weekly BHS Statistics" table whenever GlideView data 
  was uploaded for that site that week:
    • Total throughput (main line, excl. oversize)
    • Oversize line volume
    • L1 flagged for re-screen (count and %)
    • L2 confirmed reject — final (count and % of throughput)
    • Lost-bag events (count and % of tracked scans)
  Adapt tag names (T101, OS104, etc.) to whatever that site's own 
  export uses — don't assume Queenstown's tags apply elsewhere. Omit 
  the BHS section entirely for a site if no GlideView data was 
  uploaded, rather than inventing figures.
- Append a "BHS Performance Review" subsection whenever weekly 
  GlideView data is provided for that site. This is no longer gated on 
  having a full month's data — the review now covers the reporting 
  week itself, not a calendar month, and the "Monthly" framing is 
  dropped throughout. Structure, as refined to actually fit a 
  single-week export (confirmed working on the ZQN and CHC reports):

    Executive Summary (5 bullets) → 
    1. Total Throughput → 
    2. Reject Bags (L1 auto-flag vs L2 confirmed reject, explained 
       clearly — L1 is NOT a final reject) → 
    3. Fault Subsections — comparative bar chart → 
    Lost-bag events — comparative bar chart by zone → 
    Observations & Discussion Points → 
    Suggested Next Steps.

  "Comparative Percentage/trend" (week-over-week) and "Snapshot: most 
  recent full week" are DROPPED from the weekly format — a single 
  week has nothing to compare against yet. Reintroduce both once 
  several consecutive weeks of exports are on file for that site.

  In their place:
    • Section 3 (Fault Subsections) and the Lost-bag chart give a 
      *within-week comparative* view (across fault categories / 
      tracking zones) instead of a week-over-week trend. This is the 
      "comparative data across subsections" approach, not a time 
      trend.
    • Trend charts: only build a genuine day-by-day throughput trend 
      if the site's GlideView export includes a 5-minute-interval 
      time-series sheet — e.g. ZQN's "System Throughput Rates" sheet, 
      using its main-throughput-tag column (T114) aggregated to daily 
      totals. If a site's export has no such sheet (e.g. CHC), omit 
      the trend chart entirely and say so in a note rather than 
      inventing one.
    • Fault chart: pull the top ~10 categories from the Fault Table 
      sheet, summed across all devices for the week. Exclude obvious 
      non-fault heartbeat/status counters if one dominates the data 
      (e.g. ZQN's "Dieback Status", tens of thousands of hits vs. 
      hundreds for real faults) — always name what was excluded and 
      why, in a footnote, rather than silently dropping it.
    • Lost-bag chart: break the weekly total down by whatever 
      granularity that site's export actually provides (flat 
      checkpoint list for ZQN; system-zone groupings for CHC) — don't 
      force one site's zone structure onto another's.

Report-level:
- Same letterhead banner.
- Facilities contact under a site's subheading if known (e.g. Sean 
  Whittaker, Facilities Manager, for Queenstown). If not known, omit 
  the "Prepared for" line rather than guessing, and ask me once.
- Same sign-off as Part A.
- Output filename: "[SITE-OR-GROUP]_Weekly_Phone_Logs_and_BHS_Statistics.docx"

═══════════════════════════════════════════════════════════
TICKET FORMATTING
═══════════════════════════════════════════════════════════
Always use bullets for Reported / Findings / Remedy, even if the source 
ticket text (CarlSource export) is written as plain paragraphs. Convert 
to bullets every time — don't preserve the source formatting. Applies 
to both Part A and Part B.

═══════════════════════════════════════════════════════════
DATA HANDLING
═══════════════════════════════════════════════════════════
- Always cross-check GlideView totals against the source file's SUMMARY 
  sheet (if present) before finalizing.
- SUMMARY-sheet formulas can carry stale/zero cached values until the 
  workbook is recalculated (seen on the CHC export — every SUMMARY 
  total read as 0 despite valid underlying data). Recalculate the 
  workbook (e.g. via LibreOffice headless) before trusting any 
  SUMMARY-sheet figure, and flag it to me if a source file needed this 
  fix so it can be raised with whoever generates the export.
- If the phone-log week and BHS-stats week don't line up exactly, say 
  so explicitly rather than blending them silently.
- Never blend two sites' or two regions' data into one entry.
- Site-specific throughput/tag conventions, once established, carry 
  forward automatically without re-asking each week:
    • ZQN — main-line throughput = T101/T114; oversize = OS104; 
      lost-bag total = flat sum across its checkpoint list (T110PC, 
      T111PC, T112PC, T1DV1PD, T1DV1PU, T110APC, T109PC).
    • CHC — no single main-line tag exists. Main-line throughput = sum 
      of bags screened across its four parallel EDS lines 
      (X1X0100–X1X0400). Report the check-in-side total 
      (C1+C2+C3+Tranship) as a secondary cross-check figure — it won't 
      match the EDS-aggregate exactly, due to re-screening, and that's 
      expected, not an error. Lost-bag total = the 8-zone Tracking 
      Stats aggregate on its SUMMARY sheet, not a flat checkpoint sum.
    • Because these two sites' lost-bag figures come from structurally 
      different exports (flat list vs. zone aggregate), always note 
      in the report that the percentages aren't directly comparable 
      site-to-site.
    • When a new site appears for the first time, work out its own 
      convention from its export structure (following the same logic 
      as above), state the assumption made, and add it to this list 
      for future weeks.

═══════════════════════════════════════════════════════════
STYLE
═══════════════════════════════════════════════════════════
NZ English spelling throughout. Navy (#0B2545) headings and KPI table 
headers, light-grey secondary text, clean bordered tables, Alstef navy 
letterhead banner with white logo. Output as downloadable .docx file(s).

═══════════════════════════════════════════════════════════
PROJECT KNOWLEDGE ON FILE (reference these, don't ask me to re-explain)
═══════════════════════════════════════════════════════════
- TEMPLATE_Master_Weekly_PhoneLogs_Report.docx — master structural skeleton
- Weekly_Support_Phone_Logs_29_06_to_05_07_2026.docx — real finished 
  example of Part A
- TEMPLATE_Site_Weekly_PhoneLogs_and_BHS_Statistics.docx — site 
  structural skeleton
- ZQN_Weekly_Phone_Logs_Report_03_08_26_to_09_08_26.docx — real 
  finished example of Part B