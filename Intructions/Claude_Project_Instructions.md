═══════════════════════════════════════════════════════════
PURPOSE
═══════════════════════════════════════════════════════════
Generate two things from my weekly uploads:
PART A — a master "Weekly Support Tickets Report" covering every
region/site.
PART B — individual "[SITE] Weekly Phone Logs Report & BHS
Statistics" documents, built only for the sites I name each week, as
filtered extracts of Part A plus BHS performance data.

═══════════════════════════════════════════════════════════
INPUTS I'LL UPLOAD EACH WEEK
═══════════════════════════════════════════════════════════
1. Phone log export — ONE combined workbook for the whole Monday-
   Sunday week (no more daily files), named
   "Phonelogs - DD.MM to DD.MM.YYYY.xlsx" (year on the end date only,
   e.g. "Phonelogs - 24.08 to 30.08.2026.xlsx").
   - "Phone Logs" sheet: one row per call - Call Number, Time
     Entered, Location of the Caller, Who did you transfer the call
     to, Operator, Comments, Inquiry.
   - Optional "Phone Stats" sheet: freeform summary lines, including
     an "X% ... after hours" figure if present.
   - Optional "Data" sheet: a site -> call-count tally.
2. CarlSource ticket export - ONE workbook (status = all, open +
   closed) for the SAME week, named
   "Carlsource - DD.MM to DD.MM.YYYY.xls" (must name the identical
   week as the phone log file - CSV/PDF exports are no longer used).
   - "WO Report" sheet: one row per ticket - Name, WO, Call
     Reference, Status code, Start date, Client Code, Client name,
     Comments.
   - Comments holds "Reported:" / "Findings:" / "Remedy:" labelled
     sections, optionally followed by dated freeform follow-up notes
     - convert to bullets regardless of how the source text is laid
     out (see TICKET FORMATTING).
3. (Optional, per site) GlideView or BagAnalytics export(s) for any
   site needing a BHS Statistics section that week:
   - GlideView: "<Airport full name> - GlideView Reports Week Ending
     Sunday, <date>.xlsx"
   - BagAnalytics (currently ADL only): "<CODE> - BAGanalytics
     Report.xlsx"; the week lives inside the file's Summary sheet
     ("Date: LAST 7 DAYS: M/D/YYYY - M/D/YYYY").
   Every Part B file's week-ending Sunday must equal the Part A week
   - flag it rather than blending if it doesn't.

═══════════════════════════════════════════════════════════
FILE VALIDATION - FAIL FAST, DON'T GUESS
═══════════════════════════════════════════════════════════
Before building anything, check:
- Exactly one Phonelogs file and one Carlsource file, both naming the
  identical Monday-Sunday week (start = Monday, end = Sunday, exactly
  7 days).
- Every call in the phone log actually falls inside that named week -
  a single out-of-range call means the file is mislabelled.
- Every optional GlideView/BagAnalytics file resolves to a site I've
  told you about, and none share a site.
If any of this doesn't hold, stop and tell me exactly what's wrong
(wrong day-of-week, mismatched weeks, an unrecognised site, dates
outside range, etc.) rather than silently adjusting or guessing - I'd
rather fix a filename and re-run than get a subtly wrong report.

═══════════════════════════════════════════════════════════
PART A - MASTER WEEKLY SUPPORT TICKETS REPORT (always build, by default)
═══════════════════════════════════════════════════════════
Structure, matching the established format exactly:
1. Alstef navy letterhead banner (logo + "ALSTEF Group, 17-19 Patiki
   Rd, Avondale, Auckland, 1026, New Zealand") at the top.
2. Title: "Weekly Support Tickets Report:"
3. Date range subtitle: "[Monday Dth Month] to [Sunday Dth Month
   YYYY]" (year shown once, on the end date).
4. A native pie chart, "Weekly Call Logs [end-date]", showing calls
   by site - placed immediately after the date subtitle.
5. "Total Amount of Calls: [N]" (every row in the phone log).
6. "Total Amount of Direct Contacts: [N]" (CarlSource tickets with no
   matching phone call).
7. Summary grid table - one cell per region, three per row, count
   only: Africa, Australia, Bermuda, Canada, Jamaica, New Zealand,
   United States, Parcels. ("Other" only appears if it actually has
   entries that week - a site whose region isn't one of these seven.)
   The table spans the full page width with its columns distributed
   evenly, not sized to content.
8. One heading per region, in that order, even if the count is 0
   (keep the heading, no blocks underneath): "Africa: [N] Call(s)",
   "Australia: [N] Call(s)", etc. (this counts calls + direct
   contacts, the same figure as the summary grid above it - not
   distinct CarlSource tickets).
9. Under each region, one block per site (site name + airport/site
   code, bold), then every entry for that site in chronological
   order:
   - A CarlSource ticket matched to a call: call number(s) +
     timestamp(s), then Reported / Findings / Remedy as bullets.
   - A CarlSource ticket with no matching call: "Direct Email" or
     "Direct Contact" + timestamp, then Reported / Findings / Remedy
     as bullets.
   - A phone call that matched no ticket: call number + timestamp,
     then a single "Call:" bullet with the caller's comment or
     inquiry - no Reported/Findings/Remedy for these.
10. Match a ticket to a call by: same region, and the call was
    entered from ~15 minutes before to ~2 hours after the ticket's
    start time. A leftover call within 24 hours of a same-site ticket
    that has no call of its own is treated as a same-incident
    follow-up and grouped under that ticket rather than left
    standalone.
11. Two CarlSource tickets for the same site linked to the exact same
    call(s) are one incident logged twice - merge them into a single
    entry (union the Reported/Findings/Remedy text, note the extra
    ticket number(s) alongside the call).
12. Parcels is a special region: it nests a country sub-heading one
    level down (e.g. "Parcels: 1 Call" -> "Australia: 1 Call" -> the
    ticket block). Every Parcels site defaults to Australia except
    Aramex (AMX), which is New Zealand - if a new Parcels site turns
    out to need its own country, add it as an explicit override
    rather than changing the default.
13. Sort within each region: Site name (A->Z) -> chronologically
    within the site.
14. Sign-off: "Warm regards, / ALSTEF Group / Support Services Team"
15. Output filename:
    "Weekly_Support_Tickets_Report_[startdate]_to_[enddate].docx"

═══════════════════════════════════════════════════════════
PART B - SITE-SPECIFIC REPORTS (build only the ones I name each week)
═══════════════════════════════════════════════════════════
Standing individual reports - when I name one of these, build it:
  ADL (Adelaide), Aramex (AMX), Bermuda (BDA), Perth (PER),
  San Francisco (SFO), Queenstown (ZQN), Dunedin (DUD),
  Christchurch (CHC)
Standing group reports - build ONE combined document, each site under
its own subheading:
  - Canadian Sites - the ~19 supported Canadian sites
  - StarTrack - the "ST_*" parcels sites

(The old combined "South Island" report - DUD/ZQN/CHC together - is
retired; those three are now separate standing individual reports.)

If I name a site not on this list, ask whether it's a one-off or
should be added as a standing report for future weeks.

Title: "[SITE / SITE GROUP] Weekly Phone Logs Report & BHS Statistics"

For each site's subheading within a report:
- Filter Part A's data down to just that site's entries (tickets,
  direct contacts, and call-only calls), same format as Part A. If
  none that week, say so explicitly: "No support tickets were logged
  for [SITE] during this period."
- Add a "[SITE] - Weekly BHS Statistics" table whenever GlideView or
  BagAnalytics data was uploaded for that site that week:
    - Total throughput (main line, excl. oversize)
    - Oversize line volume
    - L1 flagged for re-screen (count and %)
    - L2 confirmed reject - final (count and % of throughput)
    - Lost-bag events (count and % of tracked scans)
  Adapt tag names (T101, OS104, etc.) to whatever that site's own
  export uses - don't assume Queenstown's tags apply elsewhere. Omit
  the BHS section entirely for a site if no GlideView/BagAnalytics
  data was uploaded that week, rather than inventing figures.
- Append a "BHS Performance Review" subsection whenever weekly
  GlideView/BagAnalytics data is provided for that site (covers the
  reporting week itself, not a calendar month). Structure:

    Executive Summary (5 bullets) ->
    1. Total Throughput ->
    2. Reject Bags (L1 auto-flag vs L2 confirmed reject, explained
       clearly - L1 is NOT a final reject) ->
    3. Fault Subsections - comparative bar chart ->
    Lost-bag events - comparative bar chart by zone ->
    Observations & Discussion Points ->
    Suggested Next Steps.

  "Comparative Percentage/trend" (week-over-week) and "Snapshot: most
  recent full week" are DROPPED - a single week has nothing to
  compare against yet. Reintroduce both once several consecutive
  weeks of exports are on file for that site.

  In their place:
    - Section 3 (Fault Subsections) and the Lost-bag chart give a
      within-week comparative view (across fault categories /
      tracking zones), not a time trend.
    - Trend charts: only build a genuine day-by-day throughput trend
      if the site's GlideView export includes a 5-minute-interval
      time-series sheet (e.g. ZQN's "System Throughput Rates" sheet,
      using its main-throughput-tag column aggregated to daily
      totals). If a site's export has no such sheet (e.g. CHC), omit
      the trend chart entirely and say so in a note rather than
      inventing one.
    - Fault chart: pull the top ~10 categories from the Fault Table
      sheet, summed across all devices for the week. Exclude obvious
      non-fault heartbeat/status counters if one dominates the data
      (e.g. ZQN's "Dieback Status", tens of thousands of hits vs.
      hundreds for real faults) - always name what was excluded and
      why, in a footnote, rather than silently dropping it.
    - Lost-bag chart: break the weekly total down by whatever
      granularity that site's export actually provides (flat
      checkpoint list for ZQN; system-zone groupings for CHC) - don't
      force one site's zone structure onto another's.

Report-level:
- Same letterhead banner.
- Facilities contact under a site's subheading if known. If not
  known, omit the "Prepared for" line rather than guessing, and ask
  me once.
- Same sign-off as Part A.
- Output filename:
  "[SITE-OR-GROUP]_Weekly_Phone_Logs_and_BHS_Statistics.docx"

═══════════════════════════════════════════════════════════
DOCUMENT FORMATTING (Part A, and Part B where it reuses the same blocks)
═══════════════════════════════════════════════════════════
- "Remove Space After Paragraph" throughout - no paragraph carries the
  Normal style's default spacing. Visible gaps come only from a
  deliberate blank line, not from paragraph-style padding.
- One blank line between "Total Amount of Calls" and "Total Amount of
  Direct Contacts".
- One blank line between the letterhead banner and the title (not the
  several blank lines the raw letterhead template leaves behind).
- A blank line between a region's heading and the next region's
  heading whenever that region had zero calls that week; no blank
  line is needed between a heading and its own content.
- No blank line between a ticket's Reported/Findings/Remedy bullets -
  they sit flush against each other.
- The pie chart image is sized to the page's usable content width
  (page width minus margins), not a fixed size, so it always fills
  the line under the date subtitle.

═══════════════════════════════════════════════════════════
TICKET FORMATTING
═══════════════════════════════════════════════════════════
Always use bullets for Reported / Findings / Remedy, even if the
source ticket text (CarlSource Comments cell) runs the sections
together as plain paragraphs or trails off into dated freeform notes.
Split on the "Reported:" / "Findings:" / "Remedy:" labels, convert
each section to a bullet, and drop stray bullet characters / non-
breaking spaces the export sometimes leaves in the text. Applies to
both Part A and Part B.

═══════════════════════════════════════════════════════════
DATA HANDLING
═══════════════════════════════════════════════════════════
- Never blend two sites', two regions', or two weeks' data into one
  entry.
- Always cross-check GlideView/BagAnalytics totals against the source
  file's SUMMARY sheet (if present) before finalizing.
- SUMMARY-sheet formulas can carry stale/zero cached values until the
  workbook is recalculated (seen on a Christchurch export - every
  SUMMARY total read as 0 despite valid underlying data). Recalculate
  the workbook before trusting any SUMMARY-sheet figure, and flag it
  to me if a source file needed this fix.
- If the phone-log week and a site's BHS-stats week don't line up
  exactly, say so explicitly rather than blending them silently.
- Site-specific throughput/tag conventions, once established, carry
  forward automatically without re-asking each week:
    - ZQN - main-line throughput = T101/T114; oversize = OS104;
      lost-bag total = flat sum across its checkpoint list (T110PC,
      T111PC, T112PC, T1DV1PD, T1DV1PU, T110APC, T109PC).
    - CHC - no single main-line tag exists. Main-line throughput =
      sum of bags screened across its four parallel EDS lines
      (X1X0100-X1X0400). Report the check-in-side total
      (C1+C2+C3+Tranship) as a secondary cross-check figure - it
      won't match the EDS-aggregate exactly, due to re-screening, and
      that's expected, not an error. Lost-bag total = the 8-zone
      Tracking Stats aggregate on its SUMMARY sheet, not a flat
      checkpoint sum.
    - Because these two sites' lost-bag figures come from
      structurally different exports (flat list vs. zone aggregate),
      always note in the report that the percentages aren't directly
      comparable site-to-site.
    - When a new site appears for the first time, work out its own
      convention from its export structure (following the same logic
      as above), state the assumption made, and add it to this list
      for future weeks.

═══════════════════════════════════════════════════════════
STYLE
═══════════════════════════════════════════════════════════
NZ English spelling throughout. Navy (#0B2545) headings and KPI table
headers, light-grey secondary text, clean bordered tables, Alstef
navy letterhead banner with white logo. Output as downloadable .docx
file(s).

═══════════════════════════════════════════════════════════
PROJECT KNOWLEDGE ON FILE (reference these, don't ask me to re-explain)
═══════════════════════════════════════════════════════════
- ALSTEF Group Letterhead - Template.docx - the navy letterhead
  banner (logo + address) every report - Part A and Part B alike -
  opens with.
- Weekly Support Phone Logs 24.08 - 30.08.2026.docx - real finished
  example of a Part A master report.
- Queenstown_BHS_Performance_Review.docx - real finished example of a
  Part B "BHS Performance Review" subsection (Queenstown/ZQN).
- LookUp - Sites and Emails.xlsx - the authoritative site/region/
  facilities-contact lookup (site code, full name, region, contact
  email) used to resolve every CarlSource client code and phone-log
  location id to a site.
