# Job Tracker

A self-updating spreadsheet of product roles, pulled straight from the source.

Instead of asking Google what it happened to crawl, this queries the JSON APIs
that Greenhouse, Lever, Ashby and Workday job boards are themselves built on.
Every posting in the sheet is live at fetch time, which removes the stale-listing
problem that Google `site:` searches have — the CVS req that stayed indexed for
weeks after it closed would simply drop off the Active tab on the next run.

It also fixes the Workday blind spot. Workday tenants block crawlers in
`robots.txt`, so `site:myworkdayjobs.com` under-reports badly. Their own
`/wday/cxs/` endpoint has no such restriction and returns everything, with
posted dates attached.

## Setup

Already done — the virtualenv exists and the first run has populated the sheet.

```bash
cd job-tracker && ./run.sh
```

Open **`Job Tracker.xlsx`**.

## The tabs

| Tab | What it holds |
|---|---|
| **Active** | Every currently-live match. The first four columns are yours. |
| **New This Run** | Only what appeared since the last run — the daily read. |
| **Closed** | Postings that have disappeared, with how many days they stayed up. |
| **Boards** | Every board polled, how many postings it had, how many matched, and any errors. |
| **Search Queries** | Your original Google operators, with a rolling `after:` date, as clickable links. |
| **Run Log** | Counts from the last run. |

### Your columns are safe

`Status`, `Priority`, `Applied On` and `Notes` are read back out of the workbook
before each refresh and written straight back. Type in them freely — a run will
never overwrite them, and they follow a posting to the Closed tab when the req
comes down. Close the file before running, or Excel will hold a lock on it.

Status and Priority are dropdowns. Rows first seen today are highlighted;
rows marked `Applied` fade out. Clearing a cell sticks — the blank is treated
as a deliberate edit rather than reverting to what was there before.

The workbook is the source of truth for those four columns; `state.json` is just
the durable copy behind it. If you ever want to wipe your tracking, edit the
sheet, not the JSON.

## Automating it

Daily at 7:15am:

```bash
cp com.averyburbank.jobtracker.plist ~/Library/LaunchAgents/ && launchctl load ~/Library/LaunchAgents/com.averyburbank.jobtracker.plist
```

To stop:

```bash
launchctl unload ~/Library/LaunchAgents/com.averyburbank.jobtracker.plist
```

## Adding companies

Paste a careers URL — the board is identified and verified before it is saved:

```bash
./.venv/bin/python job_tracker.py add-board https://job-boards.greenhouse.io/anthropic --name Anthropic
```

Works for `job-boards.greenhouse.io/…`, `boards.greenhouse.io/…`,
`jobs.lever.co/…`, `jobs.ashbyhq.com/…` and any
`<tenant>.wd<N>.myworkdayjobs.com/<site>` URL. List what is configured with
`./.venv/bin/python job_tracker.py boards`.

This is the part worth feeding. The tool only sees companies it has been told
about, so when a `site:` search in the Search Queries tab turns up a company
that isn't in the Boards tab yet, add it and it is covered from then on.

## Tuning `config.json`

`title_include` matches enterprise phrasing, not just literal titles — large
health plans post `Manager, Product Management` and `Product Management Senior
Advisor` far more often than `Product Manager`, and those are matched.

`title_exclude` currently drops interns, program/project managers, product
marketing, product design, analysts and engineering managers. Levels are *not*
excluded — director and principal roles come through. To strip them, add
`"director"` and `"principal"` to that list.

`locations_any` covers remote phrasings (including Workday's "Work At Home",
which does not contain the word "remote") plus Minnesota. It also allows the
broad `united states`, which lets through US roles that are not remote-tagged;
remove it to tighten to genuinely remote + Twin Cities.

`workday_category_keywords` picks which Workday job categories to sweep. This is
what makes Workday tractable: filtering CVS to its product category turns 19,000
postings into about 55.

## Known limits

- **Workday is slow.** Ten tenants, paged 20 at a time, is most of the runtime.
  Hammering it can get you throttled; once a day is comfortable.
- **Coverage is a function of the board list.** 62 boards is a starting set, not
  a market sweep. The Google operators remain the discovery tool; this is the
  tracking tool.
- **Multi-site Workday reqs show as "5 Locations"** with no list of which ones.
  Those are let through the location filter rather than dropped, since they
  often include the ones being looked for — so expect some Active rows whose
  location needs a click to resolve.
- **`Posted` on Workday is approximate.** It reports "Posted 30+ Days Ago" and
  similar, so anything older than a month is pinned at 30 days.
