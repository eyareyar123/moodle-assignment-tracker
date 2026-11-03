# Moodle Assignment Tracker

Syncs your HUJI Moodle assignments and quizzes into `ER assignments.xlsx`.

```bash
npm run sync -- login     # one-time: log in through HUJI SSO in a browser window
npm run sync              # dry run: prints every change, writes nothing
npm run sync -- --apply   # backs up to backups/, then writes the changes
```

## How it decides what to change

Rows are matched to Moodle by a stable key `assign:<id>` / `quiz:<id>` (the `id=` in the activity URL),
stored in the **Moodle Key** column. Rows without a key are matched by their **link**, then by
same course + same due date (only if unique; ambiguous cases are skipped, never duplicated).

| Column | Who owns it |
|---|---|
| Assignment, Type, Due Date, link, Moodle Key, Moodle Status, Last Synced | Moodle (kept up to date) |
| Status | Moodle may only move it forward: Not Started → In Progress → Submitted → Graded |
| Grade (%) | Filled from Moodle only if blank |
| Subject, Priority | You (set once on new rows, never overwritten) |
| Days Left, Dashboard, chart, dropdowns | Untouched |

Type: assignments and quizzes are "תרגיל בית להגשה"; a quiz whose name contains בוחן or מבחן is "בוחן".

Every sync also sorts the table: pending rows by due date (red → yellow → green in Days Left; undated
last), then Submitted, then Graded, then empty rows. Each row's cells, link and formatting move together.

Only assignments and quizzes with a due date are synced (`--include-undated` to add the rest).

## Safety

- The workbook is edited at the XML level; every part the sync doesn't need (chart, dropdowns,
  styles, other sheets) is kept byte-for-byte. ExcelJS's full save drops the chart and dropdowns.
- A timestamped backup is written to `backups/` before every save; the save is verified and atomic.
- Refuses to write while the workbook is open in Excel.
- The token is encrypted with Windows DPAPI in `%LOCALAPPDATA%\moodle-assignment-tracker\`,
  never in the project. The workbook, backups and `.env` are git-ignored.

## New academic year

HUJI uses a new Moodle site per year. Change `siteUrl` in `tracker.config.json` and run `login` again.

## Web page (for other students)

Live: **https://eyareyar123.github.io/moodle-assignment-tracker/**. Every push to `main` redeploys it
(`.github/workflows/pages.yml`).

A static page (`web/`) does the same sync in the visitor's browser and offers the workbook as a download.
No server, no password: the visitor signs in to Moodle themselves, copies the link Moodle's mobile-app
launch page shows (`launch.php?...&confirmed=1` → “Click here if the app does not open automatically”),
and pastes it. The token in it goes only to `moodle.huji.ac.il` (enforced by a Content-Security-Policy),
and is forgotten when the tab closes. A new tracker starts from `web/public/tracker-template.xlsx`; a
visitor can also upload a tracker they downloaded before to update it.

```bash
npm run web:local      # build + serve dist/ at http://localhost:4173 (what would be published)
npm run web:dev        # live-reload dev server at http://localhost:5173 (no CSP)
npm run web:build      # dist/ is the whole site; any static host works (e.g. GitHub Pages)
npm run web:template   # rebuild the blank template from your workbook (scrubbed; checked by tests)
```

Differences from the CLI: a new tracker takes every dated activity of the year (no 14-day cut-off),
there is no ignore list, and up to 12 courses get a Subject-dropdown entry and a Dashboard chart row
(more still get their rows, with a warning).

## Development

```bash
npm test             # unit + round-trip tests on a private copy of the workbook (test/fixtures, git-ignored)
npm run test:e2e     # builds the page and drives it in headless Edge against a fake Moodle
npm run typecheck
npx tsx scripts/make-sample.ts   # writes .out/sample.xlsx to inspect in Excel
```
