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

Type: assignments and quizzes are "תרגיל בית להגשה"; a quiz whose name contains בוחן or מבחן is "בוחן".

Every sync also sorts the table: pending rows by due date (red → yellow → green in Days Left; undated
last), then Submitted, then Graded, then empty rows. Each row's cells, link and formatting move together.
| Days Left, Dashboard, chart, dropdowns | Untouched |

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

## Development

```bash
npm test             # unit + round-trip tests on a private copy of the workbook (test/fixtures, git-ignored)
npm run typecheck
npx tsx scripts/make-sample.ts   # writes .out/sample.xlsx to inspect in Excel
```
