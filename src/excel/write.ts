// Apply a sync Plan to the tracker in memory. Saving/backup is handled by file.ts.

import { Sheet, children, dateToSerial } from "./ooxml.js";
import { SHEET_DASHBOARD, SHEET_LISTS, readTracker, type ColKey, type Tracker } from "./tracker.js";
import type { XlsxPackage } from "./ooxml.js";
import type { LinkChange, Plan } from "../sync/plan.js";

const DATE_FIELDS: ColKey[] = ["due", "lastSynced"];

export interface ApplyResult {
  updated: number;
  added: number;
  addedRows: number[];
  subjectsAdded: string[];
  /** Things the workbook had no room for; the rows themselves are still written. */
  warnings: string[];
}

/**
 * Register new course labels so they appear in the Subject dropdown (Lists!A, within the range the
 * dropdown reads) and in the Dashboard "By Subject" block (a free row whose count formula is ready).
 * A label that does not fit is reported, not fatal: rows still carry it, only the list/chart miss it.
 */
export async function addSubjects(pkg: XlsxPackage, t: Tracker, labels: string[], warnings: string[] = []): Promise<string[]> {
  const lists = await Sheet.open(pkg, SHEET_LISTS);
  const dash = await Sheet.open(pkg, SHEET_DASHBOARD);
  const listEnd = subjectListEnd(t);
  const added: string[] = [];
  for (const label of labels) {
    let r = 4;
    while (r <= listEnd && lists.read(`A${r}`).value) r++;
    if (r > listEnd) {
      warnings.push(`No room for subject "${label}" in the Subject dropdown (Lists!A4:A${listEnd}); add it there by hand.`);
      continue;
    }
    lists.write(`A${r}`, label);

    // Dashboard: first row with an empty label next to a COUNTIF(Table1[Subject],A<row>) formula.
    const free = children(dash.sheetData, "row")
      .map((row) => Number(row.getAttribute("r")))
      .find((n) => (dash.read(`B${n}`).formula ?? "").includes("Table1[Subject]") && !dash.read(`A${n}`).value);
    if (free !== undefined) dash.write(`A${free}`, label);
    else warnings.push(`No free row for "${label}" in the Dashboard's By Subject chart.`);
    added.push(label);
  }
  return added;
}

/** Last row of the Lists range the Subject dropdown points at (e.g. Lists!$A$4:$A$10 -> 10). */
function subjectListEnd(t: Tracker): number {
  const col = t.table.cols.subject!;
  const dvs = t.sheet.doc.getElementsByTagName("x14:dataValidation");
  for (let i = 0; i < dvs.length; i++) {
    const dv = dvs.item(i)!;
    const sq = dv.getElementsByTagName("xm:sqref").item(0)?.textContent ?? "";
    const f = dv.getElementsByTagName("xm:f").item(0)?.textContent ?? "";
    const m = /Lists!\$?A\$?\d+:\$?A\$?(\d+)/.exec(f);
    if (sq.startsWith(col) && m) return Number(m[1]);
  }
  return 10;
}

export async function applyPlan(pkg: XlsxPackage, plan: Plan): Promise<ApplyResult> {
  const t = await readTracker(pkg);
  const free = t.rows.filter((r) => r.isEmpty).map((r) => r.row);
  const adds = plan.actions.filter((a) => a.kind === "add");
  if (adds.length > free.length) {
    throw new Error(`Not enough empty table rows (${free.length}) for ${adds.length} new assignments. Run "migrate" to grow the table.`);
  }

  const result: ApplyResult = { updated: 0, added: 0, addedRows: [], subjectsAdded: [], warnings: [] };
  if (plan.newSubjects.length) result.subjectsAdded = await addSubjects(pkg, t, plan.newSubjects, result.warnings);
  for (const a of plan.actions) {
    if (a.kind === "update") {
      for (const ch of a.changes) if (ch.field !== "link") writeField(t, a.row, ch.field, ch.to);
      if (a.link) await writeLink(t, a.row, a.link);
      result.updated++;
    } else if (a.kind === "add") {
      const r = free.shift()!;
      for (const [field, v] of Object.entries(a.values) as [ColKey, string | number | null][]) writeField(t, r, field, v);
      await writeLink(t, r, a.link);
      result.added++;
      result.addedRows.push(r);
    }
  }
  if (result.updated || result.added || result.subjectsAdded.length) await pkg.requestFullRecalc();
  return result;
}

function writeField(t: Tracker, row: number, field: ColKey, value: string | number | null) {
  const col = t.table.cols[field];
  if (!col) throw new Error(`Column "${field}" is missing from the table. Run "migrate" first.`);
  let v: string | number | null = value;
  if (DATE_FIELDS.includes(field) && typeof value === "string") {
    const [y, m, d] = value.split("-").map(Number) as [number, number, number];
    v = dateToSerial(y, m, d);
  }
  t.sheet.write(`${col}${row}`, v);
}

async function writeLink(t: Tracker, row: number, link: LinkChange) {
  const ref = `${t.table.cols.link}${row}`;
  if (link.text !== undefined) t.sheet.write(ref, link.text);
  await t.sheet.setHyperlink(ref, link.url);
}
