// Keep the Assignments table in working order: pending work by due date (the Days Left colours run
// red → yellow → green), then Submitted, then Graded, then empty rows.
//
// Rows are reordered by moving cell elements between rows. Formula columns (Days Left) stay in place:
// each row's formula refers to its own row, so it follows the data automatically.

import type { Element } from "@xmldom/xmldom";
import { child, children, colToIndex, splitRef } from "./ooxml.js";
import { readTracker, type Tracker, type TrackerRow } from "./tracker.js";
import type { XlsxPackage } from "./ooxml.js";

const GROUP: Record<string, number> = { Submitted: 1, Graded: 2 };

function group(r: TrackerRow): number {
  return r.isEmpty ? 3 : (GROUP[r.status] ?? 0);
}

/** Row numbers of the table body in their sorted order (stable). */
export function sortOrder(rows: TrackerRow[]): number[] {
  return [...rows]
    .sort((a, b) => group(a) - group(b) || cmpDue(a.due, b.due) || a.row - b.row)
    .map((r) => r.row);
}

const cmpDue = (a: string | null, b: string | null) => (a === b ? 0 : a === null ? 1 : b === null ? -1 : a < b ? -1 : 1);

/** Sort the table in place. Returns how many rows changed position (0 = already sorted). */
export async function sortTable(pkg: XlsxPackage): Promise<number> {
  const t = await readTracker(pkg);
  const order = sortOrder(t.rows);
  const first = t.table.firstRow;
  const newRowOf = new Map(order.map((src, i) => [src, first + i]));
  const moved = order.filter((src, i) => src !== first + i).length;
  if (!moved) return 0;

  const movable = movableColumns(t);
  // Detach every movable cell first, then re-home it, so no row is overwritten before it is read.
  const detached = new Map<number, Element[]>();
  for (const r of t.rows) {
    const rowEl = t.sheet.rowEl(r.row);
    const cells = rowEl ? children(rowEl, "c").filter((c) => movable.has(splitRef(c.getAttribute("r")!).col)) : [];
    for (const c of cells) rowEl!.removeChild(c);
    detached.set(r.row, cells);
  }
  for (const [src, cells] of detached) {
    const dst = newRowOf.get(src)!;
    const rowEl = t.sheet.rowEl(dst, true)!;
    for (const c of cells) {
      const col = splitRef(c.getAttribute("r")!).col;
      c.setAttribute("r", `${col}${dst}`);
      const before = children(rowEl, "c").find((o) => colToIndex(splitRef(o.getAttribute("r")!).col) > colToIndex(col)) ?? null;
      rowEl.insertBefore(c, before);
    }
    rowEl.removeAttribute("spans");
  }

  // Hyperlinks are anchored by cell address; move them with their rows.
  const hl = child(t.sheet.doc.documentElement!, "hyperlinks");
  for (const h of hl ? children(hl, "hyperlink") : []) {
    const { col, row } = splitRef(h.getAttribute("ref")!);
    const dst = newRowOf.get(row);
    if (dst !== undefined && movable.has(col)) h.setAttribute("ref", `${col}${dst}`);
  }

  // A sort Excel remembered for the table no longer describes the order.
  for (const doc of [await pkg.xml(t.table.path), t.sheet.doc]) {
    const els = doc.getElementsByTagName("sortState");
    for (let i = els.length - 1; i >= 0; i--) els.item(i)!.parentNode!.removeChild(els.item(i)!);
  }
  pkg.touch(t.table.path);
  t.sheet.touch();
  await pkg.requestFullRecalc();
  return moved;
}

/** Table columns whose body holds data rather than a per-row formula. */
function movableColumns(t: Tracker): Set<string> {
  const hasFormula = (col: string) =>
    t.rows.some((r) => {
      const c = t.sheet.cellEl(`${col}${r.row}`);
      return c !== undefined && child(c, "f") !== undefined;
    });
  return new Set(t.table.columns.filter((col) => !hasFormula(col)));
}
