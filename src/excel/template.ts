// Turn a filled-in tracker into a blank template other people can start from.
//
// Everything personal goes: table rows, links, subjects, the extra personal sheet, cached formula and
// chart values (they echo your data), strings Excel keeps around after you delete them, printer
// settings and author metadata. The layout, styles, formulas, dropdowns and chart stay.

import type { Document, Element } from "@xmldom/xmldom";
import { XlsxPackage, Sheet, NS_REL, REL_HYPERLINK, child, children } from "./ooxml.js";
import { migrate, stretchSqref } from "./migrate.js";
import { SHEET_ASSIGNMENTS, SHEET_DASHBOARD, SHEET_LISTS, readTracker } from "./tracker.js";

/** Subject slots in the template: Lists!A4:A15 feeds the dropdown, Dashboard A12:B23 the chart. */
export const SUBJECT_SLOTS = 12;
const LIST_FIRST = 4;
const DASH_FIRST = 12;

const REL_PRINTER = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/printerSettings";
const REL_SHARED_STRINGS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings";

/** Sheets that are not part of the tracker and are dropped from the template. */
const KEEP_SHEETS = ["How to Use", SHEET_DASHBOARD, SHEET_ASSIGNMENTS, SHEET_LISTS];

export async function makeTemplate(bytes: Uint8Array): Promise<Uint8Array> {
  const pkg = await XlsxPackage.load(bytes);
  await migrate(pkg);
  await dropExtraSheets(pkg);
  await clearAssignments(pkg);
  await resetSubjects(pkg);
  for (const path of await sheetPaths(pkg)) {
    await dropPrinterSettings(pkg, path);
    clearCachedValues(await pkg.xml(path));
    pkg.touch(path);
  }
  await compactSharedStrings(pkg);
  await scrubMetadata(pkg);
  await pkg.requestFullRecalc();
  return pkg.save();
}

// ---------- sheets ----------

async function sheetPaths(pkg: XlsxPackage): Promise<string[]> {
  const wb = await pkg.xml("xl/workbook.xml");
  const names = children(child(wb.documentElement!, "sheets")!, "sheet").map((s) => s.getAttribute("name")!);
  return Promise.all(names.map((n) => pkg.sheetPath(n)));
}

async function dropExtraSheets(pkg: XlsxPackage) {
  const wb = await pkg.xml("xl/workbook.xml");
  const sheets = child(wb.documentElement!, "sheets")!;
  const rels = await pkg.xml("xl/_rels/workbook.xml.rels");
  for (const s of children(sheets, "sheet")) {
    const name = s.getAttribute("name")!;
    if (KEEP_SHEETS.includes(name)) continue;
    const path = await pkg.sheetPath(name);
    const rid = s.getAttributeNS(NS_REL, "id");
    for (const r of children(rels.documentElement!, "Relationship")) if (r.getAttribute("Id") === rid) rels.documentElement!.removeChild(r);
    sheets.removeChild(s);
    await dropPrinterSettings(pkg, path);
    await removePart(pkg, path);
  }
  // Point the active tab at a sheet that still exists.
  const view = wb.getElementsByTagName("workbookView").item(0);
  const active = Number(view?.getAttribute("activeTab") ?? 0);
  if (view && active >= children(sheets, "sheet").length) view.setAttribute("activeTab", "0");
  pkg.touch("xl/workbook.xml");
  pkg.touch("xl/_rels/workbook.xml.rels");

  // docProps/app.xml lists sheet titles.
  if (pkg.has("docProps/app.xml")) {
    const app = await pkg.xml("docProps/app.xml");
    const titles = app.getElementsByTagName("TitlesOfParts").item(0);
    const vec = titles && child(titles as Element, "vector");
    if (vec) {
      for (const l of children(vec, "lpstr")) if (!KEEP_SHEETS.includes(l.textContent ?? "")) vec.removeChild(l);
      vec.setAttribute("size", String(children(vec, "lpstr").length));
    }
    const counts = app.getElementsByTagName("vt:i4");
    if (counts.length === 1 && vec) counts.item(0)!.textContent = String(children(vec, "lpstr").length);
    pkg.touch("docProps/app.xml");
  }
}

async function clearAssignments(pkg: XlsxPackage) {
  const t = await readTracker(pkg);
  const sheet = t.sheet;
  for (const r of t.rows) {
    for (const col of t.table.columns) {
      const c = sheet.cellEl(`${col}${r.row}`);
      if (c && !child(c, "f")) sheet.write(`${col}${r.row}`, null);
    }
  }
  // Links and their relationships.
  const root = sheet.doc.documentElement!;
  const hl = child(root, "hyperlinks");
  if (hl) root.removeChild(hl);
  await removeRels(pkg, sheet.path, (type) => type === REL_HYPERLINK);
  // Dropdown/format ranges end where the table does.
  const fix = (s: string) => stretchSqref(s, t.table.firstRow, t.table.lastRow);
  for (const cf of children(root, "conditionalFormatting")) cf.setAttribute("sqref", fix(cf.getAttribute("sqref")!));
  const xm = sheet.doc.getElementsByTagName("xm:sqref");
  for (let i = 0; i < xm.length; i++) xm.item(i)!.textContent = fix(xm.item(i)!.textContent ?? "");
  sheet.touch();
}

/** Empty subject list and By Subject block, sized to SUBJECT_SLOTS, with the chart reading all of it. */
async function resetSubjects(pkg: XlsxPackage) {
  const lists = await Sheet.open(pkg, SHEET_LISTS);
  const listLast = LIST_FIRST + SUBJECT_SLOTS - 1;
  const listStyle = lists.read(`A${LIST_FIRST}`).style;
  for (let r = LIST_FIRST; r <= Math.max(listLast, 30); r++) {
    if (r <= listLast) lists.write(`A${r}`, null, { style: listStyle });
    else if (lists.cellEl(`A${r}`)) lists.write(`A${r}`, null);
  }

  const assignments = await Sheet.open(pkg, SHEET_ASSIGNMENTS);
  const dvs = assignments.doc.getElementsByTagName("xm:f");
  for (let i = 0; i < dvs.length; i++) {
    const f = dvs.item(i)!;
    f.textContent = (f.textContent ?? "").replace(/^Lists!\$A\$4:\$A\$\d+$/, `Lists!$A$${LIST_FIRST}:$A$${listLast}`);
  }
  assignments.touch();

  const dash = await Sheet.open(pkg, SHEET_DASHBOARD);
  const formulaRows = children(dash.sheetData, "row")
    .map((row) => Number(row.getAttribute("r")))
    .filter((n) => (dash.read(`B${n}`).formula ?? "").includes("Table1[Subject]"));
  const lastExisting = Math.max(...formulaRows);
  const aStyle = dash.read(`A${lastExisting}`).style;
  const bStyle = dash.read(`B${lastExisting}`).style;
  const rowAttrs = dash.rowEl(lastExisting)!;
  const dashLast = DASH_FIRST + SUBJECT_SLOTS - 1;
  for (let r = DASH_FIRST; r <= dashLast; r++) {
    if (!dash.rowEl(r)) {
      const row = dash.rowEl(r, true)!;
      for (const a of ["ht", "customHeight", "x14ac:dyDescent"]) {
        const v = rowAttrs.getAttribute(a);
        if (v !== null) row.setAttribute(a, v);
      }
    }
    dash.write(`A${r}`, null, { style: aStyle });
    dash.writeFormula(`B${r}`, `COUNTIF(Table1[Subject],A${r})`, { style: bStyle });
  }
  const dim = child(dash.doc.documentElement!, "dimension");
  if (dim) dim.setAttribute("ref", (dim.getAttribute("ref") ?? "A1:H1").replace(/\d+$/, (n) => String(Math.max(Number(n), dashLast))));

  // Chart: categories/values span the whole block; cached copies of your data go.
  for (const rel of await pkg.rels(dash.path)) {
    if (!rel.type.endsWith("/drawing")) continue;
    const drawing = XlsxPackage.resolveTarget(dash.path, rel.target);
    for (const crel of await pkg.rels(drawing)) {
      if (!crel.type.endsWith("/chart")) continue;
      const chartPath = XlsxPackage.resolveTarget(drawing, crel.target);
      const chart = await pkg.xml(chartPath);
      const fs = chart.getElementsByTagName("c:f");
      for (let i = 0; i < fs.length; i++) {
        const f = fs.item(i)!;
        f.textContent = (f.textContent ?? "").replace(/^(Dashboard!\$[AB]\$)(\d+):(\$[AB]\$)\d+$/, `$1${DASH_FIRST}:$3${dashLast}`);
      }
      for (const tag of ["c:strCache", "c:numCache"]) {
        const els = chart.getElementsByTagName(tag);
        for (let i = els.length - 1; i >= 0; i--) els.item(i)!.parentNode!.removeChild(els.item(i)!);
      }
      pkg.touch(chartPath);
    }
  }
}

// ---------- package hygiene ----------

/** Formula results are recomputed on open (fullCalcOnLoad); stored ones would echo your numbers. */
function clearCachedValues(doc: Document) {
  const cells = doc.getElementsByTagName("c");
  for (let i = 0; i < cells.length; i++) {
    const c = cells.item(i)!;
    if (!child(c, "f")) continue;
    const v = child(c, "v");
    if (v) c.removeChild(v);
    if (c.getAttribute("t") === "str" || c.getAttribute("t") === "e" || c.getAttribute("t") === "b") c.removeAttribute("t");
  }
}

/**
 * Rewrite the shared string table to the strings still in use. Excel never prunes it, so it still holds
 * every assignment name you ever typed.
 */
async function compactSharedStrings(pkg: XlsxPackage) {
  const path = "xl/sharedStrings.xml";
  if (!pkg.has(path)) return;
  const sst = await pkg.xml(path);
  const items = children(sst.documentElement!, "si");
  const remap = new Map<number, number>();
  const sheetDocs = await Promise.all((await sheetPaths(pkg)).map((p) => pkg.xml(p)));
  let total = 0;
  for (const doc of sheetDocs) {
    const cells = doc.getElementsByTagName("c");
    for (let i = 0; i < cells.length; i++) {
      const c = cells.item(i)!;
      if (c.getAttribute("t") !== "s") continue;
      const v = child(c, "v");
      const old = Number(v?.textContent);
      if (!v || !items[old]) continue;
      if (!remap.has(old)) remap.set(old, remap.size);
      v.textContent = String(remap.get(old));
      total++;
    }
  }
  const root = sst.documentElement!;
  for (const si of items) root.removeChild(si);
  for (const [old] of [...remap].sort((a, b) => a[1] - b[1])) root.appendChild(items[old]!);
  root.setAttribute("count", String(total));
  root.setAttribute("uniqueCount", String(remap.size));
  pkg.touch(path);
  if (remap.size === 0) {
    await removePart(pkg, path);
    await removeRels(pkg, "xl/workbook.xml", (type) => type === REL_SHARED_STRINGS);
  }
}

async function scrubMetadata(pkg: XlsxPackage) {
  if (!pkg.has("docProps/core.xml")) return;
  const core = await pkg.xml("docProps/core.xml");
  const root = core.documentElement!;
  for (const name of ["dc:creator", "cp:lastModifiedBy", "dc:title", "dc:subject", "dc:description"]) {
    const el = core.getElementsByTagName(name).item(0);
    if (el) el.textContent = "";
  }
  for (const name of ["cp:lastPrinted", "cp:keywords"]) {
    const el = core.getElementsByTagName(name).item(0);
    if (el) root.removeChild(el);
  }
  for (const name of ["dcterms:created", "dcterms:modified"]) {
    const el = core.getElementsByTagName(name).item(0);
    if (el) el.textContent = "2026-01-01T00:00:00Z";
  }
  pkg.touch("docProps/core.xml");
  // Threaded-comment authors: keep the (required) part, drop any people in it.
  if (pkg.has("xl/persons/person.xml")) {
    const p = await pkg.xml("xl/persons/person.xml");
    for (const person of children(p.documentElement!, "person")) p.documentElement!.removeChild(person);
    pkg.touch("xl/persons/person.xml");
  }
}

async function dropPrinterSettings(pkg: XlsxPackage, sheetPath: string) {
  const targets = (await pkg.rels(sheetPath)).filter((r) => r.type === REL_PRINTER);
  if (!targets.length) return;
  const doc = await pkg.xml(sheetPath);
  const ps = child(doc.documentElement!, "pageSetup");
  if (ps) ps.removeAttributeNS(NS_REL, "id");
  for (const r of targets) await removePart(pkg, XlsxPackage.resolveTarget(sheetPath, r.target));
  await removeRels(pkg, sheetPath, (type) => type === REL_PRINTER);
  pkg.touch(sheetPath);
}

async function removeRels(pkg: XlsxPackage, partPath: string, match: (type: string) => boolean) {
  const p = XlsxPackage.relsPathFor(partPath);
  if (!pkg.has(p)) return;
  const doc = await pkg.xml(p);
  for (const r of children(doc.documentElement!, "Relationship")) if (match(r.getAttribute("Type") ?? "")) doc.documentElement!.removeChild(r);
  pkg.touch(p);
}

/** Delete a part, its own relationships part, and its content-type override. */
async function removePart(pkg: XlsxPackage, path: string) {
  if (pkg.has(path)) pkg.remove(path);
  const rels = XlsxPackage.relsPathFor(path);
  if (pkg.has(rels)) pkg.remove(rels);
  const ct = await pkg.xml("[Content_Types].xml");
  for (const o of children(ct.documentElement!, "Override")) if (o.getAttribute("PartName") === `/${path}`) ct.documentElement!.removeChild(o);
  pkg.touch("[Content_Types].xml");
}
