// One-time, idempotent upgrades to the workbook layout so the sync has room to work:
//   1. add "Moodle Key", "Moodle Status", "Last Synced" columns to the table
//   2. grow the table to MIN_LAST_ROW pre-formatted rows (with the Days Left formula)
//   3. stretch conditional formatting / dropdown ranges to the new last row
//   4. point Dashboard formulas at the table (structured refs) so they grow with it —
//      this also fixes the drifted By Subject ranges (B14:B18 counted B5:B45, B6:B46, ...)

import { XlsxPackage, Sheet, NS_MAIN, child, children, colToIndex, indexToCol, splitRef } from "./ooxml.js";
import { COL, SHEET_ASSIGNMENTS, SHEET_DASHBOARD, SYNC_COLUMNS, TABLE_NAME, readTable, type ColKey } from "./tracker.js";

export const MIN_LAST_ROW = 200;

const NEW_COLUMN_WIDTH: Partial<Record<ColKey, number>> = { moodleKey: 16, moodleStatus: 15, lastSynced: 13 };

export async function migrate(pkg: XlsxPackage): Promise<string[]> {
  const log: string[] = [];
  const sheet = await Sheet.open(pkg, SHEET_ASSIGNMENTS);
  let table = await readTable(pkg, sheet.path);

  // ---- 1. sync columns ----
  const missing = SYNC_COLUMNS.filter((k) => !table.cols[k]);
  if (missing.length) {
    const tdoc = await pkg.xml(table.path);
    const troot = tdoc.documentElement!;
    const tcols = child(troot, "tableColumns")!;
    const ids = children(tcols, "tableColumn").map((c) => Number(c.getAttribute("id")));
    let nextId = Math.max(...ids) + 1;
    const [start, end] = troot.getAttribute("ref")!.split(":") as [string, string];
    let lastCol = colToIndex(splitRef(end).col);
    const headerStyle = sheet.read(`${table.cols.subject}${table.headerRow}`).style;
    for (const key of missing) {
      lastCol++;
      const tc = tdoc.createElementNS(NS_MAIN, "tableColumn");
      tc.setAttribute("id", String(nextId++));
      tc.setAttribute("name", COL[key]);
      tcols.appendChild(tc);
      sheet.write(`${indexToCol(lastCol)}${table.headerRow}`, COL[key], { style: headerStyle });
      setColWidth(sheet, lastCol, NEW_COLUMN_WIDTH[key] ?? 14);
      log.push(`Added column ${indexToCol(lastCol)} "${COL[key]}" to ${TABLE_NAME}`);
    }
    tcols.setAttribute("count", String(children(tcols, "tableColumn").length));
    const ref = `${start}:${indexToCol(lastCol)}${splitRef(end).row}`;
    troot.setAttribute("ref", ref);
    child(troot, "autoFilter")?.setAttribute("ref", ref);
    pkg.touch(table.path);
    table = await readTable(pkg, sheet.path);
  }

  // ---- 2. styles for every table column, taken from an existing body row ----
  const tmplRow = Math.max(table.firstRow, Math.min(table.lastRow - 2, 40));
  const styleOf = (key: ColKey): string | undefined => {
    if (key === "moodleKey" || key === "moodleStatus") return sheet.read(`${table.cols.type}${tmplRow}`).style;
    if (key === "lastSynced") return sheet.read(`${table.cols.due}${tmplRow}`).style;
    return sheet.read(`${table.cols[key]}${tmplRow}`).style;
  };
  const allKeys = Object.keys(table.cols) as ColKey[];

  // Give the new columns' existing body cells the right style.
  if (missing.length) {
    for (let r = table.firstRow; r <= table.lastRow; r++) {
      for (const key of missing) sheet.write(`${table.cols[key]}${r}`, null, { style: styleOf(key) });
    }
  }

  // ---- 3. grow the table ----
  const daysCol = table.cols.daysLeft!;
  const master = masterFormula(sheet, daysCol, table.firstRow);
  if (table.lastRow < MIN_LAST_ROW) {
    // Rows just below the table (43–45 in the original) hold stray copies of the formula — replaced here.
    const from = table.lastRow + 1;
    for (let r = from; r <= MIN_LAST_ROW; r++) {
      for (const key of allKeys) {
        const ref = `${table.cols[key]}${r}`;
        const c = sheet.cellEl(ref);
        if (c && child(c, "f")) c.removeChild(child(c, "f")!);
        if (key === "daysLeft") sheet.writeFormula(ref, formulaForRow(master, table.firstRow, r), { style: styleOf(key), volatile: true });
        else sheet.write(ref, null, { style: styleOf(key) });
      }
    }
    // The original last table row had an older Days Left formula; make it consistent.
    sheet.writeFormula(`${daysCol}${table.lastRow}`, formulaForRow(master, table.firstRow, table.lastRow), {
      style: styleOf("daysLeft"),
      volatile: true,
    });
    const tdoc = await pkg.xml(table.path);
    const troot = tdoc.documentElement!;
    // Rows you add by hand in Excel get the same formula as the rest of the column.
    const daysTc = children(child(troot, "tableColumns")!, "tableColumn").find((c) => c.getAttribute("name") === COL.daysLeft);
    const calc = daysTc && child(daysTc, "calculatedColumnFormula");
    if (calc) calc.textContent = master;
    const [start, end] = troot.getAttribute("ref")!.split(":") as [string, string];
    const ref = `${start}:${splitRef(end).col}${MIN_LAST_ROW}`;
    troot.setAttribute("ref", ref);
    child(troot, "autoFilter")?.setAttribute("ref", ref);
    pkg.touch(table.path);
    log.push(`Grew ${TABLE_NAME} to row ${MIN_LAST_ROW}`);
    table = await readTable(pkg, sheet.path);
  }

  // ---- 4. stretch conditional formatting and dropdown ranges ----
  const lastRow = table.lastRow;
  const root = sheet.doc.documentElement!;
  let stretched = 0;
  for (const cf of children(root, "conditionalFormatting")) {
    const s = stretchSqref(cf.getAttribute("sqref")!, table.firstRow, lastRow);
    if (s !== cf.getAttribute("sqref")) (cf.setAttribute("sqref", s), stretched++);
  }
  for (const dv of children(child(root, "dataValidations") ?? root, "dataValidation")) {
    const s = stretchSqref(dv.getAttribute("sqref")!, table.firstRow, lastRow);
    if (s !== dv.getAttribute("sqref")) (dv.setAttribute("sqref", s), stretched++);
  }
  const xmSqrefs = sheet.doc.getElementsByTagName("xm:sqref");
  for (let i = 0; i < xmSqrefs.length; i++) {
    const el = xmSqrefs.item(i)!;
    const s = stretchSqref(el.textContent ?? "", table.firstRow, lastRow);
    if (s !== el.textContent) {
      el.textContent = s;
      stretched++;
    }
  }
  if (stretched) {
    sheet.touch();
    log.push(`Extended ${stretched} formatting/dropdown ranges to row ${lastRow}`);
  }
  const dim = child(root, "dimension");
  if (dim) {
    const lastColLetter = Object.values(table.cols).sort((a, b) => colToIndex(b!) - colToIndex(a!))[0]!;
    dim.setAttribute("ref", `A1:${lastColLetter}${lastRow}`);
  }

  // ---- 5. Dashboard formulas -> structured references ----
  const dash = await Sheet.open(pkg, SHEET_DASHBOARD);
  const headerByCol = new Map(Object.entries(table.cols).map(([k, col]) => [col!, COL[k as ColKey]]));
  let rewired = 0;
  for (const row of children(dash.sheetData, "row")) {
    for (const c of children(row, "c")) {
      const f = child(c, "f");
      if (!f?.textContent) continue;
      const before = f.textContent;
      const after = before.replace(/Assignments!\$?([A-Z]+)\$?\d+:\$?\1\$?\d+/g, (m, col: string) => {
        const h = headerByCol.get(col);
        return h ? `${TABLE_NAME}[${h}]` : m;
      });
      if (after !== before) {
        f.textContent = after;
        const v = child(c, "v");
        if (v) c.removeChild(v);
        rewired++;
      }
    }
  }
  if (rewired) {
    dash.touch();
    log.push(`Pointed ${rewired} Dashboard formulas at ${TABLE_NAME} columns`);
  }

  if (log.length) await pkg.requestFullRecalc();
  return log;
}

/** The Days Left formula text as written for the first body row (resolving shared formulas). */
function masterFormula(sheet: Sheet, col: string, firstRow: number): string {
  const f = child(sheet.cellEl(`${col}${firstRow}`)!, "f");
  const text = f?.textContent;
  if (!text) throw new Error(`No Days Left formula in ${col}${firstRow}`);
  return text;
}

/** Re-target a row-relative formula written for `fromRow` to `toRow` (only same-row refs like D4). */
export function formulaForRow(formula: string, fromRow: number, toRow: number): string {
  return formula.replace(new RegExp(`(?<![A-Za-z0-9_$])([A-Z]{1,3})${fromRow}(?!\\d)`, "g"), `$1${toRow}`);
}

/** Make every range in a space-separated sqref that lives in the table body end at lastRow. */
export function stretchSqref(sqref: string, firstRow: number, lastRow: number): string {
  const out: string[] = [];
  const seenCols = new Set<string>();
  for (const tok of sqref.trim().split(/\s+/)) {
    const [a, b = a] = tok.split(":") as [string, string?];
    const s = splitRef(a);
    const e = splitRef(b!);
    const inBody = s.row >= firstRow && s.row <= firstRow + 50 && e.row < 1_000_000;
    if (!inBody || s.col !== e.col) {
      out.push(tok);
      continue;
    }
    if (seenCols.has(s.col)) continue; // "H4:H42 H44:H45" collapses into one range
    seenCols.add(s.col);
    out.push(`${s.col}${firstRow}:${s.col}${Math.max(lastRow, e.row)}`);
  }
  return out.join(" ");
}

function setColWidth(sheet: Sheet, colIndex: number, width: number) {
  const root = sheet.doc.documentElement!;
  let cols = child(root, "cols");
  if (!cols) {
    cols = sheet.doc.createElementNS(NS_MAIN, "cols");
    root.insertBefore(cols, child(root, "sheetData")!);
  }
  if (children(cols, "col").some((c) => Number(c.getAttribute("min")) <= colIndex && Number(c.getAttribute("max")) >= colIndex)) return;
  const c = sheet.doc.createElementNS(NS_MAIN, "col");
  c.setAttribute("min", String(colIndex));
  c.setAttribute("max", String(colIndex));
  c.setAttribute("width", String(width));
  c.setAttribute("customWidth", "1");
  cols.appendChild(c);
  sheet.touch();
}
