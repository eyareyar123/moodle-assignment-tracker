// The Assignments table of the tracker workbook: where it lives and how to read it.

import { XlsxPackage, Sheet, child, children, indexToCol, splitRef, serialToIsoDate } from "./ooxml.js";

export const SHEET_ASSIGNMENTS = "Assignments";
export const SHEET_LISTS = "Lists";
export const SHEET_DASHBOARD = "Dashboard";
export const TABLE_NAME = "Table1";

/** Table column headers, keyed by the name the code uses. */
export const COL = {
  assignment: "Assignment",
  subject: "Subject",
  type: "Type",
  due: "Due Date",
  priority: "Priority",
  status: "Status",
  grade: "Grade (%)",
  daysLeft: "Days Left",
  link: "link",
  moodleKey: "Moodle Key",
  moodleStatus: "Moodle Status",
  lastSynced: "Last Synced",
} as const;
export type ColKey = keyof typeof COL;

/** Columns the sync adds to the table (in this order) if they are missing. */
export const SYNC_COLUMNS: ColKey[] = ["moodleKey", "moodleStatus", "lastSynced"];

export const STATUSES = ["Not Started", "In Progress", "Submitted", "Graded"] as const;
export type TrackerStatus = (typeof STATUSES)[number];

export interface TrackerRow {
  row: number;
  assignment: string;
  subject: string;
  type: string;
  /** ISO date YYYY-MM-DD, or null when blank / not a date. */
  due: string | null;
  priority: string;
  status: string;
  grade: string | number | null;
  linkText: string;
  /** Hyperlink target if the link cell has one, else the cell text when it looks like a URL. */
  linkUrl: string | null;
  moodleKey: string;
  moodleStatus: string;
  isEmpty: boolean;
}

export interface TableInfo {
  path: string;
  headerRow: number;
  firstRow: number;
  lastRow: number;
  /** Column letter of every table column, in order. */
  columns: string[];
  /** Column letter for each header present in the table. */
  cols: Partial<Record<ColKey, string>>;
  headers: string[];
}

export interface Tracker {
  sheet: Sheet;
  table: TableInfo;
  rows: TrackerRow[];
  /** Subject labels from the Lists sheet (e.g. "57340 - סטטיסטיקה לכלכלה א"). */
  subjects: string[];
}

export async function readTable(pkg: XlsxPackage, sheetPath: string): Promise<TableInfo> {
  const path = await pkg.tablePath(sheetPath, TABLE_NAME);
  const t = (await pkg.xml(path)).documentElement!;
  const [start, end] = t.getAttribute("ref")!.split(":") as [string, string];
  const s = splitRef(start);
  const e = splitRef(end);
  const headers = children(child(t, "tableColumns")!, "tableColumn").map((c) => c.getAttribute("name")!);
  const startCol = s.col.charCodeAt(0) - 64;
  const cols: Partial<Record<ColKey, string>> = {};
  for (const [key, name] of Object.entries(COL) as [ColKey, string][]) {
    const i = headers.indexOf(name);
    if (i >= 0) cols[key] = indexToCol(startCol + i);
  }
  const columns = headers.map((_, i) => indexToCol(startCol + i));
  return { path, headerRow: s.row, firstRow: s.row + 1, lastRow: e.row, columns, cols, headers };
}

const str = (v: unknown) => (v === null || v === undefined ? "" : String(v).trim());

export async function readTracker(pkg: XlsxPackage): Promise<Tracker> {
  const sheet = await Sheet.open(pkg, SHEET_ASSIGNMENTS);
  const table = await readTable(pkg, sheet.path);
  const links = await sheet.hyperlinks();
  const get = (key: ColKey, r: number) => {
    const col = table.cols[key];
    return col ? sheet.read(`${col}${r}`).value : null;
  };

  const rows: TrackerRow[] = [];
  for (let r = table.firstRow; r <= table.lastRow; r++) {
    const dueRaw = get("due", r);
    const linkText = str(get("link", r));
    const linkCol = table.cols.link;
    const hl = linkCol ? links.get(`${linkCol}${r}`)?.target : undefined;
    const row: TrackerRow = {
      row: r,
      assignment: str(get("assignment", r)),
      subject: str(get("subject", r)),
      type: str(get("type", r)),
      due: typeof dueRaw === "number" ? serialToIsoDate(dueRaw) : null,
      priority: str(get("priority", r)),
      status: str(get("status", r)),
      grade: (get("grade", r) as string | number | null) ?? null,
      linkText,
      linkUrl: hl ?? (/^https?:\/\//i.test(linkText) ? linkText : null),
      moodleKey: str(get("moodleKey", r)),
      moodleStatus: str(get("moodleStatus", r)),
      isEmpty: false,
    };
    row.isEmpty = !row.assignment && !row.subject && !row.due && !row.status && !row.linkText && !row.moodleKey;
    rows.push(row);
  }

  const lists = await Sheet.open(pkg, SHEET_LISTS);
  const subjects: string[] = [];
  for (let r = 4; r <= 100; r++) {
    const v = str(lists.read(`A${r}`).value);
    if (v) subjects.push(v);
  }
  return { sheet, table, rows, subjects };
}
