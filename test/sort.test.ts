import { describe, expect, it } from "vitest";
import ExcelJS from "exceljs";
import { XlsxPackage } from "../src/excel/ooxml.js";
import { readTracker, type TrackerRow } from "../src/excel/tracker.js";
import { migrate } from "../src/excel/migrate.js";
import { sortOrder, sortTable } from "../src/excel/sort.js";
import { fixtureBytes } from "./helpers.js";

function row(r: number, p: Partial<TrackerRow>): TrackerRow {
  return {
    row: r, assignment: "x", subject: "", type: "", due: null, priority: "", status: "", grade: null,
    linkText: "", linkUrl: null, moodleKey: "", moodleStatus: "", isEmpty: false, ...p,
  };
}

describe("sortOrder", () => {
  it("puts pending by due date first, then Submitted, then Graded, then empty rows", () => {
    const rows = [
      row(4, { status: "Graded", due: "2026-10-01" }),
      row(5, { isEmpty: true, assignment: "" }),
      row(6, { status: "Not Started", due: "2026-11-20" }),
      row(7, { status: "Submitted", due: "2026-11-03" }),
      row(8, { status: "In Progress", due: "2026-10-30" }), // overdue: most urgent
      row(9, { status: "Not Started", due: null }),
      row(10, { status: "Submitted", due: "2026-10-15" }),
      row(11, { status: "Not Started", due: "2026-11-03" }),
    ];
    expect(sortOrder(rows)).toEqual([8, 11, 6, 9, 10, 7, 4, 5]);
  });
});

describe("sortTable on the real workbook", async () => {
  const pkg = await XlsxPackage.load(await fixtureBytes());
  await migrate(pkg);
  const before = await readTracker(pkg);
  const moved = await sortTable(pkg);
  const out = await pkg.save();
  const after = await readTracker(await XlsxPackage.load(out));
  const filled = (rs: TrackerRow[]) => rs.filter((r) => !r.isEmpty);

  it("reorders rows and keeps every row's data together, link included", () => {
    expect(moved).toBeGreaterThan(0);
    const strip = ({ row: _r, ...rest }: TrackerRow) => rest;
    const sorted = sortOrder(before.rows).map((n) => strip(before.rows.find((r) => r.row === n)!));
    expect(after.rows.map(strip)).toEqual(sorted);
    expect(filled(after.rows).map((r) => r.row)).toEqual(filled(after.rows).map((_, i) => after.table.firstRow + i));
  });

  it("leaves the Days Left formula pointing at its own row", async () => {
    const sheet = after.sheet;
    for (const r of filled(after.rows)) expect(sheet.read(`H${r.row}`).formula ?? "").not.toBe("");
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(out as unknown as ArrayBuffer);
    const f = wb.getWorksheet("Assignments")!.getCell("H10").formula ?? "";
    if (f) expect(f).toContain("D10");
  });

  it("is a no-op when already sorted", async () => {
    expect(await sortTable(await XlsxPackage.load(out))).toBe(0);
  });
});
