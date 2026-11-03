import { describe, expect, it } from "vitest";
import ExcelJS from "exceljs";
import JSZip from "jszip";
import { mkdtemp, readdir, readFile, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { XlsxPackage } from "../src/excel/ooxml.js";
import { readTracker } from "../src/excel/tracker.js";
import { migrate, formulaForRow, stretchSqref } from "../src/excel/migrate.js";
import { applyPlan } from "../src/excel/write.js";
import { saveWorkbook } from "../src/excel/file.js";
import { buildPlan } from "../src/sync/plan.js";
import type { MoodleItem } from "../src/moodle/types.js";
import { fixtureBytes } from "./helpers.js";

const BASE = "https://moodle.huji.ac.il/2026-27";

async function migrated() {
  const orig = await fixtureBytes();
  const pkg = await XlsxPackage.load(orig);
  const log = await migrate(pkg);
  return { orig, bytes: await pkg.save(), log };
}

const coreFields = (rows: Awaited<ReturnType<typeof readTracker>>["rows"]) =>
  rows.filter((r) => !r.isEmpty).map(({ row, assignment, subject, type, due, priority, status, grade, linkText, linkUrl }) =>
    ({ row, assignment, subject, type, due, priority, status, grade, linkText, linkUrl }));

describe("helpers", () => {
  it("retargets row-relative formulas", () => {
    expect(formulaForRow('IF(D4="",F4,D4-TODAY())', 4, 57)).toBe('IF(D57="",F57,D57-TODAY())');
  });
  it("stretches table-body sqrefs and leaves whole-column ones", () => {
    expect(stretchSqref("H4:H42 H44:H45", 4, 200)).toBe("H4:H200");
    expect(stretchSqref("H1:H1048576", 4, 200)).toBe("H1:H1048576");
    expect(stretchSqref("B4:B44", 4, 200)).toBe("B4:B200");
  });
});

describe("migrate", async () => {
  const { orig, bytes, log } = await migrated();
  const a = await JSZip.loadAsync(orig);
  const b = await JSZip.loadAsync(bytes);

  it("reports what it changed", () => {
    expect(log.join("\n")).toMatch(/Moodle Key/);
    expect(log.join("\n")).toMatch(/Grew Table1 to row 200/);
  });

  it("keeps chart, drawing, persons and printer settings byte-identical", async () => {
    for (const p of ["xl/charts/chart1.xml", "xl/drawings/drawing1.xml", "xl/persons/person.xml", "xl/styles.xml", "xl/sharedStrings.xml", "xl/worksheets/sheet1.xml"]) {
      expect(await b.file(p)!.async("string"), p).toBe(await a.file(p)!.async("string"));
    }
    expect(b.file("xl/calcChain.xml")).toBeNull(); // dropped on purpose; Excel rebuilds it
  });

  it("keeps all four dropdowns and stretches them to the new last row", async () => {
    const s3 = await b.file("xl/worksheets/sheet3.xml")!.async("string");
    expect(s3.match(/<x14:dataValidation /g)).toHaveLength(4);
    for (const c of "BCEF") expect(s3).toContain(`<xm:sqref>${c}4:${c}200</xm:sqref>`);
    expect(s3).toContain('sqref="F4:F200"');
  });

  it("leaves your existing data unchanged", async () => {
    const before = await readTracker(await XlsxPackage.load(orig));
    const after = await readTracker(await XlsxPackage.load(bytes));
    expect(coreFields(after.rows)).toEqual(coreFields(before.rows));
    expect(after.table.lastRow).toBe(200);
    expect(after.table.cols).toMatchObject({ moodleKey: "J", moodleStatus: "K", lastSynced: "L" });
  });

  it("rewires Dashboard formulas to the table", async () => {
    const dash = await b.file("xl/worksheets/sheet2.xml")!.async("string");
    expect(dash).not.toMatch(/Assignments!/);
    expect(dash).toContain("COUNTIF(Table1[Subject],A14)");
    expect(dash).toContain("COUNTA(Table1[Assignment])");
  });

  it("is idempotent", async () => {
    const pkg = await XlsxPackage.load(bytes);
    expect(await migrate(pkg)).toEqual([]);
  });

  it("is readable by ExcelJS (independent parser)", async () => {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(bytes as unknown as ArrayBuffer);
    const ws = wb.getWorksheet("Assignments")!;
    expect(ws.getCell("J3").value).toBe("Moodle Key");
    expect(ws.getCell("A7").value).toBe("Week 1 - תרגיל 1");
  });
});

describe("applyPlan", async () => {
  const { bytes } = await migrated();
  const pkg = await XlsxPackage.load(bytes);
  const t = await readTracker(pkg);

  const existing: MoodleItem = {
    key: "assign:154886", kind: "assign", cmid: 154886, courseId: 1, courseCode: "57121", courseName: "חדו\"א",
    name: "תרגיל בית 1", due: "2026-10-22", status: "graded", gradePercent: 100, url: `${BASE}/mod/assign/view.php?id=154886`,
  };
  const fresh: MoodleItem = {
    key: "quiz:200001", kind: "quiz", cmid: 200001, courseId: 2, courseCode: "51102", courseName: "חשיבה סטטיסטית",
    name: "תרגיל 4", due: "2026-11-12", status: "not_submitted", gradePercent: null, url: `${BASE}/mod/quiz/view.php?id=200001`,
  };
  const plan = buildPlan(t.rows, [existing, fresh], { subjects: t.subjects, today: "2026-11-01" });
  const res = await applyPlan(pkg, plan);
  const out = await pkg.save();
  const t2 = await readTracker(await XlsxPackage.load(out));

  it("adds the new item into the first empty row", () => {
    expect(res.added).toBe(1);
    const r = t2.rows.find((x) => x.row === res.addedRows[0])!;
    expect(r.row).toBe(17);
    expect(r).toMatchObject({ assignment: "תרגיל 4", type: "תרגיל בית להגשה", due: "2026-11-12", status: "Not Started", moodleKey: "quiz:200001" });
    expect(r.subject).toMatch(/^51102 - /);
    expect(r.linkUrl).toBe(fresh.url);
  });

  it("updates the matched row, taking its name from Moodle", () => {
    const before = t.rows.find((x) => x.row === 7)!;
    const r7 = t2.rows.find((x) => x.row === 7)!;
    expect(r7).toMatchObject({ assignment: "תרגיל בית 1", status: "Graded", grade: 100, moodleKey: "assign:154886", due: "2026-10-22" });
    expect({ subject: r7.subject, priority: r7.priority }).toEqual({ subject: before.subject, priority: before.priority });
  });

  it("keeps the Days Left formula in written rows", async () => {
    const s3 = await (await JSZip.loadAsync(out)).file("xl/worksheets/sheet3.xml")!.async("string");
    expect(s3).toMatch(/<c r="H17"[^>]*><f[^>]*\/>/); // shared formula reference still there
  });

  it("produces no further changes when run again", async () => {
    const again = buildPlan(t2.rows, [existing, fresh], { subjects: t2.subjects, today: "2026-11-02" });
    expect(again.actions.map((a) => a.kind)).toEqual(["unchanged", "unchanged"]);
  });

  it("ExcelJS sees the date and hyperlink", async () => {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(out as unknown as ArrayBuffer);
    const ws = wb.getWorksheet("Assignments")!;
    expect((ws.getCell("D17").value as Date).toISOString().slice(0, 10)).toBe("2026-11-12");
    expect(ws.getCell("I17").hyperlink).toBe(fresh.url);
  });
});

describe("new subjects", async () => {
  const { bytes } = await migrated();
  const pkg = await XlsxPackage.load(bytes);
  const t = await readTracker(pkg);
  const r: MoodleItem = {
    key: "quiz:92518", kind: "quiz", cmid: 92518, courseId: 9, courseCode: "51033", courseName: "תכנות R",
    name: "יחידה 2", due: "2026-11-05", status: "not_submitted", gradePercent: null, url: `${BASE}/mod/quiz/view.php?id=92518`,
  };
  const plan = buildPlan(t.rows, [r], { subjects: t.subjects, today: "2026-11-01" });
  const res = await applyPlan(pkg, plan);
  const out = await pkg.save();

  it("lands in the dropdown range and the Dashboard's free By-Subject row", async () => {
    expect(res.subjectsAdded).toEqual(["51033 - תכנות R"]);
    const t2 = await readTracker(await XlsxPackage.load(out));
    expect(t2.subjects).toContain("51033 - תכנות R");
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(out as unknown as ArrayBuffer);
    expect(wb.getWorksheet("Lists")!.getCell("A10").value).toBe("51033 - תכנות R");
    expect(wb.getWorksheet("Dashboard")!.getCell("A18").value).toBe("51033 - תכנות R");
  });

  it("refuses rather than overflowing the dropdown range", async () => {
    const pkg2 = await XlsxPackage.load(out);
    const t3 = await readTracker(pkg2);
    const other = { ...r, key: "quiz:1", courseCode: "99999", courseName: "Other" };
    const plan2 = buildPlan(t3.rows, [other], { subjects: t3.subjects, today: "2026-11-01" });
    await expect(applyPlan(pkg2, plan2)).rejects.toThrow(/No room for subject/);
  });
});

describe("saveWorkbook", () => {
  it("backs up before replacing", async () => {
    const dir = await mkdtemp(path.join(tmpdir(), "tracker-"));
    const file = path.join(dir, "book.xlsx");
    const orig = await fixtureBytes();
    await writeFile(file, orig);
    const { bytes } = await migrated();
    const bak = await saveWorkbook(file, bytes, path.join(dir, "backups"));
    expect(await readFile(bak)).toEqual(orig);
    expect(await readFile(file)).toEqual(bytes);
    expect(await readdir(dir)).not.toContain("book.xlsx.tmp");
  });

  it("refuses when Excel's lock file is present", async () => {
    const dir = await mkdtemp(path.join(tmpdir(), "tracker-"));
    const file = path.join(dir, "book.xlsx");
    await writeFile(file, await fixtureBytes());
    await writeFile(path.join(dir, "~$book.xlsx"), "");
    await expect(saveWorkbook(file, Buffer.from(""), dir)).rejects.toThrow(/open in Excel/);
  });
});
