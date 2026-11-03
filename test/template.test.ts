import { describe, expect, it } from "vitest";
import ExcelJS from "exceljs";
import JSZip from "jszip";
import { existsSync } from "node:fs";
import { readFile } from "node:fs/promises";
import path from "node:path";
import { XlsxPackage } from "../src/excel/ooxml.js";
import { readTracker } from "../src/excel/tracker.js";
import { makeTemplate, SUBJECT_SLOTS } from "../src/excel/template.js";
import { syncPackage } from "../src/sync/run.js";
import type { MoodleItem } from "../src/moodle/types.js";
import { fixtureBytes } from "./helpers.js";

const COMMITTED = path.resolve("web/public/tracker-template.xlsx");
const REAL = path.resolve("ER assignments.xlsx");

/** Strings that would identify the workbook's owner or their courses. */
async function personalStrings(bytes: Uint8Array): Promise<string[]> {
  const t = await readTracker(await XlsxPackage.load(bytes));
  const fromRows = t.rows.filter((r) => !r.isEmpty).flatMap((r) => [r.assignment, r.moodleKey, r.linkUrl ?? "", r.linkText]);
  // The owner's name, as Excel recorded it in the file properties (kept out of the source on purpose).
  const core = (await (await JSZip.loadAsync(bytes)).file("docProps/core.xml")?.async("string")) ?? "";
  const people = [...core.matchAll(/<(?:dc:creator|cp:lastModifiedBy)>([^<]+)</g)].flatMap((m) => m[1]!.split(/\s+/)).filter((w) => w !== "openpyxl");
  const out = [...fromRows, ...t.subjects, ...people, "/mod/assign/view.php", "/mod/quiz/view.php", "Number of absences"];
  return [...new Set(out.map((s) => s.trim()).filter((s) => s.length >= 4))];
}

async function allText(bytes: Uint8Array): Promise<string> {
  const zip = await JSZip.loadAsync(bytes);
  const parts = Object.values(zip.files).filter((f) => !f.dir && /\.(xml|rels)$/.test(f.name));
  return (await Promise.all(parts.map((f) => f.async("string")))).join("\n");
}

function item(n: number, p: Partial<MoodleItem> = {}): MoodleItem {
  return {
    key: `assign:${n}`, kind: "assign", cmid: n, courseId: n, courseCode: String(10000 + n), courseName: `Course ${n}`,
    name: `Task ${n}`, due: "2026-11-20", status: "not_submitted", gradePercent: null,
    url: `https://moodle.example/mod/assign/view.php?id=${n}`, ...p,
  };
}

async function checkTemplate(tpl: Uint8Array, sources: Uint8Array[]) {
  const text = await allText(tpl);
  for (const src of sources) for (const s of await personalStrings(src)) expect(text, `template still contains "${s}"`).not.toContain(s);

  const zip = await JSZip.loadAsync(tpl);
  expect(Object.keys(zip.files).filter((n) => /printerSettings|calcChain/.test(n))).toEqual([]);
  const t = await readTracker(await XlsxPackage.load(tpl));
  expect(t.rows.every((r) => r.isEmpty)).toBe(true);
  expect(t.subjects).toEqual([]);
  expect(`${t.table.headerRow}:${t.table.lastRow}`).toBe("3:200");
  expect(t.table.cols).toMatchObject({ moodleKey: "J", moodleStatus: "K", lastSynced: "L" });

  const s3 = await zip.file("xl/worksheets/sheet3.xml")!.async("string");
  expect(s3.match(/<x14:dataValidation /g)).toHaveLength(4);
  for (const c of "BCEF") expect(s3).toContain(`<xm:sqref>${c}4:${c}200</xm:sqref>`);
  expect(s3).toContain(`<xm:f>Lists!$A$4:$A$${3 + SUBJECT_SLOTS}</xm:f>`);
  expect(s3).not.toContain("<hyperlink ");

  const chart = await zip.file("xl/charts/chart1.xml")!.async("string");
  expect(chart).toContain(`Dashboard!$A$12:$A$${11 + SUBJECT_SLOTS}`);
  expect(chart).toContain(`Dashboard!$B$12:$B$${11 + SUBJECT_SLOTS}`);
  expect(chart).not.toMatch(/<c:(str|num)Cache>/);

  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(tpl as unknown as ArrayBuffer);
  expect(wb.worksheets.map((w) => w.name)).toEqual(["How to Use", "Dashboard", "Assignments", "Lists"]);
  expect(wb.getWorksheet("Dashboard")!.getCell(`B${11 + SUBJECT_SLOTS}`).formula).toBe(`COUNTIF(Table1[Subject],A${11 + SUBJECT_SLOTS})`);
  expect(wb.getWorksheet("Assignments")!.getCell("H4").formula).toMatch(/D4/);
}

describe("makeTemplate", async () => {
  const src = await fixtureBytes();
  const tpl = await makeTemplate(src);

  it("removes every trace of the owner's data and keeps the tracker's structure", async () => {
    await checkTemplate(tpl, [src]);
  });

  it(`takes up to ${SUBJECT_SLOTS} courses into the dropdown and chart, then warns`, async () => {
    const pkg = await XlsxPackage.load(tpl);
    const items = Array.from({ length: SUBJECT_SLOTS + 1 }, (_, i) => item(i + 1));
    const out = await syncPackage(pkg, items, { today: "2026-11-03" });
    expect(out.applied.added).toBe(SUBJECT_SLOTS + 1);
    expect(out.applied.subjectsAdded).toHaveLength(SUBJECT_SLOTS);
    expect(out.applied.warnings).toHaveLength(1);
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load((await pkg.save()) as unknown as ArrayBuffer);
    expect(wb.getWorksheet("Dashboard")!.getCell(`A${11 + SUBJECT_SLOTS}`).value).toBe(`${10000 + SUBJECT_SLOTS} - Course ${SUBJECT_SLOTS}`);
  });
});

describe.runIf(existsSync(COMMITTED))("committed web template", async () => {
  it("is blank, against both the original and the current workbook", async () => {
    const sources = [await fixtureBytes(), ...(existsSync(REAL) ? [await readFile(REAL)] : [])];
    await checkTemplate(await readFile(COMMITTED), sources);
  });
});
