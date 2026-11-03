import { describe, expect, it } from "vitest";
import ExcelJS from "exceljs";
import { readFile } from "node:fs/promises";
import path from "node:path";
import { XlsxPackage, Sheet } from "../src/excel/ooxml.js";
import { readTracker } from "../src/excel/tracker.js";
import { academicYear, buildTracker, downloadName, siteForYear, UserError, yearChoices } from "../src/web/core.js";
import { FAKE_TOKEN, fakeFetch } from "./fake-moodle.js";
import { STRINGS, initialLang, t } from "../web/i18n.js";

const TEMPLATE = path.resolve("web/public/tracker-template.xlsx");
const SITE = "https://moodle.huji.ac.il/2026-27";
const NOW = new Date("2026-11-03T10:00:00Z");

const template = async () => new Uint8Array(await readFile(TEMPLATE));

describe("page text", () => {
  it("has every string in both languages, with the same placeholders", () => {
    const keys = Object.keys(STRINGS.en).sort();
    expect(Object.keys(STRINGS.he).sort()).toEqual(keys);
    const holes = (s: string) => [...s.matchAll(/\{(\w+)\}/g)].map((m) => m[1]).sort();
    for (const k of keys as (keyof typeof STRINGS.en)[]) {
      expect(holes(STRINGS.he[k]), k).toEqual(holes(STRINGS.en[k]));
      expect(STRINGS.he[k].trim(), k).not.toBe("");
    }
  });
  it("has a message for every error the page can raise", () => {
    for (const code of ["badUpload", "badTemplate", "rejected", "refused", "unreachable"]) expect(STRINGS.en).toHaveProperty(`err_${code}`);
  });
  it("starts in Hebrew for Hebrew browsers unless the visitor chose otherwise", () => {
    expect(initialLang(null, ["he-IL", "en-US"])).toBe("he");
    expect(initialLang(null, ["en-US"])).toBe("en");
    expect(initialLang("en", ["he-IL"])).toBe("en");
    expect(t("he", "stageFetched", { n: 5 })).toBe("נמצאו 5 פעילויות");
  });
});

describe("site and year helpers", () => {
  it("picks HUJI's academic-year site, switching in September", () => {
    expect(academicYear(new Date(2026, 10, 3))).toBe("2026-27");
    expect(academicYear(new Date(2027, 7, 31))).toBe("2026-27");
    expect(academicYear(new Date(2027, 8, 1))).toBe("2027-28");
    expect(academicYear(new Date(2099, 9, 1))).toBe("2099-00");
    expect(yearChoices(new Date(2026, 10, 3))).toEqual(["2026-27", "2025-26"]);
    expect(siteForYear("2026-27")).toBe(SITE);
    expect(() => siteForYear("../evil")).toThrow();
  });
  it("names the download", () => {
    expect(downloadName(undefined, "2026-11-03")).toBe("Moodle assignments 2026-11-03.xlsx");
    expect(downloadName("ER assignments.xlsx", "2026-11-03")).toBe("ER assignments.xlsx");
  });
});

describe("buildTracker", async () => {
  const calls: string[] = [];
  const first = await buildTracker({ siteUrl: SITE, token: FAKE_TOKEN, template: await template(), now: NOW, fetchImpl: fakeFetch(calls) });
  const t = await readTracker(await XlsxPackage.load(first.bytes));
  const filled = t.rows.filter((r) => !r.isEmpty);

  it("fills a new tracker with every dated activity of the year, sorted", () => {
    expect(first.user).toBe("Test Student");
    expect(first.undatedSkipped).toBe(1);
    expect(filled.map((r) => [r.assignment, r.status])).toEqual([
      ["הגשת קוד", "Not Started"],
      ["יחידה 2: dplyr ו-tidyr", "Not Started"],
      ["הגשת תרגיל 2", "Submitted"],
      ["הגשת תרגיל 1", "Graded"], // past due, still added: a new tracker takes the whole year
    ]);
    expect(filled.find((r) => r.moodleKey === "assign:1001")?.grade).toBe(90);
    expect(t.subjects.sort()).toEqual(["51033 - תכנות R", "57340 - סטטיסטיקה"]);
    expect(calls).not.toContain(""); // every request named a function
  });

  it("produces a file Excel-compatible readers accept, with links", async () => {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(first.bytes as unknown as ArrayBuffer);
    const ws = wb.getWorksheet("Assignments")!;
    expect(ws.getCell("I4").hyperlink).toBe(`${SITE}/mod/assign/view.php?id=2001`);
    expect(wb.getWorksheet("Dashboard")!.getCell("A12").value).toMatch(/^(51033|57340) - /);
  });

  it("updates an uploaded tracker and keeps what the visitor typed", async () => {
    const pkg = await XlsxPackage.load(first.bytes);
    const sheet = await Sheet.open(pkg, "Assignments");
    sheet.write("E4", "High"); // priority on "הגשת קוד"
    sheet.write("G5", "N\\A"); // grade cell typed by hand
    const second = await buildTracker({
      siteUrl: SITE, token: FAKE_TOKEN, template: await template(), existing: await pkg.save(), now: NOW, fetchImpl: fakeFetch(),
    });
    const t2 = await readTracker(await XlsxPackage.load(second.bytes));
    expect(second.outcome.applied.added).toBe(0);
    expect(t2.rows.find((r) => r.moodleKey === "assign:2001")?.priority).toBe("High");
    expect(t2.rows.find((r) => r.row === 5)?.grade).toBe("N\\A");
    expect(second.outcome.changed).toBe(false);
  });

  it("includes undated activities when asked", async () => {
    const r = await buildTracker({ siteUrl: SITE, token: FAKE_TOKEN, template: await template(), includeUndated: true, now: NOW, fetchImpl: fakeFetch() });
    expect(r.undatedSkipped).toBe(0);
    expect(r.outcome.applied.added).toBe(5);
  });

  it("explains a rejected link and an unusable upload in plain words", async () => {
    const bad = buildTracker({ siteUrl: SITE, token: "f".repeat(32), template: await template(), now: NOW, fetchImpl: fakeFetch() });
    await expect(bad).rejects.toBeInstanceOf(UserError);
    await expect(bad).rejects.toMatchObject({ code: "rejected" });
    const notTracker = buildTracker({ siteUrl: SITE, token: FAKE_TOKEN, template: await template(), existing: new TextEncoder().encode("hello"), fetchImpl: fakeFetch() });
    await expect(notTracker).rejects.toMatchObject({ code: "badUpload" });
  });

  it("reports an unreachable Moodle as a connection problem", async () => {
    const down = (async () => { throw new TypeError("Failed to fetch"); }) as unknown as typeof fetch;
    const r = buildTracker({ siteUrl: SITE, token: FAKE_TOKEN, template: await template(), fetchImpl: down });
    await expect(r).rejects.toMatchObject({ code: "unreachable", detail: expect.stringContaining("Failed to fetch") });
  }, 15_000);

  it("reports progress stages in order", async () => {
    const stages: string[] = [];
    await buildTracker({ siteUrl: SITE, token: FAKE_TOKEN, template: await template(), now: NOW, fetchImpl: fakeFetch(), onStage: (s) => stages.push(s.kind) });
    expect(stages).toEqual(["signedIn", "fetched", "building"]);
  });
});
