import { describe, expect, it } from "vitest";
import { XlsxPackage } from "../src/excel/ooxml.js";
import { readTracker } from "../src/excel/tracker.js";
import { fixtureBytes } from "./helpers.js";

describe("readTracker on the real workbook", async () => {
  const t = await readTracker(await XlsxPackage.load(await fixtureBytes()));
  const filled = t.rows.filter((r) => !r.isEmpty);

  it("finds the Assignments table and its columns", () => {
    expect(t.table.headerRow).toBe(3);
    expect(t.table.firstRow).toBe(4);
    expect(t.table.cols.assignment).toBe("A");
    expect(t.table.cols.due).toBe("D");
    expect(t.table.cols.link).toBe("I");
  });

  it("reads the 13 filled rows with dates and statuses", () => {
    expect(filled).toHaveLength(13);
    const r7 = t.rows.find((r) => r.row === 7)!;
    expect(r7.assignment).toBe("Week 1 - תרגיל 1");
    expect(r7.subject).toMatch(/^57121 /);
    expect(r7.due).toMatch(/^\d{4}-\d{2}-\d{2}$/);
    expect(r7.status).toBe("Graded");
    expect(r7.grade).toBe(100);
  });

  it("prefers hyperlink targets over link-cell text", () => {
    const r4 = t.rows.find((r) => r.row === 4)!;
    expect(r4.linkText).toMatch(/^57340-/);
    expect(r4.linkUrl).toBe("https://moodle.huji.ac.il/2026-27/mod/assign/view.php?id=47410");
    const r5 = t.rows.find((r) => r.row === 5)!; // plain URL text, no hyperlink
    expect(r5.linkUrl).toBe("https://moodle.huji.ac.il/2026-27/mod/quiz/view.php?id=108780");
    expect(t.rows.find((r) => r.row === 14)!.linkUrl).toBeNull();
  });

  it("reads subject labels from Lists", () => {
    expect(t.subjects).toHaveLength(6);
    expect(t.subjects.every((s) => /^\d{5} - /.test(s))).toBe(true);
  });
});
