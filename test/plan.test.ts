import { describe, expect, it } from "vitest";
import { buildPlan, typeFor, type Action } from "../src/sync/plan.js";
import type { TrackerRow } from "../src/excel/tracker.js";
import type { MoodleItem } from "../src/moodle/types.js";

const BASE = "https://moodle.huji.ac.il/2026-27";
const SUBJECTS = ["57340 - סטטיסטיקה לכלכלה א", "51102 - חשיבה סטטיסטית בפסיכולוגיה", "57121 - חשבון דיפרנציאלי ואינטגרלי לכלכלנים א"];
const opts = { subjects: SUBJECTS, today: "2026-11-01" };

function row(r: number, p: Partial<TrackerRow>): TrackerRow {
  return {
    row: r, assignment: "", subject: "", type: "", due: null, priority: "", status: "", grade: null,
    linkText: "", linkUrl: null, moodleKey: "", moodleStatus: "", isEmpty: false, ...p,
  };
}

function item(kind: "assign" | "quiz", cmid: number, p: Partial<MoodleItem> = {}): MoodleItem {
  return {
    key: `${kind}:${cmid}`, kind, cmid, courseId: 1, courseCode: "57340", courseName: "סטטיסטיקה",
    name: `Moodle ${cmid}`, due: "2026-11-05", status: "not_submitted", gradePercent: null,
    url: `${BASE}/mod/${kind}/view.php?id=${cmid}`, ...p,
  };
}

const byKind = (actions: Action[], k: Action["kind"]) => actions.filter((a) => a.kind === k);

describe("buildPlan", () => {
  it("matches by link URL and updates only Moodle-owned fields", () => {
    const rows = [row(4, { assignment: "My own name", subject: SUBJECTS[1]!, type: "בוחן", due: "2026-11-03", status: "Not Started", priority: "High", linkUrl: `${BASE}/mod/assign/view.php?id=10`, linkText: "page title" })];
    const { actions } = buildPlan(rows, [item("assign", 10, { name: "Renamed in Moodle", due: "2026-11-06" })], opts);
    const a = actions[0]!;
    expect(a.kind).toBe("update");
    if (a.kind !== "update") return;
    expect(a.matchedBy).toBe("link");
    const to = Object.fromEntries(a.changes.map((c) => [c.field, c.to]));
    expect(to).toMatchObject({ due: "2026-11-06", assignment: "Renamed in Moodle", type: "תרגיל בית להגשה" });
    expect(to).not.toHaveProperty("priority");
    expect(to).not.toHaveProperty("subject");
  });

  it("types quizzes as homework unless named as a quiz or exam", () => {
    expect(typeFor({ kind: "quiz", name: "יחידה 2: פעולות ומשתנים" })).toBe("תרגיל בית להגשה");
    expect(typeFor({ kind: "quiz", name: "בוחן אמצע" })).toBe("בוחן");
    expect(typeFor({ kind: "assign", name: "הגשת תרגיל 1" })).toBe("תרגיל בית להגשה");
  });

  it("is idempotent: a second run over already-synced rows changes nothing", () => {
    const it1 = item("assign", 10);
    const rows = [row(4, { assignment: it1.name, type: "תרגיל בית להגשה", subject: SUBJECTS[0]!, due: it1.due, status: "Not Started", linkUrl: it1.url, moodleKey: it1.key, moodleStatus: "not_submitted" })];
    const { actions } = buildPlan(rows, [it1], opts);
    expect(actions[0]!.kind).toBe("unchanged");
  });

  it("adds unknown items once, with mapped subject and type", () => {
    const { actions } = buildPlan([], [item("quiz", 99, { courseCode: "51102" })], opts);
    const a = actions[0]!;
    expect(a.kind).toBe("add");
    if (a.kind !== "add") return;
    expect(a.values.subject).toBe(SUBJECTS[1]);
    expect(a.values.type).toBe("תרגיל בית להגשה");
    expect(a.values.status).toBe("Not Started");
    expect(a.values.moodleKey).toBe("quiz:99");
  });

  it("never moves Status backwards, but moves it forward", () => {
    const it1 = item("assign", 10, { status: "not_submitted" });
    const graded = item("assign", 11, { status: "graded", gradePercent: 87.25 });
    const rows = [
      row(4, { subject: SUBJECTS[0]!, status: "Submitted", linkUrl: it1.url, due: it1.due }),
      row(5, { subject: SUBJECTS[0]!, status: "Submitted", linkUrl: graded.url, due: graded.due }),
    ];
    const { actions } = buildPlan(rows, [it1, graded], opts);
    const a0 = actions[0]!;
    const a1 = actions[1]!;
    expect(a0.kind === "update" && a0.changes.some((c) => c.field === "status")).toBe(false);
    expect(a1.kind === "update" && a1.changes.find((c) => c.field === "status")?.to).toBe("Graded");
    expect(a1.kind === "update" && a1.changes.find((c) => c.field === "grade")?.to).toBe(87.3);
  });

  it("does not overwrite a grade you entered", () => {
    const g = item("assign", 11, { status: "graded", gradePercent: 50 });
    const { actions } = buildPlan([row(4, { subject: SUBJECTS[0]!, grade: "N\\A", linkUrl: g.url, due: g.due })], [g], opts);
    const a = actions[0]!;
    expect(a.kind === "update" && a.changes.some((c) => c.field === "grade")).toBe(false);
  });

  it("resolves a copy-pasted link by course, and rematches the other row by course+due", () => {
    // Real situation from the workbook: rows 11 (57340) and 12 (51102) both link to quiz 108782.
    const q51102 = item("quiz", 108782, { courseCode: "51102", due: "2026-10-20" });
    const a57340 = item("assign", 500, { courseCode: "57340", due: "2026-10-21" });
    const rows = [
      row(11, { subject: SUBJECTS[0]!, due: "2026-10-21", status: "Submitted", linkUrl: q51102.url }),
      row(12, { subject: SUBJECTS[1]!, due: "2026-10-20", status: "Submitted", linkUrl: q51102.url }),
    ];
    const { actions, warnings } = buildPlan(rows, [q51102, a57340], opts);
    expect(byKind(actions, "add")).toHaveLength(0);
    const u = actions.filter((a) => a.kind === "update" || a.kind === "unchanged") as Extract<Action, { row: number }>[];
    expect(u.find((a) => a.item.key === "quiz:108782")!.row).toBe(12);
    const fixed = u.find((a) => a.item.key === "assign:500")!;
    expect(fixed.row).toBe(11);
    expect(fixed.matchedBy).toBe("course+due");
    expect(warnings).toEqual([]); // row 11 found its real match, so nothing to warn about
  });

  it("warns about a shared link only when the row matched nothing else", () => {
    const q = item("quiz", 108782, { courseCode: "51102", due: "2026-10-20" });
    const rows = [
      row(11, { subject: SUBJECTS[0]!, due: "2026-10-21", linkUrl: q.url }),
      row(12, { subject: SUBJECTS[1]!, due: "2026-10-20", linkUrl: q.url }),
    ];
    const { warnings } = buildPlan(rows, [q], opts);
    expect(warnings.map((w) => w.row)).toEqual([11]);
  });

  it("skips ignored keys and items already past due when first seen", () => {
    const old = item("quiz", 1, { due: "2026-08-27" });
    const recent = item("quiz", 2, { due: "2026-10-25" }); // 7 days ago: still added
    const ign = item("assign", 3);
    const { actions } = buildPlan([], [old, recent, ign], { ...opts, ignoreKeys: ["assign:3"] });
    const kinds = Object.fromEntries(actions.map((a) => [a.item.key, a.kind]));
    expect(kinds).toEqual({ "quiz:1": "skip", "quiz:2": "add", "assign:3": "skip" });
  });

  it("still updates a tracked row even if it is long past due", () => {
    const old = item("assign", 10, { due: "2026-08-27" });
    const { actions } = buildPlan([row(4, { subject: SUBJECTS[0]!, linkUrl: old.url, due: "2026-08-26" })], [old], opts);
    expect(actions[0]!.kind).toBe("update");
  });

  it("reports subjects missing from Lists", () => {
    const { newSubjects } = buildPlan([], [item("quiz", 5, { courseCode: "51033", courseName: "תכנות R" }), item("quiz", 6)], opts);
    expect(newSubjects).toEqual(["51033 - תכנות R"]);
  });

  it("skips instead of duplicating when the heuristic is ambiguous", () => {
    const it1 = item("assign", 7);
    const rows = [row(4, { subject: SUBJECTS[0]!, due: it1.due }), row(5, { subject: SUBJECTS[0]!, due: it1.due })];
    const { actions } = buildPlan(rows, [it1], opts);
    expect(actions[0]!.kind).toBe("skip");
  });

  it("keeps link text that is a page title, replaces placeholders", () => {
    const a = item("assign", 1);
    const b = item("assign", 2, { due: "2026-11-09" });
    const rows = [
      row(4, { subject: SUBJECTS[0]!, due: a.due, linkText: "51102-... | HUJI" }),
      row(5, { subject: SUBJECTS[0]!, due: b.due, linkText: "TBD" }),
    ];
    const { actions } = buildPlan(rows, [a, b], opts);
    const [u0, u1] = actions as Extract<Action, { kind: "update" }>[];
    expect(u0!.link).toEqual({ url: a.url, text: undefined });
    expect(u1!.link).toEqual({ url: b.url, text: b.url });
  });
});
