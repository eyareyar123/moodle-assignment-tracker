// Dev helper: write a migrated copy (plus one synthetic added row) to .out/ for opening in Excel.
import { mkdir, readFile, writeFile } from "node:fs/promises";
import { XlsxPackage } from "../src/excel/ooxml.js";
import { migrate } from "../src/excel/migrate.js";
import { readTracker } from "../src/excel/tracker.js";
import { applyPlan } from "../src/excel/write.js";
import { buildPlan } from "../src/sync/plan.js";

const pkg = await XlsxPackage.load(await readFile(process.argv[2] ?? "ER assignments.xlsx"));
console.log((await migrate(pkg)).join("\n"));
const t = await readTracker(pkg);
const plan = buildPlan(t.rows, [{
  key: "quiz:200001", kind: "quiz", cmid: 200001, courseId: 2, courseCode: "51102", courseName: "x",
  name: "SAMPLE ROW (not real)", due: "2026-11-12", status: "not_submitted", gradePercent: null,
  url: "https://moodle.huji.ac.il/2026-27/mod/quiz/view.php?id=200001",
}], { subjects: t.subjects, today: "2026-11-01" });
await applyPlan(pkg, plan);
await mkdir(".out", { recursive: true });
await writeFile(".out/sample.xlsx", await pkg.save());
console.log("wrote .out/sample.xlsx");
