// Dev helper: print the tracker's filled rows.
import { readFile } from "node:fs/promises";
import { loadConfig } from "../src/config.js";
import { XlsxPackage } from "../src/excel/ooxml.js";
import { readTracker } from "../src/excel/tracker.js";

const cfg = loadConfig();
const t = await readTracker(await XlsxPackage.load(await readFile(cfg.workbook)));
for (const r of t.rows.filter((x) => !x.isEmpty)) console.log(r.row, r.moodleKey || "-", "|", r.assignment, "|", r.type, "|", r.status, "|", r.subject);
