// Build the blank workbook the web page starts from:  npx tsx scripts/make-template.ts [source.xlsx]
// Writes web/public/tracker-template.xlsx. Re-run it after changing the tracker's layout.
import { mkdir, readFile, writeFile } from "node:fs/promises";
import path from "node:path";
import { makeTemplate } from "../src/excel/template.js";
import { verifyWorkbook } from "../src/excel/file.js";

export const TEMPLATE_PATH = path.resolve("web/public/tracker-template.xlsx");

const src = path.resolve(process.argv[2] ?? "ER assignments.xlsx");
const out = await makeTemplate(await readFile(src));
await verifyWorkbook(out);
await mkdir(path.dirname(TEMPLATE_PATH), { recursive: true });
await writeFile(TEMPLATE_PATH, out);
console.log(`Wrote ${path.relative(process.cwd(), TEMPLATE_PATH)} (${out.length} bytes) from ${path.basename(src)}`);
