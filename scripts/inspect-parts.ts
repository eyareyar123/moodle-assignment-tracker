// Dev helper: print workbook structure relevant to building a blank template.
//   npx tsx scripts/inspect-parts.ts [file] [part regex...]   e.g. "drawing1" "<row r=\"1[1-8]\".*?</row>"
import { readFile } from "node:fs/promises";
import JSZip from "jszip";

const [file = "ER assignments.xlsx", ...patterns] = process.argv.slice(2);
const zip = await JSZip.loadAsync(await readFile(file));
const s = (p: string) => zip.file(p)!.async("string");

if (patterns.length) {
  // Pairs of <part path> <regex>: print every match of the regex in that part.
  for (let i = 0; i + 1 < patterns.length; i += 2) {
    const x = await s(patterns[i]!);
    console.log(`== ${patterns[i]} /${patterns[i + 1]}/`);
    for (const m of x.matchAll(new RegExp(patterns[i + 1]!, "gs"))) console.log("  " + m[0].slice(0, 600));
  }
} else {
  const wb = await s("xl/workbook.xml");
  console.log("SHEETS:", [...wb.matchAll(/<sheet [^>]*name="([^"]+)"[^>]*r:id="([^"]+)"/g)].map((m) => `${m[1]}=${m[2]}`).join(", "));
  console.log("DEFINED NAMES:", wb.match(/<definedNames>.*<\/definedNames>/)?.[0] ?? "none");
  console.log("PARTS:", Object.keys(zip.files).join(" "));
  for (const p of ["docProps/core.xml", "docProps/app.xml"]) console.log(`\n${p}:\n`, (await s(p)).slice(0, 1500));
  const chart = await s("xl/charts/chart1.xml");
  console.log("\nCHART refs:", chart.match(/<c:f>[^<]+<\/c:f>/g)?.join(" "), "| caches:", (chart.match(/<c:(str|num)Cache>/g) ?? []).length);
}
