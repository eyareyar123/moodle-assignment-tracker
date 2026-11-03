// Usage:
//   npm run sync -- login            log in via HUJI SSO in a browser window, store the token (encrypted)
//   npm run sync -- whoami           check the stored token
//   npm run sync                     dry run: show what would change (default, writes nothing)
//   npm run sync -- --apply          back up the workbook, then apply the changes
//   npm run sync -- logout           delete the stored token
// Options: --include-undated (also add activities with no due date), --json <file> (save the plan)

import { parseArgs } from "node:util";
import { mkdir, writeFile } from "node:fs/promises";
import path from "node:path";
import { loadConfig, type Config } from "./config.js";
import { TokenStore } from "./auth/secret-store.js";
import { loginWithBrowser } from "./auth/login.js";
import { MoodleClient, MoodleError } from "./moodle/client.js";
import { fetchItems, siteInfo } from "./moodle/fetch.js";
import { loadWorkbook, saveWorkbook, assertNotOpenInExcel } from "./excel/file.js";
import { syncPackage } from "./sync/run.js";

const { values, positionals } = parseArgs({
  allowPositionals: true,
  options: {
    apply: { type: "boolean", default: false },
    "include-undated": { type: "boolean", default: false },
    json: { type: "string" },
  },
});

const cfg = loadConfig();
const store = new TokenStore(cfg.stateDir);

async function client(cfg: Config): Promise<MoodleClient> {
  const t = await store.load(cfg.siteUrl);
  if (!t) throw new Error(`No Moodle token for ${cfg.siteUrl}. Run: npm run sync -- login`);
  return new MoodleClient(cfg.siteUrl, t.token);
}

async function login() {
  const r = await loginWithBrowser(cfg);
  const c = new MoodleClient(cfg.siteUrl, r.token);
  const info = await siteInfo(c); // prove the token works before saving it
  await store.save({ siteUrl: cfg.siteUrl, token: r.token, privateToken: r.privateToken, createdAt: new Date().toISOString() });
  console.log(`Logged in as ${info.fullname} on ${info.sitename}. Token saved (encrypted) in ${cfg.stateDir}`);
}

async function sync() {
  const today = new Intl.DateTimeFormat("en-CA", { timeZone: cfg.timezone }).format(new Date());
  console.log(`Fetching from ${cfg.siteUrl} ...`);
  let items = await fetchItems(await client(cfg), { timezone: cfg.timezone, log: (s) => console.log(`  ${s}`) });
  if (!values["include-undated"]) {
    const n = items.filter((i) => !i.due).length;
    items = items.filter((i) => i.due);
    if (n) console.log(`  ignoring ${n} activities without a due date (use --include-undated to add them)`);
  }

  // Applied in memory either way; only --apply saves it.
  const pkg = await loadWorkbook(cfg.workbook);
  const out = await syncPackage(pkg, items, { today, ignoreKeys: cfg.ignore, staleAfterDays: cfg.staleAfterDays });
  console.log(`\n${out.report}\n`);
  if (values.json) {
    await mkdir(path.dirname(path.resolve(values.json)), { recursive: true });
    await writeFile(values.json, JSON.stringify({ items, plan: out.plan }, null, 2));
  }

  if (!values.apply) {
    console.log(out.changed ? "Dry run — nothing written. Re-run with --apply to make these changes." : "Already up to date.");
    return;
  }
  if (!out.changed) return console.log("Already up to date — workbook not touched.");
  await assertNotOpenInExcel(cfg.workbook);
  const bak = await saveWorkbook(cfg.workbook, await pkg.save(), cfg.backupDir);
  const { updated, added } = out.applied;
  console.log(`Saved ${path.basename(cfg.workbook)}: ${updated} updated, ${added} added, ${out.moved} re-sorted. Backup: ${path.relative(process.cwd(), bak)}`);
}

async function main() {
  const cmd = positionals[0] ?? "sync";
  if (cmd === "login") return login();
  if (cmd === "logout") return (await store.clear(), console.log("Token deleted."));
  if (cmd === "whoami") {
    const i = await siteInfo(await client(cfg));
    return console.log(`${i.fullname} @ ${i.sitename}`);
  }
  if (cmd === "sync") return sync();
  throw new Error(`Unknown command "${cmd}"`);
}

main().catch((e) => {
  if (e instanceof MoodleError && e.isAuth) console.error(`Moodle rejected the token (${e.errorcode}). Run: npm run sync -- login`);
  else console.error(`Error: ${(e as Error).message}`);
  process.exitCode = 1;
});
