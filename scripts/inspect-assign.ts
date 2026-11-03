// Dev helper: print raw Moodle date fields for assignments with the given cmids.
import { loadConfig } from "../src/config.js";
import { TokenStore } from "../src/auth/secret-store.js";
import { MoodleClient } from "../src/moodle/client.js";

const cfg = loadConfig();
const t = await new TokenStore(cfg.stateDir).load(cfg.siteUrl);
const c = new MoodleClient(cfg.siteUrl, t!.token);
const info = await c.call<{ userid: number }>("core_webservice_get_site_info");
const courses = await c.call<{ id: number; fullname: string }[]>("core_enrol_get_users_courses", { userid: info.userid });
const res = await c.call<{ courses: { assignments: Record<string, unknown>[] }[] }>("mod_assign_get_assignments", { courseids: courses.map((x) => x.id) });
const want = new Set(process.argv.slice(2).map(Number));
const fmt = (ts: number) => (ts ? new Date(ts * 1000).toLocaleString("en-GB", { timeZone: cfg.timezone }) : "-");
for (const a of res.courses.flatMap((x) => x.assignments)) {
  if (!want.has(a.cmid as number)) continue;
  console.log(a.cmid, a.name, { allowsubmissionsfromdate: fmt(a.allowsubmissionsfromdate as number), duedate: fmt(a.duedate as number), cutoffdate: fmt(a.cutoffdate as number) });
}
