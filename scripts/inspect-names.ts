// Dev helper: print raw Moodle names (escaped) of every assignment and quiz, with its course.
import { loadConfig } from "../src/config.js";
import { TokenStore } from "../src/auth/secret-store.js";
import { MoodleClient } from "../src/moodle/client.js";

const cfg = loadConfig();
const t = await new TokenStore(cfg.stateDir).load(cfg.siteUrl);
const c = new MoodleClient(cfg.siteUrl, t!.token);
const info = await c.call<{ userid: number }>("core_webservice_get_site_info");
const courses = await c.call<{ id: number; shortname: string }[]>("core_enrol_get_users_courses", { userid: info.userid });
const short = new Map(courses.map((x) => [x.id, x.shortname]));
const ids = courses.map((x) => x.id);
const esc = (s: string) => JSON.stringify(s).replace(/[\u0080-￿]/g, (ch) => (/[֐-׿]/.test(ch) ? ch : `\\u${ch.charCodeAt(0).toString(16).padStart(4, "0")}`));
const a = await c.call<{ courses: { assignments: { cmid: number; course: number; name: string }[] }[] }>("mod_assign_get_assignments", { courseids: ids });
for (const x of a.courses.flatMap((y) => y.assignments)) console.log("assign", x.cmid, short.get(x.course), esc(x.name));
const q = await c.call<{ quizzes: { coursemodule: number; course: number; name: string }[] }>("mod_quiz_get_quizzes_by_courses", { courseids: ids });
for (const x of q.quizzes) console.log("quiz  ", x.coursemodule, short.get(x.course), esc(x.name));
