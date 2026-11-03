// Collect the current user's assignments and quizzes from Moodle and normalise them to MoodleItem.

import { MoodleClient, MoodleError } from "./client.js";
import { moodleKey, type MoodleItem, type MoodleStatus } from "./types.js";

interface SiteInfo {
  userid: number;
  fullname: string;
  sitename: string;
  functions?: { name: string }[];
}
interface Course {
  id: number;
  shortname: string;
  fullname: string;
  idnumber?: string;
  visible?: number;
}
interface Assign {
  id: number;
  cmid: number;
  course: number;
  name: string;
  duedate: number;
  nosubmissions: number;
  grade: number;
}
interface AssignStatus {
  lastattempt?: {
    submission?: { status: string };
    teamsubmission?: { status: string };
    gradingstatus?: string;
  };
  feedback?: { grade?: { grade: string | null } | null };
}
interface Quiz {
  id: number;
  coursemodule: number;
  course: number;
  name: string;
  timeclose: number;
  grade: number;
}
interface QuizAttempt {
  state: string;
}

/**
 * Resolve Moodle multi-language text ("{mlang he}×©×œ×•×{mlang}{mlang en}Hello{mlang}", also the
 * <span lang="he" class="multilang"> form) to one language, falling back to the first block.
 */
export function cleanText(s: string, lang = "he"): string {
  const pick = (blocks: [string, string][]) => (blocks.find(([l]) => l === lang) ?? blocks[0])?.[1] ?? "";
  const mlang = [...s.matchAll(/\{mlang\s+([\w-]+)\}([\s\S]*?)\{mlang\}/g)].map((m) => [m[1]!, m[2]!] as [string, string]);
  if (mlang.length) s = s.replace(/(\{mlang\s+[\w-]+\}[\s\S]*?\{mlang\}\s*)+/, pick(mlang));
  const spans = [...s.matchAll(/<span\s+lang="([\w-]+)"\s+class="multilang">([\s\S]*?)<\/span>/g)].map((m) => [m[1]!, m[2]!] as [string, string]);
  if (spans.length) s = s.replace(/(<span\s+lang="[\w-]+"\s+class="multilang">[\s\S]*?<\/span>\s*)+/, pick(spans));
  // A Hebrew prefix letter glued to a Latin word ("וtidyr") is reordered by bidi display; Hebrew
  // typography puts a hyphen there ("ו-tidyr").
  s = s.replace(/(^|[\s(])([ובהלמשכ]{1,2})(?=[A-Za-z0-9])/g, "$1$2-");
  return s.replace(/\s+/g, " ").trim();
}

/** Course title without its leading number ("51033 - R Programming" -> "R Programming"). */
export function courseTitle(fullname: string, code: string | null): string {
  const t = cleanText(fullname);
  return code ? t.replace(new RegExp(`^\\s*${code}\\s*[-â€“:]?\\s*`), "") || t : t;
}

export function courseCode(c: Pick<Course, "shortname" | "fullname" | "idnumber">): string | null {
  for (const s of [c.idnumber, c.shortname, c.fullname]) {
    const m = s && /^\s*(\d{4,6})(?!\d)/.exec(s);
    if (m) return m[1]!;
  }
  return null;
}

/**
 * Unix timestamp -> calendar date in the given time zone. A deadline at exactly 00:00 belongs to the
 * previous day (Moodle's "due midnight" means end of that day), so we step back one second.
 */
export function dueDate(ts: number, timeZone: string): string | null {
  if (!ts) return null;
  return new Intl.DateTimeFormat("en-CA", { timeZone, year: "numeric", month: "2-digit", day: "2-digit" }).format(new Date((ts - 1) * 1000));
}

export function assignStatus(s: AssignStatus): { status: MoodleStatus; rawGrade: number | null } {
  const g = s.feedback?.grade?.grade;
  const rawGrade = g !== null && g !== undefined && g !== "" && !Number.isNaN(Number(g)) ? Number(g) : null;
  const sub = s.lastattempt?.submission?.status ?? s.lastattempt?.teamsubmission?.status;
  if (rawGrade !== null || s.lastattempt?.gradingstatus === "released") return { status: "graded", rawGrade };
  if (sub === "submitted") return { status: "submitted", rawGrade };
  if (sub === "draft") return { status: "draft", rawGrade };
  return { status: "not_submitted", rawGrade };
}

export interface FetchOptions {
  timezone: string;
  log?: (s: string) => void;
}

export async function siteInfo(client: MoodleClient): Promise<SiteInfo> {
  return client.call<SiteInfo>("core_webservice_get_site_info");
}

export async function fetchItems(client: MoodleClient, opts: FetchOptions): Promise<MoodleItem[]> {
  const log = opts.log ?? (() => {});
  const info = await siteInfo(client);
  const courses = (await client.call<Course[]>("core_enrol_get_users_courses", { userid: info.userid })).filter((c) => c.visible !== 0);
  log(`${courses.length} courses`);
  const byId = new Map(courses.map((c) => [c.id, c]));
  const courseids = courses.map((c) => c.id);
  const items: MoodleItem[] = [];
  const base = (course: Course) => {
    const code = courseCode(course);
    return { courseId: course.id, courseCode: code, courseName: courseTitle(course.fullname, code) };
  };

  // ---- assignments ----
  const assigns = (await client.call<{ courses: { id: number; assignments: Assign[] }[] }>("mod_assign_get_assignments", { courseids })).courses.flatMap(
    (c) => c.assignments,
  );
  for (const a of assigns) {
    const course = byId.get(a.course);
    if (!course) continue;
    let status: MoodleStatus = "not_submitted";
    let gradePercent: number | null = null;
    try {
      const s = assignStatus(await client.call<AssignStatus>("mod_assign_get_submission_status", { assignid: a.id }));
      status = s.status;
      if (s.rawGrade !== null && a.grade > 0) gradePercent = (s.rawGrade / a.grade) * 100;
    } catch (e) {
      if (!(e instanceof MoodleError) || e.isAuth) throw e;
      log(`  (no submission status for "${a.name}": ${e.message})`);
    }
    items.push({
      key: moodleKey("assign", a.cmid), kind: "assign", cmid: a.cmid, ...base(course), name: cleanText(a.name),
      due: dueDate(a.duedate, opts.timezone), status, gradePercent, url: `${client.siteUrl}/mod/assign/view.php?id=${a.cmid}`,
    });
  }
  log(`${assigns.length} assignments`);

  // ---- quizzes ----
  const quizzes = (await client.call<{ quizzes: Quiz[] }>("mod_quiz_get_quizzes_by_courses", { courseids })).quizzes;
  for (const q of quizzes) {
    const course = byId.get(q.course);
    if (!course) continue;
    let status: MoodleStatus = "not_submitted";
    let gradePercent: number | null = null;
    try {
      const { attempts } = await client.call<{ attempts: QuizAttempt[] }>("mod_quiz_get_user_attempts", { quizid: q.id, status: "all", includepreviews: false });
      if (attempts.some((t) => t.state === "finished")) status = "submitted";
      else if (attempts.some((t) => t.state === "inprogress" || t.state === "overdue")) status = "draft";
      if (status === "submitted") {
        const best = await client.call<{ hasgrade: boolean; grade?: number }>("mod_quiz_get_user_best_grade", { quizid: q.id });
        if (best.hasgrade && best.grade !== undefined && q.grade > 0) {
          status = "graded";
          gradePercent = (best.grade / q.grade) * 100;
        }
      }
    } catch (e) {
      if (!(e instanceof MoodleError) || e.isAuth) throw e;
      log(`  (no attempt info for "${q.name}": ${e.message})`);
    }
    items.push({
      key: moodleKey("quiz", q.coursemodule), kind: "quiz", cmid: q.coursemodule, ...base(course), name: cleanText(q.name),
      due: dueDate(q.timeclose, opts.timezone), status, gradePercent, url: `${client.siteUrl}/mod/quiz/view.php?id=${q.coursemodule}`,
    });
  }
  log(`${quizzes.length} quizzes`);
  return items;
}
