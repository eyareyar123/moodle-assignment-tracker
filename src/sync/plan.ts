// Pure diff between Moodle and the tracker. No I/O: easy to test and to print as a dry run.
//
// Field ownership:
//   Moodle owns   Assignment name, Type, Due Date, link, Moodle Key, Moodle Status
//   Shared        Status (Moodle may only move it forward), Grade (filled only when blank)
//   You own       Subject, Priority on existing rows (set once, on new rows)

import { keyFromUrl, type MoodleItem, type MoodleStatus } from "../moodle/types.js";
import type { ColKey, TrackerRow, TrackerStatus } from "../excel/tracker.js";

const HOMEWORK = "תרגיל בית להגשה";
const QUIZ = "בוחן";

/** Most Moodle quizzes here are graded homework; only ones named as a quiz/exam count as "בוחן". */
export function typeFor(item: Pick<MoodleItem, "kind" | "name">): string {
  return item.kind === "quiz" && /בוחן|מבחן/.test(item.name) ? QUIZ : HOMEWORK;
}

const STATUS_FROM_MOODLE: Record<MoodleStatus, TrackerStatus> = {
  not_submitted: "Not Started",
  draft: "In Progress",
  submitted: "Submitted",
  graded: "Graded",
};

const STATUS_RANK: Record<string, number> = { "Not Started": 0, "In Progress": 1, Submitted: 2, Graded: 3 };

const PLACEHOLDER_LINK_TEXT = /^(|tbd|n\\a|n\/a|-)$/i;

export interface FieldChange {
  field: ColKey;
  from: string | number | null;
  to: string | number | null;
}

export type Action =
  | { kind: "update"; row: number; item: MoodleItem; matchedBy: MatchMethod; changes: FieldChange[]; link?: LinkChange }
  | { kind: "add"; item: MoodleItem; values: Partial<Record<ColKey, string | number | null>>; link: LinkChange }
  | { kind: "unchanged"; row: number; item: MoodleItem; matchedBy: MatchMethod }
  | { kind: "skip"; item: MoodleItem; reason: string };

export type MatchMethod = "moodle-key" | "link" | "course+due";

export interface LinkChange {
  url: string;
  /** New text for the link cell, or undefined to keep the existing text. */
  text?: string;
}

export interface Warning {
  row?: number;
  key?: string;
  message: string;
}

export interface Plan {
  actions: Action[];
  warnings: Warning[];
  /** Subject labels used by new rows that are not yet in the Lists sheet. */
  newSubjects: string[];
}

export interface PlanOptions {
  subjects: string[];
  /** ISO date stamp written to "Last Synced" on changed rows. */
  today: string;
  /** Moodle keys never to add or update. */
  ignoreKeys?: string[];
  /** Don't add a new row for an item that was already due more than this many days ago. */
  staleAfterDays?: number;
}

const daysBetween = (a: string, b: string) => (Date.parse(b) - Date.parse(a)) / 86_400_000;

const courseCodeOf = (label: string) => /^\s*(\d{4,6})\b/.exec(label)?.[1] ?? null;

export function subjectFor(item: MoodleItem, subjects: string[]): string {
  const hit = item.courseCode && subjects.find((s) => courseCodeOf(s) === item.courseCode);
  return hit || (item.courseCode ? `${item.courseCode} - ${item.courseName}` : item.courseName);
}

export function buildPlan(rows: TrackerRow[], allItems: MoodleItem[], opts: PlanOptions): Plan {
  const warnings: Warning[] = [];
  const actions: Action[] = [];
  const ignore = new Set(opts.ignoreKeys ?? []);
  for (const i of allItems) if (ignore.has(i.key)) actions.push({ kind: "skip", item: i, reason: "in ignore list" });
  const items = allItems.filter((i) => !ignore.has(i.key));
  const filled = rows.filter((r) => !r.isEmpty);
  const claimed = new Set<number>(); // rows already matched to an item

  // Index rows by explicit key and by link-derived key. A key may appear on several rows (copy-paste).
  const byKey = new Map<string, TrackerRow[]>();
  const byLink = new Map<string, TrackerRow[]>();
  for (const r of filled) {
    if (r.moodleKey) push(byKey, r.moodleKey, r);
    const k = keyFromUrl(r.linkUrl);
    if (k) push(byLink, k, r);
  }

  const pick = (candidates: TrackerRow[] | undefined, item: MoodleItem): TrackerRow | undefined => {
    const free = (candidates ?? []).filter((r) => !claimed.has(r.row));
    if (free.length <= 1) return free[0];
    // Several rows claim this activity: prefer the one in the right course.
    const sameCourse = free.filter((r) => courseCodeOf(r.subject) === item.courseCode);
    return sameCourse[0] ?? free[0];
  };

  // Pass 1: exact identity (Moodle Key column, then link URL).
  const matched = new Map<string, { row: TrackerRow; by: MatchMethod }>();
  for (const item of items) {
    const byK = pick(byKey.get(item.key), item);
    const hit = byK ? { row: byK, by: "moodle-key" as const } : (() => {
      const r = pick(byLink.get(item.key), item);
      return r ? { row: r, by: "link" as const } : undefined;
    })();
    if (hit) {
      claimed.add(hit.row.row);
      matched.set(item.key, hit);
      const code = courseCodeOf(hit.row.subject);
      if (code && item.courseCode && code !== item.courseCode) {
        warnings.push({
          row: hit.row.row,
          key: item.key,
          message: `Row ${hit.row.row} is in course ${code} but its link points to "${item.name}" in course ${item.courseCode}. Check the link.`,
        });
      }
    }
  }

  // Pass 2: heuristic for rows with no usable identity — same course, same due date, unique.
  // A row whose link points at an activity Moodle did not return is left alone (it may be hidden).
  const itemKeys = new Set(items.map((i) => i.key));
  const ambiguous = new Set<string>();
  for (const item of items) {
    if (matched.has(item.key) || !item.due || !item.courseCode) continue;
    const cands = filled.filter((r) => {
      const lk = keyFromUrl(r.linkUrl);
      return (
        !claimed.has(r.row) &&
        !r.moodleKey &&
        !(lk && !itemKeys.has(lk)) &&
        r.due === item.due &&
        courseCodeOf(r.subject) === item.courseCode
      );
    });
    if (cands.length === 1) {
      claimed.add(cands[0]!.row);
      matched.set(item.key, { row: cands[0]!, by: "course+due" });
    } else if (cands.length > 1) {
      ambiguous.add(item.key);
      warnings.push({
        key: item.key,
        message: `"${item.name}" (${item.key}) could be rows ${cands.map((c) => c.row).join(", ")}. Skipped; put its key in "Moodle Key" on the right row.`,
      });
    }
  }

  // Rows whose link points at an activity another row claimed, and that found no match of their own:
  // the link is probably a copy-paste error and the row is left untouched.
  for (const [key, rs] of byLink) {
    for (const r of rs) {
      if (!claimed.has(r.row) && rs.some((o) => claimed.has(o.row))) {
        warnings.push({ row: r.row, key, message: `Row ${r.row} shares its link (${key}) with another row and matched nothing else; left as is.` });
      }
    }
  }

  const stale = opts.staleAfterDays ?? 14;
  for (const item of items) {
    const m = matched.get(item.key);
    if (m) actions.push(diffRow(m.row, item, m.by, opts));
    else if (ambiguous.has(item.key)) actions.push({ kind: "skip", item, reason: "ambiguous match" });
    else if (item.due && daysBetween(item.due, opts.today) > stale) {
      actions.push({ kind: "skip", item, reason: `due ${item.due}, already past when first seen (add its key to "ignore" to silence)` });
    } else actions.push(newRow(item, opts));
  }

  const known = new Set(opts.subjects);
  const newSubjects = [
    ...new Set(actions.flatMap((a) => (a.kind === "add" && typeof a.values.subject === "string" && !known.has(a.values.subject) ? [a.values.subject] : []))),
  ];

  // Keys in the sheet that Moodle no longer returns (deleted/hidden activity) — never deleted, only reported.
  const seen = new Set(items.map((i) => i.key));
  for (const r of filled) {
    if (r.moodleKey && !seen.has(r.moodleKey)) {
      warnings.push({ row: r.row, key: r.moodleKey, message: `Row ${r.row} (${r.moodleKey}) was not returned by Moodle; left as is.` });
    }
  }
  return { actions, warnings, newSubjects };
}

function diffRow(r: TrackerRow, item: MoodleItem, by: MatchMethod, opts: PlanOptions): Action {
  const changes: FieldChange[] = [];
  const set = (field: ColKey, from: string | number | null, to: string | number | null) => {
    if ((from ?? "") !== (to ?? "")) changes.push({ field, from, to });
  };

  set("assignment", r.assignment, item.name);
  set("type", r.type, typeFor(item));
  if (item.due) set("due", r.due, item.due); // a removed Moodle due date never blanks yours
  set("moodleKey", r.moodleKey, item.key);
  set("moodleStatus", r.moodleStatus, item.status);

  const proposed = STATUS_FROM_MOODLE[item.status];
  const cur = STATUS_RANK[r.status];
  if (cur === undefined ? r.status === "" : STATUS_RANK[proposed]! > cur) set("status", r.status, proposed);

  if ((r.grade === null || r.grade === "") && item.gradePercent !== null) set("grade", r.grade, round1(item.gradePercent));

  let link: LinkChange | undefined;
  if (r.linkUrl !== item.url) {
    link = { url: item.url, text: PLACEHOLDER_LINK_TEXT.test(r.linkText) || /^https?:/i.test(r.linkText) ? item.url : undefined };
    changes.push({ field: "link", from: r.linkUrl ?? (r.linkText || null), to: item.url });
  }

  // Bookkeeping-only changes (key/moodle status on first sync) still count; "Last Synced" marks them.
  if (changes.length === 0) return { kind: "unchanged", row: r.row, item, matchedBy: by };
  changes.push({ field: "lastSynced", from: null, to: opts.today });
  return { kind: "update", row: r.row, item, matchedBy: by, changes, link };
}

function newRow(item: MoodleItem, opts: PlanOptions): Action {
  return {
    kind: "add",
    item,
    values: {
      assignment: item.name,
      subject: subjectFor(item, opts.subjects),
      type: typeFor(item),
      due: item.due,
      status: STATUS_FROM_MOODLE[item.status],
      grade: item.gradePercent === null ? null : round1(item.gradePercent),
      moodleKey: item.key,
      moodleStatus: item.status,
      lastSynced: opts.today,
    },
    link: { url: item.url, text: item.url },
  };
}

const round1 = (n: number) => Math.round(n * 10) / 10;

function push<K, V>(m: Map<K, V[]>, k: K, v: V) {
  const a = m.get(k);
  if (a) a.push(v);
  else m.set(k, [v]);
}
