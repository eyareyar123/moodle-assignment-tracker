export type MoodleKind = "assign" | "quiz";

/** Normalised submission state, independent of activity type. */
export type MoodleStatus = "not_submitted" | "draft" | "submitted" | "graded";

/** One Moodle activity, reduced to what the tracker needs. */
export interface MoodleItem {
  /** Stable identity: `${kind}:${cmid}` — cmid is the `id=` in the activity's view.php URL. */
  key: string;
  kind: MoodleKind;
  cmid: number;
  courseId: number;
  /** Leading course number, e.g. "57340", parsed from the course shortname/fullname. */
  courseCode: string | null;
  courseName: string;
  name: string;
  /** Due date in Israel time, YYYY-MM-DD, or null if the activity has none. */
  due: string | null;
  status: MoodleStatus;
  /** Grade as a percentage 0–100, when Moodle exposes one. */
  gradePercent: number | null;
  url: string;
}

export function moodleKey(kind: MoodleKind, cmid: number): string {
  return `${kind}:${cmid}`;
}

/** Parse `.../mod/<kind>/view.php?id=<cmid>` into a key, or null. */
export function keyFromUrl(url: string | null | undefined): string | null {
  const m = url && /\/mod\/(assign|quiz)\/view\.php\?(?:.*&)?id=(\d+)/.exec(url);
  return m ? moodleKey(m[1] as MoodleKind, Number(m[2])) : null;
}
