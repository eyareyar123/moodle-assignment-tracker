// Everything the web page does except touching the DOM, so it can be tested in Node.
// Runs entirely in the visitor's browser: the token goes only to Moodle, the workbook never leaves the page.

import { XlsxPackage } from "../excel/ooxml.js";
import { readTracker } from "../excel/tracker.js";
import { MoodleClient, MoodleError } from "../moodle/client.js";
import { fetchItems, siteInfo } from "../moodle/fetch.js";
import { syncPackage, type SyncOutcome } from "../sync/run.js";

export const HUJI_MOODLE = "https://moodle.huji.ac.il";
export const TIMEZONE = "Asia/Jerusalem";

/** HUJI runs one Moodle per academic year ("2026-27"); a new one opens before the year starts in October. */
export function academicYear(d: Date): string {
  const y = d.getFullYear();
  const start = d.getMonth() >= 8 ? y : y - 1; // September onwards belongs to the coming year
  return `${start}-${String((start + 1) % 100).padStart(2, "0")}`;
}

export function siteForYear(year: string): string {
  if (!/^\d{4}-\d{2}$/.test(year)) throw new Error(`Not an academic year: ${year}`);
  return `${HUJI_MOODLE}/${year}`;
}

/** The current year and the one before, newest first (last year's site stays readable for a while). */
export function yearChoices(d: Date): string[] {
  const cur = academicYear(d);
  const start = Number(cur.slice(0, 4));
  return [cur, `${start - 1}-${String(start % 100).padStart(2, "0")}`];
}

export function todayIn(timeZone: string, d = new Date()): string {
  return new Intl.DateTimeFormat("en-CA", { timeZone, year: "numeric", month: "2-digit", day: "2-digit" }).format(d);
}

export interface BuildOptions {
  siteUrl: string;
  token: string;
  /** Blank template the page ships. */
  template: Uint8Array;
  /** The visitor's own tracker to update instead of starting blank. */
  existing?: Uint8Array;
  includeUndated?: boolean;
  now?: Date;
  fetchImpl?: typeof fetch;
  log?: (line: string) => void;
  /** Coarse progress for the page to show in the visitor's language. */
  onStage?: (stage: Stage) => void;
}

export type Stage =
  | { kind: "signedIn"; user: string }
  | { kind: "fetched"; count: number }
  | { kind: "building" };

export interface BuildResult {
  bytes: Uint8Array;
  outcome: SyncOutcome;
  /** Moodle's name for the signed-in user. */
  user: string;
  undatedSkipped: number;
}

export type UserErrorCode = "badUpload" | "badTemplate" | "rejected" | "refused" | "unreachable";

/** Errors worth showing to a visitor; the page translates them by `code`. */
export class UserError extends Error {
  constructor(
    readonly code: UserErrorCode,
    message: string,
    readonly detail?: string,
  ) {
    super(message);
  }
}

export async function buildTracker(o: BuildOptions): Promise<BuildResult> {
  const log = o.log ?? (() => {});
  const stage = o.onStage ?? (() => {});
  let pkg: XlsxPackage;
  try {
    pkg = await XlsxPackage.load(o.existing ?? o.template);
    await readTracker(pkg);
  } catch {
    throw o.existing
      ? new UserError("badUpload", 'That file is not a tracker made by this page (it needs the "Assignments" sheet with its table).')
      : new UserError("badTemplate", "The blank tracker could not be loaded. Reload the page and try again.");
  }

  const client = new MoodleClient(o.siteUrl, o.token, o.fetchImpl);
  let user: string;
  let items;
  try {
    const info = await siteInfo(client);
    user = info.fullname;
    log(`Signed in to Moodle as ${user}.`);
    stage({ kind: "signedIn", user });
    items = await fetchItems(client, { timezone: TIMEZONE, log });
  } catch (e) {
    if (e instanceof MoodleError && e.isAuth) {
      throw new UserError("rejected", "Moodle did not accept the link. It may be old or for another year's site: get a fresh one.");
    }
    if (e instanceof MoodleError) throw new UserError("refused", `Moodle refused a request: ${e.message}`, e.message);
    const msg = (e as Error).message;
    throw new UserError("unreachable", `Could not reach Moodle (${msg}). Check your connection and try again.`, msg);
  }

  const undated = items.filter((i) => !i.due);
  if (!o.includeUndated) items = items.filter((i) => i.due);
  if (undated.length && !o.includeUndated) log(`Leaving out ${undated.length} activities without a due date.`);
  stage({ kind: "fetched", count: items.length });
  stage({ kind: "building" });

  // A new tracker takes the whole year so far; an existing one only adds items that are still current.
  const outcome = await syncPackage(pkg, items, {
    today: todayIn(TIMEZONE, o.now),
    staleAfterDays: o.existing ? 14 : Number.POSITIVE_INFINITY,
  });
  return { bytes: await pkg.save(), outcome, user, undatedSkipped: o.includeUndated ? 0 : undated.length };
}

/** File name for the download: keep the visitor's own name when updating their file. */
export function downloadName(existingName: string | undefined, today: string): string {
  if (existingName && /\.xlsx$/i.test(existingName)) return existingName;
  return `Moodle assignments ${today}.xlsx`;
}
