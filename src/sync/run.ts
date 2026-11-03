// One sync of Moodle items into a workbook, in memory: upgrade the table, plan, apply, sort.
// Shared by the CLI (which then saves with a backup) and the web page (which offers a download).

import type { XlsxPackage } from "../excel/ooxml.js";
import { migrate } from "../excel/migrate.js";
import { readTracker } from "../excel/tracker.js";
import { applyPlan, type ApplyResult } from "../excel/write.js";
import { sortTable } from "../excel/sort.js";
import type { MoodleItem } from "../moodle/types.js";
import { buildPlan, type Plan } from "./plan.js";
import { formatPlan } from "./report.js";

export interface SyncOptions {
  /** ISO date (YYYY-MM-DD) in the user's time zone. */
  today: string;
  ignoreKeys?: string[];
  staleAfterDays?: number;
}

export interface SyncOutcome {
  plan: Plan;
  migrationLog: string[];
  applied: ApplyResult;
  /** Rows that changed position when sorting. */
  moved: number;
  /** True when the workbook differs from what was loaded. */
  changed: boolean;
  /** Human-readable summary of everything above. */
  report: string;
}

export async function syncPackage(pkg: XlsxPackage, items: MoodleItem[], opts: SyncOptions): Promise<SyncOutcome> {
  const migrationLog = await migrate(pkg);
  const t = await readTracker(pkg);
  const plan = buildPlan(t.rows, items, { subjects: t.subjects, today: opts.today, ignoreKeys: opts.ignoreKeys, staleAfterDays: opts.staleAfterDays });
  const applied = await applyPlan(pkg, plan);
  const moved = await sortTable(pkg);
  const changed = applied.updated > 0 || applied.added > 0 || moved > 0 || migrationLog.length > 0 || applied.subjectsAdded.length > 0;

  const lines = [formatPlan(plan, migrationLog)];
  if (applied.warnings.length) lines.push("", "WORKBOOK WARNINGS:", ...applied.warnings.map((w) => `  ! ${w}`));
  if (moved) lines.push("", `SORT: ${moved} row(s) move (pending by due date, then Submitted, then Graded).`);
  return { plan, migrationLog, applied, moved, changed, report: lines.join("\n") };
}
