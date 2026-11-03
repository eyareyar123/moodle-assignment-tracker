// Human-readable dry-run output.

import type { Plan } from "./plan.js";
import { COL } from "../excel/tracker.js";

const fmt = (v: unknown) => (v === null || v === undefined || v === "" ? "∅" : String(v));

export function formatPlan(plan: Plan, migrationLog: string[] = []): string {
  const out: string[] = [];
  if (migrationLog.length) {
    out.push("Workbook upgrade (one-time):");
    for (const l of migrationLog) out.push(`  • ${l}`);
    out.push("");
  }
  const adds = plan.actions.filter((a) => a.kind === "add");
  const updates = plan.actions.filter((a) => a.kind === "update");
  const skips = plan.actions.filter((a) => a.kind === "skip");
  const same = plan.actions.filter((a) => a.kind === "unchanged");

  if (updates.length) {
    out.push(`UPDATE ${updates.length} row(s):`);
    for (const a of updates) {
      if (a.kind !== "update") continue;
      out.push(`  row ${a.row}  [${a.item.key}, matched by ${a.matchedBy}]  ${a.item.name}`);
      for (const c of a.changes) {
        if (c.field === "lastSynced") continue;
        out.push(`      ${COL[c.field].padEnd(14)} ${fmt(c.from)}  →  ${fmt(c.to)}`);
      }
    }
    out.push("");
  }
  if (adds.length) {
    out.push(`ADD ${adds.length} new row(s):`);
    for (const a of adds) {
      if (a.kind !== "add") continue;
      out.push(`  + ${fmt(a.values.due)}  ${fmt(a.values.subject)}  |  ${a.item.name}  [${a.item.key}, ${a.values.status}]`);
    }
    out.push("");
  }
  if (plan.newSubjects.length) {
    out.push("NEW SUBJECTS (added to Lists dropdown and Dashboard):");
    for (const s of plan.newSubjects) out.push(`  + ${s}`);
    out.push("");
  }
  if (skips.length) {
    out.push(`SKIPPED ${skips.length}:`);
    for (const a of skips) {
      if (a.kind !== "skip") continue;
      const reason = a.reason.startsWith("due ") ? "already past due" : a.reason;
      out.push(`  - ${fmt(a.item.due)}  ${a.item.name} [${a.item.key}]: ${reason}`);
    }
    out.push("");
  }
  if (plan.warnings.length) {
    out.push("WARNINGS:");
    for (const w of plan.warnings) out.push(`  ! ${w.message}`);
    out.push("");
  }
  out.push(`Summary: ${updates.length} update, ${adds.length} add, ${same.length} unchanged, ${skips.length} skipped.`);
  return out.join("\n");
}
