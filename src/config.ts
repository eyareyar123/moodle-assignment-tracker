// Non-secret settings live in tracker.config.json (committed). Secrets never live in the repo:
// the Moodle token and the browser profile are stored per-user under %LOCALAPPDATA%.

import { readFileSync } from "node:fs";
import os from "node:os";
import path from "node:path";

export interface Config {
  siteUrl: string;
  workbook: string;
  backupDir: string;
  timezone: string;
  browserChannel: string;
  /** Moodle keys ("assign:123", "quiz:456") never to add or update. */
  ignore: string[];
  /** New items already due more than this many days ago are not added. */
  staleAfterDays: number;
  /** Per-user directory for the encrypted token and the SSO browser profile. */
  stateDir: string;
}

export function loadConfig(file = "tracker.config.json"): Config {
  const raw = JSON.parse(readFileSync(file, "utf8")) as Partial<Config>;
  const base = process.env.LOCALAPPDATA ?? path.join(os.homedir(), "AppData", "Local");
  const cfg: Config = {
    siteUrl: (process.env.MOODLE_SITE_URL ?? raw.siteUrl ?? "").replace(/\/+$/, ""),
    workbook: path.resolve(process.env.TRACKER_WORKBOOK ?? raw.workbook ?? "ER assignments.xlsx"),
    backupDir: path.resolve(raw.backupDir ?? "backups"),
    timezone: raw.timezone ?? "Asia/Jerusalem",
    browserChannel: raw.browserChannel ?? "msedge",
    ignore: raw.ignore ?? [],
    staleAfterDays: raw.staleAfterDays ?? 14,
    stateDir: path.join(base, "moodle-assignment-tracker"),
  };
  if (!/^https:\/\//.test(cfg.siteUrl)) throw new Error(`siteUrl must be an https URL (got "${cfg.siteUrl}")`);
  return cfg;
}
