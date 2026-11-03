// Obtain a Moodle web-service token through the site's SSO, the same way the Moodle mobile app does.
//
// HUJI Moodle uses SAML ("Connect to HUJI"), so login/token.php with a password won't work. Instead:
//   1. open  <site>/admin/tool/mobile/launch.php?service=moodle_mobile_app&passport=<random>&urlscheme=moodlemobile
//   2. you log in through HUJI SSO in a real browser window (Playwright; your password is typed by you, never seen by this code)
//   3. Moodle redirects to  moodlemobile://token=<base64("md5(site+passport):::token:::privatetoken")>
//   4. we capture that redirect, verify the signature, and keep the token.

import { createHash, randomInt } from "node:crypto";
import path from "node:path";
import { chromium } from "playwright";
import type { Config } from "../config.js";
import { launchUrl, parseLaunchToken } from "../moodle/launch.js";

export interface LoginResult {
  token: string;
  privateToken?: string;
}

export function decodeLaunchUrl(url: string, siteUrl: string, passport: string): LoginResult {
  if (!url.startsWith("moodlemobile://token=")) throw new Error("Unexpected redirect from Moodle (no token)");
  const { signature, token, privateToken } = parseLaunchToken(url);
  const expected = createHash("md5").update(siteUrl + passport).digest("hex");
  // Moodle signs with its wwwroot; accept a trailing-slash variant.
  const alt = createHash("md5").update(`${siteUrl}/` + passport).digest("hex");
  if (signature !== expected && signature !== alt) throw new Error("Token signature mismatch — refusing token");
  return { token, privateToken };
}

export async function loginWithBrowser(cfg: Config, log: (s: string) => void = console.log): Promise<LoginResult> {
  const passport = String(randomInt(1, 2 ** 31));
  const launch = launchUrl(cfg.siteUrl, passport);

  // A persistent profile keeps the SSO session, so future logins are usually one click.
  const ctx = await chromium.launchPersistentContext(path.join(cfg.stateDir, "browser-profile"), {
    channel: cfg.browserChannel,
    headless: false,
    viewport: null,
  });
  try {
    const page = ctx.pages()[0] ?? (await ctx.newPage());
    const captured = new Promise<string>((resolve) => {
      const check = (u: string | null | undefined) => u?.startsWith("moodlemobile://") && resolve(u);
      ctx.on("response", (r) => check(r.headers()["location"]));
      ctx.on("request", (r) => check(r.url()));
      ctx.on("requestfailed", (r) => check(r.url()));
      page.on("framenavigated", (f) => check(f.url()));
    });
    const closed = new Promise<never>((_, reject) => ctx.on("close", () => reject(new Error("Browser closed before login finished"))));

    log("A browser window opened. Log in with your HUJI account there; it will close by itself.");
    await page.goto(launch).catch(() => {}); // navigation to moodlemobile:// aborts — expected
    const url = await Promise.race([captured, closed]);
    return decodeLaunchUrl(url, cfg.siteUrl, passport);
  } finally {
    await ctx.close().catch(() => {});
  }
}
