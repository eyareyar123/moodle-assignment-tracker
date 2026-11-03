// The built page in a real browser (Edge), served by `vite preview`, with Moodle answered by the fake.
import { afterAll, beforeAll, describe, expect, it } from "vitest";
import { build, preview, type PreviewServer } from "vite";
import { chromium, type Browser, type BrowserContext, type Page } from "playwright";
import { mkdtemp, readFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { XlsxPackage } from "../../src/excel/ooxml.js";
import { readTracker } from "../../src/excel/tracker.js";
import { academicYear } from "../../src/web/core.js";
import { FAKE_TOKEN, fakeMoodleAnswer } from "../fake-moodle.js";

const PORT = 4174; // not 4173, so a running `npm run web:local` doesn't clash
/** Screenshots for a human look (git-ignored). */
const SHOTS = path.resolve(".out/e2e");
const ORIGIN = `http://localhost:${PORT}`;
const LINK = `moodlemobile://token=${Buffer.from(`0cc175b9c0f1b6a831c399e269772661:::${FAKE_TOKEN}`).toString("base64")}`;

let server: PreviewServer;
let browser: Browser;
let ctx: BrowserContext;
let page: Page;
const hosts = new Set<string>();
const consoleErrors: string[] = [];
let dir: string;

beforeAll(async () => {
  await build({ configFile: "vite.config.ts", logLevel: "warn" });
  server = await preview({ configFile: "vite.config.ts", preview: { port: PORT, strictPort: true }, logLevel: "warn" });
  browser = await chromium.launch({ channel: "msedge", headless: true });
  ctx = await browser.newContext({ acceptDownloads: true, locale: "en-US" });
  ctx.on("request", (r) => hosts.add(new URL(r.url()).origin));
  await ctx.route("https://moodle.huji.ac.il/**", async (route) => {
    const req = route.request();
    if (req.url().includes("/webservice/rest/server.php")) {
      const params = new URLSearchParams(req.postData() ?? "");
      const body = params.get("wstoken") === FAKE_TOKEN
        ? fakeMoodleAnswer(params.get("wsfunction") ?? "", params)
        : { exception: "moodle_exception", errorcode: "invalidtoken", message: "Invalid token" };
      return route.fulfill({ status: 200, contentType: "application/json", headers: { "Access-Control-Allow-Origin": "*" }, body: JSON.stringify(body) });
    }
    return route.fulfill({ status: 200, contentType: "text/html", body: `<a href="${LINK}">Click here if the app does not open automatically.</a>` });
  });
  page = await ctx.newPage();
  page.on("console", (m) => m.type() === "error" && consoleErrors.push(m.text()));
  page.on("pageerror", (e) => consoleErrors.push(e.message));
  dir = await mkdtemp(path.join(tmpdir(), "tracker-e2e-"));
  await page.goto(ORIGIN);
});

afterAll(async () => {
  await browser?.close();
  await new Promise<void>((r) => server?.httpServer.close(() => r()));
});

async function buildAndDownload(): Promise<{ file: string; name: string }> {
  const [download] = await Promise.all([
    page.waitForEvent("download"),
    (async () => {
      await page.click("#build");
      await page.locator("#download").waitFor({ state: "visible" });
      await page.click("#download");
    })(),
  ]);
  const file = path.join(dir, `${Date.now()}-${download.suggestedFilename()}`);
  await download.saveAs(file);
  return { file, name: download.suggestedFilename() };
}

describe("web page", () => {
  it("ships with a strict Content-Security-Policy", async () => {
    const csp = await page.locator('meta[http-equiv="Content-Security-Policy"]').getAttribute("content");
    expect(csp).toContain("connect-src 'self' https://moodle.huji.ac.il");
    expect(csp).toContain("script-src 'self'");
  });

  it("offers the current academic year and a launch link for it", async () => {
    expect(await page.locator("#year").inputValue()).toBe(academicYear(new Date()));
    const href = await page.locator("#launch").getAttribute("href");
    expect(href).toMatch(new RegExp(`^https://moodle\\.huji\\.ac\\.il/${academicYear(new Date())}/admin/tool/mobile/launch\\.php\\?`));
    expect(href).toContain("confirmed=1");
    const [popup] = await Promise.all([ctx.waitForEvent("page"), page.click("#launch")]);
    await popup.waitForLoadState();
    expect(await popup.locator("a").getAttribute("href")).toBe(LINK);
    await popup.close();
  });

  it("is Made By Safam: top bar, emblem and footer, with the mustache logo, and no privacy banner", async () => {
    for (const sel of [".brand", ".ring", ".signature"]) expect(await page.textContent(sel), sel).toMatch(/Made By Safam/i);
    const logoOk = await page.$$eval("img[src$='safam-logo.svg']", (imgs) => imgs.length >= 3 && imgs.every((i) => (i as HTMLImageElement).naturalWidth > 0));
    expect(logoOk).toBe(true);
    const emblemWidth = await page.$eval(".emblem-logo", (i) => i.getBoundingClientRect().width);
    expect(emblemWidth).toBeGreaterThanOrEqual(120); // the mustache leads the page
    expect(await page.getAttribute('link[rel="icon"]', "href")).toContain("safam-logo.svg");
    expect(await page.textContent("body")).not.toMatch(/Private by design/i);
  });

  it("tells you when the pasted text is not a Moodle app link", async () => {
    await page.fill("#link", "https://moodle.huji.ac.il/2026-27/my/");
    await expect.poll(() => page.textContent("#link-status")).toMatch(/not the Moodle app link/);
    await page.click("#build");
    await expect.poll(() => page.textContent("#error")).toMatch(/Paste the Moodle link/);
  });

  it("builds a new tracker and downloads it", async () => {
    await page.fill("#link", LINK);
    await expect.poll(() => page.textContent("#link-status")).toContain("Link recognised");
    const { file, name } = await buildAndDownload();
    expect(name).toMatch(/^Moodle assignments \d{4}-\d{2}-\d{2}\.xlsx$/);
    expect(await page.textContent("#result-title")).toBe("Your tracker is ready");
    expect(await page.textContent("#stat-added")).toBe("4");
    expect(await page.locator("#new-items li").allTextContents()).toHaveLength(4);
    const t = await readTracker(await XlsxPackage.load(await readFile(file)));
    expect(t.rows.filter((r) => !r.isEmpty).map((r) => r.moodleKey)).toEqual(["assign:2001", "quiz:3001", "assign:1002", "assign:1001"]);
    expect(await page.locator("#error").isVisible()).toBe(false);
    await page.waitForTimeout(600); // let the result card's fade-in finish
    await page.screenshot({ path: path.join(SHOTS, "result-en.png"), fullPage: true });
  });

  it("updates an uploaded tracker, keeping its file name", async () => {
    const first = await buildAndDownload();
    await page.click("#seg-update");
    await expect.poll(() => page.locator("#drop").isVisible()).toBe(true);
    await page.setInputFiles("#file", { name: "My tracker.xlsx", mimeType: "application/octet-stream", buffer: await readFile(first.file) });
    expect(await page.textContent("#drop-title")).toContain("My tracker.xlsx");
    const second = await buildAndDownload();
    expect(second.name).toBe("My tracker.xlsx");
    expect(await page.textContent("#result-title")).toBe("Already up to date");
  });

  it("explains a link Moodle rejects", async () => {
    await page.click("#seg-new");
    await page.fill("#link", "f".repeat(32));
    await page.click("#build");
    await expect.poll(() => page.textContent("#error")).toMatch(/didn't accept the link/);
  });

  it("switches to Hebrew, right to left, and remembers it", async () => {
    await page.click("#lang");
    expect(await page.getAttribute("html", "dir")).toBe("rtl");
    expect(await page.getAttribute("html", "lang")).toBe("he");
    expect(await page.textContent("h1")).toContain("כל הדדליינים");
    expect(await page.textContent("#error")).toContain("המודל לא קיבל את הקישור"); // the shown error is translated too
    expect(await page.textContent("#lang")).toBe("English");
    expect(await page.textContent(".signature")).toMatch(/Made By Safam/i); // the brand stays as is in Hebrew
    // ...and the stamp's circular text still sits on its ring (rtl once pushed it off the path)
    const ringWidth = await page.$eval(".ring text", (t) => (t as SVGTextElement).getBBox().width);
    expect(ringWidth).toBeGreaterThan(150);
    await page.reload();
    expect(await page.getAttribute("html", "dir")).toBe("rtl");
    await page.fill("#link", LINK);
    await expect.poll(() => page.textContent("#link-status")).toContain("הקישור זוהה");
    await page.click("#build");
    await page.locator("#download").waitFor({ state: "visible" });
    expect(await page.textContent("#result-title")).toBe("הטבלה שלך מוכנה");
    await page.waitForTimeout(600);
    await page.screenshot({ path: path.join(SHOTS, "result-he.png"), fullPage: true });
    await page.click("#lang");
    expect(await page.getAttribute("html", "dir")).toBe("ltr");
    expect(await page.textContent("#result-title")).toBe("Your tracker is ready");
  });

  it("talks only to itself and Moodle, with no console or CSP errors", () => {
    expect([...hosts].sort()).toEqual([ORIGIN, "https://moodle.huji.ac.il"].sort());
    expect(consoleErrors.filter((e) => !/invalidtoken|did not accept/i.test(e)).map((e) => e.slice(0, 300))).toEqual([]);
  });

  it("renders Hebrew and English in the bundled Karantina and Rubik fonts", async () => {
    const r = await page.evaluate(async () => {
      await document.fonts.ready;
      await document.fonts.load('700 40px "Karantina"', "שלום Hello");
      await document.fonts.load('400 16px "Rubik Variable"', "שלום Hello");
      const hebrewFace = (fam: string) => [...document.fonts].some((f) => f.family.replace(/"/g, "") === fam && f.status === "loaded" && /590/.test(f.unicodeRange));
      return {
        karantinaHe: hebrewFace("Karantina"),
        rubikHe: hebrewFace("Rubik Variable"),
        h1: getComputedStyle(document.querySelector("h1")!).fontFamily,
        body: getComputedStyle(document.body).fontFamily,
      };
    });
    expect(r).toMatchObject({ karantinaHe: true, rubikHe: true });
    expect(r.h1).toContain("Karantina");
    expect(r.body).toContain("Rubik");
  });

  it("fits a phone screen without sideways scrolling, in both languages and colour schemes", async () => {
    await page.setViewportSize({ width: 360, height: 740 });
    for (const scheme of ["light", "dark"] as const) {
      await page.emulateMedia({ colorScheme: scheme });
      for (const _ of [0, 1]) {
        const overflow = await page.evaluate(() => document.documentElement.scrollWidth - window.innerWidth);
        expect(overflow).toBeLessThanOrEqual(0);
        await page.screenshot({ path: path.join(SHOTS, `phone-${await page.getAttribute("html", "lang")}-${scheme}.png`), fullPage: true });
        await page.click("#lang");
      }
    }
  });
});
