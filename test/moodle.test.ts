import { describe, expect, it } from "vitest";
import { createHash } from "node:crypto";
import { flatten, MoodleClient, MoodleError } from "../src/moodle/client.js";
import { assignStatus, cleanText, courseCode, courseTitle, dueDate, fetchItems } from "../src/moodle/fetch.js";
import { decodeLaunchUrl } from "../src/auth/login.js";
import { launchUrl, parseLaunchToken } from "../src/moodle/launch.js";
import { keyFromUrl } from "../src/moodle/types.js";

const SITE = "https://moodle.huji.ac.il/2026-27";

describe("moodle helpers", () => {
  it("flattens nested params the way Moodle REST expects", () => {
    expect(flatten({ courseids: [3, 4], options: [{ name: "x", value: true }] }).toString()).toBe(
      "courseids%5B0%5D=3&courseids%5B1%5D=4&options%5B0%5D%5Bname%5D=x&options%5B0%5D%5Bvalue%5D=1",
    );
  });
  it("parses course codes", () => {
    expect(courseCode({ shortname: "57340-2027", fullname: "x" })).toBe("57340");
    expect(courseCode({ shortname: "abc", fullname: "51102-חשיבה סטטיסטית" })).toBe("51102");
    expect(courseCode({ shortname: "abc", fullname: "general" })).toBeNull();
  });
  it("converts due timestamps to Israel dates, midnight belonging to the day before", () => {
    expect(dueDate(Date.UTC(2026, 10, 5, 21, 59) / 1000, "Asia/Jerusalem")).toBe("2026-11-05"); // 23:59 IST
    expect(dueDate(Date.UTC(2026, 10, 5, 22, 0) / 1000, "Asia/Jerusalem")).toBe("2026-11-05"); // 00:00 on the 6th
    expect(dueDate(0, "Asia/Jerusalem")).toBeNull();
  });
  it("maps submission status", () => {
    expect(assignStatus({ lastattempt: { submission: { status: "new" } } }).status).toBe("not_submitted");
    expect(assignStatus({ lastattempt: { submission: { status: "draft" } } }).status).toBe("draft");
    expect(assignStatus({ lastattempt: { submission: { status: "submitted" } } }).status).toBe("submitted");
    expect(assignStatus({ lastattempt: { submission: { status: "submitted" } }, feedback: { grade: { grade: "85.00000" } } })).toEqual({ status: "graded", rawGrade: 85 });
  });
  it("resolves multi-language text to Hebrew and trims the course number", () => {
    const full = "51033 - {mlang he}תכנות ויישום R במחקר בפסיכולוגיה{mlang}{mlang en}R Programming for Psychology Research{mlang}";
    expect(cleanText(full)).toBe("51033 - תכנות ויישום R במחקר בפסיכולוגיה");
    expect(cleanText(full, "en")).toBe("51033 - R Programming for Psychology Research");
    expect(courseTitle(full, "51033")).toBe("תכנות ויישום R במחקר בפסיכולוגיה");
    expect(cleanText('<span lang="en" class="multilang">Hi</span><span lang="he" class="multilang">שלום</span>')).toBe("שלום");
    expect(cleanText("  plain   name ")).toBe("plain name");
    expect(cleanText("עריכת נתונים באמצעות החבילות dplyr וtidyr")).toBe("עריכת נתונים באמצעות החבילות dplyr ו-tidyr");
    expect(cleanText("ויישום R במחקר")).toBe("ויישום R במחקר");
  });
  it("parses keys from tracker links", () => {
    expect(keyFromUrl(`${SITE}/mod/quiz/view.php?id=108782`)).toBe("quiz:108782");
    expect(keyFromUrl(`${SITE}/mod/forum/view.php?id=1`)).toBeNull();
  });
  it("verifies the SSO token signature", () => {
    const passport = "12345";
    const sig = createHash("md5").update(SITE + passport).digest("hex");
    const url = `moodlemobile://token=${Buffer.from(`${sig}:::TOKEN:::PRIV`).toString("base64")}`;
    expect(decodeLaunchUrl(url, SITE, passport)).toEqual({ token: "TOKEN", privateToken: "PRIV" });
    expect(() => decodeLaunchUrl(url, SITE, "999")).toThrow(/signature/);
  });
  it("reads a copied app link, URL-encoded or bare, and rejects anything else", () => {
    const tok = "0123456789abcdef0123456789abcdef";
    const link = `moodlemobile://token=${Buffer.from(`sig:::${tok}`).toString("base64")}`;
    expect(parseLaunchToken(link)).toEqual({ signature: "sig", token: tok, privateToken: undefined });
    expect(parseLaunchToken(`  ${encodeURIComponent(link)}\n`).token).toBe(tok);
    expect(parseLaunchToken(tok)).toEqual({ token: tok });
    expect(() => parseLaunchToken("https://moodle.huji.ac.il/2026-27/my/")).toThrow(/not a Moodle app link/);
    expect(() => parseLaunchToken(`moodlemobile://token=${Buffer.from("sig:::").toString("base64")}`)).toThrow(/no token/);
  });
  it("builds the launch URL, with the copyable page when confirmed", () => {
    expect(launchUrl(SITE, "42", { confirmed: true })).toBe(
      `${SITE}/admin/tool/mobile/launch.php?service=moodle_mobile_app&passport=42&urlscheme=moodlemobile&confirmed=1`,
    );
  });
});

describe("fetchItems against a fake Moodle", () => {
  const responses: Record<string, unknown> = {
    core_webservice_get_site_info: { userid: 7, fullname: "Test", sitename: "HUJI" },
    core_enrol_get_users_courses: [{ id: 1, shortname: "57340", fullname: "57340-סטטיסטיקה", visible: 1 }],
    mod_assign_get_assignments: { courses: [{ id: 1, assignments: [{ id: 10, cmid: 47410, course: 1, name: "תרגיל 1", duedate: 1793051940, nosubmissions: 0, grade: 100 }] }] },
    mod_assign_get_submission_status: { lastattempt: { submission: { status: "submitted" }, gradingstatus: "graded" }, feedback: { grade: { grade: "90" } } },
    mod_quiz_get_quizzes_by_courses: { quizzes: [{ id: 5, coursemodule: 108782, course: 1, name: "בוחן", timeclose: 0, grade: 10 }] },
    mod_quiz_get_user_attempts: { attempts: [{ state: "finished" }] },
    mod_quiz_get_user_best_grade: { hasgrade: true, grade: 8 },
  };
  const fakeFetch = (async (_url: string, init: RequestInit) => {
    const fn = new URLSearchParams(init.body as URLSearchParams).get("wsfunction")!;
    return new Response(JSON.stringify(responses[fn] ?? { exception: "x", errorcode: "nofunc", message: fn }));
  }) as typeof fetch;

  it("normalises assignments and quizzes", async () => {
    const items = await fetchItems(new MoodleClient(SITE, "t", fakeFetch), { timezone: "Asia/Jerusalem" });
    expect(items).toHaveLength(2);
    expect(items[0]).toMatchObject({ key: "assign:47410", courseCode: "57340", status: "graded", gradePercent: 90, url: `${SITE}/mod/assign/view.php?id=47410` });
    expect(items[1]).toMatchObject({ key: "quiz:108782", status: "graded", gradePercent: 80, due: null });
  });

  it("surfaces an invalid token as an auth error", async () => {
    const bad = (async () => new Response(JSON.stringify({ exception: "moodle_exception", errorcode: "invalidtoken", message: "Invalid token" }))) as unknown as typeof fetch;
    const err = await fetchItems(new MoodleClient(SITE, "t", bad), { timezone: "Asia/Jerusalem" }).catch((e) => e);
    expect(err).toBeInstanceOf(MoodleError);
    expect((err as MoodleError).isAuth).toBe(true);
  });
});
