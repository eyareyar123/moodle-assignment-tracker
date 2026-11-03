// A small fake of Moodle's REST web service, shared by the unit and browser tests.

export const FAKE_TOKEN = "0123456789abcdef0123456789abcdef";

const ts = (iso: string) => Date.parse(`${iso}T21:59:00Z`) / 1000; // 23:59 Israel time

/** Answers for each wsfunction; per-activity calls are keyed by their id. */
export function fakeMoodleAnswer(fn: string, params: URLSearchParams): unknown {
  switch (fn) {
    case "core_webservice_get_site_info":
      return { userid: 7, fullname: "Test Student", sitename: "HUJI" };
    case "core_enrol_get_users_courses":
      return [
        { id: 1, shortname: "57340-2027", fullname: "57340 - {mlang he}סטטיסטיקה{mlang}{mlang en}Statistics{mlang}", visible: 1 },
        { id: 2, shortname: "51033", fullname: "51033 - תכנות R", visible: 1 },
        { id: 3, shortname: "hidden", fullname: "99999 - Hidden", visible: 0 },
      ];
    case "mod_assign_get_assignments":
      return {
        courses: [
          { id: 1, assignments: [
            { id: 11, cmid: 1001, course: 1, name: "הגשת תרגיל 1", duedate: ts("2026-10-20"), nosubmissions: 0, grade: 100 },
            { id: 12, cmid: 1002, course: 1, name: "הגשת תרגיל 2", duedate: ts("2026-11-10"), nosubmissions: 0, grade: 100 },
          ] },
          { id: 2, assignments: [{ id: 21, cmid: 2001, course: 2, name: "הגשת קוד", duedate: ts("2026-11-05"), nosubmissions: 0, grade: 10 }] },
        ],
      };
    case "mod_assign_get_submission_status": {
      const id = params.get("assignid");
      if (id === "11") return { lastattempt: { submission: { status: "submitted" }, gradingstatus: "graded" }, feedback: { grade: { grade: "90" } } };
      if (id === "12") return { lastattempt: { submission: { status: "submitted" }, gradingstatus: "notgraded" } };
      return { lastattempt: { submission: { status: "new" } } };
    }
    case "mod_quiz_get_quizzes_by_courses":
      return {
        quizzes: [
          { id: 31, coursemodule: 3001, course: 2, name: "יחידה 2: dplyr וtidyr", timeclose: ts("2026-11-12"), grade: 10 },
          { id: 32, coursemodule: 3002, course: 2, name: "תרגול חופשי", timeclose: 0, grade: 0 },
        ],
      };
    case "mod_quiz_get_user_attempts":
      return { attempts: [] };
    case "mod_quiz_get_user_best_grade":
      return { hasgrade: false };
    default:
      return { exception: "moodle_exception", errorcode: "invalidrecord", message: `unknown function ${fn}` };
  }
}

/** A fetch that serves the fake; any token but FAKE_TOKEN is rejected the way Moodle does. */
export function fakeFetch(calls: string[] = []): typeof fetch {
  return (async (_url: string | URL | Request, init?: RequestInit) => {
    const params = new URLSearchParams(String(init?.body ?? ""));
    const fn = params.get("wsfunction") ?? "";
    calls.push(fn);
    const body = params.get("wstoken") !== FAKE_TOKEN
      ? { exception: "moodle_exception", errorcode: "invalidtoken", message: "Invalid token - token not found" }
      : fakeMoodleAnswer(fn, params);
    return new Response(JSON.stringify(body), { headers: { "Content-Type": "application/json" } });
  }) as typeof fetch;
}
