// Thin typed client for Moodle's REST web services (webservice/rest/server.php).

export class MoodleError extends Error {
  constructor(
    readonly fn: string,
    readonly errorcode: string,
    message: string,
  ) {
    super(`${fn}: ${message} (${errorcode})`);
  }
  get isAuth(): boolean {
    return ["invalidtoken", "accessexception", "servicerequireslogin"].includes(this.errorcode);
  }
}

type Params = Record<string, unknown>;

/** Moodle's REST format wants nested arrays/objects flattened: courseids[0]=1, options[0][name]=x */
export function flatten(params: Params, prefix = "", out = new URLSearchParams()): URLSearchParams {
  for (const [k, v] of Object.entries(params)) {
    const key = prefix ? `${prefix}[${k}]` : k;
    if (v === undefined || v === null) continue;
    if (typeof v === "object") flatten(v as Params, key, out);
    else out.append(key, typeof v === "boolean" ? (v ? "1" : "0") : String(v));
  }
  return out;
}

export class MoodleClient {
  constructor(
    readonly siteUrl: string,
    private readonly token: string,
    private readonly fetchImpl: typeof fetch = fetch,
  ) {}

  async call<T>(fn: string, params: Params = {}): Promise<T> {
    const body = flatten({ wstoken: this.token, wsfunction: fn, moodlewsrestformat: "json", ...params });
    let lastErr: unknown;
    for (let attempt = 0; attempt < 3; attempt++) {
      try {
        const res = await this.fetchImpl(`${this.siteUrl}/webservice/rest/server.php`, {
          method: "POST",
          headers: { "Content-Type": "application/x-www-form-urlencoded" },
          body,
          signal: AbortSignal.timeout(30_000),
        });
        if (res.status >= 500) throw new Error(`HTTP ${res.status}`);
        const json = (await res.json()) as T & { exception?: string; errorcode?: string; message?: string };
        if (json && typeof json === "object" && "exception" in json && json.exception) {
          throw new MoodleError(fn, json.errorcode ?? "unknown", json.message ?? json.exception);
        }
        return json;
      } catch (e) {
        if (e instanceof MoodleError) throw e; // a Moodle-level error won't fix itself on retry
        lastErr = e;
        await new Promise((r) => setTimeout(r, 1000 * (attempt + 1)));
      }
    }
    throw new Error(`${fn}: network error after 3 attempts: ${(lastErr as Error)?.message}`);
  }
}
