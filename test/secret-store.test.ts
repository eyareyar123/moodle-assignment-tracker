import { describe, expect, it } from "vitest";
import { mkdtemp, readFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { TokenStore } from "../src/auth/secret-store.js";

describe.runIf(process.platform === "win32")("TokenStore (DPAPI)", () => {
  it("round-trips a token and never writes it in clear text", async () => {
    const dir = await mkdtemp(path.join(tmpdir(), "tokens-"));
    const s = new TokenStore(dir);
    await s.save({ siteUrl: "https://x/2026-27", token: "dummy-token-abc123", createdAt: "now" });
    expect(await readFile(path.join(dir, "token.dpapi"), "utf8")).not.toContain("dummy-token");
    expect((await s.load("https://x/2026-27"))?.token).toBe("dummy-token-abc123");
    expect(await s.load("https://x/2027-28")).toBeNull();
    await s.clear();
    expect(await s.load("https://x/2026-27")).toBeNull();
  }, 30_000);
});
