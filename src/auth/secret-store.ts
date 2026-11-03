// Stores the Moodle token encrypted with Windows DPAPI (CurrentUser scope): only your Windows
// account on this machine can decrypt it. Node has no DPAPI binding, so we call PowerShell.

import { execFile } from "node:child_process";
import { mkdir, readFile, rm, writeFile } from "node:fs/promises";
import { existsSync } from "node:fs";
import path from "node:path";

function dpapi(mode: "Protect" | "Unprotect", inputB64: string): Promise<string> {
  const script =
    "Add-Type -AssemblyName System.Security;" +
    "$in=[Convert]::FromBase64String([Console]::In.ReadToEnd().Trim());" +
    `$out=[Security.Cryptography.ProtectedData]::${mode}($in,$null,[Security.Cryptography.DataProtectionScope]::CurrentUser);` +
    "[Console]::Out.Write([Convert]::ToBase64String($out))";
  return new Promise((resolve, reject) => {
    const p = execFile("powershell.exe", ["-NoProfile", "-NonInteractive", "-Command", script], (err, stdout, stderr) =>
      err ? reject(new Error(`DPAPI ${mode} failed: ${stderr || err.message}`)) : resolve(stdout.trim()),
    );
    p.stdin!.end(inputB64);
  });
}

export interface StoredToken {
  siteUrl: string;
  token: string;
  privateToken?: string;
  createdAt: string;
}

export class TokenStore {
  constructor(private readonly dir: string) {}

  private get file() {
    return path.join(this.dir, "token.dpapi");
  }

  async save(t: StoredToken): Promise<void> {
    await mkdir(this.dir, { recursive: true });
    const enc = await dpapi("Protect", Buffer.from(JSON.stringify(t), "utf8").toString("base64"));
    await writeFile(this.file, enc, "utf8");
  }

  async load(siteUrl: string): Promise<StoredToken | null> {
    if (!existsSync(this.file)) return null;
    const dec = await dpapi("Unprotect", (await readFile(this.file, "utf8")).trim());
    const t = JSON.parse(Buffer.from(dec, "base64").toString("utf8")) as StoredToken;
    return t.siteUrl === siteUrl ? t : null; // token from last year's site is useless
  }

  async clear(): Promise<void> {
    await rm(this.file, { force: true });
  }
}
