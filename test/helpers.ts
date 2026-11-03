import { copyFile, mkdir, readFile } from "node:fs/promises";
import { existsSync } from "node:fs";
import path from "node:path";

/**
 * Tests run against a private copy of the real workbook in test/fixtures (git-ignored,
 * because it holds personal data). It is refreshed from the project root when missing.
 */
export const FIXTURE = path.resolve("test/fixtures/tracker.xlsx");

export async function fixtureBytes(): Promise<Buffer> {
  if (!existsSync(FIXTURE)) {
    await mkdir(path.dirname(FIXTURE), { recursive: true });
    await copyFile(path.resolve("ER assignments.xlsx"), FIXTURE);
  }
  return readFile(FIXTURE);
}
