// Safe file handling around the workbook: lock check, timestamped backup, verify, atomic replace.

import { copyFile, mkdir, open, readFile, rename, rm, writeFile } from "node:fs/promises";
import { existsSync } from "node:fs";
import path from "node:path";
import ExcelJS from "exceljs";
import { XlsxPackage } from "./ooxml.js";
import { readTracker } from "./tracker.js";

/** Excel keeps "~$name.xlsx" next to an open workbook and holds an exclusive lock on the file. */
export async function assertNotOpenInExcel(file: string): Promise<void> {
  const owner = path.join(path.dirname(file), `~$${path.basename(file)}`);
  if (existsSync(owner)) throw new Error(`"${path.basename(file)}" looks open in Excel (found ${path.basename(owner)}). Close it and retry.`);
  try {
    const h = await open(file, "r+");
    await h.close();
  } catch (e) {
    const code = (e as NodeJS.ErrnoException).code;
    if (code === "EBUSY" || code === "EPERM" || code === "EACCES") {
      throw new Error(`"${path.basename(file)}" is locked by another program (${code}). Close Excel and retry.`);
    }
    throw e;
  }
}

export async function backup(file: string, backupDir: string): Promise<string> {
  await mkdir(backupDir, { recursive: true });
  const stamp = new Date().toISOString().replace(/[:.]/g, "-").slice(0, 19);
  const dest = path.join(backupDir, `${path.parse(file).name}.${stamp}.xlsx`);
  await copyFile(file, dest);
  return dest;
}

/** Check that our bytes parse as a valid tracker, with both our reader and ExcelJS as an independent one. */
export async function verifyWorkbook(bytes: Buffer): Promise<void> {
  await readTracker(await XlsxPackage.load(bytes));
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(bytes as unknown as ArrayBuffer);
  if (!wb.getWorksheet("Assignments")) throw new Error("Verification failed: Assignments sheet missing");
}

/** Back up, verify, then atomically replace the workbook. Returns the backup path. */
export async function saveWorkbook(file: string, bytes: Buffer, backupDir: string): Promise<string> {
  await assertNotOpenInExcel(file);
  await verifyWorkbook(bytes);
  const bak = await backup(file, backupDir);
  const tmp = `${file}.tmp`;
  await writeFile(tmp, bytes);
  try {
    await rename(tmp, file);
  } catch (e) {
    await rm(tmp, { force: true });
    throw e;
  }
  return bak;
}

export async function loadWorkbook(file: string): Promise<XlsxPackage> {
  return XlsxPackage.load(await readFile(file));
}
