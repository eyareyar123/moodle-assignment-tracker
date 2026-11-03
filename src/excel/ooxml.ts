// Minimal, surgical access to an .xlsx package.
//
// ExcelJS rewrites the whole workbook on save and silently drops charts, x14 data
// validations, threaded-comment persons, etc. Instead we load the zip, edit only
// the XML parts we need through a DOM, and write every other part back untouched.

import JSZip from "jszip";
import { DOMParser, XMLSerializer, type Document, type Element } from "@xmldom/xmldom";

export const NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
export const NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
export const NS_PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships";
export const REL_HYPERLINK = `${NS_REL}/hyperlink`;
export const REL_TABLE = `${NS_REL}/table`;

const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n';

// ---------- A1 address helpers ----------

export function colToIndex(col: string): number {
  let n = 0;
  for (const ch of col.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n;
}

export function indexToCol(n: number): string {
  let s = "";
  while (n > 0) {
    const r = (n - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

export function splitRef(ref: string): { col: string; row: number } {
  const m = /^([A-Z]+)(\d+)$/.exec(ref.replace(/\$/g, ""));
  if (!m) throw new Error(`Bad cell reference: ${ref}`);
  return { col: m[1]!, row: Number(m[2]) };
}

// ---------- Excel date serials ----------

/** Serial number for a calendar date (no time part), Excel 1900 date system. */
export function dateToSerial(y: number, m: number, d: number): number {
  return Date.UTC(y, m - 1, d) / 86_400_000 + 25_569;
}

/** Calendar date "YYYY-MM-DD" for a serial (fractional part ignored). */
export function serialToIsoDate(serial: number): string {
  return new Date(Math.round(Math.floor(serial) - 25_569) * 86_400_000).toISOString().slice(0, 10);
}

// ---------- DOM helpers ----------

export function children(el: Element, localName?: string): Element[] {
  const out: Element[] = [];
  for (let n = el.firstChild; n; n = n.nextSibling) {
    if (n.nodeType === 1 && (!localName || (n as Element).localName === localName)) out.push(n as Element);
  }
  return out;
}

export function child(el: Element, localName: string): Element | undefined {
  return children(el, localName)[0];
}

export function textOf(el: Element | undefined): string {
  return el?.textContent ?? "";
}

// ---------- Package ----------

export class XlsxPackage {
  private docs = new Map<string, Document>();

  private constructor(readonly zip: JSZip) {}

  static async load(data: Uint8Array | ArrayBuffer): Promise<XlsxPackage> {
    return new XlsxPackage(await JSZip.loadAsync(data));
  }

  has(path: string): boolean {
    return this.zip.file(path) !== null;
  }

  async xml(path: string): Promise<Document> {
    let doc = this.docs.get(path);
    if (!doc) {
      const f = this.zip.file(path);
      if (!f) throw new Error(`Missing package part: ${path}`);
      doc = new DOMParser().parseFromString(await f.async("string"), "text/xml") as unknown as Document;
      this.docs.set(path, doc);
    }
    return doc;
  }

  /** Mark a part as edited; it will be re-serialised on save. Parts never touched stay byte-identical. */
  private dirty = new Set<string>();
  touch(path: string): void {
    this.dirty.add(path);
  }

  remove(path: string): void {
    this.zip.remove(path);
    this.docs.delete(path);
    this.dirty.delete(path);
  }

  async save(): Promise<Uint8Array> {
    for (const path of this.dirty) {
      const body = new XMLSerializer().serializeToString(this.docs.get(path)!).replace(/^<\?xml[^>]*\?>\s*/, "");
      this.zip.file(path, XML_DECL + body);
    }
    return this.zip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
  }

  // ----- relationships -----

  static relsPathFor(partPath: string): string {
    const i = partPath.lastIndexOf("/");
    return `${partPath.slice(0, i)}/_rels/${partPath.slice(i + 1)}.rels`;
  }

  static resolveTarget(fromPart: string, target: string): string {
    if (target.startsWith("/")) return target.slice(1);
    const parts = fromPart.split("/").slice(0, -1);
    for (const seg of target.split("/")) {
      if (seg === "..") parts.pop();
      else if (seg !== ".") parts.push(seg);
    }
    return parts.join("/");
  }

  async rels(partPath: string): Promise<{ id: string; type: string; target: string; external: boolean }[]> {
    const p = XlsxPackage.relsPathFor(partPath);
    if (!this.has(p)) return [];
    const doc = await this.xml(p);
    return children(doc.documentElement!, "Relationship").map((r) => ({
      id: r.getAttribute("Id")!,
      type: r.getAttribute("Type")!,
      target: r.getAttribute("Target")!,
      external: r.getAttribute("TargetMode") === "External",
    }));
  }

  async addExternalRel(partPath: string, type: string, target: string): Promise<string> {
    const p = XlsxPackage.relsPathFor(partPath);
    const doc = await this.xml(p);
    const root = doc.documentElement!;
    const used = new Set(children(root, "Relationship").map((r) => r.getAttribute("Id")));
    let n = 1;
    while (used.has(`rId${n}`)) n++;
    const rel = doc.createElementNS(NS_PKG_REL, "Relationship");
    rel.setAttribute("Id", `rId${n}`);
    rel.setAttribute("Type", type);
    rel.setAttribute("Target", target);
    rel.setAttribute("TargetMode", "External");
    root.appendChild(rel);
    this.touch(p);
    return `rId${n}`;
  }

  async setRelTarget(partPath: string, id: string, target: string): Promise<void> {
    const p = XlsxPackage.relsPathFor(partPath);
    const doc = await this.xml(p);
    const rel = children(doc.documentElement!, "Relationship").find((r) => r.getAttribute("Id") === id);
    if (!rel) throw new Error(`No relationship ${id} in ${p}`);
    rel.setAttribute("Target", target);
    this.touch(p);
  }

  // ----- workbook structure -----

  async sheetPath(sheetName: string): Promise<string> {
    const wb = await this.xml("xl/workbook.xml");
    const sheet = children(child(wb.documentElement!, "sheets")!, "sheet").find((s) => s.getAttribute("name") === sheetName);
    if (!sheet) throw new Error(`Sheet not found: ${sheetName}`);
    const rid = sheet.getAttributeNS(NS_REL, "id");
    const rel = (await this.rels("xl/workbook.xml")).find((r) => r.id === rid);
    if (!rel) throw new Error(`No relationship for sheet ${sheetName}`);
    return XlsxPackage.resolveTarget("xl/workbook.xml", rel.target);
  }

  async tablePath(sheetPath: string, tableName: string): Promise<string> {
    for (const rel of await this.rels(sheetPath)) {
      if (rel.type !== REL_TABLE) continue;
      const p = XlsxPackage.resolveTarget(sheetPath, rel.target);
      if ((await this.xml(p)).documentElement!.getAttribute("name") === tableName) return p;
    }
    throw new Error(`Table ${tableName} not found on ${sheetPath}`);
  }

  async sharedStrings(): Promise<string[]> {
    if (!this.has("xl/sharedStrings.xml")) return [];
    const doc = await this.xml("xl/sharedStrings.xml");
    // A string item is either <t> or rich-text runs <r><t/></r>; phonetic <rPh> is ignored.
    return children(doc.documentElement!, "si").map((si) =>
      children(si)
        .filter((n) => n.localName === "t" || n.localName === "r")
        .map((n) => (n.localName === "t" ? textOf(n) : textOf(child(n, "t"))))
        .join(""),
    );
  }

  /** Force Excel to recalculate everything on open, and drop the (now stale) calc chain. */
  async requestFullRecalc(): Promise<void> {
    const wb = await this.xml("xl/workbook.xml");
    let calcPr = child(wb.documentElement!, "calcPr");
    if (!calcPr) {
      calcPr = wb.createElementNS(NS_MAIN, "calcPr");
      wb.documentElement!.appendChild(calcPr);
    }
    calcPr.setAttribute("fullCalcOnLoad", "1");
    this.touch("xl/workbook.xml");

    if (this.has("xl/calcChain.xml")) {
      this.remove("xl/calcChain.xml");
      const rels = await this.xml("xl/_rels/workbook.xml.rels");
      for (const r of children(rels.documentElement!, "Relationship")) {
        if (r.getAttribute("Target")?.endsWith("calcChain.xml")) rels.documentElement!.removeChild(r);
      }
      this.touch("xl/_rels/workbook.xml.rels");
      const ct = await this.xml("[Content_Types].xml");
      for (const o of children(ct.documentElement!, "Override")) {
        if (o.getAttribute("PartName") === "/xl/calcChain.xml") ct.documentElement!.removeChild(o);
      }
      this.touch("[Content_Types].xml");
    }
  }
}

// ---------- Worksheet cell access ----------

export type CellValue = string | number | boolean | null;

export interface CellRead {
  value: CellValue;
  formula?: string;
  style?: string;
}

/** Wraps one worksheet DOM, with cell read/write that preserves styles and ordering. */
export class Sheet {
  readonly sheetData: Element;

  constructor(
    readonly pkg: XlsxPackage,
    readonly path: string,
    readonly doc: Document,
    private readonly shared: string[],
  ) {
    this.sheetData = child(doc.documentElement!, "sheetData")!;
  }

  static async open(pkg: XlsxPackage, sheetName: string): Promise<Sheet> {
    const path = await pkg.sheetPath(sheetName);
    return new Sheet(pkg, path, await pkg.xml(path), await pkg.sharedStrings());
  }

  touch(): void {
    this.pkg.touch(this.path);
  }

  rowEl(r: number, create = false): Element | undefined {
    let before: Element | null = null;
    for (const row of children(this.sheetData, "row")) {
      const n = Number(row.getAttribute("r"));
      if (n === r) return row;
      if (n > r) {
        before = row;
        break;
      }
    }
    if (!create) return undefined;
    const row = this.doc.createElementNS(NS_MAIN, "row");
    row.setAttribute("r", String(r));
    this.sheetData.insertBefore(row, before);
    return row;
  }

  cellEl(ref: string, create = false): Element | undefined {
    const { col, row } = splitRef(ref);
    const rowEl = this.rowEl(row, create);
    if (!rowEl) return undefined;
    const ci = colToIndex(col);
    let before: Element | null = null;
    for (const c of children(rowEl, "c")) {
      const cc = splitRef(c.getAttribute("r")!).col;
      if (cc === col) return c;
      if (colToIndex(cc) > ci) {
        before = c;
        break;
      }
    }
    if (!create) return undefined;
    const c = this.doc.createElementNS(NS_MAIN, "c");
    c.setAttribute("r", ref);
    rowEl.insertBefore(c, before);
    rowEl.removeAttribute("spans"); // spans is only an optimisation hint; drop it rather than keep it wrong
    return c;
  }

  read(ref: string): CellRead {
    const c = this.cellEl(ref);
    if (!c) return { value: null };
    const t = c.getAttribute("t") ?? "n";
    const f = child(c, "f");
    const v = child(c, "v");
    const style = c.getAttribute("s") ?? undefined;
    let value: CellValue = null;
    if (t === "inlineStr") value = textOf(child(child(c, "is")!, "t")) || textOf(child(c, "is"));
    else if (v && textOf(v) !== "") {
      const raw = textOf(v);
      if (t === "s") value = this.shared[Number(raw)] ?? "";
      else if (t === "str" || t === "e") value = raw;
      else if (t === "b") value = raw === "1";
      else value = Number(raw);
    }
    return { value, formula: f ? textOf(f) || "(shared)" : undefined, style };
  }

  /** Write a literal value. Keeps the cell's style; refuses to overwrite a formula. */
  write(ref: string, value: CellValue, opts: { style?: string } = {}): void {
    const c = this.cellEl(ref, true)!;
    if (child(c, "f")) throw new Error(`Refusing to overwrite formula in ${this.path}!${ref}`);
    while (c.firstChild) c.removeChild(c.firstChild);
    c.removeAttribute("t");
    if (opts.style !== undefined) c.setAttribute("s", opts.style);
    if (value === null || value === "") {
      // empty cell keeps its style only
    } else if (typeof value === "number") {
      this.append(c, "v", String(value));
    } else if (typeof value === "boolean") {
      c.setAttribute("t", "b");
      this.append(c, "v", value ? "1" : "0");
    } else {
      c.setAttribute("t", "inlineStr");
      const is = this.doc.createElementNS(NS_MAIN, "is");
      const t = this.doc.createElementNS(NS_MAIN, "t");
      if (/^\s|\s$/.test(value)) t.setAttribute("xml:space", "preserve");
      t.appendChild(this.doc.createTextNode(value));
      is.appendChild(t);
      c.appendChild(is);
    }
    this.touch();
  }

  /** Write (or replace) a formula; the cached value is cleared so Excel recalculates it. */
  writeFormula(ref: string, formula: string, opts: { style?: string; volatile?: boolean } = {}): void {
    const c = this.cellEl(ref, true)!;
    while (c.firstChild) c.removeChild(c.firstChild);
    c.removeAttribute("t");
    if (opts.style !== undefined) c.setAttribute("s", opts.style);
    const f = this.append(c, "f", formula);
    if (opts.volatile) f.setAttribute("ca", "1");
    this.touch();
  }

  private append(parent: Element, name: string, text: string): Element {
    const el = this.doc.createElementNS(NS_MAIN, name);
    el.appendChild(this.doc.createTextNode(text));
    parent.appendChild(el);
    return el;
  }

  // ----- hyperlinks -----

  async hyperlinks(): Promise<Map<string, { relId: string; target: string }>> {
    const out = new Map<string, { relId: string; target: string }>();
    const hl = child(this.doc.documentElement!, "hyperlinks");
    if (!hl) return out;
    const rels = new Map((await this.pkg.rels(this.path)).map((r) => [r.id, r.target]));
    for (const h of children(hl, "hyperlink")) {
      const relId = h.getAttributeNS(NS_REL, "id");
      if (relId && rels.has(relId)) out.set(h.getAttribute("ref")!, { relId, target: rels.get(relId)! });
    }
    return out;
  }

  async setHyperlink(ref: string, url: string): Promise<void> {
    const existing = (await this.hyperlinks()).get(ref);
    if (existing) {
      if (existing.target !== url) await this.pkg.setRelTarget(this.path, existing.relId, url);
      return;
    }
    const relId = await this.pkg.addExternalRel(this.path, REL_HYPERLINK, url);
    const root = this.doc.documentElement!;
    let hl = child(root, "hyperlinks");
    if (!hl) {
      // CT_Worksheet order: ... conditionalFormatting, dataValidations, hyperlinks, printOptions, pageMargins ...
      hl = this.doc.createElementNS(NS_MAIN, "hyperlinks");
      const after = ["printOptions", "pageMargins", "pageSetup", "headerFooter", "drawing", "tableParts", "extLst"];
      const anchor = children(root).find((e) => after.includes(e.localName!)) ?? null;
      root.insertBefore(hl, anchor);
    }
    const h = this.doc.createElementNS(NS_MAIN, "hyperlink");
    h.setAttribute("ref", ref);
    h.setAttributeNS(NS_REL, "r:id", relId);
    hl.appendChild(h);
    this.touch();
  }
}
