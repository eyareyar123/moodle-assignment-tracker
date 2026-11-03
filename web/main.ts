// The page's wiring. Sync logic lives in src/web/core.ts, text in web/i18n.ts.

// Self-hosted fonts, so the page still only talks to itself and Moodle. Both cover Hebrew and Latin.
import "@fontsource/karantina/400.css";
import "@fontsource/karantina/700.css";
import "@fontsource-variable/rubik";

import { launchUrl, parseLaunchToken } from "../src/moodle/launch.js";
import { buildTracker, downloadName, siteForYear, todayIn, TIMEZONE, UserError, yearChoices, type BuildResult } from "../src/web/core.js";
import { initialLang, t, type Key, type Lang } from "./i18n.js";

const $ = <T extends HTMLElement>(id: string) => document.getElementById(id) as T;
const form = $<HTMLFormElement>("form");
const year = $<HTMLSelectElement>("year");
const launch = $<HTMLAnchorElement>("launch");
const link = $<HTMLInputElement>("link");
const linkStatus = $<HTMLParagraphElement>("link-status");
const drop = $<HTMLLabelElement>("drop");
const file = $<HTMLInputElement>("file");
const undated = $<HTMLInputElement>("undated");
const build = $<HTMLButtonElement>("build");
const progress = $<HTMLElement>("progress");
const stagesList = $<HTMLOListElement>("stages");
const result = $<HTMLElement>("result");
const errorBox = $<HTMLParagraphElement>("error");

type Msg = { key: Key; vars?: Record<string, string | number> };
const state: {
  lang: Lang;
  stages: (Msg & { done: boolean })[];
  error?: Msg;
  result?: { r: BuildResult; existingName?: string; fileName: string };
} = { lang: "en", stages: [] };
let downloadUrl: string | undefined;

// ---------- language ----------

const storage = {
  get: (k: string) => { try { return localStorage.getItem(k); } catch { return null; } },
  set: (k: string, v: string) => { try { localStorage.setItem(k, v); } catch { /* private mode: fine */ } },
};

const tr = (m: Msg) => t(state.lang, m.key, m.vars);

function render() {
  const L = state.lang;
  document.documentElement.lang = L;
  document.documentElement.dir = L === "he" ? "rtl" : "ltr";
  document.title = t(L, "docTitle");
  for (const el of document.querySelectorAll<HTMLElement>("[data-i18n]")) el.textContent = t(L, el.dataset.i18n as Key);
  for (const el of document.querySelectorAll<HTMLElement>("[data-i18n-html]")) el.innerHTML = t(L, el.dataset.i18nHtml as Key);
  const langBtn = $<HTMLButtonElement>("lang");
  langBtn.textContent = t(L, "switchLang");
  langBtn.setAttribute("aria-label", t(L, "switchLangLabel"));
  link.placeholder = t(L, "pastePlaceholder");
  $("build-label").textContent = t(L, build.disabled ? "building" : "build");
  checkLink();
  renderFile();
  renderStages();
  renderResult();
  errorBox.hidden = !state.error;
  if (state.error) errorBox.textContent = tr(state.error);
}

$("lang").addEventListener("click", () => {
  state.lang = state.lang === "en" ? "he" : "en";
  storage.set("lang", state.lang);
  render();
});

// ---------- step 1 & 2: site and link ----------

// Isolated left-to-right, so "2026–27" doesn't read "27–2026" in Hebrew.
for (const y of yearChoices(new Date())) year.add(new Option(`⁦${y.replace("-", "–")}⁩`, y));
// A fresh passport per page load; Moodle only uses it to sign the link it hands back.
const passport = String(crypto.getRandomValues(new Uint32Array(1))[0]);
const site = () => siteForYear(year.value);
const updateLaunch = () => (launch.href = launchUrl(site(), passport, { confirmed: true }));
year.addEventListener("change", updateLaunch);
updateLaunch();

function checkLink(): string | null {
  linkStatus.className = "hint";
  linkStatus.textContent = "";
  if (!link.value.trim()) return null;
  try {
    const { token } = parseLaunchToken(link.value);
    linkStatus.textContent = t(state.lang, "linkOk");
    linkStatus.classList.add("ok");
    return token;
  } catch {
    linkStatus.textContent = t(state.lang, "linkBad");
    linkStatus.classList.add("bad");
    return null;
  }
}
link.addEventListener("input", checkLink);

$("paste").addEventListener("click", async () => {
  try {
    link.value = (await navigator.clipboard.readText()).trim();
    checkLink();
  } catch {
    link.focus(); // clipboard blocked: let the visitor paste by hand
  }
});

// ---------- step 3: new or update ----------

const mode = () => (form.elements.namedItem("mode") as RadioNodeList).value;
form.addEventListener("change", (e) => {
  if ((e.target as HTMLInputElement).name === "mode") drop.hidden = mode() !== "update";
});

function renderFile() {
  const f = file.files?.[0];
  drop.classList.toggle("chosen", !!f);
  $("drop-title").textContent = f ? t(state.lang, "dropChosen", { name: f.name }) : t(state.lang, "dropTitle");
}
file.addEventListener("change", renderFile);
drop.addEventListener("dragover", (e) => (e.preventDefault(), drop.classList.add("over")));
drop.addEventListener("dragleave", () => drop.classList.remove("over"));
drop.addEventListener("drop", (e) => {
  e.preventDefault();
  drop.classList.remove("over");
  if (e.dataTransfer?.files.length) {
    file.files = e.dataTransfer.files;
    renderFile();
  }
});

// ---------- progress and result ----------

function stage(key: Key, vars?: Msg["vars"]) {
  for (const s of state.stages) s.done = true;
  state.stages.push({ key, vars, done: false });
  renderStages();
}

function renderStages() {
  stagesList.replaceChildren(
    ...state.stages.map((s) => {
      const li = document.createElement("li");
      li.textContent = tr(s);
      li.classList.toggle("done", s.done);
      return li;
    }),
  );
}

function renderResult() {
  const res = state.result;
  result.hidden = !res;
  if (!res) return;
  const { r, existingName, fileName } = res;
  const acts = r.outcome.plan.actions;
  const adds = acts.filter((a) => a.kind === "add").sort((a, b) => (a.item.due ?? "9999").localeCompare(b.item.due ?? "9999"));
  $("result-title").textContent = t(state.lang, r.outcome.changed ? "readyTitle" : "upToDateTitle");
  $("stat-added").textContent = String(r.outcome.applied.added);
  $("stat-updated").textContent = String(r.outcome.applied.updated);
  $("stat-total").textContent = String(acts.filter((a) => a.kind !== "skip").length);
  const notes: string[] = [];
  if (r.undatedSkipped) notes.push(t(state.lang, "undatedLeft", { n: r.undatedSkipped }));
  if (existingName) notes.push(t(state.lang, "saveOver"));
  $("result-note").textContent = notes.join(" ");
  $("download-name").textContent = fileName;

  const fmt = new Intl.DateTimeFormat(state.lang === "he" ? "he-IL" : "en-GB", { day: "numeric", month: "short" });
  const shown = adds.slice(0, 8);
  $("new-items-box").hidden = adds.length === 0;
  $("new-items").replaceChildren(
    ...shown.map((a) => {
      const li = document.createElement("li");
      const name = document.createElement("span");
      name.textContent = a.item.name;
      name.dir = "auto";
      const due = document.createElement("span");
      due.className = "due";
      due.textContent = a.item.due ? fmt.format(new Date(`${a.item.due}T12:00:00`)) : "";
      li.append(name, due);
      return li;
    }),
    ...(adds.length > shown.length ? [Object.assign(document.createElement("li"), { textContent: t(state.lang, "andMore", { n: adds.length - shown.length }) })] : []),
  );
}

function fail(m: Msg) {
  state.error = m;
  render();
}

form.addEventListener("submit", async (ev) => {
  ev.preventDefault();
  state.error = undefined;
  state.result = undefined;
  render();
  const token = checkLink();
  if (!token) return fail({ key: "errNeedLink" });
  const existingFile = mode() === "update" ? file.files?.[0] : undefined;
  if (mode() === "update" && !existingFile) return fail({ key: "errNeedFile" });

  build.disabled = true;
  state.stages = [];
  progress.hidden = false;
  render();
  try {
    stage("stageLoading");
    const template = new Uint8Array(await (await fetch(`${import.meta.env.BASE_URL}tracker-template.xlsx`)).arrayBuffer());
    const existing = existingFile ? new Uint8Array(await existingFile.arrayBuffer()) : undefined;
    stage("stageFetching");
    const r = await buildTracker({
      siteUrl: site(), token, template, existing, includeUndated: undated.checked,
      onStage: (s) => {
        if (s.kind === "signedIn") stage("stageSignedIn", { user: s.user });
        else if (s.kind === "fetched") stage("stageFetched", { n: s.count });
        else stage("stageBuilding");
      },
    });

    if (downloadUrl) URL.revokeObjectURL(downloadUrl);
    downloadUrl = URL.createObjectURL(new Blob([r.bytes as Uint8Array<ArrayBuffer>], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }));
    const a = $<HTMLAnchorElement>("download");
    const fileName = downloadName(existingFile?.name, todayIn(TIMEZONE));
    a.href = downloadUrl;
    a.download = fileName;
    state.result = { r, existingName: existingFile?.name, fileName };
    progress.hidden = true;
    render();
    result.scrollIntoView({ behavior: "smooth", block: "start" });
    a.focus({ preventScroll: true });
  } catch (e) {
    progress.hidden = true;
    console.error(e);
    if (e instanceof UserError) fail({ key: `err_${e.code}` as Key, vars: { detail: e.detail ?? "" } });
    else fail({ key: "errGeneric", vars: { detail: (e as Error).message } });
  } finally {
    build.disabled = false;
    render();
  }
});

state.lang = initialLang(storage.get("lang"), navigator.languages ?? [navigator.language]);
render();
