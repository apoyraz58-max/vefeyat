/**
 * Vefeyât pane. Opened by the ribbon like a VBA window: it shows the view the
 * button asked for. Settings live in the ribbon's Tercihler menus; the pane
 * only displays them. Writes happen only on an explicit click.
 */
import { Catalog, CatalogError, type Person, nameOptions, personKey, selectionAlias, textName, withEnding } from "../core/index.js";
import {
  DEFAULT_SETTINGS,
  type Launch,
  type Lookup,
  type Settings,
  type View,
  candidateLabel,
  dateTag,
  earlierSpellings,
  lookup,
  matchHint,
  orderByNameForm,
  parseLaunch,
  planDate,
  quickInsertBlocker,
  scopeRule,
  settingsSummary,
} from "./logic.js";
import { type StoredCatalog, clearCatalog, fetchEmbeddedCatalog, loadCatalog, saveCatalog } from "./store.js";
import { findEarlier, insertDate, onSelectionChanged, readSelection, readSettings, refreshSettings, replaceSelection } from "./word.js";

declare const __VERSION__: string;

interface State {
  catalog: Catalog | null;
  catalogSource: "user" | "embedded" | null;
  alias: string | null;
  found: Lookup | null;
  chosen: Person | null;
  settings: Settings;
  busy: boolean;
}

const state: State = {
  catalog: null,
  catalogSource: null,
  alias: null,
  found: null,
  chosen: null,
  settings: DEFAULT_SETTINGS,
  busy: false,
};

const $ = <T extends HTMLElement>(id: string): T => {
  const element = document.getElementById(id);
  if (!element) throw new Error(`Missing #${id}`);
  return element as T;
};

function setMessage(text: string, kind: "info" | "warning" | "error" | "success" = "info"): void {
  const box = $("message");
  box.textContent = text;
  box.className = `message ${kind}`;
  box.hidden = text.length === 0;
}

function showView(view: View): void {
  document.querySelectorAll<HTMLElement>(".view").forEach((element) => {
    const active = element.id === `view-${view}`;
    element.hidden = !active;
    if (active) $("view-title").textContent = element.dataset.title ?? "Vefeyât";
  });
  $("back").hidden = view === "secili";
}

async function loadSettings(): Promise<void> {
  await refreshSettings();
  state.settings = readSettings();
  $("settings-summary").textContent = settingsSummary(state.settings);
}

// ---- Catalog ----

function describeCatalog(fileName: string, count: number, source: "user" | "embedded", persisted: boolean): void {
  const origin = source === "embedded" ? "eklentiyle gelen katalog" : `sizin seçtiğiniz dosya: ${fileName}`;
  $("catalog-status").textContent = `${count.toLocaleString("tr-TR")} şahıs · ${origin}${persisted ? "" : " (yalnız bu oturum)"}`;
  $("reset-catalog").hidden = source !== "user";
  $("no-catalog").hidden = true;
}

function useCatalog(csv: string, fileName: string, source: "user" | "embedded", persisted: boolean): void {
  state.catalog = Catalog.fromCsv(csv);
  state.catalogSource = source;
  describeCatalog(fileName, state.catalog.people.length, source, persisted);
}

/** User's own CSV first; otherwise the catalog published with the add-in. */
async function restoreCatalog(): Promise<void> {
  const stored = await loadCatalog();
  if (stored?.source === "user") {
    try {
      useCatalog(stored.csv, stored.fileName, "user", true);
      return;
    } catch (error) {
      console.error(error);
    }
  }
  const embedded = await fetchEmbeddedCatalog();
  const candidates: Array<StoredCatalog | null> = [
    embedded === null ? null : { csv: embedded, fileName: "catalog.csv", savedAt: new Date().toISOString(), source: "embedded" },
    stored, // offline fallback: last copy
  ];
  for (const candidate of candidates) {
    if (!candidate) continue;
    try {
      useCatalog(candidate.csv, candidate.fileName, candidate.source ?? "embedded", true);
      if (candidate === candidates[0]) void saveCatalog(candidate);
      return;
    } catch (error) {
      console.error(error);
    }
  }
  state.catalog = null;
  $("catalog-status").textContent = "Katalog yüklenemedi. Bilgisayarınızdaki DİA CSV dosyasını seçin.";
  $("no-catalog").hidden = false;
}

async function chooseCatalog(file: File): Promise<void> {
  try {
    const csv = await file.text();
    Catalog.fromCsv(csv); // validate before replacing anything
    const persisted = await saveCatalog({ csv, fileName: file.name, savedAt: new Date().toISOString(), source: "user" });
    useCatalog(csv, file.name, "user", persisted);
    setMessage("Katalog yüklendi. CSV dosyanız bilgisayarınızda kalır.", "success");
    state.alias = null;
    showView("secili");
    await refresh();
  } catch (error) {
    const detail = error instanceof CatalogError ? error.message : String(error);
    setMessage(`CSV okunamadı: ${detail}`, "error");
  }
}

async function resetCatalog(): Promise<void> {
  await clearCatalog();
  await restoreCatalog();
  setMessage(state.catalogSource === "embedded" ? "Eklentinin kataloğuna dönüldü." : "Eklentiyle gelen katalog bulunamadı.", "info");
}

// ---- Selected name ----

function renderCandidates(): void {
  const list = $("candidates");
  list.replaceChildren();
  const found = state.found;
  if (!found) return;
  if (found.candidates.length === 0) {
    const empty = document.createElement("p");
    empty.className = "muted";
    empty.textContent = "Katalogda bu adla eşleşen şahıs yok.";
    list.append(empty);
    return;
  }
  found.candidates.forEach((person, index) => {
    const id = `candidate-${index}`;
    const row = document.createElement("label");
    row.className = "candidate";
    row.htmlFor = id;
    const radio = document.createElement("input");
    radio.type = "radio";
    radio.name = "candidate";
    radio.id = id;
    radio.checked = state.chosen !== null && personKey(state.chosen) === personKey(person);
    radio.addEventListener("change", () => {
      state.chosen = person;
      hideEarlier();
      renderDetails();
    });
    const text = document.createElement("span");
    const title = document.createElement("strong");
    title.textContent = candidateLabel(person);
    text.append(title);
    const note = matchHint(found.scores.get(personKey(person)) ?? 0);
    if (note) {
      const hint = document.createElement("span");
      hint.className = "hint";
      hint.textContent = ` · ${note}`;
      text.append(hint);
    }
    if (person.description) {
      const description = document.createElement("small");
      description.textContent = person.description.length > 140 ? `${person.description.slice(0, 140)}…` : person.description;
      text.append(document.createElement("br"), description);
    }
    row.append(radio, text);
    list.append(row);
  });
}

function renderDetails(): void {
  const person = state.chosen;
  $("details").hidden = person === null;
  $<HTMLButtonElement>("insert").disabled = person === null || state.busy;
  const spelling = $<HTMLSelectElement>("spelling");
  spelling.replaceChildren();
  $<HTMLButtonElement>("respell").disabled = true;
  if (!person || !state.alias) return;
  $("detail-full").textContent = textName(person.fullName) || "—";
  $("detail-original").textContent = person.originalName || "—";
  $("detail-date").textContent = person.deathDate || "Katalogda tarih yok";
  const link = $<HTMLAnchorElement>("detail-link");
  link.hidden = !person.source.startsWith("https://islamansiklopedisi.org.tr/");
  link.href = link.hidden ? "#" : person.source;
  for (const option of orderByNameForm(nameOptions(person, state.alias), state.settings.nameForm)) {
    if (option.value === state.alias) continue;
    const element = document.createElement("option");
    element.value = option.value;
    element.textContent = `${option.value} (${option.label})`;
    spelling.append(element);
  }
  $<HTMLButtonElement>("respell").disabled = spelling.options.length === 0 || state.busy;
}

function hideEarlier(): void {
  $("earlier").hidden = true;
}

async function refresh(): Promise<void> {
  if (state.busy) return;
  hideEarlier();
  try {
    await loadSettings();
    const selection = await readSelection();
    const alias = selectionAlias(selection.text);
    if (alias === state.alias && state.found) {
      renderDetails(); // settings such as Ad biçimi may have changed
      return;
    }
    if (alias !== state.alias) setMessage("");
    state.alias = alias;
    state.found = null;
    state.chosen = null;
    $("selected-name").textContent = alias ?? "Metinde bir şahıs adı seçin.";
    if (alias && state.catalog) {
      state.found = lookup(state.catalog, alias);
      state.chosen = state.found.automatic;
    }
    renderCandidates();
    renderDetails();
  } catch (error) {
    console.error(error);
  }
}

async function run(task: () => Promise<void>): Promise<void> {
  if (state.busy) return;
  state.busy = true;
  renderDetails();
  try {
    await task();
  } catch (error) {
    setMessage(`Word işlemi tamamlanamadı: ${String(error)} [Vefeyât ${__VERSION__}]`, "error");
  } finally {
    state.busy = false;
    renderDetails();
  }
}

/** The selection must still be the name the pane shows. */
async function currentSelection(): Promise<{ alias: string; following: string } | null> {
  const selection = await readSelection();
  const alias = selectionAlias(selection.text);
  if (!alias || alias !== state.alias) {
    setMessage("Seçim değişti. Adı yeniden seçin.", "warning");
    await refresh();
    return null;
  }
  return { alias, following: selection.following };
}

async function addDate(force: boolean): Promise<void> {
  const person = state.chosen;
  if (!person) return;
  await loadSettings();
  const rule = scopeRule(state.settings.scope);
  if (!rule.allowed) {
    setMessage(rule.reason ?? "", "warning");
    return;
  }
  const current = await currentSelection();
  if (!current) return;
  const plan = planDate(person, state.settings.prefix, current.following);
  if (!plan.ok) {
    setMessage(plan.reason, "warning");
    return;
  }
  if (!force && rule.checkEarlier && (await findEarlier(earlierSpellings(person, current.alias)))) {
    $("earlier").hidden = false;
    setMessage("", "info");
    return;
  }
  await insertDate(plan.text, dateTag(person), `Vefeyât: ${textName(person.shortName)}`);
  hideEarlier();
  setMessage(`Tarih eklendi: ${plan.text}`, "success");
  state.alias = null; // the caret moved; read the next selection afresh
}

async function respell(): Promise<void> {
  const spelling = $<HTMLSelectElement>("spelling").value;
  if (!spelling) return;
  const selection = await readSelection();
  if (selection.text !== state.alias) {
    setMessage("Yalnız adı seçin (boşluk veya noktalama olmadan); yazım yalnız seçili metinde değişir.", "warning");
    return;
  }
  const replacement = withEnding(selection.text, spelling);
  await replaceSelection(replacement);
  setMessage(`Yazım değiştirildi: ${replacement}`, "success");
  state.alias = null;
  await refresh();
}

/** Carries out what the ribbon or context-menu button asked for. */
async function launch(request: Launch): Promise<void> {
  showView(request.view);
  if (request.tool) document.getElementById(`tool-${request.tool}`)?.classList.add("focus");
  if (request.action === "insert") {
    const blocker = quickInsertBlocker(state.catalog !== null, state.alias, state.found);
    if (blocker) setMessage(blocker, "info");
    else await run(() => addDate(false));
  } else if (request.action === "spelling") {
    if (state.chosen && !$<HTMLButtonElement>("respell").disabled) $<HTMLSelectElement>("spelling").focus();
    else setMessage(quickInsertBlocker(state.catalog !== null, state.alias, state.found) ?? "Bu ad için farklı bir yazım önerisi yok.", "info");
  }
}

function initialise(): void {
  $("version").textContent = __VERSION__;
  const file = $<HTMLInputElement>("catalog-file");
  $("choose-catalog").addEventListener("click", () => file.click());
  $("choose-catalog-inline").addEventListener("click", () => file.click());
  $("reset-catalog").addEventListener("click", () => void resetCatalog());
  file.addEventListener("change", () => {
    const chosen = file.files?.[0];
    if (chosen) void chooseCatalog(chosen);
    file.value = "";
  });
  $("back").addEventListener("click", () => {
    showView("secili");
    state.alias = null;
    void refresh();
  });
  $("insert").addEventListener("click", () => void run(() => addDate(false)));
  $("insert-anyway").addEventListener("click", () => void run(() => addDate(true)));
  $("go-earlier").addEventListener("click", () =>
    void run(async () => {
      const person = state.chosen;
      if (person && state.alias) await findEarlier(earlierSpellings(person, state.alias), true);
      hideEarlier();
    }),
  );
  $("respell").addEventListener("click", () => void run(respell));
  let timer: number | undefined;
  onSelectionChanged(() => {
    window.clearTimeout(timer);
    timer = window.setTimeout(() => void refresh(), 250);
  });
  const request = parseLaunch(window.location.search);
  showView(request.view);
  void restoreCatalog()
    .then(refresh)
    .then(() => launch(request));
}

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) initialise();
  else setMessage("Vefeyât yalnız Word içinde çalışır.", "error");
});
