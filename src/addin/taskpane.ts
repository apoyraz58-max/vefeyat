/**
 * Vefeyât task pane. Reads the selected name, shows catalog candidates and
 * writes a date or a spelling only on an explicit click.
 */
import { Catalog, CatalogError, type Person, type PrefixKey, nameOptions, personKey, selectionAlias, textName, withEnding } from "../core/index.js";
import { type Launch, type Lookup, type View, candidateLabel, dateTag, earlierSpellings, lookup, matchHint, parseLaunch, planDate, quickInsertBlocker } from "./logic.js";
import { loadCatalog, saveCatalog } from "./store.js";
import { findEarlier, insertDate, onSelectionChanged, readPrefix, readSelection, replaceSelection, writePrefix } from "./word.js";

declare const __VERSION__: string;

interface State {
  catalog: Catalog | null;
  alias: string | null;
  found: Lookup | null;
  chosen: Person | null;
  prefix: PrefixKey;
  busy: boolean;
}

const state: State = { catalog: null, alias: null, found: null, chosen: null, prefix: "o", busy: false };

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
    element.hidden = element.id !== `view-${view}`;
  });
  document.querySelectorAll<HTMLButtonElement>(".tabs button").forEach((tab) => {
    tab.setAttribute("aria-selected", String(tab.dataset.view === view));
    tab.tabIndex = tab.dataset.view === view ? 0 : -1;
  });
}

function showCatalogNotice(): void {
  $("no-catalog").hidden = state.catalog !== null;
}

function hideEarlier(): void {
  $("earlier").hidden = true;
}

function describeCatalog(fileName: string, count: number, persisted: boolean): void {
  $("catalog-status").textContent = `${count.toLocaleString("tr-TR")} şahıs · ${fileName}${persisted ? "" : " (yalnız bu oturum)"}`;
}

async function restoreCatalog(): Promise<void> {
  const stored = await loadCatalog();
  if (!stored) {
    $("catalog-status").textContent = "Katalog seçilmedi. Bilgisayarınızdaki DİA CSV dosyasını seçin.";
    return;
  }
  try {
    state.catalog = Catalog.fromCsv(stored.csv);
    describeCatalog(stored.fileName, state.catalog.people.length, true);
    showCatalogNotice();
  } catch (error) {
    $("catalog-status").textContent = "Kayıtlı katalog okunamadı; CSV dosyasını yeniden seçin.";
    console.error(error);
  }
}

async function chooseCatalog(file: File): Promise<void> {
  try {
    const csv = await file.text();
    const catalog = Catalog.fromCsv(csv);
    state.catalog = catalog;
    const persisted = await saveCatalog({ csv, fileName: file.name, savedAt: new Date().toISOString() });
    describeCatalog(file.name, catalog.people.length, persisted);
    showCatalogNotice();
    showView("secili");
    setMessage("Katalog yüklendi. CSV dosyanız bilgisayarınızda kalır.", "success");
    await refresh();
  } catch (error) {
    const detail = error instanceof CatalogError ? error.message : String(error);
    setMessage(`CSV okunamadı: ${detail}`, "error");
  }
}

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
  for (const option of nameOptions(person, state.alias)) {
    if (option.value === state.alias) continue;
    const element = document.createElement("option");
    element.value = option.value;
    element.textContent = `${option.value} (${option.label})`;
    spelling.append(element);
  }
  $<HTMLButtonElement>("respell").disabled = spelling.options.length === 0 || state.busy;
}

async function refresh(): Promise<void> {
  if (state.busy) return;
  hideEarlier();
  try {
    const selection = await readSelection();
    const alias = selectionAlias(selection.text);
    if (alias === state.alias && state.found) return;
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

/** The selection must still be the name the panel shows. */
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
  const current = await currentSelection();
  if (!current) return;
  const plan = planDate(person, state.prefix, current.following);
  if (!plan.ok) {
    setMessage(plan.reason, "warning");
    return;
  }
  if (!force && (await findEarlier(earlierSpellings(person, current.alias)))) {
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
  state.prefix = readPrefix() ?? "o";
  const prefix = $<HTMLSelectElement>("prefix");
  prefix.value = state.prefix;
  prefix.addEventListener("change", () => {
    state.prefix = prefix.value as PrefixKey;
    void writePrefix(state.prefix).catch((error) => setMessage(`Ayar kaydedilemedi: ${String(error)}`, "error"));
  });
  const file = $<HTMLInputElement>("catalog-file");
  $("choose-catalog").addEventListener("click", () => file.click());
  $("choose-catalog-inline").addEventListener("click", () => file.click());
  document.querySelectorAll<HTMLButtonElement>(".tabs button").forEach((tab) => {
    tab.addEventListener("click", () => showView(tab.dataset.view as View));
  });
  file.addEventListener("change", () => {
    const chosen = file.files?.[0];
    if (chosen) void chooseCatalog(chosen);
    file.value = "";
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
  $("refresh").addEventListener("click", () => {
    state.alias = null;
    void refresh();
  });
  let timer: number | undefined;
  onSelectionChanged(() => {
    window.clearTimeout(timer);
    timer = window.setTimeout(() => void refresh(), 250);
  });
  const request = parseLaunch(window.location.search);
  showView(request.view);
  void restoreCatalog()
    .then(() => {
      showCatalogNotice();
      return refresh();
    })
    .then(() => launch(request));
}

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) initialise();
  else setMessage("Vefeyât yalnız Word içinde çalışır.", "error");
});
