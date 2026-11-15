/**
 * Vefeyât pane. Opened by the ribbon like a VBA window: it shows the view the
 * button asked for. Settings live in the ribbon's Tercihler menus; the pane
 * only displays them. Writes happen only on an explicit click.
 */
import { Catalog, CatalogError, NameIndex, type Person, nameOptions, personKey, selectionAlias, textName, withEnding } from "../core/index.js";
import {
  DEFAULT_SETTINGS,
  type Launch,
  type Lookup,
  type Settings,
  type View,
  SOON,
  auditRows,
  addConfirmation,
  confirmedPerson,
  mentionKey,
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
import { type InventoryRow, type Segment, buildInventory, catalogLabel, statusLabel, summarize } from "./inventory.js";
import { type StoredCatalog, clearCatalog, fetchEmbeddedCatalog, loadCatalog, saveCatalog } from "./store.js";
import { consistencyReport, personReportDocx, toBase64 } from "./submit.js";
import { documentName, finalizeDocument, findEarlier, insertDate, openNewDocument, onSelectionChanged, readConfirmations, readSegments, readSelection, readSettings, refreshSettings, replaceSelection, selectMention, writeConfirmations } from "./word.js";

declare const __VERSION__: string;

interface State {
  catalog: Catalog | null;
  catalogSource: "user" | "embedded" | null;
  alias: string | null;
  found: Lookup | null;
  chosen: Person | null;
  settings: Settings;
  busy: boolean;
  index: NameIndex | null;
  segments: Segment[];
  rows: InventoryRow[];
  page: number;
  selectedRow: number | null;
  audit: InventoryRow[] | null;
  auditAt: number;
  mention: string | null;
  confirmed: boolean;
}

const state: State = {
  catalog: null,
  catalogSource: null,
  alias: null,
  found: null,
  chosen: null,
  settings: DEFAULT_SETTINGS,
  busy: false,
  index: null,
  segments: [],
  rows: [],
  page: 0,
  selectedRow: null,
  audit: null,
  auditAt: 0,
  mention: null,
  confirmed: false,
};

const PAGE_SIZE = 250;

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

function showSoon(key: string): void {
  const entry = SOON[key];
  if (!entry) return;
  $("view-yakinda").dataset.title = entry[0];
  $("soon-description").textContent = entry[1];
}

function runSearch(): void {
  const query = $<HTMLInputElement>("search-query").value;
  const list = $("search-results");
  list.replaceChildren();
  if (!state.catalog) {
    $("search-status").textContent = "Katalog yüklenmedi.";
    return;
  }
  const found = state.catalog.find(query, true);
  $("search-status").textContent = query.trim() ? `${found.length} sonuç${found.length > 50 ? " (ilk 50 gösteriliyor)" : ""}` : "";
  for (const person of found.slice(0, 50)) {
    const item = document.createElement("li");
    const title = document.createElement("strong");
    title.textContent = candidateLabel(person);
    item.append(title);
    if (person.source.startsWith("https://islamansiklopedisi.org.tr/")) {
      const link = document.createElement("a");
      link.href = person.source;
      link.target = "_blank";
      link.rel = "noopener noreferrer";
      link.textContent = "DİA";
      item.append(link);
    }
    if (person.description) {
      const description = document.createElement("small");
      description.textContent = person.description.length > 160 ? `${person.description.slice(0, 160)}…` : person.description;
      item.append(document.createElement("br"), description);
    }
    list.append(item);
  }
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
  state.index = null; // rebuilt on the next inventory
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
    const confirmedHere = state.confirmed && state.chosen !== null && personKey(state.chosen) === personKey(person);
    const note = confirmedHere ? "bu geçiş için onayladınız" : matchHint(found.scores.get(personKey(person)) ?? 0);
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

async function refresh(force = false): Promise<void> {
  if (state.busy && !force) return;
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
    state.confirmed = false;
    state.mention = alias ? mentionKey(alias, selection.preceding, selection.following) : null;
    $("selected-name").textContent = alias ?? "Metinde bir şahıs adı seçin.";
    if (alias && state.catalog) {
      state.found = lookup(state.catalog, alias);
      state.chosen = state.found.automatic;
      if (!state.chosen && state.mention) {
        state.chosen = confirmedPerson(readConfirmations(), state.mention, state.found.candidates);
        state.confirmed = state.chosen !== null;
      }
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
  await rememberChoice();
  await insertDate(plan.text, dateTag(person), `Vefeyât: ${textName(person.shortName)}`);
  hideEarlier();
  setMessage(`Tarih eklendi: ${plan.text}`, "success");
  state.alias = null; // the caret moved; read the next selection afresh
}

/** Remembers the user's choice for an ambiguous mention in the document. */
async function rememberChoice(): Promise<void> {
  if (!state.chosen || !state.mention || !state.found || state.found.automatic) return;
  await writeConfirmations(addConfirmation(readConfirmations(), state.mention, personKey(state.chosen)));
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
  if (state.chosen && state.found && !state.found.automatic && state.mention) {
    // The sentence keeps its words; only the name changes, so key the new spelling.
    const [before, , after] = state.mention.split("|");
    const key = `${before ?? ""}|${replacement}|${after ?? ""}`;
    await writeConfirmations(addConfirmation(readConfirmations(), key, personKey(state.chosen)));
  }
  await replaceSelection(replacement);
  setMessage(`Yazım değiştirildi: ${replacement}`, "success");
  state.alias = null;
  await refresh();
}

// ---- Şahıs dökümü ----

const LOCATION = { main: "Metin", note: "Dipnot" } as const;

async function prepareInventory(): Promise<void> {
  if (!state.catalog) {
    $("inventory-summary").textContent = "Katalog yüklenmedi; döküm hazırlanamaz.";
    return;
  }
  $("inventory-summary").textContent = "Döküm hazırlanıyor…";
  const read = await readSegments();
  state.index ??= new NameIndex(state.catalog.people);
  state.segments = read.segments;
  state.rows = buildInventory(read.segments, state.catalog, state.index, readConfirmations());
  state.page = 0;
  state.selectedRow = null;
  const sum = summarize(state.rows);
  $("inventory-summary").textContent =
    `${sum.mentions} geçiş · ${sum.people} kesin şahıs · ${sum.uncertain} doğrulanacak · ${sum.firstWithoutDate} ilk geçişte tarih yok · ${sum.spelling} yazım uyarısı` +
    (read.footnotes ? "" : " · Bu Word sürümünde dipnotlar okunamıyor");
  renderInventory();
}

function renderInventory(): void {
  const list = $("inventory-list");
  list.replaceChildren();
  const pageRows = state.rows.slice(state.page * PAGE_SIZE, (state.page + 1) * PAGE_SIZE);
  for (const row of pageRows) {
    const item = document.createElement("li");
    item.setAttribute("role", "option");
    item.setAttribute("aria-selected", String(state.selectedRow === row.id));
    item.tabIndex = 0;
    const name = document.createElement("strong");
    name.textContent = row.text;
    const badge = document.createElement("span");
    badge.className = "badge";
    badge.textContent = LOCATION[row.part];
    const meta = document.createElement("span");
    meta.className = "meta";
    meta.textContent = `${catalogLabel(row)} · ${row.date || "tarih yok"}`;
    const status = document.createElement("span");
    status.className = row.status === "match" ? "meta" : "meta warn";
    status.textContent = statusLabel(row);
    item.append(name, badge, meta, status);
    item.addEventListener("click", () => void run(() => chooseRow(row)));
    item.addEventListener("dblclick", () => void actOnRow());
    list.append(item);
  }
  const pages = Math.ceil(state.rows.length / PAGE_SIZE);
  $("inventory-pages").hidden = pages <= 1;
  $("inventory-page").textContent = pages > 1 ? `${state.page + 1} / ${pages}` : "";
  $<HTMLButtonElement>("inventory-act").disabled = state.selectedRow === null;
}

async function chooseRow(row: InventoryRow): Promise<void> {
  const segment = state.segments[row.segment];
  if (!segment) return;
  const ok = await selectMention(segment, row.start, row.text);
  if (!ok) {
    setMessage("Bu geçiş metinde bulunamadı; metin değişmiş olabilir. Dökümü yenileyin.", "warning");
    return;
  }
  state.selectedRow = row.id;
  renderInventory();
}

async function actOnRow(): Promise<void> {
  if (state.selectedRow === null) return;
  showView("secili");
  state.alias = null;
  await refresh();
}

// ---- Tüm belgeyi denetle ----

async function startAudit(): Promise<void> {
  await loadSettings();
  if (!state.catalog) {
    setMessage("Katalog yüklenmedi; denetim başlatılamaz.", "warning");
    return;
  }
  setMessage("Belge taranıyor…", "info");
  const read = await readSegments();
  state.index ??= new NameIndex(state.catalog.people);
  state.segments = read.segments;
  state.rows = buildInventory(read.segments, state.catalog, state.index, readConfirmations());
  state.audit = auditRows(state.rows, state.settings);
  state.auditAt = 0;
  setMessage("", "info");
  if (state.audit.length === 0) {
    state.audit = null;
    setMessage("Denetlenecek geçiş yok.", "success");
    renderAudit();
    return;
  }
  await gotoAudit(0);
}

async function gotoAudit(at: number): Promise<void> {
  if (!state.audit) return;
  const row = state.audit[at];
  if (!row) return;
  state.auditAt = at;
  const segment = state.segments[row.segment];
  const ok = segment ? await selectMention(segment, row.start, row.text) : false;
  if (!ok) setMessage("Bu geçiş metinde bulunamadı; metin değişmiş olabilir. Sonraki ile devam edin.", "warning");
  renderAudit();
  state.alias = null;
  await refresh(true); // runs inside a busy task
}

function renderAudit(): void {
  const bar = $("audit-bar");
  bar.hidden = state.audit === null;
  if (!state.audit) return;
  const row = state.audit[state.auditAt];
  $("audit-position").textContent = `Denetim: ${state.auditAt + 1} / ${state.audit.length}` + (row ? ` · ${LOCATION[row.part]} · ${statusLabel(row)}` : "");
  $<HTMLButtonElement>("audit-prev").disabled = state.auditAt === 0;
  $<HTMLButtonElement>("audit-next").textContent = state.auditAt + 1 >= state.audit.length ? "Bitir" : "Sonraki ▶";
}


// ---- Teslim ve dizin ----

async function showConsistency(): Promise<void> {
  const { segments, footnotes } = await readSegments();
  const report = consistencyReport(segments);
  const out = $("consistency-out");
  out.hidden = false;
  out.replaceChildren();
  const line = (text: string, cls = "") => {
    const p = document.createElement("p");
    p.textContent = text;
    if (cls) p.className = cls;
    out.append(p);
  };
  if (report.total === 0) {
    line("Belgede (ö. …) biçiminde tarih bulunamadı.");
  } else {
    line(`${report.total} tarih incelendi${footnotes ? " (dipnotlar dahil)" : " (dipnotlar bu Word sürümünde okunamadı)"}.`);
    for (const [title, entries, mixed] of [["İşaret", report.markers, report.mixedMarkers], ["Biçim", report.kinds, report.mixedKinds]] as const) {
      line(`${title}: ${entries.map((e) => `${e.label} ${e.count}`).join(" · ")}${mixed ? " — karışık" : " — tutarlı"}`, mixed ? "warning" : "");
      for (const e of entries) for (const x of e.examples) line(`  ${x}`, "muted");
    }
  }
  setMessage(report.mixedMarkers || report.mixedKinds ? "Tarih yazımında karışıklık var; örnekler aşağıda." : "Tutarlılık raporu hazır.", report.mixedMarkers || report.mixedKinds ? "warning" : "success");
}

async function createPersonReport(): Promise<void> {
  if (!state.catalog) {
    setMessage("Katalog yüklenmedi.", "warning");
    return;
  }
  const read = await readSegments();
  state.index ??= new NameIndex(state.catalog.people);
  const rows = buildInventory(read.segments, state.catalog, state.index, readConfirmations());
  await openNewDocument(toBase64(personReportDocx(documentName(), rows)));
  setMessage("Şahıs raporu yeni bir Word belgesi olarak açıldı.", "success");
}

let cleanupArmed = false;
async function cleanup(): Promise<void> {
  const button = $<HTMLButtonElement>("run-temizlik");
  if (!cleanupArmed) {
    cleanupArmed = true;
    button.textContent = "Eminim, temizle";
    setMessage("Bu işlem geri alınamaz; teslim kopyasında çalıştığınızdan emin olun ve tekrar tıklayın.", "warning");
    return;
  }
  cleanupArmed = false;
  button.textContent = "Temizle";
  const result = await finalizeDocument();
  setMessage(`Temizlendi: ${result.controls} tarih denetimi metne dönüştürüldü, ${result.settings} belge ayarı kaldırıldı.`, "success");
}

/** Carries out what the ribbon or context-menu button asked for. */
async function launch(request: Launch): Promise<void> {
  showView(request.view);
  if (request.tool) document.getElementById(`tool-${request.tool}`)?.classList.add("focus");
  if (request.view === "ara") $<HTMLInputElement>("search-query").focus();
  if (request.view === "dokum") await run(prepareInventory);
  if (request.tool === "tutarlilik") await run(showConsistency);
  if (request.tool === "rapor") await run(createPersonReport);
  if (request.action === "audit") await run(startAudit);
  if (request.action === "dia") {
    const person = state.chosen;
    if (person && person.source.startsWith("https://islamansiklopedisi.org.tr/")) {
      setMessage("DİA maddesini açmak için aşağıdaki bağlantıya tıklayın.", "info");
      $("detail-link").focus();
    } else {
      setMessage(quickInsertBlocker(state.catalog !== null, state.alias, state.found) ?? "Bu şahıs için DİA bağlantısı yok.", "info");
    }
  }
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
  $("run-tutarlilik").addEventListener("click", () => void run(showConsistency));
  $("run-rapor").addEventListener("click", () => void run(createPersonReport));
  $("run-temizlik").addEventListener("click", () => void run(cleanup));
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
  $("inventory-refresh").addEventListener("click", () => void run(prepareInventory));
  $("inventory-act").addEventListener("click", () => void actOnRow());
  $("inventory-prev").addEventListener("click", () => {
    if (state.page > 0) state.page--;
    renderInventory();
  });
  $("inventory-next").addEventListener("click", () => {
    if ((state.page + 1) * PAGE_SIZE < state.rows.length) state.page++;
    renderInventory();
  });
  $("audit-prev").addEventListener("click", () => void run(() => gotoAudit(state.auditAt - 1)));
  $("audit-next").addEventListener("click", () =>
    void run(async () => {
      if (state.audit && state.auditAt + 1 < state.audit.length) await gotoAudit(state.auditAt + 1);
      else {
        state.audit = null;
        renderAudit();
        setMessage("Denetim tamamlandı.", "success");
      }
    }),
  );
  $("audit-stop").addEventListener("click", () => {
    state.audit = null;
    renderAudit();
  });
  $("search-form").addEventListener("submit", (event) => {
    event.preventDefault();
    runSearch();
  });
  const request = parseLaunch(window.location.search);
  if (request.soon) showSoon(request.soon);
  showView(request.view);
  void restoreCatalog()
    .then(() => refresh())
    .then(() => launch(request));
}

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) initialise();
  else setMessage("Vefeyât yalnız Word içinde çalışır.", "error");
});
