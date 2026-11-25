/** Vefeyât pane — Şahıs dökümü and Tüm belgeyi denetle / Denetime devam et. */
import { buildXlsx, inventoryTable, toTsv } from "../submit.js";
import { NameIndex } from "../../core/index.js";
import { auditRows, exclusionKey } from "../logic.js";
import { type InventoryFilter, type InventoryRow, FILTER_LABELS, buildInventoryAsync, filterRows, inFilter, catalogLabel, hasDateIssue, statusLabel, summarize } from "../inventory.js";
import { parseBookmark, resumeIndex, rowKey } from "../bulk.js";
import { readSpellings, readExclusions, readAuditBookmark, writeAuditBookmark, readConfirmations, readDocument, documentVersion, selectMention, documentName } from "../word.js";

import { $, PAGE_SIZE, documentRows, loadSettings, noteCounts, placeOfRow, run, setMessage, showView, state, versions, progress, ensureIndex } from "./core.js";
import { refresh } from "./selected.js";
import { renderBulkOffer } from "./wikidata-bulk.js";

export async function prepareInventory(): Promise<void> {
  if (!state.catalog) {
    $("inventory-summary").textContent = "Katalog yüklenmedi; döküm hazırlanamaz.";
    return;
  }
  $("inventory-summary").textContent = "Döküm hazırlanıyor…";
  const read = await readDocument();
  await ensureIndex();
  state.segments = read.segments;
  state.rows = await buildInventoryAsync(read.segments, state.catalog, state.index ?? undefined, readConfirmations(), readExclusions(), readSpellings(), progress("Şahıslar taranıyor"));
  versions.rowsBuiltFor = `${documentVersion()}|${versions.identity}|${JSON.stringify([readConfirmations(), readExclusions(), readSpellings()])}`;
  state.page = 0;
  state.selectedRow = null;
  const sum = summarize(state.rows, state.settings.scope);
  $("inventory-summary").textContent =
    `${sum.mentions} geçiş · ${sum.people} kesin şahıs · ${sum.uncertain} doğrulanacak · ${sum.firstWithoutDate} ilk geçişte tarih yok · ${sum.spelling} yazım uyarısı` +
    (sum.dateIssues ? ` · ${sum.dateIssues} tarih uyarısı` : "") +
    (read.footnotes ? ` · okunan: ${noteCounts(read.segments)}` : " · Bu Word sürümünde dipnot ve sonnotlar okunamıyor");
  renderInventory();
  renderBulkOffer();
}

/** The rows the Şahıs dökümü süzgeci and the name search leave. */
export function visibleRows(): InventoryRow[] {
  return filterRows(state.rows, state.filter, state.query, state.settings.scope);
}

/** Süzgeç options with their counts ("Sorunlular (12)"). */
function renderFilter(): void {
  const select = $<HTMLSelectElement>("inventory-filter");
  select.replaceChildren();
  for (const [value, label] of Object.entries(FILTER_LABELS) as Array<[InventoryFilter, string]>) {
    const count = value === "all" ? state.rows.length : state.rows.filter((r) => inFilter(r, value, state.settings.scope)).length;
    if (count === 0 && value !== "all" && value !== state.filter) continue;
    const option = document.createElement("option");
    option.value = value;
    option.textContent = `${label} (${count})`;
    option.selected = value === state.filter;
    select.append(option);
  }
}

export function setInventoryFilter(filter: InventoryFilter, query: string): void {
  state.filter = filter;
  state.query = query;
  state.page = 0;
  renderInventory();
}

export function renderInventory(): void {
  const list = $("inventory-list");
  list.replaceChildren();
  renderFilter();
  const shown = visibleRows();
  if (shown.length === 0 && state.rows.length > 0) {
    const empty = document.createElement("li");
    empty.className = "muted";
    empty.textContent = "Bu süzgece uyan geçiş yok.";
    list.append(empty);
  }
  const pageRows = shown.slice(state.page * PAGE_SIZE, (state.page + 1) * PAGE_SIZE);
  for (const row of pageRows) {
    const item = document.createElement("li");
    item.setAttribute("role", "option");
    item.setAttribute("aria-selected", String(state.selectedRow === row.id));
    item.tabIndex = 0;
    const name = document.createElement("strong");
    name.textContent = row.text;
    const badge = document.createElement("span");
    badge.className = "badge";
    badge.textContent = placeOfRow(row);
    const meta = document.createElement("span");
    meta.className = "meta";
    meta.textContent = `${catalogLabel(row)} · ${row.date || "tarih yok"}`;
    const status = document.createElement("span");
    status.className = row.status === "match" && !hasDateIssue(row) ? "meta" : "meta warn";
    status.textContent = statusLabel(row, state.settings.scope);
    item.append(name, badge, meta, status);
    item.addEventListener("click", () => void run(() => chooseRow(row)));
    item.addEventListener("dblclick", () => void actOnRow());
    list.append(item);
  }
  const pages = Math.ceil(shown.length / PAGE_SIZE);
  $("inventory-pages").hidden = pages <= 1;
  $("inventory-page").textContent = pages > 1 ? `${state.page + 1} / ${pages}` : "";
  $<HTMLButtonElement>("inventory-act").disabled = state.selectedRow === null;
}

export async function chooseRow(row: InventoryRow): Promise<void> {
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

export async function actOnRow(): Promise<void> {
  if (state.selectedRow === null) return;
  showView("secili");
  state.alias = null;
  await refresh();
}

// ---- Tüm belgeyi denetle ----

export async function startAudit(): Promise<void> {
  await loadSettings();
  if (!state.catalog) {
    setMessage("Katalog yüklenmedi; denetim başlatılamaz.", "warning");
    return;
  }
  setMessage("Belge taranıyor…", "info");
  const read = await readDocument();
  await ensureIndex();
  state.segments = read.segments;
  state.rows = await buildInventoryAsync(read.segments, state.catalog, state.index ?? undefined, readConfirmations(), readExclusions(), readSpellings(), progress("Şahıslar taranıyor"));
  versions.rowsBuiltFor = `${documentVersion()}|${versions.identity}|${JSON.stringify([readConfirmations(), readExclusions(), readSpellings()])}`;
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

export async function resumeAudit(): Promise<void> {
  const bookmark = parseBookmark(readAuditBookmark());
  if (!bookmark) {
    setMessage("Duraklatılmış bir denetim yok; baştan başlatılıyor.", "info");
    await startAudit();
    return;
  }
  await loadSettings();
  if (!state.catalog) {
    setMessage("Katalog yüklenmedi; denetim sürdürülemez.", "warning");
    return;
  }
  const read = await readDocument();
  await ensureIndex();
  state.segments = read.segments;
  state.rows = await buildInventoryAsync(read.segments, state.catalog, state.index ?? undefined, readConfirmations(), readExclusions(), readSpellings(), progress("Şahıslar taranıyor"));
  versions.rowsBuiltFor = `${documentVersion()}|${versions.identity}|${JSON.stringify([readConfirmations(), readExclusions(), readSpellings()])}`;
  state.audit = auditRows(state.rows, state.settings);
  const at = resumeIndex(state.audit, state.segments, bookmark);
  if (at < 0) {
    state.audit = null;
    await writeAuditBookmark(null);
    setMessage("Denetlenecek geçiş kalmadı.", "success");
    renderAudit();
    return;
  }
  await gotoAudit(at);
}

export async function gotoAudit(at: number): Promise<void> {
  if (!state.audit) return;
  const row = state.audit[at];
  if (!row) return;
  state.auditAt = at;
  void writeAuditBookmark(JSON.stringify({ at, key: rowKey(row, state.segments), total: state.audit.length }));
  const segment = state.segments[row.segment];
  const ok = segment ? await selectMention(segment, row.start, row.text) : false;
  if (!ok) setMessage("Bu geçiş metinde bulunamadı; metin değişmiş olabilir. Sonraki ile devam edin.", "warning");
  renderAudit();
  state.alias = null;
  await refresh(true); // runs inside a busy task
}

/**
 * The audit's stops read again from the document as it is now: a date added,
 * a phrase marked Şahıs değil or a mention visited out of turn (Önceki geçişe
 * git) changes them. Returns the place of the current stop among them, or of
 * the stop just before where it was when it is gone (-1 before the first).
 */
async function freshAudit(): Promise<number> {
  if (!state.audit) return -1;
  const current = state.audit[state.auditAt];
  const key = current ? rowKey(current, state.segments) : "";
  // Where the stop is in the document: paragraph, note, offset (kept across a new reading).
  const where = (row: InventoryRow, segments: typeof state.segments): [number, number, number] => {
    const segment = segments[row.segment];
    return [segment?.anchor ?? 0, segment?.order ?? 0, row.start];
  };
  const at = current ? where(current, state.segments) : [0, 0, 0];
  const rows = await documentRows();
  const audit = auditRows(rows, state.settings);
  state.audit = audit;
  const same = audit.findIndex((row) => rowKey(row, state.segments) === key);
  if (same >= 0) return same;
  // Gone (dated under "Yalnız tarihsizler", or marked Şahıs değil): the last stop before its place.
  const earlier = (row: InventoryRow): boolean => {
    const [a1, o1, s1] = where(row, state.segments);
    return a1 < (at[0] as number) || (a1 === at[0] && (o1 < (at[1] as number) || (o1 === at[1] && s1 < (at[2] as number))));
  };
  let before = -1;
  audit.forEach((row, k) => {
    if (earlier(row)) before = k;
  });
  return before;
}

/** Sonraki / Önceki in the audit, over the stops as they are now; past the last one it finishes. */
export async function stepAudit(delta: 1 | -1): Promise<void> {
  if (!state.audit) return;
  const at = await freshAudit();
  const target = delta === 1 ? at + 1 : Math.max(at - 1, 0);
  if (state.audit && target < state.audit.length && state.audit.length > 0) {
    await gotoAudit(target);
    return;
  }
  state.audit = null;
  void writeAuditBookmark(null);
  renderAudit();
  setMessage("Denetim tamamlandı.", "success");
}

/** The audit goes on to the next stop (after a date was added, or a phrase was marked Şahıs değil). */
export async function advanceAudit(): Promise<void> {
  await stepAudit(1);
}

/** Şahıs değil during an audit: the phrase's later stops go too. */
export function dropFromAudit(key: string): void {
  if (!state.audit) return;
  const at = state.auditAt;
  state.audit = state.audit.filter((row, k) => k <= at || exclusionKey(row.text) !== key);
}

export function renderAudit(): void {
  const bar = $("audit-bar");
  bar.hidden = state.audit === null;
  if (!state.audit) return;
  const row = state.audit[state.auditAt];
  $("audit-position").textContent = `Denetim: ${state.auditAt + 1} / ${state.audit.length}` + (row ? ` · ${placeOfRow(row)} · ${statusLabel(row, state.settings.scope)}` : "");
  $<HTMLButtonElement>("audit-prev").disabled = state.auditAt === 0;
  $<HTMLButtonElement>("audit-next").textContent = state.auditAt + 1 >= state.audit.length ? "Bitir" : "Sonraki ▶";
}

// ---- Excel'e aktar ----

/** The döküm as a table: the rows the süzgeç and the search leave (all of them when nothing is filtered). */
function inventoryAsTable(): string[][] {
  return inventoryTable(visibleRows(), state.segments, (row) => statusLabel(row, state.settings.scope));
}

/** A download name browsers accept everywhere: Turkish letters folded, no odd characters. */
export function safeFileName(value: string): string {
  const folded = value
    .replace(/ı/g, "i").replace(/İ/g, "I").replace(/ş/g, "s").replace(/Ş/g, "S").replace(/ğ/g, "g").replace(/Ğ/g, "G")
    .replace(/ç/g, "c").replace(/Ç/g, "C").replace(/ö/g, "o").replace(/Ö/g, "O").replace(/ü/g, "u").replace(/Ü/g, "U")
    .normalize("NFD").replace(/\p{M}/gu, "");
  return folded.replace(/[^A-Za-z0-9._-]+/g, "-").replace(/-+/g, "-").replace(/^-|-$/g, "").slice(0, 80) || "belge";
}

/** Saves the inventory as an .xlsx file (the browser's download). */
export function exportInventory(): void {
  if (state.rows.length === 0) {
    setMessage("Önce Şahıs dökümünü hazırlayın.", "info");
    return;
  }
  const bytes = buildXlsx("Şahıs dökümü", inventoryAsTable());
  const url = URL.createObjectURL(new Blob([bytes], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }));
  const link = document.createElement("a");
  link.href = url;
  link.download = `${safeFileName(documentName().replace(/\.docx?$/i, ""))}-sahis-dokumu.xlsx`;
  document.body.append(link);
  link.click();
  link.remove();
  setTimeout(() => URL.revokeObjectURL(url), 2000);
  setMessage("Excel dosyası indirildi. İndirme açılmadıysa Panoya kopyala ile Excel'e yapıştırın.", "success");
}

/** Copies the inventory as tab-separated text: pastes into Excel or a Word table as cells. */
export async function copyInventory(): Promise<void> {
  if (state.rows.length === 0) {
    setMessage("Önce Şahıs dökümünü hazırlayın.", "info");
    return;
  }
  try {
    await navigator.clipboard.writeText(toTsv(inventoryAsTable()));
    setMessage(`${visibleRows().length} satır panoya kopyalandı. Excel'de bir hücreye ya da Word'de boş bir yere yapıştırın.`, "success");
  } catch {
    setMessage("Pano kullanılamadı. Excel'e aktar düğmesini deneyin.", "warning");
  }
}
