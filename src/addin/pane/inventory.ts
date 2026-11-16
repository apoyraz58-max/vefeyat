/** Vefeyât pane — Şahıs dökümü and Tüm belgeyi denetle / Denetime devam et. */
import { NameIndex } from "../../core/index.js";
import { auditRows } from "../logic.js";
import { type InventoryRow, buildInventory, catalogLabel, hasDateIssue, statusLabel, summarize } from "../inventory.js";
import { parseBookmark, resumeIndex, rowKey } from "../bulk.js";
import { readSpellings, readExclusions, readAuditBookmark, writeAuditBookmark, readConfirmations, readDocument, documentVersion, selectMention } from "../word.js";

import { $, PAGE_SIZE, loadSettings, noteCounts, placeOfRow, run, setMessage, showView, state, versions } from "./core.js";
import { refresh } from "./selected.js";

export async function prepareInventory(): Promise<void> {
  if (!state.catalog) {
    $("inventory-summary").textContent = "Katalog yüklenmedi; döküm hazırlanamaz.";
    return;
  }
  $("inventory-summary").textContent = "Döküm hazırlanıyor…";
  const read = await readDocument();
  state.index ??= new NameIndex(state.catalog.people);
  state.segments = read.segments;
  state.rows = buildInventory(read.segments, state.catalog, state.index, readConfirmations(), readExclusions(), readSpellings());
  versions.rowsBuiltFor = `${documentVersion()}|${versions.identity}|${JSON.stringify([readConfirmations(), readExclusions(), readSpellings()])}`;
  state.page = 0;
  state.selectedRow = null;
  const sum = summarize(state.rows);
  $("inventory-summary").textContent =
    `${sum.mentions} geçiş · ${sum.people} kesin şahıs · ${sum.uncertain} doğrulanacak · ${sum.firstWithoutDate} ilk geçişte tarih yok · ${sum.spelling} yazım uyarısı` +
    (sum.dateIssues ? ` · ${sum.dateIssues} tarih uyarısı` : "") +
    (read.footnotes ? ` · okunan: ${noteCounts(read.segments)}` : " · Bu Word sürümünde dipnot ve sonnotlar okunamıyor");
  renderInventory();
}

export function renderInventory(): void {
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
    badge.textContent = placeOfRow(row);
    const meta = document.createElement("span");
    meta.className = "meta";
    meta.textContent = `${catalogLabel(row)} · ${row.date || "tarih yok"}`;
    const status = document.createElement("span");
    status.className = row.status === "match" && !hasDateIssue(row) ? "meta" : "meta warn";
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
  state.index ??= new NameIndex(state.catalog.people);
  state.segments = read.segments;
  state.rows = buildInventory(read.segments, state.catalog, state.index, readConfirmations(), readExclusions(), readSpellings());
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
  state.index ??= new NameIndex(state.catalog.people);
  state.segments = read.segments;
  state.rows = buildInventory(read.segments, state.catalog, state.index, readConfirmations(), readExclusions(), readSpellings());
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

export function renderAudit(): void {
  const bar = $("audit-bar");
  bar.hidden = state.audit === null;
  if (!state.audit) return;
  const row = state.audit[state.auditAt];
  $("audit-position").textContent = `Denetim: ${state.auditAt + 1} / ${state.audit.length}` + (row ? ` · ${placeOfRow(row)} · ${statusLabel(row)}` : "");
  $<HTMLButtonElement>("audit-prev").disabled = state.auditAt === 0;
  $<HTMLButtonElement>("audit-next").textContent = state.auditAt + 1 >= state.audit.length ? "Bitir" : "Sonraki ▶";
}
