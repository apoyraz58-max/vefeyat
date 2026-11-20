/** Vefeyât pane — İşlem geçmişi list, Toplu düzeltme, Belge istisnaları, Kimlik eşleşmeleri, Denetim kapsamı, Tercihler. */
import { NameIndex, personKey, textName } from "../../core/index.js";
import { type SettingName, listConfirmations, dateTag } from "../logic.js";
import { type InventoryRow, buildInventoryAsync } from "../inventory.js";
import { HISTORY_LABEL, type Suggestion, applyOrder, bulkSuggestions, describeSuggestion, firstUseSuggestions } from "../bulk.js";
import { exportPrefs, mergePrefs, parsePrefs } from "../prefs.js";
import { writeSetting, readSpellings, writeSpellings, toggleDateFrames, readCoverage, readExclusions, writeExclusions, selectDateControl, replaceDateAfterSelection, canSkipProofing, skipProofing, insertDate, readConfirmations, readDocument, documentVersion, replaceSelection, selectMention, writeConfirmations, readLocalPersons, writeLocalPersons, removeDateAfter, hasPendingChanges, documentRead, watchDocument } from "../word.js";

import { applyLocalPersons } from "./catalog.js";
import { $, documentRows, history, loadSettings, logHistory, markSelected, placeOfRow, progress, run, setMessage, state, versions, ensureIndex } from "./core.js";

export function renderHistory(): void {
  $("history-summary").textContent = history.length
    ? `Bu oturumda ${history.length} işlem. Tarih satırına tıklayınca tarih belgede seçilir.`
    : "Bu oturumda henüz Vefeyât işlemi yapılmadı.";
  const list = $("history-list");
  list.replaceChildren();
  for (const entry of history) {
    const item = document.createElement("li");
    const name = document.createElement("strong");
    name.textContent = entry.text;
    const meta = document.createElement("span");
    meta.className = "meta";
    meta.textContent = `${entry.time} · ${HISTORY_LABEL[entry.kind]}`;
    item.append(name, meta);
    if (entry.control !== undefined) {
      const id = entry.control;
      item.addEventListener("click", () => {
        markSelected(item);
        void run(async () => {
          if (!(await selectDateControl(id))) setMessage("Bu tarih artık belgede yok.", "warning");
        });
      });
    }
    list.append(item);
  }
}

// ---- Toplu düzeltme önizlemesi ----

export let suggestions: Suggestion[] = [];

export const bulkChosen = new Set<string>();

export async function prepareBulk(): Promise<void> {
  await loadSettings();
  if (!state.catalog) {
    $("bulk-summary").textContent = "Katalog yüklenmedi.";
    return;
  }
  $("bulk-summary").textContent = "Öneriler hazırlanıyor…";
  const read = await readDocument();
  await ensureIndex();
  state.segments = read.segments;
  state.rows = await buildInventoryAsync(read.segments, state.catalog, state.index ?? undefined, readConfirmations(), readExclusions(), readSpellings(), progress("Şahıslar taranıyor"));
  versions.rowsBuiltFor = `${documentVersion()}|${versions.identity}|${JSON.stringify([readConfirmations(), readExclusions(), readSpellings()])}`;
  suggestions = bulkSuggestions(state.rows, state.settings, readSpellings());
  bulkChosen.clear();
  for (const s of suggestions) if (s.preselected) bulkChosen.add(s.id);
  renderBulk();
}

export function renderBulk(): void {
  const dates = suggestions.filter((s) => s.kind === "date").length;
  const fixes = suggestions.filter((s) => s.kind === "fixdate").length;
  const upkeep = suggestions.filter((s) => s.kind === "move" || s.kind === "drop").length;
  const uncertain = state.rows.filter((r) => r.status === "near" || r.status === "choose").length;
  $("bulk-summary").textContent = suggestions.length
    ? (upkeep ? `İlk geçiş değişikliği: ${upkeep} öneri (taşı / fazlayı kaldır). ` : "") +
      `${dates} tarih, ${suggestions.length - dates - fixes - upkeep} yazım önerisi` +
      (fixes ? `, ${fixes} yanlış tarih (işaretsiz; bilerek başka kaynak kullandıysanız dokunmayın)` : "") +
      ". Yalnız işaretliler uygulanır." +
      (uncertain ? ` Kimliği belirsiz ${uncertain} geçiş listede yok; onları Tüm belgeyi denetle ile tek tek seçin.` : "")
    : "Uygulanacak öneri yok." + (uncertain ? ` Kimliği belirsiz ${uncertain} geçiş için Tüm belgeyi denetle'yi kullanın.` : "");
  const list = $("bulk-list");
  list.replaceChildren();
  for (const s of suggestions) {
    const item = document.createElement("li");
    const box = document.createElement("input");
    box.type = "checkbox";
    box.checked = bulkChosen.has(s.id);
    box.addEventListener("click", (event) => {
      event.stopPropagation();
      if (box.checked) bulkChosen.add(s.id);
      else bulkChosen.delete(s.id);
      $<HTMLButtonElement>("bulk-apply").disabled = bulkChosen.size === 0;
    });
    const name = document.createElement("strong");
    name.textContent = describeSuggestion(s);
    const badge = document.createElement("span");
    badge.className = "badge";
    badge.textContent = placeOfRow(s.row);
    const meta = document.createElement("span");
    meta.className = "meta";
    meta.textContent =
      s.kind === "move"
        ? `İlk geçiş değişti: tarihi ilk geçişe taşı · ${textName(s.person.shortName)} · ${placeOfRow(s.row)}`
        : s.kind === "drop"
          ? `Fazla tarih: ilk geçişte zaten var · ${textName(s.person.shortName)} · ${placeOfRow(s.row)}`
          : s.kind === "date"
        ? `Tarih ekle · ${textName(s.person.shortName)}`
        : s.kind === "fixdate"
          ? `Tarihi katalogdakiyle değiştir · ${textName(s.person.shortName)} · ${s.row.dateCheck === "inconsistent" ? "hicrî ve milâdî yıl tutmuyor" : "katalogdan farklı"}`
          : `Kaynak yazımı · ${textName(s.person.shortName)}${s.row.confirmed ? " · onaylı kimlik" : ""}`;
    item.append(box, name, badge, meta);
    item.addEventListener("click", () => {
      markSelected(item);
      void run(() => chooseRowQuiet(s.row));
    });
    list.append(item);
  }
  $<HTMLButtonElement>("bulk-apply").disabled = bulkChosen.size === 0;
}

export async function chooseRowQuiet(row: InventoryRow): Promise<void> {
  const segment = state.segments[row.segment];
  if (!segment || !(await selectMention(segment, row.start, row.text))) setMessage("Bu geçiş metinde bulunamadı; listeyi yenileyin.", "warning");
}

export async function applyBulk(): Promise<void> {
  const chosen = applyOrder(suggestions.filter((s) => bulkChosen.has(s.id)));
  let done = 0;
  let missed = 0;
  for (const s of chosen) {
    const segment = state.segments[s.row.segment];
    if (!segment || !(await selectMention(segment, s.row.start, s.row.text))) {
      missed++;
      continue;
    }
    if (s.kind === "date") {
      const control = await insertDate(s.value, dateTag(s.person), `Vefeyât: ${textName(s.person.shortName)}`);
      logHistory("date", `${s.row.text} ${s.value}`, control);
    } else if (s.kind === "move" || s.kind === "drop") {
      // Selecting the receiving mention above found it; take the date off the later one first (it stands after).
      const from = s.from as InventoryRow;
      const fromSegment = state.segments[from.segment];
      if (!fromSegment || !(await removeDateAfter(fromSegment, from.start, from.text, s.old ?? ""))) {
        missed++;
        continue;
      }
      if (s.kind === "move") {
        if (!(await selectMention(segment, s.row.start, s.row.text))) {
          missed++;
          continue;
        }
        const control = await insertDate(s.value, dateTag(s.person), `Vefeyât: ${textName(s.person.shortName)}`);
        logHistory("date", `${s.row.text} ${s.value} (ilk geçişe taşındı)`, control);
      } else {
        logHistory("repair", `${from.text} ${s.old} kaldırıldı (ilk geçişte tarih var)`);
      }
    } else if (s.kind === "fixdate") {
      if (!(await replaceDateAfterSelection(s.old ?? "", s.value))) {
        missed++;
        continue;
      }
      logHistory("repair", `${s.row.text} ${s.old} → ${s.value}`);
    } else {
      await replaceSelection(s.value);
      logHistory("spelling", `${s.row.text} → ${s.value}`);
    }
    done++;
  }
  logHistory("bulk", `${done} öneri uygulandı`);
  await prepareBulk();
  setMessage(`${done} öneri uygulandı.` + (missed ? ` ${missed} geçiş metinde bulunamadı.` : ""), missed ? "warning" : "success");
}

export function renderExclusions(): void {
  const list = readExclusions();
  $("exclusion-summary").textContent = list.length
    ? `${list.length} ifade bu belgede şahıs adı sayılmıyor. Kaldırdığınız ifade yeniden taranır.`
    : "İstisna yok. Seçili ad penceresinde \"Şahıs değil olarak işaretle\" ile ekleyebilirsiniz.";
  const out = $("exclusion-list");
  out.replaceChildren();
  for (const key of list) {
    const item = document.createElement("li");
    const name = document.createElement("strong");
    name.textContent = key;
    const remove = document.createElement("button");
    remove.type = "button";
    remove.className = "link";
    remove.textContent = "Kaldır";
    remove.addEventListener("click", () =>
      void run(async () => {
        await writeExclusions(readExclusions().filter((k) => k !== key));
        renderExclusions();
      }),
    );
    item.append(name, remove);
    out.append(item);
  }
}

export function renderIdentities(): void {
  const views = listConfirmations(readConfirmations());
  const people = new Map((state.catalog?.people ?? []).map((p) => [personKey(p), p] as const));
  $("identity-summary").textContent = views.length
    ? `${views.length} geçiş için kimliği siz seçtiniz. Yanlış olanı silerseniz o geçiş yeniden "seçim gerekli" olur.`
    : "Kayıtlı kimlik onayı yok. Birden fazla adayı olan bir adda kişiyi seçip tarih eklediğinizde ya da “Bu yazımın bütün belirsiz geçişlerine uygula” dediğinizde burada görünür.";
  $("identity-clear").hidden = views.length === 0;
  const out = $("identity-list");
  out.replaceChildren();
  for (const view of views) {
    const person = people.get(view.personKey);
    const item = document.createElement("li");
    const name = document.createElement("strong");
    name.textContent = `${view.literal} → ${person ? textName(person.shortName) : "katalogda bulunamadı"}`;
    const meta = document.createElement("span");
    meta.className = "meta";
    meta.textContent = `…${view.before} [${view.literal}] ${view.after}…` + (person?.deathDate ? ` · ${person.deathDate}` : "");
    const remove = document.createElement("button");
    remove.type = "button";
    remove.className = "link";
    remove.textContent = "Onayı sil";
    remove.addEventListener("click", () =>
      void run(async () => {
        const map = { ...readConfirmations() };
        delete map[view.key];
        await writeConfirmations(map);
        renderIdentities();
      }),
    );
    item.append(name, meta, remove);
    out.append(item);
  }
}

export async function showCoverage(): Promise<void> {
  const c = await readCoverage();
  const lines: Array<[string, string]> = [
    ["Taranır", `Ana metin: ${c.paragraphs} paragraf (tablolar dahil).`],
    c.footnotes === null ? ["Taranmaz", "Dipnotlar: bu Word sürümü dipnot okumayı desteklemiyor (WordApi 1.5)."] : ["Taranır", `Dipnotlar: ${c.footnotes}.`],
    c.endnotes === null ? ["Taranmaz", "Sonnotlar: bu Word sürümünde okunamıyor."] : ["Taranır", `Sonnotlar: ${c.endnotes}.`],
    ["Taranmaz", `Üst ve alt bilgiler: ${c.headers} dolu bölüm.`],
    ["Taranmaz", c.comments === null ? "Açıklamalar: bu Word sürümünde okunamıyor." : `Açıklamalar: ${c.comments}.`],
    ["Taranmaz", "Metin kutuları ve şekiller içindeki metin (Office.js bunlara erişemiyor)."],
  ];
  const out = $("coverage-list");
  out.replaceChildren();
  for (const [kind, text] of lines) {
    const item = document.createElement("li");
    const label = document.createElement("strong");
    label.textContent = `${kind}: `;
    item.append(label, text);
    if (kind === "Taranmaz") item.className = "muted";
    out.append(item);
  }
  const exclusions = readExclusions().length;
  if (exclusions) {
    const item = document.createElement("li");
    item.textContent = `Atlanan: ${exclusions} ifade "Şahıs değil" olarak işaretli (Belge istisnaları).`;
    out.append(item);
  }
}

export function renderSpellings(): void {
  const map = readSpellings();
  const people = new Map((state.catalog?.people ?? []).map((p) => [personKey(p), p] as const));
  const entries = Object.entries(map);
  $("spellings-summary").textContent = entries.length
    ? `${entries.length} şahıs için belgeye özgü yazım var.`
    : "Yazım tercihi yok. Bir adı seçip Vefeyât → Ayarlar → Seçili yazımı tercih et ile ekleyebilirsiniz.";
  $("spellings-clear").hidden = entries.length === 0;
  const out = $("spellings-list");
  out.replaceChildren();
  for (const [key, spelling] of entries) {
    const person = people.get(key);
    const item = document.createElement("li");
    const name = document.createElement("strong");
    name.textContent = spelling;
    const meta = document.createElement("span");
    meta.className = "meta";
    meta.textContent = person ? `${textName(person.shortName)} · ${person.deathDate || "tarih yok"}` : "katalogda bulunamadı";
    const remove = document.createElement("button");
    remove.type = "button";
    remove.className = "link";
    remove.textContent = "Kaldır";
    remove.addEventListener("click", () =>
      void run(async () => {
        const next = { ...readSpellings() };
        delete next[key];
        await writeSpellings(next);
        renderSpellings();
      }),
    );
    item.append(name, meta, remove);
    out.append(item);
  }
}

export async function clearSpellings(): Promise<void> {
  versions.identity++; // identities may change
  const count = Object.keys(readSpellings()).length;
  await writeSpellings({});
  renderSpellings();
  setMessage(count ? `${count} yazım tercihi silindi.` : "Silinecek yazım tercihi yok.", "success");
}

export function prefsText(): string {
  return exportPrefs(state.settings, readSpellings(), readExclusions(), readConfirmations(), readLocalPersons());
}

export function downloadPrefs(): void {
  const url = URL.createObjectURL(new Blob([prefsText()], { type: "application/json" }));
  const link = document.createElement("a");
  link.href = url;
  link.download = "vefeyat-tercihler.json";
  document.body.append(link);
  link.click();
  link.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
  setMessage("Tercih dosyası indirildi. İndirme açılmadıysa Panoya kopyala'yı kullanın.", "success");
}

export async function copyPrefs(): Promise<void> {
  const text = prefsText();
  const area = $<HTMLTextAreaElement>("prefs-text");
  try {
    await navigator.clipboard.writeText(text);
    setMessage("Tercihler panoya kopyalandı.", "success");
  } catch {
    area.value = text;
    area.select();
    setMessage("Pano kullanılamadı; metin aşağıda seçili, Ctrl+C ile kopyalayın.", "info");
  }
}

export async function applyPrefs(): Promise<void> {
  versions.identity++; // identities may change
  const parsed = parsePrefs($<HTMLTextAreaElement>("prefs-text").value);
  if (!parsed.ok) {
    setMessage(parsed.reason, "warning");
    return;
  }
  const merged = mergePrefs({ spellings: readSpellings(), exclusions: readExclusions(), confirmations: readConfirmations(), persons: readLocalPersons() }, parsed.file);
  for (const [name, value] of Object.entries(parsed.file.settings)) await writeSetting(name as SettingName, value as string);
  await writeSpellings(merged.spellings);
  await writeExclusions(merged.exclusions);
  await writeConfirmations(merged.confirmations);
  await writeLocalPersons(merged.persons);
  applyLocalPersons();
  await loadSettings();
  renderSpellings();
  setMessage(
    `Tercihler uygulandı: ${Object.keys(parsed.file.settings).length} ayar, ${Object.keys(parsed.file.spellings).length} yazım, ${parsed.file.exclusions.length} istisna, ${Object.keys(parsed.file.confirmations).length} kimlik onayı, ${parsed.file.persons.length} kendi şahıs kaydı.`,
    "success",
  );
}

export async function noProofing(): Promise<void> {
  if (!canSkipProofing()) {
    setMessage("Bu Word sürümü yazım denetimi ayarını değiştirmeye izin vermiyor (WordApiDesktop 1.3). Microsoft 365 masaüstü Word'ü güncelleyin.", "warning");
    return;
  }
  const rows = (await documentRows()).filter((r) => r.person !== null || r.candidates.length > 0);
  if (rows.length === 0) {
    setMessage("Belgede şahıs adı bulunamadı.", "info");
    return;
  }
  const marked = await skipProofing(state.segments, rows, progress("Şahıs adları işaretleniyor"));
  logHistory("repair", `${marked} şahıs adı yazım denetiminden çıkarıldı`);
  setMessage(`${marked} şahıs adı yazım denetiminden çıkarıldı. Bu adlara sağ tıklayınca Vefeyât menüsü görünür.`, "success");
}

export async function toggleFrames(): Promise<void> {
  const result = await toggleDateFrames();
  setMessage(
    result.count === 0 ? "Belgede Vefeyât tarihi yok." : result.visible ? `${result.count} tarih alanı çerçeveyle gösteriliyor.` : `${result.count} tarih alanının çerçevesi gizlendi.`,
    "success",
  );
}

// ---- İlk geçiş bakımı: the document changed ----

/**
 * After an edit, quietly re-reads the changed paragraphs and tells when a
 * person's first mention moved (a new earlier mention, or a date left on a
 * later one). Nothing changes until the user reviews and applies it.
 */
export async function checkFirstUse(): Promise<void> {
  if (state.busy || !state.catalog || (state.settings.scope !== "first" && state.settings.scope !== "chapter")) return;
  // The first quiet read primes the stored reading; afterwards only edits trigger a look.
  if (documentRead() && !hasPendingChanges()) return;
  if (!(await watchDocument())) return; // without paragraph events every look would be a full read
  const rows = await documentRows();
  const found = firstUseSuggestions(rows, state.settings);
  const notice = $("firstuse-notice");
  notice.hidden = found.length === 0;
  if (found.length === 0) return;
  const names = [...new Set(found.map((s) => textName(s.person.shortName)))];
  $("firstuse-text").textContent =
    `İlk geçiş değişti: ${names.slice(0, 3).join(", ")}${names.length > 3 ? ` ve ${names.length - 3} kişi daha` : ""}. ` +
    `${found.filter((s) => s.kind === "move").length} tarih ilk geçişe taşınmalı, ${found.filter((s) => s.kind === "drop").length} fazla tarih var. Onayınız olmadan bir şey değişmez.`;
}
