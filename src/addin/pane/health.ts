/** Vefeyât pane — Belge sağlığı, VBA tarihlerini dönüştür, Tarihin kaynağı. */
import { NameIndex, textName } from "../../core/index.js";
import { SETTING_CHOICES, dateTag } from "../logic.js";
import { DATE_TAG_PREFIX, HEALTH_LABEL, type HealthRow, TagIndex, checkControls, repairText, repairable } from "../health.js";
import { readDateSpacing, fixDateSpacing, type SpacingIssue, removeEmptyDateControls, removeEmptyDateControlsIn, countLegacyDates, convertLegacyDates, readDateControls, selectDateControl, selectedDateControl, writeDateControl, readSettings, refreshSettings } from "../word.js";

import { $, logHistory, markSelected, progress, run, setMessage, state } from "./core.js";

// ---- Belge sağlığı ve Tarihin kaynağı ----

export let tagIndex: TagIndex | null = null;

export let healthRows: HealthRow[] = [];

export const healthChosen = new Set<number>();

export function tags(): TagIndex | null {
  if (!state.catalog) return null;
  tagIndex ??= new TagIndex(state.catalog);
  return tagIndex;
}

/**
 * Removes the empty controls left when a Vefeyât date was deleted by hand
 * (Word shows "Metin girmek için buraya tıklayın" there). Visible text does not change.
 */
export async function tidyEmptyDates(ids?: readonly string[]): Promise<number> {
  const removed = await (ids ? removeEmptyDateControlsIn(ids).catch(() => removeEmptyDateControls()) : removeEmptyDateControls()).catch(() => 0);
  if (removed) logHistory("repair", `Silinmiş tarihlerin boş kalan ${removed} alanı kaldırıldı`);
  return removed;
}

export async function checkHealth(): Promise<void> {
  const index = tags();
  if (!index) {
    setMessage("Katalog yüklenmedi.", "warning");
    return;
  }
  await tidyEmptyDates();
  await refreshSettings();
  state.settings = readSettings();
  healthRows = checkControls(await readDateControls(), index, state.settings.prefix, state.settings.dateForm);
  const legacy = await countLegacyDates().catch(() => 0);
  $("legacy").hidden = legacy === 0;
  $("legacy-text").textContent =
    `Belgede VBA sürümünün eklediği ${legacy} tarih var (Word alanı). Bunlar yeni sürümde denetlenemez ve teslim temizliğinde kalır. ` +
    "Dönüştürünce aynı tarih metni, adın arkasında Vefeyât tarihi olarak kalır; adı bulunamayan alana dokunulmaz.";
  spacing = await readDateSpacing().catch(() => []);
  renderSpacing();
  healthChosen.clear();
  for (const row of healthRows) if (row.health === "prefix" || row.health === "empty") healthChosen.add(row.id);
  renderHealth();
}

let spacing: SpacingIssue[] = [];

function renderSpacing(): void {
  $("spacing").hidden = spacing.length === 0;
  if (!spacing.length) return;
  const after = spacing.filter((s) => s.missingAfter).length;
  const before = spacing.filter((s) => s.doubleBefore || s.missingBefore).length;
  $("spacing-text").textContent =
    `${spacing.length} tarihin çevresinde boşluk bozuk` +
    ` (${[after ? `${after} tarih sonraki kelimeye bitişik` : "", before ? `${before} tarihin önünde çift ya da eksik boşluk` : ""].filter(Boolean).join(", ")}).` +
    ` Örnek: ${spacing[0]!.sample}`;
}

/** "Boşlukları düzelt": one space before and after every Vefeyât date. */
export async function fixSpacing(): Promise<void> {
  const fixed = await fixDateSpacing(spacing.map((s) => s.id));
  logHistory("repair", `${fixed} tarihin çevresindeki boşluk düzeltildi`);
  await checkHealth();
  setMessage(`${fixed} tarihin çevresindeki boşluk düzeltildi.`, "success");
}

export function renderHealth(): void {
  const problems = healthRows.filter((r) => r.health !== "ok");
  $("health-summary").textContent =
    healthRows.length === 0
      ? "Belgede Vefeyât'ın eklediği tarih yok."
      : `${healthRows.length} tarih denetlendi: ${healthRows.length - problems.length} sağlam, ${problems.length} sorunlu.` +
        (problems.length ? " Onarılacakları işaretleyin; işaretsizlere dokunulmaz." : "");
  const list = $("health-list");
  list.replaceChildren();
  for (const row of problems) {
    const item = document.createElement("li");
    const box = document.createElement("input");
    box.type = "checkbox";
    box.disabled = !repairable(row);
    box.checked = healthChosen.has(row.id);
    box.addEventListener("click", (event) => {
      event.stopPropagation();
      if (box.checked) healthChosen.add(row.id);
      else healthChosen.delete(row.id);
      $<HTMLButtonElement>("health-repair").disabled = healthChosen.size === 0;
    });
    const name = document.createElement("strong");
    name.textContent = row.person ? textName(row.person.shortName) : "Bilinmeyen kayıt";
    const badge = document.createElement("span");
    badge.className = "badge";
    badge.textContent = row.part === "note" ? "Not" : "Metin";
    const meta = document.createElement("span");
    meta.className = "meta";
    meta.textContent = `Belgede: ${row.text.trim() || "(boş)"}` + (row.expected ? ` → ${row.expected}` : "");
    const status = document.createElement("span");
    status.className = "meta warn";
    status.textContent = HEALTH_LABEL[row.health];
    item.append(box, name, badge, meta, status);
    item.addEventListener("click", () => {
      markSelected(item);
      void run(async () => {
        if (!(await selectDateControl(row.id))) setMessage("Bu tarih artık belgede yok; yeniden denetleyin.", "warning");
      });
    });
    list.append(item);
  }
  $<HTMLButtonElement>("health-repair").disabled = healthChosen.size === 0;
}

export async function convertLegacy(): Promise<void> {
  if (!state.catalog) {
    setMessage("Katalog yüklenmedi.", "warning");
    return;
  }
  state.index ??= new NameIndex(state.catalog.people);
  const result = await convertLegacyDates(state.index, (person) => (person ? dateTag(person) : `${DATE_TAG_PREFIX}00000000`), progress("VBA tarihleri dönüştürülüyor"));
  logHistory("repair", `${result.converted} VBA tarihi dönüştürüldü`);
  await checkHealth();
  setMessage(
    `${result.converted} VBA tarihi dönüştürüldü.` +
      (result.unresolved ? ` ${result.unresolved} tanesinin kişisi belirlenemedi; listede "Katalogda kaydı yok" olarak görünür.` : "") +
      (result.skipped ? ` ${result.skipped} alanın önündeki ad bulunamadı; olduğu gibi bırakıldı.` : ""),
    result.skipped ? "warning" : "success",
  );
}

export async function repairHealth(): Promise<void> {
  let fixed = 0;
  for (const row of healthRows) {
    if (!healthChosen.has(row.id) || !repairable(row)) continue;
    if (await writeDateControl(row.id, repairText(row, row.text))) {
      fixed++;
      logHistory("repair", `${row.text.trim() || "(boş)"} → ${row.expected}`, row.id);
    }
  }
  await checkHealth();
  setMessage(`${fixed} tarih onarıldı.`, "success");
}

export let sourceRow: HealthRow | null = null;

export async function showSource(): Promise<void> {
  const index = tags();
  if (!index) return;
  const control = await selectedDateControl();
  sourceRow = control ? (checkControls([control], index, state.settings.prefix, state.settings.dateForm)[0] ?? null) : null;
  $("source-details").hidden = sourceRow === null;
  $("source-repair").hidden = !(sourceRow && repairable(sourceRow));
  if (!sourceRow) {
    $("source-summary").textContent = "İmleci belgedeki bir Vefeyât tarihinin içine (ya da tarihi olan adın üzerine) koyun.";
    return;
  }
  const person = sourceRow.person;
  $("source-summary").textContent = person ? "Bu tarih Vefeyât tarafından aşağıdaki katalog kaydından eklendi." : "Bu tarihin kaydı yüklü katalogda yok; katalog değişmiş olabilir.";
  $("source-person").textContent = person ? `${textName(person.shortName)}${person.fullName ? ` (${textName(person.fullName)})` : ""}` : "—";
  $("source-text").textContent = sourceRow.text.trim() || "(boş)";
  $("source-catalog").textContent = person?.deathDate || "—";
  $("source-health").textContent = HEALTH_LABEL[sourceRow.health];
  const link = $<HTMLAnchorElement>("source-link");
  link.textContent = person?.source || "—";
  if (person?.source.startsWith("https://")) link.href = person.source;
  else link.removeAttribute("href");
}

export async function repairSource(): Promise<void> {
  if (!sourceRow || !repairable(sourceRow)) return;
  if (await writeDateControl(sourceRow.id, repairText(sourceRow, sourceRow.text))) {
    logHistory("repair", `${sourceRow.text.trim() || "(boş)"} → ${sourceRow.expected}`, sourceRow.id);
    setMessage("Tarih onarıldı.", "success");
  }
  await showSource();
}

export function resetTags(): void {
  tagIndex = null;
}

// ---- Kısaltma değişince ----

let prefixRows: HealthRow[] = [];

/**
 * After Kısaltma changes (ribbon Ayarlar menu, Tercihleri al): when Vefeyât
 * dates in the document still carry another abbreviation, offers to convert
 * them. Hand-written dates are not Vefeyât dates and stay as they are.
 */
export async function offerPrefix(): Promise<void> {
  const index = tags();
  if (!index) return;
  await refreshSettings();
  state.settings = readSettings();
  prefixRows = checkControls(await readDateControls(), index, state.settings.prefix, state.settings.dateForm).filter((r) => r.health === "prefix" && repairable(r));
  $("prefix-offer").hidden = prefixRows.length === 0;
  if (!prefixRows.length) return;
  const label = SETTING_CHOICES.prefix[state.settings.prefix];
  $("prefix-offer-text").textContent =
    `Kısaltma: ${label}. Belgedeki ${prefixRows.length} Vefeyât tarihi başka kısaltmayla yazılı` +
    ` (ör. ${prefixRows[0]!.text.trim()} → ${prefixRows[0]!.expected}). Bunlar da çevrilsin mi? Elle yazılmış tarihler değişmez.`;
}

export async function convertPrefix(): Promise<void> {
  $("prefix-offer").hidden = true;
  let fixed = 0;
  for (const row of prefixRows) {
    if (await writeDateControl(row.id, repairText(row, row.text))) {
      fixed++;
      logHistory("repair", `${row.text.trim()} → ${row.expected}`, row.id);
    }
  }
  prefixRows = [];
  setMessage(`${fixed} tarih yeni kısaltmaya çevrildi.`, "success");
}

export function declinePrefix(): void {
  $("prefix-offer").hidden = true;
  prefixRows = [];
  setMessage("Yalnız ayar değişti; belgedeki tarihler olduğu gibi kaldı. Sonradan Belge sağlığı'ndan çevirebilirsiniz.", "info");
}
