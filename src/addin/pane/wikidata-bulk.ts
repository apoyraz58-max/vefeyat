/**
 * Katalogda yok → Wikidata'da toplu ara: every name of the Şahıs dökümü that no
 * catalog has is looked up on Wikidata, one after another; the user checks the
 * proposed person for each and adds the chosen ones to Kendi şahıslarım in one go.
 */
import { nameQuery, normalize } from "../../core/index.js";
import { type WikidataHit, searchWikidata } from "../wikidata-live.js";
import { $, setMessage, state } from "./core.js";
import { applyLocalPersons } from "./catalog.js";
import { addLocals, localFromHit, standardLabel } from "./wikidata.js";

/** At most this many names in one search: Wikidata is asked once per name. */
export const BULK_LIMIT = 40;

interface Found {
  readonly name: string;
  readonly count: number;
  readonly hits: readonly WikidataHit[];
}

let found: Found[] = [];
let searching = false;

/** The names no catalog has, most frequent first. */
export function unknownNames(): Array<{ name: string; count: number }> {
  const counts = new Map<string, { name: string; count: number }>();
  for (const row of state.rows) {
    if (row.status !== "unknown") continue;
    const name = nameQuery(row.text).trim();
    const key = normalize(name);
    if (!key) continue;
    const known = counts.get(key);
    if (known) known.count++;
    else counts.set(key, { name, count: 1 });
  }
  return [...counts.values()].sort((a, b) => b.count - a.count || a.name.localeCompare(b.name, "tr"));
}

/** The Katalogda yok section of Şahıs dökümü: shown when the document has such names. */
export function renderBulkOffer(): void {
  const names = unknownNames();
  const section = $("wikidata-bulk");
  section.hidden = names.length === 0;
  if (names.length === 0) return;
  if (!searching && found.length === 0) {
    $("wikidata-bulk-summary").textContent =
      `Belgede hiçbir katalogda olmayan ${names.length} ad var. Wikidata'da aranıp ölüm tarihi bilinen kişiler önerilir; ` +
      `siz onaylayınca kendi şahıslarınıza eklenir.` + (names.length > BULK_LIMIT ? ` (En sık geçen ${BULK_LIMIT} ad aranır.)` : "");
  }
}

/** A hit that is surely the name: every word of the name is in its label or English name. */
function sure(name: string, hit: WikidataHit): boolean {
  const words = normalize(name).split(" ").filter((w) => w.length > 2);
  const label = normalize(`${hit.label} ${hit.english} ${standardLabel(hit)}`);
  return words.length > 0 && words.every((w) => label.includes(w));
}

export async function searchUnknown(): Promise<void> {
  if (searching) return;
  const names = unknownNames().slice(0, BULK_LIMIT);
  if (names.length === 0) return;
  searching = true;
  found = [];
  const go = $<HTMLButtonElement>("wikidata-bulk-search");
  go.disabled = true;
  $("wikidata-bulk-list").replaceChildren();
  $<HTMLButtonElement>("wikidata-bulk-add").hidden = true;
  let failed = 0;
  try {
    for (const [k, entry] of names.entries()) {
      $("wikidata-bulk-summary").textContent = `Wikidata'da aranıyor: ${k + 1} / ${names.length} (${entry.name})…`;
      try {
        found.push({ ...entry, hits: (await searchWikidata(entry.name)).slice(0, 4) });
      } catch {
        failed++;
        found.push({ ...entry, hits: [] });
        // Wikidata is not answering: stop rather than wait on every name.
        if (failed >= 3 && failed === k + 1) break;
      }
    }
  } finally {
    searching = false;
    go.disabled = false;
  }
  renderFound(failed);
}

function renderFound(failed: number): void {
  const list = $("wikidata-bulk-list");
  list.replaceChildren();
  const withHits = found.filter((f) => f.hits.length > 0);
  $("wikidata-bulk-summary").textContent =
    failed === found.length
      ? "Wikidata şu an yanıt vermiyor ya da internete ulaşılamıyor. Biraz sonra yeniden deneyin."
      : `${found.length} ad arandı, ${withHits.length} ad için öneri var. Doğru kişiyi seçin ("Eklenmesin" seçili olanlar eklenmez).` +
        (failed ? ` ${failed} ad için Wikidata yanıt vermedi.` : "");
  for (const [index, entry] of found.entries()) {
    const item = document.createElement("li");
    const title = document.createElement("strong");
    title.textContent = entry.name;
    const count = document.createElement("span");
    count.className = "meta";
    count.textContent = ` ${entry.count} geçiş`;
    item.append(title, count, document.createElement("br"));
    if (entry.hits.length === 0) {
      const none = document.createElement("span");
      none.className = "muted";
      none.textContent = "Wikidata'da ölüm tarihi bilinen bir kişi bulunamadı.";
      item.append(none);
    } else {
      const select = document.createElement("select");
      select.className = "wide";
      select.id = `wikidata-bulk-${index}`;
      select.setAttribute("aria-label", `${entry.name} için Wikidata önerisi`);
      const skip = new Option("Eklenmesin", "");
      select.append(skip);
      entry.hits.forEach((hit, k) => {
        const label = standardLabel(hit);
        const text = `${label} ${hit.deathDate}${hit.description ? ` · ${hit.description.slice(0, 70)}` : ""}`;
        select.append(new Option(text, String(k)));
      });
      // Only a hit that surely bears the name is proposed; the rest wait for the user.
      const first = entry.hits[0];
      select.value = first && sure(entry.name, first) ? "0" : "";
      item.append(select);
    }
    list.append(item);
  }
  $<HTMLButtonElement>("wikidata-bulk-add").hidden = withHits.length === 0;
}

export async function addFound(): Promise<void> {
  const records = [];
  const reasons: string[] = [];
  for (const [index, entry] of found.entries()) {
    const select = document.getElementById(`wikidata-bulk-${index}`) as HTMLSelectElement | null;
    if (!select || select.value === "") continue;
    const hit = entry.hits[Number(select.value)];
    if (!hit) continue;
    const record = localFromHit(entry.name, hit);
    if (typeof record === "string") reasons.push(`${entry.name}: ${record}`);
    else records.push(record);
  }
  if (records.length === 0) {
    setMessage(reasons[0] ?? "Eklenecek kişi seçilmedi.", "warning");
    return;
  }
  await addLocals(records);
  applyLocalPersons();
  found = [];
  $("wikidata-bulk-list").replaceChildren();
  $<HTMLButtonElement>("wikidata-bulk-add").hidden = true;
  setMessage(`${records.length} kişi kendi şahıslarınıza eklendi; döküm yenileniyor.` + (reasons.length ? ` ${reasons.length} kişi eklenemedi.` : ""), "success");
}
