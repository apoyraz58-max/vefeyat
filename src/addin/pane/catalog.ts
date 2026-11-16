/** Vefeyât pane — the catalog (embedded or the user's CSV) and Şahıs ara. */
import { Catalog, CatalogError } from "../../core/index.js";
import { candidateLabel } from "../logic.js";
import { type StoredCatalog, clearCatalog, fetchEmbeddedCatalog, loadCatalog, saveCatalog } from "../store.js";

import { $, setMessage, showView, state, versions } from "./core.js";
import { resetTags } from "./health.js";
import { refresh } from "./selected.js";

export function runSearch(): void {
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

// ---- Catalog ----

export function describeCatalog(fileName: string, count: number, source: "user" | "embedded", persisted: boolean): void {
  const origin = source === "embedded" ? "eklentiyle gelen katalog" : `sizin seçtiğiniz dosya: ${fileName}`;
  $("catalog-status").textContent = `${count.toLocaleString("tr-TR")} şahıs · ${origin}${persisted ? "" : " (yalnız bu oturum)"}`;
  $("reset-catalog").hidden = source !== "user";
  $("no-catalog").hidden = true;
}

export function useCatalog(csv: string, fileName: string, source: "user" | "embedded", persisted: boolean): void {
  state.catalog = Catalog.fromCsv(csv);
  versions.identity++;
  resetTags();
  state.index = null; // rebuilt on the next inventory
  state.catalogSource = source;
  describeCatalog(fileName, state.catalog.people.length, source, persisted);
}

/** User's own CSV first; otherwise the catalog published with the add-in. */
export async function restoreCatalog(): Promise<void> {
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
  resetTags();
  $("catalog-status").textContent = "Katalog yüklenemedi. Bilgisayarınızdaki DİA CSV dosyasını seçin.";
  $("no-catalog").hidden = false;
}

export async function chooseCatalog(file: File): Promise<void> {
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

export async function resetCatalog(): Promise<void> {
  await clearCatalog();
  await restoreCatalog();
  setMessage(state.catalogSource === "embedded" ? "Eklentinin kataloğuna dönüldü." : "Eklentiyle gelen katalog bulunamadı.", "info");
}
