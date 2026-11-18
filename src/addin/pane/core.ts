/** Vefeyât pane — shared state, messages, the busy bar and long-operation runner, the cached reading of the document, İşlem geçmişi. */
import { Catalog, NameIndex, type Person } from "../../core/index.js";
import { DEFAULT_SETTINGS, type Lookup, type Settings, type View, settingsSummary } from "../logic.js";
import { type InventoryFilter, type InventoryRow, type Segment, buildInventoryAsync, placeOf } from "../inventory.js";
import { type HistoryEntry, type HistoryKind, addHistory, parseHistory } from "../bulk.js";
import { readSpellings, readExclusions, readConfirmations, readDocument, documentVersion, readSettings, refreshSettings } from "../word.js";

import { renderDetails } from "./selected.js";
import { renderHistory } from "./tools.js";

declare const __VERSION__: string;

export interface State {
  catalog: Catalog | null;
  /** The catalog from the CSV alone; `catalog` adds Kendi şahıslarım to it. */
  baseCatalog: Catalog | null;
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
  /** Şahıs dökümü süzgeci and name search. */
  filter: InventoryFilter;
  query: string;
  selectedRow: number | null;
  audit: InventoryRow[] | null;
  auditAt: number;
  mention: string | null;
  confirmed: boolean;
}

export const state: State = {
  catalog: null,
  baseCatalog: null,
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
  filter: "all",
  query: "",
  selectedRow: null,
  audit: null,
  auditAt: 0,
  mention: null,
  confirmed: false,
};

export const PAGE_SIZE = 250;

export const $ = <T extends HTMLElement>(id: string): T => {
  const element = document.getElementById(id);
  if (!element) throw new Error(`Missing #${id}`);
  return element as T;
};

export function setMessage(text: string, kind: "info" | "warning" | "error" | "success" = "info"): void {
  const box = $("message");
  box.textContent = text;
  box.className = `message ${kind}`;
  box.hidden = text.length === 0;
}

export function showView(view: View): void {
  document.querySelectorAll<HTMLElement>(".view").forEach((element) => {
    const active = element.id === `view-${view}`;
    element.hidden = !active;
    if (active) $("view-title").textContent = element.dataset.title ?? "Vefeyât";
  });
  $("back").hidden = view === "secili";
}

export async function loadSettings(): Promise<void> {
  await refreshSettings();
  state.settings = readSettings();
  $("settings-summary").textContent = settingsSummary(state.settings);
}

export let busyLabel = "";

export let busyStarted = 0;

export let busyTimer = 0;

export let cancelRequested = false;

/** Thrown from a progress callback when the user pressed İptal. */
export class Cancelled extends Error {}

/** Progress callback for long Word operations: updates the bar, honours İptal. */
export function progress(label: string): (done: number, total: number) => void {
  const started = Date.now();
  return (done, total) => {
    if (cancelRequested) throw new Cancelled();
    let eta = "";
    const elapsed = (Date.now() - started) / 1000;
    if (done > 0 && done < total && elapsed > 5) {
      const left = Math.round((elapsed / done) * (total - done));
      eta = left >= 60 ? ` · kalan ~${Math.round(left / 60)} dk` : ` · kalan ~${left} sn`;
    }
    $("busy-text").textContent = `${label}… ${done.toLocaleString("tr-TR")} / ${total.toLocaleString("tr-TR")}${eta}`;
  };
}

export function showBusy(label: string): void {
  busyLabel = label;
  busyStarted = Date.now();
  cancelRequested = false;
  $("busy").hidden = false;
  $("busy-text").textContent = `${label}…`;
  $("busy-time").textContent = "";
  window.clearInterval(busyTimer);
  busyTimer = window.setInterval(() => {
    const seconds = Math.round((Date.now() - busyStarted) / 1000);
    $("busy-time").textContent = seconds >= 60 ? `${Math.floor(seconds / 60)} dk ${seconds % 60} sn` : `${seconds} sn`;
  }, 1000);
}

export function hideBusy(): void {
  window.clearInterval(busyTimer);
  $("busy").hidden = true;
}

export async function run(task: () => Promise<void>, label = ""): Promise<void> {
  if (state.busy) {
    // Never ignore a click silently: say what is still running.
    setMessage(`${busyLabel || "Önceki işlem"} sürüyor. Bitmesini bekleyin ya da İptal'e basın.`, "warning");
    $("busy").hidden = !busyLabel;
    return;
  }
  state.busy = true;
  if (label) showBusy(label);
  renderDetails();
  try {
    await task();
  } catch (error) {
    if (error instanceof Cancelled || (error as { cause?: unknown }).cause instanceof Cancelled) {
      setMessage("İptal edildi. Yarım kalan işin izlerini ilgili komutla temizleyebilirsiniz (ör. Dizini tamamen kaldır).", "warning");
      return;
    }
    const office = ((error as { cause?: unknown }).cause ?? error) as { code?: string; debugInfo?: { errorLocation?: string; statement?: string } };
    const detail = [office.code, office.debugInfo?.errorLocation].filter(Boolean).join(" · ");
    console.error(error);
    setMessage(`Word işlemi tamamlanamadı: ${String(error)}${detail ? ` (${detail})` : ""} [Vefeyât ${__VERSION__}]`, "error");
  } finally {
    state.busy = false;
    busyLabel = "";
    hideBusy();
    renderDetails();
  }
}

/** Reading the inventory was built from: document version and identity settings version. */

/** The document read as a whole (cached for two minutes; inserting dates does not change identities). */
/** Reading the inventory was built from, and identity-settings changes since. */
export const versions = { rowsBuiltFor: "", identity: 0 };

export async function documentRows(): Promise<InventoryRow[]> {
  if (!state.catalog) return [];
  const read = await readDocument();
  // Confirmations, exclusions and preferred spellings change identities too.
  const stamp = `${documentVersion()}|${versions.identity}|${JSON.stringify([readConfirmations(), readExclusions(), readSpellings()])}`;
  if (state.rows.length > 0 && stamp === versions.rowsBuiltFor) return state.rows;
  state.index ??= new NameIndex(state.catalog.people);
  state.segments = read.segments;
  state.rows = await buildInventoryAsync(read.segments, state.catalog, state.index, readConfirmations(), readExclusions(), readSpellings());
  versions.rowsBuiltFor = stamp;
  return state.rows;
}

// ---- Şahıs dökümü ----

/** "3 dipnot, 1 sonnot": what was read, so a missing note shows at once. */
export function noteCounts(segments: readonly Segment[]): string {
  const foot = segments.filter((s) => placeOf(s) === "Dipnot").length;
  const end = segments.filter((s) => placeOf(s) === "Sonnot").length;
  return `${foot} dipnot, ${end} sonnot`;
}

export const placeOfRow = (row: { segment: number }) => placeOf(state.segments[row.segment]);

/** The row just clicked in a list stays marked, the others are unmarked. */
export function markSelected(item: HTMLElement): void {
  for (const sibling of item.parentElement?.children ?? []) sibling.setAttribute("aria-selected", "false");
  item.setAttribute("aria-selected", "true");
}

// ---- İşlem geçmişi ----

export const HISTORY_KEY = "vefeyat.history";

export function loadHistory(): HistoryEntry[] {
  try {
    return parseHistory(window.sessionStorage.getItem(HISTORY_KEY));
  } catch {
    return [];
  }
}

export let history: HistoryEntry[] = loadHistory();

export function logHistory(kind: HistoryKind, text: string, control?: number): void {
  history = addHistory(history, control === undefined ? { kind, text } : { kind, text, control });
  try {
    window.sessionStorage.setItem(HISTORY_KEY, JSON.stringify(history));
  } catch {
    /* private mode: kept in memory only */
  }
  if (!$("view-gecmis").hidden) renderHistory();
}

export function clearHistory(): void {
  history = [];
  try {
    window.sessionStorage.removeItem(HISTORY_KEY);
  } catch {
    /* ignored */
  }
}

export function requestCancel(): void {
  cancelRequested = true;
}
