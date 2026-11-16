/**
 * Toplu düzeltme önizlemesi, Denetime devam et and İşlem geçmişi:
 * Word-independent helpers (VBA 0.3.41 `VfbPreview`, `VfwResume`, `VfgLog`).
 */
import { type Person, personKey, formatDate, isValidDate, nameQuery, sourceName, textName, withEnding } from "../core/index.js";
import type { InventoryRow, Segment } from "./inventory.js";
import { type Settings, mentionKey } from "./logic.js";

export interface Suggestion {
  readonly id: string;
  readonly kind: "date" | "spelling" | "fixdate" | "move" | "drop";
  readonly row: InventoryRow;
  readonly person: Person;
  /** Date to insert, the replacement spelling (with the Turkish ending), or the corrected date. */
  readonly value: string;
  /** For "fixdate": the date as written now. For "move"/"drop": the date to take away. */
  readonly old?: string;
  /** For "move" and "drop": the later mention that carries the date now (`row` is where it goes, for "move"). */
  readonly from?: InventoryRow;
  /** Checked by default: safe changes only. */
  readonly preselected: boolean;
}

/**
 * Spelling fixes for certain mentions written differently from the catalog,
 * and dates for first mentions without one (unless Tarih ekleme is
 * "Yalnız denetle"; with "Bütün kullanımlar", every undated certain mention).
 */
/** Identities the reading only guessed: their changes always wait for an explicit tick. */
const guessed = (row: InventoryRow) => row.resolved === "context" || row.resolved === "topic";

/**
 * İlk geçiş bakımı ("Yalnız ilk kullanım"): in the main text and in the notes
 * separately, a person's date belongs to the first mention. A date found on a
 * later mention is moved to the first (kept as written) when the first has
 * none, and dropped when the first already has one. Only dates that agree with
 * the catalog are touched.
 */
export function firstUseSuggestions(rows: readonly InventoryRow[], settings: Settings): Suggestion[] {
  if (settings.scope !== "first") return [];
  const groups = new Map<string, InventoryRow[]>();
  for (const row of rows) {
    if (!row.person) continue;
    const key = `${row.part}|${personKey(row.person)}`;
    groups.set(key, [...(groups.get(key) ?? []), row]);
  }
  const out: Suggestion[] = [];
  for (const group of groups.values()) {
    const first = group[0] as InventoryRow;
    const person = first.person as Person;
    const later = group.slice(1).filter((r) => r.date && r.dateCheck === "ok");
    if (later.length === 0) continue;
    let rest = later;
    if (!first.date) {
      const source = later[0] as InventoryRow;
      out.push({ id: `m${source.id}`, kind: "move", row: first, from: source, person, value: source.date, old: source.date, preselected: !guessed(first) && !guessed(source) });
      rest = later.slice(1);
    }
    for (const extra of rest) {
      out.push({ id: `x${extra.id}`, kind: "drop", row: extra, from: extra, person, value: "", old: extra.date, preselected: !guessed(extra) });
    }
  }
  return out;
}

export function bulkSuggestions(rows: readonly InventoryRow[], settings: Settings, spellings: Readonly<Record<string, string>> = {}): Suggestion[] {
  const out: Suggestion[] = firstUseSuggestions(rows, settings);
  // A first mention that receives a moved date needs no new one.
  const receiving = new Set(out.filter((s) => s.kind === "move").map((s) => s.row.id));
  for (const row of rows) {
    const person = row.person;
    if (!person) continue;
    if (row.status === "spelling") {
      const spelling = spellings[personKey(person)] ?? sourceName(person, nameQuery(row.text));
      const value = spelling ? withEnding(row.text, spelling) : "";
      if (value && value !== row.text) out.push({ id: `s${row.id}`, kind: "spelling", row, person, value, preselected: !row.confirmed });
    }
    // A written date that the catalog contradicts: offered, never preselected (another source may be meant).
    if (row.date && (row.dateCheck === "different" || row.dateCheck === "inconsistent") && isValidDate(person.deathDate)) {
      const value = formatDate(person.deathDate, settings.prefix);
      if (value !== row.date) out.push({ id: `f${row.id}`, kind: "fixdate", row, person, value, old: row.date, preselected: false });
    }
    const wantsDate = settings.scope === "all" || (settings.scope === "first" && row.first);
    if (wantsDate && !row.date && isValidDate(person.deathDate) && !receiving.has(row.id)) {
      out.push({ id: `d${row.id}`, kind: "date", row, person, value: formatDate(person.deathDate, settings.prefix), preselected: true });
    }
  }
  return out;
}

export function describeSuggestion(s: Suggestion): string {
  if (s.kind === "fixdate") return `${s.row.text} ${s.old} → ${s.row.text} ${s.value}`;
  if (s.kind === "move") return `${s.row.text} → ${s.row.text} ${s.value} (sonraki "${s.from?.text} ${s.old}" geçişinden taşınır)`;
  if (s.kind === "drop") return `${s.row.text} ${s.old} → ${s.row.text} (ilk geçişte tarih zaten var)`;
  return s.kind === "date" ? `${s.row.text} → ${s.row.text} ${s.value}` : `${s.row.text} → ${s.value}`;
}

/**
 * Order to apply in: last position first, so the changes do not shift the
 * mentions still to be found; a date goes in before its name is respelled.
 */
export function applyOrder(chosen: readonly Suggestion[]): Suggestion[] {
  // Last position first, so changes never shift mentions still to be found;
  // a move counts at its later end (the date taken away). Dates sit after the
  // name: change them before the name is respelled.
  const at = (s: Suggestion) => Math.max(s.row.id, s.from?.id ?? -1);
  const rank = (s: Suggestion) => (s.kind === "spelling" ? 1 : 0);
  return [...chosen].sort((a, b) => at(b) - at(a) || rank(a) - rank(b));
}

// ---- Denetime devam et ----

export interface AuditBookmark {
  readonly at: number;
  readonly key: string;
  readonly total: number;
}

export function rowKey(row: InventoryRow, segments: readonly Segment[]): string {
  const text = segments[row.segment]?.text ?? "";
  return mentionKey(row.text, text.slice(0, row.start), text.slice(row.start + row.length));
}

export function parseBookmark(raw: unknown): AuditBookmark | null {
  if (typeof raw !== "string") return null;
  try {
    const value = JSON.parse(raw) as Partial<AuditBookmark>;
    if (typeof value.at === "number" && typeof value.key === "string" && typeof value.total === "number") return { at: value.at, key: value.key, total: value.total };
  } catch {
    /* ignored */
  }
  return null;
}

/** Where a paused audit continues: the same mention if it is still there, otherwise the same position. */
export function resumeIndex(audit: readonly InventoryRow[], segments: readonly Segment[], bookmark: AuditBookmark): number {
  if (audit.length === 0) return -1;
  const same = audit.findIndex((row) => rowKey(row, segments) === bookmark.key);
  if (same >= 0) return same;
  return Math.min(Math.max(bookmark.at, 0), audit.length - 1);
}

// ---- İşlem geçmişi ----

export type HistoryKind = "date" | "spelling" | "repair" | "index" | "cleanup" | "bulk";

export interface HistoryEntry {
  readonly time: string;
  readonly kind: HistoryKind;
  readonly text: string;
  /** Date content control to go to, when there is one. */
  readonly control?: number;
}

export const HISTORY_LIMIT = 200;

export function addHistory(list: readonly HistoryEntry[], entry: Omit<HistoryEntry, "time">, now = new Date()): HistoryEntry[] {
  const time = now.toLocaleTimeString("tr-TR", { hour: "2-digit", minute: "2-digit" });
  return [{ time, ...entry }, ...list].slice(0, HISTORY_LIMIT);
}

export function parseHistory(raw: string | null): HistoryEntry[] {
  if (!raw) return [];
  try {
    const value = JSON.parse(raw) as unknown;
    return Array.isArray(value) ? (value.filter((e) => e && typeof e.text === "string" && typeof e.kind === "string") as HistoryEntry[]).slice(0, HISTORY_LIMIT) : [];
  } catch {
    return [];
  }
}

export const HISTORY_LABEL: Record<HistoryKind, string> = {
  date: "Tarih eklendi",
  spelling: "Yazım değişti",
  repair: "Tarih onarıldı",
  index: "Dizin",
  cleanup: "Teslim temizliği",
  bulk: "Toplu düzeltme",
};

export function personLabel(person: Person): string {
  return textName(person.shortName);
}
