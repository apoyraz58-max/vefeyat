/**
 * Toplu düzeltme önizlemesi, Denetime devam et and İşlem geçmişi:
 * Word-independent helpers (VBA 0.3.41 `VfbPreview`, `VfwResume`, `VfgLog`).
 */
import { type Person, formatDate, isValidDate, nameQuery, sourceName, textName, withEnding } from "../core/index.js";
import type { InventoryRow, Segment } from "./inventory.js";
import { type Settings, mentionKey } from "./logic.js";

export interface Suggestion {
  readonly id: string;
  readonly kind: "date" | "spelling";
  readonly row: InventoryRow;
  readonly person: Person;
  /** Date to insert, or the replacement spelling (with the Turkish ending). */
  readonly value: string;
  /** Checked by default: safe changes only. */
  readonly preselected: boolean;
}

/**
 * Spelling fixes for certain mentions written differently from the catalog,
 * and dates for first mentions without one (unless Tarih ekleme is
 * "Yalnız denetle"; with "Bütün kullanımlar", every undated certain mention).
 */
export function bulkSuggestions(rows: readonly InventoryRow[], settings: Settings): Suggestion[] {
  const out: Suggestion[] = [];
  for (const row of rows) {
    const person = row.person;
    if (!person) continue;
    if (row.status === "spelling") {
      const spelling = sourceName(person, nameQuery(row.text));
      const value = spelling ? withEnding(row.text, spelling) : "";
      if (value && value !== row.text) out.push({ id: `s${row.id}`, kind: "spelling", row, person, value, preselected: !row.confirmed });
    }
    const wantsDate = settings.scope === "all" || (settings.scope === "first" && row.first);
    if (wantsDate && !row.date && isValidDate(person.deathDate)) {
      out.push({ id: `d${row.id}`, kind: "date", row, person, value: formatDate(person.deathDate, settings.prefix), preselected: true });
    }
  }
  return out;
}

export function describeSuggestion(s: Suggestion): string {
  return s.kind === "date" ? `${s.row.text} → ${s.row.text} ${s.value}` : `${s.row.text} → ${s.value}`;
}

/**
 * Order to apply in: last position first, so the changes do not shift the
 * mentions still to be found; a date goes in before its name is respelled.
 */
export function applyOrder(chosen: readonly Suggestion[]): Suggestion[] {
  return [...chosen].sort((a, b) => b.row.id - a.row.id || (a.kind === "date" ? -1 : 1));
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
