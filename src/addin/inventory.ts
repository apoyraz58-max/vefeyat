/**
 * Şahıs dökümü: every person mention in the main text and footnotes, with its
 * catalog match, the date written after it and a status. Word-independent;
 * word.ts supplies the paragraph texts. Follows VBA 0.3.41 `VfjPrepare` and
 * `VfjRefreshItem`.
 */
import {
  type Catalog,
  type NameHit,
  NameIndex,
  type Person,
  alreadyHasDate,
  citationSpans,
  containsSpan,
  findMatches,
  isValidDate,
  nameQuery,
  personKey,
  singleSafe,
  sourceName,
  startsWithDate,
  textName,
} from "../core/index.js";

/** A paragraph of the main text or of a footnote, in reading order. */
export interface Segment {
  readonly part: "main" | "note";
  /** Main-text paragraph index (for a footnote: the paragraph holding its reference). */
  readonly anchor: number;
  /** Position inside the anchor: 0 for the main paragraph, 1.. for its footnotes. */
  readonly order: number;
  readonly text: string;
}

export type Status = "match" | "spelling" | "near" | "choose" | "unknown";

export interface InventoryRow {
  readonly id: number;
  readonly segment: number;
  readonly part: "main" | "note";
  /** 0-based offset in the segment text. */
  readonly start: number;
  readonly length: number;
  /** Literal text, including a Turkish ending. */
  readonly text: string;
  readonly candidates: readonly Person[];
  /** Set when the identity is certain (single strong match). */
  readonly person: Person | null;
  /** Date already written right after the name, or "". */
  readonly date: string;
  readonly status: Status;
  /** First certain mention of this person in reading order. */
  readonly first: boolean;
}

/** Date text right after a name: "(ö. 505/1111)" etc., or "". */
export function followingDate(after: string): string {
  const trimmed = after.replace(/^[  ]+/, "");
  if (!startsWithDate(trimmed)) return "";
  const closing = trimmed.indexOf(")");
  return closing > 0 && closing < 80 ? trimmed.slice(0, closing + 1) : "";
}

/** Whether the selected spelling already is the catalog spelling (`VfjSpellingMatches`). */
export function spellingMatches(person: Person, literal: string): boolean {
  const bare = nameQuery(literal);
  if ([person.shortName, person.fullName, person.originalName].some((name) => name === bare || textName(name) === bare)) return true;
  return sourceName(person, bare) === bare;
}

export function buildInventory(segments: readonly Segment[], catalog: Catalog, index = new NameIndex(catalog.people)): InventoryRow[] {
  const ordered = segments
    .map((segment, i) => ({ segment, i }))
    .sort((a, b) => a.segment.anchor - b.segment.anchor || a.segment.order - b.segment.order);
  const rows: InventoryRow[] = [];
  const seen = new Set<string>();
  for (const { segment, i } of ordered) {
    const hits = findMatches(segment.text, index, true, true);
    // Footnotes: a name inside an author–title reference is a citation.
    const spans =
      segment.part === "note" ? citationSpans(segment.text, hits.map((h): NameHit => [h.start, h.length, h.text])) : [];
    for (const hit of hits) {
      if (containsSpan(spans, hit.start, hit.length)) continue;
      const candidates = index.rank(hit.key, nameQuery(hit.text));
      const person = singleSafe(candidates, hit.text) ? (candidates[0] as Person) : null;
      const after = segment.text.slice(hit.start + hit.length);
      let date = followingDate(after);
      if (!date && person && alreadyHasDate(after, person.deathDate)) date = person.deathDate;
      let status: Status;
      if (candidates.length === 0) status = "unknown";
      else if (person) status = spellingMatches(person, hit.text) ? "match" : "spelling";
      else if (candidates.length === 1) status = "near";
      else status = "choose";
      let first = false;
      if (person) {
        const key = personKey(person);
        first = !seen.has(key);
        seen.add(key);
      }
      rows.push({ id: rows.length, segment: i, part: segment.part, start: hit.start, length: hit.length, text: hit.text, candidates, person, date, status, first });
    }
  }
  return rows;
}

/** Catalog column text for a row. */
export function catalogLabel(row: InventoryRow): string {
  if (row.person) return textName(row.person.shortName);
  if (row.candidates.length === 1) return `${textName((row.candidates[0] as Person).shortName)} (?)`;
  if (row.candidates.length > 1) return `${row.candidates.length} aday; seçim gerekli`;
  return "Katalogda yok";
}

/** Status text for a row, with date hints as in the VBA inventory. */
export function statusLabel(row: InventoryRow): string {
  const base: Record<Status, string> = {
    match: "Kaynak yazımıyla uyumlu",
    spelling: "Uyarı: kaynak yazımı farklı",
    near: "Yakın eşleşme; doğrulayın",
    choose: "Birden fazla şahıs eşleşiyor",
    unknown: "Katalogda yok",
  };
  let text = base[row.status];
  if (row.person && !isValidDate(row.person.deathDate)) text += " · Kaynakta tarih yok";
  else if (row.person && row.first && !row.date) text += " · İlk geçişte tarih yok";
  return text;
}

export interface InventorySummary {
  readonly mentions: number;
  readonly people: number;
  readonly firstWithoutDate: number;
  readonly spelling: number;
  readonly uncertain: number;
}

export function summarize(rows: readonly InventoryRow[]): InventorySummary {
  const people = new Set(rows.filter((r) => r.person).map((r) => personKey(r.person as Person)));
  return {
    mentions: rows.length,
    people: people.size,
    firstWithoutDate: rows.filter((r) => r.first && !r.date && r.person && isValidDate(r.person.deathDate)).length,
    spelling: rows.filter((r) => r.status === "spelling").length,
    uncertain: rows.filter((r) => r.status === "near" || r.status === "choose").length,
  };
}
