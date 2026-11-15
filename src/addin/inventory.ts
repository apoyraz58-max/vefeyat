/**
 * Şahıs dökümü: every person mention in the main text and footnotes, with its
 * catalog match, the date written after it and a status. Word-independent;
 * word.ts supplies the paragraph texts. Follows VBA 0.3.41 `VfjPrepare` and
 * `VfjRefreshItem`.
 */
import {
  type Catalog,
  type Match,
  type NameHit,
  NameIndex,
  type Person,
  alreadyHasDate,
  bibliographyFlags,
  citationSpans,
  containsSpan,
  findMatches,
  findUnknownPeople,
  isValidDate,
  nameQuery,
  personKey,
  singleSafe,
  sourceName,
  startsWithDate,
  textName,
} from "../core/index.js";
import { type Confirmations, confirmedPerson, exclusionKey, mentionKey } from "./logic.js";

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
  /** Identity taken from a confirmation saved in the document. */
  readonly confirmed: boolean;
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

/**
 * Person mentions of one paragraph. Main text: strict matches with context
 * rules. Footnotes (VBA test harness `FootnotePeople`): citation islands are
 * computed from all catalog spellings and strict matches inside them dropped.
 */
export function notePeopleOrMatches(segment: Pick<Segment, "part" | "text">, index: NameIndex): Match[] {
  // Citation islands ("Gazzâlî, el-Mustasfâ, I, 45", "(Gazzâlî, 1993: 45)")
  // are left out in both parts; the rest of the paragraph or footnote is read.
  const names = findMatches(segment.text, index, false).map((h): NameHit => [h.start, h.length, h.text]);
  let spans = citationSpans(segment.text, names);
  if (segment.part === "main") {
    // Main text narrates more than it cites: only islands with a locator
    // (page, volume, year) count, and "Gazzâlî (1993: 45) vurgular" keeps
    // its subject.
    spans = spans.filter(([a, b]) => {
      const island = segment.text.slice(a, b);
      return /\d/.test(island) && !/^[^,(]+\(\s*\d{4}/.test(island);
    });
  }
  return findMatches(segment.text, index, true, segment.part !== "note").filter((hit) => !containsSpan(spans, hit.start, hit.length));
}

export function buildInventory(
  segments: readonly Segment[],
  catalog: Catalog,
  index = new NameIndex(catalog.people),
  confirmations: Confirmations = {},
  exclusions: readonly string[] = [],
  spellings: Readonly<Record<string, string>> = {},
): InventoryRow[] {
  const excluded = new Set(exclusions);
  const ordered = segments
    .map((segment, i) => ({ segment, i }))
    .sort((a, b) => a.segment.anchor - b.segment.anchor || a.segment.order - b.segment.order);
  const rows: InventoryRow[] = [];
  const seen = new Set<string>();
  const mains = ordered.filter((o) => o.segment.part === "main");
  const flags = bibliographyFlags(mains.map((o) => o.segment.text));
  const bibliography = new Set(mains.filter((_, k) => flags[k]).map((o) => o.segment.anchor));
  for (const { segment, i } of ordered) {
    if (bibliography.has(segment.anchor)) continue;
    const known = notePeopleOrMatches(segment, index);
    const unknown = findUnknownPeople(segment.text, index).filter(
      (u) => !known.some((k) => u.start < k.start + k.length && k.start < u.start + u.length),
    );
    const hits = [...known, ...unknown].sort((a, b) => a.start - b.start);
    for (const hit of hits) {
      if (excluded.size > 0 && excluded.has(exclusionKey(hit.text))) continue;
      const candidates = hit.key ? index.rank(hit.key, nameQuery(hit.text)) : [];
      let person = singleSafe(candidates, hit.text) ? (candidates[0] as Person) : null;
      let confirmed = false;
      if (!person) {
        const key = mentionKey(hit.text, segment.text.slice(0, hit.start), segment.text.slice(hit.start + hit.length));
        person = confirmedPerson(confirmations, key, candidates);
        confirmed = person !== null;
      }
      const after = segment.text.slice(hit.start + hit.length);
      let date = followingDate(after);
      if (!date && person && alreadyHasDate(after, person.deathDate)) date = person.deathDate;
      let status: Status;
      if (candidates.length === 0) status = "unknown";
      else if (person) {
        const preferred = spellings[personKey(person)];
        const ok = preferred ? nameQuery(hit.text) === preferred : spellingMatches(person, hit.text);
        status = ok ? "match" : "spelling";
      }
      else if (candidates.length === 1) status = "near";
      else status = "choose";
      let first = false;
      if (person) {
        const key = personKey(person);
        first = !seen.has(key);
        seen.add(key);
      }
      rows.push({ id: rows.length, segment: i, part: segment.part, start: hit.start, length: hit.length, text: hit.text, candidates, person, date, status, first, confirmed });
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
  if (row.confirmed) text += " · onaylı kimlik";
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
