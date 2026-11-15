/**
 * Belge sağlığı and Tarihin kaynağı: checks Vefeyât's date content controls
 * (tag vefeyat:date:<hash>) against the catalog. Word-independent; port of
 * VBA 0.3.41 `VfhCheck`. Nothing is repaired without the user's choice.
 */
import { type Catalog, type NameIndex, type Person, PREFIX_KEYS, type PrefixKey, alreadyHasDate, findMatches, formatDate, isValidDate, nameQuery, safeCandidate, startsWithDate } from "../core/index.js";
import { dateTag } from "./logic.js";

export const DATE_TAG_PREFIX = "vefeyat:date:";

/** Catalog person of a date tag, or null when the catalog has no such record. */
export class TagIndex {
  private readonly map = new Map<string, Person>();
  constructor(catalog: Catalog) {
    for (const person of catalog.people) this.map.set(dateTag(person), person);
  }
  person(tag: string): Person | null {
    return this.map.get(tag) ?? null;
  }
}

export type Health = "ok" | "prefix" | "changed" | "empty" | "broken" | "unknown";

export interface DateControl {
  readonly id: number;
  readonly part: "main" | "note";
  readonly tag: string;
  readonly text: string;
}

export interface HealthRow extends DateControl {
  readonly person: Person | null;
  readonly health: Health;
  /** Text the control should hold with the current prefix setting ("" if unknown). */
  readonly expected: string;
}

const squash = (s: string) => s.replace(/[  \t]+/g, " ").trim();

export function checkDate(text: string, person: Person | null, prefix: PrefixKey): { health: Health; expected: string } {
  if (!person || !isValidDate(person.deathDate)) return { health: "unknown", expected: "" };
  const expected = formatDate(person.deathDate, prefix);
  const value = squash(text);
  if (value.length === 0) return { health: "empty", expected };
  if (value === squash(expected)) return { health: "ok", expected };
  if (PREFIX_KEYS.some((k) => value === squash(formatDate(person.deathDate, k)))) return { health: "prefix", expected };
  if (startsWithDate(value)) return { health: "changed", expected };
  return { health: "broken", expected };
}

export function checkControls(controls: readonly DateControl[], tags: TagIndex, prefix: PrefixKey): HealthRow[] {
  return controls.map((c) => {
    const person = tags.person(c.tag);
    return { ...c, person, ...checkDate(c.text, person, prefix) };
  });
}

export const HEALTH_LABEL: Record<Health, string> = {
  ok: "Sağlam",
  prefix: "Önek ayardan farklı",
  changed: "Tarih katalogdan farklı",
  empty: "Tarih silinmiş",
  broken: "Tarih bozulmuş",
  unknown: "Katalogda kaydı yok",
};

/** Rows that a repair can rewrite with the expected text. */
export function repairable(row: HealthRow): boolean {
  return row.expected.length > 0 && row.health !== "ok";
}

/** Text written by a repair: the leading space insertDate uses, then the date. */
export function repairText(row: HealthRow, original: string): string {
  return /^[  ]/.test(original) || original.length === 0 ? ` ${row.expected}` : row.expected;
}

// ---- VBA belgeleri: ADDIN Vefeyat.Date alanları ----

/** Field code of a date written by the VBA add-in (0.3.x). */
export function isLegacyDateField(code: string): boolean {
  return /^\s*ADDIN\s+Vefeyat\.Date\.[A-Za-z0-9]+(\.[A-Za-z0-9]+)?\s*$/i.test(code);
}

export interface LegacyResolution {
  readonly person: Person | null;
  /** The name right before the date, as written ("Gazzâlî'nin" etc.). */
  readonly literal: string;
  /** 0-based start of that name in the paragraph text before the date. */
  readonly start: number;
}

/**
 * The person a VBA date belongs to: the catalog name ending right before it
 * (`before` is the paragraph text up to the date). Among homonyms the one
 * whose catalog date the field shows is taken; otherwise none.
 */
export function legacyPerson(before: string, dateText: string, index: NameIndex): LegacyResolution | null {
  const hits = findMatches(before, index, false);
  const last = hits[hits.length - 1];
  if (!last || before.slice(last.start + last.length).trim().length > 1) return null;
  const candidates = index.rank(last.key, nameQuery(last.text));
  const dated = candidates.filter((c) => isValidDate(c.deathDate) && alreadyHasDate(dateText, c.deathDate));
  const person = dated.length === 1 ? (dated[0] as Person) : safeCandidate(candidates, last.text);
  return { person, literal: last.text, start: last.start };
}
