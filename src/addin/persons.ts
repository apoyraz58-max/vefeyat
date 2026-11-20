/**
 * Kendi şahıslarım: people the DİA catalog lacks (modern authors, scholars
 * without an article), kept in the document and used like catalog entries.
 */
import { Catalog, type Person, deathYears, isValidAlias, normalize } from "../core/index.js";

export const LOCAL_SOURCE = "vefeyat:yerel:";
export const LOCAL_LIMIT = 500;

export interface LocalPerson {
  readonly shortName: string;
  readonly fullName: string;
  readonly deathDate: string;
  readonly description: string;
}

export function isLocal(person: Pick<Person, "source">): boolean {
  return person.source.startsWith(LOCAL_SOURCE);
}

/** A stable source for the record: its normalized short name. */
export function localSource(shortName: string): string {
  return LOCAL_SOURCE + normalize(shortName).replace(/ /g, "-");
}

export function toPerson(local: LocalPerson): Person {
  return { shortName: local.shortName, fullName: local.fullName, originalName: "", deathDate: local.deathDate, description: local.description, source: localSource(local.shortName) };
}

/**
 * The date as Vefeyât writes it: "1950" → "(ö. 1950)", "1890-1950" →
 * "(1890-1950)", "ö. 505/1111" → "(ö. 505/1111)"; a date already in
 * parentheses is kept. Null when no year can be read.
 */
export function normalizeDate(input: string): string | null {
  const value = input.trim().replace(/\s+/g, " ");
  if (!value) return "";
  if (!deathYears(value)) return null;
  if (value.startsWith("(") && value.endsWith(")")) return value;
  if (/^(ö|öl|v|d|m\.\s?ö)\./i.test(value)) return `(${value})`;
  if (/^\d{3,4}\s*-\s*\d{2,4}$/.test(value)) return `(${value.replace(/\s/g, "")})`;
  return `(ö. ${value})`;
}

/** Checks one record; returns the cleaned record or the reason it is refused. */
export function validateLocal(input: LocalPerson): { ok: true; person: LocalPerson } | { ok: false; reason: string } {
  const shortName = input.shortName.trim().replace(/\s+/g, " ");
  if (!isValidAlias(shortName) || shortName.length < 2) return { ok: false, reason: "Kısa ad geçerli değil (en az iki harf, tek satır)." };
  const deathDate = normalizeDate(input.deathDate);
  if (deathDate === null) return { ok: false, reason: "Ölüm tarihinde yıl okunamadı. Örnek: 1950, ö. 505/1111 ya da 1890-1950." };
  return {
    ok: true,
    person: { shortName, fullName: input.fullName.trim().replace(/\s+/g, " "), deathDate, description: input.description.trim().slice(0, 300) },
  };
}

export function parseLocalPersons(stored: unknown): LocalPerson[] {
  if (typeof stored !== "string") return [];
  try {
    const value: unknown = JSON.parse(stored);
    if (!Array.isArray(value)) return [];
    const out: LocalPerson[] = [];
    for (const item of value.slice(0, LOCAL_LIMIT)) {
      if (!item || typeof item !== "object") continue;
      const r = item as Record<string, unknown>;
      const checked = validateLocal({
        shortName: String(r.shortName ?? ""),
        fullName: String(r.fullName ?? ""),
        deathDate: String(r.deathDate ?? ""),
        description: String(r.description ?? ""),
      });
      if (checked.ok) out.push(checked.person);
    }
    return out;
  } catch {
    return [];
  }
}

/** Adds or replaces (same short name) a record. */
export function upsertLocal(list: readonly LocalPerson[], person: LocalPerson): LocalPerson[] {
  const key = localSource(person.shortName);
  return [...list.filter((p) => localSource(p.shortName) !== key), person].slice(-LOCAL_LIMIT);
}

/** The catalog with the document's own people added. */
export function withLocal(base: Catalog, locals: readonly LocalPerson[]): Catalog {
  return locals.length === 0 ? base : Catalog.combine([base], locals.map(toPerson));
}
