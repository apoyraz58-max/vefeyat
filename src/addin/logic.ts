/**
 * Task pane decisions that do not need Word. Word calls live in word.ts so
 * that everything here is unit tested.
 */
import {
  type Catalog,
  type Person,
  type PrefixKey,
  alreadyHasDate,
  confirmationCandidates,
  formatDate,
  isValidAlias,
  isValidDate,
  nameCandidates,
  nameQuery,
  normalize,
  personKey,
  singleSafe,
  startsWithDate,
  strength,
  textName,
} from "../core/index.js";

export interface Lookup {
  /** Selected name without a Turkish ending. */
  readonly query: string;
  /** Strong candidates (same name), then partial ones. */
  readonly candidates: readonly Person[];
  readonly strong: ReadonlySet<string>;
  /** Best name score per person key: 4/3 same name, 2 part of the name, 1 similar spelling. */
  readonly scores: ReadonlyMap<string, number>;
  /** Set when exactly one strong candidate exists: no choice needed. */
  readonly automatic: Person | null;
}

/**
 * Candidates for a selected name, as the VBA "Ölüm tarihi ekle" command
 * gathers them: name matches only (no description matches), one entry per
 * person, strong matches first.
 */
export function lookup(catalog: Catalog, alias: string): Lookup {
  const found = confirmationCandidates(catalog.find(alias, false), alias);
  const { exact, partial } = nameCandidates(found, alias);
  const strong = new Set(exact.map(personKey));
  return {
    query: nameQuery(alias),
    candidates: [...exact, ...partial],
    strong,
    scores: new Map(found.map((person) => [personKey(person), strength(person, alias)])),
    automatic: singleSafe(found, alias) ? (found[0] as Person) : null,
  };
}

export type DatePlan =
  | { readonly ok: true; readonly text: string }
  | { readonly ok: false; readonly reason: string };

/**
 * Whether a date may be written after the name, and its text. `following` is
 * the rest of the paragraph after the selected name.
 */
export function planDate(person: Person, prefix: PrefixKey, following: string): DatePlan {
  if (!isValidDate(person.deathDate)) {
    return { ok: false, reason: "Katalogda bu şahıs için kullanılabilir bir ölüm tarihi yok; tarih eklenmedi." };
  }
  if (startsWithDate(following) || alreadyHasDate(following, person.deathDate)) {
    return { ok: false, reason: "Adın hemen arkasında zaten bir tarih var; mevcut tarihe dokunulmadı." };
  }
  return { ok: true, text: formatDate(person.deathDate, prefix) };
}

/**
 * Literal spellings searched before the selection to warn about an earlier
 * mention: the selected spelling, the three catalog names, their reading
 * order and the family part of the short name ("Râzî" of "Fahreddin er-Râzî").
 * Very short or invalid values are skipped; at most 12 spellings.
 */
export function earlierSpellings(person: Person, alias: string): string[] {
  const result: string[] = [];
  const seen = new Set<string>();
  const add = (value: string): void => {
    value = value.trim();
    const key = normalize(value);
    if (value.length < 3 || value.length > 200 || key.length < 3 || !isValidAlias(value) || seen.has(value)) return;
    seen.add(value);
    result.push(value);
  };
  add(nameQuery(alias));
  for (const name of [person.shortName, person.fullName, person.originalName]) {
    add(name.includes(",") ? textName(name) : name);
  }
  const short = person.shortName.includes(",")
    ? person.shortName.slice(0, person.shortName.indexOf(","))
    : (person.shortName.trim().split(/[\s\-–]+/).pop() ?? "");
  add(short);
  return result.slice(0, 12);
}

/** Content control tag of a Vefeyât date: short, stable per person. */
export function dateTag(person: Person): string {
  return `vefeyat:date:${hash(personKey(person))}`;
}

function hash(value: string): string {
  // FNV-1a 32-bit; only used to group dates of one person, not for security.
  let h = 0x811c9dc5;
  for (let i = 0; i < value.length; i++) {
    h ^= value.charCodeAt(i);
    h = Math.imul(h, 0x01000193) >>> 0;
  }
  return h.toString(16).padStart(8, "0");
}

/** Short note explaining a non-exact candidate, or "" for the same name. */
export function matchHint(score: number): string {
  if (score >= 3) return "";
  if (score === 2) return "adın bir parçası";
  return "benzer yazım, doğrulayın";
}

/** One-line summary of a candidate for the list. */
export function candidateLabel(person: Person): string {
  const name = textName(person.shortName);
  return person.deathDate ? `${name} ${person.deathDate}` : name;
}
