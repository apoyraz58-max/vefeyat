/**
 * Spelling suggestions for a selected name. Ported from `VfnNameOptions`,
 * `VfnSourceName`, `VfnSpellingScore` and `VfnNameBoundary` (VBA 0.3.41).
 */
import { MAX_NAME_LENGTH, isSpace, isValidAlias } from "./alias.js";
import type { Person } from "./catalog.js";
import { nameQuery, nearWord, textName } from "./names.js";
import { collapseLetters, normalize } from "./normalize.js";
import { vbTrim } from "./text.js";

export interface NameOption {
  /** Spelling to write into the document. */
  readonly value: string;
  /** Turkish label shown to the user. */
  readonly label: string;
}

function isNameBoundary(ch: string): boolean {
  return isSpace(ch) || ch === "-" || ch === "–" || ch === "—";
}

/** `VfnSpellingScore`: 4 same name, 3 same without spaces, 2 same with collapsed letters, 1 near. */
export function spellingScore(candidate: string, query: string, key: string): number {
  const normalized = normalize(candidate);
  const compact = normalized.split(" ").join("");
  if (normalized === normalize(query)) return 4;
  if (compact === key) return 3;
  if (collapseLetters(compact) === collapseLetters(key)) return 2;
  if (key.length >= 5 && nearWord(compact, key, key.length >= 9 ? 2 : 1)) return 1;
  return 0;
}

/**
 * `VfnSourceName`: the catalog spelling of the part of a source name that the
 * selection refers to ("Gazali" -> "Gazzâlî", "Kıvamussunne" -> "Kıvâmüssünne"
 * inside a longer full name). Empty when nothing is close enough.
 */
export function sourceName(person: Person, query: string): string {
  const key = normalize(query).split(" ").join("");
  if (key.length === 0) return "";
  const raw = [person.shortName, person.fullName, person.originalName];
  let best = 0;
  let result = "";
  for (let i = 0; i < 3; i++) {
    let name = raw[i] as string;
    if (i < 2) {
      name = textName(name);
      if (name !== raw[i] && normalize(query) === normalize(raw[i] as string) && best < 4) {
        best = 4;
        result = name;
      }
    }
    if (name.length > 2000) continue;
    for (let first = 0; first < name.length; first++) {
      const startsWord = first === 0 || isNameBoundary(name[first - 1] as string);
      if (!startsWord || isNameBoundary(name[first] as string)) continue;
      for (let last = first; last < name.length; last++) {
        if (last - first + 1 > MAX_NAME_LENGTH) break;
        const endsWord = last === name.length - 1 || isNameBoundary(name[last + 1] as string);
        if (!endsWord) continue;
        const candidate = name.slice(first, last + 1);
        const score = spellingScore(candidate, query, key);
        if (score > best) {
          best = score;
          result = candidate;
        }
      }
    }
  }
  return result;
}

/**
 * `VfnNameOptions`: distinct spellings to offer for a selected name, closest
 * first. `query` is the selected name; a Turkish ending is ignored.
 */
export function nameOptions(person: Person, query: string): NameOption[] {
  const bare = nameQuery(query);
  const options: NameOption[] = [];
  const seen = new Set<string>();
  const add = (value: string, label: string): void => {
    if (value === "---" || !isValidAlias(value) || seen.has(value)) return;
    seen.add(value);
    options.push({ value, label });
  };
  add(sourceName(person, bare), "Seçili ada yakın yazım");
  add(textName(vbTrim(person.shortName)), "Kısa ad");
  add(textName(vbTrim(person.fullName)), "Tam / uzun ad");
  add(vbTrim(person.originalName), "Orijinal ad");
  return options;
}

/**
 * Replacement text for a selection: the chosen spelling followed by the
 * selection's Turkish ending ("Gazali'nin" + "Gazzâlî" -> "Gazzâlî'nin").
 */
export function withEnding(selected: string, spelling: string): string {
  return spelling + selected.slice(nameQuery(selected).length);
}
