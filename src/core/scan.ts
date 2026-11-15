/**
 * Name scanner: builds a lookup of catalog spellings and finds their
 * occurrences in running text, including Turkish case endings
 * ("Gazzâlî'nin") and multi-word names ("Fahreddin er-Râzî").
 * Ported from the VBA 0.3.41 `Vfb*` procedures.
 */
import { MAX_NAME_LENGTH } from "./alias.js";
import type { Person } from "./catalog.js";
import { personKey } from "./catalog.js";
import { excludeMention, hasSuffix, schoolRoot } from "./context.js";
import { nameQuery, nameScore, prominence } from "./names.js";
import { collapseLetters, isSeparator, nameKey, normalize } from "./normalize.js";
import { vbLower, vbTrim } from "./text.js";
import { mid, upper } from "./vb.js";

/** A token: lookup key, 0-based start and length in the source text. */
export interface Token {
  readonly key: string;
  readonly start: number;
  readonly length: number;
}

/** `VfbBreak`: token separator. Hyphens and apostrophes stay inside names. */
export function isBreak(ch: string): boolean {
  if (["-", "'", "’", "‘", "ʻ", "ʼ", "ʿ"].includes(ch)) return false;
  if (ch.length === 0) return true;
  let code = ch.charCodeAt(0);
  if (code >= 65 && code <= 90) code += 32;
  return isSeparator(code);
}

/** `VfbTokens`. */
export function tokenize(text: string, cache: Map<string, string> = new Map()): Token[] {
  const result: Token[] = [];
  let first = -1;
  for (let i = 0; i <= text.length; i++) {
    const ch = i < text.length ? (text[i] as string) : " ";
    if (isBreak(ch)) {
      if (first >= 0) {
        const value = text.slice(first, i);
        let key = cache.get(value);
        if (key === undefined) {
          key = nameKey(value);
          cache.set(value, key);
        }
        if (key.length > 0) result.push({ key, start: first, length: i - first });
        first = -1;
      }
    } else if (first < 0) {
      first = i;
    }
  }
  return result;
}

const NOISE = new Set("ibn|ibni|bin|bint|ebu|ebul|el|er|es|et|ez|en|ed|bey|han|sah|seyh|imam|haci|seyyid|sair".split("|"));
const COMMON = new Set(
  ("tarif|tanim|ilim|bilim|isim|ad|bilgi|tarih|eser|kitap|usul|fikih|hukuk|hadis|tefsir|kelam|felsefe|mantik|delil|akil|nazar|mana|anlam|hak|" +
    "hakikat|adalet|hayat|olum|dogum|zaman|yer|islam|kuran|sunet|tasavuf|ahlak|fikir|dusunce|mesele|bahis|bab|fasl|fasil|cilt|sayfa|baski|yayin|tercume|" +
    "serh|hasiye|kavram|terim|kisi|sahis|yazar|arastirma|metin|deneme|konu|once|sonra|gore|icin|ile|ve|veya|yahut|ama|fakat|bu|su|bir|" +
    "iki|uc|her|olan|olarak|oldu|dedi|der|gibi|uzere|hakinda").split("|"),
);
const NAME_LINKS = new Set("b|ibn|ibni|bin|bint|ebu|ebul|el|er|es|et|ez|en|ed".split("|"));

export const isNoise = (key: string): boolean => NOISE.has(key);
export const isCommonWord = (key: string): boolean => COMMON.has(key);
export const isNameLink = (key: string): boolean => NAME_LINKS.has(key);

/** `VfbWordCount`. */
export function wordCount(value: string): number {
  let count = 0;
  let inWord = false;
  for (const ch of value) {
    if (isBreak(ch)) inWord = false;
    else if (!inWord) {
      count++;
      inWord = true;
    }
  }
  return count;
}

/** `VfbInitialUpper`: the first cased letter is a capital (uncased scripts count as capital). */
export function initialUpper(value: string): boolean {
  if (["el-", "er-", "es-", "et-", "ez-", "en-", "ed-"].includes(vbLower(value.slice(0, 3)))) value = value.slice(3);
  for (const ch of value) {
    if (upper(ch) !== vbLower(ch)) return ch === upper(ch);
    const code = ch.charCodeAt(0);
    if ((code >= 0x590 && code <= 0x1fff) || code >= 0x2e80) return true;
  }
  return false;
}

/** `VfbWithoutArticles`: normalized name without Arabic article tokens. */
export function withoutArticles(value: string): string {
  let result = ` ${normalize(value)} `;
  for (const article of ["el", "er", "es", "et", "ez", "en", "ed"]) result = result.split(` ${article} `).join(" ");
  return vbTrim(result);
}

/** `VfbNaturalOrder`: "Râzî, Fahreddin" -> "Fahreddin Râzî". */
export function naturalOrder(value: string): string {
  const comma = value.indexOf(",");
  if (comma < 0 || value.indexOf(",", comma + 1) >= 0) return "";
  const first = vbTrim(value.slice(comma + 1));
  const last = vbTrim(value.slice(0, comma));
  return first.length > 0 && last.length > 0 ? `${first} ${last}` : "";
}

function compoundAlias(value: string): boolean {
  if (!value.includes("-")) return false;
  if (!value.includes("'") && !value.includes("’")) return false;
  const key = nameKey(value);
  return !(key.startsWith("ebu") || key.startsWith("abu"));
}

function independentAlias(value: string, shortName: string): boolean {
  const key = nameKey(value);
  return schoolRoot(key).length > 0 ? key === nameKey(shortName) : true;
}

/** Lookup of catalog spellings (`VfbBuildPatterns`, `VfbIndexRow`, `VfbAddPattern`). */
export class NameIndex {
  /** Spelling key -> catalog row indices. */
  readonly patterns = new Map<string, Set<number>>();
  /** First three letters of every key, to stop scanning early. */
  readonly heads = new Set<string>();
  /** Spelling key -> word counts it was seen with (strict mode). */
  readonly shapes = new Map<string, Set<number>>();
  readonly cache = new Map<string, string>();

  constructor(readonly people: readonly Person[]) {
    people.forEach((person, i) => this.indexRow(person, i));
  }

  private add(value: string, index: number): void {
    let key = this.cache.get(value);
    if (key === undefined) {
      key = nameKey(value);
      this.cache.set(value, key);
    }
    if (key.length < 3 || value.length > MAX_NAME_LENGTH || value === "---") return;
    if (isNoise(key) || isCommonWord(key)) return;
    let ids = this.patterns.get(key);
    if (!ids) {
      ids = new Set();
      this.patterns.set(key, ids);
      this.heads.add(key.slice(0, 3));
    }
    ids.add(index);
    let counts = this.shapes.get(key);
    if (!counts) {
      counts = new Set();
      this.shapes.set(key, counts);
    }
    counts.add(wordCount(value));
  }

  private indexRow(person: Person, i: number): void {
    const values = [person.shortName, person.fullName, person.originalName];
    values.forEach((value, j) => {
      this.add(value, i);
      this.add(withoutArticles(value), i);
      if (j === 0) {
        const natural = naturalOrder(value);
        if (natural.length > 0) {
          this.add(natural, i);
          this.add(withoutArticles(natural), i);
        }
      }
      for (const token of tokenize(value, this.cache)) {
        const alias = value.slice(token.start, token.start + token.length);
        // Short catalog names supply independent aliases; components of long
        // names do not, except joined forms such as "Kıvâmü's-sünne".
        if ((j === 0 || compoundAlias(alias)) && independentAlias(alias, person.shortName)) this.add(alias, i);
      }
      if (j === 0) {
        for (const part of normalize(value).split(" ")) {
          if (independentAlias(part, person.shortName)) this.add(part, i);
        }
      }
    });
  }

  has(key: string): boolean {
    return this.patterns.has(key);
  }

  shapeAllowed(key: string, count: number): boolean {
    return this.shapes.get(key)?.has(count) ?? false;
  }

  /** `VfbRankRows`: catalog rows for a matched key, best first, one per person. */
  rank(key: string, query: string): Person[] {
    const ids = this.patterns.get(key);
    if (!ids) return [];
    const seen = new Set<string>();
    const ranked: Array<{ rank: number; person: Person }> = [];
    for (const index of ids) {
      const person = this.people[index] as Person;
      const id = personKey(person);
      if (seen.has(id)) continue;
      seen.add(id);
      const best = Math.max(nameScore(person.shortName, query), nameScore(person.fullName, query), nameScore(person.originalName, query));
      const rank = best * 1000 + prominence(person.shortName, nameKey(query));
      let position = ranked.length;
      while (position > 0 && (ranked[position - 1] as { rank: number }).rank < rank) position--;
      ranked.splice(position, 0, { rank, person });
    }
    return ranked.map((entry) => entry.person);
  }
}

/** `VfbJoinable`: tokens of one name are separated by spaces only ("b." "i." "v." allowed). */
function joinable(gap: string, previous: string): boolean {
  for (const ch of gap) {
    if (ch === "." && (previous === "b" || previous === "i" || previous === "v")) continue;
    if (ch !== " " && ch !== "\t" && ch !== " ") return false;
  }
  return true;
}

function gapBetween(text: string, previous: Token, token: Token): string {
  return text.slice(previous.start + previous.length, token.start);
}

/** `VfbNameSpanEnd`: last token index of the capitalised name run starting at `first`. */
function nameSpanEnd(text: string, tokens: readonly Token[], first: number, index: NameIndex): number {
  let end = first;
  for (let i = first; i < tokens.length; i++) {
    const token = tokens[i] as Token;
    const raw = text.slice(token.start, token.start + token.length);
    let key = token.key;
    if (hasSuffix(raw)) key = nameKey(nameQuery(raw));
    if (isCommonWord(key)) break;
    if (isNoise(key) && !isNameLink(key)) break;
    if (!isNameLink(key) && !index.has(key) && !initialUpper(raw)) break;
    if (i > first && !joinable(gapBetween(text, tokens[i - 1] as Token, token), (tokens[i - 1] as Token).key)) break;
    if (!isNameLink(key)) end = i;
    if (hasSuffix(raw)) break;
  }
  return end;
}

/** A name occurrence found in text. */
export interface Match {
  /** 0-based start in the scanned text. */
  readonly start: number;
  readonly length: number;
  /** Literal text, including a Turkish ending. */
  readonly text: string;
  /** Matched spelling key. */
  readonly key: string;
}

/**
 * `VfbTextMatches`: longest catalog spellings in `text`. In `strict` mode
 * (person discovery) the whole capitalised name run must match, single words
 * must be capitalised and not common words, and context rules (school
 * adjectives, terms, citations) exclude non-person uses.
 */
export function findMatches(text: string, index: NameIndex, strict = false, citationCheck = true): Match[] {
  const result: Match[] = [];
  const tokens = tokenize(text, index.cache);
  let i = 0;
  while (i < tokens.length) {
    const firstToken = tokens[i] as Token;
    const first = firstToken.start;
    let key = "";
    let bestEnd = 0;
    let bestIndex = i;
    let bestKey = "";
    const spanEnd = strict ? nameSpanEnd(text, tokens, i, index) : i;
    for (let j = i; j < tokens.length; j++) {
      const token = tokens[j] as Token;
      const finish = token.start + token.length;
      if (finish - first > MAX_NAME_LENGTH) break;
      if (j > i && !joinable(gapBetween(text, tokens[j - 1] as Token, token), (tokens[j - 1] as Token).key)) break;
      let tokenKey = token.key;
      const raw = text.slice(token.start, finish);
      let inflected = false;
      if (hasSuffix(raw) && !index.has(collapseLetters(key + tokenKey))) {
        tokenKey = nameKey(nameQuery(raw));
        inflected = true;
      }
      key = collapseLetters(key + tokenKey);
      if (key.length >= 3 && !index.heads.has(key.slice(0, 3))) break;
      if (index.has(key)) {
        bestEnd = finish;
        bestIndex = j;
        bestKey = key;
      }
      if (inflected) break;
    }
    if (bestEnd > first) {
      let accepted = true;
      if (strict) {
        if (!index.shapeAllowed(bestKey, bestIndex - i + 1)) accepted = false;
        if (spanEnd > bestIndex) accepted = false;
        if (bestIndex === i) {
          const raw = mid(text, first + 1, bestEnd - first);
          if (isCommonWord(nameKey(nameQuery(raw))) || !initialUpper(raw)) accepted = false;
        }
        if (accepted) accepted = !excludeMention(text, first, bestEnd - first, bestKey, citationCheck);
      }
      if (accepted) result.push({ start: first, length: bestEnd - first, text: text.slice(first, bestEnd), key: bestKey });
      i = Math.max(bestIndex, spanEnd) + 1;
    } else {
      i = spanEnd + 1;
    }
  }
  return result;
}
