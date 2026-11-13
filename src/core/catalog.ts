/**
 * Local DİA catalog: CSV parsing, person identity and lookup.
 * Ported from `VfParseCsv`, `VfAppendCsvRecord`, `VfValidateIdentities`,
 * `VfReadCatalog`, `VfBuildIndex`, `VefeyatFind`, `VfpPersonKey`,
 * `VfqIdentity`, `VfdArticleUrl`, `VfaNameCandidates`,
 * `VfaConfirmationCandidates`, `VfqStrength` and `VfqSingleSafe`.
 */
import { collapseLetters, normalize } from "./normalize.js";
import { containsTokens, looseNameMatch, lookupTokens, nameMatch, nameQuery, nameScore, prominence } from "./names.js";
import { vbTrim } from "./text.js";

/** One catalog row. Field order matches the CSV and the VBA arrays (0..5). */
export interface Person {
  readonly shortName: string;
  readonly fullName: string;
  readonly originalName: string;
  readonly deathDate: string;
  readonly description: string;
  readonly source: string;
}

export const CSV_HEADERS = ["Kısa Ad", "Tam Ad", "Orijinal Ad", "Vefat Tarihi", "Tanıtım", "Kaynak Link"] as const;

/** VBA row order: short, full, original, date, description, source. */
export function personFields(person: Person): readonly [string, string, string, string, string, string] {
  return [person.shortName, person.fullName, person.originalName, person.deathDate, person.description, person.source];
}

function names(person: Person): readonly [string, string, string] {
  return [person.shortName, person.fullName, person.originalName];
}

export class CatalogError extends Error {
  constructor(message: string, readonly line?: number) {
    super(line === undefined ? message : `CSV line ${line}: ${message}`);
    this.name = "CatalogError";
  }
}

const keyPart = (value: string): string => `${value.length}:${value}`;

/**
 * `VfpPersonKey`: stable identity of a catalog person. Date and description
 * are deliberately not part of it, so a corrected date keeps the identity.
 */
export function personKey(person: Person): string {
  return `Vfp1:${keyPart(person.source)}${keyPart(person.shortName)}${keyPart(person.fullName)}${keyPart(person.originalName)}`;
}

const DIA_ROOT = "https://islamansiklopedisi.org.tr/";

/** `VfdArticleUrl`: canonical DİA article address, or "" when not a DİA article. */
export function articleUrl(value: string): string {
  value = vbTrim(value);
  const www = "https://www.islamansiklopedisi.org.tr/";
  if (value.slice(0, www.length).toLowerCase() === www) value = `https://${value.slice(12)}`;
  if (value.length > 2048) return "";
  if (value.slice(0, DIA_ROOT.length).toLowerCase() !== DIA_ROOT) return "";
  const path = value.slice(DIA_ROOT.length);
  if (path.length === 0) return "";
  if (!/^[a-z0-9\-_/#!]+$/i.test(path)) return "";
  return DIA_ROOT + path;
}

/** `VfqIdentity`: "DIA:<article url>" when available, otherwise the person key. */
export function identity(person: Person): string {
  let url = articleUrl(person.source);
  if (url.length > 0) {
    const fragment = url.indexOf("#");
    if (fragment >= 0) url = url.slice(0, fragment);
    while (url.endsWith("/")) url = url.slice(0, -1);
    if (url.length <= "https://islamansiklopedisi.org.tr".length) url = "";
  }
  return url.length > 0 ? `DIA:${url.toLowerCase()}` : personKey(person);
}

/** `VfParseCsv`: semicolon-separated UTF-8 CSV with the six expected headers. */
export function parseCatalogCsv(content: string): Person[] {
  if (content.startsWith("﻿")) content = content.slice(1);
  const rows: Person[] = [];
  let fields: string[] = [];
  let value = "";
  let state = 0; // 0 unquoted, 1 inside quotes, 2 after a closing quote
  let lineNumber = 1;
  let recordLine = 1;
  let recordStarted = false;
  let haveHeader = false;

  const append = (line: number): void => {
    if (fields.length !== 6) throw new CatalogError("Expected exactly six semicolon-separated columns.", line);
    if (!haveHeader) {
      CSV_HEADERS.forEach((header, i) => {
        if (fields[i] !== header) {
          throw new CatalogError("The six CSV headers do not match the expected catalog format.", line);
        }
      });
      haveHeader = true;
      return;
    }
    const cells = fields.map((cell) => {
      const trimmed = vbTrim(cell);
      return trimmed === "---" ? "" : trimmed;
    }) as [string, string, string, string, string, string];
    if (cells[0].length === 0) throw new CatalogError("A short name is required.", line);
    rows.push({
      shortName: cells[0],
      fullName: cells[1],
      originalName: cells[2],
      deathDate: cells[3],
      description: cells[4],
      source: cells[5],
    });
  };

  for (let i = 0; i < content.length; i++) {
    const ch = content[i] as string;
    if (state === 1) {
      if (ch === '"') {
        if (content[i + 1] === '"') {
          value += '"';
          i++;
        } else {
          state = 2;
        }
      } else {
        value += ch;
        if (ch === "\n") lineNumber++;
      }
      continue;
    }
    if (state === 2 && ch !== ";" && ch !== "\r" && ch !== "\n") {
      throw new CatalogError("Unexpected text after a closing quote.", lineNumber);
    }
    switch (ch) {
      case '"':
        if (value.length !== 0) throw new CatalogError("Quote inside an unquoted field.", lineNumber);
        state = 1;
        recordStarted = true;
        break;
      case ";":
        fields.push(value);
        value = "";
        state = 0;
        recordStarted = true;
        break;
      case "\r":
      case "\n":
        if (recordStarted || fields.length > 0) {
          fields.push(value);
          append(recordLine);
        }
        fields = [];
        value = "";
        state = 0;
        recordStarted = false;
        if (ch === "\r" && content[i + 1] === "\n") i++;
        lineNumber++;
        recordLine = lineNumber;
        break;
      default:
        value += ch;
        recordStarted = true;
    }
  }
  if (state === 1) throw new CatalogError("Unclosed quoted field.", recordLine);
  if (recordStarted || fields.length > 0) {
    fields.push(value);
    append(recordLine);
  }
  if (!haveHeader) throw new CatalogError("The CSV file has no header.");
  if (rows.length === 0) throw new CatalogError("The CSV file contains no people.");
  validateIdentities(rows);
  return rows;
}

/** `VfValidateIdentities`: one person must not carry two different dates. */
function validateIdentities(rows: readonly Person[]): void {
  const dates = new Map<string, string>();
  for (const row of rows) {
    const key = personKey(row);
    const known = dates.get(key);
    if (known === undefined) dates.set(key, row.deathDate);
    else if (known !== row.deathDate) {
      throw new CatalogError(
        "The catalog has conflicting dates for the same person. Review the CSV; no catalog was loaded.",
      );
    }
  }
}

interface IndexedPerson {
  readonly person: Person;
  readonly text: readonly [string, string, string];
  readonly keys: readonly [string, string, string];
  readonly folded: readonly [string, string, string];
  readonly search: string;
}

/** In-memory catalog with the VBA lookup order (`VefeyatFind`). */
export class Catalog {
  private readonly index: readonly IndexedPerson[];

  constructor(readonly people: readonly Person[]) {
    this.index = people.map((person) => {
      const text = names(person).map((value) => normalize(value)) as [string, string, string];
      const keys = text.map((value) => value.split(" ").join("")) as [string, string, string];
      const folded = keys.map((value) => collapseLetters(value)) as [string, string, string];
      const combined = [person.shortName, person.fullName, person.originalName, person.deathDate, person.description].join(" ");
      return { person, text, keys, folded, search: normalize(combined) };
    });
  }

  static fromCsv(content: string): Catalog {
    return new Catalog(parseCatalogCsv(content));
  }

  /**
   * Candidates for a name, best first. Ties keep catalog order. With
   * `includeDescriptions`, rows that mention every query token only in their
   * description or date are appended with the lowest rank.
   */
  find(query: string, includeDescriptions = true): Person[] {
    const term = normalize(nameQuery(query));
    if (term.length === 0 || this.index.length === 0) return [];
    const compact = term.split(" ").join("");
    const folded = collapseLetters(compact);
    const ranked: Array<{ rank: number; person: Person }> = [];
    const insert = (rank: number, person: Person): void => {
      // Stable descending insertion, as in the VBA loop.
      let position = ranked.length;
      while (position > 0 && (ranked[position - 1] as { rank: number }).rank < rank) position--;
      ranked.splice(position, 0, { rank, person });
    };
    for (const row of this.index) {
      let best = 0;
      for (let j = 0; j < 3; j++) {
        const score = nameMatch(row.text[j] as string, row.keys[j] as string, term, compact, row.folded[j] as string, folded);
        if (score > best) best = score;
      }
      if (best > 0) insert(best * 1000 + prominence(row.person.shortName, compact), row.person);
      else if (includeDescriptions && containsTokens(row.search, term)) insert(1, row.person);
    }
    if (ranked.length === 0 && lookupTokens(term).includes(" ")) {
      for (const row of this.index) {
        let best = 0;
        for (let j = 0; j < 3; j++) {
          const score = looseNameMatch(row.text[j] as string, term);
          if (score > best) best = score;
        }
        if (best > 0) insert(best * 1000 + prominence(row.person.shortName, compact), row.person);
      }
    }
    return ranked.map((entry) => entry.person);
  }
}

/** `VfqStrength`: best name score of a person for a query. */
export function strength(person: Person, query: string): number {
  return Math.max(...names(person).map((value) => nameScore(value, query)));
}

/** `VfqSingleSafe`: exactly one candidate and it matches strongly (>= 3). */
export function singleSafe(candidates: readonly Person[], query: string): boolean {
  return candidates.length === 1 && strength(candidates[0] as Person, query) >= 3;
}

function uniqueByKey(people: Iterable<Person>): Person[] {
  const seen = new Set<string>();
  const result: Person[] = [];
  for (const person of people) {
    const key = personKey(person);
    if (!seen.has(key)) {
      seen.add(key);
      result.push(person);
    }
  }
  return result;
}

/** `VfaNameCandidates`: split candidates into strong (score >= 3) and partial matches. */
export function nameCandidates(candidates: readonly Person[], query: string): { exact: Person[]; partial: Person[] } {
  const exact: Person[] = [];
  const partial: Person[] = [];
  for (const person of candidates) {
    const scores = names(person).map((value) => nameScore(value, query));
    if (scores.some((score) => score >= 3)) exact.push(person);
    else if (scores.some((score) => score > 0)) partial.push(person);
  }
  return { exact: uniqueByKey(exact), partial: uniqueByKey(partial) };
}

/** `VfaConfirmationCandidates`: every distinct person whose names match the query at all. */
export function confirmationCandidates(candidates: readonly Person[], query: string): Person[] {
  return uniqueByKey(candidates.filter((person) => names(person).some((value) => nameScore(value, query) > 0)));
}
