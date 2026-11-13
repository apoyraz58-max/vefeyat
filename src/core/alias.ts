/**
 * Selected-name ("alias") rules. Ported from `VfaIsSpace`, `VfaTrimSpaces`,
 * `VfaNormalizeWhitespace`, `VfaSingleLine`, `VfaValidAlias`, `VfaAliasKey`,
 * `VfaSelectionAlias` and `VfaOrderBefore` (VBA 0.3.41).
 */
import { normalize } from "./normalize.js";
import { nameQuery } from "./names.js";
import { codeAt, vbLower } from "./text.js";

/** Longest literal name the add-in accepts (`VfaMaxName`). */
export const MAX_NAME_LENGTH = 200;
/** Longest date expression the add-in accepts (`VfaMaxDate`). */
export const MAX_DATE_LENGTH = 500;

/** `VfaIsSpace`: ASCII, no-break and Unicode spaces (not tab or line breaks). */
export function isSpace(ch: string): boolean {
  const code = codeAt(ch, 0);
  return (
    code === 32 || code === 160 || code === 5760 || (code >= 8192 && code <= 8202) ||
    code === 8239 || code === 8287 || code === 12288
  );
}

/** `VfaTrimSpaces`. */
export function trimSpaces(value: string): string {
  let first = 0;
  let last = value.length;
  while (first < last && isSpace(value[first] as string)) first++;
  while (last > first && isSpace(value[last - 1] as string)) last--;
  return value.slice(first, last);
}

/** `VfaNormalizeWhitespace`: runs of spaces and tabs become one space; edges trimmed. */
export function normalizeWhitespace(value: string): string {
  let result = "";
  let pending = false;
  for (let i = 0; i < value.length; i++) {
    const ch = value[i] as string;
    if (isSpace(ch) || ch === "\t") {
      if (result.length > 0) pending = true;
    } else {
      if (pending) result += " ";
      result += ch;
      pending = false;
    }
  }
  return result;
}

/** `VfaSingleLine`: no C0 controls, DEL, NEL or Unicode line/paragraph separators. */
export function isSingleLine(value: string): boolean {
  for (let i = 0; i < value.length; i++) {
    const code = codeAt(value, i);
    if (code < 32 || code === 127 || code === 133 || code === 8232 || code === 8233) return false;
  }
  return true;
}

/** `VfaValidAlias`. */
export function isValidAlias(value: string): boolean {
  if (value.length === 0 || value.length > MAX_NAME_LENGTH) return false;
  if (!isSingleLine(value)) return false;
  return normalize(value).length > 0;
}

/** `VfaValidDate`. */
export function isValidDate(value: string): boolean {
  if (trimSpaces(value).length === 0 || value.length > MAX_DATE_LENGTH) return false;
  return isSingleLine(value);
}

/** `VfaAliasKey`: literal spelling key (no accent folding). */
export function aliasKey(value: string): string {
  return vbLower(trimSpaces(value));
}

/**
 * `VfaSelectionAlias`: the literal name inside a Word selection, or null.
 * Edge spaces and one trailing paragraph mark are removed; trailing
 * punctuation (",;:!?" and a sentence period after a Turkish ending) is dropped.
 */
export function selectionAlias(selected: string): string | null {
  let candidate = trimSpaces(selected);
  if (candidate.length === 0) return null;
  if (candidate.endsWith("\r")) candidate = trimSpaces(candidate.slice(0, -1));
  if (candidate.length === 0 || candidate.length > MAX_NAME_LENGTH) return null;
  for (let i = 0; i < candidate.length; i++) {
    const code = codeAt(candidate, i);
    if (code >= 127 && code <= 159) return null;
  }
  if (!isValidAlias(candidate)) return null;
  while (candidate.length > 1) {
    const ch = candidate[candidate.length - 1] as string;
    const trimmed = candidate.slice(0, -1);
    if (",;:!?".includes(ch)) candidate = trimmed;
    else if (ch === "." && nameQuery(trimmed).length < trimmed.length) candidate = trimmed;
    else break;
  }
  return candidate;
}

/** Position of a name: main-text anchor, part (0 main, 1 footnote) and offset inside the part. */
export interface DocumentPosition {
  readonly anchor: number;
  readonly part: 0 | 1;
  readonly local: number;
}

/** `VfaOrderBefore`: reading order of main text and footnotes. */
export function orderBefore(a: DocumentPosition, b: DocumentPosition): boolean {
  if (a.anchor !== b.anchor) return a.anchor < b.anchor;
  if (a.part !== b.part) return a.part < b.part;
  return a.local < b.local;
}
