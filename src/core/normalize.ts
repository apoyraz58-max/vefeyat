/**
 * Search normalization, ported from `VefeyatNormalize`, `VfIsSeparator`,
 * `VfCollapseLetters` and `VfbKey` (VBA 0.3.41).
 *
 * Normalized text is lowercase ASCII-like Latin with Turkish and
 * transliteration diacritics folded away, separators collapsed to single
 * spaces and no leading or trailing space. Non-Latin letters (for example
 * Arabic) are kept; Arabic vowel marks are dropped.
 */

import { vbTrim } from "./text.js";

const FOLDS: ReadonlyArray<readonly [string, string]> = [
  ["àáâãäåāăąǎ", "a"],
  ["çćĉċč", "c"],
  ["ďđḍ", "d"],
  ["èéêëēĕėęě", "e"],
  ["ĝğġģḡ", "g"],
  ["ĥħḥ", "h"],
  ["ìíîïĩīĭįǐ", "i"],
  ["ĵ", "j"],
  ["ķḳ", "k"],
  ["ĺļľł", "l"],
  ["ñńņňṅṇ", "n"],
  ["òóôõöøōŏőǒ", "o"],
  ["ŕŗřṛ", "r"],
  ["śŝşšṣ", "s"],
  ["ţťŧṭ", "t"],
  ["ùúûüũūŭůűųǔ", "u"],
  ["ýÿŷ", "y"],
  ["źżžẓ", "z"],
];

const FOLD_MAP: ReadonlyMap<string, string> = (() => {
  const map = new Map<string, string>();
  for (const [letters, plain] of FOLDS) {
    for (const letter of letters) map.set(letter, plain);
  }
  map.set("æ", "ae");
  map.set("œ", "oe");
  map.set("ß", "ss");
  return map;
})();

/** `VfIsSeparator`: characters that end a search token. */
export function isSeparator(code: number): boolean {
  if (code <= 32 || code === 127 || code === 0xa0 || code === 0xad) return true;
  if (code < 128) return !((code >= 48 && code <= 57) || (code >= 97 && code <= 122));
  if ((code >= 0x2000 && code <= 0x206f) || (code >= 0x2e00 && code <= 0x2e7f)) return true;
  if (code === 0x60c || code === 0x61b || code === 0x61f || code === 0x640) return true;
  if (code === 0xb4 || code === 0x2bb || code === 0x2bc || code === 0x2be || code === 0x2bf) return true;
  if (code === 0xfeff || code === 0x3000 || code === 0xff0c || code === 0xff1b) return true;
  return false;
}

function isIgnoredMark(code: number): boolean {
  // Combining accents and Arabic vocalization marks are not search terms.
  return (code >= 0x300 && code <= 0x36f) || (code >= 0x64b && code <= 0x65f);
}

/** `VefeyatNormalize`. */
export function normalize(text: string): string {
  // Fold Turkish I before lowercasing to avoid locale-specific results.
  let lowered = "";
  for (const ch of text) {
    if (ch === "İ" || ch === "ı" || ch === "I") {
      lowered += "i";
    } else {
      const low = ch.toLowerCase();
      lowered += low.length === ch.length ? low : ch;
    }
  }
  let result = "";
  let pendingSpace = false;
  for (const original of lowered) {
    const ch = FOLD_MAP.get(original) ?? original;
    // A fold may produce two letters ("ae"); each behaves as a plain letter.
    const code = ch.length === 1 ? ch.charCodeAt(0) : 97;
    if (ch.length === 1 && isIgnoredMark(code)) continue;
    if (ch.length === 1 && isSeparator(code)) {
      if (result.length > 0) pendingSpace = true;
      continue;
    }
    if (pendingSpace) result += " ";
    result += ch;
    pendingSpace = false;
  }
  result = ` ${result} `;
  while (result.includes(" ibni ")) result = result.split(" ibni ").join(" ibn ");
  return vbTrim(result);
}

/** `VfCollapseLetters`: collapses runs of the same character ("gazzali" -> "gazali"). */
export function collapseLetters(value: string): string {
  let result = "";
  let previous = "";
  for (let i = 0; i < value.length; i++) {
    const ch = value[i] as string;
    if (ch !== previous) result += ch;
    previous = ch;
  }
  return result;
}

/** `VfbKey`: compact, collapsed lookup key of a name. */
export function nameKey(value: string): string {
  return collapseLetters(normalize(value).split(" ").join(""));
}
