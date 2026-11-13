/**
 * Name matching, ported from the VBA 0.3.41 functions noted on each export.
 *
 * Scores (`nameMatch`):
 *   4 identical normalized name
 *   3 identical once spaces are removed ("Kıvâmü's-sünne" = "Kıvâmüssünne")
 *   2 every query token appears in the name
 *   1 substring of the compact name, or a near spelling (edit distance 1–2)
 *   0 no match
 */
import { collapseLetters, normalize } from "./normalize.js";
import { vbTrim } from "./text.js";

const TURKISH_ENDINGS: ReadonlySet<string> = new Set([
  "in", "nin", "un", "nun", "a", "e", "ya", "ye", "na", "ne", "i", "u", "yi", "yu",
  "da", "de", "ta", "te", "nda", "nde", "dan", "den", "tan", "ten", "ndan", "nden",
  "la", "le", "yla", "yle", "daki", "deki", "taki", "teki", "ndaki", "ndeki",
  "ca", "ce", "casina", "cesine", "inki", "ninki", "unki", "nunki",
]);

/**
 * `VfNameQuery`: removes a Turkish case ending written after an apostrophe
 * ("Gazzâlî'nin" -> "Gazzâlî"). Transliterated parts such as "Ebu'l-…" or
 * "…'s-sünne" are kept.
 */
export function nameQuery(value: string): string {
  const plain = value.replace(/[’‘ʼ]/g, "'");
  const at = plain.lastIndexOf("'") + 1; // 1-based like InStrRev
  if (at <= 1) return value;
  const ending = plain.slice(at);
  // Circumflex vowels belong to a transliterated name, not a Turkish ending.
  if (/[âîûÂÎÛ]/.test(ending)) return value;
  return TURKISH_ENDINGS.has(normalize(ending)) ? value.slice(0, at - 1) : value;
}

/** `VfnTextName`: "Râzî, Fahreddin" -> "Fahreddin Râzî"; other values unchanged. */
export function textName(value: string): string {
  value = vbTrim(value);
  const comma = value.indexOf(",") + 1;
  if (comma < 2) return value;
  if (value.indexOf(",", comma) >= 0) return value;
  for (const marker of ["(", ")", ":", ";", "/", "\r", "\n"]) {
    if (value.includes(marker)) return value;
  }
  const left = vbTrim(value.slice(0, comma - 1));
  const right = vbTrim(value.slice(comma));
  return left.length > 0 && right.length > 0 ? `${right} ${left}` : value;
}

/** `VfLookupTokens`: drops Arabic article tokens. */
export function lookupTokens(value: string): string {
  return value
    .split(" ")
    .filter((part) => !["", "el", "er", "es", "et", "ez", "en", "ed", "al"].includes(part))
    .join(" ");
}

/** `VfContainsTokens`: every space-separated token of `term` occurs in `name`. */
export function containsTokens(name: string, term: string): boolean {
  if (term.length === 0) return false;
  return term.split(" ").every((token) => name.includes(token));
}

/** `VfSharesBigram`. */
export function sharesBigram(name: string, term: string): boolean {
  for (let i = 0; i < term.length - 1; i++) {
    if (name.includes(term.slice(i, i + 2))) return true;
  }
  return false;
}

/** `VfNearWord`: Levenshtein distance of `word` and `term` is at most `limit`. */
export function nearWord(word: string, term: string, limit: number): boolean {
  if (Math.abs(word.length - term.length) > limit || word.length === 0) return false;
  let previous = Array.from({ length: term.length + 1 }, (_, j) => j);
  let current = new Array<number>(term.length + 1).fill(0);
  for (let i = 1; i <= word.length; i++) {
    current[0] = i;
    let best = i;
    for (let j = 1; j <= term.length; j++) {
      const cost = word[i - 1] === term[j - 1] ? 0 : 1;
      let value = (previous[j] as number) + 1;
      if ((current[j - 1] as number) + 1 < value) value = (current[j - 1] as number) + 1;
      if ((previous[j - 1] as number) + cost < value) value = (previous[j - 1] as number) + cost;
      current[j] = value;
      if (value < best) best = value;
    }
    if (best > limit) return false;
    [previous, current] = [current, previous];
  }
  return (previous[term.length] as number) <= limit;
}

/** `VfSpellingMatch`. `term` is the compact (space-free) query. */
export function spellingMatch(
  name: string,
  compactName: string,
  term: string,
  foldedName = "",
  foldedTerm = "",
): boolean {
  if (term.length < 5 || term.length > 80) return false;
  if (foldedTerm.length === 0) foldedTerm = collapseLetters(term);
  if (foldedName.length === 0) foldedName = collapseLetters(compactName);
  if (foldedName.includes(foldedTerm)) return true;
  // At lengths >= 5 (distance 1) or >= 9 (distance 2) a possible edit match
  // must retain a bigram; this only rules out impossible matches.
  if (!sharesBigram(compactName, term)) return false;
  const limit = term.length >= 9 ? 2 : 1;
  if (nearWord(compactName, term, limit)) return true;
  return name.split(" ").some((token) => nearWord(token, term, limit));
}

/** `VfNameMatch` on already normalized values. */
export function nameMatch(
  name: string,
  compactName: string,
  term: string,
  compactTerm: string,
  foldedName = "",
  foldedTerm = "",
): number {
  if (term.length === 0 || name.length === 0 || compactTerm.length === 0) return 0;
  if (name === term) return 4;
  if (compactName === compactTerm) return 3;
  if (containsTokens(name, term)) return 2;
  if (compactName.includes(compactTerm)) return 1;
  if (spellingMatch(name, compactName, compactTerm, foldedName, foldedTerm)) return 1;
  return 0;
}

/** `VfLooseNameMatch`: multi-word queries whose words appear in any order. */
export function looseNameMatch(name: string, query: string): number {
  query = lookupTokens(query);
  if (!query.includes(" ")) return 0;
  if (containsTokens(name, query)) return 2;
  const wanted = query.split(" ");
  const words = lookupTokens(name).split(" ");
  if (wanted.length > words.length) return 0;
  const used = new Array<boolean>(words.length).fill(false);
  const matched = new Array<boolean>(wanted.length).fill(false);
  wanted.forEach((want, i) => {
    const j = words.findIndex((word, k) => !used[k] && word === want);
    if (j >= 0) {
      used[j] = true;
      matched[i] = true;
    }
  });
  let approximate = false;
  for (let i = 0; i < wanted.length; i++) {
    if (matched[i]) continue;
    const want = wanted[i] as string;
    if (want.length < 5) return 0;
    const limit = want.length >= 9 ? 2 : 1;
    const j = words.findIndex((word, k) => !used[k] && nearWord(word, want, limit));
    if (j < 0) return 0;
    used[j] = true;
    approximate = true;
  }
  return approximate ? 1 : 2;
}

/** `VefeyatNameMatch`: score of a raw catalog name against a raw query. */
export function nameScore(name: string, query: string): number {
  const term = normalize(nameQuery(query));
  const normalizedName = normalize(name);
  const score = nameMatch(normalizedName, normalizedName.split(" ").join(""), term, term.split(" ").join(""));
  return score === 0 ? looseNameMatch(normalizedName, term) : score;
}

/**
 * `VfProminence`: well-known bearers of a short name rank first
 * ("Râzî" -> Fahreddin er-Râzî). `query` is the compact normalized query.
 */
export function prominence(shortName: string, query: string): number {
  const short = normalize(shortName);
  const preferred: Record<string, string> = {
    razi: "fahreddin er razi",
    gazali: "gazzali",
    buhari: "buhari muhammed b ismail",
    muslim: "muslim b haccac",
    taberi: "taberi muhammed b cerir",
  };
  let value = 0;
  const wanted = preferred[collapseLetters(query)];
  if (wanted !== undefined && short === wanted) value = 5000;
  if (short === query) value += 100;
  return value;
}
