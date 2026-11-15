/**
 * Context rules that keep a catalog spelling from counting as a person:
 * school adjectives ("Hanefî fıkhı"), the terms "usûlî/usul" and "aynı",
 * "… adlı eserinde", and author–title citations. Ported from the VBA 0.3.41
 * `Vfc*` procedures; offsets are 0-based like the VBA arguments.
 */
import { nameQuery } from "./names.js";
import { collapseLetters, nameKey, normalize } from "./normalize.js";
import { vbLower, vbTrim } from "./text.js";
import { charAt, inStr, isNumericToken, left, mid, right, rTrim, upper } from "./vb.js";

/** `VfnHasSuffix`: the value ends in a Turkish case ending after an apostrophe. */
export function hasSuffix(value: string): boolean {
  return nameQuery(value).length < value.length;
}

/** `VfcBoundary`: sentence boundary at a 1-based position. */
export function isBoundary(text: string, at: number): boolean {
  let ch = charAt(text, at);
  if (ch === "\r" || ch === "\n" || ch === "\u000b" || ch === ";" || ch === "!" || ch === "?") return true;
  if (ch !== ".") return false;
  let i = at - 1;
  while (i > 0 && at - i <= 6) {
    ch = charAt(text, i);
    if (ch === " " || ch === "\t" || ch === "," || ch === "(") break;
    i--;
  }
  const leftPart = vbLower(mid(text, i + 1, at - i - 1));
  if (["b", "s", "ss", "c", "bk", "kr", "vd", "hz", "dr", "prof", "thk", "nsr", "trc", "ed", "vol", "pp", "p"].includes(leftPart)) return false;
  i = at + 1;
  let rightPart = "";
  while (i <= text.length && i - at <= 8) {
    rightPart = charAt(text, i);
    if (rightPart !== " " && rightPart !== "\t") break;
    i++;
  }
  if (leftPart.length > 0 && rightPart.length > 0) {
    const last = right(leftPart, 1);
    if (last >= "0" && last <= "9" && vbLower(rightPart) === rightPart && upper(rightPart) !== rightPart) return false;
  }
  return true;
}

/** `VfcWindow`: the clause around a name; returns the text and its 0-based offset. */
export function contextWindow(text: string, first: number, length: number): { window: string; offset: number } {
  let low = first - 320;
  if (low < 0) low = 0;
  let high = first + length + 480;
  if (high > text.length) high = text.length;
  let a = first;
  while (a > low) {
    const ch = charAt(text, a);
    if (isBoundary(text, a) || ch === "(" || ch === "[" || ch === ")" || ch === "]") break;
    a--;
  }
  let b = first + length;
  let depth = 0;
  while (b < high) {
    const ch = charAt(text, b + 1);
    if (isBoundary(text, b + 1)) break;
    if (ch === "(" || ch === "[") depth++;
    if (ch === ")" || ch === "]") {
      if (depth === 0) break;
      depth--;
    }
    b++;
  }
  return { window: mid(text, a + 1, b - a), offset: a };
}

/** `VfcWords`: " word word " with collapsed letters (numbers kept). */
export function words(value: string): string {
  let result = "";
  for (const part of normalize(value).split(" ")) {
    if (part.length > 0) result += `${isNumericToken(part) ? part : collapseLetters(part)} `;
  }
  return ` ${result}`;
}

export function hasWord(wordList: string, value: string): boolean {
  return wordList.includes(` ${value} `);
}

const NARRATIVE =
  "gore|soyle|der|dedi|demistir|belirtir|belirtmistir|savunur|savunmustur|aktarir|aciklar|yazar|yazmistir|vefat|dogdu|dogmustur|yilinda|eserinde|eserinin|sayfasinda|gorusune|tarafindan|dusuncesi|oldugunu|oldugu|ise|fakat|ancak|tartisir|inceler|degerlendirir|kaydeder|nakleder|soylemistir|gorur|olmustur|yazdi|karsilastirir|kabul|rededer|tanimlar|tanimlamistir|zikreder|dusunur|vurgular|ifade".split("|");

/** `VfcNarrative`: the clause speaks about a person (verbs of saying, dying …). */
export function isNarrative(wordList: string): boolean {
  return NARRATIVE.some((item) => hasWord(wordList, item));
}

export function isDigit(ch: string): boolean {
  return ch.length === 1 && ch >= "0" && ch <= "9";
}

/** `VfcLocator`: volume/page, page marker or publication year. */
export function isLocator(value: string): boolean {
  for (let i = 2; i <= value.length - 1; i++) {
    const ch = charAt(value, i);
    if (ch !== "/" && ch !== ":") continue;
    let a = i - 1;
    let b = i + 1;
    while (a > 0 && charAt(value, a) === " ") a--;
    while (b <= value.length && charAt(value, b) === " ") b++;
    if (a > 0 && b <= value.length) {
      const rightPart = charAt(value, b);
      const leftPart = vbLower(charAt(value, a));
      if (isDigit(rightPart) && (isDigit(leftPart) || "ivxlcdm".includes(leftPart))) return true;
    }
  }
  const list = words(value);
  for (let i = 1; i <= value.length; i++) {
    if (isDigit(charAt(value, i))) {
      if (["s", "ss", "sayfa", "p", "pp"].some((w) => hasWord(list, w))) return true;
      break;
    }
  }
  const publication = "beyrut|kahire|istanbul|ankara|press|yayinlari|yayinevi|nesriyat".split("|").some((w) => hasWord(list, w));
  if (publication) {
    for (const part of list.split(" ")) {
      if (part.length === 4 && isNumericToken(part)) return true;
    }
  }
  return false;
}

function isRoman(value: string): boolean {
  return value.length > 0 && [...value].every((ch) => "ivxlcdm".includes(ch));
}

/** `VfcTitleSignal`: a page marker or at least one real title word. */
export function hasTitleSignal(after: string): boolean {
  const list = words(after);
  if (["s", "ss", "sayfa", "p", "pp"].some((w) => hasWord(list, w))) return true;
  return list.split(" ").some((part) => part.length >= 3 && !isNumericToken(part) && part !== "cilt" && !isRoman(part));
}

/** `VfcAuthorComma`: 1-based position of the comma after an author name, or 0. */
export function authorComma(after: string): number {
  let i = 1;
  while (i <= after.length) {
    const ch = charAt(after, i);
    if (ch !== " " && ch !== "\t" && ch !== " ") break;
    i++;
  }
  if (charAt(after, i) === ",") return i;
  const comma = inStr(i, left(after, 120), ",");
  if (comma === 0) return 0;
  const prefix = words(mid(after, i, comma - i));
  return left(prefix, 4) === " ve " || charAt(after, i) === "&" ? comma : 0;
}

/** `VfcCitation`: "Name, Title, 2/34" without narrative verbs. */
export function isCitation(window: string, localFirst: number, length: number): boolean {
  let after = mid(window, localFirst + length + 1);
  const comma = authorComma(after);
  if (comma === 0) return false;
  after = mid(after, comma);
  if (!isLocator(after) || !hasTitleSignal(after)) return false;
  return !isNarrative(words(window));
}

const SCHOOL_SUFFIXES = new Set([
  "", "ler", "lar", "lerde", "larda", "lerden", "lardan", "lerin", "larin", "lere", "lara",
  "lerce", "larca", "lerince", "larinca", "nin", "ni", "ye", "yi", "den", "de",
]);

/** `VfcSchoolRoot`: "hanefi", "safi" … when the key is a school adjective. */
export function schoolRoot(key: string): string {
  for (const root of ["hanefi", "safi", "maliki", "hanbeli", "zahiri", "caferi", "esari", "maturidi"]) {
    if (key.startsWith(root) && SCHOOL_SUFFIXES.has(key.slice(root.length))) return root;
  }
  return "";
}

function endsWithWord(before: string, stem: string): boolean {
  return right(before, stem.length + 1) === ` ${stem}`;
}

function firstWord(after: string): string {
  return vbTrim(after).split(" ")[0] ?? "";
}

/** `VfcTerm`: "usûlî düşünce", "usul fıkhı" … are terms, not the scholar. */
export function isTerm(window: string, localFirst: number, length: number, key: string): boolean {
  if (key !== "usuli" && key !== "usul") return false;
  if (hasSuffix(mid(window, localFirst + 1, length))) return false;
  const before = rTrim(words(left(window, localFirst)));
  if (["sair", "sairi", "mutasavif"].some((stem) => endsWithWord(before, stem))) return false;
  const afterWords = vbTrim(words(mid(window, localFirst + length + 1)));
  if (afterWords.length === 0) return false;
  const after = firstWord(afterWords);
  if (after.startsWith("fikih") || after.startsWith("fikh") || after.startsWith("hadis") || after.startsWith("tefsir") || after === "din" || after.startsWith("dine")) return true;
  return "dusunce|yaklasim|ilke|mesele|tartisma|cerceve|yontem|anlayis|birikim|gelenek|geleneg|gorus|delil|kaide|degerlendirme|tahlil|inceleme"
    .split("|")
    .some((stem) => after.startsWith(stem));
}

/** `VfcOrdinary`: "aynı" (same) is not Aynî the scholar. */
export function isOrdinary(window: string, localFirst: number, length: number, key: string): boolean {
  if (key !== "ayni") return false;
  let literal = mid(window, localFirst + 1, length);
  if (hasSuffix(literal)) return false;
  const before = rTrim(words(left(window, localFirst)));
  if (["imam", "alame", "hafiz", "alim", "muhadis", "sair"].some((stem) => endsWithWord(before, stem))) return false;
  // Dotless "aynı" (same) and marked "Aynî" share a search key, not a sense.
  literal = vbLower(literal.replace(/I/g, "ı"));
  if (literal.includes("î") || literal.includes("ī")) return false;
  if (literal === "aynı") return true;
  const afterWords = vbTrim(words(mid(window, localFirst + length + 1)));
  if (afterWords.length === 0) return false;
  const after = firstWord(afterWords);
  return "muelif|yazar|eser|kitap|gorus|sekil|donem|zaman|konu|mesele|dusunce|kaynak|metin|kisi|sahis|nokta|sonuc|anlam|hukum|yol|tavir"
    .split("|")
    .some((stem) => after.startsWith(stem));
}

const PERSON_PREDICATES = new Set(
  "der|dedi|demistir|belirtir|savunur|savunmustur|aktarir|aciklar|yazar|vefat|dogdu|dusunur|tartisir|inceler|vurgular|rededer|tanimlar|kaydeder|nakleder|soylemistir".split("|"),
);

function isPersonPredicate(after: string): boolean {
  let count = 0;
  let found = false;
  for (const part of vbTrim(after).split(" ")) {
    if (part.length === 0) continue;
    count++;
    if (count > 6) break;
    if (schoolRoot(part).length > 0) return false;
    if (part.startsWith("mezhep") || part.startsWith("mezheb") || part.startsWith("ekol")) return false;
    if (PERSON_PREDICATES.has(part)) found = true;
  }
  return found;
}

/** `VfcSchool`: a school adjective or label, not a person. */
export function isSchool(window: string, localFirst: number, length: number, literal: string): boolean {
  const key = nameKey(nameQuery(literal));
  const root = schoolRoot(key);
  if (root.length === 0) return false;
  if (key !== root) return true;
  const before = rTrim(words(left(window, localFirst)));
  const after = words(mid(window, localFirst + length + 1));
  if (["imam", "fakih", "alim", "muhadis"].some((item) => endsWithWord(before, item))) return false;
  // "Eş'arî (ö. 324/935)", "Ebü'l-Hasan el-Eş'arî": the scholar himself.
  if (/^\s*\((ö|öl|v)\.\s*\d/.test(mid(window, localFirst + length + 1))) return false;
  if (/^(el|al)[-‐]/i.test(literal) || /\b(el|al)[-‐]$/i.test(left(window, localFirst))) return false;
  if (vbTrim(before).length > 0) {
    const prior = vbTrim(before).split(" ");
    let index = prior.length - 1;
    if (prior[index] === "ve" || prior[index] === "ile") index--;
    if (index >= 0 && schoolRoot(prior[index] as string).length > 0) return true;
  }
  if (root === "safi" && hasSuffix(literal)) return false;
  if ((root === "safi" || root === "esari" || root === "maturidi") && isPersonPredicate(after)) return false;
  // "Şâfiî, Hanefî fakihlerinden farklı olarak … reddeder": the scholar as subject.
  if ((root === "safi" || root === "esari" || root === "maturidi") && /^\s*,/.test(mid(window, localFirst + length + 1)) && isNarrative(words(window))) return false;
  // Direct person predicates win over a more distant school reference.
  const next = firstWord(after);
  if (next !== "gore" && isNarrative(` ${next} `)) return false;
  for (const part of words(window).split(" ")) {
    if (part.startsWith("mezhep") || part.startsWith("mezheb") || part.startsWith("ekol")) return true;
  }
  // A bare school label is insufficient evidence for any of the many scholars
  // sharing that nisba; explicit personal names remain eligible.
  return true;
}

/** `VfcWorkUsage`: "… adlı eserinde", "… isimli kitabın". */
export function isWorkUsage(after: string): boolean {
  const key = vbTrim(words(after));
  if (key.length === 0) return false;
  const list = key.split(" ");
  if (list.length < 2) return false;
  if (list[0] !== "adli" && list[0] !== "isimli") return false;
  return ["eser", "eserinde", "eserin", "kitap", "kitabinda", "kitabin", "makale", "makalesinde", "risale", "risalesinde"].includes(list[1] as string);
}

/** `VfcNeedsContext`. */
export function needsContext(text: string, first: number, length: number, key: string): boolean {
  if (isWorkUsage(mid(text, first + length + 1, 100))) return true;
  if (schoolRoot(key).length > 0 || key === "usuli" || key === "usul" || key === "ayni") return true;
  return authorComma(mid(text, first + length + 1, 160)) > 0;
}

/** Nouns that make a preceding nisba an adjective: "Nakşibendî tarikatı", "Abbâsî halifesi". */
const GROUP_NOUNS =
  "tarikat|dervis|seyh|sufi|mutasavvif|silsile|tekke|dergah|zaviye|asitane|ayin|mensup|mensub|mezhep|mezheb|ekol|firka|cemaat|zumre|topluluk|cevre|gelenek|geleneg|anlayis|dusunce|kelam|fikih|fikh|akide|akaid|itikad|inanc|alim|ulema|fakih|fukaha|kelamci|mutekellim|muellif|muhaddis|muctehid|halife|hanedan|devlet|sultan|emir|vezir|ordu|donem|devir|devri|saray|mimari|sanat|uslup|uslub|edebiyat|hareket|akim|egilim|yorum|okul|kol|mesrep|mesreb|adab|erkan|kultur|toplum|kesim|grup|taraftar|camia|muhit|nesil|kusak|vakif|medrese|cami|mescit|imparatorluk|beylik|aile|soy|boy|kabile|asiret|tarih|kaynak|metin|literatur|risale|eser"
    .split("|")
    .map(collapseLetters); // compared with words(), which collapses doubled letters

const TITLES = new Set("imam|seyh|hz|hazret|hazreti|mevlana|hace|haci|molla|baba|dede|sultan|emir|kadi|hafiz|alim|fakih|muhaddis|sair|mutasavvif|ebu|ibn|ibni".split("|").map(collapseLetters));

function startsGroupNoun(word: string): boolean {
  return GROUP_NOUNS.some((stem) => word.startsWith(stem));
}

/**
 * A one-word nisba used as an adjective of a group, order, dynasty or school:
 * "Nakşibendî tarikatı", "Kādirî ve Rifâî tarikatları", "Abbâsî halifesi",
 * "Sünnî ve Şiî müellifler"; also an izafet like "Ehl-i hadîs". A title
 * before it ("Şeyh", "İmam") or a death date after it keeps the person.
 */
export function isGroupAdjective(text: string, first: number, length: number, literal: string): boolean {
  const bare = nameQuery(literal);
  if (bare.length < literal.length) return false; // "Nakşibendî'nin": the person
  const after = text.slice(first + length);
  if (/^[-‐][iıuü]\s+\p{Ll}/u.test(after) || /[-‐][iıuü]$/.test(bare) && /^\s+\p{Ll}/u.test(after)) return true;
  if (/\s/.test(bare.trim())) return false; // "Bahâeddin Nakşibend": an explicit name
  const key = nameKey(bare);
  if (!/[iy]$/.test(key)) return false;
  if (/^\s*\((ö|öl|v|d|m)\.|^\s*\(\d/.test(after)) return false;
  const before = words(text.slice(Math.max(0, first - 40), first)).trim().split(" ");
  const prior = before[before.length - 1] ?? "";
  if (TITLES.has(prior)) return false;
  const clause = after.slice(0, 160).split(/[,;:.()]/)[0] ?? "";
  const next = words(clause).trim().split(" ").filter(Boolean);
  // "X tarikatı", "X ve Y tarikatları", "X, Y ve Z şeyhleri"
  for (let i = 0; i < Math.min(next.length, 7); i++) {
    const word = next[i] as string;
    if (startsGroupNoun(word)) return true;
    if (word === "ve" || word === "ile" || word === "veya") continue;
    // Another nisba in the list ("Rifâî"); stop at anything else.
    if (/[iy]$/.test(word) && word.length >= 2) continue;
    break;
  }
  return false;
}

/**
 * `VfcExclude`: whether a matched catalog spelling at 0-based `first` should
 * not count as a person mention.
 */
export function excludeMention(text: string, first: number, length: number, key = "", citationCheck = true): boolean {
  if (key.length === 0) key = nameKey(nameQuery(mid(text, first + 1, length)));
  if (isGroupAdjective(text, first, length, mid(text, first + 1, length))) return true;
  if (!needsContext(text, first, length, key)) return false;
  const { window, offset } = contextWindow(text, first, length);
  const literal = mid(text, first + 1, length);
  if (isWorkUsage(mid(text, first + length + 1, 100))) return true;
  if (citationCheck && isCitation(window, first - offset, length)) return true;
  if (isOrdinary(window, first - offset, length, key)) return true;
  if (isTerm(window, first - offset, length, key)) return true;
  return isSchool(window, first - offset, length, literal);
}
