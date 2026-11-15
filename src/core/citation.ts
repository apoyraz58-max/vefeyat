/**
 * Citation islands: spans such as "Gazzâlî, el-Müstasfâ, 2/34", "bk. Râzî,
 * Mefâtîh, s. 12" or "Yazar (2010: 45)" in which a name refers to a work,
 * not to the person. Ported from the VBA 0.3.41 `Vfx*` procedures. Units and
 * spans use the VBA convention: [0-based start, 0-based exclusive end].
 */
import { hasSuffix, hasTitleSignal, hasWord, isBoundary, isDigit, words, authorComma } from "./context.js";
import { initialUpper, isBreak, isCommonWord, isNameLink, tokenize, wordCount } from "./scan.js";
import { nameKey } from "./normalize.js";
import { vbLower, vbTrim } from "./text.js";
import { charAt, inStr, left, mid, upper } from "./vb.js";

export type Span = readonly [start: number, end: number];

/** A name occurrence: 0-based start, length, literal text. */
export type NameHit = readonly [start: number, length: number, text: string];

function isWhite(ch: string): boolean {
  return ch === " " || ch === "\t" || ch === "\r" || ch === "\n" || ch === "\u000b" || ch === " ";
}

/** `VfxQuoteClose`: 1-based position of the matching closing quote, or 0. */
export function quoteClose(text: string, at: number): number {
  let closing: string;
  switch (charAt(text, at)) {
    case '"':
      closing = '"';
      break;
    case "“":
    case "„":
      closing = "”";
      break;
    case "«":
      closing = "»";
      break;
    default:
      return 0;
  }
  return inStr(at + 1, text, closing);
}

const ABBREVIATIONS = new Set(
  "age|agm|amlf|bkz|krs|a|i|ii|iii|iv|v|dn|nr|no|bs|eds|haz|cf|vgl|b|s|ss|c|bk|kr|vd|hz|dr|prof|thk|nsr|trc|ed|vol|pp|p".split("|"),
);

/** `VfxDotBoundary`: a period that ends a sentence (not an abbreviation or initial). */
function dotBoundary(text: string, at: number): boolean {
  if (!isBoundary(text, at)) return false;
  if (at < text.length && at > 1) {
    const nextChar = charAt(text, at + 1);
    if (!isBreak(nextChar) && at + 2 <= text.length && charAt(text, at + 2) === ".") return false;
  }
  let i = at - 1;
  while (i > 0 && at - i <= 16) {
    const ch = charAt(text, i);
    if (isWhite(ch) || "(,;[".includes(ch)) break;
    i--;
  }
  const raw = mid(text, i + 1, at - i - 1);
  if (raw.length === 1 && upper(raw) === raw && vbLower(raw) !== raw) return false;
  if (raw.includes(".") && raw.split(".").every((part) => part.length === 1)) return false;
  return !ABBREVIATIONS.has(nameKey(raw));
}

/** `VfxUnits`: sentence-like units, ignoring boundaries inside quotes. */
export function units(text: string): Span[] {
  const result: Span[] = [];
  let start = 0;
  let quotedUntil = 0;
  for (let i = 1; i <= text.length; i++) {
    const ch = charAt(text, i);
    let boundary = false;
    if (i > quotedUntil) {
      quotedUntil = quoteClose(text, i);
      if (quotedUntil === 0) {
        boundary = ch === ";" || ch === "؛" || ch === "!" || ch === "?";
        if (ch === ".") boundary = dotBoundary(text, i);
      }
    }
    if (boundary) {
      if (i - 1 > start) result.push([start, i - 1]);
      start = i;
    }
  }
  if (text.length > start) result.push([start, text.length]);
  return result;
}

const NARRATION =
  "der|dedi|demistir|soylemistir|soylemektedir|savunur|savunmustur|belirtir|belirtmistir|aktarir|aciklar|yazar|yazmistir|dusunur|tartisir|inceler|vurgular|zikreder|kaydeder|nakleder|rededer|vefat|dogdu|olmustur|yilinda|sayfasinda|eserinde|gore".split("|");

function isNarration(value: string): boolean {
  const list = words(value);
  return NARRATION.some((item) => hasWord(list, item));
}

function narrationOutsideQuotes(value: string): boolean {
  let i = 1;
  let first = 1;
  let outside = "";
  while (i <= value.length) {
    const closing = quoteClose(value, i);
    if (closing > i) {
      outside += `${mid(value, first, i - first)} `;
      i = closing + 1;
      first = i;
    } else {
      i++;
    }
  }
  return isNarration(outside + mid(value, first));
}

function prefixIsName(value: string): boolean {
  if (value.length > 300 || isNarration(value)) return false;
  const stop = new Set(["bu", "su", "icin", "gibi", "hakkinda", "uzerine", "konuda", "konuyla", "gorus", "gorusu", "karsilastirma", "eserlerinde", "kitaplarinda", "dusuncesi"]);
  return !words(value).split(" ").some((part) => stop.has(part));
}

function pageEnd(text: string, at: number): number {
  let i = at;
  while (i <= text.length) {
    const ch = charAt(text, i);
    if (!isDigit(ch) && ch !== "-" && ch !== "–" && ch !== "—") break;
    i++;
  }
  return i - 1;
}

type Locator = readonly [start: number, end: number, kind: "page" | "marked" | "year" | "bare"];

/** `VfxLocators`: volume/page ("2/34", "II: 12"), page markers ("s. 12"), years, bare pages. */
export function locators(text: string): Locator[] {
  const result: Locator[] = [];
  const list = words(text);
  const publication = "dimask|beyrut|kahire|istanbul|ankara|london|oxford|cambridge|press|matbaa|yayinlari|yayinevi|nesriyat|thk|nsr|trc|haz"
    .split("|")
    .some((part) => hasWord(list, part));
  let i = 1;
  while (i <= text.length) {
    const quotedUntil = quoteClose(text, i);
    if (quotedUntil > i) {
      i = quotedUntil + 1;
      continue;
    }
    if (!isDigit(charAt(text, i))) {
      i++;
      continue;
    }
    const start = i - 1;
    let finish = i;
    while (finish <= text.length && isDigit(charAt(text, finish))) finish++;
    let b = finish;
    while (b <= text.length && isWhite(charAt(text, b))) b++;
    let ch = charAt(text, b);
    if (ch === "/" || ch === ":") {
      b++;
      while (b <= text.length && isWhite(charAt(text, b))) b++;
      if (isDigit(charAt(text, b))) {
        finish = pageEnd(text, b);
        result.push([start, finish, "page"]);
        i = finish + 1;
        continue;
      }
    }
    let a = i - 1;
    while (a > 0 && isWhite(charAt(text, a))) a--;
    if (a > 0) {
      ch = charAt(text, a);
      if (ch === "/" || ch === ":") {
        b = a - 1;
        while (b > 0 && isWhite(charAt(text, b))) b--;
        a = b;
        while (a > 0 && "IVXLCDMivxlcdm".includes(charAt(text, a))) a--;
        if (b > a) {
          finish = pageEnd(text, i);
          result.push([a, finish, "page"]);
          i = finish + 1;
          continue;
        }
      }
      b = a;
      if (charAt(text, b) === ".") b--;
      a = b;
      while (a > 0 && !isBreak(charAt(text, a))) a--;
      const word = nameKey(mid(text, a + 1, b - a));
      if (["s", "ss", "p", "pp", "sayfa"].includes(word)) {
        finish = pageEnd(text, i);
        result.push([a, finish, "marked"]);
        i = finish + 1;
        continue;
      }
    }
    if (publication && finish - i === 4) {
      result.push([start, finish - 1, "year"]);
    } else {
      a = i - 1;
      while (a > 0 && isWhite(charAt(text, a))) a--;
      if (a > 0 && charAt(text, a) === ",") {
        b = finish;
        while (b <= text.length && isWhite(charAt(text, b))) b++;
        ch = charAt(text, b);
        if (ch === "-" || ch === "–" || ch === "—") {
          b++;
          while (b <= text.length && isWhite(charAt(text, b))) b++;
          if (isDigit(charAt(text, b))) {
            finish = pageEnd(text, b);
            result.push([start, finish, "bare"]);
            i = finish + 1;
            continue;
          }
        } else if (ch === "" || ch === "." || ch === "," || ch === ")") {
          result.push([start, finish - 1, "bare"]);
        }
      }
    }
    i = finish;
  }
  return result;
}

/** `VfxWorkLead`: text that starts like a work title. */
function workLead(value: string): boolean {
  let clean = vbTrim(value);
  while (clean.length > 0 && isWhite(left(clean, 1))) clean = clean.slice(1);
  if (clean.length === 0) return false;
  let lead = ["el-", "et-", "er-", "es-", "ez-", "en-", "ed-"].includes(vbLower(left(clean, 3)));
  if (initialUpper(clean)) lead = true;
  if (quoteClose(clean, 1) > 1) lead = true;
  return lead;
}

/** `VfxAuthorStart`: start of a capitalised author name before a comma, or -1. */
function authorStart(text: string, first: number, comma: number): number {
  let finish = comma;
  while (finish > first && isWhite(charAt(text, finish))) finish--;
  if (finish - first > 400) first = finish - 400;
  if (charAt(text, finish) === ")") {
    let depth = 0;
    for (let i = finish; i > first; i--) {
      const ch = charAt(text, i);
      if (ch === ")") depth++;
      if (ch === "(") {
        depth--;
        if (depth === 0) {
          finish = i - 1;
          break;
        }
      }
    }
  }
  const prefix = mid(text, first + 1, finish - first);
  const tokens = tokenize(prefix);
  let start = -1;
  for (let i = tokens.length - 1; i >= 0; i--) {
    const token = tokens[i]!;
    const raw = mid(prefix, token.start + 1, token.length);
    const key = token.key;
    if (isCommonWord(key) && key !== "ve" && key !== "ile") {
      if (key !== "yazar" || !initialUpper(raw)) break;
    }
    if (!initialUpper(raw) && !isNameLink(key) && key !== "ve" && key !== "ile") break;
    start = first + token.start;
  }
  return start;
}

function probe(text: string, unit: Span, hit: NameHit, names: readonly NameHit[]): Span | null {
  const nameStart = hit[0];
  const nameEnd = nameStart + hit[1];
  const tail = mid(text, nameEnd + 1, unit[1] - nameEnd);
  const comma = authorComma(tail);
  let titleStart: number;
  if (comma > 0) {
    titleStart = nameEnd + comma;
  } else {
    const prefix = words(mid(text, unit[0] + 1, nameStart - unit[0]));
    if (!hasWord(prefix, "bk") && !hasWord(prefix, "bkz") && !hasWord(prefix, "krs")) return null;
    if (!workLead(tail)) return null;
    titleStart = nameEnd;
  }
  const title = mid(text, titleStart + 1, unit[1] - titleStart);
  const found = locators(title);
  let finish: number;
  if (found.length === 0) {
    if (comma === 0 || !workLead(title) || isNarration(title)) return null;
    if (wordCount(title) > 16 || !hasTitleSignal(title)) return null;
    finish = unit[1];
  } else {
    let chosen = false;
    let localEnd = 0;
    for (const locator of found) {
      if (!chosen) {
        const gap = left(title, locator[0]);
        if (!hasTitleSignal(gap) && locator[2] !== "marked") return null;
        if (locator[2] === "bare" && !workLead(title)) return null;
        if (narrationOutsideQuotes(gap)) return null;
        chosen = true;
      } else {
        const gap = mid(title, localEnd + 1, locator[0] - localEnd);
        if (isNarration(gap)) break;
      }
      localEnd = locator[1];
    }
    finish = titleStart + localEnd;
  }
  if (finish <= titleStart) return null;
  let start = nameStart;
  if (prefixIsName(mid(text, unit[0] + 1, nameStart - unit[0]))) start = unit[0];
  const gap = mid(text, finish + 1, unit[1] - finish);
  if (isNarration(gap)) {
    const nextSubject = names.some((item) => item[0] >= finish && item[0] < unit[1]);
    if (!nextSubject) start = titleStart;
  }
  if (hasSuffix(hit[2])) start = titleStart;
  return finish > start ? [start, finish] : null;
}

/** `VfxContains`. */
export function containsSpan(spans: readonly Span[], first: number, length: number): boolean {
  return spans.some((span) => span[0] <= first && span[1] >= first + length);
}

/** `VfxAuthorYearEnd`: "(2010: 45)" right after a name; 1-based end in `tail` or 0. */
function authorYearEnd(tail: string): number {
  let i = 1;
  while (i <= tail.length && isWhite(charAt(tail, i))) i++;
  if (charAt(tail, i) !== "(") return 0;
  const first = i + 1;
  for (i = first; i <= first + 3; i++) {
    if (!isDigit(charAt(tail, i))) return 0;
  }
  while (i <= tail.length && isWhite(charAt(tail, i))) i++;
  let ch = charAt(tail, i);
  if (ch !== ":" && ch !== ",") return 0;
  const closing = inStr(i + 1, tail, ")");
  if (closing === 0 || closing - i > 32) return 0;
  let page = false;
  for (let j = i + 1; j <= closing - 1; j++) {
    ch = charAt(tail, j);
    if (isDigit(ch)) page = true;
    else if (!isWhite(ch) && ch !== "-" && ch !== "–") return 0;
  }
  return page ? closing : 0;
}

/**
 * `VfxClassify`: citation spans in a text given its name hits (sorted by
 * start). Also finds author–title citations of names absent from the catalog.
 */
export function citationSpans(text: string, names: readonly NameHit[]): Span[] {
  const result: Span[] = [];
  const unitList = units(text);
  let index = 0;
  for (const hit of names) {
    while (index < unitList.length && hit[0] >= (unitList[index] as Span)[1]) index++;
    if (index >= unitList.length) break;
    const unit = unitList[index] as Span;
    if (containsSpan(result, hit[0], hit[1])) continue;
    let span = probe(text, unit, hit, names);
    if (!span) {
      const authorYear = authorYearEnd(mid(text, hit[0] + hit[1] + 1, 100));
      if (authorYear > 0) span = [hit[0], hit[0] + hit[1] + authorYear];
    }
    if (span) result.push(span);
  }
  // Structural author/work evidence also covers authors absent from the
  // catalog and names occurring only inside titles or editor credits.
  for (const unit of unitList) {
    for (let i = unit[0] + 1; i <= unit[1]; i++) {
      if (charAt(text, i) !== ",") continue;
      if (containsSpan(result, i - 1, 1)) continue;
      const start = authorStart(text, unit[0], i - 1);
      if (start >= unit[0] && start < i - 1) {
        const pseudo: NameHit = [start, i - 1 - start, mid(text, start + 1, i - 1 - start)];
        const span = probe(text, unit, pseudo, names);
        if (span) result.push(span);
      }
    }
  }
  return result;
}
