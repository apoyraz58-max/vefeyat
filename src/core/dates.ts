/**
 * Death-date expressions. Ported from `VfpFormatDate`, `VfpMarkerLength`,
 * `VfpRequirePrefix`, `VfqDateNote`, `VfmDateText`, `VfaAlreadyHasDate`,
 * `VfcDateKind` and `VfdChronoYear` (VBA 0.3.41).
 *
 * The add-in never invents or converts dates: the catalog expression is kept
 * byte for byte and only its leading marker (ö. / öl. / v.) may be replaced.
 */
import { normalize } from "./normalize.js";
import { normalizeWhitespace } from "./alias.js";
import { vbLower, vbTrim } from "./text.js";

/** Death-date marker setting: as in the catalog, "ö.", "öl." or "v.". */
export type PrefixKey = "original" | "o" | "ol" | "v";

export const PREFIX_KEYS: readonly PrefixKey[] = ["original", "o", "ol", "v"];

export class InvalidSettingError extends Error {
  constructor(message: string) {
    super(message);
    this.name = "InvalidSettingError";
  }
}

/** `VfpRequirePrefix`: accepts " OL " etc.; rejects unknown values. */
export function requirePrefix(value: string): PrefixKey {
  const key = vbLower(vbTrim(value));
  if ((PREFIX_KEYS as readonly string[]).includes(key)) return key as PrefixKey;
  throw new InvalidSettingError("Unknown death-date prefix.");
}

/** Visible marker text of a prefix setting. */
export function prefixLabel(prefix: PrefixKey): string {
  switch (prefix) {
    case "o":
      return "ö.";
    case "ol":
      return "öl.";
    case "v":
      return "v.";
    default:
      return "";
  }
}

/** `VfpIsSpace`. */
function isDateSpace(ch: string | undefined): boolean {
  if (ch === undefined || ch.length !== 1) return false;
  return [9, 10, 13, 32, 160, 8199, 8239].includes(ch.charCodeAt(0));
}

/** `VfpSkipSpaces` on a 0-based index. */
function skipSpaces(value: string, start: number): number {
  while (start < value.length && isDateSpace(value[start])) start++;
  return start;
}

/** `VfpMarkerBoundary`: a marker must be followed by a space, a digit or "?". */
function markerBoundary(value: string, next: number): boolean {
  const ch = value[next];
  if (ch === undefined) return false;
  return isDateSpace(ch) || (ch >= "0" && ch <= "9") || ch === "?";
}

/** `VfpMarkerLength`: length of a leading "öl." (3) or "ö."/"v." (2) marker at `start`. */
export function markerLength(value: string, start: number): number {
  if (start < 0 || start >= value.length) return 0;
  const candidate = vbLower(value.slice(start, start + 3));
  let length = 0;
  if (candidate.startsWith("öl.")) length = 3;
  else if (candidate.startsWith("ö.") || candidate.startsWith("v.")) length = 2;
  else return 0;
  return markerBoundary(value, start + length) ? length : 0;
}

/**
 * `VfpFormatDate`: replaces only the leading marker of a catalog date.
 * Lifespans, B.C. dates, other events and unknown dates are returned unchanged.
 */
export function formatDate(raw: string, prefix: string): string {
  const key = requirePrefix(prefix);
  if (key === "original" || raw.length === 0) return raw;
  let start = skipSpaces(raw, 0);
  if (start >= raw.length) return raw;
  if (raw[start] === "(") start = skipSpaces(raw, start + 1);
  const length = markerLength(raw, start);
  if (length === 0) return raw;
  return raw.slice(0, start) + prefixLabel(key) + raw.slice(start + length);
}

/** `VfqDateNote`: warnings shown with a catalog date (Turkish UI text). */
export function dateNote(raw: string): string {
  const key = normalize(raw);
  for (const part of ["yaklasik", "once", "sonra", "civari", "tahminen", "bilinmiyor", "ihtilaf", "muhtemel"]) {
    if (key.includes(part)) return "Belirsizlik içeren kaynak tarihi aynen korunur; kesin tarihe dönüştürülmez.";
  }
  let note = "";
  if (raw.includes("?") || raw.includes(";")) note = "Tarih belirsizlik veya alternatif içerebilir. Ham kaynak ifadesini kontrol edin.";
  if (raw.includes("/")) note += "\r\nKaynak takvimleri arasında yeni bir dönüşüm hesaplanmadı.";
  return note;
}

/**
 * `VfmDateText`: whether the text right after a name already starts with a
 * date: "(ö. …", "(öl. …", "(v. …", "(d. …" or "(505/1111)"-like years.
 * A numbered comment or page locator such as "(12)" is not a date.
 */
export function startsWithDate(following: string): boolean {
  let text = vbLower(trimDateSpaces(following));
  for (const prefix of ["(o.", "(ö.", "(öl.", "(v.", "(d."]) {
    if (text.startsWith(prefix)) return true;
  }
  if (!text.startsWith("(")) return false;
  text = trimDateSpaces(text.slice(1));
  let at = 0;
  while (at < text.length && (text[at] as string) >= "0" && (text[at] as string) <= "9") at++;
  if (at < 3 || at > 4) return false;
  text = trimDateSpaces(text.slice(at));
  const ch = text[0];
  return ch === undefined || [")", "/", "-", "–", "—", "?", "["].includes(ch);
}

/** `VfaTrimSpaces` with the alias space set. */
function trimDateSpaces(value: string): string {
  return value.replace(/^[    -   　]+|[    -   　]+$/g, "");
}

/** `VfaDateBoundary`. */
function dateBoundary(ch: string): boolean {
  return ' .,;:!?)]}"'.includes(ch);
}

/**
 * `VfaAlreadyHasDate` on the text after a name: true when that text starts
 * with the catalog date in any marker style (case-insensitive), up to the end
 * of the paragraph, table cell or line.
 */
export function alreadyHasDate(afterText: string, rawDate: string): boolean {
  const stop = afterText.search(/[\r\u0007\u000b]/);
  if (stop >= 0) afterText = afterText.slice(0, stop);
  afterText = normalizeWhitespace(afterText);
  for (const key of PREFIX_KEYS) {
    const dateText = normalizeWhitespace(formatDate(rawDate, key));
    if (dateText.length === 0 || afterText.length < dateText.length) continue;
    if (vbLower(afterText.slice(0, dateText.length)) !== vbLower(dateText)) continue;
    if (afterText.length === dateText.length) return true;
    if (dateBoundary(afterText[dateText.length] as string)) return true;
  }
  return false;
}

/** Consistency report category of a date body (text after the marker). */
export type DateKind = "hijriGregorian" | "single" | "range" | "other";

/** `VfcDateKind`: "505/1111", "1111", "1100-1111" or other; null when not a year. */
export function dateKind(body: string): DateKind | null {
  body = vbTrim(body);
  const semicolon = body.indexOf(";");
  if (semicolon >= 0) body = vbTrim(body.slice(0, semicolon));
  if (body.length === 0 || !/^[0-9]/.test(body)) return null;
  if (body.includes("/")) return "hijriGregorian";
  if (body.includes("-") || body.includes("–")) return "range";
  return /^[0-9]+$/.test(body) ? "single" : "other";
}

/**
 * `VfdChronoYear`: Gregorian year used to order people. "505/1111" -> 1111
 * (certain); a single number is returned but marked uncertain (it may be
 * Hijri); "m.ö. 384" -> -384; null when no year can be read.
 */
export function chronoYear(raw: string): { year: number; certain: boolean } | null {
  const slash = raw.indexOf("/");
  const tail = slash >= 0 ? raw.slice(slash + 1) : raw;
  const digits = /[0-9]+/.exec(tail)?.[0] ?? "";
  if (digits.length === 0 || digits.length > 4) return null;
  const year = Number(digits);
  if (` ${normalize(raw)} `.includes(" m o ")) return { year: -year, certain: true };
  return { year, certain: slash >= 0 };
}
