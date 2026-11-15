/**
 * 1-based VBA string helpers used by the scanner ports, so that index
 * arithmetic can be copied from the VBA source without off-by-one changes.
 */

/** VBA `Mid$(text, start, length)`; start is 1-based. Out of range gives "". */
export function mid(text: string, start: number, length?: number): string {
  if (start < 1) start = 1;
  return length === undefined ? text.slice(start - 1) : text.slice(start - 1, start - 1 + Math.max(0, length));
}

/** One character at a 1-based position, or "" outside the string (VBA `Mid$(text, i, 1)`). */
export function charAt(text: string, position: number): string {
  return position >= 1 && position <= text.length ? (text[position - 1] as string) : "";
}

/** VBA `Left$`. */
export function left(text: string, length: number): string {
  return text.slice(0, Math.max(0, length));
}

/** VBA `Right$`. */
export function right(text: string, length: number): string {
  return length <= 0 ? "" : text.slice(-length);
}

/** VBA `InStr(start, text, find)`: 1-based position, 0 when absent. */
export function inStr(start: number, text: string, find: string): number {
  return text.indexOf(find, start - 1) + 1;
}

/** VBA `RTrim$` (ASCII spaces only). */
export function rTrim(value: string): string {
  return value.replace(/ +$/, "");
}

/** VBA `IsNumeric` for the plain digit tokens produced by normalization. */
export function isNumericToken(value: string): boolean {
  return /^[0-9]+$/.test(value);
}

/** VBA `UCase$` on one character, length preserving. */
export function upper(ch: string): string {
  const u = ch.toUpperCase();
  return u.length === ch.length ? u : ch;
}
