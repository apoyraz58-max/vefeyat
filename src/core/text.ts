/**
 * VBA string semantics used by the original add-in.
 *
 * Offsets and lengths are UTF-16 code units, exactly like VBA `Len`/`Mid$`
 * and Word ranges, so results can be compared with the VBA version 1:1.
 */

/** VBA `Trim$`: removes ASCII spaces only (not tabs or Unicode spaces). */
export function vbTrim(value: string): string {
  let start = 0;
  let end = value.length;
  while (start < end && value.charCodeAt(start) === 32) start++;
  while (end > start && value.charCodeAt(end - 1) === 32) end--;
  return value.slice(start, end);
}

/**
 * VBA `LCase$`: lowercases while keeping the UTF-16 length unchanged.
 * JavaScript turns "İ" into "i̇" (two code units); VBA yields one character.
 */
export function vbLower(value: string): string {
  let result = "";
  for (const ch of value) {
    if (ch === "İ") {
      result += "i";
      continue;
    }
    const lowered = ch.toLowerCase();
    result += lowered.length === ch.length ? lowered : ch;
  }
  return result;
}

/** Unsigned UTF-16 code of the character at `index` (VBA `AscW And &HFFFF&`). */
export function codeAt(value: string, index: number): number {
  return value.charCodeAt(index) & 0xffff;
}
