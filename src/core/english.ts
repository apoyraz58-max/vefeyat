/**
 * English (IJMES-style) transliteration of DİA's names, so that a text written
 * in English finds them: "Tāj al-Dīn al-Subkī" for "Sübkî, Tâceddin". The DİA
 * catalog gives the name in Turkish transliteration and its Arabic original
 * (تاج الدين السبكي); the vowels come from the Turkish spelling, the
 * consonants from the Arabic (ق q, خ kh, ث th, ذ dh, غ gh, ش sh, ج j, و w, ع ʿ),
 * the article is written "al-" as in English ("es-Sübkî" → "al-Subkī") and a
 * Turkish compound is taken apart as the Arabic is ("Tâceddin" → "Tāj al-Dīn",
 * "Abdülvehhâb" → "ʿAbd al-Wahhāb"). When the two spellings do not line up
 * letter by letter no English form is made.
 */
import type { Person } from "./catalog.js";
import { textName } from "./names.js";

/** Turkish consonants and the Arabic letters each may stand for. */
const CONSONANTS: Record<string, string> = {
  b: "ب", p: "پب", t: "تطة", s: "سصث", c: "ج", ç: "چج", h: "هحخ", d: "دضذ", z: "زذظض", r: "ر", ş: "ش", g: "غقگك", ğ: "غ",
  f: "ف", k: "كقخ", l: "ل", m: "م", n: "ن", v: "و", y: "ي", j: "ژج",
};

/** Each Arabic letter in English transliteration. */
const ENGLISH: Record<string, string> = {
  ب: "b", پ: "p", ت: "t", ة: "t", ث: "th", ج: "j", چ: "ch", ح: "ḥ", خ: "kh", د: "d", ذ: "dh", ر: "r", ز: "z", ژ: "zh",
  س: "s", ش: "sh", ص: "ṣ", ض: "ḍ", ط: "ṭ", ظ: "ẓ", غ: "gh", ف: "f", ق: "q", ك: "k", گ: "g", ل: "l", م: "m", ن: "n",
  ه: "h", و: "w", ي: "y",
};

const VOWEL: Record<string, string> = {
  a: "a", e: "a", â: "ā", ê: "ā", ā: "ā", i: "i", ı: "i", î: "ī", ī: "ī", u: "u", ü: "u", o: "u", ö: "u", û: "ū", ô: "ū", ū: "ū",
};
const LONG: Record<string, string> = { a: "ā", i: "ī", u: "ū" };
const SILENT = "عءئؤ";

/** Arabic without vowel marks and with one form of each letter. */
function arabicWords(arabic: string): string[] {
  return arabic
    .replace(/[ً-ٰٟـ]/g, "")
    .replace(/[أإٱ]/g, "ا")
    .replace(/ی/g, "ي")
    .replace(/ک/g, "ك")
    .replace(/ہ/g, "ه")
    .replace(/(^|\s)عبدال/g, "$1عبد ال")
    .split(/\s+/)
    .filter((w) => w && w !== "،");
}

const capital = (value: string): string => {
  if (value.startsWith("ʿ") || value.startsWith("ʾ")) return value.charAt(0) + capital(value.slice(1));
  return value.charAt(0).toUpperCase() + value.slice(1);
};

/** The English form of a name in Turkish transliteration, given its Arabic original; null when they do not pair. */
export function englishName(latin: string, arabic: string): string | null {
  return englishParts(latin, arabic)?.join(" ") ?? null;
}

/** The English of each Latin word (a künye "Ebü’l-…" counts as two); null when they do not pair. */
export function englishParts(latin: string, arabic: string): string[] | null {
  if (!/[؀-ۿ]/.test(arabic)) return null;
  const words = arabicWords(arabic);
  const latinWords = latin
    // "Ebü’l-Ferec" is "Ebü" and "el-Ferec", as the Arabic أبو الفرج.
    .replace(/(\p{L})[’'‘]([lrsştdzn])[-‐]/gu, "$1 e$2-")
    .replace(/[‘’'ʿʾ`]/g, "")
    // "Mehdî-Billâh": two words.
    .replace(/(^|\s)(\p{Lu}\p{L}*)[-‐](\p{Lu})/gu, "$1$2 $3")
    .split(/\s+/)
    .filter(Boolean);
  const out: string[] = [];
  let w = 0;
  for (const raw of latinWords) {
    const arabicWord = words[w];
    if (arabicWord === undefined) return null;
    // "b." / "bin" and their Arabic.
    if (/^(b\.|bin)$/i.test(raw)) {
      if (arabicWord !== "بن" && arabicWord !== "ابن") return null;
      out.push("b.");
      w++;
      continue;
    }
    // A regnal number: as it is, with no Arabic word of its own.
    if (/^[IVX]+$/.test(raw)) {
      out.push(raw);
      continue;
    }
    // "Ebû" in the genitive (أبي): Abī.
    if (/^Eb[ûuü]$/i.test(raw) && arabicWord === "ابي") {
      out.push("Abī");
      w++;
      continue;
    }
    // "Billâh" (بالله): billāh.
    if (/^Bill[âa]h$/i.test(raw) && arabicWord === "بالله") {
      out.push("billāh");
      w++;
      continue;
    }
    const core = raw.replace(/^(el|er|es|et|ez|en|ed|eş)[-‐]/i, "");
    const result = word(core, words, w);
    if (!result) return null;
    out.push(result.text);
    w = result.next;
  }
  // Every Arabic word used: the two are the same name.
  return w === words.length ? out : null;
}

/** One Latin word against the Arabic words from `start`; it may take several (a compound). */
function word(latin: string, words: readonly string[], start: number): { text: string; next: number } | null {
  const letters = [...latin.normalize("NFC")];
  const lower = letters.map((ch) => ch.toLocaleLowerCase("tr"));
  let w = start;
  let a = words[w] as string;
  let j = 0;
  let text = "";
  let piece = ""; // the current English word, capitalised when done
  const flush = (): void => {
    text += capital(piece);
    piece = "";
  };
  if (a.length > 3 && a.startsWith("ال")) {
    text += "al-";
    j = 2;
  }
  for (let i = 0; i < letters.length; i++) {
    const ch = lower[i] as string;
    const next = a[j] ?? "";
    if (ch in VOWEL) {
      // The end of an Arabic word inside the Latin one: a compound with the article ("Tâc-ed-din", "Abd-ül-vehhâb").
      if (j >= a.length && i === letters.length - 1) {
        // A case vowel the Arabic does not write: "İbnü" (ابن), "Ümmü" (أمّ).
        continue;
      }
      if (j >= a.length && words[w + 1] !== undefined && !words[w + 1]?.startsWith("ال")) {
        // Two Arabic words written as one Latin word: "Abdümenâf" (عبد مناف), "Ferruhsiyer" (فرّخ سير).
        flush();
        w++;
        a = words[w] as string;
        j = 0;
        text += " ";
        if (lower[i - 1] !== undefined && !(lower[i - 1] as string in VOWEL) && (ch === "ü" || ch === "u" || ch === "i" || ch === "ı" || ch === "e")) {
          // The connecting vowel of a Turkish compound: not in the Arabic.
          const after = a[0] ?? "";
          if (!(after === "ا" || after === "ع" || after === "آ")) continue;
        }
        i--;
        continue;
      }
      if (j >= a.length && words[w + 1]?.startsWith("ال")) {
        flush();
        w++;
        a = words[w] as string;
        if (a === "الله") {
          // "Abdullah", "Nasrullah": Allāh, whatever the Latin vowels.
          text += " Allāh";
          j = a.length;
          i = letters.length;
          break;
        }
        j = 2;
        text += " al-";
        // The article's sound in Latin: the vowel and an "l" or the doubled sun letter.
        const following = lower[i + 1];
        if (following === "l" && a[2] !== "ل") i++;
        else if (following && following === lower[i + 2]) i++;
        else return null;
        continue;
      }
      let vowel = VOWEL[ch] as string;
      if (piece === "" && (next === "ا" || next === "ع" || next === "آ")) {
        if (next === "ع") piece += "ʿ";
        if (next === "آ") vowel = LONG[vowel] ?? vowel;
        j++;
        // "Âişe" عائشة, "İnak" ايناق: the long vowel after the first letter.
        const then = a[j] ?? "";
        const plain = vowel.normalize("NFD").charAt(0);
        if ((then === "ا" && plain === "a") || (then === "ي" && plain === "i" && lower[i + 1] !== "y") || (then === "و" && plain === "u" && lower[i + 1] !== "v")) {
          vowel = LONG[vowel] ?? vowel;
          j++;
        }
      } else if ((next === "ا" || next === "ى" || next === "آ") && (vowel === "a" || vowel === "ā")) {
        vowel = "ā";
        j++;
      } else if (next === "ى" && (vowel === "i" || vowel === "ī") && j === a.length - 1) {
        // A final ى read as ī (Persian, Ottoman): "Lârî" لارى.
        vowel = "ī";
        j++;
      } else if (next === "ي" && (vowel === "i" || vowel === "ī") && lower[i + 1] !== "y" && !(lower[i + 1] && lower[i + 1] as string in VOWEL)) {
        vowel = "ī";
        j++;
      } else if (next === "و" && (vowel === "u" || vowel === "ū") && lower[i + 1] !== "v" && !(lower[i + 1] && lower[i + 1] as string in VOWEL)) {
        vowel = "ū";
        j++;
      }
      piece += vowel;
      // An ayn or hamza after the vowel: "Saʿīd", "Shāfiʿī".
      while (j < a.length && SILENT.includes(a[j] as string) && !(lower[i + 1] && CONSONANTS[lower[i + 1] as string]?.includes(a[j] as string))) {
        piece += a[j] === "ع" ? "ʿ" : "ʾ";
        j++;
      }
      continue;
    }
    const options = CONSONANTS[ch];
    if (!options) return null;
    while (j < a.length && SILENT.includes(a[j] as string) && !options.includes(a[j] as string)) {
      piece += a[j] === "ع" ? "ʿ" : "ʾ";
      j++;
    }
    const letter = a[j];
    // "Takıyy" over تقيّ: the doubled y of a final ī.
    if (ch === "y" && lower[i + 1] === "y" && letter === "ي" && j === a.length - 1 && piece.endsWith("i")) {
      piece = `${piece.slice(0, -1)}ī`;
      j++;
      i++;
      continue;
    }
    if (letter === undefined || !options.includes(letter)) return null;
    j++;
    const english = ENGLISH[letter] ?? ch;
    piece += english;
    // A doubled consonant is one Arabic letter with shadda.
    if (lower[i + 1] === ch) {
      piece += english;
      i++;
    }
  }
  while (j < a.length && (SILENT + "ةه").includes(a[j] as string)) j++;
  // "Amr" عمرو: the waw is not read.
  if (j === a.length - 1 && a === "عمرو") j++;
  if (j < a.length) return null;
  flush();
  return { text, next: w + 1 };
}

/** "al-Subkī" → "as-Subkī", "al-Dīn" → "ad-Dīn": the article as some English texts write it. */
export function assimilated(name: string): string {
  return name.replace(/\bal-(\p{L})(\p{L}?)/gu, (whole, first: string, second: string) => {
    const pair = `${first}${second}`.toLowerCase();
    if (pair === "sh" || pair === "th" || pair === "dh") return `a${pair}-${first}${second}`;
    return "tdrzsṣḍṭẓnl".includes(first.toLowerCase()) ? `a${first.toLowerCase()}-${first}${second}` : whole;
  });
}

const cache = new WeakMap<Person, readonly string[]>();
/** The English forms of a DİA person's name (with and without the assimilated article); none for others. */
export function englishAliases(person: Person): readonly string[] {
  const known = cache.get(person);
  if (known) return known;
  let list: string[] = [];
  if (person.source.startsWith("https://islamansiklopedisi.org.tr/") && person.originalName) {
    const natural = textName(person.shortName);
    const parts = englishParts(natural, person.originalName);
    if (parts) {
      const forms = [parts.join(" ")];
      // "İbn Teymiyye, Takıyyüddin": the part before the comma alone too ("Ibn Taymiyya"), as texts name him.
      const comma = person.shortName.indexOf(",");
      if (comma > 0 && natural !== person.shortName) {
        const head = person.shortName.slice(0, comma).trim().split(/\s+/).length;
        if (head < parts.length) forms.push(parts.slice(parts.length - head).join(" "));
      }
      list = [...new Set(forms.flatMap((f) => [f, assimilated(f)]))];
    }
  }
  cache.set(person, list);
  return list;
}
