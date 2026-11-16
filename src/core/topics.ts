/**
 * Konu: the fields a sentence speaks about and the fields a catalog
 * description names ("müfessir" → tefsir, "hekim" → tıp). Used to tell
 * homonyms apart: "Râzî tefsirinde…" is the müfessir, "Râzî tıpta…" the hekim.
 */
import { words } from "./context.js";
import { collapseLetters } from "./normalize.js";

/** Field → word stems (as words() writes them: normalized, doubled letters collapsed). */
const FIELDS: Record<string, readonly string[]> = {
  tefsir: ["tefsir", "mufesir", "ayet", "sure", "kuran"],
  tıp: ["tip", "tib", "hekim", "tabib", "tabip", "ilac", "hastalik", "tedavi", "anatomi", "eczaci"],
  kelâm: ["kelam", "mutekelim", "akaid", "akide", "itikad", "esari", "maturidi", "mutezil"],
  fıkıh: ["fikih", "fikh", "fakih", "fukaha", "kadi", "mufti", "ictihad", "fetva", "furu"],
  hadis: ["hadis", "muhadis", "rivayet", "ravi", "sened", "isnad", "hafiz"],
  felsefe: ["felsefe", "filozof", "feylesof", "mantik", "mesai", "metafizik"],
  tasavvuf: ["tasavuf", "mutasavif", "sufi", "tarikat", "zikir", "vahdet", "dervis", "seyh"],
  şiir: ["siir", "sair", "divan", "beyit", "kaside", "gazel", "mesnevi", "nazim"],
  tarih: ["tarih", "tarihci", "muverih", "vakayiname", "vekayi"],
  dil: ["nahiv", "nahv", "sarf", "lugat", "sozluk", "belagat", "dilci", "edip", "edib", "edebiyat", "gramer"],
  kıraat: ["kiraat", "kura", "kari", "tecvid"],
  siyaset: ["hukumdar", "sultan", "vezir", "sadrazam", "emir", "halife", "devlet", "siyaset", "komutan"],
  astronomi: ["astronomi", "heyet", "felek", "zic", "riyazi", "matematik", "hendese", "cebir"],
  mûsiki: ["musiki", "bestekar", "makam", "beste"],
  hat: ["hatat", "kitabet", "tezhip"],
};

const STEMS: ReadonlyArray<readonly [string, string]> = Object.entries(FIELDS).flatMap(([field, stems]) =>
  stems.map((stem) => [collapseLetters(stem), field] as const),
);

const SHORT_ENDINGS = /^(ta|da|ti|i|in|ki|ci|ta|tan|dan|la|le|bi)$/;

/** The sentence around a mention (between . ! ? ; boundaries), for its fields. */
export function sentenceAround(text: string, start: number, length: number): string {
  const before = text.slice(0, start);
  const after = text.slice(start + length);
  const from = Math.max(before.lastIndexOf(". "), before.lastIndexOf("! "), before.lastIndexOf("? "), before.lastIndexOf("; "));
  const ends = [". ", "! ", "? ", "; "].map((m) => after.indexOf(m)).filter((k) => k >= 0);
  const to = ends.length ? Math.min(...ends) : after.length;
  return text.slice(from < 0 ? 0 : from + 2, start + length + to);
}

/** Fields a text mentions. Stems match word starts ("tefsirinde" → tefsir); short stems must match whole words. */
export function fieldsOf(text: string): Set<string> {
  const out = new Set<string>();
  for (const word of words(text).trim().split(" ")) {
    if (word.length < 3) continue;
    for (const [stem, field] of STEMS) {
      if (stem.length <= 3 ? word === stem || SHORT_ENDINGS.test(word.slice(stem.length)) && word.startsWith(stem) : word.startsWith(stem)) out.add(field);
    }
  }
  return out;
}

/** Number of fields shared by a sentence and a description. */
export function sharedFields(sentence: ReadonlySet<string>, description: string): string[] {
  return [...fieldsOf(description)].filter((f) => sentence.has(f));
}
