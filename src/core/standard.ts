/**
 * Standart yazım: names from outside DİA (Wikidata) written the way DİA writes
 * them. The word list is DİA's own: every capitalised word of its short and
 * full names, under a key without circumflexes, macrons, the ayn and Turkish
 * letters ("Türâbî", "Turabi" → "turabi"). A word is rewritten only when DİA
 * writes it one way clearly more often than any other ("Yusuf" → "Yûsuf",
 * "Abdulaziz" → "Abdülazîz"); a word DİA writes both ways ("Ali" / "Alî") or
 * does not have is left as it is. The Arabic article is assimilated as in DİA
 * ("el-Turabi" → "et-Turabi", "al-Nur" → "en-Nur") and "bin" / "ibn" between
 * names becomes "b.".
 *
 * The Arabic spelling tells apart names Latin letters write alike: DİA's short
 * names are paired word by word with their Arabic original ("Hâmid" حامد,
 * "Hamîd" حميد). Given the Arabic name of the person, the form DİA uses for
 * that Arabic word is taken; without it, a key that stands for two different
 * Arabic words is left alone.
 */
import type { Person } from "./catalog.js";

export interface SpellingStandard {
  /** Word key → DİA's form ("" when DİA writes it more than one way). */
  readonly forms: ReadonlyMap<string, string>;
  /** Word key + Arabic skeleton → DİA's form for that Arabic word. */
  readonly arabic: ReadonlyMap<string, string>;
  /** Word keys that stand for more than one Arabic word in DİA (Hâmid / Hamîd). */
  readonly split: ReadonlySet<string>;
}

/** An Arabic (or Persian) word without vowel marks, the article and letter variants: "الحميد" → "حميد". */
export function arabicSkeleton(word: string): string {
  const plain = word
    .replace(/[\u064B-\u065F\u0670\u0640]/g, "")
    .replace(/[أإآٱ]/g, "ا")
    .replace(/ة/g, "ه")
    .replace(/[ىی]/g, "ي")
    .replace(/ک/g, "ك");
  return plain.length > 3 && plain.startsWith("ال") ? plain.slice(2) : plain;
}

const isArabicLink = (w: string): boolean => w === "بن" || w === "ابن" || w === "بنت";
const isLatinLink = (w: string): boolean => /^(b\.|bin|ibn|ibni|bint)$/i.test(w);

/** The Latin words of a name paired with its Arabic words, when both have as many (links left out). */
export function pairWords(latin: string, arabic: string): Array<[string, string]> | null {
  const l = latin.split(/\s+/).filter((w) => w && !isLatinLink(w));
  const a = arabic.split(/\s+/).filter((w) => w && !isArabicLink(w));
  if (l.length === 0 || l.length !== a.length || !/[\u0600-\u06FF]/.test(arabic)) return null;
  return l.map((w, k) => [w, a[k] ?? ""]);
}

/** "Türâbî", "Turabi", "Tûrâbî" → "turabi". */
export function spellingKey(word: string): string {
  return word
    .toLocaleLowerCase("tr")
    .replace(/ı/g, "i")
    .normalize("NFD")
    .replace(/[̀-ͯ]/g, "")
    .replace(/[‘’'ʿʾ`]/g, "");
}

const ARTICLE = /^(?:el|al|er|ar|et|at|es|as|eş|aş|ed|ad|ez|az|en|an)-(.+)$/i;
/** Turkish letters of the Arabic "sun letters": the article takes their sound. */
const SUN = new Set(["t", "s", "ş", "d", "z", "r", "n", "l"]);

/** The DİA spelling of each word key, from the DİA people of a catalog. */
export function buildSpellingStandard(people: readonly Person[], dominance = 0.7): SpellingStandard {
  const counts = new Map<string, Map<string, number>>();
  const count = (text: string, weight: number): void => {
    if (!text || text === "---") return;
    for (const raw of text.replace(/[,()]/g, " ").split(/\s+/)) {
      const word = raw.replace(ARTICLE, "$1");
      if (word.length < 2 || !/^\p{Lu}/u.test(word) || /[.\d]/.test(word)) continue;
      const key = spellingKey(word);
      const forms = counts.get(key) ?? new Map<string, number>();
      forms.set(word, (forms.get(word) ?? 0) + weight);
      counts.set(key, forms);
    }
  };
  for (const person of people) {
    if (!person.source.startsWith("https://islamansiklopedisi.org.tr/")) continue;
    // Short names are DİA's usual spelling; full names (transliterated) weigh less.
    count(person.shortName, 3);
    count(person.fullName, 1);
  }
  const out = new Map<string, string>();
  for (const [key, forms] of counts) {
    let best = "";
    let top = 0;
    let total = 0;
    for (const [form, n] of forms) {
      total += n;
      if (n > top) {
        best = form;
        top = n;
      }
    }
    // A word DİA writes more than one way is known ("") but not rewritten.
    out.set(key, top / total >= dominance ? best : "");
  }
  // DİA's short names with their Arabic originals, word by word.
  const pairs = new Map<string, Map<string, number>>();
  const skeletons = new Map<string, Set<string>>();
  for (const person of people) {
    if (!person.source.startsWith("https://islamansiklopedisi.org.tr/") || person.shortName.includes(",")) continue;
    for (const [latin, arabic] of pairWords(person.shortName, person.originalName) ?? []) {
      const word = latin.replace(ARTICLE, "$1");
      if (word.length < 2 || !/^\p{Lu}/u.test(word)) continue;
      const key = spellingKey(word);
      const skeleton = arabicSkeleton(arabic);
      const id = `${key}|${skeleton}`;
      const forms = pairs.get(id) ?? new Map<string, number>();
      forms.set(word, (forms.get(word) ?? 0) + 1);
      pairs.set(id, forms);
      const seen = skeletons.get(key) ?? new Set<string>();
      seen.add(skeleton);
      skeletons.set(key, seen);
    }
  }
  const arabic = new Map<string, string>();
  for (const [id, forms] of pairs) {
    const [best] = [...forms.entries()].sort((x, y) => y[1] - x[1]);
    if (best) arabic.set(id, best[0]);
  }
  // Two Arabic words under one key, written differently in DİA: a real difference of names.
  const split = new Set<string>();
  for (const [key, seen] of skeletons) {
    const written = new Set([...seen].map((sk) => arabic.get(`${key}|${sk}`)));
    if (seen.size > 1 && written.size > 1) split.add(key);
  }
  return { forms: out, arabic, split };
}

/** The article as DİA writes it before the word: "et-Türâbî", "el-Cezerî". */
function article(word: string): string {
  const first = spellingKey(word).charAt(0);
  const sun = word.toLocaleLowerCase("tr").startsWith("ş") ? "ş" : SUN.has(first) ? first : "l";
  return `e${sun}-`;
}

/** What DİA's spelling may add to a letter: length marks and the Turkish front vowels. */
const ADDS: Record<string, string> = {
  a: "âā", e: "ē", i: "îī", o: "öō", u: "ûūü",
  A: "ÂĀ", I: "ÎĪ", O: "ÖŌ", U: "ÛŪÜ",
  â: "ā", î: "ī", û: "ū", ü: "ū",
};

/**
 * DİA's form only adds to the written word: marks of length, ü/ö, the ayn. A
 * form that drops or changes a letter ("Kürt" → "Kurt", "Smith" → "Smıth") is
 * not taken.
 */
export function onlyAdds(written: string, form: string): boolean {
  const a = [...written.normalize("NFC")];
  const b = [...form.normalize("NFC").replace(/[‘’ʿʾ]/g, "")];
  if (a.length !== b.length) return false;
  return a.every((ch, k) => ch === b[k] || (ADDS[ch] ?? "").includes(b[k] ?? ""));
}

function standardWord(word: string, standard: SpellingStandard, arabic = ""): string {
  if (!/^\p{Lu}/u.test(word) || NOT_ARABIC.has(spellingKey(word))) return word;
  const key = spellingKey(word);
  const paired = arabic ? standard.arabic.get(`${key}|${arabicSkeleton(arabic)}`) : undefined;
  // Without the Arabic word, a key DİA uses for two different names is not guessed.
  const form = paired ?? (standard.split.has(key) ? "" : standard.forms.get(key));
  return form && onlyAdds(word, form) ? form : word;
}

/** Words that look like DİA's but name someone else: the Greek "Mihail" is not the Arabic "Mîhâîl". */
const NOT_ARABIC = new Set(["mihail"]);

/**
 * The name in DİA's spelling. Only a name at least `share` of whose words DİA
 * has (an article or "bin" counts) is rewritten: "Mihail Attaleiates" stays.
 */
export function standardName(name: string, standard: SpellingStandard, share = 0.5, arabicName = ""): string {
  const words = name.split(" ");
  const names = words.filter((w) => /^\p{Lu}/u.test(w) || ARTICLE.test(w) || /^(bin|ibn|ibni)$/i.test(w));
  const known = names.filter((w) => ARTICLE.test(w) || /^(bin|ibn|ibni)$/i.test(w) || (standard.forms.has(spellingKey(w)) && !NOT_ARABIC.has(spellingKey(w))));
  if (names.length === 0 || known.length < names.length * share) return name;
  // The Arabic word of each Latin word, when the two names pair up.
  const arabicOf = new Map<number, string>();
  const pairs = arabicName ? pairWords(name, arabicName) : null;
  if (pairs) {
    let k = 0;
    words.forEach((word, index) => {
      if (!word || isLatinLink(word)) return;
      arabicOf.set(index, pairs[k]?.[1] ?? "");
      k += 1;
    });
  }
  const out: string[] = [];
  words.forEach((word, index) => {
    const lower = word.toLocaleLowerCase("tr");
    const withArticle = ARTICLE.exec(word);
    if (withArticle && withArticle[1]) {
      const core = standardWord(withArticle[1].charAt(0).toLocaleUpperCase("tr") + withArticle[1].slice(1), standard, arabicOf.get(index));
      out.push(article(core) + core);
    } else if (index > 0 && index < words.length - 1 && (lower === "bin" || lower === "ibn" || lower === "ibni")) {
      out.push("b.");
    } else {
      out.push(standardWord(word, standard, arabicOf.get(index)));
    }
  });
  return out.join(" ");
}

/**
 * Whose names take DİA's spelling: citizens (Wikidata P27) of the Muslim world
 * (the Organisation of Islamic Cooperation, Mandatory Palestine; not India,
 * whose people are of many faiths), and Ottoman subjects who died before the
 * Republic's spelling took over (1928). Westerners, Turks of the Republic and
 * everyone else keep their usual spelling.
 */
export const STANDARD_COUNTRIES: ReadonlySet<string> = new Set([
  "Q889", "Q222", "Q262", "Q227", "Q398", "Q902", "Q225", "Q921", "Q965", "Q657", "Q970", "Q977", "Q79", "Q1005", "Q1006",
  "Q252", "Q794", "Q796", "Q810", "Q232", "Q1246", "Q817", "Q813", "Q822", "Q1016", "Q833", "Q826", "Q912", "Q1025", "Q1028",
  "Q1032", "Q1033", "Q842", "Q843", "Q219060", "Q846", "Q851", "Q1041", "Q1044", "Q1045", "Q1049", "Q858", "Q863", "Q948",
  "Q874", "Q878", "Q265", "Q805", "Q193714", "Q23681",
]);
export const OTTOMAN = "Q12560";
const OTTOMAN_UNTIL = 1928;

export function wantsStandard(countries: readonly string[], deathYear: number | null): boolean {
  if (countries.some((c) => STANDARD_COUNTRIES.has(c))) return true;
  return countries.includes(OTTOMAN) && deathYear !== null && deathYear < OTTOMAN_UNTIL;
}

const built = new WeakMap<readonly Person[], SpellingStandard>();
/** The standard of a catalog's DİA people, built once per catalog. */
export function diaStandard(people: readonly Person[]): SpellingStandard {
  let standard = built.get(people);
  if (!standard) {
    standard = buildSpellingStandard(people);
    built.set(people, standard);
  }
  return standard;
}

/**
 * The person with a standard short name; the spelling it had is kept among its
 * names (" | "), so a text written that way still finds it.
 */
export function standardPerson(person: Person, standard: SpellingStandard, share = 0.5): Person {
  const shortName = standardName(person.shortName, standard, share, person.originalName);
  if (shortName === person.shortName) return person;
  const others = person.fullName && person.fullName !== "---" ? person.fullName.split(" | ") : [];
  const fullName = [person.shortName, ...others.filter((n) => n !== person.shortName && n !== shortName)].join(" | ");
  return { ...person, shortName, fullName };
}
