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
import { nearWord } from "./names.js";

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

/** "عبد العزيز" and "نصر الله" are one name each, as DİA's "Abdülazîz" is one word. */
function joinArabicAbd(words: string[]): string[] {
  const out: string[] = [];
  for (let k = 0; k < words.length; k++) {
    const w = words[k] as string;
    if (w === "عبد" && k + 1 < words.length) {
      out.push(w + (words[k + 1] as string));
      k++;
    } else if (words[k + 1] === "الله") {
      // "نصر الله" is Nasrallah, one word.
      out.push(w + "الله");
      k++;
    } else out.push(w);
  }
  return out;
}

/** The Latin words of a name paired with its Arabic words, when both have as many (links left out). */
export function pairWords(latin: string, arabic: string): Array<[string, string]> | null {
  const l = latin.split(/\s+/).filter((w) => w && !isLatinLink(w));
  const a = joinArabicAbd(arabic.split(/\s+/).filter((w) => w && !isArabicLink(w)));
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

/** "Abdul" on its own: DİA never writes it apart, so it is not rewritten alone. */
const ABD_ALONE = new Set(["abd", "abdul", "abdel", "abdal", "abdur", "abdus"]);

function standardWord(word: string, standard: SpellingStandard, arabic = "", long = false): string {
  if (!/^\p{Lu}/u.test(word) || NOT_ARABIC.has(spellingKey(word)) || ABD_ALONE.has(spellingKey(word))) return word;
  const key = spellingKey(word);
  const paired = arabic ? standard.arabic.get(`${key}|${arabicSkeleton(arabic)}`) : undefined;
  // Without the Arabic word, a key DİA uses for two different names is not guessed.
  const form = paired ?? (standard.split.has(key) ? "" : standard.forms.get(key));
  if (form && onlyAdds(word, form)) {
    // DİA's form for the key, but the person's Arabic says otherwise: "Hassan" حسن is Hasan, not Hassân حسّان.
    const read = !paired && arabic ? longVowels(word, arabic) : null;
    if (read !== null && lengthMarks(read) !== lengthMarks(form)) return long && onlyAdds(word, read) ? read : word;
    return form;
  }
  // A word DİA does not have, or writes two ways: its long vowels read from the Arabic ("Turabi", الترابي → "Turâbî").
  if (long && arabic && !form) {
    const lengthened = longVowels(word, arabic);
    if (lengthened && onlyAdds(word, lengthened)) return lengthened;
  }
  return word;
}

/** Where a word has long vowels: "Hassân" → "....L.". */
function lengthMarks(word: string): string {
  return [...word.normalize("NFC").replace(/[‘’ʿʾ]/g, "")].map((ch) => ("âîûāīūÂÎÛĀĪŪ".includes(ch) ? "L" : ".")).join("");
}

/** Latin consonants and the Arabic letters each may stand for. */
const CONSONANTS: Record<string, string> = {
  b: "ب", p: "پب", t: "تطة", s: "سصث", c: "ج", ç: "چج", h: "هحخ", d: "دض", z: "زذظض", r: "ر", ş: "ش", g: "غقگ", ğ: "غ",
  f: "ف", k: "كقخ", l: "ل", m: "م", n: "ن", v: "و", y: "ي", j: "ژج",
};
const VOWELS = "aeiıoöuü";
/** Hamza seats and the ayn: written in Arabic, silent in the Latin word. */
const SILENT = "عءئؤ";

/**
 * The long vowels of a Latin word read from its Arabic spelling: an "a" over
 * an alif is "â", an "i" over a yâ "î", a "u" over a vâv "û" ("Turabi" +
 * الترابي → "Turâbî"). Null when the two do not line up letter by letter.
 * Measured on DİA's own short names, it writes 95% of the rare Arabic words as
 * DİA does; Turkic names (whose short vowels Arabic letters also show) are
 * left to the callers, who use it only for the Arab world, Iran, Afghanistan
 * and Pakistan.
 */
export function longVowels(word: string, arabicWord: string): string | null {
  let a = arabicWord
    .replace(/[\u064B-\u065F\u0670\u0640]/g, "")
    .replace(/[أإٱ]/g, "ا")
    .replace(/ة/g, "ه")
    .replace(/[ىی]/g, "ي")
    .replace(/ک/g, "ك");
  if (a.length > 3 && a.startsWith("ال")) a = a.slice(2);
  const letters = [...word.normalize("NFC")];
  const lower = letters.map((ch) => ch.toLocaleLowerCase("tr"));
  const out: string[] = [];
  let j = 0;
  for (let i = 0; i < letters.length; i++) {
    const ch = lower[i] as string;
    const written = letters[i] as string;
    // An ayn between a vowel and what follows is silent in the Latin word: "Said", سعيد.
    if (VOWELS.includes(ch) && i > 0 && a[j] === "ع") j++;
    const next = a[j] ?? "";
    if (VOWELS.includes(ch)) {
      const upper = written !== ch;
      if (next === "آ" && (ch === "a" || ch === "e")) {
        out.push(upper ? "Â" : "â");
        j++;
      } else if (i === 0 && (next === "ا" || next === "ع")) {
        out.push(written);
        j++;
      } else if (next === "ا" && (ch === "a" || ch === "e")) {
        out.push(upper ? "Â" : "â");
        j++;
      } else if (next === "ي" && (ch === "i" || ch === "ı") && lower[i + 1] !== "y" && !(lower[i + 1] && VOWELS.includes(lower[i + 1] as string))) {
        out.push(upper ? "Î" : "î");
        j++;
      } else if (next === "و" && "uüoö".includes(ch) && lower[i + 1] !== "v" && !(lower[i + 1] && VOWELS.includes(lower[i + 1] as string))) {
        out.push(upper ? "Û" : "û");
        j++;
      } else {
        if (SILENT.slice(1).includes(next)) j++;
        out.push(written);
      }
      continue;
    }
    const options = CONSONANTS[ch];
    if (!options) return null;
    while (j < a.length && SILENT.includes(a[j] as string) && !options.includes(a[j] as string)) j++;
    if (j >= a.length || !options.includes(a[j] as string)) return null;
    j++;
    out.push(written);
    // A doubled consonant is one Arabic letter (with shadda).
    if (lower[i + 1] === ch) {
      out.push(letters[i + 1] as string);
      i++;
    }
  }
  while (j < a.length && (SILENT + "ه").includes(a[j] as string)) j++;
  return j === a.length ? out.join("") : null;
}

const FOREIGN_SURNAME = /(yan|yân|ian|ov|ev|opoulos|poulos|idis|akis|escu|ović|evič|evich|ski|sky|ska)$/i;

/** Words that look like DİA's but name someone else: the Greek "Mihail" is not the Arabic "Mîhâîl". */
const NOT_ARABIC = new Set(["mihail"]);

/**
 * The name in DİA's spelling. Only a name at least `share` of whose words DİA
 * has (an article or "bin" counts) is rewritten: "Mihail Attaleiates" stays.
 */
/**
 * "Abdul Aziz", "Abd al-Rahman", "Abdel Kader": one name in DİA ("Abdülazîz",
 * "Abdurrahmân"). Joined only when DİA has the joined word.
 */
function joinAbd(words: string[], standard: SpellingStandard): string[] {
  const out: string[] = [];
  for (let k = 0; k < words.length; k++) {
    const w = words[k] as string;
    const next = words[k + 1];
    if (next && /^(abd|abdul|abdül|abdel|abdal|abdur|abdür|abdus|abdüs)$/i.test(w)) {
      const rest = next.replace(ARTICLE, "$1");
      const first = spellingKey(rest).charAt(0);
      const link = SUN.has(first) ? `u${first}` : "ul";
      const keys = [`abd${link}${spellingKey(rest)}`, `abdul${spellingKey(rest)}`];
      const key = keys.find((k) => standard.forms.has(k));
      if (key) {
        // DİA's joined form; one it writes two ways is joined and left to the Arabic ("Abdülhalim").
        const sound = link === "ul" ? "ül" : link.replace("u", "ü");
        out.push(standard.forms.get(key) || `Abd${sound}${rest.toLocaleLowerCase("tr")}`);
        k++;
        continue;
      }
    }
    out.push(w);
  }
  return out;
}

export function standardName(name: string, standard: SpellingStandard, share = 0.5, arabicName = "", long = false): string {
  const words = joinAbd(name.split(" "), standard);
  const names = words.filter((w) => /^\p{Lu}/u.test(w) || ARTICLE.test(w) || /^(bin|ibn|ibni)$/i.test(w));
  const known = names.filter((w) => ARTICLE.test(w) || /^(bin|ibn|ibni)$/i.test(w) || (standard.forms.has(spellingKey(w)) && !NOT_ARABIC.has(spellingKey(w))));
  if (names.length === 0 || known.length < names.length * share) return name;
  // An Armenian, Slavic, Greek or Romanian surname: a Christian or a Soviet-era name, spelled as its people spell it
  // ("Serkis Minasyan", "Kuzman Şapkarev", "Tahir Salahov"), not in DİA's manner.
  if (words.slice(1).some((w) => w.length >= 5 && /^\p{Lu}/u.test(w) && FOREIGN_SURNAME.test(w))) return name;
  // The Arabic word of each Latin word, when the two names pair up.
  const arabicOf = new Map<number, string>();
  const pairs = arabicName ? pairWords(words.join(" "), arabicName) : null;
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
      const core = standardWord(withArticle[1].charAt(0).toLocaleUpperCase("tr") + withArticle[1].slice(1), standard, arabicOf.get(index), long);
      // DİA does not begin a short name with the article: "Ad-Dardir" is "Dardir".
      out.push(index === 0 ? core : article(core) + core);
    } else if (index > 0 && index < words.length - 1 && (lower === "bin" || lower === "ibn" || lower === "ibni")) {
      out.push("b.");
    } else {
      out.push(standardWord(word, standard, arabicOf.get(index), long));
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

/**
 * Where long vowels may be read from the Arabic spelling of a word DİA does not
 * have: the Arab world, Iran, Afghanistan, Pakistan. (Turkic names write their
 * short vowels with letters too, so not there.)
 */
export const LONG_VOWEL_COUNTRIES: ReadonlySet<string> = new Set([
  "Q262", "Q398", "Q970", "Q977", "Q79", "Q796", "Q810", "Q817", "Q822", "Q1016", "Q1025", "Q1028", "Q842", "Q219060",
  "Q846", "Q851", "Q1045", "Q1049", "Q858", "Q948", "Q878", "Q805", "Q193714", "Q794", "Q889", "Q843",
]);

export function wantsLongVowels(countries: readonly string[]): boolean {
  return countries.some((c) => LONG_VOWEL_COUNTRIES.has(c));
}

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

/**
 * Countries that write their people's names in Latin letters, as Turkish keeps
 * them ("Schleiermacher", "Descartes"): Western Europe, the Americas, Oceania
 * and the states before them.
 */
export const LATIN_WEST: ReadonlySet<string> = new Set([
  "Q183", "Q40", "Q39", "Q142", "Q31", "Q55", "Q145", "Q161885", "Q174193", "Q27", "Q30", "Q16", "Q408", "Q664", "Q38",
  "Q29", "Q45", "Q34", "Q20", "Q35", "Q33", "Q189", "Q36", "Q213", "Q214", "Q28", "Q27306", "Q12548", "Q151624", "Q43287",
  "Q41304", "Q7318", "Q70972", "Q28513", "Q131964", "Q172579", "Q179876", "Q230791", "Q155", "Q414", "Q96", "Q298", "Q32",
  "Q154195", "Q153015", "Q4948", "Q170072",
]);

const TURKISH_LETTERS = /[çşğıöüİÇŞĞÖÜ]/;

/**
 * A slip in a Western name's Turkish label ("Friedrich Scleirmacher"): for a
 * person all of whose citizenships are of Latin-alphabet countries, a label written with no Turkish letter,
 * the same as no other language's, one or two letters from a spelling that at
 * least two languages share ("Friedrich Schleiermacher" in German, English,
 * French …) is that spelling. Turkish renderings ("Mihail Lomonosov", "Platon")
 * are left: their people are not of those countries, or another language
 * writes them so. Null when nothing is to change.
 */
export function westernSpelling(written: string, labels: readonly string[], countries: readonly string[]): string | null {
  // Every citizenship Western: an Ottoman Armenian who died in Paris ("Sogomon Tehliryan") keeps his Turkish rendering.
  if (countries.length === 0 || !countries.every((c) => LATIN_WEST.has(c)) || TURKISH_LETTERS.test(written)) return null;
  if (labels.includes(written)) return null;
  const words = written.split(" ").length;
  const key = spellingKey(written);
  const near = new Map<string, number>();
  for (const label of labels) {
    if (label.split(" ").length !== words || /[^\p{Script=Latin}\s.'’-]/u.test(label)) continue;
    const other = spellingKey(label);
    if (other !== key && nearWord(other, key, 2)) near.set(label, (near.get(label) ?? 0) + 1);
  }
  const [best] = [...near.entries()].sort((a, b) => b[1] - a[1]);
  return best && best[1] >= 2 ? best[0] : null;
}
