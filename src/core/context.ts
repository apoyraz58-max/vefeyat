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
  if (/^\s*\((ö|öl|v|d)\.\s*\d/.test(mid(window, localFirst + length + 1))) return false;
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
  // "… âlimleri Mâtürîdî idi", "Eş‘arî olarak bilinir": a predicate, not the imam.
  if (["idi", "idiler", "dir", "dirler", "dur", "oldu", "olan", "olarak", "oldugu", "olmustur", "sayilir", "kabul"].includes(firstWord(after))) return true;
  // Şâfiî, Eş‘arî and Mâtürîdî are eponyms: written alone, without a school
  // word or group noun after them (handled above and by isGroupAdjective),
  // they are the imam himself, also in a list of names. Hanefî, Mâlikî … are
  // adjectives (the imams are Ebû Hanîfe, Mâlik), so for them the label stays.
  return !(root === "safi" || root === "esari" || root === "maturidi");
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
  ("filozof|doktrin|ogreti|dusunur|sair|han|hukumdar|aile|dai|fikir|fikr|cereyan|cizgi|dil|lehce|edebiyat|siir|uslub|sebk|tavir|cevre|emir|beg|bey|" +
  "tarikat|dervis|seyh|sufi|mutasavvif|silsile|tekke|dergah|zaviye|asitane|ayin|mensup|mensub|mezhep|mezheb|ekol|firka|cemaat|zumre|topluluk|cevre|gelenek|geleneg|anlayis|dusunce|kelam|fikih|fikh|akide|akaid|itikad|inanc|alim|ulema|fakih|fukaha|kelamci|mutekellim|muellif|muhaddis|muctehid|halife|hanedan|devlet|sultan|emir|vezir|ordu|donem|devir|devri|saray|mimari|sanat|uslup|uslub|edebiyat|hareket|akim|egilim|yorum|okul|kol|mesrep|mesreb|adab|erkan|kultur|toplum|kesim|grup|taraftar|camia|muhit|nesil|kusak|vakif|medrese|cami|mescit|imparatorluk|beylik|aile|soy|boy|kabile|asiret|tarih|kaynak|metin|literatur|risale|eser")
    .split("|")
    .map(collapseLetters); // compared with words(), which collapses doubled letters

const TITLES = new Set("imam|seyh|hz|hazret|hazreti|mevlana|hace|haci|molla|baba|dede|sultan|emir|kadi|hafiz|alim|fakih|muhaddis|sair|mutasavvif|ebu|ibn|ibni".split("|").map(collapseLetters));

const key = (value: string): string => words(value).trim();

/**
 * Nisbas that are also names of orders, schools, sects, dynasties and
 * regions; only these are read as adjectives before a group noun. A proper
 * nisba like Gazzâlî stays a person in "Gazzâlî medrese hocalığı yaptı".
 */
const GROUP_NISBAS = new Set(
  ("Nakşibendî Kādirî Kâdirî Rifâî Mevlevî Bektaşî Halvetî Sühreverdî Çiştî Şâzelî Celvetî Bayramî Gülşenî Şa‘bânî Uşşâkī Sünbülî Cerrâhî Kübrevî Yesevî Ekberî Melâmî Kalenderî Haydarî Hurûfî Nûrbahşî " +
    "Sünnî Şiî Abbâsî Fâtımî Emevî Eyyûbî Selçukî Memlûkî Safevî Kaçar Osmanlı Gaznelî Karahanlı Tolunî Ihşîdî Hamdânî Büveyhî Muvahhidî Murâbıtî Ağlebî İdrîsî Rüstemî Zeydî İsmâilî Câferî Nizârî Karmatî Hâricî İbâzî Vehhâbî Selefî Mu‘tezilî Mutezilî " +
    "Suûdî Suudî Özbek Sâmî Hint Moğol Tatar Kıpçak Oğuz Türkmen Kürt Arap Acem Fars " +
    "Hanefî Şâfiî Mâlikî Hanbelî Zâhirî Eş‘arî Mâtürîdî Sûfî Hindî Yemenî Endelüsî Mısrî Şâmî Irâkī Kureşî Arabî Türkî Fârisî Horasânî Mağribî")
    .split(" ")
    .map(key),
);

const OCCUPATIONS = new Set(
  "Kadı Kādî Fakih Âlim Edip Edîb Şâir Muhaddis Müfessir Kâtib Hattat Tabip Hekim Müderris Vâiz İmam Şeyh Mutasavvıf Mütekellim Mukri Hâfız Müezzin Mimar Mûsikişinas"
    .split(" ")
    .map(key),
);

/** Catalog spellings that are also ordinary adjectives or nouns: "Usûlî açıdan", "Hakîm bir kişi". */
const DICTIONARY_WORDS = new Set(
  ("Usûlî Usûl Kelâmî Fıkhî Mantıkî Felsefî Lügavî Nahvî Edîb Hakîm Kâtib Şâir Âlim Fakih Muhaddis Kebîr Emîn Mü’min Fahrî Bâkî Hâlî Kâmî Kânî Mecnûn Musannif Münşî Muallimî Kâşî Zâhidî Vahyî Vücûdî Şeyhî Hilâlî Nebâtî Riyâzî Hayâlî Harbî Sâbit Sürûr Derd Bahâr " +
    "Sünnî Şiî Sûfî Hindî Yemenî Endelüsî Mısrî Şâmî Kureşî Arabî Türkî Fârisî İsmâilî Selefî Zâhirî")
    .split(" ")
    .map(key),
);

// Endings a short stem may take ("soyu", "kolları", "dilinde"), back and front
// forms apart: "soy" + "le" is "şöyle", not a lineage; "kolay", "beyan" are other words.
const BACK_ENDINGS = new Set("|i|u|a|in|un|nin|nun|ina|una|ini|unu|inda|unda|indan|undan|da|ta|dan|tan|lar|lari|larin|larina|larinda|larindan|la|yla|ya|yi|yu|si|su".split("|"));
const FRONT_ENDINGS = new Set("|i|u|e|in|un|nin|nun|ine|une|ini|unu|inde|unde|inden|unden|de|te|den|ten|ler|leri|lerin|lerine|lerinde|lerinden|le|yle|ye|yi|yu|si|su".split("|"));

function startsGroupNoun(word: string): boolean {
  return GROUP_NOUNS.some((stem) => {
    if (!word.startsWith(stem)) return false;
    if (stem.length >= 4) return true;
    const ending = word.slice(stem.length);
    // The stem's last vowel picks the endings; "i" stands for ı and i alike.
    const last = /[aeiou](?=[^aeiou]*$)/.exec(stem)?.[0];
    if (last === "e") return FRONT_ENDINGS.has(ending);
    if (last === "i") return FRONT_ENDINGS.has(ending) || BACK_ENDINGS.has(ending);
    return BACK_ENDINGS.has(ending);
  });
}

/**
 * A one-word nisba used as an adjective of a group, order, dynasty or school:
 * "Nakşibendî tarikatı", "Kādirî ve Rifâî tarikatları", "Abbâsî halifesi",
 * "Sünnî ve Şiî müellifler"; also an izafet like "Ehl-i hadîs". A title
 * before it ("Şeyh", "İmam") or a death date after it keeps the person.
 */
export function isGroupAdjective(text: string, first: number, length: number, literal: string): boolean {
  const bare = nameQuery(literal);
  // Persian izafet term: "sebk-i Hindî", "tarz-ı Arabî" (also with an ending).
  if (/\p{Ll}[-‐][iıuü]\s+$/u.test(text.slice(Math.max(0, first - 20), first))) return true;
  if (bare.length < literal.length) return false; // "Nakşibendî'nin": the person
  const after = text.slice(first + length);
  if (/^[-‐][iıuü]\s+\p{Ll}/u.test(after) || /[-‐][iıuü]$/.test(bare) && /^\s+\p{Ll}/u.test(after)) return true;
  if (/\s/.test(bare.trim())) return false; // "Bahâeddin Nakşibend": an explicit name
  if (!GROUP_NISBAS.has(key(bare))) return false;
  if (/^\s*\((ö|öl|v|d|m)\.|^\s*\(\d/.test(after)) return false;
  // A title right before it ("Şeyh Şâzelî", "İmam Şâfiî"), not across a comma ("Muhaddis, Şâfiî fakihi").
  const lead = text.slice(Math.max(0, first - 40), first);
  const before = words(lead).trim().split(" ");
  const prior = before[before.length - 1] ?? "";
  if (TITLES.has(prior) && /\p{L}[.\s]*$/u.test(lead)) return false;
  const clause = after.slice(0, 160).split(/[,;:.()]/)[0] ?? "";
  const next = words(clause).trim().split(" ").filter(Boolean);
  // "X tarikatı", "X ve Y tarikatları", "X, Y ve Z şeyhleri"
  for (let i = 0; i < Math.min(next.length, 7); i++) {
    const word = next[i] as string;
    if (startsGroupNoun(word)) return true;
    if (word === "ve" || word === "ile" || word === "veya") continue;
    // Another nisba of the list ("Kādirî ve Rifâî tarikatları"); stop at anything else.
    if (GROUP_NISBAS.has(word)) continue;
    break;
  }
  return false;
}

/** Words after a sentence-initial adjective: "Usûlî açıdan", "Usûlî bir yaklaşım", "Vücûdî tevhid". */
const ADJECTIVE_CONTEXT =
  "kisi|insan|adam|yazar|sahsiyet|acidan|acisindan|bakimdan|bakimindan|yonden|yonuyle|anlamda|anlami|anlamin|manada|manasi|sekilde|bicimde|tarzda|yaklasim|delil|mesele|tahlil|kaide|kavram|terim|ilke|yontem|metot|metod|tartisma|degerlendirme|cerceve|hukum|ilim|ilm|bilgi|kural|terminoloji|gorus|tefsir|serh|tevhid|baglam|duzey|seviye|boyut|nitelik|karakter|yapi|ozellik|unsur|kategori|tasnif|istilah|kullanim|eser|kitap|literatur|calisma|arastirma|soru|sorun|problem|konu"
    .split("|")
    .map(collapseLetters);

/** Sentence start: paragraph start, or after . ! ? : ; or an opening quote/bracket. */
export function isSentenceInitial(text: string, first: number): boolean {
  let i = first - 1;
  while (i >= 0 && /[\s\u00a0]/.test(text.charAt(i))) i--;
  if (i < 0) return true;
  const ch = text.charAt(i);
  if ("!?:;\"“«([‘'".includes(ch)) return true;
  if (ch !== ".") return false;
  return isBoundary(text, i + 1);
}

/**
 * A one-word catalog spelling that opens a sentence as an adjective:
 * "Usûlî açıdan bakıldığında…", "Hindî tefsir geleneği…". Evidence is
 * either the next word (a manner/topic noun, or "bir" + noun) or the same
 * word written in lower case elsewhere in the document (`lowerWords`: keys
 * of lower-case words, as words() writes them). A death date, a suffix or a
 * narrative verb right after keeps the person.
 */
export function isAdjectiveUse(text: string, first: number, length: number, literal: string, lowerWords: ReadonlySet<string> = new Set()): boolean {
  const bare = nameQuery(literal);
  if (bare.length < literal.length || /[\s‐-]/.test(bare.trim())) return false;
  const initial = isSentenceInitial(text, first);
  const after = text.slice(first + length, first + length + 120);
  // "Kadı, edip ve şair": a profession heading a list, no verb about a person.
  if (initial && OCCUPATIONS.has(key(bare)) && /^\s*(,|ve\s)/.test(after) && !isNarrative(words(after.slice(0, 80)))) return true;
  if (/^\s*\((ö|öl|v|d|m)\.|^\s*\(\d|^\s*[,;]/.test(after)) return false;
  const next = words(after.split(/[.;:!?()]/)[0] ?? "").trim().split(" ").filter(Boolean);
  const word = next[0];
  if (!word) return false;
  if (!/^\s*\p{Ll}/u.test(after)) return false; // "Usûlî Şeyh Ahmed": a name goes on
  if (isNarrative(` ${next.slice(0, 3).join(" ")} `)) return false;
  const own = key(bare);
  // Only words that really are ordinary words: listed, or written in lower case elsewhere in the document.
  if (!DICTIONARY_WORDS.has(own) && !lowerWords.has(own)) return false;
  const head = word === "bir" ? next[1] ?? "" : word;
  const context = ADJECTIVE_CONTEXT.some((stem) => head.startsWith(stem)) || GROUP_NOUNS.some((stem) => head.startsWith(stem));
  // Mid-sentence a capital usually means a name: only a manner/topic noun after it ("…, Usûlî açıdan") says otherwise.
  if (!initial) return context;
  if (lowerWords.has(own)) return true;
  return context;
}

/** Keys (as words() writes them) of every word written in lower case in the texts. */
export function lowerCaseWords(texts: readonly string[]): Set<string> {
  const out = new Set<string>();
  for (const text of texts) {
    for (const match of text.matchAll(/(?<![\p{L}\p{M}'’‘])\p{Ll}[\p{L}\p{M}]{2,}/gu)) {
      const key = words(match[0]).trim();
      if (key) out.add(key);
    }
  }
  return out;
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

/**
 * A heading-like line: short, without a closing period, and written in
 * Title Case or capitals ("Hayatı ve Eserleri", "IRAK'TAKİ MEDRESELER").
 * Capitals there say nothing about names.
 */
export function isTitleLike(text: string): boolean {
  const value = text.trim();
  if (value.length === 0 || value.length > 160 || /[.!?…]$/.test(value) && !/^\d+(\.\d+)*\.?\s/.test(value)) return false;
  const tokens = value.split(/\s+/).filter((w) => /\p{L}{3,}/u.test(w));
  if (tokens.length < 1) return false;
  if (value === value.toLocaleUpperCase("tr") && /\p{Lu}{3}/u.test(value)) return true;
  if (tokens.length < 2) return false;
  const upper = tokens.filter((w) => /^[^\p{L}]*\p{Lu}/u.test(w)).length;
  return upper / tokens.length >= 0.7;
}
